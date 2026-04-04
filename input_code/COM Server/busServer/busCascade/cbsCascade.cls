VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "cbsCascade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const m_strCMODULENAME As String = "cbsCascade"

' Must reset this after each public method or beginning
Private m_strErrorDescription As String
Private m_rsErrors As ADODB.Recordset
Private m_lngUserEmpNum As Long
Private m_rsEmployees As ADODB.Recordset


'****************************************************************
' Public Procedures
'****************************************************************

Public Function GetExceptionConflicts( _
                    ByVal v_rsEmpExcIntervals As ADODB.Recordset, _
                    ByVal v_lngScheduleSourceLevel As Long, _
                    ByRef r_rsExceptionsSPA As ADODB.Recordset, _
                    ByRef r_rsAvailSchedulesSPA As ADODB.Recordset, _
                    Optional ByVal v_rsWGBaseSchedChanges As ADODB.Recordset, _
                    Optional ByVal v_rsWGLineSchedChanges As ADODB.Recordset, _
                    Optional ByVal v_rsWGAssignChanges As ADODB.Recordset, _
                    Optional ByVal v_rsWGAssignOtherChanges As ADODB.Recordset, _
                    Optional ByVal v_rsEmpLineAssignChanges As ADODB.Recordset, _
                    Optional ByVal v_rsPersonalChanges As ADODB.Recordset, _
                    Optional ByVal v_rsLineTradeChanges As ADODB.Recordset, _
                    Optional ByVal v_rsEmpTradeChanges As ADODB.Recordset) _
                As ADODB.Recordset
' The purpose of this procedure is to get a conflict list, having cascade
' determine the new schedule with help from busSchedule.  This procedure combines prep,
' determination of the new schedule, building a list of exceptions to
' reschedule or delete, i.e. conflict list.
'
    Dim l_rsExceptionShifts As ADODB.Recordset
    Dim l_rsConflictList As ADODB.Recordset
    Dim l_lngScheduleSourceLevel As Long

On Error GoTo eh

    If v_lngScheduleSourceLevel = trWorkGroupBaseScheduleLevel And _
        Not v_rsWGLineSchedChanges Is Nothing Then
        l_lngScheduleSourceLevel = trDemandLineBaseSchedLevel
    Else
        l_lngScheduleSourceLevel = v_lngScheduleSourceLevel
    End If
' Must always consider line trades when affecting demand line schedules.
' Otherwise, line assignment changes may not pick up exceptions on a traded shift.
' e.g.(1) emp A on line 1, emp B on line 2, emp C on line 3.
' A traded with C on day 1.
' B has an exception on day 1 and is reassigned to line 1 (Emp A line).
' Must consider trades while updating the schedule and source, otherwise
' B exception will not be updated to traded shift, assumed from the line trade,
' resulting in schedule updated to line schedule, not the traded schedule.
' Also, must consider trades when getting candidate list of exceptions based
' on source.  Otherwise, if A is reassigned to line 2 and has an exception on the
' traded shift, it will not be picked up, resulting in not updating the exception.
    If l_lngScheduleSourceLevel = trDemandLineBaseSchedLevel Then
        l_lngScheduleSourceLevel = trDemandLineTrades
    End If

    Set l_rsExceptionShifts = _
        GetScheduledExceptionShifts(v_rsEmpExcIntervals, l_lngScheduleSourceLevel)

'    Guard against no exceptions scheduled
    If Not RecordsExist(l_rsExceptionShifts) Then
        GoTo xit
    End If

' Determine new schedule code
    UpdateNewScheduleShift l_lngScheduleSourceLevel, v_rsEmpExcIntervals, l_rsExceptionShifts, _
            v_rsWGBaseSchedChanges, v_rsWGLineSchedChanges, _
            v_rsWGAssignChanges, v_rsWGAssignOtherChanges, v_rsEmpLineAssignChanges, _
            v_rsPersonalChanges, v_rsLineTradeChanges, v_rsEmpTradeChanges

' Build the conflict list
    Set l_rsConflictList = GetExceptionConflictList( _
                l_rsExceptionShifts, r_rsExceptionsSPA, r_rsAvailSchedulesSPA)
                
    Set GetExceptionConflicts = l_rsConflictList
    
xit:
    Set l_rsExceptionShifts = Nothing
    Set l_rsConflictList = Nothing
    GetObjectContext.SetComplete
    Exit Function

eh:
    GetObjectContext.SetAbort
    Set r_rsExceptionsSPA = Nothing
    Set r_rsAvailSchedulesSPA = Nothing
    Set l_rsExceptionShifts = Nothing
    Set l_rsConflictList = Nothing
    TrsRaiseError m_strCMODULENAME, "GetExceptionConflicts", , , , , True
End Function


Public Function GetScheduledExceptionShifts( _
                    ByVal v_rsEmpExcInterval As ADODB.Recordset, _
                    ByVal v_lngScheduleSourceLevel As Long) _
                As ADODB.Recordset
'
' Compare exception date with request date range and filter out when outside.
' This will prevent extra exceptions from being returned to client while allowing
' 1 db access to be used.
'
'1. Get all the dates exceptions are scheduled, but not processed,
' for 1 or more employees. This should also include the schedule code,
' schedule type and the schedule source(hiearchy level from which
' the schedule code was assigned).
'
'2. Get all employees schedule data for the date range from the first exception
' to the last exception found.  TRIS 4.0 will store the schedule source
' with the exception, however there will be a period that the source will
' have to be determined using GetAdjustedBaseSchedule.
'
'3. Determine the schedule source of the exception, when necessary.
' This is done by matching the schedule code and charge date of the schedule
' and exception.
'
'4. Filter the exception date list before returning by comparing the
' schedule source.  If the schedule source of the exception is
' equal to or lower than the given schedule source, add it to the
' list of exceptions that need to be cascaded.  This means that any exception
' scheduled for a double shift will be filtered out.  Currently, only know
' of 1 exception for double shift; HDS.  This process will return exceptions
' that do not have a matching schedule and the source/level is not valid on
' the schedule.  In these scenarios, the SCHED_SRC_CD will be returned zero.
' Cascade will have to deal with these conditions later.
'
' Should be able to remove the 20(double shifts) from the db request and
' simply use the level that is on the schedule data returned from GetAdjustedBaseSchedule.
' For each exception, lookup up the schedule record based on emp_num,
' charge_date and sched_cd.  The double will have a level of 42 and therefore
' will be higher than the given level and skipped.
'
' Client will provide the NEW_SCHED_CD and NEW_SCHED_SRC_CD
'
' Trade client will set v_lngScheduleSourceLevel to 50 (line or employee)
'
    Dim l_cdCascade As dbCascade.cdCascade
    Dim l_cbsSchedule As busSchedule.cbsSchedule
    Dim l_rsExcDates As ADODB.Recordset
    Dim l_rsEmpSchedules As ADODB.Recordset
    Dim l_rsFilteredExcDates As ADODB.Recordset
    Dim l_varEmployees As Variant   'array of longs
    Dim l_intSchedSourceExc As Integer
    Dim l_varGabEmployees As Variant   'array of longs
    Dim l_datGabStartDate As Date
    Dim l_datGabEndDate As Date
    
    Dim l_datStartDate As Date
    Dim l_datEndDate As Date
    Dim l_blnValidRequest As Boolean
    
    
On Error GoTo eh

' Validate request and convert each array element into a long.
    l_blnValidRequest = ValidateGetSchedExcDtRequest(v_rsEmpExcInterval, l_varEmployees, l_datStartDate, l_datEndDate)
' how should the error be returned????
'
'
    Set l_cdCascade = CreateObject("dbCascade.cdCascade")
    Set l_rsExcDates = _
        l_cdCascade.GetScheduledExceptionShifts(l_varEmployees, l_datStartDate, l_datEndDate)
    Set l_cdCascade = Nothing
    
    Set l_rsFilteredExcDates = CreateSchedExcDtRecordset
'If there are no exceptions, just return empty output recordset
    If Not RecordsExist(l_rsExcDates) Then
        Set GetScheduledExceptionShifts = l_rsFilteredExcDates
        Exit Function
    End If

' The scope of this request may be reduced since the results are only used to
' determine the schedule source and TRIS 4.0 will store this on the exception.
    DetermineGetAdjustedBaseScope l_rsExcDates, _
        l_varEmployees, l_datStartDate, l_datEndDate, _
        l_varGabEmployees, l_datGabStartDate, l_datGabEndDate
    If Not IsEmpty(l_varGabEmployees) Then
        Set l_cbsSchedule = CreateObject("busSchedule.cbsSchedule")
        Set l_rsEmpSchedules = _
            l_cbsSchedule.GetAdjustedBaseSchedule( _
                l_varGabEmployees, l_datGabStartDate, l_datGabEndDate, Nothing, trBaseSchedNoDetails)
        Set l_cbsSchedule = Nothing
    End If

    Do Until l_rsExcDates.EOF
        l_intSchedSourceExc = DetermineSchedExcSource(l_rsExcDates, l_rsEmpSchedules)
        If l_intSchedSourceExc <= v_lngScheduleSourceLevel Then
            If IsExceptionDateInScope(v_rsEmpExcInterval, _
                    l_rsExcDates!EMP_NUM, l_rsExcDates!CHARGE_DATE) Then
                l_rsFilteredExcDates.AddNew
                l_rsFilteredExcDates!EMP_NUM = l_rsExcDates!EMP_NUM
                l_rsFilteredExcDates!CHARGE_DATE = l_rsExcDates!CHARGE_DATE
                ' Ensure shift num is valid
                l_rsFilteredExcDates!SHIFT_NUM = _
                    IIf(l_rsExcDates!SHIFT_NUM = 0, 1, l_rsExcDates!SHIFT_NUM)
                l_rsFilteredExcDates!SCHED_CD = l_rsExcDates!SCHED_CD
                
                'should SCHED_CD be set to zero when out of sync(l_intSchedSourceExc = 0) ?
                
                l_rsFilteredExcDates!SCHED_SRC_CD = l_intSchedSourceExc
                l_rsFilteredExcDates!NEW_SCHED_CD = 0
                l_rsFilteredExcDates!NEW_SCHED_SRC_CD = 0
                l_rsFilteredExcDates.Update
            End If
        End If
        l_rsExcDates.MoveNext
    Loop

    Set GetScheduledExceptionShifts = l_rsFilteredExcDates
    Set l_rsFilteredExcDates = Nothing
    Set l_rsExcDates = Nothing
    Set l_rsEmpSchedules = Nothing
    GetObjectContext.SetComplete
    Exit Function

eh:
    GetObjectContext.SetAbort
    Set l_cdCascade = Nothing
    Set l_cbsSchedule = Nothing
    Set l_rsExcDates = Nothing
    Set l_rsEmpSchedules = Nothing
    Set l_rsFilteredExcDates = Nothing
    TrsRaiseError m_strCMODULENAME, "GetScheduledExceptionShifts", , , , , True
End Function


Public Function GetExceptionConflictList( _
                    ByVal v_rsExcScheduledShifts As ADODB.Recordset, _
                    ByRef r_rsExceptions As ADODB.Recordset, _
                    ByRef r_rsAvailSchedules As ADODB.Recordset) _
                As ADODB.Recordset
' RED281
' v_rsExcScheduledShifts has the current and new scheduled shift for each employee/charge_date.
' If the new scheduled shift is zero, ignore it since it was not provided by the client.
' r_rsExceptions and r_rsAvailSchedules are output recordsets
'
'  Ensure do not delete an HB when becomes an off day since employee may want to bank
'  instead of paid.
'
'  Create a conflict list that returns a full shift exception for sets of partial shift exceptions
'  that when combined complete a full shift exception. i.e. a set of partial exceptions with the same
'  emp_num, charge_date, and sched_cd are considered a full shift exception when the sum of the
'  hours equal the total shift hours.
'
'  This process will be driven by shift, i.e. charge_date and sched_cd.

    Dim l_rsConflicts As ADODB.Recordset

On Error GoTo eh
' Get exception and schedule details
    Set r_rsExceptions = GetScheduledExceptions(v_rsExcScheduledShifts)
    Set r_rsAvailSchedules = GetAvailableSchedules(v_rsExcScheduledShifts)

    Set l_rsConflicts = CreateConflictRecordset

    Do Until v_rsExcScheduledShifts.EOF
        If Not BypassShift(v_rsExcScheduledShifts) Then
            ProcessShift l_rsConflicts, v_rsExcScheduledShifts, r_rsExceptions, r_rsAvailSchedules
        End If
        v_rsExcScheduledShifts.MoveNext
    Loop

    Set GetExceptionConflictList = l_rsConflicts
    Set l_rsConflicts = Nothing
    ClearFilter r_rsExceptions
    ClearFilter r_rsAvailSchedules

    GetObjectContext.SetComplete
    Exit Function

eh:
    GetObjectContext.SetAbort
    Set l_rsConflicts = Nothing
    Set r_rsExceptions = Nothing
    Set r_rsAvailSchedules = Nothing
    TrsRaiseError m_strCMODULENAME, "GetExceptionConflictList", , , , , True
End Function


Private Function BypassShift(ByRef r_rsExcScheduledShifts As ADODB.Recordset) As Boolean
' Bypass exception shifts that the client did not provide a new schedule.
' Bypass exception shifts that the schedule/source did not change.

On Error GoTo eh

    If r_rsExcScheduledShifts!NEW_SCHED_CD = 0 Or _
        (r_rsExcScheduledShifts!SCHED_CD = r_rsExcScheduledShifts!NEW_SCHED_CD And _
        r_rsExcScheduledShifts!SCHED_SRC_CD = r_rsExcScheduledShifts!NEW_SCHED_SRC_CD) Then
        BypassShift = True
    End If
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "BypassShift"
End Function


Private Function GetScheduledExceptions( _
            ByRef r_rsNewScheduleDates As ADODB.Recordset) As ADODB.Recordset
' Retrieve all exceptions in scope
    Dim l_varEmployees As Variant
    Dim l_datStartDate As Date
    Dim l_datEndDate As Date
    Dim l_cdCascade As dbCascade.cdCascade

On Error GoTo eh

    DetermineGetExceptionsScope r_rsNewScheduleDates, l_varEmployees, l_datStartDate, l_datEndDate
    
    GetEmployees l_varEmployees

    Set l_cdCascade = CreateObject("dbCascade.cdCascade")
    Set GetScheduledExceptions = _
        l_cdCascade.GetScheduledExceptions(l_varEmployees, l_datStartDate, l_datEndDate)

    Set l_cdCascade = Nothing
    Exit Function

eh:
    Set l_cdCascade = Nothing
    TrsRaiseError m_strCMODULENAME, "GetScheduledExceptions"
End Function


Private Function GetAvailableSchedules( _
            ByRef r_rsNewScheduleShifts As ADODB.Recordset) As ADODB.Recordset
' Retrieve all schedule details in scope
    Dim l_cdCascade As dbCascade.cdCascade
    Dim l_varSchedCodes As Variant

On Error GoTo eh

    l_varSchedCodes = DetermineGetSchedulesScope(r_rsNewScheduleShifts)
    
    Set l_cdCascade = CreateObject("dbCascade.cdCascade")
    Set GetAvailableSchedules = _
        l_cdCascade.GetAvailableSchedules(l_varSchedCodes)

    Set l_cdCascade = Nothing
    Exit Function

eh:
    Set l_cdCascade = Nothing
    TrsRaiseError m_strCMODULENAME, "GetAvailableSchedules"
End Function


Private Sub ProcessShift( _
                ByRef r_rsConflicts As ADODB.Recordset, _
                ByRef r_rsExcScheduledShifts As ADODB.Recordset, _
                ByRef r_rsExceptions As ADODB.Recordset, _
                ByRef r_rsAvailSchedules As ADODB.Recordset)
                
    Dim l_udtCurShiftDetails As ScheduledShift
    Dim l_udtNewShiftDetails As ScheduledShift
    Dim l_strExcCd As String
    Dim l_lngEmpNum As Long
    
On Error GoTo eh
    l_udtCurShiftDetails = _
        AllocateScheduledShift(r_rsAvailSchedules, _
            r_rsExcScheduledShifts!SCHED_CD, r_rsExcScheduledShifts!CHARGE_DATE)
    l_udtNewShiftDetails = _
        AllocateScheduledShift(r_rsAvailSchedules, _
            r_rsExcScheduledShifts!NEW_SCHED_CD, r_rsExcScheduledShifts!CHARGE_DATE)
'
' Do i need to know the shift number for a schedule shift?
'
    FilterOnShift r_rsExceptions, _
                r_rsExcScheduledShifts!EMP_NUM, _
                r_rsExcScheduledShifts!CHARGE_DATE, _
                r_rsExcScheduledShifts!SHIFT_NUM

    Do Until r_rsExceptions.EOF
        'Decision table processing with 4 possible scenarios
        'WORK/WORK - WORK/OFF - OFF/WORK - OFF/OFF
        l_strExcCd = Trim(r_rsExceptions!EXCEPTION_CD)
        If IsScheduledWorkDay(l_udtCurShiftDetails) Then
            If IsScheduledWorkDay(l_udtNewShiftDetails) Then    'WORK/WORK
                If l_udtCurShiftDetails.Hours = l_udtNewShiftDetails.Hours Then
                    'reschedule
                    'JimB 01/21/2003
                    'The If Statement was added to allow bad Holiday data to
                    ' be fixed by simply changing the schedule.
                    If IsAHolidayorHolidayWorked(l_strExcCd, r_rsExceptions!GROUP_IND) Then
                        AppendConflict r_rsConflicts, r_rsExcScheduledShifts, r_rsExceptions, _
                            CRESCHEDEXC, l_udtCurShiftDetails, l_udtNewShiftDetails, l_strExcCd
                    Else
                        AppendConflict r_rsConflicts, r_rsExcScheduledShifts, r_rsExceptions, _
                            CRESCHEDEXC, l_udtCurShiftDetails, l_udtNewShiftDetails
                    End If
                Else
                    'delete            'Jimb
                    l_lngEmpNum = r_rsExceptions!EMP_NUM
                    If FlexPolicyApplies(l_lngEmpNum) Then
                        If l_udtNewShiftDetails.Hours = 9 And l_udtCurShiftDetails.Hours = 8 And l_strExcCd = "H" Then
                            AppendConflict2 r_rsConflicts, r_rsExcScheduledShifts, r_rsExceptions, _
                                CRESCHEDEXC, l_udtCurShiftDetails, l_udtNewShiftDetails
                            AppendConflict3 r_rsConflicts, r_rsExcScheduledShifts, r_rsExceptions, _
                                "I", l_udtCurShiftDetails, l_udtNewShiftDetails, "FLX"
                        Else
                            If l_udtNewShiftDetails.Hours = 8 And l_udtCurShiftDetails.Hours = 9 And l_strExcCd = "H" Then
                                AppendConflict2 r_rsConflicts, r_rsExcScheduledShifts, r_rsExceptions, _
                                    CRESCHEDEXC, l_udtCurShiftDetails, l_udtNewShiftDetails
                            Else
                                AppendConflict r_rsConflicts, r_rsExcScheduledShifts, r_rsExceptions, _
                                    CDELETEEXC, l_udtCurShiftDetails, l_udtNewShiftDetails
                            End If
                        End If
                    
                    Else
                        AppendConflict r_rsConflicts, r_rsExcScheduledShifts, r_rsExceptions, _
                            CDELETEEXC, l_udtCurShiftDetails, l_udtNewShiftDetails
                    End If
                End If
            Else                                                'WORK/OFF
                If IsAHolidayorHolidayWorked(l_strExcCd, r_rsExceptions!GROUP_IND) Then
                    'reclassify HR
        'If HR Exception already on conflict list
        'then the next one is a delete   'jimb2
                    If Not DoesHRExist(r_rsConflicts, r_rsExceptions) Then
                        l_lngEmpNum = r_rsExceptions!EMP_NUM
                        If FlexPolicyApplies(l_lngEmpNum) Then
                             AppendConflict r_rsConflicts, r_rsExcScheduledShifts, r_rsExceptions, _
                                CRESCHEDEXC, l_udtCurShiftDetails, l_udtNewShiftDetails, CHOLIDAYBANKED
                        Else
                            AppendConflict r_rsConflicts, r_rsExcScheduledShifts, r_rsExceptions, _
                                CRESCHEDEXC, l_udtCurShiftDetails, l_udtNewShiftDetails, CHOLIDAYONREST
                        End If
                    Else
                        AppendConflict r_rsConflicts, r_rsExcScheduledShifts, r_rsExceptions, _
                        CDELETEEXC, l_udtCurShiftDetails, l_udtNewShiftDetails
                    End If
                    
' 10/16/02 RSG - HR is the only exception valid on an off day.  All others must be deleted.
'                ElseIf l_strExcCd = CHOLIDAYBANKED Then
'                    'reschedule
'                    AppendConflict r_rsConflicts, r_rsExcScheduledShifts, r_rsExceptions, _
'                        CRESCHEDEXC, l_udtCurShiftDetails, l_udtNewShiftDetails
                Else
                    If l_strExcCd = CHOLIDAYBANKED Then
                    'JimB 02/03/2003
                    'Holiday Banked is now allowed on an off day
                    'In this situation give the user the option to change, delete, or leave the HB alone.
                        AppendConflict r_rsConflicts, r_rsExcScheduledShifts, r_rsExceptions, _
                            CRESCHEDEXC, l_udtCurShiftDetails, l_udtNewShiftDetails, CHOLIDAYBANKED
                    Else
                    'delete
                        AppendConflict r_rsConflicts, r_rsExcScheduledShifts, r_rsExceptions, _
                            CDELETEEXC, l_udtCurShiftDetails, l_udtNewShiftDetails
                    End If
                End If
            End If
        Else
            If IsScheduledWorkDay(l_udtNewShiftDetails) Then    'OFF/WORK
                If (l_strExcCd = CHOLIDAYONREST Or l_strExcCd = CHOLIDAYBANKED) And l_udtNewShiftDetails.Hours = CDec(CHRHOURS) Then
                    'reclassify H when shift hours = default HR
                    'it seems like H will have to be deleted anyway to schedule HW...
                    AppendConflict r_rsConflicts, r_rsExcScheduledShifts, r_rsExceptions, _
                        CRESCHEDEXC, l_udtCurShiftDetails, l_udtNewShiftDetails, CHOLIDAY
                Else
'                    If l_strExcCd = CHOLIDAYBANKED And l_udtNewShiftDetails.Hours = CDec(CHRHOURS) Then
'                    'JimB 02/03/2003
'                    'Holiday Banked is now allowed on an off day
'                    'In this situation give the user the option to change, delete, or leave the HB alone.
'                        AppendConflict r_rsConflicts, r_rsExcScheduledShifts, r_rsExceptions, _
'                            CRESCHEDEXC, l_udtCurShiftDetails, l_udtNewShiftDetails, CHOLIDAYBANKED
'                    Else
                    'delete   'jimb3
                        l_lngEmpNum = r_rsExceptions!EMP_NUM
                        If FlexPolicyApplies(l_lngEmpNum) Then
                            If (l_strExcCd = CHOLIDAYONREST Or l_strExcCd = CHOLIDAYBANKED) And l_udtNewShiftDetails.Hours = 9 Then
                                AppendConflict2 r_rsConflicts, r_rsExcScheduledShifts, r_rsExceptions, _
                                    CRESCHEDEXC, l_udtCurShiftDetails, l_udtNewShiftDetails
                                AppendConflict3 r_rsConflicts, r_rsExcScheduledShifts, r_rsExceptions, _
                                    CRESCHEDEXC, l_udtCurShiftDetails, l_udtNewShiftDetails
                            Else
                                AppendConflict r_rsConflicts, r_rsExcScheduledShifts, r_rsExceptions, _
                                    CDELETEEXC, l_udtCurShiftDetails, l_udtNewShiftDetails
                            End If
                        Else
                            AppendConflict r_rsConflicts, r_rsExcScheduledShifts, r_rsExceptions, _
                                CDELETEEXC, l_udtCurShiftDetails, l_udtNewShiftDetails
                        End If
'                    End If
                End If
            Else                                                'OFF/OFF
                'reschedule
                AppendConflict r_rsConflicts, r_rsExcScheduledShifts, r_rsExceptions, _
                    CRESCHEDEXC, l_udtCurShiftDetails, l_udtNewShiftDetails
            End If
        End If
        r_rsExceptions.MoveNext
    Loop

    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "ProcessShift"
End Sub


Public Function UpdateExceptions( _
                    ByVal v_rsConflicts As ADODB.Recordset, _
                    ByVal v_rsExceptions As ADODB.Recordset, _
                    ByVal v_rsAvailSchedules As ADODB.Recordset, _
                    ByVal v_blnApplyUpdates As Boolean, _
                    ByRef r_varUpdates As Variant, _
                    ByRef r_rsErrors As ADODB.Recordset, _
                    ByVal v_lngUserEmpNum As Long) As Variant
' RED286
    Dim l_cdCascade As dbCascade.cdCascade
    Dim l_cdMsgDispatcher As dbMessageDispatcher.cdMessageDispatcher
    Dim l_cdWorkgroup As dbWorkGroup.cdWorkGroup
    Dim l_varEmployees As Variant
    Dim l_datStartDate As Date
    Dim l_datEndDate As Date
    Dim l_varSchedCodes As Variant
    Dim l_rsEmpWorkgroups As ADODB.Recordset
    Dim l_colEMPRVTrans As Collection
    Dim l_varEMPRVTrans As Variant
    Dim l_rsAccBankUpdates As ADODB.Recordset
    Dim l_cbsAccBankMgr As busAccrualBankMgr.cbsAccrualBankMgr

On Error GoTo eh

    m_lngUserEmpNum = v_lngUserEmpNum
    If Not RecordsExist(v_rsConflicts) Then
        GoTo xit
    Else
        v_rsConflicts.MoveFirst
    End If
    Set m_rsErrors = CreateNewBusinessErrorSet
    Set l_rsAccBankUpdates = CreateAccrualBankMgrRecordset

' Retrieve all exceptions, if necessary
    If v_rsExceptions Is Nothing Then
        DetermineGetExceptionsScope v_rsConflicts, l_varEmployees, l_datStartDate, l_datEndDate
        Set l_cdCascade = CreateObject("dbCascade.cdCascade")
        Set v_rsExceptions = _
            l_cdCascade.GetScheduledExceptions(l_varEmployees, l_datStartDate, l_datEndDate)
        Set l_cdCascade = Nothing
    End If

' Retrieve all available schedule details, if necessary
    If v_rsAvailSchedules Is Nothing Then
        l_varSchedCodes = DetermineSchedCodesforUpdateExceptions(v_rsConflicts, v_rsExceptions)
        Set l_cdCascade = CreateObject("dbCascade.cdCascade")
        Set v_rsAvailSchedules = _
            l_cdCascade.GetAvailableSchedules(l_varSchedCodes)
        Set l_cdCascade = Nothing
    End If
    
' Retrieve workgroups for EMPRV employees
    DetermineEMPRVEmployees v_rsExceptions, l_varEmployees, l_datStartDate, l_datEndDate
    If Not IsEmpty(l_varEmployees) Then
        Set l_cdWorkgroup = CreateObject("dbWorkGroup.cdWorkGroup")
        Set l_rsEmpWorkgroups = _
            l_cdWorkgroup.GetEmployeeWorkGroups( _
                    l_varEmployees, l_datStartDate, l_datEndDate, True)
        Set l_cdWorkgroup = Nothing
    End If
    
    If m_rsEmployees Is Nothing Then
        DetermineGetExceptionsScope v_rsConflicts, l_varEmployees, l_datStartDate, l_datEndDate
        GetEmployees l_varEmployees
    End If
      
    Set l_colEMPRVTrans = _
            ProcessExceptions(v_rsConflicts, v_rsExceptions, v_rsAvailSchedules, l_rsEmpWorkgroups, l_rsAccBankUpdates)
    
    Set l_cdCascade = CreateObject("dbCascade.cdCascade")
    If v_blnApplyUpdates Then
        l_cdCascade.UpdateCascadeExceptions v_rsExceptions, m_lngUserEmpNum
        r_varUpdates = Empty
    Else
        r_varUpdates = l_cdCascade.GetCascadeExceptionUpdates(v_rsExceptions, m_lngUserEmpNum)
    End If
        
    If RecordsExist(l_rsAccBankUpdates) Then
        Set l_cbsAccBankMgr = _
            CreateObject("busAccrualBankMgr.cbsAccrualBankMgr")
        l_cbsAccBankMgr.UpdateBanksForCasacade l_rsAccBankUpdates, Nothing, m_lngUserEmpNum
        Set l_cbsAccBankMgr = Nothing
    End If
        
    If l_colEMPRVTrans.Count > 0 Then
        l_varEMPRVTrans = ArrayFromCollection(l_colEMPRVTrans)
        Set l_cdMsgDispatcher = _
            CreateObject("dbMessageDispatcher.cdMessageDispatcher")
        l_cdMsgDispatcher.AddEMPRVTransactions l_varEMPRVTrans
        Set l_cdMsgDispatcher = Nothing
    End If
   
xit:
    Set l_cdCascade = Nothing
    Set l_cdWorkgroup = Nothing
    Set l_rsEmpWorkgroups = Nothing
    Set l_cbsAccBankMgr = Nothing
    Set l_colEMPRVTrans = Nothing
    Set v_rsExceptions = Nothing
    Set v_rsAvailSchedules = Nothing
    
    If Not m_rsErrors Is Nothing Then
        If m_rsErrors.RecordCount = 0 Then
            Set r_rsErrors = Nothing
        Else
            Set r_rsErrors = m_rsErrors
        End If
    Else
        Set r_rsErrors = Nothing
    End If
    Set m_rsErrors = Nothing
    GetObjectContext.SetComplete
    Exit Function

eh:
    GetObjectContext.SetAbort
    Set l_cdCascade = Nothing
    Set l_cdWorkgroup = Nothing
    Set l_rsEmpWorkgroups = Nothing
    Set l_cdMsgDispatcher = Nothing
    Set l_cbsAccBankMgr = Nothing
    Set l_colEMPRVTrans = Nothing
    Set v_rsExceptions = Nothing
    Set v_rsAvailSchedules = Nothing
    Set r_rsErrors = Nothing
    Set m_rsErrors = Nothing
    r_varUpdates = Empty
    TrsRaiseError m_strCMODULENAME, "UpdateExceptions", , , , , True
End Function


'****************************************************************
' Private Procedures
'****************************************************************

Private Function ProcessExceptions( _
                    ByRef r_rsConflicts As ADODB.Recordset, _
                    ByRef r_rsExceptions As ADODB.Recordset, _
                    ByRef r_rsAvailSchedules As ADODB.Recordset, _
                    ByRef r_rsEmpWorkgroups As ADODB.Recordset, _
                    ByRef r_rsAccBankUpdates As ADODB.Recordset) As Collection
' RED286
' Answer a collection of EMPRV transactions (variant arrays)
    Dim l_varEMPRVTranRec As Variant
    Dim l_colEMPRVTrans As Collection

On Error GoTo eh
    Set l_colEMPRVTrans = New Collection

    r_rsConflicts.MoveFirst
    Do Until r_rsConflicts.EOF
        If r_rsConflicts!CASCADE_TYPE = CRESCHEDEXC Then
'JIMB 02/07/2003
' Making changes to allow HB to be rescheduled as an H
' Adding code to update banks
            ProcessRescheduleException r_rsConflicts, r_rsExceptions, r_rsAvailSchedules, r_rsAccBankUpdates
'JIMB 02/07/2003
' If we allow exceptions other than holiday to be changed, we need to pass them
' along to EMPRV.
'The code is in ProcessRescheduleEMPRV , but is not being implemented at this time
'We probably need to consider interfacing with exception scheduling
' to reschedule and delete exce[tions
'            ProcessRescheduleEMPRV l_colEMPRVTrans, r_rsExceptions, r_rsEmpWorkgroups
        Else
            If r_rsConflicts!CASCADE_TYPE = CDELETEEXC Then
                l_varEMPRVTranRec = ProcessRemoveException( _
                    r_rsConflicts, r_rsExceptions, r_rsAvailSchedules, r_rsEmpWorkgroups, r_rsAccBankUpdates)
                If Not IsEmpty(l_varEMPRVTranRec) Then
                    l_colEMPRVTrans.Add l_varEMPRVTranRec
                End If
            Else
                ProcessAddException r_rsConflicts, r_rsExceptions
                If Not IsEmpty(l_varEMPRVTranRec) Then
                    l_colEMPRVTrans.Add l_varEMPRVTranRec
                End If
            End If
        End If
        r_rsConflicts.MoveNext
    Loop

    Set ProcessExceptions = l_colEMPRVTrans
    Set l_colEMPRVTrans = Nothing
    Exit Function

eh:
    Set l_colEMPRVTrans = Nothing
    TrsRaiseError m_strCMODULENAME, "ProcessExceptions"
End Function


Private Function DetermineEMPRVEmployees( _
            ByRef r_rsExceptions As ADODB.Recordset, _
            ByRef r_varEmployees As Variant, _
            ByRef r_datStartDate As Date, _
            ByRef r_datEndDate As Date)
'r_varEmployees, r_datStartDate, r_datEndDate are output arguments
'
'r_rsExceptions has the INTERFACE_TO_CD, as a result of join with TTR_DEPT,
'of the employees DEPT_NUM or PREV_DEPT_NUM, based on comparison with DEPT_EFF_DT
'and CHARGE_DATE.  Therefore, INTERFACE_TO_CD is the correct indicator for the
'CHARGE_DATE of the exception.
'
    Dim l_colDistinctEmpNums As cuSet

On Error GoTo eh

    r_varEmployees = Empty
    r_datStartDate = CLOWDATE
    r_datEndDate = CHIGHDATE
    Set l_colDistinctEmpNums = New cuSet
    
    r_rsExceptions.MoveFirst
    Do Until r_rsExceptions.EOF
        If IsEmployeeEMPRV(r_rsExceptions) Then
            If l_colDistinctEmpNums Is Nothing Then
                Set l_colDistinctEmpNums = New cuSet
            End If
            l_colDistinctEmpNums.Add CLng(r_rsExceptions!EMP_NUM.Value)
            r_datStartDate = MinDate(r_rsExceptions!CHARGE_DATE.Value, r_datStartDate)
            r_datEndDate = MaxDate(r_rsExceptions!CHARGE_DATE.Value, r_datEndDate)
        End If
        r_rsExceptions.MoveNext
    Loop
    If Not l_colDistinctEmpNums Is Nothing Then
        r_varEmployees = ArrayFromCollection(l_colDistinctEmpNums.Contents)
        Set l_colDistinctEmpNums = Nothing
    End If
    r_rsExceptions.MoveFirst
    
    Exit Function

eh:
    Set l_colDistinctEmpNums = Nothing
    TrsRaiseError m_strCMODULENAME, "DetermineEMPRVEmployees"
End Function


Private Function UpdateNewScheduleShift( _
                    ByVal v_intSchedStartLevel As Integer, _
                    ByRef r_rsEmpExcIntervals As ADODB.Recordset, _
                    ByRef r_rsSchedExcShifts As ADODB.Recordset, _
                    ByRef v_rsWGBaseSchedChanges As ADODB.Recordset, _
                    ByRef v_rsWGLineSchedChanges As ADODB.Recordset, _
                    ByRef r_rsWGAssignChanges As ADODB.Recordset, _
                    ByRef r_rsWGAssignOtherChanges As ADODB.Recordset, _
                    ByRef r_rsEmpLineAssignChanges As ADODB.Recordset, _
                    ByRef r_rsPersonalChanges As ADODB.Recordset, _
                    ByRef r_rsLineTradeChanges As ADODB.Recordset, _
                    ByRef r_rsEmpTradeChanges As ADODB.Recordset) _
' New schedule must be determined by calling GetAdjustedBase requesting
'  schedules from the given level, v_intSchedStartLevel.
' 10/16/02 RSG - Set NEW_SCHED_CD to off day(-2) when no match found on schedule.
    Dim l_cbsSchedule As busSchedule.cbsSchedule
    Dim l_rsEmpSchedules As ADODB.Recordset
    Dim l_rsEmpSchedExc As ADODB.Recordset
    Dim l_strFilter As String

On Error GoTo eh

    Set l_rsEmpSchedExc = GetEmpSchedExcDates(r_rsSchedExcShifts)
    
'    Where will the WGAssignmentSPA be passed from?  Where previously retrieved?
    
    Set l_cbsSchedule = CreateObject("busSchedule.cbsSchedule")
    Set l_rsEmpSchedules = l_cbsSchedule.GetPendingAdjustedBaseSchedule( _
            l_rsEmpSchedExc, Nothing, False, trBaseSchedNoDetails, v_intSchedStartLevel, _
            Nothing, v_rsWGBaseSchedChanges, v_rsWGLineSchedChanges, _
            r_rsWGAssignChanges, r_rsWGAssignOtherChanges, r_rsEmpLineAssignChanges, _
            r_rsPersonalChanges, r_rsLineTradeChanges, r_rsEmpTradeChanges)
    Set l_cbsSchedule = Nothing
    
    r_rsSchedExcShifts.MoveFirst
    Do Until r_rsSchedExcShifts.EOF
        l_strFilter = "EMP_NUM  = " & CStr(r_rsSchedExcShifts!EMP_NUM) & _
            " AND CHARGE_DATE = " & CStr(CDate(r_rsSchedExcShifts!CHARGE_DATE)) & _
            " AND SHIFT_NUM = " & CStr(r_rsSchedExcShifts!SHIFT_NUM)
        l_strFilter = MDAC25FilterFix(l_strFilter)
        l_rsEmpSchedules.Filter = l_strFilter
        If RecordsExist(l_rsEmpSchedules) Then
            If l_rsEmpSchedules!SCHED_CD = r_rsSchedExcShifts!SCHED_CD And _
                l_rsEmpSchedules!SCHED_SRC_CD = r_rsSchedExcShifts!SCHED_SRC_CD Then
                r_rsSchedExcShifts!NEW_SCHED_CD = 0
                r_rsSchedExcShifts!NEW_SCHED_SRC_CD = 0
            Else
                r_rsSchedExcShifts!NEW_SCHED_CD = l_rsEmpSchedules!SCHED_CD
                r_rsSchedExcShifts!NEW_SCHED_SRC_CD = l_rsEmpSchedules!SCHED_SRC_CD
            End If
        Else
            r_rsSchedExcShifts!NEW_SCHED_CD = -2
            r_rsSchedExcShifts!NEW_SCHED_SRC_CD = 0
        End If
        r_rsSchedExcShifts.Update
        r_rsSchedExcShifts.MoveNext
    Loop

    Set l_rsEmpSchedExc = Nothing
    Set l_rsEmpSchedules = Nothing
    Exit Function
    
eh:
    Set l_cbsSchedule = Nothing
    Set l_rsEmpSchedExc = Nothing
    Set l_rsEmpSchedules = Nothing
    TrsRaiseError m_strCMODULENAME, "UpdateNewScheduleShift"
End Function


Private Function GetEmpSchedExcDates( _
                    ByRef r_rsSchedExcShifts As ADODB.Recordset) _
                As ADODB.Recordset
'    Build a recordset with intervals for each day an exception is scheduled.
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

    r_rsSchedExcShifts.MoveFirst
    Do Until r_rsSchedExcShifts.EOF
        l_rsEmpDates.AddNew
        l_rsEmpDates!EMP_NUM = r_rsSchedExcShifts!EMP_NUM
        l_rsEmpDates!START_DATE = r_rsSchedExcShifts!CHARGE_DATE
        l_rsEmpDates!END_DATE = r_rsSchedExcShifts!CHARGE_DATE
        l_rsEmpDates.Update
        r_rsSchedExcShifts.MoveNext
    Loop
    r_rsSchedExcShifts.MoveFirst
    
    Set GetEmpSchedExcDates = l_rsEmpDates
    Set l_rsEmpDates = Nothing

    Exit Function

eh:
    Set l_rsEmpDates = Nothing
    TrsRaiseError m_strCMODULENAME, "GetEmpSchedExcDates"
End Function


Private Function IsEmployeeEMPRV( _
                    ByRef r_rsExceptions As ADODB.Recordset) As Boolean

On Error GoTo eh
    
    IsEmployeeEMPRV = (r_rsExceptions!INTERFACE_TO_CD = "E")

    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "IsEmployeeEMPRV"
End Function


Private Function ProcessRescheduleException( _
                ByRef r_rsConflicts As ADODB.Recordset, _
                ByRef r_rsExceptions As ADODB.Recordset, _
                ByRef r_rsAvailSchedules As ADODB.Recordset, _
                ByRef r_rsAccBankUpdates As ADODB.Recordset _
                ) As Collection
'RED284
    Dim l_strExcCd As String
    Dim l_lngNewSchedCd As Long
    Dim l_udtNewSchedShift As ScheduledShift
    Dim l_datChargeDate As Date
    Dim l_strDeptNum As String
    Dim l_lngEmpNum As Long
    Dim l_lngChargeDate As Date


On Error GoTo eh

    l_lngEmpNum = r_rsConflicts!EMP_NUM
    l_lngChargeDate = r_rsConflicts!CHARGE_DATE
    FindException r_rsExceptions, _
        l_lngEmpNum, l_lngChargeDate, r_rsConflicts!SHIFT_NUM, _
        r_rsConflicts.Fields("HOURS_OFFSET_NUM").Value, r_rsConflicts!CREATED_TIMESTAMP

'Holiday on Day of Rest is always 8 hours
'If the previous holiday was 12 hours, the HR is set to 8 hours.

    If Trim(r_rsConflicts!HOLIDAY_CHG_IND) = "Y" Then
        l_strExcCd = Trim(r_rsConflicts!NEW_EXCEPTION_CD)
        l_strDeptNum = r_rsExceptions!DEPT_NUM
        If IsNewExceptionCodeValid(l_strExcCd, Trim(r_rsConflicts!SCHED_WORK), l_strDeptNum) Then
'JimB Begin 01/21/2003
            If l_strExcCd = CHOLIDAYONREST Then
                    r_rsExceptions("HOURS") = CHRHOURS
            Else
        'The else part of this code is to clean up bad data
        'It will allow the user to clean up the data by changing the
        'employee's schedule.
        'After data has been cleaned up this code is not needed.
                l_datChargeDate = r_rsConflicts!CHARGE_DATE
                l_lngNewSchedCd = r_rsConflicts!NEW_SCHED_CD
                l_udtNewSchedShift = _
                    AllocateScheduledShift(r_rsAvailSchedules, l_lngNewSchedCd, l_datChargeDate)
                If Trim(r_rsConflicts!EXCEPTION_CD) = CHOLIDAYONREST Then
                    r_rsExceptions("HOURS") = l_udtNewSchedShift.Hours
                End If
                If l_udtNewSchedShift.Hours < r_rsExceptions("HOURS") Then
                    If FlexPolicyApplies(l_lngEmpNum) Then
                        If l_strExcCd = CHOLIDAYBANKED Then
                            r_rsExceptions("HOURS") = CHRHOURS
                        Else
                            r_rsExceptions("HOURS") = l_udtNewSchedShift.Hours
                        End If
                    Else
                        r_rsExceptions("HOURS") = l_udtNewSchedShift.Hours
                    End If
                End If
            End If
'JimB End 01/21/2003
            r_rsExceptions("EXCEPTION_CD") = l_strExcCd
        Else
            Err.Raise trErrbusCascadeException, m_strCMODULENAME
        End If
    End If
                
    r_rsExceptions("SCHED_CD") = r_rsConflicts!NEW_SCHED_CD
    r_rsExceptions("SCHED_SRC_CD") = r_rsConflicts!NEW_SCHED_SRC_CD
    r_rsExceptions("LAST_MOD_EMP_NUM") = m_lngUserEmpNum
    r_rsExceptions.Update

    AppendExceptionRescheduledBRV r_rsExceptions, r_rsConflicts, r_rsAvailSchedules
'Jimb 02/07/2003
'Added code to update accrual banks
'Needed when converting HB to H or HW codes.
'It does open this up for all exception codes
    AppendAccrualBankUpdateReschedule r_rsAccBankUpdates, r_rsExceptions
        
    Exit Function

eh:

    Dim l_colDescArgs As Collection
    Dim l_lngErrNum As Long
    
    l_lngErrNum = Err.Number
    Select Case l_lngErrNum
        Case trErrbusCascadeException
            Set l_colDescArgs = New Collection
            l_colDescArgs.Add l_strExcCd
            l_colDescArgs.Add CStr(l_lngEmpNum)
            l_colDescArgs.Add CStr(l_lngChargeDate)
            Err.Source = FormatBRV(TrType_TC, l_lngEmpNum, l_lngChargeDate, l_lngChargeDate, "", "", 0, l_strExcCd, "")
            TrsRaiseError m_strCMODULENAME, "ProcessRescheduleException", l_lngErrNum, , l_colDescArgs
        Case Else
            TrsRaiseError m_strCMODULENAME, "ProcessRescheduleException"
    End Select
End Function


Private Sub ClearFilter(ByRef r_rsData As ADODB.Recordset)

On Error GoTo eh
    If Not r_rsData Is Nothing Then
        r_rsData.Filter = 0
    End If
    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "ClearFilter"
End Sub


Private Function ProcessRemoveException( _
                ByRef r_rsConflicts As ADODB.Recordset, _
                ByRef r_rsExceptions As ADODB.Recordset, _
                ByRef r_rsAvailSchedules As ADODB.Recordset, _
                ByRef r_rsEmpWorkgroups As ADODB.Recordset, _
                ByRef r_rsAccBankUpdates As ADODB.Recordset) As Variant
' Answer an EMPRV transaction array when EMPRV employee, otherwise Empty
' RED284
    Dim l_blnEMPRV As Boolean
    Dim l_intLocCd As Integer

On Error GoTo eh

    FindException r_rsExceptions, _
        r_rsConflicts!EMP_NUM, r_rsConflicts!CHARGE_DATE, r_rsConflicts!SHIFT_NUM, _
        r_rsConflicts!HOURS_OFFSET_NUM, r_rsConflicts!CREATED_TIMESTAMP
    
    l_blnEMPRV = IsEmployeeEMPRV(r_rsExceptions)
    If l_blnEMPRV Then
        l_intLocCd = EmployeeWorkgroupLocation(r_rsEmpWorkgroups, _
            r_rsExceptions!WG_NUM, r_rsExceptions!EMP_NUM, r_rsExceptions!CHARGE_DATE)
        ProcessRemoveException = BuildEMPRVTransaction(r_rsExceptions, l_intLocCd)
    Else
        ProcessRemoveException = Empty
    End If

    AppendExceptionRemovedBRV r_rsExceptions, r_rsConflicts, r_rsAvailSchedules
    AppendAccrualBankUpdate r_rsAccBankUpdates, r_rsExceptions

    r_rsExceptions.Delete
    
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "ProcessRemoveException"
End Function


Private Function BuildEMPRVTransaction( _
                    ByRef r_rsExceptions As ADODB.Recordset, _
                    ByVal v_intLocationCd As Integer, _
                    Optional v_strTransType As String = "D", _
                    Optional v_blnRescheduleRequest As Boolean = False) As Variant
    Dim l_varEMPRVTran(0 To 6) As Variant

On Error GoTo eh

    l_varEMPRVTran(etEMPNUM) = r_rsExceptions!EMP_NUM
    l_varEMPRVTran(etDEPTNUM) = r_rsExceptions!DEPT_NUM
    l_varEMPRVTran(etCHGDATE) = r_rsExceptions!CHARGE_DATE
    l_varEMPRVTran(etHOURS) = r_rsExceptions!Hours
    l_varEMPRVTran(etLOCCD) = v_intLocationCd
    l_varEMPRVTran(etTRANCD) = v_strTransType
    If Trim(r_rsExceptions!OVER_EXC_CD) <> "" Then
        l_varEMPRVTran(etEXCCD) = Trim(r_rsExceptions!OVER_EXC_CD)
    Else
        l_varEMPRVTran(etEXCCD) = Trim(r_rsExceptions!EXCEPTION_CD)
    End If
    If v_blnRescheduleRequest Then
        If Trim(r_rsExceptions("OVER_EXC_CD").OriginalValue) <> "" Then
            l_varEMPRVTran(etEXCCD) = Trim(r_rsExceptions("OVER_EXC_CD").OriginalValue)
        Else
            l_varEMPRVTran(etEXCCD) = Trim(r_rsExceptions("EXCEPTION_CD").OriginalValue)
        End If
    End If

    BuildEMPRVTransaction = l_varEMPRVTran
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "BuildEMPRVTransaction"
End Function


Private Function EmployeeWorkgroupLocation( _
                ByRef r_rsEmpWorkgroups As ADODB.Recordset, _
                ByVal v_lngWgNum As Long, _
                ByVal v_lngEmpNum As Long, _
                ByVal v_datChargeDate As Date) As Long
'If WG_NUM is > 0 on exception, look up LOC_CD for it and answer it.
' Otherwise determine the primary workgroup on the charge date and look up LOC_CD.
' cdWorkgroup caches workgroup data in SPM, as requested.  May need to do some
' fast initialization of the SPM cache earlier in ProcessExceptions to speed things up here.
    Dim l_cdWorkgroup As dbWorkGroup.cdWorkGroup

On Error GoTo eh

    If v_lngWgNum > 0 Then
        Set l_cdWorkgroup = CreateObject("dbWorkGroup.cdWorkGroup")
        EmployeeWorkgroupLocation = _
            CLng(l_cdWorkgroup.LocationCode(v_lngWgNum))
        Set l_cdWorkgroup = Nothing
    Else
        PositionOnWorkgroup r_rsEmpWorkgroups, v_lngEmpNum, v_datChargeDate
        If RecordsExist(r_rsEmpWorkgroups) Then
            EmployeeWorkgroupLocation = r_rsEmpWorkgroups!LOC_CD
        Else
            EmployeeWorkgroupLocation = 0
        End If
    End If

    Exit Function

eh:
    Set l_cdWorkgroup = Nothing
    TrsRaiseError m_strCMODULENAME, "EmployeeWorkgroupLocation"
End Function


Private Function IsAHolidayorHolidayWorked( _
                    ByVal v_strExcCd As String, ByVal v_strGroupInd As String) _
                As Boolean
' Answer True when this exception should be reclassified as HR (Holiday on Rest).
' Do not reclassify HB since it should be rescheduled when a work day or an off day;
'  i.e. do not delete an HB when becomes an off day since employee may want to bank
'  instead of paid.
' Do not reclassify HD since it should be deleted when day scheduled becomes an off day.
' RED282
' 10/16/2002 RSG - Do not reclassify HD since it should be deleted when
' a double scheduled becomes an off day.
    Dim l_strExcCd As String

On Error GoTo eh
    l_strExcCd = Trim(v_strExcCd)
    IsAHolidayorHolidayWorked = _
        Trim(v_strGroupInd = "H") And l_strExcCd <> "HB" And v_strExcCd <> "HD" And v_strExcCd <> "HDS"
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "IsAHolidayorHolidayWorked"
End Function


Private Function CreateAccrualBankMgrRecordset() As ADODB.Recordset
    Dim l_rsRequest As ADODB.Recordset

On Error GoTo eh

    Set l_rsRequest = New ADODB.Recordset
    With l_rsRequest
      .Fields.Append "EMP_NUM", adInteger, , adFldUpdatable
      .Fields.Append "CHARGE_DATE", adDate, , adFldUpdatable
      .Fields.Append "EXCEPTION_CD", adChar, 3, adFldUpdatable
      .Fields.Append "OVER_EXC_CD", adChar, 3, adFldUpdatable
      .Fields.Append "HOURS", adDecimal, , adFldUpdatable
      .Fields.Item("HOURS").Precision = 3
      .Fields.Item("HOURS").NumericScale = 1
      .Fields.Append "BANK_CODE", adChar, 2, adFldUpdatable
      .Fields.Append "FMLA_IND", adChar, 1, adFldUpdatable
      .Fields.Append "REC_STATUS", adChar, 1, adFldUpdatable

      .CursorLocation = adUseClientBatch
      .CursorType = adOpenStatic
      .LockType = adLockBatchOptimistic
      .Open
    End With
    Set CreateAccrualBankMgrRecordset = l_rsRequest
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "CreateAccrualBankMgrRecordset"
End Function


Private Function CreateConflictRecordset() As ADODB.Recordset
' Create an empty recordset to store exceptions that need to be
' rescheduled or deleted.
    Dim l_rsConflicts As ADODB.Recordset
    
On Error GoTo eh

    Set l_rsConflicts = New ADODB.Recordset
    With l_rsConflicts
        .Fields.Append "CASCADE_TYPE", adChar, 1, adFldUpdatable       'R(reschedule) or D(delete)
        .Fields.Append "SHIFT_TYPE", adChar, 1, adFldUpdatable          'F,P
        .Fields.Append "EMP_NUM", adInteger, , adFldUpdatable
        .Fields.Append "EMP_NAME", adChar, 30, adFldUpdatable
        .Fields.Append "CHARGE_DATE", adDate, , adFldUpdatable
        .Fields.Append "SHIFT_NUM", adInteger, , adFldUpdatable
        .Fields.Append "EXCEPTION_CD", adChar, 3, adFldUpdatable
        .Fields.Append "NEW_EXCEPTION_CD", adChar, 3, adFldUpdatable
        .Fields.Append "EXCEPTION_DESC", adChar, 30, adFldUpdatable
        .Fields.Append "HOLIDAY_CHG_IND", adChar, 1, adFldUpdatable
        .Fields.Append "START_TIME", adChar, 5, adFldUpdatable
        .Fields.Append "END_TIME", adChar, 5, adFldUpdatable
        .Fields.Append "SCHED_CD", adInteger, , adFldUpdatable
        .Fields.Append "NEW_SCHED_CD", adInteger, , adFldUpdatable
        .Fields.Append "NEW_SCHED_SRC_CD", adInteger, , adFldUpdatable
        .Fields.Append "SCHED_WORK", adChar, 1, adFldUpdatable
        .Fields.Append "WG_NUM", adInteger, , adFldUpdatable
        .Fields.Append "JOB_NUM", adChar, 6, adFldUpdatable
        .Fields.Append "RELIEF_COMP", adChar, 3, adFldUpdatable
        .Fields.Append "CASE_NUM", adInteger, , adFldUpdatable
        .Fields.Append "COMP_CD", adChar, 3, adFldUpdatable
        .Fields.Append "OVER_EXC_CD", adChar, 3, adFldUpdatable
        .Fields.Append "FMLA_IND", adChar, 1, adFldUpdatable
        .Fields.Append "SHIFT_COMP", adChar, 3, adFldUpdatable
        .Fields.Append "HOURS_OFFSET_NUM", adDecimal, , adFldUpdatable
            .Fields.Item("HOURS_OFFSET_NUM").Precision = 3
            .Fields.Item("HOURS_OFFSET_NUM").NumericScale = 1
        .Fields.Append "CREATED_TIMESTAMP", adChar, 26, adFldUpdatable
        .Fields.Append "LAST_MOD_TIMESTAMP", adChar, 26, adFldUpdatable
        .Fields.Append "HOURS", adDecimal, , adFldUpdatable
        .Fields.Item("HOURS").Precision = 3
        .Fields.Item("HOURS").NumericScale = 1
        .Open
    End With
    
    Set CreateConflictRecordset = l_rsConflicts
    Set l_rsConflicts = Nothing
    Exit Function

eh:
    Set l_rsConflicts = Nothing
    TrsRaiseError m_strCMODULENAME, "CreateConflictRecordset"
End Function


Private Sub AppendConflict( _
                ByRef r_rsConflicts As ADODB.Recordset, _
                ByRef r_rsSchedExceptions As ADODB.Recordset, _
                ByRef r_rsExceptionDetails As ADODB.Recordset, _
                ByVal v_strCascadeType As String, _
                ByRef v_udtCurShift As ScheduledShift, _
                ByRef v_udtNewShift As ScheduledShift, _
                Optional ByVal v_strNewExcCd As String)
    Dim l_strExcDesc As String
    Dim l_strShiftType As String
    
On Error GoTo eh

    l_strShiftType = FullOrPartialExcType(r_rsExceptionDetails, v_udtCurShift.Hours)

    With r_rsConflicts
        .AddNew
        !CASCADE_TYPE = v_strCascadeType
        !SHIFT_TYPE = l_strShiftType
        !EMP_NUM = r_rsExceptionDetails!EMP_NUM
        !EMP_NAME = r_rsExceptionDetails!EMP_NAME
        !CHARGE_DATE = r_rsExceptionDetails!CHARGE_DATE
        !SHIFT_NUM = r_rsExceptionDetails!SHIFT_NUM
        
        !HOLIDAY_CHG_IND = "N"
        !EXCEPTION_CD = r_rsExceptionDetails!EXCEPTION_CD
        If v_strNewExcCd <> "" Then
            !NEW_EXCEPTION_CD = v_strNewExcCd
            If v_strNewExcCd = CHOLIDAY Or v_strNewExcCd = CHOLIDAYONREST Or v_strNewExcCd = CHOLIDAYBANKED Then
                !HOLIDAY_CHG_IND = "Y"
            End If
        Else
            If Trim(!EXCEPTION_CD) = CHOLIDAY Then
                !NEW_EXCEPTION_CD = CHOLIDAY
            End If
        End If
        
        
        Select Case v_strNewExcCd
            Case CHOLIDAY
                If r_rsExceptionDetails!EXCEPTION_CD = CHOLIDAYONREST Then
                    l_strExcDesc = "HR->H: Holiday"
                Else
                    l_strExcDesc = Trim(r_rsExceptionDetails!EXCEPTION_CD) _
                        & "->H: Holiday"
                End If
            Case CHOLIDAYONREST
                l_strExcDesc = Trim(r_rsExceptionDetails!EXCEPTION_CD) _
                    & "->HR: Holiday on Rest"
            Case CHOLIDAYBANKED
                l_strExcDesc = Trim(r_rsExceptionDetails!EXCEPTION_CD) _
                    & "->HB: Holiday Banked"
            
            Case Else
                If Trim(r_rsExceptionDetails!OVER_EXC_CD) <> "" Then
                    l_strExcDesc = Trim(r_rsExceptionDetails!OVER_EXC_DESC)
                Else
                    l_strExcDesc = Trim(r_rsExceptionDetails!EXCEPTION_DESC)
                End If
        End Select
                
        !EXCEPTION_DESC = l_strExcDesc

' 6/25/02 RSG - Added this to support HR edit during update.
        If IsScheduledWorkDay(v_udtNewShift) Then
            !SCHED_WORK = "Y"
        Else
            !SCHED_WORK = "N"
        End If

        If IsScheduledWorkDay(v_udtCurShift) Then
            !START_TIME = ShiftTime(v_udtCurShift.StartTime, r_rsExceptionDetails!HOURS_OFFSET_NUM, v_udtCurShift.LunchStart, v_udtCurShift.LunchEnd)
            !END_TIME = ShiftTime(!START_TIME, r_rsExceptionDetails!Hours, v_udtCurShift.LunchStart, v_udtCurShift.LunchEnd, True)
        Else
            !START_TIME = CHRSTARTTIME
            !END_TIME = CHRENDTIME
        End If
        
        !SCHED_CD = r_rsExceptionDetails!SCHED_CD
        !NEW_SCHED_CD = r_rsSchedExceptions!NEW_SCHED_CD
        !NEW_SCHED_SRC_CD = r_rsSchedExceptions!NEW_SCHED_SRC_CD
        
        !JOB_NUM = r_rsExceptionDetails!JOB_NUM
        !WG_NUM = r_rsExceptionDetails!WG_NUM
        !RELIEF_COMP = r_rsExceptionDetails!RELIEF_COMP
        !CASE_NUM = r_rsExceptionDetails!CASE_NUM
        !COMP_CD = r_rsExceptionDetails!COMP_CD
        !OVER_EXC_CD = r_rsExceptionDetails!OVER_EXC_CD
        !FMLA_IND = r_rsExceptionDetails!FMLA_IND
        !SHIFT_COMP = r_rsExceptionDetails!SHIFT_COMP
        !HOURS_OFFSET_NUM = r_rsExceptionDetails!HOURS_OFFSET_NUM
        !CREATED_TIMESTAMP = r_rsExceptionDetails!CREATED_TIMESTAMP
        .Update
    End With

    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "AppendConflict"
End Sub


Private Function AppendExceptionRescheduledBRV( _
            ByRef r_rsExceptions As ADODB.Recordset, _
            ByRef r_rsConflicts As ADODB.Recordset, _
            ByRef r_rsAvailSchedules As ADODB.Recordset) As Boolean
            
    Dim l_bolError As Boolean
    Dim l_lngEmpNum As Long
    Dim l_intShiftNum As Long
    Dim l_datChargeDate As Date
    Dim l_lngOldSchedCd As Long
    Dim l_lngNewSchedCd As Long
    Dim l_strExcCd As String
    Dim l_udtOldSchedShift As ScheduledShift
    Dim l_udtNewSchedShift As ScheduledShift
    Dim l_strErrorDesc As String
    Dim l_strExcStart As String
    Dim l_strExcStop As String
    Dim l_strNewExcStart As String
    Dim l_strNewExcStop As String
    
On Error GoTo eh
   
    l_lngEmpNum = r_rsConflicts!EMP_NUM
    l_datChargeDate = r_rsConflicts!CHARGE_DATE
    l_intShiftNum = r_rsConflicts!SHIFT_NUM
    l_lngOldSchedCd = r_rsConflicts!SCHED_CD
    l_lngNewSchedCd = r_rsConflicts!NEW_SCHED_CD
    l_strExcCd = Trim(r_rsConflicts!EXCEPTION_CD)
    
' Skip information messages when schedule does not change, i.e. source only change.
    If l_lngOldSchedCd = l_lngNewSchedCd Then
        AppendExceptionRescheduledBRV = True
        Exit Function
    End If
    
    l_udtOldSchedShift = _
        AllocateScheduledShift(r_rsAvailSchedules, l_lngOldSchedCd, l_datChargeDate)
    l_udtNewSchedShift = _
        AllocateScheduledShift(r_rsAvailSchedules, l_lngNewSchedCd, l_datChargeDate)
        
    If IsScheduledWorkDay(l_udtOldSchedShift) Then
        l_strExcStart = ShiftTime(l_udtOldSchedShift.StartTime, r_rsExceptions!HOURS_OFFSET_NUM, l_udtOldSchedShift.LunchStart, l_udtOldSchedShift.LunchEnd)
        l_strExcStop = ShiftTime(l_strExcStart, r_rsExceptions!Hours, l_udtOldSchedShift.LunchStart, l_udtOldSchedShift.LunchEnd, True)
    Else
        l_strExcStart = ShiftTime(CHRSTARTTIME, r_rsExceptions!HOURS_OFFSET_NUM)
        l_strExcStop = ShiftTime(l_strExcStart, r_rsExceptions!Hours, , , True)
    End If

    If IsScheduledWorkDay(l_udtNewSchedShift) Then
        l_strNewExcStart = ShiftTime(l_udtNewSchedShift.StartTime, r_rsExceptions!HOURS_OFFSET_NUM, l_udtNewSchedShift.LunchStart, l_udtNewSchedShift.LunchEnd)
        l_strNewExcStop = ShiftTime(l_strNewExcStart, r_rsExceptions!Hours, l_udtNewSchedShift.LunchStart, l_udtNewSchedShift.LunchEnd, True)
    Else
        l_strNewExcStart = ShiftTime(CHRSTARTTIME, r_rsExceptions!HOURS_OFFSET_NUM)
        l_strNewExcStop = ShiftTime(l_strNewExcStart, r_rsExceptions!Hours, , , True)
    End If
        
    l_strErrorDesc = _
        "Rescheduled " & Trim(r_rsConflicts!EXCEPTION_CD) & " for " & _
        l_lngEmpNum & "[" & Trim(r_rsConflicts!EMP_NAME) & "] " & _
        " on " & Format(l_datChargeDate, "mm/dd/yyyy") & " from schedule " & _
         r_rsConflicts!SCHED_CD & "(" & l_strExcStart & "-" & l_strExcStop & ") to " & _
         r_rsConflicts!NEW_SCHED_CD & "(" & l_strNewExcStart & "-" & l_strNewExcStop & ")"

    l_bolError = AppendBusinessError( _
            m_rsErrors, _
            LongAppErr(trInfCascadeResult), _
            mstrCBUSINESSINF, _
            l_strErrorDesc, _
            m_strCMODULENAME, _
            TrType_CA, , , "EMP_NUM;CHARGE_DATE;SHIFT_NUM;EXCEPTION_CD", _
            CStr(l_lngEmpNum), _
            Format(l_datChargeDate, "mm/dd/yyyy"), _
            CStr(l_intShiftNum), _
            l_strExcCd)

    AppendExceptionRescheduledBRV = l_bolError
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "AppendExceptionRescheduledBRV"
End Function


Private Function AppendExceptionRemovedBRV( _
            ByRef r_rsExceptions As ADODB.Recordset, _
            ByRef r_rsConflicts As ADODB.Recordset, _
            ByRef r_rsAvailSchedules As ADODB.Recordset) As Boolean
            
    Dim l_bolError As Boolean
    Dim l_lngEmpNum As Long
    Dim l_intShiftNum As Long
    Dim l_datChargeDate As Date
    Dim l_lngSchedCd As Long
    Dim l_strExcCd As String
    Dim l_udtOldSchedShift As ScheduledShift
    Dim l_strErrorDesc As String
    Dim l_strExcStart As String
    Dim l_strExcStop As String
    Dim l_strOverExcCd As String
    Dim l_strExcCdMsg As String
    Dim l_udtNewSchedShift As ScheduledShift
    Dim l_strNewExcStart As String
    Dim l_strNewExcStop As String
    Dim l_lngNewSchedCd As Long
    
On Error GoTo eh
   
    l_lngEmpNum = r_rsConflicts!EMP_NUM
    l_datChargeDate = r_rsConflicts!CHARGE_DATE
    l_intShiftNum = r_rsConflicts!SHIFT_NUM
    l_lngSchedCd = r_rsConflicts!SCHED_CD
    l_strExcCd = Trim(r_rsConflicts!EXCEPTION_CD)
    l_strOverExcCd = Trim(r_rsConflicts!OVER_EXC_CD)
    l_lngNewSchedCd = r_rsConflicts!NEW_SCHED_CD
    
    If l_strOverExcCd <> "" Then
        l_strExcCdMsg = l_strExcCd & "/" & l_strOverExcCd
    Else
        l_strExcCdMsg = l_strExcCd
    End If
    
    l_udtOldSchedShift = _
        AllocateScheduledShift(r_rsAvailSchedules, l_lngSchedCd, l_datChargeDate)
    l_udtNewSchedShift = _
        AllocateScheduledShift(r_rsAvailSchedules, l_lngNewSchedCd, l_datChargeDate)
    
    If IsScheduledWorkDay(l_udtOldSchedShift) Then
        l_strExcStart = ShiftTime(l_udtOldSchedShift.StartTime, r_rsExceptions!HOURS_OFFSET_NUM, l_udtOldSchedShift.LunchStart, l_udtOldSchedShift.LunchEnd)
        l_strExcStop = ShiftTime(l_strExcStart, r_rsExceptions!Hours, l_udtOldSchedShift.LunchStart, l_udtOldSchedShift.LunchEnd, True)
    Else
        l_strExcStart = ShiftTime(CHRSTARTTIME, r_rsExceptions!HOURS_OFFSET_NUM)
        l_strExcStop = ShiftTime(l_strExcStart, r_rsExceptions!Hours, , , True)
    End If
    
    If IsScheduledWorkDay(l_udtNewSchedShift) Then
        l_strNewExcStart = ShiftTime(l_udtNewSchedShift.StartTime, r_rsExceptions!HOURS_OFFSET_NUM, l_udtNewSchedShift.LunchStart, l_udtNewSchedShift.LunchEnd)
        l_strNewExcStop = ShiftTime(l_strNewExcStart, r_rsExceptions!Hours, l_udtNewSchedShift.LunchStart, l_udtNewSchedShift.LunchEnd, True)
    Else
        l_strNewExcStart = "Off"
        l_strNewExcStop = "day"
    End If
    
    Select Case l_strExcCd
        Case "VCO"
'            l_strErrorDesc = "VCO exception must be rescheduled as VCO in the same week as the other VCO exceptions.  "
            l_strErrorDesc = "Vacation scheduled on " & Format(l_datChargeDate, "mm/dd/yyyy") _
                    & " is part of a partial week of vacation that has been carried over from the previous year. " _
                    & " This partial week of vacation cannot be further split and must be scheduled in a block."
        Case "VEF"
'            l_strErrorDesc = "VEF exception must be rescheduled as VEF in the same week as the other VEF exceptions.  "
            l_strErrorDesc = "Vacation scheduled on " & Format(l_datChargeDate, "mm/dd/yyyy") _
                    & " is part of a third week of vacation that has been split for vacation in lieu of sick leave. " _
                    & " The remaining hours of vacation in this week cannot be moved except to be taken as vacation in lieu of sick leave up to this date."
        Case Else
            l_strErrorDesc = ""
    End Select
    
    If l_strErrorDesc <> "" Then
        l_bolError = AppendBusinessError( _
            m_rsErrors, _
            LongAppErr(trInfCascadeResult), _
            mstrCBUSINESSWAR, _
            l_strErrorDesc, _
            m_strCMODULENAME, _
            TrType_CA, , , "EMP_NUM;CHARGE_DATE;SHIFT_NUM;EXCEPTION_CD", _
            CStr(l_lngEmpNum), _
            Format(l_datChargeDate, "mm/dd/yyyy"), _
            CStr(l_intShiftNum), _
            l_strExcCd)
    End If
    l_strErrorDesc = _
        "Deleted " & l_strExcCdMsg & " for " & _
        l_lngEmpNum & "[" & Trim(r_rsConflicts!EMP_NAME) & "] " & _
        " on " & Format(l_datChargeDate, "mm/dd/yyyy") & " from schedule " & _
         r_rsConflicts!SCHED_CD & "(" & l_strExcStart & "-" & l_strExcStop & ") to " & _
         r_rsConflicts!NEW_SCHED_CD & "(" & l_strNewExcStart & "-" & l_strNewExcStop & ")"
    
    l_bolError = AppendBusinessError( _
            m_rsErrors, _
            LongAppErr(trInfCascadeResult), _
            mstrCBUSINESSINF, _
            l_strErrorDesc, _
            m_strCMODULENAME, _
            TrType_CA, , , "EMP_NUM;CHARGE_DATE;SHIFT_NUM;EXCEPTION_CD", _
            CStr(l_lngEmpNum), _
            Format(l_datChargeDate, "mm/dd/yyyy"), _
            CStr(l_intShiftNum), _
            l_strExcCd)

    AppendExceptionRemovedBRV = l_bolError
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "AppendExceptionRemovedBRV"
End Function


Private Sub AppendAccrualBankUpdate( _
            ByRef r_rsAccBankUpdates As ADODB.Recordset, _
            ByRef r_rsExceptions As ADODB.Recordset)
    
On Error GoTo eh
   
    With r_rsAccBankUpdates
        .AddNew
        !EMP_NUM = r_rsExceptions!EMP_NUM
        !CHARGE_DATE = r_rsExceptions!CHARGE_DATE
        !EXCEPTION_CD = r_rsExceptions!EXCEPTION_CD
        !OVER_EXC_CD = r_rsExceptions!OVER_EXC_CD
        !Hours = r_rsExceptions!Hours
        !BANK_CODE = ExceptionBankCode(r_rsExceptions!EXCEPTION_CD)
        !FMLA_IND = r_rsExceptions!FMLA_IND
        !REC_STATUS = "D"
        .Update
    End With
        
    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "AppendAccrualBankUpdate"
End Sub


Private Function ExceptionBankCode( _
            ByVal v_strExcCd As String) As String
    Dim l_cdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh

    Set l_cdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    ExceptionBankCode = _
        l_cdCodeTables.ExceptionBankCode(Trim(v_strExcCd))
    Set l_cdCodeTables = Nothing
    Exit Function

eh:
    Set l_cdCodeTables = Nothing
    TrsRaiseError m_strCMODULENAME, "ExceptionBankCode"
End Function


Private Function IsNewExceptionCodeValid( _
                    ByVal v_strExcCd As String, _
                    ByVal v_strSchedWork As String, _
                    ByVal v_strDeptNum As String _
                 ) As Boolean
    Dim l_cdCodeTables As dbCodeTables.cdCodeTables
    Dim l_blnIsValid As Boolean
    
On Error GoTo eh

    If v_strExcCd = CHOLIDAYONREST And v_strSchedWork = "Y" Then
        IsNewExceptionCodeValid = False
        Exit Function
    End If

    Set l_cdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    l_blnIsValid = _
        l_cdCodeTables.ExceptionCodeIsValid(v_strExcCd)
    If l_blnIsValid Then
        If Left(v_strDeptNum, 1) >= "5" Then
            l_blnIsValid = l_cdCodeTables.ExceptionCodeIsValidFor(v_strExcCd, "E")
        Else
            l_blnIsValid = l_cdCodeTables.ExceptionCodeIsValidFor(v_strExcCd, "C")
        End If
    End If
    IsNewExceptionCodeValid = l_blnIsValid
    Set l_cdCodeTables = Nothing
    Exit Function

eh:
    Set l_cdCodeTables = Nothing
    TrsRaiseError m_strCMODULENAME, "IsNewExceptionCodeValid"
End Function


Private Function ValidateGetSchedExcDtRequest( _
                    ByRef v_rsEmpExcInterval As ADODB.Recordset, _
                    ByRef r_varEmployees As Variant, _
                    ByRef r_datStartDate As Date, _
                    ByRef r_datEndDate As Date) As Boolean
' v_varEmployees must be either;
' a) an integer, long, or string that is upto 6 numerals long; i.e. valid empNums
' b) an array of valid empNums
'
' v_datStartDate and v_datEndDate must be a date,but not 12/30/1899 or 12/31/1899.
' (vbNull or vbEmpty)
'
' Answer r_varEmployees as a variant array of at least 1 long when valid.
'  Otherwise, answer an undimensioned or empty array.
'
    Dim l_lngValResult As Long
    Dim l_objSet As cuSet
    
    r_datStartDate = CHIGHDATE
    r_datEndDate = CLOWDATE

    If Not RecordsExist(v_rsEmpExcInterval) Then
        ValidateGetSchedExcDtRequest = False
        Exit Function
    Else
        v_rsEmpExcInterval.MoveFirst
    End If

    Set l_objSet = New cuSet
    Do Until v_rsEmpExcInterval.EOF
        l_lngValResult = ValidateEmployeeNumber(v_rsEmpExcInterval!EMP_NUM)
        If l_lngValResult = -1 Then
            ValidateGetSchedExcDtRequest = False
            Exit Function
        Else
            l_objSet.Add l_lngValResult
        End If
        r_datStartDate = MinDate(v_rsEmpExcInterval!START_DATE.Value, r_datStartDate)
        r_datEndDate = MaxDate(v_rsEmpExcInterval!END_DATE.Value, r_datEndDate)
        v_rsEmpExcInterval.MoveNext
    Loop

    If l_objSet.Contents.Count < 1 Then
        ValidateGetSchedExcDtRequest = False
    Else
        r_varEmployees = ArrayFromCollection(l_objSet.Contents)
        ValidateGetSchedExcDtRequest = True
    End If
    v_rsEmpExcInterval.MoveFirst
    Set l_objSet = Nothing
    Exit Function

eh:
    Set l_objSet = Nothing
    TrsRaiseError m_strCMODULENAME, "ValidateGetSchedExcDtRequest"
End Function


Private Function ValidateEmployeeNumber(v_varEmpNum As Variant) As Long
'Answer employee number as long.
'Answer -1 when v_varEmpNum is not a valid employee number.
    Dim l_lngEmpNum As Long

On Error GoTo eh
    Select Case VarType(v_varEmpNum)
        Case vbInteger, vbLong
            l_lngEmpNum = CLng(v_varEmpNum)
            If l_lngEmpNum > 0 And l_lngEmpNum <= 999999 Then
                ValidateEmployeeNumber = l_lngEmpNum
            Else
                m_strErrorDescription = "Employee number: " & CStr(l_lngEmpNum) & " not in valid range 1 to 999999"
                ValidateEmployeeNumber = -1
            End If
        Case vbString
            If Len(v_varEmpNum) <= 6 And IsAllNumerals(v_varEmpNum) Then
                ValidateEmployeeNumber = CLng(v_varEmpNum)
            Else
                m_strErrorDescription = "Employee number: " & v_varEmpNum & " is not a valid string"
                ValidateEmployeeNumber = -1
            End If
        Case Else
            m_strErrorDescription = "Employee number: " & CStr(v_varEmpNum) & " is not valid"
            ValidateEmployeeNumber = -1
    End Select
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "ValidateEmployeeNumber"
End Function

Private Function DetermineGetAdjustedBaseScope( _
                    ByRef r_rsExcDates As ADODB.Recordset, _
                    ByVal v_varEmployees As Variant, _
                    ByVal v_datStartDate As Date, _
                    ByVal v_datEndDate As Date, _
                    ByRef r_varGabEmployees As Variant, _
                    ByRef r_datGabStartDate As Date, _
                    ByRef r_datGabEndDate As Date)
' The scope of the GetAdjustedBase request may be reduced since the results are only
' used to determine the schedule source and TRIS 4.0 will store this on the exception.
' Scan the exceptions for SCHED_SRC_CD
    Dim l_datStartDate As Date
    Dim l_datEndDate As Date
    Dim l_colDistinctEmpNums As cuSet
    Const l_datLowDate As Date = #1/1/1900#
    Const l_datHighDate As Date = #12/31/2079#

On Error GoTo eh

'guard for empty recordset
    If Not RecordsExist(r_rsExcDates) Then
        r_varGabEmployees = v_varEmployees
        r_datGabStartDate = v_datStartDate
        r_datGabEndDate = v_datEndDate
        Exit Function
    Else
        r_rsExcDates.MoveFirst
    End If

' seed the dates with high and low dates
' there should be at least 1 valid record
    l_datStartDate = l_datHighDate
    l_datEndDate = l_datLowDate
    Set l_colDistinctEmpNums = New cuSet
    Do Until r_rsExcDates.EOF
    'if SCHED_SRC_CD > 0 assume a valid value
        If r_rsExcDates!SCHED_SRC_CD <= 0 Then
            l_colDistinctEmpNums.Add CLng(r_rsExcDates!EMP_NUM.Value)
            l_datStartDate = MinDate(r_rsExcDates!CHARGE_DATE.Value, l_datStartDate)
            l_datEndDate = MaxDate(r_rsExcDates!CHARGE_DATE.Value, l_datEndDate)
        End If
        r_rsExcDates.MoveNext
    Loop

    r_datGabStartDate = l_datStartDate
    r_datGabEndDate = l_datEndDate
    r_varGabEmployees = ArrayFromCollection(l_colDistinctEmpNums.Contents)
    r_rsExcDates.MoveFirst
    Set l_colDistinctEmpNums = Nothing
    Exit Function

eh:
    Set l_colDistinctEmpNums = Nothing
    TrsRaiseError m_strCMODULENAME, "DetermineGetAdjustedBaseScope"
End Function

                    
Private Sub DetermineGetEmployeeScope( _
                    ByRef r_rsConflicts As ADODB.Recordset, _
                    ByRef r_varEmployees As Variant)
    Dim l_colDistinctEmpNums As cuSet
    
On Error GoTo eh
'guard for empty recordset
    If Not RecordsExist(r_rsConflicts) Then
        r_varEmployees = Empty
        Exit Sub
    Else
        r_rsConflicts.MoveFirst
    End If

    Set l_colDistinctEmpNums = New cuSet
    Do Until r_rsConflicts.EOF
        l_colDistinctEmpNums.Add CLng(r_rsConflicts!EMP_NUM.Value)
        r_rsConflicts.MoveNext
    Loop
    
    r_varEmployees = ArrayFromCollection(l_colDistinctEmpNums.Contents)
    r_rsConflicts.MoveFirst
    Set l_colDistinctEmpNums = Nothing
    Exit Sub

eh:
    Set l_colDistinctEmpNums = Nothing
    TrsRaiseError m_strCMODULENAME, "DetermineGetEmployeeScope"
End Sub
                    
                  
Private Sub DetermineGetExceptionsScope( _
                    ByRef r_rsExcDates As ADODB.Recordset, _
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
    If Not RecordsExist(r_rsExcDates) Then
        r_varEmployees = Empty
        r_datStartDate = 0
        r_datEndDate = 0
        Exit Sub
    Else
        r_rsExcDates.MoveFirst
    End If

' seed the dates with high and low dates
' there should be at least 1 valid record
    l_datStartDate = l_datHighDate
    l_datEndDate = l_datLowDate
    Set l_colDistinctEmpNums = New cuSet
    Do Until r_rsExcDates.EOF
        l_colDistinctEmpNums.Add CLng(r_rsExcDates!EMP_NUM.Value)
        l_datStartDate = MinDate(r_rsExcDates!CHARGE_DATE.Value, l_datStartDate)
        l_datEndDate = MaxDate(r_rsExcDates!CHARGE_DATE.Value, l_datEndDate)
        r_rsExcDates.MoveNext
    Loop

    r_datStartDate = l_datStartDate
    r_datEndDate = l_datEndDate
    r_varEmployees = ArrayFromCollection(l_colDistinctEmpNums.Contents)
    r_rsExcDates.MoveFirst
    Set l_colDistinctEmpNums = Nothing
    Exit Sub

eh:
    Set l_colDistinctEmpNums = Nothing
    TrsRaiseError m_strCMODULENAME, "DetermineGetExceptionsScope"
End Sub
                                     

Private Function DetermineGetSchedulesScope( _
                    ByRef r_rsExcDates As ADODB.Recordset) As Variant
' Answer a variant array of schedule codes.
    Dim l_colDistinctSchedCodes As cuSet

On Error GoTo eh

    Set l_colDistinctSchedCodes = New cuSet
    
    If RecordsExist(r_rsExcDates) Then
        r_rsExcDates.MoveFirst
        Do Until r_rsExcDates.EOF
            AddNonZero l_colDistinctSchedCodes, CLng(r_rsExcDates!SCHED_CD.Value)
            AddNonZero l_colDistinctSchedCodes, CLng(r_rsExcDates!NEW_SCHED_CD.Value)
            r_rsExcDates.MoveNext
        Loop
        r_rsExcDates.MoveFirst
    End If
        
    DetermineGetSchedulesScope = ArrayFromCollection(l_colDistinctSchedCodes.Contents)
    Set l_colDistinctSchedCodes = Nothing
    Exit Function

eh:
    Set l_colDistinctSchedCodes = Nothing
    TrsRaiseError m_strCMODULENAME, "DetermineGetSchedulesScope"
End Function


Private Sub AddNonZero(ByRef r_objSet As cuSet, ByVal v_lngSchedCd As Long)
'Only add non-zero schedule codes
    Dim l_lngSchedCd As Long

    If v_lngSchedCd <> 0 Then
        r_objSet.Add v_lngSchedCd
    End If
    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "AddNonZero"
End Sub


'Temporary
'Copy of DetermineGetSchedulesScope
'If the NEW_SCHED_CD was SCHED_CD and SCHED_CD was OLD_SCHED_CD,
'it could use the original function.
'JimB 04/21/2002
' I do not understand comment above.
' The r_rsConflicts contains the NEW_SCHED_CD as the new and SCHED_CD as the old schedule
' The code will be left in for now, but changed name to avoid some confusion.
Private Function DetermineSchedCodesforUpdateExceptions( _
                    ByRef r_rsConflicts As ADODB.Recordset, _
                    ByRef r_rsExceptions As ADODB.Recordset _
                ) As Variant
' Answer a variant array of schedule codes.
'
' 10/10/01 RSG
' Should this be scope be reduced by start and end dates by employee?
' How many extraneous exceptions are returned?  Could determine a range
' based on the scheduled exception dates.
'
    Dim l_colDistinctSchedCodes As cuSet

On Error GoTo eh

    Set l_colDistinctSchedCodes = New cuSet
    
    If RecordsExist(r_rsConflicts) Then
        r_rsConflicts.MoveFirst
        Do Until r_rsConflicts.EOF
            l_colDistinctSchedCodes.Add CLng(r_rsConflicts!NEW_SCHED_CD.Value)
            r_rsConflicts.MoveNext
        Loop
        r_rsConflicts.MoveFirst
    End If
    
    If RecordsExist(r_rsExceptions) Then
        r_rsExceptions.MoveFirst
        Do Until r_rsExceptions.EOF
            l_colDistinctSchedCodes.Add CLng(r_rsExceptions!SCHED_CD.Value)
            r_rsExceptions.MoveNext
        Loop
        r_rsExceptions.MoveFirst
    End If
        
    DetermineSchedCodesforUpdateExceptions = ArrayFromCollection(l_colDistinctSchedCodes.Contents)
    Set l_colDistinctSchedCodes = Nothing
    Exit Function

eh:
    Set l_colDistinctSchedCodes = Nothing
    TrsRaiseError m_strCMODULENAME, "DetermineSchedCodesforUpdateExceptions"
End Function

                    
Private Function MinDate(ByVal v_datDate1 As Date, ByVal v_datDate2 As Date) As Date
'compare v_datDate1 and v_datDate2, returning the minimum value.
On Error GoTo eh

    If v_datDate1 < v_datDate2 Then
        MinDate = v_datDate1
    Else
        MinDate = v_datDate2
    End If
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "MinDate"
End Function


Private Function MaxDate(ByVal v_datDate1 As Date, ByVal v_datDate2 As Date) As Date
'compare v_datDate1 and v_datDate2, returning the maximum value.
On Error GoTo eh
    
    If v_datDate1 > v_datDate2 Then
        MaxDate = v_datDate1
    Else
        MaxDate = v_datDate2
    End If
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "MaxDate"
End Function


Private Function CreateSchedExcDtRecordset() As ADODB.Recordset
' Create an empty recordset to store exceptions that will be returned.
    Dim l_rsExcSchedDates As ADODB.Recordset
    
On Error GoTo eh

    Set l_rsExcSchedDates = New ADODB.Recordset
    With l_rsExcSchedDates
        .Fields.Append "EMP_NUM", adInteger, , adFldUpdatable
        .Fields.Append "CHARGE_DATE", adDate, , adFldUpdatable
        .Fields.Append "SHIFT_NUM", adInteger, , adFldUpdatable
        .Fields.Append "SCHED_CD", adInteger, , adFldUpdatable
        .Fields.Append "SCHED_SRC_CD", adInteger, , adFldUpdatable
        .Fields.Append "NEW_SCHED_CD", adInteger, , adFldUpdatable
        .Fields.Append "NEW_SCHED_SRC_CD", adInteger, , adFldUpdatable
        .Open
    End With
    
    Set CreateSchedExcDtRecordset = l_rsExcSchedDates
    Set l_rsExcSchedDates = Nothing
    Exit Function

eh:
    Set l_rsExcSchedDates = Nothing
    TrsRaiseError m_strCMODULENAME, "CreateSchedExcDtRecordset"
End Function


Private Function DetermineSchedExcSource( _
                    ByRef r_rsExcDates As ADODB.Recordset, _
                    ByRef r_rsEmpSched As ADODB.Recordset) As Integer
' r_rsExcDates is positioned on the current exception record
    Dim l_strFilter As String
    Dim l_intSchedSource As Integer
    
On Error GoTo eh

    If r_rsExcDates!SCHED_SRC_CD > 0 Then
        l_intSchedSource = r_rsExcDates!SCHED_SRC_CD
    Else
        If r_rsEmpSched Is Nothing Then
            'raise application error - gab not called when SCHED_SRC_CD <= 0
            Err.Raise vbObjectError + 10020
        End If
        l_strFilter = _
            "EMP_NUM = " & r_rsExcDates!EMP_NUM & _
            " AND CHARGE_DATE = #" & Format(r_rsExcDates!CHARGE_DATE, "mm/dd/yyyy") & "#" & _
            " AND SCHED_CD = r_rsExcDates!SCHED_CD"
        l_strFilter = MDAC25FilterFix(l_strFilter)
        r_rsEmpSched.Filter = l_strFilter
        If RecordsExist(r_rsEmpSched) Then
            If r_rsEmpSched!SCHED_SRC_CD > 0 Then
                l_intSchedSource = r_rsEmpSched!SCHED_SRC_CD
            Else
                'source not valid on schedule data
                l_intSchedSource = 0
            End If
        Else
            'exception does not match schedule
            l_intSchedSource = 0
        End If
    End If
    
    DetermineSchedExcSource = l_intSchedSource
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "DetermineSchedExcSource"
End Function


Private Sub FindException( _
                ByRef r_rsExceptions As ADODB.Recordset, _
                ByVal v_lngEmpNum As Long, _
                ByVal v_datChargeDate As Date, _
                ByVal v_intShiftNum As Integer, _
                ByVal v_intOffset As Variant, _
                ByVal v_strCreatedTS As String)
    Dim l_strFilter As String
    
On Error GoTo eh

    r_rsExceptions.MoveFirst
    Do Until r_rsExceptions.EOF
        If r_rsExceptions!EMP_NUM = v_lngEmpNum And _
            CDate(r_rsExceptions!CHARGE_DATE) = v_datChargeDate And _
            r_rsExceptions!SHIFT_NUM = v_intShiftNum And _
            r_rsExceptions!HOURS_OFFSET_NUM = v_intOffset And _
            Trim(r_rsExceptions!CREATED_TIMESTAMP) = Trim(v_strCreatedTS) Then
            Exit Do
        End If
        r_rsExceptions.MoveNext
    Loop
   
    If r_rsExceptions.EOF Then
'     exception for original shift/schedule code not found
'     raise a commit conflict exception
        Err.Raise vbObjectError + 10002
    End If
       
    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "FindException"
End Sub


Private Sub FilterOnShift( _
                ByRef r_rsExceptions As ADODB.Recordset, _
                ByVal v_lngEmpNum As Long, _
                ByVal v_datChargeDate As Date, _
                ByVal v_intShiftNum As Integer)
    Dim l_strFilter As String
    
On Error GoTo eh
    l_strFilter = _
        "EMP_NUM = " & v_lngEmpNum & _
        " AND CHARGE_DATE = " & "#" & Format(v_datChargeDate, "MM/DD/YYYY") & "#" & _
        " AND SHIFT_NUM = " & v_intShiftNum
    l_strFilter = MDAC25FilterFix(l_strFilter)
    r_rsExceptions.Filter = l_strFilter
          
    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "FilterOnShift"
End Sub


Private Sub FilterOnException( _
                ByRef r_rsExceptions As ADODB.Recordset, _
                ByVal v_lngEmpNum As Long, _
                ByVal v_datChargeDate As Date, _
                ByVal v_intShiftNum As Integer, _
                ByVal v_intOffset As Integer, _
                ByVal v_strCreatedTS As String)
    Dim l_strFilter As String
    
On Error GoTo eh
    l_strFilter = _
        "EMP_NUM = " & v_lngEmpNum & _
        " AND CHARGE_DATE = " & "#" & Format(v_datChargeDate, "MM/DD/YYYY") & "#" & _
        " AND SHIFT_NUM = " & v_intShiftNum & _
        " AND HOURS_OFFSET_NUM = " & v_intOffset & _
        " AND CREATED_TIMESTAMP = " & v_strCreatedTS
    l_strFilter = MDAC25FilterFix(l_strFilter)
    r_rsExceptions.Filter = l_strFilter
          
    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "FilterOnException"
End Sub


Private Sub PositionOnWorkgroup( _
                ByRef r_rsEmpWorkgroups As ADODB.Recordset, _
                ByVal v_lngEmpNum As Long, _
                ByVal v_datChargeDate As Date)
    Dim l_strFilter As String
    
On Error GoTo eh
    l_strFilter = "EMP_NUM = " & v_lngEmpNum & " AND " & _
        "EFF_START_DATE <= " & _
        "#" & Format(v_datChargeDate, "MM/DD/YYYY") & "#" & _
        " AND EFF_END_DATE >= " & _
        "#" & Format(v_datChargeDate, "MM/DD/YYYY") & "#"
    l_strFilter = MDAC25FilterFix(l_strFilter)
    r_rsEmpWorkgroups.Filter = l_strFilter
    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "PositionOnWorkgroup"
End Sub


Private Function IsExceptionDateInScope( _
                ByRef r_rsEmpExcInterval As ADODB.Recordset, _
                ByVal v_lngEmpNum As Long, _
                ByVal v_datChargeDate As Date) As Boolean
    Dim l_strFilter As String
    
On Error GoTo eh
    l_strFilter = "EMP_NUM = " & v_lngEmpNum & " AND " & _
        "START_DATE <= " & _
        "#" & Format(v_datChargeDate, "MM/DD/YYYY") & "#" & _
        " AND END_DATE >= " & _
        "#" & Format(v_datChargeDate, "MM/DD/YYYY") & "#"
    l_strFilter = MDAC25FilterFix(l_strFilter)
    r_rsEmpExcInterval.Filter = l_strFilter
    IsExceptionDateInScope = RecordsExist(r_rsEmpExcInterval)
    r_rsEmpExcInterval.Filter = 0
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "IsExceptionDateInScope"
End Function


Private Function FullOrPartialExcType( _
                    ByRef r_rsExceptions As ADODB.Recordset, _
                    ByVal v_varShiftHours As Variant) As String
' Answer F or P
' v_varShiftHours should be a decimal.
' Now that exceptions are not split for lunch or midnight, just compare
' the hours to the shift hours.
    
On Error GoTo eh
    
    If CDec(r_rsExceptions.Fields("HOURS")) < v_varShiftHours Then
        FullOrPartialExcType = CPARTTYPE
    Else
        FullOrPartialExcType = CFULLTYPE
    End If

    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "FullOrPartialExcType"
End Function

Private Function DoesHRExist( _
    ByVal r_rsConflicts As ADODB.Recordset, _
    ByVal r_rsExceptions As ADODB.Recordset _
    ) As Boolean
    
    Dim l_datChargeDate As Date
    Dim l_lngEmpNum As Long
    Dim l_strFilter As String
    
    
 On Error GoTo eh
 
    l_lngEmpNum = r_rsExceptions!EMP_NUM
    l_datChargeDate = r_rsExceptions!CHARGE_DATE
    l_strFilter = "EMP_NUM = " & str(l_lngEmpNum) & _
            " AND CHARGE_DATE = " & l_datChargeDate & _
            " AND NEW_EXCEPTION_CD = 'HR'"
    l_strFilter = MDAC25FilterFix(l_strFilter)
    r_rsConflicts.Filter = l_strFilter
    
    DoesHRExist = RecordsExist(r_rsConflicts)
    
    r_rsConflicts.Filter = ""
 
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "DoesHRExist"
End Function

Private Sub AppendAccrualBankUpdateReschedule( _
            ByRef r_rsAccBankUpdates As ADODB.Recordset, _
            ByRef r_rsExceptions As ADODB.Recordset)
    
Dim l_strExcCd As String
Dim l_strOrigExcCd As String

On Error GoTo eh
   
    l_strExcCd = Trim(r_rsExceptions!EXCEPTION_CD)
    l_strOrigExcCd = Trim(r_rsExceptions("EXCEPTION_CD").OriginalValue)
   
    If l_strExcCd = l_strOrigExcCd Then
        Exit Sub
    End If
   
    With r_rsAccBankUpdates
        .AddNew
        !EMP_NUM = r_rsExceptions!EMP_NUM
        !CHARGE_DATE = r_rsExceptions!CHARGE_DATE
        !EXCEPTION_CD = l_strOrigExcCd
        !Hours = r_rsExceptions("Hours").OriginalValue
        !BANK_CODE = ExceptionBankCode(l_strOrigExcCd)
        !FMLA_IND = r_rsExceptions("FMLA_IND").OriginalValue
        !REC_STATUS = "D"
        .Update
    End With
    
    With r_rsAccBankUpdates
        .AddNew
        !EMP_NUM = r_rsExceptions!EMP_NUM
        !CHARGE_DATE = r_rsExceptions!CHARGE_DATE
        !EXCEPTION_CD = l_strExcCd
        !Hours = r_rsExceptions!Hours
        !BANK_CODE = ExceptionBankCode(l_strExcCd)
        !FMLA_IND = r_rsExceptions!FMLA_IND
        !REC_STATUS = "I"
        .Update
    End With
        
    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "AppendAccrualBankUpdateReschedule"
End Sub

Private Sub ProcessRescheduleEMPRV( _
    ByRef r_colEMPRVTrans As Collection, _
    ByRef r_rsExceptions As ADODB.Recordset, _
    ByRef r_rsEmpWorkgroups As ADODB.Recordset)
    
'JIMB 02/07/2003
' If we allow exceptions other than holiday to be changed, we need to pass them
' along to EMPRV.
'This code is not being implemented at this time
'We probably need to consider interfacing with exception scheduling
' to reschedule and delete exce[tions
    
Dim l_strExcCd As String
Dim l_strOrigExcCd As String
Dim l_strOverExcCd As String
Dim l_strOrigOverExcCd As String
Dim l_blnEMPRV As Boolean
Dim l_varEMPRVTrans As Variant
Dim l_intLocCd As Integer

On Error GoTo eh
   
    l_blnEMPRV = IsEmployeeEMPRV(r_rsExceptions)
    If Not l_blnEMPRV Then
        Exit Sub
    End If
    
    l_strExcCd = Trim(r_rsExceptions!EXCEPTION_CD)
    l_strOrigExcCd = Trim(r_rsExceptions("EXCEPTION_CD").OriginalValue)
    l_strOverExcCd = Trim(r_rsExceptions!OVER_EXC_CD)
    l_strOrigOverExcCd = Trim(r_rsExceptions("OVER_EXC_CD").OriginalValue)
   
    If l_strExcCd = l_strOrigExcCd And _
        l_strOverExcCd = l_strOrigOverExcCd Then
        Exit Sub
    End If
    
    l_intLocCd = EmployeeWorkgroupLocation(r_rsEmpWorkgroups, _
        r_rsExceptions!WG_NUM, r_rsExceptions!EMP_NUM, r_rsExceptions!CHARGE_DATE)
    l_varEMPRVTrans = BuildEMPRVTransaction(r_rsExceptions, l_intLocCd, "I")
    r_colEMPRVTrans.Add l_varEMPRVTrans
    l_varEMPRVTrans = BuildEMPRVTransaction(r_rsExceptions, l_intLocCd, "D", True)
    r_colEMPRVTrans.Add l_varEMPRVTrans

    
    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "ProcessRescheduleEMPRV"
    
End Sub

Private Function FlexPolicyApplies(ByVal v_lngEmpNum As Long) As Boolean

    m_rsEmployees.Filter = "EMP_NUM = " & v_lngEmpNum
    If m_rsEmployees.BOF And m_rsEmployees.EOF Then
        FlexPolicyApplies = False
    Else
        If m_rsEmployees!FLEX_IND > 0 Then
            FlexPolicyApplies = True
        Else
            FlexPolicyApplies = False
        End If
    
    End If
    

End Function

Private Function DeterimineFlexPolicyInd(ByVal l_intLocationCode As Integer, _
                                        ByVal l_strDepartmentCode As String, _
                                        ByVal l_strExecContractInd As String, _
                                        ByVal l_strOTExclusion As String) As Long
    
    On Error GoTo eh
    
        
    If Trim(l_strExecContractInd) <> "E" Then
        DeterimineFlexPolicyInd = 0
        Exit Function
    End If
    
    If DoesRuleApply(601, trDepartmentNumber, Trim(l_strDepartmentCode)) And l_strOTExclusion = "0000" Then
        DeterimineFlexPolicyInd = 0
        Exit Function
    End If
    
    If DoesRuleApply(602, trDepartmentNumber, Trim(l_strDepartmentCode)) And l_strOTExclusion = "0000" Then
        DeterimineFlexPolicyInd = 0
        Exit Function
    End If
    
    If DoesRuleApply(600, trLocationCode, Trim(str(l_intLocationCode))) Then
        DeterimineFlexPolicyInd = 0
        Exit Function
    End If

    If DoesRuleApply(600, trDepartmentNumber, Trim(l_strDepartmentCode)) Then
        DeterimineFlexPolicyInd = 0
        Exit Function
    End If

    DeterimineFlexPolicyInd = 1
    Exit Function

eh:

    TrsRaiseError m_strCMODULENAME _
        , "DeterimineFlexPolicyInd"


 End Function
 
 Private Sub GetEmployees(l_varEmployees)
 
 Dim l_objDBEmployee As dbEmployee.cdEmployee
 Dim l_rsEmps As ADODB.Recordset
 Dim l_lngEmpNum As Long
 Dim l_lngFlexInd As Long
 Dim l_intLocationCode As Integer
 Dim l_strDepartmentCode As String
 Dim l_strExecContractInd As String
 Dim l_strOTExclusion As String
 
 On Error GoTo eh
 
    Set l_objDBEmployee = CreateObject("dbEmployee.cdEmployee")
    Set l_rsEmps = l_objDBEmployee.GetEmployee(l_varEmployees)
    Set l_objDBEmployee = Nothing
    
    Set m_rsEmployees = CreateFlexRS()
    If l_rsEmps.BOF And l_rsEmps.EOF Then
        Exit Sub
    End If

    l_rsEmps.MoveFirst
    While Not l_rsEmps.EOF
        l_lngEmpNum = l_rsEmps!EMP_NUM
        l_intLocationCode = l_rsEmps!LOC_CD
        l_strDepartmentCode = l_rsEmps!DEPT_NUM
        l_strExecContractInd = l_rsEmps!EXEC_CONTRACT_IND
        l_strOTExclusion = Trim(l_rsEmps!AR_OT_EXCLUSN)
        l_lngFlexInd = DeterimineFlexPolicyInd(l_intLocationCode, l_strDepartmentCode, l_strExecContractInd, l_strOTExclusion)
        m_rsEmployees.AddNew
        m_rsEmployees!EMP_NUM = l_lngEmpNum
        m_rsEmployees!FLEX_IND = l_lngFlexInd
        l_rsEmps.MoveNext
    Wend

    m_rsEmployees.Update

    Exit Sub

eh:

    TrsRaiseError m_strCMODULENAME _
        , "GetEmployees"

 
 End Sub
 
 Private Function CreateFlexRS() As ADODB.Recordset
' Create an empty recordset to store exceptions that will be returned.
    Dim l_rsFlex As ADODB.Recordset
    
On Error GoTo eh

    Set l_rsFlex = New ADODB.Recordset
    With l_rsFlex
        .Fields.Append "EMP_NUM", adInteger, , adFldUpdatable
        .Fields.Append "FLEX_IND", adInteger, , adFldUpdatable
        .Open
    End With
    
    Set CreateFlexRS = l_rsFlex
    Set l_rsFlex = Nothing
    Exit Function

eh:
    Set l_rsFlex = Nothing
    TrsRaiseError m_strCMODULENAME, "CreateFlexRS"
End Function


Private Sub AppendConflict2( _
                ByRef r_rsConflicts As ADODB.Recordset, _
                ByRef r_rsSchedExceptions As ADODB.Recordset, _
                ByRef r_rsExceptionDetails As ADODB.Recordset, _
                ByVal v_strCascadeType As String, _
                ByRef v_udtCurShift As ScheduledShift, _
                ByRef v_udtNewShift As ScheduledShift, _
                Optional ByVal v_strNewExcCd As String)
    Dim l_strExcDesc As String
    Dim l_strShiftType As String
    
On Error GoTo eh

    l_strShiftType = CPARTTYPE

    With r_rsConflicts
        .AddNew
        !CASCADE_TYPE = v_strCascadeType
        !SHIFT_TYPE = l_strShiftType
        !EMP_NUM = r_rsExceptionDetails!EMP_NUM
        !EMP_NAME = r_rsExceptionDetails!EMP_NAME
        !CHARGE_DATE = r_rsExceptionDetails!CHARGE_DATE
        !SHIFT_NUM = r_rsExceptionDetails!SHIFT_NUM
        
        !HOLIDAY_CHG_IND = "N"
        !EXCEPTION_CD = r_rsExceptionDetails!EXCEPTION_CD
        If v_strNewExcCd <> "" Then
            !NEW_EXCEPTION_CD = v_strNewExcCd
            If v_strNewExcCd = CHOLIDAY Or v_strNewExcCd = CHOLIDAYONREST Then
                !HOLIDAY_CHG_IND = "Y"
            End If
        Else
            If Trim(!EXCEPTION_CD) = CHOLIDAY Then
                !NEW_EXCEPTION_CD = CHOLIDAY
            End If
        End If
        
        
        Select Case v_strNewExcCd
            Case CHOLIDAY
                If r_rsExceptionDetails!EXCEPTION_CD = CHOLIDAYONREST Then
                    l_strExcDesc = "HR->H: Holiday"
                Else
                    l_strExcDesc = Trim(r_rsExceptionDetails!EXCEPTION_CD) _
                        & "->H: Holiday"
                End If
            Case CHOLIDAYONREST
                l_strExcDesc = Trim(r_rsExceptionDetails!EXCEPTION_CD) _
                    & "->HR: Holiday on Rest"
            Case Else
                If Trim(r_rsExceptionDetails!OVER_EXC_CD) <> "" Then
                    l_strExcDesc = Trim(r_rsExceptionDetails!OVER_EXC_DESC)
                Else
                    l_strExcDesc = Trim(r_rsExceptionDetails!EXCEPTION_DESC)
                End If
        End Select
                
        !EXCEPTION_DESC = l_strExcDesc

' 6/25/02 RSG - Added this to support HR edit during update.
        If IsScheduledWorkDay(v_udtNewShift) Then
            !SCHED_WORK = "Y"
        Else
            !SCHED_WORK = "N"
        End If

        If IsScheduledWorkDay(v_udtCurShift) Then
            !START_TIME = ShiftTime(v_udtCurShift.StartTime, r_rsExceptionDetails!HOURS_OFFSET_NUM, v_udtCurShift.LunchStart, v_udtCurShift.LunchEnd)
            !END_TIME = ShiftTime(!START_TIME, 8, v_udtCurShift.LunchStart, v_udtCurShift.LunchEnd, True)
        Else
            !START_TIME = CHRSTARTTIME
            !END_TIME = CHRENDTIME
        End If
        
        !SCHED_CD = r_rsExceptionDetails!SCHED_CD
        !NEW_SCHED_CD = r_rsSchedExceptions!NEW_SCHED_CD
        !NEW_SCHED_SRC_CD = r_rsSchedExceptions!NEW_SCHED_SRC_CD
        
        !JOB_NUM = r_rsExceptionDetails!JOB_NUM
        !WG_NUM = r_rsExceptionDetails!WG_NUM
        !RELIEF_COMP = r_rsExceptionDetails!RELIEF_COMP
        !CASE_NUM = r_rsExceptionDetails!CASE_NUM
        !COMP_CD = r_rsExceptionDetails!COMP_CD
        !OVER_EXC_CD = r_rsExceptionDetails!OVER_EXC_CD
        !FMLA_IND = r_rsExceptionDetails!FMLA_IND
        !SHIFT_COMP = r_rsExceptionDetails!SHIFT_COMP
        !HOURS_OFFSET_NUM = r_rsExceptionDetails!HOURS_OFFSET_NUM
        !CREATED_TIMESTAMP = r_rsExceptionDetails!CREATED_TIMESTAMP
        .Update
    End With

    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "AppendConflict2"
End Sub

Private Sub AppendConflict3( _
                ByRef r_rsConflicts As ADODB.Recordset, _
                ByRef r_rsSchedExceptions As ADODB.Recordset, _
                ByRef r_rsExceptionDetails As ADODB.Recordset, _
                ByVal v_strCascadeType As String, _
                ByRef v_udtCurShift As ScheduledShift, _
                ByRef v_udtNewShift As ScheduledShift, _
                Optional ByVal v_strNewExcCd As String)
    Dim l_strExcDesc As String
    Dim l_strShiftType As String
    Dim l_strCreatedTimeStamp As String
    Dim intPos As Integer
    Dim strPos As String
    
On Error GoTo eh

    l_strShiftType = CPARTTYPE

    With r_rsConflicts
        .AddNew
        !CASCADE_TYPE = v_strCascadeType
        !SHIFT_TYPE = l_strShiftType
        !EMP_NUM = r_rsExceptionDetails!EMP_NUM
        !EMP_NAME = r_rsExceptionDetails!EMP_NAME
        !CHARGE_DATE = r_rsExceptionDetails!CHARGE_DATE
        !SHIFT_NUM = r_rsExceptionDetails!SHIFT_NUM
        
        !HOLIDAY_CHG_IND = "N"
        !EXCEPTION_CD = "FLX"
        If v_strNewExcCd <> "" Then
            !NEW_EXCEPTION_CD = v_strNewExcCd
            If v_strNewExcCd = CHOLIDAY Or v_strNewExcCd = CHOLIDAYONREST Then
                !HOLIDAY_CHG_IND = "Y"
            End If
        Else
            If Trim(!EXCEPTION_CD) = CHOLIDAY Then
                !NEW_EXCEPTION_CD = CHOLIDAY
            End If
        End If
        
        
        Select Case v_strNewExcCd
            Case CHOLIDAY
                If r_rsExceptionDetails!EXCEPTION_CD = CHOLIDAYONREST Then
                    l_strExcDesc = "HR->H: Holiday"
                Else
                    l_strExcDesc = Trim(r_rsExceptionDetails!EXCEPTION_CD) _
                        & "->H: Holiday"
                End If
            Case CHOLIDAYONREST
                l_strExcDesc = Trim(r_rsExceptionDetails!EXCEPTION_CD) _
                    & "->HR: Holiday on Rest"
            Case CHOLIDAYBANKED
                l_strExcDesc = Trim(r_rsExceptionDetails!EXCEPTION_CD) _
                    & "->HB: Holiday Banked"
            Case "FLX"
                l_strExcDesc = Trim(r_rsExceptionDetails!EXCEPTION_CD) _
                    & "->FLX: Flex Time"
            Case Else
                If Trim(r_rsExceptionDetails!OVER_EXC_CD) <> "" Then
                    l_strExcDesc = Trim(r_rsExceptionDetails!OVER_EXC_DESC)
                Else
                    l_strExcDesc = Trim(r_rsExceptionDetails!EXCEPTION_DESC)
                End If
        End Select
                
        !EXCEPTION_DESC = l_strExcDesc

' 6/25/02 RSG - Added this to support HR edit during update.
        If IsScheduledWorkDay(v_udtNewShift) Then
            !SCHED_WORK = "Y"
        Else
            !SCHED_WORK = "N"
        End If

        If IsScheduledWorkDay(v_udtCurShift) Then
            !START_TIME = ShiftTime(v_udtCurShift.StartTime, 8, v_udtCurShift.LunchStart, v_udtCurShift.LunchEnd)
            !END_TIME = ShiftTime(!START_TIME, 1, v_udtCurShift.LunchStart, v_udtCurShift.LunchEnd, True)
        Else
            !START_TIME = CHRSTARTTIME
            !END_TIME = CHRENDTIME
        End If
        
        !SCHED_CD = r_rsExceptionDetails!SCHED_CD
        !NEW_SCHED_CD = r_rsSchedExceptions!NEW_SCHED_CD
        !NEW_SCHED_SRC_CD = r_rsSchedExceptions!NEW_SCHED_SRC_CD
        
        !JOB_NUM = r_rsExceptionDetails!JOB_NUM
        !WG_NUM = r_rsExceptionDetails!WG_NUM
        !RELIEF_COMP = r_rsExceptionDetails!RELIEF_COMP
        !CASE_NUM = r_rsExceptionDetails!CASE_NUM
        !COMP_CD = r_rsExceptionDetails!COMP_CD
        !OVER_EXC_CD = r_rsExceptionDetails!OVER_EXC_CD
        !FMLA_IND = r_rsExceptionDetails!FMLA_IND
        !SHIFT_COMP = r_rsExceptionDetails!SHIFT_COMP
        !HOURS_OFFSET_NUM = 8
        !Hours = 1
' The Client is expecting a uniquie created timestamp
' The FLX time doesn't need a created timestamp because it is an insert
' If the created timestamp was cleared it will affect the client
        l_strCreatedTimeStamp = r_rsExceptionDetails!CREATED_TIMESTAMP
        intPos = r_rsExceptionDetails.AbsolutePosition
        If intPos > 0 And intPos < 10 Then
            strPos = "00" + Trim(str(intPos))
        Else
            If intPos > 9 And intPos < 99 Then
                strPos = "0" + Trim(str(intPos))
            Else
                strPos = Trim(str(intPos))
            End If
        End If
        l_strCreatedTimeStamp = Left(l_strCreatedTimeStamp, 23) + strPos
        If l_strCreatedTimeStamp = r_rsExceptionDetails!CREATED_TIMESTAMP Then
            intPos = Int(Right(l_strCreatedTimeStamp, 1))
            If intPos < 9 Then
                intPos = intPos - 1
                l_strCreatedTimeStamp = Left(l_strCreatedTimeStamp, 25) + Trim(str(intPos))
            Else
                l_strCreatedTimeStamp = Left(l_strCreatedTimeStamp, 25) + "0"
            End If
        End If
        !CREATED_TIMESTAMP = l_strCreatedTimeStamp
        .Update
    End With

    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "AppendConflict3"
End Sub

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

Private Function ProcessAddException( _
                ByRef r_rsConflicts As ADODB.Recordset, _
                ByRef r_rsExceptions As ADODB.Recordset _
                ) As Collection


    Dim l_strExcCd As String
    Dim l_datChargeDate As Date
    Dim l_strDeptNum As String
    Dim l_lngEmpNum As Long

On Error GoTo eh

        l_strExcCd = r_rsConflicts!EXCEPTION_CD
        l_lngEmpNum = r_rsConflicts!EMP_NUM
        l_datChargeDate = r_rsConflicts!CHARGE_DATE
        
        r_rsExceptions.AddNew
        r_rsExceptions!EMP_NUM = r_rsConflicts!EMP_NUM
        r_rsExceptions!CHARGE_DATE = r_rsConflicts!CHARGE_DATE
        r_rsExceptions!SHIFT_NUM = r_rsConflicts!SHIFT_NUM
        r_rsExceptions!HOURS_OFFSET_NUM = r_rsConflicts!HOURS_OFFSET_NUM
        r_rsExceptions!Hours = r_rsConflicts!Hours
        r_rsExceptions!EMP_NUM = r_rsConflicts!EMP_NUM
        r_rsExceptions!EXCEPTION_CD = r_rsConflicts!EXCEPTION_CD
        r_rsExceptions!OVER_EXC_CD = r_rsConflicts!OVER_EXC_CD
        r_rsExceptions!FMLA_IND = r_rsConflicts!FMLA_IND
        r_rsExceptions!CASE_NUM = r_rsConflicts!CASE_NUM
        r_rsExceptions!WG_NUM = r_rsConflicts!WG_NUM
        r_rsExceptions!JOB_NUM = r_rsConflicts!JOB_NUM
        r_rsExceptions!RELIEF_COMP = r_rsConflicts!RELIEF_COMP
        r_rsExceptions!COMP_CD = r_rsConflicts!COMP_CD
        r_rsExceptions!SHIFT_COMP = r_rsConflicts!SHIFT_COMP
        r_rsExceptions!SCHED_CD = r_rsConflicts!NEW_SCHED_CD
        r_rsExceptions!SCHED_SRC_CD = r_rsConflicts!NEW_SCHED_SRC_CD
        r_rsExceptions!CREATED_EMP_NUM = m_lngUserEmpNum
        r_rsExceptions!LAST_MOD_EMP_NUM = m_lngUserEmpNum
        r_rsExceptions!CREATED_TIMESTAMP = r_rsConflicts!CREATED_TIMESTAMP
        r_rsExceptions!LAST_MOD_TIMESTAMP = r_rsConflicts!LAST_MOD_TIMESTAMP
        r_rsExceptions!EMP_NAME = r_rsConflicts!EMP_NAME
'                r_rsExceptions!B_DEPT_EFF_DATE = r_rsConflicts!B_DEPT_EFF_DATE
        r_rsExceptions!DEPT_NUM = ""
        r_rsExceptions!INTERFACE_TO_CD = ""
        r_rsExceptions!EXCEPTION_TYPE = ""
        r_rsExceptions!EXCEPTION_DESC = ""
        r_rsExceptions!GROUP_IND = ""
        r_rsExceptions!OVER_EXC_DESC = ""
        r_rsExceptions.Update

        Exit Function

eh:

    Dim l_colDescArgs As Collection
    Dim l_lngErrNum As Long
    
    l_lngErrNum = Err.Number
    Select Case l_lngErrNum
        Case trErrbusCascadeException
            Set l_colDescArgs = New Collection
            l_colDescArgs.Add l_strExcCd
            l_colDescArgs.Add CStr(l_lngEmpNum)
            l_colDescArgs.Add CStr(l_datChargeDate)
            Err.Source = FormatBRV(TrType_TC, l_lngEmpNum, l_datChargeDate, l_datChargeDate, "", "", 0, l_strExcCd, "")
            TrsRaiseError m_strCMODULENAME, "ProcessAddException", l_lngErrNum, , l_colDescArgs
        Case Else
            TrsRaiseError m_strCMODULENAME, "ProcessAddException"
    End Select
End Function

