VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cbsSchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const mstrNotAvailable As String = "NA"
Private Const mstrNotAvailable2 As String = "Z#"
Private Const mstrDefaultDate As String = "01/01/0100"
Private Const m_strCMODULENAME As String = "busSchedule"

Public Enum TRSReturnOption
    trBaseSchedNoDetails = 1
    trAdjustedBaseSched = 2
    trAdjBaseSchedWithExceptions = 3
    trAdjBaseSchedWithExceptionsAndOvertime = 4
End Enum
Public Enum TRSSchedSourceStartLevel
    trAllScheduleLevels = 50
    trDemandLineTrades = 50
    trDemandLineBaseSchedLevel = 40
    trEmployeeTradesLevel = 30
    trNonDemandLevel = 20
    trWorkGroupBaseScheduleLevel = 10
End Enum

Public Enum TRSLineReturnOption
    trLineBaseSchedNoDetails = 1
    trAdjustedLineBaseSched = 2
End Enum
Public Enum TRSLineSchedSourceStartLevel
    trLineAllLevels = 50
    trLineTrades = 50
    trLineBaseSched = 40
End Enum

Public Enum TRSSchedLineAssignmentSource
    trRelief = 20
    trDemandScheduling = 10
End Enum

Public Enum TRSScheduledExceptionsOption
    trAllExceptions = 1
    trOnlyPendingExceptions = 2
End Enum

Public Enum TRSScheduleTypes
    trScheduleChange = 10
    trTradeDoubleShift = 20
    trPartialShiftException = 110
    trFullShiftException = 115
    trOvertimeException = 120
End Enum

Public Enum TRSScheduleSources
    trWorkGroupBaseSchedule = 10
    trNonDemandSchedule = 20
    trDemandSchedule = 30
    trScheduleChanged = 41
    trTradedDouble = 42
End Enum

'do not use this public enum - it will be deleted begin
Public Enum TRSbusScheduleErrorEnum
    trErrInvalidRequestOnStartDate = 20002 + vbObjectError
    trErrInvalidRequestOption = 20003 + vbObjectError
    trErrEmployeeDoesNotExist = 30043 + vbObjectError
    trErrScheduleCodeDoesNotExist = 30044 + vbObjectError
    trErrEmpExceptionDoesNotMatchSchedule = 30045 + vbObjectError
End Enum
'do not use this public enum - it will be deleted end

Private m_rsRotations As ADODB.Recordset
Private m_varRepeatingShiftData(26) As Variant
'5/24/02 RSG m_rsWGEmpAssignments and m_rsWGBaseScheduleOnly were added to support
'cascade and applying pending workgroup changes.
'm_rsWGEmpAssignments contains employee primary workgroup assignments
'that is optionally passed.
'm_rsWGBaseScheduleOnly contains workgroup base schedules only, i.e. not join with assigments.
'This is used with separate employee primary workgroup assignments, m_rsWGEmpAssignments
'to apply pending changes, then joined and set to m_rsWGBaseSchedule.
Private m_rsWGEmpAssignments As ADODB.Recordset
Private m_rsWGBaseScheduleOnly As ADODB.Recordset
Private m_rsDemandLineAssignments As ADODB.Recordset
    'these are line assignments based on the employee
Private m_rsDemandLineBaseSchedule As ADODB.Recordset
Private m_rsDemandLineTrades As ADODB.Recordset
Private m_rsEmployeeTrades As ADODB.Recordset
Private m_rsPersonalSchedule As ADODB.Recordset
'm_rsWGBaseSchedule contains the result of a join on employee workgroup assignments
'and workgroup base schedules
Private m_rsWGBaseSchedule As ADODB.Recordset
Private m_rsExceptionSched As ADODB.Recordset
Private m_rsOvertimeSched As ADODB.Recordset
Private m_rsEmpJobNumByChargeDate As ADODB.Recordset
Private m_rsAvailSched As ADODB.Recordset
Private m_colSchedCodes As Collection
Private m_rsEmployeesInWorkGroup As ADODB.Recordset
Private m_rsWorkGroupLineAssignments As ADODB.Recordset
            'This is employees that have a line assignment in
            'selected work group
Private m_rsWorkGroupLineBaseSched As ADODB.Recordset
            'This is the base schedule of the line
            'regardless of who is assigned
            'This is needed to determine the schedule of
            'an unmanned line
Private m_rsWorkGroupLineTrade As ADODB.Recordset
            'This is needed to determine the schedule of
            'an unmanned line
Private m_aryEmployees As Variant
'Profiler
Private m_rsEmployeeData As ADODB.Recordset
'Optimistic locking
Private m_rsSchedTimeStamps As ADODB.Recordset
Private m_strScheduleTimeStamp As String
            






Private Function IsInList( _
        ByVal v_lngCode As Long, _
        ByVal v_colCodes As Collection _
        ) As Boolean

    Dim l_intIndex As Integer
    Dim l_intCount As Integer
    Dim l_bolInList As Boolean
    
    l_bolInList = False
    l_intIndex = 1
    l_intCount = v_colCodes.Count
    
    While Not l_bolInList And l_intIndex <= l_intCount
        If v_colCodes.Item(l_intIndex) = v_lngCode Then
            l_bolInList = True
        End If
        l_intIndex = l_intIndex + 1
    Wend
    
    IsInList = l_bolInList
    

End Function


Private Function GetUniqueSchedCodes(ByVal v_rsSchedCodeByDate As ADODB.Recordset) As Collection

    Dim l_intSchedCode As Integer
    Dim l_colSchedCodes As Collection
    
    On Error GoTo eh
    
    Set l_colSchedCodes = New Collection
    If RecordsExist(v_rsSchedCodeByDate) Then
        v_rsSchedCodeByDate.MoveFirst
        While Not v_rsSchedCodeByDate.EOF
            l_intSchedCode = v_rsSchedCodeByDate!SCHED_CD
            If IsInList(l_intSchedCode, l_colSchedCodes) Then
            Else
               l_colSchedCodes.Add l_intSchedCode
            End If
            l_intSchedCode = v_rsSchedCodeByDate!SCHED_CD_FOR_SHIFT_COMP
            If IsInList(l_intSchedCode, l_colSchedCodes) Then
            Else
               l_colSchedCodes.Add l_intSchedCode
            End If
            v_rsSchedCodeByDate.MoveNext
        Wend
    End If

    Set GetUniqueSchedCodes = l_colSchedCodes
    
    Exit Function
    
eh:

    TrsRaiseError m_strCMODULENAME, "GetUniqueSchedCodes"

End Function



Private Sub ApplyEmployeeBaseJobNum( _
                ByVal v_lngEmpNum As Long, _
                ByVal v_datStartDate As Date, _
                ByVal v_datEndDate As Date, _
                ByRef r_rsAdjustedBaseSched As ADODB.Recordset)

    Dim l_strJobNum As String
    Dim l_strSchedJobNum As String
    Dim l_datChargeDate As Date
    Dim l_strFilter As String

    On Error GoTo eh

    m_rsEmpJobNumByChargeDate.Filter = "EMP_NUM = " & str(v_lngEmpNum)
    If RecordsExist(m_rsEmpJobNumByChargeDate) Then
        m_rsEmpJobNumByChargeDate.Sort = "CHARGE_DATE"
        m_rsEmpJobNumByChargeDate.MoveFirst
    Else
        Exit Sub
    End If
    
    While Not m_rsEmpJobNumByChargeDate.EOF
        l_strJobNum = Trim(m_rsEmpJobNumByChargeDate!JOB_NUM)
        l_datChargeDate = m_rsEmpJobNumByChargeDate!CHARGE_DATE
        l_strFilter = "EMP_NUM = " & str(v_lngEmpNum) & _
                    " AND CHARGE_DATE = " & str(l_datChargeDate)
        l_strFilter = MDAC25FilterFix(l_strFilter)
        r_rsAdjustedBaseSched.Filter = l_strFilter
        If RecordsExist(r_rsAdjustedBaseSched) Then
            r_rsAdjustedBaseSched.MoveFirst
            While Not r_rsAdjustedBaseSched.EOF
                l_strSchedJobNum = Trim(r_rsAdjustedBaseSched!SCHED_JOB_NUM)
                If l_strSchedJobNum = "" Then
                    l_strSchedJobNum = l_strJobNum
                End If
                If l_strJobNum = l_strSchedJobNum Then
                    r_rsAdjustedBaseSched!SCHED_JOB_NUM = ""
                End If
                r_rsAdjustedBaseSched!EMP_JOB_NUM = l_strJobNum
                r_rsAdjustedBaseSched.Update
                r_rsAdjustedBaseSched.MoveNext
            Wend
        End If
        r_rsAdjustedBaseSched.Filter = ""
        m_rsEmpJobNumByChargeDate.MoveNext
    Wend
    
        Exit Sub
    
eh:

    TrsRaiseError m_strCMODULENAME, "ApplyEmployeeBaseJobNum"

End Sub
Private Sub ApplyOvertimeExceptions( _
                ByRef r_rsAdjustedBaseSched As ADODB.Recordset, _
                ByVal v_rsExceptionSched As ADODB.Recordset, _
                Optional v_rsSpecificDates As ADODB.Recordset = Nothing)
                            
    Dim l_datChargeDate As Date
    Dim l_datDateWorked As Date
    Dim l_strStartTime As String
    Dim l_strEndTime As String
    Dim l_decHours As String
    Dim l_strShiftComponent As String
    Dim l_strFilter As String
    Dim l_bolCrossesMidnight As Boolean
    Dim l_lngEmpNum As Long
    Dim l_lngWgNum As Long
    Dim l_lngLineNum As Long
    Dim l_blnApplyException As Boolean
    
    v_rsExceptionSched.MoveFirst
    While Not v_rsExceptionSched.EOF
        l_datChargeDate = CDate(v_rsExceptionSched!CHARGE_DATE)
        If v_rsSpecificDates Is Nothing Then
            l_blnApplyException = True
        Else
            l_strFilter = "EMP_NUM = " & str(l_lngEmpNum)
            l_strFilter = l_strFilter & "  AND  "
            l_strFilter = l_strFilter & "START_DATE <= " & str(l_datChargeDate)
            l_strFilter = l_strFilter & "  AND  "
            l_strFilter = l_strFilter & "END_DATE >= " & str(l_datChargeDate)
            v_rsSpecificDates.Filter = MDAC25FilterFix(l_strFilter)
            If RecordsExist(v_rsSpecificDates) Then
                l_blnApplyException = True
            Else
                l_blnApplyException = False
            End If
        End If
        
        If l_blnApplyException Then
            l_datDateWorked = CDate(v_rsExceptionSched!Date)
            l_strStartTime = v_rsExceptionSched!START_TIME
            l_strEndTime = v_rsExceptionSched!END_TIME
            l_lngEmpNum = v_rsExceptionSched!EMP_NUM
            If l_strEndTime = "00:00" Then
                l_strEndTime = "24:00"
            End If
            If l_strStartTime > l_strEndTime Then
                l_bolCrossesMidnight = True
            Else
                l_bolCrossesMidnight = False
            End If
            l_strFilter = "EMP_NUM = " & str(l_lngEmpNum)
            l_strFilter = l_strFilter & "  AND  "
            l_strFilter = l_strFilter & "CHARGE_DATE = " & str(l_datChargeDate)
            r_rsAdjustedBaseSched.Filter = ""
     'Get Shift Component from base Schedule for that day
            If RecordsExist(r_rsAdjustedBaseSched) Then
                r_rsAdjustedBaseSched.MoveFirst
                r_rsAdjustedBaseSched.Filter = MDAC25FilterFix(l_strFilter)
                If RecordsExist(r_rsAdjustedBaseSched) Then
                    InitializeExceptionRepeatingShiftData r_rsAdjustedBaseSched, True
                    l_strShiftComponent = r_rsAdjustedBaseSched!COMP_SHIFT_CD
                Else
                    l_strShiftComponent = ""
                End If
            End If
            
            
            If l_bolCrossesMidnight Then
                l_strEndTime = "24:00"
            End If
            r_rsAdjustedBaseSched.AddNew
            SetOvertimeRepeatingShiftData r_rsAdjustedBaseSched
            r_rsAdjustedBaseSched!CHARGE_DATE = l_datChargeDate
            r_rsAdjustedBaseSched!DATE_WORKED = l_datDateWorked
            r_rsAdjustedBaseSched!START_TIME = l_strStartTime
            r_rsAdjustedBaseSched!END_TIME = l_strEndTime
            If l_strEndTime = "24:00" Then
                l_decHours = 24 - (Hour(l_strStartTime) + Minute(l_strStartTime) / 60)
            Else
                l_decHours = (Hour(l_strEndTime) + Minute(l_strEndTime) / 60) - _
                    (Hour(l_strStartTime) + Minute(l_strStartTime) / 60)
            End If
    
            r_rsAdjustedBaseSched!HOURS = Round(l_decHours, 1)
            r_rsAdjustedBaseSched!TRAV_PAY_HOURS = CDec(v_rsExceptionSched!TRAV_PAY_HRS)
            r_rsAdjustedBaseSched!OT_OPP_NUM = v_rsExceptionSched!OT_OPP_NUM
            SetExceptionData r_rsAdjustedBaseSched, v_rsExceptionSched
            If Trim(r_rsAdjustedBaseSched!COMP_SHIFT_CD) = "" Then
                r_rsAdjustedBaseSched!COMP_SHIFT_CD = l_strShiftComponent
            End If
            r_rsAdjustedBaseSched.Update
            
            If l_bolCrossesMidnight Then
                l_strStartTime = "00:00"
                l_strEndTime = v_rsExceptionSched!END_TIME
                r_rsAdjustedBaseSched.AddNew
                InitializeExceptionRepeatingShiftData r_rsAdjustedBaseSched
                r_rsAdjustedBaseSched!CHARGE_DATE = l_datChargeDate
                If l_datDateWorked = l_datChargeDate Then
                    l_datDateWorked = DateAdd("d", 1, l_datDateWorked)
                End If
                r_rsAdjustedBaseSched!DATE_WORKED = l_datDateWorked
                r_rsAdjustedBaseSched!START_TIME = l_strStartTime
                r_rsAdjustedBaseSched!END_TIME = l_strEndTime
                l_decHours = (Hour(l_strEndTime) + Minute(l_strEndTime) / 60) - _
                                (Hour(l_strStartTime) + Minute(l_strStartTime) / 60)
                r_rsAdjustedBaseSched!HOURS = Round(l_decHours, 1)
                r_rsAdjustedBaseSched!TRAV_PAY_HOURS = CDec(v_rsExceptionSched!TRAV_PAY_HRS)
                r_rsAdjustedBaseSched!OT_OPP_NUM = v_rsExceptionSched!OT_OPP_NUM
                SetExceptionData r_rsAdjustedBaseSched, v_rsExceptionSched
                If Trim(r_rsAdjustedBaseSched!COMP_SHIFT_CD) = "" Then
                    r_rsAdjustedBaseSched!COMP_SHIFT_CD = l_strShiftComponent
                End If
                r_rsAdjustedBaseSched.Update
            End If
        End If
        v_rsExceptionSched.MoveNext
    Wend
    
    r_rsAdjustedBaseSched.Filter = ""

End Sub


Private Sub ApplyExceptions( _
                ByRef r_rsAdjustedBaseSched As ADODB.Recordset, _
                ByVal v_rsExceptionSched As ADODB.Recordset, _
                Optional v_rsSpecificDates As ADODB.Recordset = Nothing)
                            
' Exceptions entered in Exception Scheduling, stored on ttr_emp_sched_exc,
'  are applied on top of the schedule.  Basically providing more specific
'  information such as Vacation (V) or Illness (E) to the schedule.
' The schedule (shift_num) for the Exception must match a schedule on the
'  same charge date.



    Dim l_datChargeDate As Date
    Dim l_intSchedCode As Integer
    Dim l_intSchedType As Integer
    Dim l_strFilter As String
    Dim l_colDescArgs As Collection
    Dim l_strException As String
    Dim l_intAdjBaseSchedCd As Integer
    Dim l_intShiftNum As Integer
    Dim l_decExceptionHours As Variant
    Dim l_decExceptionOffSetHours As Variant
    Dim l_decShiftHours As Variant
    Dim l_decHoursApplied As Variant
    Dim l_lngEmpNum As Long
    Dim l_blnApplyException As Boolean
    Dim l_colErrArgs As Collection
    
On Error GoTo ApplyExceptions_EH
    
    v_rsExceptionSched.MoveFirst
    While Not v_rsExceptionSched.EOF
        l_lngEmpNum = v_rsExceptionSched!EMP_NUM
        l_datChargeDate = v_rsExceptionSched!CHARGE_DATE
        If v_rsSpecificDates Is Nothing Then
            l_blnApplyException = True
        Else
            l_strFilter = "EMP_NUM = " & str(l_lngEmpNum)
            l_strFilter = l_strFilter & "  AND  "
            l_strFilter = l_strFilter & "START_DATE <= " & str(l_datChargeDate)
            l_strFilter = l_strFilter & "  AND  "
            l_strFilter = l_strFilter & "END_DATE >= " & str(l_datChargeDate)
            v_rsSpecificDates.Filter = MDAC25FilterFix(l_strFilter)
            If RecordsExist(v_rsSpecificDates) Then
                l_blnApplyException = True
            Else
                l_blnApplyException = False
            End If
        End If
        
        If l_blnApplyException Then
            l_intSchedCode = v_rsExceptionSched!SCHED_CD
            l_strException = v_rsExceptionSched!EXCEPTION_CD
            l_intShiftNum = v_rsExceptionSched!SHIFT_NUM
            l_decExceptionOffSetHours = v_rsExceptionSched!HOURS_OFFSET_NUM
            l_decExceptionHours = v_rsExceptionSched!HOURS
            l_strFilter = "EMP_NUM = " & str(l_lngEmpNum)
            l_strFilter = l_strFilter & "  AND  "
            l_strFilter = l_strFilter & "CHARGE_DATE = " & str(l_datChargeDate)
            l_strFilter = l_strFilter & "  AND  "
            l_strFilter = l_strFilter & "SHIFT_NUM = " & str(l_intShiftNum)
            If RecordsExist(r_rsAdjustedBaseSched) Then
                r_rsAdjustedBaseSched.MoveFirst
                r_rsAdjustedBaseSched.Filter = MDAC25FilterFix(l_strFilter)
                If Not RecordsExist(r_rsAdjustedBaseSched) Then
                        r_rsAdjustedBaseSched.Filter = "CHARGE_DATE = " & str(l_datChargeDate)
                        If Not RecordsExist(r_rsAdjustedBaseSched) Then
                            l_intAdjBaseSchedCd = 0
                        Else
                            l_intAdjBaseSchedCd = r_rsAdjustedBaseSched!SCHED_CD
                        End If
                        Err.Raise trErrbusEmpExceptionDoesNotMatchSchedule
                End If
                
                l_decShiftHours = r_rsAdjustedBaseSched!SHIFT_HOURS
                If l_decShiftHours = 0 Then
                    If Trim(l_strException) = "HR" Or _
                        Trim(l_strException) = "H" Or _
                        Trim(l_strException) = "HB" Then
                        SetExceptionData r_rsAdjustedBaseSched, v_rsExceptionSched, True
                    Else
                        Err.Raise 30000 + vbObjectError, "busSchedule.ApplyExceptions", "Exception not valid on an off day."
                    End If
                Else
                    If Trim(l_strException) = "HR" Then
                        Err.Raise 30000 + vbObjectError, "busSchedule.ApplyExceptions", "HR (Holiday on a day of Rest) is only valid on an off day."
                    End If
                End If
                
                If (l_decExceptionHours + l_decExceptionOffSetHours) _
                    > l_decShiftHours And l_decShiftHours <> 0 Then
                    r_rsAdjustedBaseSched.Filter = "CHARGE_DATE = " & str(l_datChargeDate)
                    If Not RecordsExist(r_rsAdjustedBaseSched) Then
                        l_intAdjBaseSchedCd = 0
                    Else
                        l_intAdjBaseSchedCd = r_rsAdjustedBaseSched!SCHED_CD
                    End If
                    Err.Raise trErrbusEmpExceptionDoesNotMatchSchedule
                End If
                
                If l_decExceptionHours = l_decShiftHours Then
                    While Not r_rsAdjustedBaseSched.EOF
                        SetExceptionData r_rsAdjustedBaseSched, v_rsExceptionSched, True
                        r_rsAdjustedBaseSched.MoveNext
                    Wend
                Else
                    If l_decShiftHours > 0 Then
                        ApplyPartialShiftException r_rsAdjustedBaseSched, v_rsExceptionSched
                    End If
                End If
            End If
        End If
        r_rsAdjustedBaseSched.Filter = ""
        v_rsExceptionSched.MoveNext
    Wend
    Exit Sub
    
ApplyExceptions_EH:
    
    Set l_colErrArgs = New Collection
    l_colErrArgs.Add "l_lngEmpNum=" & CStr(l_lngEmpNum)
    l_colErrArgs.Add "l_datChargeDate=" & Format(l_datChargeDate, "mm/dd/yyyy")
    l_colErrArgs.Add "l_intSchedCode=" & CStr(l_intSchedCode)
    l_colErrArgs.Add "l_strException=" & l_strException
    l_colErrArgs.Add "l_intAdjBaseSchedCd=" & CStr(l_intAdjBaseSchedCd)
    Select Case Err.Number
        Case trErrbusEmpExceptionDoesNotMatchSchedule
            Set l_colDescArgs = New Collection
            l_colDescArgs.Add str(l_intSchedCode)
            l_colDescArgs.Add l_strException
            l_colDescArgs.Add str(l_intAdjBaseSchedCd)
            l_colDescArgs.Add str(l_lngEmpNum)
            l_colDescArgs.Add str(l_datChargeDate)
            TrsRaiseError m_strCMODULENAME, "ApplyExceptions", trErrbusEmpExceptionDoesNotMatchSchedule, , l_colDescArgs, , True
        Case 30000 + vbObjectError
            TrsRaiseError m_strCMODULENAME, "ApplyExceptions", , , , l_colErrArgs, True
        Case Else
            TrsRaiseError m_strCMODULENAME, "ApplyExceptions", , , , l_colErrArgs
    End Select

    
End Sub


Private Sub ApplyPartialShiftException(ByRef r_rsAdjustedBaseSched As ADODB.Recordset, ByVal v_rsExceptionSched As ADODB.Recordset)

'This procedure will apply the partial shift exception to the shift
'It will loop thru each existing schedule split and decide if the
' partial needs to be applied to it.
'If the partial is only for part of the existing schedule split, it will
'another split
'Exceptions that cross midnight are processed as 2 exceptions (Before and After midnight)


    Dim l_strSchedStartTime As String
    Dim l_strSchedEndTime As String
    Dim l_strExceptionStartTime As String
    Dim l_strExceptionEndTime As String
    Dim l_strSplitExceptionEndTime As String
    Dim l_bolExceptionCrossesMidnight As Boolean
    Dim l_intExceptions As Integer
    Dim l_intIndex As Integer
    Dim l_intRecords As Integer
    Dim l_intPosition As Integer
    
    InitializeExceptionRepeatingShiftData r_rsAdjustedBaseSched
    DetermineExceptionTimes l_strExceptionStartTime, l_strExceptionEndTime, _
                            r_rsAdjustedBaseSched, v_rsExceptionSched
                            
    If l_strExceptionEndTime = "00:00" Then
        l_strExceptionEndTime = "24:00"
    End If
    
    If l_strExceptionEndTime < l_strExceptionStartTime Then 'Exception Crosses Midnight
                                                        'will process as 2 exceptions
        l_bolExceptionCrossesMidnight = True
        l_intExceptions = 2
    Else
        l_bolExceptionCrossesMidnight = False
        l_intExceptions = 1
    End If
    
    l_intRecords = r_rsAdjustedBaseSched.RecordCount
    For l_intIndex = 1 To l_intExceptions
        If l_bolExceptionCrossesMidnight Then
            If l_intIndex = 1 Then 'before midnight
                l_strSplitExceptionEndTime = "24:00"
            Else 'after midnight
                l_strSplitExceptionEndTime = l_strExceptionEndTime
                l_strExceptionStartTime = "00:00"
            End If
        Else
             l_strSplitExceptionEndTime = l_strExceptionEndTime
        End If
        l_intPosition = 1
        While l_intPosition <= l_intRecords
            r_rsAdjustedBaseSched.AbsolutePosition = l_intPosition
            If Trim(r_rsAdjustedBaseSched!SYSTEM_EXC_CD) = "SLT" Then
            Else
                l_strSchedStartTime = r_rsAdjustedBaseSched!START_TIME
                l_strSchedEndTime = r_rsAdjustedBaseSched!END_TIME
                If l_strExceptionStartTime <= l_strSchedStartTime And l_strSplitExceptionEndTime >= l_strSchedEndTime Then
                  'Schedule is totally within the exception time
                    SetExceptionData r_rsAdjustedBaseSched, v_rsExceptionSched
                Else 'Check to see if the Schedule is totally outside the Exception
                     'Note start time is always less than end time
                    If l_strExceptionStartTime >= l_strSchedEndTime _
                    Or l_strSchedStartTime >= l_strSplitExceptionEndTime Then
                    'Exception does not need to be applied to this portion of the schedule
                    Else 'Exception causes the schedule to split
                        SplitOnPartialShiftException r_rsAdjustedBaseSched, v_rsExceptionSched, l_strExceptionStartTime, l_strSplitExceptionEndTime
                    End If
                End If
            End If
            l_intPosition = l_intPosition + 1
        Wend
    Next l_intIndex
    
    
End Sub
Private Sub SetExceptionData( _
                ByRef r_rsAdjustedBaseSched As ADODB.Recordset, _
                ByVal v_rsExceptionSched As ADODB.Recordset, _
                Optional v_blnFullShiftExc As Boolean = False)

    Dim l_strExceptionCode As String
    Dim l_strOverExcCode As String
    Dim l_lngWorkGroup As Long
    Dim l_lngCaseNumber As Long
    Dim l_strJobNumber As String
    Dim l_strFMLAind As String
    Dim l_strReliefCode As String
    Dim l_strShiftComponent As String
    Dim l_strComponentCode As String  'used to give pay increase for hazard pay(ie climbing utility pole)
    
    If Trim(r_rsAdjustedBaseSched!SYSTEM_EXC_CD) <> "SLT" _
        And Trim(r_rsAdjustedBaseSched!SYSTEM_EXC_CD) <> "SSO" Then
        l_strExceptionCode = v_rsExceptionSched!EXCEPTION_CD
        r_rsAdjustedBaseSched!EXCEPTION_CD = l_strExceptionCode
'Per Greg Barks Request System Exception Code will not be populated with the exception code
'The system exception code will be SWT, SLT, or SSO
'        r_rsAdjustedBaseSched!SYSTEM_EXC_CD = l_strExceptionCode
        r_rsAdjustedBaseSched!OVER_EXC_CD = v_rsExceptionSched!OVER_EXC_CD
        l_strJobNumber = Trim(v_rsExceptionSched!JOB_NUM)
        If l_strJobNumber <> "" Then
            r_rsAdjustedBaseSched!SCHED_JOB_NUM = l_strJobNumber
        End If
        l_strShiftComponent = Trim(v_rsExceptionSched!SHIFT_COMP)
        If l_strShiftComponent <> "" Then
            r_rsAdjustedBaseSched!COMP_SHIFT_CD = l_strShiftComponent
        End If
        l_strReliefCode = Trim(v_rsExceptionSched!RELIEF_COMP)
        If l_strReliefCode <> "" Then
            r_rsAdjustedBaseSched!REL_CD = l_strReliefCode
        End If
        If Trim(v_rsExceptionSched!COMP_CD) <> "" Then
            r_rsAdjustedBaseSched!COMP_CD = Trim(v_rsExceptionSched!COMP_CD)
        End If
        If v_rsExceptionSched!WG_NUM > 0 Then
            r_rsAdjustedBaseSched!WG_NUM = v_rsExceptionSched!WG_NUM
        End If
        r_rsAdjustedBaseSched!CASE_NUM = v_rsExceptionSched!CASE_NUM
        r_rsAdjustedBaseSched!FMLA_IND = v_rsExceptionSched!FMLA_IND
        
    'Need to handle holiday on a day of rest
'        If IsAHoliday(l_strExceptionCode) And
        If (Trim(l_strExceptionCode) = "HR" Or _
            Trim(l_strExceptionCode) = "H") And _
            CDec(r_rsAdjustedBaseSched!SHIFT_HOURS) = 0 Then
            r_rsAdjustedBaseSched!START_TIME = "08:00"
            r_rsAdjustedBaseSched!END_TIME = "16:00"
            r_rsAdjustedBaseSched!HOURS = 8
            r_rsAdjustedBaseSched!EXCEPTION_CD = "HR"
        End If
        If Trim(l_strExceptionCode) = "HB" And _
            CDec(r_rsAdjustedBaseSched!SHIFT_HOURS) = 0 Then
            r_rsAdjustedBaseSched!START_TIME = "08:00"
            r_rsAdjustedBaseSched!END_TIME = "16:00"
            r_rsAdjustedBaseSched!HOURS = 8
            r_rsAdjustedBaseSched!EXCEPTION_CD = "HB"
        End If
        
        If v_blnFullShiftExc Then
            r_rsAdjustedBaseSched!FULL_SHIFT_EXC_IND = "Y"
        End If
        r_rsAdjustedBaseSched!LAST_MOD_EMP_NUM = v_rsExceptionSched!LAST_MOD_EMP_NUM
        r_rsAdjustedBaseSched!LAST_MOD_TIMESTAMP = v_rsExceptionSched!LAST_MOD_TIMESTAMP
        
        r_rsAdjustedBaseSched.Update
    End If
    
    
 
End Sub

Private Sub SplitOnPartialShiftException(ByRef r_rsAdjustedBaseSched As ADODB.Recordset, _
                                         ByVal v_rsExceptionSched As ADODB.Recordset, _
                                         ByVal v_strExceptionStartTime As String, _
                                         ByVal v_strExceptionEndTime As String)

    Dim l_strSchedStartTime As String
    Dim l_strSchedEndTime As String
    Dim l_strExceptionStartTime As String
    Dim l_strExceptionEndTime As String
    Dim l_datDateWorked As String
    Dim l_strSchedExceptionCode As String
    Dim l_strSystemExceptionCode As String
    Dim l_SplitType As Integer
'The partial exception will cause the schedule to be split one of 3 ways
' 1. Middle Split (Scheduled Time, Partial Exception, Scheduled Time)
' 2. End of shift (Scheduled Time, Partial Exception)
' 3. Begin of Shift (Partial Exception, Scheduled Time)
    
    Dim l_strStartTime As String
    Dim l_strEndTime As String
    Dim l_decHours As Variant
    
    l_strExceptionStartTime = v_strExceptionStartTime
    l_strExceptionEndTime = v_strExceptionEndTime
    l_strSchedStartTime = r_rsAdjustedBaseSched!START_TIME
    l_strSchedEndTime = r_rsAdjustedBaseSched!END_TIME
    
    l_datDateWorked = r_rsAdjustedBaseSched!DATE_WORKED
    l_strSchedExceptionCode = r_rsAdjustedBaseSched!EXCEPTION_CD
    l_strSystemExceptionCode = r_rsAdjustedBaseSched!SYSTEM_EXC_CD

'exception falls within a previous schedule split
    If l_strExceptionStartTime > l_strSchedStartTime And _
       l_strExceptionEndTime < l_strSchedEndTime Then 'middle split
        l_SplitType = 1    'Middle Split
    End If
    If l_strExceptionStartTime > l_strSchedStartTime And _
       l_strExceptionEndTime >= l_strSchedEndTime Then 'Begin split
        l_SplitType = 2   'Exception at End of Schedule time
    End If
    If l_strExceptionStartTime <= l_strSchedStartTime And _
       l_strExceptionEndTime < l_strSchedEndTime Then 'Begin split
        l_SplitType = 3   'Exception at Beginning of Schedule time
    End If
    
    Select Case l_SplitType  'Deterimining 1st Part of Split Start And End Times
        Case 1, 2
            l_strStartTime = l_strSchedStartTime
            l_strEndTime = l_strExceptionStartTime
        Case 3
            l_strStartTime = l_strSchedStartTime
            l_strEndTime = l_strExceptionEndTime
    End Select
    
'Update First Part of Split
    If l_strEndTime <> "24:00" Then
        l_decHours = (Hour(l_strEndTime) + Minute(l_strEndTime) / 60) - _
            (Hour(l_strStartTime) + Minute(l_strStartTime) / 60)
    Else
        l_decHours = 24 - (Hour(l_strStartTime) + Minute(l_strStartTime) / 60)
    End If
    r_rsAdjustedBaseSched!START_TIME = l_strStartTime
    r_rsAdjustedBaseSched!END_TIME = l_strEndTime
    r_rsAdjustedBaseSched!HOURS = Round(l_decHours, 1)
    If l_SplitType = 3 Then ' First Part of Split is the Partial Exception Code
        SetExceptionData r_rsAdjustedBaseSched, v_rsExceptionSched
    End If
    r_rsAdjustedBaseSched.Update
    
'Add Second Part of Split Part of Split
    r_rsAdjustedBaseSched.AddNew
    SetRepeatingShiftData r_rsAdjustedBaseSched
    r_rsAdjustedBaseSched!DATE_WORKED = l_datDateWorked
    r_rsAdjustedBaseSched!EXCEPTION_CD = l_strSchedExceptionCode
    r_rsAdjustedBaseSched!SYSTEM_EXC_CD = l_strSystemExceptionCode

    Select Case l_SplitType  'Deterimining 2nd Part of Split Start And End Times
        Case 1
            l_strStartTime = l_strExceptionStartTime
            l_strEndTime = l_strExceptionEndTime
        Case 2
            l_strStartTime = l_strExceptionStartTime
            l_strEndTime = l_strSchedEndTime
        Case 3
            l_strStartTime = l_strExceptionEndTime
            l_strEndTime = l_strSchedEndTime
    End Select

    r_rsAdjustedBaseSched!START_TIME = l_strStartTime
    r_rsAdjustedBaseSched!END_TIME = l_strEndTime
    If l_strEndTime <> "24:00" Then
        l_decHours = (Hour(l_strEndTime) + Minute(l_strEndTime) / 60) - _
            (Hour(l_strStartTime) + Minute(l_strStartTime) / 60)
    Else
        l_decHours = 24 - (Hour(l_strStartTime) + Minute(l_strStartTime) / 60)
    End If
    r_rsAdjustedBaseSched!HOURS = Round(l_decHours, 1)
    If l_SplitType <> 3 Then ' Second Part of Split is the Partial Exception Code
        SetExceptionData r_rsAdjustedBaseSched, v_rsExceptionSched
    End If
    r_rsAdjustedBaseSched.Update
    
'Add Third Part of Split for a MiddleSplit
    If l_SplitType = 1 Then
        r_rsAdjustedBaseSched.AddNew
        SetRepeatingShiftData r_rsAdjustedBaseSched
        r_rsAdjustedBaseSched!DATE_WORKED = l_datDateWorked
        r_rsAdjustedBaseSched!EXCEPTION_CD = l_strSchedExceptionCode
        r_rsAdjustedBaseSched!SYSTEM_EXC_CD = l_strSystemExceptionCode
        l_strStartTime = l_strExceptionEndTime
        l_strEndTime = l_strSchedEndTime
        If l_strEndTime <> "24:00" Then
            l_decHours = (Hour(l_strEndTime) + Minute(l_strEndTime) / 60) - _
                (Hour(l_strStartTime) + Minute(l_strStartTime) / 60)
        Else
            l_decHours = 24 - (Hour(l_strStartTime) + Minute(l_strStartTime) / 60)
        End If
        r_rsAdjustedBaseSched!START_TIME = l_strStartTime
        r_rsAdjustedBaseSched!END_TIME = l_strEndTime
        r_rsAdjustedBaseSched!HOURS = Round(l_decHours, 1)
        r_rsAdjustedBaseSched.Update
    End If
   
End Sub


Private Function CreateGetAdjBaseRecordset(ByVal v_ReturnOption As TRSReturnOption) As ADODB.Recordset
    Dim l_rsSchedule As ADODB.Recordset
    
    Set l_rsSchedule = New ADODB.Recordset
    With l_rsSchedule
      .Fields.Append "EMP_NUM", adInteger, , adFldUpdatable
      .Fields.Append "CHARGE_DATE", adDate, , adFldUpdatable
      .Fields.Append "DATE_WORKED", adDate, , adFldUpdatable
      .Fields.Append "SHIFT_START_TIME", adChar, 5, adFldUpdatable
      .Fields.Append "SHIFT_START_DATE", adDate, , adFldUpdatable
      .Fields.Append "SHIFT_END_TIME", adChar, 5, adFldUpdatable
      .Fields.Append "SHIFT_END_DATE", adDate, , adFldUpdatable
      .Fields.Append "SHIFT_HOURS", adDecimal, adFldUpdatable
      .Fields.Item("SHIFT_HOURS").Precision = 3
      .Fields.Item("SHIFT_HOURS").NumericScale = 1
      .Fields.Append "START_TIME", adChar, 5, adFldUpdatable
      .Fields.Append "END_TIME", adChar, 5, adFldUpdatable
      .Fields.Append "HOURS", adDecimal, adFldUpdatable
      .Fields.Item("HOURS").Precision = 3
      .Fields.Item("HOURS").NumericScale = 1
      .Fields.Append "EXCEPTION_CD", adChar, 3, adFldUpdatable
      .Fields.Append "OVER_EXC_CD", adChar, 3, adFldUpdatable
      .Fields.Append "SCHED_JOB_NUM", adChar, 6, adFldUpdatable
      .Fields.Append "EMP_JOB_NUM", adChar, 6, adFldUpdatable
      .Fields.Append "COMP_SHIFT_CD", adChar, 3, adFldUpdatable
      .Fields.Append "SYS_COMP_SHIFT_CD", adChar, 3, adFldUpdatable
      .Fields.Append "WG_NUM", adInteger, 7, adFldUpdatable
      .Fields.Append "CASE_NUM", adInteger, , adFldUpdatable
      .Fields.Append "FMLA_IND", adChar, 1, adFldUpdatable
      .Fields.Append "COMP_CD", adChar, 3, adFldUpdatable
      .Fields.Append "REL_CD", adChar, 3, adFldUpdatable
      .Fields.Append "SYSTEM_EXC_CD", adChar, 3, adFldUpdatable
      .Fields.Append "SHIFT_CODE", adChar, 1, adFldUpdatable
      .Fields.Append "SCHED_SRC_CD", adInteger, , adFldUpdatable
      .Fields.Append "SCHED_CD", adInteger, , adFldUpdatable
      .Fields.Append "SCHED_CHANGE_PMT", adChar, 1, adFldUpdatable
      If v_ReturnOption >= trAdjBaseSchedWithExceptionsAndOvertime Then
        .Fields.Append "OT_OPP_NUM", adInteger, 7, adFldUpdatable
        .Fields.Append "TRAV_PAY_HOURS", adDecimal, adFldUpdatable
        .Fields.Item("TRAV_PAY_HOURS").Precision = 3
        .Fields.Item("TRAV_PAY_HOURS").NumericScale = 1
      End If
      .Fields.Append "LINE_NUM", adInteger, , adFldUpdatable
      .Fields.Append "REQ_LINE_IND", adChar, 1, adFldUpdatable
      .Fields.Append "SHIFT_START_END", adChar, 11, adFldUpdatable
      .Fields.Append "SCHED_CD_FOR_SHIFT_COMP", adInteger, , adFldUpdatable
      .Fields.Append "BASE_SCHED_SRC_CD", adInteger, , adFldUpdatable
      .Fields.Append "LINE_SRC_CD", adInteger, , adFldUpdatable
      .Fields.Append "SHIFT_NUM", adInteger, , adFldUpdatable
      .Fields.Append "SHIFT_LUNCH_START_TIME", adChar, 5, adFldUpdatable
      .Fields.Append "SHIFT_LUNCH_END_TIME", adChar, 5, adFldUpdatable
      .Fields.Append "TRADE_NUM", adInteger, , adFldUpdatable
      .Fields.Append "WG_LINE_IND", adChar, 1, adFldUpdatable
      .Fields.Append "FULL_SHIFT_EXC_IND", adChar, 1, adFldUpdatable
      .Fields.Append "DOUBLE_SHIFT_IND", adChar, 1, adFldUpdatable
      .Fields.Append "LAST_MOD_SCHED_TM", adChar, 26, adFldUpdatable
      .Fields.Append "LINE_WG", adInteger, , adFldUpdatable
      .Fields.Append "LAST_MOD_EMP_NUM", adInteger, , adFldUpdatable
      .Fields.Append "LAST_MOD_TIMESTAMP", adChar, 26, adFldUpdatable
      .Fields.Append "LINE_JOB_NUM", adChar, 6, adFldUpdatable
      .Fields.Append "SCHED_CHG_EXC_CD", adChar, 3, adFldUpdatable
      .Fields.Append "ROT_CD", adInteger, , adFldUpdatable
      .Fields.Append "ROT_WEEK", adInteger, , adFldUpdatable

      
      .CursorLocation = adUseClient
      .CursorType = adOpenDynamic
      .LockType = adLockOptimistic
      .Open
    End With
    
    Set CreateGetAdjBaseRecordset = l_rsSchedule
    

End Function

Private Function CreateBaseScheduleRecordset() As ADODB.Recordset
    Dim l_rsSchedule As ADODB.Recordset
    
    Set l_rsSchedule = New ADODB.Recordset
    With l_rsSchedule
      .Fields.Append "EMP_NUM", adInteger, , adFldUpdatable
      .Fields.Append "CHARGE_DATE", adDate, , adFldUpdatable
      .Fields.Append "SCHED_CD", adInteger, , adFldUpdatable
      .Fields.Append "SHIFT_CODE", adChar, 1, adFldUpdatable
      .Fields.Append "SCHED_SRC_CD", adInteger, , adFldUpdatable
      .Fields.Append "LINE_SRC_CD", adInteger, , adFldUpdatable
      .Fields.Append "SHIFT_NUM", adInteger, , adFldUpdatable
      .Fields.Append "SHIFT_START_TIME", adChar, 5, adFldUpdatable
      .Fields.Append "SHIFT_START_DATE", adDate, , adFldUpdatable
      .Fields.Append "SHIFT_END_TIME", adChar, 5, adFldUpdatable
      .Fields.Append "SHIFT_END_DATE", adDate, , adFldUpdatable
      .Fields.Append "SHIFT_HOURS", adDecimal, adFldUpdatable
      .Fields.Item("SHIFT_HOURS").Precision = 3
      .Fields.Item("SHIFT_HOURS").NumericScale = 1
      .Fields.Append "COMP_SHIFT_CD", adChar, 3, adFldUpdatable
      .Fields.Append "SYS_COMP_SHIFT_CD", adChar, 3, adFldUpdatable
      .Fields.Append "SCHED_CD_FOR_SHIFT_COMP", adInteger, , adFldUpdatable
      .Fields.Append "BASE_SCHED_SRC_CD", adInteger, , adFldUpdatable
      .Fields.Append "SHIFT_LUNCH_START_TIME", adChar, 5, adFldUpdatable
      .Fields.Append "SHIFT_LUNCH_END_TIME", adChar, 5, adFldUpdatable
      .Fields.Append "TRADE_NUM", adInteger, , adFldUpdatable
      .Fields.Append "DOUBLE_SHIFT_IND", adChar, 1, adFldUpdatable
      .Fields.Append "LAST_MOD_SCHED_TM", adChar, 26, adFldUpdatable
      .Fields.Append "LINE_NUM", adInteger, , adFldUpdatable
      .Fields.Append "WG_NUM", adInteger, 7, adFldUpdatable
      .Fields.Append "LINE_WG", adInteger, , adFldUpdatable
      .Fields.Append "WG_LINE_IND", adChar, 1, adFldUpdatable
      .Fields.Append "OT_START_TIME", adChar, 5, adFldUpdatable
      .Fields.Append "OT_END_TIME", adChar, 5, adFldUpdatable
      .Fields.Append "OT_HOURS", adDecimal, adFldUpdatable
      .Fields.Item("OT_HOURS").Precision = 3
      .Fields.Item("OT_HOURS").NumericScale = 1
      .CursorLocation = adUseClient
      .CursorType = adOpenDynamic
      .LockType = adLockOptimistic
      .Open
    End With
    '      .Fields.Append "SHIFT_COMP", adChar, 1, adFldUpdatable
    
    Set CreateBaseScheduleRecordset = l_rsSchedule
    

End Function

Private Function AllSchedulesAreMissing() As Boolean

    Dim l_intScheduleCount As Integer
    Dim l_intIndex As Integer
    Dim l_bolScheduleMissing As Boolean
    
    l_intScheduleCount = m_colSchedCodes.Count
    l_intIndex = 1
    l_bolScheduleMissing = True
    While l_bolScheduleMissing And l_intIndex <= l_intScheduleCount
        If m_colSchedCodes.Item(l_intIndex) <> -1 Then
            l_bolScheduleMissing = False
        Else
            l_intIndex = l_intIndex + 1
        End If
    Wend

    AllSchedulesAreMissing = l_bolScheduleMissing

End Function

Private Sub BuildMissingSchedule(ByRef r_rsGetAdjBaseSched As ADODB.Recordset, ByVal v_rsSchedCodeByDate As ADODB.Recordset)


    If Not RecordsExist(v_rsSchedCodeByDate) Then
        Exit Sub
    End If


    v_rsSchedCodeByDate.MoveFirst
    While Not v_rsSchedCodeByDate.EOF
        r_rsGetAdjBaseSched.AddNew
        r_rsGetAdjBaseSched!EMP_NUM = v_rsSchedCodeByDate!EMP_NUM
        r_rsGetAdjBaseSched!CHARGE_DATE = v_rsSchedCodeByDate!CHARGE_DATE
        r_rsGetAdjBaseSched!SCHED_SRC_CD = v_rsSchedCodeByDate!SCHED_SRC_CD
        r_rsGetAdjBaseSched!SCHED_CD = v_rsSchedCodeByDate!SCHED_CD
        r_rsGetAdjBaseSched.Update
        v_rsSchedCodeByDate.MoveNext
    Wend
        

End Sub
Private Sub InitializeExceptionVariables( _
                    ByRef r_rsGetAdjBaseSched As ADODB.Recordset)
                    
    'Intialize Characters to blank and Numbers to 0
        r_rsGetAdjBaseSched!EXCEPTION_CD = " "
        r_rsGetAdjBaseSched!OVER_EXC_CD = " "
        r_rsGetAdjBaseSched!FMLA_IND = " "
        r_rsGetAdjBaseSched!COMP_CD = " "
        r_rsGetAdjBaseSched!CASE_NUM = 0
    'View Calendar extras
        r_rsGetAdjBaseSched!LAST_MOD_EMP_NUM = 0
        r_rsGetAdjBaseSched!LAST_MOD_TIMESTAMP = ""
        
        r_rsGetAdjBaseSched.Update
End Sub

Private Sub SetShiftData(ByRef r_rsGetAdjBaseSched As ADODB.Recordset, _
                         ByVal v_ReturnOption As TRSReturnOption)
                         
'gets shift data from ttr_avail_sched and perform necessary splits

    Dim l_strShiftStartDay As String
    Dim l_strDayofWeek As String 'SUN,MON, ... SAT - Representation of Charge Date
    Dim l_strShiftCode As String
    Dim l_strShiftComponent As String
    Dim l_strShiftStartTime As String
    Dim l_datShiftStartDate As Date
    Dim l_strShiftEndTime As String
    Dim l_datShiftEndDate As Date
    Dim l_decShiftHours As Variant
    Dim l_strLunchStart As String
    Dim l_strLunchEnd As String
    Dim l_strOTStart As String
    Dim l_strOTEnd As String
    Dim l_decOTHours As Variant
    Dim l_strFieldName As String
    Dim l_decHours As Variant
    Dim l_strStartTime As String
    Dim l_strEndTime As String
    Dim l_lngEmpNum As Long
    Dim l_datChargeDate As Date
    Dim l_datDateWorked As Date
    Dim l_intSchedCode As Integer
    Dim l_strJobNum As String
    Dim l_lngWgNum As Long
    Dim l_intLineNum As Integer
    Dim l_strReqInd As String
    Dim l_intSchedSource As Integer
    Dim l_strReliefComponent As String
    Dim l_strShiftComponentOverride As String  'users have the ability to supply a shift component
                                                ' on a schedule change.  This may be different than the one
                                                ' on the schedule.  Usually trades fall in this category
    Dim l_strSystemExceptionCode As String      ' This is needed to apply business rules for trades and relief
                                                ' Trades and relief get paid at straight time, but have rules
                                                ' for premimum pay on Sunday and Job upgrades
    Dim l_intSchedCodeForShiftComponent As Integer
    Dim l_intSchedLineSource As Integer
    Dim l_intShiftNum As Integer
    Dim l_strFilter As String
    Dim l_lngTradeNum As Long
    Dim l_intBaseSchedSource As Integer
    Dim l_intRotCode As Integer
    Dim l_intRotWeek As Integer
    Dim l_strSysShiftComponent As String
    
    l_lngEmpNum = r_rsGetAdjBaseSched!EMP_NUM
    l_datChargeDate = r_rsGetAdjBaseSched!CHARGE_DATE
    l_datDateWorked = l_datChargeDate
    l_intSchedCode = r_rsGetAdjBaseSched!SCHED_CD
    l_intSchedSource = r_rsGetAdjBaseSched!SCHED_SRC_CD
    l_intBaseSchedSource = r_rsGetAdjBaseSched!BASE_SCHED_SRC_CD
    l_lngTradeNum = r_rsGetAdjBaseSched!TRADE_NUM

'Get Data from Avail Available Schedule
    l_strDayofWeek = UCase(WeekdayName(Weekday(l_datChargeDate), True))   'returns day name SUN,MON,TUE, ..., SAT
    l_strShiftCode = Trim(m_rsAvailSched(l_strDayofWeek & "_SHIFT_CD"))
    l_strShiftComponent = Trim(m_rsAvailSched(l_strDayofWeek & "_COMP_SHIFT_CD"))
    l_strSysShiftComponent = l_strShiftComponent
    l_strShiftStartTime = Trim(m_rsAvailSched(l_strDayofWeek & "_START_TIME"))
    l_strShiftEndTime = Trim(m_rsAvailSched(l_strDayofWeek & "_END_TIME"))
    l_decShiftHours = m_rsAvailSched(l_strDayofWeek & "_STR_HRS")
    l_strLunchStart = Trim(m_rsAvailSched(l_strDayofWeek & "_LUNCH_START"))
    l_strLunchEnd = Trim(m_rsAvailSched(l_strDayofWeek & "_LUNCH_END"))
    l_strOTStart = Trim(m_rsAvailSched(l_strDayofWeek & "_OT_START_TIME"))
    l_strOTEnd = Trim(m_rsAvailSched(l_strDayofWeek & "_OT_END_TIME"))
    l_decOTHours = m_rsAvailSched(l_strDayofWeek & "_OT_HOURS")
    l_strShiftStartDay = m_rsAvailSched(l_strDayofWeek & "_ACT_DAY")

    l_intSchedCodeForShiftComponent = r_rsGetAdjBaseSched!SCHED_CD_FOR_SHIFT_COMP
    l_intShiftNum = r_rsGetAdjBaseSched!SHIFT_NUM
    l_intSchedLineSource = r_rsGetAdjBaseSched!LINE_SRC_CD
    If l_intSchedCodeForShiftComponent = l_intSchedCode Then
    Else
        If KeepsShiftComponent(l_lngEmpNum, l_datChargeDate) Then
            l_strFilter = "SCHED_CD = " & str(l_intSchedCodeForShiftComponent)
            m_rsAvailSched.Filter = l_strFilter
            If RecordsExist(m_rsAvailSched) Then
                l_strShiftComponent = Trim(m_rsAvailSched(l_strDayofWeek & "_COMP_SHIFT_CD"))
            End If
        End If
    End If
    
    DetermineShiftDates l_strShiftStartDay, l_strDayofWeek, l_datChargeDate, l_strShiftStartTime, l_strShiftEndTime, l_datShiftStartDate, l_datShiftEndDate
    If v_ReturnOption = trBaseSchedNoDetails Then
        r_rsGetAdjBaseSched!SHIFT_CODE = l_strShiftCode
        r_rsGetAdjBaseSched!SHIFT_START_TIME = l_strShiftStartTime
        r_rsGetAdjBaseSched!SHIFT_END_TIME = l_strShiftEndTime
        r_rsGetAdjBaseSched!SHIFT_START_DATE = l_datShiftStartDate
        r_rsGetAdjBaseSched!SHIFT_END_DATE = l_datShiftEndDate
        r_rsGetAdjBaseSched!SHIFT_HOURS = Round(l_decShiftHours, 1)
        r_rsGetAdjBaseSched!COMP_SHIFT_CD = l_strShiftComponent
        r_rsGetAdjBaseSched!SYS_COMP_SHIFT_CD = l_strSysShiftComponent
'        r_rsGetAdjBaseSched!SHIFT_COMP = l_strShiftComponent
        r_rsGetAdjBaseSched!SHIFT_LUNCH_START_TIME = l_strLunchStart
        r_rsGetAdjBaseSched!SHIFT_LUNCH_END_TIME = l_strLunchEnd
        r_rsGetAdjBaseSched!DOUBLE_SHIFT_IND = "N"
        r_rsGetAdjBaseSched!OT_START_TIME = l_strOTStart
        r_rsGetAdjBaseSched!OT_END_TIME = l_strOTEnd
        r_rsGetAdjBaseSched!OT_HOURS = Round(l_decOTHours, 1)
        r_rsGetAdjBaseSched.Update
        Exit Sub
    End If
    
    l_intRotCode = r_rsGetAdjBaseSched!ROT_CD
    l_intRotWeek = r_rsGetAdjBaseSched!ROT_WEEK
    l_lngWgNum = r_rsGetAdjBaseSched!WG_NUM
    l_intLineNum = r_rsGetAdjBaseSched!LINE_NUM
    l_strReqInd = r_rsGetAdjBaseSched!REQ_LINE_IND
    l_strJobNum = Trim(r_rsGetAdjBaseSched!SCHED_JOB_NUM)
    l_strReliefComponent = r_rsGetAdjBaseSched!REL_CD
    l_strShiftComponentOverride = Trim(r_rsGetAdjBaseSched!COMP_SHIFT_CD)
    l_strSystemExceptionCode = Trim(r_rsGetAdjBaseSched!SYSTEM_EXC_CD)
    l_strSystemExceptionCode = _
        DetermineTradeExceptionCode(l_strSystemExceptionCode, l_decShiftHours)
    If l_strSystemExceptionCode = "" Then
        l_strSystemExceptionCode = "SWT"
    End If
    r_rsGetAdjBaseSched!SYSTEM_EXC_CD = l_strSystemExceptionCode
    
    If l_strShiftComponentOverride > "" Then
        l_strShiftComponent = l_strShiftComponentOverride
    End If

    If l_strShiftCode = "" Then 'DAY OFF   Note: Data Analysis proves that there are 0 schedules where the shift code is blank and there are hours scheduled
                                'Also Shift hours = 0 and OT hours = 0 also indicate a day off
        l_strShiftStartTime = ""
        l_datShiftStartDate = l_datChargeDate
        l_strShiftEndTime = ""
        l_datShiftEndDate = l_datChargeDate
        l_decShiftHours = 0
        l_strLunchStart = ""
        l_strLunchEnd = ""
        l_strOTStart = ""
        l_strOTEnd = ""
        l_decOTHours = 0
    Else
        If l_decShiftHours = 0 Then  'No Straight time on day only OT
            l_strShiftStartTime = ""
            l_datShiftStartDate = l_datChargeDate
            l_strShiftEndTime = ""
            l_datShiftEndDate = l_datChargeDate
            l_strLunchStart = ""
            l_strLunchEnd = ""
        End If
    End If
'    DetermineShiftDates l_strShiftStartDay, l_strDayofWeek, l_datChargeDate, l_strShiftStartTime, l_strShiftEndTime, l_datShiftStartDate, l_datShiftEndDate
         
    InitializeRepeatingShiftData _
        l_lngEmpNum, l_datChargeDate, l_intSchedCode, _
        l_intSchedSource, l_strShiftStartTime, l_strShiftEndTime, _
        l_datShiftStartDate, l_datShiftEndDate, l_decShiftHours, _
        l_strShiftCode, l_strJobNum, l_intLineNum, l_lngWgNum, _
        l_strReqInd, l_strShiftComponent, l_strReliefComponent, _
        l_intSchedCodeForShiftComponent, l_intShiftNum, l_intSchedLineSource, _
        l_strLunchStart, l_strLunchEnd, l_lngTradeNum, l_intBaseSchedSource, _
        l_intRotCode, l_intRotWeek, l_strSysShiftComponent
        
    SetRepeatingShiftData r_rsGetAdjBaseSched

    l_datDateWorked = l_datShiftStartDate
    r_rsGetAdjBaseSched!DATE_WORKED = l_datDateWorked

    If l_strLunchStart > "" Then ' A LUNCH IS IN THE SCHEDULE
        l_strStartTime = l_strShiftStartTime
        l_strEndTime = l_strLunchStart
        l_decHours = (Hour(l_strEndTime) + Minute(l_strEndTime) / 60) - _
                     (Hour(l_strStartTime) + Minute(l_strStartTime) / 60)
        r_rsGetAdjBaseSched!HOURS = Round(CDec(l_decHours), 1)
        r_rsGetAdjBaseSched!START_TIME = l_strStartTime
        r_rsGetAdjBaseSched!END_TIME = l_strEndTime
'per Greg Barks Request do not populate exception code
'        r_rsGetAdjBaseSched!EXCEPTION_CD = "ST"
        r_rsGetAdjBaseSched!SYSTEM_EXC_CD = l_strSystemExceptionCode
        r_rsGetAdjBaseSched.Update
        If l_strStartTime > l_strEndTime Then
            BreakOnMidnight r_rsGetAdjBaseSched, l_strShiftStartDay, l_datDateWorked
        Else
            ShiftComponentSplits r_rsGetAdjBaseSched
        End If
        r_rsGetAdjBaseSched.AddNew
        SetRepeatingShiftData r_rsGetAdjBaseSched
        r_rsGetAdjBaseSched!DATE_WORKED = l_datDateWorked
' Do not populate job_num on lunch time
        r_rsGetAdjBaseSched!SCHED_JOB_NUM = ""
        l_strStartTime = l_strLunchStart
        l_strEndTime = l_strLunchEnd
        l_decHours = (Hour(l_strEndTime) + Minute(l_strEndTime) / 60) - _
                     (Hour(l_strStartTime) + Minute(l_strStartTime) / 60)
        r_rsGetAdjBaseSched!HOURS = Round(l_decHours, 1)
        r_rsGetAdjBaseSched!START_TIME = l_strStartTime
        r_rsGetAdjBaseSched!END_TIME = l_strEndTime
'per Greg Barks Request do not populate exception code
'        r_rsGetAdjBaseSched!EXCEPTION_CD = "LU"
        r_rsGetAdjBaseSched!SYSTEM_EXC_CD = "SLT"
        r_rsGetAdjBaseSched.Update
        If l_strStartTime > l_strEndTime Then
            BreakOnMidnight r_rsGetAdjBaseSched, l_strShiftStartDay, l_datDateWorked
        Else
            ShiftComponentSplits r_rsGetAdjBaseSched
        End If
        r_rsGetAdjBaseSched.AddNew
        SetRepeatingShiftData r_rsGetAdjBaseSched
        r_rsGetAdjBaseSched!DATE_WORKED = l_datDateWorked
        l_strStartTime = l_strLunchEnd
        l_strEndTime = l_strShiftEndTime
        l_decHours = (Hour(l_strEndTime) + Minute(l_strEndTime) / 60) - _
                     (Hour(l_strStartTime) + Minute(l_strStartTime) / 60)
        r_rsGetAdjBaseSched!HOURS = Round(l_decHours, 1)
        r_rsGetAdjBaseSched!START_TIME = l_strStartTime
        r_rsGetAdjBaseSched!END_TIME = l_strEndTime
'per Greg Barks Request do not populate exception code
'        r_rsGetAdjBaseSched!EXCEPTION_CD = "ST"
        r_rsGetAdjBaseSched!SYSTEM_EXC_CD = l_strSystemExceptionCode
        r_rsGetAdjBaseSched.Update
        If l_strStartTime > l_strEndTime Then
            BreakOnMidnight r_rsGetAdjBaseSched, l_strShiftStartDay, l_datDateWorked
        Else
            ShiftComponentSplits r_rsGetAdjBaseSched
        End If
    Else
        If l_decShiftHours > 0 Then  'Straight Time - NO LUNCH IS IN THE SCHEDULE
            l_strStartTime = l_strShiftStartTime
            l_strEndTime = l_strShiftEndTime
            r_rsGetAdjBaseSched!HOURS = Round(l_decShiftHours, 1)
            r_rsGetAdjBaseSched!START_TIME = l_strStartTime
            r_rsGetAdjBaseSched!END_TIME = l_strEndTime
'per Greg Barks Request do not populate exception code
'            r_rsGetAdjBaseSched!EXCEPTION_CD = "ST"
            r_rsGetAdjBaseSched!SYSTEM_EXC_CD = l_strSystemExceptionCode
            r_rsGetAdjBaseSched.Update
            If l_strStartTime > l_strEndTime Then
                BreakOnMidnight r_rsGetAdjBaseSched, l_strShiftStartDay, l_datDateWorked
            Else
                ShiftComponentSplits r_rsGetAdjBaseSched
            End If
        Else
            r_rsGetAdjBaseSched!START_TIME = ""
            r_rsGetAdjBaseSched!END_TIME = ""
            r_rsGetAdjBaseSched!HOURS = 0
            r_rsGetAdjBaseSched.Update
        End If
    End If
    
    If l_decOTHours > 0 Then ' OT in base schedule
        If l_decShiftHours > 0 Then
            r_rsGetAdjBaseSched.AddNew
        End If
        SetRepeatingShiftData r_rsGetAdjBaseSched
'If the Overtime was prior to the shift
'Code assumes overtime prior to shift must be
'contiguous to the straight time
'if no straight time, the shift start date is for overtime
        If l_strOTEnd = l_strShiftStartTime Then
            If l_strOTStart > l_strOTEnd Then
            'The OT Shift prior backs up to cross midnight
                l_datDateWorked = DateAdd("d", -1, l_datShiftStartDate)
            Else
                l_datDateWorked = l_datShiftStartDate
            End If
        End If
        r_rsGetAdjBaseSched!DATE_WORKED = l_datDateWorked
        r_rsGetAdjBaseSched!HOURS = Round(l_decOTHours, 1)
        r_rsGetAdjBaseSched!START_TIME = l_strOTStart
        r_rsGetAdjBaseSched!END_TIME = l_strOTEnd
        r_rsGetAdjBaseSched!EXCEPTION_CD = "SSO"
        r_rsGetAdjBaseSched!SYSTEM_EXC_CD = "SSO"
        InitializeExceptionVariables r_rsGetAdjBaseSched
        ShiftComponentSplits r_rsGetAdjBaseSched
        r_rsGetAdjBaseSched.Update
        If l_strOTStart > l_strOTEnd Then
            BreakOnMidnight r_rsGetAdjBaseSched, l_strShiftStartDay, l_datDateWorked
        End If
    End If
    

End Sub
Private Sub BreakOnMidnight( _
                ByRef r_rsGetAdjBaseSched As ADODB.Recordset, _
                ByVal l_strShiftStartDay As String, _
                ByRef r_datDateWorked As Date)

'Splits hours that cross midnight and determines date worked

    Dim l_datChargeDate As Date
    Dim l_datDateWorked As Date
    Dim l_decHours As Variant
    Dim l_strStartTime As String
    Dim l_strEndTime As String
    Dim l_strExceptionCode As String
    Dim l_strSystemExceptionCode As String
    Dim l_strDayofWeek As String  'SUN,MON, ... SAT - Representation of Charge Date

 
    l_datChargeDate = r_rsGetAdjBaseSched!CHARGE_DATE
    l_datDateWorked = r_rsGetAdjBaseSched!DATE_WORKED
    l_decHours = CDec(r_rsGetAdjBaseSched!HOURS)
    l_strStartTime = r_rsGetAdjBaseSched!START_TIME
    l_strEndTime = r_rsGetAdjBaseSched!END_TIME
    l_strExceptionCode = r_rsGetAdjBaseSched!EXCEPTION_CD
    l_strSystemExceptionCode = r_rsGetAdjBaseSched!SYSTEM_EXC_CD
    
    l_strDayofWeek = UCase(WeekdayName(Weekday(l_datChargeDate), True))


'change first piece of split
    If l_strShiftStartDay = l_strDayofWeek Then
        l_datDateWorked = l_datChargeDate
    Else
        l_datDateWorked = DateAdd("d", -1, l_datChargeDate)
    End If
    r_rsGetAdjBaseSched!DATE_WORKED = l_datDateWorked
    r_rsGetAdjBaseSched!END_TIME = "24:00"
    l_decHours = 24 - (Hour(l_strStartTime) + Minute(l_strStartTime) / 60)
    r_rsGetAdjBaseSched!HOURS = Round(CDec(l_decHours), 1)
    InitializeExceptionVariables r_rsGetAdjBaseSched
    ShiftComponentSplits r_rsGetAdjBaseSched
    r_rsGetAdjBaseSched.Update

'Add Second piece of split
    If l_strEndTime = "00:00" Then 'Do not add 2nd piece of time if end time is midnight
    Else
        r_rsGetAdjBaseSched.AddNew
        If l_strShiftStartDay = l_strDayofWeek Then
            l_datDateWorked = DateAdd("d", 1, l_datChargeDate)
        Else
            l_datDateWorked = l_datChargeDate
        End If
        SetRepeatingShiftData r_rsGetAdjBaseSched
        r_rsGetAdjBaseSched!DATE_WORKED = l_datDateWorked
        r_rsGetAdjBaseSched!END_TIME = l_strEndTime
        r_rsGetAdjBaseSched!EXCEPTION_CD = l_strExceptionCode
        r_rsGetAdjBaseSched!SYSTEM_EXC_CD = l_strSystemExceptionCode
        
        r_rsGetAdjBaseSched!DATE_WORKED = l_datDateWorked
        r_rsGetAdjBaseSched!START_TIME = "00:00"
        l_decHours = Hour(l_strEndTime) + Minute(l_strEndTime) / 60
        r_rsGetAdjBaseSched!HOURS = Round(CDec(l_decHours), 1)
        ShiftComponentSplits r_rsGetAdjBaseSched
        r_rsGetAdjBaseSched.Update
        
    End If
    
    r_datDateWorked = l_datDateWorked

End Sub




Private Function GetAvailSchedInfo(ByVal v_colSchedCodes As Collection) As ADODB.Recordset

On Error GoTo GetAvailSchedInfo_EH


    
    Dim l_dbSchedule As dbSchedule.cdSchedule
         
    
    Set l_dbSchedule = CreateObject("dbSchedule.cdSchedule")
    
    Set GetAvailSchedInfo = l_dbSchedule.GetAvailSchedInfo(v_colSchedCodes)
         
    Set l_dbSchedule = Nothing
    
    Exit Function

GetAvailSchedInfo_EH:
    Set l_dbSchedule = Nothing
    TrsRaiseError m_strCMODULENAME, "GetAvailSchedInfo"


 
End Function



Private Sub AddEmployeeSchedule( _
                ByRef r_rsAllEmployeesAdjBaseSched As ADODB.Recordset, _
                ByVal v_rsEmployeeAdjBaseSched As ADODB.Recordset)
                
    Dim l_fldField As ADODB.Field
    Dim l_strFieldName As String
                
    If Not RecordsExist(v_rsEmployeeAdjBaseSched) Then
        Exit Sub
    End If
                
    v_rsEmployeeAdjBaseSched.MoveFirst
    While Not v_rsEmployeeAdjBaseSched.EOF
        r_rsAllEmployeesAdjBaseSched.AddNew
        For Each l_fldField In v_rsEmployeeAdjBaseSched.Fields
            l_strFieldName = l_fldField.Name
            r_rsAllEmployeesAdjBaseSched(l_strFieldName).Value = v_rsEmployeeAdjBaseSched(l_strFieldName).Value
        Next
        v_rsEmployeeAdjBaseSched.MoveNext
    Wend
    

End Sub

Public Function GetAdjustedBaseSchedule _
             (ByVal v_aryEmployees As Variant, _
              ByVal v_datStartDate As Date, _
              ByVal v_datEndDate As Date, _
              ByRef r_rsScheduleWithoutExceptions As ADODB.Recordset, _
              Optional ByVal v_blnReturnScheduleWithoutExceptions As Boolean = False, _
              Optional ByVal v_ReturnOption As TRSReturnOption = trAdjustedBaseSched, _
              Optional ByVal v_ScheduleSourceStartLevel As TRSSchedSourceStartLevel = trAllScheduleLevels, _
              Optional ByVal v_SchedLineSource As TRSSchedLineAssignmentSource, _
              Optional ByVal v_ScheduleOption As TRSScheduledExceptionsOption = trOnlyPendingExceptions) _
              As ADODB.Recordset


    On Error GoTo GetAdjustedBaseSchedule_EH

    Dim l_lngI As Long
    Dim l_lngEmpNum As Long
    Dim l_lngStart As Long
    Dim l_lngEnd As Long
    Dim l_rsGetAdjBaseSched As ADODB.Recordset
    Dim l_rsScheduleWithoutExceptions As ADODB.Recordset
    Dim l_aryEmployees As Variant
    Dim l_intReturnOption As Integer
    Dim l_intScheduleSourceStartLevel As Integer

    l_intReturnOption = v_ReturnOption
    l_intScheduleSourceStartLevel = v_ScheduleSourceStartLevel
    ValidateRequest 0, v_datStartDate, v_datEndDate, l_intReturnOption, l_intScheduleSourceStartLevel
    InitializeRecordsets l_intReturnOption, v_blnReturnScheduleWithoutExceptions, l_rsGetAdjBaseSched, l_rsScheduleWithoutExceptions
    
    If IsArray(v_aryEmployees) Then
        l_aryEmployees = v_aryEmployees
    Else
        ReDim l_aryEmployees(0) As String
        l_aryEmployees(0) = str(CLng(v_aryEmployees))
    End If
    
'Retrieve Employee Schedule Data
'It place data in private global variables to reduce data passing
    GetEmployeeScheduleData l_aryEmployees, v_datStartDate, _
        v_datEndDate, v_ReturnOption, v_ScheduleSourceStartLevel
        
    Set m_rsRotations = GetRotationSchedules()
    
    GetAdjustedBaseSchedulePrivate l_aryEmployees, v_datStartDate, _
        v_datEndDate, Nothing, l_rsGetAdjBaseSched, l_rsScheduleWithoutExceptions, _
        v_blnReturnScheduleWithoutExceptions, l_intReturnOption, l_intScheduleSourceStartLevel
                
    Set r_rsScheduleWithoutExceptions = l_rsScheduleWithoutExceptions
    Set GetAdjustedBaseSchedule = l_rsGetAdjBaseSched
    
    ClearGlobalRecordSets
    Set l_rsGetAdjBaseSched = Nothing
    Set l_rsScheduleWithoutExceptions = Nothing
    Exit Function
    
GetAdjustedBaseSchedule_EH:
    ClearGlobalRecordSets
    Set r_rsScheduleWithoutExceptions = Nothing
    Set l_rsGetAdjBaseSched = Nothing
    Set l_rsScheduleWithoutExceptions = Nothing
    Dim l_colDescArgs As Collection
    Set l_colDescArgs = New Collection
    l_colDescArgs.Add str(v_datStartDate)
    l_colDescArgs.Add str(v_datEndDate)
    Select Case Err.Number
        Case trErrbusGetAdjustedBaseSchedInvalidDate
            TrsRaiseError m_strCMODULENAME, "GetAdjustedBaseSchedule", Err.Number, False, l_colDescArgs, , True
        Case Else
            TrsRaiseError m_strCMODULENAME, "GetAdjustedBaseSchedule", , , , , True
    End Select

End Function

Private Function DetermineEmployeeSchedules( _
            ByVal v_aryEmployees As Variant, _
            ByVal v_datStartDate As Date, _
            ByVal v_datEndDate As Date, _
            ByVal v_ScheduleSourceStartLevel As Integer, _
            Optional v_rsEmployeesAndDates As ADODB.Recordset = Nothing _
            ) As ADODB.Recordset
                                    
''''''''''
'  This Function is a wrapper that will call DetermineGetAdjBase
'   for each employee
'  The business rule for deriving a schedule
'       are in DetermineGetAdjBase
' v_rsEmployeesAndDates is used to output only on certain dates


    Dim l_rsAdjBaseSchedbyChargeDate As ADODB.Recordset
    Dim l_rsSchedCodesByChargeDate As ADODB.Recordset
    Dim l_lngStartEmp As Long
    Dim l_lngEndEmp As Long
    Dim l_lngI As Long
    Dim l_lngEmpNum As Long
    Dim l_datStartDate As Date
    Dim l_datEndDate As Date

    On Error GoTo eh

    Set l_rsAdjBaseSchedbyChargeDate = CreateGetAdjBaseRecordset(trAdjustedBaseSched)
    
    If v_rsEmployeesAndDates Is Nothing Then
        l_lngStartEmp = LBound(v_aryEmployees)
        l_lngEndEmp = UBound(v_aryEmployees)
        For l_lngI = l_lngStartEmp To l_lngEndEmp
             DetermineAdjBaseSched v_aryEmployees(l_lngI), v_datStartDate, v_datEndDate, v_ScheduleSourceStartLevel, l_rsAdjBaseSchedbyChargeDate
        Next l_lngI
    Else
        v_rsEmployeesAndDates.Filter = ""
        v_rsEmployeesAndDates.MoveFirst
        While Not v_rsEmployeesAndDates.EOF
            l_lngEmpNum = v_rsEmployeesAndDates!EMP_NUM
            l_datStartDate = v_rsEmployeesAndDates!START_DATE
            l_datEndDate = v_rsEmployeesAndDates!END_DATE
            DetermineAdjBaseSched l_lngEmpNum, l_datStartDate, l_datEndDate, v_ScheduleSourceStartLevel, l_rsAdjBaseSchedbyChargeDate
            v_rsEmployeesAndDates.MoveNext
        Wend
    End If

    Set DetermineEmployeeSchedules = l_rsAdjBaseSchedbyChargeDate
    Set l_rsAdjBaseSchedbyChargeDate = Nothing
    Set l_rsSchedCodesByChargeDate = Nothing
    
    Exit Function
    
eh:
    Set l_rsAdjBaseSchedbyChargeDate = Nothing
    Set l_rsSchedCodesByChargeDate = Nothing
    TrsRaiseError m_strCMODULENAME, "DetermineEmployeeSchedules"
    
    
End Function

'''''Private Function GetEmployeeSchedule _
'''''             (ByVal v_lngEmpNum As Long, _
'''''              ByVal v_datStartDate As Date, _
'''''              ByVal v_datEndDate As Date, _
'''''              ByVal v_rsSchedCodesByChargeDate As ADODB.Recordset, _
'''''              ByRef r_rsScheduleWithoutExceptions As ADODB.Recordset, _
'''''              ByVal v_ReturnOption As TRSReturnOption, _
'''''              ByVal v_blnReturnScheduleWithoutExceptions As Boolean) _
'''''              As ADODB.Recordset
'''''
''''''As of 04/17/2002 this is dead code JIMB
''''''May want to goback to the 04/15/2002 release
''''''This function built a schedule for 1 employee
''''''The caller looped thru the employees dumping the
''''''results of this function into it.
''''''There was no need to be copying the recordsets when filters will do
'''''
'''''
'''''    On Error GoTo GetEmployeeSchedule_EH
'''''
'''''    Dim l_rsAdjustedBaseSched As ADODB.Recordset
'''''    Dim l_strSort As String
'''''
''''''Need to get employee's job number for each date in the period
''''''Need to do this first because an error will be raised if it is an invalid employee
'''''    v_rsSchedCodesByChargeDate.Filter = "EMP_NUM = " & str(v_lngEmpNum)
'''''    v_rsSchedCodesByChargeDate.Sort = "CHARGE_DATE"
'''''
'''''' Fill in schedule details
'''''' if return option is AdjustedBaseSchedand,then perform splits over lunch and midnight
'''''    Set l_rsAdjustedBaseSched = BuildGetAdjBase(v_lngEmpNum, v_datStartDate, v_datEndDate, v_rsSchedCodesByChargeDate, v_ReturnOption)
'''''
''''''Save the employees schedule without applying exceptions
''''''This will allow the GUI to supply state to the server\
'''''' during the same time request.
'''''    If v_blnReturnScheduleWithoutExceptions Then
'''''        If v_ReturnOption = trBaseSchedNoDetails Then
'''''            Set r_rsScheduleWithoutExceptions = CreateBaseScheduleRecordset
'''''        Else
'''''            Set r_rsScheduleWithoutExceptions = CreateGetAdjBaseRecordset(v_ReturnOption)
'''''        End If
'''''        AddEmployeeSchedule r_rsScheduleWithoutExceptions, l_rsAdjustedBaseSched
'''''    Else
'''''         Set r_rsScheduleWithoutExceptions = Nothing
'''''    End If
''''''Apply full and partial exceptions to schedule
'''''
''''''JIMB Need Requirements
'''''' ask Greg about this stuff
'''''    If v_ReturnOption >= trAdjBaseSchedWithExceptions Then
'''''        m_rsExceptionSched.Filter = "EMP_NUM = " & str(v_lngEmpNum)
'''''        If RecordsExist(m_rsExceptionSched) Then
'''''            ApplyExceptions v_datStartDate, v_datEndDate, l_rsAdjustedBaseSched, m_rsExceptionSched, v_lngEmpNum
'''''        End If
'''''    End If
'''''
'''''    If v_ReturnOption >= trAdjBaseSchedWithExceptionsAndOvertime Then
'''''        m_rsOvertimeSched.Filter = "EMP_NUM = " & str(v_lngEmpNum)
'''''        If RecordsExist(m_rsOvertimeSched) Then
'''''            ApplyOvertimeExceptions v_datStartDate, v_datEndDate, l_rsAdjustedBaseSched, m_rsOvertimeSched
'''''        End If
'''''    End If
'''''    If v_ReturnOption > trBaseSchedNoDetails Then
'''''        ApplyEmployeeBaseJobNum v_lngEmpNum, v_datStartDate, v_datEndDate, l_rsAdjustedBaseSched
'''''    End If
'''''    If v_ReturnOption >= trAdjustedBaseSched Then
'''''        l_strSort = "CHARGE_DATE , DATE_WORKED, START_TIME, END_TIME"
'''''    Else
'''''        l_strSort = "CHARGE_DATE , SHIFT_CODE"
'''''    End If
'''''    l_rsAdjustedBaseSched.Sort = l_strSort
'''''
'''''    l_rsAdjustedBaseSched.MoveFirst
'''''
'''''
'''''    Set GetEmployeeSchedule = l_rsAdjustedBaseSched
'''''    Set l_rsAdjustedBaseSched = Nothing
'''''    Exit Function
'''''
'''''GetEmployeeSchedule_EH:
'''''    Set l_rsAdjustedBaseSched = Nothing
'''''
'''''    TrsRaiseError m_strCMODULENAME, "GetEmployeeSchedule"
'''''
'''''End Function



Private Function GetScheduleChangePayments( _
                    ByVal v_aryEmployees As Variant, _
                    ByVal v_datStartDate As String, _
                    ByVal v_datEndDate As String _
                    ) As ADODB.Recordset

    On Error GoTo GetScheduleChangePayments_EH
    
    
    Dim l_dbSchedule As dbSchedule.cdSchedule
         
    
    Set l_dbSchedule = CreateObject("dbSchedule.cdSchedule")
    
    Set GetScheduleChangePayments = l_dbSchedule.GetScheduleChangePayments(v_aryEmployees, v_datStartDate, v_datEndDate)
             
    Set l_dbSchedule = Nothing
    
    Exit Function

GetScheduleChangePayments_EH:
    Set l_dbSchedule = Nothing

    TrsRaiseError m_strCMODULENAME, "GetScheduleChangePayments"
End Function


Private Function GetWGBaseScheduleOnly( _
            ByVal v_aryWorkgroups As Variant, _
            ByVal v_datStartDate As Date, _
            ByVal v_datEndDate As Date, _
            Optional ByVal v_adLockType As ADODB.LockTypeEnum = adLockReadOnly _
        ) As ADODB.Recordset

On Error GoTo eh
    Dim l_dbSchedule As dbSchedule.cdSchedule
         
    Set l_dbSchedule = CreateObject("dbSchedule.cdSchedule")
    Set GetWGBaseScheduleOnly = _
        l_dbSchedule.GetWorkGroupBaseSchedules( _
                v_aryWorkgroups, v_datStartDate, v_datEndDate, v_adLockType)
    Set l_dbSchedule = Nothing
    
    Exit Function
    
eh:
    Set l_dbSchedule = Nothing
    TrsRaiseError m_strCMODULENAME, "GetWGBaseScheduleOnly"
End Function


Private Function GetPrimaryWorkgroupAssignments( _
            ByVal v_aryEmployees As Variant, _
            ByVal v_datStartDate As Date, _
            ByVal v_datEndDate As Date, _
            Optional ByVal v_adLockType As ADODB.LockTypeEnum = adLockReadOnly _
        ) As ADODB.Recordset

On Error GoTo eh
    Dim l_dbSchedule As dbSchedule.cdSchedule
         
    Set l_dbSchedule = CreateObject("dbSchedule.cdSchedule")
    Set GetPrimaryWorkgroupAssignments = _
        l_dbSchedule.GetPrimaryWorkgroupAssignments( _
                v_aryEmployees, v_datStartDate, v_datEndDate, v_adLockType)
    Set l_dbSchedule = Nothing
    
    Exit Function
    
eh:
    Set l_dbSchedule = Nothing
    TrsRaiseError m_strCMODULENAME, "GetPrimaryWorkgroupAssignments"
End Function


Private Function GetWGBaseSchedule( _
            ByVal v_aryEmployees As Variant, _
            ByVal v_datStartDate As Date, _
            ByVal v_datEndDate As Date, _
            Optional ByVal v_adLockType As ADODB.LockTypeEnum = adLockReadOnly _
        ) As ADODB.Recordset

    On Error GoTo GetWGBaseSchedule_EH
    
    
    Dim l_dbSchedule As dbSchedule.cdSchedule
         
    
    Set l_dbSchedule = CreateObject("dbSchedule.cdSchedule")
    
    Set GetWGBaseSchedule = _
        l_dbSchedule.GetEmployeePrimaryWGBaseSchedule( _
                v_aryEmployees, v_datStartDate, v_datEndDate, v_adLockType)
       
    Set l_dbSchedule = Nothing
    
    Exit Function
    
GetWGBaseSchedule_EH:
    Set l_dbSchedule = Nothing
    TrsRaiseError m_strCMODULENAME, "GetWGBaseSchedule"

End Function

Private Function GetExceptions( _
            ByVal v_aryEmployees As Variant, _
            ByVal v_datStartDate As Date, _
            ByVal v_datEndDate As Date _
            ) As ADODB.Recordset

    On Error GoTo GetExceptions_EH
         
    
    Dim l_dbSchedule As dbSchedule.cdSchedule
    Dim l_colSchedTypes As Collection
             
    Set l_dbSchedule = CreateObject("dbSchedule.cdSchedule")
    Set GetExceptions = l_dbSchedule.GetEmployeeScheduledExceptions(v_aryEmployees, v_datStartDate, v_datEndDate)
       
    Set l_dbSchedule = Nothing
    
    Exit Function
    
GetExceptions_EH:
    Set l_dbSchedule = Nothing
    TrsRaiseError m_strCMODULENAME, "GetExceptions"
     
End Function



Private Function BuildGetAdjBase( _
            ByVal v_rsSchedCodeByDate As ADODB.Recordset, _
            ByVal v_ReturnOption As TRSReturnOption _
            ) As ADODB.Recordset
            
            
'NEW CODE ________________________________

                                 
'This Function will build the Schedule details for each date
'If the return option is 1 (BaseSchedNoDetails),
'  then the shift code is the only detail added
'Otherwise this will split over lunch, midnight, overtime, and shifts
'The Shift component, exception code , shift start time, shift end time, ect.
' see CreateGetAdjBaseRecordset for the record layou and definitions

    
    Dim l_intIndex As Integer
    Dim l_rsGetAdjBaseSched As ADODB.Recordset
    Dim l_datChargeDate As Date
    Dim l_datDateWorked As Date
    Dim l_intSchedCd As Integer
    Dim l_strStartTime As String
    Dim l_strEndTime As String
    Dim l_strExceptionCd As String
    Dim l_strFilter As String
    Dim l_lngEmpNum As Long
    Dim l_lngPrevEmpNum As Long
    
On Error GoTo BuildGetAdjBase_EH

    If v_ReturnOption = trBaseSchedNoDetails Then
        Set l_rsGetAdjBaseSched = CreateBaseScheduleRecordset
    Else
        Set l_rsGetAdjBaseSched = CreateGetAdjBaseRecordset(v_ReturnOption)
    End If

' If no schedules are returned check to see if all schedule codes = -1
' If all schedules are not missing (sched code = -1) and no records
' retrieved from available schedule, then bad data schedule does not exist
' discountinue process
    If Not RecordsExist(m_rsAvailSched) Then
        If AllSchedulesAreMissing() Then
            BuildMissingSchedule l_rsGetAdjBaseSched, v_rsSchedCodeByDate
            Set BuildGetAdjBase = l_rsGetAdjBaseSched
            Exit Function
        Else
            Err.Raise trErrbusScheduleCodeDoesNotExist
        End If
    End If
    
    If Not RecordsExist(v_rsSchedCodeByDate) Then
        Err.Raise trErrbusScheduleCodeDoesNotExist
        Exit Function
    End If
        
    l_lngPrevEmpNum = -1
    v_rsSchedCodeByDate.MoveFirst
    While Not v_rsSchedCodeByDate.EOF
        l_rsGetAdjBaseSched.AddNew
        l_lngEmpNum = v_rsSchedCodeByDate!EMP_NUM
        l_datChargeDate = v_rsSchedCodeByDate!CHARGE_DATE
        l_datDateWorked = l_datChargeDate
        l_rsGetAdjBaseSched!EMP_NUM = l_lngEmpNum
        l_rsGetAdjBaseSched!CHARGE_DATE = l_datChargeDate
        l_rsGetAdjBaseSched!SCHED_SRC_CD = v_rsSchedCodeByDate!SCHED_SRC_CD
        l_rsGetAdjBaseSched!BASE_SCHED_SRC_CD = v_rsSchedCodeByDate!BASE_SCHED_SRC_CD
        l_rsGetAdjBaseSched!SCHED_CD = v_rsSchedCodeByDate!SCHED_CD
        If l_lngPrevEmpNum <> l_lngEmpNum Then
            m_strScheduleTimeStamp = ScheduleTimeStamp(l_lngEmpNum)
            l_lngPrevEmpNum = l_lngEmpNum
        End If
        
        If v_ReturnOption = trBaseSchedNoDetails Then
            l_rsGetAdjBaseSched!SHIFT_NUM = v_rsSchedCodeByDate!SHIFT_NUM
            l_rsGetAdjBaseSched!LINE_SRC_CD = v_rsSchedCodeByDate!LINE_SRC_CD
            l_rsGetAdjBaseSched!SCHED_CD_FOR_SHIFT_COMP = v_rsSchedCodeByDate!SCHED_CD_FOR_SHIFT_COMP
            l_rsGetAdjBaseSched!LAST_MOD_SCHED_TM = m_strScheduleTimeStamp
            l_rsGetAdjBaseSched!LINE_NUM = v_rsSchedCodeByDate!LINE_NUM
            l_rsGetAdjBaseSched!WG_NUM = v_rsSchedCodeByDate!WG_NUM
            l_rsGetAdjBaseSched!LINE_WG = v_rsSchedCodeByDate!LINE_WG
        Else
            l_rsGetAdjBaseSched!DATE_WORKED = l_datDateWorked
            LoadRepeatingShiftData v_rsSchedCodeByDate, l_rsGetAdjBaseSched
        End If
            l_strFilter = "SCHED_CD = " & str(v_rsSchedCodeByDate!SCHED_CD)
            m_rsAvailSched.Filter = l_strFilter
        ' if not found then Check to see if a schedule exists on that day
            If Not RecordsExist(m_rsAvailSched) Then
                If v_rsSchedCodeByDate!SCHED_CD = -1 Then
                Else
            ' BigTime Error - Schedule code does not exist
                    Err.Raise trErrbusScheduleCodeDoesNotExist
                End If
            Else
                SetShiftData l_rsGetAdjBaseSched, v_ReturnOption
            End If
        
        
        l_rsGetAdjBaseSched.Update
        
        v_rsSchedCodeByDate.MoveNext
        
    Wend

    
    Set BuildGetAdjBase = l_rsGetAdjBaseSched
    Set l_rsGetAdjBaseSched = Nothing
    
    Exit Function
    

BuildGetAdjBase_EH:
    Set l_rsGetAdjBaseSched = Nothing
    
    Select Case Err.Number
        Case trErrbusScheduleCodeDoesNotExist
            TrsRaiseError m_strCMODULENAME, "BuildGetAdjBase", trErrbusScheduleCodeDoesNotExist, , , , True
        Case Else
            TrsRaiseError m_strCMODULENAME, "BuildGetAdjBase", , , , , True
    End Select


End Function

Private Sub ApplyScheduleChangePayments( _
                ByVal v_aryEmployees As Variant, _
                ByVal v_datStartDate As Date, _
                ByVal v_datEndDate As Date, _
                ByRef r_rsAdjustedBaseSched As ADODB.Recordset, _
                ByRef r_rsABSWithNoExceptions As ADODB.Recordset)
                
                
    Dim l_rsSchedChangePMTs As ADODB.Recordset
    Dim l_datChargeDate As Date
    Dim l_lngEmpNum As Long
    Dim l_strFilter As String
    
    Set l_rsSchedChangePMTs = GetScheduleChangePayments(v_aryEmployees, v_datStartDate, v_datEndDate)
    
    If Not RecordsExist(l_rsSchedChangePMTs) Then
        Exit Sub
    End If
    
    l_rsSchedChangePMTs.MoveFirst
    While Not l_rsSchedChangePMTs.EOF
        l_datChargeDate = l_rsSchedChangePMTs!CHARGE_DATE
        l_lngEmpNum = l_rsSchedChangePMTs!EMP_NUM
        l_strFilter = "EMP_NUM = " & str(l_lngEmpNum) & _
                    " and CHARGE_DATE = " & str(l_datChargeDate)
        r_rsAdjustedBaseSched.Filter = MDAC25FilterFix(l_strFilter)
        If RecordsExist(r_rsAdjustedBaseSched) Then
            r_rsAdjustedBaseSched.MoveFirst
            While Not r_rsAdjustedBaseSched.EOF
                r_rsAdjustedBaseSched!SCHED_CHANGE_PMT = "Y"
                r_rsAdjustedBaseSched!SCHED_CHG_EXC_CD = l_rsSchedChangePMTs!EXCEPTION_CD
                r_rsAdjustedBaseSched.Update
                r_rsAdjustedBaseSched.MoveNext
            Wend
        End If
        If RecordsExist(r_rsABSWithNoExceptions) Then
            r_rsABSWithNoExceptions.Filter = MDAC25FilterFix(l_strFilter)
            If RecordsExist(r_rsABSWithNoExceptions) Then
                r_rsABSWithNoExceptions.MoveFirst
                While Not r_rsABSWithNoExceptions.EOF
                    r_rsABSWithNoExceptions!SCHED_CHANGE_PMT = "Y"
                    r_rsABSWithNoExceptions!SCHED_CHG_EXC_CD = l_rsSchedChangePMTs!EXCEPTION_CD
                    r_rsABSWithNoExceptions.Update
                    r_rsABSWithNoExceptions.MoveNext
                Wend
            End If
        End If
        l_rsSchedChangePMTs.MoveNext
    Wend
    r_rsAdjustedBaseSched.Filter = ""
    If RecordsExist(r_rsABSWithNoExceptions) Then
        r_rsABSWithNoExceptions.Filter = ""
    End If
    Set l_rsSchedChangePMTs = Nothing
                
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
    If v_strShiftStartTime <= v_strSiftEndTime Then
        r_datShiftEndDate = r_datShiftStartDate
    Else
        r_datShiftEndDate = v_datChargeDate
    End If

         
End Sub



Private Function GetRotationSchedules( _
            ) As ADODB.Recordset

            
    Dim l_intRotationCode As Integer
    Dim l_colRotationCodes As Collection
    
    On Error GoTo eh
    
    'Getting list of rotation codes
    Set l_colRotationCodes = New Collection
    If RecordsExist(m_rsPersonalSchedule) Then
        m_rsPersonalSchedule.MoveFirst
        While Not m_rsPersonalSchedule.EOF
            l_intRotationCode = m_rsPersonalSchedule!ROT_CD
            If l_intRotationCode > 0 Then
                If Not InRotationCodeList(l_intRotationCode, l_colRotationCodes) Then
                    l_colRotationCodes.Add l_intRotationCode
                End If
            End If
            m_rsPersonalSchedule.MoveNext
        Wend
    End If
    If RecordsExist(m_rsWGBaseSchedule) Then
        m_rsWGBaseSchedule.MoveFirst
        While Not m_rsWGBaseSchedule.EOF
            l_intRotationCode = m_rsWGBaseSchedule!ROT_CD
            If l_intRotationCode > 0 Then
                If Not InRotationCodeList(l_intRotationCode, l_colRotationCodes) Then
                    l_colRotationCodes.Add l_intRotationCode
                End If
            End If
            m_rsWGBaseSchedule.MoveNext
        Wend
    End If
    If RecordsExist(m_rsDemandLineBaseSchedule) Then
        m_rsDemandLineBaseSchedule.MoveFirst
        While Not m_rsDemandLineBaseSchedule.EOF
            l_intRotationCode = m_rsDemandLineBaseSchedule!ROT_CD
            If l_intRotationCode > 0 Then
                If Not InRotationCodeList(l_intRotationCode, l_colRotationCodes) Then
                    l_colRotationCodes.Add l_intRotationCode
                End If
            End If
            m_rsDemandLineBaseSchedule.MoveNext
        Wend
    End If
    If RecordsExist(m_rsWorkGroupLineBaseSched) Then
        m_rsWorkGroupLineBaseSched.MoveFirst
        While Not m_rsWorkGroupLineBaseSched.EOF
            l_intRotationCode = m_rsWorkGroupLineBaseSched!ROT_CD
            If l_intRotationCode > 0 Then
                If Not InRotationCodeList(l_intRotationCode, l_colRotationCodes) Then
                    l_colRotationCodes.Add l_intRotationCode
                End If
            End If
            m_rsWorkGroupLineBaseSched.MoveNext
        Wend
    End If

    Set GetRotationSchedules = GetRotationSchedulesDB(l_colRotationCodes)

    Exit Function
        
eh:

    TrsRaiseError m_strCMODULENAME, "GetRotationSchedules"
                   
            
End Function


Private Function GetRotationSchedulesDB( _
            ByVal l_colRotationCodes As Collection _
            ) As ADODB.Recordset
            
    Dim l_obj As dbSchedule.cdSchedule
    
    Set l_obj = CreateObject("dbSchedule.cdSchedule")
    Set GetRotationSchedulesDB = l_obj.GetRotationSchedules(l_colRotationCodes)
    
    Set l_obj = Nothing
    
    Exit Function
    
eh:

    TrsRaiseError m_strCMODULENAME, "GetRotationSchedulesDB"
       
End Function

Private Function InRotationCodeList( _
            ByVal v_intRotationCode, _
            ByVal v_colRotationCodes As Collection) As Boolean

    Dim l_intIndex As Integer
    Dim l_intCount As Integer
    Dim l_bolInList As Boolean
    
    On Error GoTo eh
    
    l_bolInList = False
    l_intIndex = 1
    l_intCount = v_colRotationCodes.Count
    
    While Not l_bolInList And l_intIndex <= l_intCount
        If v_colRotationCodes.Item(l_intIndex) = v_intRotationCode Then
            l_bolInList = True
        End If
        l_intIndex = l_intIndex + 1
    Wend
    
    InRotationCodeList = l_bolInList
    Exit Function
    
eh:

    TrsRaiseError m_strCMODULENAME, "InRotationCodeList"
    

End Function


Private Function DetermineRotationScheduleCode( _
            ByVal v_datChargeDate As Date, _
            ByVal v_intRotationCode As Integer, _
            ByVal v_datEffDate As Date, _
            ByVal v_intStartPosition As Integer, _
            ByRef r_intRecordPosition As Integer _
            ) As Integer
            
'This Function determines the schedule code of a rotation on a charge date
'The employee (or work group) starts on the rotation on an effective date (v_datEffDate).
'    This is the schedule_date for a Non Demand or Demand schedule assignment.
'    This is the eff_rotation_date for Work group base schedules
'The assignment to a rotation schedule allows an offset into the rotation (v_intStartPosition)
'    For example, there a 3 schedules in the rotation.
'       A rot_start_pos (offset) of 3 means that the employee
'       will start on the 3rd schedule of the rotation on the effective date.
'       The employee will then follow the rotation from there.
'       Do not get this confused with number of weeks.
'       The offset is on schedules in rotation only.
            
    Dim l_lngOffSetWeeks As Long
    Dim l_lngTotalWeeksOfRotation As Long
    Dim l_lngNumofWeeks As Long
    Dim l_lngRotationWeek As Long
    Dim l_blnFoundOffset As Boolean
    Dim l_blnSchedFound As Boolean
    
    
    On Error GoTo eh
    
    If Not RecordsExist(m_rsRotations) Then
        DetermineRotationScheduleCode = -1
        r_intRecordPosition = 0
        Exit Function
    End If
    
    m_rsRotations.Filter = "ROT_CD = " & str(v_intRotationCode)
    m_rsRotations.Sort = "ROT_WEEK"
    
    If Not RecordsExist(m_rsRotations) Then
        DetermineRotationScheduleCode = -1
        r_intRecordPosition = 0
        Exit Function
    End If
            
    'Calculate Offset into  rotation
    l_lngOffSetWeeks = v_intStartPosition - 1
    m_rsRotations.MoveFirst
    If v_datEffDate >= v_datChargeDate Then
        r_intRecordPosition = v_intStartPosition
        m_rsRotations.AbsolutePosition = v_intStartPosition
        DetermineRotationScheduleCode = m_rsRotations!SCHED_CD
        Exit Function
    End If
    
    'finding position in rotation
    l_lngNumofWeeks = DateDiff("ww", v_datEffDate, v_datChargeDate) _
                        + l_lngOffSetWeeks
    l_lngTotalWeeksOfRotation = 0
    
    'Add Total weeks of rotation
    m_rsRotations.MoveFirst
    While Not m_rsRotations.EOF
        l_lngTotalWeeksOfRotation = l_lngTotalWeeksOfRotation + 1
'''''        l_lngTotalWeeksOfRotation = l_lngTotalWeeksOfRotation + m_rsRotations!NBR_OF_WEEKS
        m_rsRotations.MoveNext
    Wend
    
    'Calulate number of weeks into the rotation
    l_lngRotationWeek = l_lngNumofWeeks Mod l_lngTotalWeeksOfRotation
    l_lngRotationWeek = l_lngRotationWeek + 1
    
    'Find Schedule for that week
    m_rsRotations.MoveFirst
    l_blnSchedFound = False
    While Not m_rsRotations.EOF And Not l_blnSchedFound
            l_lngRotationWeek = l_lngRotationWeek - 1
'''''        l_lngRotationWeek = l_lngRotationWeek - m_rsRotations!NBR_OF_WEEKS
        If l_lngRotationWeek <= 0 Then 'Rotation found
            l_blnSchedFound = True
        Else
            m_rsRotations.MoveNext
        End If
    Wend
    
    r_intRecordPosition = m_rsRotations.AbsolutePosition
    DetermineRotationScheduleCode = m_rsRotations!SCHED_CD

    Exit Function
    
eh:

    TrsRaiseError m_strCMODULENAME, "DetermineRotationScheduleCode"
    
End Function

Private Sub ClearRepeatingShiftData()

    Dim l_intI As Integer
    
    For l_intI = 0 To 26
        m_varRepeatingShiftData(l_intI) = ""
    Next l_intI

End Sub

Private Sub InitializeRepeatingShiftData( _
        ByVal v_lngEmpNum As Long, _
        ByVal v_datChargeDate As Date, _
        ByVal v_intSchedCode As Integer, _
        ByVal v_intSchedSource As Integer, _
        ByVal v_strShiftStartTime As String, _
        ByVal v_strShiftEndTime As String, _
        ByVal v_datShiftStartDate As Date, _
        ByVal v_datShiftEndDate As Date, _
        ByVal v_decShiftHours As Variant, _
        ByVal v_strShiftCode As String, _
        ByVal v_strJobNum As String, _
        ByVal v_intLineNum As String, _
        ByVal v_lngWgNum As Long, _
        ByVal v_strReqInd As String, _
        ByVal v_strShiftComponent As String, _
        ByVal v_strReliefComponent As String, _
        ByVal v_intSchedCodeForShiftComp As Integer, _
        ByVal v_intShiftNum As Integer, _
        ByVal v_intSchedLineSource As Integer, _
        ByVal v_strShiftLunchStartTime As String, ByVal v_strShiftLunchEndTime As String, _
        ByVal v_lngTradeNum As Long, ByVal v_intBaseSchedSource As Integer, _
        ByVal v_intRotCode As Integer, ByVal v_intRotWeek As Integer, _
        ByVal v_strSysShiftComponent As String)

    ClearRepeatingShiftData

    m_varRepeatingShiftData(0) = v_lngEmpNum
    m_varRepeatingShiftData(1) = v_datChargeDate
    m_varRepeatingShiftData(2) = v_intSchedCode
    m_varRepeatingShiftData(3) = v_intSchedSource
    m_varRepeatingShiftData(4) = v_strShiftStartTime
    m_varRepeatingShiftData(5) = v_strShiftEndTime
    m_varRepeatingShiftData(6) = v_datShiftStartDate
    m_varRepeatingShiftData(7) = v_datShiftEndDate
    m_varRepeatingShiftData(8) = v_decShiftHours
    m_varRepeatingShiftData(9) = v_strShiftCode
    m_varRepeatingShiftData(10) = v_strJobNum
    m_varRepeatingShiftData(11) = v_intLineNum
    m_varRepeatingShiftData(12) = v_lngWgNum
    m_varRepeatingShiftData(13) = v_strReqInd
    m_varRepeatingShiftData(14) = v_strShiftComponent
    m_varRepeatingShiftData(15) = v_strReliefComponent
    m_varRepeatingShiftData(16) = v_intSchedCodeForShiftComp
    m_varRepeatingShiftData(17) = v_intShiftNum
    m_varRepeatingShiftData(18) = v_intSchedLineSource
    m_varRepeatingShiftData(19) = v_strShiftLunchStartTime
    m_varRepeatingShiftData(20) = v_strShiftLunchEndTime
    m_varRepeatingShiftData(21) = v_lngTradeNum
    m_varRepeatingShiftData(22) = v_intBaseSchedSource
'Profiler
    m_varRepeatingShiftData(23) = GetEffectiveJob(v_lngEmpNum, v_datChargeDate)
    m_varRepeatingShiftData(24) = v_intRotCode
    m_varRepeatingShiftData(25) = v_intRotWeek
    m_varRepeatingShiftData(26) = v_strSysShiftComponent

End Sub

Private Sub SetRepeatingShiftData( _
        ByRef r_rsGetAdjBaseSched As ADODB.Recordset)

    InitializeExceptionVariables r_rsGetAdjBaseSched
    r_rsGetAdjBaseSched!EMP_NUM = m_varRepeatingShiftData(0)
    r_rsGetAdjBaseSched!CHARGE_DATE = m_varRepeatingShiftData(1)
    r_rsGetAdjBaseSched!SCHED_CD = m_varRepeatingShiftData(2)
    r_rsGetAdjBaseSched!SCHED_SRC_CD = m_varRepeatingShiftData(3)
    r_rsGetAdjBaseSched!SHIFT_START_TIME = m_varRepeatingShiftData(4)
    r_rsGetAdjBaseSched!SHIFT_END_TIME = m_varRepeatingShiftData(5)
    r_rsGetAdjBaseSched!SHIFT_START_DATE = m_varRepeatingShiftData(6)
    r_rsGetAdjBaseSched!SHIFT_END_DATE = m_varRepeatingShiftData(7)
    r_rsGetAdjBaseSched!SHIFT_HOURS = Round(m_varRepeatingShiftData(8), 1)
    r_rsGetAdjBaseSched!SHIFT_CODE = m_varRepeatingShiftData(9)
    r_rsGetAdjBaseSched!SCHED_JOB_NUM = m_varRepeatingShiftData(10)
    r_rsGetAdjBaseSched!LINE_NUM = m_varRepeatingShiftData(11)
    r_rsGetAdjBaseSched!WG_NUM = m_varRepeatingShiftData(12)
    r_rsGetAdjBaseSched!REQ_LINE_IND = m_varRepeatingShiftData(13)
    r_rsGetAdjBaseSched!COMP_SHIFT_CD = m_varRepeatingShiftData(14)
    r_rsGetAdjBaseSched!REL_CD = m_varRepeatingShiftData(15)
    r_rsGetAdjBaseSched!SHIFT_START_END = m_varRepeatingShiftData(4) & "-" & m_varRepeatingShiftData(5)
    r_rsGetAdjBaseSched!SCHED_CD_FOR_SHIFT_COMP = m_varRepeatingShiftData(16)
    r_rsGetAdjBaseSched!SHIFT_NUM = m_varRepeatingShiftData(17)
    r_rsGetAdjBaseSched!LINE_SRC_CD = m_varRepeatingShiftData(18)
    r_rsGetAdjBaseSched!SHIFT_LUNCH_START_TIME = m_varRepeatingShiftData(19)
    r_rsGetAdjBaseSched!SHIFT_LUNCH_END_TIME = m_varRepeatingShiftData(20)
    r_rsGetAdjBaseSched!TRADE_NUM = m_varRepeatingShiftData(21)
    r_rsGetAdjBaseSched!BASE_SCHED_SRC_CD = m_varRepeatingShiftData(22)
    r_rsGetAdjBaseSched!SCHED_CHANGE_PMT = "N"
    r_rsGetAdjBaseSched!WG_LINE_IND = "N"
    r_rsGetAdjBaseSched!FULL_SHIFT_EXC_IND = "N"
    r_rsGetAdjBaseSched!DOUBLE_SHIFT_IND = "N"
'profiler
    r_rsGetAdjBaseSched!EMP_JOB_NUM = m_varRepeatingShiftData(23)
'Optimistic locking
    r_rsGetAdjBaseSched!LAST_MOD_SCHED_TM = m_strScheduleTimeStamp
'View Schedule wants the work group of the line
    r_rsGetAdjBaseSched!LINE_WG = m_varRepeatingShiftData(12)
    r_rsGetAdjBaseSched!LINE_JOB_NUM = m_varRepeatingShiftData(10)
    
    r_rsGetAdjBaseSched!SCHED_CHG_EXC_CD = ""
    r_rsGetAdjBaseSched!ROT_CD = m_varRepeatingShiftData(24)
    r_rsGetAdjBaseSched!ROT_WEEK = m_varRepeatingShiftData(25)
    r_rsGetAdjBaseSched!SYS_COMP_SHIFT_CD = m_varRepeatingShiftData(26)

End Sub


Public Function GetEffectiveWorkGroupBaseSchedule( _
            ByVal v_lngWorkGroup As Long, _
            ByVal v_datDate As Date _
            ) As ADODB.Recordset
            
    Dim l_colWg As Collection
    Dim l_rsWorkGroup As ADODB.Recordset
    
    On Error GoTo GetEffectiveWorkGroupBaseSchedule_EH
    
    Set l_colWg = New Collection
    l_colWg.Add v_lngWorkGroup
    Set l_rsWorkGroup = GetWorkGroupAttributes(v_lngWorkGroup)
    If UCase(l_rsWorkGroup!DEMAND_IND) = "Y" Then
        Set GetEffectiveWorkGroupBaseSchedule = GetWGDemandLineBaseSchedules(l_colWg, v_datDate)
    Else
        Set GetEffectiveWorkGroupBaseSchedule = GetWorkGroupBaseSchedules(l_colWg, v_datDate)
    End If
    Set l_colWg = Nothing
    Exit Function
    
GetEffectiveWorkGroupBaseSchedule_EH:
    Set l_colWg = Nothing
    TrsRaiseError m_strCMODULENAME, "GetEffectiveWorkGroupBaseSchedule", , , , , True
         
End Function

Private Function GetWorkGroupBaseSchedules( _
            ByVal v_colWorkGroups As Collection, _
            ByVal v_datDate As Date _
            ) As ADODB.Recordset
    
    Dim l_dbSchedule As dbSchedule.cdSchedule
    Dim v_datChargeDate As Date
    Dim v_datEndDate As Date

    On Error GoTo GetWorkGroupBaseSchedules_EH
        
    Set l_dbSchedule = CreateObject("dbSchedule.cdSchedule")
    Set GetWorkGroupBaseSchedules = l_dbSchedule.GetWGBaseSchedules(v_colWorkGroups, v_datDate)
    Set l_dbSchedule = Nothing
    Exit Function
    
GetWorkGroupBaseSchedules_EH:
    Set l_dbSchedule = Nothing
    TrsRaiseError m_strCMODULENAME, "GetWorkGroupBaseSchedules"

End Function


Private Function GetPersonalSchedules( _
            ByVal v_aryEmployees As Variant, _
            ByVal v_datStartDate As String, _
            ByVal v_datEndDate As String, _
            Optional ByVal v_LockType As ADODB.LockTypeEnum = adLockReadOnly _
            ) As ADODB.Recordset

    On Error GoTo GetPersonalSchedules_EH


    Dim l_dbSchedule As dbSchedule.cdSchedule
         
    
    Set l_dbSchedule = CreateObject("dbSchedule.cdSchedule")
    
    Set GetPersonalSchedules = _
        l_dbSchedule.GetEmployeePersonalSchedule( _
                v_aryEmployees, v_datStartDate, v_datEndDate, v_LockType)
         
    Set l_dbSchedule = Nothing
    
    Exit Function
    
GetPersonalSchedules_EH:
    Set l_dbSchedule = Nothing
    TrsRaiseError m_strCMODULENAME, "GetPersonalSchedules"
End Function

Private Function GetDemandLineAssignments( _
            ByVal v_aryEmployees As Variant, _
            ByVal v_datStartDate As String, _
            ByVal v_datEndDate As String, _
            Optional ByVal v_adLockType As ADODB.LockTypeEnum = adLockReadOnly _
            ) As ADODB.Recordset

    On Error GoTo GetDemandLineAssignments_EH

    Dim l_dbSchedule As dbSchedule.cdSchedule


    Set l_dbSchedule = CreateObject("dbSchedule.cdSchedule")

    Set GetDemandLineAssignments = _
        l_dbSchedule.GetEmployeeLineAssignments(v_aryEmployees, v_datStartDate, v_datEndDate, v_adLockType)

    Set l_dbSchedule = Nothing

    Exit Function

GetDemandLineAssignments_EH:
    Set l_dbSchedule = Nothing
    TrsRaiseError m_strCMODULENAME, "GetDemandLineAssignments"

End Function

Private Function GetDemandLineBaseSchedules( _
            ByVal v_datStartDate As String, _
            ByVal v_datEndDate As String, _
            Optional ByVal v_adLockType As ADODB.LockTypeEnum = adLockReadOnly, _
            Optional ByVal v_aryPendingWorkgroups As Variant, _
            Optional ByVal v_aryPendingLines As Variant _
            ) As ADODB.Recordset

    Dim l_colWorkGroups As cuSet
    Dim l_colLines As cuSet
    Dim l_aryWorkgroups As Variant
    Dim l_aryLines As Variant
    Dim i As Integer

    On Error GoTo GetDemandLineBaseSchedules_EH

    Set l_colWorkGroups = New cuSet
    Set l_colLines = New cuSet
    If RecordsExist(m_rsDemandLineAssignments) Then
        m_rsDemandLineAssignments.MoveFirst
        While Not m_rsDemandLineAssignments.EOF
            l_colWorkGroups.Add CLng(m_rsDemandLineAssignments!WG_NUM)
            l_colLines.Add CLng(m_rsDemandLineAssignments!LINE_NUM)
            m_rsDemandLineAssignments.MoveNext
        Wend
    Else
        l_colWorkGroups.Add -1
        l_colLines.Add -1
    End If
    
    If Not (IsMissing(v_aryPendingWorkgroups) Or IsEmpty(v_aryPendingWorkgroups)) Then
        If IsArray(v_aryPendingWorkgroups) Then
            For i = LBound(v_aryPendingWorkgroups) To UBound(v_aryPendingWorkgroups)
                l_colWorkGroups.Add CLng(v_aryPendingWorkgroups(i))
            Next
        Else
            l_colWorkGroups.Add CLng(v_aryPendingWorkgroups)
        End If
    End If
    
    If Not (IsMissing(v_aryPendingLines) Or IsEmpty(v_aryPendingLines)) Then
        If IsArray(v_aryPendingLines) Then
            For i = LBound(v_aryPendingLines) To UBound(v_aryPendingLines)
                l_colLines.Add CLng(v_aryPendingLines(i))
            Next
        Else
            l_colLines.Add CLng(v_aryPendingLines)
        End If
    End If
    
    l_aryWorkgroups = ArrayFromCollection(l_colWorkGroups.Contents)
    l_aryLines = ArrayFromCollection(l_colLines.Contents)

    Set GetDemandLineBaseSchedules = _
        GetDemandLineBaseScheduleDB(l_aryWorkgroups, l_aryLines, v_datStartDate, v_datEndDate, v_adLockType)

    Set l_colWorkGroups = Nothing
    Set l_colLines = Nothing

    Exit Function

GetDemandLineBaseSchedules_EH:

    Set l_colWorkGroups = Nothing
    Set l_colLines = Nothing
    TrsRaiseError m_strCMODULENAME, "GetDemandLineBaseSchedules"

End Function

Private Function GetDemandLineBaseScheduleDB( _
            ByVal v_aryWorkgroups As Variant, _
            ByVal v_aryLines As Variant, _
            ByVal v_datStartDate As String, _
            ByVal v_datEndDate As String, _
            Optional ByVal v_adLockType As ADODB.LockTypeEnum = adLockReadOnly _
            ) As ADODB.Recordset

    On Error GoTo GetDemandLineBaseScheduleDB_EH

    Dim l_dbSchedule As dbSchedule.cdSchedule


    Set l_dbSchedule = CreateObject("dbSchedule.cdSchedule")

    Set GetDemandLineBaseScheduleDB = _
        l_dbSchedule.GetDemandLineSchedules(v_aryWorkgroups, v_aryLines, v_datStartDate, v_datEndDate, v_adLockType)

    Set l_dbSchedule = Nothing

    Exit Function

GetDemandLineBaseScheduleDB_EH:
    Set l_dbSchedule = Nothing
    TrsRaiseError m_strCMODULENAME, "GetDemandLineBaseScheduleDB"

End Function

Private Function GetDemandLineTrades( _
            ByVal v_datStartDate As String, _
            ByVal v_datEndDate As String, _
            Optional ByVal v_adLockType As ADODB.LockTypeEnum = adLockReadOnly _
            ) As ADODB.Recordset

    Dim l_colWorkGroups As Collection
    Dim l_colLines As Collection
    Dim l_aryWorkgroups As Variant
    Dim l_aryLines As Variant
    Dim l_lngWorkGroup As Long
    Dim l_lngLine As Long

    On Error GoTo GetDemandLineTrades_EH

    Set l_colWorkGroups = New Collection
    Set l_colLines = New Collection
    If RecordsExist(m_rsDemandLineAssignments) Then
        m_rsDemandLineAssignments.MoveFirst
        While Not m_rsDemandLineAssignments.EOF
            l_lngWorkGroup = m_rsDemandLineAssignments!WG_NUM
            l_lngLine = m_rsDemandLineAssignments!LINE_NUM
            If IsInList(l_lngWorkGroup, l_colWorkGroups) Then
            Else
               l_colWorkGroups.Add l_lngWorkGroup
            End If
            If IsInList(l_lngLine, l_colLines) Then
            Else
               l_colLines.Add l_lngLine
            End If
            m_rsDemandLineAssignments.MoveNext
        Wend
    Else
        l_colWorkGroups.Add -1
        l_colLines.Add -1
    End If
    
    l_aryWorkgroups = ArrayFromCollection(l_colWorkGroups)
    l_aryLines = ArrayFromCollection(l_colLines)
    

    Set GetDemandLineTrades = _
        GetDemandLineTradesDB(l_aryWorkgroups, l_aryLines, v_datStartDate, v_datEndDate, v_adLockType)

    Exit Function

GetDemandLineTrades_EH:

    TrsRaiseError m_strCMODULENAME, "GetDemandLineTrades"

End Function

Private Function GetDemandLineTradesDB( _
            ByVal v_aryWorkgroups As Variant, _
            ByVal v_aryLines As Variant, _
            ByVal v_datStartDate As String, _
            ByVal v_datEndDate As String, _
            Optional ByVal v_adLockType As ADODB.LockTypeEnum = adLockReadOnly _
            ) As ADODB.Recordset

    On Error GoTo GetDemandLineTradesDB_EH

    Dim l_dbSchedule As dbSchedule.cdSchedule


    Set l_dbSchedule = CreateObject("dbSchedule.cdSchedule")

    Set GetDemandLineTradesDB = _
        l_dbSchedule.GetDemandLineTrades( _
            v_aryWorkgroups, v_aryLines, v_datStartDate, v_datEndDate, v_adLockType)

    Set l_dbSchedule = Nothing

    Exit Function

GetDemandLineTradesDB_EH:
    Set l_dbSchedule = Nothing
    TrsRaiseError m_strCMODULENAME, "GetDemandLineTradesDB"

End Function


Private Function GetEmployeeTrades( _
            ByVal v_aryEmployees As Variant, _
            ByVal v_datStartDate As String, _
            ByVal v_datEndDate As String, _
            Optional ByVal v_adLockType As ADODB.LockTypeEnum = adLockReadOnly _
            ) As ADODB.Recordset

    On Error GoTo GetEmployeeTrades_EH


    Dim l_dbSchedule As dbSchedule.cdSchedule
         
    
    Set l_dbSchedule = CreateObject("dbSchedule.cdSchedule")
    
    Set GetEmployeeTrades = _
        l_dbSchedule.GetEmployeeTrades( _
            v_aryEmployees, v_datStartDate, v_datEndDate, v_adLockType)
         
    Set l_dbSchedule = Nothing
    
    Exit Function
    
GetEmployeeTrades_EH:
    Set l_dbSchedule = Nothing
    TrsRaiseError m_strCMODULENAME, "GetEmployeeTrades"


End Function


Private Sub GetEmployeeScheduleData( _
    ByVal v_aryEmployees As Variant, _
    ByVal v_datStartDate As Date, _
    ByVal v_datEndDate As Date, _
    Optional ByVal v_ReturnOption As TRSReturnOption = trAdjustedBaseSched, _
    Optional ByVal v_ScheduleSourceStartLevel As TRSSchedSourceStartLevel = trAllScheduleLevels, _
    Optional ByVal v_adLockType As ADODB.LockTypeEnum = adLockReadOnly, _
    Optional ByVal v_blnWGAssignPendingChanges As Boolean = False, _
    Optional ByVal v_aryPendingWorkgroups As Variant, _
    Optional ByVal v_aryPendingLines As Variant)
    
 'This method will only get the data needed to satisfy the request
'    5/6/2002 RSG added v_adLockType.  This will allow the schedule recordsets to be
'    updated before gab processing.  This is needed to avoid uncommitted read for
'    cascade processing.

    Dim v_aryWorkgroups As Variant

    On Error GoTo eh
    
    'Recordset below needed from optimistic locking
    Set m_rsSchedTimeStamps = GetScheduleTimeStamps(v_aryEmployees)
    
    ' Profiler
    If v_ReturnOption > trBaseSchedNoDetails Then
        Set m_rsEmployeeData = GetEmployeeData(v_aryEmployees)
'        Set m_rsEmpJobNumByChargeDate = GetEmpJobNumByChargeDate(v_aryEmployees, v_datStartDate, v_datEndDate)
    Else
        Set m_rsEmployeeData = Nothing
    End If
    If v_ScheduleSourceStartLevel >= trNonDemandLevel Then
        Set m_rsPersonalSchedule = GetPersonalSchedules(v_aryEmployees, v_datStartDate, v_datEndDate, v_adLockType)
    End If
    If v_ScheduleSourceStartLevel >= trDemandLineBaseSchedLevel Then
        Set m_rsDemandLineAssignments = GetDemandLineAssignments(v_aryEmployees, v_datStartDate, v_datEndDate, v_adLockType)
        Set m_rsDemandLineBaseSchedule = GetDemandLineBaseSchedules(v_datStartDate, v_datEndDate, v_adLockType, v_aryPendingWorkgroups, v_aryPendingLines)
    End If
    
    'If WGBaseSchedule or WGAssignments pending changes are passed, must use separate
    'data service methods; GetWorkGroupBaseSchedules, GetPrimaryWorkGroupAssignments.
    'If m_rsWGEmpAssignments is passed, use it instead of GetPrimaryWorkGroupAssignments.
    If v_blnWGAssignPendingChanges Then
        If m_rsWGEmpAssignments Is Nothing Then
            Set m_rsWGEmpAssignments = GetPrimaryWorkgroupAssignments( _
                    v_aryEmployees, v_datStartDate, v_datEndDate, v_adLockType)
            v_aryWorkgroups = DetermineAllWorkgroupsInScope(v_aryPendingWorkgroups, m_rsWGEmpAssignments)
        Else
            v_aryWorkgroups = v_aryPendingWorkgroups
        End If
        Set m_rsWGBaseScheduleOnly = GetWGBaseScheduleOnly(v_aryWorkgroups, v_datStartDate, v_datEndDate, v_adLockType)
    Else
        Set m_rsWGBaseSchedule = GetWGBaseSchedule(v_aryEmployees, v_datStartDate, v_datEndDate, v_adLockType)
    End If
    
    If v_ScheduleSourceStartLevel >= trDemandLineTrades Then
        Set m_rsDemandLineTrades = GetDemandLineTrades(v_datStartDate, v_datEndDate, v_adLockType)
    End If
    If v_ScheduleSourceStartLevel >= trEmployeeTradesLevel Then
        Set m_rsEmployeeTrades = _
            GetEmployeeTrades(v_aryEmployees, v_datStartDate, v_datEndDate, v_adLockType)
    End If
    
    If v_ReturnOption >= trAdjBaseSchedWithExceptions Then
        Set m_rsExceptionSched = GetExceptions(v_aryEmployees, v_datStartDate, v_datEndDate)
    End If
    If v_ReturnOption >= trAdjBaseSchedWithExceptionsAndOvertime Then
        Set m_rsOvertimeSched = GetOvertime(v_aryEmployees, v_datStartDate, v_datEndDate)
    End If
    
    
    Exit Sub
    
eh:
    Set m_rsDemandLineAssignments = Nothing
    Set m_rsDemandLineBaseSchedule = Nothing
    Set m_rsDemandLineTrades = Nothing
    Set m_rsEmployeeTrades = Nothing
    Set m_rsPersonalSchedule = Nothing
    Set m_rsWGBaseSchedule = Nothing
    Set m_rsExceptionSched = Nothing
    Set m_rsEmpJobNumByChargeDate = Nothing
    Set m_rsAvailSched = Nothing
    
    TrsRaiseError m_strCMODULENAME, "GetEmployeeScheduleData"

    
End Sub


Private Sub ClearGlobalRecordSets()

        Set m_rsRotations = Nothing
        Set m_rsDemandLineAssignments = Nothing
        Set m_rsDemandLineBaseSchedule = Nothing
        Set m_rsDemandLineTrades = Nothing
        Set m_rsEmployeeTrades = Nothing
        Set m_rsPersonalSchedule = Nothing
        Set m_rsWGBaseSchedule = Nothing
        Set m_rsExceptionSched = Nothing
        Set m_rsEmpJobNumByChargeDate = Nothing
        Set m_rsAvailSched = Nothing
        Set m_colSchedCodes = Nothing
        Set m_rsWorkGroupLineAssignments = Nothing
        Set m_rsWorkGroupLineBaseSched = Nothing
        Set m_rsWorkGroupLineTrade = Nothing
        Set m_rsWGEmpAssignments = Nothing
        Set m_rsWGBaseScheduleOnly = Nothing


End Sub


Private Sub ProcessTrades( _
            ByVal v_lngEmpNum As Long, _
            ByVal v_datChargeDate As Date, _
            ByRef r_rsAdjBaseSchedbyChargeDate As ADODB.Recordset _
            )
            
    'This function returns true if a trade exist for the
    ' employee's 1st shift (normal schedule)
    'If found the employee's schedule on the charge date is the trade
    'An employee may have 2 shifts on a day when trading for a double
    'A double is denoted as the 2nd shift ( sometimes known as additional shift)
    'The function returns false if no trades are found
    ' or if the only trade is a double
    
    
    Dim l_strFilter As String
    
    On Error GoTo eh
    
    l_strFilter = "EMP_NUM = " & str(v_lngEmpNum) & _
            " AND CHARGE_DATE = " & v_datChargeDate
    l_strFilter = MDAC25FilterFix(l_strFilter)
    m_rsEmployeeTrades.Filter = l_strFilter
    
    If RecordsExist(m_rsEmployeeTrades) Then
        m_rsEmployeeTrades.Sort = "SHIFT_NUM"
        m_rsEmployeeTrades.MoveFirst
    Else
'Trade or Trade Double does not exist on charge date
        Exit Sub
    End If
    
    'Trade or Trade Double exists on charge date
    While Not m_rsEmployeeTrades.EOF
        If m_rsEmployeeTrades!SHIFT_NUM = 2 Then 'Double Shift Schedule
            AddTradedDoubleSchedule v_lngEmpNum, v_datChargeDate, r_rsAdjBaseSchedbyChargeDate
        Else
            r_rsAdjBaseSchedbyChargeDate!SCHED_CD = m_rsEmployeeTrades!SCHED_CD
            r_rsAdjBaseSchedbyChargeDate!SCHED_SRC_CD = 30
            r_rsAdjBaseSchedbyChargeDate!TRADE_NUM = m_rsEmployeeTrades!TRADE_NUM
            r_rsAdjBaseSchedbyChargeDate!SYSTEM_EXC_CD = _
                m_rsEmployeeTrades!ROLE_TYPE_CD & m_rsEmployeeTrades!WEEK_TRADE_IND
            r_rsAdjBaseSchedbyChargeDate!ROT_CD = 0
            r_rsAdjBaseSchedbyChargeDate!ROT_WEEK = 0
            r_rsAdjBaseSchedbyChargeDate.Update
        End If
        m_rsEmployeeTrades.MoveNext
    Wend
    
    Exit Sub
                
eh:
    Dim l_strErrorMessage As String
    
    l_strErrorMessage = "An error has occured while process employee trades.  " & _
            " It was done for employee (" & str(v_lngEmpNum) & _
            ") on charge date " & str(v_datChargeDate) & "."
    

    TrsRaiseError m_strCMODULENAME, "ProcessTrades", , , , DescArgsFrom(l_strErrorMessage), True
    
            
End Sub


Private Function IsEmployeeAssigenedToALine( _
            ByVal v_lngEmpNum As Long, _
            ByVal v_datDate As Date _
            ) As Boolean
            
    Dim l_strFilter As String
    Dim l_blnAssigmentFound As Boolean
    Dim l_datStart As Date
    Dim l_datEnd As Date
    
    
    On Error GoTo eh
    
    m_rsDemandLineAssignments.Filter = ""
    If RecordsExist(m_rsDemandLineAssignments) Then
    Else
        IsEmployeeAssigenedToALine = False
        Exit Function
    End If
    
    l_strFilter = "EMP_NUM = " & str(v_lngEmpNum)
    m_rsDemandLineAssignments.Filter = l_strFilter
    If Not RecordsExist(m_rsDemandLineAssignments) Then
        IsEmployeeAssigenedToALine = False
        Exit Function
    End If
    m_rsDemandLineAssignments.Sort = "START_DATE DESC"
    m_rsDemandLineAssignments.MoveFirst

    l_blnAssigmentFound = False
    
    While Not m_rsDemandLineAssignments.EOF And Not l_blnAssigmentFound
        l_datStart = m_rsDemandLineAssignments!START_DATE
        l_datEnd = m_rsDemandLineAssignments!END_DATE
        If v_datDate >= l_datStart _
            And v_datDate <= l_datEnd Then
            l_blnAssigmentFound = True
        Else
            m_rsDemandLineAssignments.MoveNext
        End If
    Wend

    IsEmployeeAssigenedToALine = l_blnAssigmentFound
    Exit Function

eh:

    TrsRaiseError m_strCMODULENAME, "IsEmployeeAssigenedToALine"
            
End Function


Private Sub ProcessDemandLineSchedule( _
            ByVal v_lngEmpNum As Long, _
            ByVal v_datDate As Date, _
            ByRef r_rsAdjBaseSchedbyChargeDate As ADODB.Recordset, _
            ByVal v_intScheduleSourceStartLevel As Integer _
            )
                            
    Dim l_lngLineWG As Long
    Dim l_lngLineNum As Long
    Dim l_strReliefComponent As String


    On Error GoTo eh
    
    'Getting information of the Line Assignment
    l_lngLineWG = m_rsDemandLineAssignments!WG_NUM
    l_lngLineNum = m_rsDemandLineAssignments!LINE_NUM
    l_strReliefComponent = m_rsDemandLineAssignments!RELIEF_COMP
    
    
    DetermineLineSchedule v_lngEmpNum, l_strReliefComponent, l_lngLineWG, l_lngLineNum, _
        v_datDate, m_rsDemandLineBaseSchedule, m_rsDemandLineTrades, _
        r_rsAdjBaseSchedbyChargeDate, v_intScheduleSourceStartLevel

    Exit Sub
    
eh:

    Dim l_strErrorMessage As String
    Dim l_strAssignMsg As String
    l_strErrorMessage = "An error occured while getting the demand " & _
            "schedule for employee (" & str(v_lngEmpNum) & ") " & _
            "on the date " & str(v_datDate) & "."
            
    l_strAssignMsg = "Demand Line Assignments. "

    
    l_strErrorMessage = l_strErrorMessage & _
            "The employee was assigned via " & l_strAssignMsg & _
            "The employee was on line " & str(l_lngLineWG) & _
            "  in work group number " & str(l_lngLineNum) & "."
            
    TrsRaiseError m_strCMODULENAME, "ProcessDemandLineSchedule", , , , DescArgsFrom(l_strErrorMessage), True
            
            
End Sub

Private Function ProcessPersonalSchedule( _
            ByVal v_lngEmpNum As Long, _
            ByVal v_datDate As Date, _
            ByRef r_rsAdjBaseSchedbyChargeDate As ADODB.Recordset _
            ) As Boolean
                            
    Dim l_strFilter As String
    Dim l_blnSchedFound As Boolean
    Dim l_intPersonalSchedCode As Integer
    Dim l_intRotationCode As Integer
    Dim l_intRotationStartWeek As Integer
    Dim l_datScheduleEffectiveDate As Date
    Dim l_intRotationWeek As Integer

    On Error GoTo eh
       
    'Getting Personal Schedule
    l_strFilter = _
        "EMP_NUM = " & str(v_lngEmpNum) & _
        " AND START_DATE <= " & str(v_datDate) & _
        " AND END_DATE >= " & str(v_datDate)
    l_strFilter = MDAC25FilterFix(l_strFilter)
    m_rsPersonalSchedule.Filter = l_strFilter
    If RecordsExist(m_rsPersonalSchedule) Then
        l_blnSchedFound = True
        l_datScheduleEffectiveDate = m_rsPersonalSchedule!START_DATE
        l_intRotationCode = m_rsPersonalSchedule!ROT_CD
        l_intRotationStartWeek = m_rsPersonalSchedule!ROT_WEEK
        If l_intRotationCode > 0 Then
            l_intPersonalSchedCode = DetermineRotationScheduleCode( _
                        v_datDate, l_intRotationCode, _
                        l_datScheduleEffectiveDate, l_intRotationStartWeek, l_intRotationWeek)
        Else
            l_intPersonalSchedCode = m_rsPersonalSchedule!SCHED_CD
            l_intRotationCode = 0
            l_intRotationWeek = 0
        End If
            AddPersonalOrWGSched _
            v_lngEmpNum, v_datDate, l_intPersonalSchedCode _
            , l_intRotationCode, l_intRotationWeek _
            , False, r_rsAdjBaseSchedbyChargeDate
    Else
        l_blnSchedFound = False
    End If
    
    ProcessPersonalSchedule = l_blnSchedFound

    Exit Function
    
eh:

    Dim l_strErrorMessage As String

    l_strErrorMessage = "An error occured while getting the personal " & _
            "schedule for employee (" & str(v_lngEmpNum) & ") " & _
            "on the date " & str(v_datDate) & "."
            
    TrsRaiseError m_strCMODULENAME, "ProcessPersonalSchedule", , , , DescArgsFrom(l_strErrorMessage), True
            
            
End Function

Private Function ProcessWorkGroupSchedule( _
            ByVal v_lngEmpNum As Long, _
            ByVal v_datDate As Date, _
            ByRef r_rsAdjBaseSchedbyChargeDate As ADODB.Recordset _
            ) As Boolean
                            
    Dim l_strFilter As String
    Dim l_blnSchedFound As Boolean
    Dim l_intWGSchedCode As Integer
    Dim l_intRotationCode As Integer
    Dim l_intRotationStartWeek As Integer
    Dim l_datScheduleEffectiveDate As Date
    Dim l_intRotationWeek As Integer

    On Error GoTo eh
       
    'Getting Primary Work Group Schedule
    l_strFilter = _
        "EMP_NUM = " & str(v_lngEmpNum) & _
        " AND EFF_START_DATE <= " & str(v_datDate) & _
        " AND EFF_END_DATE >= " & str(v_datDate) & _
        " AND EFF_DATE <= " & str(v_datDate) & _
        " AND END_DATE >= " & str(v_datDate)
    l_strFilter = MDAC25FilterFix(l_strFilter)
    m_rsWGBaseSchedule.Filter = l_strFilter
    If RecordsExist(m_rsWGBaseSchedule) Then
        l_blnSchedFound = True
        l_datScheduleEffectiveDate = m_rsWGBaseSchedule!EFF_DATE
        l_intRotationCode = m_rsWGBaseSchedule!ROT_CD
        l_intRotationStartWeek = m_rsWGBaseSchedule!ROT_WEEK
        If l_intRotationCode > 0 Then
            l_intWGSchedCode = DetermineRotationScheduleCode( _
                        v_datDate, l_intRotationCode, _
                        l_datScheduleEffectiveDate, l_intRotationStartWeek, l_intRotationWeek)
        Else
            l_intWGSchedCode = m_rsWGBaseSchedule!SCHED_CD
            l_intRotationCode = 0
            l_intRotationWeek = 0
        End If
    Else
        l_intWGSchedCode = -1
        l_intRotationCode = 0
        l_intRotationWeek = 0
        l_blnSchedFound = False
    End If
    
    AddPersonalOrWGSched _
            v_lngEmpNum, v_datDate, l_intWGSchedCode _
            , l_intRotationCode, l_intRotationWeek _
            , True, r_rsAdjBaseSchedbyChargeDate
    
    ProcessWorkGroupSchedule = l_blnSchedFound

    Exit Function
    
eh:

    Dim l_strErrorMessage As String

    l_strErrorMessage = "An error occured while getting the work group " & _
            "schedule for employee (" & str(v_lngEmpNum) & ") " & _
            "on the date " & str(v_datDate) & "."
            
    TrsRaiseError m_strCMODULENAME, "ProcessWorkGroupSchedule", , , , DescArgsFrom(l_strErrorMessage), True
            
            
End Function


            
Private Sub AddDemandLineBaseSchedule( _
                ByVal v_lngEmpNum As Long, _
                ByVal v_datChargeDate As Date, _
                ByVal v_lngLineWG As Long, _
                ByVal v_lngLineNum As Long, _
                ByVal v_strReliefComponent As String, _
                ByVal v_strJob As String, _
                ByVal v_strRequiredInd As String, _
                ByVal v_intBaseSchedCode As Integer, _
                ByVal v_intRotationCode As Integer, _
                ByVal v_intRotationWeek As Integer, _
                ByRef r_rsAdjBaseSchedbyChargeDate As ADODB.Recordset)
                
' The parameter v_strReliefComponent
'   comes from the line assignments
                
' The parameters v_strJob, v_strRequiredInd, and v_intBaseSchedCode
'    come from the base schedule of the line assignment.
' The employee keeps their attributes of the line when
'   making a trade.
' The v_intBaseSchedCode is needed to determine the shift component
'   the employee is due.


   
    On Error GoTo eh
    
    r_rsAdjBaseSchedbyChargeDate.AddNew
    r_rsAdjBaseSchedbyChargeDate!EMP_NUM = v_lngEmpNum
    r_rsAdjBaseSchedbyChargeDate!CHARGE_DATE = v_datChargeDate
    r_rsAdjBaseSchedbyChargeDate!DATE_WORKED = v_datChargeDate
    r_rsAdjBaseSchedbyChargeDate!SCHED_CD = v_intBaseSchedCode
    r_rsAdjBaseSchedbyChargeDate!SCHED_JOB_NUM = v_strJob
    r_rsAdjBaseSchedbyChargeDate!REQ_LINE_IND = v_strRequiredInd
    r_rsAdjBaseSchedbyChargeDate!WG_NUM = v_lngLineWG
    r_rsAdjBaseSchedbyChargeDate!LINE_NUM = v_lngLineNum
    r_rsAdjBaseSchedbyChargeDate!REL_CD = v_strReliefComponent
    r_rsAdjBaseSchedbyChargeDate!SCHED_SRC_CD = 40
    r_rsAdjBaseSchedbyChargeDate!LINE_SRC_CD = 10
    r_rsAdjBaseSchedbyChargeDate!SHIFT_NUM = 1
    r_rsAdjBaseSchedbyChargeDate!SCHED_CD_FOR_SHIFT_COMP = v_intBaseSchedCode
    r_rsAdjBaseSchedbyChargeDate!BASE_SCHED_SRC_CD = 40
    r_rsAdjBaseSchedbyChargeDate!SYSTEM_EXC_CD = ""
    r_rsAdjBaseSchedbyChargeDate!TRADE_NUM = 0
    r_rsAdjBaseSchedbyChargeDate!DOUBLE_SHIFT_IND = "N"
    r_rsAdjBaseSchedbyChargeDate!ROT_CD = v_intRotationCode
    r_rsAdjBaseSchedbyChargeDate!ROT_WEEK = v_intRotationWeek
    
    r_rsAdjBaseSchedbyChargeDate.Update
    Exit Sub
    
eh:

    TrsRaiseError m_strCMODULENAME, "AddDemandLineBaseSchedule"
            
End Sub



Private Sub AddDemandLineTrade( _
                ByVal v_lngEmpNum As Long, _
                ByVal v_strReliefComponent As String, _
                ByVal v_strJob As String, _
                ByVal v_strRequiredInd As String, _
                ByVal v_intBaseSchedCode As Integer, _
                ByVal v_intRotationCode As Integer, _
                ByVal v_intRotationWeek As Integer, _
                ByVal v_rsLineTrades As ADODB.Recordset, _
                ByRef r_rsAdjBaseSchedbyChargeDate As ADODB.Recordset)

' The parameter v_strReliefComponent
'   comes from the line assignments
                
' The parameters v_strJob, v_strRequiredInd, and v_intBaseSchedCode
'    come from the base schedule of the line assignment.
' The employee keeps their attributes of the line when
'   making a trade.
' The v_intBaseSchedCode is needed to determine the shift component
'   the employee is due.

   
    On Error GoTo eh
    
    
    r_rsAdjBaseSchedbyChargeDate.AddNew
    r_rsAdjBaseSchedbyChargeDate!EMP_NUM = v_lngEmpNum
    r_rsAdjBaseSchedbyChargeDate!CHARGE_DATE = v_rsLineTrades!CHARGE_DATE
    r_rsAdjBaseSchedbyChargeDate!DATE_WORKED = v_rsLineTrades!CHARGE_DATE
    r_rsAdjBaseSchedbyChargeDate!SCHED_CD = v_rsLineTrades!SCHED_CD
    r_rsAdjBaseSchedbyChargeDate!SCHED_JOB_NUM = v_strJob
    r_rsAdjBaseSchedbyChargeDate!REQ_LINE_IND = v_strRequiredInd
    r_rsAdjBaseSchedbyChargeDate!WG_NUM = v_rsLineTrades!WG_NUM
    r_rsAdjBaseSchedbyChargeDate!LINE_NUM = v_rsLineTrades!LINE_NUM
    r_rsAdjBaseSchedbyChargeDate!REL_CD = v_strReliefComponent
    r_rsAdjBaseSchedbyChargeDate!SCHED_SRC_CD = 50
    r_rsAdjBaseSchedbyChargeDate!LINE_SRC_CD = 10
    r_rsAdjBaseSchedbyChargeDate!SHIFT_NUM = v_rsLineTrades!SHIFT_NUM
    r_rsAdjBaseSchedbyChargeDate!SCHED_CD_FOR_SHIFT_COMP = v_intBaseSchedCode
    r_rsAdjBaseSchedbyChargeDate!BASE_SCHED_SRC_CD = 40
'JIMB Need Requirement
'Currently the trade exception code is not stored
    r_rsAdjBaseSchedbyChargeDate!SYSTEM_EXC_CD = _
        v_rsLineTrades!ROLE_TYPE_CD & v_rsLineTrades!WEEK_TRADE_IND
    r_rsAdjBaseSchedbyChargeDate!TRADE_NUM = v_rsLineTrades!TRADE_NUM
    r_rsAdjBaseSchedbyChargeDate!DOUBLE_SHIFT_IND = "N"
    r_rsAdjBaseSchedbyChargeDate!ROT_CD = v_intRotationCode
    r_rsAdjBaseSchedbyChargeDate!ROT_WEEK = v_intRotationWeek
    
    r_rsAdjBaseSchedbyChargeDate.Update
    
    Exit Sub
    
eh:

    TrsRaiseError m_strCMODULENAME, "AddDemandLineTrade"
            
End Sub
            
Private Sub AddPersonalOrWGSched( _
                ByVal v_lngEmpNum As Long, _
                ByVal v_datChargeDate As Date, _
                ByVal v_intSchedCode As Integer, _
                ByVal v_intRotationCode As Integer, _
                ByVal v_intRotationWeek As Integer, _
                ByVal v_blnWGSchedule As Boolean, _
                ByRef r_rsAdjBaseSchedbyChargeDate As ADODB.Recordset)
    
    On Error GoTo eh
    
    r_rsAdjBaseSchedbyChargeDate.AddNew
    r_rsAdjBaseSchedbyChargeDate!EMP_NUM = v_lngEmpNum
    r_rsAdjBaseSchedbyChargeDate!CHARGE_DATE = v_datChargeDate
    r_rsAdjBaseSchedbyChargeDate!DATE_WORKED = v_datChargeDate
    r_rsAdjBaseSchedbyChargeDate!SCHED_CD = v_intSchedCode
    r_rsAdjBaseSchedbyChargeDate!SCHED_JOB_NUM = ""
    r_rsAdjBaseSchedbyChargeDate!REQ_LINE_IND = ""
    r_rsAdjBaseSchedbyChargeDate!WG_NUM = 0
    r_rsAdjBaseSchedbyChargeDate!LINE_NUM = 0
    If v_blnWGSchedule Then
        r_rsAdjBaseSchedbyChargeDate!SCHED_SRC_CD = 10
        r_rsAdjBaseSchedbyChargeDate!BASE_SCHED_SRC_CD = 10
    Else
        r_rsAdjBaseSchedbyChargeDate!SCHED_SRC_CD = 20
        r_rsAdjBaseSchedbyChargeDate!BASE_SCHED_SRC_CD = 20
    End If
    r_rsAdjBaseSchedbyChargeDate!LINE_SRC_CD = 0
    r_rsAdjBaseSchedbyChargeDate!REL_CD = ""
    r_rsAdjBaseSchedbyChargeDate!SHIFT_NUM = 1
    r_rsAdjBaseSchedbyChargeDate!SCHED_CD_FOR_SHIFT_COMP = v_intSchedCode
    r_rsAdjBaseSchedbyChargeDate!SYSTEM_EXC_CD = ""
    r_rsAdjBaseSchedbyChargeDate!TRADE_NUM = 0
    r_rsAdjBaseSchedbyChargeDate!DOUBLE_SHIFT_IND = "N"
    r_rsAdjBaseSchedbyChargeDate!ROT_CD = v_intRotationCode
    r_rsAdjBaseSchedbyChargeDate!ROT_WEEK = v_intRotationWeek
    
        
    r_rsAdjBaseSchedbyChargeDate.Update
    
        Exit Sub
    
eh:

    TrsRaiseError m_strCMODULENAME, "AddPersonalOrWGSched"
            
    
End Sub
            
Private Sub AddTradedDoubleSchedule( _
        ByVal v_lngEmpNum As Long, _
        ByVal v_datChargeDate As Date, _
        ByRef r_rsAdjBaseSchedbyChargeDate As ADODB.Recordset)
        
    Dim l_intPriorSchedCode As Integer
    Dim l_intPriorSchedCodeSource As Integer
      
    On Error GoTo eh
    
    l_intPriorSchedCode = r_rsAdjBaseSchedbyChargeDate!SCHED_CD_FOR_SHIFT_COMP
    l_intPriorSchedCodeSource = r_rsAdjBaseSchedbyChargeDate!BASE_SCHED_SRC_CD
    r_rsAdjBaseSchedbyChargeDate.AddNew
    r_rsAdjBaseSchedbyChargeDate!EMP_NUM = v_lngEmpNum
    r_rsAdjBaseSchedbyChargeDate!CHARGE_DATE = v_datChargeDate
    r_rsAdjBaseSchedbyChargeDate!DATE_WORKED = v_datChargeDate
    r_rsAdjBaseSchedbyChargeDate!SCHED_CD = m_rsEmployeeTrades!SCHED_CD
    r_rsAdjBaseSchedbyChargeDate!SCHED_JOB_NUM = ""
    r_rsAdjBaseSchedbyChargeDate!COMP_SHIFT_CD = ""
'JIMB Need Requirement
'Currently the trade exception code is not stored
    r_rsAdjBaseSchedbyChargeDate!SYSTEM_EXC_CD = _
        m_rsEmployeeTrades!ROLE_TYPE_CD & m_rsEmployeeTrades!WEEK_TRADE_IND
    r_rsAdjBaseSchedbyChargeDate!SCHED_CD_FOR_SHIFT_COMP = l_intPriorSchedCode
    r_rsAdjBaseSchedbyChargeDate!BASE_SCHED_SRC_CD = l_intPriorSchedCodeSource
    r_rsAdjBaseSchedbyChargeDate!LINE_SRC_CD = 0
'JIMB Need Requirement
'Does the employee keep relief component on a traded double
'Also this is a Not Demand employee
    r_rsAdjBaseSchedbyChargeDate!REL_CD = ""
    r_rsAdjBaseSchedbyChargeDate!SHIFT_NUM = 2
    r_rsAdjBaseSchedbyChargeDate!SCHED_JOB_NUM = ""
    r_rsAdjBaseSchedbyChargeDate!REQ_LINE_IND = ""
    r_rsAdjBaseSchedbyChargeDate!WG_NUM = 0
    r_rsAdjBaseSchedbyChargeDate!LINE_NUM = 0
    r_rsAdjBaseSchedbyChargeDate!SCHED_SRC_CD = 30
    r_rsAdjBaseSchedbyChargeDate!TRADE_NUM = m_rsEmployeeTrades!TRADE_NUM
    r_rsAdjBaseSchedbyChargeDate!DOUBLE_SHIFT_IND = "N"
    r_rsAdjBaseSchedbyChargeDate!ROT_CD = 0
    r_rsAdjBaseSchedbyChargeDate!ROT_WEEK = 0
    
    r_rsAdjBaseSchedbyChargeDate.Update
    
    Exit Sub
    
eh:

    TrsRaiseError m_strCMODULENAME, "AddTradedDoubleSchedule"
            

End Sub


Public Function GetSchedulesInRotation( _
            ByVal v_intRotationCode As Integer, _
            ByVal v_datDate As Date, _
            ByVal v_datEffectiveDate As Date, _
            ByVal v_intStartWeek As Integer _
            ) As ADODB.Recordset
            
    Dim l_intSchedCode As Integer
    Dim l_intActiveSchedPos As Integer
    
    On Error GoTo eh
    
    Set m_rsRotations = GetSchedulesInRotationDB(v_intRotationCode)
    If RecordsExist(m_rsRotations) Then
        l_intSchedCode = DetermineRotationScheduleCode(v_datDate, v_intRotationCode, v_datEffectiveDate, v_intStartWeek, l_intActiveSchedPos)
        m_rsRotations.AbsolutePosition = l_intActiveSchedPos
        m_rsRotations!CURRENT_SCHED = "Y"
        m_rsRotations.Update
    End If
    
    Set GetSchedulesInRotation = m_rsRotations
    Exit Function
    
eh:

    TrsRaiseError m_strCMODULENAME, "GetSchedulesInRotation", , , , , True
                        
End Function


Private Function GetSchedulesInRotationDB( _
            ByVal l_intRotationCode As Integer _
            ) As ADODB.Recordset
            
    Dim l_obj As dbSchedule.cdSchedule
    
    On Error GoTo eh
    
    Set l_obj = CreateObject("dbSchedule.cdSchedule")
    Set GetSchedulesInRotationDB = l_obj.GetSchedulesInRotation(l_intRotationCode)
    
    Set l_obj = Nothing
    
    Exit Function
    
eh:

    TrsRaiseError m_strCMODULENAME, "GetSchedulesInRotation"
       
End Function


Private Function GetWorkGroupAttributes( _
            ByVal v_lngWorkGroup As Long _
            ) As ADODB.Recordset
            
    Dim l_obj As dbWorkGroup.cdWorkGroup
    
    On Error GoTo eh
    
    Set l_obj = CreateObject("dbWorkGroup.cdWorkGroup")
    Set GetWorkGroupAttributes = l_obj.GetWorkgroupRecord(v_lngWorkGroup)
    
    Set l_obj = Nothing
    
    Exit Function
    
eh:

    TrsRaiseError m_strCMODULENAME, "GetWorkGroupAttributes"
                   
            
End Function



Private Function GetWGDemandLineBaseSchedules( _
            ByVal v_colWorkGroups As Collection, _
            ByVal v_datDate As Date _
            ) As ADODB.Recordset
    
    Dim l_dbSchedule As dbSchedule.cdSchedule
    Dim v_datChargeDate As Date
    Dim v_datEndDate As Date

    On Error GoTo GetWGDemandLineBaseSchedules_EH
        
    Set l_dbSchedule = CreateObject("dbSchedule.cdSchedule")
    Set GetWGDemandLineBaseSchedules = l_dbSchedule.GetWGDemandLineBaseSchedules(v_colWorkGroups, v_datDate)
    Set l_dbSchedule = Nothing
    Exit Function
    
GetWGDemandLineBaseSchedules_EH:
    Set l_dbSchedule = Nothing
    TrsRaiseError m_strCMODULENAME, "GetWGDemandLineBaseSchedules"

End Function

Private Sub LoadRepeatingShiftData( _
        ByVal v_rsSchedCodeByDate As ADODB.Recordset, _
        ByRef r_rsGetAdjBaseSched As ADODB.Recordset)
        
        r_rsGetAdjBaseSched!SCHED_JOB_NUM = v_rsSchedCodeByDate!SCHED_JOB_NUM
        r_rsGetAdjBaseSched!COMP_SHIFT_CD = v_rsSchedCodeByDate!COMP_SHIFT_CD
        r_rsGetAdjBaseSched!REL_CD = v_rsSchedCodeByDate!REL_CD
        r_rsGetAdjBaseSched!SYSTEM_EXC_CD = v_rsSchedCodeByDate!SYSTEM_EXC_CD
        r_rsGetAdjBaseSched!WG_NUM = v_rsSchedCodeByDate!WG_NUM
        r_rsGetAdjBaseSched!LINE_NUM = v_rsSchedCodeByDate!LINE_NUM
        r_rsGetAdjBaseSched!REQ_LINE_IND = v_rsSchedCodeByDate!REQ_LINE_IND
        
        r_rsGetAdjBaseSched!SCHED_CD_FOR_SHIFT_COMP = v_rsSchedCodeByDate!SCHED_CD_FOR_SHIFT_COMP
        r_rsGetAdjBaseSched!SHIFT_NUM = v_rsSchedCodeByDate!SHIFT_NUM
        r_rsGetAdjBaseSched!LINE_SRC_CD = v_rsSchedCodeByDate!LINE_SRC_CD
        r_rsGetAdjBaseSched!ROT_CD = v_rsSchedCodeByDate!ROT_CD
        r_rsGetAdjBaseSched!ROT_WEEK = v_rsSchedCodeByDate!ROT_WEEK
    
        
End Sub


Private Sub DetermineExceptionTimes( _
    ByRef r_strExceptionStartTime As String, _
    ByRef r_strExceptionEndTime As String, _
    ByVal v_rsAdjustedBaseSched As ADODB.Recordset, _
    ByVal v_rsExceptionSched As ADODB.Recordset)
    
    
    Dim l_decExceptionHours As Variant
    Dim l_decExceptionOffSetHours As Variant
    Dim l_decShiftHours As Variant
    Dim l_strShiftStartTime As String
    Dim l_strShiftEndTime As String
    Dim l_datShiftStartDate As Date
    Dim l_datShiftStartDateTime As Date
    Dim l_datShiftEndDate As Date
    Dim l_intExceptionHours As Integer
    Dim l_intExceptionMinutes As Integer
    Dim l_datExceptionStartDateTime As Date
    Dim l_datExceptionEndDateTime As Date
    Dim l_strLunchStartTime As String
    Dim l_strLunchEndTime As String
    Dim l_decHoursOfLunch As Variant
    Dim l_decHoursBeforeLunch As Variant
    Dim l_datLunchStartDate As Date
    Dim l_datLunchStartDateTime As Date
    Dim l_datLunchEndDate As Date
    Dim l_datLunchEndDateTime As Date
    Dim l_decExceptionPhysicalHours As Variant
    

    On Error GoTo eh

    l_datShiftStartDate = v_rsAdjustedBaseSched!SHIFT_START_DATE
    l_strShiftStartTime = Trim(v_rsAdjustedBaseSched!SHIFT_START_TIME)
    l_strLunchStartTime = Trim(v_rsAdjustedBaseSched!SHIFT_LUNCH_START_TIME)
    l_strLunchEndTime = Trim(v_rsAdjustedBaseSched!SHIFT_LUNCH_END_TIME)
    If l_strLunchStartTime = "" Or l_strLunchEndTime = "" Then
        l_strLunchStartTime = l_strShiftStartTime
        l_strLunchEndTime = l_strShiftStartTime
    End If
    l_decExceptionOffSetHours = CDec(v_rsExceptionSched!HOURS_OFFSET_NUM)
    l_decExceptionHours = v_rsExceptionSched!HOURS
    If l_strShiftStartTime > l_strLunchStartTime Then
        l_datLunchStartDate = DateAdd("d", 1, l_datShiftStartDate)
    Else
        l_datLunchStartDate = l_datShiftStartDate
    End If
'There are currently no lunches that cross midnight (2002-04-12)
    If l_strLunchStartTime > l_strLunchEndTime Then
        l_datLunchEndDate = DateAdd("d", 1, l_datLunchStartDate)
    Else
        l_datLunchEndDate = l_datLunchStartDate
    End If
    l_datShiftStartDateTime = DateTimeFrom(l_datShiftStartDate, l_strShiftStartTime)
    l_datLunchStartDateTime = DateTimeFrom(l_datLunchStartDate, l_strLunchStartTime)
    l_decHoursBeforeLunch = CDec(DateDiff("n", l_datShiftStartDateTime, l_datLunchStartDateTime) / 60)
    l_datLunchEndDateTime = DateTimeFrom(l_datLunchEndDate, l_strLunchEndTime)
    l_decHoursOfLunch = CDec(DateDiff("n", l_datLunchStartDateTime, l_datLunchEndDateTime) / 60)
    
    
    If l_decExceptionOffSetHours >= l_decHoursBeforeLunch Then
    'Exception starts after lunch
        l_datExceptionStartDateTime = DateAdd("n", ((l_decExceptionOffSetHours + l_decHoursOfLunch) * 60), l_datShiftStartDateTime)
        l_datExceptionEndDateTime = DateAdd("n", l_decExceptionHours * 60, l_datExceptionStartDateTime)
    Else
    'Exception starts before lunch
        l_datExceptionStartDateTime = DateAdd("n", (l_decExceptionOffSetHours * 60), l_datShiftStartDateTime)
        'Does exception end before lunch
        If (l_decExceptionHours + l_decExceptionOffSetHours) <= l_decHoursBeforeLunch Then
        'Yes have the correct hours to calcualate end time
            l_decExceptionPhysicalHours = l_decExceptionHours
        Else
            'NO, must add lunch time to calculate end time
            l_decExceptionPhysicalHours = l_decHoursOfLunch + l_decExceptionHours
        End If
        l_datExceptionEndDateTime = _
                DateAdd("n", l_decExceptionPhysicalHours * 60, _
                          l_datExceptionStartDateTime)
    End If

    r_strExceptionStartTime = Format(l_datExceptionStartDateTime, "HH:MM")
    r_strExceptionEndTime = Format(l_datExceptionEndDateTime, "HH:MM")
    
    
    Exit Sub
    
eh:

    TrsRaiseError m_strCMODULENAME, "DetermineExceptionTimes"
    
    
End Sub

Private Function GetOvertime( _
            ByVal v_aryEmployees As Variant, _
            ByVal v_datStartDate As Date, _
            ByVal v_datEndDate As Date _
            ) As ADODB.Recordset

    On Error GoTo GetOvertime_EH
         
    
    Dim l_dbSchedule As dbSchedule.cdSchedule
    Dim l_colSchedTypes As Collection
             
    Set l_dbSchedule = CreateObject("dbSchedule.cdSchedule")
    Set GetOvertime = l_dbSchedule.GetEmployeeScheduledOvertime(v_aryEmployees, v_datStartDate, v_datEndDate)
       
    Set l_dbSchedule = Nothing
    
    Exit Function
    
GetOvertime_EH:
    Set l_dbSchedule = Nothing
    TrsRaiseError m_strCMODULENAME, "GetOvertime"
     
End Function


Private Sub InitializeExceptionRepeatingShiftData( _
        ByRef r_rsGetAdjBaseSched As ADODB.Recordset, _
        Optional v_blnOT As Boolean = False)

    
    m_varRepeatingShiftData(0) = r_rsGetAdjBaseSched!EMP_NUM
    m_varRepeatingShiftData(1) = r_rsGetAdjBaseSched!CHARGE_DATE
    m_varRepeatingShiftData(2) = r_rsGetAdjBaseSched!SCHED_CD
    m_varRepeatingShiftData(3) = r_rsGetAdjBaseSched!SCHED_SRC_CD
    m_varRepeatingShiftData(4) = r_rsGetAdjBaseSched!SHIFT_START_TIME
    m_varRepeatingShiftData(5) = r_rsGetAdjBaseSched!SHIFT_END_TIME
    m_varRepeatingShiftData(6) = r_rsGetAdjBaseSched!SHIFT_START_DATE
    m_varRepeatingShiftData(7) = r_rsGetAdjBaseSched!SHIFT_END_DATE
    m_varRepeatingShiftData(8) = Round(CDec(r_rsGetAdjBaseSched!SHIFT_HOURS), 1)
    m_varRepeatingShiftData(9) = r_rsGetAdjBaseSched!SHIFT_CODE
    m_varRepeatingShiftData(10) = r_rsGetAdjBaseSched!SCHED_JOB_NUM
    m_varRepeatingShiftData(11) = r_rsGetAdjBaseSched!LINE_NUM
    m_varRepeatingShiftData(12) = r_rsGetAdjBaseSched!WG_NUM
    m_varRepeatingShiftData(13) = r_rsGetAdjBaseSched!REQ_LINE_IND
    m_varRepeatingShiftData(14) = r_rsGetAdjBaseSched!COMP_SHIFT_CD
    m_varRepeatingShiftData(15) = r_rsGetAdjBaseSched!REL_CD
    m_varRepeatingShiftData(16) = r_rsGetAdjBaseSched!SCHED_CD_FOR_SHIFT_COMP
    m_varRepeatingShiftData(17) = r_rsGetAdjBaseSched!SHIFT_NUM
    m_varRepeatingShiftData(18) = r_rsGetAdjBaseSched!LINE_SRC_CD
    m_varRepeatingShiftData(19) = r_rsGetAdjBaseSched!SHIFT_LUNCH_START_TIME
    m_varRepeatingShiftData(20) = r_rsGetAdjBaseSched!SHIFT_LUNCH_END_TIME
    m_varRepeatingShiftData(21) = r_rsGetAdjBaseSched!TRADE_NUM
    
    m_varRepeatingShiftData(22) = r_rsGetAdjBaseSched!BASE_SCHED_SRC_CD
    m_varRepeatingShiftData(23) = r_rsGetAdjBaseSched!EMP_JOB_NUM
    m_varRepeatingShiftData(24) = r_rsGetAdjBaseSched!ROT_CD
    m_varRepeatingShiftData(25) = r_rsGetAdjBaseSched!ROT_WEEK
    m_varRepeatingShiftData(26) = r_rsGetAdjBaseSched!SYS_COMP_SHIFT_CD
    If v_blnOT Then
        m_varRepeatingShiftData(4) = ""
        m_varRepeatingShiftData(5) = ""
        m_varRepeatingShiftData(8) = 0
        m_varRepeatingShiftData(9) = ""
        m_varRepeatingShiftData(17) = 0
        m_varRepeatingShiftData(19) = ""
        m_varRepeatingShiftData(20) = ""
    End If

End Sub

Private Sub SetOvertimeRepeatingShiftData( _
        ByRef r_rsGetAdjBaseSched As ADODB.Recordset)
        
 'The requirements for scheduled overtime are not final
 'This routine will allow certains fields to be reintialized
 ' after the bulk of the repeating data is set
        
 'Call this first
        SetRepeatingShiftData r_rsGetAdjBaseSched
        
 'Reintialized fields if necessary
        
End Sub


Public Function GetAdjustedBaseWorkGroupSchedules( _
            ByVal v_aryWorkgroups As Variant, _
            ByVal v_datStartDate As Date, _
            ByVal v_datEndDate As Date, _
            Optional ByVal v_ReturnOption As TRSReturnOption = trAdjustedBaseSched, _
            Optional ByVal v_ScheduleSourceStartLevel As TRSSchedSourceStartLevel = trAllScheduleLevels _
            ) As ADODB.Recordset
          
    On Error GoTo GetAdjustedBaseWorkGroupSchedules_EH
          
    Dim l_rsGetAdjBaseSched As ADODB.Recordset
    Dim l_rsGetAdjBaseLineSched As ADODB.Recordset
    Dim l_rsSchedCodesByChargeDate As ADODB.Recordset
    Dim l_rsScheduleWithoutExceptions As ADODB.Recordset
    Dim l_rsEmptyLines As ADODB.Recordset
    Dim l_aryWorkgroups As Variant
    Dim l_intReturnOption As Integer
    Dim l_intScheduleSourceStartLevel As Integer

    l_intReturnOption = v_ReturnOption
    l_intScheduleSourceStartLevel = v_ScheduleSourceStartLevel
    ValidateRequest 0, v_datStartDate, v_datEndDate, l_intReturnOption, l_intScheduleSourceStartLevel
    InitializeRecordsets l_intReturnOption, False, l_rsGetAdjBaseSched, l_rsScheduleWithoutExceptions
    
    If IsArray(v_aryWorkgroups) Then
        l_aryWorkgroups = v_aryWorkgroups
    Else
        ReDim l_aryWorkgroups(0) As String
        l_aryWorkgroups(0) = str(CLng(v_aryWorkgroups))
    End If
    
'Retrieve WorkGroup Schedule Data
'It place data in private global variables to reduce data passing
    GetWorkGroupScheduleData l_aryWorkgroups, v_datStartDate, _
        v_datEndDate, l_intReturnOption, l_intScheduleSourceStartLevel
        
' Determine Employees on lines Schedule by Charge Date
    If IsArray(m_aryEmployees) Then
        GetAdjustedBaseSchedulePrivate _
            m_aryEmployees, v_datStartDate, v_datEndDate, Nothing, _
            l_rsGetAdjBaseSched, l_rsScheduleWithoutExceptions, False, _
            l_intReturnOption, l_intScheduleSourceStartLevel
    End If
    
'Apply an indicator if the employee is in the same work group
' that they are on a demand line assignment for
    ApplyWorkGroupLineInd l_rsGetAdjBaseSched

'Determine Schedule of Line when no employee is assigned
    
    Set l_rsEmptyLines = UnassignedDemandLines(l_rsGetAdjBaseSched, v_datStartDate, v_datEndDate)
    If RecordsExist(l_rsEmptyLines) Then
        Set l_rsSchedCodesByChargeDate = _
            GetLineSchedules(l_rsEmptyLines, l_intScheduleSourceStartLevel)
        Set m_colSchedCodes = GetUniqueSchedCodes(l_rsSchedCodesByChargeDate)
        Set m_rsAvailSched = GetAvailSchedInfo(m_colSchedCodes)
        
        Set l_rsGetAdjBaseLineSched = BuildGetAdjBase(l_rsSchedCodesByChargeDate, l_intReturnOption)
    'Add the Schedule to the output
        AddEmployeeSchedule l_rsGetAdjBaseSched, l_rsGetAdjBaseLineSched
    End If
    

    
    
    Set GetAdjustedBaseWorkGroupSchedules = l_rsGetAdjBaseSched
    
    ClearGlobalRecordSets
    Set l_rsGetAdjBaseSched = Nothing
    Set l_rsScheduleWithoutExceptions = Nothing
    Set l_rsSchedCodesByChargeDate = Nothing
    Exit Function
    
GetAdjustedBaseWorkGroupSchedules_EH:
    ClearGlobalRecordSets
    Set l_rsGetAdjBaseSched = Nothing
    Set l_rsScheduleWithoutExceptions = Nothing
    Set l_rsSchedCodesByChargeDate = Nothing
    Dim l_colDescArgs As Collection
    Set l_colDescArgs = New Collection
    l_colDescArgs.Add str(v_datStartDate)
    l_colDescArgs.Add str(v_datEndDate)
    Select Case Err.Number
        Case trErrbusGetAdjustedBaseSchedInvalidDate
            TrsRaiseError m_strCMODULENAME, "GetAdjustedBaseWorkGroupSchedules", Err.Number, False, l_colDescArgs, , True
        Case Else
            TrsRaiseError m_strCMODULENAME, "GetAdjustedBaseWorkGroupSchedules", , , , , True
    End Select

End Function
    

Private Sub GetWorkGroupScheduleData( _
    ByVal v_aryWorkgroups As Variant, _
    ByVal v_datStartDate As Date, _
    ByVal v_datEndDate As Date, _
    ByVal v_ReturnOption As Integer, _
    ByVal v_ScheduleSourceStartLevel As Integer)
    
 'This method will only get the data needed to satisfy the request
    
    On Error GoTo eh
    
    Set m_rsEmployeesInWorkGroup = _
        GetWorkGroupAssignments(v_aryWorkgroups, v_datStartDate, v_datEndDate)
    Set m_rsWorkGroupLineAssignments = _
        GetWorkGroupLineAssignments(v_aryWorkgroups, v_datStartDate, v_datEndDate)
    Set m_rsWorkGroupLineBaseSched = _
        GetWorkGroupLineBaseSched(v_aryWorkgroups, v_datStartDate, v_datEndDate)
    Set m_rsWorkGroupLineTrade = _
        GetWorkGroupLineTrades(v_aryWorkgroups, v_datStartDate, v_datEndDate)
    If RecordsExist(m_rsWorkGroupLineAssignments) Or _
       RecordsExist(m_rsEmployeesInWorkGroup) Then
        m_aryEmployees = GetEmployees(m_rsWorkGroupLineAssignments, m_rsEmployeesInWorkGroup)
        GetEmployeeScheduleData m_aryEmployees, v_datStartDate, v_datEndDate, _
            v_ReturnOption, v_ScheduleSourceStartLevel
    Else
        m_aryEmployees = Empty
    End If

    Set m_rsRotations = GetRotationSchedules()
    
    Exit Sub
    
eh:
    Set m_rsRotations = Nothing

    
    TrsRaiseError m_strCMODULENAME, "GetWorkGroupScheduleData"

    
End Sub
Private Function GetWorkGroupAssignments( _
            ByVal v_aryWorkgroups As Variant, _
            ByVal v_datStartDate As String, _
            ByVal v_datEndDate As String _
            ) As ADODB.Recordset

    On Error GoTo eh

    Dim l_obj As dbSchedule.cdSchedule
    
    Set l_obj = CreateObject("dbSchedule.cdSchedule")

    Set GetWorkGroupAssignments = l_obj.GetWorkGroupAssignments(v_aryWorkgroups, v_datStartDate, v_datEndDate)

    Set l_obj = Nothing

    Exit Function

eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "GetWorkGroupAssignments"

End Function

Private Function GetWorkGroupLineAssignments( _
            ByVal v_aryWorkgroups As Variant, _
            ByVal v_datStartDate As String, _
            ByVal v_datEndDate As String _
            ) As ADODB.Recordset

    On Error GoTo eh

    Dim l_dbSchedule As dbSchedule.cdSchedule


    Set l_dbSchedule = CreateObject("dbSchedule.cdSchedule")

    Set GetWorkGroupLineAssignments = l_dbSchedule.GetWorkGroupLineAssignments(v_aryWorkgroups, v_datStartDate, v_datEndDate)

    Set l_dbSchedule = Nothing

    Exit Function

eh:
    Set l_dbSchedule = Nothing
    TrsRaiseError m_strCMODULENAME, "GetWorkGroupLineAssignments"

End Function
Private Function GetWorkGroupLineBaseSched( _
            ByVal v_aryWorkgroups As Variant, _
            ByVal v_datStartDate As String, _
            ByVal v_datEndDate As String _
            ) As ADODB.Recordset

    On Error GoTo eh

    Dim l_dbSchedule As dbSchedule.cdSchedule


    Set l_dbSchedule = CreateObject("dbSchedule.cdSchedule")

    Set GetWorkGroupLineBaseSched = l_dbSchedule.GetWorkGroupLineBaseSched(v_aryWorkgroups, v_datStartDate, v_datEndDate)

    Set l_dbSchedule = Nothing

    Exit Function

eh:
    Set l_dbSchedule = Nothing
    TrsRaiseError m_strCMODULENAME, "GetWorkGroupLineBaseSched"

End Function
Private Function GetWorkGroupLineTrades( _
            ByVal v_aryWorkgroups As Variant, _
            ByVal v_datStartDate As String, _
            ByVal v_datEndDate As String _
            ) As ADODB.Recordset

    On Error GoTo eh

    Dim l_dbSchedule As dbSchedule.cdSchedule


    Set l_dbSchedule = CreateObject("dbSchedule.cdSchedule")

    Set GetWorkGroupLineTrades = l_dbSchedule.GetWorkGroupLineTrades(v_aryWorkgroups, v_datStartDate, v_datEndDate)

    Set l_dbSchedule = Nothing

    Exit Function

eh:
    Set l_dbSchedule = Nothing
    TrsRaiseError m_strCMODULENAME, "GetWorkGroupLineTrades"

End Function

Private Function GetEmployees( _
            ByVal v_rs As ADODB.Recordset, _
            ByVal v_rs2 As ADODB.Recordset _
            ) As Variant
            
    Dim l_colEmployee As Collection
    Dim l_lngEmployee As Long

    On Error GoTo eh

    Set l_colEmployee = New Collection
    If RecordsExist(v_rs) Then
        v_rs.MoveFirst
        While Not v_rs.EOF
            l_lngEmployee = v_rs!EMP_NUM
            If IsInList(l_lngEmployee, l_colEmployee) Then
            Else
               l_colEmployee.Add l_lngEmployee
            End If
            v_rs.MoveNext
        Wend
        v_rs.Filter = ""
    End If
    If RecordsExist(v_rs2) Then
        v_rs2.MoveFirst
        While Not v_rs2.EOF
            l_lngEmployee = v_rs2!EMP_NUM
            If IsInList(l_lngEmployee, l_colEmployee) Then
            Else
               l_colEmployee.Add l_lngEmployee
            End If
            v_rs2.MoveNext
        Wend
        v_rs2.Filter = ""
    End If
    
    GetEmployees = ArrayFromCollection(l_colEmployee, 0)

    
Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "GetEmployees"


End Function

Private Function UnassignedDemandLines( _
            ByVal v_rsSchedule As ADODB.Recordset, _
            ByVal v_datStartDate As Date, _
            ByVal v_datEndDate As Date _
            ) As ADODB.Recordset
            
    Dim l_rsUnassignedLines As ADODB.Recordset
    Dim l_datChargeDate As Date
    Dim l_intIndex As Integer
    Dim l_lngNumOfDays As Long
    Dim l_lngWgNum As Long
    Dim l_lngLineNum As Long
    Dim l_datLineSchedStart As Date
    Dim l_datLineSchedEnd As Date
    Dim l_strFilter As String
               
    On Error GoTo eh
    
    If RecordsExist(m_rsWorkGroupLineBaseSched) Then
        Set l_rsUnassignedLines = CreateDemandLineRS()
    Else
        Set UnassignedDemandLines = Nothing
        Exit Function
    End If
    
    l_lngNumOfDays = DateDiff("d", v_datStartDate, v_datEndDate)
    m_rsWorkGroupLineBaseSched.Sort = "WG_NUM DESC, LINE_NUM, START_DATE"
    m_rsWorkGroupLineBaseSched.MoveFirst
    While Not m_rsWorkGroupLineBaseSched.EOF
        For l_intIndex = 0 To l_lngNumOfDays
            l_datChargeDate = DateAdd("d", l_intIndex, v_datStartDate)
            l_datLineSchedStart = m_rsWorkGroupLineBaseSched!START_DATE
            l_datLineSchedEnd = m_rsWorkGroupLineBaseSched!END_DATE
            If l_datChargeDate >= l_datLineSchedStart And _
               l_datChargeDate <= l_datLineSchedEnd Then
               l_lngWgNum = m_rsWorkGroupLineBaseSched!WG_NUM
               l_lngLineNum = m_rsWorkGroupLineBaseSched!LINE_NUM
               l_rsUnassignedLines.AddNew
               l_rsUnassignedLines!CHARGE_DATE = l_datChargeDate
               l_rsUnassignedLines!WG_NUM = l_lngWgNum
               l_rsUnassignedLines!LINE_NUM = l_lngLineNum
               l_rsUnassignedLines.Update
            End If
        Next l_intIndex
        m_rsWorkGroupLineBaseSched.MoveNext
    Wend
  
    If RecordsExist(v_rsSchedule) Then
        v_rsSchedule.Sort = "WG_NUM DESC,LINE_NUM, CHARGE_DATE"
        v_rsSchedule.MoveFirst
        l_rsUnassignedLines.Sort = "WG_NUM DESC,LINE_NUM, CHARGE_DATE"
        l_rsUnassignedLines.MoveFirst
        While Not v_rsSchedule.EOF
            l_datChargeDate = v_rsSchedule!CHARGE_DATE
            l_lngWgNum = v_rsSchedule!WG_NUM
            l_lngLineNum = v_rsSchedule!LINE_NUM
            If l_lngWgNum > 0 And l_lngLineNum > 0 Then
                l_strFilter = "WG_NUM = " & str(l_lngWgNum) & _
                    " AND LINE_NUM = " & str(l_lngLineNum) & _
                    " AND CHARGE_DATE = " & str(l_datChargeDate)
                l_strFilter = MDAC25FilterFix(l_strFilter)
                l_rsUnassignedLines.Filter = l_strFilter
                If RecordsExist(l_rsUnassignedLines) Then
                    l_rsUnassignedLines.Delete
                End If
            End If
            v_rsSchedule.MoveNext
        Wend
    End If
    
    l_rsUnassignedLines.Filter = ""
    
    
    Set UnassignedDemandLines = l_rsUnassignedLines
        
Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "UnassignedDemandLines"


End Function


Private Function CreateDemandLineRS() As ADODB.Recordset

    Dim l_rs As ADODB.Recordset
    
    On Error GoTo eh
    
    Set l_rs = New ADODB.Recordset

    With l_rs
      .Fields.Append "WG_NUM", adInteger, , adFldUpdatable
      .Fields.Append "LINE_NUM", adInteger, , adFldUpdatable
      .Fields.Append "CHARGE_DATE", adDate, , adFldUpdatable
      .CursorLocation = adUseClient
      .CursorType = adOpenDynamic
      .LockType = adLockOptimistic
      .Open
    End With
    
    Set CreateDemandLineRS = l_rs
    
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "CreateDemandLineRS"

    
    
End Function


Private Function GetLineSchedules( _
            ByVal v_rsLines As ADODB.Recordset, _
            ByVal v_ScheduleSourceStartLevel As Integer _
            ) As ADODB.Recordset

    Dim l_lngWgNum As Long
    Dim l_lngLineNum As Long
    Dim l_datChargeDate As Date
    Dim l_rsLineSchedule As ADODB.Recordset

    On Error GoTo eh

    Set l_rsLineSchedule = CreateGetAdjBaseRecordset(trAdjustedBaseSched)
    v_rsLines.Sort = "WG_NUM, LINE_NUM, CHARGE_DATE"
    v_rsLines.MoveFirst
    While Not v_rsLines.EOF
        l_lngWgNum = v_rsLines!WG_NUM
        l_lngLineNum = v_rsLines!LINE_NUM
        l_datChargeDate = v_rsLines!CHARGE_DATE
        DetermineLineSchedule 0, "", l_lngWgNum, l_lngLineNum, l_datChargeDate, _
                m_rsWorkGroupLineBaseSched, m_rsWorkGroupLineTrade, _
                l_rsLineSchedule, v_ScheduleSourceStartLevel
        v_rsLines.MoveNext
    Wend
    
    Set GetLineSchedules = l_rsLineSchedule
    
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "GetLineSchedules"

    
            
End Function


Private Sub DetermineLineSchedule( _
            ByVal v_lngEmpNum As Long, _
            ByVal v_strReliefComponent As String, _
            ByVal v_lngWgNum As Long, _
            ByVal v_lngLineNum As Long, _
            ByVal v_datDate As Date, _
            ByVal v_rsWgLineBaseSched As ADODB.Recordset, _
            ByVal v_rsLineTrades As ADODB.Recordset, _
            ByRef r_rsAdjBaseSchedbyChargeDate As ADODB.Recordset, _
            ByVal v_intScheduleSourceStartLevel As Integer _
            )
                            
    Dim l_strFilter As String
    Dim l_blnSchedFound As Boolean
    Dim l_intDemandBaseSchedCode As Integer
    Dim l_strBaseJobNumber As String
    Dim l_strBaseRequiredLineInd As String
    Dim l_intRotationCode As Integer
    Dim l_intRotationStartWeek As Integer
    Dim l_datScheduleEffectiveDate As Date
    Dim l_intRotationWeek As Integer
    Dim l_lngLineWG As Long
    Dim l_lngLineNum As Long
    Dim l_strReliefComponent As String


    On Error GoTo eh
    
    l_lngLineWG = v_lngWgNum
    l_lngLineNum = v_lngLineNum
    l_strReliefComponent = v_strReliefComponent
     
    'Getting Demand Line Base Schedule
    l_strFilter = "START_DATE <= " & str(v_datDate) & _
                " AND END_DATE >= " & str(v_datDate) & _
                " AND WG_NUM = " & str(l_lngLineWG) & _
                " AND LINE_NUM = " & str(l_lngLineNum)
    l_strFilter = MDAC25FilterFix(l_strFilter)
    v_rsWgLineBaseSched.Filter = l_strFilter
    If RecordsExist(v_rsWgLineBaseSched) Then
        l_datScheduleEffectiveDate = v_rsWgLineBaseSched!START_DATE
        l_intRotationCode = v_rsWgLineBaseSched!ROT_CD
        l_intRotationStartWeek = v_rsWgLineBaseSched!ROT_WEEK
        l_strBaseJobNumber = v_rsWgLineBaseSched!JOB_NUM
        l_strBaseRequiredLineInd = v_rsWgLineBaseSched!REQ_LINE_IND
        If l_intRotationCode > 0 Then
            l_intDemandBaseSchedCode = DetermineRotationScheduleCode( _
                        v_datDate, l_intRotationCode, _
                        l_datScheduleEffectiveDate, l_intRotationStartWeek, l_intRotationWeek)
        Else
            l_intDemandBaseSchedCode = v_rsWgLineBaseSched!SCHED_CD
            l_intRotationCode = 0
            l_intRotationWeek = 0
        End If
    Else
        'Should you raise an error or give -1 schedule
        'I believe a -1 error
        l_strBaseJobNumber = ""
        l_intDemandBaseSchedCode = -1
        l_intRotationCode = 0
        l_intRotationWeek = 0
    End If
    

'Getting Demand Trades
    l_blnSchedFound = False
    If v_intScheduleSourceStartLevel >= trDemandLineTrades Then
        l_strFilter = "CHARGE_DATE = " & str(v_datDate) & _
                    " AND WG_NUM = " & str(l_lngLineWG) & _
                    " AND LINE_NUM = " & str(l_lngLineNum)
        l_strFilter = MDAC25FilterFix(l_strFilter)
        v_rsLineTrades.Filter = l_strFilter
        If RecordsExist(v_rsLineTrades) Then
            v_rsLineTrades.Sort = "SHIFT_NUM DESC"
            v_rsLineTrades.MoveFirst
            While Not v_rsLineTrades.EOF And Not l_blnSchedFound
                AddDemandLineTrade v_lngEmpNum, l_strReliefComponent, _
                         l_strBaseJobNumber, l_strBaseRequiredLineInd _
                        , l_intDemandBaseSchedCode, _
                        l_intRotationCode, l_intRotationWeek, _
                        v_rsLineTrades, r_rsAdjBaseSchedbyChargeDate
                If v_rsLineTrades!SHIFT_NUM = 2 Then
                Else
                    l_blnSchedFound = True
                End If
                v_rsLineTrades.MoveNext
            Wend
        End If
    End If
    
    If l_blnSchedFound Then
        Exit Sub
    End If
    
    AddDemandLineBaseSchedule _
            v_lngEmpNum, v_datDate, l_lngLineWG, l_lngLineNum _
            , l_strReliefComponent _
            , l_strBaseJobNumber, l_strBaseRequiredLineInd _
            , l_intDemandBaseSchedCode, _
            l_intRotationCode, l_intRotationWeek, _
            r_rsAdjBaseSchedbyChargeDate

    Exit Sub
    
eh:

    Dim l_strErrorMessage As String
    Dim l_strAssignMsg As String
    l_strErrorMessage = "An error occured while getting the demand " & _
            "schedule for employee (" & str(v_lngEmpNum) & ") " & _
            "on the date " & str(v_datDate) & "."
            
    l_strAssignMsg = "Demand Line Assignments. "
    
    l_strErrorMessage = l_strErrorMessage & _
            "The employee was assigned via " & l_strAssignMsg & _
            "The employee was on line " & str(l_lngLineNum) & _
            "  in work group number " & str(l_lngLineWG) & "."
            
    TrsRaiseError m_strCMODULENAME, "DetermineLineSchedule", , , , DescArgsFrom(l_strErrorMessage), True
            
            
End Sub

Private Function IsAHoliday( _
            ByVal v_strException As String _
            ) As Boolean
            
'May want to change implentation to look at busCodeTables
' before we go live with TRIS 4
            
    Dim l_strExcCD As String
    
    l_strExcCD = Trim(v_strException)
    
    Select Case l_strExcCD
        Case "H", "HR"
            IsAHoliday = True
        Case Else
            IsAHoliday = False
    End Select

End Function

Private Sub ValidateRequest( _
                ByVal v_intType As Integer, _
                ByVal v_datStartDate As Date, _
                ByVal v_datEndDate As Date, _
                ByRef r_intReturnOption As Integer, _
                ByRef r_intSchedStartLevel As Integer)
                
                
  On Error GoTo eh
                
    If v_datStartDate > v_datEndDate Then
       Err.Raise trErrbusGetAdjustedBaseSchedInvalidDate '20002 + vbObjectError..
    End If
    
'Set Optional Parameters to Return Option 2 Adjusted Base Schedule
'Set Schedule Source Start Level to the top of the hierarchy
'when they are not specified or out of range

' Possibly add logic for requesting a bad return option
' Currently, Sets ReturnOption = 2 when it is bad or not entered
    If r_intReturnOption <= 0 Or r_intReturnOption > 4 Then
        r_intReturnOption = 2
    End If
    
' Possibly add logic for requesting a bad Schedule Start Level option
' Currently, Sets ScheduleSourceStartLevel = 50 when it is bad or not entered
    If r_intSchedStartLevel < trWorkGroupBaseScheduleLevel Or r_intSchedStartLevel > trAllScheduleLevels Then 'if start level not set start at top of heirarchy
        r_intSchedStartLevel = trAllScheduleLevels
    End If
    
    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "ValidateRequest"

           
End Sub

Private Sub InitializeRecordsets( _
                ByVal v_intReturnOption As Integer, _
                ByVal v_blnReturnScheduleWithoutExceptions As Boolean, _
                ByRef r_rsAdjustedBaseSched As ADODB.Recordset, _
                ByRef r_rsSchedWithoutExceptions As ADODB.Recordset)
                
    On Error GoTo eh
                
    Set r_rsSchedWithoutExceptions = Nothing
    If v_intReturnOption = trBaseSchedNoDetails Then
        Set r_rsAdjustedBaseSched = CreateBaseScheduleRecordset
        If v_blnReturnScheduleWithoutExceptions Then
            Set r_rsSchedWithoutExceptions = CreateBaseScheduleRecordset
        End If
    Else
        Set r_rsAdjustedBaseSched = CreateGetAdjBaseRecordset(v_intReturnOption)
        If v_blnReturnScheduleWithoutExceptions Then
            Set r_rsSchedWithoutExceptions = CreateGetAdjBaseRecordset(v_intReturnOption)
        End If
    End If
                
    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "ValidateRequest"
                            
End Sub



Private Sub IntitializeSpecialFields( _
                ByRef r_rsGetAdjBaseSched As ADODB.Recordset, _
                ByRef r_rsScheduleWithoutExceptions As ADODB.Recordset)
                
 'This routine will not be needed with ADO 2.5 (MDAC) and above
 'ADO 2.5 will allow the field to be intialized upon the .AddNew
 'When defining the field with the
 '.Fields.Append method
                
 On Error GoTo eh
 
    Exit Sub
 
    r_rsGetAdjBaseSched.Filter = ""
    If RecordsExist(r_rsGetAdjBaseSched) Then
        r_rsGetAdjBaseSched.MoveFirst
        While Not r_rsGetAdjBaseSched.EOF
            r_rsGetAdjBaseSched!SCHED_CHANGE_PMT = "N"
            r_rsGetAdjBaseSched!WG_LINE_IND = "N"
            r_rsGetAdjBaseSched.Update
            r_rsGetAdjBaseSched.MoveNext
        Wend
    End If
    
    If Not r_rsScheduleWithoutExceptions Is Nothing Then
        r_rsScheduleWithoutExceptions.Filter = ""
        If RecordsExist(r_rsScheduleWithoutExceptions) Then
            r_rsScheduleWithoutExceptions.MoveFirst
            While Not r_rsScheduleWithoutExceptions.EOF
                r_rsScheduleWithoutExceptions!SCHED_CHANGE_PMT = "N"
                r_rsScheduleWithoutExceptions!WG_LINE_IND = "N"
                r_rsScheduleWithoutExceptions.Update
                r_rsScheduleWithoutExceptions.MoveNext
            Wend
        End If
    End If
                 
    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "IntitializeSpecialFields"
                
End Sub

Private Sub ApplySorts( _
                ByVal v_ReturnOption As Integer, _
                ByRef r_rsGetAdjBaseSched As ADODB.Recordset, _
                ByRef r_rsScheduleWithoutExceptions As ADODB.Recordset, _
                ByVal v_blnReturnScheduleWithoutExceptions As Boolean)
                    
    Dim l_strSort As String
                
    On Error GoTo eh
    
                
    If v_ReturnOption >= trAdjustedBaseSched Then
        l_strSort = "EMP_NUM, CHARGE_DATE , DATE_WORKED, START_TIME, END_TIME"
    Else
        l_strSort = "EMP_NUM, CHARGE_DATE , SHIFT_CODE"
    End If
    
    r_rsGetAdjBaseSched.Sort = l_strSort
    If v_blnReturnScheduleWithoutExceptions Then
        r_rsScheduleWithoutExceptions.Sort = l_strSort
    End If
    
    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "ApplySorts"
                
End Sub

Private Sub GetAdjustedBaseSchedulePrivate _
             (ByVal v_aryEmployees As Variant, _
              ByVal v_datStartDate As Date, _
              ByVal v_datEndDate As Date, _
              ByVal v_rsEmployeeAndDates As ADODB.Recordset, _
              ByRef r_rsGetAdjBaseSched As ADODB.Recordset, _
              ByRef r_rsScheduleWithoutExceptions As ADODB.Recordset, _
              ByVal v_blnReturnScheduleWithoutExceptions As Boolean, _
              ByVal v_ReturnOption As Integer, _
              ByVal v_ScheduleSourceStartLevel As Integer)
              
              
    Dim l_lngI As Long
    Dim l_lngEmpNum As Long
    Dim l_lngStart As Long
    Dim l_lngEnd As Long
    Dim l_rsSchedCodesByChargeDate As ADODB.Recordset
    
    On Error GoTo eh
 
' Determine Schedule by Charge Date
    Set l_rsSchedCodesByChargeDate = DetermineEmployeeSchedules(v_aryEmployees, v_datStartDate, v_datEndDate, v_ScheduleSourceStartLevel, v_rsEmployeeAndDates)
    Set m_colSchedCodes = GetUniqueSchedCodes(l_rsSchedCodesByChargeDate)
    Set m_rsAvailSched = GetAvailSchedInfo(m_colSchedCodes)
    
'Split out the details
    l_rsSchedCodesByChargeDate.Filter = ""
    Set r_rsGetAdjBaseSched = BuildGetAdjBase(l_rsSchedCodesByChargeDate, v_ReturnOption)
        
'Apply Employees' base job number

'''''    l_lngStart = LBound(v_aryEmployees)
'''''    l_lngEnd = UBound(v_aryEmployees)
'''''    For l_lngI = l_lngStart To l_lngEnd
'''''        If v_ReturnOption > trBaseSchedNoDetails Then
'''''            ApplyEmployeeBaseJobNum v_aryEmployees(l_lngI), v_datStartDate, v_datEndDate, r_rsGetAdjBaseSched
'''''        End If
'''''    Next l_lngI
    
     'Copy to No Exceptions Recordset if needed
    If v_blnReturnScheduleWithoutExceptions Then
        AddEmployeeSchedule r_rsScheduleWithoutExceptions, r_rsGetAdjBaseSched
    End If
    
'Apply Exceptions and Overtime if needed
    If v_ReturnOption >= trAdjBaseSchedWithExceptions Then
        m_rsExceptionSched.Filter = ""
        If RecordsExist(m_rsExceptionSched) Then
            ApplyExceptions r_rsGetAdjBaseSched, m_rsExceptionSched, v_rsEmployeeAndDates
        End If
    End If
    If v_ReturnOption >= trAdjBaseSchedWithExceptionsAndOvertime Then
        m_rsOvertimeSched.Filter = ""
        If RecordsExist(m_rsOvertimeSched) Then
            ApplyOvertimeExceptions r_rsGetAdjBaseSched, m_rsOvertimeSched, v_rsEmployeeAndDates
        End If
    End If
    
    If v_ReturnOption > trBaseSchedNoDetails Then
'        IntitializeSpecialFields r_rsGetAdjBaseSched, r_rsScheduleWithoutExceptions
        ApplyScheduleChangePayments v_aryEmployees, v_datStartDate, v_datEndDate, r_rsGetAdjBaseSched, r_rsScheduleWithoutExceptions
    End If
    
    ApplySorts v_ReturnOption, r_rsGetAdjBaseSched, r_rsScheduleWithoutExceptions, v_blnReturnScheduleWithoutExceptions
             
    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "GetAdjustedBaseSchedulePrivate"
    
End Sub

Private Sub ApplyWorkGroupLineInd( _
                ByRef r_rsGetAdjBaseSched As ADODB.Recordset)
                
    On Error GoTo eh
    
    Dim l_strFilter As String
    Dim l_lngEmpNum As String
    Dim l_datChargeDate As Date
    
    If Not RecordsExist(m_rsEmployeesInWorkGroup) Then
            Exit Sub
        'No Employees in selected work group(s)
    End If
    
    m_rsDemandLineAssignments.Filter = ""
    If Not RecordsExist(m_rsDemandLineAssignments) Then
        Exit Sub
        'No Employees on a demand line assignment
    End If
    
    l_strFilter = "WG_NUM > 0 AND LINE_NUM > 0"
    l_strFilter = MDAC25FilterFix(l_strFilter)
    r_rsGetAdjBaseSched.Filter = l_strFilter
    If Not RecordsExist(r_rsGetAdjBaseSched) Then
        r_rsGetAdjBaseSched.Filter = ""
        Exit Sub
        'No Employees on a demand line assignment
        'This codition should not be met, but is here just incase
    End If
    
    r_rsGetAdjBaseSched.Sort = "EMP_NUM, CHARGE_DATE"
    r_rsGetAdjBaseSched.MoveFirst
    While Not r_rsGetAdjBaseSched.EOF
        l_lngEmpNum = r_rsGetAdjBaseSched!EMP_NUM
        l_datChargeDate = r_rsGetAdjBaseSched!CHARGE_DATE
        l_strFilter = "EMP_NUM = " & str(l_lngEmpNum) & _
            " AND EFF_START_DATE <= " & str(l_datChargeDate) & _
            " AND EFF_END_DATE >= " & str(l_datChargeDate)
        l_strFilter = MDAC25FilterFix(l_strFilter)
        m_rsEmployeesInWorkGroup.Filter = l_strFilter
        If RecordsExist(m_rsEmployeesInWorkGroup) Then
            If m_rsEmployeesInWorkGroup!WG_NUM = _
               r_rsGetAdjBaseSched!WG_NUM Then
                r_rsGetAdjBaseSched!WG_LINE_IND = "Y"
                r_rsGetAdjBaseSched.Update
            End If
        End If
        r_rsGetAdjBaseSched.MoveNext
    Wend
    
    r_rsGetAdjBaseSched.Filter = ""
    
    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "ApplyWorkGroupLineInd"
                
End Sub

Private Sub DetermineAdjBaseSched( _
            ByVal v_lngEmpNum As Long, _
            ByVal v_datStartDate As Date, _
            ByVal v_datEndDate As Date, _
            ByVal v_ScheduleSourceStartLevel As Integer, _
            ByRef r_rsSchedule As ADODB.Recordset)
                                    
''''''''''
'  This Function returns a record set that contains the schedule
' codes for each charge date between the start and end date.
' System determines the schedule for each day from the
' workgroup schedule, demand and non-demand schedules
' and schedule changes and traded doubles.
'There are 2 paths in the Algorithm
'   1. Assigned to a Line
'   2. Not Assigned to a Line
'
'   1. Assigned to a Line
'     a. Deterimine Line Assigned to
'       1) Relief (ttr_line_relief)
'       2) Demand Scheduling (ttr_line_asgnm)
'     b. Determine Schedule of the line
'       1) Line Trade (ttr_line_trades)
'       2) Line Base Schedule (ttr_wg_line_sched)
'   1. Not Assigned to a line
'      a. Trade (ttr_emp_trades)
'      b. Personal Schedule (ttr_emp_sched)
'      c. Primary Work Group Base Schedule (ttr_wg_base_sched)
'
' Please note that the above may change to
'   1. Trade (ttr_emp_trades)
'   2. Assigned to a Line
'     a. Deterimine Line Assigned to
'       1) Relief (ttr_line_relief)
'       2) Demand Scheduling (ttr_line_asgnm)
'     b. Determine Schedule of the line
'       1) Line Trade (ttr_line_trades)
'       2) Line Base Schedule (ttr_wg_line_sched)
'   3. Personal Schedule (ttr_emp_sched)
'   4. Primary Work Group Base Schedule (ttr_wg_base_sched)
'
' The system starts at the top of the hierarchy checking for
' existence of a schedule at each level.  The first schedule found
' is the regular shift schedule for the charge date.
' Traded double shifts are added to the employees schedule.


    Dim l_lngNumOfDays As Long
    Dim l_intIndex As Long
    Dim l_datChargeDate As Date
    
    l_lngNumOfDays = DateDiff("d", v_datStartDate, v_datEndDate)
    
    For l_intIndex = 0 To l_lngNumOfDays
        l_datChargeDate = DateAdd("d", l_intIndex, v_datStartDate)
        DetermineAdjBasSchedOnDate v_lngEmpNum, l_datChargeDate, _
            v_ScheduleSourceStartLevel, r_rsSchedule

    Next l_intIndex
        
   
    
End Sub

Private Sub DetermineAdjBasSchedOnDate( _
                ByVal v_lngEmpNum As Long, _
                ByVal v_datChargeDate As Date, _
                ByVal v_ScheduleSourceStartLevel As Integer, _
                ByRef r_rsSchedule As ADODB.Recordset)
                
    Dim l_bolSchedFoundOnChargeDate As Boolean
    Dim l_blnAssignedToALine As Boolean
    
    l_bolSchedFoundOnChargeDate = False
    If v_ScheduleSourceStartLevel >= trDemandLineBaseSchedLevel Then
        l_blnAssignedToALine = IsEmployeeAssigenedToALine( _
                    v_lngEmpNum, v_datChargeDate)
        If l_blnAssignedToALine Then
            l_bolSchedFoundOnChargeDate = True
            ProcessDemandLineSchedule v_lngEmpNum, v_datChargeDate, r_rsSchedule, v_ScheduleSourceStartLevel
        End If
    End If
       
' Since the employee gets his original shift component prior to a
' trade, I plan on adding the employee's schedule prior to the trade
' and then apply the trades on top of that
    If Not l_bolSchedFoundOnChargeDate Then
        If v_ScheduleSourceStartLevel >= trNonDemandLevel Then
            l_bolSchedFoundOnChargeDate = ProcessPersonalSchedule(v_lngEmpNum, v_datChargeDate, r_rsSchedule)
            If Not l_bolSchedFoundOnChargeDate Then
                l_bolSchedFoundOnChargeDate = ProcessWorkGroupSchedule(v_lngEmpNum, v_datChargeDate, r_rsSchedule)
            End If
        Else
            l_bolSchedFoundOnChargeDate = ProcessWorkGroupSchedule(v_lngEmpNum, v_datChargeDate, r_rsSchedule)
        End If
    End If
       
'Apply Employee trades to schedule ignoring employee assigned to a line
'Trades are changing a schedule or adding a schedule in the case of a double
'The codes assumes that the employee had a schedule
' prior to the trade
    If Not l_blnAssignedToALine And _
        l_bolSchedFoundOnChargeDate And _
        v_ScheduleSourceStartLevel >= trEmployeeTradesLevel Then
        ProcessTrades v_lngEmpNum, v_datChargeDate, r_rsSchedule
    End If
            
End Sub


Public Function GABEmployeeSchedOnSpecificDates( _
            ByVal v_rsEmployeesAndDates As ADODB.Recordset, _
            ByRef r_rsScheduleWithoutExceptions As ADODB.Recordset, _
            Optional ByVal v_blnReturnScheduleWithoutExceptions As Boolean = False, _
            Optional ByVal v_ReturnOption As TRSReturnOption = trAdjustedBaseSched, _
            Optional ByVal v_ScheduleSourceStartLevel As TRSSchedSourceStartLevel = trAllScheduleLevels _
            ) As ADODB.Recordset
            
    On Error GoTo GABEmployeeSchedOnSpecificDates_EH
            
    Dim l_rsGetAdjBaseSched As ADODB.Recordset
    Dim l_rsGetAdjBaseLineSched As ADODB.Recordset
    Dim l_rsScheduleWithoutExceptions As ADODB.Recordset
    Dim l_aryWorkgroups As Variant
    Dim l_intReturnOption As Integer
    Dim l_intScheduleSourceStartLevel As Integer
    Dim l_datStartDate As Date
    Dim l_datEndDate As Date
    Dim l_aryEmployees As Variant

    If Not RecordsExist(v_rsEmployeesAndDates) Then
        Set GABEmployeeSchedOnSpecificDates = Nothing
        Exit Function
    End If
    
    v_rsEmployeesAndDates.Sort = "START_DATE"
    v_rsEmployeesAndDates.MoveFirst
    l_datStartDate = v_rsEmployeesAndDates!START_DATE
    v_rsEmployeesAndDates.Sort = "END_DATE DESC"
    v_rsEmployeesAndDates.MoveFirst
    l_datEndDate = v_rsEmployeesAndDates!END_DATE
    l_intReturnOption = v_ReturnOption
    l_intScheduleSourceStartLevel = v_ScheduleSourceStartLevel
    ValidateRequest 0, l_datStartDate, l_datEndDate, l_intReturnOption, l_intScheduleSourceStartLevel
    InitializeRecordsets l_intReturnOption, False, l_rsGetAdjBaseSched, l_rsScheduleWithoutExceptions
    
    l_aryEmployees = GetEmployees(v_rsEmployeesAndDates, Nothing)
    
    'Retrieve Employee Schedule Data
'It place data in private global variables to reduce data passing
    GetEmployeeScheduleData l_aryEmployees, l_datStartDate, _
        l_datEndDate, l_intReturnOption, l_intScheduleSourceStartLevel
        
    Set m_rsRotations = GetRotationSchedules()
    
    v_rsEmployeesAndDates.Sort = "EMP_NUM,START_DATE"

    GetAdjustedBaseSchedulePrivate _
        l_aryEmployees, l_datStartDate, l_datEndDate, v_rsEmployeesAndDates, _
        l_rsGetAdjBaseSched, l_rsScheduleWithoutExceptions, v_blnReturnScheduleWithoutExceptions, _
        l_intReturnOption, l_intScheduleSourceStartLevel
    
    Set GABEmployeeSchedOnSpecificDates = l_rsGetAdjBaseSched
    Set r_rsScheduleWithoutExceptions = l_rsScheduleWithoutExceptions
    
    ClearGlobalRecordSets
    Set l_rsGetAdjBaseSched = Nothing
    Exit Function
    
GABEmployeeSchedOnSpecificDates_EH:
    ClearGlobalRecordSets
    Set l_rsGetAdjBaseSched = Nothing
    Set r_rsScheduleWithoutExceptions = Nothing
    Set l_rsScheduleWithoutExceptions = Nothing
    Dim l_colDescArgs As Collection
    Set l_colDescArgs = New Collection
    l_colDescArgs.Add str(l_datStartDate)
    l_colDescArgs.Add str(l_datEndDate)
    Select Case Err.Number
        Case trErrbusGetAdjustedBaseSchedInvalidDate
            TrsRaiseError m_strCMODULENAME, "GABEmployeeSchedOnSpecificDates", Err.Number, False, l_colDescArgs, , True
        Case Else
            TrsRaiseError m_strCMODULENAME, "GABEmployeeSchedOnSpecificDates", , , , , True
    End Select
         
End Function


Public Function GetPendingAdjustedBaseSchedule( _
            ByVal v_rsEmployeesAndDates As ADODB.Recordset, _
            ByRef r_rsScheduleWithoutExceptions As ADODB.Recordset, _
            Optional ByVal v_blnReturnScheduleWithoutExceptions As Boolean = False, _
            Optional ByVal v_ReturnOption As TRSReturnOption = trAdjustedBaseSched, _
            Optional ByVal v_ScheduleSourceStartLevel As TRSSchedSourceStartLevel = trAllScheduleLevels, _
            Optional ByVal v_rsWGAssignmentSPA As ADODB.Recordset, _
            Optional ByVal v_rsWGBaseSchedChanges As ADODB.Recordset, _
            Optional ByVal v_rsWGLineSchedChanges As ADODB.Recordset, _
            Optional ByVal v_rsWGAssignChanges As ADODB.Recordset, _
            Optional ByVal v_rsWGAssignOtherChanges As ADODB.Recordset, _
            Optional ByVal v_rsEmpLineAssignChanges As ADODB.Recordset, _
            Optional ByVal v_rsPersonalChanges As ADODB.Recordset, _
            Optional ByVal v_rsLineTradeChanges As ADODB.Recordset, _
            Optional ByVal v_rsEmpTradeChanges As ADODB.Recordset _
            ) As ADODB.Recordset
            
'    v_rsWGAssignmentSPA is previously retrieved employee primary workgroup assignment
'    data that can be used instead of accessing the database again.
            
On Error GoTo eh
            
    Dim l_rsGetAdjBaseSched As ADODB.Recordset
    Dim l_rsGetAdjBaseLineSched As ADODB.Recordset
    Dim l_rsScheduleWithoutExceptions As ADODB.Recordset
    Dim l_intReturnOption As Integer
    Dim l_intScheduleSourceStartLevel As Integer
    Dim l_datStartDate As Date
    Dim l_datEndDate As Date
    Dim l_aryEmployees As Variant
    Dim l_aryWorkgroups As Variant
    Dim l_aryDemandLines As Variant
    Dim l_blnWGPendingChanges As Boolean

    If Not RecordsExist(v_rsEmployeesAndDates) Then
        Set GetPendingAdjustedBaseSchedule = Nothing
        Exit Function
    End If
    
    If Not v_rsWGAssignmentSPA Is Nothing Then
        Set m_rsWGEmpAssignments = v_rsWGAssignmentSPA
    End If
    
    l_blnWGPendingChanges = PendingWGAssignmentChanges( _
            v_rsWGBaseSchedChanges, v_rsWGAssignChanges, v_rsWGAssignOtherChanges)
    
'    5/14/02 RSG - Determine start and end date for data retrieval as the
'    min start and max end date from pending changes and request.
'    Necessary to apply pending changes.
    DetermineStartEndDates l_datStartDate, l_datEndDate, v_rsEmployeesAndDates, _
            v_rsWGBaseSchedChanges, v_rsWGLineSchedChanges, _
            v_rsWGAssignChanges, v_rsWGAssignOtherChanges, _
            v_rsEmpLineAssignChanges, v_rsPersonalChanges, _
            v_rsLineTradeChanges, v_rsEmpTradeChanges
            
    l_aryWorkgroups = DeterminePendingWorkgroupsInScope( _
            v_rsWGBaseSchedChanges, v_rsWGLineSchedChanges, _
            v_rsWGAssignChanges, v_rsWGAssignOtherChanges, v_rsEmpLineAssignChanges)
            
    l_aryDemandLines = DeterminePendingDemandLinesInScope( _
                v_rsWGLineSchedChanges, v_rsEmpLineAssignChanges)

    l_intReturnOption = v_ReturnOption
    l_intScheduleSourceStartLevel = v_ScheduleSourceStartLevel
    ValidateRequest 0, l_datStartDate, l_datEndDate, l_intReturnOption, l_intScheduleSourceStartLevel
    InitializeRecordsets l_intReturnOption, False, l_rsGetAdjBaseSched, l_rsScheduleWithoutExceptions
    
    l_aryEmployees = GetEmployees(v_rsEmployeesAndDates, Nothing)
    
    'Retrieve Employee Schedule Data
'It place data in private global variables to reduce data passing
    GetEmployeeScheduleData l_aryEmployees, l_datStartDate, l_datEndDate, _
         l_intReturnOption, l_intScheduleSourceStartLevel, adLockBatchOptimistic, l_blnWGPendingChanges, l_aryWorkgroups, l_aryDemandLines
    ApplyPendingChanges v_rsWGBaseSchedChanges, v_rsWGLineSchedChanges, _
            v_rsWGAssignChanges, v_rsWGAssignOtherChanges, v_rsEmpLineAssignChanges, _
            v_rsPersonalChanges, v_rsLineTradeChanges, v_rsEmpTradeChanges
        
    Set m_rsRotations = GetRotationSchedules()
    
    v_rsEmployeesAndDates.Sort = "EMP_NUM,START_DATE"

    GetAdjustedBaseSchedulePrivate _
        l_aryEmployees, l_datStartDate, l_datEndDate, v_rsEmployeesAndDates, _
        l_rsGetAdjBaseSched, l_rsScheduleWithoutExceptions, v_blnReturnScheduleWithoutExceptions, _
        l_intReturnOption, l_intScheduleSourceStartLevel
    
    Set GetPendingAdjustedBaseSchedule = l_rsGetAdjBaseSched
    Set r_rsScheduleWithoutExceptions = l_rsScheduleWithoutExceptions
    
    ClearGlobalRecordSets
    Set l_rsGetAdjBaseSched = Nothing
    Exit Function
    
eh:
    ClearGlobalRecordSets
    Set l_rsGetAdjBaseSched = Nothing
    Set r_rsScheduleWithoutExceptions = Nothing
    Set l_rsScheduleWithoutExceptions = Nothing
    Dim l_colDescArgs As Collection
    Set l_colDescArgs = New Collection
    l_colDescArgs.Add str(l_datStartDate)
    l_colDescArgs.Add str(l_datEndDate)
    Select Case Err.Number
        Case trErrbusGetAdjustedBaseSchedInvalidDate
            TrsRaiseError m_strCMODULENAME, "GetPendingAdjustedBaseSchedule", Err.Number, False, l_colDescArgs, , True
        Case Else
            TrsRaiseError m_strCMODULENAME, "GetPendingAdjustedBaseSchedule", , , , , True
    End Select
             
         
End Function
                
                
Private Function PendingWGAssignmentChanges( _
            ByRef r_rsWGBaseSchedChanges As ADODB.Recordset, _
            ByRef r_rsWGAssignChanges As ADODB.Recordset, _
            ByRef r_rsWGAssignOtherChanges As ADODB.Recordset) As Boolean
' Answer TRUE when any of the given recordsets have pending records;
' otherwise FALSE.
On Error GoTo eh

    If RecordsExist(r_rsWGBaseSchedChanges) Then
        r_rsWGBaseSchedChanges.Filter = adFilterPendingRecords
        If r_rsWGBaseSchedChanges.RecordCount > 0 Then
            r_rsWGBaseSchedChanges.Filter = adFilterNone
            PendingWGAssignmentChanges = True
            Exit Function
        End If
        r_rsWGBaseSchedChanges.Filter = adFilterNone
    End If
    
    If RecordsExist(r_rsWGAssignChanges) Then
        r_rsWGAssignChanges.Filter = adFilterPendingRecords
        If r_rsWGAssignChanges.RecordCount > 0 Then
            r_rsWGAssignChanges.Filter = adFilterNone
            PendingWGAssignmentChanges = True
            Exit Function
        End If
        r_rsWGAssignChanges.Filter = adFilterNone
    End If
    
    If RecordsExist(r_rsWGAssignOtherChanges) Then
        r_rsWGAssignOtherChanges.Filter = adFilterPendingRecords
        If r_rsWGAssignOtherChanges.RecordCount > 0 Then
            r_rsWGAssignOtherChanges.Filter = adFilterNone
            PendingWGAssignmentChanges = True
            Exit Function
        End If
        r_rsWGAssignOtherChanges.Filter = adFilterNone
    End If

    PendingWGAssignmentChanges = False
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "PendingWGAssignmentChanges"
End Function

                
Private Function DetermineStartEndDates( _
            ByRef r_datStartDate As Date, _
            ByRef r_datEndDate As Date, _
            ByRef r_rsEmployeesAndDates As ADODB.Recordset, _
            ByRef r_rsWGBaseChanges As ADODB.Recordset, _
            ByRef r_rsWGLineChanges As ADODB.Recordset, _
            ByRef v_rsWGAssignChanges As ADODB.Recordset, _
            ByRef v_rsWGAssignOtherChanges As ADODB.Recordset, _
            ByRef r_rsEmpLineAssignChanges As ADODB.Recordset, _
            ByRef r_rsPersonalChanges As ADODB.Recordset, _
            ByRef r_rsLineTradeChanges As ADODB.Recordset, _
            ByRef r_rsEmpTradeChanges As ADODB.Recordset)

On Error GoTo eh
'    Min/max from schedule request
    r_rsEmployeesAndDates.Sort = "START_DATE"
    r_rsEmployeesAndDates.MoveFirst
    r_datStartDate = r_rsEmployeesAndDates!START_DATE
    r_rsEmployeesAndDates.Sort = "END_DATE DESC"
    r_rsEmployeesAndDates.MoveFirst
    r_datEndDate = r_rsEmployeesAndDates!END_DATE

'    Min/max from pending changes
    If RecordsExist(r_rsPersonalChanges) Then
        DetermineMinMaxPersonal r_datStartDate, r_datEndDate, r_rsPersonalChanges
    End If
    
    If RecordsExist(r_rsEmpLineAssignChanges) Then
        DetermineMinMaxDemand r_datStartDate, r_datEndDate, r_rsEmpLineAssignChanges
    End If
    
    If RecordsExist(v_rsWGAssignChanges) Then
        DetermineMinMaxWGAssign r_datStartDate, r_datEndDate, v_rsWGAssignChanges
    End If
    
    If RecordsExist(v_rsWGAssignOtherChanges) Then
        DetermineMinMaxWGAssign r_datStartDate, r_datEndDate, v_rsWGAssignOtherChanges
    End If

    If RecordsExist(r_rsWGLineChanges) Then
        DetermineMinMaxWGLine r_datStartDate, r_datEndDate, r_rsWGLineChanges
    End If
    
    If RecordsExist(r_rsWGBaseChanges) Then
        DetermineMinMaxWGBase r_datStartDate, r_datEndDate, r_rsWGBaseChanges
    End If

    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "DetermineStartEndDates"
End Function
                
                
Private Sub DetermineMinMaxPersonal( _
            ByRef r_datStartDate As Date, _
            ByRef r_datEndDate As Date, _
            ByRef r_rsPendingChanges As ADODB.Recordset)

    Dim l_strRecStatus As String

On Error GoTo eh
    r_rsPendingChanges.MoveFirst
    Do Until r_rsPendingChanges.EOF
        l_strRecStatus = Trim(r_rsPendingChanges!RECORD_STATUS)
        Select Case l_strRecStatus
            Case "I"
                r_datStartDate = MinDate(r_rsPendingChanges!START_DATE, r_datStartDate)
                r_datEndDate = MaxDate(r_rsPendingChanges!END_DATE, r_datEndDate)
            Case "C"
                r_datStartDate = MinDate(r_rsPendingChanges!START_DATE, r_datStartDate)
                r_datStartDate = MinDate(r_rsPendingChanges!KEY_START_DATE, r_datStartDate)
                r_datEndDate = MaxDate(r_rsPendingChanges!END_DATE, r_datEndDate)
                r_datEndDate = MaxDate(r_rsPendingChanges!KEY_END_DATE, r_datEndDate)
            Case "D"
                r_datStartDate = MinDate(r_rsPendingChanges!KEY_START_DATE, r_datStartDate)
                r_datEndDate = MaxDate(r_rsPendingChanges!KEY_END_DATE, r_datEndDate)
        End Select
        r_rsPendingChanges.MoveNext
    Loop
    r_rsPendingChanges.MoveFirst
    
    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "DetermineMinMaxPersonal"
End Sub


Private Sub DetermineMinMaxDemand( _
            ByRef r_datStartDate As Date, _
            ByRef r_datEndDate As Date, _
            ByRef r_rsPendingChanges As ADODB.Recordset)
'    This looks the same as DetermineMinMaxPersonal.
'    If still the same after integration, rename it and both use it.
    Dim l_strRecStatus As String

On Error GoTo eh
    r_rsPendingChanges.MoveFirst
    Do Until r_rsPendingChanges.EOF
        l_strRecStatus = Trim(r_rsPendingChanges!RECORD_STATUS)
        Select Case l_strRecStatus
            Case "I"
                r_datStartDate = MinDate(r_rsPendingChanges!START_DATE, r_datStartDate)
                r_datEndDate = MaxDate(r_rsPendingChanges!END_DATE, r_datEndDate)
            Case "C"
                r_datStartDate = MinDate(r_rsPendingChanges!START_DATE, r_datStartDate)
                r_datStartDate = MinDate(r_rsPendingChanges!KEY_START_DATE, r_datStartDate)
                r_datEndDate = MaxDate(r_rsPendingChanges!END_DATE, r_datEndDate)
                r_datEndDate = MaxDate(r_rsPendingChanges!KEY_END_DATE, r_datEndDate)
            Case "D"
                r_datStartDate = MinDate(r_rsPendingChanges!KEY_START_DATE, r_datStartDate)
                r_datEndDate = MaxDate(r_rsPendingChanges!KEY_END_DATE, r_datEndDate)
        End Select
        r_rsPendingChanges.MoveNext
    Loop
    r_rsPendingChanges.MoveFirst
    
    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "DetermineMinMaxDemand"
End Sub


Private Sub DetermineMinMaxWGLine( _
            ByRef r_datStartDate As Date, _
            ByRef r_datEndDate As Date, _
            ByRef r_rsPendingChanges As ADODB.Recordset)
'    This looks the same as xxxx.
'    If still the same after integration, rename it and both use it.
    Dim l_strRecStatus As String

On Error GoTo eh
    r_rsPendingChanges.Filter = adFilterPendingRecords
    Do Until r_rsPendingChanges.EOF
        Select Case r_rsPendingChanges.Status
            Case adRecNew
                r_datStartDate = MinDate(r_rsPendingChanges.Fields("START_DATE").Value, r_datStartDate)
                r_datEndDate = MaxDate(r_rsPendingChanges.Fields("END_DATE").Value, r_datEndDate)
            Case adRecModified
                r_datStartDate = MinDate(r_rsPendingChanges.Fields("START_DATE").Value, r_datStartDate)
                r_datStartDate = MinDate(r_rsPendingChanges.Fields("START_DATE").OriginalValue, r_datStartDate)
                r_datEndDate = MaxDate(r_rsPendingChanges.Fields("END_DATE").Value, r_datEndDate)
                r_datEndDate = MaxDate(r_rsPendingChanges.Fields("END_DATE").OriginalValue, r_datEndDate)
            Case adRecDeleted
                r_datStartDate = MinDate(r_rsPendingChanges.Fields("START_DATE").OriginalValue, r_datStartDate)
                r_datEndDate = MaxDate(r_rsPendingChanges.Fields("END_DATE").OriginalValue, r_datEndDate)
        End Select
        r_rsPendingChanges.MoveNext
    Loop
    r_rsPendingChanges.Filter = adFilterNone
    
    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "DetermineMinMaxWGLine"
End Sub


Private Sub DetermineMinMaxWGBase( _
            ByRef r_datStartDate As Date, _
            ByRef r_datEndDate As Date, _
            ByRef r_rsPendingChanges As ADODB.Recordset)
'    This looks the same as xxxx.
'    If still the same after integration, rename it and both use it.
    Dim l_strRecStatus As String

On Error GoTo eh
    r_rsPendingChanges.Filter = adFilterPendingRecords
    Do Until r_rsPendingChanges.EOF
        Select Case r_rsPendingChanges.Status
            Case adRecNew
                r_datStartDate = MinDate(r_rsPendingChanges.Fields("EFF_DATE").Value, r_datStartDate)
                r_datEndDate = MaxDate(r_rsPendingChanges.Fields("END_DATE").Value, r_datEndDate)
            Case adRecModified
                r_datStartDate = MinDate(r_rsPendingChanges.Fields("EFF_DATE").Value, r_datStartDate)
                r_datStartDate = MinDate(r_rsPendingChanges.Fields("EFF_DATE").OriginalValue, r_datStartDate)
                r_datEndDate = MaxDate(r_rsPendingChanges.Fields("END_DATE").Value, r_datEndDate)
                r_datEndDate = MaxDate(r_rsPendingChanges.Fields("END_DATE").OriginalValue, r_datEndDate)
            Case adRecDeleted
                r_datStartDate = MinDate(r_rsPendingChanges.Fields("EFF_DATE").OriginalValue, r_datStartDate)
                r_datEndDate = MaxDate(r_rsPendingChanges.Fields("END_DATE").OriginalValue, r_datEndDate)
        End Select
        r_rsPendingChanges.MoveNext
    Loop
    r_rsPendingChanges.Filter = adFilterNone
    
    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "DetermineMinMaxWGBase"
End Sub


Private Sub DetermineMinMaxWGAssign( _
            ByRef r_datStartDate As Date, _
            ByRef r_datEndDate As Date, _
            ByRef r_rsPendingChanges As ADODB.Recordset)
'    This looks the same as xxxx.
'    If still the same after integration, rename it and both use it.
    Dim l_strRecStatus As String

On Error GoTo eh
    r_rsPendingChanges.Filter = adFilterPendingRecords
    Do Until r_rsPendingChanges.EOF
        Select Case r_rsPendingChanges.Status
            Case adRecNew
                r_datStartDate = MinDate(r_rsPendingChanges.Fields("EFF_START_DATE").Value, r_datStartDate)
                r_datEndDate = MaxDate(r_rsPendingChanges.Fields("EFF_END_DATE").Value, r_datEndDate)
            Case adRecModified
                r_datStartDate = MinDate(r_rsPendingChanges.Fields("EFF_START_DATE").Value, r_datStartDate)
                r_datStartDate = MinDate(r_rsPendingChanges.Fields("EFF_START_DATE").OriginalValue, r_datStartDate)
                r_datEndDate = MaxDate(r_rsPendingChanges.Fields("EFF_END_DATE").Value, r_datEndDate)
                r_datEndDate = MaxDate(r_rsPendingChanges.Fields("EFF_END_DATE").OriginalValue, r_datEndDate)
            Case adRecDeleted
                r_datStartDate = MinDate(r_rsPendingChanges.Fields("EFF_START_DATE").OriginalValue, r_datStartDate)
                r_datEndDate = MaxDate(r_rsPendingChanges.Fields("EFF_END_DATE").OriginalValue, r_datEndDate)
        End Select
        r_rsPendingChanges.MoveNext
    Loop
    r_rsPendingChanges.Filter = adFilterNone
    
    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "DetermineMinMaxWGAssign"
End Sub


Private Function DeterminePendingWorkgroupsInScope( _
                    ByRef r_rsWGBaseChanges As ADODB.Recordset, _
                    ByRef r_rsWGLineChanges As ADODB.Recordset, _
                    ByRef v_rsWGAssignChanges As ADODB.Recordset, _
                    ByRef v_rsWGAssignOtherChanges As ADODB.Recordset, _
                    ByRef v_rsEmpLineAssignChanges As ADODB.Recordset) As Variant
    Dim l_cuSet As utilTRISCommonLib.cuSet

On Error GoTo eh
    Set l_cuSet = New cuSet
    
    AppendFieldToSet "WG_NUM", r_rsWGBaseChanges, l_cuSet
    AppendFieldToSet "WG_NUM", r_rsWGLineChanges, l_cuSet
    AppendFieldToSet "WG_NUM", v_rsWGAssignChanges, l_cuSet
    AppendFieldToSet "WG_NUM", v_rsWGAssignOtherChanges, l_cuSet
    AppendDICFieldToSet "WG_NUM", "KEY_WG_NUM", v_rsEmpLineAssignChanges, l_cuSet

    DeterminePendingWorkgroupsInScope = ArrayFromCollection(l_cuSet.Contents)
    Set l_cuSet = Nothing
    Exit Function
    
eh:
    Set l_cuSet = Nothing
    TrsRaiseError m_strCMODULENAME, "DeterminePendingWorkgroupsInScope"
End Function


Private Function DeterminePendingDemandLinesInScope( _
                    ByRef r_rsWGLineChanges As ADODB.Recordset, _
                    ByRef v_rsEmpLineAssignChanges As ADODB.Recordset) As Variant
    Dim l_cuSet As utilTRISCommonLib.cuSet

On Error GoTo eh
    Set l_cuSet = New cuSet
    
    AppendFieldToSet "LINE_NUM", r_rsWGLineChanges, l_cuSet
    AppendDICFieldToSet "LINE_NUM", "KEY_LINE_NUM", v_rsEmpLineAssignChanges, l_cuSet

    DeterminePendingDemandLinesInScope = ArrayFromCollection(l_cuSet.Contents)
    Set l_cuSet = Nothing
    Exit Function
    
eh:
    Set l_cuSet = Nothing
    TrsRaiseError m_strCMODULENAME, "DeterminePendingDemandLinesInScope"
End Function


Private Function DetermineAllWorkgroupsInScope( _
                    ByRef r_aryPendingWorkgroups As Variant, _
                    ByRef r_rsWGEmpAssignments As ADODB.Recordset) As Variant
    Dim l_cuSet As utilTRISCommonLib.cuSet
    Dim i As Integer
    Dim l_strErrDesc As String
    Dim l_colErrArgs As Collection

On Error GoTo eh
    Set l_cuSet = New cuSet
    If IsArray(r_aryPendingWorkgroups) Then
        For i = LBound(r_aryPendingWorkgroups) To UBound(r_aryPendingWorkgroups)
            l_cuSet.Add r_aryPendingWorkgroups(i)
        Next
    Else
        Set l_colErrArgs = New Collection
        l_colErrArgs.Add AsString(r_aryPendingWorkgroups)
        l_strErrDesc = _
            TrsGetErrDescription(LongAppErr(trErrbusNoPendingWGDetected), l_colErrArgs)
        Set l_colErrArgs = Nothing
        Err.Raise trErrbusNoPendingWGDetected, m_strCMODULENAME, l_strErrDesc
    End If
    
    AppendFieldToSet "WG_NUM", r_rsWGEmpAssignments, l_cuSet, False
    
    DetermineAllWorkgroupsInScope = ArrayFromCollection(l_cuSet.Contents)
    Set l_cuSet = Nothing
    Exit Function
    
eh:
    Set l_cuSet = Nothing
    Set l_colErrArgs = Nothing
    TrsRaiseError m_strCMODULENAME, "DetermineAllWorkgroupsInScope"
End Function


Private Function AsString(ByRef r_aryPendingWorkgroups As Variant) As String
' Answer a string representation of r_aryPendingWorkgroups.
    Dim l_intVarType As Integer
    
On Error GoTo eh
    l_intVarType = VarType(r_aryPendingWorkgroups)
    Select Case l_intVarType
        Case vbNull
            AsString = "Null"
        Case vbEmpty
            AsString = "Empty"
        Case vbDate
            AsString = Format(r_aryPendingWorkgroups, "mm/dd/yyyy")
        Case vbObject
            AsString = TypeName(r_aryPendingWorkgroups)
        Case Else
            If IsArray(r_aryPendingWorkgroups) Then
                AsString = "an Array"
            Else
                AsString = CStr(r_aryPendingWorkgroups)
            End If
    End Select

    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "AsString"
End Function


Private Function AppendFieldToSet( _
                    ByVal v_strFieldName As String, _
                    ByRef r_rsPendingChanges As ADODB.Recordset, _
                    ByRef r_cuSet As utilTRISCommonLib.cuSet, _
                    Optional ByVal v_blnPendingOnly As Boolean = True)
' Assuming all types of r_rsPendingChanges recordsets have been modified after
' retrieved from database with batch optimistic, vs DIC recordstatus.

On Error GoTo eh
    
    If RecordsExist(r_rsPendingChanges) Then
        If v_blnPendingOnly Then
            r_rsPendingChanges.Filter = adFilterPendingRecords
        Else
            r_rsPendingChanges.Filter = adFilterNone
            r_rsPendingChanges.MoveFirst
        End If
        Do Until r_rsPendingChanges.EOF
            If r_rsPendingChanges.Status = adRecNew Then
                r_cuSet.Add r_rsPendingChanges.Fields(v_strFieldName).Value
            Else
                r_cuSet.Add r_rsPendingChanges.Fields(v_strFieldName).OriginalValue
            End If
            r_rsPendingChanges.MoveNext
        Loop
        r_rsPendingChanges.Filter = adFilterNone
        r_rsPendingChanges.MoveFirst
    End If
    
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "AppendFieldToSet"
End Function


Private Function AppendDICFieldToSet( _
                    ByVal v_strFieldName As String, _
                    ByVal v_strKeyFieldName As String, _
                    ByRef r_rsPendingChanges As ADODB.Recordset, _
                    ByRef r_cuSet As utilTRISCommonLib.cuSet)
' Assuming all types of r_rsPendingChanges recordsets have DIC recordstatus.

On Error GoTo eh
    
    If r_rsPendingChanges Is Nothing Then
        Exit Function
    End If
    
    r_rsPendingChanges.Filter = adFilterNone
    Do Until r_rsPendingChanges.EOF
        If Trim(r_rsPendingChanges!RECORD_STATUS) = "I" Then
            r_cuSet.Add r_rsPendingChanges.Fields(v_strFieldName).Value
        Else
            r_cuSet.Add r_rsPendingChanges.Fields(v_strKeyFieldName).Value
        End If
        r_rsPendingChanges.MoveNext
    Loop
    
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "AppendDICFieldToSet"
End Function

Private Function ApplyPendingChanges( _
            ByVal v_rsWGBaseSchedChanges As ADODB.Recordset, _
            ByVal v_rsWGLineSchedChanges As ADODB.Recordset, _
            ByVal v_rsWGAssignChanges As ADODB.Recordset, _
            ByVal v_rsWGAssignOtherChanges As ADODB.Recordset, _
            ByVal v_rsEmpLineAssignChanges As ADODB.Recordset, _
            ByVal v_rsPersonalChanges As ADODB.Recordset, _
            ByVal v_rsLineTradeChanges As ADODB.Recordset, _
            ByVal v_rsEmpTradeChanges As ADODB.Recordset)

On Error GoTo eh

    If Not v_rsWGBaseSchedChanges Is Nothing Then
        ApplyPendingWGBaseChanges v_rsWGBaseSchedChanges
    End If

    If Not v_rsWGLineSchedChanges Is Nothing Then
        ApplyPendingWGLineChanges v_rsWGLineSchedChanges
    End If
    
    If Not v_rsWGAssignChanges Is Nothing Then
        ApplyPendingWGAssignChanges v_rsWGAssignChanges
    End If

'    Other refers to workgroup assignments otherthan the workgroup
'    that is being maintained.  Apply pending changes to the same
'    employee workgroup assignments recordset; m_rsWGEmpAssignments.
    If Not v_rsWGAssignOtherChanges Is Nothing Then
        ApplyPendingWGAssignChanges v_rsWGAssignOtherChanges
    End If
    
    If m_rsWGBaseSchedule Is Nothing Then
        Set m_rsWGBaseSchedule = _
            JoinEMP_WG_BASE_SCHED(m_rsWGEmpAssignments, m_rsWGBaseScheduleOnly)
    End If
    
    If Not v_rsEmpLineAssignChanges Is Nothing Then
        ApplyPendingEmpLineAssignChanges v_rsEmpLineAssignChanges
    End If
    
    If Not v_rsPersonalChanges Is Nothing Then
        ApplyPendingPersonalChanges v_rsPersonalChanges
    End If
    
    If (Not v_rsLineTradeChanges Is Nothing) And (Not m_rsDemandLineTrades Is Nothing) Then
        ApplyPendingLineTradeChanges v_rsLineTradeChanges
    End If
    
    If (Not v_rsEmpTradeChanges Is Nothing) And (Not m_rsEmployeeTrades Is Nothing) Then
        ApplyPendingEmpTradeChanges v_rsEmpTradeChanges
    End If

    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "ApplyPendingChanges"
End Function


Private Function ApplyPendingWGBaseChanges( _
            ByVal v_rsPendingChanges As ADODB.Recordset)
'    Can only apply changes to m_rsWGBaseScheduleOnly, not joined
'    recordset; m_rsWGBaseSchedule.
'    Dim l_strRecStatus As String

On Error GoTo eh

    If m_rsWGBaseScheduleOnly Is Nothing Then
        Set m_rsWGBaseScheduleOnly = v_rsPendingChanges
        Exit Function
    End If
    
    v_rsPendingChanges.Filter = adFilterPendingRecords
    Do Until v_rsPendingChanges.EOF
'        l_strRecStatus = Trim(v_rsPendingChanges!RECORD_STATUS)
        Select Case v_rsPendingChanges.Status
            Case adRecNew
                With m_rsWGBaseScheduleOnly
                    .AddNew
                    !WG_NUM = v_rsPendingChanges!WG_NUM
                    !EFF_DATE = CDate(v_rsPendingChanges!EFF_DATE)
                    !END_DATE = CDate(v_rsPendingChanges!END_DATE)
                    !ROT_CD = v_rsPendingChanges!ROT_CD
                    !ROT_WEEK = v_rsPendingChanges!ROT_WEEK
                    !SCHED_CD = v_rsPendingChanges!SCHED_CD
                    .Update
                End With
            Case adRecModified
                FindWGBaseSchedRecord m_rsWGBaseScheduleOnly, v_rsPendingChanges
                If Not m_rsWGBaseScheduleOnly.EOF Then
                    With m_rsWGBaseScheduleOnly
                        !EFF_DATE = CDate(v_rsPendingChanges!EFF_DATE)
                        !END_DATE = CDate(v_rsPendingChanges!END_DATE)
                        !ROT_CD = v_rsPendingChanges!ROT_CD
                        !ROT_WEEK = v_rsPendingChanges!ROT_WEEK
                        !SCHED_CD = v_rsPendingChanges!SCHED_CD
                        .Update
                    End With
                End If
            Case adRecDeleted
                FindWGBaseSchedRecord m_rsWGBaseScheduleOnly, v_rsPendingChanges
                If Not m_rsWGBaseScheduleOnly.EOF Then
                    m_rsWGBaseScheduleOnly.Delete
                End If
        End Select
        v_rsPendingChanges.MoveNext
    Loop
    v_rsPendingChanges.Filter = adFilterNone
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "ApplyPendingWGBaseChanges"
End Function


Private Function ApplyPendingWGLineChanges( _
            ByVal v_rsPendingChanges As ADODB.Recordset)
'    Dim l_strRecStatus As String

On Error GoTo eh

    If m_rsDemandLineBaseSchedule Is Nothing Then
        Set m_rsDemandLineBaseSchedule = v_rsPendingChanges
        Exit Function
    End If

    v_rsPendingChanges.Filter = adFilterPendingRecords
'    Process deletes first, this will ensure delete is processed in case of a change
'    that updates the start_date or end_date.

'
' 5/24/02 RSG This should not be needed if the created_timestamp is available.
' Verify in testing and remove if unnecessary.

    Do Until v_rsPendingChanges.EOF
'        l_strRecStatus = Trim(v_rsPendingChanges!RECORD_STATUS)
        Select Case v_rsPendingChanges.Status
            Case adRecDeleted
                FindLineSchedRecord m_rsDemandLineBaseSchedule, v_rsPendingChanges
                If Not m_rsDemandLineBaseSchedule.EOF Then
                    m_rsDemandLineBaseSchedule.Delete
                End If
        End Select
        v_rsPendingChanges.MoveNext
    Loop
    
    v_rsPendingChanges.MoveFirst
    Do Until v_rsPendingChanges.EOF
'        l_strRecStatus = Trim(v_rsPendingChanges!RECORD_STATUS)
        Select Case v_rsPendingChanges.Status
            Case adRecNew
                With m_rsDemandLineBaseSchedule
                    .AddNew
                    !WG_NUM = v_rsPendingChanges!WG_NUM
                    !LINE_NUM = v_rsPendingChanges!LINE_NUM
                    !START_DATE = CDate(v_rsPendingChanges!START_DATE)
                    !END_DATE = CDate(v_rsPendingChanges!END_DATE)
                    !ROT_CD = v_rsPendingChanges!ROT_CD
                    !ROT_WEEK = v_rsPendingChanges!ROT_WEEK
                    !SCHED_CD = v_rsPendingChanges!SCHED_CD
                    !JOB_NUM = v_rsPendingChanges!JOB_NUM
                    !REQ_LINE_IND = v_rsPendingChanges!REQ_LINE_IND
                    .Update
                End With
            Case adRecModified
                FindLineSchedRecord m_rsDemandLineBaseSchedule, v_rsPendingChanges
                If Not m_rsDemandLineBaseSchedule.EOF Then
                    With m_rsDemandLineBaseSchedule
                        !START_DATE = CDate(v_rsPendingChanges!START_DATE)
                        !END_DATE = CDate(v_rsPendingChanges!END_DATE)
                        !ROT_CD = v_rsPendingChanges!ROT_CD
                        !ROT_WEEK = v_rsPendingChanges!ROT_WEEK
                        !SCHED_CD = v_rsPendingChanges!SCHED_CD
                        !JOB_NUM = v_rsPendingChanges!JOB_NUM
                        !REQ_LINE_IND = v_rsPendingChanges!REQ_LINE_IND
                        .Update
                    End With
                End If
        End Select
        v_rsPendingChanges.MoveNext
    Loop

    v_rsPendingChanges.Filter = adFilterNone
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "ApplyPendingWGLineChanges"
End Function


Private Function ApplyPendingWGAssignChanges( _
            ByVal v_rsPendingChanges As ADODB.Recordset)
'    Can only apply changes to m_rsWGEmpAssignments, not joined
'    recordset; m_rsWGBaseSchedule.
' 9/24/02 RSG
' Must support changing employee assignments primary indicator.
' The persistent employee assignments retrieved are primary assignments only.
' Therefore, when changing a secondary to primary, m_rsWGEmpAssignments will not
' include the persistent record and must be added.
' When changing a primary to secondary, m_rsWGEmpAssignments will include the
' persistent record, but is no longer pertinent to the scheduled and must be deleted.

On Error GoTo eh

    If m_rsWGEmpAssignments Is Nothing Then
        Set m_rsWGEmpAssignments = v_rsPendingChanges
        Exit Function
    End If
    
    v_rsPendingChanges.Filter = adFilterPendingRecords
    Do Until v_rsPendingChanges.EOF
        Select Case v_rsPendingChanges.Status
            Case adRecNew
                With m_rsWGEmpAssignments
                    .AddNew
                    !EMP_NUM = v_rsPendingChanges!EMP_NUM
                    !WG_NUM = v_rsPendingChanges!WG_NUM
                    !PRIMARY_IND = v_rsPendingChanges!PRIMARY_IND
                    !EFF_START_DATE = CDate(v_rsPendingChanges!EFF_START_DATE)
                    !EFF_END_DATE = CDate(v_rsPendingChanges!EFF_END_DATE)
                    .Update
                End With
            Case adRecModified
                FindWGEmpAssignRecord m_rsWGEmpAssignments, v_rsPendingChanges
                If Not m_rsWGEmpAssignments.EOF Then
                    If ChangedPrimaryToSecondary(v_rsPendingChanges) Then
                        m_rsWGEmpAssignments.Delete
                    Else
                        With m_rsWGEmpAssignments
                            !EFF_START_DATE = CDate(v_rsPendingChanges!EFF_START_DATE)
                            !EFF_END_DATE = CDate(v_rsPendingChanges!EFF_END_DATE)
                            .Update
                        End With
                    End If
                ElseIf ChangedSecondaryToPrimary(v_rsPendingChanges) Then
                    With m_rsWGEmpAssignments
                        .AddNew
                        !EMP_NUM = v_rsPendingChanges!EMP_NUM
                        !WG_NUM = v_rsPendingChanges!WG_NUM
                        !PRIMARY_IND = v_rsPendingChanges!PRIMARY_IND
                        !EFF_START_DATE = CDate(v_rsPendingChanges!EFF_START_DATE)
                        !EFF_END_DATE = CDate(v_rsPendingChanges!EFF_END_DATE)
                        .Update
                    End With
                End If
            Case adRecDeleted
                FindWGEmpAssignRecord m_rsWGEmpAssignments, v_rsPendingChanges
                If Not m_rsWGEmpAssignments.EOF Then
                    m_rsWGEmpAssignments.Delete
                End If
        End Select
        v_rsPendingChanges.MoveNext
    Loop
    v_rsPendingChanges.Filter = adFilterNone

    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "ApplyPendingWGAssignChanges"
End Function


Private Function ChangedSecondaryToPrimary( _
                    ByVal v_rsPendingChanges As ADODB.Recordset) As Boolean
    Dim l_strOrigPrimaryInd As String
    
On Error GoTo eh

    l_strOrigPrimaryInd = Trim(v_rsPendingChanges.Fields("PRIMARY_IND").OriginalValue)
    If l_strOrigPrimaryInd = "N" Then
        ChangedSecondaryToPrimary = _
            Trim(v_rsPendingChanges.Fields("PRIMARY_IND").Value) = "Y"
    End If

    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "ChangedSecondaryToPrimary"
End Function


Private Function ChangedPrimaryToSecondary( _
                    ByVal v_rsPendingChanges As ADODB.Recordset) As Boolean
    Dim l_strOrigPrimaryInd As String
    
On Error GoTo eh

    l_strOrigPrimaryInd = Trim(v_rsPendingChanges.Fields("PRIMARY_IND").OriginalValue)
    If l_strOrigPrimaryInd = "Y" Then
        ChangedPrimaryToSecondary = _
            Trim(v_rsPendingChanges.Fields("PRIMARY_IND").Value) = "N"
    End If

    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "ChangedPrimaryToSecondary"
End Function


Private Function ApplyPendingEmpLineAssignChanges( _
            ByVal v_rsPendingChanges As ADODB.Recordset)
            
    Dim l_strRecStatus As String
On Error GoTo eh

    If m_rsDemandLineAssignments Is Nothing Then
        Set m_rsDemandLineAssignments = v_rsPendingChanges
        Exit Function
    End If
    
    v_rsPendingChanges.Filter = adFilterPendingRecords
'    Process deletes first, this will ensure delete is processed in case of a change
'    that updates the start_date or end_date.
    Do Until v_rsPendingChanges.EOF
        l_strRecStatus = Trim(v_rsPendingChanges!RECORD_STATUS)
        Select Case l_strRecStatus
            Case "D"
                FindDemandSchedRecord m_rsDemandLineAssignments, v_rsPendingChanges
                If Not m_rsDemandLineAssignments.EOF Then
                    m_rsDemandLineAssignments.Delete
                End If
        End Select
        v_rsPendingChanges.MoveNext
    Loop
    
    v_rsPendingChanges.MoveFirst
    Do Until v_rsPendingChanges.EOF
        l_strRecStatus = Trim(v_rsPendingChanges!RECORD_STATUS)
        Select Case l_strRecStatus
            Case "I"
                With m_rsDemandLineAssignments
                    .AddNew
                    !WG_NUM = v_rsPendingChanges!WG_NUM
                    !LINE_NUM = v_rsPendingChanges!LINE_NUM
                    !START_DATE = CDate(v_rsPendingChanges!START_DATE)
                    !END_DATE = CDate(v_rsPendingChanges!END_DATE)
                    !EMP_NUM = v_rsPendingChanges!EMP_NUM
                    !RELIEF_COMP = v_rsPendingChanges!RELIEF_COMP
                    .Update
                End With
            Case "C"
                FindDemandSchedRecord m_rsDemandLineAssignments, v_rsPendingChanges
                If Not m_rsDemandLineAssignments.EOF Then
                    With m_rsDemandLineAssignments
                        !RELIEF_COMP = v_rsPendingChanges!RELIEF_COMP
                        .Update
                    End With
                End If
        End Select
        v_rsPendingChanges.MoveNext
    Loop

    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "ApplyPendingEmpLineAssignChanges"
End Function


Private Function ApplyPendingPersonalChanges( _
            ByVal v_rsPendingChanges As ADODB.Recordset)
    Dim l_strRecStatus As String
            
On Error GoTo eh
    If m_rsPersonalSchedule Is Nothing Then
        Set m_rsPersonalSchedule = v_rsPendingChanges
        Exit Function
    End If

    v_rsPendingChanges.Filter = adFilterPendingRecords
'    Process deletes first, this will ensure delete is processed in case of a change
'    that updates the start_date or end_date.
    Do Until v_rsPendingChanges.EOF
        l_strRecStatus = Trim(v_rsPendingChanges!RECORD_STATUS)
        Select Case l_strRecStatus
            Case "D"
                FindPersonalSchedRecord m_rsPersonalSchedule, v_rsPendingChanges
                If Not m_rsPersonalSchedule.EOF Then
                    m_rsPersonalSchedule.Delete
                End If
        End Select
        v_rsPendingChanges.MoveNext
    Loop
    
    v_rsPendingChanges.MoveFirst
    Do Until v_rsPendingChanges.EOF
        l_strRecStatus = Trim(v_rsPendingChanges!RECORD_STATUS)
        Select Case l_strRecStatus
            Case "I"
                With m_rsPersonalSchedule
                    .AddNew
                    !EMP_NUM = v_rsPendingChanges!EMP_NUM
                    !START_DATE = v_rsPendingChanges!START_DATE
                    !END_DATE = v_rsPendingChanges!END_DATE
                    !ROT_CD = v_rsPendingChanges!ROT_CD
                    !ROT_WEEK = v_rsPendingChanges!ROT_WEEK
                    !SCHED_CD = v_rsPendingChanges!SCHED_CD
                    .Update
                End With
            Case "C"
                FindPersonalSchedRecord m_rsPersonalSchedule, v_rsPendingChanges
                If Not m_rsPersonalSchedule.EOF Then
                    With m_rsPersonalSchedule
                        !START_DATE = v_rsPendingChanges!START_DATE
                        !END_DATE = v_rsPendingChanges!END_DATE
                        !ROT_CD = v_rsPendingChanges!ROT_CD
                        !ROT_WEEK = v_rsPendingChanges!ROT_WEEK
                        !SCHED_CD = v_rsPendingChanges!SCHED_CD
                        .Update
                    End With
                End If
        End Select
        v_rsPendingChanges.MoveNext
    Loop

    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "ApplyPendingPersonalChanges"
End Function


Private Function JoinEMP_WG_BASE_SCHED( _
                    ByRef r_rsEMP_WG As ADODB.Recordset, _
                    ByRef r_rsBASE_SCHED As ADODB.Recordset) _
                As ADODB.Recordset
                
    Dim l_rsJoined As ADODB.Recordset
    Dim l_strFilter As String
    
    Set l_rsJoined = New Recordset
    With l_rsJoined
        .Fields.Append "WG_NUM", adInteger, , mCFLDATTRUPDATE
        .Fields.Append "EFF_START_DATE", adDate, , mCFLDATTRUPDATE
        .Fields.Append "EFF_END_DATE", adDate, , mCFLDATTRUPDATE
        .Fields.Append "SCHED_CD", adInteger, , mCFLDATTRUPDATE
        .Fields.Append "EFF_DATE", adDate, , mCFLDATTRUPDATE
        .Fields.Append "EMP_NUM", adInteger, , mCFLDATTRUPDATE
        .Fields.Append "ROT_CD", adInteger, , mCFLDATTRUPDATE
        .Fields.Append "ROT_WEEK", adInteger, , mCFLDATTRUPDATE
        .Fields.Append "END_DATE", adDate, , mCFLDATTRUPDATE
        .Open
    End With
                
'    inner join will answer empty recordset when either source is empty
    If (Not RecordsExist(r_rsEMP_WG)) Or _
        (Not RecordsExist(r_rsBASE_SCHED)) Then
        Set JoinEMP_WG_BASE_SCHED = l_rsJoined
        Set l_rsJoined = Nothing
        Exit Function
    End If

'    r_rsEMP_WG is the driver
    r_rsEMP_WG.MoveFirst
    Do Until r_rsEMP_WG.EOF
        l_strFilter = "WG_NUM = " & r_rsEMP_WG!WG_NUM _
            & " AND EFF_DATE <= #" & Format(r_rsEMP_WG!EFF_END_DATE, "yyyy-mm-dd") & "#" _
            & " AND END_DATE >= #" & Format(r_rsEMP_WG!EFF_START_DATE, "yyyy-mm-dd") & "#"
        r_rsBASE_SCHED.Filter = l_strFilter
        If RecordsExist(r_rsBASE_SCHED) Then
            Do Until r_rsBASE_SCHED.EOF
                With l_rsJoined
                    .AddNew
                    !WG_NUM = r_rsEMP_WG!WG_NUM
                    !EMP_NUM = r_rsEMP_WG!EMP_NUM
                    !EFF_START_DATE = CDate(r_rsEMP_WG!EFF_START_DATE)
                    !EFF_END_DATE = CDate(r_rsEMP_WG!EFF_END_DATE)
                    !EFF_DATE = CDate(r_rsBASE_SCHED!EFF_DATE)
                    !END_DATE = CDate(r_rsBASE_SCHED!END_DATE)
                    !SCHED_CD = r_rsBASE_SCHED!SCHED_CD
                    !ROT_CD = r_rsBASE_SCHED!ROT_CD
                    !ROT_WEEK = r_rsBASE_SCHED!ROT_WEEK
                    .Update
                    r_rsBASE_SCHED.MoveNext
                End With
            Loop
        End If
        r_rsEMP_WG.MoveNext
    Loop
    
    Set JoinEMP_WG_BASE_SCHED = l_rsJoined
    Set l_rsJoined = Nothing
    Exit Function
    
eh:
    Set l_rsJoined = Nothing
    TrsRaiseError m_strCMODULENAME, "JoinEMP_WG_BASE_SCHED"
End Function


Private Function FindWGBaseSchedRecord( _
                    ByVal v_rsWGBaseSchedule As ADODB.Recordset, _
                    ByVal v_rsPendingChanges As ADODB.Recordset)

On Error GoTo eh

    If RecordsExist(v_rsWGBaseSchedule) Then
        v_rsWGBaseSchedule.MoveFirst
    End If
    
    Do Until v_rsWGBaseSchedule.EOF
        If v_rsWGBaseSchedule!WG_NUM = v_rsPendingChanges.Fields("WG_NUM").OriginalValue And _
            v_rsWGBaseSchedule!EFF_DATE = CDate(v_rsPendingChanges.Fields("EFF_DATE").OriginalValue) And _
            Trim(v_rsWGBaseSchedule!CREATED_TIMESTAMP) = Trim(v_rsPendingChanges.Fields("CREATED_TIMESTAMP").OriginalValue) Then
            Exit Do
        End If
        v_rsWGBaseSchedule.MoveNext
    Loop

    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "FindWGBaseSchedRecord"
End Function


Private Function FindLineSchedRecord( _
                    ByVal v_rsLineSchedule As ADODB.Recordset, _
                    ByVal v_rsPendingChanges As ADODB.Recordset)

On Error GoTo eh

    If RecordsExist(v_rsLineSchedule) Then
        v_rsLineSchedule.MoveFirst
    End If
    
    Do Until v_rsLineSchedule.EOF
        If v_rsLineSchedule!WG_NUM = v_rsPendingChanges.Fields("WG_NUM").OriginalValue And _
            v_rsLineSchedule!LINE_NUM = v_rsPendingChanges.Fields("LINE_NUM").OriginalValue And _
            v_rsLineSchedule!START_DATE = CDate(v_rsPendingChanges.Fields("START_DATE").OriginalValue) Then
'            Trim(v_rsLineSchedule!CREATED_TIMESTAMP) = Trim(v_rsPendingChanges.Fields("CREATED_TIMESTAMP").OriginalValue) Then
            Exit Do
        End If
        v_rsLineSchedule.MoveNext
    Loop

    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "FindLineSchedRecord"
End Function


Private Function FindWGEmpAssignRecord( _
                    ByVal v_rsWGEmpAssignments As ADODB.Recordset, _
                    ByVal v_rsPendingChanges As ADODB.Recordset)

On Error GoTo eh

    If RecordsExist(v_rsWGEmpAssignments) Then
        v_rsWGEmpAssignments.MoveFirst
    End If
    
    Do Until v_rsWGEmpAssignments.EOF
        If v_rsWGEmpAssignments!EMP_NUM = v_rsPendingChanges.Fields("EMP_NUM").OriginalValue And _
            v_rsWGEmpAssignments!WG_NUM = v_rsPendingChanges.Fields("WG_NUM").OriginalValue And _
            Trim(v_rsWGEmpAssignments!CREATED_TIMESTAMP) = Trim(v_rsPendingChanges.Fields("CREATED_TIMESTAMP").OriginalValue) Then
            Exit Do
        End If
        v_rsWGEmpAssignments.MoveNext
    Loop

    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "FindWGEmpAssignRecord"
End Function


Private Function FindDemandSchedRecord( _
                    ByVal v_rsDemandSchedule As ADODB.Recordset, _
                    ByVal v_rsPendingChanges As ADODB.Recordset)

On Error GoTo eh

    If RecordsExist(v_rsDemandSchedule) Then
        v_rsDemandSchedule.MoveFirst
    End If
    
    Do Until v_rsDemandSchedule.EOF
        If v_rsDemandSchedule!EMP_NUM = v_rsPendingChanges!KEY_EMP_NUM And _
            v_rsDemandSchedule!START_DATE = CDate(v_rsPendingChanges!KEY_START_DATE) And _
            v_rsDemandSchedule!END_DATE = CDate(v_rsPendingChanges!KEY_END_DATE) Then
            Exit Do
        End If
        v_rsDemandSchedule.MoveNext
    Loop

    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "FindDemandSchedRecord"
End Function


Private Function FindPersonalSchedRecord( _
                    ByVal v_rsPersonalSchedule As ADODB.Recordset, _
                    ByVal v_rsPendingChanges As ADODB.Recordset)

On Error GoTo eh

    If RecordsExist(v_rsPersonalSchedule) Then
        v_rsPersonalSchedule.MoveFirst
    End If
    Do Until v_rsPersonalSchedule.EOF
        If v_rsPersonalSchedule!EMP_NUM = v_rsPendingChanges!KEY_EMP_NUM And _
            v_rsPersonalSchedule!START_DATE = CDate(v_rsPendingChanges!KEY_START_DATE) And _
            v_rsPersonalSchedule!END_DATE = CDate(v_rsPendingChanges!KEY_END_DATE) Then
            Exit Do
        End If
        v_rsPersonalSchedule.MoveNext
    Loop

    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "FindPersonalSchedRecord"
End Function


Private Function ApplyPendingLineTradeChanges( _
            ByVal v_rsPendingChanges As ADODB.Recordset)
            
On Error GoTo eh


    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "ApplyPendingLineTradeChanges"
End Function


Private Function ApplyPendingEmpTradeChanges( _
            ByVal v_rsPendingChanges As ADODB.Recordset)
            
On Error GoTo eh


    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "ApplyPendingEmpTradeChanges"
End Function


Private Function GetEmployeeData( _
    ByVal v_aryEmps As Variant _
    ) As ADODB.Recordset

    Dim l_obj As dbEmployee.cdEmployee
    
    
    On Error GoTo eh

    Set l_obj = CreateObject("dbEmployee.cdEmployee")

    Set GetEmployeeData = l_obj.GetEmployeeExcSchedInfo(v_aryEmps)

    Set l_obj = Nothing

    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "GetEmployeeData"
    
End Function

Private Function GetEffectiveJob( _
    ByVal v_lngEmpNum As Long, _
    ByVal v_datChargeDate As Date _
    ) As String
    
    
    Dim l_datJobEffDate As Date
    Dim l_strJobNum As String
    Dim l_strPriorJobNum As String
    Dim l_datPriorJobEffDate As Date
    Dim l_obj As busEmployee.cbsEmployee
    
    On Error GoTo eh
    
    If m_rsEmployeeData Is Nothing Then
        GetEffectiveJob = ""
        Exit Function
    End If
    
    m_rsEmployeeData.Filter = "EMP_NUM = " & str(v_lngEmpNum)
    If Not RecordsExist(m_rsEmployeeData) Then
        GetEffectiveJob = ""
        Exit Function
    End If
    
    l_datJobEffDate = m_rsEmployeeData!JOB_EFF_DATE
    l_strJobNum = m_rsEmployeeData!JOB_NUM
    l_strPriorJobNum = m_rsEmployeeData!PRIOR_JOB_NUM
    l_datPriorJobEffDate = m_rsEmployeeData!PRIOR_JOB_EFF_DT
    
    Set l_obj = CreateObject("busEmployee.cbsEmployee")

    GetEffectiveJob = l_obj.EmployeeJobNum( _
        l_datJobEffDate, l_strJobNum, l_strPriorJobNum, l_datPriorJobEffDate, v_datChargeDate)

    Set l_obj = Nothing
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "GetEffectiveJob"
    
    
End Function

Private Function DetermineTradeExceptionCode( _
    ByVal v_strTradeCodes As String, _
    ByVal v_decShiftHours As Variant _
    ) As String
    
  Dim l_strRoleType As String
  Dim l_strWeekTradeInd As String
    
  On Error GoTo eh
  
    If Trim(v_strTradeCodes) = "" Then
        DetermineTradeExceptionCode = ""
        Exit Function
    End If
  
    l_strRoleType = Left(v_strTradeCodes, 1)
    l_strWeekTradeInd = Right(v_strTradeCodes, 1)
    
    Select Case l_strRoleType
        Case "S"
            DetermineTradeExceptionCode = "TWS"
        Case "A"
            If v_decShiftHours > 0 Then
                DetermineTradeExceptionCode = "TSA"
            Else
                DetermineTradeExceptionCode = "TOA"
            End If
        Case "R"
            If v_decShiftHours > 0 Then
                DetermineTradeExceptionCode = "TSR"
            Else
                If l_strWeekTradeInd = "Y" Then
                    DetermineTradeExceptionCode = "RWT"
                Else
                    DetermineTradeExceptionCode = "TOR"
                End If
            End If
  
  End Select
  
  Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "DetermineTradeExceptionCode"
    
        
    
End Function


Private Function GetScheduleTimeStamps( _
    ByVal v_aryEmployees As Variant _
    ) As ADODB.Recordset
    
    On Error GoTo eh
    
    
    Dim l_dbSchedule As dbSchedule.cdSchedule
         
    
    Set l_dbSchedule = CreateObject("dbSchedule.cdSchedule")
    
    Set GetScheduleTimeStamps = _
        l_dbSchedule.GetScheduleTimeStamps(v_aryEmployees)
       
    Set l_dbSchedule = Nothing
    
    Exit Function
    
eh:
    Set l_dbSchedule = Nothing
    TrsRaiseError m_strCMODULENAME, "GetScheduleTimeStamps"
    
End Function

Private Function ScheduleTimeStamp( _
    ByVal v_lngEmpNum As Long _
    ) As String
    
    On Error GoTo eh
    
    If m_rsSchedTimeStamps Is Nothing Then
        ScheduleTimeStamp = ""
        Exit Function
    End If
    
    m_rsSchedTimeStamps.Filter = "EMP_NUM = " & str(v_lngEmpNum)
    If RecordsExist(m_rsSchedTimeStamps) Then
        m_rsSchedTimeStamps.MoveFirst
        ScheduleTimeStamp = Trim(m_rsSchedTimeStamps!LAST_MOD_SCHED_TM)
    Else
        ScheduleTimeStamp = ""
    End If
        
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "ScheduleTimeStamp"
        
End Function


Public Function GetRotationScheduleCode( _
            ByVal v_datChargeDate As Date, _
            ByVal v_intRotationCode As Integer, _
            ByVal v_datEffDate As Date, _
            ByVal v_intStartPosition As Integer, _
            ByRef r_intRecordPosition As Integer _
            ) As Integer
            
    Dim l_colRotationCodes As Collection
    
    On Error GoTo eh
            
    Set l_colRotationCodes = New Collection
    l_colRotationCodes.Add v_intRotationCode
    Set m_rsRotations = GetRotationSchedulesDB(l_colRotationCodes)
    Set l_colRotationCodes = Nothing
    
    GetRotationScheduleCode = DetermineRotationScheduleCode( _
        v_datChargeDate, v_intRotationCode, v_datEffDate _
        , v_intStartPosition, r_intRecordPosition)
    
    Set m_rsRotations = Nothing
    Exit Function
    
eh:
    Set m_rsRotations = Nothing
    TrsRaiseError m_strCMODULENAME, "GetRotationScheduleCode"
        
            

End Function


Public Function GetAdjustedBaseDepartmentSchedules( _
            ByVal v_aryDepartments As Variant, _
            ByVal v_datStartDate As Date, _
            ByVal v_datEndDate As Date, _
            Optional ByVal v_ReturnOption As TRSReturnOption = trAdjustedBaseSched, _
            Optional ByVal v_ScheduleSourceStartLevel As TRSSchedSourceStartLevel = trAllScheduleLevels _
            ) As ADODB.Recordset
          
    On Error GoTo GetAdjustedBaseDepartmentSchedules_EH
          
    Dim l_rsGetAdjBaseSched As ADODB.Recordset
    Dim l_rsGetAdjBaseLineSched As ADODB.Recordset
    Dim l_rsSchedCodesByChargeDate As ADODB.Recordset
    Dim l_rsScheduleWithoutExceptions As ADODB.Recordset
    Dim l_rsEmptyLines As ADODB.Recordset
    Dim l_aryDepartments As Variant
    Dim l_intReturnOption As Integer
    Dim l_intScheduleSourceStartLevel As Integer

    l_intReturnOption = v_ReturnOption
    l_intScheduleSourceStartLevel = v_ScheduleSourceStartLevel
    ValidateRequest 0, v_datStartDate, v_datEndDate, l_intReturnOption, l_intScheduleSourceStartLevel
    InitializeRecordsets l_intReturnOption, False, l_rsGetAdjBaseSched, l_rsScheduleWithoutExceptions
    
    If IsArray(v_aryDepartments) Then
        l_aryDepartments = v_aryDepartments
    Else
        ReDim l_aryDepartments(0) As String
        l_aryDepartments(0) = Trim(v_aryDepartments)
    End If
    
    m_aryEmployees = GetEmployeesInDepartment(l_aryDepartments, v_datStartDate, v_datEndDate)
        
' Determine Employees on lines Schedule by Charge Date
    If IsArray(m_aryEmployees) Then
        GetEmployeeScheduleData _
            m_aryEmployees, v_datStartDate, v_datEndDate, _
            l_intReturnOption, l_intScheduleSourceStartLevel
    Set m_rsRotations = GetRotationSchedules()
        GetAdjustedBaseSchedulePrivate _
            m_aryEmployees, v_datStartDate, v_datEndDate, Nothing, _
            l_rsGetAdjBaseSched, l_rsScheduleWithoutExceptions, False, _
            l_intReturnOption, l_intScheduleSourceStartLevel
    End If
    
    Set GetAdjustedBaseDepartmentSchedules = l_rsGetAdjBaseSched
    
    ClearGlobalRecordSets
    Set l_rsGetAdjBaseSched = Nothing
    Set l_rsScheduleWithoutExceptions = Nothing
    Set l_rsSchedCodesByChargeDate = Nothing
    Exit Function
    
GetAdjustedBaseDepartmentSchedules_EH:
    ClearGlobalRecordSets
    Set l_rsGetAdjBaseSched = Nothing
    Set l_rsScheduleWithoutExceptions = Nothing
    Set l_rsSchedCodesByChargeDate = Nothing
    Dim l_colDescArgs As Collection
    Set l_colDescArgs = New Collection
    l_colDescArgs.Add str(v_datStartDate)
    l_colDescArgs.Add str(v_datEndDate)
    Select Case Err.Number
        Case trErrbusGetAdjustedBaseSchedInvalidDate
            TrsRaiseError m_strCMODULENAME, "GetAdjustedBaseDepartmentSchedules", Err.Number, False, l_colDescArgs, , True
        Case Else
            TrsRaiseError m_strCMODULENAME, "GetAdjustedBaseDepartmentSchedules", , , , , True
    End Select

End Function

Private Function GetEmployeesInDepartment( _
    ByVal v_aryDepartments As Variant, _
    ByVal v_datStart As Date, _
    ByVal v_datEnd As Date _
    ) As Variant
    
    Dim l_obj As dbEmployee.cdEmployee
    Dim l_rs As ADODB.Recordset
    
    
    On Error GoTo eh
    
    Set l_obj = CreateObject("dbEmployee.cdEmployee")
    Set l_rs = l_obj.GetEmployeesInDepartment(v_aryDepartments, True, v_datStart, v_datEnd)
    Set l_obj = Nothing
    
    GetEmployeesInDepartment = GetEmployees(l_rs, Nothing)
    Set l_rs = Nothing
    
    Exit Function
    
eh:
    Set l_obj = Nothing
    Set l_rs = Nothing
    TrsRaiseError m_strCMODULENAME, "GetEmployeesInDepartment"
     
    
    
End Function

'//for unit testing
Public Function KeepsShiftComponentTest( _
    ByVal v_lngEmpNum As Long, _
    ByVal v_datChargeDate As Date _
    ) As Boolean
    
On Error GoTo eh

    Set m_rsEmployeeData = GetEmployeeData(v_lngEmpNum)

    KeepsShiftComponentTest = KeepsShiftComponent(v_lngEmpNum, v_datChargeDate)
    
    Set m_rsEmployeeData = Nothing
    Exit Function
eh:
    Set m_rsEmployeeData = Nothing
    TrsRaiseError m_strCMODULENAME, "KeepsShiftComponentTest"
End Function

Private Function KeepsShiftComponent( _
    ByVal v_lngEmpNum As Long, _
    ByVal v_datChargeDate As Date _
    ) As Boolean

'It was originally coded that all employees keep their shift component
'from their schedule prior to a trade
'For example employee on a PM shift trades for an AM shift,
'  the employee keeps his PM shift Component

' This is not true for AmerenCIPS and 1455 West (TA) employees.
'They get the shift component based on shift worked.
'In the example above, they would lose their PM component
' as a result of the trade.
    
    Dim l_datJobEffDate As Date
    Dim l_strUnionCode As String
    Dim l_strPriorUnionCode As String
    Dim l_datDeptEffDate As Date
    Dim l_strDeptNum As String
    Dim l_strPriorDeptNum As String
    Dim l_strCompanyCode As String
    Dim l_strPriorCompanyCode  As String
    Dim l_strPayrollGroupCd As String

    On Error GoTo eh

    If m_rsEmployeeData Is Nothing Then
        KeepsShiftComponent = True
        Exit Function
    End If
    
    m_rsEmployeeData.Filter = "EMP_NUM = " & str(v_lngEmpNum)
    If Not RecordsExist(m_rsEmployeeData) Then
        KeepsShiftComponent = True
        Exit Function
    End If
    
    l_datJobEffDate = m_rsEmployeeData!JOB_EFF_DATE
    l_strUnionCode = m_rsEmployeeData!UNION_CD
    l_strPriorUnionCode = m_rsEmployeeData!PRIOR_UNION_CD
    l_datDeptEffDate = m_rsEmployeeData!DEPT_EFF_DATE
    l_strDeptNum = m_rsEmployeeData!DEPT_NUM
    l_strPriorDeptNum = m_rsEmployeeData!PREV_DEPT_NUM
    l_strCompanyCode = m_rsEmployeeData!CMPY_CD
    l_strPriorCompanyCode = m_rsEmployeeData!PREV_CMPY_CD
    l_strPayrollGroupCd = m_rsEmployeeData!PAYROLL_GROUP_CD
    
    If IsAmerenCIPSorCILCO(dbRules.TRSRuleCode.trIsAmerenCIPS, l_datDeptEffDate, l_strDeptNum, l_strPriorDeptNum, _
            l_strCompanyCode, l_strPriorCompanyCode, v_datChargeDate, l_strUnionCode, l_strPayrollGroupCd) Then
        KeepsShiftComponent = False
        Exit Function
    End If
    
    If IsAmerenCIPSorCILCO(dbRules.TRSRuleCode.trIsAmerenCILCO, l_datDeptEffDate, l_strDeptNum, l_strPriorDeptNum, _
            l_strCompanyCode, l_strPriorCompanyCode, v_datChargeDate, l_strUnionCode, l_strPayrollGroupCd) Then
'        If Not IsAmerenCILCODispatch(v_lngEmpNum, l_datDeptEffDate, _
'             l_strDeptNum, l_strPriorDeptNum, v_datChargeDate) Then
            KeepsShiftComponent = False
            Exit Function
'        End If
    End If
    
    If IsCallawaySecurity(v_lngEmpNum, l_datDeptEffDate, l_strDeptNum, l_strPriorDeptNum, v_datChargeDate) Then
            KeepsShiftComponent = False
            Exit Function
    End If
    
    Select Case EmployeeUnion(l_datJobEffDate, l_strUnionCode, _
                            l_strPriorUnionCode, v_datChargeDate)
                            
        Case "TA", "TB"
            KeepsShiftComponent = False
            
        Case Else
            KeepsShiftComponent = True
            
     End Select
     
    Exit Function
    
eh:

    TrsRaiseError m_strCMODULENAME, "KeepsShiftComponent"
     

End Function

'Private Function IsAmerenCIPS( _
'            ByVal v_datDeptEffDate As Date, _
'            ByVal v_strDeptNum As String, _
'            ByVal v_strPriorDeptNum As String, _
'            Optional ByVal v_strCompanyCode As String = mstrNotAvailable, _
'            Optional ByVal v_strPriorCompanyCode As String = mstrNotAvailable, _
'            Optional ByVal v_datChgDate As Date = mstrDefaultDate) As Boolean
'
'     Dim l_obj As busEmployee.cbsEmployee
'     Dim l_datChgDate As Date
'
'On Error GoTo eh
'
'    Set l_obj = CreateObject("busEmployee.cbsEmployee")
'
'    IsAmerenCIPS = l_obj.IsAmerenCIPS( _
'        v_datDeptEffDate, v_strDeptNum, v_strPriorDeptNum, _
'        v_strCompanyCode, v_strPriorCompanyCode, v_datChgDate)
'
'    Set l_obj = Nothing
'
'    Exit Function
'
'eh:
'    Set l_obj = Nothing
'    TrsRaiseError m_strCMODULENAME, "IsAmerenCIPS", , , , , True
'
'
'End Function

'Private Function IsAmerenCILCO( _
'            ByVal v_datDeptEffDate As Date, _
'            ByVal v_strDeptNum As String, _
'            ByVal v_strPriorDeptNum As String, _
'            Optional ByVal v_strCompanyCode As String = mstrNotAvailable, _
'            Optional ByVal v_strPriorCompanyCode As String = mstrNotAvailable, _
'            Optional ByVal v_datChgDate As Date = mstrDefaultDate) As Boolean
'
'     Dim l_obj As busEmployee.cbsEmployee
'     Dim l_datChgDate As Date
'     Dim l_strCompany As String
'     Dim l_strDept As String
'
''Once all is cleaned up l_obj will change to busEmployee.cbsEmployee
''The IsAmerenCILCO function will have the same arguements passed to this function
'
'On Error GoTo eh
'
'    Set l_obj = CreateObject("busEmployee.cbsEmployee")
'
'    IsAmerenCILCO = l_obj.IsAmerenCILCO( _
'        v_datDeptEffDate, v_strDeptNum, v_strPriorDeptNum, _
'        v_strCompanyCode, v_strPriorCompanyCode, v_datChgDate)
'
'    Set l_obj = Nothing
'
'    Exit Function
'
'eh:
'    Set l_obj = Nothing
'    TrsRaiseError m_strCMODULENAME, "IsAmerenCILCO", , , , , True
'
'
'End Function

Private Function IsAmerenCIPSorCILCO( _
            ByVal v_lngRuleNum As Long, _
            ByVal v_datDeptEffDate As Date, _
            ByVal v_strDeptNum As String, _
            ByVal v_strPriorDeptNum As String, _
            Optional ByVal v_strCompanyCode As String = mstrNotAvailable, _
            Optional ByVal v_strPriorCompanyCode As String = mstrNotAvailable, _
            Optional ByVal v_datChgDate As Date = mstrDefaultDate, _
            Optional ByVal v_strUnionCode As String = mstrNotAvailable2, _
            Optional ByVal v_strPayrollGroupCd As String = mstrNotAvailable2) As Boolean
            
     Dim l_obj As busEmployee.cbsEmployee
On Error GoTo eh
   
    Set l_obj = CreateObject("busEmployee.cbsEmployee")
    
    IsAmerenCIPSorCILCO = l_obj.IsAmerenCIPSorCILCO(v_lngRuleNum, _
        v_datDeptEffDate, v_strDeptNum, v_strPriorDeptNum, _
        v_strCompanyCode, v_strPriorCompanyCode, v_datChgDate, _
        v_strUnionCode, v_strPayrollGroupCd)

    Set l_obj = Nothing
    
    Exit Function
eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "IsAmerenCIPSorCILCO", , , , , True
End Function

Private Function IsAmerenCILCODispatch( _
            ByVal v_lngEmpNum As Long, _
            ByVal v_datDeptEffDate As Date, _
            ByVal v_strDeptNum As String, _
            ByVal v_strPriorDeptNum As String, _
            Optional ByVal v_datChgDate As Date = mstrDefaultDate) As Boolean
            
     Dim l_obj As busEmployee.cbsEmployee
     Dim l_datChgDate As Date
     
On Error GoTo eh
    
    Set l_obj = CreateObject("busEmployee.cbsEmployee")
    
    IsAmerenCILCODispatch = l_obj.IsAmerenCILCODispatch( _
        v_lngEmpNum, v_datDeptEffDate, v_strDeptNum, _
        v_strPriorDeptNum, v_datChgDate)
    
    Set l_obj = Nothing
    
    Exit Function

eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "IsAmerenCILCODispatch", , , , , True
    
            
End Function

Private Function EmployeeUnion( _
            ByVal v_datJobEffDate As Date, _
            ByVal v_strUnionCode As String, _
            ByVal v_strPriorUnionCode As String, _
            Optional ByVal v_strChgDate As Date = mstrDefaultDate) As String

    Dim l_datChgDate As Date
    Dim l_obj As busEmployee.cbsEmployee

On Error GoTo eh

    Set l_obj = CreateObject("busEmployee.cbsEmployee")
    EmployeeUnion = l_obj.EmployeeUnion( _
        v_datJobEffDate, v_strUnionCode, v_strPriorUnionCode, v_strChgDate)
    
    Set l_obj = Nothing
    Exit Function

eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "EmployeeUnion", , , , , True

End Function

Private Function IsEdwardsOperations( _
            ByVal v_lngEmpNum As Long, _
            ByVal v_datDeptEffDate As Date, _
            ByVal v_strDeptNum As String, _
            ByVal v_strPriorDeptNum As String, _
            Optional ByVal v_datChgDate As Date = mstrDefaultDate) As Boolean
            
     Dim l_obj As busEmployee.cbsEmployee
     Dim l_datChgDate As Date
     
On Error GoTo eh
    
    Set l_obj = CreateObject("busEmployee.cbsEmployee")
    
    IsEdwardsOperations = l_obj.IsEdwardsOperations( _
        v_lngEmpNum, v_datDeptEffDate, v_strDeptNum, _
        v_strPriorDeptNum, v_datChgDate)
    
    Set l_obj = Nothing
    
    Exit Function

eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "IsEdwardsOperations", , , , , True
    
            
End Function

Private Function IsCoffeenOperations( _
            ByVal v_lngEmpNum As Long, _
            ByVal v_datDeptEffDate As Date, _
            ByVal v_strDeptNum As String, _
            ByVal v_strPriorDeptNum As String, _
            Optional ByVal v_datChgDate As Date = mstrDefaultDate) As Boolean
            
     Dim l_obj As busEmployee.cbsEmployee
     Dim l_datChgDate As Date
     
On Error GoTo eh
    
    Set l_obj = CreateObject("busEmployee.cbsEmployee")
    
    IsCoffeenOperations = l_obj.IsCoffeenOperations( _
        v_lngEmpNum, v_datDeptEffDate, v_strDeptNum, _
        v_strPriorDeptNum, v_datChgDate)
    
    Set l_obj = Nothing
    
    Exit Function

eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "IsCoffeenOperations", , , , , True
    
            
End Function

Private Function IsCoffeenMaintenance( _
            ByVal v_lngEmpNum As Long, _
            ByVal v_datDeptEffDate As Date, _
            ByVal v_strDeptNum As String, _
            ByVal v_strPriorDeptNum As String, _
            Optional ByVal v_datChgDate As Date = mstrDefaultDate) As Boolean
            
     Dim l_obj As busEmployee.cbsEmployee
     Dim l_datChgDate As Date
     
On Error GoTo eh
    
    Set l_obj = CreateObject("busEmployee.cbsEmployee")
    
    IsCoffeenMaintenance = l_obj.IsCoffeenMaintenance( _
        v_lngEmpNum, v_datDeptEffDate, v_strDeptNum, _
        v_strPriorDeptNum, v_datChgDate)
    
    Set l_obj = Nothing
    
    Exit Function

eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "IsCoffeenMaintenance", , , , , True
    
            
End Function

Private Sub ShiftComponentSplits( _
    ByRef r_rsGetAdjBaseSched As ADODB.Recordset)
    
    Dim l_lngEmpNum As Long
    Dim l_strDeptNum As String
    Dim l_strPriorDeptNum As String
    Dim l_datDeptEffDate As Date
    Dim l_blnIsCoffeen As Boolean
    Dim l_blnIsEdwardsOperations As Boolean
    Dim l_datChargeDate As Date
    Dim l_aryShifts As Variant
    
    On Error GoTo eh
    Exit Sub
    l_lngEmpNum = r_rsGetAdjBaseSched!EMP_NUM
    l_datChargeDate = r_rsGetAdjBaseSched!CHARGE_DATE
    m_rsEmployeeData.Filter = "EMP_NUM = " & str(l_lngEmpNum)
    If Not RecordsExist(m_rsEmployeeData) Then
        m_rsEmployeeData.Filter = ""
        Exit Sub
    End If
    
    l_datDeptEffDate = m_rsEmployeeData!DEPT_EFF_DATE
    l_strDeptNum = m_rsEmployeeData!DEPT_NUM
    l_strPriorDeptNum = m_rsEmployeeData!PREV_DEPT_NUM
    
    l_blnIsCoffeen = IsCoffeenOperations(l_lngEmpNum, l_datDeptEffDate, l_strDeptNum, l_strPriorDeptNum, l_datChargeDate) _
          Or IsCoffeenMaintenance(l_lngEmpNum, l_datDeptEffDate, l_strDeptNum, l_strPriorDeptNum, l_datChargeDate)
    
    l_blnIsEdwardsOperations = IsEdwardsOperations(l_lngEmpNum, l_datDeptEffDate, l_strDeptNum, l_strPriorDeptNum, l_datChargeDate)
    
    If l_blnIsCoffeen Or l_blnIsEdwardsOperations Then
    'Continue
    Else
        Exit Sub
    End If
    
    If l_blnIsCoffeen Then
        l_aryShifts = GetCoffeenShifts()
    Else
        l_aryShifts = GetEdwardsOperationsShifts()
    End If
    
    ShiftComponentSplitsDetermine l_aryShifts, r_rsGetAdjBaseSched

    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "ShiftComponentSplits"
    

End Sub


Private Sub ShiftComponentSplitsDetermine( _
    ByRef r_aryShifts As Variant, _
    ByRef r_rsGetAdjBaseSched As ADODB.Recordset)
    
    
'Please note that time that crosses midnight will be split prior to this method
'Start time will always be less than end time

    Dim l_strStartTime As String
    Dim l_strEndTime As String
    Dim l_aryShifts(3, 2) As String
    Dim l_strCompStartTime As String
    Dim l_strCompEndTime As String
    Dim l_strShiftComp As String
    Dim l_blnContinue As Boolean
    Dim l_datDateWorked As Date
    Dim l_decHours As Variant
    Dim l_strSystemExceptionCode As String

    
    On Error GoTo eh
    
    l_strStartTime = r_rsGetAdjBaseSched!START_TIME
    
    If Trim(l_strStartTime) = "" Then
        r_rsGetAdjBaseSched!COMP_SHIFT_CD = ""
        Exit Sub
    End If
    
    l_blnContinue = True
    While l_blnContinue
        l_strStartTime = r_rsGetAdjBaseSched!START_TIME
        l_strEndTime = r_rsGetAdjBaseSched!END_TIME
        GetShiftCompAndTimes l_strStartTime, l_strShiftComp, l_strCompStartTime, l_strCompEndTime, r_aryShifts
        r_rsGetAdjBaseSched!COMP_SHIFT_CD = l_strShiftComp
        
        If l_strEndTime > l_strCompEndTime Then
            r_rsGetAdjBaseSched!END_TIME = l_strCompEndTime
            l_decHours = DetermineHours(l_strStartTime, l_strCompEndTime)
            r_rsGetAdjBaseSched!HOURS = l_decHours
            l_datDateWorked = r_rsGetAdjBaseSched!DATE_WORKED
            r_rsGetAdjBaseSched.Update
            InitializeExceptionRepeatingShiftData r_rsGetAdjBaseSched
            l_strSystemExceptionCode = r_rsGetAdjBaseSched!SYSTEM_EXC_CD
            r_rsGetAdjBaseSched.AddNew
            SetRepeatingShiftData r_rsGetAdjBaseSched
            r_rsGetAdjBaseSched!SYSTEM_EXC_CD = l_strSystemExceptionCode
            r_rsGetAdjBaseSched!DATE_WORKED = l_datDateWorked
            r_rsGetAdjBaseSched!START_TIME = l_strCompEndTime
            r_rsGetAdjBaseSched!END_TIME = l_strEndTime
            l_decHours = DetermineHours(l_strCompEndTime, l_strEndTime)
            r_rsGetAdjBaseSched!HOURS = l_decHours
        Else
            l_blnContinue = False
        End If
    Wend
        
    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "ShiftComponentSplitsDetermine"
    
    
End Sub


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
    TrsRaiseError m_strCMODULENAME, "GetShiftCompAndTimes"
    
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
    
    DetermineHours = l_decHours
    
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "DetermineHours"


End Function

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
    TrsRaiseError m_strCMODULENAME, "GetCoffeenShifts"
    
End Function

Private Function GetEdwardsOperationsShifts() As Variant

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
    
    GetEdwardsOperationsShifts = l_aryShifts
    

  Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "GetEdwardsOperationsShifts"
    
End Function

Public Function GetScheduleForOTCanvassing _
             (ByVal v_aryEmployees As Variant, _
              ByVal v_datStartDate As Date, _
              ByVal v_datEndDate As Date) _
              As ADODB.Recordset


    On Error GoTo eh

    Dim l_lngI As Long
    Dim l_lngEmpNum As Long
    Dim l_lngStart As Long
    Dim l_lngEnd As Long
    Dim l_rsGetAdjBaseSched As ADODB.Recordset
    Dim l_aryEmployees As Variant
    Dim l_rsActTimeExc As ADODB.Recordset

    On Error GoTo eh

    ValidateRequest 0, v_datStartDate, v_datEndDate, trAdjBaseSchedWithExceptionsAndOvertime, trAllScheduleLevels
    InitializeRecordsets trAdjBaseSchedWithExceptionsAndOvertime, False, l_rsGetAdjBaseSched, Nothing
    
    If IsArray(v_aryEmployees) Then
        l_aryEmployees = v_aryEmployees
    Else
        ReDim l_aryEmployees(0) As String
        l_aryEmployees(0) = str(CLng(v_aryEmployees))
    End If
    
'Retrieve Employee Schedule Data
'It place data in private global variables to reduce data passing
    GetEmployeeScheduleData l_aryEmployees, v_datStartDate, _
        v_datEndDate, trAdjBaseSchedWithExceptionsAndOvertime, trAllScheduleLevels, adLockBatchOptimistic
        
    Set m_rsRotations = GetRotationSchedules()
    
    GetAdjustedBaseSchedulePrivate l_aryEmployees, v_datStartDate, _
        v_datEndDate, Nothing, l_rsGetAdjBaseSched, Nothing, _
        False, trAdjBaseSchedWithExceptionsAndOvertime, trAllScheduleLevels
           
    Set GetScheduleForOTCanvassing = l_rsGetAdjBaseSched
           
    Set l_rsActTimeExc = GetReportedOTHolAndVacs(l_aryEmployees, v_datStartDate, v_datEndDate)
    If RecordsExist(l_rsActTimeExc) Then
        ApplyActualTimeExceptions l_rsGetAdjBaseSched, l_rsActTimeExc
    End If
    
    l_rsGetAdjBaseSched.Filter = ""
    Exit Function
    
eh:
    ClearGlobalRecordSets
    Set l_rsGetAdjBaseSched = Nothing
    Dim l_colDescArgs As Collection
    Set l_colDescArgs = New Collection
    l_colDescArgs.Add str(v_datStartDate)
    l_colDescArgs.Add str(v_datEndDate)
    Select Case Err.Number
        Case trErrbusGetAdjustedBaseSchedInvalidDate
            TrsRaiseError m_strCMODULENAME, "GetScheduleForOTCanvassing", Err.Number, False, l_colDescArgs, , True
        Case Else
            TrsRaiseError m_strCMODULENAME, "GetScheduleForOTCanvassing", , , , , True
    End Select

End Function

Private Function GetReportedOTHolAndVacs( _
    ByVal v_aryEmployees As Variant, _
    ByVal v_datStartDate As Date, _
    ByVal v_datEndDate As Date _
    ) As ADODB.Recordset
    
    Dim l_obj As dbSchedule.cdSchedule
    Dim l_rs As ADODB.Recordset
    Dim l_strEmployee As String
    Dim l_colEmployees As Collection
    
On Error GoTo eh

    
    Set l_obj = CreateObject("dbSchedule.cdSchedule")
    
    Set l_rs = l_obj.GetReportedOTHolAndVacs(v_aryEmployees, v_datStartDate, v_datEndDate)
    
    Set l_obj = Nothing
    
    Set GetReportedOTHolAndVacs = l_rs
    

    Exit Function
eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "GetReportedOTHolAndVacs"
    
End Function




Private Sub ApplyActualTimeExceptions( _
    ByRef r_rsSched As ADODB.Recordset, _
    ByVal v_rsActTime As ADODB.Recordset)
    
    Dim l_lngEmpNum As Long
    Dim l_datChargeDate As Date
    Dim l_intSchedCode As Integer
    Dim l_strException As String
    Dim l_strFilter As String
    Dim l_lngOTOppNum As Long
    
    On Error GoTo eh
    
    
    v_rsActTime.MoveFirst
    While Not v_rsActTime.EOF
        l_lngEmpNum = v_rsActTime!EMP_NUM
        l_datChargeDate = v_rsActTime!CHARGE_DATE
        l_lngOTOppNum = v_rsActTime!OT_OPP_NUM
        If l_lngOTOppNum > 0 Then
            l_strFilter = "EMP_NUM = " & str(l_lngEmpNum) & _
                " AND   CHARGE_DATE = " & str(l_datChargeDate)
            l_strFilter = MDAC25FilterFix(l_strFilter)
            r_rsSched.Filter = l_strFilter
            InitializeRepeatingShiftData2 r_rsSched
            AddOvertimeToSchedule r_rsSched, v_rsActTime
        Else
            l_intSchedCode = v_rsActTime!SCHED_CD
            l_strException = v_rsActTime!EXCEPTION_CD
            l_strFilter = "EMP_NUM = " & str(l_lngEmpNum) & _
                " AND   CHARGE_DATE = " & str(l_datChargeDate) & _
                " AND SCHED_CD = " & str(l_intSchedCode) & _
                " AND  OT_OPP_NUM = 0 " & _
                " AND  SYSTEM_EXC_CD <> 'SLT' "
            l_strFilter = MDAC25FilterFix(l_strFilter)
            r_rsSched.Filter = l_strFilter
            If RecordsExist(r_rsSched) Then
                ApplyActTimeExc r_rsSched, v_rsActTime
            Else
                l_strFilter = "EMP_NUM = " & str(l_lngEmpNum) & _
                    " AND   CHARGE_DATE = " & str(l_datChargeDate) & _
                    " AND  OT_OPP_NUM = 0 " & _
                    " AND  SYSTEM_EXC_CD <> 'SLT' "
                l_strFilter = MDAC25FilterFix(l_strFilter)
                r_rsSched.Filter = l_strFilter
                If RecordsExist(r_rsSched) Then
                    ApplyActTimeExc r_rsSched, v_rsActTime
                End If
            End If
        End If
        v_rsActTime.MoveNext
    Wend
        
    Exit Sub
eh:
    TrsRaiseError m_strCMODULENAME, "ApplyActualTimeExceptions"
        
End Sub

Private Sub ApplyActTimeExc( _
    ByRef r_rsSched As ADODB.Recordset, _
    ByVal v_rsActTime As ADODB.Recordset)
    
    Dim l_lngEmpNum As Long
    Dim l_datChargeDate As Date
    Dim l_strException As String
    Dim l_intIntersect As Integer
    Dim l_datActStartDateTime As Date
    Dim l_datActEndDateTime As Date
    Dim l_datSchedStartDateTime As Date
    Dim l_datSchedEndDateTime As Date
    Dim l_strTime As String
    Dim l_datDate As Date
    Dim l_datDateTime As Date
    Dim l_blnSchedDayOff As Boolean
    Dim l_decHours As Variant
    Dim l_intPos As Integer
    Dim l_intI As Integer
    Dim l_intCount As Integer
'    Dim l_rsSched As ADODB.Recordset
    
    
    
    On Error GoTo eh
    
    l_strTime = v_rsActTime!START_TIME
    l_datDate = v_rsActTime!DATE_WORKED
    l_datActStartDateTime = DateTimeFrom(l_datDate, l_strTime)
    l_strTime = v_rsActTime!END_TIME
    l_datActEndDateTime = DateTimeFrom(l_datDate, l_strTime, True)
    
    r_rsSched.Sort = "DATE_WORKED, START_TIME"
    r_rsSched.MoveFirst
    l_blnSchedDayOff = True
    While Not r_rsSched.EOF
        l_strTime = Trim(r_rsSched!START_TIME)
        If l_strTime <> "" Then
            l_blnSchedDayOff = False
        End If
        r_rsSched.MoveNext
    Wend
    r_rsSched.MoveFirst
    If l_blnSchedDayOff Then
        r_rsSched!HOURS = v_rsActTime!HOURS
        r_rsSched!START_TIME = v_rsActTime!START_TIME
        r_rsSched!END_TIME = v_rsActTime!END_TIME
        r_rsSched!EXCEPTION_CD = v_rsActTime!EXCEPTION_CD
        r_rsSched!DATE_WORKED = v_rsActTime!DATE_WORKED
        Exit Sub
    End If
    
    r_rsSched.MoveFirst
    l_intCount = r_rsSched.RecordCount
    For l_intI = 1 To l_intCount
        r_rsSched.AbsolutePosition = l_intI
        l_strTime = Trim(r_rsSched!START_TIME)
        If l_strTime <> "" Then
            l_datDate = r_rsSched!DATE_WORKED
            l_datSchedStartDateTime = DateTimeFrom(l_datDate, l_strTime)
            l_strTime = r_rsSched!END_TIME
            l_datSchedEndDateTime = DateTimeFrom(l_datDate, l_strTime)
            l_intIntersect = DateIntersection(l_datActStartDateTime, l_datActEndDateTime, _
                l_datSchedStartDateTime, l_datSchedEndDateTime)
            Select Case l_intIntersect
                Case 0
                'The Schedule does not overlapp
                'This will happen because reported exceptions spilt over lunch
                Case 1
                'This should be the case 99% of the time for Union OA and
                ' an extremly high majority for everyone else
                'Usually Report Exception time matches the schedule exactly
                    r_rsSched!EXCEPTION_CD = v_rsActTime!EXCEPTION_CD
                Case 2
                'This causes a split in the middle
                    InitializeRepeatingShiftData2 r_rsSched
                    ChangeSchedule r_rsSched, l_datSchedStartDateTime, l_datActStartDateTime
                    AddToSchedule r_rsSched, l_datActStartDateTime, l_datActEndDateTime
                    r_rsSched!EXCEPTION_CD = v_rsActTime!EXCEPTION_CD
                    AddToSchedule r_rsSched, l_datActEndDateTime, l_datSchedEndDateTime
                Case 3
                'This causes a split in the beginning
                    If l_datActEndDateTime = l_datSchedStartDateTime Then
                    'DO NOTHING
                    Else
                        InitializeRepeatingShiftData2 r_rsSched
                        ChangeSchedule r_rsSched, l_datActStartDateTime, l_datSchedStartDateTime
                        AddToSchedule r_rsSched, l_datActStartDateTime, l_datActEndDateTime
                        r_rsSched!EXCEPTION_CD = v_rsActTime!EXCEPTION_CD
                        AddToSchedule r_rsSched, l_datActEndDateTime, l_datSchedEndDateTime
                    End If
                Case 4
                'This causes a split in the end
                    If l_datActStartDateTime = l_datSchedEndDateTime Then
                    'DO NOTHING
                    Else
                        InitializeRepeatingShiftData2 r_rsSched
                        ChangeSchedule r_rsSched, l_datSchedStartDateTime, l_datActStartDateTime
                        AddToSchedule r_rsSched, l_datActStartDateTime, l_datSchedEndDateTime
                        r_rsSched!EXCEPTION_CD = v_rsActTime!EXCEPTION_CD
                        AddToSchedule r_rsSched, l_datSchedEndDateTime, l_datActEndDateTime
                        r_rsSched!EXCEPTION_CD = v_rsActTime!EXCEPTION_CD
                    End If
                Case 5
                 'Do Nothing
            
            End Select
        End If
    Next l_intI
    Exit Sub
eh:
    TrsRaiseError m_strCMODULENAME, "ApplyActTimeExc"
        
End Sub

Private Function CrossesMidnight( _
    ByVal v_datStartDateTime As Date, _
    ByVal v_datEndDateTime As Date _
    ) As Boolean
    
    Dim l_datStart As Date
    Dim l_datEnd As Date
    
    On Error GoTo eh
    
    l_datStart = Format(v_datStartDateTime, "YYYY-MM-DD")
    l_datEnd = Format(v_datEndDateTime, "YYYY-MM-DD")
    
    If l_datEnd > l_datStart Then
        CrossesMidnight = True
    Else
        CrossesMidnight = False
    End If
    
  Exit Function
  
eh:
    TrsRaiseError m_strCMODULENAME, "CrossesMidnight"
    
End Function

Private Sub InitializeRepeatingShiftData2( _
       ByVal v_rsSched As ADODB.Recordset)

    On Error GoTo eh

    m_varRepeatingShiftData(0) = v_rsSched!EMP_NUM
    m_varRepeatingShiftData(1) = v_rsSched!CHARGE_DATE
    m_varRepeatingShiftData(2) = v_rsSched!SCHED_CD
    m_varRepeatingShiftData(3) = v_rsSched!SCHED_SRC_CD
    m_varRepeatingShiftData(4) = v_rsSched!SHIFT_START_TIME
    m_varRepeatingShiftData(5) = v_rsSched!SHIFT_END_TIME
    m_varRepeatingShiftData(6) = v_rsSched!SHIFT_START_DATE
    m_varRepeatingShiftData(7) = v_rsSched!SHIFT_END_DATE
    m_varRepeatingShiftData(8) = v_rsSched!SHIFT_HOURS
    m_varRepeatingShiftData(9) = v_rsSched!SHIFT_CODE
    m_varRepeatingShiftData(10) = v_rsSched!SCHED_JOB_NUM
    m_varRepeatingShiftData(11) = v_rsSched!LINE_NUM
    m_varRepeatingShiftData(12) = v_rsSched!WG_NUM
    m_varRepeatingShiftData(13) = v_rsSched!REQ_LINE_IND
    m_varRepeatingShiftData(14) = v_rsSched!COMP_SHIFT_CD
    m_varRepeatingShiftData(15) = v_rsSched!REL_CD
    m_varRepeatingShiftData(16) = v_rsSched!SCHED_CD_FOR_SHIFT_COMP
    m_varRepeatingShiftData(17) = v_rsSched!SHIFT_NUM
    m_varRepeatingShiftData(18) = v_rsSched!LINE_SRC_CD
    m_varRepeatingShiftData(19) = v_rsSched!SHIFT_LUNCH_START_TIME
    m_varRepeatingShiftData(20) = v_rsSched!SHIFT_LUNCH_END_TIME
    m_varRepeatingShiftData(21) = v_rsSched!TRADE_NUM
    m_varRepeatingShiftData(22) = v_rsSched!BASE_SCHED_SRC_CD
'Profiler
    m_varRepeatingShiftData(23) = v_rsSched!EMP_JOB_NUM
    m_varRepeatingShiftData(24) = v_rsSched!ROT_CD
    m_varRepeatingShiftData(25) = v_rsSched!ROT_WEEK
    m_varRepeatingShiftData(26) = v_rsSched!SYS_COMP_SHIFT_CD
    
    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "InitializeRepeatingShiftData2"

End Sub

Private Sub SetRepeatingShiftData2( _
        ByRef r_rsSched As ADODB.Recordset)

    On Error GoTo eh
    
    r_rsSched!EMP_NUM = m_varRepeatingShiftData(0)
    r_rsSched!CHARGE_DATE = m_varRepeatingShiftData(1)
    r_rsSched!SCHED_CD = m_varRepeatingShiftData(2)
    r_rsSched!SCHED_SRC_CD = m_varRepeatingShiftData(3)
    r_rsSched!SHIFT_START_TIME = m_varRepeatingShiftData(4)
    r_rsSched!SHIFT_END_TIME = m_varRepeatingShiftData(5)
    r_rsSched!SHIFT_START_DATE = m_varRepeatingShiftData(6)
    r_rsSched!SHIFT_END_DATE = m_varRepeatingShiftData(7)
    r_rsSched!SHIFT_HOURS = Round(m_varRepeatingShiftData(8), 1)
    r_rsSched!SHIFT_CODE = m_varRepeatingShiftData(9)
    r_rsSched!SCHED_JOB_NUM = m_varRepeatingShiftData(10)
    r_rsSched!LINE_NUM = m_varRepeatingShiftData(11)
    r_rsSched!WG_NUM = m_varRepeatingShiftData(12)
    r_rsSched!REQ_LINE_IND = m_varRepeatingShiftData(13)
    r_rsSched!COMP_SHIFT_CD = m_varRepeatingShiftData(14)
    r_rsSched!REL_CD = m_varRepeatingShiftData(15)
    r_rsSched!SHIFT_START_END = m_varRepeatingShiftData(4) & "-" & m_varRepeatingShiftData(5)
    r_rsSched!SCHED_CD_FOR_SHIFT_COMP = m_varRepeatingShiftData(16)
    r_rsSched!SHIFT_NUM = m_varRepeatingShiftData(17)
    r_rsSched!LINE_SRC_CD = m_varRepeatingShiftData(18)
    r_rsSched!SHIFT_LUNCH_START_TIME = m_varRepeatingShiftData(19)
    r_rsSched!SHIFT_LUNCH_END_TIME = m_varRepeatingShiftData(20)
    r_rsSched!TRADE_NUM = m_varRepeatingShiftData(21)
    r_rsSched!BASE_SCHED_SRC_CD = m_varRepeatingShiftData(22)
    r_rsSched!SCHED_CHANGE_PMT = "N"
    r_rsSched!WG_LINE_IND = "N"
    r_rsSched!FULL_SHIFT_EXC_IND = "N"
    r_rsSched!DOUBLE_SHIFT_IND = "N"
'profiler
    r_rsSched!EMP_JOB_NUM = m_varRepeatingShiftData(23)
'Optimistic locking
    r_rsSched!LAST_MOD_SCHED_TM = ""
'View Schedule wants the work group of the line
    r_rsSched!LINE_WG = m_varRepeatingShiftData(12)
    r_rsSched!LINE_JOB_NUM = m_varRepeatingShiftData(10)
    
    r_rsSched!SCHED_CHG_EXC_CD = ""
    r_rsSched!ROT_CD = m_varRepeatingShiftData(24)
    r_rsSched!ROT_WEEK = m_varRepeatingShiftData(25)
    r_rsSched!SYS_COMP_SHIFT_CD = m_varRepeatingShiftData(26)

    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "SetRepeatingShiftData2"
End Sub

Private Sub AddToSchedule( _
    ByRef r_rsSched As ADODB.Recordset, _
    ByVal v_datStartDateTime As Date, _
    ByVal v_datEndDateTime As Date)
    
    Dim l_strStartTime As String
    Dim l_strEndTime As String
    Dim l_datDateWorked As Date
    Dim l_decHours As Variant
    Dim l_datDate As Date
    Dim l_strTime As String
    Dim l_datDateTime As Date
    
    On Error GoTo eh
    
    r_rsSched.AddNew
    SetRepeatingShiftData2 r_rsSched
    InitializeExceptionVariables r_rsSched
    ChangeSchedule r_rsSched, v_datStartDateTime, v_datEndDateTime
    
    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "AddToSchedule"
    
End Sub

Private Sub ChangeSchedule( _
    ByRef r_rsSched As ADODB.Recordset, _
    ByVal v_datStartDateTime As Date, _
    ByVal v_datEndDateTime As Date)
    
    Dim l_strStartTime As String
    Dim l_strEndTime As String
    Dim l_datDateWorked As Date
    Dim l_decHours As Variant
    Dim l_datDate As Date
    Dim l_strTime As String
    Dim l_datDateTime As Date
    
    On Error GoTo eh
    
    If CrossesMidnight(v_datStartDateTime, v_datEndDateTime) Then
        l_strStartTime = Format(v_datStartDateTime, "HH:MM")
        l_strEndTime = "24:00"
        l_datDateWorked = Format(v_datStartDateTime, "YYYY-MM-DD")
        l_datDateTime = DateTimeFrom(l_datDateWorked, "24:00")
        l_decHours = DateDiff("n", v_datStartDateTime, l_datDateTime) / 60
        l_decHours = Round(l_decHours, 1)
        r_rsSched!START_TIME = l_strStartTime
        r_rsSched!END_TIME = l_strEndTime
        r_rsSched!DATE_WORKED = l_datDateWorked
        r_rsSched!HOURS = l_decHours
        r_rsSched.AddNew
        SetRepeatingShiftData r_rsSched
        InitializeExceptionVariables r_rsSched
        l_strStartTime = "00:00"
        l_strEndTime = Format(v_datEndDateTime, "HH:MM")
        l_datDateWorked = Format(v_datEndDateTime, "YYYY-MM-DD")
        l_datDateTime = DateTimeFrom(l_datDateWorked, "00:00")
        l_decHours = DateDiff("n", l_datDateTime, v_datEndDateTime) / 60
        l_decHours = Round(l_decHours, 1)
        r_rsSched!START_TIME = l_strStartTime
        r_rsSched!END_TIME = l_strEndTime
        r_rsSched!DATE_WORKED = l_datDateWorked
        r_rsSched!HOURS = l_decHours
    Else
        l_strStartTime = Format(v_datStartDateTime, "HH:MM")
        l_strEndTime = Format(v_datEndDateTime, "HH:MM")
        l_datDateWorked = Format(v_datStartDateTime, "YYYY-MM-DD")
        l_decHours = DateDiff("n", v_datStartDateTime, v_datEndDateTime) / 60
        l_decHours = Round(l_decHours, 1)
        r_rsSched!START_TIME = l_strStartTime
        r_rsSched!END_TIME = l_strEndTime
        r_rsSched!DATE_WORKED = l_datDateWorked
        r_rsSched!HOURS = l_decHours
    End If
    
        
    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "ChangeSchedule"

End Sub


Private Sub AddOvertimeToSchedule( _
    ByRef r_rsSched As ADODB.Recordset, _
    ByVal v_rsActTime As ADODB.Recordset)
    
    Dim l_strStartTime As String
    Dim l_strEndTime As String
    Dim l_datDateWorked As Date
    Dim l_decHours As Variant
    Dim l_datDate As Date
    Dim l_strTime As String
    Dim l_datDateTime As Date
    
    On Error GoTo eh
    
    r_rsSched.AddNew
    SetRepeatingShiftData2 r_rsSched
    InitializeExceptionVariables r_rsSched
    r_rsSched!START_TIME = v_rsActTime!START_TIME
    r_rsSched!END_TIME = v_rsActTime!END_TIME
    r_rsSched!DATE_WORKED = v_rsActTime!DATE_WORKED
    r_rsSched!HOURS = v_rsActTime!HOURS
    r_rsSched!EXCEPTION_CD = v_rsActTime!EXCEPTION_CD
    r_rsSched!OT_OPP_NUM = v_rsActTime!OT_OPP_NUM
    r_rsSched!TRAV_PAY_HOURS = v_rsActTime!TRAVEL_PAY_HRS
    r_rsSched!SHIFT_START_TIME = ""
    r_rsSched!SHIFT_END_TIME = ""
    r_rsSched!SHIFT_HOURS = 0
    r_rsSched!SHIFT_CODE = ""
    r_rsSched!SHIFT_NUM = 0
    r_rsSched!SHIFT_LUNCH_START_TIME = ""
    r_rsSched!SHIFT_LUNCH_END_TIME = ""
    r_rsSched!SHIFT_START_DATE = v_rsActTime!DATE_WORKED
    r_rsSched!SHIFT_END_DATE = v_rsActTime!DATE_WORKED
    
    
    
    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "AddOvertimeToSchedule"
    
End Sub

Private Function IsCallawaySecurity( _
            ByVal v_lngEmpNum As Long, _
            ByVal v_datDeptEffDate As Date, _
            ByVal v_strDeptNum As String, _
            ByVal v_strPriorDeptNum As String, _
            Optional ByVal v_datChgDate As Date = mstrDefaultDate) As Boolean
            
     Dim l_obj As busEmployee.cbsEmployee
     Dim l_datChgDate As Date
     
On Error GoTo eh
    
    Set l_obj = CreateObject("busEmployee.cbsEmployee")
    
    IsCallawaySecurity = l_obj.IsCallawaySecurity( _
        v_lngEmpNum, _
        v_datDeptEffDate, v_strDeptNum, _
        v_strPriorDeptNum, v_datChgDate)
    
    Set l_obj = Nothing
    
    Exit Function

eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "IsCallawaySecurity", , , , , True
    
            
End Function
