VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cboScheduledShift"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'
'This class is public because references to instances are
' returned from public methods on cboSchedule.
'
Private Const mstrCMODULENAME As String = "cboScheduleShift"

Private mstrShiftStart As String
Private mstrShiftLunchStart As String
Private mstrShiftLunchEnd As String
Private mstrShiftEnd As String
Private mintShiftDay As Integer
Private mintScheduleCode As Long
Private mstrShiftCode As String
Private mstrShiftComponentCode As String
Private mdteChargeDate As Date
Private mdteWorkDate As Date
Private mvarShiftHours As Variant
Private mstrSystemExceptionCode As String
Private mstrReliefCode As String
Private mintScheduleSource As Integer
Private mvarShiftLunchHours As Variant
Private mlngSchedForShiftComp As Long
Private mlngLineNumber As Long
Private mintShiftNumber As Integer
Private mstrReqLineInd As String
Private mstrExceptionCode As String
Private mstrOverrideExceptionCode As String

Public Property Let ScheduleSource(ByVal vData As Integer)
    mintScheduleSource = vData
End Property

Public Property Get ScheduleSource() As Integer
    ScheduleSource = mintScheduleSource
End Property

Public Property Let ShiftHours(ByVal vData As Variant)
' Decimal 3,1
    mvarShiftHours = vData
End Property

Public Property Get ShiftHours() As Variant
' Decimal 3,1
    ShiftHours = mvarShiftHours
End Property

Public Property Let ShiftChargeDate(ByVal vData As Date)
    mdteChargeDate = vData
End Property

Public Property Get ShiftChargeDate() As Date
    ShiftChargeDate = mdteChargeDate
End Property

Public Property Let ShiftWorkDate(ByVal vData As Date)
    mdteWorkDate = vData
End Property

Public Property Get ShiftWorkDate() As Date
    ShiftWorkDate = mdteWorkDate
End Property

Public Property Let ScheduleCode(ByVal vData As Integer)
    mintScheduleCode = vData
End Property

Public Property Get ScheduleCode() As Integer
    ScheduleCode = mintScheduleCode
End Property

Public Property Let ShiftComponentCode(ByVal vData As String)
    mstrShiftComponentCode = Trim(vData)
End Property

Public Property Get ShiftComponentCode() As String
    ShiftComponentCode = mstrShiftComponentCode
End Property

Public Property Let ShiftCode(ByVal vData As String)
    mstrShiftCode = vData
End Property

Public Property Get ShiftCode() As String
    ShiftCode = mstrShiftCode
End Property

Public Property Let SystemExceptionCode(ByVal vData As String)
    mstrSystemExceptionCode = Trim(vData)
End Property

Public Property Get SystemExceptionCode() As String
    SystemExceptionCode = mstrSystemExceptionCode
End Property

Public Property Let ReliefCode(ByVal vData As String)
    mstrReliefCode = Trim(vData)
End Property

Public Property Get ReliefCode() As String
    ReliefCode = mstrReliefCode
End Property

Public Property Let ShiftDay(ByVal vData As Integer)
    mintShiftDay = vData
End Property

Public Property Get ShiftDay() As Integer
    ShiftDay = mintShiftDay
End Property

Public Property Let ShiftEnd(ByVal vData As String)
    mstrShiftEnd = Trim(vData)
End Property

Public Property Get ShiftEnd() As String
    ShiftEnd = mstrShiftEnd
End Property

Public Property Let ShiftStart(ByVal vData As String)
    mstrShiftStart = Trim(vData)
End Property

Public Property Get ShiftStart() As String
    ShiftStart = mstrShiftStart
End Property

Public Function ShiftStartDateTime() As Date
    Dim lcolErrDescArgs As New Collection
    
On Error GoTo eh
    If IsTimeValid(ShiftStart) Then
        ShiftStartDateTime = _
            DateTimeFrom(ShiftWorkDate, ShiftStart)
    Else
        lcolErrDescArgs.Add ShiftStart
        Err.Raise trErrbusShiftInvalidStartTime
    End If
    Exit Function
    
eh:
    Select Case Err.Number
        Case trErrbusShiftInvalidStartTime
            TrsRaiseError mstrCMODULENAME, "ShiftStartDateTime", Err.Number, , lcolErrDescArgs
        Case Else
            TrsRaiseError mstrCMODULENAME, "ShiftStartDateTime"
    End Select

End Function

Public Function ShiftEndDateTime() As Date
    If ShiftSpansMidnight Then
        ShiftEndDateTime = _
            DateTimeFrom(DateAdd("d", 1, ShiftWorkDate), ShiftEnd)
    Else
        ShiftEndDateTime = _
            DateTimeFrom(ShiftWorkDate, ShiftEnd)
    End If
End Function

Public Function LunchStartDateTime() As Date
    If ShiftLunchStart < ShiftStart Then
        LunchStartDateTime = _
            DateTimeFrom(DateAdd("d", 1, ShiftWorkDate), ShiftLunchStart)
    Else
        LunchStartDateTime = _
            DateTimeFrom(ShiftWorkDate, ShiftLunchStart)
    End If
End Function

Public Function LunchEndDateTime() As Date
    If ShiftLunchEnd < ShiftStart Then
        LunchEndDateTime = _
            DateTimeFrom(DateAdd("d", 1, ShiftWorkDate), ShiftLunchEnd)
    Else
        LunchEndDateTime = _
            DateTimeFrom(ShiftWorkDate, ShiftLunchEnd)
    End If
End Function

Public Function ShiftSpansMidnight() As Boolean
' Assumes work date is the date that shift physically starts.
' Therefore, work date is always less than or equal to
' charge date.
'
    ShiftSpansMidnight = ShiftEnd < ShiftStart
End Function

Public Property Let ShiftLunchEnd(ByVal vData As String)
    mstrShiftLunchEnd = Trim(vData)
End Property

Public Property Get ShiftLunchEnd() As String
    ShiftLunchEnd = mstrShiftLunchEnd
End Property

Public Property Let ShiftLunchStart(ByVal vData As String)
    mstrShiftLunchStart = Trim(vData)
End Property

Public Property Get ShiftLunchStart() As String
    ShiftLunchStart = mstrShiftLunchStart
End Property

Public Function IsScheduledLunch() As Boolean
    IsScheduledLunch = _
        (ShiftLunchStart <> "") And (ShiftLunchEnd <> "")
End Function

Public Function IsScheduledWorkDay() As Boolean
    IsScheduledWorkDay = (Me.ShiftStart <> "")
End Function

Public Function IsScheduledHoliday() As Boolean
    
    IsScheduledHoliday = IsAHolidayException(Me.ExceptionCode)
        
End Function

Public Function IsScheduledVacation() As Boolean

    IsScheduledVacation = IsAVacationException(Me.ExceptionCode)

End Function

Public Function IsScheduledSpecialOverride() As Boolean

    IsScheduledSpecialOverride = IsSpecialOverrideCode(Me.OverrideExceptionCode)

End Function

Public Function IsScheduledFirstDayOff() As Boolean
    IsScheduledFirstDayOff = IsScheduledWorkDay And ShiftDay = 1
End Function

Public Function IsScheduledSecondDayOff() As Boolean
    IsScheduledSecondDayOff = IsScheduledWorkDay And ShiftDay = 7
End Function

Public Function IsSelfTrade() As Boolean
    IsSelfTrade = SystemExceptionCode = "TWS"
End Function

Public Function IsTradedWorkShift() As Boolean
    IsTradedWorkShift = _
        SystemExceptionCode = "TSR" Or _
        SystemExceptionCode = "TSA" Or _
        SystemExceptionCode = "TWS"
End Function

Public Function FormatTradeExceptionCode() As String
' Answer shift hours padded with zeroes upto 4 digits
' without decimal point, concatenated with system
' exception code. e.g. 8.0 is answered 0080TSA
    If IsTradedWorkShift Then
        FormatTradeExceptionCode = _
            Format(ShiftHours, "00") & _
            Right(Format(ShiftHours, "0.00"), 2) & _
            SystemExceptionCode
    Else
        FormatTradeExceptionCode = ""
    End If
    
End Function

'
'Public Function IsScheduledSunday() As Boolean
'    IsScheduledSunday = IsScheduledWorkDay And ShiftDay = 1
'End Function
'
'Public Function IsScheduledSaturday() As Boolean
'    IsScheduledSaturday = IsScheduledWorkDay And ShiftDay = 7
'End Function



Public Function IsDayOff() As Boolean
    IsDayOff = Not IsScheduledWorkDay
End Function

Public Function LunchPeriod() As Collection
    Dim colLunch As Collection
    
    Set colLunch = New Collection
    colLunch.Add ShiftLunchStart
    colLunch.Add ShiftLunchEnd
    Set LunchPeriod = colLunch

End Function

Public Function LunchPeriodAsDates() As Collection
    Dim colLunch As Collection
    
    Set colLunch = New Collection
    colLunch.Add LunchStartDateTime
    colLunch.Add LunchEndDateTime
    Set LunchPeriodAsDates = colLunch

End Function

'Public Function LunchPeriodAsMinutes() As Integer
''4/9/01 RSG untested, will include in next release
'On Error GoTo eh
'    If IsScheduledLunch Then
'        LunchPeriodAsMinutes = _
'            DateDiff("n", LunchStartDateTime, LunchEndDateTime)
'    Else
'        LunchPeriodAsMinutes = 0
'    End If
'    Exit Function
'
'eh:
'    TrsRaiseError mstrCMODULENAME, "LunchPeriodAsMinutes"
'
'End Function

Public Function IsIntervalWithinShift( _
            ByVal v_datStartDateTime As Date, _
            ByVal v_datEndDateTime As Date) As Boolean

    IsIntervalWithinShift = _
        TimePeriodsSubset(v_datStartDateTime, v_datEndDateTime, _
                ShiftStartDateTime, ShiftEndDateTime)

End Function

Public Property Let ShiftLunchHours(ByVal vData As Variant)
' Decimal 3,1
    mvarShiftLunchHours = vData
End Property

Public Property Get ShiftLunchHours() As Variant
' Decimal 3,1
    ShiftLunchHours = mvarShiftLunchHours
End Property

Public Property Let ScheduleForShiftComp(ByVal vData As Long)
    mlngSchedForShiftComp = vData
End Property

Public Property Get ScheduleForShiftComp() As Long
    ScheduleForShiftComp = mlngSchedForShiftComp
End Property

Public Property Let LineNumber(ByVal vData As Long)
    mlngLineNumber = vData
End Property

Public Property Get LineNumber() As Long
    LineNumber = mlngLineNumber
End Property

Public Property Let ShiftNumber(ByVal vData As Integer)
    mintShiftNumber = vData
End Property

Public Property Get ShiftNumber() As Integer
    ShiftNumber = mintShiftNumber
End Property

Public Property Let RequiredLineInd(ByVal vData As String)
    mstrReqLineInd = Trim(vData)
End Property

Public Property Get RequiredLineInd() As String
    RequiredLineInd = mstrReqLineInd
End Property

Public Property Let ExceptionCode(ByVal vData As String)
    mstrExceptionCode = Trim(vData)
End Property

Public Property Get ExceptionCode() As String
    ExceptionCode = mstrExceptionCode
End Property

Public Property Let OverrideExceptionCode(ByVal vData As String)
    mstrOverrideExceptionCode = Trim(vData)
End Property

Public Property Get OverrideExceptionCode() As String
    OverrideExceptionCode = mstrOverrideExceptionCode
End Property
