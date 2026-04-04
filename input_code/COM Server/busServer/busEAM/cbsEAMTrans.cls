VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "cbsEAMTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Project:   Post TRIS 5.5 Employee Absence Management
'Author:    Prem Chander
'Date:      August 2005
'Description:

Option Explicit

Private Const trEMPLOYEE As Integer = 1
Private Const trABSENCE As Integer = 2
Private Const trEXCEPTION As Integer = 3
Private Const trLOCATION As Integer = 4
Private Const trFMLAREASON As Integer = 5
Private Const trPROVIDER As Integer = 6
Private Const trABSENCECONTACT As Integer = 7
Private Const trPORTFOLIO As Integer = 8
Private Const trWORKGROUP As Integer = 9

Dim madrsScheduleWithNoExceptions As ADODB.Recordset

'//for testing only
'Private madrsAbsenceStartEndDetails As ADODB.Recordset

Private Const MODULENAME As String = "busEAM.cbsEAMTrans"
Private Const ADOFIELDATTR = adFldUpdatable Or adFldIsNullable

Private Const m_DBSCRIPTSROOT As String = "C:\TRIS SOLUTION\dbScripts\TRIS_db2"
Private Const m_DBSCRIPTSPATH As String = "\busEAM\cbsEAMTrans"
Private m_ResourceMgr As utilADOHelper.cuResourceMgr

Private Function DetermineProbableReturnToWorkOld(ByVal rlAbsenceNumber As Long) As Date
On Error GoTo eh
    'Flag
    Dim lbCheck As Boolean

    'Lock Status
    Dim ladrsLockStatus As ADODB.Recordset

    'Time
    Dim ladrsTime As ADODB.Recordset
    Dim ladrsSchedule As ADODB.Recordset
    Dim ladrsAbsenceTime As ADODB.Recordset
    Dim ldtStart As Date, ldtEnd As Date, ldtExtendedStart As Date
    Dim sErrorDates As String

    'Absence Start and End
    Dim ladrsAbsenceStartEnd As ADODB.Recordset

    'Absence
    Dim ladrsInitial As ADODB.Recordset
    Dim ladrsDate As ADODB.Recordset
    Dim ladrsReturnToWork As ADODB.Recordset

    Dim lcolDates As Collection
    'Components
    Dim lobjbusEAM As busEAM.cbsEAM
    Dim lobjdbCodeTables As dbCodeTables.cdCodeTables
    
    Dim vadrsAbsenceDates As ADODB.Recordset
    Dim vadrsInitialAbsence As ADODB.Recordset
    Dim vadrsAbsenceProvider As ADODB.Recordset
    Dim vadrsAbsenceContact As ADODB.Recordset
    Dim vadrsAbsenceReturnToWork As ADODB.Recordset
    Dim vadrsScheduledExceptions As ADODB.Recordset
    Dim vadrsAbsenceCallLog As ADODB.Recordset
    Dim vadrsExceptions As ADODB.Recordset
    Dim vlEmployeeNumber As Long
    lbCheck = True
    
    DetermineProbableReturnToWorkOld = #1/1/100#
    Set lobjbusEAM = New busEAM.cbsEAM
    lobjbusEAM.GetAbsenceInformation rlAbsenceNumber, "E26895", vadrsAbsenceDates, vadrsInitialAbsence, vadrsAbsenceProvider, vadrsAbsenceContact, _
                                    vadrsAbsenceReturnToWork, vadrsAbsenceCallLog, vadrsScheduledExceptions, Nothing, Nothing
    
    Set lobjbusEAM = Nothing
    
    '//some absence times are in the db2 format '07.00.00' rather than '07:00" and from char(8) to char(5)
    If RecordsExist(vadrsInitialAbsence) Then
        vadrsInitialAbsence.MoveFirst
        Do While Not vadrsInitialAbsence.EOF
            vadrsInitialAbsence!ABS_CALLIN_TM = ConvertDB2Time(vadrsInitialAbsence!ABS_CALLIN_TM)
            vadrsInitialAbsence!ABS_ST_TM = ConvertDB2Time(vadrsInitialAbsence!ABS_ST_TM)
            vadrsInitialAbsence!ABS_END_TM = ConvertDB2Time(vadrsInitialAbsence!ABS_END_TM)
            vadrsInitialAbsence!RTN_TO_WRK_TM = ConvertDB2Time(vadrsInitialAbsence!RTN_TO_WRK_TM)
            
            vadrsInitialAbsence.Update
        
            vadrsInitialAbsence.MoveNext
        Loop
        vadrsInitialAbsence.MoveFirst
    End If
    
    Dim vadrsInitialAbsence2 As ADODB.Recordset
    
    Set vadrsInitialAbsence2 = CopyRecordset(vadrsInitialAbsence, 0)
    vadrsInitialAbsence2.Fields("ABS_CALLIN_TM").DefinedSize = 5
    vadrsInitialAbsence2.Fields("ABS_ST_TM").DefinedSize = 5
    vadrsInitialAbsence2.Fields("ABS_END_TM").DefinedSize = 5
    vadrsInitialAbsence2.Fields("RTN_TO_WRK_TM").DefinedSize = 5
    
    vadrsInitialAbsence2.Open
    AppendRecordset2 vadrsInitialAbsence, vadrsInitialAbsence2
    
    Set vadrsInitialAbsence = vadrsInitialAbsence2
    
    Set lobjdbCodeTables = New dbCodeTables.cdCodeTables
    Set vadrsExceptions = lobjdbCodeTables.GetAllExceptions()
    Set lobjdbCodeTables = Nothing
    
    ldtStart = vbNull
    ldtEnd = vbNull
    
    If RecordsExist(vadrsAbsenceDates) And RecordsExist(vadrsInitialAbsence) Then
        vadrsInitialAbsence.MoveFirst
        vlEmployeeNumber = vadrsInitialAbsence!EMP_NUM
        ldtStart = vadrsInitialAbsence!ABS_ST_DT
    
        With vadrsAbsenceDates
            .Sort = "ABS_DATE"

            .MoveFirst
            If (ldtStart > !ABS_DATE) Then
                ldtStart = !ABS_DATE
            End If
            
            .MoveLast
            ldtEnd = !ABS_DATE

            .MoveFirst
        End With

        Set ladrsLockStatus = GetLockStatus(vlEmployeeNumber, ldtStart, ldtEnd)

        '//Set ladrsSchedule = GetEmployeeSchedule(vlEmployeeNumber, vadrsInitialAbsence, vadrsAbsenceDates)
        Set ladrsSchedule = GetEmployeeSchedule(vlEmployeeNumber, vadrsInitialAbsence, vadrsAbsenceDates)
    End If
    
    '//check for gaps or overlaps in schedule exception requests with dates
    If lbCheck Then
        If RecordsExist(vadrsAbsenceDates) Then
            'NOTE: Schedule should have info atleast 30 days over the Absence End Date
            '//Set ladrsTime = GetEmployeeTimeDetails(vlEmployeeNumber, vadrsInitialAbsence, vadrsAbsenceDates)
            Set ladrsTime = GetEmployeeTimeDetails(vlEmployeeNumber, vadrsInitialAbsence, vadrsAbsenceDates)
                
            'Determine if non consecutive dates were selected
            If RecordsExist(ladrsSchedule) Or RecordsExist(ladrsTime) Then
                Set ladrsAbsenceTime = AddTimeAndSchExceptionsOld(vadrsInitialAbsence _
                                                , vadrsScheduledExceptions _
                                                , ladrsSchedule _
                                                , ladrsTime _
                                                , ladrsLockStatus)
                If RecordsExist(ladrsAbsenceTime) Then '//fails on .MoveLast when no records
                    '//actual time seems to get deleted
                    ladrsAbsenceTime.MoveFirst
                    Do While Not ladrsAbsenceTime.EOF
                        If Trim(ladrsAbsenceTime!REC_STATUS) = "" Then
                            ladrsAbsenceTime!REC_STATUS = "L"
                            ladrsAbsenceTime.Update
                        End If
                        ladrsAbsenceTime.MoveNext
                    Loop
                    
                    Set ladrsAbsenceStartEnd = DetermineAbsenceStartAndEndTemp(ladrsAbsenceTime _
                                                    , ladrsSchedule _
                                                    , vadrsExceptions _
                                                    , ladrsLockStatus)
                End If
                
                If RecordsExist(ladrsAbsenceStartEnd) Then
                    ladrsAbsenceStartEnd.MoveFirst
                    
                    DetermineProbableReturnToWorkOld = DateTimeFrom(ladrsAbsenceStartEnd!RTN_TO_WRK_DT, ladrsAbsenceStartEnd!RTN_TO_WRK_TM)
                End If
            End If
        End If
    End If
    
    Exit Function
eh:
    Set lobjbusEAM = Nothing
    Set lobjdbCodeTables = Nothing
    TrsRaiseError MODULENAME, "DetermineProbableReturnToWorkOld"
    Exit Function
End Function

'//used primarily as a test method
Private Function DetermineProbableReturnToWork(ByVal rlAbsenceNumber As Long) As Date
On Error GoTo eh
    'Flag
    Dim lbCheck As Boolean

    'Lock Status
    Dim ladrsLockStatus As ADODB.Recordset

    'Time
    Dim ladrsTime As ADODB.Recordset
    Dim ladrsTimeForSchedule As ADODB.Recordset
    Dim ladrsSchedule As ADODB.Recordset
    Dim ladrsAbsenceTime As ADODB.Recordset
    Dim ldtStart As Date, ldtEnd As Date, ldtExtendedStart As Date
    Dim sErrorDates As String

    'Absence Start and End
    Dim ladrsAbsenceStartEnd As ADODB.Recordset

    'Absence
    Dim ladrsInitial As ADODB.Recordset
    Dim ladrsDate As ADODB.Recordset
    Dim ladrsReturnToWork As ADODB.Recordset

    Dim lcolDates As Collection
    'Components
    Dim lobjbusEAM As busEAM.cbsEAM
    Dim lobjdbCodeTables As dbCodeTables.cdCodeTables
    
    Dim vadrsAbsenceDates As ADODB.Recordset
    Dim vadrsInitialAbsence As ADODB.Recordset
    Dim vadrsAbsenceProvider As ADODB.Recordset
    Dim vadrsAbsenceContact As ADODB.Recordset
    Dim vadrsAbsenceReturnToWork As ADODB.Recordset
    Dim vadrsScheduledExceptions As ADODB.Recordset
    Dim vadrsAbsenceCallLog As ADODB.Recordset
    Dim vadrsExceptions As ADODB.Recordset
    Dim vlEmployeeNumber As Long
    lbCheck = True
    
    DetermineProbableReturnToWork = #1/1/100#
        
    Set lobjbusEAM = New busEAM.cbsEAM
    lobjbusEAM.GetAbsenceInformation rlAbsenceNumber, "E26895", vadrsAbsenceDates, vadrsInitialAbsence, vadrsAbsenceProvider, vadrsAbsenceContact, _
                                    vadrsAbsenceReturnToWork, vadrsAbsenceCallLog, vadrsScheduledExceptions, Nothing, Nothing
    
    Set lobjbusEAM = Nothing
    
    '//some absence times are in the db2 format '07.00.00' rather than '07:00" and from char(8) to char(5)
    If RecordsExist(vadrsInitialAbsence) Then
        vadrsInitialAbsence.MoveFirst
        Do While Not vadrsInitialAbsence.EOF
            vadrsInitialAbsence!ABS_CALLIN_TM = ConvertDB2Time(vadrsInitialAbsence!ABS_CALLIN_TM)
            vadrsInitialAbsence!ABS_ST_TM = ConvertDB2Time(vadrsInitialAbsence!ABS_ST_TM)
            vadrsInitialAbsence!ABS_END_TM = ConvertDB2Time(vadrsInitialAbsence!ABS_END_TM)
            vadrsInitialAbsence!RTN_TO_WRK_TM = ConvertDB2Time(vadrsInitialAbsence!RTN_TO_WRK_TM)
            
            vadrsInitialAbsence.Update
        
            vadrsInitialAbsence.MoveNext
        Loop
        vadrsInitialAbsence.MoveFirst
    End If
    
    Dim vadrsInitialAbsence2 As ADODB.Recordset
    
    Set vadrsInitialAbsence2 = CopyRecordset(vadrsInitialAbsence, 0)
    vadrsInitialAbsence2.Fields("ABS_CALLIN_TM").DefinedSize = 5
    vadrsInitialAbsence2.Fields("ABS_ST_TM").DefinedSize = 5
    vadrsInitialAbsence2.Fields("ABS_END_TM").DefinedSize = 5
    vadrsInitialAbsence2.Fields("RTN_TO_WRK_TM").DefinedSize = 5
    
    vadrsInitialAbsence2.Open
    AppendRecordset2 vadrsInitialAbsence, vadrsInitialAbsence2
    
    Set vadrsInitialAbsence = vadrsInitialAbsence2
    
    Set lobjdbCodeTables = New dbCodeTables.cdCodeTables
    Set vadrsExceptions = lobjdbCodeTables.GetAllExceptions()
    Set lobjdbCodeTables = Nothing
    
    '//get the schedule
    ldtStart = vbNull
    ldtEnd = vbNull
    
    If RecordsExist(vadrsAbsenceDates) And RecordsExist(vadrsInitialAbsence) Then
        vadrsInitialAbsence.MoveFirst
        vlEmployeeNumber = vadrsInitialAbsence!EMP_NUM
        ldtStart = vadrsInitialAbsence!ABS_ST_DT
    
        With vadrsAbsenceDates
            .Sort = "ABS_DATE"

            .MoveFirst
            If (ldtStart > !ABS_DATE) Then
                ldtStart = !ABS_DATE
            End If
            
            .MoveLast
            ldtEnd = !ABS_DATE

            .MoveFirst
        End With

        Set ladrsLockStatus = GetLockStatus(vlEmployeeNumber, ldtStart, ldtEnd)

        '//Set ladrsSchedule = GetEmployeeSchedule(vlEmployeeNumber, vadrsInitialAbsence, vadrsAbsenceDates)
        Set ladrsSchedule = GetEmployeeSchedule(vlEmployeeNumber, vadrsInitialAbsence, vadrsAbsenceDates)
        '//if absence is unextended, sched exc from db may prevent the probable RTW date from being determined properly
        '//try removing deleted exceptions from the schedule, this will allow the RTW to be sooner
        '//ApplyPendingChanges vadrsScheduledExceptions, ladrsSchedule
    End If
    
    '//TODO: try to modify the absence end date to be date worked based on schedule
    Dim ldtAbsStartDateTime As Date, ldtAbsEndDateTime As Date ', ldtAbsRTWDateTime As Date
'    Dim ldtAbsStartDateTimeOrig As Date, ldtAbsEndDateTimeOrig As Date, ldtAbsRTWDateTimeOrig As Date
'
'    ldtAbsStartDateTime = DateTimeFrom(vadrsInitialAbsence!ABS_ST_DT, vadrsInitialAbsence!ABS_ST_TM)
'    ldtAbsEndDateTime = DateTimeFrom(vadrsInitialAbsence!ABS_END_DT, vadrsInitialAbsence!ABS_END_TM, True)
'    ldtAbsRTWDateTime = DateTimeFrom(vadrsInitialAbsence!RTN_TO_WRK_DT, vadrsInitialAbsence!RTN_TO_WRK_TM)
'    ldtAbsStartDateTimeOrig = ldtAbsStartDateTime
'    ldtAbsEndDateTimeOrig = ldtAbsEndDateTime
'    ldtAbsRTWDateTimeOrig = ldtAbsRTWDateTime
'
''    DateTimeCrossesMidnight ldtAbsStartDateTime, ldtAbsEndDateTime
    
    '//TODO: may not need to do this, if the absence initial has been modified
    '//adjust the end datetime to be dateworked
    If RecordsExist(vadrsAbsenceDates) Then
        'NOTE: Schedule should have info atleast 30 days over the Absence End Date
        '//Set ladrsTime = GetEmployeeTimeDetails(vlEmployeeNumber, vadrsInitialAbsence, vadrsAbsenceDates)
        Set ladrsTime = GetEmployeeTimeDetails(vlEmployeeNumber, vadrsInitialAbsence, vadrsAbsenceDates, True)
        Set ladrsTimeForSchedule = GetEmployeeTimeDetails(vlEmployeeNumber, vadrsInitialAbsence, vadrsAbsenceDates, True, True)
    End If
    
    Set vadrsInitialAbsence = UpdateAbsenceWithDateWorked(vadrsInitialAbsence, ladrsTime)
    
    ldtAbsStartDateTime = DateTimeFrom(vadrsInitialAbsence!ABS_ST_DT, vadrsInitialAbsence!ABS_ST_TM)
    ldtAbsEndDateTime = DateTimeFrom(vadrsInitialAbsence!ABS_END_DT, vadrsInitialAbsence!ABS_END_TM, True)
    
'    If RecordsExist(ladrsTime) Then
'        'TODO: Refactor - FindDateAndTimeWorked(byval vdtLookup as date, byval vadrsTime as adodb.recordset, byval vblnStartTime as boolean ) as Date
'        ladrsTime.Filter = StringFormat("(TIME_TYPE = 2 AND OT_OPP_NUM <= 0 AND EXCEPTION_CD <> 'ONP' AND CHARGE_DATE = #{0}# AND START_TIME = '{1}')" & _
'                                            " OR (TIME_TYPE = 3 AND OT_OPP_NUM <= 0 AND EXCEPTION_CD <> 'ONP' AND CHARGE_DATE = #{0}# AND START_TIME = '{1}')", _
'                                            vadrsInitialAbsence!ABS_ST_DT, vadrsInitialAbsence!ABS_ST_TM)
'        If Not ladrsTime.EOF Then
'            If ladrsTime!CHARGE_DATE <> ladrsTime!DATE_WORKED Then
'                ldtAbsStartDateTime = DateTimeFrom(ladrsTime!DATE_WORKED, vadrsInitialAbsence!ABS_ST_TM)
'            End If
'        End If
'
'        ladrsTime.Filter = StringFormat("(TIME_TYPE = 2 AND OT_OPP_NUM <= 0 AND EXCEPTION_CD <> 'ONP' AND CHARGE_DATE = #{0}# AND END_TIME = '{1}')" & _
'                                            " OR (TIME_TYPE = 3 AND OT_OPP_NUM <= 0 AND EXCEPTION_CD <> 'ONP' AND CHARGE_DATE = #{0}# AND END_TIME = '{1}')", _
'                                            vadrsInitialAbsence!ABS_END_DT, vadrsInitialAbsence!ABS_END_TM)
'        If Not ladrsTime.EOF Then
'            If ladrsTime!CHARGE_DATE <> ladrsTime!DATE_WORKED Then
'                ldtAbsEndDateTime = DateTimeFrom(ladrsTime!DATE_WORKED, vadrsInitialAbsence!ABS_END_TM, True)
'            End If
'        End If
'        '//TODO: Need to retrieve time for the return to work in order to determine the date worked
'        ladrsTime.Filter = StringFormat("(TIME_TYPE = 2 AND EXCEPTION_CD <> 'ONP' AND CHARGE_DATE = #{0}# AND START_TIME = '{1}')" & _
'                                            " OR (TIME_TYPE = 3 AND EXCEPTION_CD <> 'ONP' AND CHARGE_DATE = #{0}# AND START_TIME = '{1}')", _
'                                            vadrsInitialAbsence!RTN_TO_WRK_DT, vadrsInitialAbsence!RTN_TO_WRK_TM)
'        If Not ladrsTime.EOF Then
'            If ladrsTime!CHARGE_DATE <> ladrsTime!DATE_WORKED Then
'                ldtAbsRTWDateTime = DateTimeFrom(ladrsTime!DATE_WORKED, vadrsInitialAbsence!RTN_TO_WRK_TM)
'            End If
'        Else
'            ladrsTime.Filter = StringFormat("(TIME_TYPE = 2 AND EXCEPTION_CD <> 'ONP' AND CHARGE_DATE = #{0}# AND END_TIME = '{1}')" & _
'                                                " OR (TIME_TYPE = 3 AND EXCEPTION_CD <> 'ONP' AND CHARGE_DATE = #{0}# AND END_TIME = '{1}')", _
'                                                vadrsInitialAbsence!RTN_TO_WRK_DT, vadrsInitialAbsence!RTN_TO_WRK_TM)
'            If Not ladrsTime.EOF Then
'                If ladrsTime!CHARGE_DATE <> ladrsTime!DATE_WORKED Then
'                    ldtAbsRTWDateTime = DateTimeFrom(ladrsTime!DATE_WORKED, vadrsInitialAbsence!RTN_TO_WRK_TM, True)
'                End If
'            End If
'        End If
'
'        '//the code that determines the probable return to work date is dependent on the RTW date (should it be using only the end date ?)
'        '//so if it is wrong to begin with, it will probably still be wrong afterward
'        '//making this correction will at least make it look like the end date
'        '//for partial absences, the end time is the end of the day's shift and the rtw is the end time
'        If ldtAbsRTWDateTimeOrig < ldtAbsEndDateTimeOrig Then
'            ldtAbsEndDateTime = ldtAbsRTWDateTime
'        End If
'        '//KLUDGE:
'        If ldtAbsStartDateTimeOrig = ldtAbsEndDateTimeOrig Then '//won't get any time that overlaps
'            ldtAbsEndDateTime = ldtAbsRTWDateTime
'        End If
'
'        If ldtAbsRTWDateTime < ldtAbsEndDateTime Then
'            ldtAbsEndDateTime = ldtAbsRTWDateTime
'        End If
'
'        ladrsTime.Filter = adFilterNone
'    End If
'
'    '//in case the datetimes don't match the schedule
'    DateTimeCrossesMidnight ldtAbsStartDateTime, ldtAbsEndDateTime
'
'    vadrsInitialAbsence!ABS_ST_DT = MidnightDate(ldtAbsStartDateTime)
'    vadrsInitialAbsence!ABS_ST_TM = FormatTime(ldtAbsStartDateTime)
'    vadrsInitialAbsence!ABS_END_DT = MidnightDate(ldtAbsEndDateTime)
'    vadrsInitialAbsence!ABS_END_TM = FormatTime(ldtAbsEndDateTime)
'    vadrsInitialAbsence!RTN_TO_WRK_DT = MidnightDate(ldtAbsRTWDateTime)
'    vadrsInitialAbsence!RTN_TO_WRK_TM = FormatTime(ldtAbsRTWDateTime)
'    vadrsInitialAbsence.Update
    
    '//exclude any exceptions not tied to absence, either past the return to work date or exception code based (date worked)
    '//TODO: Should this be using the absence end date or the return to work date (either way it needs to be date worked rather than charge date)
    '//the time is in this format "00.00.00"
'    ldtAbsEndDateTime = DateTimeFrom(vadrsInitialAbsence!ABS_END_DT, Format(vadrsInitialAbsence!ABS_END_TM, "hh:mm"))
    If RecordsExist(vadrsScheduledExceptions) Then
        vadrsScheduledExceptions.MoveLast
        Do While Not vadrsScheduledExceptions.BOF
            If DateTimeFrom(vadrsScheduledExceptions!CHARGE_DATE, vadrsScheduledExceptions!START_TIME) >= ldtAbsEndDateTime Or _
                DateTimeFrom(vadrsScheduledExceptions!CHARGE_DATE, vadrsScheduledExceptions!END_TIME) <= ldtAbsStartDateTime Then
'                    If Trim(radrsScheduledExceptions!EXCEPTION_CD) = "H" Then '//TODO: This could include other exceptions
                    vadrsScheduledExceptions.Delete
'                    End If
            Else
                Exit Do
            End If
            vadrsScheduledExceptions.MovePrevious
        Loop
        vadrsScheduledExceptions.MoveFirst
    End If
    
    '//check for gaps or overlaps in schedule exception requests with dates
    If lbCheck Then
        If RecordsExist(vadrsAbsenceDates) Then
            'NOTE: Schedule should have info atleast 30 days over the Absence End Date
            '//Set ladrsTime = GetEmployeeTimeDetails(vlEmployeeNumber, vadrsInitialAbsence, vadrsAbsenceDates)
'            Set ladrsTime = GetEmployeeTimeDetails(vlEmployeeNumber, vadrsInitialAbsence, vadrsAbsenceDates)
                
            'Determine if non consecutive dates were selected
            If RecordsExist(ladrsSchedule) Or RecordsExist(ladrsTime) Then
                Set ladrsAbsenceTime = AddTimeAndSchExceptions(vadrsInitialAbsence _
                                                , vadrsScheduledExceptions _
                                                , ladrsSchedule _
                                                , ladrsTime _
                                                , ladrsLockStatus)
                If RecordsExist(ladrsAbsenceTime) Then '//fails on .MoveLast when no records
                    '//actual time seems to get deleted
                    ladrsAbsenceTime.MoveFirst
                    Do While Not ladrsAbsenceTime.EOF
                        If Trim(ladrsAbsenceTime!REC_STATUS) = "" Then
                            ladrsAbsenceTime!REC_STATUS = "L"
                            ladrsAbsenceTime.Update
                        End If
                        ladrsAbsenceTime.MoveNext
                    Loop
                    
                    Set ladrsAbsenceStartEnd = DetermineAbsenceStartAndEnd(ladrsAbsenceTime _
                                                    , ladrsTimeForSchedule _
                                                    , vadrsExceptions _
                                                    , ladrsLockStatus _
                                                    , ladrsTime)
                End If
                
                If RecordsExist(ladrsAbsenceStartEnd) Then
                    ladrsAbsenceStartEnd.MoveFirst
                    
                    DetermineProbableReturnToWork = DateTimeFrom(ladrsAbsenceStartEnd!RTN_TO_WRK_DT, ladrsAbsenceStartEnd!RTN_TO_WRK_TM)
                End If
            End If
        End If
    End If
    
    Exit Function
eh:
    Set lobjbusEAM = Nothing
    Set lobjdbCodeTables = Nothing
    TrsRaiseError MODULENAME, "DetermineProbableReturnToWork"
    Exit Function
End Function

'//add date worked columns for start, end, and return to work
Private Function UpdateAbsenceWithDateWorked(ByVal vadrsInitialAbsence As ADODB.Recordset, ByVal ladrsTime As ADODB.Recordset) As ADODB.Recordset
    '//TODO: try to modify the absence end date to be date worked based on schedule
    Dim ldtAbsStartDateTime As Date, ldtAbsEndDateTime As Date, ldtAbsRTWDateTime As Date
    Dim ldtAbsStartDateTimeOrig As Date, ldtAbsEndDateTimeOrig As Date, ldtAbsRTWDateTimeOrig As Date
    
    '//copy recordset and add new columns
    Dim ladrsInitialAbsence As ADODB.Recordset
    
    Set ladrsInitialAbsence = CopyRecordset(vadrsInitialAbsence, 0)
    ladrsInitialAbsence.Fields.Append "ST_DATE_WORKED", adDate
    ladrsInitialAbsence.Fields.Append "END_DATE_WORKED", adDate
    ladrsInitialAbsence.Fields.Append "RTW_DATE_WORKED", adDate
    
    ladrsInitialAbsence.Open
    AppendRecordset2 vadrsInitialAbsence, ladrsInitialAbsence
    
    If Not ladrsInitialAbsence.EOF Then
        ladrsInitialAbsence.MoveFirst
    End If
    
    ldtAbsStartDateTimeOrig = DateTimeFrom2(ladrsInitialAbsence!ABS_ST_DT, ladrsInitialAbsence!ABS_ST_TM)
    '//even though this is an end time, don't adjust the date for midnight (end time is "00:00"), since it is a charge date
    ldtAbsEndDateTimeOrig = DateTimeFrom2(ladrsInitialAbsence!ABS_END_DT, ladrsInitialAbsence!ABS_END_TM)
    ldtAbsRTWDateTimeOrig = DateTimeFrom2(ladrsInitialAbsence!RTN_TO_WRK_DT, ladrsInitialAbsence!RTN_TO_WRK_TM)
    
    ldtAbsStartDateTime = FindDateAndTimeWorked(ladrsInitialAbsence!ABS_ST_DT, ladrsInitialAbsence!ABS_ST_TM, ladrsTime)
    ldtAbsEndDateTime = FindDateAndTimeWorked(ladrsInitialAbsence!ABS_END_DT, ladrsInitialAbsence!ABS_END_TM, ladrsTime)
    ldtAbsRTWDateTime = FindDateAndTimeWorked(ladrsInitialAbsence!RTN_TO_WRK_DT, ladrsInitialAbsence!RTN_TO_WRK_TM, ladrsTime)
       
'    '//the code that determines the probable return to work date is dependent on the RTW date (should it be using only the end date ?)
'    '//so if it is wrong to begin with, it will probably still be wrong afterward
'    '//making this correction will at least make it look like the end date
'    '//for partial absences, the end time is the end of the day's shift and the rtw is the end time
'    If ldtAbsRTWDateTimeOrig < ldtAbsEndDateTimeOrig Then
'        ldtAbsEndDateTime = ldtAbsRTWDateTime
'    End If
'    '//KLUDGE:
'    If ldtAbsStartDateTimeOrig = ldtAbsEndDateTimeOrig Then '//won't get any time that overlaps
'        ldtAbsEndDateTime = ldtAbsRTWDateTime
'    End If
'
'    If ldtAbsRTWDateTime < ldtAbsEndDateTime Then
'        ldtAbsEndDateTime = ldtAbsRTWDateTime
'    End If
'
'    If ldtAbsEndDateTimeOrig = #1/1/100# Then
'        ldtAbsEndDateTime = ldtAbsRTWDateTime
'    End If
    
    '//in case the datetimes don't match the schedule
'    DateTimeCrossesMidnight ldtAbsStartDateTime, ldtAbsEndDateTime
'    DateTimeCrossesMidnight ldtAbsStartDateTime, ldtAbsRTWDateTime
    
'    ladrsInitialAbsence!ABS_ST_DT = MidnightDate(ldtAbsStartDateTimeOrig)
    ladrsInitialAbsence!ST_DATE_WORKED = MidnightDate(ldtAbsStartDateTime)
'    ladrsInitialAbsence!ABS_ST_TM = FormatTime(ldtAbsStartDateTimeOrig)
    
'    ladrsInitialAbsence!ABS_END_DT = MidnightDate(ldtAbsEndDateTimeOrig)
    ladrsInitialAbsence!END_DATE_WORKED = MidnightDate(ldtAbsEndDateTime)
'    ladrsInitialAbsence!ABS_END_TM = FormatTime(ldtAbsEndDateTimeOrig)
    
'    ladrsInitialAbsence!RTN_TO_WRK_DT = MidnightDate(ldtAbsRTWDateTimeOrig)
    ladrsInitialAbsence!RTW_DATE_WORKED = MidnightDate(ldtAbsRTWDateTime)
'    ladrsInitialAbsence!RTN_TO_WRK_TM = FormatTime(ldtAbsRTWDateTimeOrig)
    ladrsInitialAbsence.Update
    
    Set UpdateAbsenceWithDateWorked = ladrsInitialAbsence
    Exit Function
eh:
    TrsRaiseError MODULENAME, "UpdateAbsenceWithDateWorked"
    Exit Function
End Function

'//-----end-----//
'Private Enum TimeIt
'    StartTime = 0
'    EndTime
'    EitherStartOrEnd
'End Enum

Private Function FindDateTimeWorked(ByVal vdtChargeDate As Date, ByVal vadrsTime As ADODB.Recordset) As Date
On Error GoTo eh
    FindDateTimeWorked = FindDateAndTimeWorked(MidnightDate(vdtChargeDate), FormatTime(vdtChargeDate), vadrsTime)
    Exit Function
eh:
    TrsRaiseError MODULENAME, "FindDateAndTimeWorked"
    Exit Function
End Function

Private Function FindDateAndTimeWorked(ByVal vdtChargeDate As Date, ByVal vsTime As String, ByVal vadrsTime As ADODB.Recordset) As Date
On Error GoTo eh
    Dim ldtDateWorked As Date
    Dim lblnFound As Boolean
    
    ldtDateWorked = DateTimeFrom2(vdtChargeDate, vsTime)
    lblnFound = False
    
    If RecordsExist(vadrsTime) Then
        If Not lblnFound Then
            vadrsTime.Filter = StringFormat("(TIME_TYPE = 2 AND OT_OPP_NUM <= 0 AND EXCEPTION_CD <> 'ONP' AND CHARGE_DATE = #{0}# AND START_TIME = '{1}')" & _
                                                " OR (TIME_TYPE = 3 AND OT_OPP_NUM <= 0 AND EXCEPTION_CD <> 'ONP' AND CHARGE_DATE = #{0}# AND START_TIME = '{1}')", _
                                                vdtChargeDate, vsTime)
            If Not vadrsTime.EOF Then
                lblnFound = True
'                If vadrsTime!CHARGE_DATE <> vadrsTime!DATE_WORKED Then
                    ldtDateWorked = DateTimeFrom2(vadrsTime!DATE_WORKED, vsTime)
'                End If
            End If
            vadrsTime.Filter = adFilterNone
        End If
        
        If Not lblnFound Then
            vadrsTime.Filter = StringFormat("(TIME_TYPE = 2 AND OT_OPP_NUM <= 0 AND EXCEPTION_CD <> 'ONP' AND CHARGE_DATE = #{0}# AND END_TIME = '{1}')" & _
                                                " OR (TIME_TYPE = 3 AND OT_OPP_NUM <= 0 AND EXCEPTION_CD <> 'ONP' AND CHARGE_DATE = #{0}# AND END_TIME = '{1}')", _
                                                vdtChargeDate, vsTime)
            If Not vadrsTime.EOF Then
                lblnFound = True
'                If vadrsTime!CHARGE_DATE <> vadrsTime!DATE_WORKED Then
                    '//even though this is date worked and an end time, the date worked really applies to the start time, so the end time needs reevaluated
                    '//Example: Time info is: Charge_Date 8/14/2013, Date_Worked 8/14/2013, Start_Time 20:00, End_Time 00:00
                    '//Example: Adjust the end date worked to 8/15/2013 (meaning end of day for 8/14/2013)
                    ldtDateWorked = DateTimeFrom2(vadrsTime!DATE_WORKED, vsTime, True)
'                End If
            End If
            vadrsTime.Filter = adFilterNone
        End If
    
        If Not lblnFound Then
            vadrsTime.Filter = StringFormat("(TIME_TYPE = 2 AND OT_OPP_NUM <= 0 AND EXCEPTION_CD <> 'ONP' AND CHARGE_DATE = #{0}#)" & _
                                                " OR (TIME_TYPE = 3 AND OT_OPP_NUM <= 0 AND EXCEPTION_CD <> 'ONP' AND CHARGE_DATE = #{0}#)", _
                                                vdtChargeDate)
            If Not vadrsTime.EOF Then
                Dim ldtChargeDateTime As Date
                
                ldtChargeDateTime = DateTimeFrom2(vdtChargeDate, vsTime)
                vadrsTime.MoveFirst
                Do While Not vadrsTime.EOF And Not lblnFound
                    Dim dtStartTime As Date, dtEndTime As Date
                    '//TODO: Don't look at TTA where the time will be invalid and cause a type mismatch
                    dtStartTime = DateTimeFrom2(vadrsTime!CHARGE_DATE, vadrsTime!START_TIME)
                    '//TEST:
                    '//even though this is date worked and an end time, the date worked really applies to the start time, so the end time needs reevaluated
                    dtEndTime = DateTimeFrom2(vadrsTime!CHARGE_DATE, vadrsTime!END_TIME, True)
                    If ldtChargeDateTime >= dtStartTime And ldtChargeDateTime <= dtEndTime Then
                        If vadrsTime!CHARGE_DATE <> vadrsTime!DATE_WORKED Then
                            ldtDateWorked = DateTimeFrom2(vadrsTime!DATE_WORKED, vsTime)
                        End If
                    
                        lblnFound = True
                        Exit Do
                    End If
                    vadrsTime.MoveNext
                Loop
            End If
            vadrsTime.Filter = adFilterNone
        End If
    End If
    
    FindDateAndTimeWorked = ldtDateWorked
    Exit Function
eh:
    TrsRaiseError MODULENAME, "FindDateAndTimeWorked"
    Exit Function
End Function

'//added a call to this from PrepareDeleteExcTimeRS, but it was failing with error 13 type mismatch,
'//because the time field was formated normally with a colon, so allow either period or colon

Private Function ConvertDB2Time(ByVal db2Time As String) As String
    '(\d{1,2}\.\d{1,2}\.\d{1,2})
    '((?<Hours>\d{1,2})\.(?<Minutes>\d{1,2})\.(?<Seconds>\d{1,2}))
    '${Hours}:${Minutes}
    '"${2}:${3}"
    Dim g As Variant
On Error GoTo eh

    g = GetRegExGroups(db2Time, "((\d{1,2})[.:](\d{1,2})[.:]?(\d{1,2})?)")
    
    If UBound(g) >= 2 Then
        ConvertDB2Time = g(1) & ":" & g(2)
    End If
    
    Exit Function
eh:
    TrsRaiseError MODULENAME, "ConvertDB2Time"
    Exit Function
End Function

Private Function GetRegExGroups(ByVal Source As String, ByVal Pattern As String) As Variant
    Dim re As RegExp
    Dim matches As MatchCollection
    Dim groups() As Variant
    Dim i As Integer
    
On Error GoTo eh
    Set re = New RegExp
  
    re.IgnoreCase = True
    re.Global = True
    re.Pattern = Pattern
    
    ReDim groups(0)
    
    Set matches = re.Execute(Source)
    
    If matches.Count > 0 Then
        'Print matches.Item(0).SubMatches.Item(1)
        If matches.Item(0).SubMatches.Count > 0 Then
            ReDim groups(matches.Item(0).SubMatches.Count - 1)
            For i = 0 To matches.Item(0).SubMatches.Count - 1
                groups(i) = matches.Item(0).SubMatches.Item(i)
            Next i
        End If
    End If
    
    GetRegExGroups = groups
    Set re = Nothing
    Exit Function
eh:
    TrsRaiseError MODULENAME, "GetRegExGroups"
    Exit Function
End Function

'//for testing (used for testing PreSaveValidate2)
Public Function ExtendAbsence(ByVal vlAbsenceNumber As Long _
                , ByVal vdtExtendToDate As Date _
                , ByVal vlTRISUserNumber As Long _
                , ByVal vsTRISUserID As String _
                , ByVal vblnValidateOnly As Boolean _
                , ByRef radrsBRV As ADODB.Recordset _
                ) As Boolean
On Error GoTo eh
    Dim lEmployeeNumber As Long, sExceptionCode As String
    Dim ladrsAbsenceDates As ADODB.Recordset
    Dim ladrsInitialAbsence As ADODB.Recordset
    Dim ladrsAbsenceProvider As ADODB.Recordset
    Dim ladrsAbsenceContact As ADODB.Recordset
    Dim ladrsAbsenceReturnToWork As ADODB.Recordset
    Dim ladrsScheduledExceptions As ADODB.Recordset
    Dim ladrsAbsenceCallLog As ADODB.Recordset
    Dim ladrsExceptions As ADODB.Recordset
    Dim ladrsEmployeeInformation As ADODB.Recordset
    Dim ldbCodeTables As dbCodeTables.cdCodeTables
    Dim lbusEAM As busEAM.cbsEAM
    Dim ladrsRTWExemptExceptions As ADODB.Recordset, ladrsWCExceptions As ADODB.Recordset
    Dim ladrsGetAbsenceDates As ADODB.Recordset
    Dim ladrsGetAbsInitial As ADODB.Recordset
    Dim ladrsGetAbsProvider As ADODB.Recordset
    Dim ladrsGetAbsContact As ADODB.Recordset
    Dim ladrsGetAbsReturnToWork As ADODB.Recordset
    Dim ladrsGetAbsCallLog As ADODB.Recordset
    Dim ladrsGetScheduledExceptions As ADODB.Recordset
    Dim ladrsGetTimeDetails As ADODB.Recordset
    Dim dtReturnToWork As Date
        
    Set ladrsInitialAbsence = m_ResourceMgr.GetRecordSetResource(m_DBSCRIPTSROOT, m_DBSCRIPTSPATH, "ExtendAbsence_InitialAbsence.xml")
    Set ladrsAbsenceDates = m_ResourceMgr.GetRecordSetResource(m_DBSCRIPTSROOT, m_DBSCRIPTSPATH, "ExtendAbsence_AbsenceDates.xml")
    Set ladrsScheduledExceptions = m_ResourceMgr.GetRecordSetResource(m_DBSCRIPTSROOT, m_DBSCRIPTSPATH, "ExtendAbsence_ScheduledExceptions.xml")
    
    Set lbusEAM = CreateObject("busEAM.cbsEAM")
    lbusEAM.GetAbsenceInformation vlAbsenceNumber _
        , "" _
        , ladrsGetAbsenceDates _
        , ladrsGetAbsInitial _
        , ladrsGetAbsProvider _
        , ladrsGetAbsContact _
        , ladrsGetAbsReturnToWork _
        , ladrsGetAbsCallLog _
        , ladrsGetScheduledExceptions _
        , ladrsGetTimeDetails _
        , Nothing
    Set lbusEAM = Nothing
       
'    dtReturnToWork = DetermineProbableReturnToWork(vlAbsenceNumber)
    
    AppendRecordsetByName ladrsGetAbsInitial, ladrsInitialAbsence
    '//change the format for time fields (e.g. '07.00' to '07:00')
    If RecordsExist(ladrsInitialAbsence) Then
        ladrsInitialAbsence.MoveFirst
        Do While Not ladrsInitialAbsence.EOF
'            ladrsInitialAbsence.Fields("EMP_NUM").Value = Null
            ladrsInitialAbsence.Fields("ABS_ST_TM").Value = ConvertDB2Time(ladrsInitialAbsence.Fields("ABS_ST_TM").Value)
            ladrsInitialAbsence.Fields("ABS_END_TM").Value = ConvertDB2Time(ladrsInitialAbsence.Fields("ABS_END_TM").Value)
            ladrsInitialAbsence.Fields("ABS_CALLIN_TM").Value = ConvertDB2Time(ladrsInitialAbsence.Fields("ABS_CALLIN_TM").Value)
            '//ladrsInitialAbsence.Fields("RTN_TO_WRK_TM").Value = ConvertDB2Time(ladrsInitialAbsence.Fields("RTN_TO_WRK_TM").Value)
            ladrsInitialAbsence.Fields("RTN_TO_WRK_DT").Value = MidnightDate(dtReturnToWork)
            ladrsInitialAbsence.Fields("RTN_TO_WRK_TM").Value = FormatTime(dtReturnToWork)
            ladrsInitialAbsence.Fields("RTW_NOT_REQD_IND").Value = IIf(ladrsInitialAbsence.Fields("RTW_NOT_REQD_IND").Value = "0", "N", "Y")
            ladrsInitialAbsence.Fields("RECORD_STATUS").Value = "C"
            ladrsInitialAbsence.Update
            
            ladrsInitialAbsence.MoveNext
        Loop
    End If
    
    If RecordsExist(ladrsInitialAbsence) Then
        ladrsInitialAbsence.MoveFirst
        lEmployeeNumber = ladrsInitialAbsence.Fields("EMP_NUM").Value
        sExceptionCode = ladrsInitialAbsence.Fields("EXCEPTION_CD").Value
    End If
    
    Set ladrsAbsenceProvider = Nothing
    Set ladrsAbsenceContact = Nothing
    Set ladrsAbsenceReturnToWork = Nothing
    Set ladrsAbsenceCallLog = Nothing
    
    Set ldbCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    Set ladrsExceptions = ldbCodeTables.GetAllExceptions
    Set ldbCodeTables = Nothing
    
    Set lbusEAM = CreateObject("busEAM.cbsEAM")
    Set ladrsEmployeeInformation = lbusEAM.GetEAMEmployeeInformation3(lEmployeeNumber, "", 0, ladrsRTWExemptExceptions, ladrsWCExceptions)
    Set lbusEAM = Nothing
        
    Dim ladrsSchedule As ADODB.Recordset, ladrsTime As ADODB.Recordset, ladrsTimeForSchedule As ADODB.Recordset
    Set ladrsSchedule = GetEmployeeSchedule(lEmployeeNumber, ladrsInitialAbsence, ladrsGetAbsenceDates)

    Dim ldtAbsStartDateTime As Date, ldtAbsEndDateTime As Date
    Dim lsStartTime As String

    '//TODO: may not need to do this, if the absence initial has been modified
    '//adjust the end datetime to be dateworked
    If RecordsExist(ladrsGetAbsenceDates) Then
        'NOTE: Schedule should have info atleast 30 days over the Absence End Date
        Set ladrsTime = GetEmployeeTimeDetails(lEmployeeNumber, ladrsInitialAbsence, ladrsGetAbsenceDates, True)
        Set ladrsTimeForSchedule = GetEmployeeTimeDetails(lEmployeeNumber, ladrsInitialAbsence, ladrsGetAbsenceDates, True, True)
    End If
    
'    Set ladrsInitialAbsence = UpdateAbsenceWithDateWorked(ladrsInitialAbsence, ladrsTime)
    
    ldtAbsStartDateTime = DateTimeFrom(ladrsInitialAbsence!ABS_ST_DT, ladrsInitialAbsence!ABS_ST_TM)
    ldtAbsEndDateTime = DateTimeFrom(ladrsInitialAbsence!ABS_END_DT, ladrsInitialAbsence!ABS_END_TM, True)
    
    Dim ldtStart As Date, ldtEnd As Date
    'Dim sCheckForDuplicate As String
    Dim sLastDataSaved(0 To 3) As Variant
    
    lsStartTime = ""
    '//create scheduled exceptions from schedule only (need to have only one row per day)
    If RecordsExist(ladrsSchedule) Then
        ladrsSchedule.Sort = "EMP_NUM, CHARGE_DATE, SHIFT_START_DATE, SHIFT_START_TIME"
        ladrsSchedule.MoveFirst
        Do While Not ladrsSchedule.EOF
            If ladrsSchedule.Fields("SHIFT_HOURS").Value > 0 Then '//working days
'                If ladrsSchedule.Fields("TIME_TYPE").Value = "3" Then '//schedule (mimic the client issue)
                    '//filter between abs end datetime and extended datetime
                    ldtStart = DateTimeFrom(ladrsSchedule!SHIFT_START_DATE, ladrsSchedule!SHIFT_START_TIME)
                    '//TODO: Does this need to be specified as an end time, since it is a dateworked ???
                    ldtEnd = DateTimeFrom(ladrsSchedule!SHIFT_END_DATE, ladrsSchedule!SHIFT_END_TIME, True)
                    If TimePeriodsIntersect(ldtStart, ldtEnd, ldtAbsEndDateTime, DateTimeFrom(vdtExtendToDate, "24:00")) Or Trim(lsStartTime) <> "" Then
                        '//add to sched except and suppress duplicates
                        If sLastDataSaved(0) <> lEmployeeNumber Or _
                            sLastDataSaved(1) <> ladrsSchedule.Fields("CHARGE_DATE").Value Or _
                            sLastDataSaved(2) <> ladrsSchedule.Fields("SHIFT_START_TIME").Value Or _
                            sLastDataSaved(3) <> ladrsSchedule.Fields("SHIFT_END_TIME").Value Then
                            
                            ladrsScheduledExceptions.AddNew
                            ladrsScheduledExceptions.Fields("ABS_NUM").Value = vlAbsenceNumber
                            ladrsScheduledExceptions.Fields("EMP_NUM").Value = lEmployeeNumber
                            ladrsScheduledExceptions.Fields("CHARGE_DATE").Value = ladrsSchedule.Fields("CHARGE_DATE").Value
                            ladrsScheduledExceptions.Fields("SHIFT_NUM").Value = ladrsSchedule.Fields("SHIFT_NUM").Value
                            ladrsScheduledExceptions.Fields("FULL_SHIFT_EXC_IND").Value = "N"
                            ladrsScheduledExceptions.Fields("START_TIME").Value = ladrsSchedule.Fields("SHIFT_START_TIME").Value

                            ladrsScheduledExceptions.Fields("END_TIME").Value = ladrsSchedule.Fields("SHIFT_END_TIME").Value
                            ladrsScheduledExceptions.Fields("EXCEPTION_CD").Value = sExceptionCode
                            '//ladrsScheduledExceptions.Fields("OVER_EXC_CD").Value = /ladrsSchedule.Fields("OVER_EXC_CD").Value
                            '//ladrsScheduledExceptions.Fields("KEY_CHARGE_DATE").Value = vbNull
                            ladrsScheduledExceptions.Fields("KEY_CHARGE_DATE").Value = DateAdd("d", -1, vbNull)
                            
                            ladrsScheduledExceptions.Fields("DELETE_ACTUAL_TIME").Value = "Y"
                            ladrsScheduledExceptions.Fields("APPR_AUTH_IND").Value = "Y"
                            ladrsScheduledExceptions.Fields("CREATED_TIMESTAMP").Value = "" 'Now()
                            ladrsScheduledExceptions.Fields("REC_STATUS").Value = "I"
                            ladrsScheduledExceptions.Update
                            
                            '//Save for next row and then check if duplicate
                            sLastDataSaved(0) = ladrsScheduledExceptions.Fields("EMP_NUM").Value
                            sLastDataSaved(1) = ladrsScheduledExceptions.Fields("CHARGE_DATE").Value
                            '//sLastDataSaved(2) = ladrsScheduledExceptions.Fields("DATE_WORKED").Value
                            sLastDataSaved(2) = ladrsScheduledExceptions.Fields("START_TIME").Value
                            sLastDataSaved(3) = ladrsScheduledExceptions.Fields("END_TIME").Value

                            lsStartTime = ""
                        End If
                    End If
                    
                    If dtReturnToWork = DateAdd("d", -1, vbNull) Then '//do once
                        If ldtStart >= DateTimeFrom(vdtExtendToDate, "24:00") Then
                            dtReturnToWork = DateTimeFrom(ladrsSchedule!SHIFT_START_DATE, ladrsSchedule!SHIFT_START_TIME)
                        End If
                    End If
'                End If
            End If
            ladrsSchedule.MoveNext
        Loop
    End If
    
    If RecordsExist(ladrsInitialAbsence) Then
        ladrsInitialAbsence.MoveFirst
        Do While Not ladrsInitialAbsence.EOF
            ladrsInitialAbsence.Fields("EMP_NUM").Value = Null '//mimic client
            ladrsInitialAbsence.Fields("RTN_TO_WRK_DT").Value = MidnightDate(dtReturnToWork)
            ladrsInitialAbsence.Fields("RTN_TO_WRK_TM").Value = FormatTime(dtReturnToWork)
            ladrsInitialAbsence.Fields("RECORD_STATUS").Value = "C"
            ladrsInitialAbsence.Update
            
            ladrsInitialAbsence.MoveNext
        Loop
    End If
    
'    '//create scheduled exceptions from merged schedule and time (trouble is that it contains more than one row per day)
'    If RecordsExist(ladrsTimeForSchedule) Then
'        ladrsTimeForSchedule.Sort = "EMP_NUM, CHARGE_DATE, DATE_WORKED, START_TIME"
'        ladrsTimeForSchedule.MoveFirst
'        Do While Not ladrsTimeForSchedule.EOF
'            If ladrsTimeForSchedule.Fields("HOURS").Value > 0 Then
''                If ladrsTimeForSchedule.Fields("TIME_TYPE").Value = "3" Then '//schedule (mimic the client issue)
'                    '//filter between abs end datetime and extended datetime
'                    ldtStart = DateTimeFrom(ladrsTimeForSchedule!DATE_WORKED, ladrsTimeForSchedule!START_TIME)
'                    '//TODO: Does this need to be specified as an end time, since it is a dateworked ???
'                    ldtEnd = DateTimeFrom(ladrsTimeForSchedule!DATE_WORKED, ladrsTimeForSchedule!END_TIME, True)
'                    If TimePeriodsIntersect(ldtStart, ldtEnd, ldtAbsEndDateTime, DateTimeFrom(vdtExtendToDate, "24:00")) Or Trim(lsStartTime) <> "" Then
'                        '//add to sched except
'                        If ladrsTimeForSchedule.Fields("END_TIME").Value = "00:00" Then '//suppress until next row
'                            lsStartTime = ladrsTimeForSchedule.Fields("START_TIME").Value
'                        Else
'                            ladrsScheduledExceptions.AddNew
'                            ladrsScheduledExceptions.Fields("ABS_NUM").Value = vlAbsenceNumber
'                            ladrsScheduledExceptions.Fields("EMP_NUM").Value = lEmployeeNumber
'                            ladrsScheduledExceptions.Fields("CHARGE_DATE").Value = ladrsTimeForSchedule.Fields("CHARGE_DATE").Value
'                            If ladrsTimeForSchedule.Fields("TIME_TYPE").Value = 2 Then '//actual time
'                                ladrsScheduledExceptions.Fields("SHIFT_NUM").Value = "1"
'                            Else
'                                ladrsScheduledExceptions.Fields("SHIFT_NUM").Value = ladrsTimeForSchedule.Fields("SHIFT_NUM").Value
'                            End If
'                            ladrsScheduledExceptions.Fields("FULL_SHIFT_EXC_IND").Value = "N"
'                            If ladrsTimeForSchedule.Fields("START_TIME").Value = "00:00" Then '//use previous value before midnight
'                                ladrsScheduledExceptions.Fields("START_TIME").Value = lsStartTime
'                            Else
'                                ladrsScheduledExceptions.Fields("START_TIME").Value = ladrsTimeForSchedule.Fields("START_TIME").Value
'                            End If
'                            ladrsScheduledExceptions.Fields("END_TIME").Value = ladrsTimeForSchedule.Fields("END_TIME").Value
'                            ladrsScheduledExceptions.Fields("EXCEPTION_CD").Value = sExceptionCode
'                            ladrsScheduledExceptions.Fields("OVER_EXC_CD").Value = ladrsTimeForSchedule.Fields("OVER_EXC_CD").Value
'                            ladrsScheduledExceptions.Fields("KEY_CHARGE_DATE").Value = vbNull
'
'                            ladrsScheduledExceptions.Fields("DELETE_ACTUAL_TIME").Value = "Y"
'                            ladrsScheduledExceptions.Fields("CREATED_TIMESTAMP").Value = "" 'Now()
'                            ladrsScheduledExceptions.Fields("REC_STATUS").Value = "I"
'                            ladrsScheduledExceptions.Update
'                            lsStartTime = ""
'                        End If
'                    End If
''                End If
'            End If
'            ladrsTimeForSchedule.MoveNext
'        Loop
'    End If
    
    If RecordsExist(ladrsScheduledExceptions) Then
        ladrsScheduledExceptions.MoveFirst
        Do While Not ladrsScheduledExceptions.EOF
            '//add to dates
            ladrsAbsenceDates.Find "ABS_DATE = '" & ladrsScheduledExceptions!CHARGE_DATE & "'"
            If ladrsAbsenceDates.EOF Then
                ladrsAbsenceDates.AddNew
                ladrsAbsenceDates.Fields("ABS_NUM").Value = vlAbsenceNumber
                ladrsAbsenceDates.Fields("ABS_DATE").Value = ladrsScheduledExceptions!CHARGE_DATE
                ladrsAbsenceDates.Fields("CREATED_TIMESTAMP").Value = Null '"" 'Now()
                ladrsAbsenceDates.Fields("RECORD_STATUS").Value = "I"
                ladrsAbsenceDates.Update
            End If
            ladrsScheduledExceptions.MoveNext
        Loop
    End If
    
    If vblnValidateOnly Then
        ExtendAbsence = PreSaveValidate2(lEmployeeNumber, ladrsAbsenceDates, ladrsInitialAbsence _
                    , ladrsAbsenceProvider, ladrsAbsenceContact, ladrsAbsenceReturnToWork, ladrsScheduledExceptions _
                    , ladrsAbsenceCallLog, ladrsExceptions, ladrsEmployeeInformation, vlTRISUserNumber, vsTRISUserID, "Warning", "Error", vlAbsenceNumber, radrsBRV)
    Else
        ExtendAbsence = ProcessEAMRequest(lEmployeeNumber, ladrsAbsenceDates, ladrsInitialAbsence _
                    , ladrsAbsenceProvider, ladrsAbsenceContact, ladrsAbsenceReturnToWork, ladrsScheduledExceptions _
                    , ladrsAbsenceCallLog, ladrsExceptions, ladrsEmployeeInformation, vlTRISUserNumber, vsTRISUserID, vlAbsenceNumber, radrsBRV)
    End If
    Exit Function
eh:
    ExtendAbsence = False
    
    Select Case (Err.Number)
    Case trErrutOpenConnection
        TrsRaiseError MODULENAME, "ExtendAbsence"
    Case Else
        TrsRaiseError MODULENAME, "ExtendAbsence", Err.Number, , , , True
    End Select

    Exit Function
End Function

'//copy by name instead of by index
Private Sub AppendRecordsetByName( _
            ByVal vrsSource As ADODB.Recordset, _
            ByRef vrsTarget As ADODB.Recordset, _
            Optional ByVal vintAffect As Integer = 2, _
            Optional ByVal Position As Integer = 0)
' copy records from the source recordset to the target recordset.
' both recordsets don't need to have the same definition and open.
' vintAffect determines the extent of the copy operation.
' 0 - copy metadata only, no effect for append operation
' 1 - copy current record only
' 2 - copy all records visible
    Dim lrs As ADODB.Recordset
    Dim i As Integer
    Dim sColumnName As String
On Error GoTo eh
    If RecordsExist(vrsSource) Then
        vrsSource.MoveFirst
    Else
        Exit Sub
    End If

    If vintAffect = 0 Then
        'do nothing
    ElseIf vintAffect = 2 Then
        'CopyAll
        Dim j As Integer
        For j = 1 To vrsSource.RecordCount
            vrsTarget.AddNew
            For i = 0 To vrsSource.Fields.Count - 1
                sColumnName = vrsSource.Fields(i).Name
                '//TODO: make more efficient
                If mADOExtensions.FieldExists(sColumnName, vrsTarget) Then
                    If vrsSource.Fields(sColumnName).Type = adChar And vrsTarget.Fields(sColumnName).DefinedSize < vrsSource.Fields(sColumnName).DefinedSize Then
                        vrsTarget.Fields(sColumnName).Value = Left(vrsSource.Fields(sColumnName).Value, vrsTarget.Fields(sColumnName).DefinedSize)
                    Else
                        vrsTarget.Fields(sColumnName).Value = vrsSource.Fields(sColumnName).Value
                    End If
                End If
            Next
            vrsTarget.Update
            vrsSource.MoveNext
        Next
    Else
        'Copy only the position
        vrsSource.AbsolutePosition = Position
        vrsTarget.AddNew
        For i = 0 To vrsSource.Fields.Count - 1
            sColumnName = vrsSource.Fields(i).Name
            '//TODO: make more efficient
            If mADOExtensions.FieldExists(sColumnName, vrsTarget) Then
                If vrsSource.Fields(sColumnName).Type = adChar And vrsTarget.Fields(sColumnName).DefinedSize < vrsSource.Fields(sColumnName).DefinedSize Then
                    vrsTarget.Fields(sColumnName).Value = Left(vrsSource.Fields(sColumnName).Value, vrsTarget.Fields(sColumnName).DefinedSize)
                Else
                    vrsTarget.Fields(sColumnName).Value = vrsSource.Fields(sColumnName).Value
                End If
            End If
        Next
        vrsTarget.Update
    End If
    
    Exit Sub
    
eh:
    Err.Raise Err.Number, "AppendRecordsetByName", Err.Description
End Sub

Public Function PreSaveValidate(ByVal vlEmployeeNumber As Long _
                , ByVal vadrsAbsenceDates As ADODB.Recordset _
                , ByVal vadrsInitialAbsence As ADODB.Recordset _
                , ByVal vadrsAbsenceProvider As ADODB.Recordset _
                , ByVal vadrsAbsenceContact As ADODB.Recordset _
                , ByVal vadrsAbsenceReturnToWork As ADODB.Recordset _
                , ByVal vadrsScheduledExceptions As ADODB.Recordset _
                , ByVal vadrsAbsenceCallLog As ADODB.Recordset _
                , ByVal vadrsExceptions As ADODB.Recordset _
                , ByVal vadrsEmployeeInformation As ADODB.Recordset _
                , ByVal vlTRISUserNumber As Long _
                , ByVal vsTRISUserID As String _
                , ByRef rlAbsenceNumber As Long _
                , ByRef radrsBRV As ADODB.Recordset _
                ) As Boolean
                
        PreSaveValidate = PreSaveValidate2(vlEmployeeNumber, vadrsAbsenceDates, vadrsInitialAbsence _
                , vadrsAbsenceProvider, vadrsAbsenceContact, vadrsAbsenceReturnToWork, vadrsScheduledExceptions _
                , vadrsAbsenceCallLog, vadrsExceptions, vadrsEmployeeInformation, vlTRISUserNumber, vsTRISUserID, "Warning", "Error", rlAbsenceNumber, radrsBRV)
End Function

Public Function PreSaveValidate2(ByVal vlEmployeeNumber As Long _
                , ByVal vadrsAbsenceDates As ADODB.Recordset _
                , ByVal vadrsInitialAbsence As ADODB.Recordset _
                , ByVal vadrsAbsenceProvider As ADODB.Recordset _
                , ByVal vadrsAbsenceContact As ADODB.Recordset _
                , ByVal vadrsAbsenceReturnToWork As ADODB.Recordset _
                , ByVal vadrsScheduledExceptions As ADODB.Recordset _
                , ByVal vadrsAbsenceCallLog As ADODB.Recordset _
                , ByVal vadrsExceptions As ADODB.Recordset _
                , ByVal vadrsEmployeeInformation As ADODB.Recordset _
                , ByVal vlTRISUserNumber As Long _
                , ByVal vsTRISUserID As String _
                , ByVal vsBRVLevel As String _
                , ByVal vsFailureLevel As String _
                , ByRef rlAbsenceNumber As Long _
                , ByRef radrsBRV As ADODB.Recordset _
                ) As Boolean
On Error GoTo eh
    'Flag
    Dim lbCheck As Boolean

    'Recordsets
    'Lock Status
    Dim ladrsLockStatus As ADODB.Recordset

    'Time
    Dim ladrsTime As ADODB.Recordset
    Dim ladrsTimeForSchedule As ADODB.Recordset
    Dim ladrsSchedule As ADODB.Recordset
    Dim ladrsAbsenceTime As ADODB.Recordset
    Dim ldtStart As Date, ldtEnd As Date, ldtExtendedStart As Date
    Dim BRVCollection As Collection
    Dim sErrorDates As String
'    Dim ldtLastProcessedDate As Date

    'Absence Start and End
    Dim ladrsAbsenceStartEnd As ADODB.Recordset
    Dim CollectionHelper As utilTRISCommonLib.cuContainer

    'Absence
    Dim ladrsInitial As ADODB.Recordset
    Dim ladrsDate As ADODB.Recordset
'    Dim ladrsProvider As ADODB.Recordset
'    Dim ladrsContact As ADODB.Recordset
    Dim ladrsReturnToWork As ADODB.Recordset
    Dim ladrsDatesAlreadyPresent As ADODB.Recordset

    Dim lcolDates As Collection
    'Components
    Dim lobjdbEAM As dbEAM.cdEAM

    PreSaveValidate2 = False
    Set radrsBRV = CreateNewBusinessErrorSet
    Set CollectionHelper = New utilTRISCommonLib.cuContainer
            
'//TODO: *** for testing only ***
If RecordsExist(vadrsInitialAbsence) Then
If CanLogBrv(mstrCBUSINESSINF, vsBRVLevel) Then
AppendBusinessError radrsBRV, trErrbusEAM, mstrCBUSINESSINF, StringFormat("Absence start time ({0}), end time ({1}), return to work ({2}) (listed as charge date) (start of validation, from client).", _
            FormatShortDateAndTime(vadrsInitialAbsence!ABS_ST_DT, vadrsInitialAbsence!ABS_ST_TM), _
            FormatShortDateAndTime(vadrsInitialAbsence!ABS_END_DT, vadrsInitialAbsence!ABS_END_TM), _
            FormatShortDateAndTime(vadrsInitialAbsence!RTN_TO_WRK_DT, vadrsInitialAbsence!RTN_TO_WRK_TM)), MODULENAME & ": PreSaveValidate2"
End If
End If

    If Not (RecordsExist(vadrsAbsenceDates) Or _
            RecordsExist(vadrsInitialAbsence) Or _
            RecordsExist(vadrsAbsenceProvider) Or _
            RecordsExist(vadrsAbsenceContact) Or _
            RecordsExist(vadrsAbsenceReturnToWork) Or _
            RecordsExist(vadrsScheduledExceptions)) Then
        If CanLogBrv(mstrCBUSINESSERR, vsBRVLevel) Then
            AppendBusinessError radrsBRV, trErrbusEAM, mstrCBUSINESSERR, "There are no absence entries to process.", MODULENAME & ": PreSaveValidate2"
        End If
        GoTo REPORTBRV
    End If
    
    If Not RecordsExist(vadrsAbsenceDates) Then
        If CanLogBrv(mstrCBUSINESSWAR, vsBRVLevel) Then
            AppendBusinessError radrsBRV, trErrbusEAM, mstrCBUSINESSWAR, "There are no absence date entries to process.", MODULENAME & ": PreSaveValidate2"
        End If
    End If
    If Not RecordsExist(vadrsInitialAbsence) Then
        If CanLogBrv(mstrCBUSINESSWAR, vsBRVLevel) Then
            AppendBusinessError radrsBRV, trErrbusEAM, mstrCBUSINESSWAR, "There is no absence information to process.", MODULENAME & ": PreSaveValidate2"
        End If
    End If
    If Not RecordsExist(vadrsScheduledExceptions) Then
        If CanLogBrv(mstrCBUSINESSWAR, vsBRVLevel) Then
            AppendBusinessError radrsBRV, trErrbusEAM, mstrCBUSINESSWAR, "There are no scheduled exception entries to process.", MODULENAME & ": PreSaveValidate2"
        End If
    End If
    
    If GetDistinctCount(vadrsAbsenceDates, "ABS_DATE") <> vadrsAbsenceDates.RecordCount Then
        Dim sDuplicateDates As String
        
        sDuplicateDates = GetDuplicateValues(vadrsAbsenceDates, "ABS_DATE")
        If CanLogBrv(mstrCBUSINESSWAR, vsBRVLevel) Then
            AppendBusinessError radrsBRV, trErrbusEAM, mstrCBUSINESSWAR, "There are duplicate absence date entries to process (" & sDuplicateDates & ").", MODULENAME & ": PreSaveValidate2"
        End If
    End If
    
    lbCheck = False

    If RecordsExist(vadrsInitialAbsence) Then
        With vadrsInitialAbsence
            .MoveFirst

            If (!RECORD_STATUS = "I") Then
                lbCheck = True
            End If
        End With
    End If

    If RecordsExist(vadrsAbsenceDates) Then
        With vadrsAbsenceDates
            .MoveFirst

            .Find "RECORD_STATUS = 'I'"
            If Not (.EOF) Then
                lbCheck = True
            End If

            If Not (lbCheck) Then
                .MoveFirst

                .Find "RECORD_STATUS = 'D'"
                If Not (.EOF) Then
                    lbCheck = True
                End If
            End If

            .MoveFirst
        End With

        'Simple Check for Valid Dates
        If Not ValidateDates(vadrsAbsenceDates, vadrsEmployeeInformation, radrsBRV) Then
            GoTo REPORTBRV
        End If
    End If
    
    ldtStart = vbNull
    ldtEnd = vbNull
        
    If RecordsExist(vadrsAbsenceDates) And RecordsExist(vadrsInitialAbsence) Then
        vadrsInitialAbsence.MoveFirst
        ldtStart = vadrsInitialAbsence!ABS_ST_DT
    
        With vadrsAbsenceDates
            .Sort = "ABS_DATE"

            .MoveFirst
            If (ldtStart > !ABS_DATE) Then
                ldtStart = !ABS_DATE
            End If
            
            .MoveLast
            ldtEnd = !ABS_DATE

            .MoveFirst
        End With

        Set ladrsLockStatus = GetLockStatus(vlEmployeeNumber, ldtStart, ldtEnd)

'        ldtStart = vbNull
'        ldtEnd = vbNull

        '//Set ladrsSchedule = GetEmployeeSchedule(vlEmployeeNumber, vadrsInitialAbsence, vadrsAbsenceDates)
        Set ladrsSchedule = GetEmployeeSchedule(vlEmployeeNumber, vadrsInitialAbsence, vadrsAbsenceDates)
        
        '//if absence is unextended, sched exc from db may prevent the probable RTW date from being determined properly
        '//try removing deleted exceptions from the schedule, this will allow the RTW to be sooner
        ApplyPendingChanges vadrsScheduledExceptions, ladrsSchedule, vadrsInitialAbsence, vadrsAbsenceDates
    End If
       
    '//TODO: Refactor and use when loading an absence, otherwise garbage in garbage out (GIGO)
    '//TODO: Borrow from tnvlib.Utility
    
    '//TODO: Move this code after DetermineAbsenceStartAndEnd
    '//compare schedule exception requests with schedule
    '//absence may have been create ealier and now extended, so bump up the start date
    If RecordsExist(vadrsScheduledExceptions) Then
        vadrsScheduledExceptions.Sort = "EMP_NUM, CHARGE_DATE, START_TIME"
        ldtExtendedStart = vadrsScheduledExceptions!CHARGE_DATE
    End If
    
    If ldtExtendedStart <> vbNull And ldtEnd <> vbNull Then
        Dim rsScheduleForValidation As ADODB.Recordset
        If RecordsExist(ladrsSchedule) Then
            ladrsSchedule.Filter = StringFormat("HOURS > 0 AND CHARGE_DATE >= #{0}# AND CHARGE_DATE <= #{1}#", ldtExtendedStart, ldtEnd)
            Set rsScheduleForValidation = CreateScheduleForValidation(ladrsSchedule)
            ladrsSchedule.Filter = adFilterNone
        End If
        
        If RecordsExist(vadrsScheduledExceptions) Then
            Dim rsSchedExcSplitForValidation As ADODB.Recordset
            
            Set rsSchedExcSplitForValidation = SplitSchedExcOverMidnight(vadrsScheduledExceptions)
                Dim oUtil As utilADOHelper.cuUtility
                Dim sKeyColumns(3) As String
                Dim sIgnoreColumns(1) As String
                Dim rsDiffLeft As ADODB.Recordset, rsDiffRight As ADODB.Recordset, rsDiffSummary As ADODB.Recordset
                Dim rc As Long
                
                Set oUtil = New utilADOHelper.cuUtility
                sKeyColumns(0) = "EMP_NUM"
                sKeyColumns(1) = "CHARGE_DATE"
                sKeyColumns(2) = "DATE_WORKED"
                sKeyColumns(3) = "START_TIME"
                
                sIgnoreColumns(0) = "EXCEPTION_CD"
                sIgnoreColumns(1) = "OVER_EXC_CD"
                
                rc = oUtil.rsCompare(rsScheduleForValidation, rsSchedExcSplitForValidation, sKeyColumns, sIgnoreColumns, rsDiffLeft, rsDiffRight, rsDiffSummary)
                If rc > 0 Then
                    If RecordsExist(rsDiffSummary) Then
                        '//check for a total mismatch of sched excp with schedule
                        rsDiffLeft.Filter = "DiffError <> ''"
                        rsDiffRight.Filter = "DiffError <> ''"
                        
                        '//TODO: Tweak this check
'                        If rsScheduleForValidation.RecordCount = rsDiffLeft.RecordCount And _
'                            rsSchedExcSplitForValidation.RecordCount = rsDiffRight.RecordCount Then
                            rsDiffSummary.Filter = "DiffSource = 'L' or DiffSource = 'R'"
                            Set lcolDates = GetUniqueValues(rsDiffSummary, "CHARGE_DATE")
                            If lcolDates.Count = 1 Then
                                Dim sSchedTime As String, sSchedExcTime As String
                                Dim blnScheduleOnLeft As Boolean '//left is schedule, right is sched exc
                                Dim lngScheduleDiffRecordCount As Long, lngSchedExcDiffRecordCount As Long
                                
                                rsDiffSummary.Filter = "DiffSource = 'L'"
                                lngScheduleDiffRecordCount = rsDiffSummary.RecordCount
                                rsDiffSummary.Filter = "DiffSource = 'R'"
                                lngSchedExcDiffRecordCount = rsDiffSummary.RecordCount
                                blnScheduleOnLeft = (lngScheduleDiffRecordCount > lngSchedExcDiffRecordCount)
                                                                
                                rsDiffSummary.Filter = StringFormat("(DiffSource = '{0}' AND DiffError = 'E') OR (DiffSource = '{0}' AND DiffError = 'I')", IIf(blnScheduleOnLeft, "L", "R"))
                                If RecordsExist(rsDiffSummary) Then
                                    rsDiffSummary.MoveFirst
                                    Dim rsDiffSummaryClone As ADODB.Recordset
                                    
                                    Set rsDiffSummaryClone = rsDiffSummary.Clone
                                    '//TODO: Test with more than one mismatch
                                    Do While Not rsDiffSummary.EOF
                                        'Set lcolDateTimeRange = GetUniqueValues2(rsDiffLeft, Array("CHARGE_DATE", "START_TIME", "CHARGE_DATE", "END_TIME"))
                                        sSchedTime = StringFormat("{0} {1} - {2}", rsDiffSummary!CHARGE_DATE, rsDiffSummary!START_TIME, rsDiffSummary!END_TIME)
                               
                                        rsDiffSummaryClone.Filter = StringFormat("DiffSource = '{0}' AND DiffError = '{1}' AND DiffRow = {2}", IIf(blnScheduleOnLeft, "R", "L"), rsDiffSummary!DiffError, rsDiffSummary!DiffRow)
                                        If Not rsDiffSummaryClone.EOF Then
                                            sSchedExcTime = StringFormat("{0} {1} - {2}", rsDiffSummaryClone!CHARGE_DATE, rsDiffSummaryClone!START_TIME, rsDiffSummaryClone!END_TIME)
                                        Else
                                            sSchedExcTime = "unknown"
                                        End If
                                                                               
                                        '//for partial exceptions, bypass brv if sched exception is a subset of the schedule
                                        '//TODO: Refactor to TimePeriodsIntersect3
                                        Dim ldt_StartDate1 As Date, ldt_EndDate1 As Date
                                        Dim ldt_StartDate2 As Date, ldt_EndDate2 As Date
                                        
                                        '//scheduled exceptions
                                        If Not rsDiffSummaryClone.EOF Then
                                            DateAndTimeRangeFrom rsDiffSummaryClone!CHARGE_DATE, rsDiffSummaryClone!START_TIME, _
                                                                rsDiffSummaryClone!CHARGE_DATE, rsDiffSummaryClone!END_TIME, _
                                                                ldt_StartDate1, ldt_EndDate1
                                        Else
                                            ldt_StartDate1 = #1/1/100#
                                            ldt_EndDate1 = #1/1/100#
                                        End If
                                        
                                        '//schedule
                                        DateAndTimeRangeFrom rsDiffSummary!CHARGE_DATE, rsDiffSummary!START_TIME, _
                                                            rsDiffSummary!CHARGE_DATE, rsDiffSummary!END_TIME, _
                                                            ldt_StartDate2, ldt_EndDate2
                                        If Not blnScheduleOnLeft Then
                                            Dim sSave As String
                                            
                                            sSave = sSchedExcTime
                                            sSchedExcTime = sSchedTime
                                            sSchedTime = sSave
                                        End If
                                        
                                        If blnScheduleOnLeft Then
                                            '//Is schedule exception a subset of the schedule ?
                                            If TimePeriodsSubset(ldt_StartDate1, ldt_EndDate1, ldt_StartDate2, ldt_EndDate2) Then
                                                If ldt_StartDate1 <> ldt_StartDate2 Or ldt_EndDate1 <> ldt_EndDate2 Then
                                                    If CanLogBrv(mstrCBUSINESSINF, vsBRVLevel) Then
                                                        AppendBusinessError radrsBRV, trErrbusEAM, mstrCBUSINESSINF, StringFormat("The schedule exception entry ({0}) is a subset of the schedule ({1}).  (If this is a partial shift absence, please disregard.)", _
                                                        sSchedExcTime, sSchedTime), MODULENAME & ": PreSaveValidate2"
                                                    End If
                                                End If
                                            Else
                                                If CanLogBrv(mstrCBUSINESSWAR, vsBRVLevel) Then
                                                    If sSchedExcTime <> "unknown" Then
                                                        AppendBusinessError radrsBRV, trErrbusEAM, mstrCBUSINESSWAR, StringFormat("The schedule exception entry ({0}) do not match the schedule ({1}).", _
                                                        sSchedExcTime, sSchedTime), MODULENAME & ": PreSaveValidate2"
                                                    Else
                                                        AppendBusinessError radrsBRV, trErrbusEAM, mstrCBUSINESSWAR, StringFormat("The schedule work day ({0}) does not have a corresponding schedule exception entry.", _
                                                        sSchedTime), MODULENAME & ": PreSaveValidate2"
                                                    End If
                                                End If
                                            End If
                                        Else '//switch, so that left is now is sched exc and right (clone) is schedule
                                            '//Is schedule a subset of the schedule exception ?
                                            If TimePeriodsSubset(ldt_StartDate2, ldt_EndDate2, ldt_StartDate1, ldt_EndDate1) Then
                                                If ldt_StartDate1 <> ldt_StartDate2 Or ldt_EndDate1 <> ldt_EndDate2 Then
                                                    If CanLogBrv(mstrCBUSINESSINF, vsBRVLevel) Then
                                                        AppendBusinessError radrsBRV, trErrbusEAM, mstrCBUSINESSINF, StringFormat("The schedule exception entry ({0}) is a subset of the schedule ({1}).  (If this is a partial shift absence, please disregard.)", _
                                                        sSchedExcTime, sSchedTime), MODULENAME & ": PreSaveValidate2"
                                                    End If
                                                End If
                                            Else
                                                If CanLogBrv(mstrCBUSINESSWAR, vsBRVLevel) Then
                                                    If sSchedExcTime <> "unknown" Then
                                                        AppendBusinessError radrsBRV, trErrbusEAM, mstrCBUSINESSWAR, StringFormat("The schedule ({0}) does not match the schedule exception entry ({1}).", _
                                                        sSchedTime, sSchedExcTime), MODULENAME & ": PreSaveValidate2"
                                                    Else
                                                        AppendBusinessError radrsBRV, trErrbusEAM, mstrCBUSINESSWAR, StringFormat("The schedule exception entry ({0}) does not have a corresponding schedule work day.", _
                                                        sSchedExcTime), MODULENAME & ": PreSaveValidate2"
                                                    End If
                                                End If
                                            End If
                                        End If
                                        
'                                        AppendBusinessError radrsBRV, trErrbusEAM, mstrCBUSINESSWAR, StringFormat("The schedule exception entries ({0}) do not match the schedule ({1}) for the following date: ({2}).", _
'                                            sSchedExcTime, sSchedTime), MODULENAME & ": PreSaveValidate2"
                                        rsDiffSummary.MoveNext
                                    Loop
                                End If
                            Else
                                Set lcolDates = CollectionAsDateRange(lcolDates)
                                sErrorDates = CollectionAsDelimitedString(lcolDates)
                                
                                If CanLogBrv(mstrCBUSINESSWAR, vsBRVLevel) Then
                                    AppendBusinessError radrsBRV, trErrbusEAM, mstrCBUSINESSWAR, "The schedule exception entries do not match the schedule for the following dates: (" & sErrorDates & ").", MODULENAME & ": PreSaveValidate2"
                                End If
                            End If
'                        Else
'                            rsDiffLeft.Filter = adFilterNone
'                            rsDiffRight.Filter = adFilterNone
'
'                            rsDiffSummary.MoveFirst
'                            rsDiffSummary.Filter = "DiffSource = 'L'"
'                            If RecordsExist(rsDiffSummary) Then
'                                '//sErrorDates = GetUniqueValues(rsDiffSummary, "CHARGE_DATE")
'                                Set lcolDates = GetUniqueValues(rsDiffSummary, "CHARGE_DATE")
'                                Set lcolDates = CollectionAsDateRange(lcolDates)
'                                sErrorDates = CollectionAsDelimitedString(lcolDates)
'
'                                If CanLogBrv(mstrCBUSINESSWAR, vsBRVLevel) Then
'                                    AppendBusinessError radrsBRV, trErrbusEAM, mstrCBUSINESSWAR, "There are schedule work days without corresponding schedule exception entries for the following dates: (" & sErrorDates & ").", MODULENAME & ": PreSaveValidate2"
'                                End If
'                            End If
'                            rsDiffSummary.Filter = adFilterNone
'
'                            rsDiffSummary.MoveFirst
'                            rsDiffSummary.Filter = "DiffSource = 'R'"
'                            If RecordsExist(rsDiffSummary) Then
'                                '//sErrorDates = GetUniqueValues(rsDiffSummary, "CHARGE_DATE")
'                                Set lcolDates = GetUniqueValues(rsDiffSummary, "CHARGE_DATE")
'                                Set lcolDates = CollectionAsDateRange(lcolDates)
'                                sErrorDates = CollectionAsDelimitedString(lcolDates)
'
'                                If CanLogBrv(mstrCBUSINESSWAR, vsBRVLevel) Then
'                                    AppendBusinessError radrsBRV, trErrbusEAM, mstrCBUSINESSWAR, "There are schedule exception entries without a corresponding schedule for the following dates: (" & sErrorDates & ").", MODULENAME & ": PreSaveValidate2"
'                                End If
'                            End If
'                            rsDiffSummary.Filter = adFilterNone
'                        End If
                    End If
                'Else
                    'AppendBusinessError radrsBRV, trErrbusEAM, mstrCBUSINESSWAR, "Unable to compare schedule exceptions with schedule.", MODULENAME & ": PreSaveValidate2"
                End If
            
                Set oUtil = Nothing
                
'            If RecordsExist(rsSchedExcSplit) Then
'                rsSchedExcSplit.Sort "EMP_NUM, DATE_WORKED, START_TIME"
'                rsSchedExcSplit.MoveFirst
'                ladrsSchedule.Sort "EMP_NUM, DATE_WORKED, START_TIME"
'                ladrsSchedule.MoveFirst
'
'                Do While Not rsSchedExcSplit.EOF
'                    '//not ideal, but...
'                    ladrsSchedule.Filter StringFormat("EMP_NUM = {0} AND DATE_WORKED = #{1}# AND START_TIME={2}", _
'                                        rsSchedExcSplit.Fields("EMP_NUM").Value, rsSchedExcSplit.Fields("DATE_WORKED").Value, rsSchedExcSplit.Fields("START_TIME").Value)
'                    If ladrsSchedule.EOF Then
'
'                    End If
'                    rsSchedExcSplit.MoveNext
'                Loop
'            End If
            
'            vadrsScheduledExceptions
            Set rsSchedExcSplitForValidation = Nothing
        End If
    End If
    
    '//compare schedule with schedule exception requests (gaps, no sched exception for abs)
    
    '//compare schedule exception requests with dates
    Dim blnDateError As Boolean
    
    '//assumption: any inserts with dates should also have inserts of scheduled exceptions
    Set BRVCollection = New Collection
    
    vadrsAbsenceDates.Filter = "RECORD_STATUS = 'I'"
    vadrsScheduledExceptions.Filter = "REC_STATUS = 'I'"
    If RecordsExist(vadrsAbsenceDates) Then
        vadrsAbsenceDates.MoveFirst
        Do While Not vadrsAbsenceDates.EOF
            blnDateError = True
            If RecordsExist(vadrsScheduledExceptions) Then
                vadrsScheduledExceptions.MoveFirst
                vadrsScheduledExceptions.Find StringFormat("CHARGE_DATE = #{0}#", vadrsAbsenceDates.Fields("ABS_DATE").Value)
                blnDateError = vadrsScheduledExceptions.EOF
            End If
        
            If blnDateError Then
                BRVCollection.Add vadrsAbsenceDates.Fields("ABS_DATE").Value
            End If
            vadrsAbsenceDates.MoveNext
        Loop
    End If
    
    If BRVCollection.Count > 0 Then
        If CanLogBrv(mstrCBUSINESSWAR, vsBRVLevel) Then
            sErrorDates = CollectionHelper.AsDelimitedString(BRVCollection, False, ", ", "")
            AppendBusinessError radrsBRV, trErrbusEAM, mstrCBUSINESSWAR, "There are absence date entries without corresponding schedule exception entries for the following dates: (" & sErrorDates & ").", MODULENAME & ": PreSaveValidate2"
        End If
    End If

    '//assumption: any deletes with dates should also have deletes of scheduled exceptions
    Set BRVCollection = New Collection
    
    vadrsAbsenceDates.Filter = "RECORD_STATUS = 'D'"
    vadrsScheduledExceptions.Filter = "REC_STATUS = 'D'"
    If RecordsExist(vadrsAbsenceDates) Then
        vadrsAbsenceDates.MoveFirst
        Do While Not vadrsAbsenceDates.EOF
            blnDateError = True
            If RecordsExist(vadrsScheduledExceptions) Then
                vadrsScheduledExceptions.MoveFirst
                vadrsScheduledExceptions.Find StringFormat("CHARGE_DATE = #{0}#", vadrsAbsenceDates.Fields("ABS_DATE").Value)
                blnDateError = vadrsScheduledExceptions.EOF
            End If
        
            If blnDateError Then
                BRVCollection.Add vadrsAbsenceDates.Fields("ABS_DATE").Value
            End If
            vadrsAbsenceDates.MoveNext
        Loop
    End If
    
    vadrsAbsenceDates.Filter = adFilterNone
    vadrsScheduledExceptions.Filter = adFilterNone
    
    If BRVCollection.Count > 0 Then
        If CanLogBrv(mstrCBUSINESSWAR, vsBRVLevel) Then
            sErrorDates = CollectionHelper.AsDelimitedString(BRVCollection, False, ", ", "")
            '//TODO: Try this for formating 07/14/2012, 07/16/2012, 07/17/2012 to 7/14, 7/16 - 7/17
'            Set lcolDates = CollectionAsDateRange(lcolDates)
'            sErrorDates = CollectionAsDelimitedString(lcolDates)
            AppendBusinessError radrsBRV, trErrbusEAM, mstrCBUSINESSWAR, "There are deleted absence date entries without corresponding deleted schedule exception entries for the following dates: (" & sErrorDates & ").", MODULENAME & ": PreSaveValidate2"
        End If
    End If

    '//check for gaps or overlaps in schedule exception requests with dates
    
    If lbCheck Then
        If RecordsExist(vadrsAbsenceDates) Then
            Dim ldtAbsStartDateTimeOrig As Date, ldtAbsEndDateTimeOrig As Date, ldtAbsRTWDateTimeOrig As Date
            
            Set ladrsDatesAlreadyPresent = Nothing
    
            If rlAbsenceNumber > -1 Then
                Set lobjdbEAM = CreateObject("dbEAM.cdEAM")
                Set ladrsDatesAlreadyPresent = lobjdbEAM.GetAbsenceDatesByAbsence(rlAbsenceNumber)
                Set lobjdbEAM = Nothing
            End If
    
            'NOTE: Schedule should have info atleast 30 days over the Absence End Date
            '//Set ladrsTime = GetEmployeeTimeDetails(vlEmployeeNumber, vadrsInitialAbsence, vadrsAbsenceDates)
            Set ladrsTime = GetEmployeeTimeDetails(vlEmployeeNumber, vadrsInitialAbsence, vadrsAbsenceDates, True)
'            ApplyPendingChanges vadrsScheduledExceptions, ladrsTime, vadrsInitialAbsence, vadrsAbsenceDates
            
''            '//TODO: Does FindDateAndTimeWorked and UpdateAbsenceWithDateWorked need more time (ladrsTimeForSchedule)
'            ldtAbsStartDateTimeOrig = FindDateAndTimeWorked(vadrsInitialAbsence!ABS_ST_DT, vadrsInitialAbsence!ABS_ST_TM, ladrsTime)
''            '//If ldtAbsEndDateTimeOrig <> #1/1/100# Then
'                ldtAbsEndDateTimeOrig = FindDateAndTimeWorked(vadrsInitialAbsence!ABS_END_DT, vadrsInitialAbsence!ABS_END_TM, ladrsTime)
''            '//End If
'            ldtAbsRTWDateTimeOrig = FindDateAndTimeWorked(vadrsInitialAbsence!RTN_TO_WRK_DT, vadrsInitialAbsence!RTN_TO_WRK_TM, ladrsTime)
            
            '//TEST:
            If IsNull(ldtAbsEndDateTimeOrig) Or ldtAbsEndDateTimeOrig = #1/1/100# Then
                If CanLogBrv(mstrCBUSINESSWAR, vsBRVLevel) Then
                    AppendBusinessError radrsBRV, trErrbusEAM, mstrCBUSINESSWAR, StringFormat("The absence end time has not been specified.  (The last absence time (non-working) before the return to work time ({0}) will be used.)", _
                        FormatShortDateTime(ldtAbsRTWDateTimeOrig)), MODULENAME & ": PreSaveValidate2"
                End If
            End If
              
            Set vadrsInitialAbsence = UpdateAbsenceWithDateWorked(vadrsInitialAbsence, ladrsTime)
            
            ldtAbsStartDateTimeOrig = DateTimeFrom(vadrsInitialAbsence!ST_DATE_WORKED, vadrsInitialAbsence!ABS_ST_TM)
'            If ldtAbsEndDateTimeOrig <> #1/1/100# Then
            '//since this is already date worked, (don't automatically adjust to the next date for midnight, since it has already been done)
                ldtAbsEndDateTimeOrig = DateTimeFrom2(vadrsInitialAbsence!END_DATE_WORKED, vadrsInitialAbsence!ABS_END_TM)
'            End If
            ldtAbsRTWDateTimeOrig = DateTimeFrom(vadrsInitialAbsence!RTW_DATE_WORKED, vadrsInitialAbsence!RTN_TO_WRK_TM)
            
            '//TEST:
            '//start of datetime checking: abs_start < abs_end <= rtw
            vadrsInitialAbsence.MoveFirst
            If CanLogBrv(mstrCBUSINESSWAR, vsBRVLevel) Then
'            If vadrsInitialAbsence!ABS_END_DT <> #1/1/100# Then
                If vadrsInitialAbsence!ABS_END_DT <> #1/1/100# Then
                    If DateTimeFrom2(vadrsInitialAbsence!ABS_ST_DT, vadrsInitialAbsence!ABS_ST_TM) > DateTimeFrom2(vadrsInitialAbsence!ABS_END_DT, vadrsInitialAbsence!ABS_END_TM) Then
                        AppendBusinessError radrsBRV, trErrbusEAM, mstrCBUSINESSWAR, StringFormat("The absence start time ({0}) is greater than the absence end time ({1}).", _
                                    FormatShortDateAndTime(vadrsInitialAbsence!ABS_ST_DT, vadrsInitialAbsence!ABS_ST_TM), _
                                    FormatShortDateAndTime(vadrsInitialAbsence!ABS_END_DT, vadrsInitialAbsence!ABS_END_TM)), MODULENAME & ": PreSaveValidate2"
                    End If
                End If
            
                If DateTimeFrom2(vadrsInitialAbsence!ABS_ST_DT, vadrsInitialAbsence!ABS_ST_TM) = DateTimeFrom2(vadrsInitialAbsence!ABS_END_DT, vadrsInitialAbsence!ABS_END_TM) Then
                    AppendBusinessError radrsBRV, trErrbusEAM, mstrCBUSINESSWAR, StringFormat("The absence start time is equal to the absence end time ({0}).", _
                                FormatShortDateAndTime(vadrsInitialAbsence!ABS_ST_DT, vadrsInitialAbsence!ABS_ST_TM)), MODULENAME & ": PreSaveValidate2"
                End If
'            End If
                                     
                If DateTimeFrom2(vadrsInitialAbsence!RTN_TO_WRK_DT, vadrsInitialAbsence!RTN_TO_WRK_TM) < DateTimeFrom2(vadrsInitialAbsence!ABS_ST_DT, vadrsInitialAbsence!ABS_ST_TM) Then
                    AppendBusinessError radrsBRV, trErrbusEAM, mstrCBUSINESSWAR, StringFormat("The return to work time ({0}) is before the absence start time ({1}).", _
                                FormatShortDateAndTime(vadrsInitialAbsence!RTN_TO_WRK_DT, vadrsInitialAbsence!RTN_TO_WRK_TM), _
                                FormatShortDateAndTime(vadrsInitialAbsence!ABS_ST_DT, vadrsInitialAbsence!ABS_ST_TM)), MODULENAME & ": PreSaveValidate2"
                End If
                If DateTimeFrom2(vadrsInitialAbsence!RTN_TO_WRK_DT, vadrsInitialAbsence!RTN_TO_WRK_TM) = DateTimeFrom2(vadrsInitialAbsence!ABS_ST_DT, vadrsInitialAbsence!ABS_ST_TM) Then
                    AppendBusinessError radrsBRV, trErrbusEAM, mstrCBUSINESSWAR, StringFormat("The return to work time ({0}) is equal to the absence start time ({1}).", _
                                FormatShortDateAndTime(vadrsInitialAbsence!RTN_TO_WRK_DT, vadrsInitialAbsence!RTN_TO_WRK_TM), _
                                FormatShortDateAndTime(vadrsInitialAbsence!ABS_ST_DT, vadrsInitialAbsence!ABS_ST_TM)), MODULENAME & ": PreSaveValidate2"
                End If
    
    '            If vadrsInitialAbsence!ABS_END_DT <> #1/1/100# Then
                    If DateTimeFrom2(vadrsInitialAbsence!RTN_TO_WRK_DT, vadrsInitialAbsence!RTN_TO_WRK_TM) < DateTimeFrom2(vadrsInitialAbsence!ABS_END_DT, vadrsInitialAbsence!ABS_END_TM) Then
                        AppendBusinessError radrsBRV, trErrbusEAM, mstrCBUSINESSWAR, StringFormat("The return to work time ({0}) is before the absence end time ({1}).", _
                                    FormatShortDateAndTime(vadrsInitialAbsence!RTN_TO_WRK_DT, vadrsInitialAbsence!RTN_TO_WRK_TM), _
                                    FormatShortDateAndTime(vadrsInitialAbsence!ABS_END_DT, vadrsInitialAbsence!ABS_END_TM)), MODULENAME & ": PreSaveValidate2"
                    End If
'            End If
            End If
            
            '//TEST:
            'Determine if non consecutive dates were selected
            If RecordsExist(ladrsSchedule) Or RecordsExist(ladrsTime) Then
                '//more datetime checking: abs end = rtw, and rtw is a start time
                If vadrsInitialAbsence!ABS_END_DT <> #1/1/100# Then
                    If DateTimeFrom(vadrsInitialAbsence!ABS_END_DT, vadrsInitialAbsence!ABS_END_TM) = DateTimeFrom(vadrsInitialAbsence!RTN_TO_WRK_DT, vadrsInitialAbsence!RTN_TO_WRK_TM) Then
                        Dim lblnRTWShiftFound As Boolean
                        
                        lblnRTWShiftFound = False
                        If RecordsExist(ladrsTime) Then
                            ladrsTime.Filter = StringFormat("(TIME_TYPE = 2 AND CHARGE_DATE = #{0}# AND START_TIME = '{1}')" & _
                                            " OR (TIME_TYPE = 3 AND CHARGE_DATE = #{0}# AND START_TIME = '{1}')", _
                                            vadrsInitialAbsence!ABS_ST_DT, vadrsInitialAbsence!ABS_ST_TM)
                            If Not ladrsTime.EOF Then
                                lblnRTWShiftFound = True
                            End If
                            ladrsTime.Filter = adFilterNone
                        End If
                        
                        If Not lblnRTWShiftFound Then
                            If RecordsExist(ladrsSchedule) Then
                                ladrsSchedule.Filter = StringFormat("CHARGE_DATE = #{0}# AND START_TIME = '{1}'", _
                                                vadrsInitialAbsence!ABS_ST_DT, vadrsInitialAbsence!ABS_ST_TM)
                                If Not ladrsSchedule.EOF Then
                                    lblnRTWShiftFound = True
                                End If
                                ladrsSchedule.Filter = adFilterNone
                            End If
                        End If
                        
                        If Not lblnRTWShiftFound Then
                            If CanLogBrv(mstrCBUSINESSWAR, vsBRVLevel) Then
                                AppendBusinessError radrsBRV, trErrbusEAM, mstrCBUSINESSWAR, StringFormat("Unable to determine return to work shift on schedule ({0}).", _
                                        FormatShortDateAndTime(vadrsInitialAbsence!RTN_TO_WRK_DT, vadrsInitialAbsence!RTN_TO_WRK_TM)), MODULENAME & ": PreSaveValidate2"
                            End If
                        End If
                    End If
                End If
                            
                '//If RecordsExist(ladrsAbsenceTime) Then '//fails on .MoveLast when no records
                    Set ladrsTimeForSchedule = MergeScheduleWithTime(vlEmployeeNumber, vadrsInitialAbsence, vadrsAbsenceDates, vadrsScheduledExceptions)
                
                    Set ladrsAbsenceTime = AddTimeAndSchExceptions(vadrsInitialAbsence _
                                                    , vadrsScheduledExceptions _
                                                    , ladrsSchedule _
                                                    , ladrsTimeForSchedule _
                                                    , ladrsLockStatus)

                    '//pre-save is using TimeForSchedule, save is using Schedule
                    Set ladrsAbsenceStartEnd = DetermineAbsenceStartAndEnd(ladrsAbsenceTime _
                                                    , ladrsTimeForSchedule _
                                                    , vadrsExceptions _
                                                    , ladrsLockStatus _
                                                    , ladrsTime)
                '//End If
                
                If (RecordsExist(ladrsAbsenceStartEnd)) Then
                    Dim ldtAbsStartDateTimeNew As Date, ldtAbsEndDateTimeNew As Date, ldtAbsRTWDateTimeNew As Date
                    ladrsAbsenceStartEnd.MoveFirst
                    
                    ldtAbsStartDateTimeNew = DateTimeFrom(ladrsAbsenceStartEnd!ST_DATE_WORKED, ladrsAbsenceStartEnd!ABS_ST_TM)
                    ldtAbsEndDateTimeNew = DateTimeFrom(ladrsAbsenceStartEnd!END_DATE_WORKED, ladrsAbsenceStartEnd!ABS_END_TM)
                    ldtAbsRTWDateTimeNew = DateTimeFrom(ladrsAbsenceStartEnd!RTW_DATE_WORKED, ladrsAbsenceStartEnd!RTN_TO_WRK_TM)
                    
'                    If ldtAbsStartDateTimeNew <> ldtAbsStartDateTimeOrig Then
'                        vadrsInitialAbsence!ABS_ST_DT = MidnightDate(ldtAbsStartDateTimeNew)
''                        vadrsInitialAbsence!ST_DATE_WORKED = MidnightDate(ladrsAbsenceStartEnd!ST_DATE_WORKED)
'                        vadrsInitialAbsence!ABS_ST_TM = FormatTime(ldtAbsStartDateTimeNew)
'                    End If
'
'                    If ldtAbsEndDateTimeNew <> ldtAbsEndDateTimeOrig Then
'                        vadrsInitialAbsence!ABS_END_DT = MidnightDate(ldtAbsEndDateTimeNew)
'                        vadrsInitialAbsence!ABS_END_TM = FormatTime(ldtAbsEndDateTimeNew)
'                    End If
'
'                    If ldtAbsRTWDateTimeNew <> ldtAbsRTWDateTimeOrig Then
'                        If CanLogBrv(mstrCBUSINESSINF, vsBRVLevel) Then
'                            AppendBusinessError radrsBRV, trErrbusEAM, mstrCBUSINESSINF, StringFormat("The probable return to work date is now ({0}).  The original return to work date was ({1})", _
'                                FormatShortDateTime(ldtAbsRTWDateTimeNew), FormatShortDateTime(ldtAbsRTWDateTimeOrig)), MODULENAME & ": PreSaveValidate2"
'                        End If
'
'                        vadrsInitialAbsence!RTN_TO_WRK_DT = MidnightDate(ldtAbsRTWDateTimeNew)
'                        vadrsInitialAbsence!RTN_TO_WRK_TM = FormatTime(ldtAbsRTWDateTimeNew)
'                    Else
''                        If CanLogBrv(mstrCBUSINESSINF, vsBRVLevel) Then
''                            AppendBusinessError radrsBRV, trErrbusEAM, mstrCBUSINESSINF, StringFormat("The probable return to work date is ({0})", _
''                                FormatShortDateTime(ldtAbsRTWDateTimeNew)), MODULENAME & ": PreSaveValidate2"
''                        End If
'                    End If
                    
                    '//TEST:
                    'Should SchedExceptions start and end times be between absence start and end times
                    vadrsScheduledExceptions.Filter = "REC_STATUS <> 'D'"
                    If RecordsExist(vadrsScheduledExceptions) Then
                        Dim ldtAbsStartDateTime As Date, ldtAbsEndDateTime As Date
                        Dim ldtSchedExcStartDateTime As Date, ldtSchedExcEndDateTime As Date
                        Dim blnTimePeriodsSubsetValid As Boolean
                        
'                        DateAndTimeRangeFrom vadrsInitialAbsence!ABS_ST_DT, vadrsInitialAbsence!ABS_ST_TM, _
'                                            vadrsInitialAbsence!ABS_END_DT, vadrsInitialAbsence!ABS_END_TM, _
'                                            ldtAbsStartDateTime, ldtAbsEndDateTime
'
'                        ldtAbsStartDateTime = FindDateTimeWorked(ldtAbsStartDateTime, ladrsTime)
                        
                        If ldtAbsEndDateTime <> #1/1/100# Then
                            vadrsScheduledExceptions.MoveFirst
                            Do While Not vadrsScheduledExceptions.EOF
                                DateAndTimeRangeFrom vadrsScheduledExceptions!CHARGE_DATE, vadrsScheduledExceptions!START_TIME, _
                                                    vadrsScheduledExceptions!CHARGE_DATE, vadrsScheduledExceptions!END_TIME, _
                                                    ldtSchedExcStartDateTime, ldtSchedExcEndDateTime, False
                                ldtSchedExcStartDateTime = FindDateTimeWorked(ldtSchedExcStartDateTime, ladrsTime)
                                ldtSchedExcEndDateTime = FindDateTimeWorked(ldtSchedExcEndDateTime, ladrsTime)
                                
                                '//there may be multiple absences (because of a split) so validate against any of them
                                blnTimePeriodsSubsetValid = False
                                ladrsAbsenceStartEnd.MoveFirst
                                Do While Not ladrsAbsenceStartEnd.EOF
                                    DateAndTimeRangeFrom ladrsAbsenceStartEnd!ST_DATE_WORKED, ladrsAbsenceStartEnd!ABS_ST_TM, _
                                                        ladrsAbsenceStartEnd!ABS_END_DT, ladrsAbsenceStartEnd!ABS_END_TM, _
                                                        ldtAbsStartDateTime, ldtAbsEndDateTime

'                                    ldtAbsStartDateTime = FindDateTimeWorked(ldtAbsStartDateTime, ladrsTime)
                                
                                    If TimePeriodsSubset(ldtSchedExcStartDateTime, ldtSchedExcEndDateTime, ldtAbsStartDateTime, ldtAbsEndDateTime) Then
                                        blnTimePeriodsSubsetValid = True
                                        Exit Do
                                    End If
                                    ladrsAbsenceStartEnd.MoveNext
                                Loop
                                
                                '//TODO: Create range of dates for multiple absences
                                If Not blnTimePeriodsSubsetValid Then
                                    If CanLogBrv(mstrCBUSINESSWAR, vsBRVLevel) Then
                                        AppendBusinessError radrsBRV, trErrbusEAM, mstrCBUSINESSWAR, StringFormat("Schedule exception entry ({0}) is not within range of absence ({1}).", _
                                                FormatShortDateTimeRange(ldtSchedExcStartDateTime, ldtSchedExcEndDateTime), _
                                                FormatShortDateTimeRange(ldtAbsStartDateTime, ldtAbsEndDateTime)), MODULENAME & ": PreSaveValidate2"
                                    End If
                                End If
                                vadrsScheduledExceptions.MoveNext
                            Loop
                        End If
                    End If
                    vadrsScheduledExceptions.Filter = adFilterNone
                                        
                    'Create Possible Multiple
                    CreateAndAddAllAbsenceRecords vlEmployeeNumber _
                        , vadrsInitialAbsence _
                        , vadrsAbsenceDates _
                        , vadrsAbsenceReturnToWork _
                        , ladrsDatesAlreadyPresent _
                        , ladrsSchedule _
                        , ladrsAbsenceStartEnd _
                        , ladrsInitial _
                        , ladrsDate _
                        , ladrsReturnToWork
    
                    Set ladrsDatesAlreadyPresent = Nothing
                Else
                    If CanLogBrv(mstrCBUSINESSWAR, vsBRVLevel) Then
                        AppendBusinessError radrsBRV, trErrbusEAM, mstrCBUSINESSWAR, StringFormat("Unable to determine the probable return to work date.  The original return to work date is ({0})", _
                            FormatShortDateTime(ldtAbsRTWDateTimeOrig)), MODULENAME & ": PreSaveValidate2"
                    End If
                End If
                
                'Validate
                If RecordsExist(ladrsInitial) Then
                    If Not ValidateInitialAbsence(ladrsInitial, radrsBRV) Then
'                        GoTo REPORTBRV
                    Else
        '                If (ladrsInitial.RecordCount > 1) Then
        '                    Set ladrsPortfolio = AddPortfolioInformation(1 _
        '                                            , ladrsInitial)
        '                    Set ladrsPortfolioDetails = AddPortfolioInformation(2 _
        '                                                    , ladrsInitial)
        '                End If
                    End If
                    
                    vadrsInitialAbsence.MoveFirst
                    ladrsInitial.MoveFirst
                    Do While Not ladrsInitial.EOF
                        If vadrsInitialAbsence.Fields("ABS_NUM").Value > 0 And ladrsInitial.Fields("ABS_NUM").Value < 0 Then '//-1, -2, -3, etc.
                            If ladrsAbsenceStartEnd.RecordCount <= 1 Then '//suppress for gaps that will be reported
                                AppendBusinessError radrsBRV, trErrbusEAM, mstrCBUSINESSERR, _
                                    "Absence will be split rather than extended, please check for any reported time or schedule exceptions during the absence period.  (Could be a partial time entry, entries crossing midnight, or a date worked issue.)" _
                                    , MODULENAME & ": PreSaveValidate2"
                            End If
                        End If
                        ladrsInitial.MoveNext
                    Loop
                End If
                
                
                
                If RecordsExist(ladrsAbsenceStartEnd) Then
                    If ladrsAbsenceStartEnd.RecordCount > 1 Then
                        Dim rsGAP As ADODB.Recordset, rsScheduleInGAP As ADODB.Recordset
                                                
                        '//TODO: Refactor to more meaningful names (CreateTimeRangeRS, CreateSplitTimeRangeRS)
                        Set rsGAP = CreateGAP() '//time range (start date, start time, end date, end time)
                        Set rsScheduleInGAP = CreateScheduleInGAP() '//time range split over midnight (charge date, dateworked, start time, end time)
    
                        'Dim sStartDate As String, sEndDate As String
                        Dim dtStartDate As Date, dtEndDate As Date
                        Dim sDateRangeList As String, sGapDateRangeList As String
                        
                        '//build list of start and end dates, there is probably a gap that is causing the abs to be split
                        sDateRangeList = ""
                        sGapDateRangeList = ""
                        ladrsAbsenceStartEnd.MoveFirst
                        Do While Not ladrsAbsenceStartEnd.EOF
                            dtStartDate = DateTimeFrom(ladrsAbsenceStartEnd.Fields("ST_DATE_WORKED").Value, ladrsAbsenceStartEnd.Fields("ABS_ST_TM").Value)
                            '//could use END_DATE_WORKED
                            dtEndDate = DateTimeFrom(ladrsAbsenceStartEnd.Fields("ABS_END_DT").Value, ladrsAbsenceStartEnd.Fields("ABS_END_TM").Value, True)
                            
                            DateTimeCrossesMidnight dtStartDate, dtEndDate
                            sDateRangeList = sDateRangeList & FormatShortDateTimeRange(dtStartDate, dtEndDate) 'FormatShortDateTime(dtStartDate) & " - " & FormatShortDateTime(dtEndDate)
                            If sGapDateRangeList <> "" Then 'skip 1st row
                                sGapDateRangeList = sGapDateRangeList & FormatShortDateTime(dtStartDate)
                                If Not rsGAP.EOF Then
                                    rsGAP.Fields("GAP_END_DATE").Value = MidnightDate(dtStartDate)
                                    rsGAP.Fields("GAP_END_TIME").Value = FormatTime(dtStartDate)
                                    rsGAP.Update
                                End If
                            End If
                            
                            ladrsAbsenceStartEnd.MoveNext
                            
                            If Not ladrsAbsenceStartEnd.EOF Then
                                sDateRangeList = sDateRangeList & ", absence dates: "
                                If sGapDateRangeList <> "" Then 'skip 1st row
                                    sGapDateRangeList = sGapDateRangeList & ", "
                                End If
                                sGapDateRangeList = sGapDateRangeList & FormatShortDateTime(dtEndDate) & " - "
                                rsGAP.AddNew
                                rsGAP.Fields("GAP_START_DATE").Value = MidnightDate(dtEndDate)
                                rsGAP.Fields("GAP_START_TIME").Value = FormatTime(dtEndDate)
                                rsGAP.Update
                            End If
                        Loop
                        
    '//for a better brv message, try to match gaps to schedule (working days)
    '//hopefully this will indicate if a working day was omitted from the absence which would probably cause the absence to be split
                        Dim blnWorkingDay As Boolean
                        '//loop through gap and filter schedule
                        If RecordsExist(rsGAP) Then
                        If CanLogBrv(mstrCBUSINESSERR, vsBRVLevel) Then
                            rsGAP.MoveFirst
                            Do While Not rsGAP.EOF
                                ladrsSchedule.Filter = adFilterNone
                                ladrsSchedule.Filter = StringFormat("SHIFT_HOURS > 0.0 AND SYSTEM_EXC_CD <> 'SLT'" & _
                                                " AND EXCEPTION_CD <> 'ONP'" & _
                                                " AND DATE_WORKED >= #{0}# " & _
                                                " AND DATE_WORKED <= #{1}#", _
                                                rsGAP.Fields("GAP_START_DATE").Value, rsGAP.Fields("GAP_END_DATE").Value)
                        
                                If Not ladrsSchedule.EOF Then
                                    ladrsSchedule.MoveFirst
                                    Do While Not ladrsSchedule.EOF
                                        '//TODO: Does this need to be specified as an end time, since it is a dateworked ???
                                        If DateTimeFrom(ladrsSchedule.Fields("DATE_WORKED").Value, ladrsSchedule.Fields("START_TIME").Value) > _
                                            DateTimeFrom(rsGAP.Fields("GAP_START_DATE").Value, rsGAP.Fields("GAP_START_TIME").Value) And _
                                            DateTimeFrom(ladrsSchedule.Fields("DATE_WORKED").Value, ladrsSchedule.Fields("END_TIME").Value, True) < _
                                            DateTimeFrom(rsGAP.Fields("GAP_END_DATE").Value, rsGAP.Fields("GAP_END_TIME").Value, True) Then
    
                                            '//determine if schedule is a working day
                                            blnWorkingDay = False
                                            If (Trim(ladrsSchedule!OVER_EXC_CD) <> "" Or _
                                                Trim(ladrsSchedule!EXCEPTION_CD) <> "") Then
                                                vadrsExceptions.Filter = adFilterNone
                
                                                vadrsExceptions.Filter = ("(EXCEPTION_CD = '" & Trim(ladrsSchedule!OVER_EXC_CD) & _
                                                                    "' AND EXCEPTION_BANK_CD <> 'O')" & _
                                                                    " OR (EXCEPTION_CD = '" & Trim(ladrsSchedule!EXCEPTION_CD) & _
                                                                    "' AND EXCEPTION_BANK_CD <> 'O')")
                                                If Not vadrsExceptions.EOF Then
                                                    vadrsExceptions.MoveFirst
                
                                                    vadrsExceptions.Find "WORK_REST_IND = 'R'"
                                                    If vadrsExceptions.EOF Then '//no rest found only work
                                                        blnWorkingDay = True
                                                    End If
                                                End If
                
                                                vadrsExceptions.Filter = adFilterNone
                                            Else
                                                blnWorkingDay = True
                                            End If
                                            
                                            If blnWorkingDay Then
    Debug.Print StringFormat("ladrsSchedule==>{0} {1} {2} {4}", ladrsSchedule.Fields("CHARGE_DATE").Value, ladrsSchedule.Fields("DATE_WORKED").Value, ladrsSchedule.Fields("START_TIME").Value, ladrsSchedule.Fields("END_TIME").Value)
                                                rsScheduleInGAP.AddNew
                                                rsScheduleInGAP.Fields("CHARGE_DATE").Value = ladrsSchedule.Fields("CHARGE_DATE").Value
                                                rsScheduleInGAP.Fields("DATE_WORKED").Value = ladrsSchedule.Fields("DATE_WORKED").Value
                                                rsScheduleInGAP.Fields("START_TIME").Value = ladrsSchedule.Fields("START_TIME").Value
                                                rsScheduleInGAP.Fields("END_TIME").Value = ladrsSchedule.Fields("END_TIME").Value
                                                rsScheduleInGAP.Update
                                            End If
                                        End If
                                        ladrsSchedule.MoveNext
                                    Loop '//Do While Not ladrsSchedule.EOF
                                End If
                                
                                rsGAP.MoveNext
                            Loop '//Do While Not rsGAP.EOF
                            
                            If RecordsExist(rsScheduleInGAP) Then
                                Dim sScheduleInGAP As String
                                '//sScheduleInGAP = GetUniqueValues(rsScheduleInGAP, "CHARGE_DATE")
                                Set lcolDates = GetUniqueValues(rsScheduleInGAP, "CHARGE_DATE")
                                Set lcolDates = CollectionAsDateRange(lcolDates)
                                sScheduleInGAP = CollectionAsDelimitedString(lcolDates)
    
                                AppendBusinessError radrsBRV, trErrbusEAM, mstrCBUSINESSERR, _
                                    "Absence will be split rather than extended, please review the following working dates that may not be included on the absence. (" & sScheduleInGAP & ")" _
                                    , MODULENAME & ": PreSaveValidate2"
                            End If
                        End If '//If CanLogBrv(mstrCBUSINESSERR, vsBRVLevel) Then
                        End If '//If RecordsExist(rsGAP) Then
                        
                        If CanLogBrv(mstrCBUSINESSERR, vsBRVLevel) Then
                            AppendBusinessError radrsBRV, trErrbusEAM, mstrCBUSINESSERR, _
                                "Absence will be split rather than extended, please review any gaps in the following date ranges. (absence dates: " _
                                & sDateRangeList & "; possible gaps:" & sGapDateRangeList & ")", MODULENAME & ": PreSaveValidate2"
                        End If
                    End If
                End If
                
            End If

            Set ladrsLockStatus = Nothing
            '//TODO: *** for testing only ***
If RecordsExist(vadrsInitialAbsence) Then
vadrsInitialAbsence.MoveFirst
If CanLogBrv(mstrCBUSINESSINF, vsBRVLevel) Then
'//convert to dateworked
'ldtAbsStartDateTimeNew = FindDateTimeWorked(ldtAbsStartDateTimeNew, ladrsTime)
'ldtAbsEndDateTimeNew = FindDateTimeWorked(ldtAbsEndDateTimeNew, ladrsTime)
'ldtAbsRTWDateTimeNew = FindDateTimeWorked(ldtAbsRTWDateTimeNew, ladrsTime)

AppendBusinessError radrsBRV, trErrbusEAM, mstrCBUSINESSINF, StringFormat("Adjusted start time ({0}){3}, end time ({1}){4}, return to work ({2}){5} (listed as dateworked) (end of validation, from server).", _
            FormatShortDateTime(ldtAbsStartDateTimeNew), _
            FormatShortDateTime(ldtAbsEndDateTimeNew), _
            FormatShortDateTime(ldtAbsRTWDateTimeNew), _
            IIf(ldtAbsStartDateTimeNew <> ldtAbsStartDateTimeOrig, "*", ""), _
            IIf(ldtAbsEndDateTimeNew <> ldtAbsEndDateTimeOrig, "*", ""), _
            IIf(ldtAbsRTWDateTimeNew <> ldtAbsRTWDateTimeOrig, "*", "")), MODULENAME & ": PreSaveValidate2"
'//If absence is going to be split, display the additional splits
If RecordsExist(ladrsAbsenceStartEnd) Then
    ladrsAbsenceStartEnd.MoveFirst
    If ladrsAbsenceStartEnd.RecordCount >= 2 Then
        ladrsAbsenceStartEnd.MoveNext
        Do While Not ladrsAbsenceStartEnd.EOF
ldtAbsStartDateTimeNew = DateTimeFrom(ladrsAbsenceStartEnd.Fields("ST_DATE_WORKED").Value, ladrsAbsenceStartEnd.Fields("ABS_ST_TM").Value)
ldtAbsEndDateTimeNew = DateTimeFrom(ladrsAbsenceStartEnd.Fields("END_DATE_WORKED").Value, ladrsAbsenceStartEnd.Fields("ABS_END_TM").Value)
ldtAbsRTWDateTimeNew = DateTimeFrom(ladrsAbsenceStartEnd.Fields("RTW_DATE_WORKED").Value, ladrsAbsenceStartEnd.Fields("RTN_TO_WRK_TM").Value)

'ldtAbsStartDateTimeNew = FindDateTimeWorked(ldtAbsStartDateTimeNew, ladrsTime)
'ldtAbsEndDateTimeNew = FindDateTimeWorked(ldtAbsEndDateTimeNew, ladrsTime)
'ldtAbsRTWDateTimeNew = FindDateTimeWorked(ldtAbsRTWDateTimeNew, ladrsTime)

AppendBusinessError radrsBRV, trErrbusEAM, mstrCBUSINESSINF, StringFormat("Additional splits start time ({0}){3}, end time ({1}){4}, return to work ({2}){5} (listed as dateworked) (end of validation, from server).", _
            FormatShortDateTime(ldtAbsStartDateTimeNew), _
            FormatShortDateTime(ldtAbsEndDateTimeNew), _
            FormatShortDateTime(ldtAbsRTWDateTimeNew), _
            IIf(ldtAbsStartDateTimeNew <> ldtAbsStartDateTimeOrig, "*", ""), _
            IIf(ldtAbsEndDateTimeNew <> ldtAbsEndDateTimeOrig, "*", ""), _
            IIf(ldtAbsRTWDateTimeNew <> ldtAbsRTWDateTimeOrig, "*", "")), MODULENAME & ": PreSaveValidate2"
        
            ladrsAbsenceStartEnd.MoveNext
        Loop
    End If
End If

End If
End If
        End If '//RecordsExist(vadrsAbsenceDates)
    End If '//lbCheck

REPORTBRV:

    radrsBRV.Filter = adFilterNone
    If RecordsExist(radrsBRV) Then
        radrsBRV.MoveFirst

        'Check for errors
        If vsFailureLevel = "None" Then
            PreSaveValidate2 = True
        Else
            Select Case vsFailureLevel
            Case "Information"
                radrsBRV.Filter = "ERR_TYPE = 'BI' Or ERR_TYPE = 'BW' Or ERR_TYPE = 'BE'"
            Case "Warning"
                radrsBRV.Filter = "ERR_TYPE = 'BW' Or ERR_TYPE = 'BE'"
            Case "Error"
                radrsBRV.Filter = "ERR_TYPE = 'BE'"
            End Select
            'true = success, false = fail
            PreSaveValidate2 = Not RecordsExist(radrsBRV)
        End If
        
        radrsBRV.Filter = adFilterNone
        radrsBRV.MoveFirst
    Else
        PreSaveValidate2 = True
    End If
    
    Set CollectionHelper = Nothing
    Set BRVCollection = Nothing
    Exit Function
eh:
    PreSaveValidate2 = False
    Set CollectionHelper = Nothing
    Set BRVCollection = Nothing
    
    Select Case (Err.Number)
    Case trErrutOpenConnection
        TrsRaiseError MODULENAME, "PreSaveValidate2"
    Case Else
        TrsRaiseError MODULENAME, "PreSaveValidate2", Err.Number, , , , True
    End Select

    Exit Function
End Function

'//vsBRVErrorType = "BI", "BW", or "BE"
'//vsBRVLevel = None, Information, Warning, or Error
Private Function CanLogBrv(ByVal vsBRVErrorType As String, ByVal vsBRVLevel As String) As Boolean
On Error GoTo eh

    CanLogBrv = True
    
    Select Case vsBRVLevel
    Case "None"
        CanLogBrv = False
    Case "Information"
        CanLogBrv = True
    Case "Warning"
        Select Case vsBRVErrorType
        Case mstrCBUSINESSERR '//"BE"
            CanLogBrv = True
        Case mstrCBUSINESSWAR '//"BW"
            CanLogBrv = True
        Case Else '//"BI"
            CanLogBrv = False
        End Select
    Case "Error"
        Select Case vsBRVErrorType
        Case mstrCBUSINESSERR '//"BE"
            CanLogBrv = True
        Case Else '//"BW", "BI"
            CanLogBrv = False
        End Select
    End Select
    
Exit Function
eh:
    TrsRaiseError MODULENAME, "CanLogBrv"
    Exit Function
End Function

'//if absence is unextended, sched exc from db may prevent the probable RTW date from being determined properly
'//try removing deleted exceptions from the schedule, this will allow the RTW to be sooner
'//if abs has 5/2, 5/3, 5/4 the RTW would probable be 5/5
'//if the abs is unextended by removing 5/4, the RTW should be 5/4, but there could be a sched exc in the db for 5/4 that may interfere
'//the exception for this is holidays that are not selected on the client or the change management on the client has removed them, then leave them alone here
Private Function ApplyPendingChanges(ByVal vadrsScheduledExceptions As ADODB.Recordset, ByVal ladrsSchedule As ADODB.Recordset, _
                                    ByVal vadrsInitialAbsence As ADODB.Recordset, ByVal vadrsAbsenceDates As ADODB.Recordset) As Boolean
On Error GoTo eh
    Dim lobjdbCodeTables As dbCodeTables.cdCodeTables
    
    If RecordsExist(vadrsScheduledExceptions) Then
        Set lobjdbCodeTables = New dbCodeTables.cdCodeTables
        vadrsScheduledExceptions.Filter = "REC_STATUS = 'D'"
        If RecordsExist(vadrsScheduledExceptions) Then
            vadrsScheduledExceptions.MoveFirst
            
            Do While Not vadrsScheduledExceptions.EOF
                '//TODO: Not sure which group to use, 1 and 2 exist in prod
                If Not lobjdbCodeTables.IsExceptionInGroup(1, Trim(vadrsScheduledExceptions.Fields("EXCEPTION_CD").Value), "H") Then '//Exclude holidays
    '                ladrsSchedule.Filter = StringFormat("EMP_NUM={0} AND CHARGE_DATE=#{1}# AND SHIFT_START_TIME='{2}' AND SHIFT_END_TIME='{3}'", _
    '                        vadrsScheduledExceptions!EMP_NUM, vadrsScheduledExceptions!CHARGE_DATE, vadrsScheduledExceptions!START_TIME, vadrsScheduledExceptions!END_TIME)
                    ladrsSchedule.Filter = StringFormat("EMP_NUM={0} AND CHARGE_DATE=#{1}#", vadrsScheduledExceptions!EMP_NUM, vadrsScheduledExceptions!CHARGE_DATE)
                    If RecordsExist(ladrsSchedule) Then
                        ladrsSchedule.MoveFirst
                        Do While Not ladrsSchedule.EOF
                            If Trim(ladrsSchedule!START_TIME) <> "" And Trim(ladrsSchedule!END_TIME) <> "" Then '// caused a type mismatch in DateTimeFrom->IsTimeValid
                            If DateTimeFrom(ladrsSchedule!CHARGE_DATE, ladrsSchedule!START_TIME, False) >= DateTimeFrom(vadrsScheduledExceptions!CHARGE_DATE, ConvertDB2Time(vadrsScheduledExceptions!START_TIME), False) And _
                                DateTimeFrom(ladrsSchedule!CHARGE_DATE, ladrsSchedule!END_TIME, True) <= DateTimeFrom(vadrsScheduledExceptions!CHARGE_DATE, ConvertDB2Time(vadrsScheduledExceptions!END_TIME), True) Then
                        
                                ladrsSchedule!EXCEPTION_CD = ""
                                ladrsSchedule!OVER_EXC_CD = ""
                                ladrsSchedule.Update
                            End If
                            End If
                            ladrsSchedule.MoveNext
                        Loop
                    End If
                End If
                ladrsSchedule.Filter = adFilterNone
                vadrsScheduledExceptions.MoveNext
            Loop
        End If
        vadrsScheduledExceptions.Filter = adFilterNone
        Set lobjdbCodeTables = Nothing
    End If
    
    '//if the absence time is on actual time, then use absence dates to indicate a deletion
    If RecordsExist(vadrsAbsenceDates) Then
        vadrsAbsenceDates.Filter = "RECORD_STATUS = 'D'"
        If RecordsExist(vadrsAbsenceDates) Then
            vadrsAbsenceDates.MoveFirst
            
            Do While Not vadrsAbsenceDates.EOF
                ladrsSchedule.Filter = StringFormat("CHARGE_DATE = #{0}#", vadrsAbsenceDates!ABS_DATE)

                If RecordsExist(ladrsSchedule) Then
                    ladrsSchedule.MoveFirst
                    Do While Not ladrsSchedule.EOF
                        If Trim(ladrsSchedule!START_TIME) <> "" And Trim(ladrsSchedule!END_TIME) <> "" Then '// caused a type mismatch in DateTimeFrom->IsTimeValid
                        If DateTimeFrom(ladrsSchedule!CHARGE_DATE, ladrsSchedule!START_TIME, False) >= DateTimeFrom(vadrsInitialAbsence!ABS_ST_DT, ConvertDB2Time(vadrsInitialAbsence!ABS_ST_TM), False) And _
                            DateTimeFrom(ladrsSchedule!CHARGE_DATE, ladrsSchedule!END_TIME, True) <= DateTimeFrom(vadrsInitialAbsence!ABS_END_DT, ConvertDB2Time(vadrsInitialAbsence!ABS_END_TM), True) Then

                            ladrsSchedule!EXCEPTION_CD = ""
                            ladrsSchedule!OVER_EXC_CD = ""
                            ladrsSchedule.Update
                        End If
                        End If
                        
                        ladrsSchedule.MoveNext
                    Loop
                End If
                ladrsSchedule.Filter = adFilterNone
                vadrsAbsenceDates.MoveNext
            Loop
        End If
        vadrsAbsenceDates.Filter = adFilterNone
    End If

    ApplyPendingChanges = True
    Exit Function
eh:
    ApplyPendingChanges = False
    Set lobjdbCodeTables = Nothing
    TrsRaiseError MODULENAME, "ApplyPendingChanges", Err.Number, , , , True
    Exit Function
End Function

Private Function GetUniqueValues(ByVal rs As ADODB.Recordset, ByVal sColumnName As String) As Collection
    Dim UniqueCollection As utilTRISCommonLib.cuSet
    
    Set UniqueCollection = New utilTRISCommonLib.cuSet
'    GetUniqueValues = ""
    
    If RecordsExist(rs) Then
        rs.MoveFirst
        Do While Not rs.EOF
            UniqueCollection.Add (rs.Fields(sColumnName).Value)
            rs.MoveNext
        Loop
    End If
    
    Set GetUniqueValues = UniqueCollection.Contents
    
    Set UniqueCollection = Nothing
End Function

'//TODO: Use parameter sColumnName As object
Private Function GetUniqueValues2(ByVal rs As ADODB.Recordset, ByVal sColumnNames As Variant) As Collection
    Dim UniqueCollection As utilTRISCommonLib.cuSet
    Dim sColumnName As String
    Dim vColumnValues() As Variant
    Dim sColumnValues As String
    Dim i As Long
    Set UniqueCollection = New utilTRISCommonLib.cuSet
'    GetUniqueValues = ""
    
    If RecordsExist(rs) Then
        ReDim vColumnValues(0 To UBound(sColumnNames) - LBound(sColumnNames))
        rs.MoveFirst
        Do While Not rs.EOF
            For i = 0 To UBound(sColumnNames) - LBound(sColumnNames)
            'For Each sColumnName In sColumnNames
                sColumnName = sColumnNames(i)
                vColumnValues(i) = rs.Fields(sColumnName).Value
            Next
            
            '//cuSet cannot handle comparison between arrays
            sColumnValues = Join(vColumnValues, ";")
            UniqueCollection.Add sColumnValues
            rs.MoveNext
        Loop
    End If
    
    Set GetUniqueValues2 = UniqueCollection.Contents
    
    Set UniqueCollection = Nothing
End Function

Private Function CollectionAsDateRange(ByVal col As Collection) As Collection
    'Dim CollectionHelper As utilTRISCommonLib.cuContainer
    Dim colRange As Collection
    Dim datRangeStart As Date, datRangeEnd As Date
    Dim datCurrent As Date, datPrevious As Date
    
    Dim i As Long
    
    'Set CollectionHelper = New utilTRISCommonLib.cuContainer
    Set colRange = New Collection
    
    If col.Count > 0 Then
        datRangeStart = col(1)
        datRangeEnd = datRangeStart
        datPrevious = datRangeStart
        
        For i = 2 To col.Count
            datCurrent = col.Item(i)
            If DateDiff("d", datPrevious, datCurrent) = 1 Then
                datRangeEnd = datCurrent
            Else '//flush, ideally this would be a Range object with a StartDate and EndDate
                If datRangeStart = datRangeEnd Then
                    colRange.Add StringFormat("{0}", Format(datRangeStart, "m/d"))
                Else
                    colRange.Add StringFormat("{0} - {1}", Format(datRangeStart, "m/d"), Format(datRangeEnd, "m/d"))
                End If
                datRangeStart = datCurrent
                datRangeEnd = datRangeStart
            End If
            datPrevious = datCurrent
        Next i
        If datRangeStart = datRangeEnd Then
            colRange.Add StringFormat("{0}", Format(datRangeStart, "m/d"))
        Else
            colRange.Add StringFormat("{0} - {1}", Format(datRangeStart, "m/d"), Format(datRangeEnd, "m/d"))
        End If
    End If
    
    Set CollectionAsDateRange = colRange
'    CollectionAsDelimitedString = CollectionHelper.AsDelimitedString(col, False, ", ", "")
    
    'Set CollectionHelper = Nothing
End Function

'date time date time
Private Function CollectionAsDateTimeRange(ByVal col As Collection) As Collection
    'Dim CollectionHelper As utilTRISCommonLib.cuContainer
    Dim colRange As Collection
    Dim datRangeStart As Date, datRangeEnd As Date
    Dim datCurrent As Date, datPrevious As Date
    
    Dim i As Long
    
    'Set CollectionHelper = New utilTRISCommonLib.cuContainer
    Set colRange = New Collection
    
    If col.Count > 0 Then
        datRangeStart = col(1)
        datRangeEnd = datRangeStart
        datPrevious = datRangeStart
        
        For i = 2 To col.Count
            datCurrent = col.Item(i)
            If DateDiff("d", datPrevious, datCurrent) = 1 Then
                datRangeEnd = datCurrent
            Else '//flush, ideally this would be a Range object with a StartDate and EndDate
                If datRangeStart = datRangeEnd Then
                    colRange.Add StringFormat("{0}", Format(datRangeStart, "m/d"))
                Else
                    colRange.Add StringFormat("{0} - {1}", Format(datRangeStart, "m/d"), Format(datRangeEnd, "m/d"))
                End If
                datRangeStart = datCurrent
                datRangeEnd = datRangeStart
            End If
            datPrevious = datCurrent
        Next i
        If datRangeStart = datRangeEnd Then
            colRange.Add StringFormat("{0}", Format(datRangeStart, "m/d"))
        Else
            colRange.Add StringFormat("{0} - {1}", Format(datRangeStart, "m/d"), Format(datRangeEnd, "m/d"))
        End If
    End If
    
    Set CollectionAsDateTimeRange = colRange
'    CollectionAsDelimitedString = CollectionHelper.AsDelimitedString(col, False, ", ", "")
    
    'Set CollectionHelper = Nothing
End Function

'Private Function CollectionRangeAsDelimitedString(ByVal col As Collection) As String
''    Dim CollectionHelper As utilTRISCommonLib.cuContainer
''
''    Set CollectionHelper = New utilTRISCommonLib.cuContainer
''
''    CollectionRangeAsDelimitedString = CollectionHelper.AsDelimitedString(col, False, ", ", "")
'
'    For i = 1 To UniqueCollection.Count
'        CollectionRangeAsDelimitedString = CollectionRangeAsDelimitedString & UniqueCollection.Item(i)
'        If i < UniqueCollection.Count Then 'not the last one
'            CollectionRangeAsDelimitedString = CollectionAsDelimitedString & ", "
'        End If
'    Next
''    Set CollectionHelper = Nothing
'End Function

Private Function CollectionAsDelimitedString(ByVal col As Collection) As String
    Dim CollectionHelper As utilTRISCommonLib.cuContainer
    
    Set CollectionHelper = New utilTRISCommonLib.cuContainer
       
    CollectionAsDelimitedString = CollectionHelper.AsDelimitedString(col, False, ", ", "")
    
'    For i = 1 To UniqueCollection.Count
'        CollectionAsDelimitedString  = CollectionAsDelimitedString & UniqueCollection.Item(i)
'        If i < UniqueCollection.Count Then 'not the last one
'            CollectionAsDelimitedString = CollectionAsDelimitedString & ", "
'        End If
'    Next
    Set CollectionHelper = Nothing
End Function

Private Function GetDuplicateValues(rs As ADODB.Recordset, sColumnName As String) As String
    Dim CollectionHelper As utilTRISCommonLib.cuContainer
    Dim AllValuesCollection As Collection
    Dim DuplicateValuesCollection As utilTRISCommonLib.cuSet 'Collection
    Dim sValue As String
    Dim i As Long
    
    Set CollectionHelper = New utilTRISCommonLib.cuContainer
    Set AllValuesCollection = New Collection
    Set DuplicateValuesCollection = New utilTRISCommonLib.cuSet 'Collection
    
    GetDuplicateValues = ""
    
    If RecordsExist(rs) Then
        rs.MoveFirst
        Do While Not rs.EOF
            sValue = rs.Fields(sColumnName).Value
            If CollectionHelper.IncludesString(sValue, AllValuesCollection) Then
                DuplicateValuesCollection.Add sValue
'                GetDuplicateValues = GetDuplicateValues & sValue
''                If i < UniqueCollection.Count Then 'not the last one
'                    GetDuplicateValues = GetDuplicateValues & ", "
''                End If
            End If
            AllValuesCollection.Add sValue
            rs.MoveNext
        Loop
    End If
    
'    If Len(GetDuplicateValues) >= 2 Then
'        GetDuplicateValues = Left(GetDuplicateValues, Len(GetDuplicateValues) - 2)
'    End If
    
    GetDuplicateValues = CollectionHelper.AsDelimitedString(DuplicateValuesCollection.Contents, False, ", ", "")
    
    Set CollectionHelper = Nothing
    Set AllValuesCollection = Nothing
    Set DuplicateValuesCollection = Nothing
End Function

Private Function GetDistinctCount(rs As ADODB.Recordset, sColumnName As String) As Long
    Dim UniqueCollection As utilTRISCommonLib.cuSet
    Dim i As Long
    
    Set UniqueCollection = New utilTRISCommonLib.cuSet
    GetDistinctCount = -1
    
    If RecordsExist(rs) Then
        rs.MoveFirst
        Do While Not rs.EOF
            UniqueCollection.Add (rs.Fields(sColumnName).Value)
            rs.MoveNext
        Loop
    End If
    
    GetDistinctCount = UniqueCollection.Count
    
    Set UniqueCollection = Nothing
End Function

'//time range
Private Function CreateGAP() As ADODB.Recordset
    Dim rs As New ADODB.Recordset

    With rs
        '.Fields.Append "EMP_NUM", adInteger, , adFldUpdatable
        '.Fields.Append "CHARGE_DATE", adDate, , adFldUpdatable
        '.Fields.Append "DATE_WORKED", adDate, , adFldUpdatable
        .Fields.Append "GAP_START_DATE", adDate, , adFldUpdatable
        .Fields.Append "GAP_START_TIME", adChar, 5, adFldUpdatable
        .Fields.Append "GAP_END_DATE", adDate, , adFldUpdatable
        .Fields.Append "GAP_END_TIME", adChar, 5, adFldUpdatable
        '.Fields.Append "SHIFT_HOURS", 14, 16, adFldUpdatable
        '.Fields.Append "START_TIME", adChar, 5, adFldUpdatable
        '.Fields.Append "END_TIME", adChar, 5, adFldUpdatable
        '.Fields.Append "HOURS", 14, 16, adFldUpdatable
        '.Fields.Append "EXCEPTION_CD", adChar, 3, adFldUpdatable
        '.Fields.Append "OVER_EXC_CD", adChar, 3, adFldUpdatable

        .CursorLocation = adUseClient
        .CursorType = adOpenDynamic
        .LockType = adLockBatchOptimistic
        .Open
    End With

    Set CreateGAP = rs
    Set rs = Nothing
    Exit Function
End Function

'//time range split over midnight
Private Function CreateScheduleInGAP() As ADODB.Recordset
    Dim rs As New ADODB.Recordset

    With rs
        .Fields.Append "CHARGE_DATE", adDate, , adFldUpdatable
        .Fields.Append "DATE_WORKED", adDate, , adFldUpdatable
        .Fields.Append "START_TIME", adChar, 5, adFldUpdatable
        .Fields.Append "END_TIME", adChar, 5, adFldUpdatable

        .CursorLocation = adUseClient
        .CursorType = adOpenDynamic
        .LockType = adLockBatchOptimistic
        .Open
    End With

    Set CreateScheduleInGAP = rs
    Set rs = Nothing
    Exit Function
End Function

Private Function CreateSchedExcForValidation() As ADODB.Recordset
    Dim rs As New ADODB.Recordset

    With rs
        .Fields.Append "EMP_NUM", adInteger, , adFldUpdatable
        .Fields.Append "CHARGE_DATE", adDate, , adFldUpdatable
        .Fields.Append "DATE_WORKED", adDate, , adFldUpdatable
        .Fields.Append "START_TIME", adChar, 5, adFldUpdatable
        .Fields.Append "END_TIME", adChar, 5, adFldUpdatable
        .Fields.Append "HOURS", adDecimal, 16, adFldUpdatable
        .Fields.Item("HOURS").Precision = 3
        .Fields.Item("HOURS").NumericScale = 1
        
        .Fields.Append "EXCEPTION_CD", adChar, 3, adFldUpdatable
        .Fields.Append "OVER_EXC_CD", adChar, 3, adFldUpdatable

        .CursorLocation = adUseClient
        .CursorType = adOpenDynamic
        .LockType = adLockBatchOptimistic
        .Open
    End With

    Set CreateSchedExcForValidation = rs
    Set rs = Nothing
    Exit Function
End Function

Private Function CreateScheduleForValidation(ByVal vadrsSchedule As ADODB.Recordset) As ADODB.Recordset
    Dim rsScheduleForValidation As ADODB.Recordset
    Dim ChargeDate As Date, DateWorked As Date
    Dim StartTime As String, EndTime As String
    Dim Hours As Double
    Dim Bookmark As Variant
    Dim blnInit As Boolean, blnFlush As Boolean
    Dim SummaryStartTime As String, SummaryEndTime As String
    Dim SummaryHours As Double
    
    Set rsScheduleForValidation = CreateSchedExcForValidation()

    If RecordsExist(vadrsSchedule) Then
        vadrsSchedule.Sort = "EMP_NUM, CHARGE_DATE, DATE_WORKED, START_TIME"
        vadrsSchedule.MoveFirst
        
        '//need to summarize the schedule in order to match the scheduled exceptions
        ChargeDate = vadrsSchedule.Fields("CHARGE_DATE").Value
        DateWorked = vadrsSchedule.Fields("DATE_WORKED").Value
        StartTime = vadrsSchedule.Fields("START_TIME").Value
        EndTime = vadrsSchedule.Fields("END_TIME").Value
        Hours = vadrsSchedule.Fields("HOURS").Value
        Bookmark = vadrsSchedule.Bookmark
        
        blnInit = True
        blnFlush = False
        SummaryStartTime = StartTime
        SummaryEndTime = EndTime
        SummaryHours = Hours
        
        Do While Not vadrsSchedule.EOF
            If vadrsSchedule.Bookmark <> Bookmark Then
                '//compare for contiguous
                If ChargeDate = vadrsSchedule.Fields("CHARGE_DATE").Value And DateWorked = vadrsSchedule.Fields("DATE_WORKED").Value Then
                    If vadrsSchedule.Fields("START_TIME").Value = EndTime Then
                        SummaryHours = SummaryHours + vadrsSchedule.Fields("HOURS").Value
                        SummaryEndTime = vadrsSchedule.Fields("END_TIME").Value
                        EndTime = vadrsSchedule.Fields("END_TIME").Value
                        blnFlush = False
                    Else
                        blnFlush = True
                    End If
                Else
                    blnFlush = True
                End If
            End If
            
            '//flush
            If blnFlush Then
                rsScheduleForValidation.AddNew
                rsScheduleForValidation.Fields("EMP_NUM").Value = vadrsSchedule.Fields("EMP_NUM").Value
                rsScheduleForValidation.Fields("CHARGE_DATE").Value = ChargeDate
                rsScheduleForValidation.Fields("DATE_WORKED").Value = DateWorked
                rsScheduleForValidation.Fields("START_TIME").Value = SummaryStartTime
                rsScheduleForValidation.Fields("END_TIME").Value = SummaryEndTime
                rsScheduleForValidation.Fields("HOURS").Value = SummaryHours
                rsScheduleForValidation.Fields("EXCEPTION_CD").Value = vadrsSchedule.Fields("EXCEPTION_CD").Value
                rsScheduleForValidation.Fields("OVER_EXC_CD").Value = vadrsSchedule.Fields("OVER_EXC_CD").Value
                rsScheduleForValidation.Update
                
                ChargeDate = vadrsSchedule.Fields("CHARGE_DATE").Value
                DateWorked = vadrsSchedule.Fields("DATE_WORKED").Value
                StartTime = vadrsSchedule.Fields("START_TIME").Value
                EndTime = vadrsSchedule.Fields("END_TIME").Value
                Hours = vadrsSchedule.Fields("HOURS").Value
                Bookmark = vadrsSchedule.Bookmark

                blnFlush = False
                blnInit = True
                SummaryStartTime = StartTime
                SummaryEndTime = EndTime
                SummaryHours = Hours
            End If
            
            vadrsSchedule.MoveNext
        Loop
        
        If SummaryHours > 0 Then
            vadrsSchedule.MoveLast
            
            rsScheduleForValidation.AddNew
            rsScheduleForValidation.Fields("EMP_NUM").Value = vadrsSchedule.Fields("EMP_NUM").Value
            rsScheduleForValidation.Fields("CHARGE_DATE").Value = ChargeDate
            rsScheduleForValidation.Fields("DATE_WORKED").Value = DateWorked
            rsScheduleForValidation.Fields("START_TIME").Value = SummaryStartTime
            rsScheduleForValidation.Fields("END_TIME").Value = SummaryEndTime
            rsScheduleForValidation.Fields("HOURS").Value = SummaryHours
            rsScheduleForValidation.Fields("EXCEPTION_CD").Value = vadrsSchedule.Fields("EXCEPTION_CD").Value
            rsScheduleForValidation.Fields("OVER_EXC_CD").Value = vadrsSchedule.Fields("OVER_EXC_CD").Value
            rsScheduleForValidation.Update
        End If
    End If
    
    Set CreateScheduleForValidation = rsScheduleForValidation
End Function

Private Function SplitSchedExcOverMidnight(ByVal vadrsScheduledExceptions As ADODB.Recordset) As ADODB.Recordset
    Dim rsSchedExcSplit As ADODB.Recordset
    Dim dtStartDate1 As Date, dtEndDate1 As Date, dtStartDate2 As Date, dtEndDate2 As Date
    Set rsSchedExcSplit = CreateSchedExcForValidation()
    Dim dblHours1 As Double, dblHours2 As Double
    
    If RecordsExist(vadrsScheduledExceptions) Then
        vadrsScheduledExceptions.MoveFirst
    
        Do While Not vadrsScheduledExceptions.EOF
            If SplitOverMidnight(vadrsScheduledExceptions.Fields("CHARGE_DATE").Value, vadrsScheduledExceptions.Fields("START_TIME").Value, vadrsScheduledExceptions.Fields("END_TIME").Value, _
                                dtStartDate1, dtEndDate1, dtStartDate2, dtEndDate2) Then
                dblHours1 = HoursDiff(FormatTime(dtStartDate1), FormatTime(dtEndDate1))
                dblHours2 = HoursDiff(FormatTime(dtStartDate2), FormatTime(dtEndDate2))
                
                rsSchedExcSplit.AddNew
                rsSchedExcSplit.Fields("EMP_NUM").Value = vadrsScheduledExceptions.Fields("EMP_NUM").Value
                rsSchedExcSplit.Fields("CHARGE_DATE").Value = vadrsScheduledExceptions.Fields("CHARGE_DATE").Value
                rsSchedExcSplit.Fields("DATE_WORKED").Value = AddDays(IIf(dblHours1 >= dblHours2, 0, -1), MidnightDate(dtStartDate1))
                rsSchedExcSplit.Fields("START_TIME").Value = FormatTime(dtStartDate1)
                rsSchedExcSplit.Fields("END_TIME").Value = FormatTime(dtEndDate1)
                rsSchedExcSplit.Fields("HOURS").Value = dblHours1 'HoursDiff(rsSchedExcSplit.Fields("START_TIME").Value, rsSchedExcSplit.Fields("END_TIME").Value)
                rsSchedExcSplit.Fields("EXCEPTION_CD").Value = vadrsScheduledExceptions.Fields("EXCEPTION_CD").Value
                rsSchedExcSplit.Fields("OVER_EXC_CD").Value = vadrsScheduledExceptions.Fields("OVER_EXC_CD").Value
                rsSchedExcSplit.Update
                
                rsSchedExcSplit.AddNew
                rsSchedExcSplit.Fields("EMP_NUM").Value = vadrsScheduledExceptions.Fields("EMP_NUM").Value
                rsSchedExcSplit.Fields("CHARGE_DATE").Value = vadrsScheduledExceptions.Fields("CHARGE_DATE").Value
                rsSchedExcSplit.Fields("DATE_WORKED").Value = AddDays(IIf(dblHours1 >= dblHours2, 0, -1), MidnightDate(dtStartDate2))
                rsSchedExcSplit.Fields("START_TIME").Value = FormatTime(dtStartDate2)
                rsSchedExcSplit.Fields("END_TIME").Value = FormatTime(dtEndDate2)
                rsSchedExcSplit.Fields("HOURS").Value = dblHours2 'HoursDiff(rsSchedExcSplit.Fields("START_TIME").Value, rsSchedExcSplit.Fields("END_TIME").Value)
                rsSchedExcSplit.Fields("EXCEPTION_CD").Value = vadrsScheduledExceptions.Fields("EXCEPTION_CD").Value
                rsSchedExcSplit.Fields("OVER_EXC_CD").Value = vadrsScheduledExceptions.Fields("OVER_EXC_CD").Value
                rsSchedExcSplit.Update
            Else
                rsSchedExcSplit.AddNew
                rsSchedExcSplit.Fields("EMP_NUM").Value = vadrsScheduledExceptions.Fields("EMP_NUM").Value
                rsSchedExcSplit.Fields("CHARGE_DATE").Value = vadrsScheduledExceptions.Fields("CHARGE_DATE").Value
                rsSchedExcSplit.Fields("DATE_WORKED").Value = vadrsScheduledExceptions.Fields("CHARGE_DATE").Value
                rsSchedExcSplit.Fields("START_TIME").Value = vadrsScheduledExceptions.Fields("START_TIME").Value
                rsSchedExcSplit.Fields("END_TIME").Value = vadrsScheduledExceptions.Fields("END_TIME").Value
                rsSchedExcSplit.Fields("HOURS").Value = HoursDiff(rsSchedExcSplit.Fields("START_TIME").Value, rsSchedExcSplit.Fields("END_TIME").Value)
                rsSchedExcSplit.Fields("EXCEPTION_CD").Value = vadrsScheduledExceptions.Fields("EXCEPTION_CD").Value
                rsSchedExcSplit.Fields("OVER_EXC_CD").Value = vadrsScheduledExceptions.Fields("OVER_EXC_CD").Value
                rsSchedExcSplit.Update
            End If
        
            vadrsScheduledExceptions.MoveNext
        Loop
    End If
    
    Set SplitSchedExcOverMidnight = rsSchedExcSplit
End Function

'//-----start of compare

'//-----end of compare

Public Function ProcessEAMRequest(ByVal vlEmployeeNumber As Long _
                , ByVal vadrsAbsenceDates As ADODB.Recordset _
                , ByVal vadrsInitialAbsence As ADODB.Recordset _
                , ByVal vadrsAbsenceProvider As ADODB.Recordset _
                , ByVal vadrsAbsenceContact As ADODB.Recordset _
                , ByVal vadrsAbsenceReturnToWork As ADODB.Recordset _
                , ByVal vadrsScheduledExceptions As ADODB.Recordset _
                , ByVal vadrsAbsenceCallLog As ADODB.Recordset _
                , ByVal vadrsExceptions As ADODB.Recordset _
                , ByVal vadrsEmployeeInformation As ADODB.Recordset _
                , ByVal vlTRISUserNumber As Long _
                , ByVal vsTRISUserID As String _
                , ByRef rlAbsenceNumber As Long _
                , ByRef radrsBRV As ADODB.Recordset _
                ) As Boolean
On Error GoTo ProcessEAMRequest_EH
Dim llOriginalAbsenceNumber As Long
'Flag
Dim lbCheck As Boolean

'Recordsets
'Lock Status
Dim ladrsLockStatus As ADODB.Recordset

'Time
Dim ladrsTime As ADODB.Recordset
Dim ladrsTimeForSchedule As ADODB.Recordset
Dim ladrsTimeOriginal As ADODB.Recordset
Dim ladrsSchedule As ADODB.Recordset
Dim ladrsAbsenceTime As ADODB.Recordset
Dim ladrsDeleteTime As ADODB.Recordset
Dim ldtStart As Date, ldtEnd As Date
Dim ldtLastProcessedDate As Date

'Absence Start and End
Dim ladrsAbsenceStartEnd As ADODB.Recordset

'Absence
Dim ladrsInitial As ADODB.Recordset
Dim ladrsDate As ADODB.Recordset
Dim ladrsProvider As ADODB.Recordset
Dim ladrsContact As ADODB.Recordset
Dim ladrsReturnToWork As ADODB.Recordset
Dim ladrsDatesAlreadyPresent As ADODB.Recordset

'Portfolio
Dim ladrsPortfolio As ADODB.Recordset
Dim ladrsPortfolioDetails As ADODB.Recordset

'BRV
Dim ladrsBRVActual As ADODB.Recordset
Dim ladrsBRVExcSch As ADODB.Recordset

'Components
Dim lobjdbEAM As dbEAM.cdEAM
Dim lobjdbEAMTrans As dbEAM.cdEAMTrans
Dim lobjSchedExc As busSchedException.cbsSchedExcTrans
Dim lobjbusTimeKeeper As busTimekeeper.cbsTimekeeper

    ProcessEAMRequest = False

    llOriginalAbsenceNumber = rlAbsenceNumber

    Set radrsBRV = CreateNewBusinessErrorSet

    If Not (RecordsExist(vadrsAbsenceDates) Or _
        RecordsExist(vadrsInitialAbsence) Or _
        RecordsExist(vadrsAbsenceProvider) Or _
        RecordsExist(vadrsAbsenceContact) Or _
        RecordsExist(vadrsAbsenceReturnToWork) Or _
        RecordsExist(vadrsScheduledExceptions)) Then
        AppendBusinessError radrsBRV _
            , trErrbusEAM _
            , mstrCBUSINESSERR _
            , "There are no absence entries to process." _
            , MODULENAME & ": ProcessEAMRequest"

        GoTo REPORTBRV
    End If

    If Not (CheckForValidLong(vlTRISUserNumber, True)) Then
        AppendBusinessError radrsBRV _
            , trErrbusEAM _
            , mstrCBUSINESSERR _
            , "Invalid TRIS User." _
            , MODULENAME & ": ProcessEAMRequest"

        GoTo REPORTBRV
    End If

    If Not (CheckForValidLong(vlEmployeeNumber, True)) Then
        AppendBusinessError radrsBRV _
            , trErrbusEAM _
            , mstrCBUSINESSERR _
            , "Invalid Employee." _
            , MODULENAME & ": ProcessEAMRequest"

        GoTo REPORTBRV
    End If

    lbCheck = False

    If (RecordsExist(vadrsInitialAbsence)) Then
        With vadrsInitialAbsence
            .MoveFirst

            If (!RECORD_STATUS = "I") Then
                lbCheck = True
            End If
        End With
    End If

    If (RecordsExist(vadrsAbsenceDates)) Then
        With vadrsAbsenceDates
            .MoveFirst

            .Find "RECORD_STATUS = 'I'"
            If Not (.EOF) Then
                lbCheck = True
            End If

            If Not (lbCheck) Then
                .MoveFirst

                .Find "RECORD_STATUS = 'D'"
                If Not (.EOF) Then
                    lbCheck = True
                End If
            End If

            .MoveFirst
        End With

        'Simple Check for Valid Dates
        If Not (ValidateDates(vadrsAbsenceDates _
                , vadrsEmployeeInformation _
                , radrsBRV)) Then
            GoTo REPORTBRV
        End If
    End If

    If (RecordsExist(vadrsAbsenceDates)) And _
        (RecordsExist(vadrsInitialAbsence)) Then
        vadrsInitialAbsence.MoveFirst
        ldtStart = vadrsInitialAbsence!ABS_ST_DT

        With vadrsAbsenceDates
            .Sort = "ABS_DATE"

            .MoveFirst
            If (ldtStart > !ABS_DATE) Then
                ldtStart = !ABS_DATE
            End If

            .MoveLast
            ldtEnd = !ABS_DATE

            .MoveFirst
        End With

        Set ladrsLockStatus = GetLockStatus(vlEmployeeNumber _
                                , ldtStart _
                                , ldtEnd)

        ldtStart = vbNull
        ldtEnd = vbNull
               
        Set ladrsSchedule = GetEmployeeSchedule(vlEmployeeNumber, vadrsInitialAbsence, vadrsAbsenceDates)
        '//if absence is unextended, sched exc from db may prevent the probable RTW date from being determined properly
        '//try removing deleted exceptions from the schedule, this will allow the RTW to be sooner
        ApplyPendingChanges vadrsScheduledExceptions, ladrsSchedule, vadrsInitialAbsence, vadrsAbsenceDates
    End If

    If (lbCheck) Then
        Set ladrsDatesAlreadyPresent = Nothing

        If (rlAbsenceNumber > -1) Then
            Set lobjdbEAM = CreateObject("dbEAM.cdEAM")
            Set ladrsDatesAlreadyPresent = lobjdbEAM.GetAbsenceDatesByAbsence( _
                                            rlAbsenceNumber)
            Set lobjdbEAM = Nothing
        End If

        'NOTE: Schedule should have info at least 30 days over the Absence End Date
        Set ladrsTime = GetEmployeeTimeDetails(vlEmployeeNumber _
                                    , vadrsInitialAbsence _
                                    , vadrsAbsenceDates)
        Set ladrsTimeOriginal = CopyRecordset(ladrsTime)
        '//TODO: TEST
        ApplyPendingChanges vadrsScheduledExceptions, ladrsTime, vadrsInitialAbsence, vadrsAbsenceDates

        Set vadrsInitialAbsence = UpdateAbsenceWithDateWorked(vadrsInitialAbsence, ladrsTime)
        
        Set ladrsTimeForSchedule = MergeScheduleWithTime(vlEmployeeNumber, vadrsInitialAbsence, vadrsAbsenceDates, vadrsScheduledExceptions)
        
        'Determine if non consecutive dates were selected
        If (RecordsExist(ladrsSchedule)) Or _
            (RecordsExist(ladrsTime)) Then
            Set ladrsAbsenceTime = AddTimeAndSchExceptions(vadrsInitialAbsence _
                                        , vadrsScheduledExceptions _
                                        , ladrsSchedule _
                                        , ladrsTimeForSchedule _
                                        , ladrsLockStatus)

'            Set ladrsAbsenceStartEnd = DetermineAbsenceStartAndEndOld(ladrsAbsenceTime _
'                                        , ladrsSchedule _
'                                        , vadrsExceptions _
'                                        , ladrsLockStatus)

            Set ladrsAbsenceStartEnd = DetermineAbsenceStartAndEnd(ladrsAbsenceTime _
                                        , ladrsTimeForSchedule _
                                        , vadrsExceptions _
                                        , ladrsLockStatus _
                                        , ladrsTime)

            Set ladrsLockStatus = Nothing

            If (RecordsExist(ladrsAbsenceStartEnd)) Then
                'Create Possible Multiple
                CreateAndAddAllAbsenceRecords vlEmployeeNumber _
                    , vadrsInitialAbsence _
                    , vadrsAbsenceDates _
                    , vadrsAbsenceReturnToWork _
                    , ladrsDatesAlreadyPresent _
                    , ladrsSchedule _
                    , ladrsAbsenceStartEnd _
                    , ladrsInitial _
                    , ladrsDate _
                    , ladrsReturnToWork

                Set ladrsDatesAlreadyPresent = Nothing
            End If

            If Not (vadrsEmployeeInformation Is Nothing) Then
                If Not (vadrsEmployeeInformation.EOF And vadrsEmployeeInformation.BOF) Then
                    ldtLastProcessedDate = vadrsEmployeeInformation!LAST_PP_PROC_DATE
                End If
            End If

            If (RecordsExist(ladrsDate)) Then
                ladrsDate.MoveFirst

                ladrsDate.Find "RECORD_STATUS = 'D'"
                If Not (ladrsDate.EOF) Then
                    '//found an issue when deleting a day from an absence
                    '//the absence end date had already been re-calced, so that the deleted day did not fit within the
                    '//abs start and end dates (unfortunately emp_num is not appear to be populated)
                    '//so use the original absence request recordset for this
                    If RecordsExist(vadrsInitialAbsence) Then
                        vadrsInitialAbsence.MoveFirst
                        Do While Not vadrsInitialAbsence.EOF
                            vadrsInitialAbsence!EMP_NUM = vlEmployeeNumber
                            vadrsInitialAbsence.Update
                            
                            vadrsInitialAbsence.MoveNext
                        Loop
                    End If
                    
                    PrepareDeleteExcTimeRS vadrsInitialAbsence _
                        , Nothing _
                        , ladrsTimeOriginal _
                        , ladrsSchedule _
                        , ladrsDate _
                        , "" _
                        , ldtLastProcessedDate _
                        , ladrsDeleteTime _
                        , Nothing
                End If

                ladrsDate.MoveFirst
            End If
        End If

        Set ladrsDatesAlreadyPresent = Nothing
        Set ladrsTime = Nothing
        Set ladrsAbsenceTime = Nothing
        Set ladrsAbsenceStartEnd = Nothing

        'Validate
        If (RecordsExist(ladrsInitial)) Then
            If Not (ValidateInitialAbsence(ladrsInitial _
                , radrsBRV)) Then
                GoTo REPORTBRV
            Else
                If (ladrsInitial.RecordCount > 1) Then
                    Set ladrsPortfolio = AddPortfolioInformation(1 _
                                            , ladrsInitial)
                    Set ladrsPortfolioDetails = AddPortfolioInformation(2 _
                                                    , ladrsInitial)
                End If
            End If
        End If

        'Validate
        If (RecordsExist(ladrsReturnToWork)) Then
            If Not (ValidateAbsenceReturnToWork(ladrsReturnToWork _
                , radrsBRV)) Then
                GoTo REPORTBRV
            End If
        End If
    Else
        If (RecordsExist(vadrsInitialAbsence)) Then
            If Not (ValidateInitialAbsence(vadrsInitialAbsence _
                , radrsBRV)) Then
                GoTo REPORTBRV
            Else
                If (vadrsInitialAbsence.RecordCount > 1) Then
                    Set ladrsPortfolio = AddPortfolioInformation(1 _
                                            , vadrsInitialAbsence)
                    Set ladrsPortfolioDetails = AddPortfolioInformation(2 _
                                                    , vadrsInitialAbsence)
                End If
            End If
        End If

        'Validate
        If (RecordsExist(vadrsAbsenceReturnToWork)) Then
            If Not (ValidateAbsenceReturnToWork(vadrsAbsenceReturnToWork _
                , radrsBRV)) Then
                GoTo REPORTBRV
            End If
        End If
    End If

    'Validate
    If (RecordsExist(vadrsAbsenceProvider)) Then
        If Not (ValidateAbsenceProvider(vadrsAbsenceProvider _
            , radrsBRV)) Then
            GoTo REPORTBRV
        End If
    End If

    'Validate
    If (RecordsExist(vadrsAbsenceContact)) Then
        If Not (ValidateAbsenceContact(vadrsAbsenceContact _
            , radrsBRV _
            , vlTRISUserNumber)) Then
            GoTo REPORTBRV
        End If
    End If

    If (RecordsExist(vadrsAbsenceCallLog)) Then
        If Not (ValidateAbsenceCallLog(vadrsAbsenceCallLog _
            , radrsBRV)) Then
            GoTo REPORTBRV
        End If
    End If

    Set ladrsSchedule = Nothing

    'EAM
    Set lobjdbEAMTrans = CreateObject("dbEAM.cdEAMTrans")
    If (lbCheck) Then
        If Not (lobjdbEAMTrans.dbProcessEAMRequest(vlEmployeeNumber _
            , ladrsDate _
            , ladrsInitial _
            , vadrsAbsenceProvider _
            , vadrsAbsenceContact _
            , ladrsReturnToWork _
            , vadrsAbsenceCallLog _
            , vlTRISUserNumber _
            , vsTRISUserID _
            , rlAbsenceNumber _
            , radrsBRV _
            , ladrsPortfolio _
            , ladrsPortfolioDetails)) Then
            Set lobjdbEAMTrans = Nothing
            Set ladrsInitial = Nothing
            Set ladrsDate = Nothing

            GoTo REPORTBRV
        End If
    Else
        If Not (lobjdbEAMTrans.dbProcessEAMRequest(vlEmployeeNumber _
            , vadrsAbsenceDates _
            , vadrsInitialAbsence _
            , vadrsAbsenceProvider _
            , vadrsAbsenceContact _
            , vadrsAbsenceReturnToWork _
            , vadrsAbsenceCallLog _
            , vlTRISUserNumber _
            , vsTRISUserID _
            , rlAbsenceNumber _
            , radrsBRV _
            , ladrsPortfolio _
            , ladrsPortfolioDetails)) Then
            Set lobjdbEAMTrans = Nothing

            GoTo REPORTBRV
        End If
    End If
    Set ladrsInitial = Nothing
    Set ladrsDate = Nothing
    Set ladrsReturnToWork = Nothing
    Set ladrsPortfolio = Nothing
    Set ladrsPortfolioDetails = Nothing

    Set lobjdbEAMTrans = Nothing

    'Actual Time
    If (RecordsExist(ladrsDeleteTime)) And _
        (RecordsExist(vadrsScheduledExceptions)) Then
        EnsureTimeDeletion vadrsScheduledExceptions _
            , ladrsDeleteTime
    End If

    '//not sure what conditions will cause this to execute and what situations this handles
    '//adding a scheduled exception does cause an actual time row to be deleted
    '//so, in other words, this has not be tested since ProcessTimeEntries3 was called
    If (RecordsExist(ladrsDeleteTime)) Then
        ladrsDeleteTime.Sort = "KEY_CHARGE_DATE"

        ladrsDeleteTime.MoveFirst
        ldtStart = ladrsDeleteTime!KEY_CHARGE_DATE
        ladrsDeleteTime.MoveLast
        ldtEnd = ladrsDeleteTime!KEY_CHARGE_DATE

        ladrsDeleteTime.MoveFirst

        Set lobjbusTimeKeeper = CreateObject("busTimekeeper.cbsTimekeeper")
        If Not (lobjbusTimeKeeper.ProcessTimeEntries3(ladrsDeleteTime _
                , Nothing _
                , Nothing _
                , ldtStart _
                , ldtEnd _
                , Nothing _
                , Nothing _
                , madrsScheduleWithNoExceptions _
                , ladrsBRVActual _
                , CStr(vlTRISUserNumber))) Then
            Set lobjbusTimeKeeper = Nothing

            If (RecordsExist(ladrsBRVActual)) Then
                AddBRVRecords ladrsBRVActual _
                    , radrsBRV _
                    , True
            End If

            GoTo REPORTBRV
        Else
            Set lobjbusTimeKeeper = Nothing

            If (RecordsExist(ladrsBRVActual)) Then
                AddBRVRecords ladrsBRVActual _
                    , radrsBRV _
                    , True
            End If
        End If

        Set lobjbusTimeKeeper = Nothing
    End If

    Set ladrsDeleteTime = Nothing
    Set madrsScheduleWithNoExceptions = Nothing

    'Exception Scheduling
    If (RecordsExist(vadrsScheduledExceptions)) Then
        Set lobjSchedExc = CreateObject("busSchedException.cbsSchedExcTrans")
        If Not (lobjSchedExc.SaveExceptions(vadrsScheduledExceptions _
                , ladrsBRVExcSch _
                , CStr(vlTRISUserNumber))) Then
            Set lobjSchedExc = Nothing

            If (RecordsExist(ladrsBRVExcSch)) Then
                AddBRVRecords ladrsBRVExcSch _
                    , radrsBRV
            End If

            GoTo REPORTBRV
        Else
            If (RecordsExist(ladrsBRVExcSch)) Then
                AddBRVRecords ladrsBRVExcSch _
                    , radrsBRV
            End If
        End If

        Set lobjSchedExc = Nothing
    End If

    ProcessEAMRequest = True

    GetObjectContext.SetComplete

    Exit Function
ProcessEAMRequest_EH:
    'Abort Transaction
    GetObjectContext().SetAbort

    rlAbsenceNumber = llOriginalAbsenceNumber

    ProcessEAMRequest = False

    Set lobjdbEAM = Nothing
    Set lobjdbEAMTrans = Nothing
    Set lobjSchedExc = Nothing
    Set lobjbusTimeKeeper = Nothing

    Set ladrsDatesAlreadyPresent = Nothing
    Set ladrsSchedule = Nothing
    Set ladrsTime = Nothing
    Set ladrsAbsenceTime = Nothing
    Set ladrsInitial = Nothing
    Set ladrsDate = Nothing
    Set ladrsReturnToWork = Nothing
    Set ladrsPortfolio = Nothing
    Set ladrsPortfolioDetails = Nothing
    Set ladrsDeleteTime = Nothing
    Set ladrsLockStatus = Nothing
    Set madrsScheduleWithNoExceptions = Nothing

    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "ProcessEAMRequest"
            Case Else
                TrsRaiseError MODULENAME _
                    , "ProcessEAMRequest" _
                    , Err.Number _
                    , , , , _
                    True
    End Select

    Exit Function
REPORTBRV:
Dim llIndex As Long

    rlAbsenceNumber = llOriginalAbsenceNumber

    Set lobjdbEAM = Nothing
    Set lobjdbEAMTrans = Nothing
    Set lobjSchedExc = Nothing
    Set lobjbusTimeKeeper = Nothing

    Set ladrsDatesAlreadyPresent = Nothing
    Set ladrsSchedule = Nothing
    Set ladrsTime = Nothing
    Set ladrsAbsenceTime = Nothing
    Set ladrsInitial = Nothing
    Set ladrsDate = Nothing
    Set ladrsReturnToWork = Nothing
    Set ladrsPortfolio = Nothing
    Set ladrsPortfolioDetails = Nothing
    Set ladrsDeleteTime = Nothing
    Set ladrsLockStatus = Nothing
    Set madrsScheduleWithNoExceptions = Nothing

    If (RecordsExist(radrsBRV)) Then
        With radrsBRV
            .MoveFirst

            'Check for errors
            .Find "ERR_TYPE = 'BE'"
            'If found abort else complete
            If Not (.EOF) Then
                .Filter = adFilterNone
                .Filter = "ERR_TYPE = 'BI'"
                If Not (.EOF) Then
                    .MoveLast
                    For llIndex = .RecordCount To 1 Step -1
                        .Delete adAffectCurrent

                        .MovePrevious
                    Next llIndex
                End If

                .Filter = adFilterNone

                'Abort Transaction
                GetObjectContext().SetAbort
                ProcessEAMRequest = False
            Else
                'Complete Transaction
                GetObjectContext().SetComplete
                ProcessEAMRequest = True
            End If

            .MoveFirst
        End With
    Else
        'Complete Transaction
        GetObjectContext().SetComplete
        ProcessEAMRequest = True
    End If

    Exit Function

End Function

'//for testing only
'Public Function ProcessEAMRequestTest(ByVal vlEmployeeNumber As Long _
'                , ByVal vadrsAbsenceDates As ADODB.Recordset _
'                , ByVal vadrsInitialAbsence As ADODB.Recordset _
'                , ByVal vadrsAbsenceProvider As ADODB.Recordset _
'                , ByVal vadrsAbsenceContact As ADODB.Recordset _
'                , ByVal vadrsAbsenceReturnToWork As ADODB.Recordset _
'                , ByVal vadrsScheduledExceptions As ADODB.Recordset _
'                , ByVal vadrsAbsenceCallLog As ADODB.Recordset _
'                , ByVal vadrsExceptions As ADODB.Recordset _
'                , ByVal vadrsEmployeeInformation As ADODB.Recordset _
'                , ByVal vlTRISUserNumber As Long _
'                , ByVal vsTRISUserID As String _
'                , ByRef rlAbsenceNumber As Long _
'                , ByRef radrsBRV As ADODB.Recordset _
'                , ByVal vstrTestCase As String _
'                , ByRef vobjTestCaseResults() As Object _
'                ) As Boolean
'
'On Error GoTo eh
'    Dim llOriginalAbsenceNumber As Long
'    'Flag
'    Dim lbCheck As Boolean
'
'    'Recordsets
'    'Lock Status
'    Dim ladrsLockStatus As ADODB.Recordset
'
'    'Time
'    Dim ladrsTime As ADODB.Recordset
'    Dim ladrsSchedule As ADODB.Recordset
'    Dim ladrsAbsenceTime As ADODB.Recordset
'    Dim ladrsDeleteTime As ADODB.Recordset
'    Dim ldtStart As Date, ldtEnd As Date
'    Dim ldtLastProcessedDate As Date
'
'    'Absence Start and End
'    Dim ladrsAbsenceStartEnd As ADODB.Recordset
'
'    'Absence
'    Dim ladrsInitial As ADODB.Recordset
'    Dim ladrsDate As ADODB.Recordset
'    Dim ladrsProvider As ADODB.Recordset
'    Dim ladrsContact As ADODB.Recordset
'    Dim ladrsReturnToWork As ADODB.Recordset
'    Dim ladrsDatesAlreadyPresent As ADODB.Recordset
'
'    'Portfolio
'    Dim ladrsPortfolio As ADODB.Recordset
'    Dim ladrsPortfolioDetails As ADODB.Recordset
'
'    'BRV
'    Dim ladrsBRVActual As ADODB.Recordset
'    Dim ladrsBRVExcSch As ADODB.Recordset
'
'    'Components
'    Dim lobjdbEAM As dbEAM.cdEAM
'    Dim lobjdbEAMTrans As dbEAM.cdEAMTrans
'    Dim lobjSchedExc As busSchedException.cbsSchedExcTrans
'    Dim lobjbusTimeKeeper As busTimekeeper.cbsTimekeeper
'
'    ProcessEAMRequestTest = False
'
'    llOriginalAbsenceNumber = rlAbsenceNumber
'
'    Set radrsBRV = CreateNewBusinessErrorSet
'
'    If Not (RecordsExist(vadrsAbsenceDates) Or _
'        RecordsExist(vadrsInitialAbsence) Or _
'        RecordsExist(vadrsAbsenceProvider) Or _
'        RecordsExist(vadrsAbsenceContact) Or _
'        RecordsExist(vadrsAbsenceReturnToWork) Or _
'        RecordsExist(vadrsScheduledExceptions)) Then
'        AppendBusinessError radrsBRV, trErrbusEAM, mstrCBUSINESSERR, "There are no absence entries to process.", MODULENAME & ": ProcessEAMRequestTest"
'
'        GoTo REPORTBRV
'    End If
'
'    If Not (CheckForValidLong(vlTRISUserNumber, True)) Then
'        AppendBusinessError radrsBRV, trErrbusEAM, mstrCBUSINESSERR, "Invalid TRIS User.", MODULENAME & ": ProcessEAMRequestTest"
'
'        GoTo REPORTBRV
'    End If
'
'    If Not (CheckForValidLong(vlEmployeeNumber, True)) Then
'        AppendBusinessError radrsBRV, trErrbusEAM, mstrCBUSINESSERR, "Invalid Employee.", MODULENAME & ": ProcessEAMRequestTest"
'
'        GoTo REPORTBRV
'    End If
'
'    lbCheck = False
'
'    If (RecordsExist(vadrsInitialAbsence)) Then
'        With vadrsInitialAbsence
'            .MoveFirst
'
'            If (!RECORD_STATUS = "I") Then
'                lbCheck = True
'            End If
'        End With
'    End If
'
'    If (RecordsExist(vadrsAbsenceDates)) Then
'        With vadrsAbsenceDates
'            .MoveFirst
'
'            .Find "RECORD_STATUS = 'I'"
'            If Not (.EOF) Then
'                lbCheck = True
'            End If
'
'            If Not (lbCheck) Then
'                .MoveFirst
'
'                .Find "RECORD_STATUS = 'D'"
'                If Not (.EOF) Then
'                    lbCheck = True
'                End If
'            End If
'
'            .MoveFirst
'        End With
'
'        'Simple Check for Valid Dates
'        If Not ValidateDates(vadrsAbsenceDates, vadrsEmployeeInformation, radrsBRV) Then
'            GoTo REPORTBRV
'        End If
'    End If
'
'    If (RecordsExist(vadrsAbsenceDates)) And _
'        (RecordsExist(vadrsInitialAbsence)) Then
'        vadrsInitialAbsence.MoveFirst
'        ldtStart = vadrsInitialAbsence!ABS_ST_DT
'
'        With vadrsAbsenceDates
'            .Sort = "ABS_DATE"
'
'            .MoveFirst
'            If (ldtStart > !ABS_DATE) Then
'                ldtStart = !ABS_DATE
'            End If
'
'            .MoveLast
'            ldtEnd = !ABS_DATE
'
'            .MoveFirst
'        End With
'
'        Set ladrsLockStatus = GetLockStatus(vlEmployeeNumber, ldtStart, ldtEnd)
'
'        ldtStart = vbNull
'        ldtEnd = vbNull
'
'        Set ladrsSchedule = GetEmployeeSchedule(vlEmployeeNumber, vadrsInitialAbsence, vadrsAbsenceDates)
'        '//Set ladrsSchedule = GetEmployeeScheduleTest(vlEmployeeNumber, vadrsInitialAbsence, vadrsAbsenceDates)
'    End If
'
'    If (lbCheck) Then
'        Set ladrsDatesAlreadyPresent = Nothing
'
'        If (rlAbsenceNumber > -1) Then
'            Set lobjdbEAM = CreateObject("dbEAM.cdEAM")
'            Set ladrsDatesAlreadyPresent = lobjdbEAM.GetAbsenceDatesByAbsence(rlAbsenceNumber)
'            Set lobjdbEAM = Nothing
'        End If
'
'        'NOTE: Schedule should have info atleast 30 days over the Absence End Date
'        Set ladrsTime = GetEmployeeTimeDetails(vlEmployeeNumber, vadrsInitialAbsence, vadrsAbsenceDates)
'        '//Set ladrsTime = GetEmployeeTimeDetailsTest(vlEmployeeNumber, vadrsInitialAbsence, vadrsAbsenceDates)
'
'        'Determine if non consecutive dates were selected
'        If RecordsExist(ladrsSchedule) Or RecordsExist(ladrsTime) Then
'            ReDim vobjTestCaseResults(2)
'
'            Select Case UCase(vstrTestCase)
'            Case "OLD"
'    ''//the old
'                Set ladrsAbsenceTime = AddTimeAndSchExceptionsOld(vadrsInitialAbsence _
'                                            , vadrsScheduledExceptions _
'                                            , ladrsSchedule _
'                                            , ladrsTime _
'                                            , ladrsLockStatus)
'                '//temp fix for type mismatch
'                If RecordsExist(ladrsAbsenceTime) Then
'                    ladrsAbsenceTime.MoveFirst
'                    Do While Not ladrsAbsenceTime.EOF
'                        If IsNullOrEmpty(ladrsAbsenceTime!DATE_WORKED) Then
'                            ladrsAbsenceTime!DATE_WORKED = ladrsAbsenceTime!CHARGE_DATE
'                            ladrsAbsenceTime.Update
'                        End If
'                        ladrsAbsenceTime.MoveNext
'                    Loop
'                End If
'
''                If RecordsExist(ladrsAbsenceTime) Then
'                    Set vobjTestCaseResults(0) = ladrsAbsenceTime
''                End If
'
'    '//the old
'                Set ladrsAbsenceStartEnd = DetermineAbsenceStartAndEndOld(ladrsAbsenceTime _
'                                            , ladrsSchedule _
'                                            , vadrsExceptions _
'                                            , ladrsLockStatus)
'            Case "NEW"
'    '//the new
'                Set ladrsAbsenceTime = AddTimeAndSchExceptions(vadrsInitialAbsence _
'                                            , vadrsScheduledExceptions _
'                                            , ladrsSchedule _
'                                            , ladrsTime _
'                                            , ladrsLockStatus)
''                If RecordsExist(ladrsAbsenceTime) Then
'                    Set vobjTestCaseResults(0) = ladrsAbsenceTime
''                End If
'    '//the new
'                Set ladrsAbsenceStartEnd = DetermineAbsenceStartAndEnd(ladrsAbsenceTime _
'                                            , ladrsSchedule _
'                                            , vadrsExceptions _
'                                            , ladrsLockStatus)
'            End Select
'            Set ladrsLockStatus = Nothing
'
''            If RecordsExist(ladrsAbsenceStartEnd) Then
'                Set vobjTestCaseResults(1) = madrsAbsenceStartEndDetails
'                Set vobjTestCaseResults(2) = ladrsAbsenceStartEnd
''            End If
'        End If
'    End If
'
''    'Abort Transaction
''    GetObjectContext().SetAbort
'    ProcessEAMRequestTest = True
'
'    Exit Function
'eh:
'    'Abort Transaction
'    GetObjectContext().SetAbort
'
'    rlAbsenceNumber = llOriginalAbsenceNumber
'
'    ProcessEAMRequestTest = False
'
'    Set lobjdbEAM = Nothing
'    Set lobjdbEAMTrans = Nothing
'    Set lobjSchedExc = Nothing
'    Set lobjbusTimeKeeper = Nothing
'
'    Set ladrsDatesAlreadyPresent = Nothing
'    Set ladrsSchedule = Nothing
'    Set ladrsTime = Nothing
'    Set ladrsAbsenceTime = Nothing
'    Set ladrsInitial = Nothing
'    Set ladrsDate = Nothing
'    Set ladrsReturnToWork = Nothing
'    Set ladrsPortfolio = Nothing
'    Set ladrsPortfolioDetails = Nothing
'    Set ladrsDeleteTime = Nothing
'    Set ladrsLockStatus = Nothing
'    Set madrsScheduleWithNoExceptions = Nothing
'
'    Select Case (Err.Number)
'        Case trErrutOpenConnection
'            TrsRaiseError MODULENAME, "ProcessEAMRequestTest"
'        Case Else
'            TrsRaiseError MODULENAME, "ProcessEAMRequestTest", Err.Number, , , , True
'    End Select
'
'    Exit Function
'REPORTBRV:
'Dim llIndex As Long
'
'    rlAbsenceNumber = llOriginalAbsenceNumber
'
'    Set lobjdbEAM = Nothing
'    Set lobjdbEAMTrans = Nothing
'    Set lobjSchedExc = Nothing
'    Set lobjbusTimeKeeper = Nothing
'
'    Set ladrsDatesAlreadyPresent = Nothing
'    Set ladrsSchedule = Nothing
'    Set ladrsTime = Nothing
'    Set ladrsAbsenceTime = Nothing
'    Set ladrsInitial = Nothing
'    Set ladrsDate = Nothing
'    Set ladrsReturnToWork = Nothing
'    Set ladrsPortfolio = Nothing
'    Set ladrsPortfolioDetails = Nothing
'    Set ladrsDeleteTime = Nothing
'    Set ladrsLockStatus = Nothing
'    Set madrsScheduleWithNoExceptions = Nothing
'
'    If (RecordsExist(radrsBRV)) Then
'        With radrsBRV
'            .MoveFirst
'
'            'Check for errors
'            .Find "ERR_TYPE = 'BE'"
'            'If found abort else complete
'            If Not (.EOF) Then
'                .Filter = adFilterNone
'                .Filter = "ERR_TYPE = 'BI'"
'                If Not (.EOF) Then
'                    .MoveLast
'                    For llIndex = .RecordCount To 1 Step -1
'                        .Delete adAffectCurrent
'
'                        .MovePrevious
'                    Next llIndex
'                End If
'
'                .Filter = adFilterNone
'
'                'Abort Transaction
'                GetObjectContext().SetAbort
'                ProcessEAMRequestTest = False
'            Else
'                'Complete Transaction
'                GetObjectContext().SetComplete
'                ProcessEAMRequestTest = True
'            End If
'
'            .MoveFirst
'        End With
'    Else
'        'Complete Transaction
'        GetObjectContext().SetComplete
'        ProcessEAMRequestTest = True
'    End If
'End Function
'//for testing only end


Private Function ValidateDates(ByVal vadrsAbsenceDates As ADODB.Recordset _
                , ByVal vadrsEmployeeInformation As ADODB.Recordset _
                , ByRef radrsBRV As ADODB.Recordset _
                ) As Boolean
On Error GoTo ValidateDates_EH
Dim llIndex As Long

    ValidateDates = False

    If Not (RecordsExist(vadrsEmployeeInformation)) Then
        AppendBusinessError radrsBRV _
            , trErrbusAbsenceDates _
            , mstrCBUSINESSERR _
            , "Employee information does not exist." _
            , MODULENAME & ": ValidateDates"
        Exit Function
    End If

    If Not (RecordsExist(vadrsAbsenceDates)) Then
        AppendBusinessError radrsBRV _
            , trErrbusAbsenceDates _
            , mstrCBUSINESSERR _
            , "Date(s) of occurence was(were) not provided for the Absence." _
            , MODULENAME & ": ValidateDates"
        Exit Function
    End If

    With vadrsAbsenceDates
        .MoveFirst

        For llIndex = 1 To .RecordCount
            If Not (CheckForValidLong(!ABS_NUM)) Then
                AppendBusinessError radrsBRV _
                    , trErrbusAbsenceDates _
                    , mstrCBUSINESSERR _
                    , "Invalid Absence Number." _
                    , MODULENAME & ": ValidateDates"
            End If

            If Not (CheckForValidDate(!ABS_DATE)) Then
                AppendBusinessError radrsBRV _
                    , trErrbusAbsenceDates _
                    , mstrCBUSINESSERR _
                    , "Invalid Absence Date." _
                    , MODULENAME & ": ValidateDates"
            Else
                If (!ABS_DATE <= vadrsEmployeeInformation!LAST_PP_PROC_DATE) Then
                    AppendBusinessError radrsBRV _
                        , trErrbusAbsenceDates _
                        , mstrCBUSINESSERR _
                        , "Invalid Absence Date. Absence Date " & CStr(!ABS_DATE) & " has already been processed by Payroll." _
                        , MODULENAME & ": ValidateDates"
                End If
            End If

            Select Case (Trim(!RECORD_STATUS))
                Case "D"
                    If Not (CheckForValidLong(!ABS_NUM, True)) Then
                        AppendBusinessError radrsBRV _
                            , trErrbusAbsenceDates _
                            , mstrCBUSINESSERR _
                            , "Invalid Absence Number." _
                            , MODULENAME & ": ValidateDates"
                    End If

                    If Not (IsNull(!CREATED_TIMESTAMP)) Then
                        If (Trim(!CREATED_TIMESTAMP) <> "") Then
                            If Not (CheckForValidString(!CREATED_TIMESTAMP, , , True, 26)) Then
                                AppendBusinessError radrsBRV _
                                    , trErrbusAbsenceDates _
                                    , mstrCBUSINESSERR _
                                    , "Invalid Time Stamp." _
                                    , MODULENAME & ": ValidateDates"
                            End If
                        Else
                            AppendBusinessError radrsBRV _
                                , trErrbusAbsenceDates _
                                , mstrCBUSINESSERR _
                                , "Invalid Time Stamp." _
                                , MODULENAME & ": ValidateDates"
                        End If
                    Else
                        AppendBusinessError radrsBRV _
                            , trErrbusAbsenceDates _
                            , mstrCBUSINESSERR _
                            , "Invalid Time Stamp." _
                            , MODULENAME & ": ValidateDates"
                    End If
            End Select

            .MoveNext
        Next llIndex

        .MoveFirst
    End With

    If (RecordsExist(radrsBRV)) Then
        With radrsBRV
            .Filter = adFilterNone
            .Filter = ("ERR_TYPE = 'BE' OR ERR_TYPE = 'AE'" & _
                        " OR ERR_TYPE = 'AW'")
            If Not (.EOF) Then
                .Filter = adFilterNone
                Exit Function
            End If
        End With
    End If

    ValidateDates = True

    Exit Function
ValidateDates_EH:
    TrsRaiseError MODULENAME _
            , "ValidateDates"

End Function

Private Function GetEmployeeSchedule(ByVal vlEmployeeNumber As Long _
                    , ByVal vadrsInitialAbsence As ADODB.Recordset _
                    , ByVal vadrsAbsenceDates As ADODB.Recordset _
                    ) As ADODB.Recordset
On Error GoTo GetEmployeeSchedule_EH
'//Dim lobjSchedule As busSchedule.cbsSchedule
'Used for determining schedule
Dim ldtStartDate As Date, ldtEndDate As Date
'//Dim llRecordIndex As Long

    Set GetEmployeeSchedule = Nothing

    vadrsInitialAbsence.MoveFirst
    ldtStartDate = vadrsInitialAbsence!ABS_ST_DT

    With vadrsAbsenceDates
        .Sort = "ABS_DATE"

        .MoveFirst
        If (ldtStartDate > !ABS_DATE) Then
            ldtStartDate = !ABS_DATE
        End If

        .MoveLast
        ldtEndDate = !ABS_DATE

        .MoveFirst
    End With

    ldtEndDate = DateAdd("m", 3, ldtEndDate) '//provide extra data for extending absence or calculating return to work???
    Set GetEmployeeSchedule = GetEmployeeSchedule_Private(vlEmployeeNumber, ldtStartDate, ldtEndDate)
    
    Exit Function
GetEmployeeSchedule_EH:
'    Set lobjSchedule = Nothing

    TrsRaiseError MODULENAME _
        , "GetEmployeeSchedule"

End Function

'//for testing
'//TODO: change Public to Private
Public Function GetEmployeeSchedule2(ByVal vlEmployeeNumber As Long _
                    , ByVal vdtStartDate As Date _
                    , ByVal vdtEndDate As Date _
                    ) As ADODB.Recordset
On Error GoTo eh
    Set GetEmployeeSchedule2 = GetEmployeeSchedule_Private(vlEmployeeNumber, vdtStartDate, vdtEndDate)
    
    Exit Function
eh:
    TrsRaiseError MODULENAME, "GetEmployeeSchedule2"
End Function

Private Function GetEmployeeSchedule_Private(ByVal vlEmployeeNumber As Long _
                    , ByVal vdtStartDate As Date _
                    , ByVal vdtEndDate As Date _
                    ) As ADODB.Recordset
On Error GoTo eh
Dim lobjSchedule As busSchedule.cbsSchedule
'Used for determining schedule
'//Dim ldtStartDate As Date, ldtEndDate As Date
Dim llRecordIndex As Long

    Set GetEmployeeSchedule_Private = Nothing

    Set lobjSchedule = CreateObject("busSchedule.cbsSchedule")
    Set GetEmployeeSchedule_Private = lobjSchedule.GetAdjustedBaseSchedule( _
                            vlEmployeeNumber _
                            , vdtStartDate _
                            , vdtEndDate _
                            , madrsScheduleWithNoExceptions _
                            , True _
                            , trAdjBaseSchedWithExceptions _
                            , trAllScheduleLevels _
                            , trRelief _
                            , trAllExceptions)
    Set lobjSchedule = Nothing

    If (RecordsExist(GetEmployeeSchedule_Private)) Then
        If (GetEmployeeSchedule_Private.RecordCount > 0) Then
            GetEmployeeSchedule_Private.Filter = adFilterNone
            GetEmployeeSchedule_Private.Filter = "SHIFT_END_TIME = '24:00'" & _
                                    " OR SHIFT_LUNCH_END_TIME = '24:00'" & _
                                    " OR END_TIME = '24:00'"
            If Not (GetEmployeeSchedule_Private.EOF) Then
                GetEmployeeSchedule_Private.MoveFirst

                For llRecordIndex = 1 To GetEmployeeSchedule_Private.RecordCount
                    If (GetEmployeeSchedule_Private!SHIFT_END_TIME = "24:00") Then
                        GetEmployeeSchedule_Private!SHIFT_END_TIME = "00:00"
                        GetEmployeeSchedule_Private.Update
                    End If
                    If (GetEmployeeSchedule_Private!SHIFT_LUNCH_END_TIME = "24:00") Then
                        GetEmployeeSchedule_Private!SHIFT_LUNCH_END_TIME = "00:00"
                        GetEmployeeSchedule_Private.Update
                    End If
                    If (GetEmployeeSchedule_Private!END_TIME = "24:00") Then
                        GetEmployeeSchedule_Private!END_TIME = "00:00"
                        GetEmployeeSchedule_Private.Update
                    End If

                    GetEmployeeSchedule_Private.MoveNext
                Next llRecordIndex
            End If

            GetEmployeeSchedule_Private.Filter = adFilterNone
            GetEmployeeSchedule_Private.MoveFirst
        End If
    End If

    Exit Function
eh:
    Set lobjSchedule = Nothing

    TrsRaiseError MODULENAME _
        , "GetEmployeeSchedule_Private"

End Function

Private Function GetEmployeeScheduleTest(ByVal vlEmployeeNumber As Long _
                    , ByVal vadrsInitialAbsence As ADODB.Recordset _
                    , ByVal vadrsAbsenceDates As ADODB.Recordset _
                    ) As ADODB.Recordset
On Error GoTo eh
Dim lobjSchedule As busSchedule.cbsSchedule
'Used for determining schedule
Dim ldtStartDate As Date, ldtEndDate As Date
Dim llRecordIndex As Long

    Set GetEmployeeScheduleTest = Nothing

    vadrsInitialAbsence.MoveFirst
    ldtStartDate = vadrsInitialAbsence!ABS_ST_DT

    With vadrsAbsenceDates
        .Sort = "ABS_DATE"

        .MoveFirst
        If (ldtStartDate > !ABS_DATE) Then
            ldtStartDate = !ABS_DATE
        End If

        .MoveLast
        ldtEndDate = !ABS_DATE

        .MoveFirst
    End With

    Set lobjSchedule = CreateObject("busSchedule.cbsSchedule")
    Set GetEmployeeScheduleTest = lobjSchedule.GetAdjustedBaseSchedule( _
                            vlEmployeeNumber _
                            , ldtStartDate _
                            , DateAdd("m", 3, ldtEndDate) _
                            , madrsScheduleWithNoExceptions _
                            , True _
                            , trAdjBaseSchedWithExceptions _
                            , trAllScheduleLevels _
                            , trRelief _
                            , trAllExceptions)
    Set lobjSchedule = Nothing

    If (RecordsExist(GetEmployeeScheduleTest)) Then
        If (GetEmployeeScheduleTest.RecordCount > 0) Then
            GetEmployeeScheduleTest.Filter = adFilterNone
            GetEmployeeScheduleTest.Filter = "SHIFT_END_TIME = '24:00'" & _
                                    " OR SHIFT_LUNCH_END_TIME = '24:00'" & _
                                    " OR END_TIME = '24:00'"
            If Not (GetEmployeeScheduleTest.EOF) Then
                GetEmployeeScheduleTest.MoveFirst

                For llRecordIndex = 1 To GetEmployeeScheduleTest.RecordCount
                    If (GetEmployeeScheduleTest!SHIFT_END_TIME = "24:00") Then
                        GetEmployeeScheduleTest!SHIFT_END_TIME = "00:00"
                        GetEmployeeScheduleTest.Update
                    End If
                    If (GetEmployeeScheduleTest!SHIFT_LUNCH_END_TIME = "24:00") Then
                        GetEmployeeScheduleTest!SHIFT_LUNCH_END_TIME = "00:00"
                        GetEmployeeScheduleTest.Update
                    End If
                    If (GetEmployeeScheduleTest!END_TIME = "24:00") Then
                        GetEmployeeScheduleTest!END_TIME = "00:00"
                        GetEmployeeScheduleTest.Update
                    End If

                    GetEmployeeScheduleTest.MoveNext
                Next llRecordIndex
            End If

            GetEmployeeScheduleTest.Filter = adFilterNone
            GetEmployeeScheduleTest.MoveFirst

            '//try to remove scheduled exceptions added by the initial absence
            Do While Not GetEmployeeScheduleTest.EOF
                If Trim(GetEmployeeScheduleTest!EXCEPTION_CD) = "E" Then
                    GetEmployeeScheduleTest!EXCEPTION_CD = ""
                    GetEmployeeScheduleTest.Update
                End If

                GetEmployeeScheduleTest.MoveNext
            Loop
        End If
    End If

    Exit Function
eh:
    Set lobjSchedule = Nothing

    TrsRaiseError MODULENAME, "GetEmployeeScheduleTest"
End Function

Private Function GetEmployeeTimeDetails(ByVal vlEmployeeNumber As Long _
                    , ByVal vadrsInitialAbsence As ADODB.Recordset _
                    , ByVal vadrsAbsenceDates As ADODB.Recordset _
                    , Optional ByVal vblnIncludeRTWDate As Boolean = False _
                    , Optional ByVal vblnUseTimeForSchedule As Boolean = False _
                    ) As ADODB.Recordset
On Error GoTo GetEmployeeTimeDetails_EH
Dim lobjTime As busSchedException.cbsSchedExc
'Used for determining schedule
Dim ldtStartDate As Date, ldtEndDate As Date

    Set GetEmployeeTimeDetails = Nothing

    vadrsInitialAbsence.MoveFirst
    ldtStartDate = vadrsInitialAbsence!ABS_ST_DT

    With vadrsAbsenceDates
        .Sort = "ABS_DATE"

        .MoveFirst
        '//ldtStartDate = MinDate(vadrsInitialAbsence!ABS_ST_DT, vadrsAbsenceDates!ABS_DATE)
        If (ldtStartDate > !ABS_DATE) Then
            ldtStartDate = !ABS_DATE
        End If

        .MoveLast
        ldtEndDate = !ABS_DATE

        .MoveFirst
    End With
    
    '//need time for return to work date in order to determine date worked
    If vblnIncludeRTWDate Then
        ldtEndDate = MaxDate(ldtEndDate, vadrsInitialAbsence!RTN_TO_WRK_DT)
    End If
    
    If vblnUseTimeForSchedule Then
        ldtEndDate = DateAdd("m", 3, ldtEndDate)
    End If
    
    Set GetEmployeeTimeDetails = GetEmployeeTimeDetails_Private(vlEmployeeNumber, ldtStartDate, ldtEndDate)

    Exit Function
GetEmployeeTimeDetails_EH:
    Set lobjTime = Nothing

    TrsRaiseError MODULENAME _
        , "GetEmployeeTimeDetails"

End Function

'//for testing
'//TODO: change Public to Private
Public Function GetEmployeeTimeDetails2(ByVal vlEmployeeNumber As Long _
                    , ByVal vdtStartDate As Date _
                    , ByVal vdtEndDate As Date _
                    ) As ADODB.Recordset

On Error GoTo eh
   
    Set GetEmployeeTimeDetails2 = GetEmployeeTimeDetails_Private(vlEmployeeNumber, vdtStartDate, vdtEndDate)
    Exit Function
eh:
    TrsRaiseError MODULENAME, "GetEmployeeTimeDetails2"
End Function

Private Function GetEmployeeTimeDetails_Private(ByVal vlEmployeeNumber As Long _
                    , ByVal vdtStartDate As Date _
                    , ByVal vdtEndDate As Date _
                    ) As ADODB.Recordset
On Error GoTo eh
Dim lobjTime As busSchedException.cbsSchedExc
'Used for determining schedule
'//Dim ldtStartDate As Date, ldtEndDate As Date

    Set GetEmployeeTimeDetails_Private = Nothing
   
    Set lobjTime = CreateObject("busSchedException.cbsSchedExc")
    Set GetEmployeeTimeDetails_Private = lobjTime.GetViewCalendarDetails( _
                            vlEmployeeNumber _
                            , vdtStartDate _
                            , vdtEndDate)
    Set lobjTime = Nothing

    If (RecordsExist(GetEmployeeTimeDetails_Private)) Then
        If (GetEmployeeTimeDetails_Private.RecordCount > 0) Then
            GetEmployeeTimeDetails_Private.MoveFirst
        End If
    End If

    Exit Function
eh:
    Set lobjTime = Nothing

    TrsRaiseError MODULENAME _
        , "GetEmployeeTimeDetails_Private"

End Function

Private Function GetEmployeeTimeDetailsTest(ByVal vlEmployeeNumber As Long _
                    , ByVal vadrsInitialAbsence As ADODB.Recordset _
                    , ByVal vadrsAbsenceDates As ADODB.Recordset _
                    ) As ADODB.Recordset
On Error GoTo eh
    Dim lobjTime As busSchedException.cbsSchedExc
    Dim lobjSched As busSchedule.cbsSchedule
    'Used for determining schedule
    Dim ldtStartDate As Date, ldtEndDate As Date
    Dim rsScheduleWithoutExceptions As ADODB.Recordset
    Dim rsViewCalendarDetails As ADODB.Recordset
    Dim rsGetAdjustedBaseSchedule As ADODB.Recordset
    Dim ldtExtendedStartDate As Date, ldtExtendedEndDate As Date

    Set GetEmployeeTimeDetailsTest = Nothing

    vadrsInitialAbsence.MoveFirst
    ldtStartDate = vadrsInitialAbsence!ABS_ST_DT

    With vadrsAbsenceDates
        .Sort = "ABS_DATE"

        .MoveFirst
        If (ldtStartDate > !ABS_DATE) Then
            ldtStartDate = !ABS_DATE
        End If
        ldtExtendedStartDate = !ABS_DATE

        .MoveLast
        ldtEndDate = !ABS_DATE
        ldtExtendedEndDate = !ABS_DATE

        .MoveFirst
    End With

    Set lobjTime = CreateObject("busSchedException.cbsSchedExc")
    Set rsViewCalendarDetails = lobjTime.GetViewCalendarDetails(vlEmployeeNumber, ldtStartDate, ldtEndDate)
    Set lobjTime = Nothing

    '//try to exclude actual time and scheduled exceptions that have already been added by the absence
    '//this will allow a re-test to better mimic the data as it was for the initial absence
    Set lobjSched = CreateObject("busSchedule.cbsSchedule")
    Set rsGetAdjustedBaseSchedule = lobjSched.GetAdjustedBaseSchedule(vlEmployeeNumber, ldtStartDate, ldtEndDate, rsScheduleWithoutExceptions, False, trAdjBaseSchedWithExceptions, trAllScheduleLevels)
    Set lobjSched = Nothing

    '//scheduled
    rsViewCalendarDetails.Filter = "TIME_TYPE = '3'"
    Set GetEmployeeTimeDetailsTest = CopyRecordset(rsViewCalendarDetails)

    If RecordsExist(GetEmployeeTimeDetailsTest) Then
        GetEmployeeTimeDetailsTest.MoveFirst
        Do While Not GetEmployeeTimeDetailsTest.EOF
            If Trim(GetEmployeeTimeDetailsTest!EXCEPTION_CD) = "E" Then
                GetEmployeeTimeDetailsTest!EXCEPTION_CD = Space(3)
            End If
            GetEmployeeTimeDetailsTest!JOB_NUM = ""
            GetEmployeeTimeDetailsTest!FMLA_IND = ""

            GetEmployeeTimeDetailsTest.Update
            GetEmployeeTimeDetailsTest.MoveNext
        Loop
    End If
    '//actual time not associated with absence
    'rsViewCalendarDetails.Filter = StringFormat("(TIME_TYPE = '2' AND (CHARGE_DATE < #{0}# OR CHARGE_DATE > #{1}#))", ldtStartDate, ldtEndDate)
    rsViewCalendarDetails.Filter = StringFormat("(TIME_TYPE = '2' AND CHARGE_DATE < #{0}#) OR " & _
                                                "(TIME_TYPE = '2' AND CHARGE_DATE > #{1}#)", ldtExtendedStartDate, ldtExtendedEndDate)
    AppendRecordset rsViewCalendarDetails, GetEmployeeTimeDetailsTest
    rsViewCalendarDetails.Filter = adFilterNone

'    If RecordsExist(GetEmployeeTimeDetailsTest) Then
'        GetEmployeeTimeDetailsTest.MoveFirst
'        Do While Not GetEmployeeTimeDetailsTest.EOF
'Debug.Print StringFormat("GetEmployeeTimeDetailsTest: {0} {1} {2}", _
'                        GetEmployeeTimeDetailsTest!EMP_NUM, GetEmployeeTimeDetailsTest!CHARGE_DATE, GetEmployeeTimeDetailsTest!START_TIME)
'            If IsNullOrEmpty(GetEmployeeTimeDetailsTest!JOB_NUM) Then
'                GetEmployeeTimeDetailsTest!JOB_NUM = ""
'            End If
'            GetEmployeeTimeDetailsTest.Update
'            GetEmployeeTimeDetailsTest.MoveNext
'        Loop
'    End If

    'CopyAll with FieldsExistCheck
    Dim j As Integer, i As Integer
    Dim sColumnName As String

    rsGetAdjustedBaseSchedule.Filter = StringFormat("CHARGE_DATE >= #{0}# AND CHARGE_DATE <= #{1}#", ldtExtendedStartDate, ldtExtendedEndDate)
    If RecordsExist(rsGetAdjustedBaseSchedule) Then
        rsGetAdjustedBaseSchedule.MoveFirst
'        GetEmployeeTimeDetailsTest.Open
        For j = 1 To rsGetAdjustedBaseSchedule.RecordCount
'Debug.Print StringFormat("rsGetAdjustedBaseSchedule: {0} {1} {2}", _
'                        rsGetAdjustedBaseSchedule!EMP_NUM, rsGetAdjustedBaseSchedule!CHARGE_DATE, rsGetAdjustedBaseSchedule!START_TIME)
            '//only need to convert TIME_TYPE = 2 (actual time), so skip if it has already been added
            GetEmployeeTimeDetailsTest.Filter = StringFormat("EMP_NUM = {0} AND CHARGE_DATE = {1} AND START_TIME = '{2}'", _
                    rsGetAdjustedBaseSchedule!EMP_NUM, rsGetAdjustedBaseSchedule!CHARGE_DATE, rsGetAdjustedBaseSchedule!START_TIME)
            If Not RecordsExist(GetEmployeeTimeDetailsTest) Then
                GetEmployeeTimeDetailsTest.AddNew

                For i = 1 To rsGetAdjustedBaseSchedule.Fields.Count
                    sColumnName = rsGetAdjustedBaseSchedule.Fields(i - 1).Name
                    If FieldExists(sColumnName, GetEmployeeTimeDetailsTest) Then
                        GetEmployeeTimeDetailsTest.Fields(sColumnName).Value = rsGetAdjustedBaseSchedule.Fields(i - 1).Value
                    End If
                Next
'                If rsGetAdjustedBaseSchedule.Fields("EXCEPTION_CD").Value = "E" Then
'                    GetEmployeeTimeDetailsTest.Fields("EXCEPTION_CD").Value = ""
'                End If

                If GetEmployeeTimeDetailsTest.Fields("END_TIME").Value = "24:00" Then
                    GetEmployeeTimeDetailsTest.Fields("END_TIME").Value = "00:00"
                End If

                If rsGetAdjustedBaseSchedule!SYSTEM_EXC_CD = "TWS" Or rsGetAdjustedBaseSchedule!SYSTEM_EXC_CD = "TSR" Then
                    GetEmployeeTimeDetailsTest!TRADE_EXC_CD = rsGetAdjustedBaseSchedule!SYSTEM_EXC_CD
                End If
                '2 is for compare only
                GetEmployeeTimeDetailsTest.Fields("TIME_TYPE") = 3
                GetEmployeeTimeDetailsTest!DISTRIBUTION_ID = Space(39)
                GetEmployeeTimeDetailsTest!DISTRIBUTION_DESC = ""
                GetEmployeeTimeDetailsTest!WG_NUM = 0
'                GetEmployeeTimeDetailsTest!JOB_NUM = ""
                GetEmployeeTimeDetailsTest!DISTR_DEPT_NUM = ""
                GetEmployeeTimeDetailsTest!DEPT_NUM = ""
                GetEmployeeTimeDetailsTest!OT_REASON_CD = ""
                GetEmployeeTimeDetailsTest!TRAVEL_PAY_IND = "N"
                GetEmployeeTimeDetailsTest!INDEX_TIME = Space(26)
                GetEmployeeTimeDetailsTest!APPR_DATE = #1/1/100#
                GetEmployeeTimeDetailsTest!OT_BANK_CHARGE_DATE = #1/1/100#
                GetEmployeeTimeDetailsTest!ADJ_TYPE = ""
                GetEmployeeTimeDetailsTest!REQ_DATE = #1/1/100#
                GetEmployeeTimeDetailsTest!COMPLETED_DATE = #1/1/100#
                GetEmployeeTimeDetailsTest!FINAL_APPR_DATE = #1/1/100#
                GetEmployeeTimeDetailsTest!SPECIAL_CHECK = ""
                GetEmployeeTimeDetailsTest!ADJ_PREV_PAID_FL = ""

                '//set the defaults for values that might not be in the GAB
                rsViewCalendarDetails.Filter = StringFormat("EMP_NUM = {0} AND CHARGE_DATE = #{1}# AND START_TIME = '{2}' AND TIME_TYPE = '2'", _
                        rsGetAdjustedBaseSchedule!EMP_NUM, rsGetAdjustedBaseSchedule!CHARGE_DATE, rsGetAdjustedBaseSchedule!START_TIME)
                If RecordsExist(rsViewCalendarDetails) Then
'Debug.Print StringFormat("   rsViewCalendarDetails: {0} {1} {2} {3}", _
'                        rsViewCalendarDetails!EMP_NUM, rsViewCalendarDetails!CHARGE_DATE, rsViewCalendarDetails!START_TIME, _
'                        rsViewCalendarDetails!EXCEPTION_CD)

'                    If IsNullOrEmpty(GetEmployeeTimeDetailsTest!DISTRIBUTION_ID) Then
'                        GetEmployeeTimeDetailsTest!DISTRIBUTION_ID = rsViewCalendarDetails!DISTRIBUTION_ID
'                    End If
'                    If IsNullOrEmpty(GetEmployeeTimeDetailsTest!DISTRIBUTION_DESC) Then
'                        GetEmployeeTimeDetailsTest!DISTRIBUTION_DESC = rsViewCalendarDetails!DISTRIBUTION_DESC
'                    End If
'                    If IsNullOrEmpty(GetEmployeeTimeDetailsTest!EXCEPTION_CD) Then
'                        GetEmployeeTimeDetailsTest!EXCEPTION_CD = rsViewCalendarDetails!EXCEPTION_CD
'                    End If
'                    If IsNullOrEmpty(GetEmployeeTimeDetailsTest!JOB_NUM) Then
'                        GetEmployeeTimeDetailsTest!JOB_NUM = rsViewCalendarDetails!JOB_NUM
'                    End If
'                    If GetEmployeeTimeDetailsTest!WG_NUM = 0 Then
'                        GetEmployeeTimeDetailsTest!WG_NUM = rsViewCalendarDetails!WG_NUM
'                    End If
'                    If IsNullOrEmpty(GetEmployeeTimeDetailsTest!FMLA_IND) Then
'                        GetEmployeeTimeDetailsTest!FMLA_IND = rsViewCalendarDetails!FMLA_IND
'                    End If
'                    If IsNullOrEmpty(GetEmployeeTimeDetailsTest!DISTR_DEPT_NUM) Then
'                        GetEmployeeTimeDetailsTest!DISTR_DEPT_NUM = rsViewCalendarDetails!DISTR_DEPT_NUM
'                    End If
'                    If IsNullOrEmpty(GetEmployeeTimeDetailsTest!DEPT_NUM) Then
'                        GetEmployeeTimeDetailsTest!DEPT_NUM = rsViewCalendarDetails!DEPT_NUM
'                    End If
'                    If GetEmployeeTimeDetailsTest!LOC_CD = 0 Then
'                        GetEmployeeTimeDetailsTest!LOC_CD = rsViewCalendarDetails!LOC_CD
'                    End If
'                    If IsNullOrEmpty(GetEmployeeTimeDetailsTest!OT_REASON_CD) Then
'                        GetEmployeeTimeDetailsTest!OT_REASON_CD = rsViewCalendarDetails!OT_REASON_CD
'                    End If
'                    If IsNullOrEmpty(GetEmployeeTimeDetailsTest!TRAVEL_PAY_IND) Then
'                        GetEmployeeTimeDetailsTest!TRAVEL_PAY_IND = rsViewCalendarDetails!TRAVEL_PAY_IND
'                    End If
                    If GetEmployeeTimeDetailsTest!APPR_DATE = #12:00:00 AM# Then
                        GetEmployeeTimeDetailsTest!APPR_DATE = rsViewCalendarDetails!APPR_DATE
                    End If
                    If GetEmployeeTimeDetailsTest!OT_BANK_CHARGE_DATE = #12:00:00 AM# Then
                        GetEmployeeTimeDetailsTest!OT_BANK_CHARGE_DATE = rsViewCalendarDetails!OT_BANK_CHARGE_DATE
                    End If
                    If IsNullOrEmpty(GetEmployeeTimeDetailsTest!REQ_LINE_IND) Then
                        GetEmployeeTimeDetailsTest!REQ_LINE_IND = rsViewCalendarDetails!REQ_LINE_IND
                    End If
'                    GetEmployeeTimeDetailsTest!SHIFT_NUM = rsViewCalendarDetails!SHIFT_NUM
'                    GetEmployeeTimeDetailsTest!SHIFT_CODE = rsViewCalendarDetails!SHIFT_CODE
'                    GetEmployeeTimeDetailsTest!SHIFT_HOURS = rsViewCalendarDetails!SHIFT_HOURS
'                    GetEmployeeTimeDetailsTest!COMP_SHIFT_CD = rsViewCalendarDetails!COMP_SHIFT_CD

                    If IsNullOrEmpty(GetEmployeeTimeDetailsTest!ADJ_TYPE) Then
                        GetEmployeeTimeDetailsTest!ADJ_TYPE = rsViewCalendarDetails!ADJ_TYPE
                    End If
                    If GetEmployeeTimeDetailsTest!REQ_DATE = #12:00:00 AM# Then
                        GetEmployeeTimeDetailsTest!REQ_DATE = rsViewCalendarDetails!REQ_DATE
                    End If
                    If GetEmployeeTimeDetailsTest!COMPLETED_DATE = #12:00:00 AM# Then
                        GetEmployeeTimeDetailsTest!COMPLETED_DATE = rsViewCalendarDetails!COMPLETED_DATE
                    End If
                    If GetEmployeeTimeDetailsTest!FINAL_APPR_DATE = #12:00:00 AM# Then
                        GetEmployeeTimeDetailsTest!FINAL_APPR_DATE = rsViewCalendarDetails!FINAL_APPR_DATE
                    End If
                    If IsNullOrEmpty(GetEmployeeTimeDetailsTest!SPECIAL_CHECK) Then
                        GetEmployeeTimeDetailsTest!SPECIAL_CHECK = rsViewCalendarDetails!SPECIAL_CHECK
                    End If
                    If IsNullOrEmpty(GetEmployeeTimeDetailsTest!ADJ_PREV_PAID_FL) Then
                        GetEmployeeTimeDetailsTest!ADJ_PREV_PAID_FL = rsViewCalendarDetails!ADJ_PREV_PAID_FL
                    End If
'                    If GetEmployeeTimeDetailsTest!LAST_MOD_EMP_NUM = 0 Then
'                        GetEmployeeTimeDetailsTest!LAST_MOD_EMP_NUM = rsViewCalendarDetails!LAST_MOD_EMP_NUM
'                    End If
'                    If IsNullOrEmpty(GetEmployeeTimeDetailsTest!INDEX_TIME) Then
'                        GetEmployeeTimeDetailsTest!INDEX_TIME = rsViewCalendarDetails!INDEX_TIME
'                    End If
'                    If IsNullOrEmpty(GetEmployeeTimeDetailsTest!TRADE_EXC_CD) Then
'                        GetEmployeeTimeDetailsTest!TRADE_EXC_CD = rsViewCalendarDetails!TRADE_EXC_CD
'                    End If
                End If

                GetEmployeeTimeDetailsTest.Update
            End If
            GetEmployeeTimeDetailsTest.Filter = adFilterNone

            rsGetAdjustedBaseSchedule.MoveNext
        Next
    End If
    rsViewCalendarDetails.Filter = adFilterNone
    GetEmployeeTimeDetailsTest.Filter = adFilterNone
    If RecordsExist(GetEmployeeTimeDetailsTest) Then
        If GetEmployeeTimeDetailsTest.RecordCount > 0 Then
            GetEmployeeTimeDetailsTest.MoveFirst
        End If
    End If

    Exit Function
eh:
    Set lobjTime = Nothing
    Set lobjSched = Nothing

    TrsRaiseError MODULENAME, "GetEmployeeTimeDetailsTest"
End Function

'//the new
Private Function MergeScheduleWithTime(ByVal vlEmployeeNumber As Long, ByVal vadrsInitialAbsence As ADODB.Recordset, _
                                        ByVal vadrsAbsenceDates As ADODB.Recordset, ByVal vadrsScheduledExceptions As ADODB.Recordset) As ADODB.Recordset
    Dim ladrsSchedule As ADODB.Recordset
    Dim ladrsTimeForSchedule As ADODB.Recordset
    
On Error GoTo eh

    Set ladrsSchedule = GetEmployeeSchedule(vlEmployeeNumber, vadrsInitialAbsence, vadrsAbsenceDates)
        
    '//if absence is unextended, sched exc from db may prevent the probable RTW date from being determined properly
    '//try removing deleted exceptions from the schedule, this will allow the RTW to be sooner
    ApplyPendingChanges vadrsScheduledExceptions, ladrsSchedule, vadrsInitialAbsence, vadrsAbsenceDates
    
    Set ladrsTimeForSchedule = GetEmployeeTimeDetails(vlEmployeeNumber, vadrsInitialAbsence, vadrsAbsenceDates, True, True)
    ApplyPendingChanges vadrsScheduledExceptions, ladrsTimeForSchedule, vadrsInitialAbsence, vadrsAbsenceDates

    '//Refactor: MergeScheduleWithTime ladrsSchedule, ladrsTimeForSchedule
    '//Issue: The schedule may be missing on the day of a partial time entry (time 00:00-02:00, sched 00:00-06:00, missing 02:00-06:00)
    '//Merge partial schedule that is not on the GetEmployeeTimeDetails (view calendar)
    Set MergeScheduleWithTime = MergeScheduleWithTime_Private(vlEmployeeNumber, ladrsSchedule, ladrsTimeForSchedule)
    
'                        Dim blnAddTimeForSchedule As Boolean
'                        If ladrsSchedule.Fields("SYSTEM_EXC_CD").Value = "SLT" Then
'                            '//need lunch because the time is split (for actual time), otherwise lunch could cause a split
'                            '//scheduled exceptions are not yet split, so lunch is not needed
'                            vadrsScheduledExceptions.Filter = "CHARGE_DATE = " & ladrsSchedule.Fields("CHARGE_DATE").Value & " AND START_TIME = " & ladrsSchedule.Fields("SHIFT_START_TIME").Value & " AND END_TIME = " & ladrsSchedule.Fields("SHIFT_END_TIME").Value
'                            blnAddTimeForSchedule = vadrsScheduledExceptions.EOF  '//if found on sched except then do not add, if not found then add
'                            vadrsScheduledExceptions.Filter = adFilterNone
'                        Else
'                            blnAddTimeForSchedule = True
'                        End If
  
    vadrsScheduledExceptions.Filter = adFilterNone
    Set ladrsTimeForSchedule = Nothing
    Set ladrsSchedule = Nothing
    Exit Function
eh:
    Set ladrsTimeForSchedule = Nothing
    vadrsScheduledExceptions.Filter = adFilterNone
    Set ladrsSchedule = Nothing
    
    TrsRaiseError MODULENAME, "MergeScheduleWithTime"
End Function

'//TODO: TEST ApplyPendingChanges outside of this
'//TODO: change Public to Private
Public Function MergeScheduleWithTime2(ByVal vlEmployeeNumber As Long, ByVal vdtStartDate As Date, ByVal vdtEndDate As Date) As ADODB.Recordset
    Dim ladrsSchedule As ADODB.Recordset
    Dim ladrsTimeForSchedule As ADODB.Recordset
    
On Error GoTo eh

    Set ladrsSchedule = GetEmployeeSchedule2(vlEmployeeNumber, vdtStartDate, vdtEndDate)
        
    '//if absence is unextended, sched exc from db may prevent the probable RTW date from being determined properly
    '//try removing deleted exceptions from the schedule, this will allow the RTW to be sooner
    '//TODO:
    '//ApplyPendingChanges vadrsScheduledExceptions, ladrsSchedule, vadrsInitialAbsence, vadrsAbsenceDates
    
    Set ladrsTimeForSchedule = GetEmployeeTimeDetails2(vlEmployeeNumber, vdtStartDate, vdtEndDate)
    '//TODO:
    '//ApplyPendingChanges vadrsScheduledExceptions, ladrsTimeForSchedule, vadrsInitialAbsence, vadrsAbsenceDates

    '//Refactor: MergeScheduleWithTime ladrsSchedule, ladrsTimeForSchedule
    '//Issue: The schedule may be missing on the day of a partial time entry (time 00:00-02:00, sched 00:00-06:00, missing 02:00-06:00)
    '//Merge partial schedule that is not on the GetEmployeeTimeDetails (view calendar)
    Set MergeScheduleWithTime2 = MergeScheduleWithTime_Private(vlEmployeeNumber, ladrsSchedule, ladrsTimeForSchedule)
    
'    vadrsScheduledExceptions.Filter = adFilterNone
    Set ladrsTimeForSchedule = Nothing
    Set ladrsSchedule = Nothing
    Exit Function
eh:
    Set ladrsTimeForSchedule = Nothing
'    vadrsScheduledExceptions.Filter = adFilterNone
    Set ladrsSchedule = Nothing
    
    TrsRaiseError MODULENAME, "MergeScheduleWithTime2"
End Function

'//TODO: resolve vadrsScheduledExceptions for SLT ByVal vadrsScheduledExceptions As ADODB.Recordset
Private Function MergeScheduleWithTime_Private(ByVal vlEmployeeNumber As Long _
                                            , ByVal vadrsSchedule As ADODB.Recordset _
                                            , ByVal vadrsTimeForSchedule As ADODB.Recordset) As ADODB.Recordset
'    Dim ladrsSchedule As ADODB.Recordset
'    Dim ladrsTimeForSchedule As ADODB.Recordset
    
On Error GoTo eh

    '//Refactor: MergeScheduleWithTime ladrsSchedule, ladrsTimeForSchedule
    '//Issue: The schedule may be missing on the day of a partial time entry (time 00:00-02:00, sched 00:00-06:00, missing 02:00-06:00)
    '//Merge partial schedule that is not on the GetEmployeeTimeDetails (view calendar)
    If RecordsExist(vadrsTimeForSchedule) Then
        Dim rsTimeSummary As ADODB.Recordset
        Dim sGroupStartDate As Date, sGroupEndDate As Date, sGroupStartTime As String, sGroupEndTime As String, dtGroupEndDateTime As Date
        Dim dtDateWorked As Date, sStartTime As String, sEndTime As String
        Dim blnFirstGroup As Boolean
        
        '//Set rsTimeSummary = CreateScheduleInGAP()  '//time range split over midnight (charge date, dateworked, start time, end time)
        Set rsTimeSummary = CreateGAP() '//time range (start date, start time, end date, end time)
        
        vadrsTimeForSchedule.MoveFirst
        blnFirstGroup = True
        
        Do While Not vadrsTimeForSchedule.EOF
            If vadrsTimeForSchedule.Fields("HOURS").Value > 0 Then
                '//ready for 1st group
                If blnFirstGroup Then
                    sGroupStartDate = vadrsTimeForSchedule.Fields("DATE_WORKED").Value
                    sGroupStartTime = vadrsTimeForSchedule.Fields("START_TIME").Value
                    'dtGroupEndDateTime = DateTimeFrom(vadrsTimeForSchedule.Fields("DATE_WORKED").Value, vadrsTimeForSchedule.Fields("END_TIME").Value)
                    sGroupEndDate = vadrsTimeForSchedule.Fields("DATE_WORKED").Value 'MidnightDate(dtGroupEndDateTime)
                    sGroupEndTime = vadrsTimeForSchedule.Fields("END_TIME").Value 'FormatTime(dtGroupEndDateTime)
                End If
                
                dtDateWorked = vadrsTimeForSchedule.Fields("DATE_WORKED").Value
                sStartTime = vadrsTimeForSchedule.Fields("START_TIME").Value
                sEndTime = vadrsTimeForSchedule.Fields("END_TIME").Value

                '//first in group
'                            If sGroupStartTime = "" Then
'                                sGroupChargeDate = vadrsTimeForSchedule.Fields("CHARGE_DATE").Value
'                            Else
                If Not blnFirstGroup Then
                    '//check for contiguous
                    '//TEST:
                    If DateTimeFrom(dtDateWorked, sStartTime) = DateTimeFrom(sGroupEndDate, sGroupEndTime, True) Then
                        'dtGroupEndDateTime = DateTimeFrom(vadrsTimeForSchedule.Fields("DATE_WORKED").Value, vadrsTimeForSchedule.Fields("END_TIME").Value)
                        sGroupEndDate = vadrsTimeForSchedule.Fields("DATE_WORKED").Value 'MidnightDate(dtGroupEndDateTime)
                        sGroupEndTime = vadrsTimeForSchedule.Fields("END_TIME").Value 'FormatTime(dtGroupEndDateTime)
                        'If sGroupEndTime = "00:00" Then
                        '    sGroupEndDate = DateAdd("d", 1, sGroupEndDate)
                        'End If
                    Else '//flush group (previous rows)
                        rsTimeSummary.AddNew
                        rsTimeSummary.Fields("GAP_START_DATE").Value = sGroupStartDate '//the actual date
                        rsTimeSummary.Fields("GAP_START_TIME").Value = sGroupStartTime
                        rsTimeSummary.Fields("GAP_END_DATE").Value = sGroupEndDate
                        rsTimeSummary.Fields("GAP_END_TIME").Value = sGroupEndTime
                        rsTimeSummary.Update
                        
                        '//ready for next group
                        sGroupStartDate = dtDateWorked
                        sGroupStartTime = sStartTime
                        'dtGroupEndDateTime = DateTimeFrom(dtDateWorked, sEndTime)
                        sGroupEndDate = dtDateWorked 'MidnightDate(dtGroupEndDateTime)
                        sGroupEndTime = sEndTime 'FormatTime(dtGroupEndDateTime)
                    End If
                Else
                    blnFirstGroup = False
                End If
'                           End If
            End If
            vadrsTimeForSchedule.MoveNext
        Loop
        
        '//final flush group (previous rows)
        If Trim(sGroupStartTime) <> "" Then
            rsTimeSummary.AddNew
            rsTimeSummary.Fields("GAP_START_DATE").Value = sGroupStartDate '//the actual date
            rsTimeSummary.Fields("GAP_START_TIME").Value = sGroupStartTime
            rsTimeSummary.Fields("GAP_END_DATE").Value = sGroupEndDate
            rsTimeSummary.Fields("GAP_END_TIME").Value = sGroupEndTime
            rsTimeSummary.Update
        End If

        '//check if schedule exists in Time, if not then create a row with the Gap
        Dim dtScheduleStartDateTime As Date, dtScheduleEndDateTime As Date
        Dim dtTimeSummaryStartDateTime As Date, dtTimeSummaryEndDateTime As Date
        Dim iIntersection As Integer, blnFoundTimeSummary As Boolean
        
        If RecordsExist(vadrsSchedule) Then
            vadrsSchedule.MoveFirst
            Do While Not vadrsSchedule.EOF
                If vadrsSchedule.Fields("HOURS").Value > 0 Then
                    dtScheduleStartDateTime = DateTimeFrom(vadrsSchedule.Fields("DATE_WORKED").Value, vadrsSchedule.Fields("START_TIME").Value)
                    '//TODO: Does this need to be specified as an end time, since it is a dateworked ???
                    dtScheduleEndDateTime = DateTimeFrom(vadrsSchedule.Fields("DATE_WORKED").Value, vadrsSchedule.Fields("END_TIME").Value, True)
                    
'//Debug.Print "dtScheduleStartDateTime = " & dtScheduleStartDateTime & "; dtScheduleEndDateTime = " & dtScheduleEndDateTime
'//If dtScheduleStartDateTime = "11/18/2013 11:30:00 AM" Then
'//Dim s As String
'//s = "stop"
'//End If
                    '//rsTimeSummary.Filter = ""
                    blnFoundTimeSummary = False
                    rsTimeSummary.MoveFirst
                    Do While Not rsTimeSummary.EOF
                        dtTimeSummaryStartDateTime = DateTimeFrom(rsTimeSummary.Fields("GAP_START_DATE").Value, rsTimeSummary.Fields("GAP_START_TIME").Value)
                        dtTimeSummaryEndDateTime = DateTimeFrom(rsTimeSummary.Fields("GAP_END_DATE").Value, rsTimeSummary.Fields("GAP_END_TIME").Value, True)
                    
                        iIntersection = DateIntersection2(dtScheduleStartDateTime, dtScheduleEndDateTime, dtTimeSummaryStartDateTime, dtTimeSummaryEndDateTime)
                        Select Case iIntersection
                        Case 0, 1
                            AppendRecordsetByName vadrsSchedule, vadrsTimeForSchedule, 1, vadrsSchedule.AbsolutePosition
'                                        vadrsTimeForSchedule.Fields("CHARGE_DATE").Value = MidnightDate(dtTimeSummaryStartDateTime)
'                                        vadrsTimeForSchedule.Fields("DATE_WORKED").Value = MidnightDate(dtScheduleStartDateTime)
                            vadrsTimeForSchedule.Fields("START_TIME").Value = FormatTime(dtScheduleStartDateTime)
                            vadrsTimeForSchedule.Fields("END_TIME").Value = FormatTime(dtTimeSummaryStartDateTime)
                            vadrsTimeForSchedule.Fields("TIME_TYPE").Value = 3 '//scheduled
                            vadrsTimeForSchedule.Update
                            blnFoundTimeSummary = True
                            Exit Do
                        Case 2
                            AppendRecordsetByName vadrsSchedule, vadrsTimeForSchedule, 1, vadrsSchedule.AbsolutePosition
'                                        vadrsTimeForSchedule.Fields("CHARGE_DATE").Value = MidnightDate(dtTimeSummaryStartDateTime)
'                                        vadrsTimeForSchedule.Fields("DATE_WORKED").Value = MidnightDate(dtScheduleStartDateTime)
                            vadrsTimeForSchedule.Fields("START_TIME").Value = FormatTime(dtScheduleStartDateTime)
                            vadrsTimeForSchedule.Fields("END_TIME").Value = FormatTime(dtTimeSummaryStartDateTime)
                            vadrsTimeForSchedule.Fields("TIME_TYPE").Value = 3 '//scheduled
                            vadrsTimeForSchedule.Update

                            AppendRecordsetByName vadrsSchedule, vadrsTimeForSchedule, 1, vadrsSchedule.AbsolutePosition
'                                        vadrsTimeForSchedule.Fields("CHARGE_DATE").Value = MidnightDate(dtTimeSummaryStartDateTime)
'                                        vadrsTimeForSchedule.Fields("DATE_WORKED").Value = MidnightDate(dtScheduleStartDateTime)
                            vadrsTimeForSchedule.Fields("START_TIME").Value = FormatTime(dtTimeSummaryEndDateTime)
                            vadrsTimeForSchedule.Fields("END_TIME").Value = FormatTime(dtScheduleEndDateTime)
                            vadrsTimeForSchedule.Fields("TIME_TYPE").Value = 3 '//scheduled
                            vadrsTimeForSchedule.Update
                            blnFoundTimeSummary = True
                            Exit Do
                        Case 3, 4, 6, 7 '//subset
                            blnFoundTimeSummary = True
                            Exit Do
                        Case 5, 8
                            AppendRecordsetByName vadrsSchedule, vadrsTimeForSchedule, 1, vadrsSchedule.AbsolutePosition
'                                        vadrsTimeForSchedule.Fields("CHARGE_DATE").Value = MidnightDate(dtTimeSummaryStartDateTime)
'                                        vadrsTimeForSchedule.Fields("DATE_WORKED").Value = MidnightDate(dtScheduleStartDateTime)
                            vadrsTimeForSchedule.Fields("START_TIME").Value = FormatTime(dtTimeSummaryEndDateTime)
                            vadrsTimeForSchedule.Fields("END_TIME").Value = FormatTime(dtScheduleEndDateTime)
                            vadrsTimeForSchedule.Fields("TIME_TYPE").Value = 3 '//scheduled
                            vadrsTimeForSchedule.Update
                            blnFoundTimeSummary = True
                            Exit Do
                        End Select
                        
                        rsTimeSummary.MoveNext
                    Loop
                    
                    If Not blnFoundTimeSummary Then
                        Dim blnAddTimeForSchedule As Boolean
                        '//TODO: Do this fix else where
'                        If vadrsSchedule.Fields("SYSTEM_EXC_CD").Value = "SLT" Then
'                            '//need lunch because the time is split (for actual time), otherwise lunch could cause a split
'                            '//scheduled exceptions are not yet split, so lunch is not needed
'                            vadrsScheduledExceptions.Filter = "CHARGE_DATE = " & vadrsSchedule.Fields("CHARGE_DATE").Value & " AND START_TIME = " & vadrsSchedule.Fields("SHIFT_START_TIME").Value & " AND END_TIME = " & vadrsSchedule.Fields("SHIFT_END_TIME").Value
'                            blnAddTimeForSchedule = vadrsScheduledExceptions.EOF  '//if found on sched except then do not add, if not found then add
'                            vadrsScheduledExceptions.Filter = adFilterNone
'                        Else
                            blnAddTimeForSchedule = True
'                        End If

                        If blnAddTimeForSchedule Then
                            AppendRecordsetByName vadrsSchedule, vadrsTimeForSchedule, 1, vadrsSchedule.AbsolutePosition
'                                        vadrsTimeForSchedule.Fields("CHARGE_DATE").Value = MidnightDate(dtTimeSummaryStartDateTime)
'                                        vadrsTimeForSchedule.Fields("DATE_WORKED").Value = MidnightDate(dtScheduleStartDateTime)
                            vadrsTimeForSchedule.Fields("START_TIME").Value = FormatTime(dtScheduleStartDateTime)
                            vadrsTimeForSchedule.Fields("END_TIME").Value = FormatTime(dtScheduleEndDateTime)
                            '//if lunches need to be included, then change the exception to a rest code
                            vadrsTimeForSchedule.Fields("TIME_TYPE").Value = "3" '//Scheduled
                            If vadrsSchedule.Fields("SYSTEM_EXC_CD").Value = "SLT" Then 'Standard lunch time
'                                                vadrsTimeForSchedule.Fields("EXCEPTION_CD").Value = "LU" 'indicate rest rather than working
                                vadrsTimeForSchedule.Fields("EXCEPTION_CD").Value = vadrsSchedule.Fields("SYSTEM_EXC_CD").Value
                            End If
                            vadrsTimeForSchedule.Update
                        End If
'                                        End If
                    End If
                End If
                vadrsSchedule.MoveNext
            Loop
        End If
    End If
    
    Set MergeScheduleWithTime_Private = vadrsTimeForSchedule
    '//vadrsScheduledExceptions.Filter = adFilterNone
    '//Set vadrsTimeForSchedule = Nothing
    '//Set vadrsSchedule = Nothing
    Set rsTimeSummary = Nothing
    Exit Function
eh:
    '//Set vadrsTimeForSchedule = Nothing
    '//vadrsScheduledExceptions.Filter = adFilterNone
    '//Set vadrsSchedule = Nothing
    Set rsTimeSummary = Nothing
    
    TrsRaiseError MODULENAME, "MergeScheduleWithTime_Private"
End Function

Private Function AddTimeAndSchExceptions(ByVal vadrsInitialAbsence As ADODB.Recordset _
                    , ByVal vadrsScheduledExceptions As ADODB.Recordset _
                    , ByVal vadrsSchedule As ADODB.Recordset _
                    , ByVal vadrsTime As ADODB.Recordset _
                    , ByVal vadrsLockStatus As ADODB.Recordset _
                    ) As ADODB.Recordset
On Error GoTo AddTimeAndSchExceptions_EH
    Dim ladrs As ADODB.Recordset
    Dim llIndex As Long ', lbAddRecord As Boolean
'    Dim llFilterIndex As Long, ldtStDtCompare As Date
'    Dim ldtEndDtCompare As Date,
    Dim ldtAbsStDateTime As Date, ldtAbsEndDateTime As Date, ldtAbsRTWDateTime As Date
    
    '//KEEP:
    '//ignore changes from the client that have been processed by payroll, these are not really changes
    If (RecordsExist(vadrsScheduledExceptions)) Then
        With vadrsScheduledExceptions
            .Sort = "CHARGE_DATE, START_TIME, END_TIME"
            .MoveFirst

            For llIndex = .RecordCount To 1 Step -1
                If (!REC_STATUS = "C") Then
                    If (RecordsExist(vadrsLockStatus)) Then
                        vadrsLockStatus.MoveFirst
                        vadrsLockStatus.Find "CHARGE_DATE = #" & !CHARGE_DATE & "#"
                        If Not (vadrsLockStatus.EOF) Then
                            If (vadrsLockStatus!LOCK_STATUS <> "N") Then
                                .Delete adAffectCurrent
                                .Update
                            End If
                        End If
                    End If
                End If

                .MoveNext
            Next llIndex
        End With
    End If

    If (RecordsExist(vadrsTime)) Then
            vadrsTime.Filter = adFilterNone

            '//KEEP:
            '//date worked may not be right on absence end date, abs 50011, 03/25/2011 18:00 - 03/25/2011 6:00, end date worked should be 3/26
            If vadrsInitialAbsence!ABS_END_DT <> #1/1/100# Then
                ldtAbsEndDateTime = DateTimeFrom(vadrsInitialAbsence!ABS_END_DT, vadrsInitialAbsence!ABS_END_TM) '//end time check already done earlier
            Else
                ldtAbsEndDateTime = DateTimeFrom(vadrsInitialAbsence!RTN_TO_WRK_DT, vadrsInitialAbsence!RTN_TO_WRK_TM)
            End If
            
            ldtAbsStDateTime = FindDateAndTimeWorked(vadrsInitialAbsence!ABS_ST_DT, vadrsInitialAbsence!ABS_ST_TM, vadrsTime)
            ldtAbsEndDateTime = FindDateTimeWorked(ldtAbsEndDateTime, vadrsTime)
            ldtAbsRTWDateTime = DateTimeFrom(vadrsInitialAbsence!RTN_TO_WRK_DT, vadrsInitialAbsence!RTN_TO_WRK_TM)
                            
            Set ladrs = AddTimeAndSchExceptions_Private(ldtAbsStDateTime, ldtAbsEndDateTime, ldtAbsRTWDateTime, _
                                                        vadrsScheduledExceptions, vadrsSchedule, vadrsTime)
    End If
    
    Set AddTimeAndSchExceptions = ladrs
    Set ladrs = Nothing

    Exit Function
AddTimeAndSchExceptions_EH:
    Set ladrs = Nothing

    TrsRaiseError MODULENAME _
        , "AddTimeAndSchExceptions"
End Function

'//for testing
'//TODO: change Public to Private
Public Function AddTimeAndSchExceptions2(ByVal vlEmployeeNumber As Long _
                    , ByVal vdtStartDateTime As Date, ByVal vdtEndDateTime As Date, ByVal vdtReturnToWorkDateTime As Date _
                    , ByVal vadrsScheduledExceptions As ADODB.Recordset _
                    ) As ADODB.Recordset
On Error GoTo eh
    Dim ladrs As ADODB.Recordset
    Dim ladrsSchedule As ADODB.Recordset
    Dim ladrsTime As ADODB.Recordset

    Set ladrsSchedule = GetEmployeeSchedule2(vlEmployeeNumber, vdtStartDateTime, vdtEndDateTime)
    Set ladrsTime = GetEmployeeTimeDetails2(vlEmployeeNumber, vdtStartDateTime, vdtEndDateTime)
    
    Set ladrs = AddTimeAndSchExceptions_Private(vdtStartDateTime, vdtEndDateTime, vdtReturnToWorkDateTime, _
                                                vadrsScheduledExceptions, ladrsSchedule, ladrsTime)
    
    Set AddTimeAndSchExceptions2 = ladrs
    Set ladrs = Nothing

    Exit Function
eh:
    Set ladrs = Nothing

    TrsRaiseError MODULENAME, "AddTimeAndSchExceptions2"
End Function

Private Function AddTimeAndSchExceptions_Private(ByVal vdtStartDateTime As Date, ByVal vdtEndDateTime As Date, ByVal vdtReturnToWorkDateTime As Date _
                    , ByVal vadrsScheduledExceptions As ADODB.Recordset _
                    , ByVal vadrsSchedule As ADODB.Recordset _
                    , ByVal vadrsTime As ADODB.Recordset _
                    ) As ADODB.Recordset
On Error GoTo eh
    Dim ladrs As ADODB.Recordset
    Dim llIndex As Long, lbAddRecord As Boolean
    Dim llFilterIndex As Long, ldtStDtCompare As Date, ldtEndDtCompare As Date ', ldtAbsStDateTime As Date, ldtAbsEndDateTime As Date
    
    Set ladrs = New ADODB.Recordset
    With ladrs
        .Fields.Append "CHARGE_DATE", adDate, , adFldUpdatable
        .Fields.Append "DATE_WORKED", adDate, , adFldUpdatable
        .Fields.Append "START_TIME", adChar, 5, adFldUpdatable
        .Fields.Append "END_TIME", adChar, 5, adFldUpdatable
        .Fields.Append "EXCEPTION_CD", adChar, 3, adFldUpdatable
        .Fields.Append "OVER_EXC_CD", adChar, 3, adFldUpdatable

        .Fields.Append "REC_STATUS", adChar, 1, adFldUpdatable

        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open
    End With

'//new
    Dim ldatStartTime1 As Date, ldatEndTime1 As Date
    Dim ldatStartTime2 As Date, ldatEndTime2 As Date
    
    '//assigned date worked based on schedule
    If (RecordsExist(vadrsScheduledExceptions)) Then
        With vadrsScheduledExceptions
            .Sort = "CHARGE_DATE, START_TIME, END_TIME"
            .MoveFirst

            For llIndex = 1 To .RecordCount
'Debug.Print PadRight("vadrsScheduledExceptions", 30) & StringFormat("{0} {1} {2}", !CHARGE_DATE, !START_TIME, !END_TIME)
'//I don't think that DetermineAbsenceStartAndEnd will understand the split over midnight, it uses shift_start_time, shift_end_time
'                If SplitOverMidnight(!CHARGE_DATE, !START_TIME, !END_TIME, ldatStartTime1, ldatEndTime1, ldatStartTime2, ldatEndTime2) Then
'                    ladrs.AddNew
'                    ladrs!CHARGE_DATE = !CHARGE_DATE
'                    ladrs!START_TIME = FormatTime(ldatStartTime1)
'                    ladrs!END_TIME = FormatTime(ldatEndTime1)
'                    ladrs!DATE_WORKED = MidnightDate(ldatStartTime1)
'                    ladrs!EXCEPTION_CD = !EXCEPTION_CD
'                    ladrs!OVER_EXC_CD = !OVER_EXC_CD
'                    ladrs!REC_STATUS = !REC_STATUS
'                    ladrs.Update
'
'                    ladrs.AddNew
'                    ladrs!CHARGE_DATE = !CHARGE_DATE
'                    ladrs!START_TIME = FormatTime(ldatStartTime2)
'                    ladrs!END_TIME = FormatTime(ldatEndTime2)
'                    ladrs!DATE_WORKED = MidnightDate(ldatStartTime2)
'                    ladrs!EXCEPTION_CD = !EXCEPTION_CD
'                    ladrs!OVER_EXC_CD = !OVER_EXC_CD
'                    ladrs!REC_STATUS = !REC_STATUS
'                    ladrs.Update
'                Else '//TODO: the dateworked may not be correct
                    ladrs.AddNew
                    ladrs!CHARGE_DATE = !CHARGE_DATE
                    ladrs!START_TIME = !START_TIME
                    ladrs!END_TIME = !END_TIME
                    
                    '//determine the dateworked based on the schedule
                    If (RecordsExist(vadrsSchedule)) Then
                        vadrsSchedule.Filter = adFilterNone
    
                        vadrsSchedule.Filter = "CHARGE_DATE = #" & !CHARGE_DATE & "#" & _
                                " AND EXCEPTION_CD <> 'ONP'"
                        If Not (vadrsSchedule.EOF) Then
                            vadrsSchedule.MoveFirst
    
                            For llFilterIndex = 1 To vadrsSchedule.RecordCount
'Debug.Print PadRight("   vadrsSchedule", 30) & StringFormat("{0} {1} {2} {3}", vadrsSchedule!CHARGE_DATE, vadrsSchedule!DATE_WORKED, vadrsSchedule!START_TIME, vadrsSchedule!END_TIME)
                                '//there could be a trade in the schedule for a non-working shift that does not have a start time
                                If Trim(vadrsSchedule!START_TIME) <> "" Then
                                    ldatStartTime1 = DateTimeFrom(!CHARGE_DATE, !START_TIME)
                                    ldatStartTime2 = DateTimeFrom(vadrsSchedule!CHARGE_DATE, vadrsSchedule!START_TIME)
                                    
                                    If ldatStartTime1 >= ldatStartTime2 Then
                                        ladrs!DATE_WORKED = vadrsSchedule!DATE_WORKED
                                        Exit For
                                    End If
                                End If
    
                                vadrsSchedule.MoveNext
                            Next llFilterIndex
                            '//HACK: work around for getting a type mismatch because the date (empty) was used in a filter
                            '//TODO: the real issue was that the schedule and schedule exceptions did not match on the client request
                            If IsNullOrEmpty(ladrs!DATE_WORKED) Then
                                vadrsSchedule.MoveFirst
                                ladrs!DATE_WORKED = vadrsSchedule!DATE_WORKED
                            End If
                        End If
    
                        vadrsSchedule.Filter = adFilterNone
                    End If
    
                    '//HACK: work around for getting a type mismatch because the date (empty) was used in a filter
                    '//TODO: the real issue was that the schedule and schedule exceptions did not match on the client request
                    If IsNullOrEmpty(ladrs!DATE_WORKED) Then
                        ladrs!DATE_WORKED = !CHARGE_DATE
                    End If
                    
                    ladrs!EXCEPTION_CD = !EXCEPTION_CD
                    ladrs!OVER_EXC_CD = !OVER_EXC_CD
                    ladrs!REC_STATUS = !REC_STATUS
                    ladrs.Update
'                End If

                .MoveNext
            Next llIndex

            .MoveFirst
        End With
    End If

    If (RecordsExist(vadrsTime)) Then
        With vadrsTime
            .Filter = adFilterNone
                           
'            DateTimeCrossesMidnight DateTimeFrom(ldtAbsStDate, vadrsInitialAbsence!ABS_ST_TM), ldtAbsEndDateTime
'
'            '//TODO: may not need to do this, if the absence initial has been modified
'            '//adjust the end datetime to be dateworked
'            .Filter = StringFormat("(TIME_TYPE = 2 AND OT_OPP_NUM <= 0 AND EXCEPTION_CD <> 'ONP' AND CHARGE_DATE = #{0}# AND END_TIME = '{1}')" & _
'                    " OR (TIME_TYPE = 3 AND OT_OPP_NUM <= 0 AND EXCEPTION_CD <> 'ONP' AND CHARGE_DATE = #{0}# AND END_TIME = '{1}')", _
'                    vadrsInitialAbsence!ABS_END_DT, vadrsInitialAbsence!ABS_END_TM)
'            If Not .EOF Then
'                If !CHARGE_DATE <> !DATE_WORKED Then
'                    ldtAbsEndDateTime = DateTimeFrom(!DATE_WORKED, vadrsInitialAbsence!ABS_END_TM)
'                End If
'            End If
            
            'look at Actual Time and Schedule only
            .Filter = ("(TIME_TYPE = 2 AND OT_OPP_NUM <= 0 AND EXCEPTION_CD <> 'ONP')" & _
                    " OR (TIME_TYPE = 3 AND OT_OPP_NUM <= 0 AND EXCEPTION_CD <> 'ONP')")
            If Not (.EOF) Then
                .Sort = "CHARGE_DATE, DATE_WORKED, START_TIME, END_TIME"
                .MoveFirst

'                ldtAbsStDate = vadrsInitialAbsence!ABS_ST_DT
Debug.Print ""
'Debug.Print StringFormat("AbsNum={0}, AbsStart={1} AbsEnd={2} AbsRTW={3}", vadrsInitialAbsence!ABS_NUM, FormatDateTime(ldtAbsStDateTime), FormatDateTime(ldtAbsEndDateTime), _
'                FormatDateAndTime(vadrsInitialAbsence!RTN_TO_WRK_DT, IIf(Trim(vadrsInitialAbsence!RTN_TO_WRK_TM) = "", "00:00", vadrsInitialAbsence!RTN_TO_WRK_TM), True))
                For llIndex = 1 To .RecordCount
'Debug.Print PadRight("Schedule/Time", 30) & StringFormat("{0} {1} {2} {3}", !CHARGE_DATE, !DATE_WORKED, !START_TIME, !END_TIME)
                    If (IsDate(!START_TIME)) Or _
                        (IsDate(!END_TIME)) Then
                        '//since this is from actual time/schedule date worked should be ok for compares (splits over midnight)
                        ldtStDtCompare = !DATE_WORKED '!CHARGE_DATE
                        ldtEndDtCompare = !DATE_WORKED '!CHARGE_DATE

'                        '//the initial absence is using a charge date, but we need a date worked for comparison
'                        If !CHARGE_DATE <> !DATE_WORKED Then
'                            If !CHARGE_DATE = vadrsInitialAbsence!ABS_ST_DT And !START_TIME = vadrsInitialAbsence!ABS_ST_TM Then
'                                ldtAbsStDate = !DATE_WORKED
'                            End If
'                        End If
'
'chargedate, dateworked, st, et
'11/19/2010 11/19/2010 18:00 00:00 processed
'11/19/2010 11/20/2010 00:00 06:00 was being ignored
                        'TODO: Fix for partials, if ABS is after start time this will be false, time=11:30-15:30, abs=13:30-15:30
Debug.Print StringFormat("TimeStart={0} TimeEnd={1}", FormatDateAndTime(ldtStDtCompare, !START_TIME), FormatDateAndTime(ldtEndDtCompare, !END_TIME, True))
                        If TimePeriodsIntersect(DateTimeFrom(ldtStDtCompare, !START_TIME), DateTimeFrom(ldtEndDtCompare, !END_TIME, True), _
                                            vdtStartDateTime, vdtEndDateTime) Then
                        'If DateTimeFrom(ldtStDtCompare, !START_TIME) >= DateTimeFrom(ldtAbsStDate, vadrsInitialAbsence!ABS_ST_TM) Then
                        'Or DateTimeFrom(ldtEndDtCompare, !END_TIME) >= DateTimeFrom(ldtAbsEndDate, vadrsInitialAbsence!ABS_END_TM) Then
                            '//HACK: received type mismatch when return to work time is "     "
                            '//use end date only instead of return to work date
'                            Dim dtReturnToWork As Date
'
'                            dtReturnToWork = DateTimeFrom(vadrsInitialAbsence!RTN_TO_WRK_DT, IIf(Trim(vadrsInitialAbsence!RTN_TO_WRK_TM) = "", "00:00", vadrsInitialAbsence!RTN_TO_WRK_TM), True)
                            'TODO: Fix for partials, if RTW is before end time this will be false, time=11:30-15:30, abs=11:30-13:30
                            If DateTimeFrom(ldtEndDtCompare, !END_TIME, True) <= vdtEndDateTime Or _
                                TimePeriodsIntersect(DateTimeFrom(ldtStDtCompare, !START_TIME), DateTimeFrom(ldtEndDtCompare, !END_TIME, True), _
                                                vdtEndDateTime, vdtEndDateTime) Then

'                            If DateTimeFrom(ldtEndDtCompare, !END_TIME, True) <= dtReturnToWork Or _
'                                TimePeriodsIntersect(DateTimeFrom(ldtStDtCompare, !START_TIME), DateTimeFrom(ldtEndDtCompare, !END_TIME, True), _
'                                                dtReturnToWork, dtReturnToWork) Then
                            
                                If (RecordsExist(ladrs)) Then
                                    ladrs.Filter = adFilterNone

                                    ladrs.Filter = ("CHARGE_DATE = #" & !CHARGE_DATE & _
                                                "# AND DATE_WORKED = #" & !DATE_WORKED & _
                                                "# AND START_TIME = '" & !START_TIME & "'")
                                    If Not (ladrs.EOF) Then
                                        For llFilterIndex = 1 To ladrs.RecordCount
'Debug.Print PadRight("   AbsenceTime (ladrs)", 30) & StringFormat("{0} {1} {2} {3}", ladrs!CHARGE_DATE, ladrs!DATE_WORKED, ladrs!START_TIME, ladrs!END_TIME) & " Loop 1"
                                            If (DateTimeFrom(ladrs!DATE_WORKED, ladrs!START_TIME) <> DateTimeFrom(ldtStDtCompare, !START_TIME) Or _
                                                DateTimeFrom(ladrs!CHARGE_DATE, ladrs!END_TIME, True) <> DateTimeFrom(ldtEndDtCompare, !END_TIME, True)) And _
                                                Trim(ladrs!EXCEPTION_CD) = Trim(!EXCEPTION_CD) And Trim(ladrs!OVER_EXC_CD) = Trim(!OVER_EXC_CD) And _
                                                Trim(ladrs!REC_STATUS) = "" Then
                                                ladrs!CHARGE_DATE = !CHARGE_DATE
                                                ladrs!DATE_WORKED = !DATE_WORKED

                                                If DateTimeFrom(ladrs!DATE_WORKED, ladrs!START_TIME) > DateTimeFrom(ldtStDtCompare, !START_TIME) Then
                                                    ladrs!START_TIME = !START_TIME
                                                End If
                                                If DateTimeFrom(ladrs!DATE_WORKED, ladrs!END_TIME) < DateTimeFrom(ldtEndDtCompare, !END_TIME) Then
                                                    ladrs!END_TIME = !END_TIME
                                                End If
                                                ladrs!EXCEPTION_CD = !EXCEPTION_CD
                                                ladrs!OVER_EXC_CD = !OVER_EXC_CD
                                                ladrs!REC_STATUS = ""
                                                ladrs.Update
'Debug.Print PadRight("   AbsenceTime (ladrs)", 30) & StringFormat("{0} {1} {2} {3}", ladrs!CHARGE_DATE, ladrs!DATE_WORKED, ladrs!START_TIME, ladrs!END_TIME) & " has been updated only 1"
                                            End If

                                            ladrs.MoveNext
                                        Next llFilterIndex
                                    Else
                                        '//use end date only instead of return to work date
                                        If (!CHARGE_DATE <= MidnightDate(vdtEndDateTime)) Then
                                        'If (!CHARGE_DATE <= vadrsInitialAbsence!RTN_TO_WRK_DT) Then
                                            ladrs.Filter = adFilterNone

                                            ladrs.Filter = "CHARGE_DATE = #" & !CHARGE_DATE & "#"
                                            If Not (ladrs.EOF) Then
                                                For llFilterIndex = 1 To ladrs.RecordCount
'Debug.Print PadRight("   AbsenceTime (ladrs)", 30) & StringFormat("{0} {1} {2} {3}", ladrs!CHARGE_DATE, ladrs!DATE_WORKED, ladrs!START_TIME, ladrs!END_TIME) & " Loop 2"
                                                    '//had trouble trying to match sched exceptions that are not split with actual time/schedule that is split
                                                    '//the end time was failing the compare for something not split over midnight
'                                                    ldatStartTime1 = DateTimeFrom(ladrs!DATE_WORKED, ladrs!START_TIME)
'                                                    ldatEndTime1 = DateTimeFrom(ladrs!DATE_WORKED, ladrs!END_TIME, True)
'                                                    DateTimeCrossesMidnight ldatStartTime1, ldatEndTime1
                                                    DateAndTimeRangeFrom ladrs!DATE_WORKED, ladrs!START_TIME, _
                                                                        ladrs!DATE_WORKED, ladrs!END_TIME, _
                                                                        ldatStartTime1, ldatEndTime1
                                                    
                                                    ldatStartTime2 = DateTimeFrom(ldtStDtCompare, !START_TIME)
                                                    ldatEndTime2 = DateTimeFrom(ldtEndDtCompare, !END_TIME, True)
                                                    
                                                    '//if the time/schedule is a subset of the results then ignore it
                                                    If ldatStartTime1 <= ldatStartTime2 And ldatEndTime1 >= ldatEndTime2 And _
                                                        Trim(ladrs!EXCEPTION_CD) = Trim(!EXCEPTION_CD) And Trim(ladrs!OVER_EXC_CD) = Trim(!OVER_EXC_CD) Then
'Debug.Print PadRight("   AbsenceTime (ladrs)", 30) & StringFormat("{0} {1} {2} {3}", ladrs!CHARGE_DATE, ladrs!DATE_WORKED, ladrs!START_TIME, ladrs!END_TIME) & " Match found"
                                                        '
                                                    Else
                                                        If (llFilterIndex = ladrs.RecordCount) Then
                                                            '//TODO: Determine what execptions are absense related
                                                            '//if the end date (or rtw) is past the proper date, there could be extra rows that cause the rtw to be later
'                                                            If Trim(!EXCEPTION_CD) <> "ST" Then
                                                                ladrs.AddNew
    
                                                                ladrs!CHARGE_DATE = !CHARGE_DATE
                                                                ladrs!START_TIME = !START_TIME
                                                                ladrs!END_TIME = !END_TIME
                                                                ladrs!DATE_WORKED = !DATE_WORKED
                                                                ladrs!EXCEPTION_CD = !EXCEPTION_CD
                                                                ladrs!OVER_EXC_CD = !OVER_EXC_CD
                                                                ladrs!REC_STATUS = ""
                                                                ladrs.Update
'                                                            End If
'Debug.Print PadRight("   AbsenceTime (ladrs)", 30) & StringFormat("{0} {1} {2} {3}", ladrs!CHARGE_DATE, ladrs!DATE_WORKED, ladrs!START_TIME, ladrs!END_TIME) & " has been added 2"
                                                        End If
                                                    End If

                                                    ladrs.MoveNext
                                                Next llFilterIndex
                                            Else '//could not find the charge date, so add it
                                                ladrs.AddNew

                                                ladrs!CHARGE_DATE = !CHARGE_DATE
                                                ladrs!START_TIME = !START_TIME
                                                ladrs!END_TIME = !END_TIME
                                                ladrs!DATE_WORKED = !DATE_WORKED
                                                ladrs!EXCEPTION_CD = !EXCEPTION_CD
                                                ladrs!OVER_EXC_CD = !OVER_EXC_CD
                                                ladrs!REC_STATUS = ""
                                                ladrs.Update
'Debug.Print PadRight("   AbsenceTime (ladrs)", 30) & StringFormat("{0} {1} {2} {3}", ladrs!CHARGE_DATE, ladrs!DATE_WORKED, ladrs!START_TIME, ladrs!END_TIME) & " has been added 3"
                                            End If

                                            ladrs.Filter = adFilterNone
                                        End If
                                    End If

                                    ladrs.Filter = adFilterNone
                                Else '//first one to add to AbsenceTime, not likely to happen if sched exceptions have already been added?
                                    ladrs.AddNew

                                    ladrs!CHARGE_DATE = !CHARGE_DATE
                                    ladrs!START_TIME = !START_TIME
                                    ladrs!END_TIME = !END_TIME
                                    ladrs!DATE_WORKED = !DATE_WORKED
                                    ladrs!EXCEPTION_CD = !EXCEPTION_CD
                                    ladrs!OVER_EXC_CD = !OVER_EXC_CD
                                    ladrs!REC_STATUS = ""
                                    ladrs.Update
'Debug.Print PadRight("   AbsenceTime (ladrs)", 30) & StringFormat("{0} {1} {2} {3}", ladrs!CHARGE_DATE, ladrs!DATE_WORKED, ladrs!START_TIME, ladrs!END_TIME) & " has been added 4"
                                End If
                            Else 'TODO: Fix for partials
                                If Not RecordsExist(ladrs) Then
                                    ladrs.AddNew

                                    ladrs!CHARGE_DATE = MidnightDate(vdtEndDateTime) '!CHARGE_DATE
                                    ladrs!START_TIME = !START_TIME
                                    ladrs!END_TIME = FormatTime(vdtEndDateTime) '!END_TIME
                                    ladrs!DATE_WORKED = !DATE_WORKED
                                    ladrs!EXCEPTION_CD = !EXCEPTION_CD
                                    ladrs!OVER_EXC_CD = !OVER_EXC_CD
                                    ladrs!REC_STATUS = ""
                                    ladrs.Update
                                End If
                                Exit For
                            End If
                        End If
                    End If

                    .MoveNext
                Next llIndex
            End If

            .Filter = adFilterNone
            .MoveFirst
        End With
    End If

    If (RecordsExist(ladrs)) Then
        ladrs.Filter = adFilterNone
        ladrs.MoveFirst
    End If

'//TODO: Remove or comment out ***for testing***
Debug.Print "AddTimeAndSchExceptions_Private==>"
If RecordsExist(ladrs) Then
    ladrs.MoveFirst
    Do While Not ladrs.EOF
    
Debug.Print StringFormat("{0} {1} {2} {3} {4}", ladrs!CHARGE_DATE, ladrs!DATE_WORKED, ladrs!START_TIME, ladrs!END_TIME, ladrs!EXCEPTION_CD)
        ladrs.MoveNext
    Loop
    ladrs.MoveFirst
End If

    Set AddTimeAndSchExceptions_Private = ladrs
    Set ladrs = Nothing

    Exit Function
eh:
    Set ladrs = Nothing

    TrsRaiseError MODULENAME _
        , "AddTimeAndSchExceptions_Private"
End Function

'//the old
Private Function AddTimeAndSchExceptionsOld(ByVal vadrsInitialAbsence As ADODB.Recordset _
                    , ByVal vadrsScheduledExceptions As ADODB.Recordset _
                    , ByVal vadrsSchedule As ADODB.Recordset _
                    , ByVal vadrsTime As ADODB.Recordset _
                    , ByVal vadrsLockStatus As ADODB.Recordset _
                    ) As ADODB.Recordset
On Error GoTo AddTimeAndSchExceptionsOld_EH
Dim ladrs As ADODB.Recordset
Dim llIndex As Long, lbAddRecord As Boolean
Dim llFilterIndex As Long, ldtStDtCompare As Date
Dim ldtEndDtCompare As Date, ldtAbsStDate As Date

    Set ladrs = New ADODB.Recordset
    With ladrs
        .Fields.Append "CHARGE_DATE", adDate, , adFldUpdatable
        .Fields.Append "DATE_WORKED", adDate, , adFldUpdatable
        .Fields.Append "START_TIME", adChar, 5, adFldUpdatable
        .Fields.Append "END_TIME", adChar, 5, adFldUpdatable
        .Fields.Append "EXCEPTION_CD", adChar, 3, adFldUpdatable
        .Fields.Append "OVER_EXC_CD", adChar, 3, adFldUpdatable

        .Fields.Append "REC_STATUS", adChar, 1, adFldUpdatable

        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open
    End With

    If (RecordsExist(vadrsScheduledExceptions)) Then
        With vadrsScheduledExceptions
            .Sort = "CHARGE_DATE, START_TIME, END_TIME"
            .MoveFirst

            For llIndex = .RecordCount To 1 Step -1
                If (!REC_STATUS = "C") Then
                    If (RecordsExist(vadrsLockStatus)) Then
                        vadrsLockStatus.MoveFirst
                        vadrsLockStatus.Find "CHARGE_DATE = #" & !CHARGE_DATE & "#"
                        If Not (vadrsLockStatus.EOF) Then
                            If (vadrsLockStatus!LOCK_STATUS <> "N") Then
                                .Delete adAffectCurrent
                                .Update
                            End If
                        End If
                    End If
                End If

                .MoveNext
            Next llIndex
        End With
    End If

    If (RecordsExist(vadrsScheduledExceptions)) Then
        With vadrsScheduledExceptions
            .Sort = "CHARGE_DATE, START_TIME, END_TIME"
            .MoveFirst

            For llIndex = 1 To .RecordCount
                ladrs.AddNew

                ladrs!CHARGE_DATE = !CHARGE_DATE
                ladrs!START_TIME = !START_TIME
                ladrs!END_TIME = !END_TIME

                If (RecordsExist(vadrsSchedule)) Then
                    vadrsSchedule.Filter = adFilterNone

                    vadrsSchedule.Filter = "CHARGE_DATE = #" & !CHARGE_DATE & "#" & _
                            " AND EXCEPTION_CD <> 'ONP'"
                    If Not (vadrsSchedule.EOF) Then
                        vadrsSchedule.MoveFirst

                        For llFilterIndex = 1 To vadrsSchedule.RecordCount
                            If (CDate(!CHARGE_DATE & " " & !START_TIME) >= _
                                CDate(vadrsSchedule!DATE_WORKED & " " & vadrsSchedule!SHIFT_START_TIME)) And _
                                ((CDate(!CHARGE_DATE & " " & !END_TIME) <= _
                                CDate(vadrsSchedule!CHARGE_DATE & " " & vadrsSchedule!SHIFT_END_TIME)) Or _
                                (CDate(!CHARGE_DATE & " " & !END_TIME) >= _
                                CDate(vadrsSchedule!CHARGE_DATE & " " & vadrsSchedule!SHIFT_END_TIME))) Then
                                ladrs!DATE_WORKED = vadrsSchedule!DATE_WORKED
                                Exit For
                            End If

                            vadrsSchedule.MoveNext
                        Next llFilterIndex
                    End If

                    vadrsSchedule.Filter = adFilterNone
                End If

                ladrs!EXCEPTION_CD = !EXCEPTION_CD
                ladrs!OVER_EXC_CD = !OVER_EXC_CD
                ladrs!REC_STATUS = !REC_STATUS

                ladrs.Update

                .MoveNext
            Next llIndex

            .MoveFirst
        End With
    End If

    If (RecordsExist(vadrsTime)) Then
        With vadrsTime
            .Filter = adFilterNone

            'look at Actual Time and Schedule only
            .Filter = ("(TIME_TYPE = 2 AND OT_OPP_NUM <= 0 AND EXCEPTION_CD <> 'ONP')" & _
                    " OR (TIME_TYPE = 3 AND OT_OPP_NUM <= 0 AND EXCEPTION_CD <> 'ONP')")
            If Not (.EOF) Then
                .Sort = "CHARGE_DATE, DATE_WORKED, START_TIME, END_TIME"
                .MoveFirst

                ldtAbsStDate = vadrsInitialAbsence!ABS_ST_DT

                For llIndex = 1 To .RecordCount
                    If (IsDate(!START_TIME)) Or _
                        (IsDate(!END_TIME)) Then
                        ldtStDtCompare = !CHARGE_DATE
                        ldtEndDtCompare = !CHARGE_DATE

                        If (!CHARGE_DATE <> !DATE_WORKED) Then
                            ldtStDtCompare = !DATE_WORKED

                            If (!CHARGE_DATE = vadrsInitialAbsence!ABS_ST_DT) Then
                                ldtAbsStDate = !DATE_WORKED
                            End If
                        End If

                        If (!END_TIME >= !START_TIME) Then
                            ldtEndDtCompare = !DATE_WORKED
                        Else
                            If (!END_TIME < !START_TIME) Then
                                ldtEndDtCompare = !CHARGE_DATE
                            End If
                        End If

                        If (CDate(ldtStDtCompare & " " & !START_TIME) >= CDate(ldtAbsStDate & " " & vadrsInitialAbsence!ABS_ST_TM)) Then
                            If (CDate(ldtEndDtCompare & " " & !END_TIME) <= CDate(vadrsInitialAbsence!RTN_TO_WRK_DT & " " & vadrsInitialAbsence!RTN_TO_WRK_TM)) Then
                                If (RecordsExist(ladrs)) Then
                                    ladrs.Filter = adFilterNone

                                    ladrs.Filter = ("CHARGE_DATE = #" & !CHARGE_DATE & _
                                                "# AND DATE_WORKED = #" & !DATE_WORKED & _
                                                "# AND START_TIME = '" & !START_TIME & "'")
                                    If Not (ladrs.EOF) Then
                                        For llFilterIndex = 1 To ladrs.RecordCount
                                            If ((CDate(ladrs!DATE_WORKED & " " & ladrs!START_TIME) <> _
                                                CDate(ldtStDtCompare & " " & !START_TIME)) Or _
                                                (CDate(ladrs!CHARGE_DATE & " " & ladrs!END_TIME) <> _
                                                CDate(ldtEndDtCompare & " " & !END_TIME))) And _
                                                (Trim(ladrs!EXCEPTION_CD) = Trim(!EXCEPTION_CD)) And _
                                                (Trim(ladrs!OVER_EXC_CD) = Trim(!OVER_EXC_CD)) And (Trim(ladrs!REC_STATUS) = "") Then
                                                ladrs!CHARGE_DATE = !CHARGE_DATE
                                                ladrs!DATE_WORKED = !DATE_WORKED

                                                If (CDate(ladrs!DATE_WORKED & " " & ladrs!START_TIME) > CDate(ldtStDtCompare & " " & !START_TIME)) Then
                                                    ladrs!START_TIME = !START_TIME
                                                End If
                                                If (CDate(ladrs!DATE_WORKED & " " & ladrs!END_TIME) < CDate(ldtEndDtCompare & " " & !END_TIME)) Then
                                                    ladrs!END_TIME = !END_TIME
                                                End If
                                                ladrs!EXCEPTION_CD = !EXCEPTION_CD
                                                ladrs!OVER_EXC_CD = !OVER_EXC_CD
                                                ladrs!REC_STATUS = ""
                                                ladrs.Update
                                            End If

                                            ladrs.MoveNext
                                        Next llFilterIndex
                                    Else
                                        If (!CHARGE_DATE <= vadrsInitialAbsence!RTN_TO_WRK_DT) Then
                                            ladrs.Filter = adFilterNone

                                            ladrs.Filter = "CHARGE_DATE = #" & !CHARGE_DATE & "#"
                                            If Not (ladrs.EOF) Then
                                                For llFilterIndex = 1 To ladrs.RecordCount
                                                    If ((CDate(ladrs!DATE_WORKED & " " & ladrs!START_TIME) <= _
                                                        CDate(ldtStDtCompare & " " & !START_TIME)) And _
                                                        (CDate(ladrs!CHARGE_DATE & " " & ladrs!END_TIME) >= _
                                                        CDate(ldtEndDtCompare & " " & !END_TIME))) And _
                                                        (Trim(ladrs!EXCEPTION_CD) = Trim(!EXCEPTION_CD)) And _
                                                        (Trim(ladrs!OVER_EXC_CD) = Trim(!OVER_EXC_CD)) Then
                                                        '
                                                    Else
                                                        If (llFilterIndex = ladrs.RecordCount) Then
                                                            ladrs.AddNew

                                                            ladrs!CHARGE_DATE = !CHARGE_DATE
                                                            ladrs!START_TIME = !START_TIME
                                                            ladrs!END_TIME = !END_TIME
                                                            ladrs!DATE_WORKED = !DATE_WORKED
                                                            ladrs!EXCEPTION_CD = !EXCEPTION_CD
                                                            ladrs!OVER_EXC_CD = !OVER_EXC_CD
                                                            ladrs!REC_STATUS = ""
                                                            ladrs.Update
                                                        End If
                                                    End If

                                                    ladrs.MoveNext
                                                Next llFilterIndex
                                            Else
                                                ladrs.AddNew

                                                ladrs!CHARGE_DATE = !CHARGE_DATE
                                                ladrs!START_TIME = !START_TIME
                                                ladrs!END_TIME = !END_TIME
                                                ladrs!DATE_WORKED = !DATE_WORKED
                                                ladrs!EXCEPTION_CD = !EXCEPTION_CD
                                                ladrs!OVER_EXC_CD = !OVER_EXC_CD
                                                ladrs!REC_STATUS = ""
                                                ladrs.Update
                                            End If

                                            ladrs.Filter = adFilterNone
                                        End If
                                    End If

                                    ladrs.Filter = adFilterNone
                                Else
                                    ladrs.AddNew

                                    ladrs!CHARGE_DATE = !CHARGE_DATE
                                    ladrs!START_TIME = !START_TIME
                                    ladrs!END_TIME = !END_TIME
                                    ladrs!DATE_WORKED = !DATE_WORKED
                                    ladrs!EXCEPTION_CD = !EXCEPTION_CD
                                    ladrs!OVER_EXC_CD = !OVER_EXC_CD
                                    ladrs!REC_STATUS = ""
                                    ladrs.Update
                                End If
                            Else
                                Exit For
                            End If
                        End If
                    End If

                    .MoveNext
                Next llIndex
            End If

            .Filter = adFilterNone
            .MoveFirst
        End With
    End If

    If (RecordsExist(ladrs)) Then
        ladrs.Filter = adFilterNone
        ladrs.MoveFirst
    End If

    Set AddTimeAndSchExceptionsOld = ladrs
    Set ladrs = Nothing

    Exit Function
AddTimeAndSchExceptionsOld_EH:
    Set ladrs = Nothing

    TrsRaiseError MODULENAME _
        , "AddTimeAndSchExceptionsOld"
End Function
'//the old end

'given 12/26/2010 18:00 02:00 and 12/27/2010 00:00 02:00
'treat this internally as 12/26/2010 18:00 - 12/27/2010 02:00 and 12/27/2010 00:00 - 12/27/2010 02:00
'return 1 (overlap)
'Private Function DateIntersectionOverMidnight(ByVal v_datStartTime1 As Date, ByVal v_datEndTime1 As Date, _
'                                            ByVal v_datStartTime2 As Date, ByVal v_datEndTime2 As Date) As Integer
'    Dim liMatch As Integer
'
'    DateTimeCrossesMidnight v_datStartTime1, v_datEndTime1
'    DateTimeCrossesMidnight v_datStartTime2, v_datEndTime2
'    liMatch = DateIntersection(v_datStartTime1, v_datEndTime1, v_datStartTime2, v_datEndTime2)
'
'    DateIntersectionOverMidnight = liMatch
'End Function

Private Function FormatDateTime(ByVal v_datDate As Date) As String
On Error GoTo eh
    FormatDateTime = Format(v_datDate, "MM/dd/yyyy hh:mm")
    Exit Function
eh:
    TrsRaiseError MODULENAME, "FormatDateTime", Err.Number, , , , True
    Exit Function
End Function

Private Function FormatDateAndTime(ByVal v_datDate As Date, ByVal v_strTime As String, _
                            Optional ByVal v_blnEndTime As Boolean) As String
On Error GoTo eh
    Dim dt As Date
    
    dt = DateTimeFrom2(v_datDate, v_strTime, v_blnEndTime)
    FormatDateAndTime = Format(dt, "MM/dd/yyyy hh:mm")
    Exit Function
eh:
    TrsRaiseError MODULENAME, "FormatDateAndTime", Err.Number, , , , True
    Exit Function
End Function

Private Function FormatShortDateTime(ByVal v_datDate As Date) As String
On Error GoTo eh
    
    FormatShortDateTime = Format(v_datDate, "M/D h:mm")
    
    Exit Function
eh:
    TrsRaiseError MODULENAME, "FormatShortDateTime", Err.Number, , , , True
    Exit Function
End Function

Private Function FormatShortDateAndTime(ByVal v_datDate As Date, ByVal v_strTime As String, _
                            Optional ByVal v_blnEndTime As Boolean) As String
On Error GoTo eh
    Dim dt As Date
    
    dt = DateTimeFrom2(v_datDate, v_strTime, v_blnEndTime)
    
    FormatShortDateAndTime = Format(dt, "M/D h:mm")
    
    Exit Function
eh:
    TrsRaiseError MODULENAME, "FormatShortDateAndTime", Err.Number, , , , True
    Exit Function
End Function

Private Function FormatShortDateTimeRange(ByVal dtStartDate As Date, ByVal dtEndDate As Date) As String
On Error GoTo eh
    If MidnightDate(dtStartDate) = MidnightDate(dtEndDate) Or _
        (MidnightDate(AddDays(1, dtStartDate)) = MidnightDate(dtEndDate) And DateDiff("n", dtStartDate, dtEndDate) < 1440) Then
        FormatShortDateTimeRange = StringFormat("{0} - {1}", Format(dtStartDate, "M/D h:mm"), Format(dtEndDate, "h:mm"))
    Else
        FormatShortDateTimeRange = StringFormat("{0} - {1}", Format(dtStartDate, "M/D h:mm"), Format(dtEndDate, "M/D h:mm"))
    End If
    Exit Function
eh:
    TrsRaiseError MODULENAME, "FormatShortDateTimeRange", Err.Number, , , , True
    Exit Function
End Function

Public Function DateTimeFrom2(ByVal v_datDate As Date, ByVal v_strTime As String, _
                            Optional ByVal v_blnEndTime As Boolean) As Date
On Error GoTo eh
    If Trim(v_strTime) = "" Then
        v_strTime = "00:00"
    End If

    If v_datDate <> #1/1/100# Then
        DateTimeFrom2 = DateTimeFrom(v_datDate, v_strTime, v_blnEndTime)
    Else
        DateTimeFrom2 = #1/1/100#
    End If
    Exit Function
eh:
    TrsRaiseError MODULENAME, "DateTimeFrom2", Err.Number, , , , True
    Exit Function
End Function

'//NOTE: if FindDateTimeWorked is called after this, then do not use vbln_AdjustForCrossingMidnight = True
Private Sub DateAndTimeRangeFrom(ByVal v_datStartDate As Date, ByVal v_strStartTime As String, _
                                        ByVal v_datEndDate As Date, ByVal v_strEndTime As String, _
                                        ByRef r_datStartDate As Date, ByRef r_datEndDate As Date, _
                                        Optional ByVal vbln_AdjustForCrossingMidnight As Boolean = True)
On Error GoTo eh
    r_datStartDate = DateTimeFrom(v_datStartDate, v_strStartTime)
    
    If Trim(v_strEndTime) = "" Then
        v_strEndTime = "00:00"
    End If
    
    If v_datEndDate <> #1/1/100# Then
        r_datEndDate = DateTimeFrom(v_datEndDate, v_strEndTime, vbln_AdjustForCrossingMidnight)
    Else
        r_datEndDate = DateTimeFrom(v_datEndDate, v_strEndTime)
    End If
    
    If v_datEndDate <> #1/1/100# Then
        If vbln_AdjustForCrossingMidnight Then
            DateTimeCrossesMidnight r_datStartDate, r_datEndDate
        End If
    End If

    Exit Sub
eh:
    TrsRaiseError MODULENAME, "DateAndTimeRangeFrom", Err.Number, , , , True
    Exit Sub
End Sub

'for when the endtime does not have the correct dateworked
'given 12/26/2010 18:00 and 12/26/2010 02:00
'return 12/26/2010 18:00 and 12/27/2010 02:00
Private Function DateTimeCrossesMidnight(ByRef r_datStartTime1 As Date, ByRef r_datEndTime1 As Date) As Boolean
    If r_datEndTime1 < r_datStartTime1 And FormatTime(r_datEndTime1) <> "00:00" Then
        r_datEndTime1 = AddDays(1, r_datEndTime1)
        DateTimeCrossesMidnight = True
    Else
        DateTimeCrossesMidnight = False
    End If
End Function

'given 12/26/2010 18:00 02:00
'return/byref 12/26/2010 18:00 - 12/27/2010 00:00 and 12/27/2010 00:00 - 12/27/2010 02:00
Private Function SplitOverMidnight(ByVal v_datChargeDate As Date, ByVal v_strStartTime As String, ByVal v_strEndTime As String, _
                                ByRef r_datStartTime1 As Date, ByRef r_datEndTime1 As Date, _
                                ByRef r_datStartTime2 As Date, ByRef r_datEndTime2 As Date) As Boolean
    Dim l_datStartTime As Date, l_datEndTime As Date

    l_datStartTime = DateTimeFrom(v_datChargeDate, v_strStartTime)
    l_datEndTime = DateTimeFrom(v_datChargeDate, v_strEndTime, True)
    
    If DateTimeCrossesMidnight(l_datStartTime, l_datEndTime) Then
        r_datStartTime1 = l_datStartTime '12/26/2010 18:00
        r_datEndTime1 = MidnightDate(l_datEndTime) '12/27/2010 00:00
        r_datStartTime2 = MidnightDate(l_datEndTime) '12/27/2010 00:00
        r_datEndTime2 = l_datEndTime '12/27/2010 02:00
    
        SplitOverMidnight = True
    Else
        r_datStartTime1 = l_datStartTime
        r_datEndTime1 = l_datEndTime
        r_datStartTime2 = #1/1/1899#
        r_datEndTime2 = #1/1/1899#
        
        SplitOverMidnight = False
    End If
    
End Function

Public Function DateIntersection2( _
    ByVal v_datStartTime1 As Date, ByVal v_datEndTime1 As Date, _
    ByVal v_datStartTime2 As Date, ByVal v_datEndTime2 As Date) _
        As Integer
'Check for intersection between 2 time periods.
'Answer:
    
On Error GoTo eh

'//check for overlap
If (v_datStartTime1 <= v_datStartTime2 And v_datEndTime1 > v_datStartTime2) Or _
    (v_datStartTime2 <= v_datStartTime1 And v_datStartTime1 < v_datEndTime2) Then
    
    If v_datStartTime1 < v_datStartTime2 Then
        If v_datEndTime1 < v_datEndTime2 Then
'     ST1 ._____________. ET1
'        ST2 ._____________. ET2
            DateIntersection2 = 0
        ElseIf v_datEndTime1 = v_datEndTime2 Then
'     ST1 ._____________. ET1
'        ST2 .__________. ET2
            DateIntersection2 = 1
        ElseIf v_datEndTime1 > v_datEndTime2 Then
'     ST1 ._____________. ET1
'        ST2 ._______. ET2
            DateIntersection2 = 2
        End If
    ElseIf v_datStartTime1 = v_datStartTime2 Then
        If v_datEndTime1 < v_datEndTime2 Then
'     ST1 ._____________. ET1
'     ST2 .________________. ET2
            DateIntersection2 = 3
        ElseIf v_datEndTime1 = v_datEndTime2 Then
'     ST1 ._____________. ET1
'     ST2 ._____________. ET2
            DateIntersection2 = 4
        ElseIf v_datEndTime1 > v_datEndTime2 Then
'     ST1 ._____________. ET1
'     ST2 .__________. ET2
            DateIntersection2 = 5
        End If
    ElseIf v_datStartTime1 > v_datStartTime2 Then
        If v_datEndTime1 < v_datEndTime2 Then
'     ST1 ._____________. ET1
'  ST2 .___________________. ET2
            DateIntersection2 = 6
        ElseIf v_datEndTime1 = v_datEndTime2 Then
'     ST1 ._____________. ET1
'  ST2 .________________. ET2
            DateIntersection2 = 7
        ElseIf v_datEndTime1 > v_datEndTime2 Then
'     ST1 ._____________. ET1
'  ST2 ._____________. ET2
            DateIntersection2 = 8
        End If
    End If
Else
    DateIntersection2 = -1
End If

    Exit Function
eh:
    TrsRaiseError MODULENAME, "DateIntersection2"
End Function


'//the new
'//NOTE: This is not looking at actual time, so running after entries have moved from sched exc table to actual time table, may have different results
'//TODO: Try using Time instead of Schedule
Private Function DetermineAbsenceStartAndEnd(ByVal vadrsAbsenceTime As ADODB.Recordset _
                    , ByVal vadrsSchedule As ADODB.Recordset _
                    , ByVal vadrsAllExceptions As ADODB.Recordset _
                    , ByVal vadrsLockStatus As ADODB.Recordset _
                    , ByVal vadrsTime As ADODB.Recordset _
                    ) As ADODB.Recordset
On Error GoTo eh

    Dim ldtMaxPayrollProcessedDate As Date
    'First/Last Processed Record

    ldtMaxPayrollProcessedDate = #1/1/100#
    '//blnIsGetAdjustedBaseSchedule = FieldExists("SYSTEM_EXC_CD", vadrsSchedule)
    
    '//KEEP:
    If (RecordsExist(vadrsLockStatus)) Then
        With vadrsLockStatus
            .Filter = adFilterNone

            .Filter = "LOCK_STATUS = 'Y'"
            If Not (.EOF) Then
                .Sort = "CHARGE_DATE DESC"
                .MoveFirst

                ldtMaxPayrollProcessedDate = !CHARGE_DATE
            End If

            .Filter = adFilterNone
        End With
    End If

    Set DetermineAbsenceStartAndEnd = DetermineAbsenceStartAndEnd_Private(vadrsAbsenceTime, vadrsSchedule, vadrsAllExceptions, ldtMaxPayrollProcessedDate, vadrsTime)

    Exit Function
eh:
    TrsRaiseError MODULENAME, "DetermineAbsenceStartAndEnd"
End Function

'//for testing, vadrsAbsenceTime is result of AddTimeAndSchExceptions
'//TODO: change Public to Private
Public Function DetermineAbsenceStartAndEnd2(ByVal vlEmployeeNumber As Long _
                                            , ByVal vdtStartDateTime As Date _
                                            , ByVal vdtEndDateTime As Date _
                                            , ByVal vadrsAbsenceTime As ADODB.Recordset _
                                            , ByVal vdtMaxPayProcessedDate As Date _
                                            ) As ADODB.Recordset
On Error GoTo eh

    Dim ladrsSchedule As ADODB.Recordset
    Dim ladrsTime As ADODB.Recordset
    Dim ladrsAllExceptions As ADODB.Recordset
    Dim lobjdbCodeTables As dbCodeTables.cdCodeTables
    
    Set ladrsSchedule = GetEmployeeSchedule2(vlEmployeeNumber, vdtStartDateTime, vdtEndDateTime)
    Set ladrsTime = GetEmployeeTimeDetails2(vlEmployeeNumber, vdtStartDateTime, vdtEndDateTime)
    
    Set lobjdbCodeTables = New dbCodeTables.cdCodeTables
    Set ladrsAllExceptions = lobjdbCodeTables.GetAllExceptions
    Set lobjdbCodeTables = Nothing
    
    Set DetermineAbsenceStartAndEnd2 = DetermineAbsenceStartAndEnd_Private(vadrsAbsenceTime, ladrsSchedule, ladrsAllExceptions, vdtMaxPayProcessedDate, ladrsTime)

    Exit Function
eh:
    Set lobjdbCodeTables = Nothing
    TrsRaiseError MODULENAME, "DetermineAbsenceStartAndEnd2"
End Function


Private Function DetermineAbsenceStartAndEnd_Private(ByVal vadrsAbsenceTime As ADODB.Recordset _
                    , ByVal vadrsSchedule As ADODB.Recordset _
                    , ByVal vadrsAllExceptions As ADODB.Recordset _
                    , ByVal vdtMaxPayProcessedDate As Date _
                    , ByVal vadrsTime As ADODB.Recordset _
                    ) As ADODB.Recordset
On Error GoTo eh
'Index and Flag
Dim llSchIndex As Long, llExcIndex As Long
Dim llSchCopyIndex As Long, llExcCopyIndex As Long
Dim llDateIndex As Long, lbUpdated As Boolean
'Date/Time Hold Variables
Dim ldtStartDate As Date, lsStartTime As String
Dim ldtEndDate As Date, lsEndTime As String
Dim ldtRtnDate As Date, lsRtnTime As String
Dim ldtSchCompareDate As Date, ldtExcDateCompare As Date
Dim ldtStDateWorked As Date, lbMove As Boolean
Dim lvBookMark As Variant, llTradeCounter As Long
'Dim ldtMaxPayrollProcessedDate As Date
'First/Last Processed Record
Dim ldtFirstStartDate As Date, lsFirstStartTime As String
Dim ldtFirstStDateWorked As Date
Dim ldtLastProcessedEndDate As Date, lsLastProcessedEndTime As String
Dim ldtLastProcessedRtnDate As Date, lsLastProcessedRtnTime As String
'RS Fields
Dim lsFldName As String, llFldIndex As Long
Dim llFldSize As Long, ladFldType As ADODB.DataTypeEnum
Dim ladFldAttrib As ADODB.FieldAttributeEnum
'RS
Dim ladrsAbsExcCopy As ADODB.Recordset
Dim ladrsSchCopy As ADODB.Recordset
Dim ladrsStartAndEnd As ADODB.Recordset
Dim ladrsStartEndCopy As ADODB.Recordset
Dim blnIsGetAdjustedBaseSchedule As Boolean

'    ldtMaxPayrollProcessedDate = #1/1/100#
    blnIsGetAdjustedBaseSchedule = FieldExists("SYSTEM_EXC_CD", vadrsSchedule)
    
    'Create a copy of the Schedule
    Set ladrsSchCopy = New ADODB.Recordset
    With vadrsSchedule
        For llSchIndex = 0 To .Fields.Count - 1
            'Get the fld name, size
            lsFldName = .Fields.Item(llSchIndex).Name
            ladFldType = .Fields.Item(llSchIndex).Type
            llFldSize = .Fields.Item(llSchIndex).DefinedSize
            ladFldAttrib = .Fields.Item(llSchIndex).Attributes

            If (llFldSize > 0) Then
                ladrsSchCopy.Fields.Append lsFldName, ladFldType _
                    , llFldSize, ladFldAttrib
            Else
                ladrsSchCopy.Fields.Append lsFldName, ladFldType _
                    , , ladFldAttrib
            End If
        Next llSchIndex

        ladrsSchCopy.CursorLocation = adUseClient
        ladrsSchCopy.LockType = adLockBatchOptimistic
        ladrsSchCopy.CursorType = adOpenStatic
        ladrsSchCopy.Open

        .MoveFirst
        For llSchIndex = 1 To .RecordCount
            ladrsSchCopy.AddNew

            For llFldIndex = 0 To .Fields.Count - 1
                ladrsSchCopy.Fields.Item(llFldIndex).Value _
                    = .Fields.Item(llFldIndex).Value
            Next llFldIndex

            ladrsSchCopy.Update

            .MoveNext
        Next llSchIndex

        .MoveFirst
    End With

    'Create a copy of the Absence Exceptions
    Set ladrsAbsExcCopy = New ADODB.Recordset
    With vadrsAbsenceTime
        For llExcIndex = 0 To .Fields.Count - 1
            'Get the fld name, size
            lsFldName = .Fields.Item(llExcIndex).Name
            ladFldType = .Fields.Item(llExcIndex).Type
            llFldSize = .Fields.Item(llExcIndex).DefinedSize
            ladFldAttrib = .Fields.Item(llExcIndex).Attributes

            If (llFldSize > 0) Then
                ladrsAbsExcCopy.Fields.Append lsFldName, ladFldType _
                    , llFldSize, ladFldAttrib
            Else
                ladrsAbsExcCopy.Fields.Append lsFldName, ladFldType _
                    , , ladFldAttrib
            End If
        Next llExcIndex

        ladrsAbsExcCopy.CursorLocation = adUseClient
        ladrsAbsExcCopy.LockType = adLockBatchOptimistic
        ladrsAbsExcCopy.CursorType = adOpenStatic
        ladrsAbsExcCopy.Open

        .MoveLast
        For llExcIndex = .RecordCount To 1 Step -1
            If (Trim(!EXCEPTION_CD) <> "") Or (Trim(!OVER_EXC_CD) <> "") Then
                ladrsAbsExcCopy.AddNew
                For llFldIndex = 0 To .Fields.Count - 1
                    ladrsAbsExcCopy.Fields.Item(llFldIndex).Value _
                        = .Fields.Item(llFldIndex).Value
                Next llFldIndex
                ladrsAbsExcCopy.Update
            Else
                If (Trim(!REC_STATUS) = "") Then
                    .Delete adAffectCurrent
                    .Update
                End If
            End If

            .MovePrevious
        Next llExcIndex

        .MoveFirst
    End With

    'Create a recordset to hold Start, End  and Return information
    Set ladrsStartAndEnd = New ADODB.Recordset
    With ladrsStartAndEnd
        .Fields.Append "ABS_ST_DT", adDate, , ADOFIELDATTR
        .Fields.Append "ST_DATE_WORKED", adDate, , ADOFIELDATTR
        .Fields.Append "ABS_ST_TM", adChar, 5, ADOFIELDATTR

        .Fields.Append "ABS_END_DT", adDate, , ADOFIELDATTR
        .Fields.Append "END_DATE_WORKED", adDate, , ADOFIELDATTR
        .Fields.Append "ABS_END_TM", adChar, 5, ADOFIELDATTR
        
        .Fields.Append "RTN_TO_WRK_DT", adDate, , ADOFIELDATTR
        .Fields.Append "RTW_DATE_WORKED", adDate, , ADOFIELDATTR
        .Fields.Append "RTN_TO_WRK_TM", adChar, 5, ADOFIELDATTR

        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open
    End With

    'Loop through the absence exceptions and schedule to determine
    'the end and return information
    With vadrsAbsenceTime
        .Sort = "CHARGE_DATE, DATE_WORKED, START_TIME, END_TIME"
        .MoveFirst
        
        For llExcIndex = 1 To .RecordCount
            If (!REC_STATUS <> "D") Then
                ladrsStartAndEnd.AddNew
                ladrsStartAndEnd!ABS_ST_DT = !CHARGE_DATE
                ladrsStartAndEnd!ABS_ST_TM = !START_TIME
                ladrsStartAndEnd!ST_DATE_WORKED = !DATE_WORKED
                ladrsStartAndEnd!ABS_END_DT = !CHARGE_DATE
                ladrsStartAndEnd!ABS_END_TM = !END_TIME
                ladrsStartAndEnd!END_DATE_WORKED = MidnightDate(FindDateAndTimeWorked(ladrsStartAndEnd!ABS_END_DT, ladrsStartAndEnd!ABS_END_TM, vadrsTime))
                '//ladrsStartAndEnd!RTW_DATE_WORKED = MidnightDate(FindDateAndTimeWorked(ladrsStartAndEnd!RTN_TO_WRK_DT, ladrsStartAndEnd!RTN_TO_WRK_TM, vadrsTime))
                ladrsStartAndEnd.Update

'                '//sched exc is one row 12/26 12/26 18:00 02:00, whereas sched is two rows split over midnight
'                '//12/26 12/26 18:00 00:00 and 12/26 12/27 00:00 02:00
'                '//TODO: refactor to TimeCrossesMidnight byref date1, time1, byref date2, time2
                Dim ldtStartTime As Date, ldtEndTime As Date
'
''                ldtStartTime = DateTimeFrom(!CHARGE_DATE, !START_TIME)
''                ldtEndTime = DateTimeFrom(!DATE_WORKED, !END_TIME, True)
''                DateTimeCrossesMidnight ldtStartTime, ldtEndTime '//add 1 day to date worked if time spans midnight
'
                DateAndTimeRangeFrom !DATE_WORKED, !START_TIME, _
                                    !DATE_WORKED, !END_TIME, _
                                    ldtStartTime, ldtEndTime, True
'                ladrsStartAndEnd!ABS_END_DT = MidnightDate(ldtEndTime)
'                ladrsStartAndEnd!ABS_END_TM = FormatTime(ldtEndTime)
'                ladrsStartAndEnd.Update
                
                vadrsSchedule.Filter = adFilterNone
                '//get adjusted base schedule
                If blnIsGetAdjustedBaseSchedule Then
                    vadrsSchedule.Filter = ("SHIFT_HOURS > 0.0 AND SYSTEM_EXC_CD <> 'SLT'" & _
                                        " AND EXCEPTION_CD <> 'ONP'" & _
                                        " AND CHARGE_DATE >= #" & !CHARGE_DATE & "#" & _
                                        " AND DATE_WORKED >= #" & !DATE_WORKED & "#")
                Else '//view calendar details
                    '//SHIFT_HOURS is not on the Time recordset, so use hours instead
                    vadrsSchedule.Filter = "HOURS > 0.0" & _
                                        " AND EXCEPTION_CD <> 'ONP'" & _
                                        " AND CHARGE_DATE >= #" & !CHARGE_DATE & "#" & _
                                        " AND DATE_WORKED >= #" & !DATE_WORKED & "#"
                End If
                                    
Debug.Print ""
Debug.Print PadRight(StringFormat("vadrsAbsenceTime[{0}]", llExcIndex), 30) & StringFormat("= {0} {1} {2} {3} {4}", vadrsAbsenceTime!CHARGE_DATE, vadrsAbsenceTime!DATE_WORKED, vadrsAbsenceTime!START_TIME, vadrsAbsenceTime!END_TIME, vadrsAbsenceTime!EXCEPTION_CD)
                If Not (vadrsSchedule.EOF) Then
                    vadrsSchedule.Sort = "CHARGE_DATE, DATE_WORKED, START_TIME, END_TIME"
                    vadrsSchedule.MoveFirst

                    For llSchIndex = 1 To vadrsSchedule.RecordCount
Debug.Print PadRight(StringFormat("   vadrsSchedule[{0}]", llSchIndex), 30) & StringFormat("= {0} {1} {2} {3} {4}", vadrsSchedule!CHARGE_DATE, vadrsSchedule!DATE_WORKED, vadrsSchedule!START_TIME, vadrsSchedule!END_TIME, vadrsSchedule!EXCEPTION_CD)
                        lbUpdated = False
                        '//TODO: Does this need to be specified as an end time, since it is a dateworked ???
                        '//the date worked from the schedule really applies to the start time, it will be broken over midnight, so go ahead and convert for and end time = "00:00"
                        If ldtEndTime < DateTimeFrom(vadrsSchedule!DATE_WORKED, vadrsSchedule!END_TIME, True) Then
'//this is failing for 12/24 12/24 18:00 22:00 and with shift 18:00 06:00
'                            If (CDate(Format(ldtExcDateCompare & "  " & !END_TIME, "MM/DD/YYYY HH:MM")) < _
'                                CDate(Format(ldtSchCompareDate & "  " & vadrsSchedule!SHIFT_END_TIME, "MM/DD/YYYY HH:MM"))) Then
'TODO: This (exchdrid 133416) is not splitting in test with exc ST because ST is not on schedule, why was the absence split correctly in prod?
                            If (Trim(vadrsSchedule!OVER_EXC_CD) <> "" Or _
                                Trim(vadrsSchedule!EXCEPTION_CD) <> "") Then
                                vadrsAllExceptions.Filter = adFilterNone

                                vadrsAllExceptions.Filter = ("(EXCEPTION_CD = '" & _
                                                    Trim(vadrsSchedule!OVER_EXC_CD) & _
                                                    "' AND EXCEPTION_BANK_CD <> 'O')" & _
                                                    " OR (EXCEPTION_CD = '" & _
                                                    Trim(vadrsSchedule!EXCEPTION_CD) & _
                                                    "' AND EXCEPTION_BANK_CD <> 'O')")
                                If Not (vadrsAllExceptions.EOF) Then
                                    vadrsAllExceptions.MoveFirst

                                    vadrsAllExceptions.Find "WORK_REST_IND = 'R'" '//rest (non-working day)
                                    If (vadrsAllExceptions.EOF) Then '//working day
                                        If (RecordsExist(ladrsStartAndEnd)) Then
                                                ladrsStartAndEnd!RTN_TO_WRK_DT = vadrsSchedule!CHARGE_DATE
                                                ladrsStartAndEnd!RTN_TO_WRK_TM = vadrsSchedule!START_TIME
                                                ladrsStartAndEnd!RTW_DATE_WORKED = MidnightDate(FindDateAndTimeWorked(ladrsStartAndEnd!RTN_TO_WRK_DT, ladrsStartAndEnd!RTN_TO_WRK_TM, vadrsTime))
                                                
                                                ladrsStartAndEnd.Update

Debug.Print PadRight("return to work 1", 30) & StringFormat("= {0} {1} {2} {3}", ladrsStartAndEnd!RTN_TO_WRK_DT, Space(10), ladrsStartAndEnd!RTN_TO_WRK_TM, Space(5))
                                                vadrsAllExceptions.Filter = adFilterNone

                                                lbUpdated = True
                                                Exit For
                                        End If
                                    Else '//non-working day
                                    End If
                                End If

                                vadrsAllExceptions.Filter = adFilterNone
                            Else '//vadrsSchedule!OVER_EXC_CD="" And vadrsSchedule!EXCEPTION_CD=""
                                If (RecordsExist(ladrsStartAndEnd)) Then
                                        ladrsStartAndEnd!RTN_TO_WRK_DT = vadrsSchedule!CHARGE_DATE
                                        ladrsStartAndEnd!RTN_TO_WRK_TM = !END_TIME

                                    '//the schedule endtime matches the absence time, so use the next schedule row
                                    '//TODO: Does this need to be specified as an end time, since it is a dateworked ???
                                    '//the date worked from the schedule really applies to the start time, it will be broken over midnight, so go ahead and convert for and end time = "00:00"
                                    If ldtEndTime = DateTimeFrom(vadrsSchedule!DATE_WORKED, vadrsSchedule!END_TIME, True) Then
                                        '//TODO: will this ever get executed???
                                        vadrsSchedule.MoveNext

                                        If Not (vadrsSchedule.EOF) Then
                                            ladrsStartAndEnd!RTN_TO_WRK_DT = vadrsSchedule!CHARGE_DATE
                                            ladrsStartAndEnd!RTN_TO_WRK_TM = vadrsSchedule!START_TIME
                                        End If

                                        vadrsSchedule.MovePrevious
                                    Else
                                        '//the schedule is entirely beyond the absence time, so use its start time
                                        '//TODO: Does this need to be specified as an end time, since it is a dateworked ???
                                        '//the date worked from the schedule really applies to the start time, it will be broken over midnight, so go ahead and convert for and end time = "00:00"
                                        If ldtEndTime < DateTimeFrom(vadrsSchedule!DATE_WORKED, vadrsSchedule!END_TIME, True) And _
                                            DateTimeFrom(vadrsSchedule!DATE_WORKED, vadrsSchedule!START_TIME) >= ldtEndTime Then
                                            ladrsStartAndEnd!RTN_TO_WRK_DT = vadrsSchedule!CHARGE_DATE
                                            ladrsStartAndEnd!RTN_TO_WRK_TM = vadrsSchedule!START_TIME
                                        End If
                                    End If
Debug.Print PadRight("return to work 2", 30) & StringFormat("= {0} {1} {2} {3}", ladrsStartAndEnd!RTN_TO_WRK_DT, Space(10), ladrsStartAndEnd!RTN_TO_WRK_TM, Space(5))
                                    ladrsStartAndEnd!RTW_DATE_WORKED = MidnightDate(FindDateAndTimeWorked(ladrsStartAndEnd!RTN_TO_WRK_DT, ladrsStartAndEnd!RTN_TO_WRK_TM, vadrsTime))
                                    
                                    ladrsStartAndEnd.Update

                                    lbUpdated = True
                                    Exit For
                                End If
                            End If
                        Else '//ldtEndTime < DateTimeFrom(vadrsSchedule!DATE_WORKED, vadrsSchedule!END_TIME, True)
                        End If '//ldtEndTime < DateTimeFrom(vadrsSchedule!DATE_WORKED, vadrsSchedule!END_TIME, True)
        
                        If (lbUpdated) Then
                            vadrsSchedule.Filter = adFilterNone
                            Exit For
                        End If

                        vadrsSchedule.MoveNext
                        DoEvents
                    Next llSchIndex '//vadrsSchedule
                End If '//Not (vadrsSchedule.EOF)

                vadrsSchedule.Filter = adFilterNone
                    
                '//HACK: Return to Work may not be getting set and may be null
                If IsNull(ladrsStartAndEnd!RTN_TO_WRK_DT) Then
                    ladrsStartAndEnd!RTN_TO_WRK_DT = AddDays(1, MidnightDate(ldtEndTime)) '//tomorrow
                    ladrsStartAndEnd!RTN_TO_WRK_TM = "00:00" 'FormatTime(ldtEndTime, True)
                    ladrsStartAndEnd!RTW_DATE_WORKED = MidnightDate(FindDateAndTimeWorked(ladrsStartAndEnd!RTN_TO_WRK_DT, ladrsStartAndEnd!RTN_TO_WRK_TM, vadrsTime))
    
                    ladrsStartAndEnd.Update
                End If
            End If '//(!REC_STATUS <> "D")
            
            .MoveNext
            DoEvents
        Next llExcIndex '//vadrsAbsenceTime
    End With

    Set ladrsAbsExcCopy = Nothing
    Set ladrsSchCopy = Nothing

    '//for testing only
'    Set madrsAbsenceStartEndDetails = CopyRecordset(ladrsStartAndEnd)
    
    'Determine the end date (range) for absence dates
    If (RecordsExist(ladrsStartAndEnd)) Then
        Set ladrsStartEndCopy = New ADODB.Recordset
        With ladrsStartEndCopy
            .Fields.Append "ABS_ST_DT", adDate, , ADOFIELDATTR
            .Fields.Append "ST_DATE_WORKED", adDate, , ADOFIELDATTR
            .Fields.Append "ABS_ST_TM", adChar, 5, ADOFIELDATTR

            .Fields.Append "ABS_END_DT", adDate, , ADOFIELDATTR
            .Fields.Append "END_DATE_WORKED", adDate, , ADOFIELDATTR
            .Fields.Append "ABS_END_TM", adChar, 5, ADOFIELDATTR
            
            .Fields.Append "RTN_TO_WRK_DT", adDate, , ADOFIELDATTR
            .Fields.Append "RTW_DATE_WORKED", adDate, , ADOFIELDATTR
            .Fields.Append "RTN_TO_WRK_TM", adChar, 5, ADOFIELDATTR

            .CursorLocation = adUseClient
            .CursorType = adOpenStatic
            .LockType = adLockBatchOptimistic
            .Open
        End With

        With ladrsStartAndEnd
            '//collapse rows from the beginning to the payroll processed date (use start dates from 1st row, use end and return dates from last row)
            '//this does not look for any gaps or overlaps between the return to work date and the start date on the next row
            '//this could cause an absence to not be split when it should (see ExecHdrId=68168, no row for 1/31/2011)
            .MoveFirst

            ldtFirstStartDate = !ABS_ST_DT
            lsFirstStartTime = !ABS_ST_TM
            ldtFirstStDateWorked = !ST_DATE_WORKED

            .Filter = adFilterNone
            
            .Filter = "ABS_ST_DT <= #" & vdtMaxPayProcessedDate & "#"
            If Not (.EOF) Then
                .Sort = "ST_DATE_WORKED, ABS_ST_DT, ABS_ST_TM, END_DATE_WORKED, ABS_END_DT, ABS_END_TM"
                .MoveLast

                ldtLastProcessedEndDate = !ABS_END_DT
                lsLastProcessedEndTime = !ABS_END_TM
                ldtLastProcessedRtnDate = !RTN_TO_WRK_DT
                lsLastProcessedRtnTime = !RTN_TO_WRK_TM

                .MoveFirst
                For llDateIndex = 1 To .RecordCount
                    .Delete adAffectCurrent

                    If Not (.EOF) Then
                        .MoveNext
                    End If
                    DoEvents
                Next llDateIndex

                .Filter = adFilterNone

                .AddNew
                !ABS_ST_DT = ldtFirstStartDate
                !ABS_ST_TM = lsFirstStartTime
                !ST_DATE_WORKED = ldtFirstStDateWorked
                !ABS_END_DT = ldtLastProcessedEndDate
                !ABS_END_TM = lsLastProcessedEndTime
                !RTN_TO_WRK_DT = ldtLastProcessedRtnDate
                !RTN_TO_WRK_TM = lsLastProcessedRtnTime
                ladrsStartAndEnd!END_DATE_WORKED = MidnightDate(FindDateAndTimeWorked(ladrsStartAndEnd!ABS_END_DT, ladrsStartAndEnd!ABS_END_TM, vadrsTime))
                ladrsStartAndEnd!RTW_DATE_WORKED = MidnightDate(FindDateAndTimeWorked(ladrsStartAndEnd!RTN_TO_WRK_DT, ladrsStartAndEnd!RTN_TO_WRK_TM, vadrsTime))
                .Update
            End If

            .Filter = adFilterNone
            .MoveFirst

            For llDateIndex = 1 To .RecordCount
                ladrsStartEndCopy.AddNew

                For llFldIndex = 0 To .Fields.Count - 1
                    ladrsStartEndCopy.Fields.Item(llFldIndex).Value = .Fields.Item(llFldIndex).Value
                Next llFldIndex

                ladrsStartEndCopy.Update

                .MoveNext
            Next llDateIndex

            'Check for Double Shift/Owl
            'If Double Shift/Owl exists, flatten.
            '//added END_DATE_WORKED in the sort order, so that the row remains in the correct order, after the start date/time is updated
            ladrsStartEndCopy.Sort = "ST_DATE_WORKED, ABS_ST_DT, ABS_ST_TM, END_DATE_WORKED, ABS_END_DT, ABS_END_TM"
            ladrsStartEndCopy.MoveFirst

            .Sort = "ST_DATE_WORKED, ABS_ST_DT, ABS_ST_TM, END_DATE_WORKED, ABS_END_DT, ABS_END_TM"
            .MoveFirst

            llTradeCounter = 0
            Do
                ladrsStartEndCopy.Filter = adFilterNone

                ladrsStartEndCopy.Filter = "ABS_ST_DT = #" & !ABS_ST_DT & "#"
                If Not (ladrsStartEndCopy.EOF) Then
                    If (ladrsStartEndCopy.RecordCount > 1) Then
                        'Potential Double Shift
                        llTradeCounter = llTradeCounter + 1

                        ladrsStartEndCopy.MoveFirst

                        ldtStartDate = ladrsStartEndCopy!ABS_ST_DT
                        lsStartTime = ladrsStartEndCopy!ABS_ST_TM
                        ldtEndDate = ladrsStartEndCopy!ABS_END_DT
                        lsEndTime = ladrsStartEndCopy!ABS_END_TM
                        ldtRtnDate = ladrsStartEndCopy!RTN_TO_WRK_DT
                        lsRtnTime = ladrsStartEndCopy!RTN_TO_WRK_TM
                        ldtStDateWorked = ladrsStartEndCopy!ST_DATE_WORKED

                        ladrsStartEndCopy.MoveNext
                        If Not (ladrsStartEndCopy.EOF) Then
                            If (ladrsStartEndCopy!ABS_ST_DT = ldtRtnDate And ladrsStartEndCopy!ABS_ST_TM = lsRtnTime) Or _
                                (ladrsStartEndCopy!ABS_END_DT = ldtEndDate And ladrsStartEndCopy!ABS_END_TM = lsEndTime And _
                                DateTimeFrom(ladrsStartEndCopy!ST_DATE_WORKED, ladrsStartEndCopy!ABS_END_TM) > CDate(ldtStDateWorked & " " & lsStartTime)) Then
                                '//since this is already date worked, (don't automatically adjust to the next date for midnight, since it has already been done)
                                ladrsStartEndCopy!ABS_ST_DT = ldtStartDate
                                ladrsStartEndCopy!ABS_ST_TM = lsStartTime
                                ladrsStartEndCopy!ST_DATE_WORKED = ldtStDateWorked
                                ladrsStartEndCopy.Update

                                ladrsStartEndCopy.MovePrevious
                                '//TODO: Why is now happening (BOF error) and what to do on else ???
                                '//If the end time is on or after midnight (charge dcate <> date worked)
                                '//updating the start time could result in the row being repositioned (end time="00:00")
                                '//Then, if the row is at position 1, doing a move previous will end up resulting in a BOF error
                                '//need to include END_DATE_WORKED in the sort order, so that the row remains in the correct order
                                If Not (ladrsStartEndCopy.BOF) Then
                                    ladrsStartEndCopy.Delete adAffectCurrent
                                    ladrsStartEndCopy.Update
                                End If

                                If Not (ladrsStartEndCopy.EOF) Then
                                    ladrsStartEndCopy.MoveNext
                                End If
                            End If
                        End If

                        ladrsStartEndCopy.Filter = adFilterNone

                        ldtStartDate = !ABS_ST_DT
                        lsStartTime = !ABS_ST_TM
                        ldtEndDate = !ABS_END_DT
                        lsEndTime = !ABS_END_TM
                        ldtRtnDate = !RTN_TO_WRK_DT
                        lsRtnTime = !RTN_TO_WRK_TM
                        ldtStDateWorked = !ST_DATE_WORKED

                        .MoveNext
                        If Not (.EOF) Then
                            If (!ABS_ST_DT = ldtRtnDate And !ABS_ST_TM = lsRtnTime) Or _
                                (!ABS_END_DT = ldtEndDate And !ABS_END_TM = lsEndTime And _
                                DateTimeFrom(!ST_DATE_WORKED, !ABS_END_TM) > CDate(ldtStDateWorked & " " & lsStartTime)) Then
                                '//since this is already date worked, (don't automatically adjust to the next date for midnight, since it has already been done)
                                !ABS_ST_DT = ldtStartDate
                                !ABS_ST_TM = lsStartTime
                                !ST_DATE_WORKED = ldtStDateWorked
                                .Update

                                .MovePrevious
                                '//TODO: Why is this BOF and what to do on else ???
                                If Not .BOF Then
                                    .Delete adAffectCurrent
                                    .Update
                                End If
                            End If
                        End If
                    End If
                End If

                ladrsStartEndCopy.Filter = adFilterNone

                If Not (.EOF) Then
                    .MoveNext
                End If
            Loop Until (.EOF)

            'Flatten dates with return to work falling between trade flattened date
            llTradeCounter = llTradeCounter * 2

            For llExcIndex = 1 To llTradeCounter
'Debug.Print StringFormat("{0} of {1} TradeCounter Loop", llExcIndex, llTradeCounter)
                ladrsStartEndCopy.Sort = "ST_DATE_WORKED, ABS_ST_DT, ABS_ST_TM, END_DATE_WORKED, ABS_END_DT, ABS_END_TM"
                ladrsStartEndCopy.MoveFirst

                .Sort = "ST_DATE_WORKED, ABS_ST_DT, ABS_ST_TM, END_DATE_WORKED, ABS_END_DT, ABS_END_TM"

                For llSchCopyIndex = 1 To ladrsStartEndCopy.RecordCount
'Debug.Print StringFormat("   {0} of {1} ladrsStartEndCopy Loop1", llSchCopyIndex, ladrsStartEndCopy.RecordCount)
                    .MoveFirst

                    For llDateIndex = 1 To .RecordCount
'Debug.Print StringFormat("      {0} of {1} DateIndex Loop 1", llDateIndex, .RecordCount)
                        lbMove = False

                        If CDate(!ST_DATE_WORKED & " " & !ABS_ST_TM) < CDate(ladrsStartEndCopy!RTN_TO_WRK_DT & " " & ladrsStartEndCopy!RTN_TO_WRK_TM) And _
                            DateTimeFrom(!ABS_END_DT, !ABS_END_TM, True) > CDate(ladrsStartEndCopy!RTN_TO_WRK_DT & " " & ladrsStartEndCopy!RTN_TO_WRK_TM) And _
                            CDate(!ST_DATE_WORKED & " " & !ABS_ST_TM) <= CDate(ladrsStartEndCopy!ST_DATE_WORKED & " " & ladrsStartEndCopy!ABS_ST_TM) And _
                            DateTimeFrom(!ABS_END_DT, !ABS_END_TM, True) >= DateTimeFrom(ladrsStartEndCopy!ABS_END_DT, ladrsStartEndCopy!ABS_END_TM, True) And _
                            CDate(!RTN_TO_WRK_DT & " " & !RTN_TO_WRK_TM) >= CDate(ladrsStartEndCopy!RTN_TO_WRK_DT & " " & ladrsStartEndCopy!RTN_TO_WRK_TM) Then
                            'Set the original to vars.
                            ldtEndDate = !ABS_END_DT
                            lsEndTime = !ABS_END_TM
                            ldtRtnDate = !RTN_TO_WRK_DT
                            lsRtnTime = !RTN_TO_WRK_TM

                            lvBookMark = .Bookmark

                            .Filter = adFilterNone
                            .Filter = "ABS_ST_DT = #" & ladrsStartEndCopy!ABS_ST_DT & "#" & _
                                " AND ABS_ST_TM = '" & ladrsStartEndCopy!ABS_ST_TM & "'" & _
                                " AND ST_DATE_WORKED = #" & ladrsStartEndCopy!ST_DATE_WORKED & "#" & _
                                " AND ABS_END_DT = #" & ladrsStartEndCopy!ABS_END_DT & "#" & _
                                " AND ABS_END_TM = '" & ladrsStartEndCopy!ABS_END_TM & "'"
                            If Not (.EOF) Then
                                !ABS_END_DT = ldtEndDate
                                !ABS_END_TM = lsEndTime
                                !RTN_TO_WRK_DT = ldtRtnDate
                                !RTN_TO_WRK_TM = lsRtnTime
                                ladrsStartAndEnd!END_DATE_WORKED = MidnightDate(FindDateAndTimeWorked(ladrsStartAndEnd!ABS_END_DT, ladrsStartAndEnd!ABS_END_TM, vadrsTime))
                                ladrsStartAndEnd!RTW_DATE_WORKED = MidnightDate(FindDateAndTimeWorked(ladrsStartAndEnd!RTN_TO_WRK_DT, ladrsStartAndEnd!RTN_TO_WRK_TM, vadrsTime))
                                
                                .Update
                            End If

                            .Filter = adFilterNone

                            .Bookmark = lvBookMark
                            If Not (.EOF) Then
                                ladrsStartEndCopy!ABS_END_DT = ldtEndDate
                                ladrsStartEndCopy!ABS_END_TM = lsEndTime
                                ladrsStartEndCopy!RTN_TO_WRK_DT = ldtRtnDate
                                ladrsStartEndCopy!RTN_TO_WRK_TM = lsRtnTime
                                '//TODO: update the dateworked for end and return dates
                                ladrsStartEndCopy.Update

                                If Not (ladrsStartEndCopy.EOF) Then
                                    lvBookMark = ladrsStartEndCopy.Bookmark
                                    ladrsStartEndCopy.Filter = adFilterNone

                                    ladrsStartEndCopy.Filter = "ABS_ST_DT = #" & !ABS_ST_DT & "#" & _
                                                        " AND ABS_ST_TM = '" & !ABS_ST_TM & "'" & _
                                                        " AND ST_DATE_WORKED = #" & !ST_DATE_WORKED & "#" & _
                                                        " AND ABS_END_DT = #" & !ABS_END_DT & "#" & _
                                                        " AND ABS_END_TM = '" & !ABS_END_TM & "'"

                                    If Not (ladrsStartEndCopy.EOF) Then
                                        ladrsStartEndCopy.Delete adAffectCurrent
                                        ladrsStartEndCopy.Update
                                        lbMove = True
                                    End If

                                    ladrsStartEndCopy.Filter = adFilterNone

                                    ladrsStartEndCopy.Bookmark = lvBookMark
                                End If

                                .Delete adAffectCurrent
                                .Update
                            End If
                        End If

                        If Not (.EOF) Then
                            .MoveNext

                            If (.EOF) Then
                                Exit For
                            End If
                        Else
                            Exit For
                        End If

                        If (lbMove) Then
                            If Not (ladrsStartEndCopy.EOF) Then
                                ladrsStartEndCopy.MoveNext
                                If (ladrsStartEndCopy.EOF) Then
                                    Exit For
                                End If
                            Else
                                Exit For
                            End If
                        End If
                        DoEvents
                    Next llDateIndex

                    If Not (ladrsStartEndCopy.EOF) Then
                        ladrsStartEndCopy.MoveNext
                        If (ladrsStartEndCopy.EOF) Then
                            Exit For
                        End If
                    Else
                        Exit For
                    End If
                    DoEvents
                Next llSchCopyIndex
'Debug.Print "entering botton loop"
                ladrsStartEndCopy.Sort = "ST_DATE_WORKED, ABS_ST_DT, ABS_ST_TM, END_DATE_WORKED, ABS_END_DT, ABS_END_TM"
                .Sort = "ST_DATE_WORKED, ABS_ST_DT, ABS_ST_TM, END_DATE_WORKED, ABS_END_DT, ABS_END_TM"
                .MoveFirst

                For llDateIndex = 1 To .RecordCount
'Debug.Print StringFormat("   {0} of {1} DateLoop2", llDateIndex, .RecordCount)
                    ladrsStartEndCopy.MoveFirst

                    For llSchCopyIndex = 1 To ladrsStartEndCopy.RecordCount
'Debug.Print StringFormat("      {0} of {1} ladrsStartEndCopy2", llDateIndex, .RecordCount)
                        lbMove = False

                        If CDate(!RTN_TO_WRK_DT & " " & !RTN_TO_WRK_TM) > CDate(ladrsStartEndCopy!RTN_TO_WRK_DT & " " & ladrsStartEndCopy!RTN_TO_WRK_TM) And _
                            CDate(!ST_DATE_WORKED & " " & !ABS_ST_TM) <= CDate(ladrsStartEndCopy!ST_DATE_WORKED & " " & ladrsStartEndCopy!ABS_ST_TM) And _
                            DateTimeFrom(!ABS_END_DT, !ABS_END_TM, True) >= DateTimeFrom(ladrsStartEndCopy!ABS_END_DT, ladrsStartEndCopy!ABS_END_TM, True) Then
                            lvBookMark = .Bookmark

                            .Filter = adFilterNone
                            .Filter = ("ABS_ST_DT = #" & ladrsStartEndCopy!ABS_ST_DT & "#" & _
                                " AND ABS_ST_TM = '" & ladrsStartEndCopy!ABS_ST_TM & "'" & _
                                " AND ABS_END_DT = #" & ladrsStartEndCopy!ABS_END_DT & "#" & _
                                " AND ABS_END_TM = '" & ladrsStartEndCopy!ABS_END_TM & "'" & _
                                " AND RTN_TO_WRK_DT = #" & ladrsStartEndCopy!RTN_TO_WRK_DT & "#" & _
                                " AND RTN_TO_WRK_TM = '" & ladrsStartEndCopy!RTN_TO_WRK_TM & "'" & _
                                " AND ST_DATE_WORKED = #" & ladrsStartEndCopy!ST_DATE_WORKED & "#")
                            If Not (.EOF) Then
                                .Delete adAffectCurrent
                                .Update

                                ladrsStartEndCopy.Delete adAffectCurrent
                                ladrsStartEndCopy.Update
                                lbMove = True
                            End If

                            .Filter = adFilterNone
                            .Bookmark = lvBookMark
                        End If

                        If Not (ladrsStartEndCopy.EOF) Then
                            ladrsStartEndCopy.MoveNext
                        Else
                            Exit For
                        End If

                        If (lbMove) Then
                            If Not (.EOF) Then
                                .MoveNext

                                If (.EOF) Then
                                    Exit For
                                End If
                            Else
                                Exit For
                            End If
                        End If
                        DoEvents
                    Next llSchCopyIndex

                    If Not (.EOF) Then
                        .MoveNext

                        If (.EOF) Then
                            Exit For
                        End If
                    Else
                        Exit For
                    End If
                    DoEvents
                Next llDateIndex
                DoEvents
            Next llExcIndex

            'Flatten dates with same return to work date or return to work overlapping absence
            For llExcIndex = 1 To ladrsStartEndCopy.RecordCount
                ladrsStartEndCopy.Sort = "ST_DATE_WORKED, ABS_ST_DT, ABS_ST_TM, END_DATE_WORKED, ABS_END_DT, ABS_END_TM"
                ladrsStartEndCopy.MoveFirst

                For llDateIndex = 1 To ladrsStartEndCopy.RecordCount
                    .Filter = adFilterNone

                    .Filter = "(RTN_TO_WRK_DT = #" & ladrsStartEndCopy!RTN_TO_WRK_DT & "#" & _
                                " AND RTN_TO_WRK_TM = '" & ladrsStartEndCopy!RTN_TO_WRK_TM & "')"
                    If Not (.EOF) Then
                        .Sort = "ST_DATE_WORKED, ABS_ST_DT, ABS_ST_TM, END_DATE_WORKED, ABS_END_DT, ABS_END_TM"
                        .MoveLast

                        ldtEndDate = !ABS_END_DT
                        lsEndTime = !ABS_END_TM

                        .MoveFirst
                        !ABS_END_DT = ldtEndDate
                        !ABS_END_TM = lsEndTime
                        ladrsStartAndEnd!END_DATE_WORKED = MidnightDate(FindDateAndTimeWorked(ladrsStartAndEnd!ABS_END_DT, ladrsStartAndEnd!ABS_END_TM, vadrsTime))
                        
                        .Update

                        .MoveNext
                        Do While Not (.EOF)
                            .Delete adAffectCurrent
                            .Update
                            .MoveNext
                        Loop
                    End If

                    .Filter = adFilterNone

                    ladrsStartEndCopy.MoveNext
                    DoEvents
                Next llDateIndex

                'Flatten dates with same return to work date as the start date
                .Sort = "ST_DATE_WORKED, ABS_ST_DT, ABS_ST_TM, END_DATE_WORKED, ABS_END_DT, ABS_END_TM"
                .MoveFirst

                For llDateIndex = 1 To .RecordCount
                    ldtStartDate = !ABS_ST_DT
                    lsStartTime = !ABS_ST_TM
                    ldtEndDate = !ABS_END_DT
                    lsEndTime = !ABS_END_TM
                    ldtRtnDate = !RTN_TO_WRK_DT
                    lsRtnTime = !RTN_TO_WRK_TM
                    ldtStDateWorked = !ST_DATE_WORKED

                    .MoveNext
                    If Not (.EOF) Then
                        If (!ABS_ST_DT = ldtRtnDate And !ABS_ST_TM = lsRtnTime) Or _
                            (CDate(!ST_DATE_WORKED & " " & !ABS_ST_TM) <= CDate(ldtRtnDate & " " & lsRtnTime) And _
                            CDate(ldtRtnDate & " " & lsRtnTime) <= DateTimeFrom(!ABS_END_DT, !ABS_END_TM, True)) Then
                            !ABS_ST_DT = ldtStartDate
                            !ABS_ST_TM = lsStartTime
                            !ST_DATE_WORKED = ldtStDateWorked
                            .Update

                            .MovePrevious
                            .Delete adAffectCurrent
                            .Update

                            If Not (.EOF) Then
                                .MoveNext
                            End If
                        End If
                    End If
                    DoEvents
                Next llDateIndex
                DoEvents
            Next llExcIndex

            ladrsStartEndCopy.Sort = "ST_DATE_WORKED, ABS_ST_DT, ABS_ST_TM, END_DATE_WORKED, ABS_END_DT, ABS_END_TM"
            ladrsStartEndCopy.MoveFirst

            For llDateIndex = 1 To ladrsStartEndCopy.RecordCount
                .Filter = adFilterNone

                .Filter = ("RTN_TO_WRK_DT = #" & ladrsStartEndCopy!RTN_TO_WRK_DT & "#" & _
                            " AND RTN_TO_WRK_TM = '" & ladrsStartEndCopy!RTN_TO_WRK_TM & "'")
                If Not (.EOF) Then
                    .Sort = "ST_DATE_WORKED, ABS_ST_DT, ABS_ST_TM, END_DATE_WORKED, ABS_END_DT, ABS_END_TM"
                    .MoveLast

                    ldtEndDate = !ABS_END_DT
                    lsEndTime = !ABS_END_TM

                    .MoveFirst
                    !ABS_END_DT = ldtEndDate
                    !ABS_END_TM = lsEndTime
                    ladrsStartAndEnd!END_DATE_WORKED = MidnightDate(FindDateAndTimeWorked(ladrsStartAndEnd!ABS_END_DT, ladrsStartAndEnd!ABS_END_TM, vadrsTime))
                    
                    .Update

                    .MoveNext
                    Do While Not (.EOF)
                        .Delete adAffectCurrent
                        .Update
                        .MoveNext
                    Loop
                End If

                .Filter = adFilterNone

                ladrsStartEndCopy.MoveNext
                DoEvents
            Next llDateIndex

            If Not (.EOF And .BOF) Then
                .MoveFirst
            End If
        End With
    End If

    If RecordsExist(ladrsStartAndEnd) Then
        ladrsStartAndEnd.MoveFirst
        Do While Not ladrsStartAndEnd.EOF
            ladrsStartAndEnd!END_DATE_WORKED = MidnightDate(FindDateAndTimeWorked(ladrsStartAndEnd!ABS_END_DT, ladrsStartAndEnd!ABS_END_TM, vadrsTime))
            ladrsStartAndEnd!RTW_DATE_WORKED = MidnightDate(FindDateAndTimeWorked(ladrsStartAndEnd!RTN_TO_WRK_DT, ladrsStartAndEnd!RTN_TO_WRK_TM, vadrsTime))
            ladrsStartAndEnd.Update
        
            ladrsStartAndEnd.MoveNext
        Loop
        ladrsStartAndEnd.MoveFirst
    End If

    Set ladrsStartEndCopy = Nothing

    Set DetermineAbsenceStartAndEnd_Private = ladrsStartAndEnd

    Set ladrsStartAndEnd = Nothing

    Exit Function
eh:
    Set ladrsAbsExcCopy = Nothing
    Set ladrsSchCopy = Nothing

    Set ladrsStartAndEnd = Nothing
    Set ladrsStartEndCopy = Nothing

    TrsRaiseError MODULENAME, "DetermineAbsenceStartAndEnd_Private"
End Function

'//old
Private Function DetermineAbsenceStartAndEndOld(ByVal vadrsAbsenceTime As ADODB.Recordset _
                    , ByVal vadrsSchedule As ADODB.Recordset _
                    , ByVal vadrsAllExceptions As ADODB.Recordset _
                    , ByVal vadrsLockStatus As ADODB.Recordset _
                    ) As ADODB.Recordset
On Error GoTo DetermineAbsenceStartAndEnd_EH
'Index and Flag
Dim llSchIndex As Long, llExcIndex As Long
Dim llSchCopyIndex As Long, llExcCopyIndex As Long
Dim llDateIndex As Long, lbUpdated As Boolean
'Date/Time Hold Variables
Dim ldtStartDate As Date, lsStartTime As String
Dim ldtEndDate As Date, lsEndTime As String
Dim ldtRtnDate As Date, lsRtnTime As String
Dim ldtSchCompareDate As Date, ldtExcDateCompare As Date
Dim ldtStDateWorked As Date, lbMove As Boolean
Dim lvBookMark As Variant, llTradeCounter As Long
Dim ldtMaxPayrollProcessedDate As Date
'First/Last Processed Record
Dim ldtFirstStartDate As Date, lsFirstStartTime As String
Dim ldtFirstStDateWorked As Date
Dim ldtLastProcessedEndDate As Date, lsLastProcessedEndTime As String
Dim ldtLastProcessedRtnDate As Date, lsLastProcessedRtnTime As String
'RS Fields
Dim lsFldName As String, llFldIndex As Long
Dim llFldSize As Long, ladFldType As ADODB.DataTypeEnum
Dim ladFldAttrib As ADODB.FieldAttributeEnum
'RS
Dim ladrsAbsExcCopy As ADODB.Recordset
Dim ladrsSchCopy As ADODB.Recordset
Dim ladrsStartAndEnd As ADODB.Recordset
Dim ladrsStartEndCopy As ADODB.Recordset

    ldtMaxPayrollProcessedDate = #1/1/100#
    If (RecordsExist(vadrsLockStatus)) Then
        With vadrsLockStatus
            .Filter = adFilterNone

            .Filter = "LOCK_STATUS = 'Y'"
            If Not (.EOF) Then
                .Sort = "CHARGE_DATE DESC"
                .MoveFirst

                ldtMaxPayrollProcessedDate = !CHARGE_DATE
            End If

            .Filter = adFilterNone
        End With
    End If

    'Create a copy of the Schedule
    Set ladrsSchCopy = New ADODB.Recordset
    With vadrsSchedule
        For llSchIndex = 0 To .Fields.Count - 1
            'Get the fld name, size
            lsFldName = .Fields.Item(llSchIndex).Name
            ladFldType = .Fields.Item(llSchIndex).Type
            llFldSize = .Fields.Item(llSchIndex).DefinedSize
            ladFldAttrib = .Fields.Item(llSchIndex).Attributes

            If (llFldSize > 0) Then
                ladrsSchCopy.Fields.Append lsFldName, ladFldType _
                    , llFldSize, ladFldAttrib
            Else
                ladrsSchCopy.Fields.Append lsFldName, ladFldType _
                    , , ladFldAttrib
            End If
        Next llSchIndex

        ladrsSchCopy.CursorLocation = adUseClient
        ladrsSchCopy.LockType = adLockBatchOptimistic
        ladrsSchCopy.CursorType = adOpenStatic
        ladrsSchCopy.Open

        .MoveFirst
        For llSchIndex = 1 To .RecordCount
            ladrsSchCopy.AddNew

            For llFldIndex = 0 To .Fields.Count - 1
                ladrsSchCopy.Fields.Item(llFldIndex).Value _
                    = .Fields.Item(llFldIndex).Value
            Next llFldIndex

            ladrsSchCopy.Update

            .MoveNext
        Next llSchIndex

        .MoveFirst
    End With

    'Create a copy of the Absence Exceptions
    Set ladrsAbsExcCopy = New ADODB.Recordset
    With vadrsAbsenceTime
        For llExcIndex = 0 To .Fields.Count - 1
            'Get the fld name, size
            lsFldName = .Fields.Item(llExcIndex).Name
            ladFldType = .Fields.Item(llExcIndex).Type
            llFldSize = .Fields.Item(llExcIndex).DefinedSize
            ladFldAttrib = .Fields.Item(llExcIndex).Attributes

            If (llFldSize > 0) Then
                ladrsAbsExcCopy.Fields.Append lsFldName, ladFldType _
                    , llFldSize, ladFldAttrib
            Else
                ladrsAbsExcCopy.Fields.Append lsFldName, ladFldType _
                    , , ladFldAttrib
            End If
        Next llExcIndex

        ladrsAbsExcCopy.CursorLocation = adUseClient
        ladrsAbsExcCopy.LockType = adLockBatchOptimistic
        ladrsAbsExcCopy.CursorType = adOpenStatic
        ladrsAbsExcCopy.Open

        .MoveLast
        For llExcIndex = .RecordCount To 1 Step -1
            If (Trim(!EXCEPTION_CD) <> "") Or (Trim(!OVER_EXC_CD) <> "") Then
                ladrsAbsExcCopy.AddNew
                For llFldIndex = 0 To .Fields.Count - 1
                    ladrsAbsExcCopy.Fields.Item(llFldIndex).Value _
                        = .Fields.Item(llFldIndex).Value
                Next llFldIndex
                ladrsAbsExcCopy.Update
            Else
                If (Trim(!REC_STATUS) = "") Then
                    .Delete adAffectCurrent
                    .Update
                End If
            End If

            .MovePrevious
        Next llExcIndex

        .MoveFirst
    End With

    'Create a recordset to hold Start, End  and Return information
    Set ladrsStartAndEnd = New ADODB.Recordset
    With ladrsStartAndEnd
        .Fields.Append "ABS_ST_DT", adDate, , ADOFIELDATTR
        .Fields.Append "ABS_ST_TM", adChar, 5, ADOFIELDATTR
        .Fields.Append "ST_DATE_WORKED", adDate, , ADOFIELDATTR
        .Fields.Append "ABS_END_DT", adDate, , ADOFIELDATTR
        .Fields.Append "ABS_END_TM", adChar, 5, ADOFIELDATTR
        .Fields.Append "RTN_TO_WRK_DT", adDate, , ADOFIELDATTR
        .Fields.Append "RTN_TO_WRK_TM", adChar, 5, ADOFIELDATTR

        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open
    End With

    'Loop through the absence exceptions and schedule to determine
    'the end and return information
    With vadrsAbsenceTime
        .Sort = "CHARGE_DATE, DATE_WORKED, START_TIME, END_TIME"
        .MoveFirst
        
        For llExcIndex = 1 To .RecordCount
            If (!REC_STATUS <> "D") Then
                ladrsStartAndEnd.AddNew
                ladrsStartAndEnd!ABS_ST_DT = !CHARGE_DATE
                ladrsStartAndEnd!ABS_ST_TM = !START_TIME
                ladrsStartAndEnd!ST_DATE_WORKED = !DATE_WORKED
                ladrsStartAndEnd!ABS_END_DT = !CHARGE_DATE
                ladrsStartAndEnd!ABS_END_TM = !END_TIME
                ladrsStartAndEnd.Update

                '//sched exc is one row 12/26 12/26 18:00 02:00, whereas sched is two rows split over midnight
                '//12/26 12/26 18:00 00:00 and 12/26 12/27 00:00 02:00
                '//TODO: refactor to TimeCrossesMidnight byref date1, time1, byref date2, time2
                Dim ldtStartTime As Date, ldtEndTime As Date

                ldtStartTime = DateTimeFrom(!CHARGE_DATE, !START_TIME)
                ldtEndTime = DateTimeFrom(!DATE_WORKED, !END_TIME, True)
                DateTimeCrossesMidnight ldtStartTime, ldtEndTime '//add 1 day to date worked if time spans midnight
                
                vadrsSchedule.Filter = adFilterNone
                vadrsSchedule.Filter = ("SHIFT_HOURS > 0.0 AND SYSTEM_EXC_CD <> 'SLT'" & _
                                    " AND EXCEPTION_CD <> 'ONP'" & _
                                    " AND CHARGE_DATE >= #" & !CHARGE_DATE & "#" & _
                                    " AND DATE_WORKED >= #" & !DATE_WORKED & "#")
                                    
'Debug.Print ""
'Debug.Print PadRight(StringFormat("vadrsAbsenceTime[{0}]", llExcIndex), 30) & StringFormat("= {0} {1} {2} {3}", vadrsAbsenceTime!CHARGE_DATE, vadrsAbsenceTime!DATE_WORKED, vadrsAbsenceTime!START_TIME, vadrsAbsenceTime!END_TIME)
                If Not (vadrsSchedule.EOF) Then
                    vadrsSchedule.Sort = "CHARGE_DATE, DATE_WORKED, START_TIME, END_TIME"
                    vadrsSchedule.MoveFirst

                    For llSchIndex = 1 To vadrsSchedule.RecordCount
'Debug.Print PadRight(StringFormat("   vadrsSchedule[{0}]", llSchIndex), 30) & StringFormat("= {0} {1} {2} {3}", vadrsSchedule!CHARGE_DATE, vadrsSchedule!DATE_WORKED, vadrsSchedule!START_TIME, vadrsSchedule!END_TIME)
                        lbUpdated = False
                        
                        If ldtEndTime < DateTimeFrom(vadrsSchedule!DATE_WORKED, vadrsSchedule!END_TIME, True) Then
'//this is failing for 12/24 12/24 18:00 22:00 and with shift 18:00 06:00
'                            If (CDate(Format(ldtExcDateCompare & "  " & !END_TIME, "MM/DD/YYYY HH:MM")) < _
'                                CDate(Format(ldtSchCompareDate & "  " & vadrsSchedule!SHIFT_END_TIME, "MM/DD/YYYY HH:MM"))) Then
'TODO: This (exchdrid 133416) is not splitting in test with exc ST because ST is not on schedule, why was the absence split correctly in prod?
                            If (Trim(vadrsSchedule!OVER_EXC_CD) <> "" Or _
                                Trim(vadrsSchedule!EXCEPTION_CD) <> "") Then
                                vadrsAllExceptions.Filter = adFilterNone

                                vadrsAllExceptions.Filter = ("(EXCEPTION_CD = '" & _
                                                    Trim(vadrsSchedule!OVER_EXC_CD) & _
                                                    "' AND EXCEPTION_BANK_CD <> 'O')" & _
                                                    " OR (EXCEPTION_CD = '" & _
                                                    Trim(vadrsSchedule!EXCEPTION_CD) & _
                                                    "' AND EXCEPTION_BANK_CD <> 'O')")
                                If Not (vadrsAllExceptions.EOF) Then
                                    vadrsAllExceptions.MoveFirst

                                    vadrsAllExceptions.Find "WORK_REST_IND = 'R'"
                                    If (vadrsAllExceptions.EOF) Then
                                        If (RecordsExist(ladrsStartAndEnd)) Then
                                                ladrsStartAndEnd!RTN_TO_WRK_DT = vadrsSchedule!CHARGE_DATE
                                                ladrsStartAndEnd!RTN_TO_WRK_TM = vadrsSchedule!START_TIME
                                                ladrsStartAndEnd.Update

'Debug.Print PadRight("return to work 1", 30) & StringFormat("= {0} {1} {2} {3}", ladrsStartAndEnd!RTN_TO_WRK_DT, Space(10), ladrsStartAndEnd!RTN_TO_WRK_TM, Space(5))
                                                vadrsAllExceptions.Filter = adFilterNone

                                                lbUpdated = True
                                                Exit For
                                        End If
                                    End If
                                End If

                                vadrsAllExceptions.Filter = adFilterNone
                            Else '//vadrsSchedule!OVER_EXC_CD="" And vadrsSchedule!EXCEPTION_CD=""
                                If (RecordsExist(ladrsStartAndEnd)) Then
                                        ladrsStartAndEnd!RTN_TO_WRK_DT = vadrsSchedule!CHARGE_DATE
                                        ladrsStartAndEnd!RTN_TO_WRK_TM = !END_TIME

                                    '//the schedule endtime matches the absence time, so use the next schedule row
                                    If ldtEndTime = DateTimeFrom(vadrsSchedule!DATE_WORKED, vadrsSchedule!END_TIME, True) Then
                                        vadrsSchedule.MoveNext

                                        If Not (vadrsSchedule.EOF) Then
                                            ladrsStartAndEnd!RTN_TO_WRK_DT = vadrsSchedule!CHARGE_DATE
                                            ladrsStartAndEnd!RTN_TO_WRK_TM = vadrsSchedule!START_TIME
                                        End If

                                        vadrsSchedule.MovePrevious
                                    Else
                                        '//the schedule is entirely beyond the absence time, so use its start time
                                        If ldtEndTime < DateTimeFrom(vadrsSchedule!DATE_WORKED, vadrsSchedule!END_TIME, True) And _
                                            DateTimeFrom(vadrsSchedule!CHARGE_DATE, vadrsSchedule!START_TIME) >= ldtEndTime Then
                                            ladrsStartAndEnd!RTN_TO_WRK_DT = vadrsSchedule!CHARGE_DATE
                                            ladrsStartAndEnd!RTN_TO_WRK_TM = vadrsSchedule!START_TIME
                                        End If
                                    End If
'Debug.Print PadRight("return to work 2", 30) & StringFormat("= {0} {1} {2} {3}", ladrsStartAndEnd!RTN_TO_WRK_DT, Space(10), ladrsStartAndEnd!RTN_TO_WRK_TM, Space(5))
                                    ladrsStartAndEnd.Update

                                    lbUpdated = True
                                    Exit For
                                End If
                            End If
                        Else '//ldtEndTime < DateTimeFrom(vadrsSchedule!DATE_WORKED, vadrsSchedule!END_TIME, True)
                        End If '//ldtEndTime < DateTimeFrom(vadrsSchedule!DATE_WORKED, vadrsSchedule!END_TIME, True)
        
                        If (lbUpdated) Then
                            vadrsSchedule.Filter = adFilterNone
                            Exit For
                        End If

                        vadrsSchedule.MoveNext
                        DoEvents
                    Next llSchIndex '//vadrsSchedule
                End If '//Not (vadrsSchedule.EOF)

                vadrsSchedule.Filter = adFilterNone
                    
                '//HACK: Return to Work may not be getting set and may be null
                If IsNull(ladrsStartAndEnd!RTN_TO_WRK_DT) Then
                    ladrsStartAndEnd!RTN_TO_WRK_DT = AddDays(1, MidnightDate(ldtEndTime)) '//tomorrow
                    ladrsStartAndEnd!RTN_TO_WRK_TM = "00:00" 'FormatTime(ldtEndTime, True)
    
                    ladrsStartAndEnd.Update
                End If
            End If '//(!REC_STATUS <> "D")
            
            .MoveNext
            DoEvents
        Next llExcIndex '//vadrsAbsenceTime
    End With

    Set ladrsAbsExcCopy = Nothing
    Set ladrsSchCopy = Nothing

    '//for testing only
'    Set madrsAbsenceStartEndDetails = CopyRecordset(ladrsStartAndEnd)
    
    'Determine the end date (range) for absence dates
    If (RecordsExist(ladrsStartAndEnd)) Then
        Set ladrsStartEndCopy = New ADODB.Recordset
        With ladrsStartEndCopy
            .Fields.Append "ABS_ST_DT", adDate, , ADOFIELDATTR
            .Fields.Append "ABS_ST_TM", adChar, 5, ADOFIELDATTR
            .Fields.Append "ST_DATE_WORKED", adDate, , ADOFIELDATTR
            .Fields.Append "ABS_END_DT", adDate, , ADOFIELDATTR
            .Fields.Append "ABS_END_TM", adChar, 5, ADOFIELDATTR
            .Fields.Append "RTN_TO_WRK_DT", adDate, , ADOFIELDATTR
            .Fields.Append "RTN_TO_WRK_TM", adChar, 5, ADOFIELDATTR

            .CursorLocation = adUseClient
            .CursorType = adOpenStatic
            .LockType = adLockBatchOptimistic
            .Open
        End With

        With ladrsStartAndEnd
            '//collapse rows from the beginning to the payroll processed date (use start dates from 1st row, use end and return dates from last row)
            '//this does not look for any gaps or overlaps between the return to work date and the start date on the next row
            '//this could cause an absence to not be split when it should (see ExecHdrId=68168, no row for 1/31/2011)
            .MoveFirst

            ldtFirstStartDate = !ABS_ST_DT
            lsFirstStartTime = !ABS_ST_TM
            ldtFirstStDateWorked = !ST_DATE_WORKED

            .Filter = adFilterNone
            
            .Filter = "ABS_ST_DT <= #" & ldtMaxPayrollProcessedDate & "#"
            If Not (.EOF) Then
                .Sort = "ST_DATE_WORKED, ABS_ST_DT, ABS_ST_TM, ABS_END_DT, ABS_END_TM"
                .MoveLast

                ldtLastProcessedEndDate = !ABS_END_DT
                lsLastProcessedEndTime = !ABS_END_TM
                ldtLastProcessedRtnDate = !RTN_TO_WRK_DT
                lsLastProcessedRtnTime = !RTN_TO_WRK_TM

                .MoveFirst
                For llDateIndex = 1 To .RecordCount
                    .Delete adAffectCurrent

                    If Not (.EOF) Then
                        .MoveNext
                    End If
                    DoEvents
                Next llDateIndex

                .Filter = adFilterNone

                .AddNew
                !ABS_ST_DT = ldtFirstStartDate
                !ABS_ST_TM = lsFirstStartTime
                !ST_DATE_WORKED = ldtFirstStDateWorked
                !ABS_END_DT = ldtLastProcessedEndDate
                !ABS_END_TM = lsLastProcessedEndTime
                !RTN_TO_WRK_DT = ldtLastProcessedRtnDate
                !RTN_TO_WRK_TM = lsLastProcessedRtnTime
                .Update
            End If

            .Filter = adFilterNone
            .MoveFirst

            For llDateIndex = 1 To .RecordCount
                ladrsStartEndCopy.AddNew

                For llFldIndex = 0 To .Fields.Count - 1
                    ladrsStartEndCopy.Fields.Item(llFldIndex).Value = .Fields.Item(llFldIndex).Value
                Next llFldIndex

                ladrsStartEndCopy.Update

                .MoveNext
            Next llDateIndex

            'Check for Double Shift/Owl
            'If Double Shift/Owl exists, flatten.
            ladrsStartEndCopy.Sort = "ST_DATE_WORKED, ABS_ST_DT, ABS_ST_TM, ABS_END_DT, ABS_END_TM"
            ladrsStartEndCopy.MoveFirst

            .Sort = "ST_DATE_WORKED, ABS_ST_DT, ABS_ST_TM, ABS_END_DT, ABS_END_TM"
            .MoveFirst

            llTradeCounter = 0
            Do
                ladrsStartEndCopy.Filter = adFilterNone

                ladrsStartEndCopy.Filter = "ABS_ST_DT = #" & !ABS_ST_DT & "#"
                If Not (ladrsStartEndCopy.EOF) Then
                    If (ladrsStartEndCopy.RecordCount > 1) Then
                        'Potential Double Shift
                        llTradeCounter = llTradeCounter + 1

                        ladrsStartEndCopy.MoveFirst

                        ldtStartDate = ladrsStartEndCopy!ABS_ST_DT
                        lsStartTime = ladrsStartEndCopy!ABS_ST_TM
                        ldtEndDate = ladrsStartEndCopy!ABS_END_DT
                        lsEndTime = ladrsStartEndCopy!ABS_END_TM
                        ldtRtnDate = ladrsStartEndCopy!RTN_TO_WRK_DT
                        lsRtnTime = ladrsStartEndCopy!RTN_TO_WRK_TM
                        ldtStDateWorked = ladrsStartEndCopy!ST_DATE_WORKED

                        ladrsStartEndCopy.MoveNext
                        If Not (ladrsStartEndCopy.EOF) Then
                            If (ladrsStartEndCopy!ABS_ST_DT = ldtRtnDate And ladrsStartEndCopy!ABS_ST_TM = lsRtnTime) Or _
                                (ladrsStartEndCopy!ABS_END_DT = ldtEndDate And ladrsStartEndCopy!ABS_END_TM = lsEndTime And _
                                CDate(ladrsStartEndCopy!ST_DATE_WORKED & " " & ladrsStartEndCopy!ABS_END_TM) > CDate(ldtStDateWorked & " " & lsStartTime)) Then
                                ladrsStartEndCopy!ABS_ST_DT = ldtStartDate
                                ladrsStartEndCopy!ABS_ST_TM = lsStartTime
                                ladrsStartEndCopy!ST_DATE_WORKED = ldtStDateWorked
                                ladrsStartEndCopy.Update

                                ladrsStartEndCopy.MovePrevious
                                ladrsStartEndCopy.Delete adAffectCurrent
                                ladrsStartEndCopy.Update

                                If Not (ladrsStartEndCopy.EOF) Then
                                    ladrsStartEndCopy.MoveNext
                                End If
                            End If
                        End If

                        ladrsStartEndCopy.Filter = adFilterNone

                        ldtStartDate = !ABS_ST_DT
                        lsStartTime = !ABS_ST_TM
                        ldtEndDate = !ABS_END_DT
                        lsEndTime = !ABS_END_TM
                        ldtRtnDate = !RTN_TO_WRK_DT
                        lsRtnTime = !RTN_TO_WRK_TM
                        ldtStDateWorked = !ST_DATE_WORKED

                        .MoveNext
                        If Not (.EOF) Then
                            If (!ABS_ST_DT = ldtRtnDate And !ABS_ST_TM = lsRtnTime) Or _
                                (!ABS_END_DT = ldtEndDate And !ABS_END_TM = lsEndTime And _
                                CDate(!ST_DATE_WORKED & " " & !ABS_END_TM) > CDate(ldtStDateWorked & " " & lsStartTime)) Then
                                !ABS_ST_DT = ldtStartDate
                                !ABS_ST_TM = lsStartTime
                                !ST_DATE_WORKED = ldtStDateWorked
                                .Update

                                .MovePrevious
                                .Delete adAffectCurrent
                                .Update
                            End If
                        End If
                    End If
                End If

                ladrsStartEndCopy.Filter = adFilterNone

                If Not (.EOF) Then
                    .MoveNext
                End If
            Loop Until (.EOF)

            'Flatten dates with return to work falling between trade flattened date
            llTradeCounter = llTradeCounter * 2

            For llExcIndex = 1 To llTradeCounter
'Debug.Print StringFormat("{0} of {1} TradeCounter Loop", llExcIndex, llTradeCounter)
                ladrsStartEndCopy.Sort = "ST_DATE_WORKED, ABS_ST_DT, ABS_ST_TM, ABS_END_DT, ABS_END_TM"
                ladrsStartEndCopy.MoveFirst

                .Sort = "ST_DATE_WORKED, ABS_ST_DT, ABS_ST_TM, ABS_END_DT, ABS_END_TM"

                For llSchCopyIndex = 1 To ladrsStartEndCopy.RecordCount
'Debug.Print StringFormat("   {0} of {1} ladrsStartEndCopy Loop1", llSchCopyIndex, ladrsStartEndCopy.RecordCount)
                    .MoveFirst

                    For llDateIndex = 1 To .RecordCount
'Debug.Print StringFormat("      {0} of {1} DateIndex Loop 1", llDateIndex, .RecordCount)
                        lbMove = False

                        If CDate(!ST_DATE_WORKED & " " & !ABS_ST_TM) < CDate(ladrsStartEndCopy!RTN_TO_WRK_DT & " " & ladrsStartEndCopy!RTN_TO_WRK_TM) And _
                            CDate(!ABS_END_DT & " " & !ABS_END_TM) > CDate(ladrsStartEndCopy!RTN_TO_WRK_DT & " " & ladrsStartEndCopy!RTN_TO_WRK_TM) And _
                            CDate(!ST_DATE_WORKED & " " & !ABS_ST_TM) <= CDate(ladrsStartEndCopy!ST_DATE_WORKED & " " & ladrsStartEndCopy!ABS_ST_TM) And _
                            CDate(!ABS_END_DT & " " & !ABS_END_TM) >= CDate(ladrsStartEndCopy!ABS_END_DT & " " & ladrsStartEndCopy!ABS_END_TM) And _
                            CDate(!RTN_TO_WRK_DT & " " & !RTN_TO_WRK_TM) >= CDate(ladrsStartEndCopy!RTN_TO_WRK_DT & " " & ladrsStartEndCopy!RTN_TO_WRK_TM) Then
                            'Set the original to vars.
                            ldtEndDate = !ABS_END_DT
                            lsEndTime = !ABS_END_TM
                            ldtRtnDate = !RTN_TO_WRK_DT
                            lsRtnTime = !RTN_TO_WRK_TM

                            lvBookMark = .Bookmark

                            .Filter = adFilterNone
                            .Filter = "ABS_ST_DT = #" & ladrsStartEndCopy!ABS_ST_DT & "#" & _
                                " AND ABS_ST_TM = '" & ladrsStartEndCopy!ABS_ST_TM & "'" & _
                                " AND ST_DATE_WORKED = #" & ladrsStartEndCopy!ST_DATE_WORKED & "#" & _
                                " AND ABS_END_DT = #" & ladrsStartEndCopy!ABS_END_DT & "#" & _
                                " AND ABS_END_TM = '" & ladrsStartEndCopy!ABS_END_TM & "'"
                            If Not (.EOF) Then
                                !ABS_END_DT = ldtEndDate
                                !ABS_END_TM = lsEndTime
                                !RTN_TO_WRK_DT = ldtRtnDate
                                !RTN_TO_WRK_TM = lsRtnTime
                                .Update
                            End If

                            .Filter = adFilterNone

                            .Bookmark = lvBookMark
                            If Not (.EOF) Then
                                ladrsStartEndCopy!ABS_END_DT = ldtEndDate
                                ladrsStartEndCopy!ABS_END_TM = lsEndTime
                                ladrsStartEndCopy!RTN_TO_WRK_DT = ldtRtnDate
                                ladrsStartEndCopy!RTN_TO_WRK_TM = lsRtnTime
                                ladrsStartEndCopy.Update

                                If Not (ladrsStartEndCopy.EOF) Then
                                    lvBookMark = ladrsStartEndCopy.Bookmark
                                    ladrsStartEndCopy.Filter = adFilterNone

                                    ladrsStartEndCopy.Filter = "ABS_ST_DT = #" & !ABS_ST_DT & "#" & _
                                                        " AND ABS_ST_TM = '" & !ABS_ST_TM & "'" & _
                                                        " AND ST_DATE_WORKED = #" & !ST_DATE_WORKED & "#" & _
                                                        " AND ABS_END_DT = #" & !ABS_END_DT & "#" & _
                                                        " AND ABS_END_TM = '" & !ABS_END_TM & "'"

                                    If Not (ladrsStartEndCopy.EOF) Then
                                        ladrsStartEndCopy.Delete adAffectCurrent
                                        ladrsStartEndCopy.Update
                                        lbMove = True
                                    End If

                                    ladrsStartEndCopy.Filter = adFilterNone

                                    ladrsStartEndCopy.Bookmark = lvBookMark
                                End If

                                .Delete adAffectCurrent
                                .Update
                            End If
                        End If

                        If Not (.EOF) Then
                            .MoveNext

                            If (.EOF) Then
                                Exit For
                            End If
                        Else
                            Exit For
                        End If

                        If (lbMove) Then
                            If Not (ladrsStartEndCopy.EOF) Then
                                ladrsStartEndCopy.MoveNext
                                If (ladrsStartEndCopy.EOF) Then
                                    Exit For
                                End If
                            Else
                                Exit For
                            End If
                        End If
                        DoEvents
                    Next llDateIndex

                    If Not (ladrsStartEndCopy.EOF) Then
                        ladrsStartEndCopy.MoveNext
                        If (ladrsStartEndCopy.EOF) Then
                            Exit For
                        End If
                    Else
                        Exit For
                    End If
                    DoEvents
                Next llSchCopyIndex
'Debug.Print "entering botton loop"
                ladrsStartEndCopy.Sort = "ST_DATE_WORKED, ABS_ST_DT, ABS_ST_TM, ABS_END_DT, ABS_END_TM"
                .Sort = "ST_DATE_WORKED, ABS_ST_DT, ABS_ST_TM, ABS_END_DT, ABS_END_TM"
                .MoveFirst

                For llDateIndex = 1 To .RecordCount
'Debug.Print StringFormat("   {0} of {1} DateLoop2", llDateIndex, .RecordCount)
                    ladrsStartEndCopy.MoveFirst

                    For llSchCopyIndex = 1 To ladrsStartEndCopy.RecordCount
'Debug.Print StringFormat("      {0} of {1} ladrsStartEndCopy2", llDateIndex, .RecordCount)
                        lbMove = False

                        If CDate(!RTN_TO_WRK_DT & " " & !RTN_TO_WRK_TM) > CDate(ladrsStartEndCopy!RTN_TO_WRK_DT & " " & ladrsStartEndCopy!RTN_TO_WRK_TM) And _
                            CDate(!ST_DATE_WORKED & " " & !ABS_ST_TM) <= CDate(ladrsStartEndCopy!ST_DATE_WORKED & " " & ladrsStartEndCopy!ABS_ST_TM) And _
                            CDate(!ABS_END_DT & " " & !ABS_END_TM) >= CDate(ladrsStartEndCopy!ABS_END_DT & " " & ladrsStartEndCopy!ABS_END_TM) Then
                            lvBookMark = .Bookmark

                            .Filter = adFilterNone
                            .Filter = ("ABS_ST_DT = #" & ladrsStartEndCopy!ABS_ST_DT & "#" & _
                                " AND ABS_ST_TM = '" & ladrsStartEndCopy!ABS_ST_TM & "'" & _
                                " AND ABS_END_DT = #" & ladrsStartEndCopy!ABS_END_DT & "#" & _
                                " AND ABS_END_TM = '" & ladrsStartEndCopy!ABS_END_TM & "'" & _
                                " AND RTN_TO_WRK_DT = #" & ladrsStartEndCopy!RTN_TO_WRK_DT & "#" & _
                                " AND RTN_TO_WRK_TM = '" & ladrsStartEndCopy!RTN_TO_WRK_TM & "'" & _
                                " AND ST_DATE_WORKED = #" & ladrsStartEndCopy!ST_DATE_WORKED & "#")
                            If Not (.EOF) Then
                                .Delete adAffectCurrent
                                .Update

                                ladrsStartEndCopy.Delete adAffectCurrent
                                ladrsStartEndCopy.Update
                                lbMove = True
                            End If

                            .Filter = adFilterNone
                            .Bookmark = lvBookMark
                        End If

                        If Not (ladrsStartEndCopy.EOF) Then
                            ladrsStartEndCopy.MoveNext
                        Else
                            Exit For
                        End If

                        If (lbMove) Then
                            If Not (.EOF) Then
                                .MoveNext

                                If (.EOF) Then
                                    Exit For
                                End If
                            Else
                                Exit For
                            End If
                        End If
                        DoEvents
                    Next llSchCopyIndex

                    If Not (.EOF) Then
                        .MoveNext

                        If (.EOF) Then
                            Exit For
                        End If
                    Else
                        Exit For
                    End If
                    DoEvents
                Next llDateIndex
                DoEvents
            Next llExcIndex

            'Flatten dates with same return to work date or return to work overlapping absence
            For llExcIndex = 1 To ladrsStartEndCopy.RecordCount
                ladrsStartEndCopy.Sort = "ST_DATE_WORKED, ABS_ST_DT, ABS_ST_TM, ABS_END_DT, ABS_END_TM"
                ladrsStartEndCopy.MoveFirst

                For llDateIndex = 1 To ladrsStartEndCopy.RecordCount
                    .Filter = adFilterNone

                    .Filter = "(RTN_TO_WRK_DT = #" & ladrsStartEndCopy!RTN_TO_WRK_DT & "#" & _
                                " AND RTN_TO_WRK_TM = '" & ladrsStartEndCopy!RTN_TO_WRK_TM & "')"
                    If Not (.EOF) Then
                        .Sort = "ST_DATE_WORKED, ABS_ST_DT, ABS_ST_TM, ABS_END_DT, ABS_END_TM"
                        .MoveLast

                        ldtEndDate = !ABS_END_DT
                        lsEndTime = !ABS_END_TM

                        .MoveFirst
                        !ABS_END_DT = ldtEndDate
                        !ABS_END_TM = lsEndTime
                        .Update

                        .MoveNext
                        Do While Not (.EOF)
                            .Delete adAffectCurrent
                            .Update
                            .MoveNext
                        Loop
                    End If

                    .Filter = adFilterNone

                    ladrsStartEndCopy.MoveNext
                    DoEvents
                Next llDateIndex

                'Flatten dates with same return to work date as the start date
                .Sort = "ST_DATE_WORKED, ABS_ST_DT, ABS_ST_TM, ABS_END_DT, ABS_END_TM"
                .MoveFirst

                For llDateIndex = 1 To .RecordCount
                    ldtStartDate = !ABS_ST_DT
                    lsStartTime = !ABS_ST_TM
                    ldtEndDate = !ABS_END_DT
                    lsEndTime = !ABS_END_TM
                    ldtRtnDate = !RTN_TO_WRK_DT
                    lsRtnTime = !RTN_TO_WRK_TM
                    ldtStDateWorked = !ST_DATE_WORKED

                    .MoveNext
                    If Not (.EOF) Then
                        If (!ABS_ST_DT = ldtRtnDate And !ABS_ST_TM = lsRtnTime) Or _
                            (CDate(!ST_DATE_WORKED & " " & !ABS_ST_TM) <= CDate(ldtRtnDate & " " & lsRtnTime) And _
                            CDate(ldtRtnDate & " " & lsRtnTime) <= CDate(!ABS_END_DT & " " & !ABS_END_TM)) Then
                            !ABS_ST_DT = ldtStartDate
                            !ABS_ST_TM = lsStartTime
                            !ST_DATE_WORKED = ldtStDateWorked
                            .Update

                            .MovePrevious
                            .Delete adAffectCurrent
                            .Update

                            If Not (.EOF) Then
                                .MoveNext
                            End If
                        End If
                    End If
                    DoEvents
                Next llDateIndex
                DoEvents
            Next llExcIndex

            ladrsStartEndCopy.Sort = "ST_DATE_WORKED, ABS_ST_DT, ABS_ST_TM, ABS_END_DT, ABS_END_TM"
            ladrsStartEndCopy.MoveFirst

            For llDateIndex = 1 To ladrsStartEndCopy.RecordCount
                .Filter = adFilterNone

                .Filter = ("RTN_TO_WRK_DT = #" & ladrsStartEndCopy!RTN_TO_WRK_DT & "#" & _
                            " AND RTN_TO_WRK_TM = '" & ladrsStartEndCopy!RTN_TO_WRK_TM & "'")
                If Not (.EOF) Then
                    .Sort = "ST_DATE_WORKED, ABS_ST_DT, ABS_ST_TM, ABS_END_DT, ABS_END_TM"
                    .MoveLast

                    ldtEndDate = !ABS_END_DT
                    lsEndTime = !ABS_END_TM

                    .MoveFirst
                    !ABS_END_DT = ldtEndDate
                    !ABS_END_TM = lsEndTime
                    .Update

                    .MoveNext
                    Do While Not (.EOF)
                        .Delete adAffectCurrent
                        .Update
                        .MoveNext
                    Loop
                End If

                .Filter = adFilterNone

                ladrsStartEndCopy.MoveNext
                DoEvents
            Next llDateIndex

            If Not (.EOF And .BOF) Then
                .MoveFirst
            End If
        End With
    End If

    Set ladrsStartEndCopy = Nothing

    Set DetermineAbsenceStartAndEndOld = ladrsStartAndEnd

    Set ladrsStartAndEnd = Nothing

    Exit Function
DetermineAbsenceStartAndEnd_EH:
    Set ladrsAbsExcCopy = Nothing
    Set ladrsSchCopy = Nothing

    Set ladrsStartAndEnd = Nothing
    Set ladrsStartEndCopy = Nothing

    TrsRaiseError MODULENAME, "DetermineAbsenceStartAndEndOld"
End Function

'//old

'//temp for testing the old
Private Function DetermineAbsenceStartAndEndTemp(ByVal vadrsAbsenceTime As ADODB.Recordset _
                    , ByVal vadrsSchedule As ADODB.Recordset _
                    , ByVal vadrsAllExceptions As ADODB.Recordset _
                    , ByVal vadrsLockStatus As ADODB.Recordset _
                    ) As ADODB.Recordset
On Error GoTo eh
'Index and Flag
Dim llSchIndex As Long, llExcIndex As Long
Dim llSchCopyIndex As Long, llExcCopyIndex As Long
Dim llDateIndex As Long, lbUpdated As Boolean
'Date/Time Hold Variables
Dim ldtStartDate As Date, lsStartTime As String
Dim ldtEndDate As Date, lsEndTime As String
Dim ldtRtnDate As Date, lsRtnTime As String
Dim ldtSchCompareDate As Date, ldtExcDateCompare As Date
Dim ldtStDateWorked As Date, lbMove As Boolean
Dim lvBookMark As Variant, llTradeCounter As Long
Dim ldtMaxPayrollProcessedDate As Date
'First/Last Processed Record
Dim ldtFirstStartDate As Date, lsFirstStartTime As String
Dim ldtFirstStDateWorked As Date
Dim ldtLastProcessedEndDate As Date, lsLastProcessedEndTime As String
Dim ldtLastProcessedRtnDate As Date, lsLastProcessedRtnTime As String
'RS Fields
Dim lsFldName As String, llFldIndex As Long
Dim llFldSize As Long, ladFldType As ADODB.DataTypeEnum
Dim ladFldAttrib As ADODB.FieldAttributeEnum
'RS
Dim ladrsAbsExcCopy As ADODB.Recordset
Dim ladrsSchCopy As ADODB.Recordset
Dim ladrsStartAndEnd As ADODB.Recordset
Dim ladrsStartEndCopy As ADODB.Recordset

    ldtMaxPayrollProcessedDate = #1/1/100#
    If (RecordsExist(vadrsLockStatus)) Then
        With vadrsLockStatus
            .Filter = adFilterNone

            .Filter = "LOCK_STATUS = 'Y'"
            If Not (.EOF) Then
                .Sort = "CHARGE_DATE DESC"
                .MoveFirst

                ldtMaxPayrollProcessedDate = !CHARGE_DATE
            End If

            .Filter = adFilterNone
        End With
    End If

    'Create a copy of the Schedule
    Set ladrsSchCopy = New ADODB.Recordset
    With vadrsSchedule
        For llSchIndex = 0 To .Fields.Count - 1
            'Get the fld name, size
            lsFldName = .Fields.Item(llSchIndex).Name
            ladFldType = .Fields.Item(llSchIndex).Type
            llFldSize = .Fields.Item(llSchIndex).DefinedSize
            ladFldAttrib = .Fields.Item(llSchIndex).Attributes

            If (llFldSize > 0) Then
                ladrsSchCopy.Fields.Append lsFldName, ladFldType _
                    , llFldSize, ladFldAttrib
            Else
                ladrsSchCopy.Fields.Append lsFldName, ladFldType _
                    , , ladFldAttrib
            End If
        Next llSchIndex

        ladrsSchCopy.CursorLocation = adUseClient
        ladrsSchCopy.LockType = adLockBatchOptimistic
        ladrsSchCopy.CursorType = adOpenStatic
        ladrsSchCopy.Open

        .MoveFirst
        For llSchIndex = 1 To .RecordCount
            ladrsSchCopy.AddNew

            For llFldIndex = 0 To .Fields.Count - 1
                ladrsSchCopy.Fields.Item(llFldIndex).Value _
                    = .Fields.Item(llFldIndex).Value
            Next llFldIndex

            ladrsSchCopy.Update

            .MoveNext
        Next llSchIndex

        .MoveFirst
    End With

    'Create a copy of the Absence Exceptions
    Set ladrsAbsExcCopy = New ADODB.Recordset
    With vadrsAbsenceTime
        For llExcIndex = 0 To .Fields.Count - 1
            'Get the fld name, size
            lsFldName = .Fields.Item(llExcIndex).Name
            ladFldType = .Fields.Item(llExcIndex).Type
            llFldSize = .Fields.Item(llExcIndex).DefinedSize
            ladFldAttrib = .Fields.Item(llExcIndex).Attributes

            If (llFldSize > 0) Then
                ladrsAbsExcCopy.Fields.Append lsFldName, ladFldType _
                    , llFldSize, ladFldAttrib
            Else
                ladrsAbsExcCopy.Fields.Append lsFldName, ladFldType _
                    , , ladFldAttrib
            End If
        Next llExcIndex

        ladrsAbsExcCopy.CursorLocation = adUseClient
        ladrsAbsExcCopy.LockType = adLockBatchOptimistic
        ladrsAbsExcCopy.CursorType = adOpenStatic
        ladrsAbsExcCopy.Open

        .MoveLast
        For llExcIndex = .RecordCount To 1 Step -1
            If (Trim(!EXCEPTION_CD) <> "") Or (Trim(!OVER_EXC_CD) <> "") Then
                ladrsAbsExcCopy.AddNew
                For llFldIndex = 0 To .Fields.Count - 1
                    ladrsAbsExcCopy.Fields.Item(llFldIndex).Value _
                        = .Fields.Item(llFldIndex).Value
                Next llFldIndex
                ladrsAbsExcCopy.Update
            Else
                If (Trim(!REC_STATUS) = "") Then
                    .Delete adAffectCurrent
                    .Update
                End If
            End If

            .MovePrevious
        Next llExcIndex

        .MoveFirst
    End With

    'Create a recordset to hold Start, End  and Return information
    Set ladrsStartAndEnd = New ADODB.Recordset
    With ladrsStartAndEnd
        .Fields.Append "ABS_ST_DT", adDate, , ADOFIELDATTR
        .Fields.Append "ABS_ST_TM", adChar, 5, ADOFIELDATTR
        .Fields.Append "ST_DATE_WORKED", adDate, , ADOFIELDATTR
        .Fields.Append "ABS_END_DT", adDate, , ADOFIELDATTR
        .Fields.Append "ABS_END_TM", adChar, 5, ADOFIELDATTR
        .Fields.Append "RTN_TO_WRK_DT", adDate, , ADOFIELDATTR
        .Fields.Append "RTN_TO_WRK_TM", adChar, 5, ADOFIELDATTR

        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open
    End With

    'Loop through the absence exceptions and schedule to determine
    'the end and return information
    With vadrsAbsenceTime
        .Sort = "CHARGE_DATE, DATE_WORKED, START_TIME, END_TIME"
        .MoveFirst
        
        For llExcIndex = 1 To .RecordCount
            If (!REC_STATUS <> "D") Then
                ladrsStartAndEnd.AddNew
                ladrsStartAndEnd!ABS_ST_DT = !CHARGE_DATE
                ladrsStartAndEnd!ABS_ST_TM = !START_TIME
                ladrsStartAndEnd!ST_DATE_WORKED = !DATE_WORKED
'                ladrsStartAndEnd!ABS_END_DT = !CHARGE_DATE
'                ladrsStartAndEnd!ABS_END_TM = !END_TIME
'                ladrsStartAndEnd.Update

                '//sched exc is one row 12/26 12/26 18:00 02:00, whereas sched is two rows split over midnight
                '//12/26 12/26 18:00 00:00 and 12/26 12/27 00:00 02:00
                '//TODO: refactor to TimeCrossesMidnight byref date1, time1, byref date2, time2
                Dim ldtStartTime As Date, ldtEndTime As Date

                ldtStartTime = DateTimeFrom(!CHARGE_DATE, !START_TIME)
                ldtEndTime = DateTimeFrom(!DATE_WORKED, !END_TIME, True)
                DateTimeCrossesMidnight ldtStartTime, ldtEndTime '//add 1 day to date worked if time spans midnight

                ladrsStartAndEnd!ABS_END_DT = MidnightDate(ldtEndTime)
                ladrsStartAndEnd!ABS_END_TM = FormatTime(ldtEndTime)
                ladrsStartAndEnd.Update
                
                vadrsSchedule.Filter = adFilterNone
                vadrsSchedule.Filter = ("SHIFT_HOURS > 0.0 AND SYSTEM_EXC_CD <> 'SLT'" & _
                                    " AND EXCEPTION_CD <> 'ONP'" & _
                                    " AND CHARGE_DATE >= #" & !CHARGE_DATE & "#" & _
                                    " AND DATE_WORKED >= #" & !DATE_WORKED & "#")
                                    
'Debug.Print ""
'Debug.Print PadRight(StringFormat("vadrsAbsenceTime[{0}]", llExcIndex), 30) & StringFormat("= {0} {1} {2} {3}", vadrsAbsenceTime!CHARGE_DATE, vadrsAbsenceTime!DATE_WORKED, vadrsAbsenceTime!START_TIME, vadrsAbsenceTime!END_TIME)
                If Not (vadrsSchedule.EOF) Then
                    vadrsSchedule.Sort = "CHARGE_DATE, DATE_WORKED, START_TIME, END_TIME"
                    vadrsSchedule.MoveFirst

                    For llSchIndex = 1 To vadrsSchedule.RecordCount
'Debug.Print PadRight(StringFormat("   vadrsSchedule[{0}]", llSchIndex), 30) & StringFormat("= {0} {1} {2} {3}", vadrsSchedule!CHARGE_DATE, vadrsSchedule!DATE_WORKED, vadrsSchedule!START_TIME, vadrsSchedule!END_TIME)
                        lbUpdated = False
                        
                        If ldtEndTime < DateTimeFrom(vadrsSchedule!DATE_WORKED, vadrsSchedule!END_TIME, True) Then
'//this is failing for 12/24 12/24 18:00 22:00 and with shift 18:00 06:00
'                            If (CDate(Format(ldtExcDateCompare & "  " & !END_TIME, "MM/DD/YYYY HH:MM")) < _
'                                CDate(Format(ldtSchCompareDate & "  " & vadrsSchedule!SHIFT_END_TIME, "MM/DD/YYYY HH:MM"))) Then
'TODO: This (exchdrid 133416) is not splitting in test with exc ST because ST is not on schedule, why was the absence split correctly in prod?
                            If (Trim(vadrsSchedule!OVER_EXC_CD) <> "" Or _
                                Trim(vadrsSchedule!EXCEPTION_CD) <> "") Then
                                vadrsAllExceptions.Filter = adFilterNone

                                vadrsAllExceptions.Filter = ("(EXCEPTION_CD = '" & _
                                                    Trim(vadrsSchedule!OVER_EXC_CD) & _
                                                    "' AND EXCEPTION_BANK_CD <> 'O')" & _
                                                    " OR (EXCEPTION_CD = '" & _
                                                    Trim(vadrsSchedule!EXCEPTION_CD) & _
                                                    "' AND EXCEPTION_BANK_CD <> 'O')")
                                If Not (vadrsAllExceptions.EOF) Then
                                    vadrsAllExceptions.MoveFirst

                                    vadrsAllExceptions.Find "WORK_REST_IND = 'R'" '//ignore rest (non-working day)
                                    If (vadrsAllExceptions.EOF) Then '//working day
                                        If (RecordsExist(ladrsStartAndEnd)) Then
                                                ladrsStartAndEnd!RTN_TO_WRK_DT = vadrsSchedule!CHARGE_DATE
                                                ladrsStartAndEnd!RTN_TO_WRK_TM = vadrsSchedule!START_TIME
                                                ladrsStartAndEnd.Update

'Debug.Print PadRight("return to work 1", 30) & StringFormat("= {0} {1} {2} {3}", ladrsStartAndEnd!RTN_TO_WRK_DT, Space(10), ladrsStartAndEnd!RTN_TO_WRK_TM, Space(5))
                                                vadrsAllExceptions.Filter = adFilterNone

                                                lbUpdated = True
                                                Exit For
                                        End If
                                    End If
                                End If

                                vadrsAllExceptions.Filter = adFilterNone
                            Else '//vadrsSchedule!OVER_EXC_CD="" And vadrsSchedule!EXCEPTION_CD=""
                                If (RecordsExist(ladrsStartAndEnd)) Then
                                        ladrsStartAndEnd!RTN_TO_WRK_DT = vadrsSchedule!CHARGE_DATE
                                        ladrsStartAndEnd!RTN_TO_WRK_TM = !END_TIME

                                    '//the schedule endtime matches the absence time, so use the next schedule row
                                    If ldtEndTime = DateTimeFrom(vadrsSchedule!DATE_WORKED, vadrsSchedule!END_TIME, True) Then
                                        '//TODO: will this ever get executed???
                                        vadrsSchedule.MoveNext

                                        If Not (vadrsSchedule.EOF) Then
                                            ladrsStartAndEnd!RTN_TO_WRK_DT = vadrsSchedule!CHARGE_DATE
                                            ladrsStartAndEnd!RTN_TO_WRK_TM = vadrsSchedule!START_TIME
                                        End If

                                        vadrsSchedule.MovePrevious
                                    Else
                                        '//the schedule is entirely beyond the absence time, so use its start time
                                        If ldtEndTime < DateTimeFrom(vadrsSchedule!DATE_WORKED, vadrsSchedule!END_TIME, True) And _
                                            DateTimeFrom(vadrsSchedule!CHARGE_DATE, vadrsSchedule!START_TIME) >= ldtEndTime Then
                                            ladrsStartAndEnd!RTN_TO_WRK_DT = vadrsSchedule!CHARGE_DATE
                                            ladrsStartAndEnd!RTN_TO_WRK_TM = vadrsSchedule!START_TIME
                                        End If
                                    End If
'Debug.Print PadRight("return to work 2", 30) & StringFormat("= {0} {1} {2} {3}", ladrsStartAndEnd!RTN_TO_WRK_DT, Space(10), ladrsStartAndEnd!RTN_TO_WRK_TM, Space(5))
                                    ladrsStartAndEnd.Update

                                    lbUpdated = True
                                    Exit For
                                End If
                            End If
                        Else '//ldtEndTime < DateTimeFrom(vadrsSchedule!DATE_WORKED, vadrsSchedule!END_TIME, True)
                        End If '//ldtEndTime < DateTimeFrom(vadrsSchedule!DATE_WORKED, vadrsSchedule!END_TIME, True)
        
                        If (lbUpdated) Then
                            vadrsSchedule.Filter = adFilterNone
                            Exit For
                        End If

                        vadrsSchedule.MoveNext
                        DoEvents
                    Next llSchIndex '//vadrsSchedule
                End If '//Not (vadrsSchedule.EOF)

                vadrsSchedule.Filter = adFilterNone
                    
                '//HACK: Return to Work may not be getting set and may be null
                If IsNull(ladrsStartAndEnd!RTN_TO_WRK_DT) Then
                    ladrsStartAndEnd!RTN_TO_WRK_DT = AddDays(1, MidnightDate(ldtEndTime)) '//tomorrow
                    ladrsStartAndEnd!RTN_TO_WRK_TM = "00:00" 'FormatTime(ldtEndTime, True)
    
                    ladrsStartAndEnd.Update
                End If
            End If '//(!REC_STATUS <> "D")
            
            .MoveNext
            DoEvents
        Next llExcIndex '//vadrsAbsenceTime
    End With

    Set ladrsAbsExcCopy = Nothing
    Set ladrsSchCopy = Nothing

    '//for testing only
'    Set madrsAbsenceStartEndDetails = CopyRecordset(ladrsStartAndEnd)
    
    'Determine the end date (range) for absence dates
    If (RecordsExist(ladrsStartAndEnd)) Then
        Set ladrsStartEndCopy = New ADODB.Recordset
        With ladrsStartEndCopy
            .Fields.Append "ABS_ST_DT", adDate, , ADOFIELDATTR
            .Fields.Append "ABS_ST_TM", adChar, 5, ADOFIELDATTR
            .Fields.Append "ST_DATE_WORKED", adDate, , ADOFIELDATTR
            .Fields.Append "ABS_END_DT", adDate, , ADOFIELDATTR
            .Fields.Append "ABS_END_TM", adChar, 5, ADOFIELDATTR
            .Fields.Append "RTN_TO_WRK_DT", adDate, , ADOFIELDATTR
            .Fields.Append "RTN_TO_WRK_TM", adChar, 5, ADOFIELDATTR

            .CursorLocation = adUseClient
            .CursorType = adOpenStatic
            .LockType = adLockBatchOptimistic
            .Open
        End With

        With ladrsStartAndEnd
            '//collapse rows from the beginning to the payroll processed date (use start dates from 1st row, use end and return dates from last row)
            '//this does not look for any gaps or overlaps between the return to work date and the start date on the next row
            '//this could cause an absence to not be split when it should (see ExecHdrId=68168, no row for 1/31/2011)
            .MoveFirst

            ldtFirstStartDate = !ABS_ST_DT
            lsFirstStartTime = !ABS_ST_TM
            ldtFirstStDateWorked = !ST_DATE_WORKED

            .Filter = adFilterNone
            
            .Filter = "ABS_ST_DT <= #" & ldtMaxPayrollProcessedDate & "#"
            If Not (.EOF) Then
                .Sort = "ST_DATE_WORKED, ABS_ST_DT, ABS_ST_TM, ABS_END_DT, ABS_END_TM"
                .MoveLast

                ldtLastProcessedEndDate = !ABS_END_DT
                lsLastProcessedEndTime = !ABS_END_TM
                ldtLastProcessedRtnDate = !RTN_TO_WRK_DT
                lsLastProcessedRtnTime = !RTN_TO_WRK_TM

                .MoveFirst
                For llDateIndex = 1 To .RecordCount
                    .Delete adAffectCurrent

                    If Not (.EOF) Then
                        .MoveNext
                    End If
                    DoEvents
                Next llDateIndex

                .Filter = adFilterNone

                .AddNew
                !ABS_ST_DT = ldtFirstStartDate
                !ABS_ST_TM = lsFirstStartTime
                !ST_DATE_WORKED = ldtFirstStDateWorked
                !ABS_END_DT = ldtLastProcessedEndDate
                !ABS_END_TM = lsLastProcessedEndTime
                !RTN_TO_WRK_DT = ldtLastProcessedRtnDate
                !RTN_TO_WRK_TM = lsLastProcessedRtnTime
                .Update
            End If

            .Filter = adFilterNone
            .MoveFirst

            For llDateIndex = 1 To .RecordCount
                ladrsStartEndCopy.AddNew

                For llFldIndex = 0 To .Fields.Count - 1
                    ladrsStartEndCopy.Fields.Item(llFldIndex).Value = .Fields.Item(llFldIndex).Value
                Next llFldIndex

                ladrsStartEndCopy.Update

                .MoveNext
            Next llDateIndex

            'Check for Double Shift/Owl
            'If Double Shift/Owl exists, flatten.
            ladrsStartEndCopy.Sort = "ST_DATE_WORKED, ABS_ST_DT, ABS_ST_TM, ABS_END_DT, ABS_END_TM"
            ladrsStartEndCopy.MoveFirst

            .Sort = "ST_DATE_WORKED, ABS_ST_DT, ABS_ST_TM, ABS_END_DT, ABS_END_TM"
            .MoveFirst

            llTradeCounter = 0
            Do
                ladrsStartEndCopy.Filter = adFilterNone

                ladrsStartEndCopy.Filter = "ABS_ST_DT = #" & !ABS_ST_DT & "#"
                If Not (ladrsStartEndCopy.EOF) Then
                    If (ladrsStartEndCopy.RecordCount > 1) Then
                        'Potential Double Shift
                        llTradeCounter = llTradeCounter + 1

                        ladrsStartEndCopy.MoveFirst

                        ldtStartDate = ladrsStartEndCopy!ABS_ST_DT
                        lsStartTime = ladrsStartEndCopy!ABS_ST_TM
                        ldtEndDate = ladrsStartEndCopy!ABS_END_DT
                        lsEndTime = ladrsStartEndCopy!ABS_END_TM
                        ldtRtnDate = ladrsStartEndCopy!RTN_TO_WRK_DT
                        lsRtnTime = ladrsStartEndCopy!RTN_TO_WRK_TM
                        ldtStDateWorked = ladrsStartEndCopy!ST_DATE_WORKED

                        ladrsStartEndCopy.MoveNext
                        If Not (ladrsStartEndCopy.EOF) Then
                            If (ladrsStartEndCopy!ABS_ST_DT = ldtRtnDate And ladrsStartEndCopy!ABS_ST_TM = lsRtnTime) Or _
                                (ladrsStartEndCopy!ABS_END_DT = ldtEndDate And ladrsStartEndCopy!ABS_END_TM = lsEndTime And _
                                CDate(ladrsStartEndCopy!ST_DATE_WORKED & " " & ladrsStartEndCopy!ABS_END_TM) > CDate(ldtStDateWorked & " " & lsStartTime)) Then
                                ladrsStartEndCopy!ABS_ST_DT = ldtStartDate
                                ladrsStartEndCopy!ABS_ST_TM = lsStartTime
                                ladrsStartEndCopy!ST_DATE_WORKED = ldtStDateWorked
                                ladrsStartEndCopy.Update

                                ladrsStartEndCopy.MovePrevious
                                ladrsStartEndCopy.Delete adAffectCurrent
                                ladrsStartEndCopy.Update

                                If Not (ladrsStartEndCopy.EOF) Then
                                    ladrsStartEndCopy.MoveNext
                                End If
                            End If
                        End If

                        ladrsStartEndCopy.Filter = adFilterNone

                        ldtStartDate = !ABS_ST_DT
                        lsStartTime = !ABS_ST_TM
                        ldtEndDate = !ABS_END_DT
                        lsEndTime = !ABS_END_TM
                        ldtRtnDate = !RTN_TO_WRK_DT
                        lsRtnTime = !RTN_TO_WRK_TM
                        ldtStDateWorked = !ST_DATE_WORKED

                        .MoveNext
                        If Not (.EOF) Then
                            If (!ABS_ST_DT = ldtRtnDate And !ABS_ST_TM = lsRtnTime) Or _
                                (!ABS_END_DT = ldtEndDate And !ABS_END_TM = lsEndTime And _
                                CDate(!ST_DATE_WORKED & " " & !ABS_END_TM) > CDate(ldtStDateWorked & " " & lsStartTime)) Then
                                !ABS_ST_DT = ldtStartDate
                                !ABS_ST_TM = lsStartTime
                                !ST_DATE_WORKED = ldtStDateWorked
                                .Update

                                .MovePrevious
                                .Delete adAffectCurrent
                                .Update
                            End If
                        End If
                    End If
                End If

                ladrsStartEndCopy.Filter = adFilterNone

                If Not (.EOF) Then
                    .MoveNext
                End If
            Loop Until (.EOF)

            'Flatten dates with return to work falling between trade flattened date
            llTradeCounter = llTradeCounter * 2

            For llExcIndex = 1 To llTradeCounter
'Debug.Print StringFormat("{0} of {1} TradeCounter Loop", llExcIndex, llTradeCounter)
                ladrsStartEndCopy.Sort = "ST_DATE_WORKED, ABS_ST_DT, ABS_ST_TM, ABS_END_DT, ABS_END_TM"
                ladrsStartEndCopy.MoveFirst

                .Sort = "ST_DATE_WORKED, ABS_ST_DT, ABS_ST_TM, ABS_END_DT, ABS_END_TM"

                For llSchCopyIndex = 1 To ladrsStartEndCopy.RecordCount
'Debug.Print StringFormat("   {0} of {1} ladrsStartEndCopy Loop1", llSchCopyIndex, ladrsStartEndCopy.RecordCount)
                    .MoveFirst

                    For llDateIndex = 1 To .RecordCount
'Debug.Print StringFormat("      {0} of {1} DateIndex Loop 1", llDateIndex, .RecordCount)
                        lbMove = False

                        If CDate(!ST_DATE_WORKED & " " & !ABS_ST_TM) < CDate(ladrsStartEndCopy!RTN_TO_WRK_DT & " " & ladrsStartEndCopy!RTN_TO_WRK_TM) And _
                            CDate(!ABS_END_DT & " " & !ABS_END_TM) > CDate(ladrsStartEndCopy!RTN_TO_WRK_DT & " " & ladrsStartEndCopy!RTN_TO_WRK_TM) And _
                            CDate(!ST_DATE_WORKED & " " & !ABS_ST_TM) <= CDate(ladrsStartEndCopy!ST_DATE_WORKED & " " & ladrsStartEndCopy!ABS_ST_TM) And _
                            CDate(!ABS_END_DT & " " & !ABS_END_TM) >= CDate(ladrsStartEndCopy!ABS_END_DT & " " & ladrsStartEndCopy!ABS_END_TM) And _
                            CDate(!RTN_TO_WRK_DT & " " & !RTN_TO_WRK_TM) >= CDate(ladrsStartEndCopy!RTN_TO_WRK_DT & " " & ladrsStartEndCopy!RTN_TO_WRK_TM) Then
                            'Set the original to vars.
                            ldtEndDate = !ABS_END_DT
                            lsEndTime = !ABS_END_TM
                            ldtRtnDate = !RTN_TO_WRK_DT
                            lsRtnTime = !RTN_TO_WRK_TM

                            lvBookMark = .Bookmark

                            .Filter = adFilterNone
                            .Filter = "ABS_ST_DT = #" & ladrsStartEndCopy!ABS_ST_DT & "#" & _
                                " AND ABS_ST_TM = '" & ladrsStartEndCopy!ABS_ST_TM & "'" & _
                                " AND ST_DATE_WORKED = #" & ladrsStartEndCopy!ST_DATE_WORKED & "#" & _
                                " AND ABS_END_DT = #" & ladrsStartEndCopy!ABS_END_DT & "#" & _
                                " AND ABS_END_TM = '" & ladrsStartEndCopy!ABS_END_TM & "'"
                            If Not (.EOF) Then
                                !ABS_END_DT = ldtEndDate
                                !ABS_END_TM = lsEndTime
                                !RTN_TO_WRK_DT = ldtRtnDate
                                !RTN_TO_WRK_TM = lsRtnTime
                                .Update
                            End If

                            .Filter = adFilterNone

                            .Bookmark = lvBookMark
                            If Not (.EOF) Then
                                ladrsStartEndCopy!ABS_END_DT = ldtEndDate
                                ladrsStartEndCopy!ABS_END_TM = lsEndTime
                                ladrsStartEndCopy!RTN_TO_WRK_DT = ldtRtnDate
                                ladrsStartEndCopy!RTN_TO_WRK_TM = lsRtnTime
                                ladrsStartEndCopy.Update

                                If Not (ladrsStartEndCopy.EOF) Then
                                    lvBookMark = ladrsStartEndCopy.Bookmark
                                    ladrsStartEndCopy.Filter = adFilterNone

                                    ladrsStartEndCopy.Filter = "ABS_ST_DT = #" & !ABS_ST_DT & "#" & _
                                                        " AND ABS_ST_TM = '" & !ABS_ST_TM & "'" & _
                                                        " AND ST_DATE_WORKED = #" & !ST_DATE_WORKED & "#" & _
                                                        " AND ABS_END_DT = #" & !ABS_END_DT & "#" & _
                                                        " AND ABS_END_TM = '" & !ABS_END_TM & "'"

                                    If Not (ladrsStartEndCopy.EOF) Then
                                        ladrsStartEndCopy.Delete adAffectCurrent
                                        ladrsStartEndCopy.Update
                                        lbMove = True
                                    End If

                                    ladrsStartEndCopy.Filter = adFilterNone

                                    ladrsStartEndCopy.Bookmark = lvBookMark
                                End If

                                .Delete adAffectCurrent
                                .Update
                            End If
                        End If

                        If Not (.EOF) Then
                            .MoveNext

                            If (.EOF) Then
                                Exit For
                            End If
                        Else
                            Exit For
                        End If

                        If (lbMove) Then
                            If Not (ladrsStartEndCopy.EOF) Then
                                ladrsStartEndCopy.MoveNext
                                If (ladrsStartEndCopy.EOF) Then
                                    Exit For
                                End If
                            Else
                                Exit For
                            End If
                        End If
                        DoEvents
                    Next llDateIndex

                    If Not (ladrsStartEndCopy.EOF) Then
                        ladrsStartEndCopy.MoveNext
                        If (ladrsStartEndCopy.EOF) Then
                            Exit For
                        End If
                    Else
                        Exit For
                    End If
                    DoEvents
                Next llSchCopyIndex
'Debug.Print "entering botton loop"
                ladrsStartEndCopy.Sort = "ST_DATE_WORKED, ABS_ST_DT, ABS_ST_TM, ABS_END_DT, ABS_END_TM"
                .Sort = "ST_DATE_WORKED, ABS_ST_DT, ABS_ST_TM, ABS_END_DT, ABS_END_TM"
                .MoveFirst

                For llDateIndex = 1 To .RecordCount
'Debug.Print StringFormat("   {0} of {1} DateLoop2", llDateIndex, .RecordCount)
                    ladrsStartEndCopy.MoveFirst

                    For llSchCopyIndex = 1 To ladrsStartEndCopy.RecordCount
'Debug.Print StringFormat("      {0} of {1} ladrsStartEndCopy2", llDateIndex, .RecordCount)
                        lbMove = False

                        If CDate(!RTN_TO_WRK_DT & " " & !RTN_TO_WRK_TM) > CDate(ladrsStartEndCopy!RTN_TO_WRK_DT & " " & ladrsStartEndCopy!RTN_TO_WRK_TM) And _
                            CDate(!ST_DATE_WORKED & " " & !ABS_ST_TM) <= CDate(ladrsStartEndCopy!ST_DATE_WORKED & " " & ladrsStartEndCopy!ABS_ST_TM) And _
                            CDate(!ABS_END_DT & " " & !ABS_END_TM) >= CDate(ladrsStartEndCopy!ABS_END_DT & " " & ladrsStartEndCopy!ABS_END_TM) Then
                            lvBookMark = .Bookmark

                            .Filter = adFilterNone
                            .Filter = ("ABS_ST_DT = #" & ladrsStartEndCopy!ABS_ST_DT & "#" & _
                                " AND ABS_ST_TM = '" & ladrsStartEndCopy!ABS_ST_TM & "'" & _
                                " AND ABS_END_DT = #" & ladrsStartEndCopy!ABS_END_DT & "#" & _
                                " AND ABS_END_TM = '" & ladrsStartEndCopy!ABS_END_TM & "'" & _
                                " AND RTN_TO_WRK_DT = #" & ladrsStartEndCopy!RTN_TO_WRK_DT & "#" & _
                                " AND RTN_TO_WRK_TM = '" & ladrsStartEndCopy!RTN_TO_WRK_TM & "'" & _
                                " AND ST_DATE_WORKED = #" & ladrsStartEndCopy!ST_DATE_WORKED & "#")
                            If Not (.EOF) Then
                                .Delete adAffectCurrent
                                .Update

                                ladrsStartEndCopy.Delete adAffectCurrent
                                ladrsStartEndCopy.Update
                                lbMove = True
                            End If

                            .Filter = adFilterNone
                            .Bookmark = lvBookMark
                        End If

                        If Not (ladrsStartEndCopy.EOF) Then
                            ladrsStartEndCopy.MoveNext
                        Else
                            Exit For
                        End If

                        If (lbMove) Then
                            If Not (.EOF) Then
                                .MoveNext

                                If (.EOF) Then
                                    Exit For
                                End If
                            Else
                                Exit For
                            End If
                        End If
                        DoEvents
                    Next llSchCopyIndex

                    If Not (.EOF) Then
                        .MoveNext

                        If (.EOF) Then
                            Exit For
                        End If
                    Else
                        Exit For
                    End If
                    DoEvents
                Next llDateIndex
                DoEvents
            Next llExcIndex

            'Flatten dates with same return to work date or return to work overlapping absence
            For llExcIndex = 1 To ladrsStartEndCopy.RecordCount
                ladrsStartEndCopy.Sort = "ST_DATE_WORKED, ABS_ST_DT, ABS_ST_TM, ABS_END_DT, ABS_END_TM"
                ladrsStartEndCopy.MoveFirst

                For llDateIndex = 1 To ladrsStartEndCopy.RecordCount
                    .Filter = adFilterNone

                    .Filter = "(RTN_TO_WRK_DT = #" & ladrsStartEndCopy!RTN_TO_WRK_DT & "#" & _
                                " AND RTN_TO_WRK_TM = '" & ladrsStartEndCopy!RTN_TO_WRK_TM & "')"
                    If Not (.EOF) Then
                        .Sort = "ST_DATE_WORKED, ABS_ST_DT, ABS_ST_TM, ABS_END_DT, ABS_END_TM"
                        .MoveLast

                        ldtEndDate = !ABS_END_DT
                        lsEndTime = !ABS_END_TM

                        .MoveFirst
                        !ABS_END_DT = ldtEndDate
                        !ABS_END_TM = lsEndTime
                        .Update

                        .MoveNext
                        Do While Not (.EOF)
                            .Delete adAffectCurrent
                            .Update
                            .MoveNext
                        Loop
                    End If

                    .Filter = adFilterNone

                    ladrsStartEndCopy.MoveNext
                    DoEvents
                Next llDateIndex

                'Flatten dates with same return to work date as the start date
                .Sort = "ST_DATE_WORKED, ABS_ST_DT, ABS_ST_TM, ABS_END_DT, ABS_END_TM"
                .MoveFirst

                For llDateIndex = 1 To .RecordCount
                    ldtStartDate = !ABS_ST_DT
                    lsStartTime = !ABS_ST_TM
                    ldtEndDate = !ABS_END_DT
                    lsEndTime = !ABS_END_TM
                    ldtRtnDate = !RTN_TO_WRK_DT
                    lsRtnTime = !RTN_TO_WRK_TM
                    ldtStDateWorked = !ST_DATE_WORKED

                    .MoveNext
                    If Not (.EOF) Then
                        If (!ABS_ST_DT = ldtRtnDate And !ABS_ST_TM = lsRtnTime) Or _
                            (CDate(!ST_DATE_WORKED & " " & !ABS_ST_TM) <= CDate(ldtRtnDate & " " & lsRtnTime) And _
                            CDate(ldtRtnDate & " " & lsRtnTime) <= CDate(!ABS_END_DT & " " & !ABS_END_TM)) Then
                            !ABS_ST_DT = ldtStartDate
                            !ABS_ST_TM = lsStartTime
                            !ST_DATE_WORKED = ldtStDateWorked
                            .Update

                            .MovePrevious
                            .Delete adAffectCurrent
                            .Update

                            If Not (.EOF) Then
                                .MoveNext
                            End If
                        End If
                    End If
                    DoEvents
                Next llDateIndex
                DoEvents
            Next llExcIndex

            ladrsStartEndCopy.Sort = "ST_DATE_WORKED, ABS_ST_DT, ABS_ST_TM, ABS_END_DT, ABS_END_TM"
            ladrsStartEndCopy.MoveFirst

            For llDateIndex = 1 To ladrsStartEndCopy.RecordCount
                .Filter = adFilterNone

                .Filter = ("RTN_TO_WRK_DT = #" & ladrsStartEndCopy!RTN_TO_WRK_DT & "#" & _
                            " AND RTN_TO_WRK_TM = '" & ladrsStartEndCopy!RTN_TO_WRK_TM & "'")
                If Not (.EOF) Then
                    .Sort = "ST_DATE_WORKED, ABS_ST_DT, ABS_ST_TM, ABS_END_DT, ABS_END_TM"
                    .MoveLast

                    ldtEndDate = !ABS_END_DT
                    lsEndTime = !ABS_END_TM

                    .MoveFirst
                    !ABS_END_DT = ldtEndDate
                    !ABS_END_TM = lsEndTime
                    .Update

                    .MoveNext
                    Do While Not (.EOF)
                        .Delete adAffectCurrent
                        .Update
                        .MoveNext
                    Loop
                End If

                .Filter = adFilterNone

                ladrsStartEndCopy.MoveNext
                DoEvents
            Next llDateIndex

            If Not (.EOF And .BOF) Then
                .MoveFirst
            End If
        End With
    End If

    Set ladrsStartEndCopy = Nothing

    Set DetermineAbsenceStartAndEndTemp = ladrsStartAndEnd

    Set ladrsStartAndEnd = Nothing

    Exit Function
eh:
    Set ladrsAbsExcCopy = Nothing
    Set ladrsSchCopy = Nothing

    Set ladrsStartAndEnd = Nothing
    Set ladrsStartEndCopy = Nothing

    TrsRaiseError MODULENAME, "DetermineAbsenceStartAndEndTemp"
End Function

'//just needed for testing in order to handle assigning a null value
Private Function CopyRecordset2( _
            ByVal vrsBaseRecordset As ADODB.Recordset, _
            Optional ByVal vintAffect As Integer = 2, _
            Optional ByVal Position As Integer = 0, _
            Optional ByVal v_fldAttrib As Long = adFldUpdatable) As ADODB.Recordset
' vintAffect determines the extent of the copy operation.
' 0 - copy metadata only
' 1 - copy current record only
' 2 - copy all records visible
    Dim lrs As ADODB.Recordset
    Dim i As Integer
On Error GoTo eh

    If RecordsExist(vrsBaseRecordset) Then
        vrsBaseRecordset.MoveFirst
    End If
    Set lrs = New ADODB.Recordset
    For i = 1 To vrsBaseRecordset.Fields.Count
        Select Case vrsBaseRecordset.Fields(i - 1).Type
            Case adSmallInt
                lrs.Fields.Append vrsBaseRecordset.Fields(i - 1).Name, adSmallInt, , v_fldAttrib
            Case adInteger
                lrs.Fields.Append vrsBaseRecordset.Fields(i - 1).Name, adInteger, , v_fldAttrib
            Case adDate, adDBDate
                lrs.Fields.Append vrsBaseRecordset.Fields(i - 1).Name, adDate, , v_fldAttrib
            Case adDecimal, adNumeric
                lrs.Fields.Append vrsBaseRecordset.Fields(i - 1).Name, adDecimal, , v_fldAttrib
                lrs.Fields(i - 1).Precision = vrsBaseRecordset.Fields(i - 1).Precision
                lrs.Fields(i - 1).NumericScale = vrsBaseRecordset.Fields(i - 1).NumericScale
            Case adChar
                lrs.Fields.Append vrsBaseRecordset.Fields(i - 1).Name, adChar, vrsBaseRecordset.Fields(i - 1).DefinedSize, v_fldAttrib
            Case adVarChar
                lrs.Fields.Append vrsBaseRecordset.Fields(i - 1).Name, adVarChar, vrsBaseRecordset.Fields(i - 1).DefinedSize, v_fldAttrib
            Case adLongVarChar
                lrs.Fields.Append vrsBaseRecordset.Fields(i - 1).Name, adLongVarChar, vrsBaseRecordset.Fields(i - 1).DefinedSize, v_fldAttrib
            Case Else
                lrs.Fields.Append vrsBaseRecordset.Fields(i - 1).Name, adVariant, , v_fldAttrib
        End Select
    Next
    If vintAffect = 0 Then
        'CopyMetadata
    ElseIf vintAffect = 2 Then
        'CopyAll
        Dim j As Long
        lrs.Open
        For j = 1 To vrsBaseRecordset.RecordCount
            lrs.AddNew
            For i = 1 To vrsBaseRecordset.Fields.Count
'            if vrsBaseRecordset.Fields(i - 1).Attributes=
                If Not IsNull(vrsBaseRecordset.Fields(i - 1).Value) Then
                    lrs.Fields(i - 1).Value = vrsBaseRecordset.Fields(i - 1).Value
                Else
                    lrs.Fields(i - 1).Value = vbNull
                End If
            Next
            lrs.Update
            vrsBaseRecordset.MoveNext
        Next
    Else
        'Copy only the position
        vrsBaseRecordset.AbsolutePosition = Position
        lrs.Open
        lrs.AddNew
        For i = 1 To vrsBaseRecordset.Fields.Count
            lrs.Fields(i - 1).Value = vrsBaseRecordset.Fields(i - 1).Value
        Next
        lrs.Update
    End If
    
    Set CopyRecordset2 = lrs
    Set lrs = Nothing
    
Exit Function
eh:
    Set CopyRecordset2 = Nothing
    Set lrs = Nothing
    Err.Raise Err.Number, "mADOExtensions.CopyRecordset2", Err.Description
End Function

'//needed for trimming strings and for assigning a null value
Public Sub AppendRecordset2( _
            ByVal vrsSource As ADODB.Recordset, _
            ByRef vrsTarget As ADODB.Recordset, _
            Optional ByVal vintAffect As Integer = 2, _
            Optional ByVal Position As Integer = 0)
' copy records from the source recordset to the target recordset.
' both recordsets must have the same definition and open.
' vintAffect determines the extent of the copy operation.
' 0 - copy metadata only, no effect for append operation
' 1 - copy current record only
' 2 - copy all records visible
    Dim lrs As ADODB.Recordset
    Dim i As Integer
On Error GoTo eh
    If RecordsExist(vrsSource) Then
        vrsSource.MoveFirst
    Else
        Exit Sub
    End If

    If vintAffect = 0 Then
        'do nothing
    ElseIf vintAffect = 2 Then
        'CopyAll
        Dim j As Integer
        For j = 1 To vrsSource.RecordCount
            vrsTarget.AddNew
            For i = 1 To vrsSource.Fields.Count
                If Not IsNull(vrsSource.Fields(i - 1).Value) Then
                    If vrsSource.Fields(i - 1).Type = adChar Or vrsSource.Fields(i - 1).Type = adVarChar Then
                        vrsTarget.Fields(i - 1).Value = Trim(vrsSource.Fields(i - 1).Value)
                    Else
                        vrsTarget.Fields(i - 1).Value = vrsSource.Fields(i - 1).Value
                    End If
                Else
                    vrsTarget.Fields(i - 1).Value = vbNull
                End If
            Next
            vrsTarget.Update
            vrsSource.MoveNext
        Next
    Else
        'Copy only the position
        vrsSource.AbsolutePosition = Position
        vrsTarget.AddNew
        For i = 1 To vrsSource.Fields.Count
            vrsTarget.Fields(i - 1).Value = vrsSource.Fields(i - 1).Value
        Next
        vrsTarget.Update
    End If
    
    Exit Sub
    
eh:
    Err.Raise Err.Number, "mADOExtensions.AppendRecordset2", Err.Description
End Sub

'//use dateworked (ST_DATE_WORKED) instead of charge date (ABS_ST_DT) in vadrsAbsenceStartEnd
Private Sub CreateAndAddAllAbsenceRecords(ByVal vlEmployeeNumber As Long _
                , ByVal vadrsInitialAbsence As ADODB.Recordset _
                , ByVal vadrsPassedDate As ADODB.Recordset _
                , ByVal vadrsPassedRTW As ADODB.Recordset _
                , ByVal vadrsDatesAlreadyPresent As ADODB.Recordset _
                , ByVal vadrsSchedule As ADODB.Recordset _
                , ByVal vadrsAbsenceStartEnd As ADODB.Recordset _
                , ByRef radrsInitial As ADODB.Recordset _
                , ByRef radrsDate As ADODB.Recordset _
                , ByRef radrsReturnToWork As ADODB.Recordset)
On Error GoTo eh
Dim llIndex As Long, llNumDays As Long
Dim llDateIndex As Long, ladrs As ADODB.Recordset
Dim ldtMinDate As Date, ldtMaxDate As Date
Dim lbAddRecord As Boolean, llRTWIndex As Long
Dim lbRTWAvailable As Boolean, lsRecordStatus As String

    If (RecordsExist(vadrsPassedDate)) And _
        (RecordsExist(vadrsAbsenceStartEnd)) Then
        Set ladrs = New ADODB.Recordset
        With ladrs
            .Fields.Append "ABS_DATE", adDate, , ADOFIELDATTR
            .Fields.Append "CREATED_TIMESTAMP", adChar, 26, ADOFIELDATTR
            .Fields.Append "RECORD_STATUS", adChar, 1, ADOFIELDATTR

            .CursorLocation = adUseClient
            .CursorType = adOpenStatic
            .LockType = adLockBatchOptimistic
            .Open
        End With

        With vadrsPassedDate
            .MoveFirst

            For llDateIndex = 1 To .RecordCount
                vadrsAbsenceStartEnd.Filter = adFilterNone

                vadrsAbsenceStartEnd.Filter = ( _
                    "ABS_ST_DT <= #" & !ABS_DATE & _
                    "# AND RTN_TO_WRK_DT >=#" & !ABS_DATE & "#")

                If Not (vadrsAbsenceStartEnd.EOF) Then
                    ladrs.AddNew
                    ladrs!ABS_DATE = !ABS_DATE
                    ladrs!RECORD_STATUS = !RECORD_STATUS

                    If (!RECORD_STATUS = "I") Then
                        ladrs!CREATED_TIMESTAMP = ""
                    Else
                        If (!RECORD_STATUS = "D") Then
                            ladrs!CREATED_TIMESTAMP = !CREATED_TIMESTAMP
                        End If
                    End If

                    ladrs.Update
                Else
                    If (!RECORD_STATUS = "D") Then
                        ladrs.AddNew
                        ladrs!ABS_DATE = !ABS_DATE
                        ladrs!CREATED_TIMESTAMP = !CREATED_TIMESTAMP
                        ladrs!RECORD_STATUS = !RECORD_STATUS
                        ladrs.Update
                    End If
                End If

                vadrsAbsenceStartEnd.Filter = adFilterNone

                .MoveNext
            Next llDateIndex

            .MoveFirst
        End With
    
        With vadrsAbsenceStartEnd
            .Sort = "ABS_ST_DT, ABS_ST_TM, ABS_END_DT, ABS_END_TM, RTN_TO_WRK_DT, RTN_TO_WRK_TM"

            .MoveFirst
            ldtMinDate = !ABS_ST_DT
            
            .MoveLast
            ldtMaxDate = !ABS_END_DT

            .MoveFirst
        End With
    End If

    Set radrsInitial = New ADODB.Recordset
    With radrsInitial
        .Fields.Append "EMP_NUM", adInteger, , ADOFIELDATTR
        .Fields.Append "ABS_NUM", adInteger, , ADOFIELDATTR

        .Fields.Append "ABS_ST_DT", adDate, , ADOFIELDATTR
        .Fields.Append "ST_DATE_WORKED", adDate, , ADOFIELDATTR
        .Fields.Append "ABS_ST_TM", adChar, 5, ADOFIELDATTR
        
        .Fields.Append "ABS_END_DT", adDate, , ADOFIELDATTR
        .Fields.Append "END_DATE_WORKED", adDate, , ADOFIELDATTR
        .Fields.Append "ABS_END_TM", adChar, 5, ADOFIELDATTR

        .Fields.Append "ABS_RPT_TS", adChar, 26, ADOFIELDATTR

        .Fields.Append "ABS_CALLIN_DT", adDate, , ADOFIELDATTR
        .Fields.Append "ABS_CALLIN_TM", adChar, 5, ADOFIELDATTR

        .Fields.Append "ABS_LOC_CD", adInteger, , ADOFIELDATTR

        .Fields.Append "ABS_RPTR_TYPE", adVarChar, 30, ADOFIELDATTR
        .Fields.Append "ABS_RPTR_TXT", adVarChar, 255, ADOFIELDATTR

        .Fields.Append "EMP_PHONE", adChar, 10, ADOFIELDATTR
        .Fields.Append "EMP_LOC_CD", adInteger, , ADOFIELDATTR

        .Fields.Append "EXCEPTION_CD", adChar, 3, ADOFIELDATTR
        .Fields.Append "OVER_EXC_CD", adChar, 3, ADOFIELDATTR
        .Fields.Append "ADDL_TXT", adVarChar, 255, ADOFIELDATTR

        .Fields.Append "FMLA_REASON", adVarChar, 20, ADOFIELDATTR

        .Fields.Append "RTN_TO_WRK_DT", adDate, , ADOFIELDATTR
        .Fields.Append "RTW_DATE_WORKED", adDate, , ADOFIELDATTR
        .Fields.Append "RTN_TO_WRK_TM", adChar, 5, ADOFIELDATTR

        .Fields.Append "FOLW_UP_DT", adDate, , ADOFIELDATTR
        .Fields.Append "RTW_NOT_REQD_IND", adChar, 1, ADOFIELDATTR
        .Fields.Append "RPTNG_EMP_NUM", adInteger, , ADOFIELDATTR
        .Fields.Append "ABS_PRIM_CONT_EMP", adInteger, , ADOFIELDATTR
        .Fields.Append "ABS_NOTIFY_WG_NUM", adInteger, , ADOFIELDATTR
        .Fields.Append "PRVDR_NM", adVarChar, 30, ADOFIELDATTR
        .Fields.Append "PRVDR_PHN_NBR", adChar, 10, ADOFIELDATTR

        .Fields.Append "CREATED_EMP_NUM", adInteger, , ADOFIELDATTR
        .Fields.Append "CREATED_TIMESTAMP", adChar, 26, ADOFIELDATTR
        .Fields.Append "LAST_MOD_EMP_NUM", adInteger, , ADOFIELDATTR
        .Fields.Append "LAST_MOD_TIMESTAMP", adChar, 26, ADOFIELDATTR

        .Fields.Append "RECORD_STATUS", adChar, 1, ADOFIELDATTR

        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open
    End With

    Set radrsDate = New ADODB.Recordset
    With radrsDate
        .Fields.Append "ABS_NUM", adInteger, , ADOFIELDATTR
        .Fields.Append "ABS_DATE", adDate, , ADOFIELDATTR

        .Fields.Append "CREATED_EMP_NUM", adInteger, , ADOFIELDATTR
        .Fields.Append "CREATED_TIMESTAMP", adChar, 26, ADOFIELDATTR
        .Fields.Append "LAST_MOD_EMP_NUM", adInteger, , ADOFIELDATTR
        .Fields.Append "LAST_MOD_TIMESTAMP", adChar, 26, ADOFIELDATTR

        .Fields.Append "RECORD_STATUS", adChar, 1, ADOFIELDATTR

        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open
    End With

    Set radrsReturnToWork = New ADODB.Recordset
    With radrsReturnToWork
        .Fields.Append "ABS_NUM", adInteger, , ADOFIELDATTR
        .Fields.Append "RTN_TO_WRK_DT", adDate, , ADOFIELDATTR
        .Fields.Append "RTN_TO_WRK_TM", adChar, 5, ADOFIELDATTR

        .Fields.Append "CONFRM_RSN_IND", adChar, 1, ADOFIELDATTR
        .Fields.Append "ABS_REASON_TXT", adVarChar, 255, ADOFIELDATTR

        .Fields.Append "MEET_EMP_IND", adChar, 1, ADOFIELDATTR
        .Fields.Append "MEET_EMP_TXT", adVarChar, 255, ADOFIELDATTR
        .Fields.Append "PRFM_DUTY_IND", adChar, 1, ADOFIELDATTR
        .Fields.Append "PRFM_DUTY_TXT", adVarChar, 255, ADOFIELDATTR
        .Fields.Append "SIDE_EFF_IND", adChar, 1, ADOFIELDATTR
        .Fields.Append "SIDE_EFF_TXT", adVarChar, 255, ADOFIELDATTR

        .Fields.Append "ADDL_TXT", adVarChar, 255, ADOFIELDATTR

        .Fields.Append "RPTNG_EMP_NUM", adInteger, , ADOFIELDATTR
        .Fields.Append "RPT_DT", adDate, , ADOFIELDATTR
        .Fields.Append "CREATED_TIMESTAMP", adChar, 26, ADOFIELDATTR
        .Fields.Append "RECORD_STATUS", adChar, 1, adFldUpdatable

        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open
    End With

    If (RecordsExist(vadrsAbsenceStartEnd) And _
        RecordsExist(vadrsInitialAbsence)) Then
        vadrsInitialAbsence.MoveFirst
        vadrsAbsenceStartEnd.MoveFirst

        For llIndex = 1 To vadrsAbsenceStartEnd.RecordCount
            With radrsInitial
                .AddNew

                !EMP_NUM = vlEmployeeNumber
                If (vadrsInitialAbsence!ABS_NUM > 0) Then
                    If (vadrsInitialAbsence!ST_DATE_WORKED = vadrsAbsenceStartEnd!ST_DATE_WORKED) And _
                        (vadrsInitialAbsence!ABS_ST_TM = vadrsAbsenceStartEnd!ABS_ST_TM) And _
                        (vadrsInitialAbsence!RTN_TO_WRK_DT = vadrsAbsenceStartEnd!RTN_TO_WRK_DT) And _
                        (vadrsInitialAbsence!RTN_TO_WRK_TM = vadrsAbsenceStartEnd!RTN_TO_WRK_TM) Then
                        !ABS_NUM = vadrsInitialAbsence!ABS_NUM
                        !ABS_END_DT = vadrsAbsenceStartEnd!ABS_END_DT
                        !ABS_END_TM = vadrsAbsenceStartEnd!ABS_END_TM
                        !CREATED_TIMESTAMP = vadrsInitialAbsence!CREATED_TIMESTAMP
                        !RECORD_STATUS = vadrsInitialAbsence!RECORD_STATUS
                        If (RecordsExist(ladrs)) Then
                            With ladrs
                                .MoveFirst

                                For llDateIndex = 1 To .RecordCount
                                    lbAddRecord = True

                                    If (RecordsExist(radrsDate)) Then
                                        radrsDate.Filter = adFilterNone
                                        radrsDate.Filter = ("ABS_DATE = #" & !ABS_DATE & _
                                                            "# AND RECORD_STATUS = '" & _
                                                            !RECORD_STATUS & "'")
                                        If Not (radrsDate.EOF) Then
                                            lbAddRecord = False
                                        End If

                                        radrsDate.Filter = adFilterNone
                                    End If

                                    If (lbAddRecord) Then
                                        radrsDate.AddNew

                                        radrsDate!ABS_NUM = vadrsInitialAbsence!ABS_NUM
                                        radrsDate!ABS_DATE = !ABS_DATE
                                        radrsDate!CREATED_TIMESTAMP = !CREATED_TIMESTAMP
                                        radrsDate!RECORD_STATUS = !RECORD_STATUS
                                        radrsDate.Update
                                    End If

                                    .MoveNext
                                Next llDateIndex
                            End With
                        End If

                        If (ldtMaxDate > vadrsAbsenceStartEnd!ABS_END_DT) Then
                            llNumDays = Abs(DateDiff("d", ldtMaxDate _
                                            , vadrsAbsenceStartEnd!ABS_END_DT))
                            For llDateIndex = 1 To llNumDays
                                lbAddRecord = True

                                If (RecordsExist(radrsDate)) Then
                                    radrsDate.Filter = adFilterNone

                                    radrsDate.Filter = ("ABS_DATE = #" & _
                                        DateAdd("d", CDbl(llDateIndex) _
                                                , vadrsAbsenceStartEnd!ABS_END_DT) & _
                                        "# AND RECORD_STATUS = 'D'")
                                    If Not (radrsDate.EOF) Then
                                        lbAddRecord = False
                                    End If

                                    radrsDate.Filter = adFilterNone
                                End If

                                If (lbAddRecord) Then
                                    radrsDate.AddNew

                                    radrsDate!ABS_NUM = vadrsInitialAbsence!ABS_NUM
                                    radrsDate!ABS_DATE = DateAdd("d", CDbl(llDateIndex) _
                                                        , vadrsAbsenceStartEnd!ABS_END_DT)
                                    radrsDate!CREATED_TIMESTAMP = ""
                                    radrsDate!RECORD_STATUS = "D"

                                    radrsDate.Update
                                End If
                            Next llDateIndex
                        Else
                            llNumDays = Abs(DateDiff("d", vadrsAbsenceStartEnd!ABS_ST_DT _
                                            , vadrsAbsenceStartEnd!ABS_END_DT))
                            For llDateIndex = 0 To llNumDays - 1
                                lbAddRecord = True

                                If (RecordsExist(radrsDate)) Then
                                    radrsDate.Filter = adFilterNone

                                    radrsDate.Filter = ("ABS_DATE = #" & _
                                                DateAdd("d", CDbl(llDateIndex) _
                                                , vadrsAbsenceStartEnd!ABS_ST_DT) & "#")
                                    If Not (radrsDate.EOF) Then
                                        lbAddRecord = False
                                    End If

                                    radrsDate.Filter = adFilterNone
                                End If

                                If (RecordsExist(vadrsDatesAlreadyPresent)) Then
                                    vadrsDatesAlreadyPresent.Filter = adFilterNone

                                    vadrsDatesAlreadyPresent.Filter = ("ABS_DATE = #" & _
                                                DateAdd("d", CDbl(llDateIndex) _
                                                , vadrsAbsenceStartEnd!ABS_ST_DT) & "#")
                                    If Not (vadrsDatesAlreadyPresent.EOF) Then
                                        lbAddRecord = False
                                    End If

                                    vadrsDatesAlreadyPresent.Filter = adFilterNone
                                End If

                                If (RecordsExist(vadrsSchedule)) Then
                                    vadrsSchedule.Filter = adFilterNone

                                    vadrsSchedule.Filter = ("CHARGE_DATE = #" & _
                                            DateAdd("d", CDbl(llDateIndex) _
                                                , vadrsAbsenceStartEnd!ABS_ST_DT) & _
                                            "# AND SHIFT_HOURS = 0.0" & _
                                            " AND EXCEPTION_CD <> 'ONP'")
                                    If Not (vadrsSchedule.EOF) Then
                                        lbAddRecord = False
                                    End If

                                    vadrsSchedule.Filter = adFilterNone
                                End If

                                If (lbAddRecord) Then
                                    radrsDate.AddNew

                                    radrsDate!ABS_NUM = vadrsInitialAbsence!ABS_NUM
                                    radrsDate!ABS_DATE = DateAdd("d", CDbl(llDateIndex) _
                                                        , vadrsAbsenceStartEnd!ABS_ST_DT)
                                    radrsDate!CREATED_TIMESTAMP = ""
                                    radrsDate!RECORD_STATUS = "I"

                                    radrsDate.Update
                                End If
                            Next llDateIndex
                        End If
                    Else '//not a full match
                        If (vadrsInitialAbsence!ST_DATE_WORKED = vadrsAbsenceStartEnd!ST_DATE_WORKED) And _
                            (vadrsInitialAbsence!ABS_ST_TM = vadrsAbsenceStartEnd!ABS_ST_TM) Then
                            !ABS_NUM = vadrsInitialAbsence!ABS_NUM
                            !ABS_END_DT = vadrsAbsenceStartEnd!ABS_END_DT
                            !ABS_END_TM = vadrsAbsenceStartEnd!ABS_END_TM
                            !CREATED_TIMESTAMP = vadrsInitialAbsence!CREATED_TIMESTAMP
                            !RECORD_STATUS = vadrsInitialAbsence!RECORD_STATUS

                            If (RecordsExist(ladrs)) Then
                                With ladrs
                                    .MoveFirst

                                    For llDateIndex = 1 To .RecordCount
                                        If (!ABS_DATE < vadrsAbsenceStartEnd!RTN_TO_WRK_DT) Then
                                            lbAddRecord = True

                                            If (RecordsExist(radrsDate)) Then
                                                radrsDate.Filter = adFilterNone

                                                radrsDate.Filter = ("ABS_DATE = #" & _
                                                                    !ABS_DATE & _
                                                                    "# AND RECORD_STATUS = '" & _
                                                                    !RECORD_STATUS & "'")
                                                If Not (radrsDate.EOF) Then
                                                    lbAddRecord = False
                                                End If

                                                radrsDate.Filter = adFilterNone
                                            End If

                                            If (RecordsExist(vadrsSchedule)) Then
                                                vadrsSchedule.Filter = adFilterNone

                                                vadrsSchedule.Filter = ("CHARGE_DATE = #" & _
                                                        !ABS_DATE & "# AND SHIFT_HOURS = 0.0" & _
                                                        " AND EXCEPTION_CD <> 'ONP'")
                                                If Not (vadrsSchedule.EOF) Then
                                                    lbAddRecord = False
                                                End If

                                                vadrsSchedule.Filter = adFilterNone
                                            End If

                                            If (lbAddRecord) Then
                                                radrsDate.AddNew

                                                radrsDate!ABS_NUM = vadrsInitialAbsence!ABS_NUM
                                                radrsDate!ABS_DATE = !ABS_DATE
                                                radrsDate!CREATED_TIMESTAMP = !CREATED_TIMESTAMP
                                                radrsDate!RECORD_STATUS = !RECORD_STATUS

                                                radrsDate.Update
                                            End If
                                        End If

                                        .MoveNext
                                    Next llDateIndex
                                End With
                            End If

                            If (ldtMaxDate > vadrsAbsenceStartEnd!ABS_END_DT) Then
                                llNumDays = Abs(DateDiff("d", ldtMaxDate _
                                            , vadrsAbsenceStartEnd!ABS_END_DT))
                                For llDateIndex = 1 To llNumDays
                                    lbAddRecord = True

                                    If (RecordsExist(radrsDate)) Then
                                        radrsDate.Filter = adFilterNone

                                        radrsDate.Filter = ("ABS_DATE = #" & _
                                            DateAdd("d", CDbl(llDateIndex) _
                                                , vadrsAbsenceStartEnd!ABS_END_DT) & _
                                            "# AND RECORD_STATUS = 'D'")
                                        If Not (radrsDate.EOF) Then
                                            lbAddRecord = False
                                        End If

                                        radrsDate.Filter = adFilterNone
                                    End If

                                    If (lbAddRecord) Then
                                        radrsDate.AddNew

                                        radrsDate!ABS_NUM = vadrsInitialAbsence!ABS_NUM
                                        radrsDate!ABS_DATE = DateAdd("d", CDbl(llDateIndex) _
                                                        , vadrsAbsenceStartEnd!ABS_END_DT)
                                        radrsDate!CREATED_TIMESTAMP = ""
                                        radrsDate!RECORD_STATUS = "D"

                                        radrsDate.Update
                                    End If
                                Next llDateIndex
                            End If
                        Else '//not a full match and not a partial match
                            !ABS_NUM = -1 * llIndex
                            !RECORD_STATUS = "I"

                            llNumDays = Abs(DateDiff("d", vadrsAbsenceStartEnd!ABS_ST_DT, vadrsAbsenceStartEnd!ABS_END_DT))
                            '//MKS 716815 : TRIS - Employee Absence Management is not saving scheduled exception for 2nd day
                            '//It is actually including an additional day, if the end time is "00:00"
                            If vadrsAbsenceStartEnd!ABS_END_TM = "00:00" Then
                                If llNumDays >= 1 Then
                                    llNumDays = llNumDays - 1
                                End If
                            End If
                            
                            Select Case (llNumDays)
                                Case 0
                                    lbAddRecord = True

                                    If (RecordsExist(vadrsSchedule)) Then
                                        vadrsSchedule.Filter = adFilterNone

                                        vadrsSchedule.Filter = ("CHARGE_DATE = #" & _
                                                vadrsAbsenceStartEnd!ABS_ST_DT & _
                                                "# AND SHIFT_HOURS = 0.0" & _
                                                " AND EXCEPTION_CD <> 'ONP'")
                                        If Not (vadrsSchedule.EOF) Then
                                            lbAddRecord = False
                                        End If

                                        vadrsSchedule.Filter = adFilterNone
                                    End If

                                    If (lbAddRecord) Then
                                        radrsDate.AddNew
                                        radrsDate!ABS_NUM = -1 * llIndex
                                        radrsDate!ABS_DATE = vadrsAbsenceStartEnd!ABS_ST_DT
                                        radrsDate!RECORD_STATUS = "I"
                                        radrsDate.Update
                                    End If

                                Case Is > 0
                                    For llDateIndex = 0 To llNumDays
                                        lbAddRecord = True

                                        If (RecordsExist(vadrsSchedule)) Then
                                            vadrsSchedule.Filter = adFilterNone

                                            vadrsSchedule.Filter = ("CHARGE_DATE = #" & _
                                                        DateAdd("d", CDbl(llDateIndex) _
                                                            , vadrsAbsenceStartEnd!ABS_ST_DT) & _
                                                        "# AND SHIFT_HOURS = 0.0" & _
                                                        " AND EXCEPTION_CD <> 'ONP'")
                                            If Not (vadrsSchedule.EOF) Then
                                                lbAddRecord = False
                                            End If

                                            vadrsSchedule.Filter = adFilterNone
                                        End If

                                        If (lbAddRecord) Then
                                            radrsDate.AddNew
                                            radrsDate!ABS_NUM = -1 * llIndex
                                            radrsDate!ABS_DATE = DateAdd("d" _
                                                        , CDbl(llDateIndex) _
                                                        , vadrsAbsenceStartEnd!ABS_ST_DT)
                                            radrsDate!RECORD_STATUS = "I"
                                            radrsDate.Update
                                        End If
                                    Next llDateIndex
                            End Select
                        End If
                    End If
                Else '//not ABS_NUM > 0, new absence
                    !ABS_NUM = -1 * llIndex
                    !RECORD_STATUS = "I"

                    llNumDays = Abs(DateDiff("d", vadrsAbsenceStartEnd!ABS_ST_DT, vadrsAbsenceStartEnd!ABS_END_DT))
                    
                    '//MKS 716815 : TRIS - Employee Absence Management is not saving scheduled exception for 2nd day
                    '//It is actually including an additional day, if the end time is "00:00"
'                    If vadrsAbsenceStartEnd!ABS_END_TM = "00:00" Then
'                        If llNumDays >= 1 Then
'                            llNumDays = llNumDays - 1
'                        End If
'                    End If
                    
                    Select Case (llNumDays)
                        Case 0
                            lbAddRecord = True

                            If (RecordsExist(vadrsSchedule)) Then
                                vadrsSchedule.Filter = adFilterNone

                                vadrsSchedule.Filter = ("CHARGE_DATE = #" & _
                                        vadrsAbsenceStartEnd!ABS_ST_DT & _
                                        "# AND SHIFT_HOURS = 0.0" & _
                                        " AND EXCEPTION_CD <> 'ONP'")
                                If Not (vadrsSchedule.EOF) Then
                                    lbAddRecord = False
                                End If

                                vadrsSchedule.Filter = adFilterNone
                            End If

                            If (lbAddRecord) Then
                                radrsDate.AddNew
                                radrsDate!ABS_NUM = -1 * llIndex
                                radrsDate!ABS_DATE = vadrsAbsenceStartEnd!ABS_ST_DT
                                radrsDate!RECORD_STATUS = "I"
                                radrsDate.Update
                            End If

                        Case Is > 0
                            For llDateIndex = 0 To llNumDays
                                lbAddRecord = True

                                If (RecordsExist(vadrsSchedule)) Then
                                    vadrsSchedule.Filter = adFilterNone

                                    vadrsSchedule.Filter = ("CHARGE_DATE = #" & _
                                                        DateAdd("d", CDbl(llDateIndex) _
                                                            , vadrsAbsenceStartEnd!ABS_ST_DT) & _
                                                        "# AND SHIFT_HOURS = 0.0" & _
                                                        " AND EXCEPTION_CD <> 'ONP'")
                                    If Not (vadrsSchedule.EOF) Then
                                        lbAddRecord = False
                                    End If

                                    vadrsSchedule.Filter = adFilterNone
                                End If

                                If (lbAddRecord) Then
                                    radrsDate.AddNew
                                    radrsDate!ABS_NUM = -1 * llIndex
                                    radrsDate!ABS_DATE = DateAdd("d" _
                                        , CDbl(llDateIndex) _
                                        , vadrsAbsenceStartEnd!ABS_ST_DT)
                                    radrsDate!RECORD_STATUS = "I"
                                    radrsDate.Update
                                End If
                            Next llDateIndex
                    End Select
                End If

                !ABS_ST_DT = vadrsAbsenceStartEnd!ABS_ST_DT
                !ST_DATE_WORKED = vadrsAbsenceStartEnd!ST_DATE_WORKED
                !ABS_ST_TM = vadrsAbsenceStartEnd!ABS_ST_TM
                
                !ABS_END_DT = vadrsAbsenceStartEnd!ABS_END_DT
                !END_DATE_WORKED = vadrsAbsenceStartEnd!END_DATE_WORKED
                !ABS_END_TM = vadrsAbsenceStartEnd!ABS_END_TM

                !ABS_RPT_TS = vadrsInitialAbsence!ABS_RPT_TS
                !ABS_CALLIN_DT = vadrsInitialAbsence!ABS_CALLIN_DT
                !ABS_CALLIN_TM = vadrsInitialAbsence!ABS_CALLIN_TM
                !ABS_LOC_CD = vadrsInitialAbsence!ABS_LOC_CD
                !ABS_RPTR_TYPE = vadrsInitialAbsence!ABS_RPTR_TYPE
                !ABS_RPTR_TXT = vadrsInitialAbsence!ABS_RPTR_TXT
                !EMP_PHONE = vadrsInitialAbsence!EMP_PHONE
                !EMP_LOC_CD = vadrsInitialAbsence!EMP_LOC_CD
                !EXCEPTION_CD = vadrsInitialAbsence!EXCEPTION_CD
                !OVER_EXC_CD = vadrsInitialAbsence!OVER_EXC_CD
                !ADDL_TXT = vadrsInitialAbsence!ADDL_TXT
                !FMLA_REASON = vadrsInitialAbsence!FMLA_REASON

                !RTN_TO_WRK_DT = vadrsAbsenceStartEnd!RTN_TO_WRK_DT
                !RTW_DATE_WORKED = vadrsAbsenceStartEnd!RTW_DATE_WORKED
                !RTN_TO_WRK_TM = vadrsAbsenceStartEnd!RTN_TO_WRK_TM

                !RTW_NOT_REQD_IND = vadrsInitialAbsence!RTW_NOT_REQD_IND
                !FOLW_UP_DT = vadrsInitialAbsence!FOLW_UP_DT
                !RPTNG_EMP_NUM = vadrsInitialAbsence!RPTNG_EMP_NUM
                !ABS_PRIM_CONT_EMP = vadrsInitialAbsence!ABS_PRIM_CONT_EMP

                !PRVDR_NM = vadrsInitialAbsence!PRVDR_NM
                !PRVDR_PHN_NBR = vadrsInitialAbsence!PRVDR_PHN_NBR
                !ABS_NOTIFY_WG_NUM = vadrsInitialAbsence!ABS_NOTIFY_WG_NUM

                .Update
            End With

            vadrsAbsenceStartEnd.MoveNext
        Next llIndex

        vadrsAbsenceStartEnd.MoveFirst
    End If

    lbRTWAvailable = False
    If (RecordsExist(vadrsPassedRTW)) Then
        vadrsPassedRTW.MoveFirst
        lbRTWAvailable = True
    End If

    If (RecordsExist(radrsInitial)) Then
        radrsInitial.MoveFirst

        For llIndex = 1 To radrsInitial.RecordCount
            Select Case (vadrsInitialAbsence!RTW_NOT_REQD_IND)
                Case "Y", "1", "YES"
                    lbAddRecord = True

                    If (lbRTWAvailable) Then
                        vadrsPassedRTW.Filter = adFilterNone
                        vadrsPassedRTW.Filter = ("RTN_TO_WRK_DT = #" & radrsInitial!RTN_TO_WRK_DT & "#" & _
                                                    " AND RTN_TO_WRK_TM = '" & radrsInitial!RTN_TO_WRK_TM & "'")
                        If Not (vadrsPassedRTW.EOF) Then
                            With radrsReturnToWork
                                .AddNew
                                !ABS_NUM = radrsInitial!ABS_NUM
                                !RTN_TO_WRK_DT = vadrsPassedRTW!RTN_TO_WRK_DT
                                !RTN_TO_WRK_TM = vadrsPassedRTW!RTN_TO_WRK_TM
                                !CONFRM_RSN_IND = vadrsPassedRTW!CONFRM_RSN_IND
                                !ABS_REASON_TXT = vadrsPassedRTW!ABS_REASON_TXT
                                !MEET_EMP_IND = vadrsPassedRTW!MEET_EMP_IND
                                !MEET_EMP_TXT = vadrsPassedRTW!MEET_EMP_TXT
                                !PRFM_DUTY_IND = vadrsPassedRTW!PRFM_DUTY_IND
                                !PRFM_DUTY_TXT = vadrsPassedRTW!PRFM_DUTY_TXT
                                !SIDE_EFF_IND = vadrsPassedRTW!SIDE_EFF_IND
                                !SIDE_EFF_TXT = vadrsPassedRTW!SIDE_EFF_TXT
                                !ADDL_TXT = vadrsPassedRTW!ADDL_TXT
                                !RPTNG_EMP_NUM = vadrsPassedRTW!RPTNG_EMP_NUM
                                !RPT_DT = vadrsPassedRTW!RPT_DT
                                If (!ABS_NUM > 0) Then
                                    If Not (IsNull(vadrsPassedRTW!CREATED_TIMESTAMP)) Then
                                        !CREATED_TIMESTAMP = vadrsPassedRTW!CREATED_TIMESTAMP
                                    Else
                                        !CREATED_TIMESTAMP = ""
                                    End If
                                    !RECORD_STATUS = vadrsPassedRTW!RECORD_STATUS
                                Else
                                    !CREATED_TIMESTAMP = ""
                                    !RECORD_STATUS = "I"
                                End If
                                .Update

                                lbAddRecord = False
                            End With
                        Else
                            vadrsPassedRTW.Filter = adFilterNone
                            vadrsPassedRTW.MoveFirst

                            With radrsReturnToWork
                                .AddNew
                                !ABS_NUM = radrsInitial!ABS_NUM
                                !RTN_TO_WRK_DT = radrsInitial!RTN_TO_WRK_DT
                                !RTN_TO_WRK_TM = radrsInitial!RTN_TO_WRK_TM
                                !CONFRM_RSN_IND = vadrsPassedRTW!CONFRM_RSN_IND
                                !ABS_REASON_TXT = vadrsPassedRTW!ABS_REASON_TXT
                                !MEET_EMP_IND = vadrsPassedRTW!MEET_EMP_IND
                                !MEET_EMP_TXT = vadrsPassedRTW!MEET_EMP_TXT
                                !PRFM_DUTY_IND = vadrsPassedRTW!PRFM_DUTY_IND
                                !PRFM_DUTY_TXT = vadrsPassedRTW!PRFM_DUTY_TXT
                                !SIDE_EFF_IND = vadrsPassedRTW!SIDE_EFF_IND
                                !SIDE_EFF_TXT = vadrsPassedRTW!SIDE_EFF_TXT
                                !ADDL_TXT = vadrsPassedRTW!ADDL_TXT
                                !RPTNG_EMP_NUM = vadrsPassedRTW!RPTNG_EMP_NUM
                                !RPT_DT = vadrsPassedRTW!RPT_DT
                                If (radrsInitial!ABS_NUM > 0) Then
                                    If Not (IsNull(vadrsPassedRTW!CREATED_TIMESTAMP)) Then
                                        !CREATED_TIMESTAMP = vadrsPassedRTW!CREATED_TIMESTAMP
                                        !RECORD_STATUS = "C"
                                    Else
                                        !CREATED_TIMESTAMP = ""
                                        !RECORD_STATUS = "I"
                                    End If
                                Else
                                    !CREATED_TIMESTAMP = ""
                                    !RECORD_STATUS = "I"
                                End If
                                .Update
                            End With

                            lbAddRecord = False
                        End If
                    End If

                    If (lbAddRecord) Then
                        With radrsReturnToWork
                            .AddNew
                            !ABS_NUM = radrsInitial!ABS_NUM
                            !RTN_TO_WRK_DT = radrsInitial!RTN_TO_WRK_DT
                            !RTN_TO_WRK_TM = radrsInitial!RTN_TO_WRK_TM
                            !CONFRM_RSN_IND = "N"
                            !ABS_REASON_TXT = ""
                            !MEET_EMP_IND = "N"
                            !MEET_EMP_TXT = ""
                            !PRFM_DUTY_IND = "N"
                            !PRFM_DUTY_TXT = ""
                            !SIDE_EFF_IND = "N"
                            !SIDE_EFF_TXT = ""
                            !ADDL_TXT = ""
                            !RPTNG_EMP_NUM = radrsInitial!RPTNG_EMP_NUM
                            'USE CURRENT DATE
                            !RPT_DT = Date
                            !CREATED_TIMESTAMP = ""
                            !RECORD_STATUS = "I"
                            .Update

                            lbAddRecord = False
                        End With
                    End If

                Case Else
                    If (lbRTWAvailable) Then
                        Set radrsReturnToWork = vadrsPassedRTW
                    End If
            End Select

            radrsInitial.MoveNext
        Next llIndex

        radrsInitial.MoveFirst
    End If

    Set ladrs = Nothing

    Exit Sub
eh:
    Set ladrs = Nothing

    TrsRaiseError MODULENAME _
        , "CreateAndAddAllAbsenceRecords"

End Sub

Private Function ValidateInitialAbsence(ByRef radrsInitialAbsence As ADODB.Recordset _
                , ByRef radrsBRV As ADODB.Recordset _
                ) As Boolean
On Error GoTo ValidateInitialAbsence_EH
Dim llIndex As Long

    ValidateInitialAbsence = False

    With radrsInitialAbsence
        .MoveFirst

        For llIndex = 1 To .RecordCount
            If Not (CheckForValidLong(!ABS_NUM)) Then
                AppendBusinessError radrsBRV _
                    , trErrbusAbsenceInitial _
                    , mstrCBUSINESSERR _
                    , "Invalid Absence Number." _
                    , MODULENAME & ": ValidateInitialAbsence"
            End If

            If Not (IsNull(!ABS_ST_DT)) Then
                If Not (CheckForValidDate(!ABS_ST_DT)) Then
                    AppendBusinessError radrsBRV _
                        , trErrbusAbsenceInitial _
                        , mstrCBUSINESSERR _
                        , "Invalid Absence Start Date." _
                        , MODULENAME & ": ValidateInitialAbsence"
                End If
            Else
                AppendBusinessError radrsBRV _
                    , trErrbusAbsenceInitial _
                    , mstrCBUSINESSERR _
                    , "Invalid Absence Start Date." _
                    , MODULENAME & ": ValidateInitialAbsence"
            End If

            If Not (IsNull(!ABS_ST_TM)) Then
                If Not (CheckForValidString(!ABS_ST_TM, , , True, 5)) Then
                    AppendBusinessError radrsBRV _
                        , trErrbusAbsenceInitial _
                        , mstrCBUSINESSERR _
                        , "Invalid Absence Start Time." _
                        , MODULENAME & ": ValidateInitialAbsence"
                End If
            Else
                AppendBusinessError radrsBRV _
                    , trErrbusAbsenceInitial _
                    , mstrCBUSINESSERR _
                    , "Invalid Absence Start Time." _
                    , MODULENAME & ": ValidateInitialAbsence"
            End If

            If Not (IsNull(!ABS_RPT_TS)) Then
                If (!RECORD_STATUS = "I") Then
                    If (IsDate(!ABS_RPT_TS)) Then
                        !ABS_RPT_TS = Format(!ABS_RPT_TS, "YYYY-MM-DD-hh.mm.ss.000000")
                        .Update
                    Else
                        !ABS_RPT_TS = Format(Now, "YYYY-MM-DD-hh.mm.ss.000000")
                        .Update
                    End If
                End If

                If Not (CheckForValidString(!ABS_RPT_TS, , , True, 26)) Then
                    AppendBusinessError radrsBRV _
                        , trErrbusAbsenceInitial _
                        , mstrCBUSINESSERR _
                        , "Invalid Absence Reported Date/Time." _
                        , MODULENAME & ": ValidateInitialAbsence"
                End If
            Else
                AppendBusinessError radrsBRV _
                    , trErrbusAbsenceInitial _
                    , mstrCBUSINESSERR _
                    , "Invalid Absence Reported Date/Time." _
                    , MODULENAME & ": ValidateInitialAbsence"
            End If

            If Not (IsNull(!ABS_CALLIN_DT)) Then
                If Not (CheckForValidDate(!ABS_CALLIN_DT)) Then
                    AppendBusinessError radrsBRV _
                        , trErrbusAbsenceInitial _
                        , mstrCBUSINESSERR _
                        , "Invalid Absence Called In Date." _
                        , MODULENAME & ": ValidateInitialAbsence"
                End If
            Else
                AppendBusinessError radrsBRV _
                    , trErrbusAbsenceInitial _
                    , mstrCBUSINESSERR _
                    , "Invalid Absence Called In Date." _
                    , MODULENAME & ": ValidateInitialAbsence"
            End If

            If Not (IsNull(!ABS_CALLIN_TM)) Then
                If Not (CheckForValidString(!ABS_CALLIN_TM, , , True, 5)) Then
                    AppendBusinessError radrsBRV _
                        , trErrbusAbsenceInitial _
                        , mstrCBUSINESSERR _
                        , "Invalid Absence Called In Time." _
                        , MODULENAME & ": ValidateInitialAbsence"
                End If
            Else
                AppendBusinessError radrsBRV _
                    , trErrbusAbsenceInitial _
                    , mstrCBUSINESSERR _
                    , "Invalid Absence Called In Time." _
                    , MODULENAME & ": ValidateInitialAbsence"
            End If

            If Not (CheckForValidInteger(!ABS_LOC_CD, True)) Then
                AppendBusinessError radrsBRV _
                    , trErrbusAbsenceInitial _
                    , mstrCBUSINESSERR _
                    , "Invalid Absence Location." _
                    , MODULENAME & ": ValidateInitialAbsence"
            Else
                If Not (CheckForExistence(trLOCATION, , , !ABS_LOC_CD)) Then
                    AppendBusinessError radrsBRV _
                        , trErrbusAbsenceInitial _
                        , mstrCBUSINESSERR _
                        , "Invalid Absence Location." _
                        , MODULENAME & ": ValidateInitialAbsence"
                End If
            End If

            If Not (IsNull(!ABS_RPTR_TYPE)) Then
                If Not (CheckForValidString(!ABS_RPTR_TYPE, True, 10)) Then
                    AppendBusinessError radrsBRV _
                        , trErrbusAbsenceInitial _
                        , mstrCBUSINESSERR _
                        , "Invalid Contacted by Type." _
                        , MODULENAME & ": ValidateInitialAbsence"
                End If
            Else
                AppendBusinessError radrsBRV _
                    , trErrbusAbsenceInitial _
                    , mstrCBUSINESSERR _
                    , "Invalid Contacted by Type." _
                    , MODULENAME & ": ValidateInitialAbsence"
            End If

            If Not (IsNull(!ABS_RPTR_TXT)) Then
                If (Trim(!ABS_RPTR_TXT) <> "") Then
                    !ABS_RPTR_TXT = ParseString(!ABS_RPTR_TXT)
                    .Update

                    If Not (CheckForValidString(!ABS_RPTR_TXT, True, 255)) Then
                        AppendBusinessError radrsBRV _
                            , trErrbusAbsenceInitial _
                            , mstrCBUSINESSERR _
                            , "Invalid Contacted by Comments. Exceeds maximum character limit of 255." _
                            , MODULENAME & ": ValidateInitialAbsence"
                    End If
                End If
            End If

            If Not (IsNull(!EMP_PHONE)) Then
                If (Trim(!EMP_PHONE) <> "") Then
                    If Not (CheckForValidString(!EMP_PHONE, True, 10)) Then
                        AppendBusinessError radrsBRV _
                            , trErrbusAbsenceInitial _
                            , mstrCBUSINESSERR _
                            , "Invalid Employee Phone. Exceeds maximum character limit of 10." _
                            , MODULENAME & ": ValidateInitialAbsence"
                    End If
                End If
            End If

            If Not (CheckForValidInteger(!EMP_LOC_CD, True)) Then
                AppendBusinessError radrsBRV _
                    , trErrbusAbsenceInitial _
                    , mstrCBUSINESSERR _
                    , "Invalid Employee Home Location." _
                    , MODULENAME & ": ValidateInitialAbsence"
            Else
                If Not (CheckForExistence(trLOCATION, , , !EMP_LOC_CD)) Then
                    AppendBusinessError radrsBRV _
                        , trErrbusAbsenceInitial _
                        , mstrCBUSINESSERR _
                        , "Invalid Employee Home Location." _
                        , MODULENAME & ": ValidateInitialAbsence"
                End If
            End If

            If Not (IsNull(!EXCEPTION_CD)) Then
                If Not (CheckForValidString(!EXCEPTION_CD, True, 3)) Then
                    AppendBusinessError radrsBRV _
                        , trErrbusAbsenceInitial _
                        , mstrCBUSINESSERR _
                        , "Invalid Exception Code." _
                        , MODULENAME & ": ValidateInitialAbsence"
                Else
                    !EXCEPTION_CD = UCase(!EXCEPTION_CD)
                    .Update
                End If
            Else
                AppendBusinessError radrsBRV _
                    , trErrbusAbsenceInitial _
                    , mstrCBUSINESSERR _
                    , "Invalid Exception Code." _
                    , MODULENAME & ": ValidateInitialAbsence"
            End If

            If Not (IsNull(!OVER_EXC_CD)) Then
                If (Trim(!OVER_EXC_CD) <> "") Then
                    If Not (CheckForValidString(!OVER_EXC_CD, True, 3)) Then
                        AppendBusinessError radrsBRV _
                            , trErrbusAbsenceInitial _
                            , mstrCBUSINESSERR _
                            , "Invalid Override Code." _
                            , MODULENAME & ": ValidateInitialAbsence"
                    Else
                        !OVER_EXC_CD = UCase(!OVER_EXC_CD)
                        .Update
                    End If
                End If
            End If

            If Not (IsNull(!ADDL_TXT)) Then
                !ADDL_TXT = ParseString(!ADDL_TXT)
                .Update

                If Not (CheckForValidString(!ADDL_TXT, True, 255)) Then
                    AppendBusinessError radrsBRV _
                        , trErrbusAbsenceInitial _
                        , mstrCBUSINESSERR _
                        , "Invalid Reason for Absence. May exceeds maximum character limit of 255 or is empty." _
                        , MODULENAME & ": ValidateInitialAbsence"
                End If
            Else
                AppendBusinessError radrsBRV _
                    , trErrbusAbsenceInitial _
                    , mstrCBUSINESSERR _
                    , "Invalid Reason for Absence." _
                    , MODULENAME & ": ValidateInitialAbsence"
            End If

            If Not (IsNull(!FMLA_REASON)) Then
                Select Case (Trim(!FMLA_REASON))
                    Case "NOT APPLICABLE", "NA"
                        !FMLA_REASON = ""
                        .Update
                    Case "BIRTH OR ADOPTION", "BA"
                        !FMLA_REASON = "BA"
                        .Update
                    Case "CARE OF CHILD", "CC"
                        !FMLA_REASON = "CC"
                        .Update
                    Case "CARE OF SPOUSE", "CS"
                        !FMLA_REASON = "CS"
                        .Update
                    Case "CARE OF PARENT", "CP"
                        !FMLA_REASON = "CP"
                        .Update
                    Case "EMPLOYEE ILLNESS", "EI"
                        !FMLA_REASON = "EI"
                        .Update
                    Case Else
                        AppendBusinessError radrsBRV _
                            , trErrbusAbsenceInitial _
                            , mstrCBUSINESSERR _
                            , "Invalid Absence FMLA Reason." _
                            , MODULENAME & ": ValidateInitialAbsence"
                End Select
            End If

            If Not (IsNull(!FOLW_UP_DT)) Then
                If Not (CheckForValidDate(!FOLW_UP_DT)) Then
                    AppendBusinessError radrsBRV _
                        , trErrbusAbsenceInitial _
                        , mstrCBUSINESSERR _
                        , "Invalid Absence Start Date." _
                        , MODULENAME & ": ValidateInitialAbsence"
                End If
            End If

            If Not (IsNull(!RTW_NOT_REQD_IND)) Then
                Select Case (Trim(!RTW_NOT_REQD_IND))
                    Case "1"
                        !RTW_NOT_REQD_IND = "Y"
                        .Update
                    Case "0", ""
                        !RTW_NOT_REQD_IND = "N"
                        .Update
                    Case "Y", "N"
                        '
                    Case Else
                        AppendBusinessError radrsBRV _
                            , trErrbusAbsenceContact _
                            , mstrCBUSINESSERR _
                            , "Invalid, Return to Work not required indicator." _
                            , MODULENAME & ": ValidateInitialAbsence"
                End Select
            Else
                AppendBusinessError radrsBRV _
                    , trErrbusAbsenceContact _
                    , mstrCBUSINESSERR _
                    , "Invalid, Return to Work not required indicator." _
                    , MODULENAME & ": ValidateInitialAbsence"
            End If

            If Not (CheckForValidDate(!RTN_TO_WRK_DT) Or _
                CheckForValidString(!RTN_TO_WRK_TM, , , True, 5)) Then
                AppendBusinessError radrsBRV _
                    , trErrbusAbsenceInitial _
                    , mstrCBUSINESSERR _
                    , "Invalid Probable Return to Work Date." _
                    , MODULENAME & ": ValidateInitialAbsence"
            End If

            If Not (CheckForValidLong(!RPTNG_EMP_NUM, True)) Then
                AppendBusinessError radrsBRV _
                    , trErrbusAbsenceInitial _
                    , mstrCBUSINESSERR _
                    , "Invalid Reporting Employee." _
                    , MODULENAME & ": ValidateInitialAbsence"
            Else
                If Not (CheckForExistence(trEMPLOYEE, , !RPTNG_EMP_NUM)) Then
                    AppendBusinessError radrsBRV _
                        , trErrbusAbsenceInitial _
                        , mstrCBUSINESSERR _
                        , "Invalid Reporting Employee. Employee " & _
                            CStr(!RPTNG_EMP_NUM) & " does not exist." _
                        , MODULENAME & ": ValidateInitialAbsence"
                End If
            End If
    
            If Not (CheckForValidLong(!ABS_PRIM_CONT_EMP, True)) Then
                AppendBusinessError radrsBRV _
                    , trErrbusAbsenceInitial _
                    , mstrCBUSINESSERR _
                    , "Invalid Absence Contact Employee." _
                    , MODULENAME & ": ValidateInitialAbsence"
            Else
                If Not (CheckForExistence(trEMPLOYEE, , !ABS_PRIM_CONT_EMP)) Then
                    AppendBusinessError radrsBRV _
                        , trErrbusAbsenceInitial _
                        , mstrCBUSINESSERR _
                        , "Invalid Absence Contact Employee. Employee " & _
                            CStr(!ABS_PRIM_CONT_EMP) & " does not exist." _
                        , MODULENAME & ": ValidateInitialAbsence"
                End If
            End If

            If Not (IsNull(!PRVDR_NM)) Then
                If (Trim(!PRVDR_NM) <> "") Then
                    !PRVDR_NM = ParseString(!PRVDR_NM)
                    .Update

                    If Not (CheckForValidString(!PRVDR_NM, True, 30)) Then
                        AppendBusinessError radrsBRV _
                            , trErrbusAbsenceInitial _
                            , mstrCBUSINESSERR _
                            , "Invalid Provider Name." _
                            , MODULENAME & ": ValidateInitialAbsence"
                    End If
                End If
            Else
                !PRVDR_NM = ""
                .Update
            End If

            If Not (IsNull(!PRVDR_PHN_NBR)) Then
                If (Trim(!PRVDR_PHN_NBR) <> "") Then
                    If Not (CheckForValidString(!PRVDR_PHN_NBR, , , True, 10)) Then
                        AppendBusinessError radrsBRV _
                            , trErrbusAbsenceInitial _
                            , mstrCBUSINESSERR _
                            , "Invalid Provider Phone Number." _
                            , MODULENAME & ": ValidateInitialAbsence"
                    End If
                End If
            Else
                !PRVDR_PHN_NBR = ""
                .Update
            End If

            If Not (IsNull(!ABS_NOTIFY_WG_NUM)) Then
                If (!ABS_NOTIFY_WG_NUM > 0) Then
                    If Not (CheckForExistence(trWORKGROUP, , !ABS_NOTIFY_WG_NUM)) Then
                        AppendBusinessError radrsBRV _
                            , trErrbusAbsenceInitial _
                            , mstrCBUSINESSERR _
                            , "Invalid Absence Notify to Workgroup. Workgroup " & _
                                CStr(!ABS_NOTIFY_WG_NUM) & " does not exist." _
                            , MODULENAME & ": ValidateInitialAbsence"
                    End If
                End If
            Else
                !ABS_NOTIFY_WG_NUM = 0
                .Update
            End If

            Select Case (Trim(!RECORD_STATUS))
                Case "C", "D"
                    If Not (CheckForValidLong(!ABS_NUM, True)) Then
                        AppendBusinessError radrsBRV _
                            , trErrbusAbsenceInitial _
                            , mstrCBUSINESSERR _
                            , "Invalid Absence Number." _
                            , MODULENAME & ": ValidateInitialAbsence"
                    End If

                    If Not (IsNull(!CREATED_TIMESTAMP)) Then
                        If (Trim(!CREATED_TIMESTAMP) <> "") Then
                            If Not (CheckForValidString(!CREATED_TIMESTAMP, , , True, 26)) Then
                                AppendBusinessError radrsBRV _
                                    , trErrbusAbsenceInitial _
                                    , mstrCBUSINESSERR _
                                    , "Invalid Time Stamp." _
                                    , MODULENAME & ": ValidateInitialAbsence"
                            End If
                        Else
                            AppendBusinessError radrsBRV _
                                , trErrbusAbsenceInitial _
                                , mstrCBUSINESSERR _
                                , "Invalid Time Stamp." _
                                , MODULENAME & ": ValidateInitialAbsence"
                        End If
                    Else
                        AppendBusinessError radrsBRV _
                            , trErrbusAbsenceInitial _
                            , mstrCBUSINESSERR _
                            , "Invalid Time Stamp." _
                            , MODULENAME & ": ValidateInitialAbsence"
                    End If
            End Select

            .MoveNext
        Next llIndex

        .MoveFirst
    End With

    If (RecordsExist(radrsBRV)) Then
        radrsBRV.Filter = adFilterNone
        radrsBRV.Filter = "ERR_TYPE = 'BE' OR ERR_TYPE = 'AE'" & _
                " OR ERR_TYPE = 'AW'"
        If Not (radrsBRV.EOF) Then
            radrsBRV.Filter = adFilterNone
            Exit Function
        End If
        radrsBRV.Filter = adFilterNone
    End If

    ValidateInitialAbsence = True

    Exit Function
ValidateInitialAbsence_EH:
    TrsRaiseError MODULENAME _
        , "ValidateInitialAbsence"

End Function
Private Function ValidateAbsenceProvider(ByRef radrsAbsenceProvider As ADODB.Recordset _
                , ByRef radrsBRV As ADODB.Recordset _
                ) As Boolean
On Error GoTo ValidateAbsenceProvider_EH
Dim llIndex As Long

    ValidateAbsenceProvider = False

    With radrsAbsenceProvider
        .MoveFirst

        For llIndex = 1 To .RecordCount
            If Not (CheckForValidLong(!ABS_NUM)) Then
                AppendBusinessError radrsBRV _
                    , trErrbusAbsenceProvider _
                    , mstrCBUSINESSERR _
                    , "Invalid Absence Number." _
                    , MODULENAME & ": ValidateAbsenceProvider"
            End If

            If Not (CheckForValidLong(!PRVDR_NBR)) Then
                AppendBusinessError radrsBRV _
                    , trErrbusAbsenceProvider _
                    , mstrCBUSINESSERR _
                    , "Invalid Provider Number." _
                    , MODULENAME & ": ValidateAbsenceProvider"
            Else
                If (!PRVDR_NBR > 0) Then
                    If Not (CheckForExistence(trPROVIDER, , !PRVDR_NBR)) Then
                        AppendBusinessError radrsBRV _
                            , trErrbusAbsenceProvider _
                            , mstrCBUSINESSERR _
                            , "Invalid Provider Number. Provider " & CStr(!PRVDR_NBR) & _
                                " does not exist." _
                            , MODULENAME & ": ValidateAbsenceProvider"
                    End If
                End If
            End If

            Select Case (!RECORD_STATUS)
                Case "I", "C"
                    If Not (IsNull(!TYP_DESC)) Then
                        If Not (CheckForValidString(!TYP_DESC, True, 30)) Then
                            AppendBusinessError radrsBRV _
                                , trErrbusAbsenceProvider _
                                , mstrCBUSINESSERR _
                                , "Invalid Provider Type." _
                                , MODULENAME & ": ValidateAbsenceProvider"
                        Else
                            Select Case (UCase(!TYP_DESC))
                                Case "DOCTOR", "HOSPITAL" _
                                    , "NURSE PRACTITIONER", "REHAB. FACILITY" _
                                    , "SPECIALIST", "PHYSICAL THERAPIST"
                                    !TYP_DESC = Left(!TYP_DESC, 1)
                                    .Update
                                Case "PHYSICIAN ASSISTANT"
                                    !TYP_DESC = "A"
                                    .Update
                                Case "DENTIST"
                                    !TYP_DESC = "T"
                                    .Update
                                Case Else
                                    AppendBusinessError radrsBRV _
                                        , trErrbusAbsenceProvider _
                                        , mstrCBUSINESSERR _
                                        , "Invalid Provider Type." _
                                        , MODULENAME & ": ValidateAbsenceProvider"
                            End Select
                        End If
                    Else
                        AppendBusinessError radrsBRV _
                            , trErrbusAbsenceProvider _
                            , mstrCBUSINESSERR _
                            , "Invalid Provider Type." _
                            , MODULENAME & ": ValidateAbsenceProvider"
                    End If

                    If Not (IsNull(!PRVDR_NM)) Then
                        !PRVDR_NM = ParseString(!PRVDR_NM)
                        .Update

                        If Not (CheckForValidString(!PRVDR_NM, True, 30)) Then
                            AppendBusinessError radrsBRV _
                                , trErrbusAbsenceProvider _
                                , mstrCBUSINESSERR _
                                , "Invalid Provider Name." _
                                , MODULENAME & ": ValidateAbsenceProvider"
                        End If
                    Else
                        AppendBusinessError radrsBRV _
                            , trErrbusAbsenceProvider _
                            , mstrCBUSINESSERR _
                            , "Invalid Provider Name." _
                            , MODULENAME & ": ValidateAbsenceProvider"
                    End If

                    If Not (IsNull(!LN1_ADDR)) Then
                        If (Trim(!LN1_ADDR) <> "") Then
                            !LN1_ADDR = ParseString(!LN1_ADDR)
                            .Update

                            If Not (CheckForValidString(!LN1_ADDR, True, 255)) Then
                                AppendBusinessError radrsBRV _
                                    , trErrbusAbsenceProvider _
                                    , mstrCBUSINESSERR _
                                    , "Invalid Provider Address Line 1." _
                                    , MODULENAME & ": ValidateAbsenceProvider"
                            End If
                        End If
                    End If

                    If Not (IsNull(!LN2_ADDR)) Then
                        If (Trim(!LN2_ADDR) <> "") Then
                            !LN2_ADDR = ParseString(!LN2_ADDR)
                            .Update

                            If Not (CheckForValidString(!LN2_ADDR, True, 255)) Then
                                AppendBusinessError radrsBRV _
                                    , trErrbusAbsenceProvider _
                                    , mstrCBUSINESSERR _
                                    , "Invalid Provider Address Line 2." _
                                    , MODULENAME & ": ValidateAbsenceProvider"
                            End If
                        End If
                    End If

                    If Not (IsNull(!CTY_NM)) Then
                        If (Trim(!CTY_NM) <> "") Then
                            !CTY_NM = ParseString(!CTY_NM)
                            .Update

                            If Not (CheckForValidString(!CTY_NM, True, 60)) Then
                                AppendBusinessError radrsBRV _
                                    , trErrbusAbsenceProvider _
                                    , mstrCBUSINESSERR _
                                    , "Invalid Provider City." _
                                    , MODULENAME & ": ValidateAbsenceProvider"
                            End If
                        End If
                    End If

                    If Not (IsNull(!ST_CD)) Then
                        If (Trim(!ST_CD) <> "") Then
                            If Not (CheckForValidString(!ST_CD, , , True, 2)) Then
                                AppendBusinessError radrsBRV _
                                    , trErrbusAbsenceProvider _
                                    , mstrCBUSINESSERR _
                                    , "Invalid Provider State." _
                                    , MODULENAME & ": ValidateAbsenceProvider"
                            End If
                        End If
                    End If

                    If Not (IsNull(!ZIP_CD)) Then
                        If (Trim(!ZIP_CD) <> "") Then
                            If Not (CheckForValidString(!ZIP_CD, True, 10)) Then
                                AppendBusinessError radrsBRV _
                                    , trErrbusAbsenceProvider _
                                    , mstrCBUSINESSERR _
                                    , "Invalid Provider ZIP Code." _
                                    , MODULENAME & ": ValidateAbsenceProvider"
                            End If
                        End If
                    End If

                    If Not (IsNull(!PHN_NBR)) Then
                        If (Trim(!PHN_NBR) <> "") Then
                            If Not (CheckForValidString(!PHN_NBR, True, 10)) Then
                                AppendBusinessError radrsBRV _
                                    , trErrbusAbsenceProvider _
                                    , mstrCBUSINESSERR _
                                    , "Invalid Provider Phone." _
                                    , MODULENAME & ": ValidateAbsenceProvider"
                            End If
                        Else
                            AppendBusinessError radrsBRV _
                                , trErrbusAbsenceProvider _
                                , mstrCBUSINESSERR _
                                , "Invalid Provider Phone." _
                                , MODULENAME & ": ValidateAbsenceProvider"
                        End If
                    Else
                        AppendBusinessError radrsBRV _
                            , trErrbusAbsenceProvider _
                            , mstrCBUSINESSERR _
                            , "Invalid Provider Phone." _
                            , MODULENAME & ": ValidateAbsenceProvider"
                    End If

                    If Not (IsNull(!CLS_TXT)) Then
                        !CLS_TXT = ParseString(!CLS_TXT)
                        .Update

                        If (Trim(!CLS_TXT) <> "") Then
                            If Not (CheckForValidString(!CLS_TXT, True, 30)) Then
                                AppendBusinessError radrsBRV _
                                    , trErrbusAbsenceProvider _
                                    , mstrCBUSINESSERR _
                                    , "Invalid Provider Class." _
                                    , MODULENAME & ": ValidateAbsenceProvider"
                            End If
                        End If
                    End If

                    If Not (IsNull(!VST_ST_ADM_DT)) Then
                        If Not (CheckForValidDate(!VST_ST_ADM_DT)) Then
                            AppendBusinessError radrsBRV _
                                , trErrbusAbsenceProvider _
                                , mstrCBUSINESSERR _
                                , "Invalid Visit Start/Admission Date." _
                                , MODULENAME & ": ValidateAbsenceProvider"
                        End If
                    Else
                        AppendBusinessError radrsBRV _
                            , trErrbusAbsenceProvider _
                            , mstrCBUSINESSERR _
                            , "Invalid Visit Start/Admission Date." _
                            , MODULENAME & ": ValidateAbsenceProvider"
                    End If

                    If Not (IsNull(!VST_ST_ADM_TM)) Then
                        If Not (CheckForValidString(!VST_ST_ADM_TM, , , True, 5)) Then
                            AppendBusinessError radrsBRV _
                                , trErrbusAbsenceProvider _
                                , mstrCBUSINESSERR _
                                , "Invalid Visit Start/Admission Time." _
                                , MODULENAME & ": ValidateAbsenceProvider"
                        End If
                    Else
                        AppendBusinessError radrsBRV _
                            , trErrbusAbsenceProvider _
                            , mstrCBUSINESSERR _
                            , "Invalid Visit Start/Admission Time." _
                            , MODULENAME & ": ValidateAbsenceProvider"
                    End If

                    If Not (IsNull(!VST_END_DGCE_DT)) Then
                        If Not (CheckForValidDate(!VST_END_DGCE_DT)) Then
                            If Not (IsNull(!VST_ST_ADM_DT)) Then
                                If (CheckForValidDate(!VST_ST_ADM_DT)) Then
                                    !VST_END_DGCE_DT = !VST_ST_ADM_DT
                                    .Update
                                Else
                                    AppendBusinessError radrsBRV _
                                        , trErrbusAbsenceProvider _
                                        , mstrCBUSINESSERR _
                                        , "Invalid Visit End/Discharge Date." _
                                        , MODULENAME & ": ValidateAbsenceProvider"
                                End If
                            Else
                                AppendBusinessError radrsBRV _
                                    , trErrbusAbsenceProvider _
                                    , mstrCBUSINESSERR _
                                    , "Invalid Visit End/Discharge Date." _
                                    , MODULENAME & ": ValidateAbsenceProvider"
                            End If
                        End If
                    Else
                        If Not (IsNull(!VST_ST_ADM_DT)) Then
                            If (CheckForValidDate(!VST_ST_ADM_DT)) Then
                                !VST_END_DGCE_DT = !VST_ST_ADM_DT
                                .Update
                            Else
                                AppendBusinessError radrsBRV _
                                    , trErrbusAbsenceProvider _
                                    , mstrCBUSINESSERR _
                                    , "Invalid Visit End/Discharge Date." _
                                    , MODULENAME & ": ValidateAbsenceProvider"
                            End If
                        Else
                            AppendBusinessError radrsBRV _
                                , trErrbusAbsenceProvider _
                                , mstrCBUSINESSERR _
                                , "Invalid Visit End/Discharge Date." _
                                , MODULENAME & ": ValidateAbsenceProvider"
                        End If
                    End If

                    If Not (IsNull(!VST_END_DGCE_TM)) Then
                        If Not (CheckForValidString(!VST_END_DGCE_TM, , , True, 5)) Then
                            If Not (IsNull(!VST_ST_ADM_TM)) Then
                                If (CheckForValidString(!VST_ST_ADM_TM, , , True, 5)) Then
                                    !VST_END_DGCE_TM = !VST_ST_ADM_TM
                                    .Update
                                Else
                                    AppendBusinessError radrsBRV _
                                        , trErrbusAbsenceProvider _
                                        , mstrCBUSINESSERR _
                                        , "Invalid Visit End/Discharge Time." _
                                        , MODULENAME & ": ValidateAbsenceProvider"
                                End If
                            Else
                                AppendBusinessError radrsBRV _
                                    , trErrbusAbsenceProvider _
                                    , mstrCBUSINESSERR _
                                    , "Invalid Visit End/Discharge Time." _
                                    , MODULENAME & ": ValidateAbsenceProvider"
                            End If
                        End If
                    Else
                        If Not (IsNull(!VST_ST_ADM_TM)) Then
                            If (CheckForValidString(!VST_ST_ADM_TM, , , True, 5)) Then
                                !VST_END_DGCE_TM = !VST_ST_ADM_TM
                                .Update
                            Else
                                AppendBusinessError radrsBRV _
                                    , trErrbusAbsenceProvider _
                                    , mstrCBUSINESSERR _
                                    , "Invalid Visit End/Discharge Time." _
                                    , MODULENAME & ": ValidateAbsenceProvider"
                            End If
                        Else
                            AppendBusinessError radrsBRV _
                                , trErrbusAbsenceProvider _
                                , mstrCBUSINESSERR _
                                , "Invalid Visit End/Discharge Time." _
                                , MODULENAME & ": ValidateAbsenceProvider"
                        End If
                    End If

                    If Not (IsNull(!RLSE_IND)) Then
                        Select Case (Trim(!RLSE_IND))
                            Case "1"
                                !RLSE_IND = "Y"
                                .Update
                            Case "0", ""
                                !RLSE_IND = "N"
                                .Update
                            Case "Y", "N"
                                '
                            Case Else
                                AppendBusinessError radrsBRV _
                                    , trErrbusAbsenceProvider _
                                    , mstrCBUSINESSERR _
                                    , "Invalid Doctor's release Indicator." _
                                    , MODULENAME & ": ValidateAbsenceProvider"
                        End Select
                    End If

                    If Not (IsNull(!ADDL_TXT)) Then
                        If (Trim(!ADDL_TXT) <> "") Then
                            !ADDL_TXT = ParseString(!ADDL_TXT)
                            .Update

                            If Not (CheckForValidString(!ADDL_TXT, True, 255)) Then
                                AppendBusinessError radrsBRV _
                                    , trErrbusAbsenceProvider _
                                    , mstrCBUSINESSERR _
                                    , "Invalid Visit End/Discharge Time." _
                                    , MODULENAME & ": ValidateAbsenceProvider"
                            End If
                        End If
                    End If
            End Select

            Select Case (Trim(!RECORD_STATUS))
                Case "C", "D"
                    If Not (CheckForValidLong(!ABS_NUM, True)) Then
                        AppendBusinessError radrsBRV _
                            , trErrbusAbsenceProvider _
                            , mstrCBUSINESSERR _
                            , "Invalid Absence Number." _
                            , MODULENAME & ": ValidateAbsenceProvider"
                    End If

                    If Not (IsNull(!CREATED_TIMESTAMP)) Then
                        If (Trim(!CREATED_TIMESTAMP) <> "") Then
                            If Not (CheckForValidString(!CREATED_TIMESTAMP, , , True, 26)) Then
                                AppendBusinessError radrsBRV _
                                    , trErrbusAbsenceProvider _
                                    , mstrCBUSINESSERR _
                                    , "Invalid Time Stamp." _
                                    , MODULENAME & ": ValidateAbsenceProvider"
                            End If
                        Else
                            AppendBusinessError radrsBRV _
                                , trErrbusAbsenceProvider _
                                , mstrCBUSINESSERR _
                                , "Invalid Time Stamp." _
                                , MODULENAME & ": ValidateAbsenceProvider"
                        End If
                    Else
                        AppendBusinessError radrsBRV _
                            , trErrbusAbsenceProvider _
                            , mstrCBUSINESSERR _
                            , "Invalid Time Stamp." _
                            , MODULENAME & ": ValidateAbsenceProvider"
                    End If
            End Select

            .MoveNext
        Next llIndex

        .MoveFirst
    End With

    If (RecordsExist(radrsBRV)) Then
        radrsBRV.Filter = adFilterNone
        radrsBRV.Filter = "ERR_TYPE = 'BE' OR ERR_TYPE = 'AE'" & _
                " OR ERR_TYPE = 'AW'"
        If Not (radrsBRV.EOF) Then
            radrsBRV.Filter = adFilterNone
            Exit Function
        End If
        radrsBRV.Filter = adFilterNone
    End If

    ValidateAbsenceProvider = True

    Exit Function
ValidateAbsenceProvider_EH:
    TrsRaiseError MODULENAME _
        , "ValidateAbsenceProvider"

End Function
Private Function ValidateAbsenceContact(ByRef radrsAbsenceContact As ADODB.Recordset _
                , ByRef radrsBRV As ADODB.Recordset _
                , Optional ByVal vlTRISUserNumber As Long = 0 _
                ) As Boolean
On Error GoTo ValidateAbsenceContact_EH
Dim llIndex As Long

    ValidateAbsenceContact = False

    With radrsAbsenceContact
        .MoveFirst

        For llIndex = 1 To .RecordCount
            If Not (CheckForValidLong(!ABS_NUM)) Then
                AppendBusinessError radrsBRV _
                    , trErrbusAbsenceContact _
                    , mstrCBUSINESSERR _
                    , "Invalid Absence Number." _
                    , MODULENAME & ": ValidateAbsenceContact" _
                    , , , , _
                    , CStr(!ABS_NUM)
            Else
                If (!ABS_NUM > 0) Then
                    If Not (CheckForExistence(trABSENCE, , !ABS_NUM)) Then
                        AppendBusinessError radrsBRV _
                            , trErrbusAbsenceContact _
                            , mstrCBUSINESSERR _
                            , "Invalid Absence Number. Absence Number does not exist." _
                            , MODULENAME & ": ValidateAbsenceContact" _
                            , , , , _
                            , CStr(!ABS_NUM)
                    End If
                End If
            End If

            If Not (CheckForValidLong(!ABS_CONT_NUM)) Then
                AppendBusinessError radrsBRV _
                    , trErrbusAbsenceContact _
                    , mstrCBUSINESSERR _
                    , "Invalid Absence Contact Number." _
                    , MODULENAME & ": ValidateAbsenceContact" _
                    , , , , _
                    , CStr(!ABS_NUM)
            End If

            Select Case (!RECORD_STATUS)
                Case "I", "C"
                    If Not (IsNull(!ABS_CONT_DT)) Then
                        If Not (CheckForValidDate(!ABS_CONT_DT)) Then
                            AppendBusinessError radrsBRV _
                                , trErrbusAbsenceContact _
                                , mstrCBUSINESSERR _
                                , "Invalid Absence Contact Date." _
                                , MODULENAME & ": ValidateAbsenceContact" _
                                , , , , _
                                , CStr(!ABS_NUM)
                        End If
                    Else
                        AppendBusinessError radrsBRV _
                            , trErrbusAbsenceContact _
                            , mstrCBUSINESSERR _
                            , "Invalid Absence Contact Date." _
                            , MODULENAME & ": ValidateAbsenceContact" _
                            , , , , _
                            , CStr(!ABS_NUM)
                    End If

                    If Not (IsNull(!ABS_CONT_TM)) Then
                        If (Trim(!ABS_CONT_TM) <> "") Then
                            If Not (CheckForValidString(!ABS_CONT_TM, , , True, 5)) Then
                                AppendBusinessError radrsBRV _
                                    , trErrbusAbsenceContact _
                                    , mstrCBUSINESSERR _
                                    , "Invalid Absence Contact Time." _
                                    , MODULENAME & ": ValidateAbsenceContact" _
                                    , , , , _
                                    , CStr(!ABS_NUM)
                            End If
                        Else
                            AppendBusinessError radrsBRV _
                                , trErrbusAbsenceContact _
                                , mstrCBUSINESSERR _
                                , "Invalid Absence Contact Time." _
                                , MODULENAME & ": ValidateAbsenceContact" _
                                , , , , _
                                , CStr(!ABS_NUM)
                        End If
                    Else
                        AppendBusinessError radrsBRV _
                            , trErrbusAbsenceContact _
                            , mstrCBUSINESSERR _
                            , "Invalid Absence Contact Time." _
                            , MODULENAME & ": ValidateAbsenceContact" _
                            , , , , _
                            , CStr(!ABS_NUM)
                    End If

                    If Not (IsNull(!USR_INIT_IND)) Then
                        Select Case (Trim(!USR_INIT_IND))
                            Case "1"
                                !USR_INIT_IND = "Y"
                                .Update
                            Case "0", ""
                                !USR_INIT_IND = "N"
                                .Update
                            Case "Y", "N"
                                '
                            Case Else
                                AppendBusinessError radrsBRV _
                                    , trErrbusAbsenceContact _
                                    , mstrCBUSINESSERR _
                                    , "Invalid, Did you initiate contact Indicator." _
                                    , MODULENAME & ": ValidateAbsenceContact" _
                                    , , , , _
                                    , CStr(!ABS_NUM)
                        End Select
                    End If

                    If Not (IsNull(!TALK_TO_EMP_IND)) Then
                        Select Case (Trim(!TALK_TO_EMP_IND))
                            Case "1"
                                !TALK_TO_EMP_IND = "Y"
                                .Update
                            Case "0", ""
                                !TALK_TO_EMP_IND = "N"
                                .Update
                            Case "Y", "N"
                                '
                            Case Else
                                AppendBusinessError radrsBRV _
                                    , trErrbusAbsenceContact _
                                    , mstrCBUSINESSERR _
                                    , "Invalid, Did you talk employee Indicator." _
                                    , MODULENAME & ": ValidateAbsenceContact" _
                                    , , , , _
                                    , CStr(!ABS_NUM)
                        End Select
                    End If
                    If Not (IsNull(!TALK_TO_EMP_TXT)) Then
                        If (Trim(!TALK_TO_EMP_TXT) <> "") Then
                            !TALK_TO_EMP_TXT = ParseString(!TALK_TO_EMP_TXT)
                            .Update

                            If Not (CheckForValidString(!TALK_TO_EMP_TXT, True, 255)) Then
                                AppendBusinessError radrsBRV _
                                    , trErrbusAbsenceContact _
                                    , mstrCBUSINESSERR _
                                    , "Invalid, Did you talk to employee Comments." _
                                    , MODULENAME & ": ValidateAbsenceContact" _
                                    , , , , _
                                    , CStr(!ABS_NUM)
                            End If
                        End If
                    End If

                    If Not (IsNull(!EMP_TO_DR_IND)) Then
                        Select Case (Trim(!EMP_TO_DR_IND))
                            Case "1"
                                !EMP_TO_DR_IND = "Y"
                                .Update
                            Case "0", ""
                                !EMP_TO_DR_IND = "N"
                                .Update
                            Case "Y", "N"
                                '
                            Case Else
                                AppendBusinessError radrsBRV _
                                    , trErrbusAbsenceContact _
                                    , mstrCBUSINESSERR _
                                    , "Invalid, Has the employee been to a Doctor Indicator." _
                                    , MODULENAME & ": ValidateAbsenceContact" _
                                    , , , , _
                                    , CStr(!ABS_NUM)
                        End Select
                    End If
                    If Not (IsNull(!EMP_TO_DR_TXT)) Then
                        If (Trim(!EMP_TO_DR_TXT) <> "") Then
                            !EMP_TO_DR_TXT = ParseString(!EMP_TO_DR_TXT)
                            .Update

                            If Not (CheckForValidString(!EMP_TO_DR_TXT, True, 255)) Then
                                AppendBusinessError radrsBRV _
                                    , trErrbusAbsenceContact _
                                    , mstrCBUSINESSERR _
                                    , "Invalid, Has the employee been to a Doctor Comments." _
                                    , MODULENAME & ": ValidateAbsenceContact" _
                                    , , , , _
                                    , CStr(!ABS_NUM)
                            End If
                        End If
                    End If

                    If Not (IsNull(!OTH_TRTMNT_IND)) Then
                        Select Case (Trim(!OTH_TRTMNT_IND))
                            Case "1"
                                !OTH_TRTMNT_IND = "Y"
                                .Update
                            Case "0", ""
                                !OTH_TRTMNT_IND = "N"
                                .Update
                            Case "Y", "N"
                                '
                            Case Else
                                AppendBusinessError radrsBRV _
                                    , trErrbusAbsenceContact _
                                    , mstrCBUSINESSERR _
                                    , "Invalid, Has the employee had other medical treatment Indicator." _
                                    , MODULENAME & ": ValidateAbsenceContact" _
                                    , , , , _
                                    , CStr(!ABS_NUM)
                        End Select
                    End If
                    If Not (IsNull(!OTH_TRTMNT_TXT)) Then
                        If (Trim(!OTH_TRTMNT_TXT) <> "") Then
                            !OTH_TRTMNT_TXT = ParseString(!OTH_TRTMNT_TXT)
                            .Update

                            If Not (CheckForValidString(!OTH_TRTMNT_TXT, True, 255)) Then
                                AppendBusinessError radrsBRV _
                                    , trErrbusAbsenceContact _
                                    , mstrCBUSINESSERR _
                                    , "Invalid, Has the employee had other medical treatment Comments." _
                                    , MODULENAME & ": ValidateAbsenceContact" _
                                    , , , , _
                                    , CStr(!ABS_NUM)
                            End If
                        End If
                    End If

                    If Not (IsNull(!ADVS_SL_IND)) Then
                        Select Case (Trim(!ADVS_SL_IND))
                            Case "1"
                                !ADVS_SL_IND = "Y"
                                .Update
                            Case "0", ""
                                !ADVS_SL_IND = "N"
                                .Update
                            Case "Y", "N"
                                '
                            Case Else
                                AppendBusinessError radrsBRV _
                                    , trErrbusAbsenceContact _
                                    , mstrCBUSINESSERR _
                                    , "Invalid, Did you advise employee of Sick Leave policy Indicator." _
                                    , MODULENAME & ": ValidateAbsenceContact" _
                                    , , , , _
                                    , CStr(!ABS_NUM)
                        End Select
                    End If
                    If Not (IsNull(!ADVS_SL_TXT)) Then
                        If (Trim(!ADVS_SL_TXT) <> "") Then
                            !ADVS_SL_TXT = ParseString(!ADVS_SL_TXT)
                            .Update

                            If Not (CheckForValidString(!ADVS_SL_TXT, True, 255)) Then
                                AppendBusinessError radrsBRV _
                                    , trErrbusAbsenceContact _
                                    , mstrCBUSINESSERR _
                                    , "Invalid, Did you advise employee of Sick Leave policy Comments." _
                                    , MODULENAME & ": ValidateAbsenceContact" _
                                    , , , , _
                                    , CStr(!ABS_NUM)
                            End If
                        End If
                    End If

                    If Not (IsNull(!ADDL_TXT)) Then
                        If (Trim(!ADDL_TXT) <> "") Then
                            !ADDL_TXT = ParseString(!ADDL_TXT)
                            .Update

                            If Not (CheckForValidString(!ADDL_TXT, True, 255)) Then
                                AppendBusinessError radrsBRV _
                                    , trErrbusAbsenceContact _
                                    , mstrCBUSINESSERR _
                                    , "Invalid Comments." _
                                    , MODULENAME & ": ValidateAbsenceContact" _
                                    , , , , _
                                    , CStr(!ABS_NUM)
                            End If
                        Else
                            AppendBusinessError radrsBRV _
                                , trErrbusAbsenceContact _
                                , mstrCBUSINESSERR _
                                , "Invalid Comments. Comments are required." _
                                , MODULENAME & ": ValidateAbsenceContact" _
                                , , , , _
                                , CStr(!ABS_NUM)
                        End If
                    Else
                        AppendBusinessError radrsBRV _
                            , trErrbusAbsenceContact _
                            , mstrCBUSINESSERR _
                            , "Invalid Comments. Comments are required." _
                            , MODULENAME & ": ValidateAbsenceContact" _
                            , , , , _
                            , CStr(!ABS_NUM)
                    End If

                    If Not (IsNull(!PHONE_NUM)) Then
                        If (Trim(!PHONE_NUM) <> "") Then
                            If Not (CheckForValidString(!PHONE_NUM, , , True, 10)) Then
                                AppendBusinessError radrsBRV _
                                    , trErrbusAbsenceContact _
                                    , mstrCBUSINESSERR _
                                    , "Invalid Phone Number." _
                                    , MODULENAME & ": ValidateAbsenceContact" _
                                    , , , , _
                                    , CStr(!ABS_NUM)
                            End If
                        End If
                    Else
                        !PHONE_NUM = ""
                        .Update
                    End If

                    If Not (IsNull(!PRVDR_PHN_NUM)) Then
                        If (Trim(!PRVDR_PHN_NUM) <> "") Then
                            If Not (CheckForValidString(!PRVDR_PHN_NUM, , , True, 10)) Then
                                AppendBusinessError radrsBRV _
                                    , trErrbusAbsenceContact _
                                    , mstrCBUSINESSERR _
                                    , "Invalid Provider Phone Number." _
                                    , MODULENAME & ": ValidateAbsenceContact" _
                                    , , , , _
                                    , CStr(!ABS_NUM)
                            End If
                        End If
                    Else
                        !PRVDR_PHN_NUM = ""
                        .Update
                    End If

                    If Not (CheckForValidLong(!RPTNG_EMP_NUM)) Then
                        If (!RPTNG_EMP_NUM = 0) Then
                            !RPTNG_EMP_NUM = vlTRISUserNumber
                            .Update
                        Else
                            AppendBusinessError radrsBRV _
                                , trErrbusAbsenceContact _
                                , mstrCBUSINESSERR _
                                , "Invalid Contact Reported by Employee." _
                                , MODULENAME & ": ValidateAbsenceContact" _
                                , , , , _
                                , CStr(!ABS_NUM)
                        End If
                    Else
                        If (!RPTNG_EMP_NUM = 0) Then
                            !RPTNG_EMP_NUM = vlTRISUserNumber
                            .Update
                        Else
                            If Not (CheckForExistence(trEMPLOYEE, , !RPTNG_EMP_NUM)) Then
                                AppendBusinessError radrsBRV _
                                    , trErrbusAbsenceContact _
                                    , mstrCBUSINESSERR _
                                    , "Invalid Contact Reported by Employee." _
                                    , MODULENAME & ": ValidateAbsenceContact" _
                                    , , , , _
                                    , CStr(!ABS_NUM)
                            End If
                        End If
                    End If
            End Select

            Select Case (Trim(!RECORD_STATUS))
                Case "C", "D"
                    If Not (CheckForValidLong(!ABS_NUM, True)) Then
                        AppendBusinessError radrsBRV _
                            , trErrbusAbsenceContact _
                            , mstrCBUSINESSERR _
                            , "Invalid Absence Number." _
                            , MODULENAME & ": ValidateAbsenceContact" _
                            , , , , _
                            , CStr(!ABS_NUM)
                    End If

                    If Not (CheckForValidLong(!ABS_CONT_NUM, True)) Then
                        AppendBusinessError radrsBRV _
                            , trErrbusAbsenceContact _
                            , mstrCBUSINESSERR _
                            , "Invalid Absence Contact Number." _
                            , MODULENAME & ": ValidateAbsenceContact" _
                            , , , , _
                            , CStr(!ABS_NUM)
                    End If

                    If Not (IsNull(!CREATED_TIMESTAMP)) Then
                        If (Trim(!CREATED_TIMESTAMP) <> "") Then
                            If Not (CheckForValidString(!CREATED_TIMESTAMP, , , True, 26)) Then
                                AppendBusinessError radrsBRV _
                                    , trErrbusAbsenceContact _
                                    , mstrCBUSINESSERR _
                                    , "Invalid Time Stamp." _
                                    , MODULENAME & ": ValidateAbsenceContact" _
                                    , , , , _
                                    , CStr(!ABS_NUM)
                            End If
                        Else
                            AppendBusinessError radrsBRV _
                                , trErrbusAbsenceContact _
                                , mstrCBUSINESSERR _
                                , "Invalid Time Stamp." _
                                , MODULENAME & ": ValidateAbsenceContact" _
                                , , , , _
                                , CStr(!ABS_NUM)
                        End If
                    Else
                        AppendBusinessError radrsBRV _
                            , trErrbusAbsenceContact _
                            , mstrCBUSINESSERR _
                            , "Invalid Time Stamp." _
                            , MODULENAME & ": ValidateAbsenceContact" _
                            , , , , _
                            , CStr(!ABS_NUM)
                    End If
            End Select

            .MoveNext
        Next llIndex

        .MoveFirst
    End With

    If (RecordsExist(radrsBRV)) Then
        radrsBRV.Filter = adFilterNone
        radrsBRV.Filter = "ERR_TYPE = 'BE' OR ERR_TYPE = 'AE'" & _
                " OR ERR_TYPE = 'AW'"
        If Not (radrsBRV.EOF) Then
            radrsBRV.Filter = adFilterNone
            Exit Function
        End If
        radrsBRV.Filter = adFilterNone
    End If

    ValidateAbsenceContact = True

    Exit Function
ValidateAbsenceContact_EH:
    TrsRaiseError MODULENAME _
        , "ValidateAbsenceContact"

End Function
Private Function ValidateAbsenceReturnToWork(ByRef radrsAbsenceReturnToWork As ADODB.Recordset _
                , ByRef radrsBRV As ADODB.Recordset _
                ) As Boolean
On Error GoTo ValidateAbsenceReturnToWork_EH
Dim llIndex As Long

    ValidateAbsenceReturnToWork = False

    With radrsAbsenceReturnToWork
        .MoveFirst

        For llIndex = 1 To .RecordCount
            If Not (CheckForValidLong(!ABS_NUM)) Then
                AppendBusinessError radrsBRV _
                    , trErrbusAbsenceReturnToWork _
                    , mstrCBUSINESSERR _
                    , "Invalid Absence Number." _
                    , MODULENAME & ": ValidateAbsenceReturnToWork" _
                    , , , , _
                    , CStr(!ABS_NUM)
            Else
                If (!ABS_NUM > 0) Then
                    If Not (CheckForExistence(trABSENCE, , !ABS_NUM)) Then
                        AppendBusinessError radrsBRV _
                            , trErrbusAbsenceContact _
                            , mstrCBUSINESSERR _
                            , "Invalid Absence Number. Absence Number does not exist." _
                            , MODULENAME & ": ValidateAbsenceReturnToWork" _
                            , , , , _
                            , CStr(!ABS_NUM)
                    End If
                End If
            End If

            If Not (IsNull(!RTN_TO_WRK_DT)) Then
                If Not (CheckForValidDate(!RTN_TO_WRK_DT)) Then
                    AppendBusinessError radrsBRV _
                        , trErrbusAbsenceReturnToWork _
                        , mstrCBUSINESSERR _
                        , "Invalid Return to Work Date." _
                        , MODULENAME & ": ValidateAbsenceReturnToWork" _
                        , , , , _
                        , CStr(!ABS_NUM)
                End If
            Else
                AppendBusinessError radrsBRV _
                    , trErrbusAbsenceReturnToWork _
                    , mstrCBUSINESSERR _
                    , "Invalid Return to Work Date." _
                    , MODULENAME & ": ValidateAbsenceReturnToWork" _
                    , , , , _
                    , CStr(!ABS_NUM)
            End If

            If Not (IsNull(!RTN_TO_WRK_TM)) Then
                If (Trim(!RTN_TO_WRK_TM) <> "") Then
                    If Not (CheckForValidString(!RTN_TO_WRK_TM, , , True, 5)) Then
                        AppendBusinessError radrsBRV _
                            , trErrbusAbsenceReturnToWork _
                            , mstrCBUSINESSERR _
                            , "Invalid Return to Work Time." _
                            , MODULENAME & ": ValidateAbsenceReturnToWork" _
                            , , , , _
                            , CStr(!ABS_NUM)
                    End If
                Else
                    AppendBusinessError radrsBRV _
                        , trErrbusAbsenceReturnToWork _
                        , mstrCBUSINESSERR _
                        , "Invalid Return to Work Time." _
                        , MODULENAME & ": ValidateAbsenceReturnToWork" _
                        , , , , _
                        , CStr(!ABS_NUM)
                End If
            Else
                AppendBusinessError radrsBRV _
                    , trErrbusAbsenceReturnToWork _
                    , mstrCBUSINESSERR _
                    , "Invalid Contact Time." _
                    , MODULENAME & ": ValidateAbsenceReturnToWork" _
                    , , , , _
                    , CStr(!ABS_NUM)
            End If

            If Not (IsNull(!CONFRM_RSN_IND)) Then
                Select Case (Trim(!CONFRM_RSN_IND))
                    Case "1"
                        !CONFRM_RSN_IND = "Y"
                        .Update
                    Case "0", ""
                        !CONFRM_RSN_IND = "N"
                        .Update
                    Case "Y", "N"
                        '
                    Case Else
                        AppendBusinessError radrsBRV _
                            , trErrbusAbsenceReturnToWork _
                            , mstrCBUSINESSERR _
                            , "Invalid, Did you meet with the employee upon his/her return to work Indicator." _
                            , MODULENAME & ": ValidateAbsenceReturnToWork" _
                            , , , , _
                            , CStr(!ABS_NUM)
                End Select
            End If
            If Not (IsNull(!ABS_REASON_TXT)) Then
                If (Trim(!ABS_REASON_TXT) <> "") Then
                    !ABS_REASON_TXT = ParseString(!ABS_REASON_TXT)
                    .Update

                    If Not (CheckForValidString(!ABS_REASON_TXT, True, 255)) Then
                        AppendBusinessError radrsBRV _
                            , trErrbusAbsenceReturnToWork _
                            , mstrCBUSINESSERR _
                            , "Invalid Confirm Absence Reason." _
                            , MODULENAME & ": ValidateAbsenceReturnToWork" _
                            , , , , _
                            , CStr(!ABS_NUM)
                    End If
                End If
            End If

            If Not (IsNull(!MEET_EMP_IND)) Then
                Select Case (Trim(!MEET_EMP_IND))
                    Case "1"
                        !MEET_EMP_IND = "Y"
                        .Update
                    Case "0", ""
                        !MEET_EMP_IND = "N"
                        .Update
                    Case "Y", "N"
                        '
                    Case Else
                        AppendBusinessError radrsBRV _
                            , trErrbusAbsenceReturnToWork _
                            , mstrCBUSINESSERR _
                            , "Invalid, Did you meet with the employee upon his/her return to work Indicator." _
                            , MODULENAME & ": ValidateAbsenceReturnToWork" _
                            , , , , _
                            , CStr(!ABS_NUM)
                End Select
            End If
            If Not (IsNull(!MEET_EMP_TXT)) Then
                If (Trim(!MEET_EMP_TXT) <> "") Then
                    !MEET_EMP_TXT = ParseString(!MEET_EMP_TXT)
                    .Update

                    If Not (CheckForValidString(!MEET_EMP_TXT, True, 255)) Then
                        AppendBusinessError radrsBRV _
                            , trErrbusAbsenceReturnToWork _
                            , mstrCBUSINESSERR _
                            , "Invalid, Did you meet with the employee upon his/her return to work Comments." _
                            , MODULENAME & ": ValidateAbsenceReturnToWork" _
                            , , , , _
                            , CStr(!ABS_NUM)
                    End If
                End If
            End If

            If Not (IsNull(!PRFM_DUTY_IND)) Then
                Select Case (Trim(!PRFM_DUTY_IND))
                    Case "1"
                        !PRFM_DUTY_IND = "Y"
                        .Update
                    Case "0", ""
                        !PRFM_DUTY_IND = "N"
                        .Update
                    Case "Y", "N"
                        '
                    Case Else
                        AppendBusinessError radrsBRV _
                            , trErrbusAbsenceReturnToWork _
                            , mstrCBUSINESSERR _
                            , "Invalid, Has the employee been released to work with restrictions Indicator." _
                            , MODULENAME & ": ValidateAbsenceReturnToWork" _
                            , , , , _
                            , CStr(!ABS_NUM)
                End Select
            End If
            If Not (IsNull(!PRFM_DUTY_TXT)) Then
                If (Trim(!PRFM_DUTY_TXT) <> "") Then
                    !PRFM_DUTY_TXT = ParseString(!PRFM_DUTY_TXT)
                    .Update

                    If Not (CheckForValidString(!PRFM_DUTY_TXT, True, 255)) Then
                        AppendBusinessError radrsBRV _
                            , trErrbusAbsenceReturnToWork _
                            , mstrCBUSINESSERR _
                            , "Invalid, Has the employee been released to work with restrictions Comments." _
                            , MODULENAME & ": ValidateAbsenceReturnToWork" _
                            , , , , _
                            , CStr(!ABS_NUM)
                    End If
                End If
            End If

            If Not (IsNull(!SIDE_EFF_IND)) Then
                Select Case (Trim(!SIDE_EFF_IND))
                    Case "1"
                        !SIDE_EFF_IND = "Y"
                        .Update
                    Case "0", ""
                        !SIDE_EFF_IND = "N"
                        .Update
                    Case "Y", "N"
                        '
                    Case Else
                        AppendBusinessError radrsBRV _
                            , trErrbusAbsenceReturnToWork _
                            , mstrCBUSINESSERR _
                            , "Invalid, Does the Employee have any side effects from medication that would affect the performance of his/her job Indicator." _
                            , MODULENAME & ": ValidateAbsenceReturnToWork" _
                            , , , , _
                            , CStr(!ABS_NUM)
                End Select
            End If
            If Not (IsNull(!SIDE_EFF_TXT)) Then
                If (Trim(!SIDE_EFF_TXT) <> "") Then
                    !SIDE_EFF_TXT = ParseString(!SIDE_EFF_TXT)
                    .Update

                    If Not (CheckForValidString(!SIDE_EFF_TXT, True, 255)) Then
                        AppendBusinessError radrsBRV _
                            , trErrbusAbsenceReturnToWork _
                            , mstrCBUSINESSERR _
                            , "Invalid, Does the Employee have any side effects from medication that would affect the performance of his/her job Comments." _
                            , MODULENAME & ": ValidateAbsenceReturnToWork" _
                            , , , , _
                            , CStr(!ABS_NUM)
                    End If
                End If
            End If

            If Not (IsNull(!ADDL_TXT)) Then
                If (Trim(!ADDL_TXT) <> "") Then
                    !ADDL_TXT = ParseString(!ADDL_TXT)
                    .Update

                    If Not (CheckForValidString(!ADDL_TXT, True, 255)) Then
                        AppendBusinessError radrsBRV _
                            , trErrbusAbsenceReturnToWork _
                            , mstrCBUSINESSERR _
                            , "Invalid Return to Work Comments." _
                            , MODULENAME & ": ValidateAbsenceReturnToWork" _
                            , , , , _
                            , CStr(!ABS_NUM)
                    End If
                End If
            End If

            If Not (CheckForValidLong(!RPTNG_EMP_NUM)) Then
                AppendBusinessError radrsBRV _
                    , trErrbusAbsenceReturnToWork _
                    , mstrCBUSINESSERR _
                    , "Invalid Return to Work Reported by Employee." _
                    , MODULENAME & ": ValidateAbsenceReturnToWork" _
                    , , , , _
                    , CStr(!ABS_NUM)
            Else
                If Not (CheckForExistence(trEMPLOYEE, , !RPTNG_EMP_NUM)) Then
                    AppendBusinessError radrsBRV _
                        , trErrbusAbsenceReturnToWork _
                        , mstrCBUSINESSERR _
                        , "Invalid Return to Work Reported by Employee." _
                        , MODULENAME & ": ValidateAbsenceReturnToWork" _
                        , , , , _
                        , CStr(!ABS_NUM)
                End If
            End If

            If Not (IsNull(!RPT_DT)) Then
                If Not (CheckForValidDate(!RPT_DT)) Then
                    AppendBusinessError radrsBRV _
                        , trErrbusAbsenceReturnToWork _
                        , mstrCBUSINESSERR _
                        , "Invalid Return to Work Reported Date." _
                        , MODULENAME & ": ValidateAbsenceReturnToWork" _
                        , , , , _
                        , CStr(!ABS_NUM)
                End If
            Else
                AppendBusinessError radrsBRV _
                    , trErrbusAbsenceReturnToWork _
                    , mstrCBUSINESSERR _
                    , "Invalid Return to Work Reported Date." _
                    , MODULENAME & ": ValidateAbsenceReturnToWork" _
                    , , , , _
                    , CStr(!ABS_NUM)
            End If

            Select Case (Trim(!RECORD_STATUS))
                Case "C", "D"
                    If Not (CheckForValidLong(!ABS_NUM, True)) Then
                        AppendBusinessError radrsBRV _
                            , trErrbusAbsenceReturnToWork _
                            , mstrCBUSINESSERR _
                            , "Invalid Absence Number." _
                            , MODULENAME & ": ValidateAbsenceReturnToWork" _
                            , , , , _
                            , CStr(!ABS_NUM)
                    End If

                    If Not (IsNull(!CREATED_TIMESTAMP)) Then
                        If (Trim(!CREATED_TIMESTAMP) <> "") Then
                            If Not (CheckForValidString(!CREATED_TIMESTAMP, , , True, 26)) Then
                                AppendBusinessError radrsBRV _
                                    , trErrbusAbsenceReturnToWork _
                                    , mstrCBUSINESSERR _
                                    , "Invalid Time Stamp." _
                                    , MODULENAME & ": ValidateAbsenceReturnToWork" _
                                    , , , , _
                                    , CStr(!ABS_NUM)
                            End If
                        Else
                            AppendBusinessError radrsBRV _
                                , trErrbusAbsenceReturnToWork _
                                , mstrCBUSINESSERR _
                                , "Invalid Time Stamp." _
                                , MODULENAME & ": ValidateAbsenceReturnToWork" _
                                , , , , _
                                , CStr(!ABS_NUM)
                        End If
                    Else
                        AppendBusinessError radrsBRV _
                            , trErrbusAbsenceReturnToWork _
                            , mstrCBUSINESSERR _
                            , "Invalid Time Stamp." _
                            , MODULENAME & ": ValidateAbsenceReturnToWork" _
                            , , , , _
                            , CStr(!ABS_NUM)
                    End If
            End Select

            .MoveNext
        Next llIndex

        .MoveFirst
    End With

    If (RecordsExist(radrsBRV)) Then
        radrsBRV.Filter = adFilterNone
        radrsBRV.Filter = "ERR_TYPE = 'BE' OR ERR_TYPE = 'AE'" & _
                " OR ERR_TYPE = 'AW'"
        If Not (radrsBRV.EOF) Then
            radrsBRV.Filter = adFilterNone
            Exit Function
        End If
        radrsBRV.Filter = adFilterNone
    End If

    ValidateAbsenceReturnToWork = True

    Exit Function
ValidateAbsenceReturnToWork_EH:
    TrsRaiseError MODULENAME _
        , "ValidateAbsenceReturnToWork"

End Function
Private Function AddPortfolioInformation(ByVal viType As Integer _
                    , ByVal vadrsInitialAbsence As ADODB.Recordset _
                    ) As ADODB.Recordset
On Error GoTo AddPortfolioInformation_EH
Dim ladrs As ADODB.Recordset, llIndex As Long

    Set AddPortfolioInformation = Nothing

    Set ladrs = CreatePortfolioRS(viType)

    If (RecordsExist(vadrsInitialAbsence)) Then
        With vadrsInitialAbsence
            If (.RecordCount > 1) Then
                .MoveFirst

                Select Case (viType)
                    Case 1
                        ladrs.AddNew
                        ladrs!PRTFL_NUM = -1
                        ladrs!PRTFL_DESC = "SYSTEM DEFINED PORTFOLIO"
                        ladrs!SHARE_IND = "N"
                        ladrs!RECORD_STATUS = "I"
                        ladrs.Update

                    Case 2
                        For llIndex = 1 To .RecordCount
                            ladrs.AddNew
                            ladrs!PRTFL_NUM = -1
                            ladrs!ABS_NUM = !ABS_NUM
                            ladrs!RECORD_STATUS = "I"
                            ladrs.Update

                            .MoveNext
                        Next llIndex
                End Select

                .MoveFirst
            End If
        End With
    End If

    Set AddPortfolioInformation = ladrs
    Set ladrs = Nothing

    Exit Function
AddPortfolioInformation_EH:
    Set ladrs = Nothing

    TrsRaiseError MODULENAME _
        , "AddPortfolioInformation"

End Function
Private Function CreatePortfolioRS(ByVal viType As Integer _
                    ) As ADODB.Recordset
On Error GoTo CreatePortfolioRS_EH
Dim ladrs As ADODB.Recordset

    Set ladrs = New ADODB.Recordset
    With ladrs
        .Fields.Append "PRTFL_NUM", adInteger, , ADOFIELDATTR

        Select Case (viType)
            Case 1
                .Fields.Append "PRTFL_DESC", adVarChar, 30, ADOFIELDATTR
                .Fields.Append "SHARE_IND", adChar, 1, ADOFIELDATTR

            Case 2
                .Fields.Append "ABS_NUM", adInteger, , ADOFIELDATTR
        End Select

        .Fields.Append "RECORD_STATUS", adChar, 1, ADOFIELDATTR

        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open
    End With

    Set CreatePortfolioRS = ladrs
    Set ladrs = Nothing

    Exit Function
CreatePortfolioRS_EH:
    Set ladrs = Nothing

    TrsRaiseError MODULENAME _
        , "CreatePortfolioRS"

End Function
Public Function ProcessPortfolioRequest(ByVal vadrsPortfolio As ADODB.Recordset _
                , ByVal vadrsAbsences As ADODB.Recordset _
                , ByVal vlTRISUserNumber As Long _
                , ByVal vsTRISUserID As String _
                , ByRef rlPortfolioNumber As Long _
                , ByRef radrsBRV As ADODB.Recordset _
                ) As Boolean
On Error GoTo ProcessPortfolioRequest_EH
Dim lobjdbEAM As dbEAM.cdEAMTrans
Dim llOriginalPortfolio As Long

    ProcessPortfolioRequest = False

    llOriginalPortfolio = rlPortfolioNumber

    Set radrsBRV = CreateNewBusinessErrorSet

    If Not (RecordsExist(vadrsPortfolio) Or _
        RecordsExist(vadrsAbsences)) Then
        AppendBusinessError radrsBRV _
            , trErrbusAbsencePortfolio _
            , mstrCBUSINESSERR _
            , "There are no portfolio absence entries to process." _
            , MODULENAME & ": ProcessPortfolioRequest"

        GoTo REPORTBRV
    End If

    If Not (CheckForValidLong(vlTRISUserNumber, True)) Then
        AppendBusinessError radrsBRV _
            , trErrbusEAM _
            , mstrCBUSINESSERR _
            , "Invalid TRIS User." _
            , MODULENAME & ": ProcessPortfolioRequest"

        GoTo REPORTBRV
    End If

    If (RecordsExist(vadrsPortfolio)) Then
        If Not (ValidatePortfolio(vadrsPortfolio _
                , radrsBRV)) Then
            GoTo REPORTBRV
        End If
    End If

    If (RecordsExist(vadrsAbsences)) Then
        If Not (ValidatePortfolioAbsence(vadrsAbsences _
                , radrsBRV)) Then
            GoTo REPORTBRV
        End If
    End If

    Set lobjdbEAM = CreateObject("dbEAM.cdEAMTrans")

    If Not (lobjdbEAM.dbProcessPorfolioRequest(vadrsPortfolio _
        , vadrsAbsences _
        , vlTRISUserNumber _
        , vsTRISUserID _
        , rlPortfolioNumber _
        , radrsBRV)) Then
        Set lobjdbEAM = Nothing

        GoTo REPORTBRV
    End If

    Set lobjdbEAM = Nothing

    ProcessPortfolioRequest = True

    GetObjectContext.SetComplete

    Exit Function
ProcessPortfolioRequest_EH:
    GetObjectContext.SetAbort

    rlPortfolioNumber = llOriginalPortfolio

    ProcessPortfolioRequest = False

    Set lobjdbEAM = Nothing

    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "ProcessPortfolioRequest"
            Case Else
                TrsRaiseError MODULENAME _
                    , "ProcessPortfolioRequest" _
                    , Err.Number _
                    , , , , _
                    True
    End Select

    Exit Function
REPORTBRV:
    Set lobjdbEAM = Nothing

    rlPortfolioNumber = llOriginalPortfolio

    If (RecordsExist(radrsBRV)) Then
        With radrsBRV
            .MoveFirst

            'Check for errors
            .Find "ERR_TYPE = 'BE'"
            'If found abort else complete
            If Not (.EOF) Then
                'Abort Transaction
                GetObjectContext().SetAbort
                ProcessPortfolioRequest = False
            Else
                'Complete Transaction
                GetObjectContext().SetComplete
                ProcessPortfolioRequest = True
            End If

            .MoveFirst
        End With
    Else
        'Complete Transaction
        GetObjectContext().SetComplete
        ProcessPortfolioRequest = True
    End If

    Exit Function
End Function
Private Function ValidatePortfolio(ByVal vadrsPortfolio As ADODB.Recordset _
                , ByRef radrsBRV As ADODB.Recordset _
                ) As Boolean
On Error GoTo ValidatePortfolio_EH

    ValidatePortfolio = True

    With vadrsPortfolio
        .MoveFirst

        If Not (CheckForValidLong(!PRTFL_NUM)) Then
            AppendBusinessError radrsBRV _
                , trErrbusAbsencePortfolio _
                , mstrCBUSINESSERR _
                , "Invalid Portfolio Number." _
                , MODULENAME & ": ValidatePortfolio"
        End If

        If Not (IsNull(!PRTFL_DESC)) Then
            !PRTFL_DESC = ParseString(!PRTFL_DESC)
            .Update

            If Not (CheckForValidString(!PRTFL_DESC, True, 30)) Then
                AppendBusinessError radrsBRV _
                    , trErrbusAbsencePortfolio _
                    , mstrCBUSINESSERR _
                    , "Invalid Portfolio Description." _
                    , MODULENAME & ": ValidatePortfolio"
            End If
        Else
            AppendBusinessError radrsBRV _
                , trErrbusAbsencePortfolio _
                , mstrCBUSINESSERR _
                , "Invalid Portfolio Description." _
                , MODULENAME & ": ValidatePortfolio"
        End If

        If Not (IsNull(!SHARE_IND)) Then
            Select Case (Trim(!SHARE_IND))
                Case "Y", "1", "YES"
                    !SHARE_IND = "Y"
                    .Update

                Case "N", "0", "NO"
                    !SHARE_IND = "N"
                    .Update

                Case Else
                    AppendBusinessError radrsBRV _
                        , trErrbusAbsencePortfolio _
                        , mstrCBUSINESSERR _
                        , "Invalid Portfolio Share Indicator." _
                        , MODULENAME & ": ValidatePortfolio"
            End Select
        Else
            AppendBusinessError radrsBRV _
                , trErrbusAbsencePortfolio _
                , mstrCBUSINESSERR _
                , "Invalid Portfolio Share Indicator." _
                , MODULENAME & ": ValidatePortfolio"
        End If

        .MoveFirst
    End With

    If (RecordsExist(radrsBRV)) Then
        radrsBRV.Filter = adFilterNone

        radrsBRV.Filter = "ERR_TYPE = 'BE' OR ERR_TYPE = 'AE'" & _
                " OR ERR_TYPE = 'AW'"
        If Not (radrsBRV.EOF) Then
            radrsBRV.Filter = adFilterNone
            Exit Function
        End If

        radrsBRV.Filter = adFilterNone
    End If

    ValidatePortfolio = True

    Exit Function
ValidatePortfolio_EH:
    TrsRaiseError MODULENAME _
            , "ValidatePortfolio"

End Function
Private Function ValidatePortfolioAbsence(ByVal vadrsAbsences As ADODB.Recordset _
                , ByRef radrsBRV As ADODB.Recordset _
                ) As Boolean
On Error GoTo ValidatePortfolioAbsence_EH
Dim llIndex As Long

    ValidatePortfolioAbsence = False

    With vadrsAbsences
        .MoveFirst

        For llIndex = 1 To .RecordCount
            If Not (CheckForValidLong(!PRTFL_NUM)) Then
                AppendBusinessError radrsBRV _
                    , trErrbusAbsencePortfolio _
                    , mstrCBUSINESSERR _
                    , "Invalid Portfolio Number." _
                    , MODULENAME & ": ValidatePortfolioAbsence"
            End If

            If Not (CheckForValidLong(!ABS_NUM)) Then
                AppendBusinessError radrsBRV _
                    , trErrbusAbsencePortfolio _
                    , mstrCBUSINESSERR _
                    , "Invalid Absence Number." _
                    , MODULENAME & ": ValidatePortfolioAbsence"
            Else
                If (!ABS_NUM > 0) Then
                    If Not (CheckForExistence(trABSENCE, , !ABS_NUM)) Then
                        AppendBusinessError radrsBRV _
                            , trErrbusAbsencePortfolio _
                            , mstrCBUSINESSERR _
                            , "Invalid Absence. Absence " & CStr(!ABS_NUM) & " does not exist." _
                            , MODULENAME & ": ValidatePortfolioAbsence"
                    End If
                End If
            End If

            .MoveNext
        Next llIndex

        .MoveFirst
    End With

    If (RecordsExist(radrsBRV)) Then
        radrsBRV.Filter = adFilterNone

        radrsBRV.Filter = "ERR_TYPE = 'BE' OR ERR_TYPE = 'AE'" & _
                " OR ERR_TYPE = 'AW'"
        If Not (radrsBRV.EOF) Then
            radrsBRV.Filter = adFilterNone
            Exit Function
        End If

        radrsBRV.Filter = adFilterNone
    End If

    ValidatePortfolioAbsence = True

    Exit Function
ValidatePortfolioAbsence_EH:
    TrsRaiseError MODULENAME _
            , "ValidatePortfolioAbsence"

End Function
Private Function ValidateAbsenceCallLog(ByVal vadrsCallLog As ADODB.Recordset _
                , ByRef radrsBRV As ADODB.Recordset _
                ) As Boolean
On Error GoTo ValidateAbsenceCallLog_EH
Dim llIndex As Long

    ValidateAbsenceCallLog = False

    With vadrsCallLog
        .Filter = adFilterNone

        .Filter = ("RECORD_STATUS = 'I'")
        If Not (.EOF) Then
            .MoveFirst

            For llIndex = 1 To .RecordCount
                If Not (CheckForValidLong(!ABS_NUM)) Then
                    AppendBusinessError radrsBRV _
                        , trErrbusAbsencePortfolio _
                        , mstrCBUSINESSERR _
                        , "Invalid Absence Number." _
                        , MODULENAME & ": ValidateAbsenceCallLog"
                End If

                If Not (IsNull(!CALL_ST_TS)) Then
                    If Not (CheckForValidString(!CALL_ST_TS, False, , True, 26)) Then
                        AppendBusinessError radrsBRV _
                            , trErrbusAbsencePortfolio _
                            , mstrCBUSINESSERR _
                            , "Invalid Call Start Date/Time." _
                            , MODULENAME & ": ValidateAbsenceCallLog"
                    End If
                Else
                    AppendBusinessError radrsBRV _
                        , trErrbusAbsencePortfolio _
                        , mstrCBUSINESSERR _
                        , "Invalid Call Start Date/Time." _
                        , MODULENAME & ": ValidateAbsenceCallLog"
                End If

                If Not (IsNull(!CALL_END_TS)) Then
                    If Not (CheckForValidString(!CALL_END_TS, False, , True, 26)) Then
                        AppendBusinessError radrsBRV _
                            , trErrbusAbsencePortfolio _
                            , mstrCBUSINESSERR _
                            , "Invalid Call End Date/Time." _
                            , MODULENAME & ": ValidateAbsenceCallLog"
                    End If
                Else
                    AppendBusinessError radrsBRV _
                        , trErrbusAbsencePortfolio _
                        , mstrCBUSINESSERR _
                        , "Invalid Call End Date/Time." _
                        , MODULENAME & ": ValidateAbsenceCallLog"
                End If

                If Not (IsNull(!PHONE_NUM)) Then
'Begin Prem PTS 2095
                    If Not (CheckForValidString(!PHONE_NUM, True, 10)) Then
'End Prem PTS 2095
                        AppendBusinessError radrsBRV _
                            , trErrbusAbsencePortfolio _
                            , mstrCBUSINESSERR _
                            , "Invalid Phone Number." _
                            , MODULENAME & ": ValidateAbsenceCallLog"
                    End If
                Else
                    AppendBusinessError radrsBRV _
                        , trErrbusAbsencePortfolio _
                        , mstrCBUSINESSERR _
                        , "Invalid Phone Number." _
                        , MODULENAME & ": ValidateAbsenceCallLog"
                End If

                .MoveNext
            Next llIndex
        End If

        .Filter = adFilterNone
        .MoveFirst
    End With

    If (RecordsExist(radrsBRV)) Then
        radrsBRV.Filter = adFilterNone

        radrsBRV.Filter = "ERR_TYPE = 'BE' OR ERR_TYPE = 'AE'" & _
                " OR ERR_TYPE = 'AW'"
        If Not (radrsBRV.EOF) Then
            radrsBRV.Filter = adFilterNone
            Exit Function
        End If

        radrsBRV.Filter = adFilterNone
    End If

    ValidateAbsenceCallLog = True

    Exit Function
ValidateAbsenceCallLog_EH:
    TrsRaiseError MODULENAME _
            , "ValidateAbsenceCallLog"

End Function
Private Sub AddBRVRecords(ByVal vadrsSource As ADODB.Recordset _
                , ByRef radrsDestination As ADODB.Recordset _
                , Optional ByVal vbActualBRV As Boolean = False)
On Error GoTo AddBRVRecords_EH
Dim llIndex As Long, llFldIndex As Long

    If (RecordsExist(vadrsSource)) Then
        With vadrsSource
            .MoveFirst

            For llIndex = 1 To .RecordCount
                radrsDestination.AddNew

                If Not (vbActualBRV) Then
                    For llFldIndex = 0 To .Fields.Count - 1
                        radrsDestination.Fields.Item(llFldIndex).Value _
                            = .Fields.Item(llFldIndex).Value
                    Next llFldIndex
                Else
                    radrsDestination!ERR_TYPE = !ERR_TYPE
                    radrsDestination!ERR_DESCRIPTION = !ERR_DESCRIPTION
                    radrsDestination!ERR_NUMBER = !ERR_NUMBER
                    radrsDestination!KEY_FIELD1 = CStr(!CHARGE_DATE)
                    radrsDestination!KEY_FIELD2 = CStr(!DATE_WORKED)
                    radrsDestination!KEY_FIELD3 = Trim(!START_TIME)
                    If (!END_TIME = "24:00") Then
                        radrsDestination!KEY_FIELD4 = "00:00"
                    Else
                        radrsDestination!KEY_FIELD4 = Trim(!END_TIME)
                    End If
                    radrsDestination!KEY_FIELD5 = !EXCEPTION_CD
                    radrsDestination!KEY_FIELD6 = !OVER_EXC_CD
                End If

                radrsDestination.Update

                .MoveNext
            Next llIndex

            .MoveFirst
        End With
    End If

    If (RecordsExist(radrsDestination)) Then
        radrsDestination.MoveFirst
    End If

    Exit Sub
AddBRVRecords_EH:
    TrsRaiseError MODULENAME _
        , "AddBRVRecords"

End Sub
Public Function ProcessContactMgmtRequest(ByVal vadrsContacts As ADODB.Recordset _
                , ByVal vlTRISUserNumber As Long _
                , ByVal vsTRISUserID As String _
                , ByRef radrsBRV As ADODB.Recordset _
                ) As Boolean
On Error GoTo ProcessContactMgmtRequest_EH
Dim lobjdbEAM As dbEAM.cdEAMTrans

    ProcessContactMgmtRequest = False

    Set radrsBRV = CreateNewBusinessErrorSet

    If Not (RecordsExist(vadrsContacts)) Then
        AppendBusinessError radrsBRV _
            , trErrbusAbsencePortfolio _
            , mstrCBUSINESSERR _
            , "There are no contact entries to process." _
            , MODULENAME & ": ProcessContactMgmtRequest"

        GoTo REPORTBRV
    End If

    If Not (CheckForValidLong(vlTRISUserNumber, True)) Then
        AppendBusinessError radrsBRV _
            , trErrbusEAM _
            , mstrCBUSINESSERR _
            , "Invalid TRIS User." _
            , MODULENAME & ": ProcessContactMgmtRequest"

        GoTo REPORTBRV
    End If

    If (RecordsExist(vadrsContacts)) Then
        If Not (ValidateAbsenceContact(vadrsContacts _
                , radrsBRV _
                , vlTRISUserNumber)) Then
            GoTo REPORTBRV
        End If
    End If

    Set lobjdbEAM = CreateObject("dbEAM.cdEAMTrans")

    If Not (lobjdbEAM.dbProcessContactRequest(vadrsContacts _
        , vlTRISUserNumber _
        , vsTRISUserID)) Then
        Set lobjdbEAM = Nothing

        GoTo REPORTBRV
    End If

    Set lobjdbEAM = Nothing

    ProcessContactMgmtRequest = True

    GetObjectContext.SetComplete

    Exit Function
ProcessContactMgmtRequest_EH:
    GetObjectContext.SetAbort

    ProcessContactMgmtRequest = False

    Set lobjdbEAM = Nothing

    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "ProcessContactMgmtRequest"
            Case Else
                TrsRaiseError MODULENAME _
                    , "ProcessContactMgmtRequest" _
                    , , Err.Number _
                    , , , , _
                    True
    End Select

    Exit Function
REPORTBRV:
    Set lobjdbEAM = Nothing

    If (RecordsExist(radrsBRV)) Then
        With radrsBRV
            .MoveFirst

            'Check for errors
            .Find "ERR_TYPE = 'BE'"
            'If found abort else complete
            If Not (.EOF) Then
                'Abort Transaction
                GetObjectContext().SetAbort
                ProcessContactMgmtRequest = False
            Else
                'Complete Transaction
                GetObjectContext().SetComplete
                ProcessContactMgmtRequest = True
            End If

            .MoveFirst
        End With
    Else
        'Complete Transaction
        GetObjectContext().SetComplete
        ProcessContactMgmtRequest = True
    End If

    Exit Function
End Function
Public Function ProcessRTWMgmtRequest(ByVal vadrsRTW As ADODB.Recordset _
                , ByVal vlTRISUserNumber As Long _
                , ByVal vsTRISUserID As String _
                , ByRef radrsBRV As ADODB.Recordset _
                ) As Boolean
On Error GoTo ProcessRTWMgmtRequest_EH
Dim lobjdbEAM As dbEAM.cdEAMTrans

    ProcessRTWMgmtRequest = False

    Set radrsBRV = CreateNewBusinessErrorSet

    If Not (RecordsExist(vadrsRTW)) Then
        AppendBusinessError radrsBRV _
            , trErrbusAbsencePortfolio _
            , mstrCBUSINESSERR _
            , "There are no return to work entries to process." _
            , MODULENAME & ": ProcessRTWMgmtRequest"

        GoTo REPORTBRV
    End If

    If Not (CheckForValidLong(vlTRISUserNumber, True)) Then
        AppendBusinessError radrsBRV _
            , trErrbusEAM _
            , mstrCBUSINESSERR _
            , "Invalid TRIS User." _
            , MODULENAME & ": ProcessRTWMgmtRequest"

        GoTo REPORTBRV
    End If

    If (RecordsExist(vadrsRTW)) Then
        If Not (ValidateAbsenceReturnToWork(vadrsRTW _
                , radrsBRV)) Then
            GoTo REPORTBRV
        End If
    End If

    Set lobjdbEAM = CreateObject("dbEAM.cdEAMTrans")

    If Not (lobjdbEAM.dbProcessRtnToWorkRequest(vadrsRTW _
        , vlTRISUserNumber _
        , vsTRISUserID)) Then
        Set lobjdbEAM = Nothing

        GoTo REPORTBRV
    End If

    Set lobjdbEAM = Nothing

    ProcessRTWMgmtRequest = True

    GetObjectContext.SetComplete

    Exit Function
ProcessRTWMgmtRequest_EH:
    GetObjectContext.SetAbort

    ProcessRTWMgmtRequest = False

    Set lobjdbEAM = Nothing

    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "ProcessRTWMgmtRequest"
            Case Else
                TrsRaiseError MODULENAME _
                    , "ProcessRTWMgmtRequest" _
                    , Err.Number _
                    , , , , _
                    True
    End Select

    Exit Function
REPORTBRV:
    Set lobjdbEAM = Nothing

    If (RecordsExist(radrsBRV)) Then
        With radrsBRV
            .MoveFirst

            'Check for errors
            .Find "ERR_TYPE = 'BE'"
            'If found abort else complete
            If Not (.EOF) Then
                'Abort Transaction
                GetObjectContext().SetAbort
                ProcessRTWMgmtRequest = False
            Else
                'Complete Transaction
                GetObjectContext().SetComplete
                ProcessRTWMgmtRequest = True
            End If

            .MoveFirst
        End With
    Else
        'Complete Transaction
        GetObjectContext().SetComplete
        ProcessRTWMgmtRequest = True
    End If

    Exit Function
End Function
Public Function CallStartLog(ByVal vlAbsenceNumber As Long _
                , ByVal vsPhoneNumber As String _
                , ByVal vlTRISUserNumber As String _
                , ByVal vsTRISUserID As String _
                ) As ADODB.Recordset
On Error GoTo CallStartLog_EH
Dim lobjdbEAM As dbEAM.cdEAMTrans

    Set CallStartLog = Nothing

    If (vlAbsenceNumber > 0) And _
        (Trim(vsPhoneNumber) <> "") Then
        Set lobjdbEAM = CreateObject("dbEAM.cdEAMTrans")

        Set CallStartLog = lobjdbEAM.dbCallStart(vlAbsenceNumber _
                , vsPhoneNumber _
                , vlTRISUserNumber)

        Set lobjdbEAM = Nothing
    End If

    GetObjectContext.SetComplete

    Exit Function
CallStartLog_EH:
    GetObjectContext.SetAbort

    Set lobjdbEAM = Nothing

    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "CallStartLog"
            Case Else
                TrsRaiseError MODULENAME _
                    , "CallStartLog" _
                    , Err.Number _
                    , , , , _
                    True
    End Select

End Function
Public Function CallEndLog(ByVal vadrsCallStartLog As ADODB.Recordset _
                ) As String
On Error GoTo CallEndLog_EH
Dim lobjdbEAM As dbEAM.cdEAMTrans

    CallEndLog = ""

    If (RecordsExist(vadrsCallStartLog)) Then
        Set lobjdbEAM = CreateObject("dbEAM.cdEAMTrans")

        CallEndLog = lobjdbEAM.dbCallEnd(vadrsCallStartLog)

        Set lobjdbEAM = Nothing
    End If

    GetObjectContext.SetComplete

    Exit Function
CallEndLog_EH:
    GetObjectContext.SetAbort

    Set lobjdbEAM = Nothing

    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "CallEndLog"
            Case Else
                TrsRaiseError MODULENAME _
                    , "CallEndLog" _
                    , Err.Number _
                    , , , , _
                    True
    End Select

End Function
Public Function ProcessDeleteAbsence(ByVal vadrsAbsence As ADODB.Recordset _
                , ByVal vsApprovalAuthority As String _
                , ByVal vlTRISUserNumber As Long _
                , ByVal vsTRISUserID As String _
                , ByRef radrsBRV As ADODB.Recordset _
                ) As Boolean
On Error GoTo ProcessDeleteAbsence_EH
'Components
Dim lobjCollPeriod As busCollectionPeriod.cbsCollectionPeriod
Dim lobjSchedule As busSchedule.cbsSchedule
Dim lobjdbEAMSave As dbEAM.cdEAMTrans
Dim lobjSchedExc As busSchedException.cbsSchedExc
Dim lobjSchedExcSave As busSchedException.cbsSchedExcTrans
Dim lobjbusTimeKeeper As busTimekeeper.cbsTimekeeper

'RS
Dim ladrsPayPeriod As ADODB.Recordset
Dim ladrsTimeDetails As ADODB.Recordset
Dim ladrsSchedExcDtls As ADODB.Recordset
Dim ladrsAbsence As ADODB.Recordset
Dim ladrsScheduleNoExc As ADODB.Recordset
Dim ladrsSchedule As ADODB.Recordset

'Save
Dim ladrsSchedExc As ADODB.Recordset
Dim ladrsBRVSchedExc As ADODB.Recordset
Dim ladrsActual As ADODB.Recordset
Dim ladrsBRVActual As ADODB.Recordset

'Other
Dim ldtLastProcessedDate As Date

    Set radrsBRV = CreateNewBusinessErrorSet

    If (RecordsExist(vadrsAbsence)) Then
        vadrsAbsence.MoveFirst

        If Not (CheckForValidLong(vadrsAbsence!ABS_NUM, True)) Then
            AppendBusinessError radrsBRV _
                , trErrbusEAM _
                , mstrCBUSINESSERR _
                , "Invalid Absence." _
                , MODULENAME & ": ProcessDeleteAbsence"

            GoTo REPORTBRV
        Else
            If Not (CheckForExistence(trABSENCE, , vadrsAbsence!ABS_NUM)) Then
                AppendBusinessError radrsBRV _
                    , trErrbusEAM _
                    , mstrCBUSINESSERR _
                    , "Invalid Absence. Absence " & CStr(vadrsAbsence!ABS_NUM) & " does not exist." _
                    , MODULENAME & ": ProcessDeleteAbsence"

                GoTo REPORTBRV
            End If
        End If

        If Not (CheckForValidLong(vadrsAbsence!EMP_NUM, True)) Then
            AppendBusinessError radrsBRV _
                , trErrbusEAM _
                , mstrCBUSINESSERR _
                , "Invalid Employee." _
                , MODULENAME & ": ProcessDeleteAbsence"

            GoTo REPORTBRV
        Else
            If Not (CheckForExistence(trEMPLOYEE, , vadrsAbsence!EMP_NUM)) Then
                AppendBusinessError radrsBRV _
                    , trErrbusEAM _
                    , mstrCBUSINESSERR _
                    , "Invalid Employee. Employee " & CStr(vadrsAbsence!EMP_NUM) & " does not exist." _
                    , MODULENAME & ": ProcessDeleteAbsence"

                GoTo REPORTBRV
            End If
        End If
    Else
        AppendBusinessError radrsBRV _
            , trErrbusEAM _
            , mstrCBUSINESSERR _
            , "There are not entries to process." _
            , MODULENAME & ": ProcessDeleteAbsence"

        GoTo REPORTBRV
    End If

    If Not (CheckForValidLong(vlTRISUserNumber, True)) Then
        AppendBusinessError radrsBRV _
            , trErrbusEAM _
            , mstrCBUSINESSERR _
            , "Invalid TRIS User." _
            , MODULENAME & ": ProcessDeleteAbsence"

        GoTo REPORTBRV
    End If

    'Get Last Processed Pay Period
    ldtLastProcessedDate = #1/1/100#
    Set lobjCollPeriod = CreateObject("busCollectionPeriod.cbsCollectionPeriod")
    Set ladrsPayPeriod = lobjCollPeriod.GetCollectionPeriodStatusByEmployee( _
                    vadrsAbsence!EMP_NUM, CDate("12/30/1899"), 0)
    If (RecordsExist(ladrsPayPeriod)) Then
        ldtLastProcessedDate = lobjCollPeriod.GetLatestProcessedPayPeriodEndDate( _
                                    ladrsPayPeriod!PAYROLL_GROUP_CD)
    End If
    Set lobjCollPeriod = Nothing
    Set ladrsPayPeriod = Nothing

    If (vadrsAbsence!ABS_ST_DT <= ldtLastProcessedDate) Or _
        (vadrsAbsence!ABS_END_DT <= ldtLastProcessedDate) Then
        AppendBusinessError radrsBRV _
            , trErrbusEAM _
            , mstrCBUSINESSERR _
            , "Absence cannot be deleted. It has been processed by payroll." _
            , MODULENAME & ": ProcessDeleteAbsence"

        GoTo REPORTBRV
    End If

    Set lobjdbEAMSave = CreateObject("dbEAM.cdEAMTrans")
    If Not (lobjdbEAMSave.dbProcessDeleteAbsence(vadrsAbsence)) Then
        Set lobjdbEAMSave = Nothing

        Err.Raise trErrbusEAM
    End If
    Set lobjdbEAMSave = Nothing

    vadrsAbsence.MoveFirst

    Set lobjSchedExc = CreateObject("busSchedException.cbsSchedExc")
    Set ladrsSchedExcDtls = lobjSchedExc.GetScheduledExceptions( _
                            vadrsAbsence!EMP_NUM _
                            , vadrsAbsence!ABS_ST_DT _
                            , vadrsAbsence!ABS_END_DT)
    Set ladrsTimeDetails = lobjSchedExc.GetViewCalendarDetails( _
                            vadrsAbsence!EMP_NUM _
                            , vadrsAbsence!ABS_ST_DT _
                            , vadrsAbsence!ABS_END_DT)
    Set lobjSchedExc = Nothing

    Set lobjSchedule = CreateObject( _
                        "busSchedule.cbsSchedule")
    Set ladrsSchedule = lobjSchedule.GetAdjustedBaseSchedule(vadrsAbsence!EMP_NUM _
            , vadrsAbsence!ABS_ST_DT _
            , vadrsAbsence!ABS_END_DT _
            , ladrsScheduleNoExc _
            , True _
            , trAdjustedBaseSched _
            , trAllScheduleLevels _
            , trRelief _
            , trAllExceptions)
    Set lobjSchedule = Nothing
       
    PrepareDeleteExcTimeRS vadrsAbsence _
        , ladrsSchedExcDtls _
        , ladrsTimeDetails _
        , ladrsSchedule _
        , Nothing _
        , vsApprovalAuthority _
        , ldtLastProcessedDate _
        , ladrsActual _
        , ladrsSchedExc

    Set ladrsTimeDetails = Nothing
    Set ladrsSchedExcDtls = Nothing

    '//this code was effectively made inactive 11/10/2006 because it was calling ProcessTimeEntries
    '//at that time ProcessTimeEntries was obsolete and ProcessTimeEntries2 was created for the EMPRV interface
    '//unfortuneatly ProcessTimeEntries did nothing rather than redirect the call to ProcessTimeEntries2 or throw an exception
    '//now, since it is actually doing something, it does not work !!!
    '//need to add these columns (JOB_NBR, JOBTASK_NBR, JOBTASK_STEP_NBR, EMPRV_DATABASE_ID) for the EMPRV interface
    '//busDistributionMgr.cbsDistributionMgr.ValidateDistributionIDRemoveOtherPay gets a null exception with r_rsOtherPayEntries
    If (RecordsExist(ladrsActual)) Then
        Set lobjbusTimeKeeper = CreateObject("busTimekeeper.cbsTimekeeper")
        If Not (lobjbusTimeKeeper.ProcessTimeEntries3(ladrsActual _
                , Nothing _
                , Nothing _
                , vadrsAbsence!ABS_ST_DT _
                , vadrsAbsence!ABS_END_DT _
                , Nothing _
                , Nothing _
                , ladrsScheduleNoExc _
                , ladrsBRVActual _
                , CStr(vlTRISUserNumber))) Then
            Set lobjbusTimeKeeper = Nothing

            If (RecordsExist(ladrsBRVActual)) Then
                AddBRVRecords ladrsBRVActual _
                    , radrsBRV _
                    , True
            End If

            GoTo REPORTBRV
        Else
            Set lobjbusTimeKeeper = Nothing

            If (RecordsExist(ladrsBRVActual)) Then
                AddBRVRecords ladrsBRVActual _
                    , radrsBRV _
                    , True
            End If
        End If

        Set lobjbusTimeKeeper = Nothing
    End If

    If (RecordsExist(ladrsSchedExc)) Then
        Set lobjSchedExcSave = CreateObject("busSchedException.cbsSchedExcTrans")
        If Not (lobjSchedExcSave.SaveExceptions(ladrsSchedExc _
                , ladrsBRVSchedExc _
                , CStr(vlTRISUserNumber))) Then
            Set lobjSchedExcSave = Nothing

            If (RecordsExist(ladrsBRVSchedExc)) Then
                AddBRVRecords ladrsBRVSchedExc, radrsBRV
            End If

            GoTo REPORTBRV
        Else
            Set lobjSchedExcSave = Nothing

            If (RecordsExist(ladrsBRVSchedExc)) Then
                AddBRVRecords ladrsBRVSchedExc, radrsBRV
            End If
        End If

        Set lobjSchedExcSave = Nothing
    End If

    Set ladrsPayPeriod = Nothing
    Set ladrsTimeDetails = Nothing
    Set ladrsSchedExcDtls = Nothing
    Set ladrsAbsence = Nothing
    Set ladrsScheduleNoExc = Nothing
    Set ladrsSchedule = Nothing
    Set ladrsSchedExc = Nothing
    Set ladrsActual = Nothing

    Set lobjCollPeriod = Nothing
    Set lobjSchedule = Nothing
    Set lobjdbEAMSave = Nothing
    Set lobjSchedExc = Nothing
    Set lobjSchedExcSave = Nothing
    Set lobjbusTimeKeeper = Nothing

    ProcessDeleteAbsence = True

    GetObjectContext.SetComplete

    Exit Function
ProcessDeleteAbsence_EH:
    'Abort Transaction
    GetObjectContext().SetAbort

    ProcessDeleteAbsence = False

    Set ladrsPayPeriod = Nothing
    Set ladrsTimeDetails = Nothing
    Set ladrsSchedExcDtls = Nothing
    Set ladrsAbsence = Nothing
    Set ladrsScheduleNoExc = Nothing
    Set ladrsSchedule = Nothing
    Set ladrsSchedExc = Nothing
    Set ladrsBRVSchedExc = Nothing
    Set ladrsActual = Nothing
    Set ladrsBRVActual = Nothing

    Set lobjCollPeriod = Nothing
    Set lobjSchedule = Nothing
    Set lobjdbEAMSave = Nothing
    Set lobjSchedExc = Nothing
    Set lobjSchedExcSave = Nothing
    Set lobjbusTimeKeeper = Nothing

    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "ProcessDeleteAbsence"
            Case Else
                TrsRaiseError MODULENAME _
                    , "ProcessDeleteAbsence" _
                    , Err.Number _
                    , , , , _
                    True
    End Select

    Exit Function
REPORTBRV:
    Set ladrsPayPeriod = Nothing
    Set ladrsTimeDetails = Nothing
    Set ladrsSchedExcDtls = Nothing
    Set ladrsAbsence = Nothing
    Set ladrsScheduleNoExc = Nothing
    Set ladrsSchedule = Nothing
    Set ladrsSchedExc = Nothing
    Set ladrsActual = Nothing
    Set ladrsBRVSchedExc = Nothing
    Set ladrsBRVActual = Nothing

    Set lobjCollPeriod = Nothing
    Set lobjSchedule = Nothing
    Set lobjdbEAMSave = Nothing
    Set lobjSchedExc = Nothing
    Set lobjSchedExcSave = Nothing
    Set lobjbusTimeKeeper = Nothing

    If (RecordsExist(radrsBRV)) Then
        With radrsBRV
            .MoveFirst

            'Check for errors
            .Find "ERR_TYPE = 'BE'"
            'If found abort else complete
            If Not (.EOF) Then
                'Abort Transaction
                GetObjectContext().SetAbort
                ProcessDeleteAbsence = False
            Else
                'Complete Transaction
                GetObjectContext().SetComplete
                ProcessDeleteAbsence = True
            End If

            .MoveFirst
        End With
    Else
        'Complete Transaction
        GetObjectContext().SetComplete
        ProcessDeleteAbsence = True
    End If

End Function

Private Sub PrepareDeleteExcTimeRS(ByVal vadrsAbsence As ADODB.Recordset _
                    , ByVal vadrsSchedExcDtls As ADODB.Recordset _
                    , ByVal vadrsTimeDetails As ADODB.Recordset _
                    , ByVal vadrsSchedule As ADODB.Recordset _
                    , ByVal vadrsRemovedDates As ADODB.Recordset _
                    , ByVal vsApprovalAuthority As String _
                    , ByVal vdtLastProcessedDate As Date _
                    , ByRef radrsActual As ADODB.Recordset _
                    , ByRef radrsSchedExc As ADODB.Recordset)
On Error GoTo eh
'NOTE: TIME_TYPE 1 = ADJUSTMENT; 2 = ACTUAL; 3 = SCHEDULED
Dim llIndex As Long, lbAddRecord As Boolean
Dim ldtAbsenceStart As Date, lbProceed As Boolean

    If (RecordsExist(vadrsAbsence)) And _
        (RecordsExist(vadrsSchedExcDtls) Or _
        RecordsExist(vadrsTimeDetails)) Then
        vadrsAbsence.MoveFirst
        ldtAbsenceStart = vadrsAbsence!ABS_ST_DT

        '//match the abs start time to the schedule shift and use its dateworked for the abs start date
        If Not (vadrsSchedule Is Nothing) Then
            With vadrsSchedule
                If Not (.EOF And .BOF) Then
                    .Filter = adFilterNone

                    .Filter = "CHARGE_DATE = #" & vadrsAbsence!ABS_ST_DT & "#"
                    If Not (.EOF) Then
                        .Sort = "CHARGE_DATE, DATE_WORKED, START_TIME, END_TIME"
                        .MoveFirst

                        For llIndex = 1 To .RecordCount
                            If (!DATE_WORKED < !CHARGE_DATE) Then
                                If CDate(vadrsAbsence!ABS_ST_DT & " " & vadrsAbsence!ABS_ST_TM) > CDate(!DATE_WORKED & " " & !SHIFT_START_TIME) And _
                                    CDate(vadrsAbsence!ABS_ST_DT & " " & vadrsAbsence!ABS_ST_TM) > CDate(!DATE_WORKED & " " & !SHIFT_END_TIME) Then
                                    ldtAbsenceStart = !DATE_WORKED
                                    Exit For
                                End If
                            End If

                            .MoveNext
                        Next llIndex
                    End If

                    .Filter = adFilterNone
                End If
            End With
        End If

        vadrsAbsence.MoveFirst

        If (RecordsExist(vadrsTimeDetails)) Then
            With vadrsTimeDetails
                .Filter = adFilterNone

                'Actual
                .Filter = ("TIME_TYPE = 2")
                If Not (.EOF) Then
                    .Sort = "CHARGE_DATE, DATE_WORKED, START_TIME, END_TIME, HOURS"
                    .MoveFirst

                    Set radrsActual = CreateTimeRS(2)

                    For llIndex = 1 To .RecordCount
                        lbProceed = True

                        If (!CHARGE_DATE > vdtLastProcessedDate) Then
                            If (IsDate(!START_TIME)) Or _
                                (IsDate(!END_TIME)) Then
                                '//if called by ProcessedEAMRequest, then only process dates that have been removed from absence
                                If (RecordsExist(vadrsRemovedDates)) Then
                                    vadrsRemovedDates.MoveFirst

                                    vadrsRemovedDates.Find "ABS_DATE = #" & !CHARGE_DATE & "#"
                                    If Not (vadrsRemovedDates.EOF) Then
                                        If (vadrsRemovedDates!RECORD_STATUS <> "D") Then
                                            lbProceed = False
                                        End If
                                    Else
                                        lbProceed = False
                                    End If

                                    vadrsRemovedDates.MoveFirst
                                End If

                                If (lbProceed) Then
                                    If DateTimeFrom(!CHARGE_DATE, !START_TIME, False) >= DateTimeFrom(ldtAbsenceStart, ConvertDB2Time(vadrsAbsence!ABS_ST_TM), False) And _
                                        DateTimeFrom(!CHARGE_DATE, !END_TIME, True) <= DateTimeFrom(vadrsAbsence!ABS_END_DT, ConvertDB2Time(vadrsAbsence!ABS_END_TM), True) And _
                                        (Trim(!EXCEPTION_CD) <> "" Or Trim(!OVER_EXC_CD) <> "") Then
                                        lbAddRecord = True

                                        If (!APPR_BY_EMP_NUM > 0) Then
                                            If (vsApprovalAuthority <> "Y") Then
                                                lbAddRecord = False
                                            End If
                                        End If

                                        If (lbAddRecord) Then
                                            radrsActual.AddNew

                                            radrsActual!EMP_NUM = Null
                                            radrsActual!CHARGE_DATE = Null
                                            radrsActual!KEY_EMP_NUM = !EMP_NUM
                                            radrsActual!KEY_CHARGE_DATE = !CHARGE_DATE
                                            radrsActual!KEY_START_TIME = !START_TIME
                                            radrsActual!KEY_INDEX_TIME = !INDEX_TIME
                                            radrsActual!RECORD_STATUS = "D"
                                            radrsActual!SCHED_TIME_IND = "N"

                                            radrsActual.Update
                                        End If
                                    End If
                                End If
                            End If
                        End If

                        .MoveNext
                    Next llIndex
                End If

                .Filter = adFilterNone
            End With
        End If

        If (RecordsExist(vadrsSchedExcDtls)) Then
            With vadrsSchedExcDtls
                .Filter = adFilterNone
                If Not (.EOF) Then
                    .Sort = "CHARGE_DATE, START_TIME, END_TIME, HOURS"
                    .MoveFirst

                    Set radrsSchedExc = CreateTimeRS(3)

                    For llIndex = 1 To .RecordCount
                        lbProceed = True
                        
                        '//if called by ProcessedEAMRequest, then only process dates that have been removed from absence
                        If (RecordsExist(vadrsRemovedDates)) Then
                            vadrsRemovedDates.MoveFirst

                            vadrsRemovedDates.Find "ABS_DATE = #" & !CHARGE_DATE & "#"
                            If Not (vadrsRemovedDates.EOF) Then
                                If (vadrsRemovedDates!RECORD_STATUS <> "D") Then
                                    lbProceed = False
                                End If
                            Else
                                lbProceed = False
                            End If

                            vadrsRemovedDates.MoveFirst
                        End If

                        If (lbProceed) Then
                            If DateTimeFrom(!CHARGE_DATE, !START_TIME, False) >= DateTimeFrom(ldtAbsenceStart, ConvertDB2Time(vadrsAbsence!ABS_ST_TM), False) And _
                                DateTimeFrom(!CHARGE_DATE, !END_TIME, True) <= DateTimeFrom(vadrsAbsence!ABS_END_DT, ConvertDB2Time(vadrsAbsence!ABS_END_TM), True) And _
                                (Trim(!EXCEPTION_CD) <> "" Or Trim(!OVER_EXC_CD) <> "") Then
                                radrsSchedExc.AddNew

                                radrsSchedExc!EMP_NUM = vadrsAbsence!EMP_NUM
                                radrsSchedExc!CHARGE_DATE = !CHARGE_DATE
                                radrsSchedExc!START_TIME = !START_TIME
                                radrsSchedExc!END_TIME = !END_TIME
                                radrsSchedExc!EXCEPTION_CD = !EXCEPTION_CD
                                radrsSchedExc!OVER_EXC_CD = !OVER_EXC_CD
                                radrsSchedExc!CASE_NUM = !CASE_NUM
                                radrsSchedExc!KEY_CHARGE_DATE = !CHARGE_DATE
                                radrsSchedExc!KEY_START_TIME = !START_TIME
                                radrsSchedExc!KEY_END_TIME = !END_TIME
                                radrsSchedExc!KEY_EXCEPTION_CD = !EXCEPTION_CD
                                radrsSchedExc!KEY_OVER_EXC_CD = !OVER_EXC_CD
                                radrsSchedExc!CREATED_TIMESTAMP = !CREATED_TIMESTAMP
                                radrsSchedExc!DELETE_ACTUAL_TIME = "Y"
                                radrsSchedExc!APPR_AUTH_IND = vsApprovalAuthority
                                radrsSchedExc!REC_STATUS = "D"

                                radrsSchedExc.Update
                            End If
                        End If

                        .MoveNext
                    Next llIndex
                End If

                .Filter = adFilterNone
            End With
        End If
    End If

    If (RecordsExist(radrsActual)) Then
        If (radrsActual.RecordCount > 0) Then
            radrsActual.MoveFirst
        End If
    End If

    If (RecordsExist(radrsSchedExc)) Then
        If (radrsSchedExc.RecordCount > 0) Then
            radrsSchedExc.MoveFirst
        End If
    End If

    Exit Sub
eh:
    TrsRaiseError MODULENAME, "PrepareDeleteExcTimeRS"
End Sub

Private Function CreateTimeRS(ByVal viTimeType As Integer _
                ) As ADODB.Recordset
On Error GoTo CreateTimeRS_EH
Dim ladrs As ADODB.Recordset

    Set CreateTimeRS = Nothing

    Set ladrs = New ADODB.Recordset
    With ladrs
        Select Case (viTimeType)
            Case 3
                .Fields.Append "EMP_NUM", adInteger, , ADOFIELDATTR

                .Fields.Append "CHARGE_DATE", adDate, , adFldUpdatable
                .Fields.Append "SHIFT_NUM", adInteger, , adFldUpdatable
                .Fields.Append "FULL_SHIFT_EXC_IND", adChar, 1, adFldUpdatable
                .Fields.Append "START_TIME", adChar, 5, adFldUpdatable
                .Fields.Append "END_TIME", adChar, 5, adFldUpdatable
                .Fields.Append "EXCEPTION_CD", adChar, 3, adFldUpdatable
                .Fields.Append "OVER_EXC_CD", adChar, 3, adFldUpdatable
                .Fields.Append "FMLA_IND", adChar, 1, adFldUpdatable
                .Fields.Append "CASE_NUM", adInteger, , adFldUpdatable
                .Fields.Append "WG_NUM", adInteger, , adFldUpdatable
                .Fields.Append "JOB_NUM", adChar, 6, adFldUpdatable
                .Fields.Append "RELIEF_COMP", adChar, 3, adFldUpdatable
                .Fields.Append "COMP_CD", adChar, 3, adFldUpdatable
                .Fields.Append "SHIFT_COMP", adChar, 3, adFldUpdatable
                .Fields.Append "DELETE_ACTUAL_TIME", adChar, 1, adFldUpdatable
                .Fields.Append "APPR_AUTH_IND", adChar, 1, adFldUpdatable

                .Fields.Append "KEY_CHARGE_DATE", adDate, , adFldUpdatable
                .Fields.Append "KEY_START_TIME", adChar, 5, adFldUpdatable
                .Fields.Append "KEY_END_TIME", adChar, 5, adFldUpdatable
                .Fields.Append "KEY_EXCEPTION_CD", adChar, 3, adFldUpdatable
                .Fields.Append "KEY_OVER_EXC_CD", adChar, 3, adFldUpdatable
                .Fields.Append "KEY_CASE_NUM", adInteger, , adFldUpdatable
                .Fields.Append "CREATED_TIMESTAMP", adChar, 26, adFldUpdatable

                .Fields.Append "REC_STATUS", adChar, 1, adFldUpdatable

            Case 2
                .Fields.Append "EMP_NUM", adInteger, , ADOFIELDATTR
                .Fields.Append "CHARGE_DATE", adDate, , ADOFIELDATTR
                .Fields.Append "DATE_WORKED", adDate, , ADOFIELDATTR
                .Fields.Append "DISTRIBUTION_ID", adChar, 39, ADOFIELDATTR
                .Fields.Append "DISTRIBUTION_TYPE", adChar, 1, ADOFIELDATTR
                .Fields.Append "HOURS", adDecimal, , ADOFIELDATTR
                    .Fields.Item("HOURS").Precision = 3
                    .Fields.Item("HOURS").NumericScale = 1
                .Fields.Append "START_TIME", adChar, 5, ADOFIELDATTR
                .Fields.Append "END_TIME", adChar, 5, ADOFIELDATTR
                .Fields.Append "EXCEPTION_CD", adChar, 3, ADOFIELDATTR
                .Fields.Append "OVER_EXC_CD", adChar, 3, ADOFIELDATTR
                .Fields.Append "JOB_NUM", adChar, 6, ADOFIELDATTR
                .Fields.Append "OT_REASON_CD", adChar, 3, ADOFIELDATTR
                .Fields.Append "PREARRANGED_OT_IND", adChar, 1, ADOFIELDATTR
                .Fields.Append "OT_OPP_DESC", adVarChar, 60, ADOFIELDATTR
                .Fields.Append "OT_TRAVEL_PAY_IND", adChar, 1, ADOFIELDATTR
                .Fields.Append "OT_ADJUST_DESC", adChar, 30, ADOFIELDATTR
                .Fields.Append "WG_NUM", adInteger, , ADOFIELDATTR
                .Fields.Append "CASE_NUM", adInteger, , ADOFIELDATTR
                .Fields.Append "FAMILY_LEAVE_IND", adChar, 1, ADOFIELDATTR
                .Fields.Append "COMP_SHIFT_CD", adChar, 3, ADOFIELDATTR
                .Fields.Append "COMP_CD", adChar, 3, ADOFIELDATTR
                .Fields.Append "REL_CD", adChar, 3, ADOFIELDATTR
                .Fields.Append "SCHED_CHG_EXC_CD", adChar, 3, ADOFIELDATTR
                .Fields.Append "DISTR_DEPT_NUM", adChar, 3, ADOFIELDATTR
                .Fields.Append "APPR_IND", adChar, 1, ADOFIELDATTR
                'RICK SAYS PASS TIMESTAMP: NOV, 2000
                .Fields.Append "TIME_STAMP", adChar, 26, ADOFIELDATTR
                .Fields.Append "JOB_NBR", adChar, 8, ADOFIELDATTR
                .Fields.Append "JOBTASK_NBR", adChar, 4, ADOFIELDATTR
                .Fields.Append "JOBTASK_STEP_NBR", adInteger, , ADOFIELDATTR
                .Fields.Append "EMPRV_DATABASE_ID", adChar, 8, ADOFIELDATTR

                'Non-nullable information
                '   RECORD_STATUS CAN BE
                '   D - DELETE
                '   I - INSERT
                '   C - CHANGE
                '   SCHED_TIME_IND CAN BE
                '   Y - Yes Scheduled Time
                '   N - No Actual Time
                .Fields.Append "RECORD_STATUS", adChar, 1, adFldUpdatable
                .Fields.Append "SCHED_TIME_IND", adChar, 1, adFldUpdatable

                'KEY INFORMATION
                .Fields.Append "KEY_EMP_NUM", adInteger, , ADOFIELDATTR
                .Fields.Append "KEY_CHARGE_DATE", adDate, , ADOFIELDATTR
                .Fields.Append "KEY_START_TIME", adChar, 5, ADOFIELDATTR
                .Fields.Append "KEY_INDEX_TIME", adChar, 26, ADOFIELDATTR
        End Select

        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open
    End With

    Set CreateTimeRS = ladrs
    Set ladrs = Nothing

    Exit Function
CreateTimeRS_EH:
    Set ladrs = Nothing

    TrsRaiseError MODULENAME _
        , "CreateTimeRS"
End Function
Private Sub EnsureTimeDeletion(ByVal vadrsScheduledExceptions As ADODB.Recordset _
                , ByRef radrsDeleteTime As ADODB.Recordset)
On Error GoTo EnsureTimeDeletion_EH
Dim llOuterIndex As Long, llInnerIndex As Long

    If (RecordsExist(vadrsScheduledExceptions)) And _
        (RecordsExist(radrsDeleteTime)) Then
        radrsDeleteTime.MoveFirst

        With vadrsScheduledExceptions
            .Sort = "CHARGE_DATE"
            .MoveFirst

            For llOuterIndex = 1 To .RecordCount
                If (!REC_STATUS = "I") Or _
                    (!REC_STATUS = "C") Then
                    radrsDeleteTime.Filter = adFilterNone

                    radrsDeleteTime.Filter = "KEY_CHARGE_DATE = " & !CHARGE_DATE
                    If Not (radrsDeleteTime.EOF) Then
                        radrsDeleteTime.MoveFirst

                        For llInnerIndex = radrsDeleteTime.RecordCount To 1 Step -1
                            radrsDeleteTime.Delete adAffectCurrent
                            radrsDeleteTime.Update
                            radrsDeleteTime.MoveNext
                        Next llInnerIndex
                    End If

                    radrsDeleteTime.Filter = adFilterNone

                    .MoveNext
                End If
            Next llOuterIndex
        End With
    End If

    If (RecordsExist(radrsDeleteTime)) Then
        If (radrsDeleteTime.RecordCount > 0) Then
            radrsDeleteTime.Filter = adFilterNone
            radrsDeleteTime.MoveFirst
        End If
    End If

    Exit Sub
EnsureTimeDeletion_EH:
    TrsRaiseError MODULENAME _
        , "EnsureTimeDeletion"

End Sub
Public Function ProcessDeletePortfolio(ByVal vlPortfolioNumber As Long _
                , ByVal vlTRISUserNumber As Long _
                , ByVal vsTRISUserID As String _
                , ByRef radrsBRV As ADODB.Recordset _
                ) As Boolean
On Error GoTo ProcessDeletePortfolio_EH
'Components
Dim lobjdbEAMSave As dbEAM.cdEAMTrans

'NO NEED TO CHECK IF THE USER HAS AUTHORITY TO DELETE PORTFOLIO.

    Set radrsBRV = CreateNewBusinessErrorSet

    If Not (CheckForValidLong(vlPortfolioNumber, True)) Then
        AppendBusinessError radrsBRV _
            , trErrbusEAM _
            , mstrCBUSINESSERR _
            , "Invalid Portfolio." _
            , MODULENAME & ": ProcessDeletePortfolio"

        GoTo REPORTBRV
    Else
        If Not (CheckForExistence(trPORTFOLIO, , vlPortfolioNumber)) Then
            AppendBusinessError radrsBRV _
                , trErrbusEAM _
                , mstrCBUSINESSERR _
                , "Invalid Portfolio. Portfolio " & CStr(vlPortfolioNumber) & " does not exist." _
                , MODULENAME & ": ProcessDeletePortfolio"

            GoTo REPORTBRV
        End If
    End If

    If Not (CheckForValidLong(vlTRISUserNumber, True)) Then
        AppendBusinessError radrsBRV _
            , trErrbusEAM _
            , mstrCBUSINESSERR _
            , "Invalid TRIS User." _
            , MODULENAME & ": ProcessDeletePortfolio"

        GoTo REPORTBRV
    End If

    Set lobjdbEAMSave = CreateObject("dbEAM.cdEAMTrans")
    If Not (lobjdbEAMSave.dbProcessDeletePortfolio(vlPortfolioNumber)) Then
        Set lobjdbEAMSave = Nothing

        Err.Raise trErrbusAbsencePortfolio
    End If
    Set lobjdbEAMSave = Nothing

    ProcessDeletePortfolio = True

    GetObjectContext.SetComplete

    Exit Function
ProcessDeletePortfolio_EH:
    'Abort Transaction
    GetObjectContext().SetAbort

    ProcessDeletePortfolio = False

    Set lobjdbEAMSave = Nothing

    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "ProcessDeletePortfolio"
            Case Else
                TrsRaiseError MODULENAME _
                    , "ProcessDeletePortfolio" _
                    , Err.Number _
                    , , , , _
                    True
    End Select

    Exit Function
REPORTBRV:
    Set lobjdbEAMSave = Nothing

    If (RecordsExist(radrsBRV)) Then
        With radrsBRV
            .MoveFirst

            'Check for errors
            .Find "ERR_TYPE = 'BE'"
            'If found abort else complete
            If Not (.EOF) Then
                'Abort Transaction
                GetObjectContext().SetAbort
                ProcessDeletePortfolio = False
            Else
                'Complete Transaction
                GetObjectContext().SetComplete
                ProcessDeletePortfolio = True
            End If

            .MoveFirst
        End With
    Else
        'Complete Transaction
        GetObjectContext().SetComplete
        ProcessDeletePortfolio = True
    End If

End Function
Private Function GetLockStatus(ByVal vlEmployee As Long _
                    , ByVal vdtStartDate As Date _
                    , ByVal vdtEndDate As Date _
                    ) As ADODB.Recordset
On Error GoTo GetLockStatus_EH
'Objects
Dim lobjEmployee As busEmployee.cbsEmployee
Dim lobjdbEmployee As dbEmployee.cdEmployee
Dim lobjdbCollectionPeriods As dbCollectionPeriod.cdCollectionPeriod
'Recordsets
Dim ladrs As ADODB.Recordset, ladrsEmployee As ADODB.Recordset
Dim ladrsCollectionPeriods As ADODB.Recordset
'Other
Dim ldtLoopDate As Date, lsLockStatus As String
'Collection
Dim lcolPayrollGroup As Collection

    'Define RS
    Set ladrs = New ADODB.Recordset
    With ladrs
        .Fields.Append "CHARGE_DATE", adDate, , ADOFIELDATTR
        .Fields.Append "LOCK_STATUS", adChar, 1, ADOFIELDATTR

        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open
    End With

    'Get Employee Information
    Set lobjdbEmployee = CreateObject("dbEmployee.cdEmployee")
    Set ladrsEmployee = lobjdbEmployee.GetEmployeeExcSchedInfo(vlEmployee)
    Set lobjdbEmployee = Nothing

    'Get Collection Periods
    Set lcolPayrollGroup = New Collection
    Set lobjdbCollectionPeriods = CreateObject("dbCollectionPeriod.cdCollectionPeriod")
    Set ladrsCollectionPeriods = lobjdbCollectionPeriods.GetCollectionPeriods(vdtStartDate _
                                    , vdtEndDate _
                                    , lcolPayrollGroup)
    Set lobjdbCollectionPeriods = Nothing

    ldtLoopDate = vdtStartDate

    'Get Lock Status
    Set lobjEmployee = CreateObject("busEmployee.cbsEmployee")
    Do Until ldtLoopDate > vdtEndDate
        lsLockStatus = lobjEmployee.GetLockStatus(vlEmployee _
                            , ldtLoopDate _
                            , ladrsEmployee _
                            , ladrsCollectionPeriods)

        ladrs.AddNew
        ladrs!CHARGE_DATE = ldtLoopDate
        ladrs!LOCK_STATUS = lsLockStatus
        ladrs.Update

        ldtLoopDate = DateAdd("d", 1, ldtLoopDate)
    Loop
    Set lobjEmployee = Nothing

    Set lobjEmployee = Nothing
    Set lobjdbEmployee = Nothing
    Set lobjdbCollectionPeriods = Nothing

    Set ladrsEmployee = Nothing
    Set ladrsCollectionPeriods = Nothing
    Set lcolPayrollGroup = Nothing

    If (RecordsExist(ladrs)) Then
        ladrs.MoveFirst
    End If

    Set GetLockStatus = ladrs
    Set ladrs = Nothing

    Exit Function
GetLockStatus_EH:
    Set lobjEmployee = Nothing
    Set lobjdbEmployee = Nothing
    Set lobjdbCollectionPeriods = Nothing

    Set ladrsEmployee = Nothing
    Set ladrsCollectionPeriods = Nothing
    Set ladrs = Nothing
    Set lcolPayrollGroup = Nothing

    TrsRaiseError MODULENAME _
        , "GetLockStatus"
End Function
Private Function CheckForExistence(ByVal viType As Integer _
                , Optional ByVal vsStringItem As String = "" _
                , Optional ByVal vlLongItem As Long = 0 _
                , Optional ByVal viIntegerItem As Integer = 0 _
                , Optional ByVal vsAssociatedStringItem As String = "" _
                , Optional ByVal vlAssociatedLongItem As Long = 0 _
                , Optional ByVal viAssociatedIntegerItem As Integer = 0 _
                ) As Boolean
On Error GoTo CheckForExistence_EH
Dim lobjdbWorkgroup As dbWorkGroup.cdWorkGroup
Dim lobjEmployee As busEmployee.cbsEmployee
Dim lobjdbEAM As dbEAM.cdEAM
Dim lobjCodeTables As dbCodeTables.cdCodeTables

    CheckForExistence = False

    Select Case (viType)
        Case trEMPLOYEE
            Set lobjEmployee = CreateObject("busEmployee.cbsEmployee")
            CheckForExistence = lobjEmployee.IsValidEmployee(vlLongItem)
            Set lobjEmployee = Nothing

        Case trABSENCE
            Set lobjdbEAM = CreateObject("dbEAM.cdEAM")
            CheckForExistence = lobjdbEAM.IsAValidAbsence(vlLongItem)
            Set lobjdbEAM = Nothing

        Case trEXCEPTION
            Set lobjCodeTables = CreateObject("dbCodeTables.cdCodeTables")
            CheckForExistence = lobjCodeTables.ExceptionCodeIsValid(vsStringItem)
            Set lobjCodeTables = Nothing

        Case trLOCATION
            Set lobjCodeTables = CreateObject("dbCodeTables.cdCodeTables")
            CheckForExistence = lobjCodeTables.LocationCodeIsValid(CStr(viIntegerItem))
            Set lobjCodeTables = Nothing

        Case trFMLAREASON
            CheckForExistence = True
'            Set lobjCodeTables = CreateObject("dbCodeTables.cdCodeTables")
'            CheckForExistence = lobjCodeTables.ExceptionCodeIsValid(vsStringItem)
'            Set lobjCodeTables = Nothing

        Case trPROVIDER
            Set lobjdbEAM = CreateObject("dbEAM.cdEAM")
            CheckForExistence = lobjdbEAM.IsAValidProvider(vlLongItem)
            Set lobjCodeTables = Nothing

        Case trABSENCECONTACT
            Set lobjdbEAM = CreateObject("dbEAM.cdEAM")
            CheckForExistence = lobjdbEAM.IsAValidAbsenceContact(vlLongItem _
                                , vlAssociatedLongItem)
            Set lobjdbEAM = Nothing

        Case trPORTFOLIO
            Set lobjdbEAM = CreateObject("dbEAM.cdEAM")
            CheckForExistence = lobjdbEAM.IsAValidPortfolio(vlLongItem)
            Set lobjdbEAM = Nothing

        Case trWORKGROUP
            Set lobjdbWorkgroup = CreateObject("dbWorkgroup.cdWorkgroup")
            CheckForExistence = lobjdbWorkgroup.IsValidWorkgroup(vlLongItem _
                                    , Date)
            Set lobjdbWorkgroup = Nothing
    End Select

    Exit Function
CheckForExistence_EH:
    Set lobjEmployee = Nothing
    Set lobjCodeTables = Nothing
    Set lobjdbEAM = Nothing
    Set lobjdbWorkgroup = Nothing

    TrsRaiseError MODULENAME _
        , "CheckForExistence"
End Function

Private Sub Class_Initialize()
    Set m_ResourceMgr = New utilADOHelper.cuResourceMgr
End Sub

Private Sub Class_Terminate()
    Set m_ResourceMgr = Nothing
End Sub

