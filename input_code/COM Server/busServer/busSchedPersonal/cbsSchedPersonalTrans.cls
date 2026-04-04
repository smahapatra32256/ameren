VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "cbsSchedPersonalTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const mstrCMODULENAME As String = "busSchedPersonalTrans"
Private Const mstrCVARFIELDNAMES As String = "EMP_NUM;START_DATE;END_DATE"
Private Const m_datCOPTIONALSTARTDATE As Date = #1/1/100#
Private Const m_datCOPTIONALENDDATE As Date = #12/31/9999#
Private Const m_lngEMPNOCONVERSIONMINUSONE As Long = -1
Private Const m_lngEMPNOCONVERSIONID As Long = 99999  'temporarily use Greg's id
Private Const mblnAPPLYCASCADE As Boolean = True

Private m_lngRecordId As Long
Private m_lngUserId As Long
Private m_lngCreatedUserNo As Long
Private m_lngErrorNo As Long
Private m_colCascadeIntervals As Collection
Private m_HolidayCalendar As ADODB.Recordset

Private Enum TRSSchedSourceStartLevel
    trAllScheduleLevels = 50
    trDemandLineTrades = 50
    trDemandLineBaseSchedLevel = 40
    trEmployeeTradesLevel = 30
    trNonDemandLevel = 20
    trWorkGroupBaseScheduleLevel = 10
End Enum

Private Enum TrTypeOfUpdate
    DELETENDS_1 = 1
    UPDATENDS_2 = 2
    INSERTNDS_3 = 3
End Enum

'Request personal schedule for a single employee
Private Function GetPersonalSchedules _
                (ByVal v_lngEmpNum As Long _
                 , Optional ByVal v_datStartDate As Date = m_datCOPTIONALSTARTDATE _
                 , Optional ByVal v_datEndDate As Date = m_datCOPTIONALENDDATE _
                ) As ADODB.Recordset


On Error GoTo GetPersonalSchedules_EH

    Dim l_dbSchedPersonal As dbSchedPersonal.cdSchedPersonal

    Set l_dbSchedPersonal = CreateObject("dbSchedPersonal.cdSchedPersonal")

    Set GetPersonalSchedules = l_dbSchedPersonal.GetPersonalSchedules(v_lngEmpNum, v_datStartDate, v_datEndDate)

    Set l_dbSchedPersonal = Nothing
    GetObjectContext.SetComplete

Exit Function

GetPersonalSchedules_EH:
    Set l_dbSchedPersonal = Nothing
    GetObjectContext.SetAbort
    TrsRaiseError mstrCMODULENAME, "GetPersonalSchedules"

End Function
'check for a valid input data set
'do preliminary BRV check
'do the business brv
'if updates exist by original request, perform save process with result net request
Public Function RecordPersonal( _
                    ByVal r_rsPersonalSchedRequest As ADODB.Recordset, _
                    ByVal v_lngUserId As Long, _
                    ByRef r_rsErrors As ADODB.Recordset, _
                    ByRef r_rsNetPersonalSchedUpdates As ADODB.Recordset, _
                    ByRef r_rsCascadeConflicts As ADODB.Recordset, _
                    ByRef r_rsExceptionSPA As ADODB.Recordset, _
                    ByRef r_rsAvailScheduleSPA As ADODB.Recordset, _
                    ByRef r_rsEmpSchedDates As ADODB.Recordset, _
                    ByVal v_blnConfirmConflicts As Boolean, _
                    Optional ByVal v_blnBypassCascade As Boolean = False) _
                As Boolean

On Error GoTo RecordPersonal_EH

    Dim l_blnConfirmPending As Boolean
    Dim l_varCascadeUpdates As Variant
    
    If IsClearedFilterRecordsetEmpty(r_rsPersonalSchedRequest) Then
        RecordPersonal = True
        Exit Function
    Else
    End If
    
    '//-----workaround for 2.0 Gateway dynamic dispatcher
    If Not r_rsErrors Is Nothing Then
        If r_rsErrors.State <> adStateOpen Then
            Set r_rsErrors = Nothing
        End If
    End If
    
    If Val(v_lngUserId) > 0 Then             ' is the user id valid
        If IsValidEmployee(v_lngUserId) Then
            m_lngUserId = v_lngUserId
            m_lngCreatedUserNo = v_lngUserId
        Else
            Err.Raise trErrboEmpNotFound
        End If
    Else
        If Val(v_lngUserId) = m_lngEMPNOCONVERSIONMINUSONE Then 'processing conversion only
           m_lngUserId = m_lngEMPNOCONVERSIONID 'use default user id for last mod emp num
           m_lngCreatedUserNo = r_rsPersonalSchedRequest!CREATED_EMP_NUM
        Else
            Err.Raise trErrdbInvalidZerolngEmpNum
        End If
    End If
    
    RecordPersonal = True
'input records must be sorted
    r_rsPersonalSchedRequest.Sort = "EMP_NUM, START_DATE"
    Set r_rsPersonalSchedRequest = CopyRecordset(r_rsPersonalSchedRequest, , , mCFLDATTRUPDATENULL) 'copy
       
    If RecordsExist(r_rsPersonalSchedRequest) Then
        If Not ValidatePersonalSchedEntry(r_rsPersonalSchedRequest, r_rsErrors) Then
           GetObjectContext.SetAbort
           RecordPersonal = False
           GoTo xit
        End If
        If Not IsBUSValidPersonalSched(r_rsPersonalSchedRequest, r_rsNetPersonalSchedUpdates, r_rsErrors) Then
           GetObjectContext.SetAbort
           RecordPersonal = False
           GoTo xit
        End If
        
        If HasAnyPersonalfieldChanged(r_rsPersonalSchedRequest, r_rsErrors) Then
            If IsClearedFilterRecordsetEmpty(r_rsErrors) Then 'jimb
            Else
                If BRVErrorsExist(r_rsErrors) Then
                    GetObjectContext.SetAbort
                    RecordPersonal = False
                    GoTo xit
                End If
            End If
            'jimb do holiday splits here
            SplitScheduleForHolidayWeeks r_rsNetPersonalSchedUpdates, r_rsPersonalSchedRequest
            
            If Not v_blnBypassCascade Then
                If ProcessCascade( _
                            v_blnConfirmConflicts, l_blnConfirmPending, _
                            l_varCascadeUpdates, r_rsNetPersonalSchedUpdates, _
                            r_rsCascadeConflicts, r_rsExceptionSPA, r_rsAvailScheduleSPA, _
                            r_rsEmpSchedDates, r_rsErrors) Then
                    If l_blnConfirmPending Then
                        RecordPersonal = True
                        GetObjectContext.SetComplete
                        GoTo xit
                    Else
                        Set r_rsCascadeConflicts = Nothing
                        Set r_rsExceptionSPA = Nothing
                        Set r_rsAvailScheduleSPA = Nothing
                    End If
                Else
                    Set r_rsNetPersonalSchedUpdates = Nothing
                    Set r_rsCascadeConflicts = Nothing
                    Set r_rsExceptionSPA = Nothing
                    Set r_rsAvailScheduleSPA = Nothing
                    Set r_rsEmpSchedDates = Nothing
                    GetObjectContext.SetAbort
                    RecordPersonal = False
                    GoTo xit
                End If
            Else
                Set r_rsCascadeConflicts = Nothing
                Set r_rsExceptionSPA = Nothing
                Set r_rsAvailScheduleSPA = Nothing
            End If
        
            If Not ProcessScheduleChange(r_rsEmpSchedDates, r_rsErrors, r_rsNetPersonalSchedUpdates) Then
                GetObjectContext.SetAbort
                Set r_rsEmpSchedDates = Nothing
                RecordPersonal = False
                GoTo xit
            Else
                Set r_rsEmpSchedDates = Nothing
            End If
            
            If SavePersonalSched(r_rsNetPersonalSchedUpdates) = True Then
                RecordPersonal = True
            Else
                GetObjectContext.SetAbort
                RecordPersonal = False
                GoTo xit
            End If
        End If
    End If
    
'    Set l_rsNetPersonalSchedUpdates = Nothing
    GetObjectContext.SetComplete
    
xit:
    Set m_colCascadeIntervals = Nothing
    Exit Function

RecordPersonal_EH:

    RecordPersonal = False
    GetObjectContext.SetAbort
    Set r_rsPersonalSchedRequest = Nothing
    Set r_rsErrors = Nothing
    Set r_rsNetPersonalSchedUpdates = Nothing
    Set r_rsCascadeConflicts = Nothing
    Set r_rsExceptionSPA = Nothing
    Set r_rsAvailScheduleSPA = Nothing
    Set r_rsEmpSchedDates = Nothing
    Set m_colCascadeIntervals = Nothing

    Select Case Err.Number
        Case trErrbusSchedRotDuplicate
            m_lngErrorNo = trErrbusSchedRotDuplicate
            Err.Clear
            'l_strPersonalError = "cannot create duplicate detail record"
            m_lngRecordId = r_rsPersonalSchedRequest!RECORD_ID
            AppendPersonalError r_rsErrors, TrType_PER, m_lngRecordId, "START_DATE", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), r_rsPersonalSchedRequest!EMP_NUM, r_rsPersonalSchedRequest!START_DATE, r_rsPersonalSchedRequest!END_DATE
            Exit Function
        Case trErrboEmpNotFound
            TrsRaiseError mstrCMODULENAME, "RecordPersonal", trErrboEmpNotFound, , , , True
        Case trErrdbInvalidZerolngEmpNum
            TrsRaiseError mstrCMODULENAME, "RecordPersonal", trErrdbInvalidZerolngEmpNum, , , , True
'    Clear the cascade error arguments since the client does not format them.
        Case trErrAppCascadeCommitConflict
            Err.Source = Join(Array(CStr(TrType_PER), "", "", "", ""), ";")
            TrsRaiseError mstrCMODULENAME, "RecordPersonal", , , , , True
        Case Else
            TrsRaiseError mstrCMODULENAME, "RecordPersonal", , , , , True
    End Select

End Function


Public Function ResumeRecordPersonal( _
                    ByVal v_rsNetPersonalSchedUpdates As ADODB.Recordset, _
                    ByVal v_rsCascadeConflicts As ADODB.Recordset, _
                    ByVal v_rsExceptionSPA As ADODB.Recordset, _
                    ByVal v_rsAvailScheduleSPA As ADODB.Recordset, _
                    ByVal v_rsEmpSchedDates As ADODB.Recordset, _
                    ByRef r_rsErrors As ADODB.Recordset, _
                    ByVal v_lngUserId As Long) _
                As Boolean
' Do not expect any lock contention between cascade and personal schedule updates, so
' request cascade to apply the updates instead of returning them, mblnAPPLYCASCADE.
    Dim l_objCtx As ObjectContext
    Dim l_blnError As Boolean
    Dim l_cbsCascade As busCascade.cbsCascade
    Dim l_varCascadeUpdates As Variant
    Dim l_rsCascadeErrors As ADODB.Recordset

On Error GoTo eh
    SetUser v_lngUserId
    
    Set l_objCtx = GetObjectContext
    If r_rsErrors Is Nothing Then
        Set r_rsErrors = CreateNewBusinessErrorSet
    End If
    
    If Not ProcessScheduleChange(v_rsEmpSchedDates, r_rsErrors, v_rsNetPersonalSchedUpdates) Then
        l_objCtx.SetAbort
        ResumeRecordPersonal = False
        GoTo xit
    End If
           
    Set l_cbsCascade = CreateObject("busCascade.cbsCascade")
    l_cbsCascade.UpdateExceptions _
        v_rsCascadeConflicts, v_rsExceptionSPA, v_rsAvailScheduleSPA, _
        mblnAPPLYCASCADE, l_varCascadeUpdates, l_rsCascadeErrors, m_lngUserId
    Set l_cbsCascade = Nothing
    
    AppendCascadeErrors l_rsCascadeErrors, r_rsErrors
    
    Set l_rsCascadeErrors = Nothing
    
    If Not SavePersonalSched(v_rsNetPersonalSchedUpdates) Then
        l_objCtx.SetAbort
        ResumeRecordPersonal = False
        GoTo xit
    End If
    
    ResumeRecordPersonal = True
    l_objCtx.SetComplete
    
xit:
    Set l_objCtx = Nothing
    Exit Function
    
eh:
    l_objCtx.SetAbort
    Set l_objCtx = Nothing
    Set l_rsCascadeErrors = Nothing
    Set l_cbsCascade = Nothing
    Set r_rsErrors = Nothing
'    Clear the cascade error arguments since the client does not format them.
    If Err.Number = trErrAppCascadeCommitConflict Then
        Err.Source = Join(Array(CStr(TrType_PER), "", "", "", ""), ";")
    End If
    TrsRaiseError mstrCMODULENAME, "ResumeRecordPersonal", , , , , True
End Function

Private Function ExtractEmployeeNumbers( _
                    ByRef r_rsEmpSchedDates As ADODB.Recordset) As Variant
' Answer an array of employee numbers; i.e. longs.
    Dim l_cuSet As utilTRISCommonLib.cuSet

On Error GoTo eh
    If RecordsExist(r_rsEmpSchedDates) Then
        r_rsEmpSchedDates.MoveFirst
        Set l_cuSet = New cuSet
        Do Until r_rsEmpSchedDates.EOF
            l_cuSet.Add r_rsEmpSchedDates.Fields("EMP_NUM").Value
            r_rsEmpSchedDates.MoveNext
        Loop
        ExtractEmployeeNumbers = ArrayFromCollection(l_cuSet.Contents)
        r_rsEmpSchedDates.MoveFirst
        Set l_cuSet = Nothing
    Else
        ExtractEmployeeNumbers = Array()
    End If
    
    Exit Function

eh:
    Set l_cuSet = Nothing
    TrsRaiseError mstrCMODULENAME, "ExtractEmployeeNumbers"
End Function


Private Function AppendCascadeErrors( _
                    ByRef r_rsCascadeErrors As ADODB.Recordset, _
                    ByRef r_rsErrors As ADODB.Recordset)
' Append cascade errors/information to the client business error set.
' The client cannot navigate cascade errors so just ensure
' compatible datatypes are populated in the key fields.
On Error GoTo eh
    If Not RecordsExist(r_rsCascadeErrors) Then
        Exit Function
    Else
        r_rsCascadeErrors.MoveFirst
    End If

    Do Until r_rsCascadeErrors.EOF
        AppendBusinessError _
            r_rsErrors, _
            r_rsCascadeErrors!ERR_NUMBER, _
            r_rsCascadeErrors!ERR_TYPE, _
            r_rsCascadeErrors!ERR_DESCRIPTION, _
            r_rsCascadeErrors!ERR_SOURCE, _
            TrType_PER, , , , _
            r_rsCascadeErrors!KEY_FIELD1, _
            r_rsCascadeErrors!KEY_FIELD2, _
            r_rsCascadeErrors!KEY_FIELD2
        r_rsCascadeErrors.MoveNext
    Loop
    Exit Function

eh:
    TrsRaiseError mstrCMODULENAME, "AppendCascadeErrors"
End Function


Private Function ProcessScheduleChange( _
                    ByRef r_rsEmpSchedDates As ADODB.Recordset, _
                    ByRef r_rsErrors As ADODB.Recordset, _
                    ByRef r_rsPersonalSchedUpdates As ADODB.Recordset) _
                As Boolean
    Dim l_cdCascade As dbCascade.cdCascade
    Dim l_varEmps As Variant

On Error GoTo eh
    
    If r_rsEmpSchedDates Is Nothing Then
        ProcessScheduleChange = True
        Exit Function
    End If
    
    l_varEmps = ExtractEmployeeNumbers(r_rsEmpSchedDates)
    
    ' Perform schedule change payment edit before updating the employee schedule
    ' modified timestamp(TTR_EMP_MOD).  Otherwise, cbsSchedule may attempt to read TTR_EMP_MOD
    ' on a different, unenlisted connection and wait for lock to be released, eventually
    ' timing out.  cbsSchedule will only perform GAB call when schedule change payments exist.
    ScheduleChangePaymentExists r_rsEmpSchedDates, r_rsErrors, r_rsPersonalSchedUpdates
    
    ' update the emp_mod timestamp
    ' this will create a lock that time collection must wait on before saving actual time.
    Set l_cdCascade = CreateObject("dbCascade.cdCascade")
    l_cdCascade.SetEmployeeScheduleModified l_varEmps
    Set l_cdCascade = Nothing
    
    If ActualTimeExists(r_rsEmpSchedDates, r_rsErrors) Then
        ProcessScheduleChange = False
        Exit Function
    End If

    If TradeExists(r_rsEmpSchedDates, r_rsErrors) Then
        ProcessScheduleChange = False
        Exit Function
    End If
    

    ProcessScheduleChange = True
    Exit Function
    
eh:
    Set l_cdCascade = Nothing
    Set r_rsErrors = Nothing
    TrsRaiseError mstrCMODULENAME, "ProcessScheduleChange"
End Function


Private Function SetUser(ByVal v_lngUserId As Long)

On Error GoTo eh

    If Val(v_lngUserId) > 0 Then             ' is the user id valid
        If IsValidEmployee(v_lngUserId) Then
            m_lngUserId = v_lngUserId
            m_lngCreatedUserNo = v_lngUserId
        Else
            Err.Raise trErrboEmpNotFound
        End If
    Else
        If Val(v_lngUserId) = m_lngEMPNOCONVERSIONMINUSONE Then 'processing conversion only
           m_lngUserId = m_lngEMPNOCONVERSIONID 'use default user id for last mod emp num
           m_lngCreatedUserNo = v_lngUserId
        Else
            Err.Raise trErrdbInvalidZerolngEmpNum
        End If
    End If
    Exit Function

eh:
    TrsRaiseError mstrCMODULENAME, "SetUser"
End Function


Private Function ProcessCascade( _
                    ByVal v_blnConfirmConflicts As Boolean, _
                    ByRef r_blnConfirmPending As Boolean, _
                    ByRef r_varCascadeUpdates As Variant, _
                    ByRef r_rsNetPersonalSchedUpdates As ADODB.Recordset, _
                    ByRef r_rsCascadeConflicts As ADODB.Recordset, _
                    ByRef r_rsExceptionSPA As ADODB.Recordset, _
                    ByRef r_rsAvailScheduleSPA As ADODB.Recordset, _
                    ByRef r_rsEmpSchedDates As ADODB.Recordset, _
                    ByRef r_rsErrors As ADODB.Recordset) _
                As Boolean

    Dim l_rsCascadeConflicts As ADODB.Recordset
    Dim l_rsExceptionSPA As ADODB.Recordset
    Dim l_rsAvailScheduleSPA As ADODB.Recordset
    Dim l_cbsCascade As busCascade.cbsCascade
    Dim l_varCascadeUpdates As Variant
    Dim l_rsCascadeErrors As ADODB.Recordset


On Error GoTo eh

    r_varCascadeUpdates = Empty
    r_blnConfirmPending = False
    
    Set r_rsEmpSchedDates = GetEmpSchedDates()

    If r_rsEmpSchedDates Is Nothing Then
        ProcessCascade = True
        Exit Function
    End If

    Set l_cbsCascade = CreateObject("busCascade.cbsCascade")
    Set l_rsCascadeConflicts = l_cbsCascade.GetExceptionConflicts( _
                r_rsEmpSchedDates, trNonDemandLevel, _
                l_rsExceptionSPA, l_rsAvailScheduleSPA, , , , , , r_rsNetPersonalSchedUpdates)

    If RecordsExist(l_rsCascadeConflicts) Then
        If v_blnConfirmConflicts Then
            Set r_rsCascadeConflicts = l_rsCascadeConflicts
            Set r_rsExceptionSPA = l_rsExceptionSPA
            Set r_rsAvailScheduleSPA = l_rsAvailScheduleSPA
'            In case inprocess with client
            ClearAllFilters r_rsNetPersonalSchedUpdates, _
                    r_rsCascadeConflicts, r_rsExceptionSPA, r_rsAvailScheduleSPA
            r_blnConfirmPending = True
        Else
            If l_cbsCascade Is Nothing Then
                Set l_cbsCascade = CreateObject("busCascade.cbsCascade")
            End If
            l_cbsCascade.UpdateExceptions _
                l_rsCascadeConflicts, l_rsExceptionSPA, l_rsAvailScheduleSPA, _
                mblnAPPLYCASCADE, l_varCascadeUpdates, l_rsCascadeErrors, m_lngUserId
            AppendCascadeErrors l_rsCascadeErrors, r_rsErrors
            r_varCascadeUpdates = l_varCascadeUpdates
        End If
    End If
    
    Set l_cbsCascade = Nothing
    Set l_rsCascadeConflicts = Nothing
    Set l_rsExceptionSPA = Nothing
    Set l_rsAvailScheduleSPA = Nothing
    Set l_rsCascadeErrors = Nothing
    
    ProcessCascade = True
    
    Exit Function

eh:
    Set l_rsCascadeConflicts = Nothing
    Set l_rsExceptionSPA = Nothing
    Set l_rsAvailScheduleSPA = Nothing
    Set l_cbsCascade = Nothing
    Set l_rsCascadeErrors = Nothing
    TrsRaiseError mstrCMODULENAME, "ProcessCascade"
End Function


Private Function ActualTimeExists( _
                    ByRef r_rsEmpSchedDates As ADODB.Recordset, _
                    ByRef r_rsErrors As ADODB.Recordset) As Boolean
    Dim l_blnExists As Boolean
    Dim l_cdTimecard As dbTimeCard.cdTimeCard
    Dim l_rsActualTime As ADODB.Recordset
    Dim l_varEmployees As Variant
    Dim l_datStart As Date
    Dim l_datEnd As Date
    Dim l_strFilter As String
    Dim l_strErrMsg As String

On Error GoTo eh

    l_blnExists = False
    DetermineGetActualTimeScope r_rsEmpSchedDates, l_varEmployees, l_datStart, l_datEnd
    If IsEmpty(l_varEmployees) Then
        ActualTimeExists = False
        Exit Function
    End If

    Set l_cdTimecard = CreateObject("dbTimeCard.cdTimeCard")
    Set l_rsActualTime = _
            l_cdTimecard.GetActualTimeCascadeExists(l_varEmployees, l_datStart, l_datEnd)
    Set l_cdTimecard = Nothing

    If l_rsActualTime.EOF Then
        ActualTimeExists = False
        Set l_rsActualTime = Nothing
        Exit Function
    End If

    Do Until r_rsEmpSchedDates.EOF
        l_strFilter = _
            "EMP_NUM = " & r_rsEmpSchedDates!EMP_NUM & _
            " AND CHARGE_DATE >= #" & Format(r_rsEmpSchedDates!START_DATE, "mm/dd/yyyy") & "#" & _
            " AND CHARGE_DATE <= #" & Format(r_rsEmpSchedDates!END_DATE, "mm/dd/yyyy") & "#"
        l_strFilter = MDAC25FilterFix(l_strFilter)
        l_rsActualTime.Filter = l_strFilter
        If RecordsExist(l_rsActualTime) Then
            l_strErrMsg = _
                "Employee " & CStr(l_rsActualTime!EMP_NUM) & _
                "[" & Trim(l_rsActualTime!EMP_NAME) & "] has existing actual time on " & _
                Format(l_rsActualTime!CHARGE_DATE, "mm/dd/yyyy") & "."
            AppendBusinessError _
                r_rsErrors, _
                trErrbusActualTimeExists, _
                mstrCBUSINESSERR, _
                l_strErrMsg, _
                mstrCMODULENAME, _
                TrType_PER, _
                , _
                , _
                mstrCVARFIELDNAMES, _
                r_rsEmpSchedDates!EMP_NUM, _
                r_rsEmpSchedDates!START_DATE, _
                r_rsEmpSchedDates!END_DATE
            l_blnExists = True
        End If
        r_rsEmpSchedDates.MoveNext
    Loop
    r_rsEmpSchedDates.MoveFirst

    ActualTimeExists = l_blnExists
    Set l_rsActualTime = Nothing

    Exit Function

eh:
    Set l_cdTimecard = Nothing
    Set l_rsActualTime = Nothing
    TrsRaiseError mstrCMODULENAME, "ActualTimeExists"
End Function


Private Sub DetermineGetActualTimeScope( _
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
    TrsRaiseError mstrCMODULENAME, "DetermineGetActualTimeScope"
End Sub


Private Function TradeExists( _
                    ByRef r_rsEmpSchedDates As ADODB.Recordset, _
                    ByRef r_rsErrors As ADODB.Recordset) As Boolean
    Dim l_blnExists As Boolean
    Dim l_cdSchedule As dbSchedule.cdSchedule
    Dim l_rsEmpTrades As ADODB.Recordset
    Dim l_varEmployees As Variant
    Dim l_datStart As Date
    Dim l_datEnd As Date
    Dim l_strFilter As String
    Dim l_strErrMsg As String
                    
On Error GoTo eh
                    
    l_blnExists = False
    DetermineGetActualTimeScope r_rsEmpSchedDates, l_varEmployees, l_datStart, l_datEnd
    If IsEmpty(l_varEmployees) Then
        TradeExists = False
        Exit Function
    End If
                    
    Set l_cdSchedule = CreateObject("dbSchedule.cdSchedule")
    
    Set l_rsEmpTrades = _
        l_cdSchedule.GetEmployeeTrades(l_varEmployees, l_datStart, l_datEnd)
    Set l_cdSchedule = Nothing

    If l_rsEmpTrades.EOF Then
        TradeExists = False
        Set l_rsEmpTrades = Nothing
        Exit Function
    End If

    Do Until r_rsEmpSchedDates.EOF
        l_strFilter = _
            "EMP_NUM = " & r_rsEmpSchedDates!EMP_NUM & _
            " AND CHARGE_DATE >= #" & Format(r_rsEmpSchedDates!START_DATE, "mm/dd/yyyy") & "#" & _
            " AND CHARGE_DATE <= #" & Format(r_rsEmpSchedDates!END_DATE, "mm/dd/yyyy") & "#"
        l_strFilter = MDAC25FilterFix(l_strFilter)
        l_rsEmpTrades.Filter = l_strFilter
        If RecordsExist(l_rsEmpTrades) Then
            l_strErrMsg = _
                "Employee " & CStr(l_rsEmpTrades!EMP_NUM) & _
                "[" & Trim(l_rsEmpTrades!EMP_NAME) & "] has existing trade on " & _
                Format(l_rsEmpTrades!CHARGE_DATE, "mm/dd/yyyy") & "."
            AppendBusinessError _
                r_rsErrors, _
                trErrbusTradeExists, _
                mstrCBUSINESSERR, _
                l_strErrMsg, _
                mstrCMODULENAME, _
                TrType_PER, _
                , _
                , _
                mstrCVARFIELDNAMES, _
                r_rsEmpSchedDates!EMP_NUM, _
                r_rsEmpSchedDates!START_DATE, _
                r_rsEmpSchedDates!END_DATE
            l_blnExists = True
        End If
        r_rsEmpSchedDates.MoveNext
    Loop
    r_rsEmpSchedDates.MoveFirst

    TradeExists = l_blnExists
    Set l_rsEmpTrades = Nothing

    Exit Function

eh:
    Set l_cdSchedule = Nothing
    Set l_rsEmpTrades = Nothing
    TrsRaiseError mstrCMODULENAME, "TradeExists"
End Function


Private Function ScheduleChangePaymentExists( _
                    ByRef r_rsEmpSchedDates As ADODB.Recordset, _
                    ByRef r_rsErrors As ADODB.Recordset, _
                    ByRef r_rsPersonalSchedUpdates As ADODB.Recordset) As Boolean
' Check for any schedule change payments during the scope of the personal schedule change.
' If any found, determine whether the payment is scheduled on an off-day.  Must call
' GAB with the pending changes to determine the schedule based on the unsaved personal
' schedule changes.  If a payment is found on an off-day, just warn the user so they can take
' appropriate action, e.g. move the payment to the next scheduled work day.
' Answer TRUE if a payment exists for an off-day, otherwise FALSE.
    Dim l_blnExists As Boolean
    Dim l_dbSchedChgPay As dbSchedChgPay.cdSchedChgPay
    Dim l_rsSchedChgPay As ADODB.Recordset
    Dim l_varEmployees As Variant
    Dim l_datStart As Date
    Dim l_datEnd As Date
    Dim l_cbsSchedule As busSchedule.cbsSchedule
    Dim l_rsSchedule As ADODB.Recordset
    Dim l_strErrorMsg As String
    Dim l_colErrArgs As Collection
    Dim l_rsSchedChgPayDates As ADODB.Recordset
                    
On Error GoTo eh
                    
    l_blnExists = False
    DetermineGetActualTimeScope r_rsEmpSchedDates, l_varEmployees, l_datStart, l_datEnd
    
    If IsEmpty(l_varEmployees) Then
        ScheduleChangePaymentExists = False
        Exit Function
    End If
                          
    Set l_dbSchedChgPay = CreateObject("dbSchedChgPay.cdSchedChgPay")
    Set l_rsSchedChgPay = l_dbSchedChgPay.GetSchedChgPay(l_varEmployees, l_datStart, l_datEnd)
    Set l_dbSchedChgPay = Nothing
    
    If Not RecordsExist(l_rsSchedChgPay) Then
        ScheduleChangePaymentExists = False
        Exit Function
    End If

' Reduce scope of GAB by only requesting the dates schedule change payments exist.
    Set l_rsSchedChgPayDates = GetSchedChgPayDates(l_rsSchedChgPay)

    Set l_cbsSchedule = CreateObject("busSchedule.cbsSchedule")
    Set l_rsSchedule = l_cbsSchedule.GetPendingAdjustedBaseSchedule( _
            l_rsSchedChgPayDates, Nothing, False, trBaseSchedNoDetails, trAllScheduleLevels, _
            Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, _
            r_rsPersonalSchedUpdates)
    Set l_cbsSchedule = Nothing

    Do Until l_rsSchedChgPay.EOF
        If IsPaymentScheduledOnOffDay(l_rsSchedChgPay, l_rsSchedule) Then
            l_blnExists = True
            Set l_colErrArgs = New Collection
            l_colErrArgs.Add CStr(l_rsSchedChgPay!EMP_NUM)
            l_colErrArgs.Add Trim(l_rsSchedChgPay!EMP_NAME)
            l_colErrArgs.Add Format(l_rsSchedChgPay!CHARGE_DATE, "mm/dd/yyyy")
            l_strErrorMsg = TrsGetErrDescription(LongAppErr(trWarSchedChgPmtOffday), l_colErrArgs)
            Set l_colErrArgs = Nothing
            AppendBusinessError _
                r_rsErrors, _
                LongAppErr(trWarSchedChgPmtOffday), _
                mstrCBUSINESSWAR, _
                l_strErrorMsg, _
                mstrCMODULENAME, _
                TrType_PER, _
                , _
                , _
                mstrCVARFIELDNAMES, _
                l_rsSchedChgPay!EMP_NUM, _
                l_rsSchedChgPay!CHARGE_DATE, _
                l_rsSchedChgPay!CHARGE_DATE
        End If
        l_rsSchedChgPay.MoveNext
    Loop

    ScheduleChangePaymentExists = l_blnExists
    Set l_rsSchedChgPay = Nothing
    Set l_rsSchedule = Nothing
    Set l_rsSchedChgPayDates = Nothing
    Exit Function
    
eh:
    Set l_rsSchedChgPayDates = Nothing
    Set l_colErrArgs = Nothing
    Set l_dbSchedChgPay = Nothing
    Set l_rsSchedChgPay = Nothing
    Set l_rsSchedule = Nothing
    Set l_cbsSchedule = Nothing
    TrsRaiseError mstrCMODULENAME, "ScheduleChangePaymentExists"
End Function


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
    TrsRaiseError mstrCMODULENAME, "GetSchedChgPayDates"
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
    TrsRaiseError mstrCMODULENAME, "IsPaymentScheduledOnOffDay"
End Function


'perform insert, change and delete using net recordset
Private Function SavePersonalSched( _
                    ByRef r_rsPersonalSchedRequest As ADODB.Recordset) _
                As Boolean

On Error GoTo SavePersonalSched_EH
    
    Dim l_dbSchedPersonal As dbSchedPersonal.cdSchedPersonal
    SavePersonalSched = True
    
    If IsClearedFilterRecordsetEmpty(r_rsPersonalSchedRequest) Then
        Exit Function
    Else
    End If
    
    Set l_dbSchedPersonal = CreateObject("dbSchedPersonal.cdSchedPersonal")

    If l_dbSchedPersonal.SavePersonalSchedule(r_rsPersonalSchedRequest) Then
        SavePersonalSched = True
    Else
        SavePersonalSched = False
    End If

    Set l_dbSchedPersonal = Nothing

Exit Function
SavePersonalSched_EH:

    Set l_dbSchedPersonal = Nothing
    
    TrsRaiseError mstrCMODULENAME, "SavePersonalSched", , , , , True

End Function


'check for valid schedule codes
Private Function IsValidSchedCd(ByRef r_rsPersonalRequest As ADODB.Recordset) As Boolean

On Error GoTo IsValidSchedCd_EH

    Dim rs As ADODB.Recordset

    If CodeTableDataManager.ValidateScheduleCode(Trim(r_rsPersonalRequest("SCHED_CD"))) Then
       IsValidSchedCd = True
    Else
       IsValidSchedCd = False
    End If

    Set rs = Nothing

Exit Function

IsValidSchedCd_EH:

    Set rs = Nothing
    TrsRaiseError mstrCMODULENAME, "IsValidSchedCd"

End Function


'create reference to code tables
Private Function CodeTableDataManager() As busCodeTables.cbsCodeTables

On Error GoTo eh

    Set CodeTableDataManager = _
        CreateObject("busCodeTables.cbsCodeTables")

Exit Function

eh:
    TrsRaiseError mstrCMODULENAME, "CodeTableDataManager"
    
End Function

'validate preliminary BRV for all updatable input fields
Private Function ValidatePersonalSchedEntry( _
                    ByRef r_rsPersonalSchedRequest As ADODB.Recordset, _
                    ByRef r_rsErrors As ADODB.Recordset _
                    ) As Boolean
    
On Error GoTo eh
    
    Dim l_blnValidationError As Boolean
    Dim l_lngEmpNum As Long
    Dim l_datStartDateAdj As Date
    Dim l_strWarDesc As String
    
    If IsClearedFilterRecordsetEmpty(r_rsPersonalSchedRequest) Then
        ValidatePersonalSchedEntry = True
        Exit Function
    Else
    End If
    
    Do Until r_rsPersonalSchedRequest.EOF
        If Not IsNull(r_rsPersonalSchedRequest!RECORD_STATUS) Then
            If Not IsNull(r_rsPersonalSchedRequest!RECORD_ID) Then
                m_lngRecordId = Val(r_rsPersonalSchedRequest!RECORD_ID)
            Else
                m_lngRecordId = 0
            End If
            If Trim(r_rsPersonalSchedRequest!RECORD_STATUS) = "D" Then 'VALIDATE DELETE
                If Not IsNull(r_rsPersonalSchedRequest!KEY_EMP_NUM) Then
                    If Val(r_rsPersonalSchedRequest!KEY_EMP_NUM) > 0 And IsNumeric(r_rsPersonalSchedRequest!KEY_EMP_NUM) Then
                    Else
                        'trErrboEmpNotFound Invalid Employee Number
                        m_lngErrorNo = trErrboEmpNotFound
                        AppendPersonalError r_rsErrors, TrType_PER, m_lngRecordId, "EMP_NUM", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), r_rsPersonalSchedRequest!KEY_EMP_NUM, r_rsPersonalSchedRequest!KEY_START_DATE, r_rsPersonalSchedRequest!KEY_END_DATE
                        l_blnValidationError = True
                    End If
                Else
                    'trErrboEmpNotFound Invalid Employee Number
                    m_lngErrorNo = trErrboEmpNotFound
                    AppendPersonalError r_rsErrors, TrType_PER, m_lngRecordId, "EMP_NUM", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), r_rsPersonalSchedRequest!KEY_EMP_NUM, r_rsPersonalSchedRequest!KEY_START_DATE, r_rsPersonalSchedRequest!KEY_END_DATE
                    l_blnValidationError = True
                End If
                
                If Not IsNull(r_rsPersonalSchedRequest!KEY_START_DATE) Then
                    If IsDate(r_rsPersonalSchedRequest!KEY_START_DATE) Then
                    Else
                        'trErrbusSchedStartDateInvalid Invalid Schedule Date
                        m_lngErrorNo = trErrbusSchedStartDateInvalid
                        AppendPersonalError r_rsErrors, TrType_PER, m_lngRecordId, "START_DATE", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), r_rsPersonalSchedRequest!KEY_EMP_NUM, r_rsPersonalSchedRequest!KEY_START_DATE, r_rsPersonalSchedRequest!KEY_END_DATE
                        l_blnValidationError = True
                    End If
                Else
                    'trErrbusSchedStartDateInvalid Invalid Schedule Date
                    m_lngErrorNo = trErrbusSchedStartDateInvalid
                    AppendPersonalError r_rsErrors, TrType_PER, m_lngRecordId, "START_DATE", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), r_rsPersonalSchedRequest!KEY_EMP_NUM, r_rsPersonalSchedRequest!KEY_START_DATE, r_rsPersonalSchedRequest!KEY_END_DATE
                    l_blnValidationError = True
                End If
                
                If Not IsNull(r_rsPersonalSchedRequest!KEY_END_DATE) Then
                    If IsDate(r_rsPersonalSchedRequest!KEY_END_DATE) Then
                    Else
                        'trErrbusSchedEndDateInvalid Invalid Schedule End Date
                        m_lngErrorNo = trErrbusSchedEndDateInvalid
                        AppendPersonalError r_rsErrors, TrType_PER, m_lngRecordId, "END_DATE", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), r_rsPersonalSchedRequest!KEY_EMP_NUM, r_rsPersonalSchedRequest!KEY_START_DATE, r_rsPersonalSchedRequest!KEY_END_DATE
                        l_blnValidationError = True
                    End If
                Else
                    'trErrbusSchedEndDateInvalid Invalid Schedule End Date
                    m_lngErrorNo = trErrbusSchedEndDateInvalid
                    AppendPersonalError r_rsErrors, TrType_PER, m_lngRecordId, "END_DATE", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), r_rsPersonalSchedRequest!KEY_EMP_NUM, r_rsPersonalSchedRequest!KEY_START_DATE, r_rsPersonalSchedRequest!KEY_END_DATE
                    l_blnValidationError = True
                End If
            Else
                If Trim(r_rsPersonalSchedRequest!RECORD_STATUS) = "C" Then  'VALIDATE CHANGE
                    If Not IsNull(r_rsPersonalSchedRequest!KEY_EMP_NUM) Then
                        If Val(r_rsPersonalSchedRequest!KEY_EMP_NUM) > 0 And IsNumeric(r_rsPersonalSchedRequest!KEY_EMP_NUM) Then
                        Else
                            'trErrboEmpNotFound Invalid  Employee Number
                            m_lngErrorNo = trErrboEmpNotFound
                            AppendPersonalError r_rsErrors, TrType_PER, m_lngRecordId, "EMP_NUM", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), r_rsPersonalSchedRequest!EMP_NUM, r_rsPersonalSchedRequest!START_DATE, r_rsPersonalSchedRequest!END_DATE
                            l_blnValidationError = True
                        End If
                    Else
                        'trErrboEmpNotFound Invalid  Employee Number
                        m_lngErrorNo = trErrboEmpNotFound
                        AppendPersonalError r_rsErrors, TrType_PER, m_lngRecordId, "EMP_NUM", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), r_rsPersonalSchedRequest!EMP_NUM, r_rsPersonalSchedRequest!START_DATE, r_rsPersonalSchedRequest!END_DATE
                        l_blnValidationError = True
                    End If
                    If r_rsPersonalSchedRequest!KEY_EMP_NUM = r_rsPersonalSchedRequest!EMP_NUM Then
                    Else
                        'trErrbusSchedPersCannotChangeEmpNum Employer Number cannot be changed on a personal schedule.
                        m_lngErrorNo = trErrbusSchedPersCannotChangeEmpNum
                        AppendPersonalError r_rsErrors, TrType_PER, m_lngRecordId, "EMP_NUM", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), r_rsPersonalSchedRequest!EMP_NUM, r_rsPersonalSchedRequest!START_DATE, r_rsPersonalSchedRequest!END_DATE
                        l_blnValidationError = True
                    End If
                    
                    If Not IsNull(r_rsPersonalSchedRequest!KEY_START_DATE) Then
                        If IsDate(r_rsPersonalSchedRequest!KEY_START_DATE) Then
                        Else
                            'trErrbusSchedStartDateInvalid Invalid Schedule Date
                            m_lngErrorNo = trErrbusSchedStartDateInvalid
                            AppendPersonalError r_rsErrors, TrType_PER, m_lngRecordId, "START_DATE", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), r_rsPersonalSchedRequest!EMP_NUM, r_rsPersonalSchedRequest!START_DATE, r_rsPersonalSchedRequest!END_DATE
                            l_blnValidationError = True
                        End If
                    Else
                        'trErrbusSchedStartDateInvalid Invalid Schedule Date
                        m_lngErrorNo = trErrbusSchedStartDateInvalid
                        AppendPersonalError r_rsErrors, TrType_PER, m_lngRecordId, "START_DATE", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), r_rsPersonalSchedRequest!EMP_NUM, r_rsPersonalSchedRequest!START_DATE, r_rsPersonalSchedRequest!END_DATE
                        l_blnValidationError = True
                    End If
                    
                    If Not IsNull(r_rsPersonalSchedRequest!KEY_END_DATE) Then
                        If IsDate(r_rsPersonalSchedRequest!KEY_END_DATE) Then
                        Else
                            'trErrbusSchedEndDateInvalid Invalid Schedule End Date
                            m_lngErrorNo = trErrbusSchedEndDateInvalid
                            AppendPersonalError r_rsErrors, TrType_PER, m_lngRecordId, "END_DATE", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), r_rsPersonalSchedRequest!EMP_NUM, r_rsPersonalSchedRequest!START_DATE, r_rsPersonalSchedRequest!END_DATE
                            l_blnValidationError = True
                        End If
                    Else
                        'trErrbusSchedEndDateInvalid Invalid Schedule End Date
                        m_lngErrorNo = trErrbusSchedEndDateInvalid
                        AppendPersonalError r_rsErrors, TrType_PER, m_lngRecordId, "END_DATE", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), r_rsPersonalSchedRequest!EMP_NUM, r_rsPersonalSchedRequest!START_DATE, r_rsPersonalSchedRequest!END_DATE
                        l_blnValidationError = True
                    End If
                    
                    If Not IsNull(r_rsPersonalSchedRequest!EMP_NUM) Then
                        If Val(r_rsPersonalSchedRequest!EMP_NUM) > 0 And IsNumeric(r_rsPersonalSchedRequest!EMP_NUM) Then
                        Else
                            'trErrboEmpNotFound Invalid  Employee Number
                            m_lngErrorNo = trErrboEmpNotFound
                            AppendPersonalError r_rsErrors, TrType_PER, m_lngRecordId, "EMP_NUM", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), r_rsPersonalSchedRequest!EMP_NUM, r_rsPersonalSchedRequest!START_DATE, r_rsPersonalSchedRequest!END_DATE
                            l_blnValidationError = True
                        End If
                    Else
                        'trErrboEmpNotFound Invalid  Employee Number
                        m_lngErrorNo = trErrboEmpNotFound
                        AppendPersonalError r_rsErrors, TrType_PER, m_lngRecordId, "EMP_NUM", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), r_rsPersonalSchedRequest!EMP_NUM, r_rsPersonalSchedRequest!START_DATE, r_rsPersonalSchedRequest!END_DATE
                        l_blnValidationError = True
                    End If
                    
                    If Not IsNull(r_rsPersonalSchedRequest!START_DATE) Then
                        If IsDate(r_rsPersonalSchedRequest!START_DATE) Then
                        Else
                            'trErrbusSchedStartDateInvalid Invalid Schedule Date
                            m_lngErrorNo = trErrbusSchedStartDateInvalid
                            AppendPersonalError r_rsErrors, TrType_PER, m_lngRecordId, "START_DATE", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), r_rsPersonalSchedRequest!EMP_NUM, r_rsPersonalSchedRequest!START_DATE, r_rsPersonalSchedRequest!END_DATE
                            l_blnValidationError = True
                        End If
                    Else
                        'trErrbusSchedStartDateInvalid Invalid Schedule Date
                        m_lngErrorNo = trErrbusSchedStartDateInvalid
                        AppendPersonalError r_rsErrors, TrType_PER, m_lngRecordId, "START_DATE", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), r_rsPersonalSchedRequest!EMP_NUM, r_rsPersonalSchedRequest!START_DATE, r_rsPersonalSchedRequest!END_DATE
                        l_blnValidationError = True
                    End If
                    
                    If Not IsNull(r_rsPersonalSchedRequest!END_DATE) Then
                        If IsDate(r_rsPersonalSchedRequest!END_DATE) Then
                        Else
                            'trErrbusSchedEndDateInvalid Invalid Schedule Schedule End Date
                            m_lngErrorNo = trErrbusSchedEndDateInvalid
                            AppendPersonalError r_rsErrors, TrType_PER, m_lngRecordId, "END_DATE", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), r_rsPersonalSchedRequest!EMP_NUM, r_rsPersonalSchedRequest!START_DATE, r_rsPersonalSchedRequest!END_DATE
                            l_blnValidationError = True
                        End If
                    Else
                        'trErrbusSchedEndDateInvalid Invalid Schedule Schedule End Date
'                        m_lngErrorNo = trErrbusSchedEndDateInvalid
'                        AppendPersonalError r_rsErrors, TrType_PER, m_lngRecordId, "END_DATE", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), r_rsPersonalSchedRequest!EMP_NUM, r_rsPersonalSchedRequest!START_DATE, r_rsPersonalSchedRequest!END_DATE
'                        l_blnValidationError = True
'                         MsgBox "no end date in chage"
                    End If
                   
                    If Not IsNull(r_rsPersonalSchedRequest!SCHED_CD) Then
                        If Val(r_rsPersonalSchedRequest!SCHED_CD) = 0 And IsNumeric(r_rsPersonalSchedRequest!SCHED_CD) Then
                            If Not IsNull(r_rsPersonalSchedRequest!ROT_CD) Then
                                If Val(r_rsPersonalSchedRequest!ROT_CD) > 0 And IsNumeric(r_rsPersonalSchedRequest!ROT_CD) Then
                                    If Not IsNull(r_rsPersonalSchedRequest!ROT_WEEK) Then
                                        If Val(r_rsPersonalSchedRequest!ROT_WEEK) > 0 And IsNumeric(r_rsPersonalSchedRequest!ROT_WEEK) Then
                                        Else
                                            'trErrbusSchedStartPosIsZero Invalid Rotation Start Position
                                            m_lngErrorNo = trErrbusSchedStartPosIsZero
                                            AppendPersonalError r_rsErrors, TrType_PER, m_lngRecordId, "ROT_WEEK", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), r_rsPersonalSchedRequest!EMP_NUM, r_rsPersonalSchedRequest!START_DATE, r_rsPersonalSchedRequest!END_DATE
                                            l_blnValidationError = True
                                        End If
                                    Else
                                        'trErrbusSchedStartPosNull Invalid Rotation Statr Position
                                        m_lngErrorNo = trErrbusSchedStartPosNull
                                        AppendPersonalError r_rsErrors, TrType_PER, m_lngRecordId, "ROT_WEEK", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), r_rsPersonalSchedRequest!EMP_NUM, r_rsPersonalSchedRequest!START_DATE, r_rsPersonalSchedRequest!END_DATE
                                        l_blnValidationError = True
                                    End If
                                Else
                                    'trErrbusSchedRotCodeIsZero Invalid Rotation Code
                                    m_lngErrorNo = trErrbusSchedRotCodeIsZero
                                    AppendPersonalError r_rsErrors, TrType_PER, m_lngRecordId, "ROT_CD", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), r_rsPersonalSchedRequest!EMP_NUM, r_rsPersonalSchedRequest!START_DATE, r_rsPersonalSchedRequest!END_DATE
                                    l_blnValidationError = True
                                End If
                            Else
                                'trErrbusSchedRotCodeNull Invalid Rotation Code
                                m_lngErrorNo = trErrbusSchedRotCodeNull
                                AppendPersonalError r_rsErrors, TrType_PER, m_lngRecordId, "ROT_CD", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), r_rsPersonalSchedRequest!EMP_NUM, r_rsPersonalSchedRequest!START_DATE, r_rsPersonalSchedRequest!END_DATE
                                l_blnValidationError = True
                            End If
                        Else
                            If Not IsNull(r_rsPersonalSchedRequest!ROT_CD) Then
                                If Val(r_rsPersonalSchedRequest!ROT_CD) < 1 And IsNumeric(r_rsPersonalSchedRequest!ROT_CD) Then
                                    If Not IsNull(r_rsPersonalSchedRequest!ROT_WEEK) Then
                                        If Val(r_rsPersonalSchedRequest!ROT_WEEK) < 1 And IsNumeric(r_rsPersonalSchedRequest!ROT_WEEK) Then
                                        Else
                                            'trErrbusSchedPersStartPosNotZero Rotation Start Position must be zero when a valid Schedule Code is entered
                                            m_lngErrorNo = trErrbusSchedPersStartPosNotZero
                                            AppendPersonalError r_rsErrors, TrType_PER, m_lngRecordId, "ROT_WEEK", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), r_rsPersonalSchedRequest!EMP_NUM, r_rsPersonalSchedRequest!START_DATE, r_rsPersonalSchedRequest!END_DATE
                                            l_blnValidationError = True
                                        End If
                                    Else
                                        'trErrbusSchedStartPosNull Invalid Rotation Statr Position
                                        m_lngErrorNo = trErrbusSchedStartPosNull
                                        AppendPersonalError r_rsErrors, TrType_PER, m_lngRecordId, "ROT_WEEK", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), r_rsPersonalSchedRequest!EMP_NUM, r_rsPersonalSchedRequest!START_DATE, r_rsPersonalSchedRequest!END_DATE
                                        l_blnValidationError = True
                                    End If
                                Else
                                    'trErrbusSchedPersRotCdNotZero Rotation Code must be zero when a valid Schedule Code is entered
                                    m_lngErrorNo = trErrbusSchedPersRotCdNotZero
                                    AppendPersonalError r_rsErrors, TrType_PER, m_lngRecordId, "ROT_CD", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), r_rsPersonalSchedRequest!EMP_NUM, r_rsPersonalSchedRequest!START_DATE, r_rsPersonalSchedRequest!END_DATE
                                    l_blnValidationError = True
                                End If
                            Else
                                'trErrbusSchedRotCodeNull Invalid Rotation Code
                                m_lngErrorNo = trErrbusSchedRotCodeNull
                                AppendPersonalError r_rsErrors, TrType_PER, m_lngRecordId, "ROT_CD", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), r_rsPersonalSchedRequest!EMP_NUM, r_rsPersonalSchedRequest!START_DATE, r_rsPersonalSchedRequest!END_DATE
                                l_blnValidationError = True
                            End If
                        End If
                    Else
                        If Not IsNull(r_rsPersonalSchedRequest!ROT_CD) Then
                            If Val(r_rsPersonalSchedRequest!ROT_CD) > 0 And IsNumeric(r_rsPersonalSchedRequest!ROT_CD) Then
                                If Not IsNull(r_rsPersonalSchedRequest!ROT_WEEK) Then
                                    If Val(r_rsPersonalSchedRequest!ROT_WEEK) > 0 And IsNumeric(r_rsPersonalSchedRequest!ROT_WEEK) Then
                                    Else
                                        'trErrbusSchedStartPosIsZero Invalid Rotation Start Position
                                        m_lngErrorNo = trErrbusSchedStartPosIsZero
                                        AppendPersonalError r_rsErrors, TrType_PER, m_lngRecordId, "ROT_WEEK", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), r_rsPersonalSchedRequest!EMP_NUM, r_rsPersonalSchedRequest!START_DATE, r_rsPersonalSchedRequest!END_DATE
                                        l_blnValidationError = True
                                    End If
                                Else
                                    'trErrbusSchedStartPosNull Invalid Rotation Start Position
                                    m_lngErrorNo = trErrbusSchedStartPosNull
                                    AppendPersonalError r_rsErrors, TrType_PER, m_lngRecordId, "ROT_WEEK", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), r_rsPersonalSchedRequest!EMP_NUM, r_rsPersonalSchedRequest!START_DATE, r_rsPersonalSchedRequest!END_DATE
                                    l_blnValidationError = True
                                End If
                            Else
                                'trErrbusSchedRotCodeIsZero Invalid Rotation Code
                                m_lngErrorNo = trErrbusSchedRotCodeIsZero
                                AppendPersonalError r_rsErrors, TrType_PER, m_lngRecordId, "ROT_CD", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), r_rsPersonalSchedRequest!EMP_NUM, r_rsPersonalSchedRequest!START_DATE, r_rsPersonalSchedRequest!END_DATE
                                l_blnValidationError = True
                            End If
                        Else
                            'trErrbusSchedRotCodeNull Invalid Rotation Code
                            m_lngErrorNo = trErrbusSchedRotCodeNull
                            AppendPersonalError r_rsErrors, TrType_PER, m_lngRecordId, "ROT_CD", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), r_rsPersonalSchedRequest!EMP_NUM, r_rsPersonalSchedRequest!START_DATE, r_rsPersonalSchedRequest!END_DATE
                            l_blnValidationError = True
                        End If
                    End If
                Else ' VALIDATE INSERT
                    If Not IsNull(r_rsPersonalSchedRequest!EMP_NUM) Then
                        If Val(r_rsPersonalSchedRequest!EMP_NUM) > 0 And IsNumeric(r_rsPersonalSchedRequest!EMP_NUM) Then
                        Else
                            'trErrboEmpNotFound Invalid Employee Number
                            m_lngErrorNo = trErrboEmpNotFound
                            AppendPersonalError r_rsErrors, TrType_PER, m_lngRecordId, "EMP_NUM", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), r_rsPersonalSchedRequest!EMP_NUM, r_rsPersonalSchedRequest!START_DATE, r_rsPersonalSchedRequest!END_DATE
                            l_blnValidationError = True
                        End If
                    Else
                        'trErrboEmpNotFound Invalid Employee Number
                        m_lngErrorNo = trErrboEmpNotFound
                        AppendPersonalError r_rsErrors, TrType_PER, m_lngRecordId, "EMP_NUM", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), r_rsPersonalSchedRequest!EMP_NUM, r_rsPersonalSchedRequest!START_DATE, r_rsPersonalSchedRequest!END_DATE
                        l_blnValidationError = True
                    End If
                    If Not IsNull(r_rsPersonalSchedRequest!START_DATE) Then
                        If IsDate(r_rsPersonalSchedRequest!START_DATE) Then
                        Else
                            'trErrbusSchedStartDateInvalid Invalid Schedule Date
                            m_lngErrorNo = trErrbusSchedStartDateInvalid
                            AppendPersonalError r_rsErrors, TrType_PER, m_lngRecordId, "START_DATE", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), r_rsPersonalSchedRequest!EMP_NUM, r_rsPersonalSchedRequest!START_DATE, r_rsPersonalSchedRequest!END_DATE
                            l_blnValidationError = True
                        End If
                    Else
                        'trErrbusSchedStartDateInvalid Invalid Schedule Date
                        m_lngErrorNo = trErrbusSchedStartDateInvalid
                        AppendPersonalError r_rsErrors, TrType_PER, m_lngRecordId, "START_DATE", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), r_rsPersonalSchedRequest!EMP_NUM, r_rsPersonalSchedRequest!START_DATE, r_rsPersonalSchedRequest!END_DATE
                        l_blnValidationError = True
                    End If
                    If Not IsNull(r_rsPersonalSchedRequest!END_DATE) Then
                        If IsDate(r_rsPersonalSchedRequest!END_DATE) Then
                        Else
                            'trErrbusSchedEndDateInvalid Invalid Schedule End Date
                            m_lngErrorNo = trErrbusSchedEndDateInvalid
                            AppendPersonalError r_rsErrors, TrType_PER, m_lngRecordId, "END_DATE", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), r_rsPersonalSchedRequest!EMP_NUM, r_rsPersonalSchedRequest!START_DATE, r_rsPersonalSchedRequest!END_DATE
                            l_blnValidationError = True
                        End If
                    Else
                        ' perform assignment of end date when null
                        'trErrbusSchedEndDateInvalid Invalid Schedule End Date
'                        m_lngErrorNo = trErrbusSchedEndDateInvalid
'                        AppendPersonalError r_rsErrors, TrType_PER, m_lngRecordId, "END_DATE", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), r_rsPersonalSchedRequest!EMP_NUM, r_rsPersonalSchedRequest!START_DATE, r_rsPersonalSchedRequest!END_DATE
'                        l_blnValidationError = True
'                        MsgBox "No end date selected"
'                        r_rsPersonalSchedRequest!END_DATE = CDate(GetDBDefaultEndDate(r_rsPersonalSchedRequest!EMP_NUM, r_rsPersonalSchedRequest!START_DATE, r_rsPersonalSchedRequest))
'                        r_rsPersonalSchedRequest.Update
                    End If
                    If Not IsNull(r_rsPersonalSchedRequest!SCHED_CD) Then
                        If Val(r_rsPersonalSchedRequest!SCHED_CD) = 0 And IsNumeric(r_rsPersonalSchedRequest!SCHED_CD) Then
                            If Not IsNull(r_rsPersonalSchedRequest!ROT_CD) Then
                                If Val(r_rsPersonalSchedRequest!ROT_CD) > 0 And IsNumeric(r_rsPersonalSchedRequest!ROT_CD) Then
                                    If Not IsNull(r_rsPersonalSchedRequest!ROT_WEEK) Then
                                        If Val(r_rsPersonalSchedRequest!ROT_WEEK) > 0 And IsNumeric(r_rsPersonalSchedRequest!ROT_WEEK) Then
                                        Else
                                            'trErrbusSchedStartPosIsZero Invalid Rotation Start Position
                                            m_lngErrorNo = trErrbusSchedStartPosIsZero
                                            AppendPersonalError r_rsErrors, TrType_PER, m_lngRecordId, "ROT_WEEK", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), r_rsPersonalSchedRequest!EMP_NUM, r_rsPersonalSchedRequest!START_DATE, r_rsPersonalSchedRequest!END_DATE
                                            l_blnValidationError = True
                                        End If
                                    Else
                                        'trErrbusSchedStartPosNull Invalid Rotation Start Position
                                        m_lngErrorNo = trErrbusSchedStartPosNull
                                        AppendPersonalError r_rsErrors, TrType_PER, m_lngRecordId, "ROT_WEEK", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), r_rsPersonalSchedRequest!EMP_NUM, r_rsPersonalSchedRequest!START_DATE, r_rsPersonalSchedRequest!END_DATE
                                        l_blnValidationError = True
                                    End If
                                Else
                                    'trErrbusSchedRotCodeIsZero Invalid Rotation Code
                                    m_lngErrorNo = trErrbusSchedRotCodeIsZero
                                    AppendPersonalError r_rsErrors, TrType_PER, m_lngRecordId, "ROT_CD", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), r_rsPersonalSchedRequest!EMP_NUM, r_rsPersonalSchedRequest!START_DATE, r_rsPersonalSchedRequest!END_DATE
                                    l_blnValidationError = True
                                End If
                            Else
                                'trErrbusSchedRotCodeNull Invalid Rotation Code
                                m_lngErrorNo = trErrbusSchedRotCodeNull
                                AppendPersonalError r_rsErrors, TrType_PER, m_lngRecordId, "ROT_CD", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), r_rsPersonalSchedRequest!EMP_NUM, r_rsPersonalSchedRequest!START_DATE, r_rsPersonalSchedRequest!END_DATE
                                l_blnValidationError = True
                            End If
                        Else
                            If Not IsNull(r_rsPersonalSchedRequest!ROT_CD) Then
                                If Val(r_rsPersonalSchedRequest!ROT_CD) < 1 And IsNumeric(r_rsPersonalSchedRequest!ROT_CD) Then
                                    If Not IsNull(r_rsPersonalSchedRequest!ROT_WEEK) Then
                                        If Val(r_rsPersonalSchedRequest!ROT_WEEK) < 1 And IsNumeric(r_rsPersonalSchedRequest!ROT_WEEK) Then
                                        Else
                                            'trErrbusSchedPersStartPosNotZero Rotation Start Position is invalid when Schedule Code is greater than zero
                                            m_lngErrorNo = trErrbusSchedPersStartPosNotZero
                                            AppendPersonalError r_rsErrors, TrType_PER, m_lngRecordId, "ROT_WEEK", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), r_rsPersonalSchedRequest!EMP_NUM, r_rsPersonalSchedRequest!START_DATE, r_rsPersonalSchedRequest!END_DATE
                                            l_blnValidationError = True
                                        End If
                                    Else
                                        'trErrbusSchedStartPosNull Invalid Rotation Start Position
                                        m_lngErrorNo = trErrbusSchedStartPosNull
                                        AppendPersonalError r_rsErrors, TrType_PER, m_lngRecordId, "ROT_WEEK", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), r_rsPersonalSchedRequest!EMP_NUM, r_rsPersonalSchedRequest!START_DATE, r_rsPersonalSchedRequest!END_DATE
                                        l_blnValidationError = True
                                    End If
                                Else
                                    'trErrbusSchedPersRotCdNotZero Rotation Code must equal zero when Schedule Code is greater than zero
                                    m_lngErrorNo = trErrbusSchedPersRotCdNotZero
                                    AppendPersonalError r_rsErrors, TrType_PER, m_lngRecordId, "ROT_CD", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), r_rsPersonalSchedRequest!EMP_NUM, r_rsPersonalSchedRequest!START_DATE, r_rsPersonalSchedRequest!END_DATE
                                    l_blnValidationError = True
                                End If
                            Else
                                'trErrbusSchedRotCodeNull Invalid Rotation Code
                                m_lngErrorNo = trErrbusSchedRotCodeNull
                                AppendPersonalError r_rsErrors, TrType_PER, m_lngRecordId, "ROT_CD", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), r_rsPersonalSchedRequest!EMP_NUM, r_rsPersonalSchedRequest!START_DATE, r_rsPersonalSchedRequest!END_DATE
                                l_blnValidationError = True
                            End If
                        End If
                    Else
                        If Not IsNull(r_rsPersonalSchedRequest!ROT_CD) Then
                            If Val(r_rsPersonalSchedRequest!ROT_CD) > 0 And IsNumeric(r_rsPersonalSchedRequest!ROT_CD) Then
                                If Not IsNull(r_rsPersonalSchedRequest!ROT_WEEK) Then
                                    If Val(r_rsPersonalSchedRequest!ROT_WEEK) > 0 And IsNumeric(r_rsPersonalSchedRequest!ROT_WEEK) Then
                                    Else
                                        'trErrbusSchedStartPosIsZero Invalid Rotation Start Position
                                        m_lngErrorNo = trErrbusSchedStartPosIsZero
                                        AppendPersonalError r_rsErrors, TrType_PER, m_lngRecordId, "ROT_WEEK", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), r_rsPersonalSchedRequest!EMP_NUM, r_rsPersonalSchedRequest!START_DATE, r_rsPersonalSchedRequest!END_DATE
                                        l_blnValidationError = True
                                    End If
                                Else
                                    'trErrbusSchedStartPosNull Invalid Rotation Start Position
                                    m_lngErrorNo = trErrbusSchedStartPosNull
                                    AppendPersonalError r_rsErrors, TrType_PER, m_lngRecordId, "ROT_WEEK", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), r_rsPersonalSchedRequest!EMP_NUM, r_rsPersonalSchedRequest!START_DATE, r_rsPersonalSchedRequest!END_DATE
                                    l_blnValidationError = True
                                End If
                            Else
                                'trErrbusSchedRotCodeIsZero Invalid Rotation Code
                                m_lngErrorNo = trErrbusSchedRotCodeIsZero
                                AppendPersonalError r_rsErrors, TrType_PER, m_lngRecordId, "ROT_CD", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), r_rsPersonalSchedRequest!EMP_NUM, r_rsPersonalSchedRequest!START_DATE, r_rsPersonalSchedRequest!END_DATE
                                l_blnValidationError = True
                            End If
                        Else
                            'trErrbusSchedRotCodeNull Invalid Rotation Code
                            m_lngErrorNo = trErrbusSchedRotCodeNull
                            AppendPersonalError r_rsErrors, TrType_PER, m_lngRecordId, "ROT_CD", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), r_rsPersonalSchedRequest!EMP_NUM, r_rsPersonalSchedRequest!START_DATE, r_rsPersonalSchedRequest!END_DATE
                            l_blnValidationError = True
                        End If
                    End If
                    If Not l_blnValidationError Then
                        l_lngEmpNum = r_rsPersonalSchedRequest!EMP_NUM
                        l_datStartDateAdj = r_rsPersonalSchedRequest!START_DATE
                        If IsStartDateBeforeEmployeeEmploymentDate(l_lngEmpNum, l_datStartDateAdj) Then
                            l_strWarDesc = "System changed the schedule start date because it was before the employee's employment date. "
                            r_rsPersonalSchedRequest!START_DATE = l_datStartDateAdj
                            r_rsPersonalSchedRequest!KEY_START_DATE = l_datStartDateAdj
                                AppendBusinessError r_rsErrors, 20000, mstrCBUSINESSWAR, l_strWarDesc, mstrCMODULENAME, TrType_PER, m_lngRecordId, "START_DATE", _
                            mstrCVARFIELDNAMES, _
                            CStr(r_rsPersonalSchedRequest!EMP_NUM), Format(r_rsPersonalSchedRequest!START_DATE, "mm/dd/yyyy"), Format(r_rsPersonalSchedRequest!END_DATE, "mm/dd/yyyy"), ""
                        End If
                    
                    End If
                End If
            End If
        Else
            ' Invalid RecStatusNull
            m_lngErrorNo = trErrbusSchedRotRecStatusNull
            AppendPersonalError r_rsErrors, TrType_PER, m_lngRecordId, "RECORD_STATUS", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), r_rsPersonalSchedRequest!KEY_EMP_NUM, r_rsPersonalSchedRequest!KEY_START_DATE, r_rsPersonalSchedRequest!KEY_END_DATE
            l_blnValidationError = True
        End If
        r_rsPersonalSchedRequest.MoveNext
    Loop

    r_rsPersonalSchedRequest.MoveFirst
    ValidatePersonalSchedEntry = Not l_blnValidationError
    
Exit Function

eh:
    TrsRaiseError mstrCMODULENAME, "ValidatePersonalSchedEntry"

End Function

'validate employee number
Private Function IsValidEmployee _
                (ByVal v_lngEmployee As Long _
                ) As Boolean

On Error GoTo IsValidEmployee_EH

    Dim l_dbEmployee As busEmployee.cbsEmployee

    Set l_dbEmployee = CreateObject("busEmployee.cbsEmployee")

    IsValidEmployee = l_dbEmployee.IsValidEmployee(v_lngEmployee)
        
    Set l_dbEmployee = Nothing
    GetObjectContext.SetComplete

Exit Function

IsValidEmployee_EH:
    Set l_dbEmployee = Nothing
    GetObjectContext.SetAbort
    TrsRaiseError mstrCMODULENAME, "IsValidEmployee"

End Function

'build the net effect from requested personal schedules
Private Function CreateNetPersonalSchedUpdates() As ADODB.Recordset

On Error GoTo CreateNetPersonalSchedUpdates_EH

'Const l_lngCFIELDATTR = adFldUpdatable Or adFldIsNullable
Const l_lngCFIELDATTR = adFldUpdatable

    'Create Recordset
    Dim l_adrsPersonalSaveEntries As ADODB.Recordset

    Set l_adrsPersonalSaveEntries = New ADODB.Recordset

    With l_adrsPersonalSaveEntries

        .Fields.Append "EMP_NUM", adInteger, , l_lngCFIELDATTR
        .Fields.Append "START_DATE", adDate, , l_lngCFIELDATTR
        .Fields.Append "END_DATE", adDate, , l_lngCFIELDATTR
        .Fields.Append "ROT_CD", adInteger, , l_lngCFIELDATTR
        .Fields.Append "ROT_WEEK", adInteger, , l_lngCFIELDATTR
        .Fields.Append "SCHED_CD", adInteger, , l_lngCFIELDATTR
        .Fields.Append "CREATED_EMP_NUM", adInteger, , l_lngCFIELDATTR
        .Fields.Append "LAST_MOD_EMP_NUM", adInteger, , l_lngCFIELDATTR
        .Fields.Append "CREATED_TIMESTAMP", adChar, 26, l_lngCFIELDATTR
        .Fields.Append "LAST_MOD_TIMESTAMP", adChar, 26, l_lngCFIELDATTR

        .Fields.Append "KEY_EMP_NUM", adInteger, , l_lngCFIELDATTR
        .Fields.Append "KEY_START_DATE", adDate, , l_lngCFIELDATTR
        .Fields.Append "KEY_END_DATE", adDate, , l_lngCFIELDATTR
        .Fields.Append "RECORD_STATUS", adChar, 1, l_lngCFIELDATTR
        .Fields.Append "RECORD_ID", adInteger, , l_lngCFIELDATTR
        .Open

    End With
    
    Set CreateNetPersonalSchedUpdates = l_adrsPersonalSaveEntries

Exit Function

CreateNetPersonalSchedUpdates_EH:
    TrsRaiseError mstrCMODULENAME, "CreateNetPersonalSchedUpdates"

End Function

'get Scope of existing schedules where their date ranges are affected by the new start and end dates entered
Private Function GetScopePersonalSched _
                (ByVal v_lngEmpNum As Long _
                 , ByVal v_datStartDate As Date _
                 , ByVal v_datEndDate As Date _
                ) As ADODB.Recordset


On Error GoTo GetScopePersonalSched_EH

    Dim l_dbSchedPersonal As dbSchedPersonal.cdSchedPersonal

    Set l_dbSchedPersonal = CreateObject("dbSchedPersonal.cdSchedPersonal")

    Set GetScopePersonalSched = l_dbSchedPersonal.GetScopePersonalSched(v_lngEmpNum, v_datStartDate, v_datEndDate)

    Set l_dbSchedPersonal = Nothing
    GetObjectContext.SetComplete

Exit Function

GetScopePersonalSched_EH:
    Set l_dbSchedPersonal = Nothing
    GetObjectContext.SetAbort
    TrsRaiseError mstrCMODULENAME, "GetScopePersonalSched"

End Function


'if record found delete them all - dates that fall in between the affected dates
Private Function DelNetSkipExistBetwnNewReq( _
    ByVal v_intRecordId As Integer, _
    ByVal v_strFilter As String, _
    ByRef r_objRSExistingPersonalSched As ADODB.Recordset, _
    ByRef r_rsNetPersonalSchedUpdates As ADODB.Recordset) _
        As Boolean

On Error GoTo DelNetSkipExistBetwnNewReq_EH

    Dim l_datOutdate1 As Date
    Dim l_datOutdate2 As Date
    
        
    If IsRecordsInFilterRecordSet(v_strFilter, r_objRSExistingPersonalSched) Then 'updating key fields if deleting
    Else
        Exit Function
    End If


    Do Until r_objRSExistingPersonalSched.EOF
        l_datOutdate1 = r_objRSExistingPersonalSched!START_DATE
        l_datOutdate2 = r_objRSExistingPersonalSched!END_DATE
        ConstructCDIupdaterequests _
        r_rsNetPersonalSchedUpdates _
        , r_objRSExistingPersonalSched!EMP_NUM _
        , DELETENDS_1 _
        , v_intRecordId _
        , , , _
        , l_datOutdate1 _
        , l_datOutdate2 _
        , , , _
        , r_objRSExistingPersonalSched!LAST_MOD_TIMESTAMP
        'build DB2Request ("DELETENDS")
        r_objRSExistingPersonalSched!RECORD_STATUS = "X"  'flag all deleted records
        r_objRSExistingPersonalSched.Update 'added to ensure delete
        r_objRSExistingPersonalSched.MoveNext
    Loop
    
    If IsClearedFilterRecordsetEmpty(r_objRSExistingPersonalSched) Then
    End If
    
    
Exit Function

DelNetSkipExistBetwnNewReq_EH:
    GetObjectContext.SetAbort
    TrsRaiseError mstrCMODULENAME, "DelNetSkipExistBetwnNewReq"

    
End Function


'Validate business rules
Private Function IsBUSValidPersonalSched(ByRef r_rsPersonalSchedRequest As ADODB.Recordset, _
                                         ByRef r_rsNetPersonalSchedUpdates As ADODB.Recordset, _
                                         ByRef r_rsErrors As ADODB.Recordset _
                                         ) As Boolean

On Error GoTo IsBUSValidPersonalSched_EH
    
'    Dim l_rdatEarliestStartDate As Date
'    Dim l_rdatLatestEndDate As Date

    IsBUSValidPersonalSched = True
    
'    l_rdatEarliestStartDate = m_datCOPTIONALENDDATE
'    l_rdatLatestEndDate = m_datCOPTIONALSTARTDATE
    
    Set r_rsNetPersonalSchedUpdates = CreateNetPersonalSchedUpdates
    
    If Not IsBUSDeletePersSched(r_rsPersonalSchedRequest, r_rsNetPersonalSchedUpdates, r_rsErrors) Then
       GetObjectContext.SetAbort
       IsBUSValidPersonalSched = False
       Exit Function
    End If
    
    If Not ProcessEndDateNull(r_rsPersonalSchedRequest, r_rsErrors) Then
       GetObjectContext.SetAbort
       IsBUSValidPersonalSched = False
       Exit Function
    End If

    If Not IsPREVALBUSInsert(r_rsPersonalSchedRequest, r_rsErrors) Then
       GetObjectContext.SetAbort
       IsBUSValidPersonalSched = False
       Exit Function
    End If

    If Not IsPREVALBUSChange(r_rsPersonalSchedRequest, r_rsErrors) Then
       GetObjectContext.SetAbort
       IsBUSValidPersonalSched = False
       Exit Function
    End If


    If Not IsBUSInsertPersSched(r_rsPersonalSchedRequest, r_rsNetPersonalSchedUpdates, r_rsErrors) Then
       GetObjectContext.SetAbort
       IsBUSValidPersonalSched = False
       Exit Function
    End If
    
    If Not IsBUSChangePersSched(r_rsPersonalSchedRequest, r_rsNetPersonalSchedUpdates, r_rsErrors) Then
       GetObjectContext.SetAbort
       IsBUSValidPersonalSched = False
       Exit Function
    End If
    

    
    
Exit Function

IsBUSValidPersonalSched_EH:

    TrsRaiseError mstrCMODULENAME, "IsBUSValidPersonalSched"

End Function



'Business rules for delete and date overlap
Private Function IsBUSDeletePersSched(ByRef r_rsPersonalSchedRequest As ADODB.Recordset, _
                                      ByRef r_rsNetPersonalSchedUpdates As ADODB.Recordset, _
                                      ByRef r_rsErrors As ADODB.Recordset _
                                      ) As Boolean

On Error GoTo IsBUSDeletePersSched_EH
    
    Dim l_strFilter As String
    
    IsBUSDeletePersSched = True
    
    If IsClearedFilterRecordsetEmpty(r_rsPersonalSchedRequest) Then
        Exit Function
    Else
    End If
    
    'process all records to be deleted
    
    l_strFilter = "RECORD_STATUS = 'D'"

    If IsRecordsInFilterRecordSet(l_strFilter, r_rsPersonalSchedRequest) Then
    Else
        Exit Function
    End If
    
    
    Do Until r_rsPersonalSchedRequest.EOF
        ConstructCDIupdaterequests _
        r_rsNetPersonalSchedUpdates _
        , r_rsPersonalSchedRequest!EMP_NUM _
        , DELETENDS_1 _
        , r_rsPersonalSchedRequest!RECORD_ID _
        , , , _
        , r_rsPersonalSchedRequest!KEY_START_DATE _
        , r_rsPersonalSchedRequest!KEY_END_DATE _
        , , , _
        , r_rsPersonalSchedRequest!LAST_MOD_TIMESTAMP
        'build DB2Request ("DELETENDS")
        AppendCascadeDates _
            "D", r_rsPersonalSchedRequest!EMP_NUM, _
            r_rsPersonalSchedRequest!KEY_START_DATE, r_rsPersonalSchedRequest!KEY_END_DATE
        r_rsPersonalSchedRequest.MoveNext
    Loop

    If IsClearedFilterRecordsetEmpty(r_rsPersonalSchedRequest) Then
    End If
    
Exit Function

IsBUSDeletePersSched_EH:
    
    TrsRaiseError mstrCMODULENAME, "IsBUSDeletePersSched"

End Function
'Business rules to process end date when null for "C" and "I"
Private Function ProcessEndDateNull(ByRef r_rsPersonalSchedRequest As ADODB.Recordset, _
                                   ByRef r_rsErrors As ADODB.Recordset) As Boolean

On Error GoTo ProcessEndDateNull_EH
    
    Dim l_rsPSCloneRequest As ADODB.Recordset
    Dim l_strFilter As String
    
    
    ProcessEndDateNull = True
    
    If IsClearedFilterRecordsetEmpty(r_rsPersonalSchedRequest) Then
        Exit Function
    Else
    End If
    
    
    Set l_rsPSCloneRequest = r_rsPersonalSchedRequest.Clone 'full clone of request
    
    l_strFilter = "RECORD_STATUS <> 'D'"
    If IsRecordsInFilterRecordSet(l_strFilter, r_rsPersonalSchedRequest) Then
    Else
        Exit Function
    End If
    
    
    Do Until r_rsPersonalSchedRequest.EOF
        If r_rsPersonalSchedRequest!RECORD_STATUS = "I" Then
            If IsNull(r_rsPersonalSchedRequest!END_DATE) Then
                r_rsPersonalSchedRequest!END_DATE = _
                CDate(GetDBDefaultEndDate(r_rsPersonalSchedRequest!EMP_NUM _
                                         , r_rsPersonalSchedRequest!START_DATE _
                                         , l_rsPSCloneRequest))
                r_rsPersonalSchedRequest.Update
            Else
            End If
        Else
            If r_rsPersonalSchedRequest!RECORD_STATUS = "C" Then
                If IsNull(r_rsPersonalSchedRequest!END_DATE) Then
                    If IsPersonalSchedExist _
                        (r_rsPersonalSchedRequest!KEY_EMP_NUM _
                       , r_rsPersonalSchedRequest!KEY_START_DATE _
                       , r_rsPersonalSchedRequest!KEY_END_DATE) Then
                       r_rsPersonalSchedRequest!END_DATE = _
                       CDate(GetDBDefaultEndDate _
                              (r_rsPersonalSchedRequest!KEY_EMP_NUM _
                             , r_rsPersonalSchedRequest!KEY_START_DATE _
                             , l_rsPSCloneRequest)) 'skip existing record since it may be completely modified and get next db start date
                    Else
                        r_rsPersonalSchedRequest!END_DATE = _
                        CDate(GetDBDefaultEndDate _
                               (r_rsPersonalSchedRequest!KEY_EMP_NUM _
                              , r_rsPersonalSchedRequest!START_DATE _
                              , l_rsPSCloneRequest))     'use the req start date to get next db closing the gap
                    End If
                    r_rsPersonalSchedRequest.Update
                Else
                End If
            Else
            End If
        End If
        r_rsPersonalSchedRequest.MoveNext
    Loop
    
    
    If IsClearedFilterRecordsetEmpty(r_rsPersonalSchedRequest) Then
    End If
    

    Set l_rsPSCloneRequest = Nothing

Exit Function
ProcessEndDateNull_EH:

    Set l_rsPSCloneRequest = Nothing

    TrsRaiseError mstrCMODULENAME, "ProcessEndDateNull"

End Function

'Business rules for insert and date overlap
Private Function IsPREVALBUSInsert(ByRef r_rsPersonalSchedRequest As ADODB.Recordset, _
                                   ByRef r_rsErrors As ADODB.Recordset) As Boolean

On Error GoTo IsPREVALBUSInsert_EH
    
    Dim l_objDBPersonalSched As ADODB.Recordset
    Dim l_objRSchedHeader As ADODB.Recordset
    Dim l_rsPSCloneRequest As ADODB.Recordset
    Dim l_strFilter As String
    
    
    IsPREVALBUSInsert = True
    
    If IsClearedFilterRecordsetEmpty(r_rsPersonalSchedRequest) Then
        Exit Function
    Else
    End If
    
    
    Set l_rsPSCloneRequest = r_rsPersonalSchedRequest.Clone 'full clone of request
    
    l_strFilter = "RECORD_STATUS = 'I'"

    If IsRecordsInFilterRecordSet(l_strFilter, r_rsPersonalSchedRequest) Then
    Else
        Exit Function
    End If
    
    
'need to eliminate duplicate end dates equal to null
'then filter the null for assignment of valid end dates
'update end dates
'separate the first pass for prevalidation then go for the real thing

'pre business validation
    
    Do Until r_rsPersonalSchedRequest.EOF
        If r_rsPersonalSchedRequest!START_DATE <= _
           r_rsPersonalSchedRequest!END_DATE Then
        Else
            'trErrbusSchedStartGreaterEndDt Schedule date cannot be greater than sched end date
            m_lngErrorNo = trErrbusSchedStartGreaterEndDt
            AppendPersonalError r_rsErrors, TrType_PER, m_lngRecordId, "START_DATE", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), r_rsPersonalSchedRequest!EMP_NUM, r_rsPersonalSchedRequest!START_DATE, r_rsPersonalSchedRequest!END_DATE
            IsPREVALBUSInsert = False
            Exit Function
        End If
        'check for overlap request C and I   'overlap error
        If IsBUSOverlapInRequest(l_rsPSCloneRequest _
           , r_rsPersonalSchedRequest!EMP_NUM _
           , r_rsPersonalSchedRequest!RECORD_STATUS _
           , r_rsPersonalSchedRequest!START_DATE _
           , r_rsPersonalSchedRequest!END_DATE _
           , m_datCOPTIONALSTARTDATE _
           , m_datCOPTIONALENDDATE) Then
            'trErrbusSchedPersonalIsDup duplicate records cannot be added
            m_lngRecordId = r_rsPersonalSchedRequest!RECORD_ID
            m_lngErrorNo = trErrbusSchedPersonalIsDup
            AppendPersonalError r_rsErrors, TrType_PER, m_lngRecordId, "START_DATE", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), r_rsPersonalSchedRequest!EMP_NUM, r_rsPersonalSchedRequest!START_DATE, r_rsPersonalSchedRequest!END_DATE
            IsPREVALBUSInsert = False
            Exit Function
        Else
        End If
        'CHECK FOR AN ALREADY EXISTING PERSONAL SCHEDULE BRV
            
        If Val(r_rsPersonalSchedRequest!SCHED_CD) <> 0 _
           And IsNumeric(r_rsPersonalSchedRequest!SCHED_CD) Then
           If IsValidSchedCd(r_rsPersonalSchedRequest) Then
           Else
               'trErrbusSchedCodeInvalid Invalid Schedule Code
               m_lngErrorNo = trErrbusSchedCodeInvalid
               AppendPersonalError r_rsErrors, TrType_PER, m_lngRecordId, "SCHED_CD", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), r_rsPersonalSchedRequest!EMP_NUM, r_rsPersonalSchedRequest!START_DATE, r_rsPersonalSchedRequest!END_DATE
               IsPREVALBUSInsert = False
               Exit Function
           End If
        Else
            Set l_objRSchedHeader = _
                GetRotationSchedHeader(r_rsPersonalSchedRequest!ROT_CD)
            If RecordsExist(l_objRSchedHeader) Then
                Set l_objDBPersonalSched = _
                    GetRotationSchedDetail(r_rsPersonalSchedRequest!ROT_CD _
                                         , r_rsPersonalSchedRequest!ROT_WEEK)
                If RecordsExist(l_objDBPersonalSched) Then
                Else
            '      trErrbusSchedRotStartPosNotFound Invalid Start Position
                    m_lngErrorNo = trErrbusSchedRotStartPosNotFound
                    AppendPersonalError r_rsErrors, TrType_PER, m_lngRecordId, "ROT_WEEK", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), r_rsPersonalSchedRequest!EMP_NUM, r_rsPersonalSchedRequest!START_DATE, r_rsPersonalSchedRequest!END_DATE
                    IsPREVALBUSInsert = False
                    Exit Function
                End If
            Else
    '        trErrbusSchedRotHeaderNotFound    Invalid Rotation Code
                m_lngErrorNo = trErrbusSchedRotHeaderNotFound
                AppendPersonalError r_rsErrors, TrType_PER, m_lngRecordId, "ROT_CD", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), r_rsPersonalSchedRequest!EMP_NUM, r_rsPersonalSchedRequest!START_DATE, r_rsPersonalSchedRequest!END_DATE
                IsPREVALBUSInsert = False
                Exit Function
            End If
        End If
        r_rsPersonalSchedRequest.MoveNext
    Loop

    If IsClearedFilterRecordsetEmpty(r_rsPersonalSchedRequest) Then
    End If
    

    Set l_objDBPersonalSched = Nothing
    Set l_objRSchedHeader = Nothing
    Set l_rsPSCloneRequest = Nothing

Exit Function
IsPREVALBUSInsert_EH:

    Set l_objDBPersonalSched = Nothing
    Set l_objRSchedHeader = Nothing
    Set l_rsPSCloneRequest = Nothing

    TrsRaiseError mstrCMODULENAME, "IsPREVALBUSInsert"

End Function

Private Function IsPREVALBUSChange(ByRef r_rsPersonalSchedRequest As ADODB.Recordset, _
                                      ByRef r_rsErrors As ADODB.Recordset _
                                      ) As Boolean


On Error GoTo IsPREVALBUSChange_EH
    
    Dim l_objDBPersonalSched As ADODB.Recordset
    Dim l_objRSchedHeader As ADODB.Recordset
    Dim l_rsPSCloneRequest As ADODB.Recordset
    Dim l_strFilter As String

    
    IsPREVALBUSChange = True
    
    If IsClearedFilterRecordsetEmpty(r_rsPersonalSchedRequest) Then
        Exit Function
    Else
    End If
    
    Set l_rsPSCloneRequest = r_rsPersonalSchedRequest.Clone 'Full clone
    
    l_strFilter = "RECORD_STATUS = 'C'"
    If IsRecordsInFilterRecordSet(l_strFilter, r_rsPersonalSchedRequest) Then
    Else
        Exit Function
    End If
    

'pre business validation
    Do Until r_rsPersonalSchedRequest.EOF
        'check for duplicate key date send in the request recordset
        If IsduplicateDateKeysinrequest _
            (r_rsPersonalSchedRequest!KEY_EMP_NUM _
           , r_rsPersonalSchedRequest!KEY_START_DATE _
           , r_rsPersonalSchedRequest!KEY_END_DATE _
           , l_rsPSCloneRequest) Then
           m_lngErrorNo = trErrbusSchedPersonalIsDup
           AppendPersonalError r_rsErrors, TrType_PER, m_lngRecordId, "START_DATE", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), r_rsPersonalSchedRequest!EMP_NUM, r_rsPersonalSchedRequest!START_DATE, r_rsPersonalSchedRequest!END_DATE
           IsPREVALBUSChange = False
           Exit Function
        End If
        
        
        If r_rsPersonalSchedRequest!START_DATE <= _
           r_rsPersonalSchedRequest!END_DATE Then
        Else
            'trErrbusSchedStartGreaterEndDt  Schedule date cannot be greater than sched end date
            m_lngErrorNo = trErrbusSchedStartGreaterEndDt
            AppendPersonalError r_rsErrors, TrType_PER, m_lngRecordId, "START_DATE", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), r_rsPersonalSchedRequest!EMP_NUM, r_rsPersonalSchedRequest!START_DATE, r_rsPersonalSchedRequest!END_DATE
            IsPREVALBUSChange = False
            Exit Function
        End If
        
        'check for duplicate changed date requests - end dates won't be null
        If IsduplicateDateChangedrequest _
            (r_rsPersonalSchedRequest!EMP_NUM _
           , r_rsPersonalSchedRequest!START_DATE _
           , r_rsPersonalSchedRequest!END_DATE _
           , l_rsPSCloneRequest) Then
           m_lngErrorNo = trErrbusSchedPersonalIsDup
           AppendPersonalError r_rsErrors, TrType_PER, m_lngRecordId, "START_DATE", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), r_rsPersonalSchedRequest!EMP_NUM, r_rsPersonalSchedRequest!START_DATE, r_rsPersonalSchedRequest!END_DATE
           IsPREVALBUSChange = False
           Exit Function
        End If
        'check for overlap request for C and I
        If IsBUSOverlapInRequest(l_rsPSCloneRequest _
           , r_rsPersonalSchedRequest!EMP_NUM _
           , r_rsPersonalSchedRequest!RECORD_STATUS _
           , r_rsPersonalSchedRequest!START_DATE _
           , r_rsPersonalSchedRequest!END_DATE _
           , r_rsPersonalSchedRequest!KEY_START_DATE _
           , r_rsPersonalSchedRequest!KEY_END_DATE) Then
            'overlap error
            'trErrbusSchedPersonalIsDup duplicate records cannot be added
            m_lngRecordId = r_rsPersonalSchedRequest!RECORD_ID
            m_lngErrorNo = trErrbusSchedPersonalIsDup
            AppendPersonalError r_rsErrors, TrType_PER, m_lngRecordId, "START_DATE", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), r_rsPersonalSchedRequest!EMP_NUM, r_rsPersonalSchedRequest!START_DATE, r_rsPersonalSchedRequest!END_DATE
            IsPREVALBUSChange = False
            Exit Function
        Else
        End If
            
    'CHECK FOR AN ALREADY EXISTING PERSONAL SCHEDULE BRV - prevents generating duplicates
        
        If Val(r_rsPersonalSchedRequest!SCHED_CD) <> 0 _
           And IsNumeric(r_rsPersonalSchedRequest!SCHED_CD) Then
           If IsValidSchedCd(r_rsPersonalSchedRequest) Then
           Else
  '          trErrbusSchedCodeInvalid  Invalid Schedule Code
               m_lngErrorNo = trErrbusSchedCodeInvalid
               AppendPersonalError r_rsErrors, TrType_PER, m_lngRecordId, "SCHED_CD", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), r_rsPersonalSchedRequest!EMP_NUM, r_rsPersonalSchedRequest!START_DATE, r_rsPersonalSchedRequest!END_DATE
               IsPREVALBUSChange = False
               Exit Function
           End If
        Else
            Set l_objRSchedHeader = _
                GetRotationSchedHeader(r_rsPersonalSchedRequest!ROT_CD)
            If RecordsExist(l_objRSchedHeader) Then
                Set l_objDBPersonalSched = _
                    GetRotationSchedDetail _
                     (r_rsPersonalSchedRequest!ROT_CD _
                    , r_rsPersonalSchedRequest!ROT_WEEK)
                If RecordsExist(l_objDBPersonalSched) Then
                Else
            '         trErrbusSchedRotStartPosNotFound   Invalid Rotation Start Position
                    m_lngErrorNo = trErrbusSchedRotStartPosNotFound
                    AppendPersonalError r_rsErrors, TrType_PER, m_lngRecordId, "ROT_WEEK", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), r_rsPersonalSchedRequest!EMP_NUM, r_rsPersonalSchedRequest!START_DATE, r_rsPersonalSchedRequest!END_DATE
                    IsPREVALBUSChange = False
                    Exit Function
                End If
            Else
    '           trErrbusSchedRotHeaderNotFound Invalid Rotation Code
                m_lngErrorNo = trErrbusSchedRotHeaderNotFound
                AppendPersonalError r_rsErrors, TrType_PER, m_lngRecordId, "ROT_CD", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), r_rsPersonalSchedRequest!EMP_NUM, r_rsPersonalSchedRequest!START_DATE, r_rsPersonalSchedRequest!END_DATE
                IsPREVALBUSChange = False
                Exit Function
            End If
        End If
        r_rsPersonalSchedRequest.MoveNext
    Loop

    If IsClearedFilterRecordsetEmpty(r_rsPersonalSchedRequest) Then
    End If

    
    Set l_objDBPersonalSched = Nothing
    Set l_objRSchedHeader = Nothing
    Set l_rsPSCloneRequest = Nothing

Exit Function

IsPREVALBUSChange_EH:

    Set l_objDBPersonalSched = Nothing
    Set l_objRSchedHeader = Nothing
    Set l_rsPSCloneRequest = Nothing
    
    TrsRaiseError mstrCMODULENAME, "IsPREVALBUSChange"

End Function


'Business rules for insert and date overlap
Private Function IsBUSInsertPersSched(ByRef r_rsPersonalSchedRequest As ADODB.Recordset, _
                                      ByRef r_rsNetPersonalSchedUpdates As ADODB.Recordset, _
                                      ByRef r_rsErrors As ADODB.Recordset) As Boolean

On Error GoTo IsBUSInsertPersSched_EH
    
    Dim l_rsDBPersonalSchedExist As ADODB.Recordset
    Dim l_objDBPersonalSched As ADODB.Recordset
    Dim l_rsPSCloneRequest As ADODB.Recordset
    Dim l_objRSExistingPersonalSched As ADODB.Recordset
    Dim l_strFilter As String
    
    Dim l_datEarliestStartDate As Date
    Dim l_datLatestEndDate As Date
    Dim l_strFilterobjRSExist As String
    
    IsBUSInsertPersSched = True
    
    If IsClearedFilterRecordsetEmpty(r_rsPersonalSchedRequest) Then
        Exit Function
    Else
    End If
    
    Set l_rsPSCloneRequest = r_rsPersonalSchedRequest.Clone 'full clone of request
    
    l_strFilter = "RECORD_STATUS = 'I'"
    If IsRecordsInFilterRecordSet(l_strFilter, r_rsPersonalSchedRequest) Then
    Else
        Exit Function
    End If
    
    Do Until r_rsPersonalSchedRequest.EOF
        Set l_rsDBPersonalSchedExist = _
            GetExistPersonalSched(r_rsPersonalSchedRequest!EMP_NUM _
                                , r_rsPersonalSchedRequest!START_DATE _
                                , r_rsPersonalSchedRequest!END_DATE)
        If RecordsExist(l_rsDBPersonalSchedExist) Then
            If l_rsDBPersonalSchedExist!ROT_CD = r_rsPersonalSchedRequest!ROT_CD And _
               l_rsDBPersonalSchedExist!ROT_WEEK = r_rsPersonalSchedRequest!ROT_WEEK And _
               l_rsDBPersonalSchedExist!SCHED_CD = r_rsPersonalSchedRequest!SCHED_CD Then
               r_rsPersonalSchedRequest!RECORD_STATUS = " "
               r_rsPersonalSchedRequest.Update
               'trErrbusSchedPersonalIsDup duplicate records cannot be added
               'needs to be ignored since it already exist
            Else
'                'update only nonkey values
                ConstructCDIupdaterequests _
                  r_rsNetPersonalSchedUpdates _
                , r_rsPersonalSchedRequest!EMP_NUM _
                , DELETENDS_1 _
                , r_rsPersonalSchedRequest!RECORD_ID _
                , , , _
                , r_rsPersonalSchedRequest!START_DATE _
                , r_rsPersonalSchedRequest!END_DATE _
                , , , _
                , l_rsDBPersonalSchedExist!LAST_MOD_TIMESTAMP
                ConstructCDIupdaterequests _
                  r_rsNetPersonalSchedUpdates _
                , r_rsPersonalSchedRequest!EMP_NUM _
                , INSERTNDS_3 _
                , r_rsPersonalSchedRequest!RECORD_ID _
                , r_rsPersonalSchedRequest!SCHED_CD _
                , r_rsPersonalSchedRequest!ROT_CD _
                , r_rsPersonalSchedRequest!ROT_WEEK _
                , r_rsPersonalSchedRequest!START_DATE _
                , r_rsPersonalSchedRequest!END_DATE _
                , , , _
                r_rsPersonalSchedRequest!CREATED_TIMESTAMP
                'build DB2Request ("INSERTNDS") '(3) inserts new requested date range replacing any existing deleted records
            End If
        Else
            'since it does not exist continue processing
            'add establish_add updates after all the preliminary business checks were done
            'can this process be improved by extracting earliest and latest dates for same employee????
    ' for efficiency reuse
            If l_objRSExistingPersonalSched Is Nothing Then 'first time is empty
                Set l_objRSExistingPersonalSched = _
                    GetScopePersonalSched _
                    (r_rsPersonalSchedRequest!EMP_NUM _
                    , r_rsPersonalSchedRequest!START_DATE _
                    , r_rsPersonalSchedRequest!END_DATE)
            Else
                If IsClearedFilterRecordsetEmpty(l_objRSExistingPersonalSched) Then 'if empty call
                    Set l_objRSExistingPersonalSched = _
                        GetScopePersonalSched _
                        (r_rsPersonalSchedRequest!EMP_NUM _
                        , r_rsPersonalSchedRequest!START_DATE _
                        , r_rsPersonalSchedRequest!END_DATE)
                Else
                    l_strFilterobjRSExist = "START_DATE <= #" & r_rsPersonalSchedRequest!START_DATE & "#" & _
                             " AND END_DATE >= #" & r_rsPersonalSchedRequest!END_DATE & "#" & _
                             " AND EMP_NUM = " & r_rsPersonalSchedRequest!EMP_NUM
                    If IsRecordsInFilterRecordSet(l_strFilterobjRSExist, l_objRSExistingPersonalSched) Then 'if same records reuse
                    Else  'if not in this set call a new set
                        Set l_objRSExistingPersonalSched = _
                            GetScopePersonalSched _
                            (r_rsPersonalSchedRequest!EMP_NUM _
                            , r_rsPersonalSchedRequest!START_DATE _
                            , r_rsPersonalSchedRequest!END_DATE)
                    End If
                End If
            End If
            If RecordsExist(l_objRSExistingPersonalSched) Then
               BuildNetModifyExistDates _
                 r_rsPersonalSchedRequest!EMP_NUM _
               , r_rsPersonalSchedRequest!RECORD_STATUS _
               , r_rsPersonalSchedRequest!START_DATE _
               , r_rsPersonalSchedRequest!END_DATE _
               , r_rsPersonalSchedRequest!RECORD_ID _
               , l_objRSExistingPersonalSched _
               , r_rsNetPersonalSchedUpdates _
               , l_rsPSCloneRequest
            End If
            
            ConstructCDIupdaterequests _
              r_rsNetPersonalSchedUpdates _
            , r_rsPersonalSchedRequest!EMP_NUM _
            , INSERTNDS_3 _
            , r_rsPersonalSchedRequest!RECORD_ID _
            , r_rsPersonalSchedRequest!SCHED_CD _
            , r_rsPersonalSchedRequest!ROT_CD _
            , r_rsPersonalSchedRequest!ROT_WEEK _
            , r_rsPersonalSchedRequest!START_DATE _
            , r_rsPersonalSchedRequest!END_DATE _
            , , , _
            r_rsPersonalSchedRequest!CREATED_TIMESTAMP
            'build DB2Request ("INSERTNDS")
        End If
        AppendCascadeDates _
            "I", r_rsPersonalSchedRequest!EMP_NUM, _
            r_rsPersonalSchedRequest!START_DATE, r_rsPersonalSchedRequest!END_DATE
        r_rsPersonalSchedRequest.MoveNext
    Loop
    
    If IsClearedFilterRecordsetEmpty(r_rsPersonalSchedRequest) Then
    End If
    

    Set l_rsDBPersonalSchedExist = Nothing
    Set l_objDBPersonalSched = Nothing
    Set l_rsPSCloneRequest = Nothing
    Set l_objRSExistingPersonalSched = Nothing

Exit Function
IsBUSInsertPersSched_EH:

    Set l_rsDBPersonalSchedExist = Nothing
    Set l_objDBPersonalSched = Nothing
    Set l_rsPSCloneRequest = Nothing
    Set l_objRSExistingPersonalSched = Nothing

    TrsRaiseError mstrCMODULENAME, "IsBUSInsertPersSched"

End Function

Private Function IsBUSChangePersSched(ByRef r_rsPersonalSchedRequest As ADODB.Recordset, _
                                      ByRef r_rsNetPersonalSchedUpdates As ADODB.Recordset, _
                                      ByRef r_rsErrors As ADODB.Recordset _
                                      ) As Boolean


On Error GoTo IsBUSChangePersSched_EH
    
    Dim l_rsDBPersonalSchedExist As ADODB.Recordset
    Dim l_objDBPersonalSched As ADODB.Recordset
    Dim l_rsPSCloneRequest As ADODB.Recordset
    Dim l_objRSExistingPersonalSched As ADODB.Recordset
    Dim l_strFilter As String

    Dim l_datEarliestStartDate As Date
    Dim l_datLatestEndDate As Date
    Dim l_strFilterobjRSExist As String
    
    IsBUSChangePersSched = True
    
    If IsClearedFilterRecordsetEmpty(r_rsPersonalSchedRequest) Then
        Exit Function
    Else
    End If

    Set l_rsPSCloneRequest = r_rsPersonalSchedRequest.Clone 'Full clone
    
    l_strFilter = "RECORD_STATUS = 'C'"
    If IsRecordsInFilterRecordSet(l_strFilter, r_rsPersonalSchedRequest) Then
    Else
        Exit Function
    End If
    
    
    Do Until r_rsPersonalSchedRequest.EOF
        If IsExistNonkeyFieldChangeOnly _
            (r_rsPersonalSchedRequest!EMP_NUM _
           , r_rsPersonalSchedRequest!START_DATE _
           , r_rsPersonalSchedRequest!END_DATE _
           , r_rsPersonalSchedRequest!KEY_START_DATE _
           , r_rsPersonalSchedRequest!KEY_END_DATE _
           , l_rsDBPersonalSchedExist) Then
           If l_rsDBPersonalSchedExist!ROT_CD = r_rsPersonalSchedRequest!ROT_CD And _
              l_rsDBPersonalSchedExist!ROT_WEEK = r_rsPersonalSchedRequest!ROT_WEEK And _
              l_rsDBPersonalSchedExist!SCHED_CD = r_rsPersonalSchedRequest!SCHED_CD Then
              r_rsPersonalSchedRequest!RECORD_STATUS = " "
              r_rsPersonalSchedRequest.Update  'skip this duplicate record without brv since it already exists
           Else
                'change only nonkey fields "C" and "D" should not be on the same request for the same dates and same employee
               ConstructCDIupdaterequests _
                 r_rsNetPersonalSchedUpdates _
               , r_rsPersonalSchedRequest!EMP_NUM _
               , UPDATENDS_2 _
               , r_rsPersonalSchedRequest!RECORD_ID _
               , r_rsPersonalSchedRequest!SCHED_CD _
               , r_rsPersonalSchedRequest!ROT_CD _
               , r_rsPersonalSchedRequest!ROT_WEEK _
               , r_rsPersonalSchedRequest!START_DATE _
               , r_rsPersonalSchedRequest!END_DATE _
               , r_rsPersonalSchedRequest!KEY_START_DATE _
               , r_rsPersonalSchedRequest!KEY_END_DATE _
               , _
               , r_rsPersonalSchedRequest!LAST_MOD_TIMESTAMP
               AppendCascadeDates _
                    "C", r_rsPersonalSchedRequest!EMP_NUM, _
                    r_rsPersonalSchedRequest!KEY_START_DATE, r_rsPersonalSchedRequest!KEY_END_DATE
           End If
        Else
        'add establish_change
    ' for efficiency reuse
            If l_objRSExistingPersonalSched Is Nothing Then 'first time is empty
                Set l_objRSExistingPersonalSched = _
                    GetScopePersonalSched _
                    (r_rsPersonalSchedRequest!EMP_NUM _
                    , r_rsPersonalSchedRequest!START_DATE _
                    , r_rsPersonalSchedRequest!END_DATE)
            Else
                If IsClearedFilterRecordsetEmpty(l_objRSExistingPersonalSched) Then 'if empty call
                    Set l_objRSExistingPersonalSched = _
                        GetScopePersonalSched _
                        (r_rsPersonalSchedRequest!EMP_NUM _
                        , r_rsPersonalSchedRequest!START_DATE _
                        , r_rsPersonalSchedRequest!END_DATE)
                Else
                    l_strFilterobjRSExist = "START_DATE <= #" & r_rsPersonalSchedRequest!START_DATE & "#" & _
                             " AND END_DATE >= #" & r_rsPersonalSchedRequest!END_DATE & "#" & _
                             " AND EMP_NUM = " & r_rsPersonalSchedRequest!EMP_NUM
                    If IsRecordsInFilterRecordSet(l_strFilterobjRSExist, l_objRSExistingPersonalSched) Then 'if same records reuse
                    Else  'if not in this set call a new set
                        Set l_objRSExistingPersonalSched = _
                            GetScopePersonalSched _
                            (r_rsPersonalSchedRequest!EMP_NUM _
                            , r_rsPersonalSchedRequest!START_DATE _
                            , r_rsPersonalSchedRequest!END_DATE)
                    End If
                End If
            End If
            If RecordsExist(l_objRSExistingPersonalSched) Then 'db recs ' establish_change - key dates updates (2) once the existing two affected records are identified (these two records fall to the right and to the left of the new start and end dates requested) use establish_yyy to update these two records
                l_strFilter = "START_DATE >= #" & r_rsPersonalSchedRequest!START_DATE & "#" & _
                         " AND END_DATE <= #" & r_rsPersonalSchedRequest!END_DATE & "#" & _
                         " AND EMP_NUM = " & r_rsPersonalSchedRequest!EMP_NUM

                DelNetSkipExistBetwnNewReq r_rsPersonalSchedRequest!RECORD_ID _
                                         , l_strFilter _
                                         , l_objRSExistingPersonalSched _
                                         , r_rsNetPersonalSchedUpdates  '1. all existing db records that fall in between the New requested dates are skipped
                If IsClearedFilterRecordsetEmpty(l_objRSExistingPersonalSched) Then
                End If
                SkipExistEqualKeyDateReq r_rsPersonalSchedRequest!KEY_EMP_NUM _
                                       , r_rsPersonalSchedRequest!KEY_START_DATE _
                                       , r_rsPersonalSchedRequest!KEY_END_DATE _
                                       , l_objRSExistingPersonalSched '2. skip current existing record being changed so it won't process in BuildNetModifyExistDates (establish_add) ...it will be deleted by key
                'if record found delete them all - dates that fall in between the affected dates (1)
                
                l_strFilter = "RECORD_STATUS <> 'X'" 'search for existing Personal sched already  processed
                If IsRecordsInFilterRecordSet(l_strFilter, l_objRSExistingPersonalSched) Then
                    BuildNetModifyExistDates _
                      r_rsPersonalSchedRequest!EMP_NUM _
                    , r_rsPersonalSchedRequest!RECORD_STATUS _
                    , r_rsPersonalSchedRequest!START_DATE _
                    , r_rsPersonalSchedRequest!END_DATE _
                    , r_rsPersonalSchedRequest!RECORD_ID _
                    , l_objRSExistingPersonalSched _
                    , r_rsNetPersonalSchedUpdates _
                    , l_rsPSCloneRequest 'process (these two records fall to the right and to the left of the new start and end dates requested) use establish_yyy to update these two records
                Else
                End If
            End If
            '3. Key record must be deleted also
            ConstructCDIupdaterequests _
              r_rsNetPersonalSchedUpdates _
            , r_rsPersonalSchedRequest!EMP_NUM _
            , DELETENDS_1 _
            , r_rsPersonalSchedRequest!RECORD_ID _
            , , , _
            , r_rsPersonalSchedRequest!KEY_START_DATE _
            , r_rsPersonalSchedRequest!KEY_END_DATE _
            , , , _
            , r_rsPersonalSchedRequest!LAST_MOD_TIMESTAMP
            
            ConstructCDIupdaterequests _
              r_rsNetPersonalSchedUpdates _
            , r_rsPersonalSchedRequest!EMP_NUM _
            , INSERTNDS_3 _
            , r_rsPersonalSchedRequest!RECORD_ID _
            , r_rsPersonalSchedRequest!SCHED_CD _
            , r_rsPersonalSchedRequest!ROT_CD _
            , r_rsPersonalSchedRequest!ROT_WEEK _
            , r_rsPersonalSchedRequest!START_DATE _
            , r_rsPersonalSchedRequest!END_DATE
            'build DB2Request ("INSERTNDS") '4. inserts New requested date range
        End If
        If l_rsDBPersonalSchedExist Is Nothing Then
            GetExistingPersonalSchedule _
                r_rsPersonalSchedRequest!EMP_NUM, _
                r_rsPersonalSchedRequest!KEY_START_DATE, _
                r_rsPersonalSchedRequest!KEY_END_DATE, _
                l_rsDBPersonalSchedExist
        End If
        If RecordsExist(l_rsDBPersonalSchedExist) Then
           If l_rsDBPersonalSchedExist!ROT_CD = r_rsPersonalSchedRequest!ROT_CD And _
              l_rsDBPersonalSchedExist!ROT_WEEK = r_rsPersonalSchedRequest!ROT_WEEK And _
              l_rsDBPersonalSchedExist!SCHED_CD = r_rsPersonalSchedRequest!SCHED_CD Then
'               Date change only.
                AppendCascadeDateDifference r_rsPersonalSchedRequest
'               For rotation, if START_DATE changed and rotation week did not change,
'               it is implied schedule change.
                If r_rsPersonalSchedRequest!ROT_CD <> 0 And _
                    r_rsPersonalSchedRequest!START_DATE <> r_rsPersonalSchedRequest!KEY_START_DATE Then
                    AppendCascadeDates _
                        "C", r_rsPersonalSchedRequest!EMP_NUM, _
                        r_rsPersonalSchedRequest!KEY_START_DATE, r_rsPersonalSchedRequest!KEY_END_DATE
                End If
            Else
'                date and schedule change
                AppendCascadeDateDifference r_rsPersonalSchedRequest
                AppendCascadeDates _
                    "C", r_rsPersonalSchedRequest!EMP_NUM, _
                    r_rsPersonalSchedRequest!KEY_START_DATE, r_rsPersonalSchedRequest!KEY_END_DATE
            End If
        End If
        r_rsPersonalSchedRequest.MoveNext
    Loop
    
    If IsClearedFilterRecordsetEmpty(r_rsPersonalSchedRequest) Then
    End If
    
    
    Set l_rsDBPersonalSchedExist = Nothing
    Set l_objDBPersonalSched = Nothing
    Set l_rsPSCloneRequest = Nothing
    Set l_objRSExistingPersonalSched = Nothing

Exit Function

IsBUSChangePersSched_EH:

    Set l_rsDBPersonalSchedExist = Nothing
    Set l_objDBPersonalSched = Nothing
    Set l_rsPSCloneRequest = Nothing
    Set l_objRSExistingPersonalSched = Nothing
    
    TrsRaiseError mstrCMODULENAME, "IsBUSChangePersSched"

End Function


Private Function AppendCascadeDateDifference( _
                    ByRef r_rsPersonalSchedRequest As ADODB.Recordset)

On Error GoTo eh
    If r_rsPersonalSchedRequest!START_DATE < r_rsPersonalSchedRequest!KEY_START_DATE Then
        AppendCascadeDates _
            "C", r_rsPersonalSchedRequest!EMP_NUM, _
            r_rsPersonalSchedRequest!START_DATE, _
            DateAdd("d", -1, r_rsPersonalSchedRequest!KEY_START_DATE)
    ElseIf r_rsPersonalSchedRequest!START_DATE > r_rsPersonalSchedRequest!KEY_START_DATE Then
        AppendCascadeDates _
            "C", r_rsPersonalSchedRequest!EMP_NUM, _
            r_rsPersonalSchedRequest!KEY_START_DATE, _
            DateAdd("d", -1, r_rsPersonalSchedRequest!START_DATE)
    End If
    If r_rsPersonalSchedRequest!END_DATE < r_rsPersonalSchedRequest!KEY_END_DATE Then
        AppendCascadeDates _
            "C", r_rsPersonalSchedRequest!EMP_NUM, _
            DateAdd("d", 1, r_rsPersonalSchedRequest!END_DATE), _
            r_rsPersonalSchedRequest!KEY_END_DATE
    ElseIf r_rsPersonalSchedRequest!END_DATE > r_rsPersonalSchedRequest!KEY_END_DATE Then
        AppendCascadeDates _
            "C", r_rsPersonalSchedRequest!EMP_NUM, _
            DateAdd("d", 1, r_rsPersonalSchedRequest!KEY_END_DATE), _
            r_rsPersonalSchedRequest!END_DATE
    End If

    Exit Function

eh:
    TrsRaiseError mstrCMODULENAME, "AppendCascadeDateDifference"
End Function


'establish_yyy or update existing personal schedules when overlapping dates occurs
'TrCase_1 and TrCase_2 cover input req dates (start date(s) and end date(e) that fall in
'the data base record range: s => db start date and e => db end date
'TrCase_1 more specificly: : s = db start date and e > db end date (the rest is for TrCase_2! s > db start date and e => db end date)
'TrCase_3 cover input req dates (start date(s) and end date(e) that fall in
'the data base record range: s <= db start date and e => db start date and < db end date
'TrCase_4 cover input req dates (start date(s) and end date(e) that fall in
'the data base record range: s => db start date and e <= db end date !(s > db start date and e < db end date)
'TrCase_5 cover input req dates (start date(s) and end date(e) that fall in
'the data base record range: s <= db start date and e => db end date
Private Function BuildNetModifyExistDates( _
    ByVal v_lngEmpNum As Long, _
    ByVal v_strRecordStatus As String, _
    ByVal v_datNewStartdate1 As Date, _
    ByVal v_datNewEnddate1 As Date, _
    ByVal v_intRecordId As Integer, _
    ByRef r_objRSExistingPersonalSched As ADODB.Recordset, _
    ByRef r_rsNetPersonalSchedUpdates As ADODB.Recordset, _
    ByVal v_rsPSCloneRequest As ADODB.Recordset) _
        As Boolean

On Error GoTo BuildNetModifyExistDates_EH

    Dim l_rdatOutdate1 As Date
    Dim l_rdatOutdate2 As Date
    Dim l_rdatOutdate3 As Date
    Dim l_rdatOutdate4 As Date
    Dim l_vdatOldStartdate2 As Date
    Dim l_vdatOldEnddate2 As Date
    Dim l_rstrDupFoundSwitch As String
    Dim l_intNewWeekPosition As Integer
    Dim v_datXupdStartdate1 As Date
    Dim v_datXupdEnddate1 As Date
    
    If IsClearedFilterRecordsetEmpty(r_objRSExistingPersonalSched) Then
        Exit Function
    Else
    End If
    

    If IsClearedFilterRecordsetEmpty(v_rsPSCloneRequest) Then
    End If
    
    Do Until r_objRSExistingPersonalSched.EOF
        If r_objRSExistingPersonalSched!RECORD_STATUS = "X" Then
        Else
            l_vdatOldStartdate2 = r_objRSExistingPersonalSched!START_DATE
            l_vdatOldEnddate2 = r_objRSExistingPersonalSched!END_DATE
            If (v_datNewStartdate1 >= l_vdatOldStartdate2) _
                And (v_datNewStartdate1 <= l_vdatOldEnddate2) _
                And (v_datNewEnddate1 >= l_vdatOldEnddate2) Then
                If (v_datNewStartdate1 = l_vdatOldStartdate2) _
                    And (l_vdatOldEnddate2 < v_datNewEnddate1) Then
                    'BuildNetModifyExistDates = TrCase_1
                    l_rdatOutdate1 = l_vdatOldStartdate2
                    l_rdatOutdate2 = l_vdatOldEnddate2
                    ConstructCDIupdaterequests _
                    r_rsNetPersonalSchedUpdates _
                    , r_objRSExistingPersonalSched!EMP_NUM _
                    , DELETENDS_1 _
                    , v_intRecordId _
                    , , , _
                    , l_rdatOutdate1 _
                    , l_rdatOutdate2 _
                    , , , , _
                    r_objRSExistingPersonalSched!LAST_MOD_TIMESTAMP
                    'build DB2Request ("DELETENDS")
                Else
                    If SkipDELKeyForSameExistDates(v_lngEmpNum _
                                                 , l_vdatOldStartdate2 _
                                                 , l_vdatOldEnddate2 _
                                                 , r_rsNetPersonalSchedUpdates) Then 'process only if it hasn't already been deleted - this happens when auto creating "C" since it cannot be "D" and "C" for same existing DB key in a request
                       r_objRSExistingPersonalSched!RECORD_STATUS = "X"  'flag all deleted records
                       r_objRSExistingPersonalSched.Update 'added to ensure delete
                    Else
                        If IsChangeFoundSkipInsert(v_lngEmpNum _
                                                     , v_strRecordStatus _
                                                     , l_vdatOldStartdate2 _
                                                     , l_vdatOldEnddate2 _
                                                     , v_rsPSCloneRequest) Then  'this process is delayed if "C" is found for same dates
                        Else
                            If r_objRSExistingPersonalSched!RECORD_STATUS = "X" Then
                            Else
                                'BuildNetModifyExistDates = TrCase_2
                                l_rdatOutdate1 = l_vdatOldStartdate2
                                l_rdatOutdate2 = l_vdatOldEnddate2
                                l_rdatOutdate3 = l_vdatOldStartdate2
                                l_rdatOutdate4 = l_vdatOldEnddate2
'                                IsExistUpdOverlapNewReqAdd v_lngEmpNum _
'                                                         , l_rdatOutdate1 _
'                                                         , l_rdatOutdate2 _
'                                                         , v_rsPSCloneRequest
                                
                                v_datXupdStartdate1 = l_vdatOldStartdate2
                                v_datXupdEnddate1 = l_vdatOldEnddate2
                                UpdExistPersfromInputReq v_lngEmpNum, _
                                                         v_intRecordId, _
                                                         l_rdatOutdate1, _
                                                         l_rdatOutdate2, _
                                                         v_rsPSCloneRequest, _
                                                         v_datXupdStartdate1, _
                                                         v_datXupdEnddate1, _
                                                         r_objRSExistingPersonalSched, _
                                                         r_rsNetPersonalSchedUpdates
                                If l_rdatOutdate1 = v_datXupdStartdate1 And l_rdatOutdate2 = v_datXupdEnddate1 Then 'if all these fields are equal, the record was deleted
                                Else
                                    l_rdatOutdate1 = v_datXupdStartdate1
                                    l_rdatOutdate2 = v_datXupdEnddate1
                                    ConstructCDIupdaterequests _
                                      r_rsNetPersonalSchedUpdates _
                                    , r_objRSExistingPersonalSched!EMP_NUM _
                                    , UPDATENDS_2 _
                                    , v_intRecordId _
                                    , r_objRSExistingPersonalSched!SCHED_CD _
                                    , r_objRSExistingPersonalSched!ROT_CD _
                                    , r_objRSExistingPersonalSched!ROT_WEEK _
                                    , l_rdatOutdate1 _
                                    , l_rdatOutdate2 _
                                    , l_rdatOutdate3 _
                                    , l_rdatOutdate4 _
                                    , _
                                    , r_objRSExistingPersonalSched!LAST_MOD_TIMESTAMP
                                    r_objRSExistingPersonalSched!RECORD_STATUS = "X"  'flag all deleted records
                                    r_objRSExistingPersonalSched.Update 'added to ensure delete
                                    'build DB2Request ("UPDATENDS")
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                If (v_datNewEnddate1 >= l_vdatOldStartdate2) _
                    And (v_datNewEnddate1 < l_vdatOldEnddate2) _
                    And (v_datNewStartdate1 <= l_vdatOldStartdate2) Then
                    'BuildNetModifyExistDates = TrCase_3
                    If SkipDELKeyForSameExistDates(v_lngEmpNum _
                                                 , l_vdatOldStartdate2 _
                                                 , l_vdatOldEnddate2 _
                                                 , r_rsNetPersonalSchedUpdates) Then 'process only if it hasn't already been deleted - this happens when auto creating "C" since it cannot be "D" and "C" for same existing DB key in a request
                       r_objRSExistingPersonalSched!RECORD_STATUS = "X"  'flag all deleted records
                       r_objRSExistingPersonalSched.Update 'added to ensure delete
                    Else
                        If IsChangeFoundSkipInsert(v_lngEmpNum _
                                                     , v_strRecordStatus _
                                                     , l_vdatOldStartdate2 _
                                                     , l_vdatOldEnddate2 _
                                                     , v_rsPSCloneRequest) Then 'this process is delayed if "C" is found for same dates
                        Else
                            If r_objRSExistingPersonalSched!RECORD_STATUS = "X" Then
                            Else
                                l_rdatOutdate1 = l_vdatOldStartdate2
                                l_rdatOutdate2 = l_vdatOldEnddate2
                                l_rdatOutdate3 = l_vdatOldStartdate2
                                l_rdatOutdate4 = l_vdatOldEnddate2
'                                IsExistUpdOverlapNewReqAdd v_lngEmpNum _
'                                                         , l_rdatOutdate1 _
'                                                         , l_rdatOutdate2 _
'                                                         , v_rsPSCloneRequest
                                
                                v_datXupdStartdate1 = l_vdatOldStartdate2
                                v_datXupdEnddate1 = l_vdatOldEnddate2
                                UpdExistPersfromInputReq v_lngEmpNum, _
                                                         v_intRecordId, _
                                                         l_rdatOutdate1, _
                                                         l_rdatOutdate2, _
                                                         v_rsPSCloneRequest, _
                                                         v_datXupdStartdate1, _
                                                         v_datXupdEnddate1, _
                                                         r_objRSExistingPersonalSched, _
                                                         r_rsNetPersonalSchedUpdates
                                If l_rdatOutdate1 = v_datXupdStartdate1 And l_rdatOutdate2 = v_datXupdEnddate1 Then 'if all these fields are equal, the record was deleted
                                Else
                                     l_rdatOutdate1 = v_datXupdStartdate1
                                     l_rdatOutdate2 = v_datXupdEnddate1
                                     If Val(r_objRSExistingPersonalSched!ROT_CD) > 0 Then
                                         If GetRotationScheduleCode(l_rdatOutdate1 _
                                             , r_objRSExistingPersonalSched!ROT_CD _
                                             , r_objRSExistingPersonalSched!START_DATE _
                                             , r_objRSExistingPersonalSched!ROT_WEEK _
                                             , l_intNewWeekPosition) = -1 Then
                                             l_intNewWeekPosition = r_objRSExistingPersonalSched!ROT_WEEK
                                         Else 'valid l_intNewWeekPosition was returned
                                         End If 'calculate new week postion for TrCase_3
                                     Else  'if rot_cd does not exist skip
                                     End If
                                     ConstructCDIupdaterequests _
                                       r_rsNetPersonalSchedUpdates _
                                     , r_objRSExistingPersonalSched!EMP_NUM _
                                     , UPDATENDS_2 _
                                     , v_intRecordId _
                                     , r_objRSExistingPersonalSched!SCHED_CD _
                                     , r_objRSExistingPersonalSched!ROT_CD _
                                     , l_intNewWeekPosition _
                                     , l_rdatOutdate1 _
                                     , l_rdatOutdate2 _
                                     , l_rdatOutdate3 _
                                     , l_rdatOutdate4 _
                                     , _
                                     , r_objRSExistingPersonalSched!LAST_MOD_TIMESTAMP
                                     r_objRSExistingPersonalSched!RECORD_STATUS = "X"  'flag all deleted records
                                     r_objRSExistingPersonalSched.Update 'added to ensure delete
                                    'build DB2Request ("UPDATENDS")
                                End If
                            End If
                        End If
                    End If
                Else
                    If (v_datNewStartdate1 >= l_vdatOldStartdate2) _
                       And (v_datNewEnddate1 <= l_vdatOldEnddate2) Then
                        'BuildNetModifyExistDates = TrCase_4 extending old start date not quite reaching old end date
                        If SkipDELKeyForSameExistDates(v_lngEmpNum _
                                                     , l_vdatOldStartdate2 _
                                                     , l_vdatOldEnddate2 _
                                                     , r_rsNetPersonalSchedUpdates) Then 'process only if it hasn't already been deleted - this happens when auto creating "C" since it cannot be "D" and "C" for same existing DB key in a request
                            r_objRSExistingPersonalSched!RECORD_STATUS = "X"  'flag all deleted records
                            r_objRSExistingPersonalSched.Update 'added to ensure delete
                        Else
                            If IsChangeFoundSkipInsert(v_lngEmpNum _
                                                         , v_strRecordStatus _
                                                         , l_vdatOldStartdate2 _
                                                         , l_vdatOldEnddate2 _
                                                         , v_rsPSCloneRequest) Then 'this process is delayed if "C" is found for same dates
                            Else
                                If r_objRSExistingPersonalSched!RECORD_STATUS = "X" Then
                                Else
                                    l_rdatOutdate1 = l_vdatOldStartdate2
                                    l_rdatOutdate2 = l_vdatOldEnddate2
                                    l_rdatOutdate3 = l_vdatOldStartdate2
                                    l_rdatOutdate4 = l_vdatOldEnddate2
'                                    IsExistUpdOverlapNewReqAdd v_lngEmpNum _
'                                                             , l_rdatOutdate1 _
'                                                             , l_rdatOutdate2 _
'                                                             , v_rsPSCloneRequest
                                 
                                    v_datXupdStartdate1 = l_vdatOldStartdate2
                                    v_datXupdEnddate1 = l_vdatOldEnddate2
                                    UpdExistPersfromInputReq v_lngEmpNum, _
                                                             v_intRecordId, _
                                                             l_rdatOutdate1, _
                                                             l_rdatOutdate2, _
                                                             v_rsPSCloneRequest, _
                                                             v_datXupdStartdate1, _
                                                             v_datXupdEnddate1, _
                                                             r_objRSExistingPersonalSched, _
                                                             r_rsNetPersonalSchedUpdates
                                    If l_rdatOutdate1 = v_datXupdStartdate1 And l_rdatOutdate2 = v_datXupdEnddate1 Then 'if all these fields are equal, the record was deleted
                                    Else
                                        If Val(r_objRSExistingPersonalSched!ROT_CD) > 0 Then
                                            If GetRotationScheduleCode(v_datXupdStartdate1 _
                                                , r_objRSExistingPersonalSched!ROT_CD _
                                                , r_objRSExistingPersonalSched!START_DATE _
                                                , r_objRSExistingPersonalSched!ROT_WEEK _
                                                , l_intNewWeekPosition) = -1 Then
                                                l_intNewWeekPosition = r_objRSExistingPersonalSched!ROT_WEEK
                                            Else 'valid l_intNewWeekPosition was returned
                                            End If 'calculate new week postion for TrCase_3
                                        Else  'if rot_cd does not exist skip
                                        End If
                                        ConstructCDIupdaterequests _
                                          r_rsNetPersonalSchedUpdates _
                                        , r_objRSExistingPersonalSched!EMP_NUM _
                                        , UPDATENDS_2 _
                                        , v_intRecordId _
                                        , r_objRSExistingPersonalSched!SCHED_CD _
                                        , r_objRSExistingPersonalSched!ROT_CD _
                                        , l_intNewWeekPosition _
                                        , v_datXupdStartdate1 _
                                        , l_rdatOutdate2 _
                                        , l_rdatOutdate3 _
                                        , l_rdatOutdate4 _
                                        , _
                                        , r_objRSExistingPersonalSched!LAST_MOD_TIMESTAMP
                                        'build DB2Request ("UPDATENDS") 'replace with adj l_rdatOutdate2
                                        r_objRSExistingPersonalSched!RECORD_STATUS = "X"  'flag all deleted records
                                        r_objRSExistingPersonalSched.Update 'added to ensure delete
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If (v_datNewStartdate1 <= l_vdatOldStartdate2) _
                            And (v_datNewEnddate1 >= l_vdatOldEnddate2) Then
                            'BuildNetModifyExistDates = TrCase_5
                            l_rdatOutdate1 = l_vdatOldStartdate2
                            l_rdatOutdate2 = l_vdatOldEnddate2
                            ConstructCDIupdaterequests _
                              r_rsNetPersonalSchedUpdates _
                            , r_objRSExistingPersonalSched!EMP_NUM _
                            , DELETENDS_1 _
                            , v_intRecordId _
                            , , , _
                            , l_rdatOutdate1 _
                            , l_rdatOutdate2 _
                            , , , _
                            , r_objRSExistingPersonalSched!LAST_MOD_TIMESTAMP
                            'build DB2Request ("DELETENDS")
                        Else
                            'BuildNetModifyExistDates = TrCase_6 - inserts new record is done after this routine is completed
                        End If
                    End If
                End If
            End If
        End If
        r_objRSExistingPersonalSched.MoveNext
    Loop
    
    If IsClearedFilterRecordsetEmpty(v_rsPSCloneRequest) Then
    End If
    
    If IsClearedFilterRecordsetEmpty(r_objRSExistingPersonalSched) Then
    End If
    
    
Exit Function

BuildNetModifyExistDates_EH:
    TrsRaiseError mstrCMODULENAME, "BuildNetModifyExistDates"

        
End Function

'building the insert, update and delete NET recordset
Private Function ConstructCDIupdaterequests( _
    ByRef r_rsRequest As ADODB.Recordset, _
    ByVal v_lngEmpNum As Long, _
    ByVal v_lngTypeOfUpdate As TrTypeOfUpdate, _
    ByVal v_intRecordId As Integer, _
    Optional ByVal v_intSchedCd As Integer, _
    Optional ByVal v_lngRotCd As Long, _
    Optional ByVal v_lngStartPos As Long, _
    Optional ByVal v_datDate1 As Date, _
    Optional ByVal v_datDate2 As Date, _
    Optional ByVal v_datdate3 As Date, _
    Optional ByVal v_datdate4 As Date, _
    Optional ByVal v_strTimestamp As String, _
    Optional ByVal v_strLastModTimestamp As String) As Boolean

'6/10/02 RSG Added v_strLastModTimestamp, represents the concurrency key
'for optimistic locking. Required for deletes and updates.
On Error GoTo ConstructCDIupdaterequests_EH

    Dim l_strFilter As String

    If v_lngTypeOfUpdate = INSERTNDS_3 Then
        r_rsRequest.AddNew
        r_rsRequest.Fields("EMP_NUM") = CStr(v_lngEmpNum)
        r_rsRequest.Fields("SCHED_CD") = CStr(v_intSchedCd)
        r_rsRequest.Fields("ROT_CD") = CStr(v_lngRotCd)
        r_rsRequest.Fields("ROT_WEEK") = CStr(v_lngStartPos)
        r_rsRequest.Fields("START_DATE") = v_datDate1
        r_rsRequest.Fields("END_DATE") = v_datDate2
        r_rsRequest.Fields("RECORD_STATUS") = "I"
        r_rsRequest.Fields("CREATED_TIMESTAMP") = CStr(v_strTimestamp)
        r_rsRequest.Fields("LAST_MOD_TIMESTAMP") = CStr(v_strTimestamp)
        r_rsRequest.Fields("CREATED_EMP_NUM") = CStr(m_lngCreatedUserNo)
        r_rsRequest.Fields("LAST_MOD_EMP_NUM") = CStr(m_lngUserId)
        r_rsRequest.Fields("RECORD_ID") = CStr(v_intRecordId)
        r_rsRequest.Update
    Else
        If v_lngTypeOfUpdate = UPDATENDS_2 Then
            r_rsRequest.AddNew
            r_rsRequest.Fields("SCHED_CD") = CStr(v_intSchedCd)
            r_rsRequest.Fields("ROT_CD") = CStr(v_lngRotCd)
            r_rsRequest.Fields("ROT_WEEK") = CStr(v_lngStartPos)
            r_rsRequest.Fields("START_DATE") = v_datDate1
            r_rsRequest.Fields("END_DATE") = v_datDate2
            r_rsRequest.Fields("LAST_MOD_EMP_NUM") = CStr(m_lngUserId)
            r_rsRequest.Fields("KEY_EMP_NUM") = CStr(v_lngEmpNum)
            r_rsRequest.Fields("KEY_START_DATE") = v_datdate3
            r_rsRequest.Fields("KEY_END_DATE") = v_datdate4
            r_rsRequest.Fields("RECORD_STATUS") = "C"
            r_rsRequest.Fields("RECORD_ID") = CStr(v_intRecordId)
            r_rsRequest.Fields("LAST_MOD_TIMESTAMP") = v_strLastModTimestamp
            r_rsRequest.Update
        Else
        '6/14/02 RSG Check for duplicate delete before adding.
        'Duplicate causes a commit conflict error since second request is not found.
            l_strFilter = "KEY_EMP_NUM = " & CStr(v_lngEmpNum) & _
                            " AND KEY_START_DATE = #" & Format(v_datDate1, "mm/dd/yyyy") & _
                            "# AND LAST_MOD_TIMESTAMP = '" & Trim(v_strLastModTimestamp) & _
                            "' AND RECORD_STATUS = 'D'"
            l_strFilter = MDAC25FilterFix(l_strFilter)
            r_rsRequest.Filter = l_strFilter
            If Not RecordsExist(r_rsRequest) Then
                r_rsRequest.AddNew
                r_rsRequest.Fields("KEY_EMP_NUM") = CStr(v_lngEmpNum)
                r_rsRequest.Fields("KEY_START_DATE") = v_datDate1
                r_rsRequest.Fields("KEY_END_DATE") = v_datDate2
                r_rsRequest.Fields("RECORD_STATUS") = "D"
                r_rsRequest.Fields("RECORD_ID") = CStr(v_intRecordId)
                r_rsRequest.Fields("LAST_MOD_TIMESTAMP") = v_strLastModTimestamp
                r_rsRequest.Update
            End If
            r_rsRequest.Filter = adFilterNone
        End If
    End If
    
Exit Function

ConstructCDIupdaterequests_EH:
    TrsRaiseError mstrCMODULENAME, "ConstructCDIupdaterequests"


End Function



'call to get Rotation header information
Private Function GetRotationSchedHeader _
                (Optional ByVal v_intRotationCode As Integer _
                ) As ADODB.Recordset

On Error GoTo GetRotationSchedHeader_EH

    Dim l_busSchedRotation As busSchedRotation.cbsSchedRotation

    Set l_busSchedRotation = CreateObject("busSchedRotation.cbsSchedRotation")

    Set GetRotationSchedHeader = l_busSchedRotation.GetRotationSchedHeader(v_intRotationCode)

    Set l_busSchedRotation = Nothing
    GetObjectContext.SetComplete

Exit Function

GetRotationSchedHeader_EH:
    Set l_busSchedRotation = Nothing
    GetObjectContext.SetAbort
    TrsRaiseError mstrCMODULENAME, "GetRotationSchedHeader"

End Function
'call to get Rotation Detail information
Private Function GetRotationSchedDetail _
                (Optional ByVal v_intRotationCode As Integer, _
                 Optional ByVal v_intRotationOrder As Integer _
                ) As ADODB.Recordset

    On Error GoTo GetRotationSchedDetail_EH

    Dim l_busSchedRotation As busSchedRotation.cbsSchedRotation

    Set l_busSchedRotation = CreateObject("busSchedRotation.cbsSchedRotation")

    Set GetRotationSchedDetail = l_busSchedRotation.GetRotationSchedDetail(v_intRotationCode, v_intRotationOrder)

    Set l_busSchedRotation = Nothing
    GetObjectContext.SetComplete

Exit Function

GetRotationSchedDetail_EH:
    Set l_busSchedRotation = Nothing
    GetObjectContext.SetAbort
    TrsRaiseError mstrCMODULENAME, "GetRotationSchedDetail"

End Function
'detect any updates to be made
Private Function HasAnyPersonalfieldChanged( _
    ByVal v_rsPersonalScheduleRequest As ADODB.Recordset, _
    ByRef r_rsErrors As ADODB.Recordset) As Boolean
        
On Error GoTo HasAnyPersonalfieldChanged_EH

    Dim l_objDBPersonalSched As ADODB.Recordset
    Dim l_strFilter As String


    If IsClearedFilterRecordsetEmpty(v_rsPersonalScheduleRequest) Then
        HasAnyPersonalfieldChanged = False
        Exit Function
    Else
    End If


    l_strFilter = "RECORD_STATUS = 'I'" & _
              " OR RECORD_STATUS = 'D'"

       
    If IsRecordsInFilterRecordSet(l_strFilter, v_rsPersonalScheduleRequest) Then
        HasAnyPersonalfieldChanged = True
        Exit Function
    Else
    End If



    l_strFilter = "RECORD_STATUS = 'C'"
    If IsRecordsInFilterRecordSet(l_strFilter, v_rsPersonalScheduleRequest) Then
    Else
        Exit Function
    End If

    Do Until v_rsPersonalScheduleRequest.EOF
'        Set l_objDBPersonalSched = GetScopePersonalSched(v_rsPersonalScheduleRequest!EMP_NUM, v_rsPersonalScheduleRequest!KEY_START_DATE, v_rsPersonalScheduleRequest!KEY_END_DATE) 'Requested keys are used to get original request to be compared with any changes
        Set l_objDBPersonalSched = GetExistPersonalSched(v_rsPersonalScheduleRequest!EMP_NUM, v_rsPersonalScheduleRequest!KEY_START_DATE, v_rsPersonalScheduleRequest!KEY_END_DATE)  'Requested keys are used to get original request to be compared with any changes
        If RecordsExist(l_objDBPersonalSched) Then
        Else
            ' trErrbusSchedPersonalNotFound           l_strRotationError = "Rotation Detail - Record does not exist "
            m_lngErrorNo = trErrbusSchedPersonalNotFound
            AppendPersonalError r_rsErrors, TrType_PER, m_lngRecordId, "START_DATE", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), v_rsPersonalScheduleRequest!KEY_EMP_NUM, v_rsPersonalScheduleRequest!KEY_START_DATE, v_rsPersonalScheduleRequest!KEY_END_DATE
            HasAnyPersonalfieldChanged = True 'even though this is an error exit normally but check r_rsErrors for brvs after this rtn to setabort and set error brv flag for client - pts 1002
            Exit Function
        End If
        If l_objDBPersonalSched!START_DATE = v_rsPersonalScheduleRequest!START_DATE Then
        Else
            HasAnyPersonalfieldChanged = True
            If IsClearedFilterRecordsetEmpty(v_rsPersonalScheduleRequest) Then
            End If
            Exit Function
        End If
        If l_objDBPersonalSched!END_DATE = v_rsPersonalScheduleRequest!END_DATE Then
        Else
            HasAnyPersonalfieldChanged = True
            If IsClearedFilterRecordsetEmpty(v_rsPersonalScheduleRequest) Then
            End If
            Exit Function
        End If
        If l_objDBPersonalSched!ROT_CD = v_rsPersonalScheduleRequest!ROT_CD Then
        Else
            HasAnyPersonalfieldChanged = True
            If IsClearedFilterRecordsetEmpty(v_rsPersonalScheduleRequest) Then
            End If
            Exit Function
        End If
        If l_objDBPersonalSched!ROT_WEEK = v_rsPersonalScheduleRequest!ROT_WEEK Then
        Else
            HasAnyPersonalfieldChanged = True
            If IsClearedFilterRecordsetEmpty(v_rsPersonalScheduleRequest) Then
            End If
            Exit Function
        End If
        If l_objDBPersonalSched!SCHED_CD = v_rsPersonalScheduleRequest!SCHED_CD Then
        Else
            HasAnyPersonalfieldChanged = True
            If IsClearedFilterRecordsetEmpty(v_rsPersonalScheduleRequest) Then
            End If
            Exit Function
        End If
        v_rsPersonalScheduleRequest.MoveNext
    Loop
    
    If IsClearedFilterRecordsetEmpty(v_rsPersonalScheduleRequest) Then
    End If
    
Exit Function
HasAnyPersonalfieldChanged_EH:
    Set l_objDBPersonalSched = Nothing
    TrsRaiseError mstrCMODULENAME, "HasAnyPersonalfieldChanged"

End Function

'should this be used for "I"?
'this is used to determine the end date if is null or not supplied by the request
'this routine includes the input request to get the net effect
'Deals with existing records so use keys
Private Function GetDBDefaultEndDate _
                (ByVal v_lngEmpNum As Long _
                 , ByVal v_datStartDate As Date _
                 , ByVal v_rsPSCloneRequest As ADODB.Recordset) As Date

    On Error GoTo GetDBDefaultEndDate_EH
    
    Dim l_rsDBDefaultEndDate As ADODB.Recordset
    Dim l_strFilter As String
    Dim l_datLowestEndDate As Date
    Dim l_datnewdfltStartDate As Date
    
'    'check for existing key if it exist get next defaultenddate with the start date of existing key
'    'db eof check only for request clone: is the db start date and end date in the request for changes, if not present there are no changes "C"
'    'search for default date record which is not going to be deleted and if
'    'not found default to optional dates
        'perform loop to keep searching until no more recs are available
    l_datnewdfltStartDate = v_datStartDate
    Do Until IsNoDelDBRecords(v_lngEmpNum, v_rsPSCloneRequest, l_datnewdfltStartDate, l_rsDBDefaultEndDate) = True
    Loop
    
    l_strFilter = "KEY_START_DATE = #" & l_rsDBDefaultEndDate!START_DATE & "# " & _
             " AND KEY_END_DATE = #" & l_rsDBDefaultEndDate!END_DATE & "# " & _
             " AND KEY_EMP_NUM = " & v_lngEmpNum & _
             " AND RECORD_STATUS = 'C'"
    
    If IsRecordsInFilterRecordSet(l_strFilter, v_rsPSCloneRequest) Then 'db matched req keys
    Else
       If l_rsDBDefaultEndDate!START_DATE = m_datCOPTIONALSTARTDATE And _
          l_rsDBDefaultEndDate!END_DATE = m_datCOPTIONALENDDATE Then 'are returned 1/1/100 and 12/31/9999?
          'db eof returned use 12/31/9999
          GetDBDefaultEndDate = m_datCOPTIONALENDDATE
          If IsClearedFilterRecordsetEmpty(v_rsPSCloneRequest) Then
          End If
          Exit Function
       Else
          'different dates were returned that may be used as end date
       End If
    End If
    l_datLowestEndDate = DateAdd("d", -1, l_rsDBDefaultEndDate!START_DATE) 'search for all requested dates with equal or greater than db date
    'check to see if there are any new req dates greater than the one from the db so we can eliminate gaps
    l_strFilter = "START_DATE > #" & l_datLowestEndDate & "#" & _
             " AND EMP_NUM = " & v_lngEmpNum & _
             " AND RECORD_STATUS = 'C'" 'search for all requested dates with equal or greater than db date
    GetDBDefaultEndDate = FindLowestDate(l_strFilter _
                        , l_rsDBDefaultEndDate!START_DATE _
                        , v_rsPSCloneRequest)
       
    
    If IsClearedFilterRecordsetEmpty(v_rsPSCloneRequest) Then
    End If

    GetObjectContext.SetComplete

Exit Function

GetDBDefaultEndDate_EH:
    GetObjectContext.SetAbort
    TrsRaiseError mstrCMODULENAME, "GetDBDefaultEndDate"

End Function

'this is used to determine the end date if is null or not supplied by the request
'this routine includes the input request to get the net effect
Private Function FindLowestDate _
                 (ByVal v_strFilter As String _
                 , ByVal v_datDefaultEndDate As Date _
                 , ByVal v_rsPSCloneRequest As ADODB.Recordset) As Date

On Error GoTo FindLowestDate_EH
    
    Dim l_datDBEndDate As Date
    Dim l_datLowestEndDate As Date
    
    If IsRecordsInFilterRecordSet(v_strFilter, v_rsPSCloneRequest) Then '
       If v_rsPSCloneRequest.RecordCount > 1 Then
          l_datDBEndDate = v_rsPSCloneRequest!START_DATE
          Do Until v_rsPSCloneRequest.EOF
              If l_datDBEndDate > v_rsPSCloneRequest!START_DATE Then
                 l_datDBEndDate = v_rsPSCloneRequest!START_DATE
              End If
              v_rsPSCloneRequest.MoveNext
          Loop
          l_datLowestEndDate = l_datDBEndDate
          FindLowestDate = DateAdd("d", -1, l_datLowestEndDate) 'use datelastmodified instead of create date
       Else
          FindLowestDate = DateAdd("d", -1, v_rsPSCloneRequest!START_DATE) 'use datelastmodified instead of create date
       End If
    Else
       FindLowestDate = DateAdd("d", -1, v_datDefaultEndDate)  'use database start date
    End If

    If IsClearedFilterRecordsetEmpty(v_rsPSCloneRequest) Then
    End If
   
Exit Function

FindLowestDate_EH:
    
    GetObjectContext.SetAbort
    
    TrsRaiseError mstrCMODULENAME, "FindLowestDate"

End Function

'for change routine IsBUSChangePersSched - this is unlikely to occur since there can only be one key displayed at one time
'this will flag a brv if duplicate key records with same emp_num, start date and end date
Private Function IsduplicateDateKeysinrequest _
                (ByVal v_lngEmpNum As Long _
                 , ByVal v_datStartDate As Date _
                 , ByVal v_datEndDate As Date _
                 , ByVal v_rsPSCloneRequest As ADODB.Recordset _
                 ) As Boolean
        
On Error GoTo IsduplicateDateKeysinrequest_EH

    Dim l_strFilter As String
    
    If IsClearedFilterRecordsetEmpty(v_rsPSCloneRequest) Then
        IsduplicateDateKeysinrequest = False
        Exit Function
    Else
    End If
    
    
    
    'db eof check only for request clone: is the db start date and end date in the request for changes, if not present there are no changes "C"
    l_strFilter = "KEY_START_DATE = #" & v_datStartDate & "# " & _
             " AND KEY_END_DATE = #" & v_datEndDate & "# " & _
             " AND KEY_EMP_NUM = " & v_lngEmpNum & _
             " AND RECORD_STATUS = 'C'"
    If IsRecordsInFilterRecordSet(l_strFilter, v_rsPSCloneRequest) Then 'db matched req keys
       If v_rsPSCloneRequest.RecordCount > 1 Then
          IsduplicateDateKeysinrequest = True
          'yes continue process
       Else
          IsduplicateDateKeysinrequest = False
       End If
    Else
        IsduplicateDateKeysinrequest = False
    End If
        
    If IsClearedFilterRecordsetEmpty(v_rsPSCloneRequest) Then
    End If
        

    
Exit Function
IsduplicateDateKeysinrequest_EH:
    
    TrsRaiseError mstrCMODULENAME, "IsduplicateDateKeysinrequest"

End Function
        
'for change routine IsBUSChangePersSched
'this will flag a brv if duplicate records with same emp_num, start date and end date
'two existing db records could be modified to contain the same dates
Private Function IsduplicateDateChangedrequest _
                (ByVal v_lngEmpNum As Long _
                 , ByVal v_datStartDate As Date _
                 , ByVal v_datEndDate As Date _
                 , ByVal v_rsPSCloneRequest As ADODB.Recordset _
                 ) As Boolean
        
On Error GoTo IsduplicateDateChangedrequest_EH

    Dim l_strFilter As String
    
    If IsClearedFilterRecordsetEmpty(v_rsPSCloneRequest) Then
        IsduplicateDateChangedrequest = False
        Exit Function
    Else
    End If
    
    
    'db eof check only for request clone: is the db start date and end date in the request for changes, if not present there are no changes "C"
    l_strFilter = "START_DATE = #" & v_datStartDate & "# " & _
             " AND END_DATE = #" & v_datEndDate & "# " & _
             " AND EMP_NUM = " & v_lngEmpNum & _
             " AND RECORD_STATUS = 'C'"
    If IsRecordsInFilterRecordSet(l_strFilter, v_rsPSCloneRequest) Then 'db matched req keys
       If v_rsPSCloneRequest.RecordCount > 1 Then
          IsduplicateDateChangedrequest = True
          'yes continue process
       Else
          IsduplicateDateChangedrequest = False
       End If
    Else
        IsduplicateDateChangedrequest = False
    End If

    If IsClearedFilterRecordsetEmpty(v_rsPSCloneRequest) Then
    End If
    

    
Exit Function
IsduplicateDateChangedrequest_EH:
    TrsRaiseError mstrCMODULENAME, "IsduplicateDateChangedrequest"

End Function

'for change routine IsBUSChangePersSched and
'    insert routine IsBUSInsertPersSched
'Business rules for delete and date overlap
Private Function IsBUSOverlapInRequest(ByVal v_rsPSCloneRequest As ADODB.Recordset _
                                    , ByVal v_lngEmpNum As Long _
                                    , ByVal v_strRecordStatus As String _
                                    , ByVal v_datStartDate As Date _
                                    , ByVal v_datEndDate As Date _
                                    , ByVal v_datKeyStartDate As Date _
                                    , ByVal v_datKeyEndDate As Date _
                                    , Optional ByRef r_strDupFoundSwitch As String _
                                      ) As Boolean

On Error GoTo IsBUSOverlapInRequest_EH
    
    Dim l_objDBPSClone As ADODB.Recordset

    Dim l_strFilter As String
    Dim l_strFilterDups As String
    
    IsBUSOverlapInRequest = False
    
    
    If r_strDupFoundSwitch = "" Then
        r_strDupFoundSwitch = "N"
    End If
    
    'process all records to be deleted
    If IsClearedFilterRecordsetEmpty(v_rsPSCloneRequest) Then
        Exit Function
    Else
    End If
    
    l_strFilterDups = "(START_DATE >= #" & v_datStartDate & "#" & _
             " AND START_DATE <= #" & v_datEndDate & "#" & _
             " AND RECORD_STATUS = '" & v_strRecordStatus & "'" & _
             " AND EMP_NUM = " & v_lngEmpNum & ")" & _
             " OR (END_DATE >= #" & v_datStartDate & "#" & _
             " AND END_DATE <= #" & v_datEndDate & "#" & _
             " AND RECORD_STATUS = '" & v_strRecordStatus & "'" & _
             " AND EMP_NUM = " & v_lngEmpNum & ")"
             
    If IsRecordsInFilterRecordSet(l_strFilterDups, v_rsPSCloneRequest) Then 'Any dups for same update function?
        If v_rsPSCloneRequest.RecordCount > 1 Then 'insert or changed dups found
           r_strDupFoundSwitch = "Y"
           IsBUSOverlapInRequest = True
           If IsClearedFilterRecordsetEmpty(v_rsPSCloneRequest) Then
           End If
           Exit Function
        Else
        End If
    End If
    
    
    'process inserts / change overlaps
    l_strFilter = "(START_DATE >= #" & v_datStartDate & "#" & _
             " AND START_DATE <= #" & v_datEndDate & "#" & _
             " AND RECORD_STATUS <> 'D'" & _
             " AND EMP_NUM = " & v_lngEmpNum & ")" & _
             " OR (END_DATE >= #" & v_datStartDate & "#" & _
             " AND END_DATE <= #" & v_datEndDate & "#" & _
             " AND RECORD_STATUS <> 'D'" & _
             " AND EMP_NUM = " & v_lngEmpNum & ")"
    If IsRecordsInFilterRecordSet(l_strFilter, v_rsPSCloneRequest) Then 'updating key fields if deleting
    Else
        Exit Function
    End If
    
    If v_rsPSCloneRequest.RecordCount > 1 Then
    Else
        Exit Function
    End If
    
    
    Do Until v_rsPSCloneRequest.EOF
        If v_rsPSCloneRequest!RECORD_STATUS = "C" Then
            If CDate(v_rsPSCloneRequest!START_DATE) = CDate(v_datStartDate) _
                And CDate(v_rsPSCloneRequest!END_DATE) = CDate(v_datEndDate) _
                And CDate(v_rsPSCloneRequest!KEY_START_DATE) = CDate(v_datKeyStartDate) _
                And CDate(v_rsPSCloneRequest!KEY_END_DATE) = CDate(v_datKeyEndDate) _
                And v_rsPSCloneRequest!RECORD_STATUS = v_strRecordStatus _
                And v_rsPSCloneRequest!EMP_NUM = v_lngEmpNum Then 'skip same C
            Else
                IsBUSOverlapInRequest = True
                If IsClearedFilterRecordsetEmpty(v_rsPSCloneRequest) Then
                End If
                Exit Function
            End If
        Else
            If v_rsPSCloneRequest!RECORD_STATUS = "I" Then
                If CDate(v_rsPSCloneRequest!START_DATE) = CDate(v_datStartDate) _
                    And CDate(v_rsPSCloneRequest!END_DATE) = CDate(v_datEndDate) _
                    And v_rsPSCloneRequest!RECORD_STATUS = v_strRecordStatus _
                    And v_rsPSCloneRequest!EMP_NUM = v_lngEmpNum Then 'skip same I
                Else
                    IsBUSOverlapInRequest = True
                    If IsClearedFilterRecordsetEmpty(v_rsPSCloneRequest) Then
                    End If
                    Exit Function
                End If
            Else
            End If
        End If
        v_rsPSCloneRequest.MoveNext
    Loop

    If IsClearedFilterRecordsetEmpty(v_rsPSCloneRequest) Then
    End If


Exit Function

IsBUSOverlapInRequest_EH:
    
    TrsRaiseError mstrCMODULENAME, "IsBUSOverlapInRequest"

End Function
        



'for change routine only IsBUSChangePersSched
'do not process as input for BuildNetModifyExistDates - build DB2Request ("UPDATENDS") since it will be deleted by key
'skip (from establish add) existing input db record since it will be deleted
'if existing key record found in database so use start date instead of key start date therefore delete so it won't be used by requested changed key record
'for changes with same key only (not use by inserts
Private Function SkipExistEqualKeyDateReq( _
    ByVal v_lngEmpNum As Long, _
    ByVal v_datKeyStartdate1 As Date, _
    ByVal v_datKeyEnddate1 As Date, _
    ByRef r_objRSExistingPersonalSched As ADODB.Recordset) _
        As Boolean

On Error GoTo SkipExistEqualKeyDateReq_EH

    Dim l_strFilter As String
    
    If IsClearedFilterRecordsetEmpty(r_objRSExistingPersonalSched) Then
        Exit Function
    Else
    End If
    
    
   l_strFilter = "START_DATE = #" & v_datKeyStartdate1 & "#" & _
            " AND END_DATE = #" & v_datKeyEnddate1 & "#" & _
            " AND EMP_NUM = " & v_lngEmpNum


    If IsRecordsInFilterRecordSet(l_strFilter, r_objRSExistingPersonalSched) Then 'updating key fields if deleting
        r_objRSExistingPersonalSched!RECORD_STATUS = "X"  'flag all existing db skipped records
        r_objRSExistingPersonalSched.Update 'added to ensure delete
        r_objRSExistingPersonalSched.MoveNext
    Else
    End If

    If IsClearedFilterRecordsetEmpty(r_objRSExistingPersonalSched) Then
    End If

    
Exit Function

SkipExistEqualKeyDateReq_EH:
    GetObjectContext.SetAbort
    TrsRaiseError mstrCMODULENAME, "SkipExistEqualKeyDateReq"

    
End Function
'if record was deleted or already processed, mark it as "X" (already processed) so it is a more efficient processing
'for BuildNetModifyExistDates routine only 'build DB2Request ("UPDATENDS")
'this existing records are skipped and won't be processed by any auto updates by "C" or "I"
'if existing key record found in request as deleted so use  key start date therefore delete so it won't be added as a changed
Private Function SkipDELKeyForSameExistDates( _
    ByVal v_lngEmpNum As Long, _
    ByVal v_datKeyStartdate1 As Date, _
    ByVal v_datKeyEnddate1 As Date, _
    ByRef r_rsNetPersonalSchedUpdates As ADODB.Recordset) _
        As Boolean

On Error GoTo SkipDELKeyForSameExistDates_EH

    Dim l_strFilter As String
    
    If IsClearedFilterRecordsetEmpty(r_rsNetPersonalSchedUpdates) Then
        Exit Function
    Else
    End If
    
    
   l_strFilter = "(KEY_START_DATE = #" & v_datKeyStartdate1 & "#" & _
                 " AND KEY_END_DATE = #" & v_datKeyEnddate1 & "#" & _
                 " AND KEY_EMP_NUM = " & v_lngEmpNum & _
                 " AND RECORD_STATUS = 'D')" & _
                 " OR (KEY_START_DATE = #" & v_datKeyStartdate1 & "#" & _
                 " AND KEY_END_DATE = #" & v_datKeyEnddate1 & "#" & _
                 " AND KEY_EMP_NUM = " & v_lngEmpNum & _
                 " AND RECORD_STATUS = 'C')"

    If IsRecordsInFilterRecordSet(l_strFilter, r_rsNetPersonalSchedUpdates) Then 'updating key fields if deleting
        SkipDELKeyForSameExistDates = True
    Else
        SkipDELKeyForSameExistDates = False
    End If

    If IsClearedFilterRecordsetEmpty(r_rsNetPersonalSchedUpdates) Then
    End If

    
Exit Function

SkipDELKeyForSameExistDates_EH:
    GetObjectContext.SetAbort
    TrsRaiseError mstrCMODULENAME, "SkipDELKeyForSameExistDates"

    
End Function

''for BuildNetModifyExistDates routine only 'build DB2Request ("UPDATENDS")
''this provides the net effect for updates and insert of db recs: dates are generated using the input high and low dates that will modify this existing db record
''(if there is a "C" and "I", the "C" will take precedence over the "I" by performing this routine since we are using the processing order D-I-C
''Update dates may be affected by new insert date range, so modified accordingly
''this routine provides the Net effect using "C" and "I")
'Private Function IsExistUpdOverlapNewReqAdd( _
'    ByVal v_lngEmpNum As Long, _
'    ByRef v_datXupdStartdate1 As Date, _
'    ByRef v_datXupdEnddate1 As Date, _
'    ByVal v_rsPSCloneRequest As ADODB.Recordset) _
'        As Boolean
'
'On Error GoTo IsExistUpdOverlapNewReqAdd_EH
'
'    Dim l_strFilter As String
'    Dim l_datMinREQdStartdate As Date
'    Dim l_datMaxREQdEnddate As Date
'
'    If IsClearedFilterRecordsetEmpty(v_rsPSCloneRequest) Then
'        Exit Function
'    Else
'    End If
'
'
''delete = overlap generated by previous transactions I or C
''if the key is found in the other modified dates touched by a upd existing start or end date
'     l_strFilter = "(START_DATE >= #" & v_datXupdStartdate1 & "#" & _
'              " AND START_DATE <= #" & v_datXupdEnddate1 & "#" & _
'              " AND RECORD_STATUS <> 'D')" & _
'              " OR (END_DATE >= #" & v_datXupdStartdate1 & "#" & _
'              " AND END_DATE <= #" & v_datXupdEnddate1 & "#" & _
'              " AND EMP_NUM = " & v_lngEmpNum & _
'              " AND RECORD_STATUS <> 'D')"
'
''process was improved by getting earliest and latest requested input dates first then applying to existing recs
'    If IsRecordsInFilterRecordSet(l_strFilter, v_rsPSCloneRequest) Then 'updating key fields if deleting
'        l_datMinREQdStartdate = v_rsPSCloneRequest!START_DATE
'        l_datMaxREQdEnddate = v_rsPSCloneRequest!END_DATE
'        Do Until v_rsPSCloneRequest.EOF  'find earlier start date and latest end date
'            If CDate(l_datMinREQdStartdate) > v_rsPSCloneRequest!START_DATE Then
'               l_datMinREQdStartdate = v_rsPSCloneRequest!START_DATE
'            End If
'            If CDate(l_datMaxREQdEnddate) < v_rsPSCloneRequest!END_DATE Then
'               l_datMaxREQdEnddate = v_rsPSCloneRequest!END_DATE
'            End If
'            v_rsPSCloneRequest.MoveNext
'        Loop
'        If CDate(v_datXupdStartdate1) < CDate(v_datXupdEnddate1) Then
'           If CDate(l_datMinREQdStartdate) <= CDate(v_datXupdEnddate1) Then
'              If CDate(l_datMinREQdStartdate) > CDate(v_datXupdStartdate1) Then
'                  v_datXupdEnddate1 = DateAdd("d", -1, l_datMinREQdStartdate)  'for update portion end date of TrCase_2 and TrCase_4
'              Else
'                  If CDate(l_datMaxREQdEnddate) < CDate(v_datXupdEnddate1) Then
'                      v_datXupdStartdate1 = DateAdd("d", 1, l_datMaxREQdEnddate) 'for update portion start date of TrCase_3 and insert of TrCase_4
'                  End If
'              End If
'           End If
'        End If
'    End If
'
'    If IsClearedFilterRecordsetEmpty(v_rsPSCloneRequest) Then
'    End If
'
'
'Exit Function
'
'IsExistUpdOverlapNewReqAdd_EH:
'    GetObjectContext.SetAbort
'    TrsRaiseError mstrCMODULENAME, "IsExistUpdOverlapNewReqAdd"
'
'
'End Function
'DELAY PROCESS FOR "C" 'this process is delayed if "C" is found for same dates
'for BuildNetModifyExistDates routine only 'build DB2Request ("UPDATENDS")
'when I for existing check if there is a "C" key in request file
'if a "C" exist and an "I" needs to process an existing record with the "C" key, skip "I" request and let the "C" request process it later since we are using the processing order D-I-C
Private Function IsChangeFoundSkipInsert( _
    ByVal v_lngEmpNum As Long, _
    ByVal v_strRecordStatus As String, _
    ByVal v_datKeyStartdate1 As Date, _
    ByVal v_datKeyEnddate1 As Date, _
    ByVal v_rsPSCloneRequest As ADODB.Recordset) _
        As Boolean

On Error GoTo IsChangeFoundSkipInsert_EH

    Dim l_strFilter As String
    
    If v_strRecordStatus = "I" Then
    Else
        IsChangeFoundSkipInsert = False
        Exit Function
    End If
    
    l_strFilter = "RECORD_STATUS = 'C' "
    If IsRecordsInFilterRecordSet(l_strFilter, v_rsPSCloneRequest) Then 'updating key fields if deleting
    Else
        IsChangeFoundSkipInsert = False
        Exit Function
    End If
    
    l_strFilter = "KEY_START_DATE = #" & v_datKeyStartdate1 & "#" & _
                  " AND KEY_END_DATE = #" & v_datKeyEnddate1 & "#" & _
                  " AND RECORD_STATUS = 'C' " & _
                  " AND KEY_EMP_NUM = " & v_lngEmpNum

    If IsRecordsInFilterRecordSet(l_strFilter, v_rsPSCloneRequest) Then 'updating key fields if deleting
        IsChangeFoundSkipInsert = True
    Else
        IsChangeFoundSkipInsert = False
    End If

    If IsClearedFilterRecordsetEmpty(v_rsPSCloneRequest) Then
    End If
    
Exit Function

IsChangeFoundSkipInsert_EH:
    GetObjectContext.SetAbort
    TrsRaiseError mstrCMODULENAME, "IsChangeFoundSkipInsert"

    
End Function


'for change routine only IsBUSChangePersSched
'allows to see if Personal record nonkey exist, if it does no brv is returned since no changes were made
'*returns by ref the Personal schedule for comparison purposes
Private Function IsExistNonkeyFieldChangeOnly _
                (ByVal v_lngEmpNum As Long _
                 , ByVal v_datStartDate As Date _
                 , ByVal v_datEndDate As Date _
                 , ByVal v_datKeyStartDate As Date _
                 , ByVal v_datKeyEndDate As Date _
                 , ByRef r_rsDBPersonalSchedExist As ADODB.Recordset _
                 ) As Boolean
        
On Error GoTo IsExistNonkeyFieldChangeOnly_EH
    
    IsExistNonkeyFieldChangeOnly = False
    If v_datStartDate = v_datKeyStartDate _
        And v_datEndDate = v_datKeyEndDate Then
        Set r_rsDBPersonalSchedExist = GetExistPersonalSched(v_lngEmpNum, v_datKeyStartDate, v_datKeyEndDate)  'need to find the nonkey fields to see if needs updating
        If RecordsExist(r_rsDBPersonalSchedExist) Then
            IsExistNonkeyFieldChangeOnly = True
        Else
            IsExistNonkeyFieldChangeOnly = False
        End If
    Else
        IsExistNonkeyFieldChangeOnly = False
    End If

    If IsClearedFilterRecordsetEmpty(r_rsDBPersonalSchedExist) Then
    End If

    
Exit Function
IsExistNonkeyFieldChangeOnly_EH:
    TrsRaiseError mstrCMODULENAME, "IsExistNonkeyFieldChangeOnly"

End Function



'for change routine only IsBUSChangePersSched
Private Function GetExistingPersonalSchedule( _
                    ByVal v_lngEmpNum As Long, _
                    ByVal v_datKeyStartDate As Date, _
                    ByVal v_datKeyEndDate As Date, _
                    ByRef r_rsDBPersonalSchedExist As ADODB.Recordset _
                 ) As Boolean
        
On Error GoTo eh
    
    GetExistingPersonalSchedule = True
    Set r_rsDBPersonalSchedExist = GetExistPersonalSched(v_lngEmpNum, v_datKeyStartDate, v_datKeyEndDate)
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "GetExistingPersonalSchedule"
End Function


'for change routine only IsBUSChangePersSched
'allows to see if Personal record exist for the date range requested,
Private Function IsPersonalSchedExist _
                (ByVal v_lngEmpNum As Long _
                 , ByVal v_datStartDate As Date _
                 , ByVal v_datEndDate As Date _
                 ) As Boolean
        
On Error GoTo IsPersonalSchedExist_EH

    Dim l_objDBPersonalSched As ADODB.Recordset
    
    Set l_objDBPersonalSched = GetExistPersonalSched(v_lngEmpNum, v_datStartDate, v_datEndDate)
    If RecordsExist(l_objDBPersonalSched) Then
        IsPersonalSchedExist = True
    Else
    End If
    
    Set l_objDBPersonalSched = Nothing
    
Exit Function
IsPersonalSchedExist_EH:
    Set l_objDBPersonalSched = Nothing
    TrsRaiseError mstrCMODULENAME, "IsPersonalSchedExist"

End Function

'check for duplicate rows
'this reads the database to get an existing Personal schedule that matches
'emp_num, start date and end date
Private Function GetExistPersonalSched _
                (ByVal v_lngEmpNum As Long _
                 , ByVal v_datStartDate As Date _
                 , ByVal v_datEndDate As Date _
                ) As ADODB.Recordset


    On Error GoTo GetExistPersonalSched_EH

    Dim l_dbSchedPersonal As dbSchedPersonal.cdSchedPersonal

    Set l_dbSchedPersonal = CreateObject("dbSchedPersonal.cdSchedPersonal")

    Set GetExistPersonalSched = l_dbSchedPersonal.GetExistPersonalSched(v_lngEmpNum, v_datStartDate, v_datEndDate)

    Set l_dbSchedPersonal = Nothing
    GetObjectContext.SetComplete

Exit Function

GetExistPersonalSched_EH:
    Set l_dbSchedPersonal = Nothing
    GetObjectContext.SetAbort
    TrsRaiseError mstrCMODULENAME, "GetExistPersonalSched"

End Function


Private Sub AppendPersonalError( _
                ByRef r_rsErrors As ADODB.Recordset, _
                ByVal v_intErrorType As TrTypeBRV, _
                ByVal v_lngRecordId As Long, _
                ByVal v_strErrorField As String, _
                ByVal v_strErrorMsg As String, _
                Optional ByVal v_strKeyField1 As Variant, _
                Optional ByVal v_strKeyField2 As Variant, _
                Optional ByVal v_strKeyField3 As Variant, _
                Optional ByVal v_strKeyField4 As Variant, _
                Optional ByVal v_strKeyField5 As Variant, _
                Optional ByVal v_strKeyField6 As Variant, _
                Optional ByVal v_strKeyField7 As Variant, _
                Optional ByVal v_strKeyField8 As Variant, _
                Optional ByVal v_strKeyField9 As Variant, _
                Optional ByVal v_strKeyField10 As Variant)
    
On Error GoTo AppendPersonalError_EH
    If IsNull(v_strKeyField1) Or IsMissing(v_strKeyField1) Then
        v_strKeyField1 = ""
    End If
    If IsNull(v_strKeyField2) Or IsMissing(v_strKeyField2) Then
        v_strKeyField2 = ""
    End If
    If IsNull(v_strKeyField3) Or IsMissing(v_strKeyField3) Then
        v_strKeyField3 = ""
    End If
    If IsNull(v_strKeyField4) Or IsMissing(v_strKeyField4) Then
        v_strKeyField4 = ""
    End If
    If IsNull(v_strKeyField5) Or IsMissing(v_strKeyField5) Then
        v_strKeyField5 = ""
    End If
    If IsNull(v_strKeyField6) Or IsMissing(v_strKeyField6) Then
        v_strKeyField6 = ""
    End If
    If IsNull(v_strKeyField7) Or IsMissing(v_strKeyField7) Then
        v_strKeyField7 = ""
    End If
    If IsNull(v_strKeyField8) Or IsMissing(v_strKeyField8) Then
        v_strKeyField8 = ""
    End If
    If IsNull(v_strKeyField9) Or IsMissing(v_strKeyField9) Then
        v_strKeyField9 = ""
    End If
    If IsNull(v_strKeyField10) Or IsMissing(v_strKeyField10) Then
        v_strKeyField10 = ""
    End If
    
    
    If v_intErrorType = TrType_PER Then
        AppendBusinessError r_rsErrors, LongAppErr(m_lngErrorNo), mstrCBUSINESSERR, v_strErrorMsg, mstrCMODULENAME, v_intErrorType, v_lngRecordId, v_strErrorField, mstrCVARFIELDNAMES, v_strKeyField1, v_strKeyField2, v_strKeyField3, v_strKeyField4, v_strKeyField5, v_strKeyField6, v_strKeyField7, v_strKeyField8, v_strKeyField9, v_strKeyField10
    End If
    
Exit Sub

AppendPersonalError_EH:
    TrsRaiseError mstrCMODULENAME, "AppendPersonalError"

End Sub


Public Function GetSchedChange() As ADODB.Recordset


    On Error GoTo GetSchedChange_EH

    Dim l_dbSchedChange As dbSchedPersonal.cdSchedPersonal

    Set l_dbSchedChange = CreateObject("dbSchedPersonal.cdSchedPersonal")

    Set GetSchedChange = l_dbSchedChange.GetSchedChange

    Set l_dbSchedChange = Nothing
    GetObjectContext.SetComplete

Exit Function

GetSchedChange_EH:
    Set l_dbSchedChange = Nothing
    GetObjectContext.SetAbort
    TrsRaiseError mstrCMODULENAME, "GetSchedChange"

End Function


Private Sub AppendCascadeDates( _
                ByVal v_strAction As String, _
                ByVal v_lngEmpNum As Long, _
                ByVal v_datStartDate As Date, _
                ByVal v_datEndDate As Date)
' Simply add the date records, without attempting to adjust as added.
' Consolidate all the date intervals changed at the end of the process by
' employeeNumber.
'
    Dim l_colRawCascadeIntervals As Collection
    Dim l_objDateInterval As cpuDateInterval

On Error GoTo eh
    Set l_colRawCascadeIntervals = GetCascadeIntervals
    Set l_objDateInterval = New cpuDateInterval
    l_objDateInterval.EmployeeNumber = v_lngEmpNum
    l_objDateInterval.StartDate = v_datStartDate
    l_objDateInterval.EndDate = v_datEndDate
    l_colRawCascadeIntervals.Add l_objDateInterval
    
    Set l_objDateInterval = Nothing
    Set l_colRawCascadeIntervals = Nothing
    Exit Sub
    
eh:
    Set l_objDateInterval = Nothing
    Set l_colRawCascadeIntervals = Nothing
    TrsRaiseError mstrCMODULENAME, "AppendCascadeDates"
End Sub


Private Function ConsolidatedCascadeIntervals() As Collection
' Consolidate/join all the raw date updates by employee number.
' These remaining date intervals will be used for cascade request.
    Dim l_setEmp As cuSet
    Dim l_varEmp As Variant
    Dim l_colRawIntervals As Collection
    Dim l_colConIntervals As Collection
    Dim l_objInterval As cpuDateInterval

On Error GoTo eh
    Set l_colConIntervals = New Collection
' Determine unique employees to be processed.
    Set l_setEmp = New cuSet
    For Each l_objInterval In GetCascadeIntervals
        l_setEmp.Add l_objInterval.EmployeeNumber
    Next

'    For each unique employee,
'    copy the raw interval objects to a separate collection;
'    execute consolidation/join process on filtered collection;
'    copy remaining consolidated interval objects to new collection.
    For Each l_varEmp In l_setEmp
        Set l_colRawIntervals = FilterCascadeIntervals(l_varEmp)
        JoinIntervals l_colRawIntervals
        For Each l_objInterval In l_colRawIntervals
            l_colConIntervals.Add l_objInterval
        Next
    Next
    
    Set ConsolidatedCascadeIntervals = l_colConIntervals
    
    Set l_setEmp = Nothing
    Set l_colRawIntervals = Nothing
    Set l_colConIntervals = Nothing
    Set l_objInterval = Nothing
    Exit Function

eh:
    Set l_setEmp = Nothing
    Set l_colRawIntervals = Nothing
    Set l_colConIntervals = Nothing
    Set l_objInterval = Nothing
    TrsRaiseError mstrCMODULENAME, "ConsolidatedCascadeIntervals"
End Function


Private Function FilterCascadeIntervals( _
                    ByVal v_lngEmpNum As Long) _
                As Collection
' Answer a new collection with references that match employee number.
    Dim l_objInterval As cpuDateInterval
    Dim l_colFilteredIntervals As Collection

On Error GoTo eh
    Set l_colFilteredIntervals = New Collection
    For Each l_objInterval In GetCascadeIntervals
        If l_objInterval.EmployeeNumber = v_lngEmpNum Then
            l_colFilteredIntervals.Add l_objInterval
        End If
    Next
    Set FilterCascadeIntervals = l_colFilteredIntervals
    Set l_colFilteredIntervals = Nothing
    Set l_objInterval = Nothing
    Exit Function

eh:
    Set l_colFilteredIntervals = Nothing
    Set l_objInterval = Nothing
    TrsRaiseError mstrCMODULENAME, "FilterCascadeIntervals"
End Function


Private Function JoinIntervals( _
            ByRef r_colIntervals As Collection)
'    Perform this for each employee.
'    This will discard duplicates.
'    Use an object because I could only get a byval reference to the variant array.
    Dim blnDone As Boolean
    Dim i As Integer
    Dim j As Integer
    Dim intIntersection As Integer
    Dim intCount As Integer
    Dim l_colIntervals As Collection
    Dim l_objOuterInterval As cpuDateInterval
    Dim l_objInnerInterval As cpuDateInterval

On Error GoTo eh
    Set l_colIntervals = r_colIntervals        '
    
    Do Until blnDone
        intCount = intCount + 1
'        Debug.Print "Start iteration: " & intCount
        blnDone = True
        For i = 1 To l_colIntervals.Count
            For j = 1 To l_colIntervals.Count
                If Not (l_colIntervals(i) Is l_colIntervals(j)) Then
                    Set l_objOuterInterval = l_colIntervals(i)
                    Set l_objInnerInterval = l_colIntervals(j)
                    intIntersection = DateIntersection( _
                        l_objOuterInterval.StartDate, l_objOuterInterval.EndDate, _
                        l_objInnerInterval.StartDate, l_objInnerInterval.EndDate)
                    Select Case intIntersection
                        Case 1
'                            Debug.Print "Removing inner1: " & _
'                                l_objInnerInterval.EmployeeNumber & ":" & _
'                                l_objInnerInterval.StartDate & ":" & l_objInnerInterval.EndDate
                            l_colIntervals.Remove j
                            blnDone = False
                            Exit For
                        Case 2
'                            Debug.Print "Removing outer2: " & _
'                                l_objInnerInterval.EmployeeNumber & ":" & _
'                                l_objOuterInterval.StartDate & ":" & l_objOuterInterval.EndDate
                            l_colIntervals.Remove i
                            blnDone = False
                            Exit For
                        Case 3, 4, 5
                            l_objOuterInterval.StartDate = _
                                MinDate(l_objOuterInterval.StartDate, l_objInnerInterval.StartDate)
                            l_objOuterInterval.EndDate = _
                                MaxDate(l_objOuterInterval.EndDate, l_objInnerInterval.EndDate)
'                            Debug.Print "Removing inner3: " & _
'                                l_objInnerInterval.EmployeeNumber & ":" & _
'                                l_objInnerInterval.StartDate & ":" & l_objInnerInterval.EndDate
                            l_colIntervals.Remove j
                            blnDone = False
                            Exit For
                    End Select
                End If
            Next
            If Not blnDone Then
                Exit For
            End If
        Next
    Loop

    Set l_colIntervals = Nothing
    Set l_objOuterInterval = Nothing
    Set l_objInnerInterval = Nothing
    Exit Function

eh:
    Set l_colIntervals = Nothing
    Set l_objOuterInterval = Nothing
    Set l_objInnerInterval = Nothing
    TrsRaiseError mstrCMODULENAME, "JoinIntervals"
End Function


Private Function GetCascadeIntervals() As Collection
' Answer a collection of variant arrays that includes:
' empNum, startDate, endDate

On Error GoTo eh
    If m_colCascadeIntervals Is Nothing Then
        Set m_colCascadeIntervals = New Collection
    End If
    Set GetCascadeIntervals = m_colCascadeIntervals
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "GetCascadeIntervals"
End Function


Private Function GetEmpSchedDates() As ADODB.Recordset
    Dim l_strEmpDate As String
    Dim l_strFilter As String
    Dim l_rsEmpDates As ADODB.Recordset
    Dim l_colDateIntervals As Collection
    Dim l_cpuDateInterval As cpuDateInterval

On Error GoTo eh

'guard for empty recordset
    Set l_colDateIntervals = ConsolidatedCascadeIntervals
    
    If l_colDateIntervals.Count = 0 Then
        Set GetEmpSchedDates = Nothing
        Exit Function
    End If

    Set l_rsEmpDates = New ADODB.Recordset
    With l_rsEmpDates
        .Fields.Append "EMP_NUM", adInteger, , adFldUpdatable
        .Fields.Append "START_DATE", adDate, , adFldUpdatable
        .Fields.Append "END_DATE", adDate, , adFldUpdatable
        .Open
    End With

    For Each l_cpuDateInterval In l_colDateIntervals
        l_rsEmpDates.AddNew
        l_rsEmpDates!EMP_NUM = l_cpuDateInterval.EmployeeNumber
        l_rsEmpDates!START_DATE = l_cpuDateInterval.StartDate
        l_rsEmpDates!END_DATE = l_cpuDateInterval.EndDate
        l_rsEmpDates.Update
    Next
    
    Set GetEmpSchedDates = l_rsEmpDates
    Set l_rsEmpDates = Nothing
    Set l_colDateIntervals = Nothing
    Set l_cpuDateInterval = Nothing
    Exit Function

eh:
    Set l_rsEmpDates = Nothing
    Set l_colDateIntervals = Nothing
    Set l_cpuDateInterval = Nothing
    TrsRaiseError mstrCMODULENAME, "GetEmpSchedDates"
End Function

'navigate thru the database table ignoring candidate records (requests for deletion ) for deletion until
'one can be found to be used as default date (defaultdate - 1 = end_date)
'
Private Function IsNoDelDBRecords _
                (ByVal v_lngEmpNum As Long _
                 , ByVal v_rsPSCloneRequest As ADODB.Recordset _
                 , ByRef r_datStartDate As Date _
                 , ByRef r_rsDBDefaultEndDate As ADODB.Recordset) As Boolean

    On Error GoTo IsNoDelDBRecords_EH
    
    Dim l_strFilter As String
    Dim l_dbDefaultEndDate As dbSchedPersonal.cdSchedPersonal
    Dim l_datLowestEndDate As Date
    
    Set l_dbDefaultEndDate = CreateObject("dbSchedPersonal.cdSchedPersonal")
    
    Set r_rsDBDefaultEndDate = l_dbDefaultEndDate.GetDBDefaultEndDate(v_lngEmpNum, r_datStartDate)
    
    If r_rsDBDefaultEndDate!START_DATE = m_datCOPTIONALSTARTDATE And _
       r_rsDBDefaultEndDate!END_DATE = m_datCOPTIONALENDDATE Then 'are returned 1/1/100 and 12/31/9999?
       IsNoDelDBRecords = True
       Exit Function   'no records were found - eof
    End If
    
    'check for existing key if it exist get next defaultenddate with the start date of existing key
    'db eof check only for request clone: is the db start date and end date in the request for changes, if not present there are no changes "C"
    'search for default date record which is not going to be deleted and if
    'not found default to optional dates
    
    l_strFilter = "KEY_START_DATE = #" & r_rsDBDefaultEndDate!START_DATE & "# " & _
             " AND KEY_END_DATE = #" & r_rsDBDefaultEndDate!END_DATE & "# " & _
             " AND KEY_EMP_NUM = " & v_lngEmpNum & _
             " AND RECORD_STATUS = 'D'"
    If IsRecordsInFilterRecordSet(l_strFilter, v_rsPSCloneRequest) Then 'db matched req keys
        'perform loop to keep searching until no more recs are available
        r_datStartDate = r_rsDBDefaultEndDate!START_DATE
        IsNoDelDBRecords = False
    Else
        IsNoDelDBRecords = True
    End If
    
    Set l_dbDefaultEndDate = Nothing
    
    If IsClearedFilterRecordsetEmpty(v_rsPSCloneRequest) Then
    End If
    
    GetObjectContext.SetComplete

Exit Function

IsNoDelDBRecords_EH:
    Set l_dbDefaultEndDate = Nothing
    GetObjectContext.SetAbort
    TrsRaiseError mstrCMODULENAME, "IsNoDelDBRecords"

End Function

Private Function GetRotationScheduleCode( _
                                        ByVal v_datChargeDate As Date, _
                                        ByVal v_intRotationCode As Integer, _
                                        ByVal v_datEffDate As Date, _
                                        ByVal v_intStartPosition As Integer, _
                                        ByRef r_intRecordPosition As Integer _
                                        ) As Integer
On Error GoTo GetRotationScheduleCode_EH

    Dim rs As ADODB.Recordset

    GetRotationScheduleCode = BusScheduleDataManager.GetRotationScheduleCode( _
                                           v_datChargeDate _
                                         , v_intRotationCode _
                                         , v_datEffDate _
                                         , v_intStartPosition _
                                         , r_intRecordPosition)

    Set rs = Nothing

Exit Function

GetRotationScheduleCode_EH:

    Set rs = Nothing
    TrsRaiseError mstrCMODULENAME, "GetRotationScheduleCode"

End Function


'create reference to busSchedule
Private Function BusScheduleDataManager() As busSchedule.cbsSchedule

On Error GoTo eh

    Set BusScheduleDataManager = _
        CreateObject("busSchedule.cbsSchedule")

Exit Function

eh:
    TrsRaiseError mstrCMODULENAME, "BusScheduleDataManager"
    
End Function




Private Function UpdExistPersfromInputReq( _
    ByVal v_lngEmpNum As Long, _
    ByVal v_intRecordId As Integer, _
    ByVal v_rdatOutdate1 As Date, _
    ByVal v_rdatOutdate2 As Date, _
    ByVal v_rsPSCloneRequest As ADODB.Recordset, _
    ByRef r_datXupdStartdate1 As Date, _
    ByRef r_datXupdEnddate1 As Date, _
    ByRef r_objRSExistingPersonalSched As ADODB.Recordset, _
    ByRef r_rsNetPersonalSchedUpdates As ADODB.Recordset) _
        As Boolean

On Error GoTo UpdExistPersfromInputReq_EH

    Dim l_vdatOldStartdate2 As Date
    Dim l_vdatOldEnddate2 As Date
    Dim l_intNewWeekPosition As Integer
    
    Dim l_strFilter As String
    Dim v_datDupdStartdate1 As Date
    Dim v_datDupdEnddate1 As Date
    Dim v_datRqCloneStartdate1 As Date
    Dim v_datRqCloneEnddate1 As Date
    
    
'delete = overlap generated by previous transactions I or C
'if the key is found in the other modified dates touched by a upd existing start or end date
     l_strFilter = "(START_DATE >= #" & r_datXupdStartdate1 & "#" & _
              " AND START_DATE <= #" & r_datXupdEnddate1 & "#" & _
              " AND EMP_NUM = " & v_lngEmpNum & _
              " AND RECORD_STATUS <> 'D')" & _
              " OR (END_DATE >= #" & r_datXupdStartdate1 & "#" & _
              " AND END_DATE <= #" & r_datXupdEnddate1 & "#" & _
              " AND EMP_NUM = " & v_lngEmpNum & _
              " AND RECORD_STATUS <> 'D')"
    v_datDupdEnddate1 = r_datXupdEnddate1
    v_datDupdStartdate1 = r_datXupdStartdate1
    If IsRecordsInFilterRecordSet(l_strFilter, v_rsPSCloneRequest) Then 'updating key fields if deleting
        Do Until v_rsPSCloneRequest.EOF  'find earlier start date and latest end date
           v_datRqCloneStartdate1 = v_rsPSCloneRequest!START_DATE
           v_datRqCloneEnddate1 = v_rsPSCloneRequest!END_DATE
                If CDate(v_datRqCloneStartdate1) <= CDate(r_datXupdStartdate1) And CDate(v_datRqCloneEnddate1) >= CDate(r_datXupdEnddate1) Then 'A)=/= and G)=/> and B)</=  and H)</> if end date was reached, delete at this point 'ie 7/12/2002 for all
                    ConstructCDIupdaterequests _
                    r_rsNetPersonalSchedUpdates _
                    , r_objRSExistingPersonalSched!EMP_NUM _
                    , DELETENDS_1 _
                    , v_intRecordId _
                    , , , _
                    , v_rdatOutdate1 _
                    , v_rdatOutdate2 _
                    , , , , _
                    r_objRSExistingPersonalSched!LAST_MOD_TIMESTAMP
                    'build DB2Request ("DELETENDS")
                    r_datXupdStartdate1 = v_rdatOutdate1
                    r_datXupdEnddate1 = v_rdatOutdate2
                    r_objRSExistingPersonalSched!RECORD_STATUS = "X"  'flag all deleted records
                    r_objRSExistingPersonalSched.Update 'added to ensure delete
                    Exit Do
                Else
                   If CDate(v_datRqCloneStartdate1) = CDate(r_datXupdStartdate1) And CDate(v_datRqCloneEnddate1) < CDate(r_datXupdEnddate1) Then 'D)=/< inside range of db rec
                   ' 'use the req end date plus 1 for dbstart date , keep moving to the next explicit req record until end date is reached or eof
                       v_datDupdStartdate1 = DateAdd("d", 1, v_datRqCloneEnddate1) 'use the req end date for the new db date range, ie 7/1/2002 - 7/12/2002- keeping track of temp key db record for update (UPD) or delete
                       r_datXupdStartdate1 = v_datDupdStartdate1
                   Else
                      If CDate(v_datRqCloneStartdate1) > CDate(r_datXupdStartdate1) And CDate(v_datRqCloneEnddate1) >= CDate(r_datXupdEnddate1) Then  'C)>/= and I)>/> if end date was reached, tommorrow the start will be answered
                        ' 'use the req start date subtract -1 for dbend date
                          v_datDupdEnddate1 = DateAdd("d", -1, v_datRqCloneStartdate1)  'for update portion end date
                          r_datXupdEnddate1 = v_datDupdEnddate1
                             'insert implicit end date use CDate(v_datRqCloneStartdate1) - 1 day and for start date use CDate(r_datXupdStartdate1)
                             'for r_datXupdStartdate1 with requested use CDate(v_datRqCloneEnddate1) + 1 day
                            ''calculate new week postion for implicit insert
    '                             v_datDupdEnddate1 = DateAdd("d", -1, v_datRqCloneStartdate1)
                            ConstructCDIupdaterequests _
                              r_rsNetPersonalSchedUpdates _
                            , r_objRSExistingPersonalSched!EMP_NUM _
                            , INSERTNDS_3 _
                            , v_intRecordId _
                            , r_objRSExistingPersonalSched!SCHED_CD _
                            , r_objRSExistingPersonalSched!ROT_CD _
                            , r_objRSExistingPersonalSched!ROT_WEEK _
                            , r_datXupdStartdate1 _
                            , v_datDupdEnddate1 _
                            , , , _
                            , r_objRSExistingPersonalSched!LAST_MOD_TIMESTAMP
                            'build DB2Request ("INSERTNDS") 'replace with adj v_rdatOutdate1
                              ConstructCDIupdaterequests _
                              r_rsNetPersonalSchedUpdates _
                              , r_objRSExistingPersonalSched!EMP_NUM _
                              , DELETENDS_1 _
                              , v_intRecordId _
                              , , , _
                              , v_rdatOutdate1 _
                              , v_rdatOutdate2 _
                              , , , , _
                              r_objRSExistingPersonalSched!LAST_MOD_TIMESTAMP
                              'build DB2Request ("DELETENDS")
                              r_datXupdStartdate1 = v_rdatOutdate1
                              r_datXupdEnddate1 = v_rdatOutdate2
                              r_objRSExistingPersonalSched!RECORD_STATUS = "X"  'flag all deleted records
                              r_objRSExistingPersonalSched.Update 'added to ensure delete
                      Else
                          If CDate(v_datRqCloneStartdate1) > CDate(r_datXupdStartdate1) And CDate(v_datRqCloneEnddate1) < CDate(r_datXupdEnddate1) Then 'F)>/< reaching for an end date, db rec start changes as explicit rec causes implicit recs to be generated
                             'insert implicit end date use CDate(v_datRqCloneStartdate1) - 1 day and for start date use CDate(r_datXupdStartdate1)
                             'for r_datXupdStartdate1 with requested use CDate(v_datRqCloneEnddate1) + 1 day
                            ''calculate new week postion for implicit insert
                             v_datDupdEnddate1 = DateAdd("d", -1, v_datRqCloneStartdate1)
                             If Val(r_objRSExistingPersonalSched!ROT_CD) > 0 Then
                                If GetRotationScheduleCode(r_datXupdStartdate1 _
                                                        , r_objRSExistingPersonalSched!ROT_CD _
                                                        , r_objRSExistingPersonalSched!START_DATE _
                                                        , r_objRSExistingPersonalSched!ROT_WEEK _
                                                        , l_intNewWeekPosition) = -1 Then
                                   l_intNewWeekPosition = r_objRSExistingPersonalSched!ROT_WEEK
                                Else   'valid l_intNewWeekPosition was returned
                                End If
                             Else 'if rot_cd does not exist skip
                             End If
                                ConstructCDIupdaterequests _
                                  r_rsNetPersonalSchedUpdates _
                                , r_objRSExistingPersonalSched!EMP_NUM _
                                , INSERTNDS_3 _
                                , v_intRecordId _
                                , r_objRSExistingPersonalSched!SCHED_CD _
                                , r_objRSExistingPersonalSched!ROT_CD _
                                , l_intNewWeekPosition _
                                , r_datXupdStartdate1 _
                                , v_datDupdEnddate1 _
                                , , , _
                                , r_objRSExistingPersonalSched!LAST_MOD_TIMESTAMP
                                'build DB2Request ("INSERTNDS") 'replace with adj v_rdatOutdate1
                                v_datDupdStartdate1 = DateAdd("d", 1, v_datRqCloneEnddate1)
                                r_datXupdStartdate1 = v_datDupdStartdate1
                          Else
                              If CDate(v_datRqCloneStartdate1) < CDate(r_datXupdStartdate1) And CDate(v_datRqCloneEnddate1) < CDate(r_datXupdEnddate1) Then 'E) reaching for an end date
                                 v_datDupdStartdate1 = DateAdd("d", 1, v_datRqCloneEnddate1)
                                 r_datXupdStartdate1 = v_datDupdStartdate1
                              End If
                          End If
                      End If
                   End If
                End If
            v_rsPSCloneRequest.MoveNext
        Loop
    End If
    
                            
Exit Function

UpdExistPersfromInputReq_EH:
    
    TrsRaiseError mstrCMODULENAME, "UpdExistPersfromInputReq"

        
End Function

Private Function IsStartDateBeforeEmployeeEmploymentDate(ByVal v_lngEmpNum As Long, _
                                        ByRef r_datStartDate As Date _
                                        ) As Boolean
    On Error GoTo eh
    
    Dim l_rsEmpInfo As ADODB.Recordset
    Dim l_datEmploymentDate As Date
    Dim l_objDBEmployee As dbEmployee.cdEmployee
    
    Set l_objDBEmployee = CreateObject("dbEmployee.cdEmployee")
    Set l_rsEmpInfo = l_objDBEmployee.GetEmployee(v_lngEmpNum)
    Set l_objDBEmployee = Nothing
    
    If RecordsExist(l_rsEmpInfo) Then
        l_datEmploymentDate = l_rsEmpInfo!EMPLOY_DATE
    Else
        IsStartDateBeforeEmployeeEmploymentDate = False
        Exit Function
    End If
    
    If l_datEmploymentDate > r_datStartDate Then
        r_datStartDate = l_datEmploymentDate
        IsStartDateBeforeEmployeeEmploymentDate = True
    Else
        IsStartDateBeforeEmployeeEmploymentDate = False
    End If
    Exit Function

eh:
    GetObjectContext.SetAbort
    Set l_objDBEmployee = Nothing

    TrsRaiseError mstrCMODULENAME _
        , "IsStartDateBeforeEmployeeEmploymentDate"


 End Function

Private Function BRVErrorsExist(ByVal v_rsBRVs As ADODB.Recordset) As Boolean

    v_rsBRVs.Filter = "ERR_TYPE = 'BE' or ERR_TYPE = 'AE'"
    If v_rsBRVs.EOF And v_rsBRVs.BOF Then
        BRVErrorsExist = False
    Else
       BRVErrorsExist = True
    End If

    v_rsBRVs.Filter = ""

End Function

Private Sub SplitScheduleForHolidayWeeks(ByRef r_rsNetPersonalSchedUpdates As ADODB.Recordset, ByVal v_rsPersonalSchedRequest As ADODB.Recordset)

    Dim l_lngEmpNum As Long
    Dim l_lngHoldEmpNum As Long
    Dim l_strUnion As String
    Dim l_blnFlexPolicyApplies As Boolean
    Dim l_rsWork As ADODB.Recordset
    Dim l_datStart As Date
    
    
    Set l_rsWork = CopyChanges(r_rsNetPersonalSchedUpdates)
    l_rsWork.Sort = "EMP_NUM,START_DATE"
    l_rsWork.Filter = "RECORD_STATUS = 'I' or RECORD_STATUS = 'C'"
    If l_rsWork.BOF And l_rsWork.EOF Then
        Exit Sub
    End If
'    v_rsPersonalSchedRequest.Sort = "START_DATE"
'    v_rsPersonalSchedRequest.Filter = "RECORD_STATUS = 'I' or RECORD_STATUS = 'C'"
'    l_datStart = GetStartDate(v_rsPersonalSchedRequest)
    
    Set m_HolidayCalendar = GetHolidayCalendar()
    If m_HolidayCalendar Is Nothing Then
        Exit Sub
    End If
    
    
    l_lngHoldEmpNum = 0
    l_rsWork.MoveFirst

    While Not l_rsWork.EOF
        l_lngEmpNum = l_rsWork!EMP_NUM
        If l_lngHoldEmpNum <> l_lngEmpNum Then
            l_blnFlexPolicyApplies = FlexPolicyApplies(l_lngEmpNum, l_strUnion)
            l_lngHoldEmpNum = l_lngEmpNum
        End If
        If l_blnFlexPolicyApplies Then
            m_HolidayCalendar.Filter = "UNION_CD = '" & Trim(l_strUnion) & "' "
            If Not m_HolidayCalendar.BOF And Not m_HolidayCalendar.EOF Then
               ApplyHolidaySplits l_rsWork, r_rsNetPersonalSchedUpdates
            End If
        End If
        l_rsWork.MoveNext
    Wend

End Sub

Private Function FlexPolicyApplies(ByVal v_lngEmpNum As Long, ByRef rstrUnion) As Boolean
    On Error GoTo eh
    
    Dim l_rsEmpInfo As ADODB.Recordset
    Dim l_datEmploymentDate As Date
    Dim l_objDBEmployee As dbEmployee.cdEmployee
    Dim l_intLocationCode As Integer
    Dim l_strDepartmentCode As String
    Dim l_strExecContractInd As String
    Dim l_strOTExclusion As String
    
    Set l_objDBEmployee = CreateObject("dbEmployee.cdEmployee")
    Set l_rsEmpInfo = l_objDBEmployee.GetEmployee(v_lngEmpNum)
    Set l_objDBEmployee = Nothing
    
    If l_rsEmpInfo.EOF And l_rsEmpInfo.BOF Then
        FlexPolicyApplies = False
        Exit Function
    End If
    
    rstrUnion = l_rsEmpInfo!UNION_CD
    l_intLocationCode = l_rsEmpInfo!LOC_CD
    l_strDepartmentCode = l_rsEmpInfo!DEPT_NUM
    l_strExecContractInd = l_rsEmpInfo!EXEC_CONTRACT_IND
    l_strOTExclusion = Trim(l_rsEmpInfo!AR_OT_EXCLUSN)
    
    If Trim(l_strExecContractInd) <> "E" Then
        FlexPolicyApplies = False
        Exit Function
    End If
    
    If DoesRuleApply(601, trDepartmentNumber, Trim(l_strDepartmentCode)) And l_strOTExclusion = "0000" Then
        FlexPolicyApplies = False
        Exit Function
    End If
    
    If DoesRuleApply(602, trDepartmentNumber, Trim(l_strDepartmentCode)) And l_strOTExclusion = "0000" Then
        FlexPolicyApplies = False
        Exit Function
    End If
    
    If DoesRuleApply(600, trLocationCode, Trim(str(l_intLocationCode))) Then
        FlexPolicyApplies = False
        Exit Function
    End If

    If DoesRuleApply(600, trDepartmentNumber, Trim(l_strDepartmentCode)) Then
        FlexPolicyApplies = False
        Exit Function
    End If

    FlexPolicyApplies = True
    Exit Function

eh:
    GetObjectContext.SetAbort
    Set l_objDBEmployee = Nothing

    TrsRaiseError mstrCMODULENAME _
        , "FlexPolicyApplies"


 End Function

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
    TrsRaiseError mstrCMODULENAME, "DoesRuleApply"
    
End Function

Private Function GetHolidayCalendar() As ADODB.Recordset

    Dim l_obj As dbCodeTables.cdCodeTables
    Dim l_col As Collection
    Dim l_datStart As Date
    Dim l_datEnd As Date
       
    On Error GoTo eh
    l_datStart = AddDays(-30, Now)
    l_datEnd = CDate("2032-12-31")
    Set l_col = New Collection
    l_col.Add "XA"
    l_col.Add "YA"
    l_col.Add "ZA"
    Set l_obj = CreateObject("dbCodeTables.cdCodeTables")
    Set GetHolidayCalendar = l_obj.GetHolidaysBetween(l_datStart, l_datEnd, l_col)
    Set l_obj = Nothing
    Exit Function
    
eh:
    Set l_obj = Nothing
    TrsRaiseError mstrCMODULENAME, "GetHolidayCalendar"

End Function

Private Sub ApplyHolidaySplits(ByVal v_rsWork As ADODB.Recordset, ByRef r_rsNetPersonalSchedUpdates As ADODB.Recordset)

    Dim l_lngEmpNum As Long
    Dim l_intSchedCode As Integer
    Dim l_lngRotCode As Long
    Dim l_intRotWeek As Integer
    Dim l_datHol As Date
    Dim l_datStartofWeek As Date
    Dim l_datStartofWeekHold As Date
    Dim l_intRecordId As Integer
    Dim l_datStart As Date
    Dim l_datEnd As Date
    Dim l_strTimeStamp As String

    

    On Error GoTo eh
    
    l_lngEmpNum = v_rsWork!EMP_NUM
    l_intSchedCode = v_rsWork!SCHED_CD
    l_lngRotCode = v_rsWork!ROT_CD
    l_intRotWeek = v_rsWork!ROT_WEEK
    l_intRecordId = v_rsWork!RECORD_ID
    l_datStart = v_rsWork!START_DATE
    l_datEnd = v_rsWork!END_DATE
    l_strTimeStamp = v_rsWork!LAST_MOD_TIMESTAMP
    If l_lngRotCode > 0 Then
        l_intSchedCode = GetFirstSchedInRotation(l_lngRotCode)
    End If
    If Not FortyHourWeekAlternativeSchedule(l_intSchedCode) Then
        Exit Sub
    End If
    
    l_datStartofWeekHold = CDate("2000-01-01")
    m_HolidayCalendar.MoveFirst
    l_datHol = m_HolidayCalendar!HOLIDAY_DATE
    l_datStartofWeekHold = DateAdd("d", 1 - Weekday(l_datHol), l_datHol)
    While Not m_HolidayCalendar.EOF
        l_datHol = m_HolidayCalendar!HOLIDAY_DATE
        l_datStartofWeek = DateAdd("d", 1 - Weekday(l_datHol), l_datHol)
        If l_datStartofWeekHold <> l_datStartofWeek Then
            UpdateSchedule r_rsNetPersonalSchedUpdates, l_datStartofWeekHold, l_lngEmpNum, l_intRecordId, l_intSchedCode, l_lngRotCode, l_intRotWeek, l_datStart, l_datEnd, l_strTimeStamp
            l_datStartofWeekHold = l_datStartofWeek
        End If
        m_HolidayCalendar.MoveNext
    Wend
    UpdateSchedule r_rsNetPersonalSchedUpdates, l_datStartofWeek, l_lngEmpNum, l_intRecordId, l_intSchedCode, l_lngRotCode, l_intRotWeek, l_datStart, l_datEnd, l_strTimeStamp
    Exit Sub
eh:
TrsRaiseError mstrCMODULENAME, "ApplyHolidaySplits"


End Sub

Private Sub UpdateSchedule(ByRef r_rsNetPersonalSchedUpdates As ADODB.Recordset, ByVal l_datStartofWeekHold As Date, ByVal l_lngEmpNum As Long, _
    ByVal l_intRecordId As Integer, ByVal l_intSchedCode As Integer, ByVal l_lngRotCode As Long, ByVal l_intRotWeek As Integer, _
    ByRef l_datStart As Date, ByVal l_datEnd As Date, ByVal l_strTimeStamp As String)
    
     Dim l_datEndofWeek As Date
     Dim n_datStart As Date
     Dim n_datEnd As Date
     Dim l_strFilter As String
     Dim l_blnUpdateSched As Boolean
     
     If r_rsNetPersonalSchedUpdates.EOF And r_rsNetPersonalSchedUpdates.BOF Then
        Exit Sub
     Else
        r_rsNetPersonalSchedUpdates.Filter = ""
     End If
     n_datStart = l_datStart
     l_strFilter = "EMP_NUM = " & str(l_lngEmpNum) & " and START_DATE <= " & l_datStartofWeekHold & "   and  END_DATE >= " & l_datStartofWeekHold
     
    r_rsNetPersonalSchedUpdates.Filter = l_strFilter
    If Not r_rsNetPersonalSchedUpdates.BOF And Not r_rsNetPersonalSchedUpdates.EOF Then
        l_blnUpdateSched = False
        If l_intSchedCode = r_rsNetPersonalSchedUpdates!SCHED_CD Then
            l_blnUpdateSched = True
        Else
            If l_lngRotCode > 0 Then
                If l_lngRotCode = r_rsNetPersonalSchedUpdates!ROT_CD Then
                    l_blnUpdateSched = True
                End If
            End If
        End If
        If l_blnUpdateSched Then
        l_datEndofWeek = DateAdd("d", 6, l_datStartofWeekHold)
        r_rsNetPersonalSchedUpdates!START_DATE = l_datStartofWeekHold
        r_rsNetPersonalSchedUpdates!KEY_START_DATE = l_datStartofWeekHold
        r_rsNetPersonalSchedUpdates!END_DATE = l_datEndofWeek
        r_rsNetPersonalSchedUpdates!KEY_END_DATE = l_datEndofWeek
        r_rsNetPersonalSchedUpdates!SCHED_CD = Get8HourSchedule(l_intSchedCode)
        r_rsNetPersonalSchedUpdates!ROT_CD = 0
        r_rsNetPersonalSchedUpdates!ROT_WEEK = 0
        r_rsNetPersonalSchedUpdates.Update
        If l_datStartofWeekHold = l_datStart Then 'split beginning
            n_datStart = DateAdd("d", 7, l_datStartofWeekHold)
            ConstructCDIupdaterequests r_rsNetPersonalSchedUpdates, l_lngEmpNum, INSERTNDS_3, l_intRecordId, l_intSchedCode, l_lngRotCode, l_intRotWeek, n_datStart, l_datEnd, n_datStart, l_datEnd, l_strTimeStamp, l_strTimeStamp
        Else
            If l_datEndofWeek = l_datEnd Then 'split end
                n_datEnd = DateAdd("d", -1, l_datStartofWeekHold)
                ConstructCDIupdaterequests r_rsNetPersonalSchedUpdates, l_lngEmpNum, INSERTNDS_3, l_intRecordId, l_intSchedCode, l_lngRotCode, l_intRotWeek, l_datStart, n_datEnd, l_datStart, n_datEnd, l_strTimeStamp, l_strTimeStamp
            Else  'split middle
                n_datEnd = DateAdd("d", -1, l_datStartofWeekHold)
                ConstructCDIupdaterequests r_rsNetPersonalSchedUpdates, l_lngEmpNum, INSERTNDS_3, l_intRecordId, l_intSchedCode, l_lngRotCode, l_intRotWeek, l_datStart, n_datEnd, l_datStart, n_datEnd, l_strTimeStamp, l_strTimeStamp
                n_datStart = DateAdd("d", 7, l_datStartofWeekHold)
                ConstructCDIupdaterequests r_rsNetPersonalSchedUpdates, l_lngEmpNum, INSERTNDS_3, l_intRecordId, l_intSchedCode, l_lngRotCode, l_intRotWeek, n_datStart, l_datEnd, n_datStart, l_datEnd, l_strTimeStamp, l_strTimeStamp
            End If
        End If
        End If
    End If
    l_datStart = n_datStart
    r_rsNetPersonalSchedUpdates.Filter = ""
    
End Sub

Private Function CopyChanges(ByVal r_rsNetPersonalSchedUpdates As ADODB.Recordset) As ADODB.Recordset

    Dim l_rs As ADODB.Recordset
    Dim l_lngEmpNum As Long

    r_rsNetPersonalSchedUpdates.Filter = ""
    r_rsNetPersonalSchedUpdates.MoveFirst
        
    Set l_rs = CreateNetPersonalSchedUpdates()
    
    While Not r_rsNetPersonalSchedUpdates.EOF
    
        
        If IsEmpty(r_rsNetPersonalSchedUpdates!EMP_NUM) Then
            l_lngEmpNum = 0
        Else
            l_lngEmpNum = r_rsNetPersonalSchedUpdates!EMP_NUM
        End If
        
        If l_lngEmpNum = 0 Then
            If IsEmpty(r_rsNetPersonalSchedUpdates!KEY_EMP_NUM) Then
                l_lngEmpNum = 0
            Else
                l_lngEmpNum = r_rsNetPersonalSchedUpdates!KEY_EMP_NUM
            End If
        End If
        
        If l_lngEmpNum > 0 Then
            l_rs.AddNew

            l_rs!EMP_NUM = l_lngEmpNum
            l_rs!START_DATE = r_rsNetPersonalSchedUpdates!START_DATE
            l_rs!END_DATE = r_rsNetPersonalSchedUpdates!END_DATE
            l_rs!ROT_CD = r_rsNetPersonalSchedUpdates!ROT_CD
            l_rs!ROT_WEEK = r_rsNetPersonalSchedUpdates!ROT_WEEK
            l_rs!SCHED_CD = r_rsNetPersonalSchedUpdates!SCHED_CD
            l_rs!CREATED_EMP_NUM = r_rsNetPersonalSchedUpdates!CREATED_EMP_NUM
            l_rs!LAST_MOD_EMP_NUM = r_rsNetPersonalSchedUpdates!LAST_MOD_EMP_NUM
            l_rs!CREATED_TIMESTAMP = r_rsNetPersonalSchedUpdates!CREATED_TIMESTAMP
            l_rs!LAST_MOD_TIMESTAMP = r_rsNetPersonalSchedUpdates!LAST_MOD_TIMESTAMP
            l_rs!KEY_EMP_NUM = r_rsNetPersonalSchedUpdates!KEY_EMP_NUM
            l_rs!KEY_START_DATE = r_rsNetPersonalSchedUpdates!KEY_START_DATE
            l_rs!KEY_END_DATE = r_rsNetPersonalSchedUpdates!KEY_END_DATE
            l_rs!RECORD_STATUS = r_rsNetPersonalSchedUpdates!RECORD_STATUS
            l_rs!RECORD_ID = r_rsNetPersonalSchedUpdates!RECORD_ID
        End If
    
        r_rsNetPersonalSchedUpdates.MoveNext
    Wend

    l_rs.Update
    
    Set CopyChanges = l_rs

End Function


Private Function GetFirstSchedInRotation(ByVal v_lngRotCode As Long) As Integer

    Dim l_obj As dbSchedule.cdSchedule
    Dim l_rs As ADODB.Recordset
    Dim l_col As Collection

    On Error GoTo eh
    
    Set l_col = New Collection
    l_col.Add v_lngRotCode
    Set l_obj = CreateObject("dbSchedule.cdSchedule")
    Set l_rs = l_obj.GetRotationSchedules(l_col)
    If l_rs.EOF And l_rs.BOF Then
        GetFirstSchedInRotation = 0
    Else
        GetFirstSchedInRotation = l_rs!SCHED_CD
    End If
    Set l_obj = Nothing
    Exit Function
eh:
TrsRaiseError mstrCMODULENAME, "GetFirstSchedInRotation"
End Function




Private Function FortyHourWeekAlternativeSchedule(ByVal v_intSchedCode As Long) As Boolean

    Dim l_obj As dbSchedule.cdSchedule
    Dim l_rs As ADODB.Recordset
    Dim l_col As Collection

    On Error GoTo eh
    
    If v_intSchedCode = 0 Then
        FortyHourWeekAlternativeSchedule = False
        Exit Function
    End If
    
    Set l_col = New Collection
    l_col.Add v_intSchedCode
    Set l_obj = CreateObject("dbSchedule.cdSchedule")
    Set l_rs = l_obj.GetAvailSchedInfo(l_col)
    If l_rs.EOF And l_rs.BOF Then
        FortyHourWeekAlternativeSchedule = False
    Else
        If FourTenHourShifts(l_rs) Or FourNineHourShiftsPlus4HourShift(l_rs) Then
            FortyHourWeekAlternativeSchedule = True
        Else
            FortyHourWeekAlternativeSchedule = False
        End If
    End If
    Set l_obj = Nothing
    Exit Function
eh:
TrsRaiseError mstrCMODULENAME, "FortyHourWeekScheduleSplit"
End Function

Private Function FourTenHourShifts(ByVal l_rs As ADODB.Recordset) As Boolean

    Dim l_int10HourShifts As Integer
    
    l_int10HourShifts = 0
    If l_rs!SUN_STR_HRS = 10 Then
        l_int10HourShifts = l_int10HourShifts + 1
    End If
    If l_rs!MON_STR_HRS = 10 Then
        l_int10HourShifts = l_int10HourShifts + 1
    End If
    If l_rs!TUE_STR_HRS = 10 Then
        l_int10HourShifts = l_int10HourShifts + 1
    End If
    If l_rs!WED_STR_HRS = 10 Then
        l_int10HourShifts = l_int10HourShifts + 1
    End If
    If l_rs!THU_STR_HRS = 10 Then
        l_int10HourShifts = l_int10HourShifts + 1
    End If
    If l_rs!FRI_STR_HRS = 10 Then
        l_int10HourShifts = l_int10HourShifts + 1
    End If
    If l_rs!SAT_STR_HRS = 10 Then
        l_int10HourShifts = l_int10HourShifts + 1
    End If
    
    If l_int10HourShifts = 4 Then
        FourTenHourShifts = True
    Else
        FourTenHourShifts = False
    End If

End Function

Private Function FourNineHourShiftsPlus4HourShift(ByVal l_rs As ADODB.Recordset) As Boolean

    Dim l_int9HourShifts As Integer
    Dim l_int4HourShifts As Integer
    
    l_int9HourShifts = 0
    If l_rs!SUN_STR_HRS = 9 Then
        l_int9HourShifts = l_int9HourShifts + 1
    End If
    If l_rs!MON_STR_HRS = 9 Then
        l_int9HourShifts = l_int9HourShifts + 1
    End If
    If l_rs!TUE_STR_HRS = 9 Then
        l_int9HourShifts = l_int9HourShifts + 1
    End If
    If l_rs!WED_STR_HRS = 9 Then
        l_int9HourShifts = l_int9HourShifts + 1
    End If
    If l_rs!THU_STR_HRS = 9 Then
        l_int9HourShifts = l_int9HourShifts + 1
    End If
    If l_rs!FRI_STR_HRS = 9 Then
        l_int9HourShifts = l_int9HourShifts + 1
    End If
    If l_rs!SAT_STR_HRS = 9 Then
        l_int9HourShifts = l_int9HourShifts + 1
    End If
    
    l_int4HourShifts = 0
    If l_rs!SUN_STR_HRS = 4 Then
        l_int4HourShifts = l_int4HourShifts + 1
    End If
    If l_rs!MON_STR_HRS = 4 Then
        l_int4HourShifts = l_int4HourShifts + 1
    End If
    If l_rs!TUE_STR_HRS = 4 Then
        l_int4HourShifts = l_int4HourShifts + 1
    End If
    If l_rs!WED_STR_HRS = 4 Then
        l_int4HourShifts = l_int4HourShifts + 1
    End If
    If l_rs!THU_STR_HRS = 4 Then
        l_int4HourShifts = l_int4HourShifts + 1
    End If
    If l_rs!FRI_STR_HRS = 4 Then
        l_int4HourShifts = l_int4HourShifts + 1
    End If
    If l_rs!SAT_STR_HRS = 4 Then
        l_int4HourShifts = l_int4HourShifts + 1
    End If
    
    If l_int9HourShifts = 4 And l_int4HourShifts = 1 Then
        FourNineHourShiftsPlus4HourShift = True
    Else
        FourNineHourShiftsPlus4HourShift = False
    End If

End Function

Private Function Get8HourSchedule(ByVal v_intSchedCode As Integer) As Integer

    Dim intNewSched As Integer
    Dim l_obj As dbSchedule.cdSchedule
    Dim l_rs As ADODB.Recordset
    Dim l_rsNewSchedules As ADODB.Recordset
    Dim l_col As Collection
    Dim l_strShiftStart As String
    Dim l_strLunchStart As String
    Dim l_strNewLunchStart As String
    
    On Error GoTo eh
    
    Set l_col = New Collection
    l_col.Add v_intSchedCode
    Set l_obj = CreateObject("dbSchedule.cdSchedule")
    Set l_rs = l_obj.GetAvailSchedInfo(l_col)
    
    If l_rs!SUN_STR_HRS > 0 Then
        l_strShiftStart = l_rs!SUN_START_TIME
        l_strLunchStart = l_rs!SUN_LUNCH_START
    End If
    If l_rs!MON_STR_HRS > 0 Then
        l_strShiftStart = l_rs!MON_START_TIME
        l_strLunchStart = l_rs!MON_LUNCH_START
    End If
    If l_rs!TUE_STR_HRS > 0 Then
        l_strShiftStart = l_rs!TUE_START_TIME
        l_strLunchStart = l_rs!TUE_LUNCH_START
    End If
    If l_rs!WED_STR_HRS > 0 Then
        l_strShiftStart = l_rs!WED_START_TIME
        l_strLunchStart = l_rs!WED_LUNCH_START
    End If
    
    Set l_obj = CreateObject("dbSchedule.cdSchedule")
    Set l_rsNewSchedules = l_obj.GetFlexSchedInfo(l_strShiftStart)
    
    intNewSched = 20021
    If Not l_rsNewSchedules.BOF And Not l_rsNewSchedules.EOF Then
        l_rsNewSchedules.MoveFirst
        intNewSched = l_rsNewSchedules!SCHED_CD
        While Not l_rsNewSchedules.EOF
            l_strNewLunchStart = l_rsNewSchedules!MON_LUNCH_START
            If l_strNewLunchStart = l_strLunchStart Then
                intNewSched = l_rsNewSchedules!SCHED_CD
            End If
            l_rsNewSchedules.MoveNext
        Wend
    End If
        
    Set l_obj = Nothing
    Set l_rs = Nothing
    Set l_rsNewSchedules = Nothing
    Get8HourSchedule = intNewSched
    Exit Function
eh:
TrsRaiseError mstrCMODULENAME, "Get8HourSchedule"
End Function

