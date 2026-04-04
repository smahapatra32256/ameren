VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "cbsSchedTradeTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const m_strCMODULENAME = "cbsSchedTradeTrans"
Private Enum TRSSchedSourceStartLevel
    trAllScheduleLevels = 50
    trDemandLineTrades = 50
    trDemandLineBaseSchedLevel = 40
    trEmployeeTradesLevel = 30
    trNonDemandLevel = 20
    trWorkGroupBaseScheduleLevel = 10
End Enum

Private m_lngUserNum As Long
Private m_strUserId As String
Private m_rsErrors As ADODB.Recordset
Private m_lngTradeNum As Long
Private m_datStartDate As Date
Private m_datEndDate As Date
Private m_rsTradeHeader As ADODB.Recordset
Private m_rsTradeDetails As ADODB.Recordset
Private m_rsEmpTrade As ADODB.Recordset
Private m_rsLineTrade As ADODB.Recordset
Private m_rsRequestorSchedule As ADODB.Recordset
Private m_rsAcceptorSchedule As ADODB.Recordset
Private m_colEmployees As Collection
Private m_aryEmployees As Variant
Private m_rsEmployeeData As ADODB.Recordset
Private m_colChargeDates As Collection 'The Charge Dates in the Request
Private m_aryDemandWorkGroups As Variant
Private m_aryDemandLines As Variant
Private m_rsSaveTradeRequest As ADODB.Recordset
Private m_rsOtherTradeDetails As ADODB.Recordset
Private m_rsNewSchedules As ADODB.Recordset
Private m_rsShiftsModified As ADODB.Recordset

    'Needed when deleting a trade
    'System only allows the trade on the employee's or line's
    '  schedule to be deleted
    'So if there is a subsequent trade after the trade being deleted
    '  a BRV will be displayed
    'Also this recordset will be used to repapply a previous trade when
    ' the last trade is deleted
    

Public Function DeleteTrade( _
                    ByVal v_aryTradeNums As Variant, _
                    ByVal v_datStartDate As Date, _
                    ByVal v_datEndDate As Date, _
                    ByRef r_rsErrors As ADODB.Recordset, _
                    ByRef r_rsTradeHeader As ADODB.Recordset, _
                    ByRef r_rsTradeDetails As ADODB.Recordset, _
                    ByRef r_rsEmpTrade As ADODB.Recordset, _
                    ByRef r_rsLineTrade As ADODB.Recordset, _
                    ByRef r_rsCascadeConflicts As ADODB.Recordset, _
                    ByRef r_rsExceptionSPA As ADODB.Recordset, _
                    ByRef r_rsAvailScheduleSPA As ADODB.Recordset, _
                    ByRef r_rsEmpSchedDatesSPA As ADODB.Recordset, _
                    ByVal v_blnConfirmConflicts As Boolean, _
                    ByVal v_UserId As String _
                ) As Boolean
    
    Dim l_objCtx As ObjectContext
    Dim l_blnError As Boolean
    Dim l_lngStart As Long
    Dim l_lngEnd As Long
    Dim l_lngI As Long
    Dim l_aryTradeNums As Variant
    Dim l_blnConfirmPending As Boolean
    Dim l_varCascadeUpdates As Variant
    
    m_strUserId = v_UserId
    m_lngUserNum = ParseUserID(v_UserId)
    
    Set l_objCtx = GetObjectContext
    
'Getting Existing State of trade
'    m_lngTradeNum = v_lngTradeNum
    Set m_rsTradeHeader = GetTradeHeader(v_aryTradeNums)
    Set m_rsTradeDetails = GetTradeDetails(v_aryTradeNums)

    If Not RecordsExist(m_rsTradeDetails) Then
        DeleteTrade = True
        l_objCtx.SetComplete
        GoTo xit
    End If
    m_datStartDate = v_datStartDate
    m_datEndDate = v_datEndDate
    Set m_rsErrors = CreateNewBusinessErrorSet
    
    Set m_colEmployees = GetEmployeesInRequest(m_rsTradeDetails)
    GetDemandLineAssignments m_rsTradeDetails, Nothing
'    GetStartAndEndDateOfTrade
    m_aryEmployees = ArrayFromCollection(m_colEmployees)
    Set m_rsEmpTrade = GetEmployeeTrades(m_aryEmployees)
    Set m_rsLineTrade = GetDemandLineTrades(m_aryDemandWorkGroups, m_aryDemandLines)
    Set m_rsOtherTradeDetails = GetAllTrades()
    
    If IsArray(v_aryTradeNums) Then
        l_lngStart = LBound(v_aryTradeNums)
        l_lngEnd = UBound(v_aryTradeNums)
        l_aryTradeNums = v_aryTradeNums
    Else
        l_lngStart = 0
        l_lngEnd = 0
        ReDim l_aryTradeNums(0) As Long
        l_aryTradeNums(0) = CLng(v_aryTradeNums)
    End If
        
    For l_lngI = l_lngStart To l_lngEnd
        m_lngTradeNum = l_aryTradeNums(l_lngI)
        If Not ProcessDeleteTrade() Then
            DeleteTrade = False
            l_objCtx.SetComplete
            GoTo xit
        End If
    Next l_lngI
    
    If Not PostValidateDeleteTrade() Then
        DeleteTrade = False
        l_objCtx.SetComplete
        GoTo xit
    End If
    
    If ProcessCascade( _
                v_blnConfirmConflicts, l_blnConfirmPending, l_varCascadeUpdates, _
                r_rsTradeHeader, r_rsTradeDetails, r_rsEmpTrade, r_rsLineTrade, _
                r_rsCascadeConflicts, r_rsExceptionSPA, r_rsAvailScheduleSPA, _
                r_rsEmpSchedDatesSPA) Then
        If l_blnConfirmPending Then
            DeleteTrade = True
            l_objCtx.SetComplete
            GoTo xit
        End If
    Else
        DeleteTrade = False
        l_objCtx.SetComplete
        GoTo xit
    End If
    
    If Not ProcessScheduleChange(r_rsEmpSchedDatesSPA, Nothing, m_rsErrors) Then
        l_objCtx.SetAbort
        Set r_rsEmpSchedDatesSPA = Nothing
        DeleteTrade = False
        GoTo xit
    Else
        Set r_rsEmpSchedDatesSPA = Nothing
    End If
    
    If Not SaveTradeDB(l_varCascadeUpdates) Then
        DeleteTrade = False
        l_objCtx.SetAbort
        GoTo xit
    End If
    
    DeleteTrade = True
    l_objCtx.SetComplete
xit:
    If Not m_rsErrors Is Nothing Then
        If m_rsErrors.RecordCount = 0 Then
            Set r_rsErrors = Nothing
        Else
            Set r_rsErrors = m_rsErrors
        End If
    Else
        Set r_rsErrors = Nothing
    End If
    SetGlobalObjectsToNothing
    Exit Function
    
DeleteTrade_EH:
    Set r_rsErrors = Nothing
    l_objCtx.SetAbort
    SetGlobalObjectsToNothing
'    Clear the cascade error arguments since the client does not format them.
    If Err.Number = trErrAppCascadeCommitConflict Then
        Err.Source = Join(Array(CStr(TrType_TR), "", "", "", ""), ";")
    End If
    TrsRaiseError m_strCMODULENAME, "DeleteTrade"
    
End Function


Public Function SaveTrade( _
                    ByRef r_lngTradeNum As Long, _
                    ByVal v_datStartDate As Date, _
                    ByVal v_datEndDate As Date, _
                    ByVal v_strTradeDesc As String, _
                    ByVal v_lngApprovedBy As Long, _
                    ByVal v_rsTradeDetails As ADODB.Recordset, _
                    ByVal v_rsEmp1Sched As ADODB.Recordset, _
                    ByVal v_rsEmp2Sched As ADODB.Recordset, _
                    ByRef r_rsErrors As ADODB.Recordset, _
                    ByRef r_rsTradeHeader As ADODB.Recordset, _
                    ByRef r_rsTradeDetails As ADODB.Recordset, _
                    ByRef r_rsEmpTrade As ADODB.Recordset, _
                    ByRef r_rsLineTrade As ADODB.Recordset, _
                    ByRef r_rsCascadeConflicts As ADODB.Recordset, _
                    ByRef r_rsExceptionSPA As ADODB.Recordset, _
                    ByRef r_rsAvailScheduleSPA As ADODB.Recordset, _
                    ByRef r_rsEmpSchedDatesSPA As ADODB.Recordset, _
                    ByVal v_blnConfirmConflicts As Boolean, _
                    ByVal v_UserId As String, _
                    Optional ByVal v_rsEmpSchedModTS As ADODB.Recordset) _
                As Boolean
    
    Dim l_objCtx As ObjectContext
    Dim l_blnError As Boolean

    Dim l_blnConfirmPending As Boolean
    Dim l_varCascadeUpdates As Variant

On Error GoTo SaveTrade_EH
    
    m_strUserId = v_UserId
    m_lngUserNum = ParseUserID(v_UserId)
    
    Set l_objCtx = GetObjectContext
    If Not RecordsExist(v_rsTradeDetails) Then
        SaveTrade = True
        l_objCtx.SetComplete
        GoTo xit
    End If
    Set m_rsErrors = CreateNewBusinessErrorSet
    
    m_lngTradeNum = GetNextTradeNumber()
    m_datStartDate = v_datStartDate
    m_datEndDate = v_datEndDate
    Set m_rsRequestorSchedule = v_rsEmp1Sched
    Set m_rsAcceptorSchedule = v_rsEmp2Sched
    Set m_rsSaveTradeRequest = v_rsTradeDetails
    
    If Not PreValidateSaveTrade(v_strTradeDesc, v_lngApprovedBy, v_rsTradeDetails) Then
        SaveTrade = False
        l_objCtx.SetAbort
        GoTo xit
    End If
    
    GetDemandLineAssignments v_rsEmp1Sched, v_rsEmp2Sched
    If Not GetStateOfTrade(v_rsTradeDetails) Then
        SaveTrade = False
        l_objCtx.SetAbort
        GoTo xit
    End If
    
    If Not ProcessSaveTrade(v_strTradeDesc, v_lngApprovedBy, v_rsTradeDetails) Then
        SaveTrade = False
        l_objCtx.SetAbort
        GoTo xit
    End If
    
    If Not PostValidateSaveTrade() Then
        SaveTrade = False
        l_objCtx.SetAbort
        GoTo xit
    End If

'    If cascade confirmation is not requested, there must be records in v_rsEmpSchedModTS
    If Not RecordsExist(v_rsEmpSchedModTS) And Not v_blnConfirmConflicts Then
        AppendBusinessError m_rsErrors, _
            40000, _
            mstrCAPPLICATIONERR, _
            "Employee schedule modified information not provided.", _
            m_strCMODULENAME
        SaveTrade = False
        l_objCtx.SetAbort
        GoTo xit
    End If
                
    If ProcessCascade( _
                v_blnConfirmConflicts, l_blnConfirmPending, l_varCascadeUpdates, _
                r_rsTradeHeader, r_rsTradeDetails, r_rsEmpTrade, r_rsLineTrade, _
                r_rsCascadeConflicts, r_rsExceptionSPA, r_rsAvailScheduleSPA, _
                r_rsEmpSchedDatesSPA) Then
        If l_blnConfirmPending Then
            r_lngTradeNum = m_lngTradeNum
            SaveTrade = True
            l_objCtx.SetComplete
            GoTo xit
        End If
    Else
        SaveTrade = False
        l_objCtx.SetComplete
        GoTo xit
    End If

    If Not ProcessScheduleChange(r_rsEmpSchedDatesSPA, v_rsEmpSchedModTS, r_rsErrors) Then
        l_objCtx.SetAbort
        Set r_rsEmpSchedDatesSPA = Nothing
        SaveTrade = False
        GoTo xit
    Else
        Set r_rsEmpSchedDatesSPA = Nothing
    End If

    If Not SaveTradeDB(l_varCascadeUpdates) Then
        SaveTrade = False
        l_objCtx.SetAbort
        GoTo xit
    End If

    WriteTradeNumberBRVMessage
    r_lngTradeNum = m_lngTradeNum
    SaveTrade = True
    l_objCtx.SetComplete
    
xit:
    If Not m_rsErrors Is Nothing Then
        If m_rsErrors.RecordCount = 0 Then
            Set r_rsErrors = Nothing
        Else
            Set r_rsErrors = m_rsErrors
        End If
    Else
        Set r_rsErrors = Nothing
    End If
    SetGlobalObjectsToNothing
    Exit Function
    
SaveTrade_EH:
    Set r_rsErrors = Nothing
    Set r_rsTradeHeader = Nothing
    Set r_rsTradeDetails = Nothing
    Set r_rsEmpTrade = Nothing
    Set r_rsLineTrade = Nothing
    Set r_rsCascadeConflicts = Nothing
    Set r_rsExceptionSPA = Nothing
    Set r_rsAvailScheduleSPA = Nothing
    l_objCtx.SetAbort
    SetGlobalObjectsToNothing
    
'    Clear the cascade error arguments since the client does not format them.
    If Err.Number = trErrAppCascadeCommitConflict Then
        Err.Source = Join(Array(CStr(TrType_TR), "", "", "", ""), ";")
    End If
    TrsRaiseError m_strCMODULENAME, "SaveTrade"
End Function


Public Function ResumeSaveTrade( _
        ByVal v_lngTradeNum As Long, _
        ByVal v_rsEmpSchedModTS As ADODB.Recordset, _
        ByVal v_rsTradeHeader As ADODB.Recordset, _
        ByVal v_rsTradeDetails As ADODB.Recordset, _
        ByVal v_rsEmpTrade As ADODB.Recordset, _
        ByVal v_rsLineTrade As ADODB.Recordset, _
        ByVal v_rsConflictList As ADODB.Recordset, _
        ByVal v_rsExceptionSPA As ADODB.Recordset, _
        ByVal v_rsAvailScheduleSPA As ADODB.Recordset, _
        ByVal v_rsEmpSchedDatesSPA As ADODB.Recordset, _
        ByRef r_rsErrors As ADODB.Recordset, _
        ByVal v_UserId As String) _
    As Boolean
'   v_rsEmpSchedModTS will be nothing when ResumeSaveTrade is called
'   after DeleteTrade.  It will process ok if passed, just incur an unnecessary
'   database retrieval for employee schedule modified timestamp.
    Dim l_objCtx As ObjectContext
    Dim l_blnError As Boolean
    Dim l_cbsCascade As busCascade.cbsCascade
    Dim l_varCascadeUpdates As Variant
    Dim l_rsCascadeErrors As ADODB.Recordset

    On Error GoTo eh
    
    m_lngTradeNum = v_lngTradeNum
    m_strUserId = v_UserId
    m_lngUserNum = ParseUserID(v_UserId)
    
    Set l_objCtx = GetObjectContext
    If r_rsErrors Is Nothing Then
        Set m_rsErrors = CreateNewBusinessErrorSet
    Else
        Set m_rsErrors = r_rsErrors
    End If
    
    Set m_rsTradeHeader = v_rsTradeHeader
    Set m_rsTradeDetails = v_rsTradeDetails
    Set m_rsEmpTrade = v_rsEmpTrade
    Set m_rsLineTrade = v_rsLineTrade
    
    If Not ProcessScheduleChange(v_rsEmpSchedDatesSPA, v_rsEmpSchedModTS, r_rsErrors) Then
        l_objCtx.SetAbort
        ResumeSaveTrade = False
        GoTo xit
    End If
    
    Set l_cbsCascade = CreateObject("busCascade.cbsCascade")
    l_cbsCascade.UpdateExceptions _
        v_rsConflictList, v_rsExceptionSPA, v_rsAvailScheduleSPA, _
        False, l_varCascadeUpdates, l_rsCascadeErrors, m_lngUserNum
    Set l_cbsCascade = Nothing
    
    AppendCascadeErrors l_rsCascadeErrors, m_rsErrors
    
    If Not SaveTradeDB(l_varCascadeUpdates) Then
        l_objCtx.SetAbort
        ResumeSaveTrade = False
        GoTo xit
    End If

    WriteTradeNumberBRVMessage
    ResumeSaveTrade = True
    l_objCtx.SetComplete
    
xit:
    Set l_rsCascadeErrors = Nothing
    If Not m_rsErrors Is Nothing Then
        If m_rsErrors.RecordCount = 0 Then
            Set r_rsErrors = Nothing
        Else
            Set r_rsErrors = m_rsErrors
        End If
    Else
        Set r_rsErrors = Nothing
    End If
    SetGlobalObjectsToNothing
    Set l_objCtx = Nothing
    Exit Function
    
eh:
    Set l_rsCascadeErrors = Nothing
    Set l_cbsCascade = Nothing
    Set r_rsErrors = Nothing
    l_objCtx.SetAbort
    SetGlobalObjectsToNothing
    Set l_objCtx = Nothing
'    Clear the cascade error arguments since the client does not format them.
    If Err.Number = trErrAppCascadeCommitConflict Then
        Err.Source = Join(Array(CStr(TrType_TR), "", "", "", ""), ";")
    End If
    TrsRaiseError m_strCMODULENAME, "ResumeSaveTrade"
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
            TrType_TR, , , , _
            r_rsCascadeErrors!KEY_FIELD1, _
            r_rsCascadeErrors!KEY_FIELD2, _
            r_rsCascadeErrors!KEY_FIELD3
        r_rsCascadeErrors.MoveNext
    Loop
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "AppendCascadeErrors"
End Function


Private Function ProcessScheduleChange( _
                    ByRef r_rsEmpSchedDates As ADODB.Recordset, _
                    ByRef r_rsEmpSchedModTS As ADODB.Recordset, _
                    ByRef r_rsErrors As ADODB.Recordset) _
                As Boolean
    Dim l_cdCascade As dbCascade.cdCascade
    Dim l_varEmps As Variant
    Dim l_rsEmpSchedModTS As ADODB.Recordset

On Error GoTo eh
    
    If r_rsEmpSchedDates Is Nothing Then
        ProcessScheduleChange = True
        Exit Function
    End If
    
    l_varEmps = ExtractEmployeeNumbers(r_rsEmpSchedDates)
    
    Set l_cdCascade = CreateObject("dbCascade.cdCascade")
    
    If Not r_rsEmpSchedModTS Is Nothing Then
        Set l_rsEmpSchedModTS = l_cdCascade.GetEmployeeScheduleModified(l_varEmps)
        If EmployeeScheduleChanged(r_rsEmpSchedModTS, l_rsEmpSchedModTS) Then
            ProcessScheduleChange = False
            GoTo xit
        End If
    End If
    
'    Update the emp_mod timestamp
'    This will create a lock that time collection must wait on before saving actual time.
    l_cdCascade.SetEmployeeScheduleModified l_varEmps
    Set l_cdCascade = Nothing
    
    If ActualTimeExists(r_rsEmpSchedDates) Then
        ProcessScheduleChange = False
        GoTo xit
    End If

    ProcessScheduleChange = True
    
xit:
    Set l_cdCascade = Nothing
    Set l_rsEmpSchedModTS = Nothing
    Exit Function
    
eh:
    Set l_cdCascade = Nothing
    Set r_rsErrors = Nothing
    Set l_rsEmpSchedModTS = Nothing
    TrsRaiseError m_strCMODULENAME, "ProcessScheduleChange"
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
            If r_rsEmpSchedDates.Fields("EMP_NUM").Value > 0 Then
                l_cuSet.Add r_rsEmpSchedDates.Fields("EMP_NUM").Value
            End If
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
    TrsRaiseError m_strCMODULENAME, "ExtractEmployeeNumbers"
End Function


Private Function ProcessCascade( _
                    ByVal v_blnConfirmConflicts As Boolean, _
                    ByRef r_blnConfirmPending As Boolean, _
                    ByRef r_varCascadeUpdates As Variant, _
                    ByRef r_rsTradeHeader As ADODB.Recordset, _
                    ByRef r_rsTradeDetails As ADODB.Recordset, _
                    ByRef r_rsEmpTrade As ADODB.Recordset, _
                    ByRef r_rsLineTrade As ADODB.Recordset, _
                    ByRef r_rsCascadeConflicts As ADODB.Recordset, _
                    ByRef r_rsExceptionSPA As ADODB.Recordset, _
                    ByRef r_rsAvailScheduleSPA As ADODB.Recordset, _
                    ByRef r_rsEmpTradeDates As ADODB.Recordset _
                 ) As Boolean

    Dim l_rsCascadeConflicts As ADODB.Recordset
    Dim l_rsExceptionSPA As ADODB.Recordset
    Dim l_rsAvailScheduleSPA As ADODB.Recordset
    Dim l_cbsCascade As busCascade.cbsCascade
    Dim l_varCascadeUpdates As Variant
    Dim l_rsCascadeErrors As ADODB.Recordset

On Error GoTo eh

    r_varCascadeUpdates = Empty
    r_blnConfirmPending = False
    Set r_rsEmpTradeDates = GetEmpDateTrades()

    If r_rsEmpTradeDates Is Nothing Then
        ProcessCascade = True
        Exit Function
    End If

    Set l_rsCascadeConflicts = _
        BuildCascadeConflictList(r_rsEmpTradeDates, l_rsExceptionSPA, l_rsAvailScheduleSPA)

    If RecordsExist(l_rsCascadeConflicts) Then
        If v_blnConfirmConflicts Then
            Set r_rsTradeHeader = m_rsTradeHeader
            Set r_rsTradeDetails = m_rsTradeDetails
            Set r_rsEmpTrade = m_rsEmpTrade
            Set r_rsLineTrade = m_rsLineTrade
            Set r_rsCascadeConflicts = l_rsCascadeConflicts
            Set r_rsExceptionSPA = l_rsExceptionSPA
            Set r_rsAvailScheduleSPA = l_rsAvailScheduleSPA
'            In case inprocess with client
            ClearAllFilters r_rsTradeHeader, r_rsTradeDetails, _
                    r_rsEmpTrade, r_rsLineTrade, r_rsCascadeConflicts, _
                    r_rsExceptionSPA, r_rsAvailScheduleSPA, r_rsEmpTradeDates
            r_blnConfirmPending = True
        Else
            Set l_cbsCascade = CreateObject("busCascade.cbsCascade")
            l_cbsCascade.UpdateExceptions _
                l_rsCascadeConflicts, l_rsExceptionSPA, l_rsAvailScheduleSPA, _
                False, l_varCascadeUpdates, l_rsCascadeErrors, m_lngUserNum
            AppendCascadeErrors l_rsCascadeErrors, m_rsErrors
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
    Set r_rsEmpTradeDates = Nothing
    Set l_rsCascadeConflicts = Nothing
    Set l_rsExceptionSPA = Nothing
    Set l_rsAvailScheduleSPA = Nothing
    Set l_cbsCascade = Nothing
    Set l_rsCascadeErrors = Nothing
    TrsRaiseError m_strCMODULENAME, "ProcessCascade"
End Function


Private Function GetEmpDateTrades() As ADODB.Recordset
    
    Dim l_strEmpDate As String
    Dim l_strFilter As String
    Dim l_rsEmpDates As ADODB.Recordset

On Error GoTo eh

'guard for empty recordset
    If Not RecordsExist(GetShiftsModified) Then
        Set GetEmpDateTrades = Nothing
        Exit Function
    End If

    Set l_rsEmpDates = New ADODB.Recordset
    With l_rsEmpDates
        .Fields.Append "EMP_NUM", adInteger, , adFldUpdatable
        .Fields.Append "CHARGE_DATE", adDate, , adFldUpdatable
        .Open
    End With

    Do Until m_rsShiftsModified.EOF
        l_strFilter = _
            "EMP_NUM = " & m_rsShiftsModified!EMP_NUM & _
            " AND CHARGE_DATE = #" & Format(m_rsShiftsModified!CHARGE_DATE, "mm/dd/yyyy") & "#"
        l_strFilter = MDAC25FilterFix(l_strFilter)
        l_rsEmpDates.Filter = l_strFilter
        If Not RecordsExist(l_rsEmpDates) Then
            l_rsEmpDates.AddNew
            l_rsEmpDates!EMP_NUM = m_rsShiftsModified!EMP_NUM
            l_rsEmpDates!CHARGE_DATE = CDate(m_rsShiftsModified!CHARGE_DATE)
            l_rsEmpDates.Update
        End If
        m_rsShiftsModified.MoveNext
    Loop
    
    l_rsEmpDates.Filter = 0
    Set GetEmpDateTrades = l_rsEmpDates
    Set l_rsEmpDates = Nothing
    Exit Function

eh:
    Set l_rsEmpDates = Nothing
    TrsRaiseError m_strCMODULENAME, "GetEmpDateTrades"
End Function


Private Function AppendShiftModified( _
            ByVal v_lngEmpNum As Long, _
            ByVal v_datChargeDate As Date, _
            ByVal v_intShiftNum As Integer)

    Dim l_strFilter As String
    Dim l_rsShiftsModified As ADODB.Recordset

On Error GoTo eh

    Set l_rsShiftsModified = GetShiftsModified

    l_strFilter = _
        "EMP_NUM = " & v_lngEmpNum & _
        " AND CHARGE_DATE = #" & Format(v_datChargeDate, "mm/dd/yyyy") & "#" & _
        " AND SHIFT_NUM = " & v_intShiftNum
    l_strFilter = MDAC25FilterFix(l_strFilter)
    l_rsShiftsModified.Filter = l_strFilter
    If Not RecordsExist(l_rsShiftsModified) Then
        l_rsShiftsModified.AddNew
        l_rsShiftsModified!EMP_NUM = v_lngEmpNum
        l_rsShiftsModified!CHARGE_DATE = v_datChargeDate
        l_rsShiftsModified!SHIFT_NUM = v_intShiftNum
        l_rsShiftsModified.Update
    End If

    l_rsShiftsModified.Filter = 0

    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "AppendShiftModified"
End Function

Private Function GetShiftsModified() As ADODB.Recordset

On Error GoTo eh

    If m_rsShiftsModified Is Nothing Then
        CreateShiftsModified
    End If
    Set GetShiftsModified = m_rsShiftsModified

    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "GetShiftsModified"
End Function


Private Function CreateShiftsModified()

On Error GoTo eh

    Set m_rsShiftsModified = New ADODB.Recordset
    With m_rsShiftsModified
        .Fields.Append "EMP_NUM", adInteger, , adFldUpdatable
        .Fields.Append "CHARGE_DATE", adDate, , adFldUpdatable
        .Fields.Append "SHIFT_NUM", adInteger, , adFldUpdatable
        .Open
    End With

    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "GetShiftsModified"
End Function


Private Function EmployeeScheduleChanged( _
                    ByRef r_rsEmpSchedModTS1 As ADODB.Recordset, _
                    ByRef r_rsEmpSchedModTS2 As ADODB.Recordset _
                ) As Boolean
'    Compare timestamps on each recordset and append an error when different.
'    Ignore differences when timestamp on r_rsEmpSchedModTS1 is an empty string
'    or record does not exist on r_rsEmpSchedModTS2 and empty string on r_rsEmpSchedModTS1.
    Dim l_blnChanged As Boolean
    Dim l_strFilter As String

On Error GoTo eh

    r_rsEmpSchedModTS1.MoveFirst
    Do Until r_rsEmpSchedModTS1.EOF
        If Not Trim(r_rsEmpSchedModTS1!last_mod_sched_tm) = "" Then
            l_strFilter = "EMP_NUM = " & CStr(r_rsEmpSchedModTS1!EMP_NUM)
            l_strFilter = MDAC25FilterFix(l_strFilter)
            r_rsEmpSchedModTS2.Filter = l_strFilter
            If RecordsExist(r_rsEmpSchedModTS2) Then
                If Not Trim(r_rsEmpSchedModTS1!last_mod_sched_tm) = Trim(r_rsEmpSchedModTS2!last_mod_sched_tm) Then
                    AppendTradeBRV trErrAppCommitConflict, _
                        mstrCBUSINESSERR, _
                        "Employee (" & CStr(r_rsEmpSchedModTS1!EMP_NUM) & _
                        ") schedule has been changed by another user while processing this trade. " & _
                        "Discard your changes and try again.", _
                        r_rsEmpSchedModTS1!EMP_NUM
                    l_blnChanged = True
                End If
            Else
                AppendTradeBRV trErrAppCommitConflict, _
                    mstrCBUSINESSERR, _
                    "Employee (" & CStr(r_rsEmpSchedModTS1!EMP_NUM) & _
                    ") schedule has been changed by another user while processing this trade. " & _
                    "Discard your changes and try again.", _
                    r_rsEmpSchedModTS1!EMP_NUM
                l_blnChanged = True
            End If
        End If
        r_rsEmpSchedModTS1.MoveNext
    Loop

    EmployeeScheduleChanged = l_blnChanged
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "EmployeeScheduleChanged"
End Function


Private Function ActualTimeExists( _
                    ByRef r_rsEmpTradeDates As ADODB.Recordset) As Boolean
    Dim l_blnExists As Boolean
    Dim l_cdTimecard As dbTimeCard.cdTimeCard
    Dim l_rsActualTime As ADODB.Recordset
    Dim l_varEmployees As Variant
    Dim l_datStart As Date
    Dim l_datEnd As Date
    Dim l_strFilter As String

On Error GoTo eh

    l_blnExists = False
    DetermineGetActualTimeScope r_rsEmpTradeDates, l_varEmployees, l_datStart, l_datEnd
    If IsEmpty(l_varEmployees) Then
        ActualTimeExists = False
        Exit Function
    End If

    Set l_cdTimecard = CreateObject("dbTimeCard.cdTimeCard")
    Set l_rsActualTime = _
            l_cdTimecard.GetActualTimeCascadeExists(l_varEmployees, l_datStart, l_datEnd)
    Set l_cdTimecard = Nothing

    Do Until r_rsEmpTradeDates.EOF
        l_strFilter = _
            "EMP_NUM = " & r_rsEmpTradeDates!EMP_NUM & _
            " AND CHARGE_DATE = #" & Format(r_rsEmpTradeDates!CHARGE_DATE, "mm/dd/yyyy") & "#"
        l_strFilter = MDAC25FilterFix(l_strFilter)
        l_rsActualTime.Filter = l_strFilter
        If RecordsExist(l_rsActualTime) Then
            AppendTradeBRV trErrbusActualTimeExists, _
                mstrCBUSINESSERR, _
                "Employee " & CStr(l_rsActualTime!EMP_NUM) & _
                "[" & Trim(l_rsActualTime!EMP_NAME) & "] has existing actual time on " & _
                Format(l_rsActualTime!CHARGE_DATE, "mm/dd/yyyy") & ".", _
                r_rsEmpTradeDates!EMP_NUM
            l_blnExists = True
        End If
        r_rsEmpTradeDates.MoveNext
    Loop

    ActualTimeExists = l_blnExists
    Set l_rsActualTime = Nothing

    Exit Function

eh:
    Set l_cdTimecard = Nothing
    Set l_rsActualTime = Nothing
    TrsRaiseError m_strCMODULENAME, "ActualTimeExists"
End Function


Private Sub DetermineGetActualTimeScope( _
                    ByRef r_rsEmpTradeDates As ADODB.Recordset, _
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
    If Not RecordsExist(r_rsEmpTradeDates) Then
        r_varEmployees = Empty
        r_datStartDate = 0
        r_datEndDate = 0
        Exit Sub
    Else
        r_rsEmpTradeDates.MoveFirst
    End If

' seed the dates with high and low dates
' there should be at least 1 valid record
    l_datStartDate = l_datHighDate
    l_datEndDate = l_datLowDate
    Set l_colDistinctEmpNums = New cuSet
    Do Until r_rsEmpTradeDates.EOF
        l_colDistinctEmpNums.Add CLng(r_rsEmpTradeDates!EMP_NUM.Value)
        l_datStartDate = MinDate(r_rsEmpTradeDates!CHARGE_DATE.Value, l_datStartDate)
        l_datEndDate = MaxDate(r_rsEmpTradeDates!CHARGE_DATE.Value, l_datEndDate)
        r_rsEmpTradeDates.MoveNext
    Loop

    r_datStartDate = l_datStartDate
    r_datEndDate = l_datEndDate
    r_varEmployees = ArrayFromCollection(l_colDistinctEmpNums.Contents)
    r_rsEmpTradeDates.MoveFirst
    Set l_colDistinctEmpNums = Nothing
    Exit Sub

eh:
    Set l_colDistinctEmpNums = Nothing
    TrsRaiseError m_strCMODULENAME, "DetermineGetActualTimeScope"
End Sub


Private Function BuildCascadeConflictList( _
            ByRef r_rsEmpTradeDates As ADODB.Recordset, _
            ByRef r_rsExceptionSPA As ADODB.Recordset, _
            ByRef r_rsAvailScheduleSPA As ADODB.Recordset) As ADODB.Recordset
' Answer Nothing when there are no exceptions to cascade.
' r_rsExceptionSPA and r_rsAvailScheduleSPA are output recordsets.
'
    Dim l_rsEmpInterval As ADODB.Recordset
    Dim l_cbsCascade As busCascade.cbsCascade
    Dim l_rsExceptionShifts As ADODB.Recordset
    
On Error GoTo eh

'    Guard against no trades completed.
    If Not RecordsExist(m_rsTradeDetails) Then
        GoTo xit
    End If
    
    Set l_rsEmpInterval = BuildCascadePrepRequest(r_rsEmpTradeDates)
    Set l_cbsCascade = CreateObject("busCascade.cbsCascade")
    Set l_rsExceptionShifts = _
            l_cbsCascade.GetScheduledExceptionShifts(l_rsEmpInterval, trDemandLineTrades)

'    Guard against no exceptions scheduled on traded shifts.
    If Not RecordsExist(l_rsExceptionShifts) Then
        GoTo xit
    End If

    UpdateNewScheduledShift l_rsExceptionShifts
    
    Set BuildCascadeConflictList = _
        l_cbsCascade.GetExceptionConflictList( _
            l_rsExceptionShifts, r_rsExceptionSPA, r_rsAvailScheduleSPA)

xit:
    Set l_cbsCascade = Nothing
    Set l_rsEmpInterval = Nothing
    Set l_rsExceptionShifts = Nothing
    Exit Function

eh:
    Set l_cbsCascade = Nothing
    Set l_rsEmpInterval = Nothing
    Set l_rsExceptionShifts = Nothing
    Set r_rsExceptionSPA = Nothing
    Set r_rsAvailScheduleSPA = Nothing
    TrsRaiseError m_strCMODULENAME, "BuildCascadeConflictList"
End Function


Private Function BuildCascadePrepRequest( _
            ByRef r_rsEmpTradeDates As ADODB.Recordset) As ADODB.Recordset
    Dim l_rsEmpInterval As ADODB.Recordset
    Dim i As Integer
    
On Error GoTo eh
    Set l_rsEmpInterval = New Recordset
    
    If Not RecordsExist(r_rsEmpTradeDates) Then
        Set BuildCascadePrepRequest = l_rsEmpInterval
        Exit Function
    End If
    
    l_rsEmpInterval.Fields.Append "EMP_NUM", adInteger, , adFldUpdatable
    l_rsEmpInterval.Fields.Append "START_DATE", adDate, , adFldUpdatable
    l_rsEmpInterval.Fields.Append "END_DATE", adDate, , adFldUpdatable
    l_rsEmpInterval.Open

    r_rsEmpTradeDates.MoveFirst
    Do Until r_rsEmpTradeDates.EOF
        If r_rsEmpTradeDates!EMP_NUM > 0 Then
            l_rsEmpInterval.AddNew
            l_rsEmpInterval!EMP_NUM = r_rsEmpTradeDates!EMP_NUM
            l_rsEmpInterval!START_DATE = r_rsEmpTradeDates!CHARGE_DATE
            l_rsEmpInterval!END_DATE = r_rsEmpTradeDates!CHARGE_DATE
            l_rsEmpInterval.Update
        End If
        r_rsEmpTradeDates.MoveNext
    Loop
    
    Set BuildCascadePrepRequest = l_rsEmpInterval
    Set l_rsEmpInterval = Nothing
    Exit Function
    
eh:
    Set l_rsEmpInterval = Nothing
    TrsRaiseError m_strCMODULENAME, "BuildCascadePrepRequest"
End Function


Private Function CreateEmpSchedModTS() As ADODB.Recordset
    Dim l_rsEmpSchedModTS As ADODB.Recordset
    Dim i As Integer
    
On Error GoTo eh
    Set l_rsEmpSchedModTS = New Recordset
    
    l_rsEmpSchedModTS.Fields.Append "EMP_NUM", adInteger, , adFldUpdatable
    l_rsEmpSchedModTS.Fields.Append "LAST_MOD_SCHED_TM", adChar, 26, adFldUpdatable
    l_rsEmpSchedModTS.Open

    Set CreateEmpSchedModTS = l_rsEmpSchedModTS
    Set l_rsEmpSchedModTS = Nothing
    Exit Function
    
eh:
    Set l_rsEmpSchedModTS = Nothing
    TrsRaiseError m_strCMODULENAME, "CreateEmpSchedModTS"
End Function


Private Function UpdateNewScheduledShift( _
            ByRef r_rsExceptionShifts As ADODB.Recordset)
' Lookup the schedule for each scheduled shift modified.
    Dim l_strFilter As String
    Dim l_rsShiftsModifed As ADODB.Recordset

On Error GoTo eh
    Set l_rsShiftsModifed = GetShiftsModified

    l_rsShiftsModifed.Filter = 0
    Do Until l_rsShiftsModifed.EOF
        l_strFilter = "EMP_NUM  = " & CStr(l_rsShiftsModifed!EMP_NUM) & _
            " AND CHARGE_DATE = " & CStr(CDate(l_rsShiftsModifed!CHARGE_DATE)) & _
            " AND SHIFT_NUM = " & CStr(l_rsShiftsModifed!SHIFT_NUM)
        l_strFilter = MDAC25FilterFix(l_strFilter)
        m_rsNewSchedules.Filter = l_strFilter
        If RecordsExist(m_rsNewSchedules) Then
            r_rsExceptionShifts.Filter = l_strFilter
            If RecordsExist(r_rsExceptionShifts) Then
                r_rsExceptionShifts!NEW_SCHED_CD = m_rsNewSchedules!SCHED_CD
                r_rsExceptionShifts!NEW_SCHED_SRC_CD = m_rsNewSchedules!SCHED_SRC_CD
                r_rsExceptionShifts.Update
            End If
        End If
        l_rsShiftsModifed.MoveNext
    Loop
    
    m_rsNewSchedules.Filter = 0
    r_rsExceptionShifts.Filter = 0
    l_rsShiftsModifed.MoveFirst
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "UpdateNewScheduledShift"
End Function


Private Sub ClearAllFilters( _
            ByRef r_rsData1 As ADODB.Recordset, _
            Optional ByRef r_rsData2 As ADODB.Recordset, _
            Optional ByRef r_rsData3 As ADODB.Recordset, _
            Optional ByRef r_rsData4 As ADODB.Recordset, _
            Optional ByRef r_rsData5 As ADODB.Recordset, _
            Optional ByRef r_rsData6 As ADODB.Recordset, _
            Optional ByRef r_rsData7 As ADODB.Recordset, _
            Optional ByRef r_rsData8 As ADODB.Recordset)

On Error GoTo eh
'    Clear all filters in case of debug and inprocess.
    If Not r_rsData1 Is Nothing Then
        r_rsData1.Filter = 0
    End If
    If Not r_rsData2 Is Nothing Then
        r_rsData2.Filter = 0
    End If
    If Not r_rsData3 Is Nothing Then
        r_rsData3.Filter = 0
    End If
    If Not r_rsData4 Is Nothing Then
        r_rsData4.Filter = 0
    End If
    If Not r_rsData5 Is Nothing Then
        r_rsData5.Filter = 0
    End If
    If Not r_rsData6 Is Nothing Then
        r_rsData6.Filter = 0
    End If
    If Not r_rsData7 Is Nothing Then
        r_rsData7.Filter = 0
    End If
    If Not r_rsData8 Is Nothing Then
        r_rsData8.Filter = 0
    End If

    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "ClearAllFilters"
End Sub


Private Function GetNextTradeNumber( _
            ) As Long
            
            
    Dim l_strSequenceType As String
    Dim l_obj As utilDBDataServices.cuDBDataServices
    
    On Error GoTo eh
    
    Set l_obj = CreateObject("utilDBDataServices.cuDBDataServices")
    GetNextTradeNumber = l_obj.GetNextSequenceNumber("TR")
    Set l_obj = Nothing

        
    Exit Function
    
eh:
    
  TrsRaiseError m_strCMODULENAME, "GetNextTradeNumber"
           
    
            
End Function


Private Sub SetGlobalObjectsToNothing()

    On Error GoTo eh

    Set m_rsErrors = Nothing
    Set m_colEmployees = Nothing
    Set m_rsTradeHeader = Nothing
    Set m_rsTradeDetails = Nothing
    Set m_rsEmpTrade = Nothing
    Set m_rsLineTrade = Nothing
    Set m_rsRequestorSchedule = Nothing
    Set m_rsAcceptorSchedule = Nothing
    
' RSG Added 4/21/02
    Set m_rsSaveTradeRequest = Nothing
    Set m_rsOtherTradeDetails = Nothing
    Set m_rsEmployeeData = Nothing
    Set m_colChargeDates = Nothing
    Set m_rsNewSchedules = Nothing
    Set m_rsShiftsModified = Nothing
    
' Initialize all module variables
    m_lngUserNum = 0
    m_strUserId = ""
    m_lngTradeNum = 0
    m_datStartDate = 0
    m_datEndDate = 0
    m_aryEmployees = Empty
    m_aryDemandWorkGroups = Empty
    m_aryDemandLines = Empty

    Exit Sub
eh:
    TrsRaiseError m_strCMODULENAME, "SetGlobalObjectsToNothing"

End Sub

Private Function GetStateOfTrade( _
    ByVal v_rsTradeDetails As ADODB.Recordset _
    ) As Boolean
    
    Dim l_aryEmps As Variant
    
    On Error GoTo eh

    l_aryEmps = ArrayFromCollection(m_colEmployees)
    Set m_rsEmpTrade = GetEmployeeTrades(l_aryEmps)
    Set m_rsLineTrade = GetDemandLineTrades(m_aryDemandWorkGroups, m_aryDemandLines)
    Set m_rsTradeHeader = GetTradeHeader(m_lngTradeNum)
    Set m_rsTradeDetails = GetTradeDetails(m_lngTradeNum)
    
    GetStateOfTrade = True
    
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "GetStateOfTrade"
    
End Function

Private Function GetEmployeesInRequest( _
            ByVal v_objRS As ADODB.Recordset _
            ) As Collection

    Dim l_colEmployees As Collection
    Dim l_strEmployee As String
    Dim l_aryEmployees As Variant

    On Error GoTo eh

    Set l_colEmployees = New Collection
    
    If Not RecordsExist(v_objRS) Then
        Set GetEmployeesInRequest = l_colEmployees
        Exit Function
    End If
    
    v_objRS.Sort = "ROLE_TYPE_CD Desc"
    v_objRS.MoveFirst
    While Not v_objRS.EOF
        l_strEmployee = Trim(str(v_objRS!EMP_NUM))
        If Not IncludesString(l_strEmployee, l_colEmployees) Then
            l_colEmployees.Add l_strEmployee
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


Private Function AppendTradeDetailBRV( _
            ByVal v_rsTradeDetail As ADODB.Recordset, _
            ByVal v_lngErrorNumber As Long, _
            ByVal v_strErrorType As String, _
            ByVal v_strErrorDesc As String _
            ) As Boolean
            
    Dim l_bolError As Boolean
    Dim lngKeyEmpNum As Long
    Dim lngKeyShiftNum As Long
    Dim l_datKeyChargeDate As Date
    Dim l_strKeyCreatedTime As String
    Dim lngKeyWGNum As Long
    Dim lngKeyLineNum As Long
    
    On Error GoTo AppendTradeDetailBRV_EH
    
    lngKeyEmpNum = v_rsTradeDetail!EMP_NUM
    lngKeyShiftNum = v_rsTradeDetail!SHIFT_NUM
    l_datKeyChargeDate = v_rsTradeDetail!CHARGE_DATE
    
    l_bolError = AppendBusinessError(m_rsErrors, _
            v_lngErrorNumber, _
            v_strErrorType, _
            v_strErrorDesc, _
            m_strCMODULENAME, _
            TrType_TR_DETL, , , "EMP_NUM;CHARGE_DATE,SHIFT_NUM", _
            str(lngKeyEmpNum), _
            str(l_datKeyChargeDate), _
            str(lngKeyShiftNum), , , str(m_lngTradeNum))

    AppendTradeDetailBRV = l_bolError
    Exit Function
    
AppendTradeDetailBRV_EH:
    TrsRaiseError m_strCMODULENAME, "AppendTradeDetailBRV"

End Function

Private Function AppendTradeBRV( _
            ByVal v_lngErrorNumber As Long, _
            ByVal v_strErrorType As String, _
            ByVal v_strErrorDesc As String, _
            Optional v_lngEmpNum As Long = 0 _
            ) As Boolean
            
    Dim l_bolError As Boolean
    
    On Error GoTo AppendTradeBRV_EH
    
   
    l_bolError = AppendBusinessError(m_rsErrors, _
            v_lngErrorNumber, _
            v_strErrorType, _
            v_strErrorDesc, _
            m_strCMODULENAME, _
            TrType_TR, , , "EMP_NUM", _
            str(v_lngEmpNum), , , , , str(m_lngTradeNum))

    AppendTradeBRV = l_bolError
    Exit Function
    
AppendTradeBRV_EH:
    TrsRaiseError m_strCMODULENAME, "AppendTradeBRV"

End Function



Private Function GetDemandLineTrades( _
            ByVal v_aryWorkGroups As Variant, _
            ByVal v_aryLines As Variant _
            ) As ADODB.Recordset

    On Error GoTo GetDemandLineTrades_EH

    Dim l_obj As dbSchedTrade.cdSchedTrade

    Set l_obj = CreateObject("dbSchedTrade.cdSchedTrade")
    
    Set GetDemandLineTrades = l_obj.GetDemandLineTrades(v_aryWorkGroups, v_aryLines, m_datStartDate, m_datEndDate)

    Set l_obj = Nothing
    GetObjectContext.SetComplete

    Exit Function

GetDemandLineTrades_EH:
    Set l_obj = Nothing
    GetObjectContext.SetAbort
    TrsRaiseError m_strCMODULENAME, "GetDemandLineTrades"

End Function

Private Function GetEmployeeTrades( _
            ByVal v_aryEmployees As Variant _
            ) As ADODB.Recordset

    On Error GoTo GetEmployeeTrades_EH


    Dim l_obj As dbSchedTrade.cdSchedTrade
         
    
    Set l_obj = CreateObject("dbSchedTrade.cdSchedTrade")
    
    Set GetEmployeeTrades = l_obj.GetEmployeeTrades(v_aryEmployees, m_datStartDate, m_datEndDate)
         
    Set l_obj = Nothing
    
    Exit Function
    
GetEmployeeTrades_EH:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "GetEmployeeTrades"


End Function


Private Function PreValidateSaveTrade( _
    ByVal v_strTradeDesc As String, _
    ByVal v_lngApprovedBy As Long, _
    ByVal v_rsTradeDetails As ADODB.Recordset _
    ) As Boolean
    
    Dim l_blnReqValid As Boolean
    
    On Error GoTo eh
    
    l_blnReqValid = True
    If Not ValidRoleTypes(v_rsTradeDetails) Then
        l_blnReqValid = False
    End If
    If Not ValidEmployees(v_rsTradeDetails) Then
        l_blnReqValid = False
    End If
    If Not ValidChargeDates(v_rsTradeDetails) Then
        l_blnReqValid = False
    End If
    If Not ValidShiftNumbers(v_rsTradeDetails) Then
        l_blnReqValid = False
    End If
    If Not ValidTradeType(v_rsTradeDetails) Then
        l_blnReqValid = False
    End If
    If Not ValidScheduleCodes(v_rsTradeDetails) Then
        l_blnReqValid = False
    End If
    Set m_colChargeDates = GetChargeDatesInRequest(v_rsTradeDetails)
    If Not ValidGABSchedules() Then
        l_blnReqValid = False
    End If
    
    'The Following edits just display warnings
    DoEmployeesHaveSaveJob v_rsTradeDetails, m_colChargeDates
    DoEmployeesHaveSaveUnion v_rsTradeDetails, m_colChargeDates
    DoEmployeesHaveSaveWorkGroup v_rsTradeDetails, m_colChargeDates
    
    
    PreValidateSaveTrade = l_blnReqValid
    
    
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "PreValidateSaveTrade"
    
    
End Function

Private Function ValidEmployees( _
    ByVal v_rsTradeDetails As ADODB.Recordset _
    ) As Boolean
    
    On Error GoTo eh
    
    Set m_colEmployees = GetEmployeesInRequest(v_rsTradeDetails)
    If Not IsValidNumberOfEmployees(v_rsTradeDetails) Then
        ValidEmployees = False
        Exit Function
    End If
    ValidEmployees = EmployeesExist()
    
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "ValidEmployees"

End Function

Private Function IsValidNumberOfEmployees( _
    ByVal v_rsTradeDetails As ADODB.Recordset) As Boolean

    Dim l_intCount As Integer
    Dim l_lngReqEmp As Long
    Dim l_lngAccEmp As Long
    Dim l_blnSelfTrade As Boolean
    Dim l_blnReqRoleOK As Boolean
    Dim l_blnAccRoleOK As Boolean
    Dim l_strErrorDesc As String
    Dim l_bln As Boolean
    

    On Error GoTo eh

    IsValidNumberOfEmployees = True
    l_intCount = m_colEmployees.Count
    Select Case l_intCount
        Case 0
            IsValidNumberOfEmployees = False
            AppendTradeBRV trErrbusValidationError, _
            mstrCAPPLICATIONERR, _
            "System Error. No employees in trade."
            Exit Function
        Case 1
            l_blnSelfTrade = True
            v_rsTradeDetails.MoveFirst
            While Not v_rsTradeDetails.EOF And l_blnSelfTrade
                If v_rsTradeDetails!ROLE_TYPE_CD <> "S" Then
                    l_blnSelfTrade = False
                End If
                v_rsTradeDetails.MoveNext
            Wend
            If Not l_blnSelfTrade Then
                IsValidNumberOfEmployees = False
                AppendTradeBRV trErrbusValidationError, _
                mstrCAPPLICATIONERR, _
                "System Error. Must be a self trade when only 1 employee.", _
                CLng(m_colEmployees.Item(1))
                Exit Function
            End If
        
        Case 2
            l_lngReqEmp = m_colEmployees(1)
            l_lngAccEmp = m_colEmployees(2)
            l_blnReqRoleOK = True
            l_blnAccRoleOK = True
            v_rsTradeDetails.Filter = "EMP_NUM = " & str(l_lngReqEmp)
            While Not v_rsTradeDetails.EOF And l_blnReqRoleOK
                If v_rsTradeDetails!ROLE_TYPE_CD <> "R" Then
                    l_blnReqRoleOK = False
                    l_strErrorDesc = _
                        "The requestor (" & str(l_lngReqEmp) & _
                        "must be the requestor for all " & _
                        "shifts in this trade."
                    l_bln = AppendTradeDetailBRV( _
                        v_rsTradeDetails, _
                        trErrbusValidationError, _
                        mstrCAPPLICATIONERR, _
                        l_strErrorDesc)
                End If
                v_rsTradeDetails.MoveNext
            Wend
            v_rsTradeDetails.Filter = "EMP_NUM = " & str(l_lngAccEmp)
            While Not v_rsTradeDetails.EOF And l_blnAccRoleOK
                If v_rsTradeDetails!ROLE_TYPE_CD <> "A" Then
                    l_blnAccRoleOK = False
                    l_strErrorDesc = _
                        "The acceptor (" & str(l_lngReqEmp) & _
                        "must be the acceptor for all " & _
                        "shifts in this trade."
                    l_bln = AppendTradeDetailBRV( _
                        v_rsTradeDetails, _
                        trErrbusValidationError, _
                        mstrCAPPLICATIONERR, _
                        l_strErrorDesc)
                End If
                v_rsTradeDetails.MoveNext
            Wend
            If l_blnReqRoleOK And l_blnAccRoleOK Then
                IsValidNumberOfEmployees = True
            Else
                IsValidNumberOfEmployees = False
            End If
        
        Case Else
            IsValidNumberOfEmployees = False
            AppendTradeBRV trErrbusValidationError, _
            mstrCAPPLICATIONERR, _
            "System Error. Up to 2 employees can be involved in trade."
            Exit Function
    End Select
    
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "IsValidNumberOfEmployees"



End Function

Private Function GetTradeHeader( _
    ByVal v_aryTradeNums As Variant _
    ) As ADODB.Recordset
    
    On Error GoTo GetTradeHeader_EH


    Dim l_obj As dbSchedTrade.cdSchedTrade
    
    Set l_obj = CreateObject("dbSchedTrade.cdSchedTrade")
    
    Set GetTradeHeader = l_obj.GetTradeHeader(v_aryTradeNums)
         
    Set l_obj = Nothing
    
    Exit Function
    
GetTradeHeader_EH:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "GetTradeHeader"



End Function

Private Function GetTradeDetails( _
    ByVal v_aryTradeNums As Variant _
    ) As ADODB.Recordset

    On Error GoTo GetTradeDetails_EH


    Dim l_obj As dbSchedTrade.cdSchedTrade

    Set l_obj = CreateObject("dbSchedTrade.cdSchedTrade")
    
    Set GetTradeDetails = l_obj.GetTradeDetails(v_aryTradeNums)
         
    Set l_obj = Nothing
    
    Exit Function
    
GetTradeDetails_EH:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "GetTradeDetails"


End Function

Private Function ProcessSaveTrade( _
    ByVal v_strTradeDesc As String, _
    ByVal v_lngApprovedBy As Long, _
    ByVal v_rsTradeDetails As ADODB.Recordset _
    ) As Boolean
        
    Dim l_lngWGNum As Long
    Dim l_lngLineNum As Long
'''''    Dim l_blnSaveOK as Boolean
    
        
    On Error GoTo ProcessSaveTrade_EH
    
    AddTradeHeader v_strTradeDesc, v_lngApprovedBy

    v_rsTradeDetails.Sort = "EMP_NUM, CHARGE_DATE"
    v_rsTradeDetails.MoveFirst
    
    While Not v_rsTradeDetails.EOF
        If IsEmployeeAssignedToALine(v_rsTradeDetails, l_lngWGNum, l_lngLineNum) Then
            ProcessSaveLineTrade v_rsTradeDetails, l_lngWGNum, l_lngLineNum
        Else
            l_lngWGNum = 0
            l_lngLineNum = 0
            ProcessSaveEmpTrade v_rsTradeDetails
        End If
        AddTradeDetail v_rsTradeDetails, l_lngWGNum, l_lngLineNum
        v_rsTradeDetails.MoveNext
    Wend
    
    ProcessSaveTrade = True
    

'Private m_rsLineTrade As ADODB.Recordset
    
    
    
    Exit Function
    
ProcessSaveTrade_EH:
    TrsRaiseError m_strCMODULENAME, "ProcessSaveTrade"
    
End Function

Private Function IsEmployeeAssignedToALine( _
    ByVal v_rsTradeDetails As ADODB.Recordset, _
    ByRef r_lngWgNum As Long, _
    ByRef r_lngLineNum As Long _
    ) As Boolean
    
    On Error GoTo eh
    
    Dim l_datChargeDate As Date
    Dim l_lngEmpNum As Long
    Dim l_strRole As String
    Dim l_rsSchedule As ADODB.Recordset
       
    l_datChargeDate = v_rsTradeDetails!CHARGE_DATE
    l_strRole = v_rsTradeDetails!ROLE_TYPE_CD
    If l_strRole = "A" Then
        Set l_rsSchedule = m_rsAcceptorSchedule
    Else
        Set l_rsSchedule = m_rsRequestorSchedule
    End If
    l_rsSchedule.Filter = "CHARGE_DATE = " & l_datChargeDate
    l_rsSchedule.MoveFirst
    r_lngWgNum = l_rsSchedule!WG_NUM
    r_lngLineNum = l_rsSchedule!LINE_NUM
    
    If r_lngWgNum > 0 And r_lngLineNum > 0 Then
        IsEmployeeAssignedToALine = True
    Else
        IsEmployeeAssignedToALine = False
    End If
    
    l_rsSchedule.Filter = ""
    Set l_rsSchedule = Nothing
    Exit Function
    
eh:
    Set l_rsSchedule = Nothing
    TrsRaiseError m_strCMODULENAME, "IsEmployeeAssignedToALine"
End Function

Private Sub AddTradeHeader( _
    ByVal v_strTradeDesc As String, _
    ByVal v_lngApprovedBy As Long)
    
    
    On Error GoTo eh
    
    
        'Creating Trade Header Info
    m_rsTradeHeader.AddNew
    m_rsTradeHeader!TRADE_NUM = m_lngTradeNum
    m_rsTradeHeader!TRADE_DESC = v_strTradeDesc
    m_rsTradeHeader!APPROVED_EMP_NUM = v_lngApprovedBy
    m_rsTradeHeader!CREATED_EMP_NUM = CLng(m_strUserId)
    m_rsTradeHeader.Update
    
    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "AddTradeHeader"
    
End Sub

Private Sub AddTradeDetail( _
    ByVal v_rsTradeDetails As ADODB.Recordset, _
    ByVal v_lngWGNum As Long, _
    ByVal v_lngLineNum As Long)
    
    On Error GoTo eh
    
    m_rsTradeDetails.AddNew
    m_rsTradeDetails!TRADE_NUM = m_lngTradeNum
    m_rsTradeDetails!WG_NUM = v_lngWGNum
    m_rsTradeDetails!LINE_NUM = v_lngLineNum
    m_rsTradeDetails!CHARGE_DATE = v_rsTradeDetails!CHARGE_DATE
    m_rsTradeDetails!EMP_NUM = v_rsTradeDetails!EMP_NUM
    m_rsTradeDetails!SHIFT_NUM = v_rsTradeDetails!SHIFT_NUM
    m_rsTradeDetails!SCHED_CD = v_rsTradeDetails!SCHED_CD
    m_rsTradeDetails!ROLE_TYPE_CD = v_rsTradeDetails!ROLE_TYPE_CD
    m_rsTradeDetails!WEEK_TRADE_IND = v_rsTradeDetails!WEEK_TRADE_IND
    m_rsTradeDetails.Update
    
    AppendShiftModified v_rsTradeDetails!EMP_NUM, v_rsTradeDetails!CHARGE_DATE, v_rsTradeDetails!SHIFT_NUM
    
    Exit Sub
   
eh:
    TrsRaiseError m_strCMODULENAME, "AddTradeDetail"
    
    
End Sub

Private Sub ProcessSaveEmpTrade( _
    ByVal v_rsTradeDetails As ADODB.Recordset)
    
    
    Dim l_datChargeDate As Date
    Dim l_lngEmpNum As Long
    Dim l_lngShiftNum As Long
    Dim l_strFilter As String
       
    On Error GoTo eh
       
    l_lngEmpNum = v_rsTradeDetails!EMP_NUM
    l_datChargeDate = v_rsTradeDetails!CHARGE_DATE
    l_lngShiftNum = v_rsTradeDetails!SHIFT_NUM
    l_strFilter = "EMP_NUM = " & str(l_lngEmpNum) & _
            " and CHARGE_DATE = " & l_datChargeDate & _
            " and SHIFT_NUM = " & str(l_lngShiftNum)
    l_strFilter = MDAC25FilterFix(l_strFilter)
    m_rsEmpTrade.Filter = l_strFilter
    If RecordsExist(m_rsEmpTrade) Then
        m_rsEmpTrade.Delete
    End If
    m_rsEmpTrade.AddNew
    m_rsEmpTrade!EMP_NUM = l_lngEmpNum
    m_rsEmpTrade!CHARGE_DATE = l_datChargeDate
    m_rsEmpTrade!SHIFT_NUM = l_lngShiftNum
    m_rsEmpTrade!SCHED_CD = v_rsTradeDetails!SCHED_CD
    m_rsEmpTrade!TRADE_NUM = m_lngTradeNum
    m_rsEmpTrade!CREATED_EMP_NUM = CLng(m_strUserId)
    m_rsEmpTrade.Update
    
    Exit Sub
   
eh:
    TrsRaiseError m_strCMODULENAME, "ProcessSaveEmpTrade"
    
End Sub


Private Sub ProcessSaveLineTrade( _
    ByVal v_rsTradeDetails As ADODB.Recordset, _
    ByVal v_lngWGNum As Long, _
    ByVal v_lngLineNum As Long)
    
    
    Dim l_datChargeDate As Date
    Dim l_lngShiftNum As Long
    Dim l_strFilter As String
       
    On Error GoTo eh
       
    l_datChargeDate = v_rsTradeDetails!CHARGE_DATE
    l_lngShiftNum = v_rsTradeDetails!SHIFT_NUM
    l_strFilter = "WG_NUM = " & str(v_lngWGNum) & _
            " and LINE_NUM = " & str(v_lngLineNum) & _
            " and CHARGE_DATE = " & l_datChargeDate & _
            " and SHIFT_NUM = " & str(l_lngShiftNum)
    l_strFilter = MDAC25FilterFix(l_strFilter)
    m_rsLineTrade.Filter = l_strFilter
    If RecordsExist(m_rsLineTrade) Then
        m_rsLineTrade.Delete
    End If
    m_rsLineTrade.AddNew
    m_rsLineTrade!WG_NUM = v_lngWGNum
    m_rsLineTrade!LINE_NUM = v_lngLineNum
    m_rsLineTrade!CHARGE_DATE = l_datChargeDate
    m_rsLineTrade!SHIFT_NUM = l_lngShiftNum
    m_rsLineTrade!SCHED_CD = v_rsTradeDetails!SCHED_CD
    m_rsLineTrade!TRADE_NUM = m_lngTradeNum
    m_rsLineTrade!CREATED_EMP_NUM = CLng(m_strUserId)
    m_rsLineTrade.Update
    
    Exit Sub
   
eh:
    TrsRaiseError m_strCMODULENAME, "ProcessSaveLineTrade"
    
End Sub

Private Function IsInList( _
        ByVal v_lngCode As Long, _
        ByVal v_colCodes As Collection _
        ) As Boolean

    Dim l_intIndex As Integer
    Dim l_intCount As Integer
    Dim l_bolInList As Boolean
    
    On Error GoTo eh
    
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
    
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "IsInList"
    

End Function

Private Sub GetDemandLineAssignments( _
    ByVal v_rs1 As ADODB.Recordset, _
    ByVal v_rs2 As ADODB.Recordset)
    
    
    Dim l_colWorkGroups As Collection
    Dim l_colLines As Collection
    Dim l_lngWorkGroup As Long
    Dim l_lngLine As Long

    On Error GoTo GetDemandLineAssignments_EH

    Set l_colWorkGroups = New Collection
    Set l_colLines = New Collection
    If RecordsExist(v_rs1) Then
        v_rs1.MoveFirst
        While Not v_rs1.EOF
            l_lngWorkGroup = v_rs1!WG_NUM
            l_lngLine = v_rs1!LINE_NUM
            If IsInList(l_lngWorkGroup, l_colWorkGroups) Then
            Else
               l_colWorkGroups.Add l_lngWorkGroup
            End If
            If IsInList(l_lngLine, l_colLines) Then
            Else
               l_colLines.Add l_lngLine
            End If
            v_rs1.MoveNext
        Wend
    End If
    
    If RecordsExist(v_rs2) Then
        v_rs2.MoveFirst
        While Not v_rs2.EOF
            l_lngWorkGroup = v_rs2!WG_NUM
            l_lngLine = v_rs2!LINE_NUM
            If IsInList(l_lngWorkGroup, l_colWorkGroups) Then
            Else
               l_colWorkGroups.Add l_lngWorkGroup
            End If
            If IsInList(l_lngLine, l_colLines) Then
            Else
               l_colLines.Add l_lngLine
            End If
            v_rs2.MoveNext
        Wend
    End If
    
    If l_colWorkGroups.Count = 0 Then
        l_colWorkGroups.Add -1
        Set l_colLines = New Collection
        l_colLines.Add -1
    End If
    
    If l_colLines.Count = 0 Then
        l_colLines.Add -1
        Set l_colWorkGroups = New Collection
        l_colWorkGroups.Add -1
    End If
        
    m_aryDemandWorkGroups = ArrayFromCollection(l_colWorkGroups)
    m_aryDemandLines = ArrayFromCollection(l_colLines)
    
    Exit Sub
    
GetDemandLineAssignments_EH:
    TrsRaiseError m_strCMODULENAME, "GetDemandLineAssignments"
    
    
End Sub


Private Function SaveTradeDB(Optional ByVal v_varCascadeUpdates As Variant) As Boolean

    Dim l_obj As dbSchedTrade.cdSchedTrade

    On Error GoTo SaveTradeDB_EH
    
    Set l_obj = CreateObject("dbSchedTrade.cdSchedTrade")
    
    m_rsTradeHeader.Filter = ""
    m_rsTradeDetails.Filter = ""
    m_rsEmpTrade.Filter = ""
    m_rsLineTrade.Filter = ""
    SaveTradeDB = l_obj.SaveTrade(m_rsTradeHeader, m_rsTradeDetails, _
                m_rsEmpTrade, m_rsLineTrade, m_lngUserNum, v_varCascadeUpdates)
         
    Set l_obj = Nothing
    
    Exit Function
    
SaveTradeDB_EH:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "SaveTradeDB"

End Function

Private Sub GetStartAndEndDateOfTrade()

    Dim l_intDayOfWeek As Integer
    Dim l_lngNumOfWeeks As Integer

    On Error GoTo eh

    m_rsTradeDetails.Sort = "CHARGE_DATE"
    m_rsTradeDetails.MoveFirst
    m_datStartDate = m_rsTradeDetails!CHARGE_DATE
    m_rsTradeDetails.MoveLast
    m_datEndDate = m_rsTradeDetails!CHARGE_DATE
    m_rsTradeDetails.Sort = "EMP_NUM,CHARGE_DATE,SHIFT_NUM"
    
'All Trades are for a calendar week starting on Sunday
    l_intDayOfWeek = Weekday(m_datStartDate)
    m_datStartDate = DateAdd("d", 1 - l_intDayOfWeek, m_datStartDate)
    
    l_lngNumOfWeeks = DateDiff("ww", m_datStartDate, m_datEndDate)
    If l_lngNumOfWeeks > 0 Then
'need better error message
'Basically trade must be in same week
        Error.Raise 30044 + vbObjectError
    Else
        m_datEndDate = DateAdd("d", 6, m_datStartDate)
    End If

    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "GetStartAndEndDateOfTrade"


End Sub

Private Function GetAllTrades( _
    ) As ADODB.Recordset
    
    Dim l_obj As dbSchedTrade.cdSchedTrade
    
    On Error GoTo GetAllTrades_EH

    Set l_obj = CreateObject("dbSchedTrade.cdSchedTrade")
    
'    Set GetAllTrades = l_obj.GetTradesByEmployeeAndLine( _
'            m_aryEmployees, m_aryDemandWorkGroups, m_aryDemandLines _
'            , m_datStartDate, m_datEndDate)

    Set GetAllTrades = l_obj.GetTradesByEmployee( _
            m_aryEmployees, m_datStartDate, m_datEndDate)
         
    Set l_obj = Nothing
    
    Exit Function
    
GetAllTrades_EH:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "GetAllTrades"
    
End Function


Private Function ProcessDeleteTrade() As Boolean

    Dim l_lngWGNum As Long
    Dim l_lngLineNum As Long
    Dim l_blnDeleteOK As Boolean
    
    On Error GoTo eh


    l_blnDeleteOK = True
    m_rsTradeDetails.Filter = "TRADE_NUM = " & str(m_lngTradeNum)
    If Not RecordsExist(m_rsTradeDetails) Then
        ProcessDeleteTrade = True
        Exit Function
    End If
    
    m_rsTradeDetails.MoveFirst
    While Not m_rsTradeDetails.EOF
        l_lngWGNum = m_rsTradeDetails!WG_NUM
        l_lngLineNum = m_rsTradeDetails!LINE_NUM
        If l_lngWGNum > 0 And l_lngLineNum > 0 Then
            If Not ProcessDeleteLineTrade( _
                m_rsTradeDetails, l_lngWGNum, l_lngLineNum) Then
                l_blnDeleteOK = False
            End If
        Else
            l_lngWGNum = 0
            l_lngLineNum = 0
            If Not ProcessDeleteEmpTrade( _
                m_rsTradeDetails) Then
                l_blnDeleteOK = False
            End If
        End If
        AppendShiftModified _
            m_rsTradeDetails!EMP_NUM, m_rsTradeDetails!CHARGE_DATE, m_rsTradeDetails!SHIFT_NUM
        m_rsTradeDetails.Delete
        m_rsTradeDetails.MoveNext ' Is this going to mess me up
    Wend
    
    If l_blnDeleteOK Then
        m_rsTradeHeader.Filter = "TRADE_NUM = " & str(m_lngTradeNum)
        If RecordsExist(m_rsTradeHeader) Then
            m_rsTradeHeader.Delete
        End If
        ClearTrade
    End If
    
    ProcessDeleteTrade = l_blnDeleteOK
    
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "ProcessDeleteTrade"
    

End Function

Private Function ProcessDeleteLineTrade( _
    ByVal v_rsTradeDetails As ADODB.Recordset, _
    ByVal v_lngWGNum As Long, _
    ByVal v_lngLineNum As Long _
    ) As Boolean
    
    
    Dim l_datChargeDate As Date
    Dim l_lngShiftNum As Long
    Dim l_strFilter As String
    Dim l_blnDeleteOK As String
    Dim l_strErrorDesc As String
    Dim l_strEmployeeMSG As String
    Dim l_bln As String
       
    On Error GoTo eh
       
    l_blnDeleteOK = True
    l_datChargeDate = v_rsTradeDetails!CHARGE_DATE
    l_lngShiftNum = v_rsTradeDetails!SHIFT_NUM
    l_strFilter = "WG_NUM = " & str(v_lngWGNum) & _
            " and LINE_NUM = " & str(v_lngLineNum) & _
            " and CHARGE_DATE = " & l_datChargeDate & _
            " and SHIFT_NUM = " & str(l_lngShiftNum)
    l_strFilter = MDAC25FilterFix(l_strFilter)
    m_rsLineTrade.Filter = l_strFilter
    If RecordsExist(m_rsLineTrade) Then
        m_rsLineTrade.Delete
    End If
    l_strFilter = "WG_NUM = " & str(v_lngWGNum) & _
            " and LINE_NUM = " & str(v_lngLineNum) & _
            " and CHARGE_DATE = " & l_datChargeDate & _
            " and SHIFT_NUM = " & str(l_lngShiftNum) & _
            " and TRADE_NUM <> " & str(m_lngTradeNum)
    l_strFilter = MDAC25FilterFix(l_strFilter)
    m_rsOtherTradeDetails.Filter = l_strFilter
    m_rsOtherTradeDetails.Sort = "TRADE_NUM DESC"
    If RecordsExist(m_rsOtherTradeDetails) Then
        If m_rsOtherTradeDetails!TRADE_NUM > m_lngTradeNum Then
            l_blnDeleteOK = False
            l_strEmployeeMSG = Trim(m_rsOtherTradeDetails!EMP_NAME) & _
                " (" & Trim(str(m_rsOtherTradeDetails!EMP_NUM)) & ") " & _
                "is assigned to work group " & Trim(str(v_lngWGNum)) & _
                ",line " & Trim(str(v_lngLineNum)) & ".  "
            l_strErrorDesc = "Trade " & Trim(str(m_lngTradeNum)) & _
                " cannot be deleted because a subsequent trade " & _
                Trim(str(m_rsOtherTradeDetails!TRADE_NUM)) & " " & _
                "exists for the line.  " & _
                "In order to delete trade " & Trim(str(m_lngTradeNum)) & "" & _
                ", you must include trade " & _
                Trim(str(m_rsOtherTradeDetails!TRADE_NUM)) & " " & _
                "in the delete request " & _
                l_strEmployeeMSG

'            l_strErrorDesc = "Trade cannot be delete because employee (" & _
                str(m_rsOtherTradeDetails!EMP_NUM) & ") " & _
                "has made another trade after the trade selected for deletion. " & _
                "In order to delete trade (" & str(m_lngTradeNum) & ")" & _
                ", you must first delete trade (" & _
                str(m_rsOtherTradeDetails!TRADE_NUM) & "). "
                
            l_bln = AppendTradeDeleteBRV( _
                    m_rsOtherTradeDetails, trErrbusValidationError, _
                    mstrCBUSINESSERR, l_strErrorDesc)
        Else
            m_rsLineTrade.AddNew
            m_rsLineTrade!WG_NUM = v_lngWGNum
            m_rsLineTrade!LINE_NUM = v_lngLineNum
            m_rsLineTrade!CHARGE_DATE = l_datChargeDate
            m_rsLineTrade!SHIFT_NUM = l_lngShiftNum
            m_rsLineTrade!SCHED_CD = m_rsOtherTradeDetails!SCHED_CD
            m_rsLineTrade!TRADE_NUM = m_rsOtherTradeDetails!TRADE_NUM
            m_rsLineTrade!CREATED_EMP_NUM = CLng(m_strUserId)
            m_rsLineTrade.Update
        End If
    End If
    
    ProcessDeleteLineTrade = l_blnDeleteOK
    Exit Function
   
eh:
    TrsRaiseError m_strCMODULENAME, "ProcessDeleteTrade"
    
End Function

Private Function ProcessDeleteEmpTrade( _
    ByVal v_rsTradeDetails As ADODB.Recordset _
    ) As Boolean
    
    
    Dim l_datChargeDate As Date
    Dim l_lngEmpNum As Long
    Dim l_lngShiftNum As Long
    Dim l_strFilter As String
    Dim l_blnDeleteOK As String
    Dim l_strErrorDesc As String
    Dim l_bln As String
       
    On Error GoTo eh
       
    l_blnDeleteOK = True
    l_lngEmpNum = v_rsTradeDetails!EMP_NUM
    l_datChargeDate = v_rsTradeDetails!CHARGE_DATE
    l_lngShiftNum = v_rsTradeDetails!SHIFT_NUM
    l_strFilter = "EMP_NUM = " & str(l_lngEmpNum) & _
            " and CHARGE_DATE = " & l_datChargeDate & _
            " and SHIFT_NUM = " & str(l_lngShiftNum)
    l_strFilter = MDAC25FilterFix(l_strFilter)
    m_rsEmpTrade.Filter = l_strFilter
    If RecordsExist(m_rsEmpTrade) Then
        m_rsEmpTrade.Delete
    End If
    l_strFilter = "EMP_NUM = " & str(l_lngEmpNum) & _
            " and CHARGE_DATE = " & l_datChargeDate & _
            " and SHIFT_NUM = " & str(l_lngShiftNum) & _
            " and TRADE_NUM <> " & str(m_lngTradeNum)
    l_strFilter = MDAC25FilterFix(l_strFilter)
    m_rsOtherTradeDetails.Filter = l_strFilter
    m_rsOtherTradeDetails.Sort = "TRADE_NUM DESC"
    If RecordsExist(m_rsOtherTradeDetails) Then
        If m_rsOtherTradeDetails!TRADE_NUM > m_lngTradeNum Then
            l_blnDeleteOK = False
            l_strErrorDesc = "Trade " & Trim(str(m_lngTradeNum)) & _
                " cannot be deleted because a subsequent trade " & _
                Trim(str(m_rsOtherTradeDetails!TRADE_NUM)) & " " & _
                "involving " & Chr(34) & _
                Trim(m_rsOtherTradeDetails!EMP_NAME) & Chr(34) & "(" & _
                Trim(str(m_rsOtherTradeDetails!EMP_NUM)) & ") " & _
                "exists. " & _
                "In order to delete trade " & Trim(str(m_lngTradeNum)) & "" & _
                ", you must include trade " & _
                Trim(str(m_rsOtherTradeDetails!TRADE_NUM)) & " " & _
                "in the delete request "
            l_bln = AppendTradeDeleteBRV( _
                    m_rsOtherTradeDetails, trErrbusValidationError, _
                    mstrCBUSINESSERR, l_strErrorDesc)
        Else
            m_rsEmpTrade.AddNew
            m_rsEmpTrade!EMP_NUM = l_lngEmpNum
            m_rsEmpTrade!CHARGE_DATE = l_datChargeDate
            m_rsEmpTrade!SHIFT_NUM = l_lngShiftNum
            m_rsEmpTrade!SCHED_CD = m_rsOtherTradeDetails!SCHED_CD
            m_rsEmpTrade!TRADE_NUM = m_rsOtherTradeDetails!TRADE_NUM
            m_rsEmpTrade!CREATED_EMP_NUM = CLng(m_strUserId)
            m_rsEmpTrade.Update
        End If
    End If
    
    ProcessDeleteEmpTrade = l_blnDeleteOK
    Exit Function
   
eh:
    TrsRaiseError m_strCMODULENAME, "ProcessDeleteEmpTrade"
    
End Function

Private Function AppendTradeDeleteBRV( _
            ByVal v_rsTradeDetail As ADODB.Recordset, _
            ByVal v_lngErrorNumber As Long, _
            ByVal v_strErrorType As String, _
            ByVal v_strErrorDesc As String _
            ) As Boolean
            
    Dim l_bolError As Boolean
    Dim lngKeyEmpNum As Long
    Dim lngKeyShiftNum As Long
    Dim l_datKeyChargeDate As Date
    Dim l_strKeyCreatedTime As String
    Dim lngKeyWGNum As Long
    Dim lngKeyLineNum As Long
    Dim lngKeyTradeNum As Long
    Dim l_strFilter As String
    Dim l_strOtherEmpNum As String
    Dim l_strOtherEmpName As String
    
    On Error GoTo AppendTradeDeleteBRV_EH
    
    lngKeyEmpNum = v_rsTradeDetail!EMP_NUM
    lngKeyShiftNum = v_rsTradeDetail!SHIFT_NUM
    l_datKeyChargeDate = v_rsTradeDetail!CHARGE_DATE
    lngKeyWGNum = v_rsTradeDetail!WG_NUM
    lngKeyLineNum = v_rsTradeDetail!LINE_NUM
    lngKeyTradeNum = v_rsTradeDetail!TRADE_NUM
    
    l_strFilter = "TRADE_NUM = " & str(lngKeyTradeNum) & _
         " AND EMP_NUM <> " & str(lngKeyEmpNum)
    l_strFilter = MDAC25FilterFix(l_strFilter)
    v_rsTradeDetail.Filter = l_strFilter
    If RecordsExist(v_rsTradeDetail) Then
        v_rsTradeDetail.MoveFirst
        l_strOtherEmpNum = Trim(str(v_rsTradeDetail!EMP_NUM))
        l_strOtherEmpName = Trim(v_rsTradeDetail!EMP_NAME)
    Else
        l_strOtherEmpNum = Trim(str(lngKeyEmpNum))
        l_strOtherEmpName = "Self Trade"
    End If

    l_bolError = AppendBusinessError(m_rsErrors, _
            v_lngErrorNumber, _
            v_strErrorType, _
            v_strErrorDesc, _
            m_strCMODULENAME, _
            TrType_TR_DEL, , , "EMP_NUM;CHARGE_DATE,SHIFT_NUM,WG_NUM,LINE_NUM,TRADE_NUM", _
            Trim(str(lngKeyEmpNum)), _
            Trim(str(l_datKeyChargeDate)), _
            Trim(str(lngKeyShiftNum)), _
            Trim(str(lngKeyWGNum)), _
            Trim(str(lngKeyLineNum)), _
            Trim(str(m_lngTradeNum)), _
            Trim(str(lngKeyTradeNum)), _
            l_strOtherEmpNum, _
            l_strOtherEmpName)

    AppendTradeDeleteBRV = l_bolError
    Exit Function
    
AppendTradeDeleteBRV_EH:
    TrsRaiseError m_strCMODULENAME, "AppendTradeDeleteBRV"

End Function

Private Function ValidRoleTypes( _
    ByVal v_rsTradeDetails As ADODB.Recordset _
    ) As Boolean
    
    Dim l_strRole As String
    Dim l_blnValid As Boolean
    Dim l_strErrorDesc As String
    Dim l_bln As Boolean
    
    On Error GoTo eh
    
    l_blnValid = True
    v_rsTradeDetails.MoveFirst
    While Not v_rsTradeDetails.EOF
        If IsNull(v_rsTradeDetails!ROLE_TYPE_CD) Or IsEmpty(v_rsTradeDetails!ROLE_TYPE_CD) Then
            l_blnValid = False
            l_strErrorDesc = "ROLE Type missing.  Please contact TRIS Development team"
            l_bln = AppendTradeDetailBRV( _
                        v_rsTradeDetails, _
                        trErrbusValidationError, _
                        mstrCAPPLICATIONERR, _
                        l_strErrorDesc)
        Else
            l_strRole = v_rsTradeDetails!ROLE_TYPE_CD
            Select Case l_strRole
                Case "A", "R", "S"
                
                Case Else
                    l_blnValid = False
                    l_strErrorDesc = "ROLE Type (" & l_strRole & ") in error.  Please contact TRIS Development team"
                    l_bln = AppendTradeDetailBRV( _
                        v_rsTradeDetails, _
                        trErrbusValidationError, _
                        mstrCAPPLICATIONERR, _
                        l_strErrorDesc)
            End Select
        v_rsTradeDetails.MoveNext
        End If
    Wend
    
    ValidRoleTypes = l_blnValid
    
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "ValidRoleTypes"

End Function

Private Function ValidScheduleCodes( _
    ByVal v_rsTradeDetails As ADODB.Recordset _
    ) As Boolean
    
    Dim l_lngSchedCd As Long
    Dim l_blnValid As Boolean
    Dim l_strErrorDesc As String
    Dim l_bln As Boolean
    
    On Error GoTo eh
    
    l_blnValid = True
    v_rsTradeDetails.MoveFirst
    While Not v_rsTradeDetails.EOF
        If IsNull(v_rsTradeDetails!SCHED_CD) Or IsEmpty(v_rsTradeDetails!SCHED_CD) Then
            l_blnValid = False
            l_strErrorDesc = "Schedule code missing.  Please contact TRIS Development team"
            l_bln = AppendTradeDetailBRV( _
                        v_rsTradeDetails, _
                        trErrbusValidationError, _
                        mstrCAPPLICATIONERR, _
                        l_strErrorDesc)
        Else
            l_lngSchedCd = v_rsTradeDetails!SCHED_CD
            If IsValidScheduleCode(l_lngSchedCd) Then
            Else
                l_blnValid = False
                l_strErrorDesc = "Schedule code (" & str(l_lngSchedCd) & ") is in error.  Please contact TRIS Development team"
                l_bln = AppendTradeDetailBRV( _
                    v_rsTradeDetails, _
                    trErrbusValidationError, _
                    mstrCAPPLICATIONERR, _
                    l_strErrorDesc)
            End If
        End If
        v_rsTradeDetails.MoveNext
    Wend
    
    ValidScheduleCodes = l_blnValid
    
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "ValidScheduleCodes"

End Function

Private Function IsValidScheduleCode( _
    ByVal v_lngSchedCd As Long _
    ) As Boolean
    
    Dim l_obj As dbCodeTables.cdCodeTables
    
    On Error GoTo eh
    
    If v_lngSchedCd = 0 Then
        IsValidScheduleCode = False
    ElseIf v_lngSchedCd = -2 Or v_lngSchedCd = -9 Then
        IsValidScheduleCode = True
    Else
        Set l_obj = New dbCodeTables.cdCodeTables
        IsValidScheduleCode = l_obj.ValidateScheduleCode(v_lngSchedCd)
        Set l_obj = Nothing
    End If
    
    Exit Function
eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "IsValidScheduleCode"
    
End Function

Private Function ValidShiftNumbers( _
    ByVal v_rsTradeDetails As ADODB.Recordset _
    ) As Boolean
    
    Dim l_intShiftNum As Integer
    Dim l_blnValid As Boolean
    Dim l_strErrorDesc As String
    Dim l_bln As Boolean
    
    On Error GoTo eh
    
    l_blnValid = True
    v_rsTradeDetails.MoveFirst
    While Not v_rsTradeDetails.EOF
        If IsNull(v_rsTradeDetails!SHIFT_NUM) Or IsEmpty(v_rsTradeDetails!SHIFT_NUM) Then
            l_blnValid = False
            l_strErrorDesc = "Shift number is missing.  Please contact TRIS Development team"
            l_bln = AppendTradeDetailBRV( _
                        v_rsTradeDetails, _
                        trErrbusValidationError, _
                        mstrCAPPLICATIONERR, _
                        l_strErrorDesc)
        Else
            l_intShiftNum = v_rsTradeDetails!SHIFT_NUM
            Select Case l_intShiftNum
                Case 1, 2
                
                Case Else
                    l_blnValid = False
                    l_strErrorDesc = "Shift number (" & str(l_intShiftNum) & ") is in error.  Please contact TRIS Development team"
                    l_bln = AppendTradeDetailBRV( _
                        v_rsTradeDetails, _
                        trErrbusValidationError, _
                        mstrCAPPLICATIONERR, _
                        l_strErrorDesc)
            End Select
        End If
        v_rsTradeDetails.MoveNext
    Wend
    
    ValidShiftNumbers = l_blnValid
    
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "ValidShiftNumbers"

End Function

Private Function ValidTradeType( _
    ByVal v_rsTradeDetails As ADODB.Recordset _
    ) As Boolean
    
    Dim l_strTradeType As String
    Dim l_blnValid As Boolean
    Dim l_strErrorDesc As String
    Dim l_bln As Boolean
    
    On Error GoTo eh
    
    l_blnValid = True
    v_rsTradeDetails.MoveFirst
    While Not v_rsTradeDetails.EOF
        If IsNull(v_rsTradeDetails!WEEK_TRADE_IND) Or IsEmpty(v_rsTradeDetails!WEEK_TRADE_IND) Then
            l_blnValid = False
            l_strErrorDesc = "Trade Type missing.  Please contact TRIS Development team"
            l_bln = AppendTradeDetailBRV( _
                        v_rsTradeDetails, _
                        trErrbusValidationError, _
                        mstrCAPPLICATIONERR, _
                        l_strErrorDesc)
        Else
            l_strTradeType = v_rsTradeDetails!WEEK_TRADE_IND
            Select Case l_strTradeType
                Case "Y", "N"
                    ' Y for Week Trade
                    ' N for Not a Week Trade
                
                Case Else
                    l_blnValid = False
                    l_strErrorDesc = "Trade Type (" & l_strTradeType & ") in error.  Please contact TRIS Development team"
                    l_bln = AppendTradeDetailBRV( _
                        v_rsTradeDetails, _
                        trErrbusValidationError, _
                        mstrCAPPLICATIONERR, _
                        l_strErrorDesc)
            End Select
        End If
        v_rsTradeDetails.MoveNext
    Wend
    
    ValidTradeType = l_blnValid
    
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "ValidTradeType"

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
    TrsRaiseError m_strCMODULENAME, "GetEmployeeData"


End Function

Private Function EmployeesExist() As Boolean

    Dim l_intI As Integer
    Dim l_intCnt As Integer
    Dim l_blnExists As Boolean
    Dim l_strEmp As String
    
    On Error GoTo eh
    
    m_aryEmployees = ArrayFromCollection(m_colEmployees)
    Set m_rsEmployeeData = GetEmployeeData(m_aryEmployees)
    
    l_blnExists = True
    l_intCnt = m_colEmployees.Count
    For l_intI = 1 To l_intCnt
        l_strEmp = m_colEmployees.Item(l_intI)
        m_rsEmployeeData.Filter = "EMP_NUM = " & l_strEmp
        If Not RecordsExist(m_rsEmployeeData) Then
            l_blnExists = False
            AppendTradeBRV trErrbusValidationError, _
                mstrCAPPLICATIONERR, _
                "System Error. Employee (" & l_strEmp & ") does not exist.", _
                CLng(l_strEmp)
        End If
    Next l_intI
    m_rsEmployeeData.Filter = ""
    
    EmployeesExist = l_blnExists
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "EmployeesExist"
    
End Function

Private Function ValidChargeDates( _
    ByVal v_rsTradeDetails As ADODB.Recordset _
    ) As Boolean
    
    Dim l_blnValid As Boolean
    Dim l_strErrorDesc As String
    Dim l_bln As Boolean
    
    On Error GoTo eh
    
    l_blnValid = True
    v_rsTradeDetails.MoveFirst
    While Not v_rsTradeDetails.EOF
        If IsNull(v_rsTradeDetails!CHARGE_DATE) Or IsEmpty(v_rsTradeDetails!CHARGE_DATE) Then
            l_blnValid = False
            l_strErrorDesc = "Charge Date is missing.  Please contact TRIS Development team"
            l_bln = AppendTradeDetailBRV( _
                        v_rsTradeDetails, _
                        trErrbusValidationError, _
                        mstrCAPPLICATIONERR, _
                        l_strErrorDesc)
        Else
            If Not IsDate(v_rsTradeDetails!CHARGE_DATE) Then
                    l_blnValid = False
                    l_strErrorDesc = "Charge Date is not a valid date.  Please contact TRIS Development team"
                    l_bln = AppendTradeDetailBRV( _
                        v_rsTradeDetails, _
                        trErrbusValidationError, _
                        mstrCAPPLICATIONERR, _
                        l_strErrorDesc)
            End If
        End If
        v_rsTradeDetails.MoveNext
    Wend
    
    If l_blnValid Then
        If Not IsTradeInSameWeek(v_rsTradeDetails) Then
            l_blnValid = False
        End If
    End If

    ValidChargeDates = l_blnValid
    
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "ValidChargeDates"

End Function


Private Function IsTradeInSameWeek( _
    ByVal v_rsTradeDetails As ADODB.Recordset _
    ) As Boolean
    
'Checks to see if trade falls in the same calendar week
'It uses the Min and Max Charge Dates in the trade
'If the number of weeks is greater than zero they do not fall
' in the same week.
'Please note that the VB DateDiff function is Sunday based
'If the Min Date is not a Sunday, the VB function uses the Sunday
' prior to the Min Date

    Dim l_intNumWeeks As Integer
    Dim l_strErrorDesc As String
    Dim l_datMinDate As Date
    Dim l_datMaxDate As Date
    
    On Error GoTo eh
    

    v_rsTradeDetails.Sort = "CHARGE_DATE"
    v_rsTradeDetails.MoveFirst
    l_datMinDate = v_rsTradeDetails!CHARGE_DATE
    v_rsTradeDetails.MoveLast
    l_datMaxDate = v_rsTradeDetails!CHARGE_DATE
    
    l_intNumWeeks = DateDiff("w", l_datMinDate, l_datMaxDate)
    
    If l_intNumWeeks > 0 Then
        IsTradeInSameWeek = False
        l_strErrorDesc = "Trade is not within the same calendar week. " & _
            "  The first date in the trade is " & Format(l_datMinDate, "dd/mm/yyyy") & _
            "  The last date in the trade is " & Format(l_datMaxDate, "dd/mm/yyyy") & _
            "  Please contact TRIS Development team."
        AppendTradeBRV trErrbusValidationError, _
            mstrCAPPLICATIONERR, _
            l_strErrorDesc
    Else
        IsTradeInSameWeek = True
    End If

    v_rsTradeDetails.Sort = ""
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "IsTradeInSameWeek"

    
End Function

Private Sub DoEmployeesHaveSaveJob( _
    ByVal v_rsTradeDetails As ADODB.Recordset, _
    ByVal v_colChargeDates As Collection)

    Dim l_strErrorDesc As String
    Dim l_datChargeDate As Date
    Dim l_strReqJob As String
    Dim l_strAccJob As String
    Dim l_intI As Integer
    Dim l_intCount As Integer
    Dim l_bln As Boolean
    
    On Error GoTo eh
    

    If m_colEmployees.Count < 2 Then
        Exit Sub
    End If
    
    l_intCount = v_colChargeDates.Count
    For l_intI = 1 To l_intCount
        l_datChargeDate = v_colChargeDates.Item(l_intI)
        l_strReqJob = GetEmployeeJob(l_datChargeDate, m_rsRequestorSchedule)
        l_strAccJob = GetEmployeeJob(l_datChargeDate, m_rsAcceptorSchedule)
        If l_strReqJob <> l_strAccJob Then
            l_strErrorDesc = _
                "The requestor's (" & str(m_colEmployees.Item(1)) & _
                ") job number (" & l_strReqJob & ") " & _
                "does not match the acceptor's (" & str(m_colEmployees.Item(2)) & _
                ") job number (" & l_strAccJob & ")."
            v_rsTradeDetails.Filter = "CHARGE_DATE = " & str(l_datChargeDate)
            l_bln = AppendTradeDetailBRV( _
                        v_rsTradeDetails, _
                        20000, _
                        mstrCBUSINESSWAR, _
                        l_strErrorDesc)
        End If
    Next l_intI

    v_rsTradeDetails.Filter = ""
    Exit Sub
eh:
    TrsRaiseError m_strCMODULENAME, "DoEmployeesHaveSaveJob"

    
End Sub

Private Function IsDateInList( _
        ByVal v_datDate As Date, _
        ByVal v_colCodes As Collection _
        ) As Boolean

    Dim l_intIndex As Integer
    Dim l_intCount As Integer
    Dim l_bolInList As Boolean
    
    On Error GoTo eh
    
    l_bolInList = False
    l_intIndex = 1
    l_intCount = v_colCodes.Count
    
    While Not l_bolInList And l_intIndex <= l_intCount
        If CDate(v_colCodes.Item(l_intIndex)) = v_datDate Then
            l_bolInList = True
        End If
        l_intIndex = l_intIndex + 1
    Wend
    
    IsDateInList = l_bolInList
    
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "IsDateInList"
    

End Function

Private Function GetChargeDatesInRequest( _
    ByVal v_objRS As ADODB.Recordset _
    ) As Collection


    Dim l_colChargeDates As Collection
    Dim l_datChargeDate As Date
    
    On Error GoTo eh
 
 
    Set l_colChargeDates = New Collection
    v_objRS.Sort = "CHARGE_DATE"
    v_objRS.MoveFirst
    While Not v_objRS.EOF
        l_datChargeDate = v_objRS!CHARGE_DATE
        If Not IsDateInList(l_datChargeDate, l_colChargeDates) Then
            l_colChargeDates.Add l_datChargeDate
        End If
        v_objRS.MoveNext
    Wend
    
    Set GetChargeDatesInRequest = l_colChargeDates
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "GetChargeDatesInRequest"
    
End Function
    
    
Private Function GetEmployeeJob( _
    ByVal v_datChargeDate As Date, _
    ByVal v_rsSchedule As ADODB.Recordset _
    ) As String
    
    Dim l_strSchedJob As String
    Dim l_strEmpJob As String
    
    On Error GoTo eh
    
    If Not RecordsExist(v_rsSchedule) Then
        GetEmployeeJob = ""
        Exit Function
    End If
    
    v_rsSchedule.Filter = "CHARGE_DATE = " & str(v_datChargeDate)
    
    If Not RecordsExist(v_rsSchedule) Then
        GetEmployeeJob = ""
        Exit Function
    End If
    l_strSchedJob = v_rsSchedule!SCHED_JOB_NUM
    l_strEmpJob = v_rsSchedule!EMP_JOB_NUM
    If Trim(l_strSchedJob) <> "" Then
        GetEmployeeJob = Trim(l_strSchedJob)
    Else
        GetEmployeeJob = Trim(l_strEmpJob)
    End If

    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "GetEmployeeJob"
    
End Function

Private Sub DoEmployeesHaveSaveUnion( _
    ByVal v_rsTradeDetails As ADODB.Recordset, _
    ByVal v_colChargeDates As Collection)

    Dim l_strErrorDesc As String
    Dim l_datChargeDate As Date
    Dim l_strReqUnion As String
    Dim l_strAccUnion As String
    Dim l_intI As Integer
    Dim l_intCount As Integer
    Dim l_bln As Boolean
    
    On Error GoTo eh
    

    If m_colEmployees.Count < 2 Then
        Exit Sub
    End If
    
    l_intCount = v_colChargeDates.Count
    For l_intI = 1 To l_intCount
        l_datChargeDate = v_colChargeDates.Item(l_intI)
        l_strReqUnion = Left(GetEmployeeJob(l_datChargeDate, m_rsRequestorSchedule), 2)
        l_strAccUnion = Left(GetEmployeeJob(l_datChargeDate, m_rsAcceptorSchedule), 2)
        If l_strReqUnion <> l_strAccUnion Then
            l_strErrorDesc = _
                "The requestor's (" & str(m_colEmployees.Item(1)) & _
                ") union (" & l_strReqUnion & ") " & _
                "does not match the acceptor's (" & str(m_colEmployees.Item(2)) & _
                ") union (" & l_strAccUnion & ")."
            v_rsTradeDetails.Filter = "CHARGE_DATE = " & str(l_datChargeDate)
            l_bln = AppendTradeDetailBRV( _
                        v_rsTradeDetails, _
                        20000, _
                        mstrCBUSINESSWAR, _
                        l_strErrorDesc)
        End If
    Next l_intI

    v_rsTradeDetails.Filter = ""
    Exit Sub
eh:
    TrsRaiseError m_strCMODULENAME, "DoEmployeesHaveSaveUnion"

    
End Sub


Private Function ValidGABSchedules() As Boolean

    Dim l_intCount As Integer
    Dim l_blnSchedsOK As Boolean
    Dim l_datChargeDate As Date
    Dim l_intI As Integer
    
    On Error GoTo eh
    
    l_intCount = m_colEmployees.Count
    
    If l_intCount < 1 Then
        ValidGABSchedules = True
        Exit Function
    End If
    
    l_blnSchedsOK = True
    m_rsRequestorSchedule.Filter = ""
    If Not RecordsExist(m_rsRequestorSchedule) Then
            l_blnSchedsOK = False
            AppendTradeBRV trErrbusValidationError, _
            mstrCAPPLICATIONERR, _
            "System Error. Requestor's Schedule not in request.  " & _
            "Please contact TRIS Development.", _
            CLng(m_colEmployees.Item(1))
    End If
    
    If l_intCount = 2 Then
        m_rsAcceptorSchedule.Filter = ""
        If Not RecordsExist(m_rsAcceptorSchedule) Then
            l_blnSchedsOK = False
            AppendTradeBRV trErrbusValidationError, _
            mstrCAPPLICATIONERR, _
            "System Error. Acceptor's Schedule not in request.  " & _
            "Please contact TRIS Development.", _
            CLng(m_colEmployees.Item(2))
        End If
    End If
    
    If Not l_blnSchedsOK Then
        ValidGABSchedules = False
        Exit Function
    End If
    
    'Is the requestor's Schedule Available for the week of the trade
    For l_intI = 0 To 6
        l_datChargeDate = DateAdd("d", l_intI, m_datStartDate)
        m_rsRequestorSchedule.Filter = "CHARGE_DATE = " & str(l_datChargeDate)
        If Not RecordsExist(m_rsRequestorSchedule) Then
            l_blnSchedsOK = False
            AppendTradeBRV trErrbusValidationError, _
                mstrCAPPLICATIONERR, _
                "System Error. Requestor's Schedule missing.  " & _
                " on charge date (" & Format("dd/mm/yyyy", l_datChargeDate) & _
                ").  Please contact TRIS Development.", _
            CLng(m_colEmployees.Item(1))
        End If
    Next l_intI
    
    If l_intCount = 2 Then
    'Is the acceptor's Schedule Available for the week of the trade
        For l_intI = 0 To 6
            l_datChargeDate = DateAdd("d", l_intI, m_datStartDate)
            m_rsRequestorSchedule.Filter = "CHARGE_DATE = " & str(l_datChargeDate)
            If Not RecordsExist(m_rsRequestorSchedule) Then
                l_blnSchedsOK = False
                AppendTradeBRV trErrbusValidationError, _
                    mstrCAPPLICATIONERR, _
                    "System Error. Requestor's Schedule missing.  " & _
                    " on charge date (" & Format("dd/mm/yyyy", l_datChargeDate) & _
                    ").  Please contact TRIS Development.", _
                    CLng(m_colEmployees.Item(1))
            End If
        Next l_intI
    End If

    ValidGABSchedules = l_blnSchedsOK
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "ValidGABSchedules"

    

End Function



Private Function GetEmployeeWorkGroup( _
    ByVal v_lngEmpNum As Long, _
    ByVal v_datChargeDate As Date, _
    ByVal v_rsSchedule As ADODB.Recordset, _
    ByVal v_rsWorkGroups As ADODB.Recordset, _
    ByRef r_lngSecondaryWG As Long _
    ) As Long
    
    Dim l_lngSchWG As Long
    Dim l_lngPrimaryWG As Long
    Dim l_lngSecondaryWG As Long
    Dim l_strFilter As String
    
    On Error GoTo eh
    
    If Not RecordsExist(v_rsSchedule) Or Not RecordsExist(v_rsWorkGroups) Then
        GetEmployeeWorkGroup = 0
        r_lngSecondaryWG = 0
        Exit Function
    End If
    
    v_rsSchedule.Filter = "CHARGE_DATE = " & str(v_datChargeDate)
    
    If Not RecordsExist(v_rsSchedule) Then
        GetEmployeeWorkGroup = 0
        r_lngSecondaryWG = 0
        Exit Function
    End If
    l_lngSchWG = v_rsSchedule!WG_NUM
    
    l_strFilter = "EMP_NUM = " & str(v_lngEmpNum) & _
        " AND PRIMARY_IND = 'Y' " & _
        " AND EFF_START_DATE <= " & str(v_datChargeDate) & _
        " AND EFF_END_DATE >= " & str(v_datChargeDate)
        
    l_strFilter = MDAC25FilterFix(l_strFilter)
    v_rsWorkGroups.Filter = l_strFilter
    If Not RecordsExist(v_rsWorkGroups) Then
        l_lngPrimaryWG = 0
    Else
        l_lngPrimaryWG = v_rsWorkGroups!WG_NUM
    End If
    
    l_strFilter = "EMP_NUM = " & str(v_lngEmpNum) & _
        " AND PRIMARY_IND = 'N' " & _
        " AND EFF_START_DATE <= " & str(v_datChargeDate) & _
        " AND EFF_END_DATE >= " & str(v_datChargeDate)
        
    l_strFilter = MDAC25FilterFix(l_strFilter)
    v_rsWorkGroups.Filter = l_strFilter
    If Not RecordsExist(v_rsWorkGroups) Then
        l_lngSecondaryWG = 0
    Else
        l_lngSecondaryWG = v_rsWorkGroups!WG_NUM
    End If
    
    r_lngSecondaryWG = l_lngSecondaryWG
    If l_lngSchWG > 0 Then
        GetEmployeeWorkGroup = l_lngSchWG
    Else
        GetEmployeeWorkGroup = l_lngPrimaryWG
    End If

    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "GetEmployeeWorkGroup"
    
End Function

Private Sub DoEmployeesHaveSaveWorkGroup( _
    ByVal v_rsTradeDetails As ADODB.Recordset, _
    ByVal v_colChargeDates As Collection)

    Dim l_strErrorDesc As String
    Dim l_datChargeDate As Date
    Dim l_intI As Integer
    Dim l_intCount As Integer
    Dim l_bln As Boolean
    Dim l_rsWorkGroups As ADODB.Recordset
    Dim l_lngReqEmp As Long
    Dim l_lngAccEmp As Long
    Dim l_lngReqEmpWG As Long
    Dim l_lngAccEmpWG As Long
    Dim l_lngReqEmpSecondaryWG As Long
    Dim l_lngAccEmpSecondaryWG As Long
    
    On Error GoTo eh
    

    If m_colEmployees.Count < 2 Then
        Exit Sub
    End If
    l_lngReqEmp = m_colEmployees.Item(1)
    l_lngAccEmp = m_colEmployees.Item(2)
    Set l_rsWorkGroups = GetEmployeeWorkGroups(m_aryEmployees)
    l_intCount = v_colChargeDates.Count
    For l_intI = 1 To l_intCount
        l_datChargeDate = v_colChargeDates.Item(l_intI)
        l_lngReqEmpWG = GetEmployeeWorkGroup(l_lngReqEmp, l_datChargeDate, m_rsRequestorSchedule, l_rsWorkGroups, l_lngReqEmpSecondaryWG)
        l_lngAccEmpWG = GetEmployeeWorkGroup(l_lngAccEmp, l_datChargeDate, m_rsAcceptorSchedule, l_rsWorkGroups, l_lngAccEmpSecondaryWG)
        If l_lngReqEmpWG <> l_lngAccEmpWG Then
            If l_lngReqEmpWG = l_lngAccEmpSecondaryWG _
                Or l_lngAccEmpWG = l_lngReqEmpSecondaryWG Then
            Else
            l_strErrorDesc = _
                "The requestor's (" & str(l_lngReqEmp) & _
                ") work group (" & str(l_lngReqEmpWG) & ") " & _
                "does not match the acceptor's (" & str(l_lngAccEmp) & _
                ") work group (" & str(l_lngAccEmpWG) & ")."
            v_rsTradeDetails.Filter = "CHARGE_DATE = " & str(l_datChargeDate)
            l_bln = AppendTradeDetailBRV( _
                        v_rsTradeDetails, _
                        20000, _
                        mstrCBUSINESSWAR, _
                        l_strErrorDesc)
            End If
        End If
    Next l_intI

    v_rsTradeDetails.Filter = ""
    Exit Sub
eh:
    TrsRaiseError m_strCMODULENAME, "DoEmployeesHaveSaveWorkGroup"

    
End Sub

Private Function GetEmployeeWorkGroups( _
            ByVal v_aryEmployees As Variant _
            ) As ADODB.Recordset

    On Error GoTo eh


    Dim l_obj As dbWorkGroup.cdWorkGroup
         
    
    Set l_obj = CreateObject("dbWorkGroup.cdWorkGroup")
    
    Set GetEmployeeWorkGroups = l_obj.GetEmployeeWorkGroups(v_aryEmployees, m_datStartDate, m_datEndDate)
         
    Set l_obj = Nothing
    
    Exit Function
    
eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "GetEmployeeWorkGroups"


End Function

Private Function CountNumberOfDoubleShifts( _
    ByVal v_rsSchedule As ADODB.Recordset _
    ) As Integer
    
    'This rountine counts the number of double shifts
    'that the employee has within the week of the trade
    'The schedule passed is the employee's schedule
    'after the trade has been applied
    
    'There should always be a shift 1 for all dates
    '  , though this code did not make that assumption
    'If a shift does not have any hours
    ', then the employee is off for that shift
    
    Dim l_intShiftNum As Integer
    Dim l_datChargeDate As Date
    Dim l_intNumberOfDoubleShifts As Integer
    Dim l_decShiftHours As Variant
    
    On Error GoTo eh
      
    l_intNumberOfDoubleShifts = 0
    v_rsSchedule.Sort = "CHARGE_DATE, SHIFT_NUM DESC"
    If Not RecordsExist(v_rsSchedule) Then
        CountNumberOfDoubleShifts = 0
        Exit Function
    End If
    v_rsSchedule.MoveFirst
    While Not v_rsSchedule.EOF
        l_intShiftNum = v_rsSchedule!SHIFT_NUM
        If l_intShiftNum = 1 Then
            v_rsSchedule.MoveNext
        Else
            l_datChargeDate = v_rsSchedule!CHARGE_DATE
            l_decShiftHours = v_rsSchedule!SHIFT_HOURS
            v_rsSchedule.MoveNext
            If l_decShiftHours > 0 Then
                If Not v_rsSchedule.EOF Then
                    If l_datChargeDate = v_rsSchedule!CHARGE_DATE Then
                        If v_rsSchedule!SHIFT_HOURS > 0 Then
                            l_intNumberOfDoubleShifts = l_intNumberOfDoubleShifts + 1
                        End If
                    End If
                End If
            End If
        End If
    Wend
    
    CountNumberOfDoubleShifts = l_intNumberOfDoubleShifts

    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "CountNumberOfDoubleShifts"
    
    
    End Function
    

Private Function CountScheduleHours( _
    ByVal v_rsSchedule As ADODB.Recordset, _
    Optional v_blnNewSched As Boolean = False _
    ) As Variant

    Dim l_decHours As Variant
    
    l_decHours = CDec(0)
    If Not RecordsExist(v_rsSchedule) Then
        CountScheduleHours = l_decHours
        Exit Function
    End If
        
    v_rsSchedule.MoveFirst
    While Not v_rsSchedule.EOF
        If v_blnNewSched Then
            l_decHours = l_decHours + CDec(v_rsSchedule!SHIFT_HOURS)
        Else
            If v_rsSchedule!SYSTEM_EXC_CD <> "SLT" _
                And v_rsSchedule!SYSTEM_EXC_CD <> "SSO" Then
                l_decHours = l_decHours + CDec(v_rsSchedule!HOURS)
            End If
        End If
        v_rsSchedule.MoveNext
    Wend
    
    CountScheduleHours = l_decHours
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "CountScheduleHours"
    
End Function

Private Function CreateNewSchedRecordset() As ADODB.Recordset

    Dim l_rs As ADODB.Recordset
    
    On Error GoTo eh
    
    Set l_rs = New ADODB.Recordset
    With l_rs
      .Fields.Append "EMP_NUM", adInteger, , adFldUpdatable
      .Fields.Append "CHARGE_DATE", adDate, , adFldUpdatable
      .Fields.Append "SHIFT_NUM", adInteger, , adFldUpdatable
      .Fields.Append "SCHED_CD", adInteger, , adFldUpdatable
      .Fields.Append "SCHED_SRC_CD", adInteger, , adFldUpdatable
      .Fields.Append "SHIFT_HOURS", adDecimal, , adFldUpdatable
      .Fields.Item("SHIFT_HOURS").Precision = 3
      .Fields.Item("SHIFT_HOURS").NumericScale = 1
      .Fields.Append "WG_NUM", adInteger, , adFldUpdatable
      .Fields.Append "LINE_NUM", adInteger, , adFldUpdatable
      .Fields.Append "SHIFT_START_TIME", adChar, 5, adFldUpdatable
      .Fields.Append "SHIFT_END_TIME", adChar, 5, adFldUpdatable
      .Fields.Append "SHIFT_START_DATE", adDate, , adFldUpdatable
      .Fields.Append "SHIFT_END_DATE", adDate, , adFldUpdatable
      .Fields.Append "SCHED_CD_FOR_SHIFT_COMP", adInteger, , adFldUpdatable
      .Fields.Append "BASE_SCHED_SRC_CD", adInteger, , adFldUpdatable
      .Fields.Append "SCHED_CD_MODIFIED", adInteger, , adFldUpdatable
'      .Fields.Append "ROLE_TYPE_CD", adChar, 1, adFldUpdatable
'      .Fields.Append "WEEK_TRADE_IND", adChar, 1, adFldUpdatable

      .CursorLocation = adUseClient
      .CursorType = adOpenStatic
      .LockType = adLockOptimistic
      .Open
    End With
    
    Set CreateNewSchedRecordset = l_rs
    
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "CreateNewSchedRecordset"


End Function



Private Sub CopySchedule( _
    ByVal v_lngEmpNum As Long, _
    ByVal v_rsSchedule As ADODB.Recordset, _
    ByRef r_rsNewSched As ADODB.Recordset)
    
    Dim l_datChargeDate As Date
    Dim l_lngShiftNum As Long
    Dim l_strFilter As String
    
    On Error GoTo eh
    
    v_rsSchedule.Sort = "CHARGE_DATE, SHIFT_NUM"
    v_rsSchedule.MoveFirst
    While Not v_rsSchedule.EOF
        l_datChargeDate = v_rsSchedule!CHARGE_DATE
        l_lngShiftNum = v_rsSchedule!SHIFT_NUM
        l_strFilter = "EMP_NUM = " & str(v_lngEmpNum) & _
            " AND CHARGE_DATE = " & str(l_datChargeDate) & _
            " AND SHIFT_NUM = " & str(l_lngShiftNum)
        l_strFilter = MDAC25FilterFix(l_strFilter)
        r_rsNewSched.Filter = l_strFilter
        If Not RecordsExist(r_rsNewSched) Then
            r_rsNewSched.AddNew
            r_rsNewSched!EMP_NUM = v_lngEmpNum
            r_rsNewSched!CHARGE_DATE = l_datChargeDate
            r_rsNewSched!SHIFT_NUM = l_lngShiftNum
            r_rsNewSched!SCHED_CD = v_rsSchedule!SCHED_CD
            r_rsNewSched!SCHED_SRC_CD = v_rsSchedule!SCHED_SRC_CD
            r_rsNewSched!SHIFT_HOURS = v_rsSchedule!SHIFT_HOURS
            r_rsNewSched!WG_NUM = v_rsSchedule!WG_NUM
            r_rsNewSched!LINE_NUM = v_rsSchedule!LINE_NUM
            r_rsNewSched!SHIFT_START_TIME = v_rsSchedule!SHIFT_START_TIME
            r_rsNewSched!SHIFT_END_TIME = v_rsSchedule!SHIFT_END_TIME
            r_rsNewSched!SHIFT_START_DATE = v_rsSchedule!SHIFT_START_DATE
            r_rsNewSched!SHIFT_END_DATE = v_rsSchedule!SHIFT_END_DATE
            r_rsNewSched!SCHED_CD_FOR_SHIFT_COMP = v_rsSchedule!SCHED_CD_FOR_SHIFT_COMP
            r_rsNewSched!BASE_SCHED_SRC_CD = v_rsSchedule!BASE_SCHED_SRC_CD
            
            r_rsNewSched.Update
        End If
        v_rsSchedule.MoveNext
    Wend
    r_rsNewSched.Filter = ""

    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "CopySchedule"

End Sub

Private Sub ApplyTradeToNewSched( _
    ByRef r_rsNewSched As ADODB.Recordset)
    
    Dim l_strFilter As String
    Dim l_lngEmpNum As Long
    Dim l_datChargeDate As Date
    Dim l_lngShiftNum As Long
    Dim l_blnEmpTrade As Boolean
    
    On Error GoTo eh
    
    m_rsTradeDetails.Filter = ""
    m_rsTradeDetails.Sort = "EMP_NUM,CHARGE_DATE,SHIFT_NUM"
    
    m_rsTradeDetails.MoveFirst
    While Not m_rsTradeDetails.EOF
        l_lngEmpNum = m_rsTradeDetails!EMP_NUM
        l_datChargeDate = m_rsTradeDetails!CHARGE_DATE
        l_lngShiftNum = m_rsTradeDetails!SHIFT_NUM
        l_strFilter = "EMP_NUM  = " & str(l_lngEmpNum) & _
            " AND CHARGE_DATE = " & str(l_datChargeDate) & _
            " AND SHIFT_NUM = " & str(l_lngShiftNum)
        l_strFilter = MDAC25FilterFix(l_strFilter)
        r_rsNewSched.Filter = l_strFilter
        l_blnEmpTrade = Not m_rsTradeDetails!LINE_NUM > 0
        If RecordsExist(r_rsNewSched) Then
            UpdateNewSched r_rsNewSched, m_rsTradeDetails, l_blnEmpTrade
        Else
            AddNewSched r_rsNewSched, m_rsTradeDetails, , , l_blnEmpTrade
        End If
        r_rsNewSched.Update
        m_rsTradeDetails.MoveNext
    Wend
    r_rsNewSched.Filter = ""
    
    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "ApplyTradeToNewSched"
    

End Sub

Private Sub UpdateNewSched( _
    ByRef r_rsNewSched As ADODB.Recordset, _
    ByVal v_rsTrade As ADODB.Recordset, _
    Optional ByVal v_blnEmpTrade As Boolean = False)
    
    On Error GoTo eh
    
'The -99.9 hours signifies that the shift hours
' of the schedule need to be retrieved
    r_rsNewSched!SCHED_CD_MODIFIED = 1
    r_rsNewSched!SHIFT_HOURS = -99.9
    r_rsNewSched!SCHED_CD = v_rsTrade!SCHED_CD
    If v_blnEmpTrade Then
        r_rsNewSched!SCHED_SRC_CD = trEmployeeTradesLevel
    Else
        r_rsNewSched!SCHED_SRC_CD = trDemandLineTrades
    End If
    r_rsNewSched.Update
    
    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "UpdateNewSched"

End Sub

Private Sub AddNewSched( _
    ByRef r_rsNewSched As ADODB.Recordset, _
    ByVal v_rsTrade As ADODB.Recordset, _
    Optional v_lngEmpNum As Long = 0, _
    Optional v_blnSetWGandLine As Boolean = True, _
    Optional v_blnEmpTrade As Boolean = False)

    On Error GoTo eh

'The -99.9 hours signifies that the shift hours
' of the schedule need to be retrieved
    r_rsNewSched.AddNew
    r_rsNewSched!SCHED_CD_MODIFIED = 1
    r_rsNewSched!SHIFT_HOURS = -99.9
    If v_lngEmpNum = 0 Then
        r_rsNewSched!EMP_NUM = v_rsTrade!EMP_NUM
    Else
        r_rsNewSched!EMP_NUM = v_lngEmpNum
    End If
    r_rsNewSched!CHARGE_DATE = v_rsTrade!CHARGE_DATE
    r_rsNewSched!SHIFT_NUM = v_rsTrade!SHIFT_NUM
    r_rsNewSched!SCHED_CD = v_rsTrade!SCHED_CD
    If v_blnEmpTrade Then
        r_rsNewSched!SCHED_SRC_CD = trEmployeeTradesLevel
    Else
        r_rsNewSched!SCHED_SRC_CD = trDemandLineTrades
    End If
    If v_blnSetWGandLine Then
        r_rsNewSched!WG_NUM = v_rsTrade!WG_NUM
        r_rsNewSched!LINE_NUM = v_rsTrade!LINE_NUM
    Else
        r_rsNewSched!WG_NUM = 0
        r_rsNewSched!LINE_NUM = 0
    End If
    r_rsNewSched.Update
    
    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "AddNewSched"


End Sub

Private Sub ApplyNewScheduleHours( _
    ByRef r_rsNewSched As ADODB.Recordset)
    
    Dim l_colSchedCodes As Collection
    Dim l_lngSchedCd As Long
    Dim l_rsAvailSchedules As ADODB.Recordset
    
    On Error GoTo eh
    
    r_rsNewSched.Filter = "SHIFT_HOURS = -99.99"
    If Not RecordsExist(r_rsNewSched) Then
        Exit Sub
    End If
    
    Set l_colSchedCodes = New Collection
    r_rsNewSched.MoveFirst
    While Not r_rsNewSched.EOF
        l_lngSchedCd = r_rsNewSched!SCHED_CD
        If IsInList(l_lngSchedCd, l_colSchedCodes) Then
        Else
            l_colSchedCodes.Add l_lngSchedCd
        End If
        r_rsNewSched.MoveNext
    Wend
    Set l_rsAvailSchedules = GetAvailScheduleDetails(l_colSchedCodes)
    r_rsNewSched.Sort = "CHARGE_DATE"
    r_rsNewSched.MoveFirst
    While Not r_rsNewSched.EOF
        AddScheduleDetails r_rsNewSched, l_rsAvailSchedules
        r_rsNewSched.MoveNext
    Wend
    
    
    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "ApplyNewScheduleHours"


End Sub

Private Function GetAvailScheduleDetails( _
    ByVal v_colSchedCds As Collection _
    ) As ADODB.Recordset
    
    Dim l_dbSchedule As dbSchedule.cdSchedule
    
    On Error GoTo eh
         
    
    Set l_dbSchedule = CreateObject("dbSchedule.cdSchedule")
    
    Set GetAvailScheduleDetails = l_dbSchedule.GetAvailSchedInfo(v_colSchedCds)
         
    Set l_dbSchedule = Nothing
    
    Exit Function
    
eh:
    Set l_dbSchedule = Nothing
    TrsRaiseError m_strCMODULENAME, "GetAvailScheduleDetails"

End Function


Private Sub AddScheduleDetails( _
    ByRef r_rsNewSched As ADODB.Recordset, _
    ByVal v_rsAvailSchedules As ADODB.Recordset)
    
    Dim l_lngSchedCode As Long
    Dim l_datChargeDate As Date
    Dim l_strShiftStartDay As String
    Dim l_strDayofWeek As String 'SUN,MON, ... SAT - Representation of Charge Date
    Dim l_strShiftStartTime As String
    Dim l_datShiftStartDate As Date
    Dim l_strShiftEndTime As String
    Dim l_datShiftEndDate As Date
    Dim l_decShiftHours As Variant
    
    
    On Error GoTo eh
    
    l_datChargeDate = r_rsNewSched!CHARGE_DATE
    l_lngSchedCode = r_rsNewSched!SCHED_CD
    v_rsAvailSchedules.Filter = "SCHED_CD = " & str(l_lngSchedCode)
    l_strDayofWeek = UCase(WeekdayName(Weekday(l_datChargeDate), True))   'returns day name SUN,MON,TUE, ..., SAT
    l_strShiftStartTime = Trim(v_rsAvailSchedules(l_strDayofWeek & "_START_TIME"))
    l_strShiftEndTime = Trim(v_rsAvailSchedules(l_strDayofWeek & "_END_TIME"))
    l_decShiftHours = v_rsAvailSchedules(l_strDayofWeek & "_STR_HRS")
    l_strShiftStartDay = v_rsAvailSchedules(l_strDayofWeek & "_ACT_DAY")
    DetermineShiftDates l_strShiftStartDay, l_strDayofWeek, l_datChargeDate, l_strShiftStartTime, l_strShiftEndTime, l_datShiftStartDate, l_datShiftEndDate
    
    r_rsNewSched!SHIFT_HOURS = l_decShiftHours
    r_rsNewSched!SHIFT_START_TIME = l_strShiftStartTime
    r_rsNewSched!SHIFT_END_TIME = l_strShiftEndTime
    r_rsNewSched!SHIFT_START_DATE = l_datShiftStartDate
    r_rsNewSched!SHIFT_END_DATE = l_datShiftEndDate
    r_rsNewSched.Update
    

    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "AddScheduleDetails"
    
    
End Sub

Private Sub DetermineShiftDates( _
                ByVal v_strShiftStartDay As String, _
                ByVal v_strChargeDateDayOfWeek As String, _
                ByVal v_datChargeDate As String, _
                ByVal v_strShiftStartTime As String, _
                ByVal v_strSiftEndTime As String, _
                ByRef r_datShiftStartDate As Date, _
                ByRef r_datShiftEndDate As Date)
                
    On Error GoTo eh
                
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
    
    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "DetermineShiftDates"
         
End Sub


Private Function PostValidateSaveTrade() As Boolean
    
' RSG 4/21/02 Changed scope of l_rsNewSchedules to module level, m_rsNewSchedules.
    Dim l_decRequestorHoursBeforeTrade As Variant
    Dim l_decAcceptorHoursBeforeTrade As Variant
    Dim l_lngRequestor As Long
    Dim l_lngAcceptor As Long
    Dim l_blnIsValid As Boolean
    

    On Error GoTo eh
    
    l_blnIsValid = True
    l_lngRequestor = m_colEmployees(1)
    Set m_rsNewSchedules = CreateNewSchedRecordset()
    CopySchedule l_lngRequestor, m_rsRequestorSchedule, m_rsNewSchedules
    l_decRequestorHoursBeforeTrade = CountScheduleHours(m_rsRequestorSchedule)
    If m_colEmployees.Count = 2 Then
        l_lngAcceptor = m_colEmployees(2)
        CopySchedule l_lngAcceptor, m_rsAcceptorSchedule, m_rsNewSchedules
        l_decAcceptorHoursBeforeTrade = CountScheduleHours(m_rsAcceptorSchedule)
    Else
        l_lngAcceptor = 0
    End If
    
    ApplyTradeToNewSched m_rsNewSchedules
    ApplyNewScheduleHours m_rsNewSchedules
    If Not PostValidateEmployeeScheduleValid( _
        l_lngRequestor, m_rsNewSchedules, l_decRequestorHoursBeforeTrade) Then
        l_blnIsValid = False
    End If
    If l_lngAcceptor > 0 Then
        If Not PostValidateEmployeeScheduleValid( _
            l_lngAcceptor, m_rsNewSchedules, l_decAcceptorHoursBeforeTrade) Then
            l_blnIsValid = False
        End If
    End If
    
    PostValidateEmployeeSchedChgPmt
    
    PostValidateSaveTrade = l_blnIsValid
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "PostValidateSaveTrade"

End Function


Private Function HasOverlappingShifts( _
    ByVal v_lngEmpNum As Long, _
    ByVal v_rsSchedule As ADODB.Recordset _
    ) As Boolean
    
    Dim l_datStartDate As Date
    Dim l_strStartTime As String
    Dim l_datEndDate As Date
    Dim l_strEndTime As String
    Dim l_datShiftStart1 As Date
    Dim l_datShiftStart2 As Date
    Dim l_datShiftEnd1 As Date
    Dim l_datShiftEnd2 As Date
    Dim l_blnShiftsOverlap As Boolean
    Dim l_strFilter As String
    Dim l_datChargeDate1 As Date
    Dim l_datChargeDate2 As Date
    Dim l_lngShiftNum1 As Long
    Dim l_lngShiftNum2 As Long
    

    
    On Error GoTo eh
    
    l_strFilter = "EMP_NUM = " & str(v_lngEmpNum) & _
        " AND SHIFT_HOURS > 0"
    l_strFilter = MDAC25FilterFix(l_strFilter)
    v_rsSchedule.Filter = l_strFilter
    If Not RecordsExist(v_rsSchedule) Then
        HasOverlappingShifts = False
        Exit Function
    End If
    
    l_blnShiftsOverlap = False
    v_rsSchedule.Sort = "SHIFT_START_DATE, SHIFT_START_TIME"
    v_rsSchedule.MoveFirst
    l_datStartDate = v_rsSchedule!SHIFT_START_DATE
    l_strStartTime = v_rsSchedule!SHIFT_START_TIME
    l_datEndDate = v_rsSchedule!SHIFT_END_DATE
    l_strEndTime = v_rsSchedule!SHIFT_END_TIME
    l_datShiftStart1 = DateTimeFrom(l_datStartDate, l_strStartTime)
    l_datShiftEnd1 = DateTimeFrom(l_datEndDate, l_strEndTime)
    l_datChargeDate1 = v_rsSchedule!CHARGE_DATE
    l_lngShiftNum1 = v_rsSchedule!SHIFT_NUM
    v_rsSchedule.MoveNext
    While Not v_rsSchedule.EOF
        l_datStartDate = v_rsSchedule!SHIFT_START_DATE
        l_strStartTime = v_rsSchedule!SHIFT_START_TIME
        l_datEndDate = v_rsSchedule!SHIFT_END_DATE
        l_strEndTime = v_rsSchedule!SHIFT_END_TIME
        l_datShiftStart2 = DateTimeFrom(l_datStartDate, l_strStartTime)
        l_datShiftEnd2 = DateTimeFrom(l_datEndDate, l_strEndTime)
        l_datChargeDate2 = v_rsSchedule!CHARGE_DATE
        l_lngShiftNum2 = v_rsSchedule!SHIFT_NUM
        If TimePeriodsIntersect(l_datShiftStart1, l_datShiftEnd1, l_datShiftStart2, l_datShiftEnd2) Then
            l_blnShiftsOverlap = True
            AppendOverrlapBRV trErrbusValidationError, _
            mstrCBUSINESSERR, _
            "Employee (" & str(v_lngEmpNum) & _
            ") has an overlapping shift.  " & _
            "The shift from " & Format(l_datShiftStart1, "MM/DD/YYYY HH:MM") & _
            " to " & Format(l_datShiftEnd1, "MM/DD/YYYY HH:MM") & _
            " overlaps shift from " & Format(l_datShiftStart2, "MM/DD/YYYY HH:MM") & _
            " to " & Format(l_datShiftEnd2, "MM/DD/YYYY HH:MM"), _
            v_lngEmpNum, l_datChargeDate1, l_datChargeDate2, l_lngShiftNum1, l_lngShiftNum2
        End If
        l_datShiftStart1 = l_datShiftStart2
        l_datShiftEnd1 = l_datShiftEnd2
        l_datChargeDate1 = l_datChargeDate2
        l_lngShiftNum1 = l_lngShiftNum2
        v_rsSchedule.MoveNext
    Wend
    
    HasOverlappingShifts = l_blnShiftsOverlap
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "HasOverlappingShifts"
End Function


Private Function PostValidateEmployeeScheduleValid( _
    ByVal v_lngEmpNum As Long, _
    ByVal v_rsSchedule As ADODB.Recordset, _
    ByVal v_decHoursBeforeTrade As Variant _
    ) As Boolean
    
    Dim l_decHoursAfterTrade As Variant
    Dim l_intNumOfDoubleShifts As Integer
    Dim l_blnIsValid As Boolean
    
    
    On Error GoTo eh
    
    
    l_blnIsValid = True
'Count Hours
    v_rsSchedule.Filter = "EMP_NUM = " & str(v_lngEmpNum)
    l_decHoursAfterTrade = CountScheduleHours(v_rsSchedule, True)
    If v_decHoursBeforeTrade <> l_decHoursAfterTrade Then
        AppendTradeBRV 20000, _
            mstrCBUSINESSWAR, _
            "The employee's (" & str(v_lngEmpNum) & _
            ") hours scheduled before the trade (" & _
            str(v_decHoursBeforeTrade) & _
            ") do not equal the hours after the trade (" & _
            str(l_decHoursAfterTrade) & ").", _
            v_lngEmpNum
    End If
'Count Doubles
    l_intNumOfDoubleShifts = CountNumberOfDoubleShifts(v_rsSchedule)
    If l_intNumOfDoubleShifts > 1 Then
        AppendTradeBRV 20000, _
            mstrCBUSINESSWAR, _
            "The employee (" & str(v_lngEmpNum) & _
            ") has " & str(l_intNumOfDoubleShifts) & _
            " double shifts in the week of " & _
            Format(m_datStartDate, "mm/dd/yyyy"), _
            v_lngEmpNum
    End If
'Check for overlapping shifts
    If HasOverlappingShifts(v_lngEmpNum, v_rsSchedule) Then
        l_blnIsValid = False
    End If

    PostValidateEmployeeScheduleValid = l_blnIsValid
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "PostValidateEmployeeScheduleValid"
    
    
End Function


Private Function PostValidateDeleteTrade() As Boolean
    Dim l_blnIsValid As Boolean
    Dim l_rsCurrentSchedule As ADODB.Recordset
    Dim colElement As Variant
    Dim l_aryEmpHoursBeforeTrade() As Variant '//-----EmpNum and Hours before trade-----//
    Dim i As Long
    
    On Error GoTo eh
    
    'Did not have time to test code
'Come back to this after the user meetings
    PostValidateDeleteTrade = True
'    Exit Function
    
    Set m_rsSaveTradeRequest = Nothing
    Set l_rsCurrentSchedule = GetGABSchedule()
    l_blnIsValid = True
    
    Set m_rsNewSchedules = CreateNewSchedRecordset()
       
    '//-----(i, 0) = EmpNum, (i, 1) = Hours Before Trade
    i = 0
    ReDim l_aryEmpHoursBeforeTrade(i To m_colEmployees.Count - 1, 0 To 1)
    
    For Each colElement In m_colEmployees
        l_aryEmpHoursBeforeTrade(i, 0) = colElement '//-----EmpNum-----//
        l_rsCurrentSchedule.Filter = "EMP_NUM = " & str(l_aryEmpHoursBeforeTrade(i, 0))
        CopySchedule l_aryEmpHoursBeforeTrade(i, 0), l_rsCurrentSchedule, m_rsNewSchedules
        l_aryEmpHoursBeforeTrade(i, 1) = CountScheduleHours(l_rsCurrentSchedule)
        i = i + 1
    Next
    
    PV_Remove_Trades m_rsNewSchedules
    PV_Reapply_EMP_Trades m_rsNewSchedules
    PV_Reapply_Line_Trades m_rsNewSchedules
    
    ApplyNewScheduleHours m_rsNewSchedules
    
    For i = LBound(l_aryEmpHoursBeforeTrade, 1) To UBound(l_aryEmpHoursBeforeTrade, 1)
        If Not PostValidateEmployeeScheduleValid( _
            l_aryEmpHoursBeforeTrade(i, 0), m_rsNewSchedules, l_aryEmpHoursBeforeTrade(i, 1)) Then
            l_blnIsValid = False
        End If
    Next i
        
    PostValidateEmployeeSchedChgPmt
        
    PostValidateDeleteTrade = l_blnIsValid
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "PostValidateDeleteTrade"
End Function


Private Function PostValidateEmployeeSchedChgPmt() As Boolean
      
    Dim l_blnIsValid As Boolean
    Dim l_dbSchedChgPay As dbSchedChgPay.cdSchedChgPay
    Dim l_rsSchedChgPay As ADODB.Recordset
    Dim l_aryEmpNums As Variant
        
On Error GoTo eh
    
    l_blnIsValid = True
    Set l_dbSchedChgPay = CreateObject("dbSchedChgPay.cdSchedChgPay")
    l_aryEmpNums = BuildEmployeeArray
    Set l_rsSchedChgPay = l_dbSchedChgPay.GetSchedChgPay(l_aryEmpNums, m_datStartDate, m_datEndDate)
    Set l_dbSchedChgPay = Nothing
    
    If Not RecordsExist(l_rsSchedChgPay) Then
        PostValidateEmployeeSchedChgPmt = True
        Exit Function
    End If

    Do Until l_rsSchedChgPay.EOF
        If IsPaymentScheduledOnOffDay(l_rsSchedChgPay, m_rsNewSchedules) Then
            AppendTradeBRV 20000, _
                mstrCBUSINESSWAR, _
                "Employee " & CStr(l_rsSchedChgPay!EMP_NUM) & _
                    "[" & Trim(l_rsSchedChgPay!EMP_NAME) & "]" & _
                    " has a schedule change payment scheduled for an off-day, " & _
                    Format(l_rsSchedChgPay!CHARGE_DATE, "mm/dd/yyyy"), _
                l_rsSchedChgPay!EMP_NUM
        End If
        l_rsSchedChgPay.MoveNext
    Loop

    PostValidateEmployeeSchedChgPmt = l_blnIsValid
    Set l_rsSchedChgPay = Nothing
    Exit Function
    
eh:
    Set l_dbSchedChgPay = Nothing
    Set l_rsSchedChgPay = Nothing
    TrsRaiseError m_strCMODULENAME, "PostValidateEmployeeSchedChgPmt"
End Function


Private Function BuildEmployeeArray() As Variant
    Dim l_varEmpNum As Variant
    Dim l_intEmpCount As Integer
    Dim l_aryReturn() As Variant
    Dim i As Integer
    
On Error GoTo eh

' Count non-zero employee numbers
    For Each l_varEmpNum In m_colEmployees
        If CLng(l_varEmpNum) <> 0 Then
            l_intEmpCount = l_intEmpCount + 1
        End If
    Next
    
    If l_intEmpCount = 0 Then
        BuildEmployeeArray = Empty
        Exit Function
    End If
    
    ReDim l_aryReturn(0 To l_intEmpCount - 1)

    For Each l_varEmpNum In m_colEmployees
        If CLng(l_varEmpNum) <> 0 Then
            l_aryReturn(i) = CLng(l_varEmpNum)
            i = i + 1
        End If
    Next

    BuildEmployeeArray = l_aryReturn
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "BuildEmployeeArray"
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


Private Function BuildSchedChgPayScope() As ADODB.Recordset
    Dim l_rsScope As ADODB.Recordset
    Dim l_varEmpNum As Variant

On Error GoTo eh

    Set l_rsScope = New ADODB.Recordset

    l_rsScope.Fields.Append "EMP_NUM", adInteger, , adFldUpdatable
    l_rsScope.Fields.Append "START_DATE", adDate, , adFldUpdatable
    l_rsScope.Fields.Append "END_DATE", adDate, , adFldUpdatable
    l_rsScope.Open

    For Each l_varEmpNum In m_colEmployees
        If CLng(l_varEmpNum) <> 0 Then
            l_rsScope.AddNew
            l_rsScope!EMP_NUM = CLng(l_varEmpNum)
            l_rsScope!START_DATE = m_datStartDate
            l_rsScope!END_DATE = m_datEndDate
            l_rsScope.Update
        End If
    Next

    Set BuildSchedChgPayScope = l_rsScope
    Set l_rsScope = Nothing
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "BuildSchedChgPayScope"
End Function


Private Function GetGABSchedule() As ADODB.Recordset
    Dim l_obj As busSchedule.cbsSchedule
    Dim l_rs As ADODB.Recordset
    Dim l_varEmployees As Variant
    Dim l_rsDummy As ADODB.Recordset

    On Error GoTo eh
    
    l_varEmployees = ArrayFromCollection(m_colEmployees)

    Set l_obj = CreateObject("busSchedule.cbsSchedule")
    
    Set l_rs = l_obj.GetAdjustedBaseSchedule(l_varEmployees, m_datStartDate, m_datEndDate, l_rsDummy)
         
    Set l_obj = Nothing
    
    Set GetGABSchedule = l_rs
    
    Exit Function
    
eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "GetGABSchedule"
    
End Function


Private Sub PV_Remove_Trades( _
    ByRef r_rsNewSched As ADODB.Recordset)
    
    On Error GoTo eh
    
    r_rsNewSched.Filter = ""
    r_rsNewSched.MoveFirst
    While Not r_rsNewSched.EOF
        r_rsNewSched!SCHED_SRC_CD = r_rsNewSched!BASE_SCHED_SRC_CD
        If r_rsNewSched!SCHED_CD <> _
            r_rsNewSched!SCHED_CD_FOR_SHIFT_COMP Then
            r_rsNewSched!SCHED_CD = r_rsNewSched!SCHED_CD_FOR_SHIFT_COMP
            r_rsNewSched!SHIFT_HOURS = -99.9
        End If
        If r_rsNewSched!SHIFT_NUM = 2 Then
            r_rsNewSched!SCHED_CD = -2
            r_rsNewSched!SCHED_SRC_CD = 0
            r_rsNewSched!SHIFT_HOURS = -99.9
        End If
        r_rsNewSched.Update
        r_rsNewSched.MoveNext
    Wend

    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "PV_Remove_Trades"



End Sub

Private Sub PV_Reapply_Line_Trades( _
    ByRef r_rsNewSched As ADODB.Recordset)
    
    Dim l_strFilter As String
    Dim l_lngEmpNum As Long
    Dim l_datChargeDate As Date
    Dim l_intShiftNum As Integer
    Dim l_lngWGNum As Long
    Dim l_lngLineNum As Long
    
    
    On Error GoTo eh
    
        m_rsLineTrade.Filter = ""
        If Not RecordsExist(m_rsLineTrade) Then
            Exit Sub
        End If
        m_rsLineTrade.MoveFirst
        While Not m_rsLineTrade.EOF
            l_datChargeDate = m_rsLineTrade!CHARGE_DATE
            l_intShiftNum = m_rsLineTrade!SHIFT_NUM
            l_lngWGNum = m_rsLineTrade!WG_NUM
            l_lngLineNum = m_rsLineTrade!LINE_NUM
            l_strFilter = "CHARGE_DATE = " & str(l_datChargeDate) & _
                " AND WG_NUM = " & str(l_lngWGNum) & _
                " AND LINE_NUM = " & str(l_lngLineNum)
            l_strFilter = MDAC25FilterFix(l_strFilter)
            r_rsNewSched.Filter = l_strFilter
            If RecordsExist(r_rsNewSched) Then
    'employee is assigned to a line
                l_lngEmpNum = r_rsNewSched!EMP_NUM
                l_strFilter = "CHARGE_DATE = " & str(l_datChargeDate) & _
                    " AND SHIFT_NUM = " & str(l_intShiftNum) & _
                    " AND WG_NUM = " & str(l_lngWGNum) & _
                    " AND LINE_NUM = " & str(l_lngLineNum)
                l_strFilter = MDAC25FilterFix(l_strFilter)
                r_rsNewSched.Filter = l_strFilter
                If RecordsExist(r_rsNewSched) Then
                    UpdateNewSched r_rsNewSched, m_rsLineTrade
                Else
                    AddNewSched r_rsNewSched, m_rsLineTrade, l_lngEmpNum
                End If
            End If
            m_rsLineTrade.MoveNext
        Wend
        
        
    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "PV_Reapply_Line_Trades"
    
End Sub

Private Sub PV_Reapply_EMP_Trades( _
    ByRef r_rsNewSched As ADODB.Recordset)
    
    Dim l_strFilter As String
    Dim l_lngEmpNum As Long
    Dim l_datChargeDate As Date
    Dim l_intShiftNum As Integer
    Dim l_lngWGNum As Long
    Dim l_lngLineNum As Long
    
    
    On Error GoTo eh
    
        m_rsEmpTrade.Filter = ""
        If Not RecordsExist(m_rsEmpTrade) Then
            Exit Sub
        End If
        m_rsEmpTrade.MoveFirst
        While Not m_rsEmpTrade.EOF
            l_lngEmpNum = m_rsEmpTrade!EMP_NUM
            l_datChargeDate = m_rsEmpTrade!CHARGE_DATE
            l_intShiftNum = m_rsEmpTrade!SHIFT_NUM
            l_strFilter = "CHARGE_DATE = " & str(l_datChargeDate) & _
                " AND EMP_NUM = " & str(l_lngEmpNum) & _
                " AND SHIFT_NUM = " & str(l_intShiftNum)
            l_strFilter = MDAC25FilterFix(l_strFilter)
            r_rsNewSched.Filter = l_strFilter
            If RecordsExist(r_rsNewSched) Then
                UpdateNewSched r_rsNewSched, m_rsEmpTrade, True
            Else
                AddNewSched r_rsNewSched, m_rsEmpTrade, l_lngEmpNum, False, True
            End If
            m_rsEmpTrade.MoveNext
        Wend
        
        
    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "PV_Reapply_EMP_Trades"
    
End Sub


Private Sub ClearTrade()


    On Error GoTo eh
    
    m_rsOtherTradeDetails.Filter = "TRADE_NUM = " & str(m_lngTradeNum)
    
    If Not RecordsExist(m_rsOtherTradeDetails) Then
        Exit Sub
    End If
    
    m_rsOtherTradeDetails.MoveFirst
    While Not m_rsOtherTradeDetails.EOF
        m_rsOtherTradeDetails.Delete
        m_rsOtherTradeDetails.MoveNext
    Wend

    m_rsOtherTradeDetails.Filter = ""
    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "ClearTrade"
End Sub

Private Function AppendOverrlapBRV( _
            ByVal v_lngErrorNumber As Long, _
            ByVal v_strErrorType As String, _
            ByVal v_strErrorDesc As String, _
            ByVal v_lngEmpNum As Long, _
            ByVal v_datChargeDate1 As Date, _
            ByVal v_datChargeDate2 As Date, _
            ByVal v_lngShiftNum1 As Long, _
            ByVal v_lngShiftNum2 As Long _
            ) As Boolean
            
    Dim l_bolError As Boolean
    Dim l_strChargeDate As String
    Dim l_strShiftNum As String
    
    On Error GoTo AppendOverrlapBRV_EH
    
    
    If RecordsExist(m_rsSaveTradeRequest) Then
        If Not SaveRequestFound(l_strChargeDate, l_strShiftNum, _
                v_datChargeDate1, v_datChargeDate2, _
                v_lngShiftNum1, v_lngShiftNum2, v_lngEmpNum) Then
            l_strChargeDate = ""
            l_strShiftNum = ""
        End If
    Else
        l_strChargeDate = str(v_datChargeDate1)
        l_strShiftNum = str(v_lngShiftNum1)
    End If
   
    l_bolError = AppendBusinessError(m_rsErrors, _
            v_lngErrorNumber, _
            v_strErrorType, _
            v_strErrorDesc, _
            m_strCMODULENAME, _
            TrType_TR, , , "EMP_NUM", _
            str(v_lngEmpNum), l_strChargeDate, l_strShiftNum, , , str(m_lngTradeNum))

    AppendOverrlapBRV = l_bolError
    Exit Function
    
AppendOverrlapBRV_EH:
    TrsRaiseError m_strCMODULENAME, "AppendOverrlapBRV"

End Function


Private Function SaveRequestFound( _
            ByRef r_strChargeDate As String, _
            ByRef r_strShiftNum As String, _
            ByVal v_datChargeDate1 As Date, _
            ByVal v_datChargeDate2 As Date, _
            ByVal v_lngShiftNum1 As Long, _
            ByVal v_lngShiftNum2 As Long, _
            ByVal v_lngEmpNum As Long _
            ) As Boolean
            
            
    Dim l_strFilter As String
    
    On Error GoTo eh
    
    l_strFilter = "EMP_NUM = " & str(v_lngEmpNum) & _
        " AND CHARGE_DATE = " & str(v_datChargeDate1) & _
        " AND SHIFT_NUM = " & str(v_lngShiftNum1)
    l_strFilter = MDAC25FilterFix(l_strFilter)
    m_rsSaveTradeRequest.Filter = l_strFilter
    If RecordsExist(m_rsSaveTradeRequest) Then
        r_strChargeDate = str(v_datChargeDate1)
        r_strShiftNum = str(v_lngShiftNum1)
        SaveRequestFound = True
        Exit Function
    End If
    
    l_strFilter = "EMP_NUM = " & str(v_lngEmpNum) & _
        " AND CHARGE_DATE = " & str(v_datChargeDate2) & _
        " AND SHIFT_NUM = " & str(v_lngShiftNum2)
    l_strFilter = MDAC25FilterFix(l_strFilter)
    m_rsSaveTradeRequest.Filter = l_strFilter
    If RecordsExist(m_rsSaveTradeRequest) Then
        r_strChargeDate = str(v_datChargeDate2)
        r_strShiftNum = str(v_lngShiftNum2)
        SaveRequestFound = True
    Else
        SaveRequestFound = False
    End If
    
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "SaveRequestFound"

End Function

Private Sub WriteTradeNumberBRVMessage()

    Dim l_bolError As Boolean
    Dim l_strErrorDesc As String

    On Error GoTo eh

    l_strErrorDesc = "Trade number " & _
        Trim(str(m_lngTradeNum)) & _
        " was saved successfully."
   
    l_bolError = AppendBusinessError(m_rsErrors, _
            10000, _
            mstrCBUSINESSINF, _
            l_strErrorDesc, _
            m_strCMODULENAME, _
            TrType_TR, , , "TRADE_NUM", _
            , , , , , str(m_lngTradeNum))
    
    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "WriteTradeNumberBRVMessage"

End Sub

