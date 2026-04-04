VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "cbsSchedDemandTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'COMMENTS: OCT, 2005 - CHANGE GETOBJECTCONTEXT.CREATEINSTANCE
'TO CREATEOBJECT -- PREM

Private Const m_strCMODULENAME As String = "cbsSchedDemandTrans"
Private Const m_datCOPTIONALSTARTDATE As Date = #1/1/100#
Private Const m_datCOPTIONALENDDATE As Date = #12/31/9999#
'//-----define the following so that ApplyPendingChanges only get applied to overlap in scope once-----//
Private Const m_strCPENDINGDELETE As String = "1"
Private Const m_strCPENDINGINSERT As String = "2"
Private Const m_strCPENDINGCHANGE As String = "3"
' 5/17/02 RSG Moved to mTRISConstants.bas
'Public Type tScopeParm
'    lngMinStartDate As Date
'    lngMaxEndDate As Date
'    lngWgNum As Long
'    lngLineNum As Long
'    lngEmpNum As Long
'End Type

Private Enum enumReliefType
    rtUseUnion = 1
    rtUseUnionAndPriorUnion = 2
    rtUsePriorUnion = 3
    rtErrorNothingBeforePriorUnion = 4
    rtErrorNoPriorUnion = 5
    rtErrorUnknown = 6
End Enum

' CASCADE processing
' Capture the date intervals that an employees schedule may have changed.
' First, capture the employee on the request in ProcessKeyChange.  Second,
' capture the employees with overlapping lines in AdjustOverlapByLine.
' Unassign requests (EMP_NUM = 0) will be captured in AdjustOverlapByLine also.
Private Const mblnAPPLYCASCADE As Boolean = True
Private Const mstrCVARFIELDNAMES As String = "WG_NUM;LINE_NUM;START_DATE;END_DATE;EMP_NUM"
Private m_colCascadeIntervals As Collection

'//-----for testing performance-----//
'if #debug then
'    Private Declare Function GetTickCount Lib "kernel32" () As Long
'End If

Public Function SaveAssignments( _
                    ByVal v_rsAssignmentRequest As ADODB.Recordset, _
                    ByRef r_rsErrors As ADODB.Recordset, _
                    ByVal v_lTRISUser As Long, _
                    ByRef r_rsSchedDemandUpdates As ADODB.Recordset, _
                    ByRef r_rsCascadeConflicts As ADODB.Recordset, _
                    ByRef r_rsExceptionSPA As ADODB.Recordset, _
                    ByRef r_rsAvailScheduleSPA As ADODB.Recordset, _
                    ByRef r_rsEmpSchedDates As ADODB.Recordset, _
                    ByVal v_blnConfirmConflicts As Boolean, _
                    Optional ByVal v_blnBypassCascade As Boolean = False _
                ) As Boolean
                
    Dim l_objDBSchDmdTrans As dbSchedDemand.cdSchedDemandTrans
    Dim l_rsErrors As ADODB.Recordset
    Dim bError As Boolean

    Dim l_blnConfirmPending As Boolean
    Dim l_varCascadeUpdates As Variant

    On Error GoTo eh
    bError = False

    Set l_rsErrors = CreateNewBusinessErrorSet
'//-----3/1/2002 test for method ~ of object ~ failed
    If PreValidateAssignments(v_rsAssignmentRequest, l_rsErrors, v_lTRISUser) Then
        '//Modify RECORD_STATUS = "C" requests, to "D" and "I" for requests that are changing the key columns (10/23/2001)
'//-----3/1/2002 test for method ~ of object ~ failed
        If Not ProcessKeyChange(v_rsAssignmentRequest) Then
            bError = True
        End If
            
'//-----3/1/2002 test for method ~ of object ~ failed
        If ValidateAssignments(v_rsAssignmentRequest, l_rsErrors) Then
'//-----3/1/2002 test for method ~ of object ~ failed
            If PostValidateAssignments(v_rsAssignmentRequest, l_rsErrors) Then
            Else
                bError = True
            End If
        Else
            bError = True
        End If
'//-----3/1/2002 test for method ~ of object ~ failed
    Else
        bError = True
    End If
    
'    TRIS 4 data conversion will pass -1 in v_lTRISUser
'    If Not bError And v_lTRISUser <> -1 Then
     If Not bError And Not v_blnBypassCascade Then
        Set r_rsSchedDemandUpdates = v_rsAssignmentRequest
        If ProcessCascade(v_lTRISUser, _
                    v_blnConfirmConflicts, l_blnConfirmPending, _
                    l_varCascadeUpdates, r_rsSchedDemandUpdates, _
                    r_rsCascadeConflicts, r_rsExceptionSPA, r_rsAvailScheduleSPA, _
                    r_rsEmpSchedDates, l_rsErrors) Then
            If l_blnConfirmPending Then
                SaveAssignments = True
                GetObjectContext.SetComplete
                GoTo xit
            End If
        Else
            bError = True
        End If
'    Avoid marshalling these recordsets when no confirmation is pending
        Set r_rsSchedDemandUpdates = Nothing
        Set r_rsExceptionSPA = Nothing
        Set r_rsAvailScheduleSPA = Nothing
        Set r_rsCascadeConflicts = Nothing
    End If
    
    If Not bError Then
        If Not ProcessScheduleChange(r_rsEmpSchedDates, l_rsErrors, v_rsAssignmentRequest) Then
            bError = True
        End If
    End If
    Set r_rsEmpSchedDates = Nothing
    
'Method HighManHours not needed in TRIS 5
'''''    If Not bError Then
'''''        If Not HighManHours(v_rsAssignmentRequest, v_lTRISUser) Then
'''''            bError = True
'''''        End If
'''''    End If
    
    If Not bError Then
        If RecordsExist(v_rsAssignmentRequest) Then
            Set l_objDBSchDmdTrans = CreateObject("dbSchedDemand.cdSchedDemandTrans")
            bError = Not l_objDBSchDmdTrans.SaveAssignments(v_rsAssignmentRequest, v_lTRISUser)
            Set l_objDBSchDmdTrans = Nothing
        End If
    End If

    If bError Then
        SaveAssignments = False
        GetObjectContext.SetAbort
    Else
        SaveAssignments = True
        GetObjectContext.SetComplete
    End If


xit:
    Set r_rsErrors = l_rsErrors '//-----pass back to client
'//-----3/1/2002 test for method ~ of object ~ failed
    Set l_rsErrors = Nothing
    Set m_colCascadeIntervals = Nothing
    Exit Function
    
eh:
    Set l_objDBSchDmdTrans = Nothing
    Set l_rsErrors = Nothing
    Set m_colCascadeIntervals = Nothing
    Set r_rsSchedDemandUpdates = Nothing
    Set r_rsExceptionSPA = Nothing
    Set r_rsAvailScheduleSPA = Nothing
    Set r_rsCascadeConflicts = Nothing
    Set r_rsEmpSchedDates = Nothing
    
    SaveAssignments = False
    GetObjectContext.SetAbort
    Select Case Err.Number
        Case trErrbusSchedDemandDuplicateLineAsgnm
            Err.Clear
            With v_rsAssignmentRequest
                AppendBusinessError r_rsErrors, trErrbusValidationError, mstrCBUSINESSERR, _
                    "Cannot perform " & IIf(!RECORD_STATUS = "I", "insert", "update") & _
                    " because an assignment already exists.", m_strCMODULENAME, _
                    1, 1, "LINE_NUM", _
                    Array("WG_NUM", "LINE_NUM", "START_DATE", "END_DATE", "EMP_NUM"), _
                    CStr(!WG_NUM), _
                    CStr(!LINE_NUM), _
                    Format(!START_DATE, "mm/dd/yyyy"), _
                    Format(!END_DATE, "mm/dd/yyyy"), _
                    CStr(!EMP_NUM)
            End With
            
            Exit Function
'    Clear the cascade error arguments since the client does not format them.
        Case trErrAppCascadeCommitConflict
            Err.Source = Join(Array("0", "", "", "", ""), ";")
            TrsRaiseError m_strCMODULENAME, "SaveAssignments", , , , , True
        Case Else
            TrsRaiseError m_strCMODULENAME, "SaveAssignments", , , , , True
    End Select

End Function


Public Function ResumeSaveAssignments( _
                    ByVal v_rsSchedDemandUpdates As ADODB.Recordset, _
                    ByVal v_rsCascadeConflicts As ADODB.Recordset, _
                    ByVal v_rsExceptionSPA As ADODB.Recordset, _
                    ByVal v_rsAvailScheduleSPA As ADODB.Recordset, _
                    ByVal v_rsEmpSchedDates As ADODB.Recordset, _
                    ByRef r_rsErrors As ADODB.Recordset, _
                    ByVal v_lTRISUser As Long) _
                As Boolean
' Do not expect any lock contention between cascade and demand schedule updates, so
' request cascade to apply the updates instead of returning them, mblnAPPLYCASCADE.
    Dim l_objCtx As ObjectContext
    Dim l_rsErrors As ADODB.Recordset
    Dim bError As Boolean
    Dim l_cbsCascade As busCascade.cbsCascade
    Dim l_objDBSchDmdTrans As dbSchedDemand.cdSchedDemandTrans
    Dim l_varCascadeUpdates As Variant
    Dim l_rsCascadeErrors As ADODB.Recordset

On Error GoTo eh
    Set l_objCtx = GetObjectContext
    If r_rsErrors Is Nothing Then
        Set r_rsErrors = CreateNewBusinessErrorSet
    End If
       
    If Not ProcessScheduleChange(v_rsEmpSchedDates, r_rsErrors, v_rsSchedDemandUpdates) Then
        l_objCtx.SetAbort
        ResumeSaveAssignments = False
        GoTo xit
    End If
       
    Set l_cbsCascade = CreateObject("busCascade.cbsCascade")
    l_cbsCascade.UpdateExceptions _
        v_rsCascadeConflicts, v_rsExceptionSPA, v_rsAvailScheduleSPA, _
        mblnAPPLYCASCADE, l_varCascadeUpdates, l_rsCascadeErrors, v_lTRISUser
    Set l_cbsCascade = Nothing
    
    AppendCascadeErrors l_rsCascadeErrors, r_rsErrors
    Set l_rsCascadeErrors = Nothing
    
'Method HighManHours not needed in TRIS 5
'''''    If Not HighManHours(v_rsSchedDemandUpdates, v_lTRISUser) Then
'''''        ResumeSaveAssignments = False
'''''        l_objCtx.SetAbort
'''''        GoTo xit
'''''    End If
    
    Set l_objDBSchDmdTrans = CreateObject("dbSchedDemand.cdSchedDemandTrans")
    bError = Not l_objDBSchDmdTrans.SaveAssignments(v_rsSchedDemandUpdates, v_lTRISUser)
    Set l_objDBSchDmdTrans = Nothing
    
    If bError Then
        ResumeSaveAssignments = False
        l_objCtx.SetAbort
    Else
        ResumeSaveAssignments = True
        l_objCtx.SetComplete
    End If

xit:
    Set l_objCtx = Nothing
    Exit Function
    
eh:
    l_objCtx.SetAbort
    Set l_objCtx = Nothing
    Set l_rsCascadeErrors = Nothing
    Set l_cbsCascade = Nothing
    Set l_objDBSchDmdTrans = Nothing
    Set r_rsErrors = Nothing
'    Clear the cascade error arguments since the client does not format them.
    If Err.Number = trErrAppCascadeCommitConflict Then
        Err.Source = Join(Array("0", "", "", "", ""), ";")
    End If
    TrsRaiseError m_strCMODULENAME, "ResumeSaveAssignments", , , , , True
End Function


Private Function PreValidateAssignments(ByVal v_rsAssignmentRequest As ADODB.Recordset, _
                                        ByRef r_rsErrors As ADODB.Recordset, _
                                        ByVal v_lTRISUser As Long) As Boolean
    Dim bError As Boolean
    Dim l_objDBWorkGroup As dbWorkGroup.cdWorkGroup
    Dim xx As Long, l_strDatesMissing As String, l_strDatesWithRelief As String
    Dim bUseBulk As Boolean, bValid As Boolean
    Dim l_rsLinesBulk As ADODB.Recordset
    Dim l_strBRVMessage As String
    Dim l_datWGEffStart As Date
    Dim l_datStartDateAdj As Date
    Dim l_lngEmpNum As Long
    Dim l_strWarDesc As String
    
'//-----for testing performance-----//
'Dim lngStart As Long, lngStop As Long, lngTotalLines As Long, lngTotalRelief As Long
    
    On Error GoTo eh
    bError = False
    PreValidateAssignments = False

    If RecordsExist(v_rsAssignmentRequest) Then
        With v_rsAssignmentRequest

            '//-----mainly for conversion
            If v_rsAssignmentRequest.RecordCount >= 5 Then
                bUseBulk = True
            Else
                bUseBulk = False
            End If

            If bUseBulk Then
                Set l_rsLinesBulk = GetLinesBulk(v_rsAssignmentRequest)
            End If
            
            Set l_objDBWorkGroup = CreateObject("dbWorkGroup.cdWorkGroup")
            
            .MoveFirst
            Do While Not .EOF
'                '//-----to make it easier for setting filters and BRV's (use key fields for all-----//
'                '//-----otherwise must use key fields for "C", "D" and non-key fields for "I")-----//
                If !RECORD_STATUS = "I" Then
                    If IsNull(!KEY_WG_NUM) Then
                        !KEY_WG_NUM = !WG_NUM
                    End If
                    If IsNull(!KEY_EMP_NUM) Then
                        !KEY_EMP_NUM = !EMP_NUM
                    End If
                    If IsNull(!KEY_LINE_NUM) Then
                        !KEY_LINE_NUM = !LINE_NUM
                    End If
                    If IsNull(!KEY_START_DATE) Then
                        !KEY_START_DATE = !START_DATE
                    End If
                    If IsNull(!KEY_END_DATE) Then
                        !KEY_END_DATE = !END_DATE
                    End If
                    .Update
                End If
                If !RECORD_STATUS = "D" Then
                    If IsNull(!WG_NUM) Then
                        !WG_NUM = !KEY_WG_NUM
                    End If
                    If IsNull(!EMP_NUM) Then
                        !EMP_NUM = !KEY_EMP_NUM
                    End If
                    If IsNull(!LINE_NUM) Then
                        !LINE_NUM = !KEY_LINE_NUM
                    End If
                    If IsNull(!START_DATE) Then
                        !START_DATE = !KEY_START_DATE
                    End If
                    If IsNull(!END_DATE) Then
                        !END_DATE = !KEY_END_DATE
                    End If
                    .Update
                End If

                '//-----if the EndDate is not specified, assume it is the end of time-----//
                If IsNull(!END_DATE) Then
                    !END_DATE = m_datCOPTIONALENDDATE
                    .Update
                End If
                If IsNull(!KEY_END_DATE) Then
                    !KEY_END_DATE = m_datCOPTIONALENDDATE
                    .Update
                End If
                Select Case !RECORD_STATUS
                Case "I", "C"
                    If !EMP_NUM > 0 Then
                        l_lngEmpNum = !EMP_NUM
                        l_datStartDateAdj = !START_DATE
                        If IsStartDateBeforeEmployeeEmploymentDate(l_lngEmpNum, l_datStartDateAdj) Then
                            l_strWarDesc = "System changed the Line assigmement start date because it was before the employee's employment date. "
                            !START_DATE = l_datStartDateAdj
                            !KEY_START_DATE = l_datStartDateAdj
                                AppendBusinessError r_rsErrors, 20000, mstrCBUSINESSWAR, l_strWarDesc, m_strCMODULENAME, 1, 1, "START_DATE", _
                            Array("WG_NUM", "LINE_NUM", "START_DATE", "END_DATE", "EMP_NUM"), _
                            CStr(!WG_NUM), CStr(!LINE_NUM), Format(!START_DATE, "mm/dd/yyyy"), Format(!END_DATE, "mm/dd/yyyy"), CStr(!EMP_NUM)
                        End If
                    End If
                    If !START_DATE > !END_DATE Then
                        bError = True
                        AppendBusinessError r_rsErrors, trErrbusValidationError, mstrCBUSINESSERR, _
                            "Start Date (" & !START_DATE & ") must be less than or equal to End Date (" & !END_DATE & ")." _
                            , m_strCMODULENAME, _
                            1, 1, "START_DATE", _
                            Array("WG_NUM", "LINE_NUM", "START_DATE", "END_DATE", "EMP_NUM"), _
                            CStr(!WG_NUM), _
                            CStr(!LINE_NUM), _
                            Format(!START_DATE, "mm/dd/yyyy"), _
                            Format(!END_DATE, "mm/dd/yyyy"), _
                            CStr(!EMP_NUM)
                    End If
                    
                    '//-----Validate the work group number (for Relief) PTS 1032-----//
                    If !WG_NUM = 0 Then
                        bError = True
                        AppendBusinessError r_rsErrors, trErrbusValidationError, mstrCBUSINESSERR, _
                            "WorkGroup must be specified for Relief Assignment." _
                            , m_strCMODULENAME, _
                            1, 1, "WG_NUM", _
                            Array("WG_NUM", "LINE_NUM", "START_DATE", "END_DATE", "EMP_NUM"), _
                            CStr(!WG_NUM), _
                            CStr(!LINE_NUM), _
                            Format(!START_DATE, "mm/dd/yyyy"), _
                            Format(!END_DATE, "mm/dd/yyyy"), _
                            CStr(!EMP_NUM)
                    ElseIf Not l_objDBWorkGroup.IsValidWorkgroup(!WG_NUM, !START_DATE) Then
                        l_datWGEffStart = l_objDBWorkGroup.EffectiveStartDate(!WG_NUM)
                        If Not (IsDate(l_datWGEffStart) And l_datWGEffStart <> "12:00:00 AM") Then
                            bError = True
                            AppendBusinessError r_rsErrors, trErrbusValidationError, mstrCBUSINESSERR, _
                                "WorkGroup (" & !WG_NUM & ") is not valid." _
                                , m_strCMODULENAME, _
                                1, 1, "WG_NUM", _
                                Array("WG_NUM", "LINE_NUM", "START_DATE", "END_DATE", "EMP_NUM"), _
                                CStr(!WG_NUM), _
                                CStr(!LINE_NUM), _
                                Format(!START_DATE, "mm/dd/yyyy"), _
                                Format(!END_DATE, "mm/dd/yyyy"), _
                                CStr(!EMP_NUM)
                        ElseIf !START_DATE < l_datWGEffStart Then '//-----PTS 706
                            bError = True
                            AppendBusinessError r_rsErrors, trErrbusValidationError, mstrCBUSINESSERR, _
                                "Start date (" & !START_DATE & ") cannot be less than work groups effective start date (" & l_datWGEffStart & ")." _
                                , m_strCMODULENAME, _
                                1, 1, "START_DATE", _
                                Array("WG_NUM", "LINE_NUM", "START_DATE", "END_DATE", "EMP_NUM"), _
                                CStr(!WG_NUM), _
                                CStr(!LINE_NUM), _
                                Format(!START_DATE, "mm/dd/yyyy"), _
                                Format(!END_DATE, "mm/dd/yyyy"), _
                                CStr(!EMP_NUM)
                        Else
                            bError = True
                            AppendBusinessError r_rsErrors, trErrbusValidationError, mstrCBUSINESSERR, _
                                "WorkGroup (" & !WG_NUM & ") is not valid for start date (" & !START_DATE & ")." _
                                , m_strCMODULENAME, _
                                1, 1, "WG_NUM", _
                                Array("WG_NUM", "LINE_NUM", "START_DATE", "END_DATE", "EMP_NUM"), _
                                CStr(!WG_NUM), _
                                CStr(!LINE_NUM), _
                                Format(!START_DATE, "mm/dd/yyyy"), _
                                Format(!END_DATE, "mm/dd/yyyy"), _
                                CStr(!EMP_NUM)
                        End If
                    '//-----Make sure that a line exists (ttr_wg_line_sched) for each day in date range-----//
                    ElseIf !EMP_NUM > 0 Then
                        If bUseBulk Then
                            bValid = DoLinesExistBulk(l_rsLinesBulk, !WG_NUM, !LINE_NUM, !START_DATE, !END_DATE, l_strDatesMissing)
                        Else
                            bValid = DoLinesExist(!WG_NUM, !LINE_NUM, !START_DATE, !END_DATE, l_strDatesMissing)
                        End If
                        If Not bValid Then
                            bError = True
                            AppendBusinessError r_rsErrors, trErrbusValidationError, mstrCBUSINESSERR, _
                                "Line number " & !LINE_NUM & " must exist for " & l_strDatesMissing & "." _
                                , m_strCMODULENAME, _
                                1, 1, "LINE_NUM", _
                                Array("WG_NUM", "LINE_NUM", "START_DATE", "END_DATE", "EMP_NUM"), _
                                CStr(!WG_NUM), _
                                CStr(!LINE_NUM), _
                                Format(!START_DATE, "mm/dd/yyyy"), _
                                Format(!END_DATE, "mm/dd/yyyy"), _
                                CStr(!EMP_NUM)
                        End If
                    End If
                    
                    '//-----validate RELIEF_COMP from TTR_UNION_COMP_AMT using SPMCache in busCodeTables
                    If v_lTRISUser <> -1 Then '//-----skip validation for relief for conversion-----//
                        If Not bUseBulk Then
                            If Not (IsNull(!RELIEF_COMP) Or IsEmpty(!RELIEF_COMP) Or Trim(!RELIEF_COMP) = "") Then
                                If Not IsReliefComponentValid(!EMP_NUM, !START_DATE, !END_DATE, !RELIEF_COMP, l_strBRVMessage) Then
                                    bError = True
                                    AppendBusinessError r_rsErrors, trErrbusValidationError, mstrCBUSINESSERR, _
                                        l_strBRVMessage, m_strCMODULENAME, _
                                        1, 1, "LINE_NUM", _
                                        Array("WG_NUM", "LINE_NUM", "START_DATE", "END_DATE", "EMP_NUM"), _
                                        CStr(!WG_NUM), _
                                        CStr(!LINE_NUM), _
                                        Format(!START_DATE, "mm/dd/yyyy"), _
                                        Format(!END_DATE, "mm/dd/yyyy"), _
                                        CStr(!EMP_NUM)
                                End If
                            End If
                        End If
                    End If
                Case "D"
                
                End Select
                .MoveNext
            Loop
            Set l_objDBWorkGroup = Nothing
        End With
    End If
       
    If v_lTRISUser <> -1 Then '//-----skip validation for relief for conversion-----//
        If bUseBulk Then
            v_rsAssignmentRequest.MoveFirst
    'lngStart = GetTickCount()
    '//-----Test 3/22/2002 (locking up after save???)-----//
            If Not IsReliefComponentValidBulk(v_rsAssignmentRequest, r_rsErrors) Then
                bError = True
            End If
    'lngStop = GetTickCount()
    'lngTotalRelief = lngTotalRelief + (lngStop - lngStart)
        End If
    End If
    
'//-----for testing performance-----//
'l_strBRVMessage = "TotalLines = " & lngTotalLines & "    TotalRelief = " & lngTotalRelief
'AppendBusinessError r_rsErrors, trErrbusValidationError, mstrCBUSINESSWAR, _
'    l_strBRVMessage, m_strCMODULENAME, _
'    1, 1, "LINE_NUM", _
'    Array("WG_NUM", "LINE_NUM", "START_DATE", "END_DATE", "EMP_NUM"), _
'    CStr(1), _
'    CStr(1), _
'    Format(Now(), "mm/dd/yyyy"), _
'    Format(Now(), "mm/dd/yyyy"), _
'    CStr(1)
 
    If Not bError Then
        If DoesOverlapExist(v_rsAssignmentRequest, r_rsErrors) Then
            bError = True
        Else
            bError = False
        End If
    End If
    
    If bError Then
        PreValidateAssignments = False
        GetObjectContext.SetAbort
    Else
        PreValidateAssignments = True
        GetObjectContext.SetComplete
    End If

    Set l_rsLinesBulk = Nothing
'    Set l_rsEmployeesBulk = Nothing
'//    Set l_objDBSchDmd = Nothing
    Set l_objDBWorkGroup = Nothing
    Exit Function
eh:
    Set l_rsLinesBulk = Nothing
'//    Set l_objDBSchDmd = Nothing
    Set l_objDBWorkGroup = Nothing
    GetObjectContext.SetAbort
    TrsRaiseError m_strCMODULENAME, "PreValidateAssignments", , , , , True
End Function

Private Function IsReliefComponentValid(ByVal v_lngEmpNum As Long, _
                                        ByVal v_datStartDate As Date, _
                                        ByVal v_datEndDate As Date, _
                                        ByVal v_strReliefComp As String, _
                                        ByRef r_strBRVMessage As String _
                                        ) As Boolean
    Dim l_rsEmpInfo As ADODB.Recordset
    Dim l_datJobEffectiveDate As Date, l_datPriorJobEffectiveDate As Date
    Dim l_strUnion As String, l_strPriorUnion As String
    Dim l_objDBEmployee As dbEmployee.cdEmployee
    Dim l_objBUSCodeTables As busCodeTables.cbsCodeTables
    Dim l_enumValidType As enumReliefType '//-----1) use union, 2) user prior union & union, 3) user prior union
    Dim l_bValid As Boolean, l_bValid2 As Boolean
    
    On Error GoTo eh

    l_bValid = False
    l_bValid2 = True
    
    Set l_objDBEmployee = CreateObject("dbEmployee.cdEmployee")
    Set l_rsEmpInfo = l_objDBEmployee.GetEmployee(v_lngEmpNum)
    Set l_objDBEmployee = Nothing
    
    If RecordsExist(l_rsEmpInfo) Then
        l_strUnion = l_rsEmpInfo!UNION_CD
        l_datJobEffectiveDate = l_rsEmpInfo!JOB_EFF_DATE
        l_strPriorUnion = l_rsEmpInfo!PRIOR_UNION_CD
        l_datPriorJobEffectiveDate = l_rsEmpInfo!PRIOR_JOB_EFF_DT
        
        l_enumValidType = DetermineReliefType(v_datStartDate, v_datEndDate, _
                        l_strUnion, l_datJobEffectiveDate, l_strPriorUnion, l_datPriorJobEffectiveDate)
    End If
    
    Select Case l_enumValidType
    Case rtUseUnion '//-----empnum is in 1 union for entire date range
        Set l_objBUSCodeTables = CreateObject("busCodeTables.cbsCodeTables")
        l_bValid = l_objBUSCodeTables.IsVacationReliefCode(l_strUnion, v_datStartDate, v_strReliefComp)
        Set l_objBUSCodeTables = Nothing
        If Not l_bValid Then
            r_strBRVMessage = "Relief component " & v_strReliefComp & " is not valid for employee " & v_lngEmpNum & _
                                " in union " & l_strUnion & _
                                " between " & v_datStartDate & " and " & v_datEndDate & "."
        End If
    Case rtUseUnionAndPriorUnion '//-----user prior union & union
        Set l_objBUSCodeTables = CreateObject("busCodeTables.cbsCodeTables")
        l_bValid = l_objBUSCodeTables.IsVacationReliefCode(l_strPriorUnion, v_datStartDate, v_strReliefComp)
        '//If l_strPriorUnion <> l_strUnion Then
        l_bValid2 = l_objBUSCodeTables.IsVacationReliefCode(l_strUnion, l_datJobEffectiveDate, v_strReliefComp)
        'Else
        'l_bValid2 = l_bValid
        'end if
        If l_bValid Then
            If l_bValid2 Then
            
            Else
                r_strBRVMessage = "Relief component " & v_strReliefComp & " is not valid for employee " & v_lngEmpNum & _
                                " in union " & l_strUnion & _
                                " between " & l_datJobEffectiveDate & " and " & v_datEndDate & "."
            End If
        Else
            If l_bValid2 Then
                If l_datJobEffectiveDate <> m_datCOPTIONALSTARTDATE Then
                    r_strBRVMessage = "Relief component " & v_strReliefComp & " is not valid for employee " & v_lngEmpNum & _
                                    " in prior union " & l_strPriorUnion & _
                                    " between " & v_datStartDate & " and " & DateAdd("d", -1, l_datJobEffectiveDate) & "."
                Else
                    r_strBRVMessage = "Relief component " & v_strReliefComp & " is not valid for employee " & v_lngEmpNum & _
                                    " in prior union " & l_strPriorUnion & _
                                    " between " & v_datStartDate & " and " & l_datJobEffectiveDate & "."
                End If
            Else
                r_strBRVMessage = "Relief component " & v_strReliefComp & " is not valid for employee " & v_lngEmpNum & _
                                " in union/prior union " & l_strUnion & "/" & l_strPriorUnion & _
                                " between " & v_datStartDate & " and " & v_datEndDate & "."
            End If
        End If
        Set l_objBUSCodeTables = Nothing
    Case rtUsePriorUnion '//-----use prior union
        Set l_objBUSCodeTables = CreateObject("busCodeTables.cbsCodeTables")
        l_bValid = l_objBUSCodeTables.IsVacationReliefCode(l_strPriorUnion, v_datStartDate, v_strReliefComp)
        Set l_objBUSCodeTables = Nothing
        
        If Not l_bValid Then
            r_strBRVMessage = "Relief component " & v_strReliefComp & " is not valid for employee " & v_lngEmpNum & _
                                " in prior union " & l_strPriorUnion & _
                                " between " & v_datStartDate & " and " & v_datEndDate & "."
        End If
    Case rtErrorNothingBeforePriorUnion
        If l_datPriorJobEffectiveDate <> m_datCOPTIONALSTARTDATE Then
            r_strBRVMessage = "Cannot validate Relief component " & v_strReliefComp & _
                " because unable to determine union code between " & v_datStartDate & " and " & DateAdd("d", -1, l_datPriorJobEffectiveDate) & "."
        Else
            r_strBRVMessage = "Cannot validate Relief component " & v_strReliefComp & _
                " because unable to determine union code between " & v_datStartDate & " and " & l_datPriorJobEffectiveDate & "."
        End If
    Case rtErrorNoPriorUnion
        r_strBRVMessage = "Cannot validate Relief component " & v_strReliefComp & _
        " because unable to determine union code before " & l_datJobEffectiveDate & "."
    Case rtErrorUnknown
        r_strBRVMessage = "Cannot validate Relief component " & v_strReliefComp & " because of unknown error."
    Case Else
        r_strBRVMessage = "Cannot validate Relief component " & v_strReliefComp & " because of unknown error."
    End Select
    
    IsReliefComponentValid = l_bValid And l_bValid2
    If IsReliefComponentValid Then
        GetObjectContext.SetComplete
    Else
        GetObjectContext.SetAbort
    End If
    Set l_rsEmpInfo = Nothing
    Exit Function
eh:
    Set l_objDBEmployee = Nothing
    Set l_objBUSCodeTables = Nothing
    Set l_rsEmpInfo = Nothing
    GetObjectContext.SetAbort
    TrsRaiseError m_strCMODULENAME, "IsReliefComponentValid", , , , , True
End Function

Private Function PostValidateAssignments(ByVal v_rsAssignmentRequest As ADODB.Recordset, ByRef r_rsErrors As ADODB.Recordset) As Boolean
    Dim bError As Boolean
'    Dim v_strErrorDesc As String

    On Error GoTo eh
'    If DoesOverlapExist(v_rsAssignmentRequest, r_rsErrors) Then
'        bError = True
'    Else
'        bError = False
'    End If
'
'    If Not ValidateEmployeeNumber(v_rsAssignmentRequest, v_strErrorDesc) Then
'        Err.Raise trErrbusInvalidEmpNum, m_strCMODULENAME, v_strErrorDesc
'    End If
    
    If bError Then
        PostValidateAssignments = False
        GetObjectContext.SetAbort
    Else
        PostValidateAssignments = True
        GetObjectContext.SetComplete
    End If

    Exit Function
eh:
    GetObjectContext.SetAbort
    TrsRaiseError m_strCMODULENAME, "PostValidateAssignments", , , , , True
End Function
'
' 8/15/02 RSG
' Originally intended this to cause the application to stop when attempting
' to save a line assignment with EMP_NUM = 0.
' Commented out this validation because it cannot be implemented while
' an unassign request, with EMP_NUM = 0, stays on the assignment recordset
' after processing.  Currently, the data service, etc ignore this request
' by filtering it out of scope.  But, there does not appear to be a way to
' distinguish this client request from the a data service request.
'
'Private Function ValidateEmployeeNumber( _
'            ByVal v_rsAssignmentRequest As ADODB.Recordset, _
'            ByRef r_strErrorDesc As String) As Boolean
'' Answer False if any records have an invalid employee number; otherwise True.
'' If False, answer a formated error description presentable to user and logging.
'    Dim l_blnValid As Boolean
'    Dim l_colDescArgs As Collection
'
'On Error GoTo eh
'
'    l_blnValid = True
'    If Not v_rsAssignmentRequest Is Nothing Then
'        v_rsAssignmentRequest.Filter = adFilterNone
'        If RecordsExist(v_rsAssignmentRequest) Then
'            v_rsAssignmentRequest.MoveFirst
'        Else
'            ValidateEmployeeNumber = True
'            Exit Function
'        End If
'    Else
'        ValidateEmployeeNumber = True
'        Exit Function
'    End If
'
'    Do Until v_rsAssignmentRequest.EOF
'        If v_rsAssignmentRequest!RECORD_STATUS = "I" Or _
'            v_rsAssignmentRequest!RECORD_STATUS = "C" Then
'            If v_rsAssignmentRequest!EMP_NUM <= 0 Then
'                l_blnValid = False
'                Exit Do
'            End If
'        End If
'        v_rsAssignmentRequest.MoveNext
'    Loop
'
'    If Not l_blnValid Then
'        Set l_colDescArgs = New Collection
'        l_colDescArgs.Add SaveAction(v_rsAssignmentRequest!RECORD_STATUS)
'        l_colDescArgs.Add CStr(v_rsAssignmentRequest!EMP_NUM)
'        l_colDescArgs.Add KeyValues(v_rsAssignmentRequest)
'        r_strErrorDesc = TrsGetErrDescription(LongAppErr(trErrbusInvalidEmpNum), l_colDescArgs)
'    End If
'
'    ValidateEmployeeNumber = l_blnValid
'    v_rsAssignmentRequest.MoveFirst
'    Set l_colDescArgs = Nothing
'    Exit Function
'
'eh:
'    Set l_colDescArgs = Nothing
'    TrsRaiseError m_strCMODULENAME, "ValidateEmployeeNumber"
'End Function
'
'Private Function SaveAction(ByVal v_strRecStatus As String) As String
'
'On Error GoTo eh
'
'    Select Case v_strRecStatus
'        Case "I"
'            SaveAction = "Insert"
'        Case "C"
'            SaveAction = "Change"
'        Case "D"
'            SaveAction = "Delete"
'        Case Else
'            SaveAction = "Unknown(" & v_strRecStatus & ")"
'    End Select
'    Exit Function
'
'eh:
'    TrsRaiseError m_strCMODULENAME, "SaveAction"
'End Function
'
'
'Private Function KeyValues(ByVal v_rsAssignmentRequest As ADODB.Recordset) As String
'    Dim l_strWgNum As String
'    Dim l_strLineNum As String
'    Dim l_strEmpNum As String
'    Dim l_strStartDate As String
'    Dim l_strEndDate As String
'
'On Error GoTo eh
'
'    If v_rsAssignmentRequest!RECORD_STATUS = "I" Then
'        l_strWgNum = FieldValueAsString("WG_NUM", v_rsAssignmentRequest)
'        l_strLineNum = FieldValueAsString("LINE_NUM", v_rsAssignmentRequest)
'        l_strEmpNum = FieldValueAsString("EMP_NUM", v_rsAssignmentRequest)
'        l_strStartDate = FieldValueAsString("START_DATE", v_rsAssignmentRequest)
'        l_strEndDate = FieldValueAsString("END_DATE", v_rsAssignmentRequest)
'    Else
'        l_strWgNum = FieldValueAsString("KEY_WG_NUM", v_rsAssignmentRequest)
'        l_strLineNum = FieldValueAsString("KEY_LINE_NUM", v_rsAssignmentRequest)
'        l_strEmpNum = FieldValueAsString("KEY_EMP_NUM", v_rsAssignmentRequest)
'        l_strStartDate = FieldValueAsString("KEY_START_DATE", v_rsAssignmentRequest)
'        l_strEndDate = FieldValueAsString("KEY_END_DATE", v_rsAssignmentRequest)
'    End If
'
'    KeyValues = _
'        Join( _
'            Array(l_strWgNum, l_strLineNum, l_strEmpNum, l_strStartDate, l_strEndDate), _
'            ";")
'
'    Exit Function
'
'eh:
'    TrsRaiseError m_strCMODULENAME, "KeyValues"
'End Function
'
'
'Private Function FieldValueAsString( _
'            ByVal v_strFieldName As String, _
'            ByVal v_rsAssignmentRequest As ADODB.Recordset) As String
'
'On Error GoTo eh
'
'    If IsNull(v_rsAssignmentRequest.Fields(v_strFieldName).Value) Then
'        FieldValueAsString = "NULL"
'    ElseIf IsEmpty(v_rsAssignmentRequest.Fields(v_strFieldName).Value) Then
'        FieldValueAsString = "EMPTY"
'    Else
'        FieldValueAsString = CStr(v_rsAssignmentRequest.Fields(v_strFieldName).Value)
'    End If
'
'    Exit Function
'
'eh:
'    TrsRaiseError m_strCMODULENAME, "FieldValueAsString"
'End Function


'//-----check for any overlaps within the client request-----//
Private Function DoesOverlapExist(ByVal v_rsAssignmentRequest As ADODB.Recordset, ByRef r_rsErrors As ADODB.Recordset) As Boolean
    Dim bError As Boolean
    Dim xx As Long
    Dim l_strFilter As String '//, l_strOverlapMsg As String
    Dim l_rsOverlap As ADODB.Recordset
    Dim l_varAssignmentBookmark As Variant
    
    On Error GoTo eh
    bError = False
    DoesOverlapExist = True

    '//-----exclude deletes and unassignments (empnum=0)
    l_strFilter = "RECORD_STATUS <> 'D' AND EMP_NUM > 0"
    v_rsAssignmentRequest.Filter = MDAC25FilterFix(l_strFilter)

    If RecordsExist(v_rsAssignmentRequest) Then
        With v_rsAssignmentRequest
            Set l_rsOverlap = .Clone
            .MoveFirst
            Do While Not .EOF
                l_varAssignmentBookmark = .Bookmark
                l_strFilter = "EMP_NUM = " & !EMP_NUM & _
                            " AND START_DATE <= #" & !END_DATE & "#" & _
                            " AND END_DATE >= #" & !START_DATE & "#" & _
                            " AND RECORD_STATUS <> 'D'"
                If OverlapExistFilter(l_rsOverlap, l_strFilter, l_varAssignmentBookmark) Then
                    bError = True
                    '//l_strOverlapMsg = "Overlap was found creating ambiguity with line " & l_rsOverlap!KEY_LINE_NUM
                    OverlapExistBRV v_rsAssignmentRequest, l_rsOverlap, "LineNum", r_rsErrors
                End If
                l_strFilter = "WG_NUM = " & !WG_NUM & _
                            " AND LINE_NUM = " & !LINE_NUM & _
                            " AND START_DATE <= #" & !END_DATE & "#" & _
                            " AND END_DATE >= #" & !START_DATE & "#" & _
                            " AND RECORD_STATUS <> 'D'"
                If OverlapExistFilter(l_rsOverlap, l_strFilter, l_varAssignmentBookmark) Then
                    bError = True
                    '//l_strOverlapMsg = "Overlap was found creating ambiguity with employee " & l_rsOverlap!KEY_EMP_NUM
                    OverlapExistBRV v_rsAssignmentRequest, l_rsOverlap, "EmpNum", r_rsErrors
                End If

                .MoveNext
            Loop
            
            Set l_rsOverlap = Nothing
        End With
    End If
    v_rsAssignmentRequest.Filter = adFilterNone
        
    If bError Then
        DoesOverlapExist = True
        GetObjectContext.SetAbort
    Else
        DoesOverlapExist = False
        GetObjectContext.SetComplete
    End If

    Exit Function
eh:
    GetObjectContext.SetAbort
    TrsRaiseError m_strCMODULENAME, "DoesOverlapExist", , , , , True
End Function

'//-----apply filter and check for any row other than the one matching the bookmark-----//
'//-----hey MS what about rs.Filter = " <column expression> And BOOKMARK <> " & bmkAssignment
Private Function OverlapExistFilter(ByVal v_rsOverlap As ADODB.Recordset, _
                                        ByVal v_strFilter As String, _
                                        ByVal v_varAssignmentBookmark As Variant) As Boolean
    Dim l_bolOverlapFound As Boolean
    On Error GoTo eh
    
    l_bolOverlapFound = False
    
    With v_rsOverlap
        .Filter = MDAC25FilterFix(v_strFilter)
        If Not (.BOF And .EOF) Then
            .MoveFirst
            Do While Not .EOF And Not l_bolOverlapFound
                If .Bookmark <> v_varAssignmentBookmark Then
                    l_bolOverlapFound = True
                    Exit Do
                End If
                .MoveNext
            Loop
        End If
    End With
    
    OverlapExistFilter = l_bolOverlapFound
    Exit Function
eh:
    GetObjectContext.SetAbort
    TrsRaiseError m_strCMODULENAME, "OverlapExistFilter", , , , , True
End Function

'//-----suppress duplicate error messages for overlap
Private Function OverlapExistBRV(ByVal v_rsAssignmentRequest As ADODB.Recordset, _
                                ByVal v_rsOverlap As ADODB.Recordset, _
                                ByVal v_strField As String, _
                                ByRef r_rsErrors As ADODB.Recordset) As Boolean
    Dim l_strFilter As String, l_strOverlapMsg As String
    
    On Error GoTo eh

    With v_rsAssignmentRequest
        Do While Not v_rsOverlap.EOF
            If .Bookmark <> v_rsOverlap.Bookmark Then '//-----ignore same row
                l_strFilter = "(ERR_NUMBER = " & trErrbusSchedDemandOverlap & _
                            " AND KEY_FIELD1 = " & !WG_NUM & _
                            " AND KEY_FIELD2 = " & !LINE_NUM & _
                            " AND KEY_FIELD3 = #" & !START_DATE & "#" & _
                            " AND KEY_FIELD4 = #" & !END_DATE & "#" & _
                            " AND KEY_FIELD5 = " & !EMP_NUM & _
                            " AND KEY_FIELD6 = " & v_rsOverlap!WG_NUM & _
                            " AND KEY_FIELD7 = " & v_rsOverlap!LINE_NUM & _
                            " AND KEY_FIELD8 = #" & v_rsOverlap!START_DATE & "#" & _
                            " AND KEY_FIELD9 = #" & v_rsOverlap!END_DATE & "#" & _
                            " AND KEY_FIELD10 = " & v_rsOverlap!EMP_NUM & ")" & _
                            " OR (ERR_NUMBER = " & trErrbusSchedDemandOverlap & _
                            " AND KEY_FIELD1 = " & v_rsOverlap!WG_NUM & _
                            " AND KEY_FIELD2 = " & v_rsOverlap!LINE_NUM & _
                            " AND KEY_FIELD3 = #" & v_rsOverlap!START_DATE & "#" & _
                            " AND KEY_FIELD4 = #" & v_rsOverlap!END_DATE & "#" & _
                            " AND KEY_FIELD5 = " & v_rsOverlap!EMP_NUM & _
                            " AND KEY_FIELD6 = " & !WG_NUM & _
                            " AND KEY_FIELD7 = " & !LINE_NUM & _
                            " AND KEY_FIELD8 = #" & !START_DATE & "#" & _
                            " AND KEY_FIELD9 = #" & !END_DATE & "#" & _
                            " AND KEY_FIELD10 = " & !EMP_NUM & ")"
                r_rsErrors.Filter = MDAC25FilterFix(l_strFilter)
                If Not RecordsExist(r_rsErrors) Then
                    '//-----display the actual overlap, use Maximum of Start Dates and Minimum of End Dates
                    '//-----given s1=2/1/2002, e1=12/31/9999, s2=2/5/2002, e2=2/10/2002, return 2/5 - 2/10
                    If v_strField = "LineNum" Then
                        l_strOverlapMsg = "Overlap was found creating ambiguity with line " & v_rsOverlap!LINE_NUM
                    ElseIf v_strField = "EmpNum" Then
                        l_strOverlapMsg = "Overlap was found creating ambiguity with employee " & v_rsOverlap!EMP_NUM
                    Else
                        l_strOverlapMsg = "Overlap was found creating ambiguity "
                    End If
                    
                    If v_rsOverlap!START_DATE > !START_DATE Then
                        l_strOverlapMsg = l_strOverlapMsg & " between " & v_rsOverlap!START_DATE
                    Else
                        l_strOverlapMsg = l_strOverlapMsg & " between " & !START_DATE
                    End If
                    
                    If v_rsOverlap!END_DATE < !END_DATE Then
                        l_strOverlapMsg = l_strOverlapMsg & " and " & v_rsOverlap!END_DATE & "."
                    Else
                        l_strOverlapMsg = l_strOverlapMsg & " and " & !END_DATE & "."
                    End If
                    '//-----add extra fields to BRV so that duplicate BRVs can be suppressed-----//
                    AppendBusinessError r_rsErrors, trErrbusSchedDemandOverlap, mstrCBUSINESSERR, _
                        l_strOverlapMsg, _
                        m_strCMODULENAME, _
                        1, 1, "EMP_NUM", _
                        Array("WG_NUM", "LINE_NUM", "START_DATE", "END_DATE", "EMP_NUM", _
                        "WG_NUM2", "LINE_NUM2", "START_DATE2", "END_DATE2", "EMP_NUM2"), _
                        CStr(!WG_NUM), _
                        CStr(!LINE_NUM), _
                        Format(!START_DATE, "mm/dd/yyyy"), _
                        Format(!END_DATE, "mm/dd/yyyy"), _
                        CStr(!EMP_NUM), _
                        CStr(v_rsOverlap!WG_NUM), _
                        CStr(v_rsOverlap!LINE_NUM), _
                        Format(v_rsOverlap!START_DATE, "mm/dd/yyyy"), _
                        Format(v_rsOverlap!END_DATE, "mm/dd/yyyy"), _
                        CStr(v_rsOverlap!EMP_NUM)
                End If
            End If
            v_rsOverlap.MoveNext
        Loop
    End With
    
    r_rsErrors.Filter = adFilterNone
    OverlapExistBRV = True
    
    Exit Function
eh:
    GetObjectContext.SetAbort
    TrsRaiseError m_strCMODULENAME, "OverlapExistBRV", , , , , True
End Function


Private Function ValidateAssignments(ByVal v_rsAssignmentRequest As ADODB.Recordset, ByRef r_rsErrors As ADODB.Recordset) As Boolean
    Dim bError As Boolean, bOverlap As Boolean
    Dim l_rsOverlapRequest As ADODB.Recordset
    Dim l_rsUnassignRequest As ADODB.Recordset
    Dim l_rsOverlapInScope As ADODB.Recordset
    Dim l_rsOriginalScope As ADODB.Recordset
    Dim l_strFilter As String, xx As Long
    
    bError = False
    ValidateAssignments = False
    On Error GoTo eh

    If RecordsExist(v_rsAssignmentRequest) Then
        With v_rsAssignmentRequest
    
            '//-----get all rows needed for validation once, and apply changes because of overlaps to it to prevent dirty reads-----//
            '//-----check for overlaps in other WG with EmpNum (1/25/2002)-----//
            Set l_rsOverlapInScope = GetOverlapInScope(v_rsAssignmentRequest)
            
            '//-----create copy before any changes are applied so that net deletes (difference) can be derived later-----//
            Set l_rsOriginalScope = CopyRecordset(l_rsOverlapInScope, 2)

            Set l_rsOverlapRequest = CopyRecordset(v_rsAssignmentRequest, 0) '//-----create empty rs for overlap-----//
            l_rsOverlapRequest.Open
            Set l_rsUnassignRequest = CopyRecordset(v_rsAssignmentRequest, 0) '//-----create empty rs for unassigns-----//
            l_rsUnassignRequest.Open

            '//-----at this point all requests are either "I" or "D" because of ProcessKeyChange-----//
            '//-----could do GetOverlapInScope after PreValidateAssignments and the following before ProcessKeyChange
            .MoveFirst
            Do While Not .EOF
                '//-----Check existence of change and delete requests-----//
                If !RECORD_STATUS = "C" Or !RECORD_STATUS = "D" Then
                    If Not DoesRequestExist(v_rsAssignmentRequest, l_rsOverlapInScope, r_rsErrors) Then
                        bError = True
                    End If
                End If
                .MoveNext
            Loop
            
            '//-----it is possible to delete a row that will be part of the unassign or overlap requests-----//
            '//-----this would then produce 2 delete request which will fail-----//
            '//-----Therefore, process in "D", "I", "C" order (tnv 8/12/2003)-----//
            If Not bError Then
            For xx = 1 To 3
            .Filter = "RECORD_STATUS = '" & IIf(xx = 1, "D", IIf(xx = 2, "I", "C")) & "'"
            If RecordsExist(v_rsAssignmentRequest) Then
                .MoveFirst
                Do While Not .EOF
                    If !RECORD_STATUS = "I" Or !RECORD_STATUS = "C" Then
                        '//-----Check for Overlap in StartDate/EndDate for an employee-----//
                        '//-----(don't allow employee to be assigned to more than 1 line for a date)-----//
                        If !EMP_NUM > 0 Then
                            '//-----now check for overlap by employee (emp cannot be assigned to more than 1 line for any date)-----//
                            If Not AdjustOverlapByEmp(v_rsAssignmentRequest, l_rsOverlapInScope, l_rsUnassignRequest, r_rsErrors) Then
    '                            bError = True
                                '//-----Adjust start and end dates for assignments that would overlap the unassign request-----//
                                If RecordsExist(l_rsUnassignRequest) Then
                                    l_rsUnassignRequest.MoveFirst
                                    Do While Not l_rsUnassignRequest.EOF
                                        AdjustOverlapByLine l_rsUnassignRequest, l_rsOverlapInScope, l_rsOverlapRequest, r_rsErrors, v_rsAssignmentRequest!WG_NUM
                                        ApplyPendingChanges l_rsOverlapRequest, l_rsOverlapInScope
                                        l_rsUnassignRequest.Delete
                                        l_rsUnassignRequest.MoveNext
                                    Loop
                                End If
                            End If
                        End If
                    
                        '//-----Adjust start and end dates for assignments that would overlap the request-----//
                        AdjustOverlapByLine v_rsAssignmentRequest, l_rsOverlapInScope, l_rsOverlapRequest, r_rsErrors, v_rsAssignmentRequest!WG_NUM
                        ApplyPendingChanges l_rsOverlapRequest, l_rsOverlapInScope
                    End If
                    
                    '//-----apply current record only-----//
                    ApplyPendingChange v_rsAssignmentRequest, l_rsOverlapInScope
    
                    .MoveNext
                Loop
            End If
            Next xx
            .Filter = adFilterNone
            End If
        End With
    End If
    
    If Not bError Then
        '//-----Problem: 12/27/2001 The overlap requests are based on processing order, but the db component is based on record status order (DIC)-----//
        '//-----    If a block of time gets broken for an overlap more than once, it could generate multiple delete requests
        '//-----        which will be executed before the inserts and mess things up
        '//-----Solution: for each delete request check for a corresponding insert and remove both-----//
        '//-----            Otherwise, send 2 requests rs to the SaveAssignments method. 1 for client requests and the other for overlap requests
        If RecordsExist(l_rsOverlapRequest) Then
            ApplyNetChanges l_rsOverlapRequest
            
            '//-----copy the overlap request into the original request record set-----//
            AppendRecordset l_rsOverlapRequest, v_rsAssignmentRequest
        '    AppendRecordset l_rsUnassignRequest, v_rsAssignmentRequest
        End If
        
    '//-----3/1/2002, Since relief is now a part of Demand Scheduling, the relief table is not needed-----//
    '    '//-----derive the net effect for deletes (difference), so that what is deleted can be validated-----//
    '    '//-----check to see if any relief exists for assignments that get deleted
    '    '//-----   either directly with record_status="D", "C" that gets changed to "D", "I", or from an overlap-----//
    '    If Not ValidateNetDelete(l_rsOriginalScope, l_rsOverlapInScope, r_rsErrors) Then
    '        bError = True
    '    End If
    End If
    
    If bError Then
        ValidateAssignments = False
        GetObjectContext.SetAbort
    Else
        ValidateAssignments = True
        GetObjectContext.SetComplete
    End If

    Set l_rsOverlapRequest = Nothing
    Set l_rsUnassignRequest = Nothing
    Set l_rsOverlapInScope = Nothing
    
    Exit Function
eh:
    Set l_rsOverlapRequest = Nothing
    Set l_rsUnassignRequest = Nothing
    Set l_rsOverlapInScope = Nothing

    GetObjectContext.SetAbort
    TrsRaiseError m_strCMODULENAME, "ValidateAssignments", , , , , True
End Function

'//-----find minimum and maximum values in order to retrieve all scope for validation-----//
Private Function GetOverlapInScope(ByVal v_rsAssignmentRequest As ADODB.Recordset) As ADODB.Recordset
    Dim l_rsAssignmentRequestClone As ADODB.Recordset
    Dim l_rsOverlapInScope As ADODB.Recordset
    Dim l_objDBSchDmd As dbSchedDemand.cdSchedDemand
    Dim l_datMinStartDate As Date, l_datMaxEndDate As Date
    Dim l_lngWGNumSave As Long, l_lngLineNumSave As Long, l_lngEmpNumSave As Long
    Dim l_colScopeParm As New Collection
'    Dim l_tScopeParm As tScopeParm
    Dim l_tScopeParm(4) As Variant
    
    On Error GoTo eh
                    
    Set l_rsAssignmentRequestClone = v_rsAssignmentRequest.Clone
    
    With l_rsAssignmentRequestClone
        '//-----build list of necessary lines-----//
        .Sort = "WG_NUM ASC, LINE_NUM ASC"
    
        .MoveFirst

        l_datMinStartDate = !START_DATE
        l_datMaxEndDate = !END_DATE
        l_lngWGNumSave = !WG_NUM
        l_lngLineNumSave = !LINE_NUM
            
        Do While Not .EOF
            '//-----check for change in wg/line
            If !WG_NUM <> l_lngWGNumSave Or !LINE_NUM <> l_lngLineNumSave Then
                l_tScopeParm(lngMinStartDate) = l_datMinStartDate
                l_tScopeParm(lngMaxEndDate) = l_datMaxEndDate
                l_tScopeParm(lngWgNum) = l_lngWGNumSave
                l_tScopeParm(lngLineNum) = l_lngLineNumSave
                l_tScopeParm(lngEmpNum) = 0 '//-----mutually exclusive with wg/line
                l_colScopeParm.Add l_tScopeParm
                
                l_datMinStartDate = !START_DATE '//-----reset for next set of parms-----//
                l_datMaxEndDate = !END_DATE
                l_lngWGNumSave = !WG_NUM
                l_lngLineNumSave = !LINE_NUM
            End If
            
            If !START_DATE < l_datMinStartDate Then
                l_datMinStartDate = !START_DATE
            End If
            If !END_DATE > l_datMaxEndDate Then
                l_datMaxEndDate = !END_DATE
            End If
                                      
            .MoveNext
        Loop
        l_tScopeParm(lngMinStartDate) = l_datMinStartDate
        l_tScopeParm(lngMaxEndDate) = l_datMaxEndDate
        l_tScopeParm(lngWgNum) = l_lngWGNumSave
        l_tScopeParm(lngLineNum) = l_lngLineNumSave
        l_tScopeParm(lngEmpNum) = 0
    
        l_colScopeParm.Add l_tScopeParm

        '//-----build list of necessary employees-----//
        .Sort = "EMP_NUM ASC"
    
        .MoveFirst

        l_datMinStartDate = !START_DATE
        l_datMaxEndDate = !END_DATE
        l_lngEmpNumSave = !EMP_NUM
            
        Do While Not .EOF
            '//-----check for change in employee
            If !EMP_NUM <> l_lngEmpNumSave Then
                l_tScopeParm(lngMinStartDate) = l_datMinStartDate
                l_tScopeParm(lngMaxEndDate) = l_datMaxEndDate
                l_tScopeParm(lngWgNum) = 0 '//-----mutually exclusive with EmpNum
                l_tScopeParm(lngLineNum) = 0
                l_tScopeParm(lngEmpNum) = l_lngEmpNumSave
                l_colScopeParm.Add l_tScopeParm
                
                l_datMinStartDate = !START_DATE '//-----reset for next set of parms-----//
                l_datMaxEndDate = !END_DATE
                l_lngEmpNumSave = !EMP_NUM
            End If
            
            If !START_DATE < l_datMinStartDate Then
                l_datMinStartDate = !START_DATE
            End If
            If !END_DATE > l_datMaxEndDate Then
                l_datMaxEndDate = !END_DATE
            End If
                                      
            .MoveNext
        Loop
        l_tScopeParm(lngMinStartDate) = l_datMinStartDate
        l_tScopeParm(lngMaxEndDate) = l_datMaxEndDate
        l_tScopeParm(lngWgNum) = 0 '//-----mutually exclusive with EmpNum
        l_tScopeParm(lngLineNum) = 0
        l_tScopeParm(lngEmpNum) = l_lngEmpNumSave
        l_colScopeParm.Add l_tScopeParm

        .MoveFirst
    End With
    
    Set l_rsAssignmentRequestClone = Nothing
    
    Set l_objDBSchDmd = CreateObject("dbSchedDemand.cdSchedDemand")
    Set l_rsOverlapInScope = l_objDBSchDmd.GetOverlapInScope(l_colScopeParm)
    Set l_objDBSchDmd = Nothing

    Set GetOverlapInScope = l_rsOverlapInScope
    Set l_rsOverlapInScope = Nothing
    Exit Function
eh:
    Set l_objDBSchDmd = Nothing
    Set l_rsOverlapInScope = Nothing
    
    GetObjectContext.SetAbort
    TrsRaiseError m_strCMODULENAME, "GetOverlapInScope", , , , , True
End Function

'//-----make sure that changes/deletes exist
Private Function DoesRequestExist(ByVal v_rsAssignmentRequest As ADODB.Recordset, _
                                    ByVal v_rsOverlapInScope As ADODB.Recordset, _
                                    ByRef r_rsErrors As ADODB.Recordset) As Boolean
    Dim l_strFilter As String
    Dim sBRVMsg As String
    
    On Error GoTo eh
    DoesRequestExist = True
    
    With v_rsAssignmentRequest
        l_strFilter = "KEY_WG_NUM = " & !KEY_WG_NUM & _
                    " AND KEY_START_DATE = #" & !KEY_START_DATE & "#" & _
                    " AND KEY_END_DATE = #" & !KEY_END_DATE & "#" & _
                    " AND KEY_EMP_NUM = " & !KEY_EMP_NUM & _
                    " AND KEY_LINE_NUM = " & !KEY_LINE_NUM
    
        v_rsOverlapInScope.Filter = MDAC25FilterFix(l_strFilter)
    
        If Not RecordsExist(v_rsOverlapInScope) Then
            sBRVMsg = "Invalid " & IIf(!RECORD_STATUS = "C", "change", "delete") & " request."
    
            AppendBusinessError r_rsErrors, trErrbusValidationError, mstrCBUSINESSERR, _
                sBRVMsg, m_strCMODULENAME, _
                1, 1, "START_DATE", _
                Array("KEY_WG_NUM", "KEY_LINE_NUM", "KEY_START_DATE", "KEY_END_DATE", "KEY_EMP_NUM"), _
                CStr(!KEY_WG_NUM), _
                CStr(!KEY_LINE_NUM), _
                Format(!KEY_START_DATE, "mm/dd/yyyy"), _
                Format(!KEY_END_DATE, "mm/dd/yyyy"), _
                CStr(!KEY_EMP_NUM)
            DoesRequestExist = False
        End If
        v_rsOverlapInScope.Filter = adFilterNone
    End With
    
    Exit Function
eh:
    DoesRequestExist = False
    GetObjectContext.SetAbort
    TrsRaiseError m_strCMODULENAME, "DoesRequestExist", , , , , True
End Function

Private Function AdjustOverlapByEmp(ByVal v_rsAssignmentRequest As ADODB.Recordset, _
                                    ByVal v_rsOverlapInScope As ADODB.Recordset, _
                                    ByRef r_rsUnassignRequest As ADODB.Recordset, _
                                    ByRef r_rsErrors As ADODB.Recordset) As Boolean
    Dim l_strFilter As String
    Dim sBRVMsg As String
    
    On Error GoTo eh
    AdjustOverlapByEmp = True
        
    With v_rsAssignmentRequest
        If RecordsExist(v_rsAssignmentRequest) Then
            '//-----locate matching record in scope overlap-----//
            '//-----ignore scope with same wg and line, let AdjustOverlapByLine handle those-----//
            l_strFilter = "(START_DATE <= #" & !END_DATE & "#" & _
                        " AND END_DATE >= #" & !START_DATE & "#" & _
                        " AND EMP_NUM = " & !EMP_NUM & _
                        " AND WG_NUM <> " & !WG_NUM & ") OR" & _
                        " (START_DATE <= #" & !END_DATE & "#" & _
                        " AND END_DATE >= #" & !START_DATE & "#" & _
                        " AND EMP_NUM = " & !EMP_NUM & _
                        " AND LINE_NUM <> " & !LINE_NUM & ")"
            v_rsOverlapInScope.Filter = MDAC25FilterFix(l_strFilter)
        End If
    End With
    
    If RecordsExist(v_rsOverlapInScope) Then
        With v_rsOverlapInScope
            AdjustOverlapByEmp = False
            .MoveFirst
            Do While Not .EOF
                '//-----for each overlap create an unassign request with a warning, rather than an error-----//
                '//-----try to be more specific about the unassign request (delete), include emp_num
                r_rsUnassignRequest.AddNew
                r_rsUnassignRequest!WG_NUM = !WG_NUM
                r_rsUnassignRequest!EMP_NUM = !EMP_NUM
                r_rsUnassignRequest!LINE_NUM = !LINE_NUM
                '//-----specifying the orginal request start and end date was too broad, and ended up trying to
                '//-----delete another employee on the same line (AdjustOverlapByLine)
                '//-----use the intersection rather than the orginal request-----//
                r_rsUnassignRequest!START_DATE = MaxDate(v_rsAssignmentRequest!START_DATE, !START_DATE)
                r_rsUnassignRequest!END_DATE = MinDate(v_rsAssignmentRequest!END_DATE, !END_DATE)
'//                r_rsUnassignRequest!START_DATE = v_rsAssignmentRequest!START_DATE
'//                r_rsUnassignRequest!END_DATE = v_rsAssignmentRequest!END_DATE
                r_rsUnassignRequest!KEY_WG_NUM = !WG_NUM
                r_rsUnassignRequest!KEY_EMP_NUM = !EMP_NUM
                r_rsUnassignRequest!KEY_LINE_NUM = !LINE_NUM
                r_rsUnassignRequest!KEY_START_DATE = v_rsAssignmentRequest!START_DATE
                r_rsUnassignRequest!KEY_END_DATE = v_rsAssignmentRequest!END_DATE
                If Not IsNull(v_rsAssignmentRequest!LAST_MOD_TIMESTAMP) Then
                    r_rsUnassignRequest!LAST_MOD_TIMESTAMP = v_rsAssignmentRequest!LAST_MOD_TIMESTAMP
                End If
                r_rsUnassignRequest!RECORD_STATUS = "D"

'                r_rsUnassignRequest!EMP_NUM = 0
'                r_rsUnassignRequest!LINE_NUM = v_rsOverlapInScope!LINE_NUM
'                r_rsUnassignRequest!START_DATE = v_rsAssignmentRequest!START_DATE
'                r_rsUnassignRequest!END_DATE = v_rsAssignmentRequest!END_DATE
'                r_rsUnassignRequest!RECORD_STATUS = "I"
                r_rsUnassignRequest.Update

                sBRVMsg = "Employee " & !EMP_NUM & " on line " & !LINE_NUM
                
                If v_rsAssignmentRequest!WG_NUM <> !WG_NUM Then
                    sBRVMsg = sBRVMsg & " in WG " & !WG_NUM
                End If
                
                sBRVMsg = sBRVMsg & " for " & !START_DATE & " and " & !END_DATE & " has been "
                
                If v_rsAssignmentRequest!START_DATE = !START_DATE And v_rsAssignmentRequest!END_DATE = !END_DATE Then
                    sBRVMsg = sBRVMsg & "unassigned." '//-----full unassign
                Else
                    sBRVMsg = sBRVMsg & "reassigned." '//-----partial unassign
                End If
                
                AppendBusinessError r_rsErrors, trErrbusValidationError, mstrCBUSINESSWAR, _
                    sBRVMsg, m_strCMODULENAME, 1, 1, "EMP_NUM", _
                    Array("KEY_WG_NUM", "KEY_LINE_NUM", "KEY_START_DATE", "KEY_END_DATE", "KEY_EMP_NUM"), _
                    CStr(v_rsAssignmentRequest!KEY_WG_NUM), _
                    CStr(v_rsAssignmentRequest!KEY_LINE_NUM), _
                    Format(v_rsAssignmentRequest!KEY_START_DATE, "mm/dd/yyyy"), _
                    Format(v_rsAssignmentRequest!KEY_END_DATE, "mm/dd/yyyy"), _
                    CStr(v_rsAssignmentRequest!KEY_EMP_NUM)

                .MoveNext
            Loop
            
        End With
    End If

    v_rsOverlapInScope.Filter = adFilterNone
    r_rsUnassignRequest.Filter = adFilterNone
      
    Exit Function
eh:
    GetObjectContext.SetAbort
    TrsRaiseError m_strCMODULENAME, "AdjustOverlapByEmp", , , , , True
End Function

'//-----the following modifies the data retrieved for validation to achieve the net effect-----//
Private Function ApplyPendingChanges(ByVal v_rsPendingChanges As ADODB.Recordset, _
                                    ByRef r_rsScopeData As ADODB.Recordset)
    Dim f As Field
    Dim l_strFilter As String

    On Error GoTo eh

    '//-----apply changes from overlap request to data retrieved from db for net effect-----//
    If RecordsExist(v_rsPendingChanges) Then
        With v_rsPendingChanges
            .Filter = MDAC25FilterFix("RECORD_STATUS = '" & m_strCPENDINGDELETE & "'") '//-----so this is applied once-----//
            If RecordsExist(v_rsPendingChanges) Then
                .MoveFirst
                Do While Not .EOF
                    ApplyPendingChange v_rsPendingChanges, r_rsScopeData
                    .MoveNext
                Loop
            End If
            
            .Filter = MDAC25FilterFix("RECORD_STATUS = '" & m_strCPENDINGINSERT & "'") '//-----so this is applied once-----//
            If RecordsExist(v_rsPendingChanges) Then
                .MoveFirst
                Do While Not .EOF
                    ApplyPendingChange v_rsPendingChanges, r_rsScopeData
                    .MoveNext
                Loop
            End If
            
            '//-----don't expect this code to execute, because changes made to key fields produce a delete/insert
            '//-----will get here if only the RELIEF_COMP changes but the keys remain the same
            .Filter = MDAC25FilterFix("RECORD_STATUS = '" & m_strCPENDINGCHANGE & "'") '//-----so this is applied once-----//
            If RecordsExist(v_rsPendingChanges) Then
                .MoveFirst
                Do While Not .EOF
                    ApplyPendingChange v_rsPendingChanges, r_rsScopeData
                    .MoveNext
                Loop
            End If
        End With
    End If
    
    'TEST FOR JUAN
    r_rsScopeData.Filter = adFilterNone
    v_rsPendingChanges.Filter = adFilterNone
    Exit Function
eh:
    GetObjectContext.SetAbort
    TrsRaiseError m_strCMODULENAME, "ApplyPendingChanges", , , , , True
End Function

Private Function ApplyPendingChange(ByRef r_rsPendingChange As ADODB.Recordset, _
                                    ByRef r_rsScopeData As ADODB.Recordset)
    Dim f As Field
    Dim l_strFilter As String

    On Error GoTo eh

    '//-----apply changes from current row of request (overlap or assignment) to data retrieved from db for net effect-----//
    If RecordsExist(r_rsPendingChange) Then
        With r_rsPendingChange
            Select Case !RECORD_STATUS
            Case m_strCPENDINGDELETE, "D"
                l_strFilter = "WG_NUM = " & !KEY_WG_NUM & _
                            " AND START_DATE = #" & !KEY_START_DATE & "#" & _
                            " AND END_DATE = #" & !KEY_END_DATE & "#" & _
                            " AND EMP_NUM = " & !KEY_EMP_NUM & _
                            " AND LINE_NUM = " & !KEY_LINE_NUM
                With r_rsScopeData
                    .Filter = MDAC25FilterFix(l_strFilter)
                    If RecordsExist(r_rsScopeData) Then
                        .MoveFirst
                        Do While Not .EOF
                            .Delete
                            .Update
                            .MoveNext
                        Loop
                    End If
                End With
                !RECORD_STATUS = "D" '//-----so this is applied once-----//
                .Update
            Case m_strCPENDINGINSERT, "I"
                '//-----fields in both recordsets need to match-----//
                r_rsScopeData.AddNew
                CopyRequestFields r_rsPendingChange, r_rsScopeData, "I"
                r_rsScopeData.Update
                !RECORD_STATUS = "I" '//-----so this is applied once-----//
                .Update
            '//-----don't expect this code to execute, because changes made to key fields produce a delete/insert
            Case m_strCPENDINGCHANGE, "C"
                l_strFilter = "WG_NUM = " & !KEY_WG_NUM & _
                            " AND START_DATE = #" & !KEY_START_DATE & "#" & _
                            " AND END_DATE = #" & !KEY_END_DATE & "#" & _
                            " AND EMP_NUM = " & !KEY_EMP_NUM & _
                            " AND LINE_NUM = " & !KEY_LINE_NUM
                r_rsScopeData.Filter = MDAC25FilterFix(l_strFilter)
                If RecordsExist(r_rsScopeData) Then
                    CopyRequestFields r_rsPendingChange, r_rsScopeData, "C"
                    r_rsScopeData.Update
                End If
                !RECORD_STATUS = "C" '//-----so this is applied once-----//
                .Update
            End Select
        End With
    End If
    
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
        If Not IsNull(rsSource!RELIEF_COMP) Then
            !RELIEF_COMP = rsSource!RELIEF_COMP
        End If
        Select Case v_strRecordStatus
        Case "I", m_strCPENDINGINSERT
            !WG_NUM = rsSource!WG_NUM
            !EMP_NUM = rsSource!EMP_NUM
            !LINE_NUM = rsSource!LINE_NUM
            !START_DATE = CDate(rsSource!START_DATE)
            !END_DATE = CDate(rsSource!END_DATE)

            '//-----even though the non-key fields are only defined, the key fields will be needed later-----//
            !KEY_WG_NUM = rsSource!WG_NUM
            !KEY_EMP_NUM = rsSource!EMP_NUM
            !KEY_LINE_NUM = rsSource!LINE_NUM
            !KEY_START_DATE = CDate(rsSource!START_DATE)
            !KEY_END_DATE = CDate(rsSource!END_DATE)
        Case "C", m_strCPENDINGCHANGE
            !WG_NUM = rsSource!WG_NUM
            !EMP_NUM = rsSource!EMP_NUM
            !LINE_NUM = rsSource!LINE_NUM
            !START_DATE = CDate(rsSource!START_DATE)
            !END_DATE = CDate(rsSource!END_DATE)
            
            !KEY_WG_NUM = rsSource!KEY_WG_NUM
            !KEY_EMP_NUM = rsSource!KEY_EMP_NUM
            !KEY_LINE_NUM = rsSource!KEY_LINE_NUM
            !KEY_START_DATE = CDate(rsSource!KEY_START_DATE)
            !KEY_END_DATE = CDate(rsSource!KEY_END_DATE)
        Case "D", m_strCPENDINGDELETE
            !WG_NUM = rsSource!KEY_WG_NUM
            !EMP_NUM = rsSource!KEY_EMP_NUM
            !LINE_NUM = rsSource!KEY_LINE_NUM
            !START_DATE = CDate(rsSource!KEY_START_DATE)
            !END_DATE = CDate(rsSource!KEY_END_DATE)
            
            !KEY_WG_NUM = rsSource!KEY_WG_NUM
            !KEY_EMP_NUM = rsSource!KEY_EMP_NUM
            !KEY_LINE_NUM = rsSource!KEY_LINE_NUM
            !KEY_START_DATE = CDate(rsSource!KEY_START_DATE)
            !KEY_END_DATE = CDate(rsSource!KEY_END_DATE)
        End Select
        
        !RECORD_STATUS = v_strRecordStatus
    
    End With
    CopyRequestFields = True
    Exit Function
eh:
    CopyRequestFields = False
    TrsRaiseError m_strCMODULENAME, "CopyRequestFields", , , , , True
End Function

'//-----Adjust start and end dates for assignments that would overlap the request-----//
Private Function AdjustOverlapByLine(ByVal v_rsAssignmentRequest As ADODB.Recordset, _
                                    ByVal v_rsOverlapInScope As ADODB.Recordset, _
                                    ByRef r_rsOverlapRequest As ADODB.Recordset, _
                                    ByRef r_rsErrors As ADODB.Recordset, _
                                    ByVal v_lngOriginalRequestWGNum As Long) As Boolean
    Dim l_datSplitStartDate1 As Date
    Dim l_datSplitEndDate1 As Date
    Dim l_datSplitStartDate2 As Date
    Dim l_datSplitEndDate2 As Date
    Dim l_strRecordStatus1 As String
    Dim l_strRecordStatus2 As String
    Dim l_strFilter As String
    
    On Error GoTo eh

    If RecordsExist(v_rsAssignmentRequest) Then
        With v_rsAssignmentRequest
            '//-----locate matching record in scope overlap-----//
            l_strFilter = "WG_NUM = " & !WG_NUM & _
            " AND START_DATE <= #" & !END_DATE & "#" & _
            " AND END_DATE >= #" & !START_DATE & "#" & _
            " AND LINE_NUM = " & !LINE_NUM
            
            v_rsOverlapInScope.Filter = MDAC25FilterFix(l_strFilter)
        End With
    Else
        AdjustOverlapByLine = False
        Exit Function
    End If
        
    If RecordsExist(v_rsOverlapInScope) Then
        v_rsOverlapInScope.MoveFirst
        Do While Not v_rsOverlapInScope.EOF
            '//-----the long format for verifying ALL test cases...-----//
            l_datSplitStartDate2 = m_datCOPTIONALSTARTDATE
            l_datSplitEndDate2 = m_datCOPTIONALENDDATE
            
            If v_rsAssignmentRequest!START_DATE < v_rsOverlapInScope!START_DATE Then
                If v_rsAssignmentRequest!END_DATE < v_rsOverlapInScope!END_DATE Then
                    l_datSplitStartDate1 = DateAdd("d", 1, v_rsAssignmentRequest!END_DATE)
                    l_datSplitEndDate1 = v_rsOverlapInScope!END_DATE
                    l_strRecordStatus1 = "C"
                    'append cascade (1)
                    AppendCascadeDates l_strRecordStatus1, _
                        v_rsOverlapInScope!EMP_NUM, _
                        v_rsOverlapInScope!START_DATE, _
                        v_rsAssignmentRequest!END_DATE
                ElseIf v_rsAssignmentRequest!END_DATE = v_rsOverlapInScope!END_DATE Then
                    l_datSplitStartDate1 = v_rsOverlapInScope!START_DATE
                    l_datSplitEndDate1 = v_rsOverlapInScope!END_DATE
                    l_strRecordStatus1 = "D"
                    'append cascade (4)
                    AppendCascadeDates l_strRecordStatus1, _
                        v_rsOverlapInScope!EMP_NUM, _
                        v_rsOverlapInScope!START_DATE, _
                        v_rsOverlapInScope!END_DATE
                ElseIf v_rsAssignmentRequest!END_DATE > v_rsOverlapInScope!END_DATE Then
                    l_datSplitStartDate1 = v_rsOverlapInScope!START_DATE
                    l_datSplitEndDate1 = v_rsOverlapInScope!END_DATE
                    l_strRecordStatus1 = "D"
                    'append cascade (7, same as 4)
                    AppendCascadeDates l_strRecordStatus1, _
                        v_rsOverlapInScope!EMP_NUM, _
                        v_rsOverlapInScope!START_DATE, _
                        v_rsOverlapInScope!END_DATE
                End If
            ElseIf v_rsAssignmentRequest!START_DATE = v_rsOverlapInScope!START_DATE Then
                If v_rsAssignmentRequest!END_DATE < v_rsOverlapInScope!END_DATE Then
                    l_datSplitStartDate1 = DateAdd("d", 1, v_rsAssignmentRequest!END_DATE)
                    l_datSplitEndDate1 = v_rsOverlapInScope!END_DATE
                    l_strRecordStatus1 = "C"
                    'append cascade (2)
                    AppendCascadeDates l_strRecordStatus1, _
                        v_rsOverlapInScope!EMP_NUM, _
                        v_rsOverlapInScope!START_DATE, _
                        v_rsAssignmentRequest!END_DATE
                ElseIf v_rsAssignmentRequest!END_DATE = v_rsOverlapInScope!END_DATE Then
                    l_datSplitStartDate1 = v_rsOverlapInScope!START_DATE
                    l_datSplitEndDate1 = v_rsOverlapInScope!END_DATE
                    l_strRecordStatus1 = "D"
                    'append cascade (5)
                    AppendCascadeDates l_strRecordStatus1, _
                        v_rsOverlapInScope!EMP_NUM, _
                        v_rsOverlapInScope!START_DATE, _
                        v_rsOverlapInScope!END_DATE
                ElseIf v_rsAssignmentRequest!END_DATE > v_rsOverlapInScope!END_DATE Then
                    l_datSplitStartDate1 = v_rsOverlapInScope!START_DATE
                    l_datSplitEndDate1 = v_rsOverlapInScope!END_DATE
                    l_strRecordStatus1 = "D"
                    'append cascade (8, same as 5)
                    AppendCascadeDates l_strRecordStatus1, _
                        v_rsOverlapInScope!EMP_NUM, _
                        v_rsOverlapInScope!START_DATE, _
                        v_rsOverlapInScope!END_DATE
                End If
            ElseIf v_rsAssignmentRequest!START_DATE > v_rsOverlapInScope!START_DATE Then
                If v_rsAssignmentRequest!END_DATE < v_rsOverlapInScope!END_DATE Then
                    l_datSplitStartDate1 = v_rsOverlapInScope!START_DATE
                    l_datSplitEndDate1 = DateAdd("d", -1, v_rsAssignmentRequest!START_DATE)
                    l_strRecordStatus1 = "C"
                    l_datSplitStartDate2 = DateAdd("d", 1, v_rsAssignmentRequest!END_DATE)
                    l_datSplitEndDate2 = v_rsOverlapInScope!END_DATE
                    l_strRecordStatus2 = "I"
                    'append cascade (3)
                    AppendCascadeDates l_strRecordStatus1, _
                        v_rsOverlapInScope!EMP_NUM, _
                        v_rsAssignmentRequest!START_DATE, _
                        v_rsAssignmentRequest!END_DATE
                ElseIf v_rsAssignmentRequest!END_DATE = v_rsOverlapInScope!END_DATE Then
                    l_datSplitStartDate1 = v_rsOverlapInScope!START_DATE
                    l_datSplitEndDate1 = DateAdd("d", -1, v_rsAssignmentRequest!START_DATE)
                    l_strRecordStatus1 = "C"
                    'append cascade (6)
                    AppendCascadeDates l_strRecordStatus1, _
                        v_rsOverlapInScope!EMP_NUM, _
                        v_rsAssignmentRequest!START_DATE, _
                        v_rsOverlapInScope!END_DATE
                ElseIf v_rsAssignmentRequest!END_DATE > v_rsOverlapInScope!END_DATE Then
                    l_datSplitStartDate1 = v_rsOverlapInScope!START_DATE
                    l_datSplitEndDate1 = DateAdd("d", -1, v_rsAssignmentRequest!START_DATE)
                    l_strRecordStatus1 = "C"
                    'append cascade (9, same as 6)
                    AppendCascadeDates l_strRecordStatus1, _
                        v_rsOverlapInScope!EMP_NUM, _
                        v_rsAssignmentRequest!START_DATE, _
                        v_rsOverlapInScope!END_DATE
                End If
            End If
            
            '//-----if request is a change and the overlap matches (keys), don't delete overlap, because it will be needed for the update
            If Not (v_rsAssignmentRequest!RECORD_STATUS = "C" And _
                v_rsAssignmentRequest!KEY_WG_NUM = v_rsOverlapInScope!KEY_WG_NUM And _
                v_rsAssignmentRequest!KEY_LINE_NUM = v_rsOverlapInScope!KEY_LINE_NUM And _
                v_rsAssignmentRequest!KEY_START_DATE = v_rsOverlapInScope!KEY_START_DATE And _
                v_rsAssignmentRequest!KEY_END_DATE = v_rsOverlapInScope!KEY_END_DATE And _
                v_rsAssignmentRequest!KEY_EMP_NUM = v_rsOverlapInScope!KEY_EMP_NUM) Then
                                
                CreateOverlapRequest v_rsAssignmentRequest, v_rsOverlapInScope, _
                        l_datSplitStartDate1, l_datSplitEndDate1, l_datSplitStartDate2, l_datSplitEndDate2, _
                        l_strRecordStatus1, l_strRecordStatus2, r_rsOverlapRequest, r_rsErrors, v_lngOriginalRequestWGNum
            End If

            v_rsOverlapInScope.MoveNext
        Loop
    End If
    
    v_rsOverlapInScope.Filter = adFilterNone
    
    AdjustOverlapByLine = True
    Exit Function
eh:
    AdjustOverlapByLine = False
    GetObjectContext.SetAbort
    TrsRaiseError m_strCMODULENAME, "AdjustOverlapByLine", , , , , True
End Function

Private Function CreateOverlapRequest(ByVal v_rsAssignmentRequest As ADODB.Recordset, _
                                    ByVal v_rsOverlapInScope As ADODB.Recordset, _
                                    ByVal v_datSplitStartDate1 As Date, _
                                    ByVal v_datSplitEndDate1 As Date, _
                                    ByVal v_datSplitStartDate2 As Date, _
                                    ByVal v_datSplitEndDate2 As Date, _
                                    ByVal v_strRecordStatus1 As String, _
                                    ByVal v_strRecordStatus2 As String, _
                                    ByRef r_rsOverlapRequest As ADODB.Recordset, _
                                    ByRef r_rsErrors As ADODB.Recordset, _
                                    ByVal v_lngOriginalRequestWGNum As Long) As Boolean
    Dim sBRVMsg As String
    
    On Error GoTo eh
    
    If RecordsExist(v_rsOverlapInScope) Then
        If v_strRecordStatus1 = "C" Then '//-----do a simpler ProcessKeyChange-----//
            CreateOverlapRequestRow v_rsOverlapInScope, v_rsOverlapInScope!START_DATE, v_rsOverlapInScope!END_DATE, "D", r_rsOverlapRequest
            CreateOverlapRequestRow v_rsOverlapInScope, v_datSplitStartDate1, v_datSplitEndDate1, "I", r_rsOverlapRequest
        Else
            CreateOverlapRequestRow v_rsOverlapInScope, v_datSplitStartDate1, v_datSplitEndDate1, v_strRecordStatus1, r_rsOverlapRequest
        End If
        
        If v_strRecordStatus2 <> "" Then
            CreateOverlapRequestRow v_rsOverlapInScope, v_datSplitStartDate2, v_datSplitEndDate2, v_strRecordStatus2, r_rsOverlapRequest
        End If
                    
        sBRVMsg = "Overlap with employee " & v_rsOverlapInScope!KEY_EMP_NUM
        sBRVMsg = sBRVMsg & " on line " & v_rsOverlapInScope!KEY_LINE_NUM
        If v_lngOriginalRequestWGNum <> v_rsOverlapInScope!WG_NUM Then
            sBRVMsg = sBRVMsg & " in WG " & v_rsOverlapInScope!KEY_WG_NUM
        End If
        
        sBRVMsg = sBRVMsg & " for " & v_rsOverlapInScope!START_DATE & " and " & v_rsOverlapInScope!END_DATE
'//                    r_rsOverlapRequest!KEY_START_DATE & " and " & r_rsOverlapRequest!KEY_END_DATE
        
        If v_strRecordStatus2 <> "" Then
            sBRVMsg = sBRVMsg & " was split to " & v_datSplitStartDate1 & " through " & v_datSplitEndDate1 & _
                                        " and " & v_datSplitStartDate2 & " through " & v_datSplitEndDate2
        Else
            If v_strRecordStatus1 = "C" Then
                sBRVMsg = sBRVMsg & " was split to " & v_datSplitStartDate1 & " through " & v_datSplitEndDate1
            ElseIf v_strRecordStatus1 = "D" Then
                sBRVMsg = sBRVMsg & " was deleted."
            End If
        End If
    
        AppendBusinessError r_rsErrors, trErrbusValidationError, mstrCBUSINESSWAR, _
            sBRVMsg, m_strCMODULENAME, _
            1, 1, "START_DATE", _
            Array("KEY_WG_NUM", "KEY_LINE_NUM", "KEY_START_DATE", "KEY_END_DATE", "KEY_EMP_NUM"), _
            CStr(v_rsAssignmentRequest!WG_NUM), _
            CStr(v_rsAssignmentRequest!LINE_NUM), _
            Format(v_rsAssignmentRequest!START_DATE, "mm/dd/yyyy"), _
            Format(v_rsAssignmentRequest!END_DATE, "mm/dd/yyyy"), _
            CStr(v_rsAssignmentRequest!EMP_NUM)
    End If
        
    CreateOverlapRequest = True
    Exit Function
eh:
    CreateOverlapRequest = False
    GetObjectContext.SetAbort
    TrsRaiseError m_strCMODULENAME, "CreateOverlapRequest", , , , , True
End Function

Private Function CreateOverlapRequestRow(ByVal v_rsOverlapInScope As ADODB.Recordset, _
                                    ByVal v_datSplitStartDate As Date, _
                                    ByVal v_datSplitEndDate As Date, _
                                    ByVal v_strRecordStatus As String, _
                                    ByRef r_rsOverlapRequest As ADODB.Recordset) As Boolean
    Dim l_strPendingRecordStatus As String
    
    On Error GoTo eh
    
    With r_rsOverlapRequest
        .AddNew
        
        '//-----use pending record status, so that the rows are applied once in ApplyPendingChanges-----//
        '//-----change these in one place, could have used the constants in AdjustOverlapByLine & CreateOverlapRequest-----//
        Select Case v_strRecordStatus
        Case "D"
            l_strPendingRecordStatus = m_strCPENDINGDELETE
        Case "I"
            l_strPendingRecordStatus = m_strCPENDINGINSERT
        Case "C"
            l_strPendingRecordStatus = m_strCPENDINGCHANGE
        Case Else
            l_strPendingRecordStatus = v_strRecordStatus
        End Select
        
        CopyRequestFields v_rsOverlapInScope, r_rsOverlapRequest, l_strPendingRecordStatus
        
        If l_strPendingRecordStatus <> m_strCPENDINGDELETE Then
            !START_DATE = v_datSplitStartDate
            !END_DATE = v_datSplitEndDate
            If l_strPendingRecordStatus = m_strCPENDINGINSERT Then
                !KEY_START_DATE = v_datSplitStartDate
                !KEY_END_DATE = v_datSplitEndDate
            End If
        End If
        
        .Update
    End With
        
    CreateOverlapRequestRow = True
    Exit Function
eh:
    CreateOverlapRequestRow = False
    GetObjectContext.SetAbort
    TrsRaiseError m_strCMODULENAME, "CreateOverlapRequestRow", , , , , True
End Function

'//-----Problem: 12/27/2001 The overlap requests are based on processing order, but the db component is based on record status order (DIC)-----//
'//-----    If a block of time gets broken for an overlap more than once, it could generate multiple delete requests
'//-----        which will be executed before the inserts and mess things up
'//-----Solution: for each delete request check for a corresponding insert and remove both-----//
'//-----            Otherwise, send 2 requests rs to the SaveAssignments method. 1 for client requests and the other for overlap requests

Private Function ApplyNetChanges(ByRef r_rsOverlapRequest As ADODB.Recordset) As Boolean
    Dim l_strFilter As String
    Dim l_rsClone As ADODB.Recordset
    
    On Error GoTo eh

    l_strFilter = "RECORD_STATUS = 'D'"
    r_rsOverlapRequest.Filter = MDAC25FilterFix(l_strFilter)
    If RecordsExist(r_rsOverlapRequest) Then
        With r_rsOverlapRequest
            Set l_rsClone = .Clone
            .MoveFirst
            Do While Not .EOF
                l_strFilter = "WG_NUM = " & !KEY_WG_NUM & _
                            " AND START_DATE = #" & !KEY_START_DATE & "#" & _
                            " AND END_DATE = #" & !KEY_END_DATE & "#" & _
                            " AND EMP_NUM = " & !KEY_EMP_NUM & _
                            " AND LINE_NUM = " & !KEY_LINE_NUM & _
                            " AND RECORD_STATUS = 'I'"
                l_rsClone.Filter = MDAC25FilterFix(l_strFilter)
                If RecordsExist(l_rsClone) Then
                    .Delete
                    .Update
                    l_rsClone.Delete
                    l_rsClone.Update
                End If
                
                .MoveNext
            Loop
            l_rsClone.Filter = adFilterNone
            Set l_rsClone = Nothing
        End With
    End If
    
    r_rsOverlapRequest.Filter = adFilterNone

    GetObjectContext.SetComplete
    ApplyNetChanges = True
    Exit Function
eh:
    Set l_rsClone = Nothing
    ApplyNetChanges = False
    GetObjectContext.SetAbort
    TrsRaiseError m_strCMODULENAME, "ApplyNetChanges", , , , , True
End Function


'//-----Make sure that a line exists (ttr_wg_line_sched) for each day in date range-----//
Private Function DoLinesExist(ByVal v_lngWgNum As Long, ByVal v_lngLineNum As Long, _
                                ByVal v_datStartDate As Date, ByVal v_datEndDate As Date, _
                                ByRef r_strDatesMissing As String) As Boolean
    Dim l_objRS As ADODB.Recordset
    Dim l_colDatesMissing As Collection
    Dim l_objDBSchDmd As dbSchedDemand.cdSchedDemand
    
    On Error GoTo eh

    Set l_objDBSchDmd = CreateObject("dbSchedDemand.cdSchedDemand")
    Set l_objRS = l_objDBSchDmd.GetLines(v_lngWgNum, v_datStartDate, v_datEndDate, v_lngLineNum)
    Set l_objDBSchDmd = Nothing
            
    Set l_colDatesMissing = GetMissingDates(l_objRS, v_datStartDate, v_datEndDate)
    r_strDatesMissing = DateRangeCol2String(l_colDatesMissing)
    
    '//-----no missing dates means that all lines exist (true)-----//
    DoLinesExist = (r_strDatesMissing = "")

    If DoLinesExist Then
        GetObjectContext.SetComplete
    Else
        GetObjectContext.SetAbort
    End If
    
    Set l_objRS = Nothing
    Set l_objDBSchDmd = Nothing
    Set l_colDatesMissing = Nothing
    Exit Function
eh:
    Set l_objRS = Nothing
    Set l_objDBSchDmd = Nothing
    Set l_colDatesMissing = Nothing
    GetObjectContext.SetAbort
    TrsRaiseError m_strCMODULENAME, "DoLinesExist", , , , , True
End Function

'//-----Make sure that a line exists (ttr_wg_line_sched) for each day in date range-----//
Private Function DoLinesExistBulk(ByVal v_rsLinesBulk As ADODB.Recordset, _
                            ByVal v_lngWgNum As Long, ByVal v_lngLineNum As Long, _
                            ByVal v_datStartDate As Date, ByVal v_datEndDate As Date, _
                            ByRef r_strDatesMissing As String) As Boolean
    Dim l_colDatesMissing As Collection
    Dim l_strFilter As String
    
    On Error GoTo eh

    l_strFilter = "WG_NUM = " & v_lngWgNum & _
                " AND LINE_NUM = " & v_lngLineNum & _
                " AND START_DATE <= #" & v_datEndDate & "#" & _
                " AND END_DATE >= #" & v_datStartDate & "#"
    
    v_rsLinesBulk.Filter = MDAC25FilterFix(l_strFilter)

    Set l_colDatesMissing = GetMissingDates(v_rsLinesBulk, v_datStartDate, v_datEndDate)
    r_strDatesMissing = DateRangeCol2String(l_colDatesMissing)
    
    '//-----no missing dates means that all lines exist (true)-----//
    DoLinesExistBulk = (r_strDatesMissing = "")

    If DoLinesExistBulk Then
        GetObjectContext.SetComplete
    Else
        GetObjectContext.SetAbort
    End If
    
    Set l_colDatesMissing = Nothing
    Exit Function
eh:
    Set l_colDatesMissing = Nothing
    GetObjectContext.SetAbort
    TrsRaiseError m_strCMODULENAME, "DoLinesExistBulk", , , , , True
End Function


'//-----Given StartDate = '6/01/1996' and EndDate = '6/30/1996'
'WG_NUM  LINE_NUM    START_DATE  END_DATE
'3043    1   05/26/1996  06/01/1996
'3043    1   06/03/1996  06/08/1996
'3043    1   06/09/1996  06/14/1996
'3043    1   06/18/1996  06/22/1996
'3043    1   06/23/1996  06/29/1996
'3043    1   07/01/1996  07/06/1996
'//-----return {{6/2/1996, 6/02/1996}, {6/15/1996, 6/17/1996}, {6/30/1996, 6/30/1996}} (as a collection of arrays)

Private Function GetMissingDates(ByVal v_rsValidation As ADODB.Recordset, _
                                ByVal v_datStartDate As Date, _
                                ByVal v_datEndDate As Date) As Collection
    
    Dim l_datTemp As Date
    Dim aryDateRange(0 To 1) As Date
    Dim l_colDateRange As Collection

    On Error GoTo eh
    
    Set l_colDateRange = New Collection
    
    With v_rsValidation
        If RecordsExist(v_rsValidation) Then
            .MoveFirst
            l_datTemp = v_datStartDate
            
            Do While l_datTemp <= v_datEndDate And Not .EOF
                If l_datTemp < !START_DATE Then '//-----found gap-----//
                    aryDateRange(0) = l_datTemp
                    aryDateRange(1) = DateAdd("d", -1, !START_DATE)
                    l_colDateRange.Add aryDateRange
                End If
            
                If !END_DATE < m_datCOPTIONALENDDATE Then '//-----cannot add 1 to 12/31/9999-----//
                    l_datTemp = DateAdd("d", 1, !END_DATE)
                Else
                    l_datTemp = m_datCOPTIONALENDDATE
                End If
                
                .MoveNext
            Loop
            
            If l_datTemp <= v_datEndDate And l_datTemp < m_datCOPTIONALENDDATE Then '//-----found gap-----//
                aryDateRange(0) = l_datTemp
                aryDateRange(1) = v_datEndDate
                l_colDateRange.Add aryDateRange
            End If
        Else
            aryDateRange(0) = v_datStartDate
            aryDateRange(1) = v_datEndDate
            l_colDateRange.Add aryDateRange
        End If
    End With

    Set GetMissingDates = l_colDateRange
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "GetMissingDates", , , , , True
End Function

'//-----Given
'WG_NUM  LINE_NUM    START_DATE  END_DATE    EMP_NUM RELIEF_COMP
'3043    1          06/02/1996  06/02/1996  18087
'3043    1          06/03/1996  06/03/1996  18087
'3043    1          06/04/1996  06/04/1996  18087
'3043    1          06/05/1996  06/05/1996  18087
'3043    1          06/06/1996  06/06/1996  18087
'3043    1          06/25/1996  06/25/1996  0
'3043    1          06/26/1996  06/26/1996  0
'3043    1          06/27/1996  06/27/1996  0
'3043    1          06/28/1996  06/28/1996  0
'//-----return {{6/2/1996, 6/06/1996}, {6/25/1996, 6/28/1996}} (as a collection of arrays)

Private Function GetContiguousDates(ByVal v_rsValidation As ADODB.Recordset) As Collection
    Dim l_datStart As Date, l_datEnd As Date
    Dim l_strDates As String
    Dim aryDateRange(0 To 1) As Date
    Dim l_colDateRange As Collection

    On Error GoTo eh
    
    Set l_colDateRange = New Collection

    With v_rsValidation
        If RecordsExist(v_rsValidation) Then
            .Sort = "START_DATE"
            
            .MoveFirst
            l_datStart = !START_DATE
            l_datEnd = !END_DATE
            .MoveNext
            
            Do While Not .EOF
                If !START_DATE > DateAdd("d", 1, l_datEnd) Then  '//-----found gap-----//
                    aryDateRange(0) = l_datStart
                    aryDateRange(1) = l_datEnd
                    l_colDateRange.Add aryDateRange
                    l_datStart = !START_DATE
                End If
                
                l_datEnd = !END_DATE
                .MoveNext
            Loop
            '//-----flush out the last one-----//
            aryDateRange(0) = l_datStart
            aryDateRange(1) = l_datEnd
            l_colDateRange.Add aryDateRange
        End If
    End With

    Set GetContiguousDates = l_colDateRange
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "GetContiguousDates", , , , , True
End Function

''//-----Line Date   Emp     RecordStatus
''Current data
''       1   11/12   19017
''       2   11/12   17723
''Request to swap employees
''       1   11/12   17723   C
''       2   11/12   19017   C
''ProcessKeyChange yields
''       1   11/12   19017   D
''       2   11/12   17723   D
''       1   11/12   17723   I
''       2   11/12   19017   I

Private Function ProcessKeyChange(ByRef r_rsRequest As ADODB.Recordset, _
                    Optional ByVal v_varKeyColumns As Variant, Optional ByVal v_varOriginalKeyColumns As Variant) As Boolean
    Dim lngKey As Long ', f As Field
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

        AppendCascadeDeleteInsert r_rsRequest

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
                                '//-----Add "D" row and Change "C" to "I"-----//
                                AppendCascadeChange r_rsRequest
                                
                                rsAppend.AddNew
                                CopyRequestFields r_rsRequest, rsAppend, "I"
                                rsAppend.Update
                                '//-----
                                CopyRequestFields r_rsRequest, r_rsRequest, "D" '//-----so that key fields will match non key fields-----//
                                '//!RECORD_STATUS = "D"
                                .Update
                                
'                                rsAppend.AddNew
'                                CopyRequestFields r_rsRequest, rsAppend, "D"
'                                '//rsAppend!RECORD_STATUS = "D"
'                                rsAppend.Update
'                                '//-----
'                                !RECORD_STATUS = "I"
'                                .Update
                                
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

Private Function dbFormat(ByVal v_varData As Variant) As String
    Dim l_strFormated As String, fieldType As Integer
    
    On Error GoTo eh
        
    fieldType = VarType(v_varData)
    Select Case fieldType
        Case vbString
            l_strFormated = "'" & Replace(v_varData, "'", "''") & "'"
        Case vbInteger, vbLong
            l_strFormated = CStr(v_varData)
        Case vbDate
            l_strFormated = "'" & Format(v_varData, "mm/dd/yyyy") & "'"
        Case Else
            l_strFormated = Format(v_varData)
    End Select
    dbFormat = l_strFormated
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "dbFormat", , , , , True
End Function

Private Function DateRangeCol2String(l_colDateRange As Collection) As String
    Dim l_strDateRange As String
    Dim aryDateRange As Variant '//Dim aryDateRange(0 To 1) As Date
    Dim xx As Long
    Dim l_blnMultipleDates As Boolean
    
    On Error GoTo eh
    
    l_strDateRange = ""
    
    xx = 1
    For Each aryDateRange In l_colDateRange
        If aryDateRange(0) = aryDateRange(1) Then '//-----suppress range into single value-----//
            l_strDateRange = l_strDateRange & aryDateRange(0)
        Else
            l_strDateRange = l_strDateRange & aryDateRange(0) & " through " & aryDateRange(1)
            l_blnMultipleDates = True
        End If
                
        If l_colDateRange.Count = 2 Then
            If xx = l_colDateRange.Count - 1 Then '//-----next to last one in collection-----//
                l_strDateRange = l_strDateRange & " and "
            End If
            l_blnMultipleDates = True
        ElseIf l_colDateRange.Count > 2 Then
            If xx < l_colDateRange.Count - 1 Then '//-----not the last one in collection-----//
                l_strDateRange = l_strDateRange & ", "
            End If

            If xx = l_colDateRange.Count - 1 Then '//-----next to last one in collection-----//
                l_strDateRange = l_strDateRange & ", and "
            End If
            l_blnMultipleDates = True
        End If
        
        xx = xx + 1
    Next
    If l_strDateRange <> "" Then
        If l_blnMultipleDates Then
            l_strDateRange = "dates " & l_strDateRange
        Else
            l_strDateRange = "date " & l_strDateRange
        End If
    End If
    
    DateRangeCol2String = l_strDateRange

Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "DateRangeArray2String", , , , , True
End Function

Private Function ValidateLines(ByVal v_rsAssignmentRequest As ADODB.Recordset, ByRef r_rsErrors As ADODB.Recordset) As Boolean
    Dim l_strDatesMissing As String
    Dim bError As Boolean
    
    On Error GoTo eh

    With v_rsAssignmentRequest
        .MoveFirst
        Do While Not .EOF
            If !EMP_NUM > 0 Then
                If Not DoLinesExist(!WG_NUM, !START_DATE, !END_DATE, !LINE_NUM, l_strDatesMissing) Then
                    bError = True
            
                    AppendBusinessError r_rsErrors, trErrbusValidationError, mstrCBUSINESSERR, _
                        "Line number " & !LINE_NUM & " must exist for date(s) between " & l_strDatesMissing & "." _
                        , m_strCMODULENAME, _
                        1, 1, "LINE_NUM", _
                        Array("WG_NUM", "LINE_NUM", "START_DATE", "END_DATE", "EMP_NUM"), _
                        CStr(!WG_NUM), _
                        CStr(!LINE_NUM), _
                        Format(!START_DATE, "mm/dd/yyyy"), _
                        Format(!END_DATE, "mm/dd/yyyy"), _
                        CStr(!EMP_NUM)
                End If
            End If
            .MoveNext
        Loop
    End With
    
    ValidateLines = Not bError
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "ValidateLines", , , , , True
End Function
    
Private Function ValidateLinesBulk(ByVal v_rsAssignmentRequest As ADODB.Recordset, _
                                    ByVal v_rsLinesBulk As ADODB.Recordset, _
                                    ByRef r_rsErrors As ADODB.Recordset) As Boolean
    Dim l_strDatesMissing As String
    Dim bError As Boolean
    
    On Error GoTo eh

    With v_rsAssignmentRequest
        .MoveFirst
        Do While Not .EOF
            If !EMP_NUM > 0 Then
                If Not DoLinesExistBulk(v_rsLinesBulk, !WG_NUM, !START_DATE, !END_DATE, !LINE_NUM, l_strDatesMissing) Then
                    bError = True
            
                    AppendBusinessError r_rsErrors, trErrbusValidationError, mstrCBUSINESSERR, _
                        "Line number " & !LINE_NUM & " must exist for date(s) between " & l_strDatesMissing & "." _
                        , m_strCMODULENAME, _
                        1, 1, "LINE_NUM", _
                        Array("WG_NUM", "LINE_NUM", "START_DATE", "END_DATE", "EMP_NUM"), _
                        CStr(!WG_NUM), _
                        CStr(!LINE_NUM), _
                        Format(!START_DATE, "mm/dd/yyyy"), _
                        Format(!END_DATE, "mm/dd/yyyy"), _
                        CStr(!EMP_NUM)
                End If
            End If
            .MoveNext
        Loop
    End With
    
    ValidateLinesBulk = Not bError
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "ValidateLinesBulk", , , , , True
End Function


'//-----find minimum and maximum values in order to retrieve all scope for validation-----//
Private Function GetLinesBulk(ByVal v_rsAssignmentRequest As ADODB.Recordset) As ADODB.Recordset
    Dim l_rsAssignmentRequestClone As ADODB.Recordset
    Dim l_rsLinesBulk As ADODB.Recordset
    Dim l_objDBSchDmd As dbSchedDemand.cdSchedDemand
    Dim l_datMinStartDate As Date, l_datMaxEndDate As Date
    Dim l_lngWGNumSave As Long, l_lngLineNumSave As Long, l_lngEmpNumSave As Long
    Dim l_colScopeParm As New Collection
'    Dim l_tScopeParm As tScopeParm
    Dim l_tScopeParm(4) As Variant
    Dim l_strFilter As String
    
    On Error GoTo eh
                    
    Set l_rsAssignmentRequestClone = v_rsAssignmentRequest.Clone
    
    With l_rsAssignmentRequestClone
        '//-----build list of necessary lines-----//
        .Sort = "WG_NUM ASC, LINE_NUM ASC"
        '//-----not concerned if lines exist for delete requests, (the start_date would be null anyway)-----//
        l_strFilter = "RECORD_STATUS = 'I' OR RECORD_STATUS = 'C'"
        .Filter = MDAC25FilterFix(l_strFilter)
    
        If RecordsExist(l_rsAssignmentRequestClone) Then
            .MoveFirst
    
            l_datMinStartDate = !START_DATE
            l_datMaxEndDate = !END_DATE
            l_lngWGNumSave = !WG_NUM
            l_lngLineNumSave = !LINE_NUM
                
            Do While Not .EOF
                '//-----check for change in wg/line
                If !WG_NUM <> l_lngWGNumSave Or !LINE_NUM <> l_lngLineNumSave Then
                    l_tScopeParm(lngMinStartDate) = l_datMinStartDate
                    l_tScopeParm(lngMaxEndDate) = l_datMaxEndDate
                    l_tScopeParm(lngWgNum) = l_lngWGNumSave
                    l_tScopeParm(lngLineNum) = l_lngLineNumSave
                    l_tScopeParm(lngEmpNum) = 0 '//-----mutually exclusive with wg/line
                    l_colScopeParm.Add l_tScopeParm
                    
                    l_datMinStartDate = !START_DATE '//-----reset for next set of parms-----//
                    l_datMaxEndDate = !END_DATE
                    l_lngWGNumSave = !WG_NUM
                    l_lngLineNumSave = !LINE_NUM
                End If
                
                If !START_DATE < l_datMinStartDate Then
                    l_datMinStartDate = !START_DATE
                End If
                If !END_DATE > l_datMaxEndDate Then
                    l_datMaxEndDate = !END_DATE
                End If
                                          
                .MoveNext
            Loop
            l_tScopeParm(lngMinStartDate) = l_datMinStartDate
            l_tScopeParm(lngMaxEndDate) = l_datMaxEndDate
            l_tScopeParm(lngWgNum) = l_lngWGNumSave
            l_tScopeParm(lngLineNum) = l_lngLineNumSave
            l_tScopeParm(lngEmpNum) = 0
        
            l_colScopeParm.Add l_tScopeParm
    
            Set l_rsAssignmentRequestClone = Nothing
            
            Set l_objDBSchDmd = CreateObject("dbSchedDemand.cdSchedDemand")
            Set l_rsLinesBulk = l_objDBSchDmd.GetLinesBulk(l_colScopeParm)
            Set l_objDBSchDmd = Nothing
        End If
        
        .Filter = adFilterNone
        .MoveFirst
    End With

    Set GetLinesBulk = l_rsLinesBulk
    Set l_rsLinesBulk = Nothing
    Set l_colScopeParm = Nothing
    
    Exit Function
eh:
    Set l_objDBSchDmd = Nothing
    Set l_rsLinesBulk = Nothing
    
    GetObjectContext.SetAbort
    TrsRaiseError m_strCMODULENAME, "GetLinesBulk", , , , , True
End Function

'//-----find minimum and maximum values in order to retrieve all scope for validation-----//
Private Function IsReliefComponentValidBulk(ByVal v_rsAssignmentRequest As ADODB.Recordset, ByRef r_rsErrors As ADODB.Recordset) As Boolean
    Dim l_rsEmployeesBulk As ADODB.Recordset
    Dim l_objDBEmployee As dbEmployee.cdEmployee
    Dim l_colEmployee As New Collection
    Dim l_lngEmpNumSave As Long
    Dim l_aryEmployee() As Variant
    Dim l_strFilter As String
    Dim l_rsReliefParms As ADODB.Recordset
    Dim l_enumReliefType As enumReliefType
    Dim l_objBUSCodeTables As busCodeTables.cbsCodeTables
    Dim l_strUnionSave As String, l_strReliefSave As String
    Dim l_bError As Boolean
    Dim l_strBRVMessage As String
    Dim l_rsAssignmentRequestClone As ADODB.Recordset

    On Error GoTo eh
    
    l_bError = False
    '//-----Get just the employees that are needed (union code, job eff date, prior union, prior job eff date)-----//
    Set l_rsAssignmentRequestClone = v_rsAssignmentRequest.Clone

    l_strFilter = "RELIEF_COMP > ''"
    l_rsAssignmentRequestClone.Filter = MDAC25FilterFix(l_strFilter)

    If RecordsExist(l_rsAssignmentRequestClone) Then
        With l_rsAssignmentRequestClone
            '//-----build list of necessary employees (distinct)-----//
            .Sort = "EMP_NUM ASC"
    
            .MoveFirst
    
            l_lngEmpNumSave = !EMP_NUM
    
            Do While Not .EOF
                '//-----check for change in employee
                If !EMP_NUM <> l_lngEmpNumSave Then
                    l_colEmployee.Add l_lngEmpNumSave
    
                    l_lngEmpNumSave = !EMP_NUM '//-----reset for next set of parms-----//
                End If
    
                .MoveNext
            Loop
            l_colEmployee.Add l_lngEmpNumSave
    
            .MoveFirst
        End With
        
        l_aryEmployee = ArrayFromCollection(l_colEmployee, 0)
        
        Set l_objDBEmployee = CreateObject("dbEmployee.cdEmployee")
        Set l_rsEmployeesBulk = l_objDBEmployee.GetEmployee(l_aryEmployee)
        Set l_objDBEmployee = Nothing
        
        Set l_rsReliefParms = New ADODB.Recordset
        With l_rsReliefParms
            .Fields.Append "UNION_CD", adChar, 3, adFldUpdatable
            .Fields.Append "DATE", adDBDate, , adFldUpdatable
            .Fields.Append "RELIEF_COMP", adChar, 3, adFldUpdatable
            '//-----for BRV-----//
            .Fields.Append "WG_NUM", adInteger, , adFldUpdatable
            .Fields.Append "LINE_NUM", adSmallInt, , adFldUpdatable
            .Fields.Append "START_DATE", adDBDate, , adFldUpdatable
            .Fields.Append "END_DATE", adDBDate, , adFldUpdatable
            .Fields.Append "EMP_NUM", adInteger, , adFldUpdatable
            
            .CursorLocation = adUseClient
            .CursorType = adOpenDynamic
            .LockType = adLockBatchOptimistic
            .Open
        
            '//-----Determine if union, prior union, or both are needed for the validatation-----//
            l_rsAssignmentRequestClone.MoveFirst
            l_rsEmployeesBulk.MoveFirst
            Do While Not l_rsAssignmentRequestClone.EOF
                If l_rsAssignmentRequestClone!EMP_NUM <> l_rsEmployeesBulk!EMP_NUM Then
                    l_rsEmployeesBulk.Find "EMP_NUM = " & l_rsAssignmentRequestClone!EMP_NUM
                End If
                
                If Not l_rsEmployeesBulk.EOF Then
                    l_enumReliefType = DetermineReliefType(l_rsAssignmentRequestClone!START_DATE, l_rsAssignmentRequestClone!END_DATE, _
                                            l_rsEmployeesBulk!UNION_CD, l_rsEmployeesBulk!JOB_EFF_DATE, _
                                            l_rsEmployeesBulk!PRIOR_UNION_CD, l_rsEmployeesBulk!PRIOR_JOB_EFF_DT)
'    '//-----for testing
'    l_enumReliefType = DetermineReliefType(l_rsAssignmentRequestClone!START_DATE, l_rsAssignmentRequestClone!END_DATE, _
'                            l_rsEmployeesBulk!UNION_CD, l_rsEmployeesBulk!JOB_EFF_DATE, _
'                            l_rsEmployeesBulk!UNION_CD, #1/1/1993#)

                    Select Case l_enumReliefType
                    Case rtUseUnion
                        .AddNew
                        !EMP_NUM = l_rsEmployeesBulk!EMP_NUM
                        !UNION_CD = l_rsEmployeesBulk!UNION_CD
                        !Date = l_rsAssignmentRequestClone!START_DATE
                        !RELIEF_COMP = l_rsAssignmentRequestClone!RELIEF_COMP
                        !WG_NUM = l_rsAssignmentRequestClone!WG_NUM '//-----for BRV-----//
                        !START_DATE = l_rsAssignmentRequestClone!START_DATE
                        !END_DATE = l_rsAssignmentRequestClone!END_DATE
                        !LINE_NUM = l_rsAssignmentRequestClone!LINE_NUM
                        .Update
                    Case rtUseUnionAndPriorUnion
                        .AddNew
                        !EMP_NUM = l_rsEmployeesBulk!EMP_NUM
                        !UNION_CD = l_rsEmployeesBulk!PRIOR_UNION_CD
                        !Date = l_rsAssignmentRequestClone!START_DATE
                        !RELIEF_COMP = l_rsAssignmentRequestClone!RELIEF_COMP
                        !WG_NUM = l_rsAssignmentRequestClone!WG_NUM
                        !START_DATE = l_rsAssignmentRequestClone!START_DATE
                        !END_DATE = l_rsAssignmentRequestClone!END_DATE
                        !LINE_NUM = l_rsAssignmentRequestClone!LINE_NUM
                        .Update
                        
                        .AddNew
                        !EMP_NUM = l_rsEmployeesBulk!EMP_NUM
                        !UNION_CD = l_rsEmployeesBulk!UNION_CD
                        !Date = l_rsEmployeesBulk!JOB_EFF_DATE
                        !RELIEF_COMP = l_rsAssignmentRequestClone!RELIEF_COMP
                        !WG_NUM = l_rsAssignmentRequestClone!WG_NUM
                        !START_DATE = l_rsAssignmentRequestClone!START_DATE
                        !END_DATE = l_rsAssignmentRequestClone!END_DATE
                        !LINE_NUM = l_rsAssignmentRequestClone!LINE_NUM
                        .Update
                    Case rtUsePriorUnion
                        .AddNew
                        !EMP_NUM = l_rsEmployeesBulk!EMP_NUM
                        !UNION_CD = l_rsEmployeesBulk!PRIOR_UNION_CD
                        !Date = l_rsAssignmentRequestClone!START_DATE
                        !RELIEF_COMP = l_rsAssignmentRequestClone!RELIEF_COMP
                        !WG_NUM = l_rsAssignmentRequestClone!WG_NUM
                        !START_DATE = l_rsAssignmentRequestClone!START_DATE
                        !END_DATE = l_rsAssignmentRequestClone!END_DATE
                        !LINE_NUM = l_rsAssignmentRequestClone!LINE_NUM
                        .Update
                    Case rtErrorNothingBeforePriorUnion
                        l_bError = True
                        If l_rsEmployeesBulk!PRIOR_JOB_EFF_DT <> m_datCOPTIONALSTARTDATE Then
                            l_strBRVMessage = "Cannot validate Relief component " & l_rsAssignmentRequestClone!RELIEF_COMP & _
                                " because unable to determine union code between " & l_rsAssignmentRequestClone!START_DATE & " and " & _
                                DateAdd("d", -1, l_rsEmployeesBulk!PRIOR_JOB_EFF_DT) & "."
                        Else
                            l_strBRVMessage = "Cannot validate Relief component " & l_rsAssignmentRequestClone!RELIEF_COMP & _
                                " because unable to determine union code between " & l_rsAssignmentRequestClone!START_DATE & " and " & _
                                l_rsEmployeesBulk!PRIOR_JOB_EFF_DT & "."
                        End If
                        AppendBusinessError r_rsErrors, trErrbusValidationError, mstrCBUSINESSERR, _
                            l_strBRVMessage, m_strCMODULENAME, _
                            1, 1, "LINE_NUM", _
                            Array("WG_NUM", "LINE_NUM", "START_DATE", "END_DATE", "EMP_NUM"), _
                            CStr(l_rsAssignmentRequestClone!WG_NUM), _
                            CStr(l_rsAssignmentRequestClone!LINE_NUM), _
                            Format(l_rsAssignmentRequestClone!START_DATE, "mm/dd/yyyy"), _
                            Format(l_rsAssignmentRequestClone!END_DATE, "mm/dd/yyyy"), _
                            CStr(l_rsAssignmentRequestClone!EMP_NUM)
                    Case rtErrorNoPriorUnion
                        l_bError = True
                        l_strBRVMessage = "Cannot validate Relief component " & l_rsAssignmentRequestClone!RELIEF_COMP & _
                            " because unable to determine union code before " & l_rsEmployeesBulk!JOB_EFF_DATE & "."
                        AppendBusinessError r_rsErrors, trErrbusValidationError, mstrCBUSINESSERR, _
                            l_strBRVMessage, m_strCMODULENAME, _
                            1, 1, "LINE_NUM", _
                            Array("WG_NUM", "LINE_NUM", "START_DATE", "END_DATE", "EMP_NUM"), _
                            CStr(l_rsAssignmentRequestClone!WG_NUM), _
                            CStr(l_rsAssignmentRequestClone!LINE_NUM), _
                            Format(l_rsAssignmentRequestClone!START_DATE, "mm/dd/yyyy"), _
                            Format(l_rsAssignmentRequestClone!END_DATE, "mm/dd/yyyy"), _
                            CStr(l_rsAssignmentRequestClone!EMP_NUM)
                    End Select
                Else
                    l_bError = True
                    l_strBRVMessage = "Employee information for " & l_rsAssignmentRequestClone!EMP_NUM & _
                                        " not found for relief validation.  Unable to determine union code."
                    AppendBusinessError r_rsErrors, trErrbusValidationError, mstrCBUSINESSERR, _
                        l_strBRVMessage, m_strCMODULENAME, _
                        1, 1, "LINE_NUM", _
                        Array("WG_NUM", "LINE_NUM", "START_DATE", "END_DATE", "EMP_NUM"), _
                        CStr(l_rsAssignmentRequestClone!WG_NUM), _
                        CStr(l_rsAssignmentRequestClone!LINE_NUM), _
                        Format(l_rsAssignmentRequestClone!START_DATE, "mm/dd/yyyy"), _
                        Format(l_rsAssignmentRequestClone!END_DATE, "mm/dd/yyyy"), _
                        CStr(l_rsAssignmentRequestClone!EMP_NUM)
                End If
                l_rsAssignmentRequestClone.MoveNext
            Loop
            
            If RecordsExist(l_rsReliefParms) Then
                .Sort = "UNION_CD, RELIEF_COMP, DATE"
                .MoveFirst
                Set l_objBUSCodeTables = CreateObject("busCodeTables.cbsCodeTables")
                Do While Not .EOF
                    If l_objBUSCodeTables.IsVacationReliefCode(!UNION_CD, !Date, !RELIEF_COMP) Then
                        '//-----skip to next union/relief, all other dates for this union/relief will be valid-----//
                        l_strUnionSave = !UNION_CD
                        l_strReliefSave = !RELIEF_COMP
                        Do While Not .EOF
                            If l_strUnionSave <> !UNION_CD Or l_strReliefSave <> !RELIEF_COMP Then
                                Exit Do
                            End If
                            .MoveNext
                        Loop
                    Else
                        l_bError = True
                        l_strBRVMessage = "Relief component " & !RELIEF_COMP & " is not valid for employee " & !EMP_NUM & _
                                            " in union " & !UNION_CD & _
                                            " between " & !START_DATE & " and " & !END_DATE & "."
                        AppendBusinessError r_rsErrors, trErrbusValidationError, mstrCBUSINESSERR, _
                            l_strBRVMessage, m_strCMODULENAME, _
                            1, 1, "LINE_NUM", _
                            Array("WG_NUM", "LINE_NUM", "START_DATE", "END_DATE", "EMP_NUM"), _
                            CStr(!WG_NUM), _
                            CStr(!LINE_NUM), _
                            Format(!START_DATE, "mm/dd/yyyy"), _
                            Format(!END_DATE, "mm/dd/yyyy"), _
                            CStr(!EMP_NUM)
        
                        .MoveNext
                    End If
                Loop
            End If
            Set l_objBUSCodeTables = Nothing
        End With
    Else '//-----nothing to validate-----//
            
    End If
    
'    Set GetEmployeeInfoBulk = l_rsEmployeesBulk
    l_rsAssignmentRequestClone.Filter = adFilterNone

    '//Set l_colScopeParm = Nothing
    IsReliefComponentValidBulk = Not l_bError
    If IsReliefComponentValidBulk Then
        GetObjectContext.SetComplete
    Else
        GetObjectContext.SetAbort
    End If
    
    Set l_rsEmployeesBulk = Nothing
    Set l_rsAssignmentRequestClone = Nothing
    Set l_colEmployee = Nothing
    Set l_rsReliefParms = Nothing
    Set l_objBUSCodeTables = Nothing

    Exit Function
eh:
    Set l_rsEmployeesBulk = Nothing
    Set l_rsAssignmentRequestClone = Nothing
    Set l_objDBEmployee = Nothing
    Set l_colEmployee = Nothing
    Set l_rsReliefParms = Nothing
    Set l_objBUSCodeTables = Nothing

    GetObjectContext.SetAbort
    TrsRaiseError m_strCMODULENAME, "IsReliefComponentValidBulk", , , , , True
End Function

Private Function DetermineReliefType(ByVal v_datStartDate As Date, ByVal v_datEndDate As Date, _
                                    ByVal v_strUnion As String, ByVal v_datJobEffectiveDate As Date, _
                                    ByVal v_strPriorUnion As String, ByVal v_datPriorJobEffectiveDate As Date) As enumReliefType
    Dim l_enumReliefType As enumReliefType
    
    On Error GoTo eh
        
    l_enumReliefType = rtErrorUnknown
    '//-----the long format-----//
    If v_datStartDate < v_datJobEffectiveDate Then
        If v_datStartDate >= v_datPriorJobEffectiveDate Then
            If Trim(v_strPriorUnion) <> "" Then
                If v_datEndDate < v_datJobEffectiveDate Then
                    l_enumReliefType = rtUsePriorUnion '//-----use prior union
                ElseIf v_datEndDate = v_datJobEffectiveDate Then
                    l_enumReliefType = rtUseUnionAndPriorUnion '//-----use prior union & union
                ElseIf v_datEndDate > v_datJobEffectiveDate Then
                    l_enumReliefType = rtUseUnionAndPriorUnion '//-----use prior union & union
                End If
            Else
                l_enumReliefType = rtErrorNoPriorUnion '//-----BRV
            End If
        Else
            l_enumReliefType = rtErrorNothingBeforePriorUnion '//-----BRV
        End If
    ElseIf v_datStartDate = v_datJobEffectiveDate Then
        If v_datEndDate < v_datJobEffectiveDate Then
        ElseIf v_datEndDate = v_datJobEffectiveDate Then
            l_enumReliefType = rtUseUnion  '//-----empnum is in 1 union for entire date range
        ElseIf v_datEndDate > v_datJobEffectiveDate Then
            l_enumReliefType = rtUseUnion  '//-----empnum is in 1 union for entire date range
        End If
    ElseIf v_datStartDate > v_datJobEffectiveDate Then
        If v_datEndDate < v_datJobEffectiveDate Then
        ElseIf v_datEndDate = v_datJobEffectiveDate Then
        ElseIf v_datEndDate > v_datJobEffectiveDate Then
            l_enumReliefType = rtUseUnion  '//-----empnum is in 1 union for entire date range
        End If
    End If
    
    DetermineReliefType = l_enumReliefType
    Exit Function
eh:
    GetObjectContext.SetAbort
    TrsRaiseError m_strCMODULENAME, "DetermineReliefType", , , , , True
End Function

Private Function ProcessCascade( _
                    ByVal v_lTRISUser As Long, _
                    ByVal v_blnConfirmConflicts As Boolean, _
                    ByRef r_blnConfirmPending As Boolean, _
                    ByRef r_varCascadeUpdates As Variant, _
                    ByRef r_rsSchedDemandUpdates As ADODB.Recordset, _
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

    If ActualTimeExists(r_rsEmpSchedDates, r_rsErrors) Then
        ProcessCascade = False
        Exit Function
    End If

    Set l_cbsCascade = CreateObject("busCascade.cbsCascade")
    Set l_rsCascadeConflicts = l_cbsCascade.GetExceptionConflicts( _
                r_rsEmpSchedDates, trDemandLineBaseSchedLevel, _
                l_rsExceptionSPA, l_rsAvailScheduleSPA, , , , , r_rsSchedDemandUpdates)

    If RecordsExist(l_rsCascadeConflicts) Then
        If v_blnConfirmConflicts And _
                IsConflictToConfirm(l_rsCascadeConflicts) Then
            Set r_rsCascadeConflicts = l_rsCascadeConflicts
            Set r_rsExceptionSPA = l_rsExceptionSPA
            Set r_rsAvailScheduleSPA = l_rsAvailScheduleSPA
'            In case inprocess with client
            ClearAllFilters r_rsSchedDemandUpdates, _
                    r_rsCascadeConflicts, r_rsExceptionSPA, r_rsAvailScheduleSPA
            r_blnConfirmPending = True
        Else
            If l_cbsCascade Is Nothing Then
                Set l_cbsCascade = CreateObject("busCascade.cbsCascade")
            End If
            l_cbsCascade.UpdateExceptions _
                l_rsCascadeConflicts, l_rsExceptionSPA, l_rsAvailScheduleSPA, _
                mblnAPPLYCASCADE, l_varCascadeUpdates, l_rsCascadeErrors, v_lTRISUser
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
    Set r_rsEmpSchedDates = Nothing
    Set l_rsCascadeConflicts = Nothing
    Set l_rsExceptionSPA = Nothing
    Set l_rsAvailScheduleSPA = Nothing
    Set l_cbsCascade = Nothing
    Set l_rsCascadeErrors = Nothing
    TrsRaiseError m_strCMODULENAME, "ProcessCascade"
End Function


Private Function IsConflictToConfirm( _
                    ByRef r_rsCascadeConflicts As ADODB.Recordset) As Boolean
' Answer True when there are any conflicts with a different schedule code.
' The user only wants to confirm schedule changes.  If it is a source only
' change the user does not want to see it.
                    
On Error GoTo eh
    r_rsCascadeConflicts.MoveFirst
    Do Until r_rsCascadeConflicts.EOF
        If r_rsCascadeConflicts!SCHED_CD <> r_rsCascadeConflicts!NEW_SCHED_CD Then
            IsConflictToConfirm = True
            Exit Do
        End If
        r_rsCascadeConflicts.MoveNext
    Loop
                    
    r_rsCascadeConflicts.MoveFirst
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "IsConflictToConfirm"
End Function


Private Function ProcessScheduleChange( _
                    ByRef r_rsEmpSchedDates As ADODB.Recordset, _
                    ByRef r_rsErrors As ADODB.Recordset, _
                    ByRef r_rsSchedDemandUpdates As ADODB.Recordset) _
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
    ScheduleChangePaymentExists r_rsEmpSchedDates, r_rsErrors, r_rsSchedDemandUpdates
    
    ' update the emp_mod timestamp
    ' this will create a lock that time collection must wait on before saving actual time.
    Set l_cdCascade = CreateObject("dbCascade.cdCascade")
    l_cdCascade.SetEmployeeScheduleModified l_varEmps
    Set l_cdCascade = Nothing
    
    If ActualTimeExists(r_rsEmpSchedDates, r_rsErrors) Then
        ProcessScheduleChange = False
        Exit Function
    End If

    ProcessScheduleChange = True
    Exit Function
    
eh:
    Set l_cdCascade = Nothing
    Set r_rsErrors = Nothing
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
    TrsRaiseError m_strCMODULENAME, "ExtractEmployeeNumbers"
End Function


Private Function ActualTimeExists( _
                    ByRef r_rsEmpSchedDates As ADODB.Recordset, _
                    ByRef r_rsErrors As ADODB.Recordset) As Boolean
' 6/27/02 RSG
' WG_NUM and LIN_NUM are not readily available for the BRV, just
' populate 0.  This will present a problem for navigation back to request entry,
' but will probably handle that with a RECORD_ID on request.  Whatever is used
' it will have to be carried through emp sched dates or have another way to
' match it up to the request.
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
                m_strCMODULENAME, _
                TrType_DEM, _
                , _
                , _
                mstrCVARFIELDNAMES, _
                "0", _
                "0", _
                r_rsEmpSchedDates!START_DATE, _
                r_rsEmpSchedDates!END_DATE, _
                r_rsEmpSchedDates!EMP_NUM
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
    TrsRaiseError m_strCMODULENAME, "ActualTimeExists"
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
    TrsRaiseError m_strCMODULENAME, "DetermineGetActualTimeScope"
End Sub


Private Function AppendCascadeErrors( _
                    ByRef r_rsCascadeErrors As ADODB.Recordset, _
                    ByRef r_rsErrors As ADODB.Recordset)
' Append cascade errors/information to the client business error set.
' The client cannot navigate cascade errors so just ensure
' compatible datatypes are populated in the key fields.
' WG_NUM, LINE_NUM, START_DATE, END_DATE, EMP_NUM
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
            1, 1, "", _
            Array("WG_NUM", "LINE_NUM", "START_DATE", "END_DATE", "EMP_NUM"), _
            "0", _
            "0", _
            r_rsCascadeErrors!KEY_FIELD2, _
            r_rsCascadeErrors!KEY_FIELD2, _
            r_rsCascadeErrors!KEY_FIELD1
        r_rsCascadeErrors.MoveNext
    Loop
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "AppendCascadeErrors"
End Function


Private Sub AppendCascadeDeleteInsert(ByRef r_rsRequest As ADODB.Recordset)
'    Capture entire request interval for deletes and inserts, except
'    for unassignments(EMP_NUM = 0).  These will be picked up during overlap
'    line processing.
    Dim l_varFilter As Variant
    Dim l_strRecStatus As String

On Error GoTo eh

    l_varFilter = r_rsRequest.Filter
'    Clear filter to check every record in request.
    r_rsRequest.Filter = 0
    Do Until r_rsRequest.EOF
        If r_rsRequest!EMP_NUM > 0 Then
            l_strRecStatus = Trim(r_rsRequest!RECORD_STATUS)
            If l_strRecStatus = "D" Then
                AppendCascadeDates _
                    l_strRecStatus, r_rsRequest!EMP_NUM, _
                    r_rsRequest!KEY_START_DATE, _
                    r_rsRequest!KEY_END_DATE
            ElseIf l_strRecStatus = "I" Then
                AppendCascadeDates _
                    l_strRecStatus, r_rsRequest!EMP_NUM, _
                    r_rsRequest!START_DATE, _
                    r_rsRequest!END_DATE
            End If
        End If
        r_rsRequest.MoveNext
    Loop
'    Restore original filter, if any, and reposition on first record.
    r_rsRequest.Filter = l_varFilter

    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "AppendCascadeDeleteInsert"
End Sub


Private Sub AppendCascadeChange(ByRef r_rsRequest As ADODB.Recordset)
'    Capture difference in start/end dates and capture entire
'    entire request when line change.

On Error GoTo eh
    If r_rsRequest!START_DATE < r_rsRequest!KEY_START_DATE Then
        AppendCascadeDates _
            "C", r_rsRequest!EMP_NUM, _
            r_rsRequest!START_DATE, _
            DateAdd("d", -1, r_rsRequest!KEY_START_DATE)
    ElseIf r_rsRequest!START_DATE > r_rsRequest!KEY_START_DATE Then
        AppendCascadeDates _
            "C", r_rsRequest!EMP_NUM, _
            r_rsRequest!KEY_START_DATE, _
            DateAdd("d", -1, r_rsRequest!START_DATE)
    End If
    
    If r_rsRequest!END_DATE < r_rsRequest!KEY_END_DATE Then
        AppendCascadeDates _
            "C", r_rsRequest!EMP_NUM, _
            DateAdd("d", 1, r_rsRequest!END_DATE), _
            r_rsRequest!KEY_END_DATE
    ElseIf r_rsRequest!END_DATE > r_rsRequest!KEY_END_DATE Then
        AppendCascadeDates _
            "C", r_rsRequest!EMP_NUM, _
            DateAdd("d", 1, r_rsRequest!KEY_END_DATE), _
            r_rsRequest!END_DATE
    End If

    If r_rsRequest!LINE_NUM <> r_rsRequest!KEY_LINE_NUM Then
        AppendCascadeDates _
            "C", r_rsRequest!EMP_NUM, _
            r_rsRequest!START_DATE, _
            r_rsRequest!END_DATE
        AppendCascadeDates _
            "C", r_rsRequest!EMP_NUM, _
            r_rsRequest!KEY_START_DATE, _
            r_rsRequest!KEY_END_DATE
    End If

    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "AppendCascadeChange"
End Sub



Private Sub AppendCascadeDates( _
                ByVal v_strAction As String, _
                ByVal v_lngEmpNum As Long, _
                ByVal v_datStartDate As Date, _
                ByVal v_datEndDate As Date)
' Simply add the date records, without attempting to adjust as added.
' Consolidate all the date intervals changed at the end of the process by
' employeeNumber.

    Dim l_objDateInterval As cpuDateInterval

On Error GoTo eh
    Set l_objDateInterval = New cpuDateInterval
    l_objDateInterval.EmployeeNumber = v_lngEmpNum
    l_objDateInterval.StartDate = v_datStartDate
    l_objDateInterval.EndDate = v_datEndDate
    GetCascadeIntervals.Add l_objDateInterval
    
    Set l_objDateInterval = Nothing
    Exit Sub
    
eh:
    Set l_objDateInterval = Nothing
    TrsRaiseError m_strCMODULENAME, "AppendCascadeDates"
End Sub


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
    TrsRaiseError m_strCMODULENAME, "GetCascadeIntervals"
End Function


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
    TrsRaiseError m_strCMODULENAME, "ConsolidatedCascadeIntervals"
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
    TrsRaiseError m_strCMODULENAME, "FilterCascadeIntervals"
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
    TrsRaiseError m_strCMODULENAME, "JoinIntervals"
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
        If l_cpuDateInterval.EmployeeNumber <> 0 Then
            l_rsEmpDates.AddNew
            l_rsEmpDates!EMP_NUM = l_cpuDateInterval.EmployeeNumber
            l_rsEmpDates!START_DATE = l_cpuDateInterval.StartDate
            l_rsEmpDates!END_DATE = l_cpuDateInterval.EndDate
            l_rsEmpDates.Update
        End If
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
    TrsRaiseError m_strCMODULENAME, "GetEmpSchedDates"
End Function

'Method HighManHours not needed in TRIS 5
'''''Private Function HighManHours( _
'''''    ByVal v_rsDemandUpdates As ADODB.Recordset, _
'''''    ByVal v_lTRISUser As Long _
'''''    ) As Boolean
'''''
''''''This method will set OT Banks to high hours when an
''''''employee is new to the group
''''''If the employee is already in the group he will
'''''' not be set to high hours
''''''If the employee is continuing an assignment in the same group
'''''' his OT bank will not be set to high hours
''''''Please see busOvertimeEqualizer for more details
'''''
'''''    Dim l_obj As busOvertimeEqualizer.cbsOTEqualizer
'''''
'''''On Error GoTo eh
'''''
'''''    If v_rsDemandUpdates Is Nothing Then
'''''        HighManHours = True
'''''        Exit Function
'''''    End If
'''''    v_rsDemandUpdates.Filter = ""
'''''    If Not RecordsExist(v_rsDemandUpdates) Then
'''''        HighManHours = True
'''''        Exit Function
'''''    End If
'''''    Set l_obj = CreateObject("busOvertimeEqualizer.cbsOTEqualizer")
'''''    HighManHours = l_obj.HighmanHoursDemandLineAssignments( _
'''''            v_rsDemandUpdates, v_lTRISUser)
'''''    Set l_obj = Nothing
'''''    Exit Function
'''''
'''''eh:
'''''
'''''  TrsRaiseError m_strCMODULENAME, "HighManHours"
'''''
'''''End Function


Private Function ScheduleChangePaymentExists( _
                    ByRef r_rsEmpSchedDates As ADODB.Recordset, _
                    ByRef r_rsErrors As ADODB.Recordset, _
                    ByRef r_rsSchedDemandUpdates As ADODB.Recordset _
                ) As Boolean
' Check for any schedule change payments during the scope of the demand line assignment change.
' If a payment is found on an off-day(ORPHAN_IND= Y), just warn the user so they can take
' appropriate action, e.g. move the payment to the next scheduled work day.
' Answer TRUE if a payment exists for an off-day, otherwise FALSE.
    Dim l_blnExists As Boolean
    Dim l_cbsSchedChgPay As busSchedChgPay.cbsSchedChgPay
    Dim l_rsSchedChgPay As ADODB.Recordset
                    
    Dim l_rsSchedDemandUpdates As ADODB.Recordset
                    
On Error GoTo eh
                    
    l_blnExists = False
    If IsEmpty(r_rsEmpSchedDates) Then
        ScheduleChangePaymentExists = False
        Exit Function
    End If
    
'    r_rsSchedDemandUpdates are explicit update requests, i.e. DIC, not using
'    recordset Status property.
    Set l_rsSchedDemandUpdates = r_rsSchedDemandUpdates
                          
    Set l_cbsSchedChgPay = CreateObject("busSchedChgPay.cbsSchedChgPay")
    Set l_rsSchedChgPay = l_cbsSchedChgPay.GetSchedChgPayPending(r_rsEmpSchedDates, _
            Nothing, Nothing, Nothing, Nothing, l_rsSchedDemandUpdates)
    Set l_cbsSchedChgPay = Nothing
    
    If Not RecordsExist(l_rsSchedChgPay) Then
        ScheduleChangePaymentExists = False
        Exit Function
    Else
        l_rsSchedChgPay.MoveFirst
    End If

    Do Until l_rsSchedChgPay.EOF
        If Trim(l_rsSchedChgPay!ORPHAN_IND) = "Y" Then
            l_blnExists = True
            AppendWarningScheduleChangePayment _
                    l_rsSchedChgPay!EMP_NUM, _
                    l_rsSchedChgPay!EMP_NAME, _
                    l_rsSchedChgPay!CHARGE_DATE, _
                    r_rsErrors
        End If
        l_rsSchedChgPay.MoveNext
    Loop

    ScheduleChangePaymentExists = l_blnExists
    Set l_rsSchedChgPay = Nothing
    Set l_rsSchedDemandUpdates = Nothing
    Exit Function
    
eh:
    Set l_cbsSchedChgPay = Nothing
    Set l_rsSchedChgPay = Nothing
    Set l_rsSchedDemandUpdates = Nothing
    TrsRaiseError m_strCMODULENAME, "ScheduleChangePaymentExists"
End Function


Private Function PendingRecords(ByVal v_rsRecordset As ADODB.Recordset) As ADODB.Recordset
' If v_rsRecordset has pending records, answer it; otherwise Nothing.

On Error GoTo eh

    If v_rsRecordset Is Nothing Then
        Set PendingRecords = Nothing
        Exit Function
    End If
    
    v_rsRecordset.Filter = adFilterNone
    v_rsRecordset.Filter = adFilterPendingRecords
    If v_rsRecordset.RecordCount > 0 Then
        Set PendingRecords = v_rsRecordset
    Else
        Set PendingRecords = Nothing
    End If
    
    v_rsRecordset.Filter = adFilterNone
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "PendingRecords"
End Function


Private Function AppendWarningScheduleChangePayment( _
                    ByVal v_lngEmpNum As Long, _
                    ByVal v_strEmpName As String, _
                    ByVal v_datChargeDate As Date, _
                    ByRef r_rsErrors As ADODB.Recordset)

    Dim l_strWarDesc As String
    Dim l_colDescArgs As Collection

On Error GoTo eh

    Set l_colDescArgs = New Collection
    l_colDescArgs.Add CStr(v_lngEmpNum)
    l_colDescArgs.Add Trim(v_strEmpName)
    l_colDescArgs.Add Format(v_datChargeDate, "mm/dd/yyyy")
    l_strWarDesc = TrsGetErrDescription(LongAppErr(trWarSchedChgPmtOffday), l_colDescArgs)
    Set l_colDescArgs = Nothing
    AppendBusinessError _
        r_rsErrors, _
        LongAppErr(trWarSchedChgPmtOffday), _
        mstrCBUSINESSWAR, _
        l_strWarDesc, _
        m_strCMODULENAME, _
        TrType_DEM, _
        , , _
        mstrCVARFIELDNAMES, _
        "0", _
        "0", _
        Format(v_datChargeDate, "mm/dd/yyyy"), _
        Format(v_datChargeDate, "mm/dd/yyyy"), _
        CStr(v_lngEmpNum)

    Exit Function

eh:
    Set l_colDescArgs = Nothing
    TrsRaiseError m_strCMODULENAME, "AppendWarningScheduleChangePayment"
End Function
'BEGIN PREM TRIS CONVERSION PHASE 3
Public Function ProcessDemandJobPositions(ByVal vadrsJobPositions As ADODB.Recordset _
                , ByRef radrsErrors As ADODB.Recordset _
                , ByVal vlUserEmployeeNumber As Long _
                ) As Boolean
On Error GoTo ProcessDemandJobPositions_EH
Dim llRecordIndex As Long
Dim lobjdb As dbSchedDemand.cdSchedDemandTrans

    'Prepare BRV Recordset
    Set radrsErrors = CreateDemandJobPositionBRV

    If Not (RecordsExist(vadrsJobPositions)) Then
        AppendDemandJobPositionBRV radrsErrors _
            , trErrbusDemandJobPositions _
            , mstrCBUSINESSERR _
            , "There are no Demand Job Position entries to process." _
            , m_strCMODULENAME & ": ProcessDemandJobPositions"
        GoTo REPORTBRV
    Else
        With vadrsJobPositions
            .Filter = adFilterNone
            If Not (.EOF) Then
                .MoveFirst

                If Not (ValidateJobPositionEntries( _
                    vadrsJobPositions _
                    , radrsErrors _
                    , vlUserEmployeeNumber)) Then
                    GoTo REPORTBRV
                End If

                .Filter = adFilterNone
                .MoveFirst
            End If
        End With
    End If

    Set lobjdb = CreateObject( _
                            "dbSchedDemand.cdSchedDemandTrans")
    If Not (lobjdb.dbSaveJobPosition( _
            vadrsJobPositions _
            , vlUserEmployeeNumber)) Then
        Set lobjdb = Nothing

        Err.Raise trErrbusDemandJobPositions
    End If
    Set lobjdb = Nothing

    GetObjectContext().SetComplete

    ProcessDemandJobPositions = True

    Exit Function
ProcessDemandJobPositions_EH:
    GetObjectContext().SetAbort

    Set lobjdb = Nothing

    Select Case (Err.Number)
        Case trErrutOpenConnection
            TrsRaiseError m_strCMODULENAME _
                , "ProcessDemandJobPositions"
        Case trErrdbDelDemandJobPositions, trErrdbInsDemandJobPositions _
            , trErrdbUpdDemandJobPositions, trErrdbSaveDemandJobPositions
            TrsRaiseError m_strCMODULENAME _
                , "ProcessDemandJobPositions" _
                , Err.Number _
                , , , , _
                True
        Case Else
            TrsRaiseError m_strCMODULENAME _
                , "ProcessDemandJobPositions" _
                , trErrbusDemandJobPositions _
                , , , , _
                True
    End Select

    Exit Function

REPORTBRV:
    Set lobjdb = Nothing

    If Not (radrsErrors Is Nothing) Then
        With radrsErrors
            If Not (.EOF And .BOF) Then
                .MoveFirst
                .Filter = adFilterNone

                'Check for errors
'MDAC25 FILTER FIX
                .Filter = MDAC25FilterFix("ERR_TYPE = 'BE'")
                'If found abort else complete
                If Not (.EOF) Then
                    'Abort Transaction
                    GetObjectContext().SetAbort
                    ProcessDemandJobPositions = False
                Else
                    'Complete Transaction
                    GetObjectContext().SetComplete
                    ProcessDemandJobPositions = True
                End If

                .Filter = adFilterNone
            Else
                'Complete Transaction
                GetObjectContext().SetComplete
                ProcessDemandJobPositions = True
            End If
        End With
    Else
        'Complete Transaction
        GetObjectContext().SetComplete
        ProcessDemandJobPositions = True
    End If

    Exit Function
    
End Function
Private Function CreateDemandJobPositionBRV() As ADODB.Recordset
On Error GoTo CreateDemandJobPositionBRV_EH
Dim ladrs As ADODB.Recordset
Const ADOFIELDATTR = adFldUpdatable Or adFldIsNullable

    Set ladrs = New ADODB.Recordset
    With ladrs
        .Fields.Append "WG_NUM", adInteger, , ADOFIELDATTR
        .Fields.Append "CHARGE_DATE", adDate, , ADOFIELDATTR
        .Fields.Append "LINE_NUM", adInteger, , ADOFIELDATTR

        .Fields.Append "ERR_NUMBER", adInteger, , ADOFIELDATTR
        .Fields.Append "ERR_TYPE", adChar, 2, ADOFIELDATTR
        .Fields.Append "ERR_DESCRIPTION", adChar, 300, ADOFIELDATTR
        .Fields.Append "ERR_SOURCE", adChar, 1000, ADOFIELDATTR

        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open
    End With

    Set CreateDemandJobPositionBRV = ladrs

    Set ladrs = Nothing
    Exit Function
CreateDemandJobPositionBRV_EH:
    Set ladrs = Nothing

    TrsRaiseError m_strCMODULENAME _
        , "CreateDemandJobPositionBRV"

End Function
Private Sub AppendDemandJobPositionBRV(ByRef radrsBRV As ADODB.Recordset _
                , ByVal vlErrorNumber As Long _
                , ByVal vsErrorType As String _
                , ByVal vsErrorDescription As String _
                , ByVal vsErrorSource As String _
                , Optional ByVal vlWorkgroup As Long = 0 _
                , Optional ByVal vdtChargeDate As Date = #1/1/100# _
                , Optional ByVal vlLineNumber As Long = 0)
On Error GoTo AppendDemandJobPositionBRV_EH

    With radrsBRV
        If (.State <> adStateOpen) Then
            Exit Sub
        End If

        .AddNew

        !WG_NUM = vlWorkgroup
        !CHARGE_DATE = vdtChargeDate
        !LINE_NUM = vlLineNumber

        !ERR_NUMBER = vlErrorNumber - vbObjectError
        !ERR_TYPE = Trim(vsErrorType)
        !ERR_DESCRIPTION = Trim(vsErrorDescription)
        !ERR_SOURCE = Trim(vsErrorSource)

        .Update
    End With

    Exit Sub
AppendDemandJobPositionBRV_EH:
    TrsRaiseError m_strCMODULENAME _
        , "AppendDemandJobPositionBRV"

End Sub
Private Function ValidateJobPositionEntries(ByRef radrsRSToValidate As ADODB.Recordset _
                , ByRef radrsBRV As ADODB.Recordset _
                , ByVal vlUserEmployeeNumber As Long _
                ) As Boolean
On Error GoTo ValidateJobPositionEntries_EH
Dim llRecordIndex As Long
Dim ladrs As ADODB.Recordset

    ValidateJobPositionEntries = False

    If Not (RecordsExist(radrsRSToValidate)) Then
        Exit Function
    End If

    If Not (stdmTRISCommon.CheckForValidLong(vlUserEmployeeNumber, True)) Then
        AppendDemandJobPositionBRV radrsBRV _
            , trErrbusDemandJobPositions _
            , mstrCBUSINESSERR _
            , "Invalid TRIS User." _
            , m_strCMODULENAME & ": ValidateJobPositionEntries"
    End If

    With radrsRSToValidate
        .MoveFirst

        For llRecordIndex = 1 To .RecordCount
            Select Case (UCase(Trim(!RECORD_STATUS)))
                Case "I"
                    If Not (stdmTRISCommon.CheckForValidLong(!WG_NUM, True)) Then
                        AppendDemandJobPositionBRV radrsBRV _
                            , trErrbusDemandJobPositions _
                            , mstrCBUSINESSERR _
                            , "Invalid Workgroup." _
                            , m_strCMODULENAME & ": ValidateJobPositionEntries" _
                            , _
                            , !CHARGE_DATE _
                            , !LINE_NUM
                    Else
                        If Not (CheckForItemExistenceInDB(1, !WG_NUM, !CHARGE_DATE)) Then
                            AppendDemandJobPositionBRV radrsBRV _
                                , trErrbusDemandJobPositions _
                                , mstrCBUSINESSERR _
                                , "Invalid Workgroup. Workgroup " & Trim(!WG_NUM) & " does not exist." _
                                , m_strCMODULENAME & ": ValidateJobPositionEntries" _
                                , !WG_NUM _
                                , !CHARGE_DATE _
                                , !LINE_NUM
                        End If
                    End If

                    If Not (CheckForValidDate(!CHARGE_DATE)) Then
                        AppendDemandJobPositionBRV radrsBRV _
                            , trErrbusDemandJobPositions _
                            , mstrCBUSINESSERR _
                            , "Invalid Charge Date." _
                            , m_strCMODULENAME & ": ValidateJobPositionEntries" _
                            , !WG_NUM _
                            , _
                            , !LINE_NUM
                    End If

                    If Not (stdmTRISCommon.CheckForValidLong(!LINE_NUM, True)) Then
                        AppendDemandJobPositionBRV radrsBRV _
                            , trErrbusDemandJobPositions _
                            , mstrCBUSINESSERR _
                            , "Invalid Workgroup Line." _
                            , m_strCMODULENAME & ": ValidateJobPositionEntries" _
                            , !WG_NUM _
                            , !CHARGE_DATE
                    End If

                    If Not (stdmTRISCommon.CheckForValidString(!JOB_POSITION)) Then
                            AppendDemandJobPositionBRV radrsBRV _
                                , trErrbusDemandJobPositions _
                                , mstrCBUSINESSERR _
                                , "Invalid Job Position." _
                                , m_strCMODULENAME & ": ValidateJobPositionEntries" _
                                , !WG_NUM _
                                , !CHARGE_DATE _
                                , !LINE_NUM
                    End If

                Case "C", "D"
                    If Not (stdmTRISCommon.CheckForValidLong(!KEY_WG_NUM, True)) Then
                        AppendDemandJobPositionBRV radrsBRV _
                            , trErrbusDemandJobPositions _
                            , mstrCBUSINESSERR _
                            , "Invalid Workgroup." _
                            , m_strCMODULENAME & ": ValidateJobPositionEntries" _
                            , _
                            , !KEY_CHARGE_DATE _
                            , !KEY_LINE_NUM
                    Else
                        If Not (CheckForItemExistenceInDB(1, !KEY_WG_NUM, !KEY_CHARGE_DATE)) Then
                            AppendDemandJobPositionBRV radrsBRV _
                                , trErrbusDemandJobPositions _
                                , mstrCBUSINESSERR _
                                , "Invalid Workgroup. Workgroup " & Trim(!KEY_WG_NUM) & " does not exist." _
                                , m_strCMODULENAME & ": ValidateJobPositionEntries" _
                                , !KEY_WG_NUM _
                                , !KEY_CHARGE_DATE _
                                , !KEY_LINE_NUM
                        End If
                    End If

                    If Not (CheckForValidDate(!KEY_CHARGE_DATE)) Then
                        AppendDemandJobPositionBRV radrsBRV _
                            , trErrbusDemandJobPositions _
                            , mstrCBUSINESSERR _
                            , "Invalid Charge Date." _
                            , m_strCMODULENAME & ": ValidateJobPositionEntries" _
                            , !KEY_WG_NUM _
                            , _
                            , !KEY_LINE_NUM
                    End If

                    If Not (stdmTRISCommon.CheckForValidLong(!KEY_LINE_NUM, True)) Then
                        AppendDemandJobPositionBRV radrsBRV _
                            , trErrbusDemandJobPositions _
                            , mstrCBUSINESSERR _
                            , "Invalid Workgroup Line." _
                            , m_strCMODULENAME & ": ValidateJobPositionEntries" _
                            , !KEY_WG_NUM _
                            , !KEY_CHARGE_DATE
                    End If

                    If Not (stdmTRISCommon.CheckForValidString(!TIME_STAMP _
                        , , , True, 26)) Then
                        AppendDemandJobPositionBRV radrsBRV _
                            , trErrbusDemandJobPositions _
                            , mstrCBUSINESSERR _
                            , "Invalid Change request. Time Stamp is invalid." _
                            , m_strCMODULENAME & ": ValidateJobPositionEntries" _
                            , !KEY_WG_NUM _
                            , !KEY_CHARGE_DATE _
                            , !KEY_LINE_NUM
                    End If

                    If (UCase(Trim(!RECORD_STATUS)) = "C") Then
                        If Not (stdmTRISCommon.CheckForValidString(!JOB_POSITION)) Then
                            AppendDemandJobPositionBRV radrsBRV _
                                , trErrbusDemandJobPositions _
                                , mstrCBUSINESSERR _
                                , "Invalid Job Position." _
                                , m_strCMODULENAME & ": ValidateJobPositionEntries" _
                                , !KEY_WG_NUM _
                                , !KEY_CHARGE_DATE _
                                , !KEY_LINE_NUM
                        End If
                    End If
            End Select

            .MoveNext
        Next llRecordIndex

        .MoveFirst
    End With

    radrsRSToValidate.Filter = adFilterNone
    radrsRSToValidate.MoveFirst

    ValidateJobPositionEntries = True

    If Not (radrsBRV Is Nothing) Then
        With radrsBRV
            If Not (.BOF And .EOF) Then
                .Filter = adFilterNone

                .Filter = MDAC25FilterFix("ERR_TYPE = 'BE' OR ERR_TYPE = 'AE' OR ERR_TYPE = 'AW'")
                If Not (.EOF) Then
                    ValidateJobPositionEntries = False
                End If

                .Filter = adFilterNone
            End If
        End With
    End If

    Exit Function
ValidateJobPositionEntries_EH:
    ValidateJobPositionEntries = False

    TrsRaiseError m_strCMODULENAME _
            , "ValidateJobPositionEntries"

End Function
Private Function CheckForItemExistenceInDB(ByVal viItemType As Integer _
                    , Optional ByVal vlLongItem As Long = 0 _
                    , Optional ByVal vdtDateItem As Date = #1/1/100# _
                    ) As Boolean
On Error GoTo CheckForItemExistenceInDB_EH
Dim lobjdbWorkgroup As dbWorkGroup.cdWorkGroup

    CheckForItemExistenceInDB = False

    Set lobjdbWorkgroup = CreateObject( _
                                "dbWorkGroup.cdWorkGroup")
    Select Case (viItemType)
        Case 1
            If (lobjdbWorkgroup.IsValidWorkgroup(vlLongItem, vdtDateItem)) Then
                If (lobjdbWorkgroup.DemandIndicator(vlLongItem) = "Y") Then
                    CheckForItemExistenceInDB = True
                End If
            End If

        Case 2
            '
    End Select
    Set lobjdbWorkgroup = Nothing

    GetObjectContext.SetComplete

    Exit Function
CheckForItemExistenceInDB_EH:
    GetObjectContext.SetAbort
    Set lobjdbWorkgroup = Nothing

    TrsRaiseError m_strCMODULENAME _
        , "CheckForItemExistenceInDB"

End Function
'END PREM TRIS CONVERSION PHASE 3

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

    TrsRaiseError m_strCMODULENAME _
        , "IsStartDateBeforeEmployeeEmploymentDate"


 End Function
