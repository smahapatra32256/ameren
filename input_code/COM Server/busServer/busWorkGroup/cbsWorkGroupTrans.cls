VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "cbsWorkGroupTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_strUserId As String
Private m_rsErrors As ADODB.Recordset
Private m_lngWorkGroup As Long
Private m_blnWGLineSchedChg As Boolean
Private m_blnDemandWorkGroup As Boolean
Private m_colEmployees As Collection 'This is the unique list of employees in the request
Private m_rsWorkGroupAttributes As ADODB.Recordset
Private m_rsWG_Assignments As ADODB.Recordset 'Employee Assignments in Work Group
Private m_rsWG_PAIOWG As ADODB.Recordset ' PAIOWG = Primary Assignments In Other Work Groups
Private m_rsWG_ALL_PA As ADODB.Recordset ' PA =  = Primary Assignments
                                         ' This combines the primary assignments in the
                                         ' 2 recordsets above
Private m_rsWG_Assignments_Copy As ADODB.Recordset  'cascade use only - original primary assignments
Private m_colWorkGroups As Collection 'This is the unique list of work groups in the request
Private m_rsWG_Supervisors As ADODB.Recordset  'Contains the employee information of WG_SUPERVISOR
Private m_rsProjects As ADODB.Recordset
Private m_rsWGDemandLineSchedules As ADODB.Recordset
Private m_rsHoldWGDemandLineSchedules As ADODB.Recordset
Private m_rsWGBaseSched As ADODB.Recordset
Private m_rsWorkGroupAssociations As ADODB.Recordset
Private m_rsWGDemandLineAssignments As ADODB.Recordset

' Added for Demand Line Schedule
Private m_colCascadeLineIntervals As Collection
Private m_colCascadeEmpIntervals As Collection
Private m_strErrorField As String
Private m_lngUserNum As Long
Private Const CDEMANDSCHEDFIELDNAMES = "WG_NUM;START_DATE;LINE_NUM"
Private Const CEMPACTFIELDNAMES = "EMP_NUM;START_DATE;END_DATE"
Private Const CWGMAINTFIELDNAMES = "WG_NUM;START_DATE;LINE_NUM;EMP_NUM;DISTRIBUTION_ID;ACTIVITY_ID"
Private Const mblnAPPLYCASCADE = True

Private m_rsWGDistribution As ADODB.Recordset
Private m_rsWGActivities As ADODB.Recordset
Private m_rsGLActivities As ADODB.Recordset

Private Const m_datMaxEndDate As Date = "9999-12-31"

Private Const m_strCMODULENAME As String = "cbsWorkGroupTrans"


Public Function ProcessSaveWorkGroup( _
                    ByVal v_rsWorkGroupAttributes As ADODB.Recordset, _
                    ByVal v_rsWorkGroupSchedule As ADODB.Recordset, _
                    ByVal v_rsWorkGroupAssignments As ADODB.Recordset, _
                    ByVal v_rsWorkGroupDistributionList As ADODB.Recordset, _
                    ByVal v_rsWorkGroupDistributionActivities As ADODB.Recordset, _
                    ByVal v_UserId As String, _
                    ByVal v_blnConfirmConflicts As Boolean, _
                    ByRef r_rsErrors As ADODB.Recordset, _
                    ByRef r_rsWGAttributeUpdates As ADODB.Recordset, _
                    ByRef r_rsWGBaseScheduleUpdates As ADODB.Recordset, _
                    ByRef r_rsWGLineScheduleUpdates As ADODB.Recordset, _
                    ByRef r_rsWGAssignUpdates As ADODB.Recordset, _
                    ByRef r_rsWGAssignOtherUpdates As ADODB.Recordset, _
                    ByRef r_rsWGDistributionUpdates As ADODB.Recordset, _
                    ByRef r_rsWGActivityUpdates As ADODB.Recordset, _
                    ByRef r_rsWGGLActivityUpdates As ADODB.Recordset, _
                    ByRef r_rsCascadeConflicts As ADODB.Recordset, _
                    ByRef r_rsExceptionSPA As ADODB.Recordset, _
                    ByRef r_rsAvailScheduleSPA As ADODB.Recordset, _
                    ByRef r_rsEmpSchedDatesSPA As ADODB.Recordset, _
                    ByRef r_lngWorkGroup As Long, _
                    ByRef r_blnDemandWorkGroup As Boolean _
                ) As Boolean

    Dim l_objCtx As ObjectContext
    Dim l_dat1YearPrior As Date
    Dim l_blnError As Boolean
    Dim l_lngEmpNum As Long
    Dim l_blnConfirmPending As Boolean
    Dim l_varCascadeUpdates As Variant
    
    On Error GoTo ProcessSaveWorkGroup_EH
    
    m_strUserId = v_UserId
    m_lngUserNum = ParseUserID(v_UserId)
    
    Set l_objCtx = GetObjectContext
    If ARequestExist(v_rsWorkGroupAttributes, v_rsWorkGroupSchedule, v_rsWorkGroupAssignments, v_rsWorkGroupDistributionList, v_rsWorkGroupDistributionActivities) Then
        'continue
    Else
        ProcessSaveWorkGroup = True
        l_objCtx.SetComplete
        GoTo xit
    End If
    
    m_lngWorkGroup = r_lngWorkGroup

    Set m_rsErrors = CreateNewBusinessErrorSet

    If Not GetStateOfWorkGroup(v_rsWorkGroupAssignments, v_rsWorkGroupAttributes, v_rsWorkGroupSchedule) Then
        ProcessSaveWorkGroup = False
        l_objCtx.SetAbort
        GoTo xit
    End If
    
    If ProcessWorkGroupAttributes(v_rsWorkGroupAttributes) Then
        'continue
    Else
        ProcessSaveWorkGroup = False
        l_objCtx.SetAbort
        GoTo xit
    End If
    
    If r_lngWorkGroup = -1 Then
        SetNewWorkGroupNumber v_rsWorkGroupAttributes, v_rsWorkGroupSchedule, v_rsWorkGroupAssignments, v_rsWorkGroupDistributionList, v_rsWorkGroupDistributionActivities
    End If
    
    If m_blnDemandWorkGroup Then
        If ProcessWorkgroupDemandSchedules(v_rsWorkGroupSchedule) Then
            'continue
        Else
            ProcessSaveWorkGroup = False
            l_objCtx.SetAbort
            GoTo xit
        End If
    Else
        If ProcessWorkgroupBaseSchedules(v_rsWorkGroupSchedule) Then
            'continue
        Else
            ProcessSaveWorkGroup = False
            l_objCtx.SetAbort
            GoTo xit
        End If
    End If
    
    If ProcessWorkGroupEmployeeAssignments(v_rsWorkGroupAssignments) Then
        'continue
    Else
        ProcessSaveWorkGroup = False
        l_objCtx.SetAbort
        GoTo xit
    End If
    
        If ProcessWorkGroupDistribution(v_rsWorkGroupDistributionList) Then
        'continue
    Else
        ProcessSaveWorkGroup = False
        l_objCtx.SetAbort
        GoTo xit
    End If
    
    If ProcessWorkGroupActivity(v_rsWorkGroupDistributionActivities) Then
        'continue
    Else
        ProcessSaveWorkGroup = False
        l_objCtx.SetAbort
        GoTo xit
    End If
    
    
    If PostValidateWGSave(v_rsWorkGroupAttributes, _
                    v_rsWorkGroupSchedule, _
                    v_rsWorkGroupAssignments) Then
        'continue
    Else
        ProcessSaveWorkGroup = False
        l_objCtx.SetAbort
        GoTo xit
    End If
    
    Set r_rsEmpSchedDatesSPA = DetermineEmployeeIntervals
    If EmployeeTradesExist(r_rsEmpSchedDatesSPA, m_rsErrors) Then
        ProcessSaveWorkGroup = False
        l_objCtx.SetAbort
        GoTo xit
    End If
    
' BEGIN CASCADE
    l_lngEmpNum = m_lngUserNum
    Set r_rsWGBaseScheduleUpdates = m_rsWGBaseSched
'fred
    If m_blnWGLineSchedChg Then
        If Not m_rsWGDemandLineSchedules Is Nothing Then
            m_rsWGDemandLineSchedules.Filter = ""
        End If
        Set r_rsWGLineScheduleUpdates = m_rsWGDemandLineSchedules
    Else
        Set r_rsWGLineScheduleUpdates = Nothing
    End If
    Set r_rsWGAssignUpdates = m_rsWG_Assignments
    Set r_rsWGAssignOtherUpdates = m_rsWG_PAIOWG
    If ProcessCascade(l_lngEmpNum, _
                v_blnConfirmConflicts, l_blnConfirmPending, l_varCascadeUpdates, _
                r_rsWGBaseScheduleUpdates, r_rsWGLineScheduleUpdates, _
                r_rsWGAssignUpdates, r_rsWGAssignOtherUpdates, _
                r_rsCascadeConflicts, r_rsExceptionSPA, r_rsAvailScheduleSPA, _
                r_rsEmpSchedDatesSPA, m_rsErrors) Then
        If l_blnConfirmPending Then
            r_lngWorkGroup = m_lngWorkGroup
            r_blnDemandWorkGroup = m_blnDemandWorkGroup
            Set r_rsWGAttributeUpdates = m_rsWorkGroupAttributes
            Set r_rsWGDistributionUpdates = m_rsWGDistribution
            Set r_rsWGActivityUpdates = m_rsWGActivities
            Set r_rsWGGLActivityUpdates = m_rsGLActivities
            ProcessSaveWorkGroup = True
            l_objCtx.SetComplete
            GoTo xit
        Else
            Set r_rsWGBaseScheduleUpdates = Nothing
            Set r_rsWGLineScheduleUpdates = Nothing
            Set r_rsWGAssignUpdates = Nothing
            Set r_rsWGAssignOtherUpdates = Nothing
            Set r_rsCascadeConflicts = Nothing
            Set r_rsExceptionSPA = Nothing
            Set r_rsAvailScheduleSPA = Nothing
        End If
    Else
        Set r_rsWGBaseScheduleUpdates = Nothing
        Set r_rsWGLineScheduleUpdates = Nothing
        Set r_rsWGAssignUpdates = Nothing
        Set r_rsWGAssignOtherUpdates = Nothing
        Set r_rsCascadeConflicts = Nothing
        Set r_rsExceptionSPA = Nothing
        Set r_rsAvailScheduleSPA = Nothing
        Set r_rsEmpSchedDatesSPA = Nothing
        ProcessSaveWorkGroup = False
        l_objCtx.SetAbort
        GoTo xit
    End If
    
    If Not ProcessScheduleChange(r_rsEmpSchedDatesSPA, m_rsErrors, _
            m_rsWGBaseSched, m_rsWGDemandLineSchedules, m_rsWG_Assignments, m_rsWG_PAIOWG) Then
        l_objCtx.SetAbort
        Set r_rsEmpSchedDatesSPA = Nothing
        ProcessSaveWorkGroup = False
        GoTo xit
    Else
        Set r_rsEmpSchedDatesSPA = Nothing
    End If
' END CASCADE
    
    If RecordsExist(v_rsWorkGroupAttributes) Then
        'save work group attributes
        If SaveWGAttributes() Then
            'continue
        Else
            ProcessSaveWorkGroup = False
            l_objCtx.SetAbort
            GoTo xit
        End If
    End If
    
    If RecordsExist(v_rsWorkGroupSchedule) Then
        If m_blnDemandWorkGroup Then
        'save work group demand schedules
            If SaveWGDemandSchedules() Then
                'continue
            Else
                ProcessSaveWorkGroup = False
                l_objCtx.SetAbort
                GoTo xit
            End If
        Else
            If SaveWGBaseSchedules() Then
                'continue
            Else
                ProcessSaveWorkGroup = False
                l_objCtx.SetAbort
                GoTo xit
            End If
        End If
    End If
    
    If RecordsExist(v_rsWorkGroupAssignments) Then
        'save work group attributes
        If SaveWGAssignments() Then
            'continue
        Else
            ProcessSaveWorkGroup = False
            l_objCtx.SetAbort
            GoTo xit
        End If
    End If
    
    If RecordsExist(v_rsWorkGroupDistributionList) Then
        'save work group distribution
        If SaveWGDistribution() Then
            'continue
        Else
            ProcessSaveWorkGroup = False
            l_objCtx.SetAbort
            GoTo xit
        End If
    End If
    
    If RecordsExist(v_rsWorkGroupDistributionActivities) Then
        'save work group activities
        If SaveWGActivities() Then
            'continue
        Else
            ProcessSaveWorkGroup = False
            l_objCtx.SetAbort
            GoTo xit
        End If
    End If
    
    If RecordsExist(m_rsGLActivities) Then
        'save GL activities
        If SaveGLActivities() Then
            'continue
        Else
            ProcessSaveWorkGroup = False
            l_objCtx.SetAbort
            GoTo xit
        End If
    End If
    r_lngWorkGroup = m_lngWorkGroup
    ProcessSaveWorkGroup = True
    l_objCtx.SetComplete
     
    '//-----invalidate the spm cache-----//
    InvalidateWorkGroupCache r_lngWorkGroup, v_rsWorkGroupAttributes
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
    Set l_objCtx = Nothing
    Exit Function
       
    
ProcessSaveWorkGroup_EH:
    l_objCtx.SetAbort
    Set r_rsWGAttributeUpdates = Nothing
    Set r_rsWGBaseScheduleUpdates = Nothing
    Set r_rsWGLineScheduleUpdates = Nothing
    Set r_rsWGAssignUpdates = Nothing
    Set r_rsWGAssignOtherUpdates = Nothing
    Set r_rsWGDistributionUpdates = Nothing
    Set r_rsWGActivityUpdates = Nothing
    Set r_rsWGGLActivityUpdates = Nothing
    Set r_rsCascadeConflicts = Nothing
    Set r_rsExceptionSPA = Nothing
    Set r_rsAvailScheduleSPA = Nothing
    Set r_rsEmpSchedDatesSPA = Nothing
    Set r_rsErrors = Nothing
    SetGlobalObjectsToNothing
    Set l_objCtx = Nothing
'    Clear the cascade error arguments since the client does not format them.
    If Err.Number = trErrAppCascadeCommitConflict Then
    'Rick Fred
        Err.Description = Err.Description & " " & Err.Source
        Err.Source = Join(Array(CStr(TrType_WGM), "", "", "", ""), ";")
    End If
    TrsRaiseError m_strCMODULENAME, "ProcessSaveWorkGroup", , , , , True
End Function


Public Function ResumeProcessSaveWorkGroup( _
                    ByRef r_lngWorkGroup As Long, _
                    ByVal v_blnDemandWorkGroup As Boolean, _
                    ByVal v_rsWGAttributeUpdates As ADODB.Recordset, _
                    ByVal v_rsWGBaseScheduleUpdates As ADODB.Recordset, _
                    ByVal v_rsWGLineScheduleUpdates As ADODB.Recordset, _
                    ByVal v_rsWGAssignUpdates As ADODB.Recordset, _
                    ByVal v_rsWGAssignOtherUpdates As ADODB.Recordset, _
                    ByVal v_rsWGDistributionUpdates As ADODB.Recordset, _
                    ByVal v_rsWGActivityUpdates As ADODB.Recordset, _
                    ByVal v_rsWGGLActivityUpdates As ADODB.Recordset, _
                    ByVal v_rsCascadeConflicts As ADODB.Recordset, _
                    ByVal v_rsExceptionSPA As ADODB.Recordset, _
                    ByVal v_rsAvailScheduleSPA As ADODB.Recordset, _
                    ByVal v_rsEmpSchedDatesSPA As ADODB.Recordset, _
                    ByRef r_rsErrors As ADODB.Recordset, _
                    ByVal v_UserId As String) _
                As Boolean
' Do not expect any lock contention between cascade and workgroup schedule updates, so
' request cascade to apply the updates instead of returning them, mblnAPPLYCASCADE.
    Dim l_objCtx As ObjectContext
    Dim l_rsErrors As ADODB.Recordset
    Dim bError As Boolean
    Dim l_cbsCascade As busCascade.cbsCascade
    Dim l_objDBSchDmdTrans As dbSchedDemand.cdSchedDemandTrans
    Dim l_varCascadeUpdates As Variant
    Dim l_rsCascadeErrors As ADODB.Recordset

On Error GoTo eh
    m_strUserId = v_UserId
    m_lngUserNum = ParseUserID(v_UserId)
    m_lngWorkGroup = r_lngWorkGroup
    m_blnDemandWorkGroup = v_blnDemandWorkGroup
    
    Set l_objCtx = GetObjectContext

    If r_rsErrors Is Nothing Then
        Set r_rsErrors = CreateNewBusinessErrorSet
    End If
    Set m_rsErrors = r_rsErrors
    
    If Not v_rsWGLineScheduleUpdates Is Nothing Then
        v_rsWGLineScheduleUpdates.Filter = ""
        Set m_rsWGDemandLineSchedules = v_rsWGLineScheduleUpdates
    End If
              
    If Not ProcessScheduleChange(v_rsEmpSchedDatesSPA, r_rsErrors, _
            v_rsWGBaseScheduleUpdates, v_rsWGLineScheduleUpdates, _
            v_rsWGAssignUpdates, v_rsWGAssignOtherUpdates) Then
        l_objCtx.SetAbort
        ResumeProcessSaveWorkGroup = False
        GoTo xit
    End If
       
'    set and reset filters for all module variables
    If Not v_rsWGAttributeUpdates Is Nothing Then
        v_rsWGAttributeUpdates.Filter = adFilterPendingRecords
        Set m_rsWorkGroupAttributes = v_rsWGAttributeUpdates
    End If
       
    If Not v_rsWGBaseScheduleUpdates Is Nothing Then
        v_rsWGBaseScheduleUpdates.Filter = adFilterPendingRecords
        Set m_rsWGBaseSched = v_rsWGBaseScheduleUpdates
    End If
    
    If Not v_rsWGLineScheduleUpdates Is Nothing Then
        v_rsWGLineScheduleUpdates.Filter = adFilterPendingRecords
        Set m_rsWGDemandLineSchedules = v_rsWGLineScheduleUpdates
    End If
    
    If Not v_rsWGAssignUpdates Is Nothing Then
        v_rsWGAssignUpdates.Filter = adFilterPendingRecords
        Set m_rsWG_Assignments = v_rsWGAssignUpdates
    End If
    
    If Not v_rsWGAssignOtherUpdates Is Nothing Then
        v_rsWGAssignOtherUpdates.Filter = adFilterPendingRecords
        Set m_rsWG_PAIOWG = v_rsWGAssignOtherUpdates
    End If
    
    If Not v_rsWGDistributionUpdates Is Nothing Then
        v_rsWGDistributionUpdates.Filter = adFilterPendingRecords
        Set m_rsWGDistribution = v_rsWGDistributionUpdates
    End If
    
    If Not v_rsWGActivityUpdates Is Nothing Then
        v_rsWGActivityUpdates.Filter = adFilterPendingRecords
        Set m_rsWGActivities = v_rsWGActivityUpdates
    End If
    
    If Not v_rsWGGLActivityUpdates Is Nothing Then
        v_rsWGGLActivityUpdates.Filter = adFilterPendingRecords
        Set m_rsGLActivities = v_rsWGGLActivityUpdates
    End If

    Set l_cbsCascade = CreateObject("busCascade.cbsCascade")
    l_cbsCascade.UpdateExceptions _
        v_rsCascadeConflicts, v_rsExceptionSPA, v_rsAvailScheduleSPA, _
        mblnAPPLYCASCADE, l_varCascadeUpdates, l_rsCascadeErrors, m_lngUserNum
    Set l_cbsCascade = Nothing
    
    AppendCascadeErrors l_rsCascadeErrors, r_rsErrors
    Set l_rsCascadeErrors = Nothing
    
    If RecordsExist(m_rsWorkGroupAttributes) Then
        If SaveWGAttributes() Then
            'continue
        Else
            ResumeProcessSaveWorkGroup = False
            l_objCtx.SetAbort
            GoTo xit
        End If
    End If
    
    If m_blnDemandWorkGroup Then
        If RecordsExist(m_rsWGDemandLineSchedules) Then
            If SaveWGDemandSchedules() Then
                'continue
            Else
                ResumeProcessSaveWorkGroup = False
                l_objCtx.SetAbort
                GoTo xit
            End If
        End If
    ElseIf RecordsExist(m_rsWGBaseSched) Then
        If SaveWGBaseSchedules() Then
                'continue
        Else
            ResumeProcessSaveWorkGroup = False
            l_objCtx.SetAbort
            GoTo xit
        End If
    End If
    
    If RecordsExist(m_rsWG_Assignments) Or _
            RecordsExist(m_rsWG_PAIOWG) Then
        If SaveWGAssignments() Then
            'continue
        Else
            ResumeProcessSaveWorkGroup = False
            l_objCtx.SetAbort
            GoTo xit
        End If
    End If
    
    If RecordsExist(m_rsWGDistribution) Then
        If SaveWGDistribution() Then
            'continue
        Else
            ResumeProcessSaveWorkGroup = False
            l_objCtx.SetAbort
            GoTo xit
        End If
    End If
    
    If RecordsExist(m_rsWGActivities) Then
        If SaveWGActivities() Then
            'continue
        Else
            ResumeProcessSaveWorkGroup = False
            l_objCtx.SetAbort
            GoTo xit
        End If
    End If
    
    If RecordsExist(m_rsGLActivities) Then
        If SaveGLActivities() Then
            'continue
        Else
            ResumeProcessSaveWorkGroup = False
            l_objCtx.SetAbort
            GoTo xit
        End If
    End If
    
    r_lngWorkGroup = m_lngWorkGroup
    ResumeProcessSaveWorkGroup = True
    l_objCtx.SetComplete

    '//-----invalidate the spm cache-----//
    InvalidateWorkGroupCache r_lngWorkGroup, v_rsWGAttributeUpdates
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
        'Rick Fred
        Err.Description = Err.Description & " " & Err.Source
        Err.Source = Join(Array(CStr(TrType_WGM), "", "", "", ""), ";")
    End If
    TrsRaiseError m_strCMODULENAME, "ResumeProcessSaveWorkGroup", , , , , True
End Function

Private Sub InvalidateWorkGroupCache(ByVal v_lngWorkGroup As Long, _
                                    ByVal v_rsWorkGroupAttributes As ADODB.Recordset)
    Dim l_objdbWkGp As dbWorkGroup.cdWorkGroup
    
    '//-----invalidate the spm cache-----//
    If RecordsExist(v_rsWorkGroupAttributes) Then
        Set l_objdbWkGp = CreateObject("dbWorkGroup.cdWorkGroup")
        l_objdbWkGp.InvalidateWorkGroupCache v_lngWorkGroup
        Set l_objdbWkGp = Nothing
    End If
End Sub

Private Function DetermineEmployeeIntervals() As ADODB.Recordset
    Dim l_colLineDateIntervals As Collection
    Dim l_colEmpDateIntervals As Collection

On Error GoTo eh

    Set l_colLineDateIntervals = ConsolidatedCascadeWGLineIntervals
    DetermineEmployeeAssignments l_colLineDateIntervals
    Set l_colEmpDateIntervals = ConsolidatedCascadeEmployeeIntervals
    
    Set DetermineEmployeeIntervals = GetEmpSchedDates(l_colEmpDateIntervals)
    Set l_colLineDateIntervals = Nothing
    Set l_colLineDateIntervals = Nothing
    Exit Function
    
eh:
    Set l_colLineDateIntervals = Nothing
    Set l_colLineDateIntervals = Nothing
    TrsRaiseError m_strCMODULENAME, "DetermineEmployeeIntervals"
End Function


Private Function ProcessScheduleChange( _
                    ByRef r_rsEmpSchedDates As ADODB.Recordset, _
                    ByRef r_rsErrors As ADODB.Recordset, _
                    ByRef r_rsWGBaseSched As ADODB.Recordset, _
                    ByRef r_rsWGLineSched As ADODB.Recordset, _
                    ByRef r_rsWGAssignments As ADODB.Recordset, _
                    ByRef r_rsWGOtherAssign As ADODB.Recordset _
                ) As Boolean
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
    ScheduleChangePaymentExists r_rsEmpSchedDates, r_rsErrors, _
        r_rsWGBaseSched, r_rsWGLineSched, r_rsWGAssignments, r_rsWGOtherAssign
    
    ' update the emp_mod timestamp
    ' this will create a lock that time collection must wait on before saving actual time.
    Set l_cdCascade = CreateObject("dbCascade.cdCascade")
    l_cdCascade.SetEmployeeScheduleModified l_varEmps
    Set l_cdCascade = Nothing
    
    If ActualTimeExists(r_rsEmpSchedDates, r_rsErrors, r_rsWGLineSched) Then
        ProcessScheduleChange = False
        Exit Function
    End If

' 10/7/02 RSG Replaced with LineTradeExists and EmployeeTradeExists,
' both are workgroup maintenance validations, not cascade.
'    If TradeExists(r_rsEmpSchedDates, r_rsErrors) Then
'        ProcessScheduleChange = False
'        Exit Function
'    End If

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
            TrType_WGM, , , , _
            CStr(m_lngWorkGroup), _
            r_rsCascadeErrors!KEY_FIELD2, _
            "0", _
            r_rsCascadeErrors!KEY_FIELD1
        r_rsCascadeErrors.MoveNext
    Loop
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "AppendCascadeErrors"
End Function


Private Function ProcessWorkGroupEmployeeAssignments( _
    ByVal v_rsWorkGroupAssignments As ADODB.Recordset _
    ) As Boolean
    
'This function will process the save from Assign/Remove employees
' from Group Maintenance
'The function will ensure that the request is valid (PreValidate)
'The function will update all assignments in the request, as well
' as adjacent primary assignemnts affected by the request

' Finally, the function will prepare a request set needed in the
' post validation routines.
' Finally, the function will prepare a CascadePrep request and
' call Cascadeprep, GetNewSchedule (new function), and Cascade

    
    Dim l_blnIsValid As Boolean

    
    On Error GoTo ProcessWorkGroupEmployeeAssignments_EH
     
    If Not RecordsExist(v_rsWorkGroupAssignments) Then
        ProcessWorkGroupEmployeeAssignments = True
        Exit Function
    End If
        
    l_blnIsValid = PreValidateAssignmentRequests(v_rsWorkGroupAssignments)
    If Not l_blnIsValid Then
        ProcessWorkGroupEmployeeAssignments = False
        Exit Function
    End If

    v_rsWorkGroupAssignments.MoveFirst
    While Not v_rsWorkGroupAssignments.EOF
        If v_rsWorkGroupAssignments!RECORD_STATUS = "D" Then
            v_rsWorkGroupAssignments!RECORD_STATUS = "Z"
            v_rsWorkGroupAssignments.Update
        End If
        v_rsWorkGroupAssignments.MoveNext
    Wend
    
    v_rsWorkGroupAssignments.Sort = "RECORD_STATUS DESC"
    
    ApplyAssigmentRequests v_rsWorkGroupAssignments
        
    'Place all primary assigments in one recordset for
    'validate and request for Cascade
    Set m_rsWG_ALL_PA = BuildPrimaryRecordset()
    
    l_blnIsValid = PostValidateAssignmentRequests(v_rsWorkGroupAssignments)
    If Not l_blnIsValid Then
        ProcessWorkGroupEmployeeAssignments = False
        Exit Function
    End If

    ProcessWorkGroupEmployeeAssignments = True
    Exit Function
   
       
    
ProcessWorkGroupEmployeeAssignments_EH:
    
    TrsRaiseError m_strCMODULENAME, "ProcessWorkGroupEmployeeAssignments"
    
    
End Function

Private Function PreValidateAssignmentRequests( _
            ByVal v_rsWorkGroupAssignments As ADODB.Recordset _
            ) As Boolean
            
    Dim l_blnIsValid As Boolean
    Dim l_strRecStatus As String
    Dim l_blnWriteError As Boolean
    
    On Error GoTo eh
    
    l_blnIsValid = True
    v_rsWorkGroupAssignments.MoveFirst
    While Not v_rsWorkGroupAssignments.EOF
        l_strRecStatus = v_rsWorkGroupAssignments!RECORD_STATUS
        Select Case l_strRecStatus
            Case "C", "I", "D"
            Case Else
                l_blnIsValid = False
                l_blnWriteError = AppendWGAssignmentsBRV(v_rsWorkGroupAssignments, 30000, mstrCBUSINESSERR, "Invalid Request. Please try again")
        End Select

        'Is Employee Assigning themseleves to a work group
        If IsEmployeeAssigningSelf(v_rsWorkGroupAssignments) Then
            l_blnIsValid = False
        End If
        
        If Not IsValidEffectiveDates(v_rsWorkGroupAssignments) Then
            l_blnIsValid = False
        End If
    
        v_rsWorkGroupAssignments.MoveNext
    Wend
      
    PreValidateAssignmentRequests = l_blnIsValid
    
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "PreValidateAssignmentRequests"

End Function
Private Function IsEmployeeAssigningSelf( _
            ByVal v_rsWorkGroupAssignments As ADODB.Recordset _
            ) As Boolean
            
    Dim l_blnWriteError As Long
    
    On Error GoTo eh

    If v_rsWorkGroupAssignments!EMP_NUM = m_strUserId Then
        IsEmployeeAssigningSelf = True
        l_blnWriteError = AppendWGAssignmentsBRV(v_rsWorkGroupAssignments, 30000, mstrCBUSINESSERR, "Can not assign yourself to a work group")
    Else
        IsEmployeeAssigningSelf = False
    End If
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "IsEmployeeAssigningSelf"
End Function

Private Function IsValidEffectiveDates( _
            ByVal v_rsWorkGroupAssignments As ADODB.Recordset _
            ) As Boolean
            
    Dim l_blnWriteError As Boolean
    Dim l_datStartDate As Date
    Dim l_datEndDate As Date
    Dim l_datWGStartDate As Date
    Dim l_datWGEndDate As Date
    Dim l_strErrorMessage As String
            
    On Error GoTo eh
    
    IsValidEffectiveDates = True
    If IsDate(v_rsWorkGroupAssignments!EFF_START_DATE) Then
    Else
        IsValidEffectiveDates = False
        l_blnWriteError = AppendWGAssignmentsBRV(v_rsWorkGroupAssignments, 30000, mstrCBUSINESSERR, "The effective start date is an invalid date")
        Exit Function
    End If
    
    l_datStartDate = GetEffectiveStartDate(v_rsWorkGroupAssignments)
    
'    MsgBox "IsValidEffectiveDates 1", , "JIMB"
    
    l_datEndDate = GetEffectiveEndDate(v_rsWorkGroupAssignments)
    
'    MsgBox "IsValidEffectiveDates 2", , "JIMB"
    If IsDate(l_datEndDate) Then
    Else
        IsValidEffectiveDates = False
        l_blnWriteError = AppendWGAssignmentsBRV(v_rsWorkGroupAssignments, 30000, mstrCBUSINESSERR, "The effective end date is an invalid date")
        Exit Function
    End If
    
    If l_datStartDate > l_datEndDate Then
        IsValidEffectiveDates = False
        l_strErrorMessage = "The assignment start date (" & l_datStartDate & ") can not be greater than the assignment end date (" & l_datEndDate & ")"
        l_blnWriteError = AppendWGAssignmentsBRV(v_rsWorkGroupAssignments, 30000, mstrCBUSINESSERR, l_strErrorMessage)
        Exit Function
    End If
    
'    MsgBox "IsValidEffectiveDates 3", , "JIMB"
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "IsValidEffectiveDates", , , , , True
            
End Function


Private Function AppendWGAttributesBRV( _
            ByVal v_rsWorkGroupAttributes As ADODB.Recordset, _
            ByVal v_lngErrorNumber As Long, _
            ByVal v_strErrorType As String, _
            ByVal v_strErrorDesc As String _
            ) As Boolean
            
    Dim l_bolError As Boolean
    Dim lngKeyWGNum As Long
    Dim l_datKeyStartDate As Date
    Dim l_strKeyCreatedTime As String
    
    On Error GoTo AppendWGAttributesBRV_EH
    
    If Not IsNull(v_rsWorkGroupAttributes!KEY_WG_NUM) And _
        Not IsEmpty(v_rsWorkGroupAttributes!KEY_WG_NUM) Then
        lngKeyWGNum = v_rsWorkGroupAttributes!KEY_WG_NUM
    Else
        If Not IsNull(v_rsWorkGroupAttributes!WG_NUM) And _
            Not IsEmpty(v_rsWorkGroupAttributes!WG_NUM) Then
            lngKeyWGNum = v_rsWorkGroupAttributes!WG_NUM
        Else
            lngKeyWGNum = -1
        End If
    End If
    If Not IsNull(v_rsWorkGroupAttributes!KEY_EFF_START_DATE) And _
        Not IsEmpty(v_rsWorkGroupAttributes!KEY_EFF_START_DATE) Then
        l_datKeyStartDate = v_rsWorkGroupAttributes!KEY_EFF_START_DATE
    Else
        If Not IsNull(v_rsWorkGroupAttributes!EFF_START_DATE) And _
            Not IsEmpty(v_rsWorkGroupAttributes!EFF_START_DATE) Then
            l_datKeyStartDate = v_rsWorkGroupAttributes!EFF_START_DATE
        Else
            l_datKeyStartDate = Date
        End If
    End If
    If IsNull(v_rsWorkGroupAttributes!CREATED_TIMESTAMP) Then
        l_strKeyCreatedTime = ""
    Else
        l_strKeyCreatedTime = v_rsWorkGroupAttributes!CREATED_TIMESTAMP
    End If
    
    l_bolError = AppendBusinessError(m_rsErrors, _
            v_lngErrorNumber, _
            v_strErrorType, _
            v_strErrorDesc, _
            m_strCMODULENAME, _
            TrType_WGM, , , "WG_NUM;EFF_START_DATE", _
            str(lngKeyWGNum), _
            str(l_datKeyStartDate), _
            "0", _
            "0", _
            "", _
            "")

    AppendWGAttributesBRV = l_bolError
    Exit Function
    
AppendWGAttributesBRV_EH:
    TrsRaiseError m_strCMODULENAME, "AppendWGAttributesBRV"

End Function

Private Function AppendWGAssignmentsBRV( _
            ByVal v_rsWorkGroupAssignments As ADODB.Recordset, _
            ByVal v_lngErrorNumber As Long, _
            ByVal v_strErrorType As String, _
            ByVal v_strErrorDesc As String, _
            Optional ByVal v_blnHasKey As Boolean = True _
            ) As Boolean
            
    Dim l_bolError As Boolean
    Dim l_strPrimaryInd As String
    Dim l_datStartDate As Date
    Dim l_datEndDate As Date
    Dim lngEmpNum As Long
    Dim lngKeyWGNum As Long
    Dim lngKeyEmpNum As Long
    Dim l_datKeyStartDate As Date
    Dim l_strKeyCreatedTime As String
    Dim l_strStartDate As String
    
    On Error GoTo AppendWGAssignmentsBRV_EH
    
    If v_blnHasKey Then
        lngKeyWGNum = v_rsWorkGroupAssignments!KEY_WG_NUM
        lngKeyEmpNum = v_rsWorkGroupAssignments!KEY_EMP_NUM
        l_datKeyStartDate = v_rsWorkGroupAssignments!KEY_EFF_START_DATE
    Else
        lngKeyWGNum = v_rsWorkGroupAssignments!WG_NUM
        lngKeyEmpNum = v_rsWorkGroupAssignments!EMP_NUM
        If Not IsNull(v_rsWorkGroupAssignments!EFF_START_DATE) _
            And Not IsEmpty(v_rsWorkGroupAssignments!EFF_START_DATE) Then
            l_datKeyStartDate = v_rsWorkGroupAssignments!EFF_START_DATE
        Else
            l_datKeyStartDate = 0
            l_strStartDate = ""
        End If
    End If
    
    If IsNull(v_rsWorkGroupAssignments!CREATED_TIMESTAMP) Then
        l_strKeyCreatedTime = ""
    Else
        l_strKeyCreatedTime = v_rsWorkGroupAssignments!CREATED_TIMESTAMP
    End If
    
    lngEmpNum = v_rsWorkGroupAssignments!EMP_NUM
    If IsNull(v_rsWorkGroupAssignments!EFF_START_DATE) _
        Or IsEmpty(v_rsWorkGroupAssignments!EFF_START_DATE) Then
        l_strStartDate = ""
    Else
        l_strStartDate = Format(v_rsWorkGroupAssignments!EFF_START_DATE, "mm/dd/yyyy")
    End If
    l_datStartDate = GetEffectiveStartDate(v_rsWorkGroupAssignments)
    l_datEndDate = GetEffectiveEndDate(v_rsWorkGroupAssignments)
    
    If IsEmpty(v_rsWorkGroupAssignments!PRIMARY_IND) Or IsNull(v_rsWorkGroupAssignments!PRIMARY_IND) Then
        l_strPrimaryInd = "S"
    Else
        l_strPrimaryInd = v_rsWorkGroupAssignments!PRIMARY_IND
    End If
    
    
    l_bolError = AppendBusinessError(m_rsErrors, _
            v_lngErrorNumber, _
            v_strErrorType, _
            v_strErrorDesc, _
            m_strCMODULENAME, _
            TrType_WGM_ASGMT, , , "WG_NUM;START_DATE;;EMP_NUM", _
            str(lngKeyWGNum), _
            l_strStartDate, _
            "0", _
            str(lngKeyEmpNum), _
            "", _
            "")
            
    AppendWGAssignmentsBRV = l_bolError
    Exit Function
AppendWGAssignmentsBRV_EH:
    TrsRaiseError m_strCMODULENAME, "AppendWGAssignmentsBRV"

End Function


Private Sub AppendWGAssignmentDeleteMessage( _
                    ByVal v_lngErrorNum As Long, _
                    ByVal v_strErrorType As String, _
                    ByVal v_strErrorDesc As String, _
                    ByVal v_lngExtWgNum As Long, _
                    ByVal v_datExtStartDate As Date, _
                    ByVal v_datExtEndDate As Date, _
                    ByVal v_lngWgNum As Long, _
                    ByVal v_datStartDate As Date, _
                    ByVal v_lngEmpNum As Long)
    Dim l_strErrDesc As String
    Dim l_strEmpName As String
    Dim l_cdCodeTables As dbCodeTables.cdCodeTables
    
On Error GoTo eh
    Set l_cdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    l_strEmpName = l_cdCodeTables.EmployeeName(v_lngEmpNum)
    Set l_cdCodeTables = Nothing
    
    If v_strErrorDesc = "" Then
        If v_datExtStartDate = 0 Then
            l_strErrDesc = _
                "Employee " & v_lngEmpNum & "[" & l_strEmpName & "]" & " previous workgroup assignment(" & _
                v_lngExtWgNum & ") has been extended forward to " & Format(v_datExtEndDate, "mm/dd/yyyy")
        Else
            l_strErrDesc = _
                "Employee " & v_lngEmpNum & "[" & l_strEmpName & "]" & " next workgroup assignment(" & _
                v_lngExtWgNum & ") has been extended back to " & Format(v_datExtStartDate, "mm/dd/yyyy")
        End If
    Else
        l_strErrDesc = v_strErrorDesc
    End If
    
    AppendBusinessError _
        m_rsErrors, _
        v_lngErrorNum, _
        v_strErrorType, _
        l_strErrDesc, _
        m_strCMODULENAME, _
        TrType_WGM_ASGMT, , , "WG_NUM;START_DATE;LINE_NUM;EMP_NUM", _
        CStr(v_lngWgNum), _
        Format(v_datStartDate, "mm/dd/yyyy"), _
        "0", _
        CStr(v_lngEmpNum)
            
    Exit Sub
    
eh:
    GetObjectContext.SetAbort
    Set l_cdCodeTables = Nothing
    TrsRaiseError m_strCMODULENAME, "AppendWGAssignmentDeleteMessage"
End Sub


Private Function GetWorkGroupInfo( _
            ByVal v_lngWorkGroup As Long _
            ) As ADODB.Recordset

    Dim l_rs As ADODB.Recordset
    Dim l_obj As dbWorkGroup.cdWorkGroup

On Error GoTo eh


    Set l_obj = CreateObject("dbWorkGroup.cdWorkGroup")
    Set l_rs = l_obj.GetWorkgroupForUpdate(v_lngWorkGroup)
    Set GetWorkGroupInfo = l_rs
    Exit Function
    
eh:
    GetObjectContext.SetAbort
  TrsRaiseError m_strCMODULENAME, "GetWorkGroupInfo"

End Function

Private Function GetWorkGroupStartDate() As Date

    On Error GoTo eh
    
    GetWorkGroupStartDate = m_rsWorkGroupAttributes!EFF_START_DATE
    
    Exit Function
    
eh:
    
  TrsRaiseError m_strCMODULENAME, "GetWorkGroupStartDate"
    
End Function

Private Function GetWorkGroupEndDate() As Date

    On Error GoTo eh
    
    GetWorkGroupEndDate = m_rsWorkGroupAttributes!EFF_END_DATE
    Exit Function
    
eh:
    
  TrsRaiseError m_strCMODULENAME, "GetWorkGroupEndDate"
    
End Function

Private Function GetEffectiveStartDate( _
            ByVal v_rs As ADODB.Recordset _
            ) As Date

    On Error GoTo eh

    If IsEmpty(v_rs!EFF_START_DATE) Or IsNull(v_rs!EFF_START_DATE) Then
        GetEffectiveStartDate = Date 'Current System Date
    Else
        GetEffectiveStartDate = v_rs!EFF_START_DATE
    End If
    Exit Function
    
eh:
    
  TrsRaiseError m_strCMODULENAME, "GetEffectiveStartDate"
    
End Function

Private Function GetEffectiveEndDate( _
            ByVal v_rs As ADODB.Recordset _
            ) As Date
            
On Error GoTo eh

    If IsNull(v_rs!EFF_END_DATE) Or IsEmpty(v_rs!EFF_END_DATE) Then
        GetEffectiveEndDate = m_datMaxEndDate
    Else
        If IsDate(v_rs!EFF_END_DATE) Then
            GetEffectiveEndDate = v_rs!EFF_END_DATE
        Else
            GetEffectiveEndDate = m_datMaxEndDate
        End If
    End If
    
    Exit Function
    
eh:
    
  TrsRaiseError m_strCMODULENAME, "GetEffectiveEndDate", , , , , True
    
End Function
Private Function GetAssignmentTempOtBankInd( _
            ByVal v_rsWorkGroupAssignments As ADODB.Recordset _
            ) As String
    
    Dim l_strInd As String
    
On Error GoTo eh

    If IsEmpty(v_rsWorkGroupAssignments!TEMP_OT_BANK_IND) Or IsNull(v_rsWorkGroupAssignments!TEMP_OT_BANK_IND) Then
        GetAssignmentTempOtBankInd = "N"
    Else
        l_strInd = UCase(v_rsWorkGroupAssignments!TEMP_OT_BANK_IND)
        If l_strInd = "Y" Then
            GetAssignmentTempOtBankInd = "Y"
        Else
            GetAssignmentTempOtBankInd = "N"
        End If
    End If
    
    Exit Function
    
eh:
    
  TrsRaiseError m_strCMODULENAME, "GetAssignmentTempOtBankInd"
    
End Function
'Private Function GetAssignmentRequiredInd( _
'            ByVal v_rsWorkGroupAssignments As ADODB.Recordset _
'            ) As String
'
'    Dim l_strInd As String
'
'On Error GoTo eh
'
'    If IsEmpty(v_rsWorkGroupAssignments!REQUIRED_PERSON) Or IsNull(v_rsWorkGroupAssignments!REQUIRED_PERSON) Then
'        GetAssignmentRequiredInd = "N"
'    Else
'        l_strInd = UCase(v_rsWorkGroupAssignments!REQUIRED_PERSON)
'        If l_strInd = "Y" Then
'            GetAssignmentRequiredInd = "Y"
'        Else
'            GetAssignmentRequiredInd = "N"
'        End If
'    End If
'
'    Exit Function
'
'eh:
'
'  TrsRaiseError m_strCMODULENAME, "GetAssignmentRequiredInd"
'
'End Function

Private Function GetEmployeeAssignmentsInWorkGroup( _
        ByVal v_lngWorkGroup As Long, _
        ByVal v_datDate As Date _
        ) As ADODB.Recordset
    
    Dim l_objWkGp As dbWorkGroup.cdWorkGroup
    Dim l_objRS As Recordset
    
    On Error GoTo GetEmployeeAssignmentsInWorkGroup_EH
    
    Set l_objWkGp = CreateObject("dbWorkGroup.cdWorkGroup")

    Set l_objRS = l_objWkGp.GetWorkGroupAssignments(v_lngWorkGroup, v_datDate)
    Set GetEmployeeAssignmentsInWorkGroup = l_objRS
    Set l_objRS = Nothing
    Set l_objWkGp = Nothing
    GetObjectContext.SetComplete

    Exit Function

GetEmployeeAssignmentsInWorkGroup_EH:
    GetObjectContext.SetAbort
    Set l_objRS = Nothing
    Set l_objWkGp = Nothing
    
    TrsRaiseError m_strCMODULENAME, "GetEmployeeAssignmentsInWorkGroup", , , , , True

End Function


Private Function GetEmployeePrimaryAssignments( _
            ByVal v_datDate As Date _
            ) As ADODB.Recordset

    Dim l_aryEmployees As Variant
    Dim l_rs As ADODB.Recordset
    Dim l_obj As dbWorkGroup.cdWorkGroup
    
    On Error GoTo eh

    If m_colEmployees.Count = 0 Then
        Set GetEmployeePrimaryAssignments = Nothing
        Exit Function
    End If

    l_aryEmployees = ArrayFromCollection(m_colEmployees)
    Set l_obj = CreateObject("dbWorkGroup.cdWorkGroup")
    Set l_rs = l_obj.GetEmployeeWorkGroupAssignments(l_aryEmployees, v_datDate, m_lngWorkGroup, True)
    Set GetEmployeePrimaryAssignments = l_rs
    
    Exit Function
    
eh:
    GetObjectContext.SetAbort
  TrsRaiseError m_strCMODULENAME, "GetEmployeePrimaryAssignments"


End Function

Private Sub ApplyAssigmentRequests( _
    ByVal v_rsWorkGroupAssignments As ADODB.Recordset)
        
    Dim l_strRecStatus As String
    Dim l_blnRequestProcessed As Boolean
    
On Error GoTo eh

    v_rsWorkGroupAssignments.MoveFirst
    While Not v_rsWorkGroupAssignments.EOF
        l_strRecStatus = UCase(v_rsWorkGroupAssignments!RECORD_STATUS)
        Select Case l_strRecStatus
            Case "I"
                AssignmentInsert v_rsWorkGroupAssignments
            Case "C"
                AssignmentUpdate v_rsWorkGroupAssignments
            Case "D", "Z"
                AssignmentDelete v_rsWorkGroupAssignments
        End Select
    
        v_rsWorkGroupAssignments.MoveNext
    Wend

    Exit Sub
    
eh:
  TrsRaiseError m_strCMODULENAME, "ApplyAssigmentRequests"
End Sub


Private Sub AssignmentInsert( _
                ByVal v_rsWorkGroupAssignments As ADODB.Recordset)
        
    Dim l_blnRequestProcessed As Boolean
    Dim l_strPrimaryInd As String
    
On Error GoTo eh
    AppendEmployeeCascadeInsert v_rsWorkGroupAssignments
    AdjustAdjacentPrimary v_rsWorkGroupAssignments
    l_blnRequestProcessed = AdjustedOverlappingAssignments(v_rsWorkGroupAssignments)
    If l_blnRequestProcessed Then
    Else
        AssignmentInsertBasic v_rsWorkGroupAssignments
    End If
                
    Exit Sub
    
eh:
    
  TrsRaiseError m_strCMODULENAME, "AssignmentInsert"
End Sub

                
Private Sub AssignmentInsertBasic( _
    ByVal v_rsWorkGroupAssignments As ADODB.Recordset)
    
    Dim l_strPrimaryInd As String
    Dim l_lngEmpNum As Long
    Dim l_datStartDateAdj As Date
    Dim sWarning As String
    Dim l_blnWriteError As Boolean
          
    On Error GoTo eh
    
    m_rsWG_Assignments.AddNew
    m_rsWG_Assignments!EMP_NUM = v_rsWorkGroupAssignments!EMP_NUM
    m_rsWG_Assignments!WG_NUM = m_lngWorkGroup
    l_strPrimaryInd = PrimaryIndicator(v_rsWorkGroupAssignments)
    m_rsWG_Assignments!PRIMARY_IND = l_strPrimaryInd
    m_rsWG_Assignments!EFF_START_DATE = GetEffectiveStartDate(v_rsWorkGroupAssignments)
    m_rsWG_Assignments!EFF_END_DATE = GetEffectiveEndDate(v_rsWorkGroupAssignments)
    m_rsWG_Assignments!TEMP_OT_BANK_IND = GetAssignmentTempOtBankInd(v_rsWorkGroupAssignments)
'    m_rsWG_Assignments!REQUIRED_PERSON = GetAssignmentRequiredInd(v_rsWorkGroupAssignments)
    m_rsWG_Assignments!LAST_MOD_EMP_NUM = CLng(m_strUserId)
    m_rsWG_Assignments!SUBSYS_OWNER = "A"
    m_rsWG_Assignments!CREATED_EMP_NUM = CLng(m_strUserId)
    If l_strPrimaryInd = "Y" Then
        l_lngEmpNum = v_rsWorkGroupAssignments!EMP_NUM
        l_datStartDateAdj = GetEffectiveStartDate(v_rsWorkGroupAssignments)
        If IsStartDateBeforeEmployeeEmploymentDate(l_lngEmpNum, l_datStartDateAdj) Then
            m_rsWG_Assignments!EFF_START_DATE = l_datStartDateAdj
            sWarning = "System changed the schedule start date because it was before the employee's employment date. "
            l_blnWriteError = AppendWGAssignmentsBRV(v_rsWorkGroupAssignments, 30000, mstrCBUSINESSWAR, sWarning)
        End If
    End If
    m_rsWG_Assignments.Update
 
    Exit Sub
    
eh:
    
  TrsRaiseError m_strCMODULENAME, "AssignmentInsertBasic"

End Sub

Private Sub AdjustAdjacentPrimary( _
    ByVal v_rsWorkGroupAssignments As ADODB.Recordset)
    
'This Subroutine adjusts adjacent primary assignments
'when the request intersects a primary assignment in
'another work group
'For example Jim (33724) is going from Greg's WG (11674)
' to Mark's (14189).
'Jim currently assigned to Greg's for the period
'07/01/1998 to 12/31/9999.
'Jim will be assigned to Mark's starting 11/01/2001 to 12/31/9999
'This routine will end the assignment in Greg's work group
' with an effective end date of 10/31/2001

    Dim l_strPrimaryInd As String
    Dim l_datPrimaryStartDate As Date
    Dim l_datPrimaryEndDate As Date
    Dim l_datRequestStartDate As Date
    Dim l_datRequestEndDate As Date
    Dim l_lngEmpNum As Long
    Dim l_intIntersectPeriod As Integer
    
    On Error GoTo eh
    
    l_strPrimaryInd = PrimaryIndicator(v_rsWorkGroupAssignments)
    If l_strPrimaryInd = "N" Then
        Exit Sub
    End If

'Find Existing Primary Assignments
    l_lngEmpNum = v_rsWorkGroupAssignments!EMP_NUM
    m_rsWG_PAIOWG.Filter = "EMP_NUM = " & str(l_lngEmpNum)
    If RecordsExist(m_rsWG_PAIOWG) Then
        m_rsWG_PAIOWG.Sort = "EFF_START_DATE"
    Else
        m_rsWG_PAIOWG.Filter = ""
        Exit Sub
    End If
 
    m_rsWG_PAIOWG.MoveFirst
    While Not m_rsWG_PAIOWG.EOF
        l_datRequestStartDate = GetEffectiveStartDate(v_rsWorkGroupAssignments)
        l_datRequestEndDate = GetEffectiveEndDate(v_rsWorkGroupAssignments)
        l_datPrimaryStartDate = GetEffectiveStartDate(m_rsWG_PAIOWG)
        l_datPrimaryEndDate = GetEffectiveEndDate(m_rsWG_PAIOWG)
        l_intIntersectPeriod = DetermineInsertingPeriod(l_datRequestStartDate, l_datRequestEndDate, _
                            l_datPrimaryStartDate, l_datPrimaryEndDate)
        Select Case l_intIntersectPeriod
            Case 0 'No insertection
            
            Case 1  'Primary Assignment overlapps entire existing assignment
                    'Delete existing assignment
                    'Existing Primary Assignment should be deleted
                    'for testing purposes I am setting it to Secondary
'                m_rsWG_PAIOWG!PRIMARY_IND = "N"
'                m_rsWG_PAIOWG!LAST_MOD_EMP_NUM = CLng(m_strUserId)
                m_rsWG_PAIOWG.Delete
                m_rsWG_PAIOWG.Update
                
            Case 2 'Primary Assignment entirely within existing assignmnet
                   'Primary Assignment Causes Split
                m_rsWG_PAIOWG!EFF_END_DATE = DateAdd("d", -1, l_datRequestStartDate)
                m_rsWG_PAIOWG!LAST_MOD_EMP_NUM = CLng(m_strUserId)
                AdjustAdjPrimAddNew
                m_rsWG_PAIOWG!EFF_START_DATE = DateAdd("d", 1, l_datRequestEndDate)
                m_rsWG_PAIOWG!EFF_END_DATE = l_datPrimaryEndDate
                m_rsWG_PAIOWG!LAST_MOD_EMP_NUM = CLng(m_strUserId)
                m_rsWG_PAIOWG.Update
                
            Case 3 'Overrlapps beginning
                m_rsWG_PAIOWG!EFF_START_DATE = DateAdd("d", 1, l_datRequestEndDate)
                m_rsWG_PAIOWG!LAST_MOD_EMP_NUM = CLng(m_strUserId)
                m_rsWG_PAIOWG.Update
                
            Case 4 'Overrlapps end
                m_rsWG_PAIOWG!EFF_END_DATE = DateAdd("d", -1, l_datRequestStartDate)
                m_rsWG_PAIOWG!LAST_MOD_EMP_NUM = CLng(m_strUserId)
                m_rsWG_PAIOWG.Update
            
        End Select
        
        m_rsWG_PAIOWG.MoveNext
    Wend
 
    Exit Sub
    
eh:
  TrsRaiseError m_strCMODULENAME, "AdjustAdjacentPrimary"
End Sub


Private Sub AppendEmployeeCascadeInsert( _
                ByVal v_rsWorkGroupAssignments As ADODB.Recordset)
    Dim l_strPrimaryInd As String
    Dim l_datEffStartDate As Date
    Dim l_datEffEndDate As Date
          
On Error GoTo eh
    
    l_strPrimaryInd = PrimaryIndicator(v_rsWorkGroupAssignments)
    l_datEffStartDate = GetEffectiveStartDate(v_rsWorkGroupAssignments)
    l_datEffEndDate = GetEffectiveEndDate(v_rsWorkGroupAssignments)
    
    If l_strPrimaryInd = "Y" Then
        AppendEmployeeCascadeDates "I", v_rsWorkGroupAssignments!EMP_NUM, _
            l_datEffStartDate, l_datEffEndDate
    End If
    
    Exit Sub
    
eh:
  TrsRaiseError m_strCMODULENAME, "AppendEmployeeCascadeInsert"
End Sub


Private Function PrimaryIndicator( _
                    ByVal v_rsWorkGroupAssignments As ADODB.Recordset) As String

    Dim l_strPrimaryInd As String

On Error GoTo eh
    l_strPrimaryInd = UCase(v_rsWorkGroupAssignments!PRIMARY_IND)
    Select Case l_strPrimaryInd
        Case "P", "Y" 'Primary Assignment
            l_strPrimaryInd = "Y"
        Case Else
            l_strPrimaryInd = "N"
    End Select
    PrimaryIndicator = l_strPrimaryInd
    Exit Function
    
eh:
  TrsRaiseError m_strCMODULENAME, "PrimaryIndicator"
End Function


Private Function DetermineInsertingPeriod( _
            ByVal v_datRequestStart As Date, _
            ByVal v_datRequestEnd As Date, _
            ByVal v_datPrimaryStart As Date, _
            ByVal v_datPrimaryEnd As Date _
            ) As Integer

    On Error GoTo eh

'Overlapps entire existing primary assignment
'        PS ._____________. PE
'     RS .___________________. RE
        If v_datRequestStart <= v_datPrimaryStart And _
            v_datRequestEnd >= v_datPrimaryEnd Then
            DetermineInsertingPeriod = 1
            Exit Function
        End If
        
'Entirely within existing primary assignment
'        PS ._____________. PE
'           RS .______. RE
        If v_datRequestStart > v_datPrimaryStart And _
            v_datRequestEnd < v_datPrimaryEnd And _
            v_datRequestStart < v_datPrimaryEnd And _
            v_datRequestEnd > v_datPrimaryStart Then
            DetermineInsertingPeriod = 2
            Exit Function
        End If

'Overlapps at beginning of existing primary assignment
'        PS ._____________. PE
'   RS .___<=__________. RE

        If v_datRequestStart <= v_datPrimaryStart And _
            v_datRequestEnd < v_datPrimaryEnd And _
            v_datRequestEnd >= v_datPrimaryStart Then
                DetermineInsertingPeriod = 3
            Exit Function
        End If
            
'Overlapps at end of existing primary assignment
'        PS ._____________. PE
'             RS .________=>_______. RE
        If v_datRequestStart > v_datPrimaryStart And _
            v_datRequestEnd >= v_datPrimaryEnd And _
            v_datRequestStart <= v_datPrimaryEnd Then
                DetermineInsertingPeriod = 4
            Exit Function
        End If
            
'Does not interset at all
       DetermineInsertingPeriod = 0
       
       Exit Function
       
eh:
    
  TrsRaiseError m_strCMODULENAME, "DetermineInsertingPeriod"
                    
            
End Function


Private Sub AdjustAdjPrimAddNew()

    Dim l_lngEmpNum As Long
    Dim l_lng_WgNum As Long
    Dim l_strOTBank As String
    Dim l_strRequired As String
    
    On Error GoTo eh

    l_lngEmpNum = m_rsWG_PAIOWG!EMP_NUM
    l_lng_WgNum = m_rsWG_PAIOWG!WG_NUM
    l_strOTBank = m_rsWG_PAIOWG!TEMP_OT_BANK_IND
'    l_strRequired = m_rsWG_PAIOWG!REQUIRED_PERSON
    
    m_rsWG_PAIOWG.AddNew
    
    m_rsWG_PAIOWG!EMP_NUM = l_lngEmpNum
    m_rsWG_PAIOWG!WG_NUM = l_lng_WgNum
    m_rsWG_PAIOWG!PRIMARY_IND = "Y"
    m_rsWG_PAIOWG!TEMP_OT_BANK_IND = l_strOTBank
'    m_rsWG_PAIOWG!REQUIRED_PERSON = l_strRequired
'    m_rsWG_PAIOWG!TIME_STAMP = ""
    m_rsWG_PAIOWG!CREATED_EMP_NUM = CLng(m_strUserId)
    m_rsWG_PAIOWG!LAST_MOD_EMP_NUM = CLng(m_strUserId)
    m_rsWG_PAIOWG!SUBSYS_OWNER = "A"
     
    Exit Sub
    
eh:
    
  TrsRaiseError m_strCMODULENAME, "AdjustAdjPrimAddNew"

    
    
End Sub


Private Function SaveWGAssignments() As Boolean
    
    Dim l_obj As dbWorkGroup.cdWorkGroupTrans
    Dim l_objDS As dbSecurity.cdBuiltDataScope
    'Dim l_blnInvalidateDataScope As Boolean, l_blnInvalidateDataScope2 As Boolean
    Dim l_varInvalidateDataScopeWG1 As Variant, l_varInvalidateDataScopeWG2 As Variant
    Dim l_varInvalidateDataScopeEmp1 As Variant, l_varInvalidateDataScopeEmp2 As Variant
    Dim i As Long
On Error GoTo eh

'Not Needed in TRIS 5
'''''    If Not HighManHours() Then
'''''        SaveWGAssignments = False
'''''        Exit Function
'''''    End If

    CheckAssignmentChangeForInvalidate m_rsWG_Assignments, "I", l_varInvalidateDataScopeWG1, l_varInvalidateDataScopeEmp1
'    CheckAssignmentChangeForInvalidate m_rsWG_PAIOWG, "I", l_varInvalidateDataScopeWG2, l_varInvalidateDataScopeEmp2
    
    m_rsWG_PAIOWG.Filter = ""
    m_rsWG_Assignments.Filter = ""
    Set l_obj = CreateObject("dbWorkGroup.cdWorkGroupTrans")
    SaveWGAssignments = l_obj.SaveWGAssignments( _
            m_rsWG_Assignments, m_rsWG_PAIOWG)
    Set l_obj = Nothing
    
    '//-----modify the cache immediately with an add or remove assignment rather than invalidate the entire wg
    
    '//-----invalidate Employee DataScope containing the changing WorkGroup membership
'    If Not IsEmpty(l_varInvalidateDataScopeWG1) Then
'        For i = LBound(l_varInvalidateDataScopeWG1) To UBound(l_varInvalidateDataScopeWG1)
'            Set l_objDS = CreateObject("dbSecurity.cdBuiltDataScope")
'            l_objDS.InvalidateDataScopeEmpByRefWG CLng(l_varInvalidateDataScopeWG1(i)), False
'            Set l_objDS = Nothing
'        Next i
'    End If
'
'    If Not IsEmpty(l_varInvalidateDataScopeWG2) Then
'        For i = LBound(l_varInvalidateDataScopeWG2) To UBound(l_varInvalidateDataScopeWG2)
'            Set l_objDS = CreateObject("dbSecurity.cdBuiltDataScope")
'            l_objDS.InvalidateDataScopeEmpByRefWG CLng(l_varInvalidateDataScopeWG2(i)), False
'            Set l_objDS = Nothing
'        Next i
'    End If
'
'    '//-----invalidate Employee DataScope containing the changing Employee
    If Not IsEmpty(l_varInvalidateDataScopeEmp1) Then
        For i = LBound(l_varInvalidateDataScopeEmp1) To UBound(l_varInvalidateDataScopeEmp1)
            Set l_objDS = CreateObject("dbSecurity.cdBuiltDataScope")
            l_objDS.InvalidateDataScopeWGByEffEmp CLng(l_varInvalidateDataScopeEmp1(i)), True
            Set l_objDS = Nothing
        Next i
    End If
'
'    If Not IsEmpty(l_varInvalidateDataScopeEmp2) Then
'        For i = LBound(l_varInvalidateDataScopeEmp2) To UBound(l_varInvalidateDataScopeEmp2)
'            Set l_objDS = CreateObject("dbSecurity.cdBuiltDataScope")
'            l_objDS.InvalidateDataScopeWGByEffEmp CLng(l_varInvalidateDataScopeEmp2(i)), True
'            Set l_objDS = Nothing
'        Next i
'    End If
    
    '//only invalidate the user that is making the wg assignment (grantor), rather than all users who have access to the workgroup
    If m_lngUserNum <> 0 Then
        Set l_objDS = CreateObject("dbSecurity.cdBuiltDataScope")
        l_objDS.InvalidateDataScopeWGByEffEmp m_lngUserNum, True
        Set l_objDS = Nothing
    End If
    
    Exit Function
    
eh:
    GetObjectContext.SetAbort
  TrsRaiseError m_strCMODULENAME, "SaveWGAssignments"

End Function

'//need to keep the inserts separate from the deletes, so pass in a flag/filter and make 2 calls, one for inserts and one for deletes
'//really need to put wg_num and emp_num together in the same array or collection, so that both can be passed to sql that does the insert or delete
Private Function CheckAssignmentChangeForInvalidate(ByVal v_rsWG_Assignments As ADODB.Recordset, _
                                                    ByVal vstr_Filter As String, _
                                                    ByRef varWG As Variant, _
                                                    ByRef varEmp As Variant)
    Dim l_dtStart As Date, l_dtEnd As Date, l_dtStart2 As Date, l_dtEnd2 As Date
    Dim l_cuSetWG As utilTRISCommonLib.cuSet
    Dim l_cuSetEmp As utilTRISCommonLib.cuSet
    Dim l_colErrArgs As Collection
On Error GoTo eh
    v_rsWG_Assignments.Filter = adFilterNone
    v_rsWG_Assignments.Filter = adFilterPendingRecords
    Set l_cuSetWG = New cuSet
    Set l_cuSetEmp = New cuSet
    
    If RecordsExist(v_rsWG_Assignments) Then
        v_rsWG_Assignments.MoveFirst
        Do While Not v_rsWG_Assignments.EOF
            If v_rsWG_Assignments.Status = adRecDeleted Then
                If vstr_Filter = "D" Then
                    l_dtStart = v_rsWG_Assignments.Fields("EFF_START_DATE").OriginalValue
                    l_dtEnd = v_rsWG_Assignments.Fields("EFF_END_DATE").OriginalValue
                    If IsDateRangeEffectiveToday(l_dtStart, l_dtEnd) Then
                        l_cuSetWG.Add v_rsWG_Assignments.Fields("WG_NUM").OriginalValue
                        l_cuSetEmp.Add v_rsWG_Assignments.Fields("EMP_NUM").OriginalValue
                    End If
                End If
            ElseIf v_rsWG_Assignments.Status = adRecNew Then
                If vstr_Filter = "I" Then
                    l_dtStart = v_rsWG_Assignments.Fields("EFF_START_DATE").Value
                    l_dtEnd = v_rsWG_Assignments.Fields("EFF_END_DATE").Value
                    If IsDateRangeEffectiveToday(l_dtStart, l_dtEnd) Then
                        l_cuSetWG.Add v_rsWG_Assignments.Fields("WG_NUM").Value
                        l_cuSetEmp.Add v_rsWG_Assignments.Fields("EMP_NUM").Value
                    End If
                End If
            ElseIf v_rsWG_Assignments.Status = adRecModified Then '//compare before and after
                l_dtStart = v_rsWG_Assignments.Fields("EFF_START_DATE").OriginalValue
                l_dtEnd = v_rsWG_Assignments.Fields("EFF_END_DATE").OriginalValue
            
                l_dtStart2 = v_rsWG_Assignments.Fields("EFF_START_DATE").Value
                l_dtEnd2 = v_rsWG_Assignments.Fields("EFF_END_DATE").Value
                
                '//If IsDateRangeEffectiveToday(l_dtStart, l_dtEnd) <> IsDateRangeEffectiveToday(l_dtStart2, l_dtEnd2) Then
                '//was not effective, but now is => Insert
                If Not IsDateRangeEffectiveToday(l_dtStart, l_dtEnd) And IsDateRangeEffectiveToday(l_dtStart2, l_dtEnd2) Then
                    If vstr_Filter = "I" Then
                        l_cuSetWG.Add v_rsWG_Assignments.Fields("WG_NUM").Value
                        l_cuSetEmp.Add v_rsWG_Assignments.Fields("EMP_NUM").Value
                    End If
                End If
                
                '//was effective, but now is not => Delete
                If IsDateRangeEffectiveToday(l_dtStart, l_dtEnd) And Not IsDateRangeEffectiveToday(l_dtStart2, l_dtEnd2) Then
                    If vstr_Filter = "D" Then
                        l_cuSetWG.Add v_rsWG_Assignments.Fields("WG_NUM").Value
                        l_cuSetEmp.Add v_rsWG_Assignments.Fields("EMP_NUM").Value
                    End If
                End If
                '//End If
            End If
            
            v_rsWG_Assignments.MoveNext
        Loop
    End If
    
    v_rsWG_Assignments.Filter = adFilterNone
    
    varWG = ArrayFromCollection(l_cuSetWG.Contents)
    varEmp = ArrayFromCollection(l_cuSetEmp.Contents)
    
    Set l_cuSetWG = Nothing
    Set l_cuSetEmp = Nothing
    
    Exit Function
eh:
    Set l_colErrArgs = New Collection
    l_colErrArgs.Add "m_lngWorkGroup = " & m_lngWorkGroup
    l_colErrArgs.Add "l_dtStart (from) = " & l_dtStart
    l_colErrArgs.Add "l_dtEnd          = " & l_dtEnd
    l_colErrArgs.Add "l_dtStart2 (to)  = " & l_dtStart2
    l_colErrArgs.Add "l_dtEnd2         = " & l_dtEnd
    
    Set l_cuSetWG = Nothing
    Set l_cuSetEmp = Nothing

    '//-----log any errors but allow the save to finish without errors
    CheckAssignmentChangeForInvalidate = Array(m_lngWorkGroup)
    TrsTraceError m_strCMODULENAME, "CheckAssignmentChangeForInvalidate", , , , l_colErrArgs, True
End Function

Private Function IsDateRangeEffectiveToday(ByVal v_dtStart As Date, ByVal v_dtEnd As Date) As Boolean
On Error GoTo eh
    If v_dtStart <= CDate(Format(Now, "mm/dd/yyyy")) And v_dtEnd >= CDate(Format(Now, "mm/dd/yyyy")) Then
        IsDateRangeEffectiveToday = True
    Else
        IsDateRangeEffectiveToday = False
    End If
    Exit Function
eh:
    '//-----log any errors but allow the save to finish without errors
    IsDateRangeEffectiveToday = True
    TrsTraceError m_strCMODULENAME, "IsDateRangeEffectiveToday", , , , , True
End Function

Private Sub AssignmentUpdate( _
    ByVal v_rsWorkGroupAssignments As ADODB.Recordset)
    
    Dim l_strFilter As String
    Dim l_strEmpNum As String
    Dim l_strTimeStamp As String
    Dim l_strPrimaryInd As String
    Dim l_blnPrimaryIndChanged As Boolean
          
    On Error GoTo eh
    
    l_strEmpNum = Trim(str(v_rsWorkGroupAssignments!EMP_NUM))
    l_strTimeStamp = v_rsWorkGroupAssignments!CREATED_TIMESTAMP
    l_strFilter = " EMP_NUM = " & l_strEmpNum & _
            " AND CREATED_TIMESTAMP = " & l_strTimeStamp
    l_strFilter = MDAC25FilterFix(l_strFilter)
    m_rsWG_Assignments.Filter = l_strFilter
    If Not RecordsExist(m_rsWG_Assignments) Then
        m_rsWG_Assignments.Filter = ""
        Err.Raise vbObjectError + 30000, "busWorkGroupTrans.AssignmenUpdate", "Invalid Update Request"
        'Raise error
        'Invalid change request
        'Data has been updated by someone else prior to save
        'Please Try again.
        Exit Sub
    End If
    l_strPrimaryInd = PrimaryIndicator(v_rsWorkGroupAssignments)
    l_blnPrimaryIndChanged = Trim(m_rsWG_Assignments!PRIMARY_IND) <> l_strPrimaryInd

    If (l_blnPrimaryIndChanged And l_strPrimaryInd = "N") Then  'primary to secondary
        AssignmentDeletePrimary
        AssignmentInsert v_rsWorkGroupAssignments
    ElseIf (l_blnPrimaryIndChanged And l_strPrimaryInd = "Y") Then  'secondary to primary
        AdjustAdjacentPrimary v_rsWorkGroupAssignments
        AssignmentUpdateBasic l_strPrimaryInd, l_blnPrimaryIndChanged, v_rsWorkGroupAssignments, m_rsWG_Assignments
    Else
        AdjustAdjacentPrimary v_rsWorkGroupAssignments
        AssignmentUpdateBasic l_strPrimaryInd, l_blnPrimaryIndChanged, v_rsWorkGroupAssignments, m_rsWG_Assignments
    End If

    m_rsWG_Assignments.Filter = ""
    m_rsWG_PAIOWG.Filter = ""
    Exit Sub
    
eh:
    
  TrsRaiseError m_strCMODULENAME, "AssignmentUpdate"

End Sub


Private Sub AssignmentUpdateBasic( _
                    ByVal v_strPrimaryInd As String, _
                    ByVal v_blnPrimaryIndChanged As Boolean, _
                    ByVal v_rsWorkGroupAssignments As ADODB.Recordset, _
                    ByRef r_rsWG_Assignments As ADODB.Recordset)
                    
On Error GoTo eh

    r_rsWG_Assignments!PRIMARY_IND = v_strPrimaryInd
    r_rsWG_Assignments!EFF_START_DATE = GetEffectiveStartDate(v_rsWorkGroupAssignments)
    r_rsWG_Assignments!EFF_END_DATE = GetEffectiveEndDate(v_rsWorkGroupAssignments)
    r_rsWG_Assignments!TEMP_OT_BANK_IND = GetAssignmentTempOtBankInd(v_rsWorkGroupAssignments)
    r_rsWG_Assignments!LAST_MOD_EMP_NUM = CLng(m_strUserId)
    r_rsWG_Assignments!SUBSYS_OWNER = "A"
    r_rsWG_Assignments.Update

    If v_blnPrimaryIndChanged Then
        AppendEmployeeCascadeDates _
            "C", r_rsWG_Assignments!EMP_NUM, _
            r_rsWG_Assignments.Fields("EFF_START_DATE").OriginalValue, _
            r_rsWG_Assignments.Fields("EFF_END_DATE").OriginalValue
    End If
    If v_strPrimaryInd = "Y" Then
        AppendCascadeEmployeeDateDifference _
            "C", r_rsWG_Assignments!EMP_NUM, _
            r_rsWG_Assignments.Fields("EFF_START_DATE").Value, _
            r_rsWG_Assignments.Fields("EFF_END_DATE").Value, _
            r_rsWG_Assignments.Fields("EFF_START_DATE").OriginalValue, _
            r_rsWG_Assignments.Fields("EFF_END_DATE").OriginalValue
        AdjustGapsInPrimaryAssignment
    End If

    Exit Sub
                    
eh:
  TrsRaiseError m_strCMODULENAME, "AssignmentUpdate"
End Sub

Private Function AppendCascadeEmployeeDateDifference( _
                    ByVal v_strAction As String, _
                    ByVal v_lngEmpNum As Long, _
                    ByVal v_datStartDate As Date, _
                    ByVal v_datEndDate As Date, _
                    ByVal v_datKeyStartDate As Date, _
                    ByVal v_datKeyEndDate As Date)

On Error GoTo eh
    If v_datStartDate < v_datKeyStartDate Then
        AppendEmployeeCascadeDates _
            "C", v_lngEmpNum, _
            v_datStartDate, _
            DateAdd("d", -1, v_datKeyStartDate)
    ElseIf v_datStartDate > v_datKeyStartDate Then
        AppendEmployeeCascadeDates _
            "C", v_lngEmpNum, _
            v_datKeyStartDate, _
            DateAdd("d", -1, v_datStartDate)
    End If
    If v_datEndDate < v_datKeyEndDate Then
        AppendEmployeeCascadeDates _
            "C", v_lngEmpNum, _
            DateAdd("d", 1, v_datEndDate), _
            v_datKeyEndDate
    ElseIf v_datEndDate > v_datKeyEndDate Then
        AppendEmployeeCascadeDates _
            "C", v_lngEmpNum, _
            DateAdd("d", 1, v_datKeyEndDate), _
            v_datEndDate
    End If

    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "AppendCascadeEmployeeDateDifference"
End Function


Private Sub AdjustGapsInPrimaryAssignment()
    
'This Subroutine adjusts the Gap in a primary assignments
'This happens when the user shortens an existing primary
'assignment.
'The assignment prior to the request and/or the assignment
' after the request needs changed.
'For example Jim (33724) is assigned to Greg's WG (11674)
' to Mark's (14189).
'Jim is assigned to Greg's WG (11674) for the period
'07/01/1998 to 10/31/2001 and Mark's (14189) for
'11/01/2001 to 12/31/9999.
'Kelly changed his mind so Jim will not start in Mark's WG
'until the first of the year
'There will be a change request that will show the new start
'date of 01/01/2002.  The effective end date of the prior assignment
'must be set to 12/31/2001.
'PLEASE NOTE THAT LEGACY TRIS DOES NOT ALLOW THIS


    Dim l_strPrimaryInd As String
    Dim l_datPrimaryStartDate As Date
    Dim l_datPrimaryEndDate As Date
    Dim l_datEffStartDate As Date
    Dim l_datEffEndDate As Date
    Dim l_datOrigEffStartDate As Date
    Dim l_datOrigEffEndDate As Date
    Dim l_lngEmpNum As Long
    Dim l_strFilter As String
    
    On Error GoTo eh
    
    l_lngEmpNum = m_rsWG_Assignments.Fields("EMP_NUM").Value
    l_datEffStartDate = m_rsWG_Assignments.Fields("EFF_START_DATE").Value
    l_datOrigEffStartDate = m_rsWG_Assignments.Fields("EFF_START_DATE").OriginalValue
    If l_datEffStartDate > l_datOrigEffStartDate Then
        l_datPrimaryEndDate = DateAdd("d", -1, l_datOrigEffStartDate)
        l_strFilter = "EMP_NUM = " & str(l_lngEmpNum) & _
                " AND EFF_END_DATE = " & str(l_datPrimaryEndDate)
        l_strFilter = MDAC25FilterFix(l_strFilter)
        m_rsWG_PAIOWG.Filter = l_strFilter
        If RecordsExist(m_rsWG_PAIOWG) Then
            m_rsWG_PAIOWG!EFF_END_DATE = DateAdd("d", -1, l_datEffStartDate)
            m_rsWG_PAIOWG!LAST_MOD_EMP_NUM = CLng(m_strUserId)
            m_rsWG_PAIOWG.Update
        End If
    End If

    l_datEffEndDate = m_rsWG_Assignments.Fields("EFF_END_DATE").Value
    l_datOrigEffEndDate = m_rsWG_Assignments.Fields("EFF_END_DATE").OriginalValue
    If l_datOrigEffEndDate = m_datMaxEndDate Then
    'Last assignment
    Else
    'Check to see if adjustment is needed
        If l_datEffEndDate < l_datOrigEffEndDate Then
            l_datPrimaryStartDate = DateAdd("d", 1, l_datOrigEffEndDate)
            l_strFilter = "EMP_NUM = " & str(l_lngEmpNum) & _
                    " AND EFF_START_DATE = " & str(l_datPrimaryStartDate)
            l_strFilter = MDAC25FilterFix(l_strFilter)
            m_rsWG_PAIOWG.Filter = l_strFilter
            If RecordsExist(m_rsWG_PAIOWG) Then
                m_rsWG_PAIOWG!EFF_START_DATE = DateAdd("d", 1, l_datEffEndDate)
                m_rsWG_PAIOWG!LAST_MOD_EMP_NUM = CLng(m_strUserId)
                m_rsWG_PAIOWG.Update
            End If
        End If
    End If
    m_rsWG_PAIOWG.Filter = ""

    Exit Sub
    
eh:
    
  TrsRaiseError m_strCMODULENAME, "AdjustGapsInPrimaryAssignment"

End Sub

Private Sub AssignmentDelete( _
    ByVal v_rsWorkGroupAssignments As ADODB.Recordset)
    
    Dim l_strFilter As String
    Dim l_strEmpNum As String
    Dim l_strTimeStamp As String
    Dim l_strPrimaryInd As String
    Dim l_datEffStartDate As Date
    Dim l_datEffEndDate As Date
    Dim l_datPrimaryStartDate As Date
    Dim l_datPrimaryEndDate As Date
    Dim l_blnAdjacentPrimaryChanged As Boolean
    Dim l_blnAdjPriorAssignment As Boolean
    
    Dim l_lngWgNum As Long
    Dim l_lngEmpNum As Long
          
On Error GoTo eh
    
    l_blnAdjacentPrimaryChanged = False
    l_strEmpNum = Trim(str(v_rsWorkGroupAssignments!EMP_NUM))
    l_strTimeStamp = v_rsWorkGroupAssignments!CREATED_TIMESTAMP
    l_strFilter = " EMP_NUM = " & l_strEmpNum & _
            " AND CREATED_TIMESTAMP = " & l_strTimeStamp
    l_strFilter = MDAC25FilterFix(l_strFilter)
    m_rsWG_Assignments.Filter = l_strFilter
    If Not RecordsExist(m_rsWG_Assignments) Then
        m_rsWG_Assignments.Filter = ""
        Err.Raise vbObjectError + 30000, "busWorkGroupTrans.AssignmenDelete", "Invalid Delete Request"
        'Raise error
        'Invalid change request
        'Data has been updated by someone else prior to save
        'Please Try again.
        Exit Sub
    End If
    l_strPrimaryInd = UCase(m_rsWG_Assignments!PRIMARY_IND)
    l_lngWgNum = m_rsWG_Assignments!WG_NUM
    l_lngEmpNum = m_rsWG_Assignments!EMP_NUM
    Select Case l_strPrimaryInd
        Case "P", "Y" 'Primary Assignment
            AssignmentDeletePrimary
        Case Else
'    If secondary, just delete and get out
            m_rsWG_Assignments.Delete
    End Select

    m_rsWG_Assignments.Filter = ""
    m_rsWG_PAIOWG.Filter = ""
 
    Exit Sub
    
eh:
  TrsRaiseError m_strCMODULENAME, "AssignmentDelete"
End Sub


Private Sub AssignmentDeletePrimary()
'   Process primary assignment deletion where positioned on m_rsWG_Assignments.
    
    Dim l_strFilter As String
    Dim l_strEmpNum As String
    Dim l_strTimeStamp As String
    Dim l_strPrimaryInd As String
    Dim l_datEffStartDate As Date
    Dim l_datEffEndDate As Date
    Dim l_datPrimaryStartDate As Date
    Dim l_datPrimaryEndDate As Date
    Dim l_blnAdjacentPrimaryChanged As Boolean
    Dim l_blnAdjPriorAssignment As Boolean
    Dim l_lngWgNum As Long
    Dim l_lngEmpNum As Long
          
On Error GoTo eh
    
    l_strPrimaryInd = "Y"
    l_blnAdjacentPrimaryChanged = False
    l_strEmpNum = Trim(CStr(m_rsWG_Assignments!EMP_NUM))
    l_strTimeStamp = m_rsWG_Assignments!CREATED_TIMESTAMP
    l_lngWgNum = m_rsWG_Assignments!WG_NUM
    l_lngEmpNum = m_rsWG_Assignments!EMP_NUM
    
    If l_strPrimaryInd = "Y" Then
        AppendEmployeeCascadeDates "D", m_rsWG_Assignments!EMP_NUM, _
            m_rsWG_Assignments!EFF_START_DATE, m_rsWG_Assignments!EFF_END_DATE
    End If
    
    l_datEffStartDate = m_rsWG_Assignments!EFF_START_DATE
    l_datEffEndDate = m_rsWG_Assignments!EFF_END_DATE
    
    If l_strPrimaryInd = "Y" Then
        'Extend the Prior WG Assignment
        l_datPrimaryEndDate = DateAdd("d", -1, l_datEffStartDate)
        l_strFilter = "EMP_NUM = " & l_strEmpNum & _
                " AND EFF_END_DATE = " & str(l_datPrimaryEndDate)
        l_strFilter = MDAC25FilterFix(l_strFilter)
        m_rsWG_PAIOWG.Filter = l_strFilter
        If RecordsExist(m_rsWG_PAIOWG) Then
            l_blnAdjacentPrimaryChanged = True
            m_rsWG_PAIOWG!EFF_END_DATE = l_datEffEndDate
            m_rsWG_PAIOWG!LAST_MOD_EMP_NUM = CLng(m_strUserId)
            m_rsWG_PAIOWG.Update
            AppendWGAssignmentDeleteMessage _
                trInfWGAssignExtended, mstrCBUSINESSINF, "", _
                m_rsWG_PAIOWG!WG_NUM, 0, l_datEffEndDate, _
                l_lngWgNum, l_datEffStartDate, l_lngEmpNum
        Else
            'No Prior try to adjust future assignment
            If l_datEffEndDate = m_datMaxEndDate Then
                l_datPrimaryStartDate = l_datEffEndDate
            Else
                l_datPrimaryStartDate = DateAdd("d", 1, l_datEffEndDate)
            End If
            l_strFilter = "EMP_NUM = " & l_strEmpNum & _
                " AND EFF_START_DATE = " & str(l_datPrimaryStartDate)
            m_rsWG_PAIOWG.Filter = l_strFilter
            If RecordsExist(m_rsWG_PAIOWG) Then
                l_blnAdjacentPrimaryChanged = True
                m_rsWG_PAIOWG!EFF_START_DATE = l_datEffStartDate
                m_rsWG_PAIOWG!LAST_MOD_EMP_NUM = CLng(m_strUserId)
                m_rsWG_PAIOWG.Update
                AppendWGAssignmentDeleteMessage _
                    trInfWGAssignExtended, mstrCBUSINESSINF, "", _
                    m_rsWG_PAIOWG!WG_NUM, l_datEffStartDate, 0, _
                    l_lngWgNum, l_datEffStartDate, l_lngEmpNum
            Else
            End If
        End If
    End If
      
    If l_blnAdjacentPrimaryChanged Then
        m_rsWG_Assignments.Delete
    Else
'Adjust assigments in work group
        If l_datEffEndDate = m_datMaxEndDate Then
            l_datPrimaryEndDate = DateAdd("d", -1, l_datEffStartDate)
            l_strFilter = "EMP_NUM = " & l_strEmpNum & _
                " AND EFF_END_DATE = " & str(l_datPrimaryEndDate)
            l_blnAdjPriorAssignment = True
        Else
            l_datPrimaryStartDate = DateAdd("d", 1, l_datEffEndDate)
            l_strFilter = "EMP_NUM = " & l_strEmpNum & _
                " AND EFF_START_DATE = " & str(l_datPrimaryStartDate)
            l_blnAdjPriorAssignment = False
        End If
        l_strFilter = MDAC25FilterFix(l_strFilter)
        m_rsWG_Assignments.Filter = l_strFilter
        If RecordsExist(m_rsWG_Assignments) Then
            If l_blnAdjPriorAssignment Then
                m_rsWG_Assignments!EFF_END_DATE = l_datEffEndDate
                m_rsWG_Assignments!LAST_MOD_EMP_NUM = CLng(m_strUserId)
                m_rsWG_Assignments.Update
                AppendWGAssignmentDeleteMessage _
                    trInfWGAssignExtended, mstrCBUSINESSINF, "", _
                    m_rsWG_Assignments!WG_NUM, 0, l_datEffEndDate, _
                    l_lngWgNum, l_datEffStartDate, l_lngEmpNum
            Else
                m_rsWG_Assignments!EFF_START_DATE = l_datEffStartDate
                m_rsWG_Assignments!LAST_MOD_EMP_NUM = CLng(m_strUserId)
                m_rsWG_Assignments.Update
                AppendWGAssignmentDeleteMessage _
                    trInfWGAssignExtended, mstrCBUSINESSINF, "", _
                    m_rsWG_Assignments!WG_NUM, l_datEffStartDate, 0, _
                    l_lngWgNum, l_datEffStartDate, l_lngEmpNum
            End If
        End If
        l_strFilter = " EMP_NUM = " & l_strEmpNum & _
            " AND CREATED_TIMESTAMP = " & l_strTimeStamp
        l_strFilter = MDAC25FilterFix(l_strFilter)
        m_rsWG_Assignments.Filter = l_strFilter
        m_rsWG_Assignments.Delete
    End If
 
    Exit Sub
    
eh:
  TrsRaiseError m_strCMODULENAME, "AssignmentDeletePrimary"
End Sub


Private Function PostValidateAssignmentRequests( _
            ByVal v_rsWorkGroupAssignments As ADODB.Recordset _
            ) As Boolean
            
'the function will ensure the following rules
'   1. Employee can not have overlapping assignments in
'       the same work group
'   2. Employee must be primary in 1 and only WG from the
'       1st assignement until the end of time
'   3. Can not change a primary assigment if actual time exists
'   4. Can not change an assignment that was created from the
'       Relief Maintenance subsytem
'   5. Can not change a primary assignment if the employee is
'       on a demand schedule
'   6. Can not change a secondary assignment, if the employee
'       is on a demand schedule for the work group assigments
'       being saved

'Only edits 1 and 2 are in this routine
' Edits 4,5, and 6  may not be needed due to new design
' of relief and demand scheduling
' Edit 3 will probably be delayed until CASCADE is complete
            
    Dim l_blnOK As Boolean
    Dim l_blnOverlappExists As Boolean
    
    On Error GoTo eh
    
'    l_blnOverlappExists = CheckForOverlappingAssignments(v_rsWorkGroupAssignments)
    
'    If l_blnOverlappExists Then
'        PostValidateAssignmentRequests = False
'        Exit Function
'    End If
    
    l_blnOK = AllEmployeesPrimary(v_rsWorkGroupAssignments)
    
    If l_blnOK Then
        l_blnOK = CheckForDemandInOtherWg(v_rsWorkGroupAssignments)
    End If
    
    PostValidateAssignmentRequests = l_blnOK
    
    Exit Function
    
eh:
    
  TrsRaiseError m_strCMODULENAME, "PostValidateAssignmentRequests"
            
            
End Function


Private Function AllEmployeesPrimary( _
            ByVal v_objRS As ADODB.Recordset _
            ) As Boolean
            
'RED259, RED260

    Dim l_colEmployees As Collection
    Dim l_lngEmployee As Long
    Dim l_intI, l_intCount As Integer
    Dim l_blnContigousPrimary As Boolean
    Dim l_blnAllEmployeesPrimary As Boolean

    On Error GoTo eh
    
    If m_colEmployees.Count = 0 Then
        AllEmployeesPrimary = True
        Exit Function
    End If
       
    l_intCount = m_colEmployees.Count
    l_blnAllEmployeesPrimary = True
    For l_intI = 1 To l_intCount
        l_lngEmployee = m_colEmployees(l_intI)
        l_blnContigousPrimary = IsEmployeePrimaryInOneAndOnlyOneWG(l_lngEmployee, v_objRS)
        If Not l_blnContigousPrimary Then
            l_blnAllEmployeesPrimary = False
        End If
        m_rsWG_ALL_PA.Filter = ""
    Next l_intI
    
    AllEmployeesPrimary = l_blnAllEmployeesPrimary
    
    Exit Function
    
eh:
    
  TrsRaiseError m_strCMODULENAME, "AllEmployeesPrimary"
                  
End Function
Private Function BuildPrimaryRecordset() As ADODB.Recordset

    Dim l_rsPrimary As ADODB.Recordset
    
    On Error GoTo eh
    
    Set l_rsPrimary = CreatePrimaryRecordset
    
    m_rsWG_PAIOWG.Filter = "PRIMARY_IND = 'Y'"
    If RecordsExist(m_rsWG_PAIOWG) Then
        m_rsWG_PAIOWG.MoveFirst
        While Not m_rsWG_PAIOWG.EOF
            l_rsPrimary.AddNew
            l_rsPrimary!EMP_NUM = m_rsWG_PAIOWG!EMP_NUM
            l_rsPrimary!WG_NUM = m_rsWG_PAIOWG!WG_NUM
            l_rsPrimary!EFF_START_DATE = CDate(m_rsWG_PAIOWG!EFF_START_DATE)
            l_rsPrimary!EFF_END_DATE = CDate(m_rsWG_PAIOWG!EFF_END_DATE)
            l_rsPrimary!ORIG_EFF_START_DATE = _
                    CDate(m_rsWG_PAIOWG.Fields("EFF_START_DATE").OriginalValue)
            l_rsPrimary!ORIG_EFF_END_DATE = _
                    CDate(m_rsWG_PAIOWG.Fields("EFF_END_DATE").OriginalValue)
            l_rsPrimary.Update
            m_rsWG_PAIOWG.MoveNext
        Wend
    End If
    
    m_rsWG_Assignments.Filter = "PRIMARY_IND = 'Y'"
    If RecordsExist(m_rsWG_Assignments) Then
        m_rsWG_Assignments.MoveFirst
        While Not m_rsWG_Assignments.EOF
'//            If m_rsWG_Assignments!PRIMARY_IND = "Y" Then
                l_rsPrimary.AddNew
                l_rsPrimary!EMP_NUM = m_rsWG_Assignments!EMP_NUM
                l_rsPrimary!WG_NUM = m_rsWG_Assignments!WG_NUM
                l_rsPrimary!EFF_START_DATE = CDate(m_rsWG_Assignments!EFF_START_DATE)
                l_rsPrimary!EFF_END_DATE = CDate(m_rsWG_Assignments!EFF_END_DATE)
                l_rsPrimary!ORIG_EFF_START_DATE = _
                        CDate(m_rsWG_Assignments.Fields("EFF_START_DATE").OriginalValue)
                l_rsPrimary!ORIG_EFF_END_DATE = _
                        CDate(m_rsWG_Assignments.Fields("EFF_END_DATE").OriginalValue)
                l_rsPrimary.Update
'//            End If
            m_rsWG_Assignments.MoveNext
        Wend
    End If
    
    Set BuildPrimaryRecordset = l_rsPrimary
    
    Exit Function
    
eh:
    
  TrsRaiseError m_strCMODULENAME, "BuildPrimaryRecordset"

End Function

Private Function CreatePrimaryRecordset() As ADODB.Recordset

    Dim l_rs As ADODB.Recordset
    
    On Error GoTo eh
    
    Set l_rs = New ADODB.Recordset
    With l_rs
      .Fields.Append "EMP_NUM", adInteger, , adFldUpdatable
      .Fields.Append "WG_NUM", adInteger, , adFldUpdatable
      .Fields.Append "EFF_START_DATE", adDate, , adFldUpdatable
      .Fields.Append "EFF_END_DATE", adDate, , adFldUpdatable
      .Fields.Append "ORIG_EFF_START_DATE", adDate, , adFldUpdatable
      .Fields.Append "ORIG_EFF_END_DATE", adDate, , adFldUpdatable

      .CursorLocation = adUseClient
      .CursorType = adOpenStatic
      .LockType = adLockBatchOptimistic
      .Open
    End With
    
    Set CreatePrimaryRecordset = l_rs
    
    Exit Function
    
eh:
    
  TrsRaiseError m_strCMODULENAME, "CreatePrimaryRecordset"
    
    
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
        l_colEmployees.Add -1
        Set GetEmployeesInRequest = l_colEmployees
        Exit Function
    End If
    
    v_objRS.MoveFirst
    While Not v_objRS.EOF
        l_strEmployee = str(v_objRS!EMP_NUM)
        If Not IncludesString(l_strEmployee, l_colEmployees) Then
            l_colEmployees.Add l_strEmployee
        End If
        v_objRS.MoveNext
    Wend
    
    Set GetEmployeesInRequest = l_colEmployees
            
    Exit Function
    
eh:
    
  TrsRaiseError m_strCMODULENAME, "GetEmployeesInRequest"
  
End Function


Private Function IsEmployeePrimaryInOneAndOnlyOneWG( _
            ByVal v_lngEmpNum As Long, _
            ByVal v_objRS As ADODB.Recordset _
            ) As Boolean

'This is the Golden rule of Employee Work Group Assignments
'Employee must always be Primary in a work group
'Employee can only be primary in one work group on a given day

    Dim l_blnStillPrimary As Boolean
    Dim l_datStartDate As Date
    Dim l_datPriorEndDate As Date
    Dim l_datEndDate As Date
    Dim l_blnWriteError As Boolean
    Dim l_lngCurrWG As Long
    Dim l_lngPriorWG As Long
    

    On Error GoTo eh


    m_rsWG_ALL_PA.Filter = "EMP_NUM = " & str(v_lngEmpNum)
    v_objRS.Filter = "EMP_NUM = " & str(v_lngEmpNum)
   
    If RecordsExist(m_rsWG_ALL_PA) Then
         m_rsWG_ALL_PA.Sort = "EFF_END_DATE DESC"
         m_rsWG_ALL_PA.MoveFirst
    Else
        l_blnWriteError = AppendWGAssignmentsBRV(v_objRS, 30000, mstrCBUSINESSERR, "Employee (" & str(v_lngEmpNum) & ") is not primary in any work group.")
        IsEmployeePrimaryInOneAndOnlyOneWG = False
        Exit Function
    End If
    
    l_datEndDate = m_rsWG_ALL_PA!EFF_END_DATE
    If l_datEndDate = m_datMaxEndDate Then
    'The most recent WG assignment is until the end of time
    Else
        l_blnWriteError = AppendWGAssignmentsBRV(v_objRS, 30000, mstrCBUSINESSERR, "Employee (" & str(v_lngEmpNum) & ") must always be primary in a work group. The last work group assignment ends on " & Format(l_datEndDate, "mm/dd/yyyy") & ".")
        IsEmployeePrimaryInOneAndOnlyOneWG = False
        Exit Function
    End If
    
    l_blnStillPrimary = True
    l_datStartDate = m_rsWG_ALL_PA!EFF_START_DATE
    l_lngCurrWG = m_rsWG_ALL_PA!WG_NUM
    m_rsWG_ALL_PA.MoveNext
    While Not m_rsWG_ALL_PA.EOF And l_blnStillPrimary
        l_datPriorEndDate = m_rsWG_ALL_PA!EFF_END_DATE
        l_lngPriorWG = m_rsWG_ALL_PA!WG_NUM
        If l_datStartDate = DateAdd("d", 1, l_datPriorEndDate) Then
        Else
            'Error Not Contigous
            l_blnWriteError = AppendWGAssignmentsBRV(v_objRS, 30000, mstrCBUSINESSERR, "Employee (" & str(v_lngEmpNum) & ") must always be primary in a work group. There is a gap in work group assignments. " _
                & " Assignment in work group (" & str(l_lngPriorWG) & ") ends on " & Format(l_datPriorEndDate, "mm/dd/yyyy") & "." _
                & " Assignment in work group (" & str(l_lngCurrWG) & ") begins on " & Format(l_datStartDate, "mm/dd/yyyy") & ".")
            l_blnStillPrimary = False
        End If
        l_datStartDate = m_rsWG_ALL_PA!EFF_START_DATE
        l_lngCurrWG = l_lngPriorWG
        m_rsWG_ALL_PA.MoveNext
    Wend
    
    IsEmployeePrimaryInOneAndOnlyOneWG = l_blnStillPrimary
    Exit Function

eh:
    
  TrsRaiseError m_strCMODULENAME, "IsEmployeePrimaryInOneAndOnlyOneWG"
  
End Function



Private Function AdjustedOverlappingAssignments( _
    ByVal v_rsWorkGroupAssignments As ADODB.Recordset _
    ) As Boolean
    
'This Subroutine adjusts primary assignments in the same work group
'when the request intersects a primary assignment
'The function returns false if further processing needs done

    Dim l_strPrimaryInd As String
    Dim l_datPrimaryStartDate As Date
    Dim l_datPrimaryEndDate As Date
    Dim l_datRequestStartDate As Date
    Dim l_datRequestEndDate As Date
    Dim l_lngEmpNum As Long
    Dim l_intIntersectPeriod As Integer
    
    On Error GoTo eh
    l_strPrimaryInd = PrimaryIndicator(v_rsWorkGroupAssignments)
    If l_strPrimaryInd = "N" Then
        AdjustedOverlappingAssignments = False
        Exit Function
    End If

'Find Existing Primary Assignments
    l_lngEmpNum = v_rsWorkGroupAssignments!EMP_NUM
    m_rsWG_Assignments.Filter = "EMP_NUM = " & str(l_lngEmpNum)
    If RecordsExist(m_rsWG_Assignments) Then
        m_rsWG_Assignments.Sort = "EFF_START_DATE"
    Else
        m_rsWG_Assignments.Filter = ""
        AdjustedOverlappingAssignments = False
        Exit Function
    End If
 
    m_rsWG_Assignments.MoveFirst
    While Not m_rsWG_Assignments.EOF
        l_datRequestStartDate = GetEffectiveStartDate(v_rsWorkGroupAssignments)
        l_datRequestEndDate = GetEffectiveEndDate(v_rsWorkGroupAssignments)
        l_datPrimaryStartDate = GetEffectiveStartDate(m_rsWG_Assignments)
        l_datPrimaryEndDate = GetEffectiveEndDate(m_rsWG_Assignments)
        l_intIntersectPeriod = DetermineInsertingPeriod(l_datRequestStartDate, l_datRequestEndDate, _
                            l_datPrimaryStartDate, l_datPrimaryEndDate)
        Select Case l_intIntersectPeriod
            Case 0 'No insertection
                AdjustedOverlappingAssignments = False
            
            Case 1  'Primary Assignment overlapps entire existing assignment
                    'Delete existing assignment
                    'Existing Primary Assignment should be deleted
                    'for testing purposes I am setting it to Secondary
'                m_rsWG_Assignments!EFF_START_DATE = l_datRequestStartDate
'                m_rsWG_Assignments!EFF_END_DATE = l_datRequestEndDate
'                m_rsWG_Assignments!LAST_MOD_EMP_NUM = CLng(m_strUserId)
                m_rsWG_Assignments.Delete
                AdjustedOverlappingAssignments = False
                
            Case 2 'Primary Assignment entirely within existing assignmnet
                   'No Action
                AdjustedOverlappingAssignments = True
                
            Case 3 'Overrlapps beginning
                m_rsWG_Assignments!EFF_START_DATE = l_datRequestStartDate
                m_rsWG_Assignments!LAST_MOD_EMP_NUM = CLng(m_strUserId)
                m_rsWG_Assignments.Update
                AdjustedOverlappingAssignments = True
                
            Case 4 'Overrlapps end
                m_rsWG_Assignments!EFF_END_DATE = l_datRequestEndDate
                m_rsWG_Assignments!LAST_MOD_EMP_NUM = CLng(m_strUserId)
                m_rsWG_Assignments.Update
                AdjustedOverlappingAssignments = True
            
        End Select
        
        m_rsWG_Assignments.MoveNext
    Wend
 
 
    Exit Function
    
eh:
    
  TrsRaiseError m_strCMODULENAME, "AdjustedOverlappingAssignments"

End Function

Private Function CheckForOverlappingAssignments( _
            ByVal v_rsWorkGroupAssignments As ADODB.Recordset _
            ) As Boolean
            
'Jimb 10/21/2004 this method is no longer called

  
'This Subroutine checks for overlapping assignments in the same work group
'when the request intersects a primary assignment
'The function returns true if an assignment overlapps

    Dim l_strPrimaryInd As String
    Dim l_datPrimaryStartDate As Date
    Dim l_datPrimaryEndDate As Date
    Dim l_datRequestStartDate As Date
    Dim l_datRequestEndDate As Date
    Dim l_lngEmpNum As Long
    Dim l_intIntersectPeriod As Integer
    Dim l_intI, l_intCount As Integer
    Dim l_blnOverlappExists As Boolean
    
    On Error GoTo eh
            
    CheckForOverlappingAssignments = False
    l_intCount = m_colEmployees.Count
    For l_intI = 1 To l_intCount
        l_lngEmpNum = m_colEmployees(l_intI)
        v_rsWorkGroupAssignments.Filter = "EMP_NUM = " & str(l_lngEmpNum)
        l_blnOverlappExists = OverlappingAssignmentExists(l_lngEmpNum, v_rsWorkGroupAssignments)
        If l_blnOverlappExists Then
            CheckForOverlappingAssignments = True
        End If
        m_rsWG_Assignments.Filter = ""
        v_rsWorkGroupAssignments.Filter = ""
    Next l_intI
                    
    Exit Function
    
eh:
    
  TrsRaiseError m_strCMODULENAME, "CheckForOverlappingAssignments"
            
End Function


Private Function OverlappingAssignmentExists( _
            ByVal v_lngEmpNum As Long, _
            ByVal v_rsWorkGroupAssignments As ADODB.Recordset _
            ) As Boolean
            
'JimB 10/18/2004 this method is no longer called
'Method CheckForOverlappingAssignments that called this method
' is no longer called.

    Dim l_strPrimaryInd As String
    Dim l_datStartDate As Date
    Dim l_datEndDate As Date
    Dim l_datPriorStartDate As Date
    Dim l_datPriorEndDate As Date
    Dim l_blnWriteError As Boolean
    Dim l_lngCurrWG As Long
    Dim l_lngPriorWG As Long
    Dim l_lngEmpNum As Long
    Dim l_intIntersectPeriod As Integer
    Dim l_strErrorDesc As String
    Dim l_bln As Boolean

    On Error GoTo eh
        
    m_rsWG_Assignments.Filter = "EMP_NUM = " & str(v_lngEmpNum)
    If RecordsExist(m_rsWG_Assignments) Then
        m_rsWG_Assignments.Sort = "EFF_START_DATE"
        m_rsWG_Assignments.MoveFirst
    Else
        m_rsWG_Assignments.Filter = ""
        OverlappingAssignmentExists = False
        Exit Function
    End If
 
    l_datStartDate = GetEffectiveStartDate(m_rsWG_Assignments)
    l_datEndDate = GetEffectiveEndDate(m_rsWG_Assignments)
    m_rsWG_Assignments.MoveNext
    If RecordsExist(m_rsWG_Assignments) Then
    Else
        m_rsWG_Assignments.Filter = ""
        OverlappingAssignmentExists = False
        Exit Function
    End If
    
    While Not m_rsWG_Assignments.EOF
        l_datPriorStartDate = GetEffectiveStartDate(m_rsWG_Assignments)
        l_datPriorEndDate = GetEffectiveEndDate(m_rsWG_Assignments)
        l_intIntersectPeriod = DetermineInsertingPeriod( _
                            l_datPriorStartDate, l_datPriorEndDate, _
                            l_datStartDate, l_datEndDate)
        If l_intIntersectPeriod > 0 Then
            l_strErrorDesc = "Overlapping assignment exits. " & _
                "Assignment from " & _
                Format(l_datPriorStartDate, "mm/dd/yyyy") & " to " & Format(l_datPriorEndDate, "mm/dd/yyyy") & _
                " overrlapps the assignment from " & _
                Format(l_datStartDate, "mm/dd/yyyy") & " to " & Format(l_datEndDate, "mm/dd/yyyy")
            l_bln = AppendWGAssignmentsBRV(v_rsWorkGroupAssignments, 30000, "BE", l_strErrorDesc)
            OverlappingAssignmentExists = True
            Exit Function
        End If
        m_rsWG_Assignments.MoveNext
    Wend
    
    Exit Function
    
eh:
    
  TrsRaiseError m_strCMODULENAME, "OverlappingAssignmentExists"
            
End Function
 
 
 
Private Function ProcessWorkGroupAttributes( _
    ByVal v_rsWorkGroupAttributes As ADODB.Recordset _
    ) As Boolean
    
    Dim l_blnError As Boolean
    
    On Error GoTo ProcessWorkGroupAttributes_EH
     
    If RecordsExist(v_rsWorkGroupAttributes) Then
        'continue
    Else
        ProcessWorkGroupAttributes = True
        Exit Function
    End If
    
    If PreValidateWGAttributesRequests(v_rsWorkGroupAttributes) Then
        If m_lngWorkGroup = -1 Then
            v_rsWorkGroupAttributes.MoveFirst
            m_lngWorkGroup = GetNextWorkGroupNumber( _
                v_rsWorkGroupAttributes!DEPT_NUM, _
                v_rsWorkGroupAttributes!LOC_CD, _
                v_rsWorkGroupAttributes!PROJECT_ID)
            If m_lngWorkGroup = -1 Then
                l_blnError = AppendWGAttributesBRV(v_rsWorkGroupAttributes, 30000, mstrCBUSINESSERR, "System error, unable to create work group.  Please contact TRIS Administrator")
                ProcessWorkGroupAttributes = False
                Exit Function
            End If
        End If
        ProcessWorkGroupAttributes = True
    Else
        ProcessWorkGroupAttributes = False
        Exit Function
    End If
    
    v_rsWorkGroupAttributes.MoveFirst
    SetWGAttributes v_rsWorkGroupAttributes

    Exit Function
       
    
ProcessWorkGroupAttributes_EH:
    
    TrsRaiseError m_strCMODULENAME, "ProcessWorkGroupAttributes"
    
    
End Function
 

Private Function GetWorkGroupsInRequest( _
            ByVal v_objRS As ADODB.Recordset, _
            ByVal v_objRS2 As ADODB.Recordset _
            ) As Collection

    Dim l_colWorkGroups As Collection
    Dim l_strWorkGroup As String
    Dim l_aryWorkGroups As Variant

    On Error GoTo eh

    Set l_colWorkGroups = New Collection
    
    If RecordsExist(v_objRS) Then
        v_objRS.MoveFirst
        While Not v_objRS.EOF
            l_strWorkGroup = str(v_objRS!WG_NUM)
            If Not IncludesString(l_strWorkGroup, l_colWorkGroups) Then
                l_colWorkGroups.Add l_strWorkGroup
            End If
            v_objRS.MoveNext
        Wend
    End If
    
    If RecordsExist(v_objRS2) Then
        v_objRS2.MoveFirst
        While Not v_objRS2.EOF
            l_strWorkGroup = str(v_objRS2!WG_NUM)
            If Not IncludesString(l_strWorkGroup, l_colWorkGroups) Then
                l_colWorkGroups.Add l_strWorkGroup
            End If
            v_objRS2.MoveNext
        Wend
    End If

    Set GetWorkGroupsInRequest = l_colWorkGroups
            
    Exit Function
    
eh:
    
  TrsRaiseError m_strCMODULENAME, "GetWorkGroupsInRequest"
  
End Function
 
 
'Private Function ApplyWorkGroupAttributes( _
'            ByVal v_rsRequest As ADODB.Recordset _
'            ) As Boolean
'
'    Dim l_blnIsValid As Boolean
'
'    On Error GoTo eh
'
'    l_blnIsValid = PreValidateWGAttributesRequests(v_rsRequest)
'    If Not l_blnIsValid Then
'        ApplyWorkGroupAttributes = False
'        Exit Function
'    End If
'
'    ApplyWorkGroupAttributesRequests (v_rsRequest)
'
'
'    ApplyWorkGroupAttributes = True
'    Exit Function
'
'eh:
'
'  TrsRaiseError m_strCMODULENAME, "ApplyWorkGroupAttributes"
'
'
'End Function

Private Function PreValidateWGAttributesRequests( _
            ByVal v_rsRequest As ADODB.Recordset _
            ) As Boolean
            
    Dim l_blnIsValid As Boolean
    Dim l_blnIsRequestValid As Boolean
    Dim l_strRecStatus As String
    Dim l_datEffStart As Date
    Dim l_datOrigEffStart As Date
    Dim l_lngWgNum As Long
    
    On Error GoTo eh
    
    l_blnIsValid = True
    v_rsRequest.MoveFirst
    While Not v_rsRequest.EOF
        l_blnIsRequestValid = True
        If Not IsValidRecordStatus_WGA(l_strRecStatus, v_rsRequest) Then
            l_blnIsRequestValid = False
        Else 'continue validation
            If Not IsValidStartDate_WGA(l_strRecStatus, v_rsRequest) Then
                l_blnIsRequestValid = False
            End If
            If Not IsValidWorkGroupType_WGA(l_strRecStatus, v_rsRequest) Then
                l_blnIsRequestValid = False
            End If
            If Not IsValidWorkGroupDepartment_WGA(v_rsRequest) Then
                l_blnIsRequestValid = False
            End If
            If Not IsValidWorkGroupLocation_WGA(v_rsRequest) Then
                l_blnIsRequestValid = False
            End If
            If Not IsValidWorkGroupSupervisor_WGA(v_rsRequest) Then
                l_blnIsRequestValid = False
            End If
            If Not IsValidWorkGroupDescription_WGA(v_rsRequest) Then
                l_blnIsRequestValid = False
            End If
            If Not IsValidWorkGroupProject_WGA(v_rsRequest) Then
                l_blnIsRequestValid = False
            End If
            If l_strRecStatus = "I" Then
                If Not IsValidWorkGroupAssociations_WGA(v_rsRequest) Then
                    l_blnIsRequestValid = False
                End If
            Else
                If DidWGAssociationChange(v_rsRequest) Then
                    If Not IsValidWorkGroupAssociations_WGA(v_rsRequest) Then
                        l_blnIsRequestValid = False
                    End If
                End If
            End If
        End If
        If Not l_blnIsRequestValid Then
            l_blnIsValid = False
        End If
        m_rsWorkGroupAttributes.Filter = ""
        v_rsRequest.MoveNext
    Wend
    
    PreValidateWGAttributesRequests = l_blnIsValid
    Exit Function
    
eh:
    
  TrsRaiseError m_strCMODULENAME, "PreValidateWGAttributesRequests"
    
    
End Function

Private Function IsValidRecordStatus_WGA( _
            ByRef r_strRecStatus As String, _
            ByVal v_rsRequest As ADODB.Recordset _
            ) As Boolean
            
    Dim l_strRecStatus As String
    Dim l_lngWgNum As Long
    Dim l_blnWriteError As Boolean
    
    On Error GoTo eh
    
    If IsNull(v_rsRequest!RECORD_STATUS) Or IsEmpty(v_rsRequest!RECORD_STATUS) Then
        l_strRecStatus = ""
    Else
        l_strRecStatus = v_rsRequest!RECORD_STATUS
    End If
    
    
    Select Case l_strRecStatus
        Case "I"
            IsValidRecordStatus_WGA = True
        Case "C"
            l_lngWgNum = v_rsRequest!WG_NUM
            m_rsWorkGroupAttributes.Filter = "WG_NUM = " & str(l_lngWgNum)
            If Not RecordsExist(m_rsWorkGroupAttributes) Then
                IsValidRecordStatus_WGA = False
                    l_blnWriteError = AppendWGAttributesBRV(v_rsRequest, 30000, mstrCBUSINESSERR, "Invalid change request. Work group " & str(l_lngWgNum) & " not found.")
            Else
                IsValidRecordStatus_WGA = True
            End If
            
        Case "D"
            IsValidRecordStatus_WGA = False
                l_blnWriteError = AppendWGAttributesBRV(v_rsRequest, 30000, mstrCBUSINESSERR, "You can not delete a work group.  Set the effective end date to the last day you want the work group to exist.")
        Case Else
            IsValidRecordStatus_WGA = False
                l_blnWriteError = AppendWGAttributesBRV(v_rsRequest, 30000, mstrCBUSINESSERR, "Invalid Request. Please try again")
        
    End Select
    
    r_strRecStatus = l_strRecStatus
    Exit Function
    
eh:
    
  TrsRaiseError m_strCMODULENAME, "IsValidRecordStatus_WGA"
    
            
            
End Function


Private Function IsValidStartDate_WGA( _
            ByVal v_strRecStatus As String, _
            ByVal v_rsRequest As ADODB.Recordset _
            ) As Boolean
            
'RED209
    
    On Error GoTo eh
    
    Dim l_datEffStart As Date
    Dim l_datOrigEffStart As Date
    Dim l_datToday As Date
    Dim l_lngWgNum As Long
    Dim l_blnWriteError As Boolean
    
    If IsNull(v_rsRequest!EFF_START_DATE) Or IsEmpty(v_rsRequest!EFF_START_DATE) Then
        IsValidStartDate_WGA = False
        l_blnWriteError = AppendWGAttributesBRV(v_rsRequest, 30000, mstrCBUSINESSERR, "The effective start date is missing")
        Exit Function
    End If
    
    If Not IsDate(v_rsRequest!EFF_START_DATE) Then
        IsValidStartDate_WGA = False
        l_blnWriteError = AppendWGAttributesBRV(v_rsRequest, 30000, mstrCBUSINESSERR, "The effective start date is an invalid date")
        Exit Function
    End If
    
'JIMB 08/29/2003
'Allow start date to change

'''''    If v_strRecStatus = "I" Then
'''''        IsValidStartDate_WGA = True
'''''        Exit Function
'''''    End If
'''''
'''''    l_lngWgNum = v_rsRequest!WG_NUM
'''''    l_datEffStart = GetEffectiveStartDate(v_rsRequest)
'''''    l_datToday = Date
'''''    m_rsWorkGroupAttributes.Filter = "WG_NUM = " & str(l_lngWgNum)
'''''    l_datOrigEffStart = m_rsWorkGroupAttributes!EFF_START_DATE
'''''
'''''    If l_datEffStart <> l_datOrigEffStart Then
'''''        If l_datOrigEffStart <= l_datToday Then
'''''    'Current sytem does not let the user change the start date
'''''    ' of a work group already in effect. Should the user be
'''''    ' allowed to change this if no employee are assigned
'''''    ' prior to the new start date?
'''''            IsValidStartDate_WGA = False
'''''                l_blnWriteError = AppendWGAttributesBRV(v_rsRequest, 30000, mstrCBUSINESSERR, "The start date cannot be changed.  It is already in effect.")
'''''            Exit Function
'''''        End If
'''''    End If
    
    IsValidStartDate_WGA = True
    Exit Function
    
eh:
    
  TrsRaiseError m_strCMODULENAME, "IsValidStartDate_WGA"
    
            
 
 End Function


Private Function IsValidWorkGroupType_WGA( _
            ByVal v_strRecStatus As String, _
            ByVal v_rsRequest As ADODB.Recordset _
            ) As Boolean
            
'RED210

    Dim l_strWorkGroupType As String
    Dim l_strOrigWorkGroupType As String
    Dim l_lngWgNum As Long
    Dim l_blnWriteError As Boolean
    
    On Error GoTo eh
    
    If IsNull(v_rsRequest!DEMAND_IND) Or IsEmpty(v_rsRequest!DEMAND_IND) Then
        l_strWorkGroupType = ""
    Else
        l_strWorkGroupType = v_rsRequest!DEMAND_IND
    End If
    
    
    Select Case l_strWorkGroupType
        Case "N" ' Not Demand
            IsValidWorkGroupType_WGA = True
        Case "Y" 'Demand Work Group
            IsValidWorkGroupType_WGA = True
        Case Else
            IsValidWorkGroupType_WGA = False
            l_blnWriteError = AppendWGAttributesBRV(v_rsRequest, 30000, mstrCBUSINESSERR, "Work group type must be (N,R,or D) you entered (" & l_strWorkGroupType & ").")
    End Select
    
    If v_strRecStatus = "C" Then
        l_lngWgNum = v_rsRequest!WG_NUM
        m_rsWorkGroupAttributes.Filter = "WG_NUM = " & str(l_lngWgNum)
        l_strOrigWorkGroupType = m_rsWorkGroupAttributes!DEMAND_IND
        If l_strWorkGroupType <> l_strOrigWorkGroupType Then
             IsValidWorkGroupType_WGA = False
             l_blnWriteError = AppendWGAttributesBRV(v_rsRequest, 30000, mstrCBUSINESSERR, "Work group type can not be changed")
        End If
    End If
    
    
    Exit Function
    
eh:
    
  TrsRaiseError m_strCMODULENAME, "IsValidWorkGroupType_WGA"
            
End Function

Private Function IsValidWorkGroupDescription_WGA( _
            ByVal v_rsRequest As ADODB.Recordset _
            ) As Boolean
            
'RED213

    Dim l_strWorkDescription As String
    Dim l_lngWgNum As Long
    Dim l_blnWriteError As Boolean
    
    On Error GoTo eh
    
    If IsNull(v_rsRequest!WG_DESC) Or IsEmpty(v_rsRequest!WG_DESC) Then
        l_strWorkDescription = ""
    Else
        l_strWorkDescription = Trim(v_rsRequest!WG_DESC)
    End If
    
    If l_strWorkDescription = "" Then
        IsValidWorkGroupDescription_WGA = False
        l_blnWriteError = AppendWGAttributesBRV(v_rsRequest, 30000, mstrCBUSINESSERR, "Work group description must be entered.")
    Else
        IsValidWorkGroupDescription_WGA = True
    End If
    
    Exit Function
    
eh:
    
  TrsRaiseError m_strCMODULENAME, "IsValidWorkGroupDescription_WGA"
            
End Function


Private Function IsValidWorkGroupSupervisor_WGA( _
            ByVal v_rsRequest As ADODB.Recordset _
            ) As Boolean
'It is valid if the following are true
' WHITE127 -- Supervisor is a valid employeee

    Dim l_lngEmpNum As Long
    Dim l_lngWgNum As Long
    Dim l_blnWriteError As Boolean
    
    On Error GoTo eh
    
    If IsNull(v_rsRequest!WG_SUPERVISOR) Or IsEmpty(v_rsRequest!WG_SUPERVISOR) Then
        l_lngEmpNum = 0
    Else
        l_lngEmpNum = v_rsRequest!WG_SUPERVISOR
    End If
    
    If l_lngEmpNum = 0 Then
        IsValidWorkGroupSupervisor_WGA = False
        l_blnWriteError = AppendWGAttributesBRV(v_rsRequest, 30000, mstrCBUSINESSERR, "A Work group supervisor must be entered.")
        Exit Function
    End If
    
    m_rsWG_Supervisors.Filter = "EMP_NUM = " & str(l_lngEmpNum)
    If RecordsExist(m_rsWG_Supervisors) Then
        IsValidWorkGroupSupervisor_WGA = True
    Else
        IsValidWorkGroupSupervisor_WGA = False
        l_blnWriteError = AppendWGAttributesBRV(v_rsRequest, 30000, mstrCBUSINESSERR, "A Work group supervisor does not exist.")
    End If
    
    m_rsWG_Supervisors.Filter = ""
    
    Exit Function
    
eh:
    
  TrsRaiseError m_strCMODULENAME, "IsValidWorkGroupSupervisor_WGA"
            
End Function


Private Function IsValidWorkGroupProject_WGA( _
            ByVal v_rsRequest As ADODB.Recordset _
            ) As Boolean
            
'RED214

    Dim l_lngProject As Long
    Dim l_lngWgNum As Long
    Dim l_blnWriteError As Boolean
    
    On Error GoTo eh
    
    If IsNull(v_rsRequest!PROJECT_ID) Or IsEmpty(v_rsRequest!PROJECT_ID) Then
        l_lngProject = 0
    Else
        l_lngProject = v_rsRequest!PROJECT_ID
    End If
    
    If l_lngProject = 0 Then
        IsValidWorkGroupProject_WGA = True
        Exit Function
    End If
    
    Set m_rsProjects = GetPPMProject(l_lngProject)
    If Not RecordsExist(m_rsProjects) Then
        IsValidWorkGroupProject_WGA = False
        l_blnWriteError = AppendWGAttributesBRV(v_rsRequest, 30000, mstrCBUSINESSERR, "The project (" & str(l_lngProject) & ") does not exist.")
    Else
        IsValidWorkGroupProject_WGA = True
         m_rsProjects.Filter = ""
    End If
    
    
    Exit Function
    
eh:
    
  TrsRaiseError m_strCMODULENAME, "IsValidWorkGroupProject_WGA"
            
End Function


Private Function IsValidWorkGroupDepartment_WGA( _
            ByVal v_rsRequest As ADODB.Recordset _
            ) As Boolean
            
    Dim l_strDept As String
    Dim l_lngWgNum As Long
    Dim l_obj As dbDepartment.cdDepartment
    Dim l_blnWriteError As Boolean
    
    On Error GoTo eh
    
    If IsNull(v_rsRequest!DEPT_NUM) Or IsEmpty(v_rsRequest!DEPT_NUM) Then
        l_strDept = ""
    Else
        l_strDept = Trim(v_rsRequest!DEPT_NUM)
    End If
    
    If l_strDept = "" Then
        IsValidWorkGroupDepartment_WGA = False
        l_blnWriteError = AppendWGAttributesBRV(v_rsRequest, 30000, mstrCBUSINESSERR, "The department associated to the work group was not entered. Please enter a department.")
        Exit Function
    End If
    
    Set l_obj = CreateObject("dbDepartment.cdDepartment")
    If l_obj.IsDepartmentValid(l_strDept) Then
        IsValidWorkGroupDepartment_WGA = True
    Else
        IsValidWorkGroupDepartment_WGA = False
        l_blnWriteError = AppendWGAttributesBRV(v_rsRequest, 30000, mstrCBUSINESSERR, "The department (" & l_strDept & ") does not exist.")
    End If
    
    Exit Function
    
eh:
    
  TrsRaiseError m_strCMODULENAME, "IsValidWorkGroupDepartment_WGA"
            
End Function

Private Function IsValidWorkGroupLocation_WGA( _
            ByVal v_rsRequest As ADODB.Recordset _
            ) As Boolean
            
    Dim l_lngLocationCode As Long
    Dim l_lngWgNum As Long
    Dim l_obj As dbCodeTables.cdCodeTables
    Dim l_blnWriteError As Boolean
    
    On Error GoTo eh
    
    If IsNull(v_rsRequest!LOC_CD) Or IsEmpty(v_rsRequest!LOC_CD) Then
        l_lngLocationCode = 0
    Else
        l_lngLocationCode = v_rsRequest!LOC_CD
    End If
    
    If l_lngLocationCode = 0 Then
        IsValidWorkGroupLocation_WGA = False
        l_blnWriteError = AppendWGAttributesBRV(v_rsRequest, 30000, mstrCBUSINESSERR, "The location associated to the work group was not entered. Please enter a location.")
        Exit Function
    End If
    
    Set l_obj = CreateObject("dbCodeTables.cdCodeTables")
    If l_obj.LocationCodeIsValid(str(l_lngLocationCode)) Then
        IsValidWorkGroupLocation_WGA = True
    Else
        IsValidWorkGroupLocation_WGA = False
        l_blnWriteError = AppendWGAttributesBRV(v_rsRequest, 30000, mstrCBUSINESSERR, "The location (" & str(l_lngLocationCode) & ") does not exist.")
    End If
    
    Exit Function
    
eh:
    GetObjectContext.SetAbort
  TrsRaiseError m_strCMODULENAME, "IsValidWorkGroupLocation_WGA"
            
End Function


Private Function IsValidWorkGroupWeeksToSchedule_WGA( _
            ByVal v_rsRequest As ADODB.Recordset _
            ) As Boolean
            
'This Validation is no longer needed
'RED215
            
    Dim l_strWorkGroupType As String
    Dim l_lngWeeksToSchedule As Long
    Dim l_lngWgNum As Long
    Dim l_blnWriteError As Boolean
    
    On Error GoTo eh
    
    If IsNull(v_rsRequest!DEMAND_IND) Or IsEmpty(v_rsRequest!DEMAND_IND) Then
        l_strWorkGroupType = ""
    Else
        l_strWorkGroupType = v_rsRequest!DEMAND_IND
    End If
    
    If l_strWorkGroupType <> "D" Then
        IsValidWorkGroupWeeksToSchedule_WGA = True
        Exit Function
    End If
    
    If IsNull(v_rsRequest!REC_WEEKS_TO_SCHED) Or IsEmpty(v_rsRequest!REC_WEEKS_TO_SCHED) Then
        l_lngWeeksToSchedule = 0
    Else
        l_lngWeeksToSchedule = v_rsRequest!REC_WEEKS_TO_SCHED
    End If
    
    If l_lngWeeksToSchedule = 0 Then
        IsValidWorkGroupWeeksToSchedule_WGA = False
        l_blnWriteError = AppendWGAttributesBRV(v_rsRequest, 30000, mstrCBUSINESSERR, "The weeks to schedule for the demand work group was not entered. Please enter weeks to schedule greater than zero.")
        Exit Function
    End If

    Exit Function
    
eh:
    
  TrsRaiseError m_strCMODULENAME, "IsValidWorkGroupWeeksToSchedule_WGA"
            
End Function

Private Function IsValidWorkGroupAssociations_WGA( _
            ByVal v_rsRequest As ADODB.Recordset _
            ) As Boolean
'It is valid if the following are true
' WHITE127 -- Supervisor is a valid employeee
' RED216 -- IsValidWorkGroupAssociations

    Dim l_lngEmpNum As Long
    Dim l_lngWgNum As Long
    Dim l_strWGDept As String
    Dim l_lngWGlocationcode As Long
    Dim l_strWGFun As String
    Dim l_strWGDiv As String
    Dim l_strRecStatus As String
    Dim l_strSupervisorDept As String
    Dim l_strSupervisorFun As String
    Dim l_strSupervisorDiv As String
    Dim l_blnWriteError As Boolean
    Dim l_blnIsEmpAnOfficer As Boolean
    
    On Error GoTo eh
    
    If IsNull(v_rsRequest!WG_SUPERVISOR) Or IsEmpty(v_rsRequest!WG_SUPERVISOR) Then
        l_lngEmpNum = 0
    Else
        l_lngEmpNum = v_rsRequest!WG_SUPERVISOR
    End If
    
    If l_lngEmpNum = 0 Then
        IsValidWorkGroupAssociations_WGA = False
        GoTo xit
    End If
    
    m_rsWG_Supervisors.Filter = "EMP_NUM = " & str(l_lngEmpNum)
    If RecordsExist(m_rsWG_Supervisors) Then
        l_strSupervisorDept = m_rsWG_Supervisors!DEPT_NUM
        l_strSupervisorDiv = m_rsWG_Supervisors!DIVISION_NUM
        l_strSupervisorFun = m_rsWG_Supervisors!FUNCTION_NUM
        If Trim(m_rsWG_Supervisors!OFFCR_FL) = "Y" Then
            l_blnIsEmpAnOfficer = True
        Else
            l_blnIsEmpAnOfficer = False
        End If
    Else
        IsValidWorkGroupAssociations_WGA = False
        GoTo xit
    End If
    
    'Is the divison and function of the supervisor
    ' the same as the divison and function of the department
    ' and location associated with the work group
    
    'Get Department
    If IsNull(v_rsRequest!DEPT_NUM) Or IsEmpty(v_rsRequest!DEPT_NUM) Then
        l_strWGDept = ""
    Else
        l_strWGDept = Trim(v_rsRequest!DEPT_NUM)
    End If
    
    'Get Location
    If IsNull(v_rsRequest!LOC_CD) Or IsEmpty(v_rsRequest!LOC_CD) Then
        l_lngWGlocationcode = 0
    Else
        l_lngWGlocationcode = v_rsRequest!LOC_CD
    End If
    
    If l_strWGDept = "" Or l_lngWGlocationcode = 0 Then
        IsValidWorkGroupAssociations_WGA = False
        GoTo xit
        'No need to write a BRV because it has already been done
    End If
    
    GetFunctionAndDivisionOFAssociation l_strWGFun, l_strWGDiv, l_strWGDept, l_lngWGlocationcode
    
    If l_strWGFun = "" Or l_strWGDiv = "" Then
        IsValidWorkGroupAssociations_WGA = False
        l_blnWriteError = AppendWGAttributesBRV(v_rsRequest, 30000, mstrCBUSINESSERR, "Invalid location (" & str(l_lngWGlocationcode) & ") for department (" & l_strWGDept & ").  The function and division were not found")
        GoTo xit
    End If
    
    'If the Supervisor can supervise any work group
    '  the association is valid
'    If CanSuperviseAnyWorkGroup(l_strSupervisorDept) Then
    If l_blnIsEmpAnOfficer Then
        IsValidWorkGroupAssociations_WGA = True
        GoTo xit
    End If
    
    If l_strWGFun <> l_strSupervisorFun Then
        IsValidWorkGroupAssociations_WGA = False
        l_blnWriteError = AppendWGAttributesBRV(v_rsRequest, 30000, mstrCBUSINESSERR, "The Supervisor's Function does not match the function for the work group's department and location.")
        GoTo xit
    End If
    
    If l_strWGDiv <> l_strSupervisorDiv Then
        IsValidWorkGroupAssociations_WGA = False
        l_blnWriteError = AppendWGAttributesBRV(v_rsRequest, 30000, mstrCBUSINESSERR, "The Supervisor's Division does not match the function for the work group's department and location.")
    Else
        IsValidWorkGroupAssociations_WGA = True
    End If
    
xit:
    m_rsWG_Supervisors.Filter = ""
    Exit Function
    
eh:
    
  TrsRaiseError m_strCMODULENAME, "IsValidWorkGroupAssociations_WGA"
            
End Function

Private Sub GetFunctionAndDivisionOFAssociation( _
        ByRef r_strFunction As String, _
        ByRef r_strDivison As String, _
        ByVal v_strDepartment As String, _
        ByVal v_lngLocation As Long)
        
    Dim l_obj As dbCodeTables.cdCodeTables
    Dim l_rs As ADODB.Recordset
    Dim l_colDept As Collection
    
    On Error GoTo eh
    
    Set l_colDept = New Collection
    l_colDept.Add v_strDepartment
    Set l_obj = CreateObject("dbCodeTables.cdCodeTables")
    Set l_rs = l_obj.GetLocationsAssociatedToDepartment(l_colDept)
    
    l_rs.Filter = "LOC_CD = " & str(v_lngLocation)
    If Not RecordsExist(l_rs) Then
        r_strFunction = ""
        r_strDivison = ""
    Else
        r_strFunction = l_rs!FUNCTION_NUM
        r_strDivison = l_rs!DIVISION_NUM
    End If
    Set l_rs = Nothing
    Set l_obj = Nothing
      
    Exit Sub
    
eh:
    GetObjectContext.SetAbort
  TrsRaiseError m_strCMODULENAME, "IsValidWorkGroupSupervisor_WGA"
               
End Sub
Private Function GetNextWorkGroupNumber( _
            ByVal v_strDept As String, _
            ByVal v_lngLoc As Long, _
            ByVal v_lngProjectId As Long _
            ) As Long
            
'RED218
            
    Dim l_strSequenceType As String
    Dim l_obj As utilDBDataServices.cuDBDataServices
    Dim l_lngWgNum As Long
    
    On Error GoTo eh
    
    l_strSequenceType = GetWGSequenceType(v_strDept, v_lngLoc, v_lngProjectId)
    
    If Trim(l_strSequenceType) <> "" Then
        Set l_obj = CreateObject("utilDBDataServices.cuDBDataServices")
        GetNextWorkGroupNumber = l_obj.GetNextSequenceNumber(l_strSequenceType)
    Else
        GetNextWorkGroupNumber = -1
        'Raise an Error because of bug in not returning Sequence Type
    End If
        
    Exit Function
    
eh:
    
  TrsRaiseError m_strCMODULENAME, "GetNextWorkGroupNumber"
           
    
            
End Function


Private Function GetWGSequenceType( _
            ByVal v_strDept As String, _
            ByVal v_lngLoc As Long, _
            ByVal v_lngProjectId As Long _
            ) As String
           
'RED219
'The sequence type is "WG" unless it is an AmerenUE power plant
' (excluding Callaway) or a DOJM work group.

    On Error GoTo eh
    
    If IsAmerenUEPlant(v_lngLoc) Or _
        IsDOJM(v_strDept) Or _
        IsPPM(v_strDept) Or _
        v_lngProjectId > 0 Then
    Else
        GetWGSequenceType = "WG"
        Exit Function
    End If
    
    If IsPPM(v_strDept) Or v_lngProjectId > 0 Then
        GetWGSequenceType = "PM"
        Exit Function
    End If
    
    If IsAmerenUEPlant(v_lngLoc) Then
        GetWGSequenceType = GetWGSequenceType_UEPlant(v_lngLoc)
        Exit Function
    End If
    
 'Only DOJM Departments Left
    GetWGSequenceType = GetWGSequenceType_DOJM(v_strDept)
    Exit Function
    
eh:
    
  TrsRaiseError m_strCMODULENAME, "GetWGSequenceType"
 
End Function



Private Function IsDOJM( _
            ByVal v_strDept As String _
            ) As Boolean
           
    Dim l_obj As busDepartment.cbsDepartment
                
    On Error GoTo eh
    
    Set l_obj = CreateObject("busDepartment.cbsDepartment")
    IsDOJM = l_obj.IsDOJMDepartmentNumber(v_strDept)
    Set l_obj = Nothing
    
    Exit Function
    
eh:
    
  TrsRaiseError m_strCMODULENAME, "IsDOJM"
                       
                    
End Function


Private Function DetermineCrewNumber( _
            ByVal v_lngWgNum As Long, _
            ByVal v_strSeqType As String, _
            ByVal v_strDept As String _
            ) As String
                          
'RED221
                          
    On Error GoTo eh
    
    If IsDOJM(v_strDept) Then
        If IsUEDojmWG(v_lngWgNum) Then
            DetermineCrewNumber = ""
        Else
            DetermineCrewNumber = Left(v_strSeqType, 1) & _
                Right(str(v_lngWgNum), 2)
        End If
    Else
        DetermineCrewNumber = ""
    End If
    
    Exit Function
    
eh:
    
  TrsRaiseError m_strCMODULENAME, "DetermineCrewNumber"
                       
                    
End Function


Private Function CanWorkGroupEndOnDate( _
            ByVal v_rsRequest As ADODB.Recordset, _
            ByVal v_lngWgNum As Long, _
            ByVal v_datEndDate As Date _
            ) As Boolean
            
    Dim l_blnWriteError As Boolean
    Dim l_datStart As Date
            
    On Error GoTo eh
            
    If v_datEndDate = CDate("12/31/9999") Then
        CanWorkGroupEndOnDate = True
        Exit Function
    End If
            
    If IsSystemDefinedWorkGroup(v_lngWgNum) Then
        CanWorkGroupEndOnDate = False
        l_blnWriteError = AppendWGAttributesBRV(v_rsRequest, 30000, mstrCBUSINESSERR, "Can not end a system work group.")
        Exit Function
    End If
    
    l_datStart = GetWorkGroupStartDate()
    If v_datEndDate < l_datStart Then
        CanWorkGroupEndOnDate = False
        l_blnWriteError = AppendWGAttributesBRV(v_rsRequest, 30000, mstrCBUSINESSERR, "The end date of the work group can not be less than the start date.")
        Exit Function
    End If

    If AnyEmployeeStillPrimary(v_lngWgNum, v_datEndDate) Then
        CanWorkGroupEndOnDate = False
                l_blnWriteError = AppendWGAttributesBRV(v_rsRequest, 30000, mstrCBUSINESSERR, "Can not end the work group because there is at least one employee still primary.")
        Exit Function
    End If
    
    If AnyEmployeeStillReliefing(v_lngWgNum, v_datEndDate) Then
        CanWorkGroupEndOnDate = False
                l_blnWriteError = AppendWGAttributesBRV(v_rsRequest, 30000, mstrCBUSINESSERR, "Can not end the work group because there is at least one employee assigned via relief.")
        Exit Function
    End If
    
    CanWorkGroupEndOnDate = True
    Exit Function
    
eh:
    
  TrsRaiseError m_strCMODULENAME, "CanWorkGroupEndOnDate"
                           
            
End Function


Private Function AnyEmployeeStillPrimary( _
            ByVal v_lngWgNum As Long, _
            ByVal v_datEndDate As Date _
            ) As Boolean
            
    Dim l_strFilter As String
            
    On Error GoTo eh
    
    
    l_strFilter = "WG_NUM = " & str(v_lngWgNum) & _
                " and EFF_END_DATE > " & str(v_datEndDate) & _
                " and PRIMARY_IND = 'Y'"
                
    l_strFilter = MDAC25FilterFix(l_strFilter)
                
    m_rsWG_Assignments.Filter = l_strFilter
    
    If RecordsExist(m_rsWG_Assignments) Then
        AnyEmployeeStillPrimary = True
    Else
        AnyEmployeeStillPrimary = False
    End If
    
    m_rsWG_Assignments.Filter = ""
    Exit Function
    
eh:
    
  TrsRaiseError m_strCMODULENAME, "AnyEmployeeStillPrimary"
                       
            
End Function

Private Function AnyEmployeeStillReliefing( _
            ByVal v_lngWgNum As Long, _
            ByVal v_datEndDate As Date _
            ) As Boolean
    Dim l_strFilter As String
            
    On Error GoTo eh
    
    
    l_strFilter = "WG_NUM = " & str(v_lngWgNum) & _
                " and EFF_END_DATE > " & str(v_datEndDate) & _
                " and SUBSYS_OWNER  = 'R'"
                
    l_strFilter = MDAC25FilterFix(l_strFilter)
                
    m_rsWG_Assignments.Filter = l_strFilter
    
    If RecordsExist(m_rsWG_Assignments) Then
        AnyEmployeeStillReliefing = True
    Else
        AnyEmployeeStillReliefing = False
    End If
    
    m_rsWG_Assignments.Filter = ""
    Exit Function
    
eh:
    
  TrsRaiseError m_strCMODULENAME, "AnyEmployeeStillReliefing"
                                   
            
End Function

' RSG - commented out during integration
'Private Sub ClearDemandLineAssignments( _
'            ByVal v_lngWGNum As Long, _
'            ByVal v_datEndDate As Date)
'
''RED252, RED253
'
'    Dim l_strFilter As String
'    Dim l_datNextSaturday As Date
'    Dim l_intDayOfWeek As Integer
'
'    On Error GoTo eh
'
'    Set m_rsWGDemandLineAssignments = GetDemandLineAssignments(v_lngWGNum, v_datEndDate)
'
'    If Not RecordsExist(m_rsWGDemandLineAssignments) Then
'        Exit Sub
'    End If
'
'    l_datNextSaturday = DateAdd("d", 7 - Weekday(v_datEndDate), v_datEndDate)
'    l_intDayOfWeek = Weekday(v_datEndDate)
'    l_strFilter = "WG_NUM = " & str(v_lngWGNum) & _
'                "END_DATE = " & l_datNextSaturday
'
'
'
'    Exit Sub
'eh:
'
'  TrsRaiseError m_strCMODULENAME, "ClearDemandLineAssignments"
'
'
'End Sub
'
'Private Function GetDemandLineAssignments( _
'            ByVal v_lngWGNum As Long, _
'            ByVal v_datEndDate As Date _
'            ) As ADODB.Recordset
'
'    Dim l_obj As dbWorkGroup.cdWorkGroup
'
'    On Error GoTo eh
'
'    Set l_obj = New dbWorkGroup.cdWorkGroup
'    Set GetDemandLineAssignments = l_obj.GetWGDemandLineAssigmments(v_lngWGNum, v_datEndDate)
'    Set l_obj = Nothing
'
'    Exit Function
'
'eh:
'
'  TrsRaiseError m_strCMODULENAME, "GetDemandLineAssignments"
'
'
'End Function


'Private Sub ApplyWorkGroupAttributesRequests( _
'    ByVal v_rsWorkGroupAssignments As ADODB.Recordset)
'
'    Dim l_strRecStatus As String
'    Dim l_blnRequestProcessed As Boolean
'
'On Error GoTo eh
'
'    v_rsWorkGroupAssignments.MoveFirst
'    While Not v_rsWorkGroupAssignments.EOF
'        l_strRecStatus = UCase(v_rsWorkGroupAssignments!RECORD_STATUS)
'        Select Case l_strRecStatus
'            Case "I"
'                AdjustAdjacentPrimary v_rsWorkGroupAssignments
'                l_blnRequestProcessed = AdjustedOverlappingAssignments(v_rsWorkGroupAssignments)
'                If l_blnRequestProcessed Then
'                Else
'                    AssignmentInsert v_rsWorkGroupAssignments
'                End If
'            Case "C"
'                AdjustAdjacentPrimary v_rsWorkGroupAssignments
'                AssignmentUpdate v_rsWorkGroupAssignments
'            Case "D", "Z"
'                AssignmentDelete v_rsWorkGroupAssignments
'        End Select
'
'        v_rsWorkGroupAssignments.MoveNext
'    Wend
'
'    Exit Sub
'
'eh:
'
'  TrsRaiseError m_strCMODULENAME, "ApplyWorkGroupAttributesRequests"
'
'End Sub


Private Function GetWorkGroupAssociations( _
            ) As ADODB.Recordset

    Dim l_aryWorkGroups As Variant
    Dim l_rs As ADODB.Recordset
    Dim l_obj As dbWorkGroup.cdWorkGroup
    
    On Error GoTo eh

    If m_colWorkGroups.Count = 0 Then
        Set GetWorkGroupAssociations = Nothing
        Exit Function
    End If

    l_aryWorkGroups = ArrayFromCollection(m_colWorkGroups)
    Set l_obj = CreateObject("dbWorkGroup.cdWorkGroup")
    Set l_rs = l_obj.GetWorkGroupAssociations(l_aryWorkGroups)
    Set GetWorkGroupAssociations = l_rs
    
    Exit Function
    
eh:
    GetObjectContext.SetAbort
  TrsRaiseError m_strCMODULENAME, "GetWorkGroupAssociations"


End Function
            
            
Private Sub SetGlobalObjectsToNothing()

    Set m_rsErrors = Nothing
    Set m_colEmployees = Nothing
    Set m_rsWorkGroupAttributes = Nothing
    Set m_rsWG_Assignments = Nothing
    Set m_rsWG_PAIOWG = Nothing
    Set m_rsWG_ALL_PA = Nothing
    Set m_colWorkGroups = Nothing
    Set m_rsWG_Supervisors = Nothing
    Set m_rsProjects = Nothing
    Set m_rsWGDemandLineSchedules = Nothing
    Set m_rsWGBaseSched = Nothing
    Set m_rsWorkGroupAssociations = Nothing
    Set m_rsWGDemandLineAssignments = Nothing
    Set m_rsWG_Assignments_Copy = Nothing
    Set m_colCascadeLineIntervals = Nothing
    Set m_colCascadeEmpIntervals = Nothing
    
End Sub

Private Function GetStateOfWorkGroup( _
            ByRef v_rsWorkGroupAssignments As ADODB.Recordset, _
            ByRef v_rsWorkGroupAttributes As ADODB.Recordset, _
            ByRef v_rsLineScheduleRequest As ADODB.Recordset _
            ) As Boolean

    Dim l_dat1YearPrior As Date
    
    Dim l_datMinEffDate As Date
    Dim l_datMaxEffDate As Date
    Dim l_varLineNums As Variant
    
    On Error GoTo eh
    
    Set m_rsWorkGroupAttributes = GetWorkGroupInfo(m_lngWorkGroup)
    m_blnDemandWorkGroup = IsADemandWorkGroup(v_rsWorkGroupAttributes)
    If m_blnDemandWorkGroup Then
    Else
        If m_lngWorkGroup = -1 And Not RecordsExist(v_rsLineScheduleRequest) Then
            GetStateOfWorkGroup = False
            AppendWGSchedError m_rsErrors, _
            trErrbusValidationError, mstrCBUSINESSERR, _
            "Base schedule does not exist."
            Exit Function
        End If
    End If
    
    l_dat1YearPrior = Date
    l_dat1YearPrior = DateAdd("yyyy", -1, l_dat1YearPrior)
    Set m_colEmployees = GetEmployeesInRequest(v_rsWorkGroupAssignments)
    Set m_rsWG_Assignments = GetEmployeeAssignmentsInWorkGroup(m_lngWorkGroup, l_dat1YearPrior)
    m_rsWG_Assignments.Filter = "PRIMARY_IND = 'Y'"
    Set m_rsWG_Assignments_Copy = CopyRecordset(m_rsWG_Assignments, 2)
    m_rsWG_Assignments.Filter = 0
    Set m_rsWG_PAIOWG = GetEmployeePrimaryAssignments(l_dat1YearPrior)
    Set m_rsWG_Supervisors = GetSupervisorInformation(v_rsWorkGroupAttributes)
'    Set m_rsWGDemandLineSchedules = GetWorkGroupLineSched(m_lngWorkGroup, l_dat1YearPrior)
    Set m_rsWGBaseSched = GetWorkGroupBaseSched(m_lngWorkGroup, 0)

'Ricks code below failed when there was no schedule request
' so I put an if statement around it
    If RecordsExist(v_rsLineScheduleRequest) And m_blnDemandWorkGroup Then
        DetermineDemandSchedScope _
            v_rsLineScheduleRequest, l_datMinEffDate, l_datMaxEffDate, l_varLineNums
        Set m_rsWGDemandLineSchedules = _
            GetDemandLineSchedules( _
                m_lngWorkGroup, l_datMinEffDate, l_datMaxEffDate, l_varLineNums)
        Set m_rsHoldWGDemandLineSchedules = CopyRecordset(m_rsWGDemandLineSchedules)
        l_datMaxEffDate = DetermineGetAssignmentMaxEnd(m_rsWGDemandLineSchedules)
        Set m_rsWGDemandLineAssignments = _
            GetDemandLineAssignments(m_lngWorkGroup, l_datMinEffDate, l_datMaxEffDate)
        m_blnWGLineSchedChg = True
    Else
'No changes made to schedule
'Get state that is on client for a post validation, PostValidateWGSave
        m_blnWGLineSchedChg = False
        If m_blnDemandWorkGroup Then
            l_datMinEffDate = DateAdd("m", -1, Date)
            Set m_rsWGDemandLineSchedules = GetWorkGroupLineSched(m_lngWorkGroup, l_datMinEffDate)
        Else
            Set m_rsWGDemandLineSchedules = Nothing
        End If
    End If

    Set m_rsWGDistribution = GetWorkGroupDistribution(m_lngWorkGroup)
    Set m_rsWGActivities = GetWorkGroupActivites(m_lngWorkGroup)
    
    GetStateOfWorkGroup = True

    Exit Function
    
        
eh:
    
  TrsRaiseError m_strCMODULENAME, "GetStateOfWorkGroup"
    
End Function

Private Function GetSupervisorInformation( _
            ByVal v_rsWorkGroupAttributes As ADODB.Recordset _
            ) As ADODB.Recordset
            
    Dim l_obj As dbEmployee.cdEmployee
    Dim l_lngEmpNum As Long
           
    On Error GoTo eh
    
    If RecordsExist(v_rsWorkGroupAttributes) Then
        If IsNull(v_rsWorkGroupAttributes!WG_SUPERVISOR) Or _
            IsEmpty(v_rsWorkGroupAttributes!WG_SUPERVISOR) Then
            l_lngEmpNum = 0
        Else
            l_lngEmpNum = v_rsWorkGroupAttributes!WG_SUPERVISOR
        End If
    Else
        l_lngEmpNum = 0
    End If
    
    
    Set l_obj = CreateObject("dbEmployee.cdEmployee")
    Set GetSupervisorInformation = l_obj.GetEmployee(l_lngEmpNum)
    Set l_obj = Nothing
                
    Exit Function
    
eh:
    GetObjectContext.SetAbort
  TrsRaiseError m_strCMODULENAME, "GetSupervisorInformation"
            
End Function

Private Sub SetNewWorkGroupNumber( _
    ByVal v_rsWorkGroupAttributes As ADODB.Recordset, _
    ByVal v_rsWorkGroupSchedule As ADODB.Recordset, _
    ByVal v_rsWorkGroupAssignments As ADODB.Recordset, _
    ByVal v_rsWorkGroupDistributionList As ADODB.Recordset, _
    ByVal v_rsWorkGroupDistributionActivities As ADODB.Recordset)
                
                
    On Error GoTo eh
                
    v_rsWorkGroupAttributes!WG_NUM = m_lngWorkGroup
    If RecordsExist(v_rsWorkGroupSchedule) Then
        v_rsWorkGroupSchedule.MoveFirst
        While Not v_rsWorkGroupSchedule.EOF
            v_rsWorkGroupSchedule!WG_NUM = m_lngWorkGroup
            v_rsWorkGroupSchedule!KEY_WG_NUM = m_lngWorkGroup
            v_rsWorkGroupSchedule.MoveNext
        Wend
    End If
    
    If RecordsExist(v_rsWorkGroupAssignments) Then
        v_rsWorkGroupAssignments.MoveFirst
        While Not v_rsWorkGroupAssignments.EOF
            v_rsWorkGroupAssignments!WG_NUM = m_lngWorkGroup
            v_rsWorkGroupAssignments.MoveNext
        Wend
    End If
    
    If RecordsExist(v_rsWorkGroupDistributionList) Then
        v_rsWorkGroupDistributionList.MoveFirst
        While Not v_rsWorkGroupDistributionList.EOF
            v_rsWorkGroupDistributionList!WG_NUM = m_lngWorkGroup
            v_rsWorkGroupDistributionList!KEY_WG_NUM = m_lngWorkGroup
            v_rsWorkGroupDistributionList.MoveNext
        Wend
    End If

    If RecordsExist(v_rsWorkGroupDistributionActivities) Then
        v_rsWorkGroupDistributionActivities.MoveFirst
        While Not v_rsWorkGroupDistributionActivities.EOF
            v_rsWorkGroupDistributionActivities!WG_NUM = m_lngWorkGroup
            v_rsWorkGroupDistributionActivities!KEY_WG_NUM = m_lngWorkGroup
            v_rsWorkGroupDistributionActivities.MoveNext
        Wend
    End If
    
    Exit Sub
    
eh:
    
  TrsRaiseError m_strCMODULENAME, "SetNewWorkGroupNumber"
  
End Sub

Private Function ARequestExist( _
    ByVal v_rsWorkGroupAttributes As ADODB.Recordset, _
    ByVal v_rsWorkGroupSchedule As ADODB.Recordset, _
    ByVal v_rsWorkGroupAssignments As ADODB.Recordset, _
    ByVal v_rsWorkGroupDistributionList As ADODB.Recordset, _
    ByVal v_rsWorkGroupDistributionActivities As ADODB.Recordset) As Boolean

    On Error GoTo eh

    ARequestExist = RecordsExist(v_rsWorkGroupAttributes) _
        Or RecordsExist(v_rsWorkGroupSchedule) _
        Or RecordsExist(v_rsWorkGroupAssignments) _
        Or RecordsExist(v_rsWorkGroupDistributionList) _
        Or RecordsExist(v_rsWorkGroupDistributionActivities)

    Exit Function
    
eh:
    
  TrsRaiseError m_strCMODULENAME, "ARequestExist"
  
End Function

Private Function DidWGAssociationChange( _
    ByVal v_rsRequest As ADODB.Recordset _
    ) As Boolean
    
    
    Dim l_strDept As String
    Dim l_lngLocationCode As Long
    Dim l_lngEmpNum As Long
    Dim l_strPriorDept As String
    Dim l_lngPriorLocationCode As Long
    Dim l_lngPriorEmpNum As Long
    
    
    On Error GoTo eh
    
        
    If IsNull(v_rsRequest!DEPT_NUM) Or IsEmpty(v_rsRequest!DEPT_NUM) Then
        l_strDept = m_rsWorkGroupAttributes!DEPT_NUM
    Else
        l_strDept = Trim(v_rsRequest!DEPT_NUM)
    End If
    
    If IsNull(v_rsRequest!LOC_CD) Or IsEmpty(v_rsRequest!LOC_CD) Then
        l_lngLocationCode = m_rsWorkGroupAttributes!LOC_CD
    Else
        l_lngLocationCode = v_rsRequest!LOC_CD
    End If
    
    If IsNull(v_rsRequest!WG_SUPERVISOR) Or IsEmpty(v_rsRequest!WG_SUPERVISOR) Then
        l_lngEmpNum = m_rsWorkGroupAttributes!WG_SUPERVISOR
    Else
        l_lngEmpNum = v_rsRequest!WG_SUPERVISOR
    End If
    
    l_strPriorDept = m_rsWorkGroupAttributes!DEPT_NUM
    l_lngPriorLocationCode = m_rsWorkGroupAttributes!LOC_CD
    l_lngPriorEmpNum = m_rsWorkGroupAttributes!WG_SUPERVISOR
    
    If l_strDept = l_strPriorDept And _
        l_lngLocationCode = l_lngPriorLocationCode And _
        l_lngEmpNum = l_lngPriorEmpNum Then
        DidWGAssociationChange = False
    Else
        DidWGAssociationChange = True
    End If
    
    Exit Function
    
eh:
    
  TrsRaiseError m_strCMODULENAME, "DidWGAssociationChange"
    
    
End Function

Private Sub SetWGAttributes( _
    ByVal v_rsWorkGroupAttributes As ADODB.Recordset)
    
    Dim l_strDeptNum As String
    Dim l_lngLocCd As Long
    Dim l_strSeqType As String
    Dim l_lngProjectId As Long
    
'Currently CrewNum is a non updatable field in TRIS 16-bit
    
    On Error GoTo eh
    
    l_strDeptNum = v_rsWorkGroupAttributes!DEPT_NUM
    l_lngLocCd = v_rsWorkGroupAttributes!LOC_CD
    l_lngProjectId = v_rsWorkGroupAttributes!PROJECT_ID
    If v_rsWorkGroupAttributes!RECORD_STATUS = "I" Then
        m_rsWorkGroupAttributes.AddNew
        m_rsWorkGroupAttributes!WG_NUM = m_lngWorkGroup
        m_rsWorkGroupAttributes!DEMAND_IND = v_rsWorkGroupAttributes!DEMAND_IND
        l_strSeqType = GetWGSequenceType(l_strDeptNum, l_lngLocCd, l_lngProjectId)
        m_rsWorkGroupAttributes!CREW_NUM = DetermineCrewNumber(m_lngWorkGroup, l_strSeqType, l_strDeptNum)
        m_rsWorkGroupAttributes!CREATED_EMP_NUM = CLng(m_strUserId)
'        Add this record to base schedule for cascade only
'        Base schedule recordset is not saved for demand workgroups
'        because the schedule is constant -9.
'        This is needed when a new demand workgroup is added and
'        employees are assigned.  busSchedule needs this state in order
'        to determine the employees workgroup schedule.
'        The data service handles the insert when adding the workgroup and
'        this simulates what the data service is doing.  If the data service
'        changes, this may need to change.
        If m_blnDemandWorkGroup Then
            m_rsWGBaseSched.AddNew
            m_rsWGBaseSched!WG_NUM = m_lngWorkGroup
            m_rsWGBaseSched!EFF_DATE = CDate(v_rsWorkGroupAttributes!EFF_START_DATE)
            m_rsWGBaseSched!ROT_CD = 0
            m_rsWGBaseSched!ROT_WEEK = 0
            m_rsWGBaseSched!SCHED_CD = -9
            m_rsWGBaseSched!CREATED_EMP_NUM = CLng(m_strUserId)
            m_rsWGBaseSched!LAST_MOD_EMP_NUM = CLng(m_strUserId)
            m_rsWGBaseSched!END_DATE = CDate("9999-12-31")
            m_rsWGBaseSched.Update
        End If
    End If
        
    m_rsWorkGroupAttributes!EFF_START_DATE = CDate(v_rsWorkGroupAttributes!EFF_START_DATE)
    If m_blnDemandWorkGroup Then
        m_rsWGBaseSched!EFF_DATE = CDate(v_rsWorkGroupAttributes!EFF_START_DATE)
    End If
    
    m_rsWorkGroupAttributes!WG_DESC = v_rsWorkGroupAttributes!WG_DESC
'    m_rsWorkGroupAttributes!CREW_NUM = v_rsWorkGroupAttributes!CREW_NUM
    m_rsWorkGroupAttributes!WG_SUPERVISOR = v_rsWorkGroupAttributes!WG_SUPERVISOR
    m_rsWorkGroupAttributes!PROJECT_ID = v_rsWorkGroupAttributes!PROJECT_ID
    m_rsWorkGroupAttributes!DEPT_NUM = v_rsWorkGroupAttributes!DEPT_NUM
    m_rsWorkGroupAttributes!LOC_CD = v_rsWorkGroupAttributes!LOC_CD
    If IsNull(v_rsWorkGroupAttributes!EFF_END_DATE) Or IsEmpty(v_rsWorkGroupAttributes!EFF_END_DATE) Then
        m_rsWorkGroupAttributes!EFF_END_DATE = m_datMaxEndDate
    Else
        m_rsWorkGroupAttributes!EFF_END_DATE = CDate(v_rsWorkGroupAttributes!EFF_END_DATE)
    End If
    m_rsWorkGroupAttributes!LAST_MOD_EMP_NUM = CLng(m_strUserId)
    
    Exit Sub
    
eh:
    
  TrsRaiseError m_strCMODULENAME, "SetWGAttributes"
    
End Sub

Private Function PostValidateWGSave( _
                    v_rsWorkGroupAttributes, _
                    v_rsWorkGroupSchedule, _
                    v_rsWorkGroupAssignments)
                    
    Dim l_blnValidSave As Boolean
    Dim l_datWGEndDate As Date
    Dim l_datOrigWGEndDate As Date
    Dim l_blnEndDateChanged As Boolean
    Dim l_datWGEffStartDate As Date
    
    On Error GoTo eh
    
    'The edits in the module are for the integrated save process
    'The user may update the end date of the work group as well
    ' as the end date of the assignments. So there are net effect edits
    
    'Please note that the edits in the comments may not be in
    ' this routine.  I was just typing what needs done
    
    'Following edits must be done
    ' 1. does a schedule exist for the life of a Not demand WG
    ' 2. Canwork group end on date (Done)
    ' 3. Check for employees primary after end of work group (Done)
    ' 4. Check for employees primary prior to start of work group (Done)
    
    'Following edits should be done in Rick's code
    ' 1. Can not change a wg line schedule if employee is assigned
    
    'Do we still need this edit
    ' 1. Can not change primary assignment of an employee in a
    '     demand work group
    
    'Cascade and Actual Time Checks can wait for next iteration
    
    l_datWGEffStartDate = GetEffectiveStartDate(m_rsWorkGroupAttributes)
    l_datWGEndDate = GetEffectiveEndDate(m_rsWorkGroupAttributes)
    l_datOrigWGEndDate = m_rsWorkGroupAttributes.Fields("EFF_END_DATE").OriginalValue
    If l_datWGEndDate = l_datOrigWGEndDate Then
        l_blnEndDateChanged = False
    Else
        l_blnEndDateChanged = True
    End If
    
    l_blnValidSave = True
    If l_datWGEndDate = m_datMaxEndDate Then
        l_blnValidSave = IsAssignmentDatesValid(v_rsWorkGroupAssignments)
    Else
        If l_blnEndDateChanged Then
            l_blnValidSave = CanWorkGroupEndOnDate(v_rsWorkGroupAttributes, m_lngWorkGroup, l_datWGEndDate)
        Else
            l_blnValidSave = IsAssignmentDatesValid(v_rsWorkGroupAssignments)
        End If
    End If
    
    If Not RecordsExist(m_rsWGBaseSched) And Not m_blnDemandWorkGroup Then
        l_blnValidSave = False
        AppendWGSchedError m_rsErrors, _
            trErrbusValidationError, mstrCBUSINESSERR, _
            "Base schedule does not exist."
    End If
    
    If CheckBaseSchedulesEffStartEnd(m_rsWGBaseSched, m_rsWorkGroupAttributes) Then
        l_blnValidSave = False
    End If
    
    If RecordsExist(m_rsWGDemandLineSchedules) Then
        m_rsWGDemandLineSchedules.Filter = "START_DATE < " & l_datWGEffStartDate
        If RecordsExist(m_rsWGDemandLineSchedules) Then
            l_blnValidSave = False
            m_rsWGDemandLineSchedules.MoveFirst
            While Not m_rsWGDemandLineSchedules.EOF
                AppendWGSchedLineError m_rsErrors, m_rsWGDemandLineSchedules, _
                    trErrbusValidationError, mstrCBUSINESSERR, _
                    "Line schedule start date(" & _
                    Format(m_rsWGDemandLineSchedules!START_DATE, "mm/dd/yyyy") & _
                    ") cannot be less than workgroup start date(" & _
                    Format(l_datWGEffStartDate, "mm/dd/yyyy") & ")"
                m_rsWGDemandLineSchedules.MoveNext
            Wend
        End If
    End If
    
    If Not m_rsWGDemandLineSchedules Is Nothing Then
        m_rsWGDemandLineSchedules.Filter = ""
    End If

    PostValidateWGSave = l_blnValidSave
    Exit Function
    
            
eh:
    
  TrsRaiseError m_strCMODULENAME, "PostValidateWGSave"

End Function


Private Function IsAssignmentDatesValid( _
    ByVal v_rsWorkGroupAssignments As ADODB.Recordset _
    ) As Boolean

    
    Dim l_blnAllValid As Boolean
    Dim l_datWGStart As Date
    Dim l_datWGEnd As Date
    Dim l_datAssignStart As Date
    Dim l_datAssignEnd As Date
    Dim l_lngEmpNum As Long
    Dim l_strErrorMessage As String
    Dim l_blnWriteError As Boolean

    On Error GoTo eh
    
    l_blnAllValid = True
    l_datWGStart = GetEffectiveStartDate(m_rsWorkGroupAttributes)
    l_datWGEnd = GetEffectiveEndDate(m_rsWorkGroupAttributes)
    If Not RecordsExist(m_rsWG_Assignments) Then
        IsAssignmentDatesValid = True
        Exit Function
    End If
        
    m_rsWG_Assignments.MoveFirst
    While Not m_rsWG_Assignments.EOF
        l_datAssignStart = GetEffectiveStartDate(m_rsWG_Assignments)
        l_datAssignEnd = GetEffectiveEndDate(m_rsWG_Assignments)
        l_lngEmpNum = m_rsWG_Assignments!EMP_NUM
        If l_datAssignStart < l_datWGStart Then
            l_blnAllValid = False
            SetFilterAssignmentDateError l_lngEmpNum, l_datAssignStart, v_rsWorkGroupAssignments
            l_strErrorMessage = "The employee (" & str(l_lngEmpNum) & ") can not be assigned (" & l_datAssignStart & ") prior to the start date of the work group (" & l_datWGStart & ")"
            If RecordsExist(v_rsWorkGroupAssignments) Then
                l_blnWriteError = AppendWGAssignmentsBRV(v_rsWorkGroupAssignments, 30000, mstrCBUSINESSERR, l_strErrorMessage)
            Else
                l_blnWriteError = AppendWGAssignmentsBRV(m_rsWG_Assignments, 30000, mstrCBUSINESSERR, l_strErrorMessage, False)
            End If
        End If
' The code below is not done in current system, but may be a needed requirement
' Most work group assignments are open end (Have Max end date)
' because it is easier to maintain for the user
' We may have a need to extend the work group effective date
' which in turn will extend all assignments that the effective
' end date is the same as the work group's
' Also most work groups are created open ended
        If l_datAssignEnd > l_datWGEnd Then
            l_blnAllValid = False
            SetFilterAssignmentDateError l_lngEmpNum, l_datAssignStart, v_rsWorkGroupAssignments
            l_strErrorMessage = "The employee can not be assigned (" & l_datAssignEnd & ") past to the end date of the work group (" & l_datWGEnd & ")"
            If RecordsExist(v_rsWorkGroupAssignments) Then
                l_blnWriteError = AppendWGAssignmentsBRV(v_rsWorkGroupAssignments, 30000, mstrCBUSINESSERR, l_strErrorMessage)
            Else
                l_blnWriteError = AppendWGAssignmentsBRV(m_rsWG_Assignments, 30000, mstrCBUSINESSERR, l_strErrorMessage, False)
            End If
        End If
        m_rsWG_Assignments.MoveNext
    Wend
    
    IsAssignmentDatesValid = l_blnAllValid
    Exit Function
    
eh:
    
  TrsRaiseError m_strCMODULENAME, "IsAssignmentDatesValid"

End Function


Private Sub SetFilterAssignmentDateError( _
    ByVal v_lngEmpNum As Long, _
    ByVal v_datAssignStart As Date, _
    ByVal v_rsWorkGroupAssignments As ADODB.Recordset)
    
    
    Dim l_strFilter As String
    Dim l_strRecStatus As String
    Dim l_strTimeStamp As String
    
    On Error GoTo eh
    
    If m_rsWG_Assignments.Status = adRecNew Then
        l_strRecStatus = "N"
        l_strFilter = "KEY_EMP_NUM = " & str(v_lngEmpNum) & _
                " AND KEY_EFF_START_DATE = " & v_datAssignStart
        l_strFilter = MDAC25FilterFix(l_strFilter)
    Else
        l_strTimeStamp = m_rsWG_Assignments!CREATED_TIMESTAMP
        l_strFilter = " EMP_NUM = " & str(v_lngEmpNum) & _
            " AND CREATED_TIMESTAMP = " & l_strTimeStamp
        l_strFilter = MDAC25FilterFix(l_strFilter)
    End If
    
    v_rsWorkGroupAssignments.Filter = l_strFilter
    Exit Sub

eh:
    
  TrsRaiseError m_strCMODULENAME, "SetFilterAssignmentDateError"

End Sub

Private Function SaveWGAttributes() As Boolean
    
    Dim l_obj As dbWorkGroup.cdWorkGroupTrans
    'Dim l_objDS As dbSecurity.cdBuiltDataScope
On Error GoTo eh

    m_rsWorkGroupAttributes.Filter = ""
    Set l_obj = CreateObject("dbWorkGroup.cdWorkGroupTrans")
    SaveWGAttributes = l_obj.SaveWGAttributes( _
            m_rsWorkGroupAttributes, m_strUserId)
    Set l_obj = Nothing
    
    '//-----invalidate WorkGroup DataScope containing the changing WorkGroup
'    Set l_objDS = CreateObject("dbSecurity.cdBuiltDataScope")
'    l_objDS.InvalidateDataScopeByRefWG m_lngWorkGroup, False
'    Set l_objDS = Nothing
    
    '//-----
    InvalidateDataScopeFromWGAttributes m_rsWorkGroupAttributes
    
    Exit Function
    
eh:
    GetObjectContext.SetAbort
  TrsRaiseError m_strCMODULENAME, "SaveWGAttributes"

End Function

Private Sub InvalidateDataScopeFromWGAttributes(ByVal v_rsWG_Attributes As ADODB.Recordset)
    Dim l_lngWgNum As Long
    '//Dim l_strDept1 As String, l_intLoc1 As Integer, l_strDiv1 As String, l_strFunc1 As String
    '//Dim l_strDept2 As String, l_intLoc2 As Integer, l_strDiv2 As String, l_strFunc2 As String
    Dim l_lngSupv1 As Long, l_lngSupv2 As Long
    Dim l_objDS As dbSecurity.cdBuiltDataScope
    '//Dim l_objCodes As dbCodeTables.cdCodeTables
    Dim l_colErrArgs As Collection
On Error GoTo eh
    v_rsWG_Attributes.Filter = adFilterNone
    v_rsWG_Attributes.Filter = adFilterPendingRecords
    
    If RecordsExist(v_rsWG_Attributes) Then
        '//Set l_objCodes = CreateObject("dbCodeTables.cdCodeTables")
        Set l_objDS = CreateObject("dbSecurity.cdBuiltDataScope")
        
        v_rsWG_Attributes.MoveFirst
        Do While Not v_rsWG_Attributes.EOF
            If v_rsWG_Attributes.Status = adRecDeleted Then
            ElseIf v_rsWG_Attributes.Status = adRecNew Then
                l_lngWgNum = v_rsWG_Attributes.Fields("WG_NUM").Value
                '//l_strDept1 = v_rsWG_Attributes.Fields("DEPT_NUM").Value
                '//l_intLoc1 = CLng(v_rsWG_Attributes.Fields("LOC_CD").Value)
                '//l_objCodes.GetDivisionAndFunctionAssociatedToDepartmentAndLocation l_strDept1, l_intLoc1, l_strDiv1, l_strFunc1
                l_lngSupv1 = CLng(v_rsWG_Attributes.Fields("WG_SUPERVISOR").Value)
                
                '//l_strDept2 = ""
                '//l_intLoc2 = 0
                '//l_strDiv2 = ""
                '//l_strFunc2 = ""
                l_lngSupv2 = 0
            ElseIf v_rsWG_Attributes.Status = adRecModified Then '//compare before and after
                l_lngWgNum = v_rsWG_Attributes.Fields("WG_NUM").OriginalValue
                '//l_strDept2 = v_rsWG_Attributes.Fields("DEPT_NUM").OriginalValue
                '//l_intLoc2 = CLng(v_rsWG_Attributes.Fields("LOC_CD").OriginalValue)
                '//l_objCodes.GetDivisionAndFunctionAssociatedToDepartmentAndLocation l_strDept2, l_intLoc2, l_strDiv2, l_strFunc2
                l_lngSupv2 = CLng(v_rsWG_Attributes.Fields("WG_SUPERVISOR").OriginalValue)

                '//l_strDept1 = v_rsWG_Attributes.Fields("DEPT_NUM").Value
                '//l_intLoc1 = CLng(v_rsWG_Attributes.Fields("LOC_CD").Value)
                '//l_objCodes.GetDivisionAndFunctionAssociatedToDepartmentAndLocation l_strDept1, l_intLoc1, l_strDiv1, l_strFunc1
                l_lngSupv1 = CLng(v_rsWG_Attributes.Fields("WG_SUPERVISOR").Value)
            End If
            
            '//-----invalidate DataScope containing the changing membership for Dept, Func, Div, Supervisor
'            If l_strDept1 <> l_strDept2 Then
'                l_objDS.InvalidateDataScopeByRefDept l_strDept1, True
'                If l_strDept2 <> "" Then
'                    l_objDS.InvalidateDataScopeByRefDept l_strDept2, True
'                End If
'            End If
'
'            If l_strFunc1 <> l_strFunc2 Then
'                l_objDS.InvalidateDataScopeByRefFun l_strFunc1, True
'                If l_strFunc2 <> "" Then
'                    l_objDS.InvalidateDataScopeByRefFun l_strFunc2, True
'                End If
'            End If
'
'            If l_strDiv1 <> l_strDiv2 Then
'                l_objDS.InvalidateDataScopeByRefDiv l_strDiv1, True
'                If l_strDiv2 <> "" Then
'                    l_objDS.InvalidateDataScopeByRefDiv l_strDiv2, True
'                End If
'            End If
            
            If l_lngSupv1 <> l_lngSupv2 Then
                l_objDS.InvalidateDataScopeWGByEffEmp l_lngSupv1, True
                If l_lngSupv2 <> 0 Then
                    l_objDS.InvalidateDataScopeWGByEffEmp l_lngSupv2, True
                End If
            End If
            
            v_rsWG_Attributes.MoveNext
        Loop
    End If
    
    v_rsWG_Attributes.Filter = adFilterNone
            
    Set l_objDS = Nothing
    '//Set l_objCodes = Nothing
    Exit Sub
eh:
    Set l_colErrArgs = New Collection
    l_colErrArgs.Add "l_lngWgNum = " & l_lngWgNum
'    l_colErrArgs.Add "l_strDept2 (from) = " & l_strDept2
'    l_colErrArgs.Add "l_intLoc2         = " & l_intLoc2
'    l_colErrArgs.Add "l_strDiv2         = " & l_strDiv2
'    l_colErrArgs.Add "l_strFunc2        = " & l_strFunc2
    l_colErrArgs.Add "l_lngSupv2 (from) = " & l_lngSupv2
'    l_colErrArgs.Add "l_strDept1   (to) = " & l_strDept1
'    l_colErrArgs.Add "l_intLoc1         = " & l_intLoc1
'    l_colErrArgs.Add "l_strDiv1         = " & l_strDiv1
'    l_colErrArgs.Add "l_strFunc1        = " & l_strFunc1
    l_colErrArgs.Add "l_lngSupv1 (to)   = " & l_lngSupv1
    
    Set l_objDS = Nothing
    '//Set l_objCodes = Nothing
    v_rsWG_Attributes.Filter = adFilterNone
    '//-----log any errors but allow the save to finish without errors
    TrsTraceError m_strCMODULENAME, "InvalidateDataScopeFromWGAttributes", , , , l_colErrArgs, True
End Sub

Private Function GetWorkGroupBaseSched( _
            ByVal v_lngWorkGroup As Long, _
            ByVal v_datDate As Date _
            ) As ADODB.Recordset

    Dim l_rs As ADODB.Recordset
    Dim l_obj As dbWorkGroup.cdWorkGroup

On Error GoTo eh

    Set l_obj = CreateObject("dbWorkGroup.cdWorkGroup")
    Set l_rs = l_obj.GetBaseSchedules(v_lngWorkGroup, v_datDate)
    Set GetWorkGroupBaseSched = l_rs
    Exit Function
    
eh:
    GetObjectContext.SetAbort
  TrsRaiseError m_strCMODULENAME, "GetWorkGroupBaseSched"

End Function

Private Function GetWorkGroupLineSched( _
            ByVal v_lngWorkGroup As Long, _
            ByVal v_datDate As Date _
            ) As ADODB.Recordset

    Dim l_rs As ADODB.Recordset
    Dim l_obj As dbWorkGroup.cdWorkGroup

On Error GoTo eh


    Set l_obj = New dbWorkGroup.cdWorkGroup
    Set l_rs = l_obj.GetDemandLineBaseSchedules(v_lngWorkGroup, v_datDate)
    Set GetWorkGroupLineSched = l_rs
    Exit Function

eh:
    GetObjectContext.SetAbort
  TrsRaiseError m_strCMODULENAME, "GetWorkGroupLineSched"

End Function



' RSG
' Added for Demand Line Schedule Maintenance
'
Private Function GetDemandLineSchedules( _
            ByVal v_lngWorkGroup As Long, _
            ByVal v_datMinEffDate As Date, _
            ByVal v_datMaxEffDate As Date, _
            ByVal v_varLineNums As Variant _
        ) As ADODB.Recordset
    Dim l_cdWorkgroup As dbWorkGroup.cdWorkGroup

On Error GoTo eh
    Set l_cdWorkgroup = CreateObject("dbWorkGroup.cdWorkGroup")
    Set GetDemandLineSchedules = _
        l_cdWorkgroup.GetDemandLineSchedules( _
            v_lngWorkGroup, v_datMinEffDate, v_datMaxEffDate, v_varLineNums)
    Set l_cdWorkgroup = Nothing
    Exit Function
    
eh:
    GetObjectContext.SetAbort
    Set l_cdWorkgroup = Nothing
    TrsRaiseError m_strCMODULENAME, "GetDemandLineSchedules"
End Function


Private Function GetDemandLineAssignments( _
            ByVal v_lngWgNum As Long, _
            ByVal v_datStartDate As Date, _
            ByVal v_datEndDate As Date _
            ) As ADODB.Recordset
    Dim l_cdSchedDemand As dbSchedDemand.cdSchedDemand
            
On Error GoTo eh
    Set l_cdSchedDemand = CreateObject("dbSchedDemand.cdSchedDemand")
    Set GetDemandLineAssignments = _
        l_cdSchedDemand.GetAssignments(v_lngWgNum, v_datStartDate, v_datEndDate)
    Set l_cdSchedDemand = Nothing
    Exit Function
    
eh:
    Set l_cdSchedDemand = Nothing
    TrsRaiseError m_strCMODULENAME, "GetDemandLineAssignments"
End Function


Private Function GetDemandLineAssignmentsInOtherWg( _
            ByVal v_rsWorkGroupAssignments As ADODB.Recordset _
            ) As ADODB.Recordset
    Dim l_cdSchedDemand As dbSchedDemand.cdSchedDemand
            
On Error GoTo eh
    Set l_cdSchedDemand = CreateObject("dbSchedDemand.cdSchedDemand")
    Set GetDemandLineAssignmentsInOtherWg = _
        l_cdSchedDemand.GetAssignmentsInOtherWG(v_rsWorkGroupAssignments)
    Set l_cdSchedDemand = Nothing
    Exit Function
    
eh:
    Set l_cdSchedDemand = Nothing
    TrsRaiseError m_strCMODULENAME, "GetDemandLineAssignmentsInOtherWg"
End Function

'//-----4/1/2002 (no april fool)
Private Function CheckForDemandInOtherWg(ByVal v_rsWorkGroupAssignments As ADODB.Recordset) As Boolean
    Dim l_rsDemandLineAssignmentsInOtherWg As ADODB.Recordset
'//    Dim l_cbsSchedDemand As busSchedDemand.cbsSchedDemandTrans
'    Dim l_rsDemandLineAssignmentsToEnd As ADODB.Recordset
    Dim l_blnWriteError As Boolean, l_blnOK As Boolean
    Dim sFilter As String, sWarning As String
    Dim l_rsErrors As ADODB.Recordset
    
On Error GoTo eh

    l_blnOK = True
    sFilter = "RECORD_STATUS = 'I' AND PRIMARY_IND = 'Y'"
    v_rsWorkGroupAssignments.Filter = MDAC25FilterFix(sFilter)
    
    If RecordsExist(v_rsWorkGroupAssignments) Then
        Set l_rsDemandLineAssignmentsInOtherWg = GetDemandLineAssignmentsInOtherWg(v_rsWorkGroupAssignments)
    
        If RecordsExist(l_rsDemandLineAssignmentsInOtherWg) Then
'            If Not m_blnDemandWorkGroup Then
'                Set l_rsDemandLineAssignmentsToEnd = CreateLineAssignmentsSaveRS
'            End If
    
            With l_rsDemandLineAssignmentsInOtherWg
                .MoveFirst
    
                Do While Not .EOF
                    '//-----for each warning, locate the request row (could probably use just EmpNum, RecStat, PrimInd)-----//
                    sFilter = "EMP_NUM = " & !EMP_NUM & _
                            " AND EFF_START_DATE <= " & !END_DATE & _
                            " AND RECORD_STATUS = 'I' AND PRIMARY_IND = 'Y'"
                    v_rsWorkGroupAssignments.Filter = MDAC25FilterFix(sFilter)
                    If RecordsExist(v_rsWorkGroupAssignments) Then
                        '//-----the eff_end_date could be null so it has been removed from the filter and done manually here-----//
                        '//" AND EFF_END_DATE >= " & !START_DATE & _'
                        v_rsWorkGroupAssignments.MoveFirst
                        Do While Not v_rsWorkGroupAssignments.EOF
                            '//-----if eff_end_date is null then treat it like 12/31/9999, which will be >= the start date-----//
                            If IsNull(v_rsWorkGroupAssignments!EFF_END_DATE) Then
                                Exit Do
                            End If
                            
                            If v_rsWorkGroupAssignments!EFF_END_DATE >= !START_DATE Then
                                Exit Do
                            End If
                            v_rsWorkGroupAssignments.MoveNext
                        Loop
                    Else '//-----could not match brv warning to request row???-----//
                        sFilter = "EMP_NUM = " & !EMP_NUM
                        v_rsWorkGroupAssignments.Filter = MDAC25FilterFix(sFilter)
                        If Not RecordsExist(v_rsWorkGroupAssignments) Then
                            v_rsWorkGroupAssignments.Filter = adFilterNone
                            v_rsWorkGroupAssignments.MoveFirst
                        End If
                    End If
                    If m_blnDemandWorkGroup Then
                        sWarning = "Employee " & !EMP_NUM & " is currently assigned to demand WG " & !WG_NUM & " line " & !LINE_NUM & _
                            " between " & !START_DATE & " and " & !END_DATE
                        l_blnWriteError = AppendWGAssignmentsBRV(v_rsWorkGroupAssignments, 30000, mstrCBUSINESSWAR, sWarning)
                    Else
                        sWarning = "Employee " & !EMP_NUM & " is still assigned to demand line schedule for WG " & !WG_NUM & " line " & !LINE_NUM & _
                            " between " & !START_DATE & " and " & !END_DATE & ". Assignment must be ended before new WG schedule will take effect."
                        l_blnWriteError = AppendWGAssignmentsBRV(v_rsWorkGroupAssignments, 30000, mstrCBUSINESSWAR, sWarning)
                    End If
                    .MoveNext
                Loop
            End With
        End If
    End If
    
' 5/21/02 RSG - Commented all code related to SaveAssignments call to unassign
' an employee from a demand line when moved from a demand wg to a nondemand wg.
' This was decided because the existing unassign request, for ANY employee assigned
' to the line would have unassigned half of a relief assignment, while not completely
' reverting the relief.  Since this workgroup assignment request may be done by
' the nondemand workgroup supervisor, which would not be able to clean up the dangling
' relief, it was decided to leave the unassignment and any relief cleanup to the
' demand supervisor/scheduler.  It will force the issue to be handled since the
' nondemand workgroup schedule will be ignored until the demand line assignment
' is removed.
'
'                    '//-----if an employee is on a demand line for a work group and
'                    '//-----the user moves that employee from a demand work group to another demand work group,
'                    '//-----return a BRV warning indicating the work group and line number that the employee is currently on
'                    If m_blnDemandWorkGroup Then
'                        sWarning = "Employee " & !EMP_NUM & " is currently assigned to demand WG " & !WG_NUM & " line " & !LINE_NUM & _
'                            " between " & !START_DATE & " and " & !END_DATE
'                        l_blnWriteError = AppendWGAssignmentsBRV(v_rsWorkGroupAssignments, 30000, mstrCBUSINESSWAR, sWarning)
'
'                    '//-----if an employee is on a demand line for a work group and
'                    '//-----the user moves that employee from a demand work group to a non-demand work group,
'                    '//-----end the employee's assignment in that work group for the dates indicated and
'                    '//-----return a BRV warning indicating the work group, line number and dates(s) that the employee's assignment ended
'                    Else
'                        '//-----build request for unassignmenting demand line assignment
'                        l_rsDemandLineAssignmentsToEnd.AddNew
'                        l_rsDemandLineAssignmentsToEnd!WG_NUM = !WG_NUM
'                        l_rsDemandLineAssignmentsToEnd!LINE_NUM = !LINE_NUM
'
'                        '//-----don't need to unassign everything, just the intersection, Max(StartDate) to Min(EndDate)
'                        If !START_DATE >= v_rsWorkGroupAssignments!EFF_START_DATE Then
'                            l_rsDemandLineAssignmentsToEnd!START_DATE = !START_DATE
'                        Else
'                            l_rsDemandLineAssignmentsToEnd!START_DATE = v_rsWorkGroupAssignments!EFF_START_DATE
'                        End If
'
'                        If IsNull(v_rsWorkGroupAssignments!EFF_END_DATE) Then
'                            l_rsDemandLineAssignmentsToEnd!END_DATE = !END_DATE
'                        ElseIf !END_DATE <= v_rsWorkGroupAssignments!EFF_END_DATE Then
'                            l_rsDemandLineAssignmentsToEnd!END_DATE = !END_DATE
'                        Else
'                            l_rsDemandLineAssignmentsToEnd!END_DATE = v_rsWorkGroupAssignments!EFF_END_DATE
'                        End If
'
'                        l_rsDemandLineAssignmentsToEnd!EMP_NUM = 0
'                        l_rsDemandLineAssignmentsToEnd!RECORD_STATUS = "I"
'                        l_rsDemandLineAssignmentsToEnd.Update
'
'                        sWarning = "Employee " & !EMP_NUM & " assigned to demand WG " & !WG_NUM & " line " & !LINE_NUM & _
'                            " between " & l_rsDemandLineAssignmentsToEnd!START_DATE & " and " & l_rsDemandLineAssignmentsToEnd!END_DATE & " has been ended."
'                        l_blnWriteError = AppendWGAssignmentsBRV(v_rsWorkGroupAssignments, 30000, mstrCBUSINESSWAR, sWarning)
'
'                    End If
'
'                    .MoveNext
'                Loop
'
' 5/21/02 RSG - Commented all code related to SaveAssignments call to unassign
' an employee from a demand line when moved from a demand wg to a nondemand wg.
'
'                If Not m_blnDemandWorkGroup Then
'                    If RecordsExist(l_rsDemandLineAssignmentsToEnd) Then
'                        Set l_cbsSchedDemand = CreateObject("busSchedDemand.cbsSchedDemandTrans")
'                        If Not l_cbsSchedDemand.SaveAssignments(l_rsDemandLineAssignmentsToEnd, l_rsErrors, m_strUserId) Then
'                            '//-----add any errors from demand scheduling to BRV's-----//
'                            l_rsErrors.Filter = MDAC25FilterFix("ERR_TYPE = 'BE'")
'                            AppendRecordset l_rsErrors, m_rsErrors
'                            m_rsErrors.Filter = adFilterNone
'                            l_blnOK = False
'                        End If
'                        Set l_cbsSchedDemand = Nothing
'                    End If
'                End If
'            End With
'        End If
'    End If
    
    v_rsWorkGroupAssignments.Filter = adFilterNone
    
    CheckForDemandInOtherWg = l_blnOK
    If l_blnOK Then
        GetObjectContext.SetComplete
    Else
        GetObjectContext.SetAbort
    End If
    
    Set l_rsDemandLineAssignmentsInOtherWg = Nothing
'//    Set l_cbsSchedDemand = Nothing
'    Set l_rsDemandLineAssignmentsToEnd = Nothing
    Set l_rsErrors = Nothing
    
    Exit Function
eh:
    Set l_rsDemandLineAssignmentsInOtherWg = Nothing
'//    Set l_cbsSchedDemand = Nothing
'    Set l_rsDemandLineAssignmentsToEnd = Nothing
    Set l_rsErrors = Nothing

    TrsRaiseError m_strCMODULENAME, "CheckForDemandInOtherWg"
End Function
    
'//-----copied from psoScheduling.cpScheduling
'DEMAND LINE ASSIGNMENTS
'
' 5/24/02 RSG Interred and changed to Private in case resurrected.
'
'Private Function CreateLineAssignmentsSaveRS() As ADODB.Recordset
'On Error GoTo CreateLineAssignmentsSaveRS_EH
'
'Const l_lngCFIELDATTR = adFldUpdatable Or adFldIsNullable
'Dim l_adrsSaveLineAssignment As ADODB.Recordset
'
'     'Create the empty recordset
'    Set l_adrsSaveLineAssignment = New ADODB.Recordset
'
'    With l_adrsSaveLineAssignment
'        .Fields.Append "WG_NUM", adInteger, , l_lngCFIELDATTR
'        .Fields.Append "EMP_NUM", adInteger, , l_lngCFIELDATTR
'        .Fields.Append "LINE_NUM", adInteger, , l_lngCFIELDATTR
'        .Fields.Append "START_DATE", adDate, , l_lngCFIELDATTR
'        .Fields.Append "END_DATE", adDate, , l_lngCFIELDATTR
'        .Fields.Append "RELIEF_COMP", adChar, 3, l_lngCFIELDATTR
'
'        'Key Fields
'        .Fields.Append "KEY_WG_NUM", adInteger, , l_lngCFIELDATTR
'        .Fields.Append "KEY_EMP_NUM", adInteger, , l_lngCFIELDATTR
'        .Fields.Append "KEY_LINE_NUM", adInteger, , l_lngCFIELDATTR
'        .Fields.Append "KEY_START_DATE", adDate, , l_lngCFIELDATTR
'        .Fields.Append "KEY_END_DATE", adDate, , l_lngCFIELDATTR
'
'    '//-----because an Update ("C") request changes to a delete/insert,
'    '//----- need to carry the created columns over-----//
'    '//-----also needed for conversion of 21st Watch which will preserve the last_mod columns-----//
'        .Fields.Append "CREATED_EMP_NUM", adInteger, , l_lngCFIELDATTR
'        .Fields.Append "LAST_MOD_EMP_NUM", adInteger, , l_lngCFIELDATTR
'        .Fields.Append "CREATED_TIMESTAMP", adChar, 26, l_lngCFIELDATTR
'        .Fields.Append "LAST_MOD_TIMESTAMP", adChar, 26, l_lngCFIELDATTR
'
'        .Fields.Append "RECORD_STATUS", adChar, 1, adFldUpdatable
'
'        .CursorLocation = adUseClient
'        .CursorType = adOpenStatic
'        .LockType = adLockBatchOptimistic
'
'        .Open
'    End With
'
'    Set CreateLineAssignmentsSaveRS = l_adrsSaveLineAssignment
'
'    'Destroy Object
'    Set l_adrsSaveLineAssignment = Nothing
'
'    Exit Function
'
'CreateLineAssignmentsSaveRS_EH:
'    Set l_adrsSaveLineAssignment = Nothing
'    Set CreateLineAssignmentsSaveRS = Nothing
'
'    TrsRaiseError m_strCMODULENAME, "CreateLineAssignmentsSaveRS"
'End Function


Private Sub DetermineDemandSchedScope( _
                ByRef v_rsLineScheduleRequest As ADODB.Recordset, _
                ByRef l_datMinEffDate As Date, _
                ByRef l_datMaxEffDate As Date, _
                ByRef l_varLineNums As Variant)
    Dim l_datMinDate As Date
    Dim l_datMaxDate As Date
    Dim l_colLines As cuSet
    'move this to date/time module
    Const CHIGHDATE As Date = #12/31/2079#
    
' RSG 12/20/01  - Ask Jim about this
' There may be an issue with determining the maxEffDate when the request
' has Null endDate.  May need the original endDate or just use 12/31/9999?
' This could be a large recordset depending on the number of lines, startdate.
'
On Error GoTo eh

    If Not RecordsExist(v_rsLineScheduleRequest) Then
        Exit Sub
    Else
        v_rsLineScheduleRequest.MoveFirst
    End If

    Set l_colLines = New cuSet
    l_datMinDate = CHIGHDATE
    Do Until v_rsLineScheduleRequest.EOF
' RSG - 12/20/01 Client is not providing values for Delete request.
' Should be able to remove this next week.
        If v_rsLineScheduleRequest!RECORD_STATUS = "D" Then
            v_rsLineScheduleRequest!WG_NUM = v_rsLineScheduleRequest!KEY_WG_NUM
            v_rsLineScheduleRequest!LINE_NUM = v_rsLineScheduleRequest!KEY_LINE_NUM
            v_rsLineScheduleRequest!START_DATE = CDate(v_rsLineScheduleRequest!KEY_START_DATE)
        End If
    
        If v_rsLineScheduleRequest!START_DATE < l_datMinDate Then
            l_datMinDate = v_rsLineScheduleRequest!START_DATE
        End If
        If v_rsLineScheduleRequest!KEY_START_DATE < l_datMinDate Then
            l_datMinDate = v_rsLineScheduleRequest!KEY_START_DATE
        End If
        If Not IsNull(v_rsLineScheduleRequest!END_DATE) Then
            If v_rsLineScheduleRequest!END_DATE > l_datMaxDate Then
                l_datMaxDate = v_rsLineScheduleRequest!END_DATE
            End If
        End If
        l_colLines.Add v_rsLineScheduleRequest!LINE_NUM.Value
        v_rsLineScheduleRequest.MoveNext
    Loop

    l_varLineNums = ArrayFromCollection(l_colLines.Contents)
    l_datMinEffDate = l_datMinDate
    l_datMaxEffDate = #12/31/9999#
    
    v_rsLineScheduleRequest.MoveFirst
    Set l_colLines = Nothing
    Exit Sub
    
eh:
    Set l_colLines = Nothing
    TrsRaiseError m_strCMODULENAME, "DetermineDemandSchedScope"
End Sub


Private Function DetermineGetAssignmentMaxEnd( _
                ByRef v_rsLineSchedules As ADODB.Recordset) As Date
    Dim l_datMaxDate As Date
    
On Error GoTo eh

    If Not RecordsExist(v_rsLineSchedules) Then
        DetermineGetAssignmentMaxEnd = m_datMaxEndDate
        Exit Function
    Else
        v_rsLineSchedules.MoveFirst
    End If
    
    l_datMaxDate = 0
    Do Until v_rsLineSchedules.EOF
        If v_rsLineSchedules!END_DATE > l_datMaxDate Then
            l_datMaxDate = v_rsLineSchedules!END_DATE
        End If
        v_rsLineSchedules.MoveNext
    Loop
    
    DetermineGetAssignmentMaxEnd = l_datMaxDate
    v_rsLineSchedules.MoveFirst
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "DetermineGetAssignmentMaxEnd"
End Function


Private Function SaveWGDemandSchedules() As Boolean
    Dim l_obj As dbWorkGroup.cdWorkGroupTrans

On Error GoTo eh
'    CheckForJobChanges
    Set l_obj = CreateObject("dbWorkGroup.cdWorkGroupTrans")
    SaveWGDemandSchedules = l_obj.SaveWGLineSchedules(m_rsWGDemandLineSchedules)
    Set l_obj = Nothing
    Exit Function
    
eh:
    GetObjectContext.SetAbort
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "SaveWGDemandSchedules"
End Function


Private Function SaveWGBaseSchedules() As Boolean
    Dim l_obj As dbWorkGroup.cdWorkGroupTrans

On Error GoTo eh
    Set l_obj = CreateObject("dbWorkGroup.cdWorkGroupTrans")
    SaveWGBaseSchedules = l_obj.SaveWGBaseSchedules(m_rsWGBaseSched)
    Set l_obj = Nothing
    Exit Function
    
eh:
    GetObjectContext.SetAbort
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "SaveWGBaseSchedules"
End Function

'
' Start ProcessWorkgroupDemandSchedules
'
Private Function ProcessWorkgroupDemandSchedules( _
            ByRef r_rsScheduleRequests As ADODB.Recordset _
        ) As Boolean
' 1. prevalidate input
' 2. retrieve db data;
'    existing wg_line_sched rows - Jim will have these in m_rsWGDemandLineSchedules
'    existing wg_base_sched rows - Jim will have these in m_rsWGBaseSched (NotDemand)
' 3. process each request
' 4. post validate
' 5. clean up module level variables.
' 6. save results - Jim will pass updated recordset to save with other updates
'
' 10/17/02 RSG Process change requests before inserts to ensure persistent line
' schedule is found before insert processing changes or deletes it. e.g. when
' an insert causes a split.

    Dim l_blnReturn As Boolean

On Error GoTo eh

    If RecordsExist(r_rsScheduleRequests) Then
        'Continue
    Else
        ProcessWorkgroupDemandSchedules = True
        Exit Function
    End If
    
    l_blnReturn = PreValidateDemandScheduleRequests(r_rsScheduleRequests)
    If Not l_blnReturn Then
        ProcessWorkgroupDemandSchedules = False
        Exit Function
    Else
        r_rsScheduleRequests.MoveFirst
    End If

    ProcessAllDemandSchedDeleteRequests r_rsScheduleRequests
    ProcessAllDemandSchedChangeRequests r_rsScheduleRequests
    ProcessAllDemandSchedInsertRequests r_rsScheduleRequests

    ProcessWorkgroupDemandSchedules = _
        PostValidateScheduleUpdates(m_rsWGDemandLineSchedules, r_rsScheduleRequests)
        
'    ClearModuleVariables
    
' Cascade preparation
'    ConsolidateCascadeIntervals
    
    Exit Function
    
eh:
'    ClearModuleVariables
    TrsRaiseError m_strCMODULENAME, "ProcessWorkgroupDemandSchedules"
End Function


Private Sub ProcessAllDemandSchedDeleteRequests( _
                ByRef r_rsScheduleRequests As ADODB.Recordset)
    Dim l_strWGNum As String
    Dim l_strLineNum As String
    Dim l_datEffDate As Date
    Dim l_strCreatedTS As String
    Dim l_colErrArgs As Collection

'
' If all these are same, combine into 1 and pass RECORD_STATUS to process
'

On Error GoTo eh

    r_rsScheduleRequests.Filter = "RECORD_STATUS = 'D'"
    Do Until r_rsScheduleRequests.EOF
        l_strWGNum = r_rsScheduleRequests!KEY_WG_NUM
        l_strLineNum = r_rsScheduleRequests!KEY_LINE_NUM
        l_datEffDate = r_rsScheduleRequests!KEY_START_DATE
        l_strCreatedTS = r_rsScheduleRequests!CREATED_TIMESTAMP
        DeleteLineSchedule r_rsScheduleRequests, m_rsWGDemandLineSchedules
        r_rsScheduleRequests.MoveNext
    Loop

    r_rsScheduleRequests.Filter = adFilterNone
    Exit Sub
    
eh:
    Set l_colErrArgs = New Collection
    l_colErrArgs.Add "KEY_WG_NUM = " & l_strWGNum
    l_colErrArgs.Add "KEY_LINE_NUM = " & l_strLineNum
    l_colErrArgs.Add "KEY_START_DATE = " & Format(l_datEffDate, "mm/dd/yyyy")
    l_colErrArgs.Add "CREATED_TIMESTAMP = " & l_strCreatedTS
    TrsRaiseError m_strCMODULENAME, "ProcessAllDemandSchedDeleteRequests", , , , l_colErrArgs, True
End Sub


Private Sub ProcessAllDemandSchedInsertRequests( _
                ByRef r_rsScheduleRequests As ADODB.Recordset)
    Dim l_strWGNum As String
    Dim l_strLineNum As String
    Dim l_datEffDate As Date
    Dim l_colErrArgs As Collection
    Dim l_strRecStatus As String

On Error GoTo eh

    r_rsScheduleRequests.Filter = "RECORD_STATUS = 'I'"
    Do Until r_rsScheduleRequests.EOF
        l_strWGNum = r_rsScheduleRequests!KEY_WG_NUM
        l_strLineNum = r_rsScheduleRequests!KEY_LINE_NUM
        l_datEffDate = r_rsScheduleRequests!KEY_START_DATE
        l_strRecStatus = Trim(UCase(r_rsScheduleRequests!RECORD_STATUS))
        InsertLineSchedule r_rsScheduleRequests, m_rsWGDemandLineSchedules
        r_rsScheduleRequests.MoveNext
    Loop

    r_rsScheduleRequests.Filter = adFilterNone
    Exit Sub
    
eh:
    Set l_colErrArgs = New Collection
    l_colErrArgs.Add "KEY_WG_NUM = " & l_strWGNum
    l_colErrArgs.Add "KEY_LINE_NUM = " & l_strLineNum
    l_colErrArgs.Add "KEY_START_DATE = " & Format(l_datEffDate, "mm/dd/yyyy")
    TrsRaiseError m_strCMODULENAME, "ProcessAllDemandSchedInsertRequests", , , , l_colErrArgs, True
End Sub


Private Sub ProcessAllDemandSchedChangeRequests( _
                ByRef r_rsScheduleRequests As ADODB.Recordset)
    Dim l_strWGNum As String
    Dim l_strLineNum As String
    Dim l_datEffDate As Date
    Dim l_strCreatedTS As String
    Dim l_strRecStatus As String
    Dim l_colErrArgs As Collection

On Error GoTo eh

    r_rsScheduleRequests.Filter = "RECORD_STATUS = 'C'"
    Do Until r_rsScheduleRequests.EOF
        l_strWGNum = r_rsScheduleRequests!KEY_WG_NUM
        l_strLineNum = r_rsScheduleRequests!KEY_LINE_NUM
        l_datEffDate = r_rsScheduleRequests!KEY_START_DATE
        l_strCreatedTS = r_rsScheduleRequests!CREATED_TIMESTAMP
        l_strRecStatus = Trim(UCase(r_rsScheduleRequests!RECORD_STATUS))
        UpdateLineSchedule r_rsScheduleRequests, m_rsWGDemandLineSchedules
        r_rsScheduleRequests.MoveNext
    Loop

    r_rsScheduleRequests.Filter = adFilterNone
    Exit Sub
    
eh:
    Set l_colErrArgs = New Collection
    l_colErrArgs.Add "KEY_WG_NUM = " & l_strWGNum
    l_colErrArgs.Add "KEY_LINE_NUM = " & l_strLineNum
    l_colErrArgs.Add "KEY_START_DATE = " & Format(l_datEffDate, "mm/dd/yyyy")
    l_colErrArgs.Add "CREATED_TIMESTAMP = " & l_strCreatedTS
    TrsRaiseError m_strCMODULENAME, "ProcessAllDemandSchedChangeRequests", , , , l_colErrArgs, True
End Sub


Private Sub DeleteLineSchedule( _
            ByRef r_rsScheduleRequests As ADODB.Recordset, _
            ByRef r_rsLineSchedules As ADODB.Recordset)
    Dim l_lngWgNum As Long
    Dim l_lngLineNum As Long
    Dim l_datStartDate As Date
    Dim l_strCreatedTS As String

On Error GoTo eh
    l_lngWgNum = r_rsScheduleRequests!KEY_WG_NUM
    l_lngLineNum = r_rsScheduleRequests!KEY_LINE_NUM
    l_datStartDate = r_rsScheduleRequests!KEY_START_DATE
    l_strCreatedTS = r_rsScheduleRequests!CREATED_TIMESTAMP
    
    FindLineScheduleRecord _
        l_lngWgNum, l_lngLineNum, l_datStartDate, l_strCreatedTS, r_rsLineSchedules
    If r_rsLineSchedules.EOF Then
        Err.Raise trErrbusInvalidDeleteEntryNotFound
    End If
    
'   Only allowed to delete a line schedule when an employee is not
'   assigned to the line for the dates specified. TTR_LINE_ASGNM OR TTR_LINE_RELIEF
' 8/2/02 RSG - Moved this edit to post validate.
'    If AnyLineAssignments( _
'            r_rsLineSchedules!WG_NUM, r_rsLineSchedules!LINE_NUM, _
'                r_rsLineSchedules!START_DATE, r_rsLineSchedules!END_DATE) Then
'        Err.Raise trErrbusWGLineAssignExist
'    End If
    
    AppendWGLineCascadeDates "D", r_rsLineSchedules!WG_NUM, r_rsLineSchedules!LINE_NUM, _
                r_rsLineSchedules!START_DATE, r_rsLineSchedules!END_DATE
    r_rsLineSchedules.Delete
    Exit Sub

eh:
    Select Case Err.Number
        Case trErrbusWGLineAssignExist
            Err.Source = FormatSchedBRV(TrType_WGM_SCHED, l_lngWgNum, l_lngLineNum, l_datStartDate)
            TrsRaiseError m_strCMODULENAME, "DeleteLineSchedule", Err.Number
        Case Else
            TrsRaiseError m_strCMODULENAME, "DeleteLineSchedule"
    End Select
End Sub


Private Sub InsertLineSchedule( _
            ByRef r_rsScheduleRequests As ADODB.Recordset, _
            ByRef r_rsLineSchedules As ADODB.Recordset)
' 7/22/02 RSG
' Adjust ROT_WEEK when implicit changes to START_DATE and
' implicit insert resulting from ending part of split interval.
' Implicit START_DATE changes result when previous interval END_DATE
' is extended.
' Explicit changes, i.e. user entered, to START_DATE should not affect
' the ROT_WEEK.  Simply accept what the user entered or retain the previous
' value.
    Dim l_lngWgNum As Long
    Dim l_lngLineNum As Long
    Dim l_datStartDate As Date
    Dim l_datEndDate As Date
    Dim l_strCreatedTS As String
    Dim l_blnEndDateProvided As Boolean
    Dim l_datPreviousEndDate As Date
    Dim l_blnSplitPrevious As Boolean
    
    Dim l_lngSplitWgNum As Long
    Dim l_lngSplitLineNum As Long
    Dim l_lngSplitStartDate As Date
    Dim l_lngSplitEndDate As Date
    Dim l_lngSplitRotCd As Long
    Dim l_lngSplitRotWeek As Long
    Dim l_lngSplitSchedCd As Long
    Dim l_strSplitJobNum As String
    Dim l_strSplitReqLineInd As String
    Dim l_lngSplitCreatedEmpNum As Long
    Dim l_strSplitCreatedTS As String
    
    Dim l_datOrigStartDate As Date
    Dim l_lngOrigRotWeek As Long

On Error GoTo eh
    l_datEndDate = 0
    l_datPreviousEndDate = 0
    r_rsLineSchedules.Sort = "WG_NUM ASC, LINE_NUM ASC, START_DATE ASC"
    l_lngWgNum = r_rsScheduleRequests!WG_NUM
    l_lngLineNum = r_rsScheduleRequests!LINE_NUM
    l_datStartDate = r_rsScheduleRequests!START_DATE
'    l_strCreatedTS = r_rsScheduleRequests!CREATED_TIMESTAMP
'    FindLineScheduleRecord l_lngWgNum, l_lngLineNum, l_datStartDate, r_rsLineSchedules
       
' Attempt to detect duplicate now.  I am not sure possible since other updates may
' have the same start date now.
' Fix the find to include timestamp, currently finding a different record with same startDate.
'
'    If Not r_rsLineSchedules.EOF Then
'        ' app error - duplicate insert - brv
'        Err.Raise vbObjectError + 10102
'    Else
'        r_rsLineSchedules.MoveFirst
'    End If

' If previous effective line sched record exists,
' set end date to day prior to this start date.
' If the new START_DATE is the same as existing effective, just delete it.
    FindEffectiveLineScheduleRecord l_lngWgNum, l_lngLineNum, l_datStartDate, r_rsLineSchedules
    If Not r_rsLineSchedules.EOF Then
        l_datOrigStartDate = r_rsLineSchedules!START_DATE
        l_datPreviousEndDate = r_rsLineSchedules!END_DATE
        l_datEndDate = l_datPreviousEndDate
'    Only allowed to change JOB_NUM OR REQ_LINE_IND when an employee is not
'    assigned to the line for the dates specified.
'    If there is an existing effective line schedule, determine the End date before
'    the AnyLineAssignments call.
        l_datEndDate = DetermineNewEndDate(l_datEndDate, r_rsScheduleRequests)
        If Trim(r_rsLineSchedules!REQ_LINE_IND) <> Trim(r_rsScheduleRequests!REQ_LINE_IND) Then
            If AnyLineAssignments( _
                r_rsScheduleRequests!WG_NUM, r_rsScheduleRequests!LINE_NUM, _
                        l_datStartDate, l_datEndDate) Then
                Err.Raise trErrbusWGLineAssignExist
            End If
        End If
        If l_datStartDate > l_datOrigStartDate Then
            If l_datEndDate < l_datPreviousEndDate Then
                l_blnSplitPrevious = True
                l_lngSplitWgNum = r_rsLineSchedules!WG_NUM
                l_lngSplitLineNum = r_rsLineSchedules!LINE_NUM
                l_lngSplitStartDate = l_datEndDate + 1
                l_lngSplitEndDate = l_datPreviousEndDate
                l_lngSplitRotCd = r_rsLineSchedules!ROT_CD
                l_lngSplitRotWeek = r_rsLineSchedules!ROT_WEEK
                l_lngSplitSchedCd = r_rsLineSchedules!SCHED_CD
                l_strSplitJobNum = r_rsLineSchedules!JOB_NUM
                l_strSplitReqLineInd = r_rsLineSchedules!REQ_LINE_IND
                l_lngSplitCreatedEmpNum = r_rsLineSchedules!CREATED_EMP_NUM
                l_strSplitCreatedTS = r_rsLineSchedules!CREATED_TIMESTAMP
            End If
            r_rsLineSchedules!END_DATE = l_datStartDate - 1
            r_rsLineSchedules!LAST_MOD_EMP_NUM = m_lngUserNum
            r_rsLineSchedules.Update
        Else                            'l_datStartDate = l_datOrigStartDate
            If l_datEndDate >= l_datPreviousEndDate Then
                r_rsLineSchedules.Delete
            Else
                r_rsLineSchedules!START_DATE = l_datEndDate + 1
                r_rsLineSchedules!LAST_MOD_EMP_NUM = m_lngUserNum
                If r_rsLineSchedules!ROT_CD <> 0 Then
                    l_lngOrigRotWeek = r_rsLineSchedules!ROT_WEEK
                    r_rsLineSchedules!ROT_WEEK = _
                        DetermineNewRotationWeek( _
                            r_rsLineSchedules!START_DATE, _
                            r_rsLineSchedules!ROT_CD, _
                            l_datOrigStartDate, _
                            l_lngOrigRotWeek)
                    If r_rsLineSchedules!ROT_WEEK <> l_lngOrigRotWeek Then
                        AppendWarningRotationWeekChange _
                            trWarRotWeekAdjusted, _
                            l_lngOrigRotWeek, _
                            r_rsLineSchedules.Fields("ROT_WEEK").Value, _
                            r_rsLineSchedules.Fields("WG_NUM").Value, _
                            r_rsLineSchedules.Fields("LINE_NUM").Value, _
                            r_rsLineSchedules.Fields("START_DATE").Value
                    End If
                End If
                r_rsLineSchedules.Update
            End If
        End If
    Else
'    If there is no existing effective line schedule, determine the End date
        l_datEndDate = DetermineNewEndDate(l_datEndDate, r_rsScheduleRequests)
    End If
    
' Add the new line schedule.
    r_rsLineSchedules.AddNew
    r_rsLineSchedules!WG_NUM = r_rsScheduleRequests!WG_NUM
    r_rsLineSchedules!LINE_NUM = r_rsScheduleRequests!LINE_NUM
    r_rsLineSchedules!START_DATE = l_datStartDate
    r_rsLineSchedules!END_DATE = l_datEndDate
    
    If IsMissingValue(r_rsScheduleRequests!ROT_CD) Then
        r_rsLineSchedules!ROT_CD = 0
        r_rsLineSchedules!ROT_WEEK = 0
    Else
        r_rsLineSchedules!ROT_CD = r_rsScheduleRequests!ROT_CD
        r_rsLineSchedules!ROT_WEEK = r_rsScheduleRequests!ROT_WEEK
    End If
    
    If IsMissingValue(r_rsScheduleRequests!SCHED_CD) Then
        r_rsLineSchedules!SCHED_CD = 0
    Else
        r_rsLineSchedules!SCHED_CD = r_rsScheduleRequests!SCHED_CD
    End If
    
    r_rsLineSchedules!JOB_NUM = r_rsScheduleRequests!JOB_NUM
    r_rsLineSchedules!REQ_LINE_IND = r_rsScheduleRequests!REQ_LINE_IND
    r_rsLineSchedules!CREATED_EMP_NUM = m_lngUserNum
    r_rsLineSchedules!LAST_MOD_EMP_NUM = m_lngUserNum
    r_rsLineSchedules.Update
        
    AppendWGLineCascadeDates "I", r_rsLineSchedules!WG_NUM, r_rsLineSchedules!LINE_NUM, _
                r_rsLineSchedules!START_DATE, r_rsLineSchedules!END_DATE
    
    If l_blnSplitPrevious Then
        r_rsLineSchedules.AddNew
        r_rsLineSchedules!WG_NUM = l_lngSplitWgNum
        r_rsLineSchedules!LINE_NUM = l_lngSplitLineNum
        r_rsLineSchedules!START_DATE = l_lngSplitStartDate
        r_rsLineSchedules!END_DATE = l_lngSplitEndDate
        r_rsLineSchedules!ROT_CD = l_lngSplitRotCd
        If l_lngSplitRotCd <> 0 Then
            r_rsLineSchedules!ROT_WEEK = _
                DetermineNewRotationWeek( _
                    l_lngSplitStartDate, _
                    l_lngSplitRotCd, _
                    l_datOrigStartDate, _
                    l_lngSplitRotWeek)
        Else
            r_rsLineSchedules!ROT_WEEK = l_lngSplitRotWeek
        End If
        r_rsLineSchedules!SCHED_CD = l_lngSplitSchedCd
        r_rsLineSchedules!JOB_NUM = l_strSplitJobNum
        r_rsLineSchedules!REQ_LINE_IND = l_strSplitReqLineInd
        r_rsLineSchedules!CREATED_EMP_NUM = l_lngSplitCreatedEmpNum
        r_rsLineSchedules!CREATED_TIMESTAMP = l_strSplitCreatedTS
        r_rsLineSchedules!LAST_MOD_EMP_NUM = m_lngUserNum
        r_rsLineSchedules.Update
        If r_rsLineSchedules!ROT_WEEK <> l_lngSplitRotWeek Then
            AppendWarningRotationWeekChange _
                trWarRotWeekAdjusted, _
                l_lngSplitRotWeek, _
                r_rsLineSchedules.Fields("ROT_WEEK").Value, _
                r_rsLineSchedules.Fields("WG_NUM").Value, _
                r_rsLineSchedules.Fields("LINE_NUM").Value, _
                r_rsLineSchedules.Fields("START_DATE").Value
        End If
    End If
    
    r_rsLineSchedules.MoveNext
    AdjustFutureLineSchedules l_lngLineNum, l_datStartDate, l_datEndDate, r_rsLineSchedules
    
    Exit Sub

eh:
    Select Case Err.Number
        Case trErrbusWGLineAssignExist
            Err.Source = FormatSchedBRV(TrType_WGM_SCHED, l_lngWgNum, l_lngLineNum, l_datStartDate)
            TrsRaiseError m_strCMODULENAME, "InsertLineSchedule", Err.Number
        Case Else
            TrsRaiseError m_strCMODULENAME, "InsertLineSchedule"
    End Select
End Sub

Private Function DetermineNewEndDate( _
            ByVal v_datEffEndDate As Date, _
            ByRef r_rsScheduleRequests As ADODB.Recordset) As Date
' If End date is not provided, set it to the end date of the previous effective.
' If there is no previous effective, set it to end of time.
    Dim l_datEndDate As Date
    Dim l_blnEndDateProvided As Boolean

On Error GoTo eh

    If IsNull(r_rsScheduleRequests!END_DATE) Then
        l_blnEndDateProvided = False
    ElseIf r_rsScheduleRequests!END_DATE = 0 Then
        l_blnEndDateProvided = False
    Else
        l_blnEndDateProvided = True
    End If
    
    If Not l_blnEndDateProvided Then
        If v_datEffEndDate = 0 Then
            l_datEndDate = #12/31/9999#
        Else
            l_datEndDate = v_datEffEndDate
        End If
    Else
        l_datEndDate = r_rsScheduleRequests!END_DATE
    End If
    DetermineNewEndDate = l_datEndDate
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "DetermineNewEndDate"
End Function


Private Sub UpdateLineSchedule( _
            ByRef r_rsScheduleRequests As ADODB.Recordset, _
            ByRef r_rsLineSchedules As ADODB.Recordset)
'    Only allowed to change start/end dates when an employee has not
'    reported time for the dates specified.  Must first remove actual time like
'    other potential schedule changes.
'    Assuming cascade will validate this based on the request.
'    An explicit update request may be updating a line schedule that was
'    previously updated, not saved, implicitly by another line schedule
'    update that created an overlap condition.  Some changes were required
'    to ensure the OriginalValue was compared instead of the potentially
'    updated Value.
'    However, it is assumed that an explicit update is only valid for a
'    persistent record.
            
    Dim l_lngWgNum As Long
    Dim l_lngLineNum As Long
    Dim l_datStartDate As Date
    Dim l_datEndDate As Date
    Dim l_strCreatedTS As String
    Dim l_datBeforeStartDate As Date
    Dim l_datBeforeEndDate As Date
    Dim l_blnEndDateProvided As Boolean
    Dim l_blnDateChange As Boolean
    Dim l_blnSchedChange As Boolean
    Dim l_strJobNum As String

On Error GoTo eh

    r_rsLineSchedules.Sort = "WG_NUM ASC, LINE_NUM ASC, START_DATE ASC"
    l_lngWgNum = r_rsScheduleRequests!WG_NUM
    l_lngLineNum = r_rsScheduleRequests!LINE_NUM
    l_datStartDate = r_rsScheduleRequests!START_DATE
    l_strCreatedTS = r_rsScheduleRequests!CREATED_TIMESTAMP

' If previous effective line sched record exists,
' set end date to day prior to this start date.
    FindEffectiveLineScheduleRecord l_lngWgNum, l_lngLineNum, l_datStartDate, r_rsLineSchedules
    If Not r_rsLineSchedules.EOF Then
        'Either back-dated START_DATE or forward-dated after another effective line schedule
        If Not OriginalStartDate(r_rsLineSchedules) = r_rsScheduleRequests!KEY_START_DATE Then
            If r_rsLineSchedules!START_DATE = l_datStartDate Then
                AppendWGLineCascadeDates "D", r_rsLineSchedules!WG_NUM, r_rsLineSchedules!LINE_NUM, _
                    r_rsLineSchedules!START_DATE, r_rsLineSchedules!END_DATE
                r_rsLineSchedules.Delete
            Else
                r_rsLineSchedules!END_DATE = l_datStartDate - 1
                r_rsLineSchedules!LAST_MOD_EMP_NUM = m_lngUserNum
                r_rsLineSchedules.Update
            End If
        End If
    End If
    
' If End date is not provided, set it to the end date of the next effective.
' If there is no next effective, set it to end of time.
    If IsNull(r_rsScheduleRequests!END_DATE) Then
        l_blnEndDateProvided = False
    ElseIf r_rsScheduleRequests!END_DATE = 0 Then
        l_blnEndDateProvided = False
    Else
        l_blnEndDateProvided = True
    End If
    If Not l_blnEndDateProvided Then
        FindNextLineScheduleRecord l_lngWgNum, l_lngLineNum, l_datStartDate, r_rsLineSchedules
        If r_rsLineSchedules.EOF Then
            l_datEndDate = #12/31/9999#
        Else
            l_datEndDate = r_rsLineSchedules!START_DATE - 1
        End If
    Else
        l_datEndDate = r_rsScheduleRequests!END_DATE
    End If
    
    FindLineScheduleRecord _
        l_lngWgNum, l_lngLineNum, r_rsScheduleRequests!KEY_START_DATE, l_strCreatedTS, r_rsLineSchedules
    If r_rsLineSchedules.EOF Then
        Err.Raise trErrbusInvalidChangeEntryNotFound
    End If
    
'    If the user extended END_DATE, creating overlap with next interval, the
'    system will adjust the START_DATE to avoid overlap.  The system will also
'    adjust the ROT_WEEK appropriately, so the schedule does not change
'    for employees assigned to the line.  An explicit update for the
'    previous implicit update may follow.  It is unnecessary if the adjusted ROT_WEEK,
'    from the implicit update, should stand.  However, if the user desires the
'    ROT_WEEK to not change, an explicit update is required to change the ROT_WEEK
'    back to the original ROT_WEEK.
    If r_rsLineSchedules!START_DATE <> l_datStartDate Or _
        r_rsLineSchedules!END_DATE <> l_datEndDate Then
        l_blnDateChange = True
    End If
    
'    ROT_WEEK may have been updated by previous implicit change.  If
'    this explicit update has the original ROT_WEEK, it is a schedule change.
'    If previous implicit update set the ROT_WEEK to the same ROT_WEEK because
'    the START_DATE change was the same number of weeks as the rotation, then
'    it is not a schedule change.
' JIMB Added Job Number to this check
    If IsNull(r_rsScheduleRequests!JOB_NUM) Then
        l_strJobNum = ""
    Else
        l_strJobNum = r_rsScheduleRequests!JOB_NUM
    End If
    If r_rsLineSchedules!SCHED_CD <> r_rsScheduleRequests!SCHED_CD Or _
        r_rsLineSchedules!ROT_CD <> r_rsScheduleRequests!ROT_CD Or _
        r_rsLineSchedules!ROT_WEEK <> r_rsScheduleRequests!ROT_WEEK Or _
        r_rsLineSchedules!JOB_NUM <> l_strJobNum Then
        l_blnSchedChange = True
    End If
        
'    Explicit START_DATE change.
'    For rotation, if START_DATE changed, it is implied schedule change.
'    The user could update the ROT_WEEK by the same number of weeks as the
'    START_DATE change, which would preserve the schedule.  For simplicity,
'    do not consider a ROT_WEEK change and mark entire interval as a
'    schedule change.
    If r_rsLineSchedules!ROT_CD <> 0 And _
            r_rsLineSchedules!START_DATE <> l_datStartDate Then
        l_blnSchedChange = True
    End If
        
    l_datBeforeStartDate = r_rsLineSchedules!START_DATE
    l_datBeforeEndDate = r_rsLineSchedules!END_DATE
        
'    Avoid gaps in line schedule when line assignments exist because
'    an employee can not be assigned to a line that does not have a schedule.
'    Only allowed to extend start date when an employee is not assigned
'    to the line, TTR_LINE_ASGNM.
'    Only allowed to shorten end date when an employee is not assigned
'    to the line, TTR_LINE_ASGNM.
'    If this line schedule had been implicitly updated previously, the l_datBeforeStartDate
'    will be the current START_DATE, not the persistent START_DATE.  Same for END_DATE.
'    The user may have extended the previous END_DATE, thus updating this START_DATE to
'    avoid overlap, and updated this START_DATE to be the same START_DATE.  Which is
'    required to maintain the ROT_WEEK.  This scenario will not check for assignments
'    since the START_DATEs are the same, assuming the previous END_DATE update was
'    processed first.
' 8/2/02 RSG - Moved edit to post validate.
'    If l_datStartDate > l_datBeforeStartDate Then
'        If AnyLineAssignments( _
'                r_rsLineSchedules!WG_NUM, r_rsLineSchedules!LINE_NUM, _
'                    l_datBeforeStartDate, l_datStartDate) Then
'            Err.Raise trErrbusWGLineAssignExist
'        End If
'    End If
'    If l_datEndDate < l_datBeforeEndDate Then
'        If AnyLineAssignments( _
'                r_rsLineSchedules!WG_NUM, r_rsLineSchedules!LINE_NUM, _
'                    l_datEndDate, l_datBeforeEndDate) Then
'            Err.Raise trErrbusWGLineAssignExist
'        End If
'    End If

    r_rsLineSchedules!START_DATE = l_datStartDate
    r_rsLineSchedules!END_DATE = l_datEndDate

    r_rsLineSchedules!SCHED_CD = r_rsScheduleRequests!SCHED_CD
    r_rsLineSchedules!ROT_CD = r_rsScheduleRequests!ROT_CD
'    If a warning has been issued for a ROT_WEEK change due to implicit START_DATE
'    change and this update will negate it, notify the user that the ROT_WEEK
'    was updated again.
    If r_rsLineSchedules.Fields("ROT_WEEK").OriginalValue <> _
        r_rsLineSchedules.Fields("ROT_WEEK").Value Then     'previous implicit update
        If r_rsLineSchedules.Fields("ROT_WEEK").OriginalValue = r_rsScheduleRequests!ROT_WEEK Then
            AppendWarningRotationWeekChange _
                trWarRotWeekChangedBack, _
                r_rsScheduleRequests.Fields("ROT_WEEK").Value, _
                r_rsLineSchedules.Fields("ROT_WEEK").Value, _
                r_rsLineSchedules.Fields("WG_NUM").Value, _
                r_rsLineSchedules.Fields("LINE_NUM").Value, _
                r_rsLineSchedules.Fields("START_DATE").Value
        End If
    End If
    r_rsLineSchedules!ROT_WEEK = r_rsScheduleRequests!ROT_WEEK
      
'    Only allowed to change REQ_LINE_IND when an employee is not
'    assigned to the line for the dates specified.

'JIMB removed this edit because there are
' no business rules based on the REQ_LINE_IND anymore
'''''    If Trim(r_rsLineSchedules!REQ_LINE_IND) <> Trim(r_rsScheduleRequests!REQ_LINE_IND) Then
'''''        If AnyLineAssignments( _
'''''            r_rsLineSchedules!WG_NUM, r_rsLineSchedules!LINE_NUM, _
'''''                    l_datStartDate, l_datEndDate) Then
'''''            Err.Raise trErrbusWGLineAssignExist
'''''        End If
'''''    End If
    
    If IsNull(r_rsScheduleRequests!JOB_NUM) Then
        r_rsLineSchedules!JOB_NUM = ""
    Else
        r_rsLineSchedules!JOB_NUM = r_rsScheduleRequests!JOB_NUM
    End If
    r_rsLineSchedules!REQ_LINE_IND = r_rsScheduleRequests!REQ_LINE_IND
    
    r_rsLineSchedules!LAST_MOD_EMP_NUM = m_lngUserNum
    r_rsLineSchedules.Update
    
    r_rsLineSchedules.MoveNext
    AdjustFutureLineSchedules l_lngLineNum, l_datStartDate, l_datEndDate, r_rsLineSchedules
    
    If l_blnSchedChange Then
        AppendWGLineCascadeDates "C", l_lngWgNum, l_lngLineNum, l_datBeforeStartDate, l_datBeforeEndDate
    End If
    If l_blnDateChange Then
        AppendCascadeDateDifference "C", l_lngWgNum, l_lngLineNum, _
            l_datStartDate, l_datEndDate, l_datBeforeStartDate, l_datBeforeEndDate
    End If
    
    Exit Sub

eh:
    Select Case Err.Number
        Case trErrbusWGLineAssignExist
            Err.Source = FormatSchedBRV(TrType_WGM_SCHED, l_lngWgNum, l_lngLineNum, l_datStartDate)
            TrsRaiseError m_strCMODULENAME, "UpdateLineSchedule", Err.Number
        Case trErrbusInvalidChangeEntryNotFound
            TrsRaiseError m_strCMODULENAME, "UpdateLineSchedule", Err.Number
        Case Else
            TrsRaiseError m_strCMODULENAME, "UpdateLineSchedule"
    End Select
End Sub


Private Function OriginalStartDate( _
                    ByVal v_rsLineSchedules As ADODB.Recordset) As Date
'   Answer the original value of START_DATE for the current record.
'   If current record is new, just answer the value.
On Error GoTo eh

    If v_rsLineSchedules.Status = adRecNew Then
        OriginalStartDate = v_rsLineSchedules!START_DATE
    Else
        OriginalStartDate = v_rsLineSchedules.Fields("START_DATE").OriginalValue
    End If

    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "OriginalStartDate"
End Function


Private Function AppendCascadeDateDifference( _
                    ByVal v_strAction As String, _
                    ByVal v_lngWgNum As Long, _
                    ByVal v_lngLineNum As Long, _
                    ByVal v_datStartDate As Date, _
                    ByVal v_datEndDate As Date, _
                    ByVal v_datKeyStartDate As Date, _
                    ByVal v_datKeyEndDate As Date)

On Error GoTo eh
    If v_datStartDate < v_datKeyStartDate Then
        AppendWGLineCascadeDates _
            "C", v_lngWgNum, v_lngLineNum, _
            v_datStartDate, _
            DateAdd("d", -1, v_datKeyStartDate)
    ElseIf v_datStartDate > v_datKeyStartDate Then
        AppendWGLineCascadeDates _
            "C", v_lngWgNum, v_lngLineNum, _
            v_datKeyStartDate, _
            DateAdd("d", -1, v_datStartDate)
    End If
    If v_datEndDate < v_datKeyEndDate Then
        AppendWGLineCascadeDates _
            "C", v_lngWgNum, v_lngLineNum, _
            DateAdd("d", 1, v_datEndDate), _
            v_datKeyEndDate
    ElseIf v_datEndDate > v_datKeyEndDate Then
        AppendWGLineCascadeDates _
            "C", v_lngWgNum, v_lngLineNum, _
            DateAdd("d", 1, v_datKeyEndDate), _
            v_datEndDate
    End If

    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "AppendCascadeDateDifference"
End Function


Private Function PreValidateDemandScheduleRequests( _
                ByRef r_rsScheduleRequests As ADODB.Recordset _
            ) As Boolean
    Dim l_blnIsValid As Boolean
    Dim l_strRecStatus As String
    Dim l_strReqLineInd As String
    Dim l_rsSchedReqClone As ADODB.Recordset
    Dim l_strFilter As String
    Dim l_blnEndDateProvided As Boolean
    
On Error GoTo eh
    
    l_blnIsValid = True
    
    If Not RecordsExist(r_rsScheduleRequests) Then
        AppendWGSchedError m_rsErrors, trErrbusValidationError, mstrCAPPLICATIONERR, "Empty request"
        l_blnIsValid = False
        GoTo xit
    Else
        r_rsScheduleRequests.MoveFirst
    End If
    
    Do Until r_rsScheduleRequests.EOF
        l_strRecStatus = Trim(UCase(r_rsScheduleRequests!RECORD_STATUS))
        Select Case l_strRecStatus
            Case "I", "C", "D"
                l_blnIsValid = ValidateKeyFields(r_rsScheduleRequests)
            Case Else
                AppendWGSchedLineError m_rsErrors, r_rsScheduleRequests, trErrbusValidationError, mstrCBUSINESSERR, "Invalid Record Status on the Request"
                l_blnIsValid = False
        End Select
        
        'If there are errors in the key fields, just get out now.
        If Not l_blnIsValid Then
            GoTo xit
        End If
        
        'ROT_CD and SCHED_CD are mutually exclusive
        If l_strRecStatus = "I" Or l_strRecStatus = "C" Then
            If IsMissingValue(r_rsScheduleRequests!ROT_CD) And _
                IsMissingValue(r_rsScheduleRequests!SCHED_CD) Then
                AppendWGSchedLineError m_rsErrors, r_rsScheduleRequests, _
                    trErrbusValidationError, mstrCBUSINESSERR, _
                    "Schedule is missing, please enter a Schedule code or Rotation code"
                l_blnIsValid = False
            ElseIf Not IsMissingValue(r_rsScheduleRequests!ROT_CD) And _
                Not IsMissingValue(r_rsScheduleRequests!SCHED_CD) Then
                AppendWGSchedLineError m_rsErrors, r_rsScheduleRequests, _
                    trErrbusValidationError, mstrCBUSINESSERR, _
                    "You can not enter a Schedule Code and a rotation code. " & _
                    "Please enter the correct schedule."
                l_blnIsValid = False
            End If
            If IsNull(r_rsScheduleRequests!END_DATE) Then
                l_blnEndDateProvided = False
            ElseIf r_rsScheduleRequests!END_DATE = 0 Then
                l_blnEndDateProvided = False
            Else
                l_blnEndDateProvided = True
            End If
            If l_blnEndDateProvided Then
                If r_rsScheduleRequests!END_DATE < r_rsScheduleRequests!START_DATE Then
                    AppendWGSchedLineError m_rsErrors, r_rsScheduleRequests, _
                        trErrbusValidationError, mstrCBUSINESSERR, _
                        "START_DATE cannot be less than END_DATE"
                    l_blnIsValid = False
                End If
            End If
        End If
        
        If Not IsNull(r_rsScheduleRequests!REQ_LINE_IND) Then
            l_strReqLineInd = Trim(UCase(r_rsScheduleRequests!REQ_LINE_IND))
            If Not IsYorN(l_strReqLineInd) Then
                AppendWGSchedLineError m_rsErrors, r_rsScheduleRequests, _
                    trErrbusValidationError, mstrCBUSINESSERR, _
                    "Required indicator is not Y or N"
                l_blnIsValid = False
            End If
        ElseIf r_rsScheduleRequests!RECORD_STATUS <> "D" Then
                AppendWGSchedLineError m_rsErrors, r_rsScheduleRequests, _
                    trErrbusValidationError, mstrCBUSINESSERR, _
                    "Required indicator can not be Null"
                l_blnIsValid = False
        End If
               
        r_rsScheduleRequests.MoveNext
    Loop
      
    'If there are errors, just get out now.
    If Not l_blnIsValid Then
        GoTo xit
    End If
    
    ' Finally, check for any duplicate inserts or deletes
    r_rsScheduleRequests.MoveFirst
    Set l_rsSchedReqClone = r_rsScheduleRequests.Clone
    Do Until r_rsScheduleRequests.EOF
        If r_rsScheduleRequests!RECORD_STATUS = "D" Or _
            r_rsScheduleRequests!RECORD_STATUS = "I" Then
            l_rsSchedReqClone.Filter = 0
            l_strFilter = "KEY_WG_NUM = " & r_rsScheduleRequests!KEY_WG_NUM & " AND " & _
                "KEY_LINE_NUM = " & r_rsScheduleRequests!KEY_LINE_NUM & " AND " & _
                "KEY_START_DATE = " & "#" & Format(r_rsScheduleRequests!KEY_START_DATE, "MM/DD/YYYY") & "#" & " AND " & _
                " RECORD_STATUS = " & "'" & r_rsScheduleRequests!RECORD_STATUS & "'"
            l_strFilter = MDAC25FilterFix(l_strFilter)
            l_rsSchedReqClone.Filter = l_strFilter
            If l_rsSchedReqClone.RecordCount > 1 Then
                AppendWGSchedLineError m_rsErrors, r_rsScheduleRequests, _
                    trErrbusValidationError, mstrCBUSINESSERR, _
                    "Duplicate request records found", , _
                    l_rsSchedReqClone!KEY_WG_NUM, _
                    l_rsSchedReqClone!KEY_LINE_NUM, _
                    l_rsSchedReqClone!KEY_START_DATE
                l_blnIsValid = False
            End If
        End If
        r_rsScheduleRequests.MoveNext
    Loop
      
xit:
    PreValidateDemandScheduleRequests = l_blnIsValid
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "PreValidateDemandScheduleRequests"
End Function


Private Function IsMissingValue(ByVal v_varData As Variant) As Boolean

On Error GoTo eh
    Select Case VarType(v_varData)
        Case vbEmpty
            IsMissingValue = True
        Case vbNull
            IsMissingValue = True
        Case vbInteger, vbLong
            If v_varData = 0 Then
                IsMissingValue = True
            Else
                IsMissingValue = False
            End If
        Case Else
            IsMissingValue = False
    End Select
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "IsMissingValue"
End Function


Private Function ValidateKeyFields( _
            r_rsScheduleRequests As ADODB.Recordset) As Boolean
'Answer True when valid, otherwise False.
    Dim l_blnIsValid As Boolean
    Dim l_strWGNum As String
    Dim l_strLineNum As String
    Dim l_datEffDate As Date
    Dim l_strCreatedTS As String

On Error GoTo eh

    l_blnIsValid = True
    If Not IsNull(r_rsScheduleRequests!KEY_WG_NUM) Then
        l_strWGNum = r_rsScheduleRequests!KEY_WG_NUM
' Jim will have already updated KEY_WG_NUM and WG_NUM with the new assigned number.
        If r_rsScheduleRequests!KEY_WG_NUM < 1 Then
            AppendWGSchedLineError m_rsErrors, r_rsScheduleRequests, _
                trErrbusValidationError, mstrCBUSINESSERR, _
                "KEY_WG_NUM must > 0", , _
                l_strWGNum
            l_blnIsValid = False
        End If
    Else
        l_strWGNum = "0"
        AppendWGSchedLineError m_rsErrors, r_rsScheduleRequests, _
            trErrbusValidationError, mstrCBUSINESSERR, _
            "KEY_WG_NUM cannot be Null", , l_strWGNum
        l_blnIsValid = False
    End If

    If Not IsNull(r_rsScheduleRequests!KEY_LINE_NUM) Then
        l_strLineNum = r_rsScheduleRequests!KEY_LINE_NUM
        If r_rsScheduleRequests!KEY_LINE_NUM < 1 Then
            AppendWGSchedLineError m_rsErrors, r_rsScheduleRequests, _
                trErrbusValidationError, mstrCBUSINESSERR, _
                "KEY_LINE_NUM must > 0", , _
                l_strWGNum, l_strLineNum
            l_blnIsValid = False
        End If
    Else
        l_strLineNum = "0"
        AppendWGSchedLineError m_rsErrors, r_rsScheduleRequests, _
            trErrbusValidationError, mstrCBUSINESSERR, _
            "KEY_LINE_NUM cannot be Null", , l_strWGNum, l_strLineNum
        l_blnIsValid = False
    End If

    If Not IsNull(r_rsScheduleRequests!KEY_START_DATE.Value) Then
        If Not IsDateValid(r_rsScheduleRequests!KEY_START_DATE.Value) Then
            AppendWGSchedLineError m_rsErrors, r_rsScheduleRequests, _
                trErrbusValidationError, mstrCBUSINESSERR, _
                "KEY_START_DATE is invalid", , _
                l_strWGNum, l_strLineNum, r_rsScheduleRequests!KEY_START_DATE.Value
            l_blnIsValid = False
        End If
    Else
        If Not IsNull(r_rsScheduleRequests!KEY_START_DATE.Value) Then
            AppendWGSchedLineError m_rsErrors, r_rsScheduleRequests, _
                trErrbusValidationError, mstrCBUSINESSERR, _
                "KEY_START_DATE cannot be Null", , l_strWGNum, l_strLineNum, 0
            l_blnIsValid = False
        Else
            AppendWGSchedLineError m_rsErrors, r_rsScheduleRequests, _
                trErrbusValidationError, mstrCBUSINESSERR, _
                "Please enter a start date for the line schedule", , l_strWGNum, l_strLineNum, 0
            l_blnIsValid = False
        End If
    End If
    
    If Not IsNull(r_rsScheduleRequests!KEY_WG_NUM) And _
        Not IsNull(r_rsScheduleRequests!WG_NUM) Then
        l_strWGNum = r_rsScheduleRequests!WG_NUM
        If r_rsScheduleRequests!KEY_WG_NUM <> r_rsScheduleRequests!WG_NUM Then
            AppendWGSchedLineError m_rsErrors, r_rsScheduleRequests, _
                trErrbusValidationError, mstrCBUSINESSERR, _
                "Please contact the TRIS team about this message. " & _
                "You can not change the work group of the line schedule. " & _
                "WG_NUM must = KEY_WG_NUM", , _
                l_strWGNum
            l_blnIsValid = False
        End If
    ElseIf IsNull(r_rsScheduleRequests!WG_NUM) Then
        l_strWGNum = "0"
        AppendWGSchedLineError m_rsErrors, r_rsScheduleRequests, _
            trErrbusValidationError, mstrCBUSINESSERR, _
            "WG_NUM cannot be Null", , _
            l_strWGNum
        l_blnIsValid = False
    End If
    
    If Not IsNull(r_rsScheduleRequests!KEY_LINE_NUM) And _
        Not IsNull(r_rsScheduleRequests!LINE_NUM) Then
        l_strLineNum = r_rsScheduleRequests!LINE_NUM
        If r_rsScheduleRequests!KEY_LINE_NUM <> r_rsScheduleRequests!LINE_NUM Then
            AppendWGSchedLineError m_rsErrors, r_rsScheduleRequests, _
                trErrbusValidationError, mstrCBUSINESSERR, _
                "The line number of the schedule can not change. " & _
                "Please add a new schedule line or change the schedule of line " & _
                str(r_rsScheduleRequests!LINE_NUM) & ".", , _
                l_strWGNum, l_strLineNum
            l_blnIsValid = False
        End If
    ElseIf IsNull(r_rsScheduleRequests!LINE_NUM) Then
        l_strLineNum = "0"
        AppendWGSchedLineError m_rsErrors, r_rsScheduleRequests, _
            trErrbusValidationError, mstrCBUSINESSERR, _
            "The line number of the schedule is missing", , _
            l_strWGNum, l_strLineNum
        l_blnIsValid = False
    End If
    
    If Not IsNull(r_rsScheduleRequests!KEY_START_DATE) And _
        Not IsNull(r_rsScheduleRequests!START_DATE) Then
        l_datEffDate = r_rsScheduleRequests!START_DATE
' START_DATE may change.
        If r_rsScheduleRequests!KEY_START_DATE <> r_rsScheduleRequests!START_DATE And _
                r_rsScheduleRequests!RECORD_STATUS <> "C" Then
            AppendWGSchedLineError m_rsErrors, r_rsScheduleRequests, _
                trErrbusValidationError, mstrCBUSINESSERR, _
                "START_DATE must = KEY_START_DATE", , _
                l_strWGNum, l_strLineNum, l_datEffDate
            l_blnIsValid = False
        End If
'        ElseIf r_rsScheduleRequests!KEY_START_DATE < Date Then
'
' TEST ONLY
'
'            AppendWGSchedLineError m_rsErrors, r_rsScheduleRequests, _
'                trErrbusValidationError, mstrCBUSINESSERR, _
'                "EFF_DATE must >= current date", , _
'                l_strWGNum, l_strLineNum, l_datEffDate
'            l_blnIsValid = False
'        End If
    ElseIf IsNull(r_rsScheduleRequests!START_DATE) Then
        If Not IsNull(r_rsScheduleRequests!KEY_START_DATE.Value) Then
            l_datEffDate = 0
            AppendWGSchedLineError m_rsErrors, r_rsScheduleRequests, _
                trErrbusValidationError, mstrCBUSINESSERR, _
                "Please enter a start date for the line schedule", , _
                l_strWGNum, l_strLineNum, l_datEffDate
            l_blnIsValid = False
        End If
    End If
    
    If IsNull(r_rsScheduleRequests!CREATED_TIMESTAMP) Then
        If r_rsScheduleRequests!RECORD_STATUS <> "I" Then
            l_strCreatedTS = ""
            AppendWGSchedLineError m_rsErrors, r_rsScheduleRequests, _
                trErrbusValidationError, mstrCBUSINESSERR, _
                "CREATED TIMESTAMP cannot be Null", , _
                l_strWGNum, l_strLineNum, l_datEffDate
            l_blnIsValid = False
        End If
    ElseIf Not IsDB2Timestamp(r_rsScheduleRequests!CREATED_TIMESTAMP) Then
        If r_rsScheduleRequests!RECORD_STATUS = "I" And _
            Trim(r_rsScheduleRequests!CREATED_TIMESTAMP) = "" Then
            'OK
        Else
            l_strCreatedTS = ""
            AppendWGSchedLineError m_rsErrors, r_rsScheduleRequests, _
                trErrbusValidationError, mstrCBUSINESSERR, _
                "CREATED TIMESTAMP is not valid", , _
                l_strWGNum, l_strLineNum, l_datEffDate
            l_blnIsValid = False
        End If
    End If
    
    ValidateKeyFields = l_blnIsValid
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "ValidateKeyFields"
End Function

'
' Move this to validation module
'
Private Function IsDateValid(aDate As Variant) As Boolean
'Reject aDate when unitialized(vbEmpty) or set to vbNull.
'Guard against Null.

On Error GoTo eh
    If IsNull(aDate) Then
        IsDateValid = False
    Else
        IsDateValid = _
            IsDate(aDate) And aDate <> vbEmpty And aDate <> vbNull
    End If
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "IsDateValid"
End Function


Private Function PostValidateScheduleUpdates( _
                ByRef r_rsLineSchedules As ADODB.Recordset, _
                ByRef r_rsScheduleRequests As ADODB.Recordset _
            ) As Boolean
' Return True when all changes are valid; otherwise return False (invalid
' or no changes recorded.)
    Dim l_blnIsValid As Boolean
    Dim l_intUpdateCount As Integer
    Dim l_datWGEffStartDate As Date
    
On Error GoTo eh
    
    l_blnIsValid = True
    
    If Not RecordsExist(r_rsLineSchedules) Then
        AppendWGSchedError m_rsErrors, _
            trErrbusValidationError, mstrCBUSINESSWAR, _
            "Line schedule recordset is empty."
        PostValidateScheduleUpdates = False
        Exit Function
    Else
        r_rsLineSchedules.MoveFirst
    End If
    
    If RecordsExist(m_rsWorkGroupAttributes) Then
        m_rsWorkGroupAttributes.MoveFirst
        l_datWGEffStartDate = m_rsWorkGroupAttributes!EFF_START_DATE
    End If
    
'only validate new or modified records
    r_rsLineSchedules.Filter = adFilterPendingRecords
    Do Until r_rsLineSchedules.EOF
        Select Case r_rsLineSchedules.Status
            Case adRecDeleted
                l_intUpdateCount = l_intUpdateCount + 1
            Case adRecNew, adRecModified
                l_intUpdateCount = l_intUpdateCount + 1
                If Not r_rsLineSchedules!LINE_NUM > 0 Then
                    l_blnIsValid = False
                    AppendWGSchedLineError m_rsErrors, r_rsLineSchedules, _
                        trErrbusValidationError, mstrCBUSINESSERR, _
                        "Line Number must be > 0"
                End If
                
                If Not IsValidWorkgroup(r_rsLineSchedules!WG_NUM, r_rsLineSchedules!START_DATE) Then
                    l_blnIsValid = False
                    AppendWGSchedLineError m_rsErrors, r_rsLineSchedules, _
                        trErrbusValidationError, mstrCBUSINESSERR, _
                        "Workgroup number is not valid on: " & Format(r_rsLineSchedules!START_DATE, "mm/dd/yyyy")
                End If
                
                If r_rsLineSchedules!START_DATE < l_datWGEffStartDate Then
                    l_blnIsValid = False
                    AppendWGSchedLineError m_rsErrors, r_rsLineSchedules, _
                        trErrbusValidationError, mstrCBUSINESSERR, _
                        "Line schedule start date(" & _
                        Format(r_rsLineSchedules!START_DATE, "mm/dd/yyyy") & _
                        ") cannot be less than workgroup start date(" & _
                        Format(l_datWGEffStartDate, "mm/dd/yyyy") & ")"
                End If
                
                If r_rsLineSchedules!END_DATE < r_rsLineSchedules!START_DATE Then
                    l_blnIsValid = False
                    AppendWGSchedLineError m_rsErrors, r_rsLineSchedules, _
                        trErrbusValidationError, mstrCBUSINESSERR, _
                        "Line schedule end date(" & _
                        Format(r_rsLineSchedules!END_DATE, "mm/dd/yyyy") & _
                        ") cannot be less than start date(" & _
                        Format(r_rsLineSchedules!START_DATE, "mm/dd/yyyy") & ")"
                End If
                
                If r_rsLineSchedules!SCHED_CD = 0 Then
                    If r_rsLineSchedules!ROT_CD = 0 Then
                        l_blnIsValid = False    'mutex - one must be provided
                        AppendWGSchedLineError m_rsErrors, r_rsLineSchedules, _
                            trErrbusValidationError, mstrCBUSINESSERR, _
                            "Schedule or Rotation code are required"
                    Else
                        If Not IsValidRotationCode(r_rsLineSchedules!ROT_CD) Then
                            l_blnIsValid = False
                            AppendWGSchedLineError m_rsErrors, r_rsLineSchedules, _
                                trErrbusValidationError, mstrCBUSINESSERR, _
                                "Rotation code is not valid"
                        End If
                        If Not IsValidRotationWeek(r_rsLineSchedules!ROT_CD, r_rsLineSchedules!ROT_WEEK) Then
                            l_blnIsValid = False
                            AppendWGSchedLineError m_rsErrors, r_rsLineSchedules, _
                                trErrbusValidationError, mstrCBUSINESSERR, _
                                "Rotation week is not valid"
                        End If
                    End If
                Else
                    If r_rsLineSchedules!ROT_CD = 0 Then
                        If Not (r_rsLineSchedules!ROT_WEEK = 0) Then
                            l_blnIsValid = False
                            AppendWGSchedLineError m_rsErrors, r_rsLineSchedules, _
                                trErrbusValidationError, mstrCBUSINESSERR, _
                                "Rotation week is not valid without Rotation code"
                        End If
                        If Not IsValidScheduleCode(r_rsLineSchedules!SCHED_CD) Then
                            l_blnIsValid = False
                            AppendWGSchedLineError m_rsErrors, r_rsLineSchedules, _
                                trErrbusValidationError, mstrCBUSINESSERR, _
                                "Schedule code is not valid"
                        End If
                    Else
                        l_blnIsValid = False    'mutex - both cannot be provided
                        AppendWGSchedLineError m_rsErrors, r_rsLineSchedules, _
                            trErrbusValidationError, mstrCBUSINESSERR, _
                            "Schedule or Rotation code are required, not both"
                    End If
                End If
                
                If Not IsValidJobNumber(r_rsLineSchedules!JOB_NUM, r_rsLineSchedules!START_DATE) Then
                    l_blnIsValid = False
                    AppendWGSchedLineError m_rsErrors, r_rsLineSchedules, _
                        trErrbusValidationError, mstrCBUSINESSERR, _
                        "Job number is not valid on: " & Format(r_rsLineSchedules!START_DATE, "mm/dd/yyyy")
                End If
                If Not IsYorN(UCase(Trim(r_rsLineSchedules!REQ_LINE_IND))) Then
                    l_blnIsValid = False
                    AppendWGSchedLineError m_rsErrors, r_rsLineSchedules, _
                        trErrbusValidationError, mstrCBUSINESSERR, _
                        "Required indicator must be Y or N"
                End If
        End Select
        r_rsLineSchedules.MoveNext
    Loop
    
    If LineTradesExist() Then
        l_blnIsValid = False
    End If
    
    r_rsLineSchedules.Filter = adFilterNone
    If CheckForOverlappingLineSchedules(r_rsLineSchedules) Then
        l_blnIsValid = False
    End If
    
    If CheckForGapsInLineSchedules(r_rsLineSchedules) Then
        l_blnIsValid = False
    End If
    
    If l_intUpdateCount = 0 Then
        AppendWGSchedError m_rsErrors, _
            trErrbusValidationError, mstrCBUSINESSWAR, _
            "No line schedule updates were recorded."
    End If
        
    PostValidateScheduleUpdates = l_blnIsValid
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "PostValidateScheduleUpdates"
End Function


Private Function LineTradesExist() As Boolean
'   Changing the underlying schedule of a trade is not allowed.  Get all
'   demand line trades in the scope of line schedule changes and check for
'   existence of a trade.  If found, append an error and return True.

    Dim l_blnExists As Boolean
    Dim l_rsLineTrades As ADODB.Recordset
    Dim l_datStart As Date
    Dim l_datEnd As Date
    Dim l_strFilter As String
    Dim l_strErrMsg As String
    Dim l_colLineDateIntervals As Collection
    Dim l_varWorkgroups As Variant
    Dim l_varLines As Variant
    Dim l_cdSchedTrade As dbSchedTrade.cdSchedTrade
    Dim l_cpuDateInterval As cpuDateInterval
                    
On Error GoTo eh
                    
    l_blnExists = False
    
    Set l_colLineDateIntervals = ConsolidatedCascadeWGLineIntervals
    
    DetermineGetDemandLineTradesScope l_colLineDateIntervals, l_varWorkgroups, l_varLines, l_datStart, l_datEnd
    If IsEmpty(l_varWorkgroups) Then
        LineTradesExist = False
        Exit Function
    End If
                    
    Set l_cdSchedTrade = CreateObject("dbSchedTrade.cdSchedTrade")
    
    Set l_rsLineTrades = _
        l_cdSchedTrade.GetDemandLineTrades(l_varWorkgroups, l_varLines, l_datStart, l_datEnd)
    Set l_cdSchedTrade = Nothing

    If l_rsLineTrades.EOF Then
        LineTradesExist = False
        Set l_rsLineTrades = Nothing
        Exit Function
    End If
    
    For Each l_cpuDateInterval In l_colLineDateIntervals
        If LineScheduleChanged(l_cpuDateInterval.WorkgroupNumber, _
                    l_cpuDateInterval.LineNumber, _
                    l_cpuDateInterval.StartDate, _
                    l_cpuDateInterval.EndDate) Then
            l_strFilter = _
                "WG_NUM = " & l_cpuDateInterval.WorkgroupNumber & _
                " AND LINE_NUM = " & l_cpuDateInterval.LineNumber & _
                " AND CHARGE_DATE >= #" & Format(l_cpuDateInterval.StartDate, "mm/dd/yyyy") & "#" & _
                " AND CHARGE_DATE <= #" & Format(l_cpuDateInterval.EndDate, "mm/dd/yyyy") & "#"
            l_strFilter = MDAC25FilterFix(l_strFilter)
            l_rsLineTrades.Filter = l_strFilter
            If RecordsExist(l_rsLineTrades) Then
                Do Until l_rsLineTrades.EOF
                    l_strErrMsg = _
                        "Workgroup " & CStr(l_rsLineTrades!WG_NUM) & _
                        "/Line " & CStr(l_rsLineTrades!LINE_NUM) & _
                        " has existing trade [" & l_rsLineTrades!TRADE_NUM & "] on " & _
                        Format(l_rsLineTrades!CHARGE_DATE, "mm/dd/yyyy") & "."
                    AppendBusinessError _
                        m_rsErrors, _
                        trErrbusTradeExists, _
                        mstrCBUSINESSERR, _
                        l_strErrMsg, _
                        m_strCMODULENAME, _
                        TrType_WGM_SCHED, _
                        , , _
                        CDEMANDSCHEDFIELDNAMES, _
                        CStr(l_cpuDateInterval.WorkgroupNumber), _
                        Format(l_cpuDateInterval.StartDate, "mm/dd/yyyy"), _
                        CStr(l_cpuDateInterval.LineNumber)
                    l_rsLineTrades.MoveNext
                Loop
                l_blnExists = True
            End If
        End If
    Next

    LineTradesExist = l_blnExists
    Set l_rsLineTrades = Nothing

    Exit Function

eh:
    Set l_cdSchedTrade = Nothing
    Set l_rsLineTrades = Nothing
    TrsRaiseError m_strCMODULENAME, "LineTradesExist"
End Function


Private Function EmployeeTradesExist( _
                    ByRef r_rsEmpSchedDates As ADODB.Recordset, _
                    ByRef r_rsErrors As ADODB.Recordset) As Boolean
'   Changing the underlying schedule of a trade is not allowed.  Get all
'   employee trades and determine if the workgroup update would be
'   considered an underlying change.  If the employee is assigned to
'   a line or has a personal schedule on the charge date of the trade,
'   it would not be an underlying schedule change.  Otherwise, it would
'   and an error will be returned stating a conflicting employee trade.

    Dim l_blnExists As Boolean
    Dim l_rsLineAssignments As ADODB.Recordset
    Dim l_rsPersonalSchedules As ADODB.Recordset
    Dim l_rsEmpTrades As ADODB.Recordset
    Dim l_datStart As Date
    Dim l_datEnd As Date
    Dim l_strFilter As String
    Dim l_strErrMsg As String
    Dim l_varEmployees As Variant
    Dim l_cdSchedule As dbSchedule.cdSchedule
                    
On Error GoTo eh
                    
    l_blnExists = False
    DetermineGetActualTimeScope r_rsEmpSchedDates, l_varEmployees, l_datStart, l_datEnd
    If IsEmpty(l_varEmployees) Then
        l_blnExists = False
        GoTo xit
    End If
                              
    Set l_cdSchedule = CreateObject("dbSchedule.cdSchedule")
    Set l_rsEmpTrades = _
        l_cdSchedule.GetEmployeeTrades(l_varEmployees, l_datStart, l_datEnd)
    If Not RecordsExist(l_rsEmpTrades) Then
        l_blnExists = False
        GoTo xit
    End If
                    
    Set l_rsLineAssignments = _
        l_cdSchedule.GetEmployeeLineAssignments(l_varEmployees, l_datStart, l_datEnd)
    Set l_rsPersonalSchedules = _
        l_cdSchedule.GetEmployeePersonalSchedule(l_varEmployees, l_datStart, l_datEnd)
    Set l_cdSchedule = Nothing
        
    r_rsEmpSchedDates.MoveFirst
    Do Until r_rsEmpSchedDates.EOF
        l_strFilter = _
            "EMP_NUM = " & r_rsEmpSchedDates!EMP_NUM & _
            " AND CHARGE_DATE >= #" & Format(r_rsEmpSchedDates!START_DATE, "mm/dd/yyyy") & "#" & _
            " AND CHARGE_DATE <= #" & Format(r_rsEmpSchedDates!END_DATE, "mm/dd/yyyy") & "#"
        l_strFilter = MDAC25FilterFix(l_strFilter)
        l_rsEmpTrades.Filter = l_strFilter
        If RecordsExist(l_rsEmpTrades) Then
            Do Until l_rsEmpTrades.EOF
                If TradeIsBasedOnWorkgroupSchedule( _
                        l_rsEmpTrades, l_rsLineAssignments, l_rsPersonalSchedules) Then
                    l_blnExists = True
                    l_strErrMsg = _
                        "Employee " & CStr(l_rsEmpTrades!EMP_NUM) & _
                        "[" & Trim(l_rsEmpTrades!EMP_NAME) & "] has existing employee trade on " & _
                        Format(l_rsEmpTrades!CHARGE_DATE, "mm/dd/yyyy") & "."
                    AppendBusinessError _
                        r_rsErrors, _
                        trErrbusTradeExists, _
                        mstrCBUSINESSERR, _
                        l_strErrMsg, _
                        m_strCMODULENAME, _
                        TrType_WGM_EMP_ACT, _
                        , _
                        , _
                        CWGMAINTFIELDNAMES, _
                        CStr(m_lngWorkGroup), _
                        Format(r_rsEmpSchedDates!START_DATE, "mm/dd/yyyy"), _
                        "0", _
                        CStr(r_rsEmpSchedDates!EMP_NUM)
                End If
                l_rsEmpTrades.MoveNext
            Loop
        End If
        r_rsEmpSchedDates.MoveNext
    Loop
    r_rsEmpSchedDates.MoveFirst

xit:
    Set l_cdSchedule = Nothing
    Set l_rsLineAssignments = Nothing
    Set l_rsPersonalSchedules = Nothing
    Set l_rsEmpTrades = Nothing
    EmployeeTradesExist = l_blnExists
    Exit Function

eh:
    Set l_cdSchedule = Nothing
    Set l_rsLineAssignments = Nothing
    Set l_rsPersonalSchedules = Nothing
    Set l_rsEmpTrades = Nothing
    TrsRaiseError m_strCMODULENAME, "EmployeeTradesExist"
End Function


Private Function TradeIsBasedOnWorkgroupSchedule( _
                    ByVal v_rsEmpTrades As ADODB.Recordset, _
                    ByVal v_rsLineAssignments As ADODB.Recordset, _
                    ByVal v_rsPersonalSchedules As ADODB.Recordset) As Boolean
'   If a line assignment or personal schedule exists for the charge date
'   of the trade, answer False; otherwise answer True.
    Dim l_lngEmpNum As Long
    Dim l_datChgDate As Date
    Dim l_strFilter As String

On Error GoTo eh
        
    l_lngEmpNum = v_rsEmpTrades!EMP_NUM
    l_datChgDate = v_rsEmpTrades!CHARGE_DATE
    
    If RecordsExist(v_rsLineAssignments) Then
        l_strFilter = _
            "EMP_NUM = " & l_lngEmpNum & _
            " AND START_DATE <= #" & Format(l_datChgDate, "mm/dd/yyyy") & "#" & _
            " AND END_DATE >= #" & Format(l_datChgDate, "mm/dd/yyyy") & "#"
        l_strFilter = MDAC25FilterFix(l_strFilter)
        v_rsLineAssignments.Filter = l_strFilter
        If RecordsExist(v_rsLineAssignments) Then
            TradeIsBasedOnWorkgroupSchedule = False
            v_rsLineAssignments.Filter = adFilterNone
            Exit Function
        End If
        v_rsLineAssignments.Filter = adFilterNone
    End If
        
    If RecordsExist(v_rsPersonalSchedules) Then
        l_strFilter = _
            "EMP_NUM = " & l_lngEmpNum & _
            " AND START_DATE <= #" & Format(l_datChgDate, "mm/dd/yyyy") & "#" & _
            " AND END_DATE >= #" & Format(l_datChgDate, "mm/dd/yyyy") & "#"
        l_strFilter = MDAC25FilterFix(l_strFilter)
        v_rsPersonalSchedules.Filter = l_strFilter
        If RecordsExist(v_rsPersonalSchedules) Then
            TradeIsBasedOnWorkgroupSchedule = False
            v_rsPersonalSchedules.Filter = adFilterNone
            Exit Function
        End If
        v_rsPersonalSchedules.Filter = adFilterNone
    End If

    TradeIsBasedOnWorkgroupSchedule = True
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "TradeIsBasedOnWorkgroupSchedule"
End Function


Private Sub DetermineGetDemandLineTradesScope( _
                    ByVal v_colWGLineDateIntervals As Collection, _
                    ByRef r_varWorkgroups As Variant, _
                    ByRef r_varLines As Variant, _
                    ByRef r_datStartDate As Date, _
                    ByRef r_datEndDate As Date)
    Dim l_datStartDate As Date
    Dim l_datEndDate As Date
    Dim l_colDistinctWorkgroupNums As cuSet
    Dim l_colDistinctLineNums As cuSet
    Const l_datLowDate As Date = #1/1/1900#
    Const l_datHighDate As Date = #12/31/2079#
    Dim l_cpuDateInterval As cpuDateInterval

On Error GoTo eh

'guard for empty collection
    If v_colWGLineDateIntervals.Count = 0 Then
        r_varWorkgroups = Empty
        r_varLines = Empty
        r_datStartDate = 0
        r_datEndDate = 0
        Exit Sub
    End If

' seed the dates with high and low dates
' there should be at least 1 valid record
    l_datStartDate = l_datHighDate
    l_datEndDate = l_datLowDate
    Set l_colDistinctWorkgroupNums = New cuSet
    Set l_colDistinctLineNums = New cuSet
    For Each l_cpuDateInterval In v_colWGLineDateIntervals
        l_colDistinctWorkgroupNums.Add l_cpuDateInterval.WorkgroupNumber
        l_colDistinctLineNums.Add l_cpuDateInterval.LineNumber
        l_datStartDate = MinDate(l_cpuDateInterval.StartDate, l_datStartDate)
        l_datEndDate = MaxDate(l_cpuDateInterval.EndDate, l_datEndDate)
    Next

    r_datStartDate = l_datStartDate
    r_datEndDate = l_datEndDate
    r_varWorkgroups = ArrayFromCollection(l_colDistinctWorkgroupNums.Contents)
    r_varLines = ArrayFromCollection(l_colDistinctLineNums.Contents)
    Set l_colDistinctWorkgroupNums = Nothing
    Set l_colDistinctLineNums = Nothing
    Exit Sub

eh:
    Set l_colDistinctWorkgroupNums = Nothing
    Set l_colDistinctLineNums = Nothing
    TrsRaiseError m_strCMODULENAME, "DetermineGetDemandLineTradesScope"
End Sub


Private Function CheckForGapsInLineSchedules( _
                    ByVal v_rsLineSchedules As ADODB.Recordset) As Boolean
' Ensure that gaps are not created where there are existing line assignments.
' Check for gaps created during this transaction only.  Match the wg/line intervals
' to any gaps found.
' Check for any line assignments that exist for each gap found.
' If line assignment found, write a BRV and answer True.

    Dim l_colWgLineSchedChgs As Collection
    Dim l_colWgLineGaps As Collection
    Dim l_cpuGap As cpuDateInterval
    Dim l_strErrorDesc As String

On Error GoTo eh

    CheckForGapsInLineSchedules = False
    If RecordsExist(v_rsLineSchedules) Then
        v_rsLineSchedules.MoveFirst
    Else
        'still need to check for deletes here, may be the only records in v_rsLineSchedules
        'if so, they are all gaps now.
        Set l_colWgLineGaps = GetAllWGLineDeletes(v_rsLineSchedules)
        If l_colWgLineGaps Is Nothing Then
            Exit Function
        End If
    End If

    If RecordsExist(v_rsLineSchedules) Then
        Set l_colWgLineSchedChgs = ConsolidatedCascadeWGLineIntervals
        Set l_colWgLineGaps = DetermineWGLineGaps(v_rsLineSchedules, l_colWgLineSchedChgs)
    Else    'proceed to process deletes only
    End If

    For Each l_cpuGap In l_colWgLineGaps
        If AnyLineAssignments( _
                l_cpuGap.WorkgroupNumber, l_cpuGap.LineNumber, _
                    l_cpuGap.StartDate, l_cpuGap.EndDate) Then
            l_strErrorDesc = _
                TrsGetErrDescription( _
                    LongAppErr(trErrDemandLineAssignExist), _
                    CollectionFromArray( _
                    Array(Format(l_cpuGap.StartDate, "mm/dd/yyyy"), _
                        Format(l_cpuGap.EndDate, "mm/dd/yyyy"), _
                        CStr(l_cpuGap.LineNumber)), 0))
            AppendBusinessError _
                m_rsErrors, _
                LongAppErr(trErrDemandLineAssignExist), _
                mstrCBUSINESSERR, _
                l_strErrorDesc, _
                m_strCMODULENAME & "CheckForGapsInLineSchedules", _
                TrType_WGM_SCHED, _
                , , _
                CDEMANDSCHEDFIELDNAMES, _
                CStr(l_cpuGap.WorkgroupNumber), _
                Format(l_cpuGap.StartDate, "mm/dd/yyyy"), _
                CStr(l_cpuGap.LineNumber)
            CheckForGapsInLineSchedules = True
        End If
    Next

    Set l_colWgLineSchedChgs = Nothing
    Set l_colWgLineGaps = Nothing
    Set l_cpuGap = Nothing
    Exit Function

eh:
    Set l_colWgLineSchedChgs = Nothing
    Set l_colWgLineGaps = Nothing
    Set l_cpuGap = Nothing
    TrsRaiseError m_strCMODULENAME, "CheckForGapsInLineSchedules"
End Function


Private Function GetAllWGLineDeletes( _
                    ByVal v_rsLineSchedules As ADODB.Recordset) As Collection
' No records visible, so check for deletes.  If any found, answer them as gaps.
    Dim l_colWgLineGaps As Collection
    Dim l_cpuGap As cpuDateInterval
    
On Error GoTo eh
    
    If v_rsLineSchedules Is Nothing Then
        Set GetAllWGLineDeletes = Nothing
        Exit Function
    End If

    Set l_colWgLineGaps = New Collection
    v_rsLineSchedules.Filter = adFilterPendingRecords
    Do Until v_rsLineSchedules.EOF
        Set l_cpuGap = New cpuDateInterval
        l_cpuGap.WorkgroupNumber = v_rsLineSchedules.Fields("WG_NUM").OriginalValue
        l_cpuGap.LineNumber = v_rsLineSchedules.Fields("LINE_NUM").OriginalValue
        l_cpuGap.StartDate = v_rsLineSchedules.Fields("START_DATE").OriginalValue
        l_cpuGap.EndDate = v_rsLineSchedules.Fields("END_DATE").OriginalValue
        v_rsLineSchedules.MoveNext
    Loop

    If l_colWgLineGaps.Count > 0 Then
        Set GetAllWGLineDeletes = l_colWgLineGaps
    Else
        Set GetAllWGLineDeletes = Nothing
    End If
    
    Set l_colWgLineGaps = Nothing
    Set l_cpuGap = Nothing
    v_rsLineSchedules.Filter = adFilterNone
    Exit Function

eh:
    Set l_colWgLineGaps = Nothing
    Set l_cpuGap = Nothing
    TrsRaiseError m_strCMODULENAME, "GetAllWGLineDeletes"
End Function


Private Function DetermineWGLineGaps( _
                    ByVal v_rsLineSchedules As ADODB.Recordset, _
                    ByVal v_colWgLineIntervals As Collection) As Collection
' Find gaps in line schedules.
' If a gap is found (not overlapping) and it intersects with a WgLine interval, answer it.

    Dim l_lngWgNum As Long
    Dim l_lngLineNum As Long
    Dim l_datEndDate As Date
    Dim l_cpuGap As cpuDateInterval
    Dim l_colGaps As Collection
    Dim l_colDeletes As Collection

On Error GoTo eh

    Set l_colGaps = New Collection
    Set l_colDeletes = New Collection
    
'    capture deletes
    v_rsLineSchedules.Filter = adFilterPendingRecords
    Do Until v_rsLineSchedules.EOF
        If v_rsLineSchedules.Status = adRecDeleted Then
            Set l_cpuGap = New cpuDateInterval
            l_cpuGap.WorkgroupNumber = v_rsLineSchedules.Fields("WG_NUM").OriginalValue
            l_cpuGap.LineNumber = v_rsLineSchedules.Fields("LINE_NUM").OriginalValue
            l_cpuGap.StartDate = v_rsLineSchedules.Fields("START_DATE").OriginalValue
            l_cpuGap.EndDate = v_rsLineSchedules.Fields("END_DATE").OriginalValue
            l_colDeletes.Add l_cpuGap
        End If
        v_rsLineSchedules.MoveNext
    Loop
    v_rsLineSchedules.Filter = adFilterNone
    
    v_rsLineSchedules.Sort = "WG_NUM ASC, LINE_NUM ASC, START_DATE ASC"
    v_rsLineSchedules.MoveFirst
    AdjustDeletedIntersections v_rsLineSchedules, l_colDeletes
    l_lngWgNum = v_rsLineSchedules!WG_NUM
    l_lngLineNum = v_rsLineSchedules!LINE_NUM
    l_datEndDate = v_rsLineSchedules!END_DATE
    v_rsLineSchedules.MoveNext

    Do Until v_rsLineSchedules.EOF
        AdjustDeletedIntersections v_rsLineSchedules, l_colDeletes
        If v_rsLineSchedules!WG_NUM = l_lngWgNum And _
            v_rsLineSchedules!LINE_NUM = l_lngLineNum Then
            If GapExists(l_datEndDate, v_rsLineSchedules!START_DATE) And _
                    l_datEndDate < v_rsLineSchedules!START_DATE Then
                Set l_cpuGap = GapIntersects(l_lngWgNum, l_lngLineNum, _
                        l_datEndDate, v_rsLineSchedules!START_DATE, v_colWgLineIntervals)
                If Not l_cpuGap Is Nothing Then
                    l_colGaps.Add l_cpuGap
                End If
            End If
            l_datEndDate = v_rsLineSchedules!END_DATE
        Else
            l_lngWgNum = v_rsLineSchedules!WG_NUM
            l_lngLineNum = v_rsLineSchedules!LINE_NUM
            l_datEndDate = v_rsLineSchedules!END_DATE
        End If
        v_rsLineSchedules.MoveNext
    Loop
    
'    Add any remaining deletes to l_colGaps
    For Each l_cpuGap In l_colDeletes
        l_colGaps.Add l_cpuGap
    Next

'    resort by wgNum, lineNum, startDate

    Set DetermineWGLineGaps = l_colGaps
    Set l_cpuGap = Nothing
    Set l_colGaps = Nothing
    v_rsLineSchedules.MoveFirst
    Exit Function

eh:
    Set l_cpuGap = Nothing
    Set l_colGaps = Nothing
    TrsRaiseError m_strCMODULENAME, "DetermineWGLineGaps"
End Function


Private Function AdjustDeletedIntersections( _
            ByVal v_rsLineSchedules As ADODB.Recordset, _
            ByRef r_colDeletes As Collection)
'    Each record processed must be compared to the deletes.
'    The deleted records will be adjusted or deleted based on the intersection.
    Dim l_intIntersection As Integer
    Dim l_cpuDelete As cpuDateInterval
    Dim l_colCopy As Collection
    Dim l_datEndDate As Date
    Dim l_cpuSplit As cpuDateInterval

On Error GoTo eh

    If v_rsLineSchedules.EOF Or r_colDeletes.Count = 0 Then
        Exit Function
    End If
    
    Set l_colCopy = CopyObjects(r_colDeletes)
    
    For Each l_cpuDelete In l_colCopy
        If v_rsLineSchedules!WG_NUM = l_cpuDelete.WorkgroupNumber And _
            v_rsLineSchedules!LINE_NUM = l_cpuDelete.LineNumber Then
            l_intIntersection = DateIntersection( _
                v_rsLineSchedules!START_DATE, v_rsLineSchedules!END_DATE, _
                l_cpuDelete.StartDate, l_cpuDelete.EndDate)
            If l_intIntersection = 1 Then
                RemoveObjectFrom l_cpuDelete, r_colDeletes
            ElseIf l_intIntersection = 2 Then   'split
                l_datEndDate = l_cpuDelete.EndDate
                l_cpuDelete.EndDate = AddDays(-1, v_rsLineSchedules!START_DATE)
                Set l_cpuSplit = New cpuDateInterval
                l_cpuSplit.WorkgroupNumber = l_cpuDelete.WorkgroupNumber
                l_cpuSplit.LineNumber = l_cpuDelete.LineNumber
                l_cpuSplit.StartDate = AddDays(1, v_rsLineSchedules!END_DATE)
                l_cpuSplit.EndDate = l_datEndDate
                r_colDeletes.Add l_cpuSplit
            ElseIf l_intIntersection = 3 Then
                l_cpuDelete.StartDate = AddDays(1, v_rsLineSchedules!END_DATE)
            ElseIf l_intIntersection = 4 Then
                l_cpuDelete.EndDate = AddDays(-1, v_rsLineSchedules!START_DATE)
            End If
        End If
    Next

    Set l_cpuDelete = Nothing
    Set l_colCopy = Nothing
    Set l_cpuSplit = Nothing
    Exit Function

eh:
    Set l_cpuDelete = Nothing
    Set l_colCopy = Nothing
    Set l_cpuSplit = Nothing
    Err.Raise Err.Number
End Function


Private Function GapIntersects( _
                    ByVal v_lngWgNum As Long, _
                    ByVal v_lngLineNum As Long, _
                    ByVal v_datPreviousEndDate As Date, _
                    ByVal v_datNextStartDate As Date, _
                    ByVal v_colWgLineIntervals As Collection) As cpuDateInterval

    Dim l_cpuDateInterval As cpuDateInterval
    Dim l_datGapStartDate As Date
    Dim l_datGapEndDate As Date
    Dim l_intIntersection As Integer
    Dim l_cpuGap As cpuDateInterval
    
On Error GoTo eh

    l_datGapStartDate = AddDays(1, v_datPreviousEndDate)
    l_datGapEndDate = AddDays(-1, v_datNextStartDate)
    
    For Each l_cpuDateInterval In v_colWgLineIntervals
        If l_cpuDateInterval.WorkgroupNumber = v_lngWgNum And _
            l_cpuDateInterval.LineNumber = v_lngLineNum Then
            l_intIntersection = DateIntersection( _
                l_datGapStartDate, l_datGapEndDate, _
                l_cpuDateInterval.StartDate, l_cpuDateInterval.EndDate)
            If l_intIntersection >= 1 And l_intIntersection <= 5 Then
                Set l_cpuGap = New cpuDateInterval
                l_cpuGap.WorkgroupNumber = l_cpuDateInterval.WorkgroupNumber
                l_cpuGap.LineNumber = l_cpuDateInterval.LineNumber
                l_cpuGap.StartDate = l_datGapStartDate
                l_cpuGap.EndDate = l_datGapEndDate
                Exit For
            End If
        End If
    Next

    Set GapIntersects = l_cpuGap

    Set l_cpuDateInterval = Nothing
    Set l_cpuGap = Nothing
    Exit Function

eh:
    Set l_cpuDateInterval = Nothing
    Set l_cpuGap = Nothing
    TrsRaiseError m_strCMODULENAME, "GapIntersects"
End Function


Private Function GapExists( _
                    ByVal v_datPreviousEndDate As Date, _
                    ByVal v_datNextStartDate As Date) As Boolean
' Answer False when v_datPreviousEndDate and v_datNextStartDate are the
' same or contiguous; otherwise True.
    Dim l_lngDateDiff As Long

On Error GoTo eh

    l_lngDateDiff = DateDiff("d", v_datPreviousEndDate, v_datNextStartDate)
    If l_lngDateDiff >= -1 And l_lngDateDiff <= 1 Then
        GapExists = False
    Else
        GapExists = True
    End If
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "GapExists"
End Function


Private Function IsValidWorkgroup( _
            ByVal v_lngWorkgroupNum As Long, _
            ByVal v_datEffDate As Date _
        ) As Boolean
' Jim already validated.  Just ensure same number as module variable.
On Error GoTo eh
    IsValidWorkgroup = (v_lngWorkgroupNum = m_lngWorkGroup)
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "IsValidWorkgroup"
End Function

'Private Function IsValidWorkgroup( _
'            ByVal v_lngWorkgroupNum As Long, _
'            ByVal v_datEffDate As Date _
'        ) As Boolean
'    Dim l_cdWorkgroup As dbWorkGroup.cdWorkGroup
'
'On Error GoTo eh
'    Set l_cdWorkgroup = CreateObject("dbWorkGroup.cdWorkGroup")
'    IsValidWorkgroup = _
'        l_cdWorkgroup.IsValidWorkgroup(v_lngWorkgroupNum, v_datEffDate)
'
'    Set l_cdWorkgroup = Nothing
'    Exit Function
'
'eh:
'    Set l_cdWorkgroup = Nothing
'    TrsRaiseError m_strCMODULENAME, "IsValidWorkgroup"
'End Function


Private Function IsValidJobNumber( _
            ByVal v_lngJobNum As String, _
            ByVal v_datEffDate As Date _
        ) As Boolean
    Dim l_cdCodeTables As dbCodeTables.cdCodeTables

On Error GoTo eh
    If Trim(v_lngJobNum) = "" Then
        IsValidJobNumber = True
        Exit Function
    End If

    Set l_cdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    IsValidJobNumber = _
        l_cdCodeTables.JobNumberIsValid(v_lngJobNum, v_datEffDate)
        
    Set l_cdCodeTables = Nothing
    Exit Function
    
eh:
    GetObjectContext.SetAbort
    TrsRaiseError m_strCMODULENAME, "IsValidJobNumber"
End Function


Private Function IsValidRotationCode( _
            ByVal v_lngRotCode As Long _
        ) As Boolean
    Dim l_cdCodeTables As dbCodeTables.cdCodeTables

On Error GoTo eh
    Set l_cdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    IsValidRotationCode = _
        l_cdCodeTables.RotationCodeIsValid(v_lngRotCode)
    
    Set l_cdCodeTables = Nothing
    Exit Function
    
eh:
    GetObjectContext.SetAbort
    TrsRaiseError m_strCMODULENAME, "IsValidRotationCode"
End Function


Private Function IsValidRotationWeek( _
            ByVal v_lngRotCode As Long, _
            ByVal v_lngRotWeek As Long _
        ) As Boolean
    Dim l_cdCodeTables As dbCodeTables.cdCodeTables

On Error GoTo eh
    Set l_cdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    IsValidRotationWeek = _
        l_cdCodeTables.RotationWeeksIsValid(v_lngRotCode, v_lngRotWeek)
    
    Set l_cdCodeTables = Nothing
    Exit Function
    
eh:
    GetObjectContext.SetAbort
    TrsRaiseError m_strCMODULENAME, "IsValidRotationWeek"
End Function


Private Function IsValidScheduleCode( _
            ByVal v_lngSchedCode As Long _
        ) As Boolean
    Dim l_cdCodeTables As dbCodeTables.cdCodeTables

On Error GoTo eh
    Set l_cdCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    IsValidScheduleCode = _
        l_cdCodeTables.ScheduleCodeIsValid(v_lngSchedCode)
    
    Set l_cdCodeTables = Nothing
    Exit Function
    
eh:
    GetObjectContext.SetAbort
    TrsRaiseError m_strCMODULENAME, "IsValidScheduleCode"
End Function


Private Function CheckForOverlappingLineSchedules( _
            ByRef r_rsLineSchedules As ADODB.Recordset _
            ) As Boolean
'This Subroutine checks for overlapping line schedules in the same work group.
'The function returns true if an overlap exists.
    Dim l_blnOverlapExists As Boolean
    Dim l_setWgLines As cuSet
    Dim l_varWgLine As Variant
    Dim l_varWgLineArray As Variant
    
On Error GoTo eh
    CheckForOverlappingLineSchedules = False
    If r_rsLineSchedules.EOF Then
        Exit Function
    End If
    
    Set l_setWgLines = New cuSet
    Do Until r_rsLineSchedules.EOF
        l_setWgLines.Add _
            Join(Array(r_rsLineSchedules!WG_NUM, r_rsLineSchedules!LINE_NUM), ";")
        r_rsLineSchedules.MoveNext
    Loop

' For each unique workgroup/line,
    For Each l_varWgLine In l_setWgLines
        l_varWgLineArray = Split(l_varWgLine, ";")
        l_blnOverlapExists = _
            OverlappingLineSchedulesExist(l_varWgLineArray(0), l_varWgLineArray(1), r_rsLineSchedules)
        If l_blnOverlapExists Then
            CheckForOverlappingLineSchedules = True
            Exit For
        End If
    Next
  
    r_rsLineSchedules.Filter = adFilterNone
    Set l_setWgLines = Nothing
    Exit Function
    
eh:
  TrsRaiseError m_strCMODULENAME, "CheckForOverlappingLineSchedules"
End Function


Private Function OverlappingLineSchedulesExist( _
            ByVal v_lngWgNum As Long, _
            ByVal v_lngLineNum As Long, _
            ByRef r_rsLineSchedules As ADODB.Recordset _
            ) As Boolean
    Dim l_strFilter As String
    Dim l_rsLineScheduleClone As ADODB.Recordset
    Dim l_intIntersection As Integer
    Dim l_blnOverlapFound As Boolean
    Dim l_strErrorDesc As String
    Dim l_datStart As Date

On Error GoTo eh
    l_strFilter = MDAC25FilterFix("WG_NUM = " & v_lngWgNum & " AND " _
                    & "LINE_NUM = " & v_lngLineNum)
    r_rsLineSchedules.Filter = l_strFilter
    If RecordsExist(r_rsLineSchedules) Then
        Set l_rsLineScheduleClone = r_rsLineSchedules.Clone
        l_rsLineScheduleClone.Filter = l_strFilter
        r_rsLineSchedules.MoveFirst
    Else
        r_rsLineSchedules.Filter = adFilterNone
        OverlappingLineSchedulesExist = False
        Exit Function
    End If
    
    l_blnOverlapFound = False
    Do Until r_rsLineSchedules.EOF
        l_rsLineScheduleClone.MoveFirst
        Do Until l_rsLineScheduleClone.EOF
            If Not IsSameLineSchedule(r_rsLineSchedules, l_rsLineScheduleClone) Then
                l_datStart = r_rsLineSchedules!START_DATE
                l_intIntersection = DateIntersection( _
                    r_rsLineSchedules!START_DATE, r_rsLineSchedules!END_DATE, _
                    l_rsLineScheduleClone!START_DATE, l_rsLineScheduleClone!END_DATE)
                    
                    If l_intIntersection >= 1 And l_intIntersection <= 4 Then
                        'error
                        l_strErrorDesc = "Overlapping assignment exits. " & _
                        "Assignemnt from " & _
                        Format(r_rsLineSchedules!START_DATE, "mm/dd/yyyy") & " to " & Format(r_rsLineSchedules!END_DATE, "mm/dd/yyyy") & _
                        " overrlaps the assignment from " & _
                        Format(l_rsLineScheduleClone!START_DATE, "mm/dd/yyyy") & " to " & Format(l_rsLineScheduleClone!END_DATE, "mm/dd/yyyy")
                        AppendWGSchedLineError m_rsErrors, Nothing, trErrbusValidationError, mstrCBUSINESSERR, l_strErrorDesc, , str(v_lngWgNum), str(v_lngLineNum), l_datStart
                        l_blnOverlapFound = True
                        Exit Do
                    End If
            End If
            l_rsLineScheduleClone.MoveNext
        Loop
        If l_blnOverlapFound Then
            Exit Do
        Else
            r_rsLineSchedules.MoveNext
        End If
    Loop
    
    Set l_rsLineScheduleClone = Nothing
    r_rsLineSchedules.Filter = adFilterNone
    OverlappingLineSchedulesExist = l_blnOverlapFound
    Exit Function
    
eh:
    Set l_rsLineScheduleClone = Nothing
    TrsRaiseError m_strCMODULENAME, "OverlappingLineSchedulesExist"
End Function


Private Function IsSameLineSchedule( _
            ByRef r_rsLineSchedule As ADODB.Recordset, _
            ByRef r_rsLineScheduleClone As ADODB.Recordset) As Boolean

On Error GoTo eh
    If r_rsLineSchedule.Bookmark = r_rsLineScheduleClone.Bookmark Then
        IsSameLineSchedule = True
    Else
        IsSameLineSchedule = False
    End If
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "IsSameLineSchedule"
End Function


Private Function CheckForOverlappingBaseSchedules( _
            ByRef r_rsBaseSchedules As ADODB.Recordset _
            ) As Boolean
'This Subroutine checks for overlapping base schedules in the same work group.
'The function returns true if an overlap exists.
    Dim l_blnOverlapExists As Boolean
    Dim l_setWgNums As cuSet
    Dim l_varWgNum As Variant
    Dim l_varWgLineArray As Variant
    
On Error GoTo eh
    CheckForOverlappingBaseSchedules = False
    If r_rsBaseSchedules.EOF Then
        Exit Function
    End If
    
    Set l_setWgNums = New cuSet
    Do Until r_rsBaseSchedules.EOF
        l_setWgNums.Add r_rsBaseSchedules!WG_NUM.Value
        r_rsBaseSchedules.MoveNext
    Loop

' For each unique workgroup/line,
    For Each l_varWgNum In l_setWgNums
        l_blnOverlapExists = _
            OverlappingBaseSchedulesExist(l_varWgNum, r_rsBaseSchedules)
        If l_blnOverlapExists Then
            CheckForOverlappingBaseSchedules = True
            Exit For
        End If
    Next
  
    r_rsBaseSchedules.Filter = adFilterNone
    Set l_setWgNums = Nothing
    Exit Function
    
eh:
  TrsRaiseError m_strCMODULENAME, "CheckForOverlappingBaseSchedules"
End Function


Private Function OverlappingBaseSchedulesExist( _
            ByVal v_lngWgNum As Long, _
            ByRef r_rsBaseSchedules As ADODB.Recordset _
            ) As Boolean
    Dim l_strFilter As String
    Dim l_rsBaseScheduleClone As ADODB.Recordset
    Dim l_intIntersection As Integer
    Dim l_blnOverlapFound As Boolean
    Dim l_strErrorDesc As String

On Error GoTo eh
    l_strFilter = "WG_NUM = " & v_lngWgNum
    r_rsBaseSchedules.Filter = l_strFilter
    If RecordsExist(r_rsBaseSchedules) Then
        Set l_rsBaseScheduleClone = r_rsBaseSchedules.Clone
        l_rsBaseScheduleClone.Filter = l_strFilter
        r_rsBaseSchedules.MoveFirst
    Else
        r_rsBaseSchedules.Filter = adFilterNone
        OverlappingBaseSchedulesExist = False
        Exit Function
    End If
    
    l_blnOverlapFound = False
    Do Until r_rsBaseSchedules.EOF
        l_rsBaseScheduleClone.MoveFirst
        Do Until l_rsBaseScheduleClone.EOF
            If Not IsSameLineSchedule(r_rsBaseSchedules, l_rsBaseScheduleClone) Then
                l_intIntersection = DateIntersection( _
                    r_rsBaseSchedules!EFF_DATE, r_rsBaseSchedules!END_DATE, _
                    l_rsBaseScheduleClone!EFF_DATE, l_rsBaseScheduleClone!END_DATE)
                    
                    If l_intIntersection >= 1 And l_intIntersection <= 4 Then
                        'error
                        l_strErrorDesc = "Overlapping assignment exits. " & _
                        "Assignemnt from " & _
                        Format(r_rsBaseSchedules!EFF_DATE, "mm/dd/yyyy") & " to " & Format(r_rsBaseSchedules!END_DATE, "mm/dd/yyyy") & _
                        " overrlaps the assignment from " & _
                        Format(l_rsBaseScheduleClone!EFF_DATE, "mm/dd/yyyy") & " to " & Format(l_rsBaseScheduleClone!END_DATE, "mm/dd/yyyy")
                        AppendWGSchedError m_rsErrors, trErrbusValidationError, mstrCBUSINESSERR, l_strErrorDesc
                        l_blnOverlapFound = True
                        Exit Do
                    End If
            End If
            l_rsBaseScheduleClone.MoveNext
        Loop
        If l_blnOverlapFound Then
            Exit Do
        Else
            r_rsBaseSchedules.MoveNext
        End If
    Loop
    
    Set l_rsBaseScheduleClone = Nothing
    r_rsBaseSchedules.Filter = adFilterNone
    OverlappingBaseSchedulesExist = l_blnOverlapFound
    Exit Function
    
eh:
    Set l_rsBaseScheduleClone = Nothing
    TrsRaiseError m_strCMODULENAME, "OverlappingBaseSchedulesExist"
End Function


Private Sub FindLineScheduleRecord( _
                ByVal v_lngWgNum As Long, _
                ByVal v_lngLineNum As Long, _
                ByVal v_datEffDate As Date, _
                ByVal v_strCreatedTS As String, _
                ByRef r_rsLineSchedules As ADODB.Recordset)
' Position r_rsLineSchedules on the first record that matches input.

On Error GoTo eh
    
    If RecordsExist(r_rsLineSchedules) Then
        r_rsLineSchedules.MoveFirst
    Else
        Exit Sub
    End If

    Do Until r_rsLineSchedules.EOF
        If r_rsLineSchedules!WG_NUM = v_lngWgNum And _
            r_rsLineSchedules!LINE_NUM = v_lngLineNum And _
            r_rsLineSchedules.Fields("START_DATE").OriginalValue = v_datEffDate And _
            r_rsLineSchedules!CREATED_TIMESTAMP = v_strCreatedTS Then
            Exit Do
        End If
        r_rsLineSchedules.MoveNext
    Loop
    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "FindLineScheduleRecord"
End Sub


Private Sub FindEffectiveLineScheduleRecord( _
                ByVal v_lngWgNum As Long, _
                ByVal v_lngLineNum As Long, _
                ByVal v_datEffDate As Date, _
                ByRef r_rsLineSchedules As ADODB.Recordset)
' Position r_rsLineSchedules on the first record where v_datEffDate is between
' START_DATE and END_DATE.

On Error GoTo eh

    If RecordsExist(r_rsLineSchedules) Then
'        r_rsLineSchedules.Sort = "WG_NUM DESC, LINE_NUM DESC, START_DATE DESC"
        r_rsLineSchedules.MoveFirst
    Else
        Exit Sub
    End If
    
    Do Until r_rsLineSchedules.EOF
        If r_rsLineSchedules!WG_NUM = v_lngWgNum And _
            r_rsLineSchedules!LINE_NUM = v_lngLineNum And _
            r_rsLineSchedules!START_DATE <= v_datEffDate And _
            r_rsLineSchedules!END_DATE >= v_datEffDate Then
            Exit Do
        End If
        r_rsLineSchedules.MoveNext
    Loop
    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "FindEffectiveLineScheduleRecord"
End Sub


Private Sub FindNextLineScheduleRecord( _
                ByVal v_lngWgNum As Long, _
                ByVal v_lngLineNum As Long, _
                ByVal v_datEffDate As Date, _
                ByRef r_rsLineSchedules As ADODB.Recordset)
' Position r_rsLineSchedules on the first record that is greater than v_datEffDate.

On Error GoTo eh
    If RecordsExist(r_rsLineSchedules) Then
        r_rsLineSchedules.Sort = "WG_NUM ASC, LINE_NUM ASC, START_DATE ASC"
        r_rsLineSchedules.MoveFirst
    Else
        Exit Sub
    End If
    Do Until r_rsLineSchedules.EOF
        If r_rsLineSchedules!WG_NUM = v_lngWgNum And _
            r_rsLineSchedules!LINE_NUM = v_lngLineNum And _
            r_rsLineSchedules!START_DATE > v_datEffDate Then
            Exit Do
        End If
        r_rsLineSchedules.MoveNext
    Loop
    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "FindNextLineScheduleRecord"
End Sub


Private Function AnyLineAssignments( _
                ByVal v_lngWgNum As Long, _
                ByVal v_lngLineNum As Long, _
                ByVal v_datStartDate As Date, _
                ByVal v_datEndDate As Date) As Boolean
    Dim l_strFilter As String
    Dim l_intIntersection As Integer

On Error GoTo eh
    AnyLineAssignments = False

    l_strFilter = MDAC25FilterFix("WG_NUM = " & v_lngWgNum & " AND " _
                    & "LINE_NUM = " & v_lngLineNum)
    m_rsWGDemandLineAssignments.Filter = l_strFilter
    Do Until m_rsWGDemandLineAssignments.EOF
        l_intIntersection = DateIntersection(v_datStartDate, v_datEndDate, _
                m_rsWGDemandLineAssignments!START_DATE, m_rsWGDemandLineAssignments!END_DATE)
        If l_intIntersection >= 1 And l_intIntersection <= 4 Then
            AnyLineAssignments = True
            Exit Do
        End If
        m_rsWGDemandLineAssignments.MoveNext
    Loop
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "AnyLineAssignments"
End Function


Private Function AdjustFutureLineSchedules( _
                    ByVal v_lngLineNum As Long, _
                    ByVal v_datStartDate As Date, _
                    ByVal v_datEndDate As Date, _
                    ByRef r_rsLineSchedules As ADODB.Recordset)
' Check for adjustments to future line schedules.
' If the updated line sched now entirely overlaps another line schedule, delete it to
' avoid the overlap.  If the new line sched ends after a future line sched,
' delete it since it is entirely overlapped and obsolete.
' This assumes record position is moving forward only.
    Dim l_intIntersection As Integer
    Dim l_datOrigStartDate As Date
    Dim l_intOrigRotWeek As Integer

On Error GoTo eh

    Do Until r_rsLineSchedules.EOF
        If r_rsLineSchedules!START_DATE > v_datEndDate Or _
                r_rsLineSchedules!LINE_NUM <> v_lngLineNum Then
            Exit Do
        End If
        l_intIntersection = DateIntersection(v_datStartDate, v_datEndDate, _
                                r_rsLineSchedules!START_DATE, r_rsLineSchedules!END_DATE)
        If l_intIntersection = 3 Then
            l_datOrigStartDate = r_rsLineSchedules!START_DATE
            l_intOrigRotWeek = r_rsLineSchedules!ROT_WEEK
            r_rsLineSchedules!START_DATE = v_datEndDate + 1
            r_rsLineSchedules!LAST_MOD_EMP_NUM = m_lngUserNum
            If r_rsLineSchedules!ROT_CD <> 0 Then
                r_rsLineSchedules!ROT_WEEK = _
                    DetermineNewRotationWeek( _
                        r_rsLineSchedules!START_DATE, _
                        r_rsLineSchedules!ROT_CD, _
                        l_datOrigStartDate, _
                        r_rsLineSchedules!ROT_WEEK)
            End If
            r_rsLineSchedules.Update
            AppendWGLineCascadeDates "C", r_rsLineSchedules!WG_NUM, r_rsLineSchedules!LINE_NUM, _
                    l_datOrigStartDate, v_datEndDate
            If r_rsLineSchedules!ROT_WEEK <> l_intOrigRotWeek Then
                AppendWarningRotationWeekChange _
                    trWarRotWeekAdjusted, _
                    l_intOrigRotWeek, _
                    r_rsLineSchedules.Fields("ROT_WEEK").Value, _
                    r_rsLineSchedules.Fields("WG_NUM").Value, _
                    r_rsLineSchedules.Fields("LINE_NUM").Value, _
                    r_rsLineSchedules.Fields("START_DATE").Value
            End If
            Exit Do
        ElseIf l_intIntersection = 1 Then
            AppendWGLineCascadeDates "C", r_rsLineSchedules!WG_NUM, r_rsLineSchedules!LINE_NUM, _
                    r_rsLineSchedules!START_DATE, r_rsLineSchedules!END_DATE
            r_rsLineSchedules.Delete
            r_rsLineSchedules.MoveNext
        End If
    Loop
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "AdjustFutureLineSchedules"
End Function


Private Function AppendWarningRotationWeekChange( _
                    ByVal v_lngVBErrNum As Long, _
                    ByVal v_intOrigRotWeek As Integer, _
                    ByVal v_intNewRotWeek As Integer, _
                    ByVal v_lngWgNum As Long, _
                    ByVal v_lngLineNum As Long, _
                    ByVal v_datStartDate As Date)

    Dim l_strWarDesc As String
    Dim l_colDescArgs As Collection

On Error GoTo eh

    Set l_colDescArgs = New Collection
    l_colDescArgs.Add CStr(v_intOrigRotWeek)
    l_colDescArgs.Add CStr(v_intNewRotWeek)
    l_colDescArgs.Add Format(v_datStartDate, "mm/dd/yyyy")
    l_colDescArgs.Add CStr(v_lngLineNum)
    l_strWarDesc = TrsGetErrDescription(LongAppErr(v_lngVBErrNum), l_colDescArgs)
    Set l_colDescArgs = Nothing
    AppendBusinessError _
        m_rsErrors, _
        LongAppErr(v_lngVBErrNum), _
        mstrCBUSINESSWAR, _
        l_strWarDesc, _
        m_strCMODULENAME, _
        TrType_WGM_SCHED, _
        , , _
        CDEMANDSCHEDFIELDNAMES, _
        CStr(v_lngWgNum), _
        Format(v_datStartDate, "mm/dd/yyyy"), _
        CStr(v_lngLineNum)

    Exit Function

eh:
    Set l_colDescArgs = Nothing
    TrsRaiseError m_strCMODULENAME, "AppendWarningRotationWeekChange"
End Function


Private Function DetermineNewRotationWeek( _
                    ByVal v_datChargeDate As Date, _
                    ByVal v_intRotationCode As Integer, _
                    ByVal v_datEffDate As Date, _
                    ByVal v_intStartPosition As Integer _
                ) As Integer
' Play out the rotation and answer the adjusted ROT_WEEK, offset from v_datEffDate
' and v_intStartPosition (original ROT_WEEK).
    Dim l_cbsSchedule As busSchedule.cbsSchedule
    Dim l_intNewRotWeek As Integer
    Dim l_intSchedCd As Integer

On Error GoTo eh

    Set l_cbsSchedule = CreateObject("busSchedule.cbsSchedule")
    l_intSchedCd = l_cbsSchedule.GetRotationScheduleCode( _
        v_datChargeDate, v_intRotationCode, v_datEffDate, v_intStartPosition, l_intNewRotWeek)
    Set l_cbsSchedule = Nothing
    
    DetermineNewRotationWeek = l_intNewRotWeek
    Exit Function
    
eh:
    Set l_cbsSchedule = Nothing
    TrsRaiseError m_strCMODULENAME, "DetermineNewRotationWeek"
End Function


'Private Function DateIntersection( _
'    ByVal v_datStartTime1 As Date, ByVal v_datEndTime1 As Date, _
'    ByVal v_datStartTime2 As Date, ByVal v_datEndTime2 As Date) _
'        As Integer
' 5/14/02 RSG - Moved to mDateTime.bas
'
'
'Private Function AddDays( _
'                ByVal v_lngNumber As Long, _
'                ByVal v_datDate As Date) As Date
' 5/14/02 RSG - Moved to mDateTime.bas
'End Function
'

Private Sub AppendWGSchedLineError( _
            ByRef r_rsErrors As ADODB.Recordset, _
            ByRef r_rsScheduleRequests As ADODB.Recordset, _
            ByVal v_lngErrorNumber As Long, _
            ByVal v_strErrorType As String, _
            ByVal v_strErrorDescription As String, _
            Optional ByVal v_strErrorSource As String, _
            Optional ByVal v_strWGNum As String, _
            Optional ByVal v_strLineNum As String, _
            Optional ByVal v_datStartDate As Date)

' May make some of these arguments module variables?

    Dim l_strErrorSource As String
    Dim l_strWG_Num As String
    Dim l_strLineNum As String
    Dim l_datStartDate As Date
    Dim l_strStartDate As String
    
On Error GoTo eh
    
    l_strErrorSource = IIf(v_strErrorSource = "", m_strCMODULENAME, v_strErrorSource)

    If v_strWGNum = "" Then
        If IsNull(r_rsScheduleRequests!WG_NUM) Then
            If IsNull(r_rsScheduleRequests!KEY_WG_NUM) Then
                l_strWG_Num = "0"
            Else
                l_strWG_Num = r_rsScheduleRequests!KEY_WG_NUM
            End If
        Else
            l_strWG_Num = r_rsScheduleRequests!WG_NUM
        End If
    Else
        l_strWG_Num = v_strWGNum
    End If
    
    If v_strLineNum = "" Then
        If IsNull(r_rsScheduleRequests!LINE_NUM) Then
            If IsNull(r_rsScheduleRequests!KEY_LINE_NUM) Then
                l_strLineNum = "0"
            Else
                l_strLineNum = r_rsScheduleRequests!KEY_LINE_NUM
            End If
        Else
            l_strLineNum = r_rsScheduleRequests!LINE_NUM
        End If
    Else
        l_strLineNum = v_strLineNum
    End If
        
    If v_datStartDate = vbEmpty Then
        If IsNull(r_rsScheduleRequests!START_DATE) Then
            l_strStartDate = ""
'            If IsNull(r_rsScheduleRequests!KEY_START_DATE) Then
'                l_datStartDate = 0
'            Else
'                l_datStartDate = r_rsScheduleRequests!KEY_START_DATE
'            End If
        Else
            l_datStartDate = r_rsScheduleRequests!START_DATE
            l_strStartDate = Format(l_datStartDate, "mm/dd/yyyy")
        End If
    Else
        l_datStartDate = v_datStartDate
        l_strStartDate = Format(l_datStartDate, "mm/dd/yyyy")
    End If
   
    AppendBusinessError _
        r_rsErrors, _
        v_lngErrorNumber, _
        v_strErrorType, _
        v_strErrorDescription, _
        l_strErrorSource, _
        TrType_WGM_SCHED, _
        , _
        m_strErrorField, _
        CDEMANDSCHEDFIELDNAMES, _
        l_strWG_Num, _
        l_strStartDate, _
        l_strLineNum

    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "AppendWGSchedLineError", _
        trErrbusAppendEntryError, True, , , True
End Sub


Private Sub AppendWGSchedError( _
            ByRef r_rsErrors As ADODB.Recordset, _
            ByVal v_lngErrorNumber As Long, _
            ByVal v_strErrorType As String, _
            Optional ByVal v_strErrorDescription As String, _
            Optional ByVal v_strErrorSource As String, _
            Optional ByRef v_colDescriptionArgs As Collection, _
            Optional ByVal v_strWGNum As String, _
            Optional ByVal v_datStartDate As Date)
    Dim l_strErrorSource As String
    Dim l_strErrDesc As String
    Dim l_strWG_Num As String
    Dim l_strLineNum As String
    Dim l_datStartDate As Date

On Error GoTo eh
    l_strErrorSource = IIf(v_strErrorSource = "", m_strCMODULENAME, v_strErrorSource)

    If v_strErrorDescription = "" Then
        l_strErrDesc = TrsGetErrDescription(LongAppErr(v_lngErrorNumber), v_colDescriptionArgs)
    Else
        l_strErrDesc = v_strErrorDescription
    End If
    
'    If IsNull(r_rsScheduleRequests!WG_NUM) Then
'        If IsNull(r_rsScheduleRequests!KEY_WG_NUM) Then
'            l_strWG_Num = "0"
'        Else
'            l_strWG_Num = r_rsScheduleRequests!KEY_WG_NUM
'        End If
'    Else
'        l_strWG_Num = r_rsScheduleRequests!WG_NUM
'    End If
    
'''''    If IsNull(r_rsScheduleRequests!LINE_NUM) Then
'''''        If IsNull(r_rsScheduleRequests!KEY_LINE_NUM) Then
'''''            l_strLineNum = "0"
'''''        Else
'''''            l_strLineNum = r_rsScheduleRequests!KEY_LINE_NUM
'''''        End If
'''''    Else
'''''        l_strLineNum = r_rsScheduleRequests!LINE_NUM
'''''    End If

'    If IsNull(r_rsScheduleRequests!START_DATE) Then
'        If IsNull(r_rsScheduleRequests!KEY_START_DATE) Then
'            l_datStartDate = 0
'        Else
'            l_datStartDate = r_rsScheduleRequests!KEY_START_DATE
'        End If
'    Else
'        l_datStartDate = r_rsScheduleRequests!START_DATE
'    End If


    AppendBusinessError _
        r_rsErrors, _
        v_lngErrorNumber, _
        v_strErrorType, _
        l_strErrDesc, _
        l_strErrorSource, _
        TrType_WGM_SCHED, _
        , _
        m_strErrorField, _
        "WG_NUM;START_DATE", _
        v_strWGNum, _
        Format(v_datStartDate, "mm/dd/yyyy")
        

    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "AppendWGSchedError", _
        trErrbusAppendEntryError, True, , , True
End Sub


'Private Function FormatSchedBRV( _
'            ByVal v_lngBRVType As TrTypeBRV, _
'            ByVal v_lngWGNum As Long, _
'            ByVal v_lngLineNum As Long, _
'            ByVal v_datStartDate As Date) As String
'    Dim vBRVitems As Variant
'
'    PreserveErrorContext
'    vBRVitems = Array(v_lngWGNum, v_lngLineNum, v_datStartDate)
'    FormatSchedBRV = AppendBRVitems(v_lngBRVType, vBRVitems)
'    RestoreErrorContext
'
'End Function


Private Sub ClearModuleVariables()

'    m_strUserId = ""
    m_strErrorField = ""


'release these references?  or were they acquired by higher level?
'    Set m_rsWGDemandLineSchedules = Nothing
'    Set m_rsErrors = Nothing
'
'    Set m_rsWGDemandLineAssignments = Nothing

End Sub

Private Function IsADemandWorkGroup( _
    ByVal v_rsWorkGroupAssignments As ADODB.Recordset _
    ) As Boolean
    
    If RecordsExist(m_rsWorkGroupAttributes) Then
        If m_rsWorkGroupAttributes!DEMAND_IND = "Y" Then
            IsADemandWorkGroup = True
        Else
            IsADemandWorkGroup = False
        End If
        Exit Function
    End If
    
    If RecordsExist(v_rsWorkGroupAssignments) Then
        If v_rsWorkGroupAssignments!DEMAND_IND = "Y" Then
            IsADemandWorkGroup = True
        Else
            IsADemandWorkGroup = False
        End If
    Else
        IsADemandWorkGroup = False
    End If

End Function



'
' Start ProcessWorkgroupBaseSchedules
'
Private Function ProcessWorkgroupBaseSchedules( _
            ByRef r_rsScheduleRequests As ADODB.Recordset _
        ) As Boolean
' 1. prevalidate input
' 2. retrieve db data;
'    existing wg_line_sched rows - Jim will have these in m_rsWGDemandLineSchedules
'    existing wg_base_sched rows - Jim will have these in m_rsWGBaseSched (NotDemand)
' 3. process each request
' 4. post validate
' 5. clean up module level variables.
' 6. save results - Jim will pass updated recordset to save with other updates
'
    Dim l_blnReturn As Boolean

On Error GoTo eh

    If RecordsExist(r_rsScheduleRequests) Then
        'Continue
    Else
        ProcessWorkgroupBaseSchedules = True
        Exit Function
    End If
    
    l_blnReturn = PreValidateBaseScheduleRequests(r_rsScheduleRequests)
    If Not l_blnReturn Then
        ProcessWorkgroupBaseSchedules = False
        Exit Function
    Else
        r_rsScheduleRequests.MoveFirst
    End If

    ProcessAllBaseScheduleRequests "D", r_rsScheduleRequests
    ProcessAllBaseScheduleRequests "I", r_rsScheduleRequests
    ProcessAllBaseScheduleRequests "C", r_rsScheduleRequests

    ProcessWorkgroupBaseSchedules = _
        PostValidateBaseScheduleUpdates(m_rsWGBaseSched, r_rsScheduleRequests)
        
        
'    ClearModuleVariables
    
' Cascade preparation
'    ConsolidateCascadeIntervals
    
    Exit Function
    
eh:
'    ClearModuleVariables
    TrsRaiseError m_strCMODULENAME, "ProcessWorkgroupBaseSchedules"
End Function


Private Sub ProcessAllBaseScheduleRequests( _
                ByVal v_strRecordStatus As String, _
                ByRef r_rsScheduleRequests As ADODB.Recordset)
    Dim l_strWGNum As String
    Dim l_datEffDate As Date
    Dim l_strCreatedTS As String
    Dim l_colErrArgs As Collection

On Error GoTo eh
    r_rsScheduleRequests.Filter = "RECORD_STATUS = '" & v_strRecordStatus & "'"
    Do Until r_rsScheduleRequests.EOF
        l_strWGNum = r_rsScheduleRequests!KEY_WG_NUM
        l_datEffDate = r_rsScheduleRequests!KEY_START_DATE
        If Not IsNull(r_rsScheduleRequests!CREATED_TIMESTAMP) Then
            l_strCreatedTS = r_rsScheduleRequests!CREATED_TIMESTAMP
        Else
            l_strCreatedTS = ""
        End If
        Select Case v_strRecordStatus
            Case "D"
                DeleteBaseSchedule r_rsScheduleRequests, m_rsWGBaseSched
            Case "I"
                InsertBaseSchedule r_rsScheduleRequests, m_rsWGBaseSched
            Case "C"
                UpdateBaseSchedule r_rsScheduleRequests, m_rsWGBaseSched
        End Select
        r_rsScheduleRequests.MoveNext
    Loop
    r_rsScheduleRequests.Filter = adFilterNone
    Exit Sub
    
eh:
    Set l_colErrArgs = New Collection
    l_colErrArgs.Add "RECORD_STATUS = " & v_strRecordStatus
    l_colErrArgs.Add "KEY_WG_NUM = " & l_strWGNum
    l_colErrArgs.Add "KEY_START_DATE = " & Format(l_datEffDate, "mm/dd/yyyy")
    l_colErrArgs.Add "CREATED_TIMESTAMP = " & l_strCreatedTS
    TrsRaiseError m_strCMODULENAME, "ProcessAllBaseScheduleRequests", , , , l_colErrArgs, True
End Sub


Private Function PreValidateBaseScheduleRequests( _
                ByRef r_rsScheduleRequests As ADODB.Recordset _
            ) As Boolean
    Dim l_blnIsValid As Boolean
    Dim l_strRecStatus As String
    Dim l_strReqLineInd As String
    Dim l_rsSchedReqClone As ADODB.Recordset
    Dim l_strFilter As String
    
On Error GoTo eh
    
    l_blnIsValid = True
    
    If Not RecordsExist(r_rsScheduleRequests) Then
'JIMB  This statement will never be true, so why is it here
        AppendWGSchedError m_rsErrors, trErrbusValidationError, mstrCAPPLICATIONERR, "Empty request"
        l_blnIsValid = False
        GoTo xit
    Else
        r_rsScheduleRequests.MoveFirst
    End If
    
    Do Until r_rsScheduleRequests.EOF
        l_strRecStatus = Trim(UCase(r_rsScheduleRequests!RECORD_STATUS))
        Select Case l_strRecStatus
            Case "I", "C", "D"
                l_blnIsValid = ValidateBaseKeyFields(r_rsScheduleRequests)
            Case Else
                AppendWGSchedLineError m_rsErrors, r_rsScheduleRequests, trErrbusValidationError, mstrCBUSINESSERR, "Invalid Record Status on the Request"
                l_blnIsValid = False
        End Select
        
        'If there are errors in the key fields, just get out now.
        If Not l_blnIsValid Then
            GoTo xit
        End If
        
        'ROT_CD and SCHED_CD are mutually exclusive
        If l_strRecStatus = "I" Or l_strRecStatus = "C" Then
            If IsMissingValue(r_rsScheduleRequests!ROT_CD) And _
                IsMissingValue(r_rsScheduleRequests!SCHED_CD) Then
                AppendWGSchedLineError m_rsErrors, r_rsScheduleRequests, _
                    trErrbusValidationError, mstrCBUSINESSERR, _
                    "Schedule is missing, please enter a Schedule code or Rotation code"
                l_blnIsValid = False
            ElseIf Not IsMissingValue(r_rsScheduleRequests!ROT_CD) And _
                Not IsMissingValue(r_rsScheduleRequests!SCHED_CD) Then
                AppendWGSchedLineError m_rsErrors, r_rsScheduleRequests, _
                    trErrbusValidationError, mstrCBUSINESSERR, _
                    "You can not enter a Schedule Code and a rotation code. " & _
                    "Please enter the correct schedule."
                l_blnIsValid = False
            End If
        End If
        
'        l_strReqLineInd = Trim(UCase(r_rsScheduleRequests!REQ_LINE_IND))
'        If Not IsYorN(l_strReqLineInd) Then
'            AppendWGSchedLineError m_rsErrors, r_rsScheduleRequests, _
'                trErrbusValidationError, mstrCBUSINESSERR, _
'                "Required indicator is not Y or N"
'            l_blnIsValid = False
'        End If
               
        r_rsScheduleRequests.MoveNext
    Loop
      
    'If there are errors, just get out now.
    If Not l_blnIsValid Then
        GoTo xit
    End If
    
    ' Finally, check for any duplicate inserts or deletes
    r_rsScheduleRequests.MoveFirst
    Set l_rsSchedReqClone = r_rsScheduleRequests.Clone
    Do Until r_rsScheduleRequests.EOF
        If r_rsScheduleRequests!RECORD_STATUS = "D" Or _
            r_rsScheduleRequests!RECORD_STATUS = "I" Then
            l_rsSchedReqClone.Filter = 0
            l_strFilter = "KEY_WG_NUM = " & r_rsScheduleRequests!KEY_WG_NUM & " AND " & _
                "KEY_START_DATE = " & "#" & Format(r_rsScheduleRequests!KEY_START_DATE, "MM/DD/YYYY") & "#" & " AND " & _
                " RECORD_STATUS = " & "'" & r_rsScheduleRequests!RECORD_STATUS & "'"
            l_strFilter = MDAC25FilterFix(l_strFilter)
            l_rsSchedReqClone.Filter = l_strFilter
            If l_rsSchedReqClone.RecordCount > 1 Then
                AppendWGSchedLineError m_rsErrors, r_rsScheduleRequests, _
                    trErrbusValidationError, mstrCBUSINESSERR, _
                    "Duplicate request records found", , _
                    l_rsSchedReqClone!KEY_WG_NUM, _
                    l_rsSchedReqClone!KEY_START_DATE
                l_blnIsValid = False
            End If
        End If
        r_rsScheduleRequests.MoveNext
    Loop
      
xit:
    PreValidateBaseScheduleRequests = l_blnIsValid
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "PreValidateBaseScheduleRequests"
End Function


Private Function ValidateBaseKeyFields( _
            r_rsScheduleRequests As ADODB.Recordset) As Boolean
'Answer True when valid, otherwise False.
    Dim l_blnIsValid As Boolean
    Dim l_strWGNum As String
    Dim l_strLineNum As String
    Dim l_datEffDate As Date
    Dim l_strCreatedTS As String

On Error GoTo eh

    l_blnIsValid = True
    If Not IsNull(r_rsScheduleRequests!KEY_WG_NUM) Then
        l_strWGNum = r_rsScheduleRequests!KEY_WG_NUM
' Jim will have already updated KEY_WG_NUM and WG_NUM with the new assigned number.
'        If (r_rsScheduleRequests!RECORD_STATUS = "I" And _
'                r_rsScheduleRequests!KEY_WG_NUM <> -1) Or _
'            r_rsScheduleRequests!KEY_WG_NUM < 1 Then
        If r_rsScheduleRequests!KEY_WG_NUM < 1 Then
            AppendWGSchedLineError m_rsErrors, r_rsScheduleRequests, _
                trErrbusValidationError, mstrCBUSINESSERR, _
                "KEY_WG_NUM must > 0", , _
                l_strWGNum
            l_blnIsValid = False
        End If
    Else
        l_strWGNum = "0"
        AppendWGSchedLineError m_rsErrors, r_rsScheduleRequests, _
            trErrbusValidationError, mstrCBUSINESSERR, _
            "KEY_WG_NUM cannot be Null", , l_strWGNum
        l_blnIsValid = False
    End If

'    If Not IsNull(r_rsScheduleRequests!KEY_LINE_NUM) Then
'        l_strLineNum = r_rsScheduleRequests!KEY_LINE_NUM
'        If r_rsScheduleRequests!KEY_LINE_NUM < 1 Then
'            AppendWGSchedLineError m_rsErrors, r_rsScheduleRequests, _
'                trErrbusValidationError, mstrCBUSINESSERR, _
'                "KEY_LINE_NUM must > 0", , _
'                l_strWGNum, l_strLineNum
'            l_blnIsValid = False
'        End If
'    Else
'        l_strLineNum = "0"
'        AppendWGSchedLineError m_rsErrors, r_rsScheduleRequests, _
'            trErrbusValidationError, mstrCBUSINESSERR, _
'            "KEY_LINE_NUM cannot be Null", , l_strWGNum, l_strLineNum
'        l_blnIsValid = False
'    End If

    If Not IsNull(r_rsScheduleRequests!KEY_START_DATE.Value) Then
        If Not IsDateValid(r_rsScheduleRequests!KEY_START_DATE.Value) Then
            AppendWGSchedLineError m_rsErrors, r_rsScheduleRequests, _
                trErrbusValidationError, mstrCBUSINESSERR, _
                "KEY_START_DATE is invalid", , _
                l_strWGNum, l_strLineNum, r_rsScheduleRequests!KEY_START_DATE.Value
            l_blnIsValid = False
        End If
    Else
        If Not IsNull(r_rsScheduleRequests!START_DATE.Value) Then
            AppendWGSchedLineError m_rsErrors, r_rsScheduleRequests, _
                trErrbusValidationError, mstrCBUSINESSERR, _
                "KEY_START_DATE cannot be Null", , l_strWGNum, l_strLineNum, 0
            l_blnIsValid = False
        Else
            AppendWGSchedLineError m_rsErrors, r_rsScheduleRequests, _
                trErrbusValidationError, mstrCBUSINESSERR, _
                "Please enter a start date for the schedule", , l_strWGNum, l_strLineNum, 0
            l_blnIsValid = False
        End If
    End If
    
    If Not IsNull(r_rsScheduleRequests!KEY_WG_NUM) And _
        Not IsNull(r_rsScheduleRequests!WG_NUM) Then
        l_strWGNum = r_rsScheduleRequests!WG_NUM
        If r_rsScheduleRequests!KEY_WG_NUM <> r_rsScheduleRequests!WG_NUM Then
            AppendWGSchedLineError m_rsErrors, r_rsScheduleRequests, _
                trErrbusValidationError, mstrCBUSINESSERR, _
                "WG_NUM must = KEY_WG_NUM", , _
                l_strWGNum
            l_blnIsValid = False
        End If
    ElseIf IsNull(r_rsScheduleRequests!WG_NUM) Then
        l_strWGNum = "0"
        AppendWGSchedLineError m_rsErrors, r_rsScheduleRequests, _
            trErrbusValidationError, mstrCBUSINESSERR, _
            "WG_NUM cannot be Null", , _
            l_strWGNum
        l_blnIsValid = False
    End If
    
'    If Not IsNull(r_rsScheduleRequests!KEY_LINE_NUM) And _
'        Not IsNull(r_rsScheduleRequests!LINE_NUM) Then
'        l_strLineNum = r_rsScheduleRequests!LINE_NUM
'        If r_rsScheduleRequests!KEY_LINE_NUM <> r_rsScheduleRequests!LINE_NUM Then
'            AppendWGSchedLineError m_rsErrors, r_rsScheduleRequests, _
'                trErrbusValidationError, mstrCBUSINESSERR, _
'                "LINE_NUM must = KEY_LINE_NUM", , _
'                l_strWGNum, l_strLineNum
'            l_blnIsValid = False
'        End If
'    ElseIf IsNull(r_rsScheduleRequests!LINE_NUM) Then
'        l_strLineNum = "0"
'        AppendWGSchedLineError m_rsErrors, r_rsScheduleRequests, _
'            trErrbusValidationError, mstrCBUSINESSERR, _
'            "LINE_NUM cannot be Null", , _
'            l_strWGNum, l_strLineNum
'        l_blnIsValid = False
'    End If
    
    If Not IsNull(r_rsScheduleRequests!KEY_START_DATE) And _
        Not IsNull(r_rsScheduleRequests!START_DATE) Then
        l_datEffDate = r_rsScheduleRequests!START_DATE
' START_DATE may change.
'        If r_rsScheduleRequests!KEY_START_DATE <> r_rsScheduleRequests!START_DATE Then
'            AppendWGSchedLineError m_rsErrors, r_rsScheduleRequests, _
'                trErrbusValidationError, mstrCBUSINESSERR, _
'                "START_DATE must = KEY_START_DATE", , _
'                l_strWGNum, l_strLineNum, l_datEffDate
'            l_blnIsValid = False
'        ElseIf r_rsScheduleRequests!KEY_START_DATE < Date Then
'
' TEST ONLY
'
'            AppendWGSchedLineError m_rsErrors, r_rsScheduleRequests, _
'                trErrbusValidationError, mstrCBUSINESSERR, _
'                "EFF_DATE must >= current date", , _
'                l_strWGNum, l_strLineNum, l_datEffDate
'            l_blnIsValid = False
'        End If
    ElseIf IsNull(r_rsScheduleRequests!START_DATE) Then
        l_datEffDate = 0
        If Not IsNull(r_rsScheduleRequests!KEY_START_DATE.Value) Then
            AppendWGSchedLineError m_rsErrors, r_rsScheduleRequests, _
                trErrbusValidationError, mstrCBUSINESSERR, _
                "Please enter a start date for the schedule", , _
                l_strWGNum, l_strLineNum, l_datEffDate
            l_blnIsValid = False
        End If
    End If
    
    If IsNull(r_rsScheduleRequests!CREATED_TIMESTAMP) Then
        If r_rsScheduleRequests!RECORD_STATUS <> "I" Then
            l_strCreatedTS = ""
            AppendWGSchedLineError m_rsErrors, r_rsScheduleRequests, _
                trErrbusValidationError, mstrCBUSINESSERR, _
                "CREATED TIMESTAMP cannot be Null", , _
                l_strWGNum, l_strLineNum, l_datEffDate
            l_blnIsValid = False
        End If
    ElseIf Not IsDB2Timestamp(r_rsScheduleRequests!CREATED_TIMESTAMP) Then
        If r_rsScheduleRequests!RECORD_STATUS = "I" And _
            Trim(r_rsScheduleRequests!CREATED_TIMESTAMP) = "" Then
            'OK
        Else
            l_strCreatedTS = ""
            AppendWGSchedLineError m_rsErrors, r_rsScheduleRequests, _
                trErrbusValidationError, mstrCBUSINESSERR, _
                "CREATED TIMESTAMP is not valid", , _
                l_strWGNum, l_strLineNum, l_datEffDate
            l_blnIsValid = False
        End If
    End If
    
    
    ValidateBaseKeyFields = l_blnIsValid
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "ValidateBaseKeyFields"
End Function

Private Function PostValidateBaseScheduleUpdates( _
                ByRef r_rsBaseSchedules As ADODB.Recordset, _
                ByRef r_rsScheduleRequests As ADODB.Recordset _
            ) As Boolean
' Return True when all changes are valid; otherwise return False (invalid
' or no changes recorded.)
    Dim l_blnIsValid As Boolean
    Dim l_intUpdateCount As Integer
    Dim l_datEffStart As Date
    
On Error GoTo eh
    
    l_blnIsValid = True
    
    If Not RecordsExist(r_rsBaseSchedules) Then
        AppendWGSchedError m_rsErrors, _
            trErrbusValidationError, mstrCBUSINESSERR, _
            "Base schedule does not exist."
        PostValidateBaseScheduleUpdates = False
        Exit Function
    Else
        r_rsBaseSchedules.MoveFirst
    End If
    
'only validate new and modified records, if deleted, just count it.
    r_rsBaseSchedules.Filter = adFilterPendingRecords
    Do Until r_rsBaseSchedules.EOF
        Select Case r_rsBaseSchedules.Status
            Case adRecDeleted
                l_intUpdateCount = l_intUpdateCount + 1
            Case adRecNew, adRecModified
                l_intUpdateCount = l_intUpdateCount + 1
                l_datEffStart = r_rsBaseSchedules!EFF_DATE
                If Not IsValidWorkgroup(r_rsBaseSchedules!WG_NUM, r_rsBaseSchedules!EFF_DATE) Then
                    l_blnIsValid = False
                    AppendWGSchedLineError m_rsErrors, r_rsBaseSchedules, _
                        trErrbusValidationError, mstrCBUSINESSERR, _
                        "Workgroup number is not valid on: " & Format(r_rsBaseSchedules!EFF_DATE, "mm/dd/yyyy") _
                        , , , "0", l_datEffStart
                End If
                
                If r_rsBaseSchedules!SCHED_CD = 0 Then
                    If r_rsBaseSchedules!ROT_CD = 0 Then
                        l_blnIsValid = False    'mutex - one must be provided
                        AppendWGSchedLineError m_rsErrors, r_rsBaseSchedules, _
                            trErrbusValidationError, mstrCBUSINESSERR, _
                            "Schedule or Rotation code are required" _
                            , , , "0", l_datEffStart
                    Else
                        If Not IsValidRotationCode(r_rsBaseSchedules!ROT_CD) Then
                            l_blnIsValid = False
                            AppendWGSchedLineError m_rsErrors, r_rsBaseSchedules, _
                                trErrbusValidationError, mstrCBUSINESSERR, _
                                "Rotation code is not valid" _
                                , , , "0", l_datEffStart
                        End If
                        If Not IsValidRotationWeek(r_rsBaseSchedules!ROT_CD, r_rsBaseSchedules!ROT_WEEK) Then
                            l_blnIsValid = False
                            AppendWGSchedLineError m_rsErrors, r_rsBaseSchedules, _
                                trErrbusValidationError, mstrCBUSINESSERR, _
                                "Rotation week is not valid" _
                                , , , "0", l_datEffStart
                        End If
                    End If
                Else
                    If r_rsBaseSchedules!ROT_CD = 0 Then
                        If Not (r_rsBaseSchedules!ROT_WEEK = 0) Then
                            l_blnIsValid = False
                            AppendWGSchedLineError m_rsErrors, r_rsBaseSchedules, _
                                trErrbusValidationError, mstrCBUSINESSERR, _
                                "Rotation week is not valid without Rotation code" _
                                , , , "0", l_datEffStart
                        End If
                        If Not IsValidScheduleCode(r_rsBaseSchedules!SCHED_CD) Then
                            l_blnIsValid = False
                            AppendWGSchedLineError m_rsErrors, r_rsBaseSchedules, _
                                trErrbusValidationError, mstrCBUSINESSERR, _
                                "Schedule code is not valid" _
                                , , , "0", l_datEffStart
                        End If
                    Else
                        l_blnIsValid = False    'mutex - both cannot be provided
                        AppendWGSchedLineError m_rsErrors, r_rsBaseSchedules, _
                            trErrbusValidationError, mstrCBUSINESSERR, _
                            "Schedule or Rotation code are required, not both" _
                            , , , "0", l_datEffStart
                    End If
                End If
        End Select
        r_rsBaseSchedules.MoveNext
    Loop
    
    r_rsBaseSchedules.Filter = adFilterNone
    If CheckForOverlappingBaseSchedules(r_rsBaseSchedules) Then
        l_blnIsValid = False
    End If
    
    If CheckBaseSchedulesEffStartEnd(r_rsBaseSchedules, m_rsWorkGroupAttributes) Then
        l_blnIsValid = False
    End If
    
    If l_intUpdateCount = 0 Then
        AppendWGSchedError m_rsErrors, _
            trErrbusValidationError, mstrCBUSINESSWAR, _
            "No base schedule updates were recorded."
        l_blnIsValid = False
    End If
        
    PostValidateBaseScheduleUpdates = l_blnIsValid
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "PostValidateBaseScheduleUpdates"
End Function


Private Sub DeleteBaseSchedule( _
            ByRef r_rsScheduleRequests As ADODB.Recordset, _
            ByRef r_rsBaseSchedules As ADODB.Recordset)
    Dim l_lngWgNum As Long
    Dim l_datEffDate As Date
    Dim l_strCreatedTS As String
    Dim l_datEndDate As Date
    Dim l_datNewPrevEndDate As Date
    Dim l_varBookmark As Variant
    Dim l_datOldPreviousEffDate As Date
    Dim l_datOldNextDate As Date

On Error GoTo eh
    l_lngWgNum = r_rsScheduleRequests!KEY_WG_NUM
    l_datEffDate = r_rsScheduleRequests!KEY_START_DATE
    l_strCreatedTS = r_rsScheduleRequests!CREATED_TIMESTAMP
    
' 1) Determine if old effective interval exists. i.e. an interval effective
' previous to the deleted interval.  If so, save the bookmark.
    FindPreviousBaseScheduleRecord l_lngWgNum, r_rsScheduleRequests!KEY_START_DATE, r_rsBaseSchedules
    If Not r_rsBaseSchedules.EOF Then
        l_datOldPreviousEffDate = r_rsBaseSchedules!EFF_DATE
        l_varBookmark = r_rsBaseSchedules.Bookmark
    End If
    
    FindBaseScheduleRecord _
        l_lngWgNum, l_datEffDate, l_strCreatedTS, r_rsBaseSchedules
    If r_rsBaseSchedules.EOF Then
        Err.Raise trErrbusInvalidDeleteEntryNotFound
    Else
        l_datEndDate = r_rsBaseSchedules!END_DATE
        r_rsBaseSchedules.Delete
    End If
    
' 4) End the old effective interval.
' If old interval exists, determine if the END_DATE needs to be adjusted.
' Scan forward and if next record is found, set the END_DATE to day before its start.
    If Not IsEmpty(l_varBookmark) Then
        FindNextBaseScheduleRecord l_lngWgNum, l_datOldPreviousEffDate, r_rsBaseSchedules
        If Not r_rsBaseSchedules.EOF Then
            l_datOldNextDate = r_rsBaseSchedules!EFF_DATE
            r_rsBaseSchedules.Bookmark = l_varBookmark
            r_rsBaseSchedules!END_DATE = l_datOldNextDate - 1
        Else
            r_rsBaseSchedules.Bookmark = l_varBookmark
            r_rsBaseSchedules!END_DATE = #12/31/9999#
        End If
        r_rsBaseSchedules!LAST_MOD_EMP_NUM = m_lngUserNum
        r_rsBaseSchedules.Update
    End If
        
    AppendWGLineCascadeDates "D", l_lngWgNum, 0, l_datEffDate, l_datEndDate
    
    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "DeleteBaseSchedule"
End Sub


Private Sub InsertBaseSchedule( _
            ByRef r_rsScheduleRequests As ADODB.Recordset, _
            ByRef r_rsBaseSchedules As ADODB.Recordset)
' 7/22/02 RSG
' No need to adjust ROT_WEEK for any base schedule change because
' there are not implicit changes to START_DATE, nor splits, because there
' are no START/END date intervals, only EFF_DATE which supercedes the
' previous interval.  Therefore all START_DATE changes are explicit and
' there are no splits.  Deletes will only attempt to adjust the END_DATE
' of previous interval.
    Dim l_lngWgNum As Long
    Dim l_datStartDate As Date
    Dim l_datEndDate As Date

On Error GoTo eh
    l_datEndDate = 0
    r_rsBaseSchedules.Sort = "WG_NUM ASC, EFF_DATE ASC"
    l_lngWgNum = r_rsScheduleRequests!WG_NUM
    l_datStartDate = r_rsScheduleRequests!START_DATE

' If old effective line schedule interval exists,
'  set end date to day prior to this start date.
    FindEffectiveBaseScheduleRecord l_lngWgNum, l_datStartDate, r_rsBaseSchedules
    If Not r_rsBaseSchedules.EOF Then
        l_datEndDate = r_rsBaseSchedules!END_DATE
        If r_rsBaseSchedules!EFF_DATE = l_datStartDate Then
            AppendWGLineCascadeDates "D", r_rsBaseSchedules!WG_NUM, 0, _
                r_rsBaseSchedules!EFF_DATE, r_rsBaseSchedules!END_DATE
            r_rsBaseSchedules.Delete
        Else
            r_rsBaseSchedules!END_DATE = l_datStartDate - 1
            r_rsBaseSchedules!LAST_MOD_EMP_NUM = m_lngUserNum
            r_rsBaseSchedules.Update
        End If
    Else
' If there is no previous effective interval, scan forward for the next interval.
        FindNextBaseScheduleRecord l_lngWgNum, l_datStartDate, r_rsBaseSchedules
        If r_rsBaseSchedules.EOF Then
            l_datEndDate = #12/31/9999#
        Else
            l_datEndDate = r_rsBaseSchedules!EFF_DATE - 1
        End If
    End If

' Add the new line schedule.
    r_rsBaseSchedules.AddNew
    r_rsBaseSchedules!WG_NUM = r_rsScheduleRequests!WG_NUM
    r_rsBaseSchedules!EFF_DATE = l_datStartDate
    r_rsBaseSchedules!END_DATE = l_datEndDate
    
    If IsMissingValue(r_rsScheduleRequests!ROT_CD) Then
        r_rsBaseSchedules!ROT_CD = 0
        r_rsBaseSchedules!ROT_WEEK = 0
    Else
        r_rsBaseSchedules!ROT_CD = r_rsScheduleRequests!ROT_CD
        r_rsBaseSchedules!ROT_WEEK = r_rsScheduleRequests!ROT_WEEK
    End If
    
    If IsMissingValue(r_rsScheduleRequests!SCHED_CD) Then
        r_rsBaseSchedules!SCHED_CD = 0
    Else
        r_rsBaseSchedules!SCHED_CD = r_rsScheduleRequests!SCHED_CD
    End If
    
    r_rsBaseSchedules!CREATED_EMP_NUM = m_lngUserNum
    r_rsBaseSchedules!LAST_MOD_EMP_NUM = m_lngUserNum
    r_rsBaseSchedules.Update
        
'    MsgBox "Insert Base Schedule 1", , "jimb"
    AppendWGLineCascadeDates "I", r_rsBaseSchedules!WG_NUM, 0, _
                r_rsBaseSchedules!EFF_DATE, r_rsBaseSchedules!END_DATE
'    MsgBox "Insert Base Schedule 2", , "jimb"
    
    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "InsertBaseSchedule"
End Sub


Private Sub UpdateBaseSchedule( _
            ByRef r_rsScheduleRequests As ADODB.Recordset, _
            ByRef r_rsBaseSchedules As ADODB.Recordset)
'
' If the new EFF_DATE falls exactly on existing EFF_DATE, delete it.
' Find effective schedule on new EFF_DATE, if different schedule set the END_DATE
' equal to new EFF_DATE - 1.
' Determine the END_DATE of changed schedule, may have changed due to different next
' schedule.
            
    Dim l_lngWgNum As Long
    Dim l_datStartDate As Date
    Dim l_datEndDate As Date
    Dim l_strCreatedTS As String
    Dim l_datBeforeStartDate As Date
    Dim l_datBeforeEndDate As Date
    Dim l_varBookmark As Variant
    Dim l_datOldNextDate As Date
    Dim l_datOldPreviousEffDate As Date
    Dim l_datKeyStartDate As Date
    Dim l_blnDateChange As Boolean
    Dim l_blnSchedChange As Boolean

On Error GoTo eh
    r_rsBaseSchedules.Sort = "WG_NUM ASC, EFF_DATE ASC"
    l_lngWgNum = r_rsScheduleRequests!WG_NUM
    l_datStartDate = r_rsScheduleRequests!START_DATE
    l_strCreatedTS = r_rsScheduleRequests!CREATED_TIMESTAMP

' 1) Determine if old effective interval exists. i.e. an interval effective
' previous to the updated interval.  If so, save the bookmark.
    FindPreviousBaseScheduleRecord l_lngWgNum, r_rsScheduleRequests!KEY_START_DATE, r_rsBaseSchedules
    If Not r_rsBaseSchedules.EOF Then
        l_datOldPreviousEffDate = r_rsBaseSchedules!EFF_DATE
        l_varBookmark = r_rsBaseSchedules.Bookmark
    End If

' 2) End the new effective interval.
' If previous effective line sched record exists,
' set end date to day prior to this start date.
    FindEffectiveBaseScheduleRecord l_lngWgNum, l_datStartDate, r_rsBaseSchedules
    If Not r_rsBaseSchedules.EOF Then
        'if the new effective interval found is the interval being changed, skip it.
        If Not r_rsBaseSchedules("EFF_DATE").UnderlyingValue = r_rsScheduleRequests!KEY_START_DATE Then
            If r_rsBaseSchedules!EFF_DATE = l_datStartDate Then
                AppendWGLineCascadeDates "D", r_rsBaseSchedules!WG_NUM, 0, _
                    r_rsBaseSchedules!EFF_DATE, r_rsBaseSchedules!END_DATE
                r_rsBaseSchedules.Delete
            Else
                r_rsBaseSchedules!END_DATE = l_datStartDate - 1
                r_rsBaseSchedules!LAST_MOD_EMP_NUM = m_lngUserNum
                r_rsBaseSchedules.Update
            End If
        End If
    End If
    
' 3) Determine the END_DATE of the changed interval
' First, find the next base interval that is not the the
' interval being updated and get the end date.
    FindNextBaseScheduleRecord l_lngWgNum, l_datStartDate, r_rsBaseSchedules
    If Not r_rsBaseSchedules.EOF Then
        If r_rsBaseSchedules("EFF_DATE").UnderlyingValue = r_rsScheduleRequests!KEY_START_DATE Then
            ' the next found was updated interval, get the next after it.
            FindNextBaseScheduleRecord l_lngWgNum, r_rsBaseSchedules!EFF_DATE, r_rsBaseSchedules
        End If
    End If
    If Not r_rsBaseSchedules.EOF Then
        l_datEndDate = r_rsBaseSchedules!EFF_DATE - 1
    Else
        l_datEndDate = #12/31/9999#
    End If
    
    FindBaseScheduleRecord _
        l_lngWgNum, r_rsScheduleRequests!KEY_START_DATE, l_strCreatedTS, r_rsBaseSchedules
    If r_rsBaseSchedules.EOF Then
        Err.Raise trErrbusInvalidChangeEntryNotFound
    End If
    
    If r_rsBaseSchedules!EFF_DATE <> l_datStartDate Or _
        r_rsBaseSchedules!END_DATE <> l_datEndDate Then
        l_blnDateChange = True
    End If
    
    If r_rsBaseSchedules!SCHED_CD <> r_rsScheduleRequests!SCHED_CD Or _
        r_rsBaseSchedules!ROT_CD <> r_rsScheduleRequests!ROT_CD Or _
        r_rsBaseSchedules!ROT_WEEK <> r_rsScheduleRequests!ROT_WEEK Then
        l_blnSchedChange = True
    End If
    
'    For rotation, if EFF_DATE changed, it is implied schedule change.
'    The user could update the ROT_WEEK by the same number of weeks as the
'    EFF_DATE change, which would preserve the schedule.  For simplicity,
'    do not consider a ROT_WEEK change and mark entire interval as a
'    schedule change.
    If r_rsBaseSchedules!ROT_CD <> 0 And _
            r_rsBaseSchedules!EFF_DATE <> l_datStartDate Then
        l_blnSchedChange = True
    End If
    
    l_datBeforeStartDate = r_rsBaseSchedules!EFF_DATE
    l_datBeforeEndDate = r_rsBaseSchedules!END_DATE

    r_rsBaseSchedules!EFF_DATE = l_datStartDate
    r_rsBaseSchedules!END_DATE = l_datEndDate
    r_rsBaseSchedules!LAST_MOD_EMP_NUM = m_lngUserNum
    r_rsBaseSchedules!SCHED_CD = r_rsScheduleRequests!SCHED_CD
    r_rsBaseSchedules!ROT_CD = r_rsScheduleRequests!ROT_CD
    r_rsBaseSchedules!ROT_WEEK = r_rsScheduleRequests!ROT_WEEK
    r_rsBaseSchedules.Update
    
' 4) End the old effective interval.
' If old interval exists, determine if the END_DATE needs to be adjusted.
' Scan forward and if next record is found, set the END_DATE to day before its start.
    If Not IsEmpty(l_varBookmark) Then
        FindNextBaseScheduleRecord l_lngWgNum, l_datOldPreviousEffDate, r_rsBaseSchedules
        If Not r_rsBaseSchedules.EOF Then
            l_datOldNextDate = r_rsBaseSchedules!EFF_DATE
            r_rsBaseSchedules.Bookmark = l_varBookmark
            r_rsBaseSchedules!END_DATE = l_datOldNextDate - 1
        Else
            r_rsBaseSchedules.Bookmark = l_varBookmark
            r_rsBaseSchedules!END_DATE = #12/31/9999#
        End If
        r_rsBaseSchedules!LAST_MOD_EMP_NUM = m_lngUserNum
        r_rsBaseSchedules.Update
    End If
    
    If l_blnSchedChange Then
        AppendWGLineCascadeDates "C", l_lngWgNum, 0, l_datBeforeStartDate, l_datBeforeEndDate
    End If
    If l_blnDateChange Then
        AppendCascadeDateDifference "C", l_lngWgNum, 0, _
            l_datStartDate, l_datEndDate, l_datBeforeStartDate, l_datBeforeEndDate
    End If
  
    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "UpdateBaseSchedule"
End Sub

Private Sub FindBaseScheduleRecord( _
                ByVal v_lngWgNum As Long, _
                ByVal v_datEffDate As Date, _
                ByVal v_strCreatedTS As String, _
                ByRef r_rsBaseSchedules As ADODB.Recordset)
' Position r_rsBaseSchedules on the first record that matches input.

On Error GoTo eh
    
    If RecordsExist(r_rsBaseSchedules) Then
        r_rsBaseSchedules.MoveFirst
    Else
        Exit Sub
    End If

    Do Until r_rsBaseSchedules.EOF
        If r_rsBaseSchedules!WG_NUM = v_lngWgNum And _
            r_rsBaseSchedules!EFF_DATE = v_datEffDate And _
            r_rsBaseSchedules!CREATED_TIMESTAMP = v_strCreatedTS Then
            Exit Do
        End If
        r_rsBaseSchedules.MoveNext
    Loop
    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "FindBaseScheduleRecord"
End Sub

Private Sub FindEffectiveBaseScheduleRecord( _
                ByVal v_lngWgNum As Long, _
                ByVal v_datEffDate As Date, _
                ByRef r_rsBaseSchedules As ADODB.Recordset)
' Position r_rsBaseSchedules on the first record where v_datEffDate is between
' EFF_DATE and END_DATE.

On Error GoTo eh

    If RecordsExist(r_rsBaseSchedules) Then
'        r_rsBaseSchedules.Sort = "WG_NUM DESC, LINE_NUM DESC, EFF_DATE DESC"
        r_rsBaseSchedules.MoveFirst
    Else
        Exit Sub
    End If
    
    Do Until r_rsBaseSchedules.EOF
        If r_rsBaseSchedules!WG_NUM = v_lngWgNum And _
            r_rsBaseSchedules!EFF_DATE <= v_datEffDate And _
            r_rsBaseSchedules!END_DATE >= v_datEffDate Then
            Exit Do
        End If
        r_rsBaseSchedules.MoveNext
    Loop
    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "FindEffectiveBaseScheduleRecord"
End Sub

Private Sub FindPreviousBaseScheduleRecord( _
                ByVal v_lngWgNum As Long, _
                ByVal v_datEffDate As Date, _
                ByRef r_rsBaseSchedules As ADODB.Recordset)
' Position r_rsBaseSchedules on the first record that is less than v_datEffDate.

On Error GoTo eh
    If RecordsExist(r_rsBaseSchedules) Then
        r_rsBaseSchedules.Sort = "WG_NUM DESC, EFF_DATE DESC"
        r_rsBaseSchedules.MoveFirst
    Else
        Exit Sub
    End If
    Do Until r_rsBaseSchedules.EOF
        If r_rsBaseSchedules!WG_NUM = v_lngWgNum And _
            r_rsBaseSchedules!EFF_DATE < v_datEffDate Then
            Exit Do
        End If
        r_rsBaseSchedules.MoveNext
    Loop
    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "FindPreviousBaseScheduleRecord"
End Sub


Private Sub FindNextBaseScheduleRecord( _
                ByVal v_lngWgNum As Long, _
                ByVal v_datEffDate As Date, _
                ByRef r_rsBaseSchedules As ADODB.Recordset)
' Position r_rsBaseSchedules on the first record that is greater than v_datEffDate.

On Error GoTo eh
    If RecordsExist(r_rsBaseSchedules) Then
        r_rsBaseSchedules.Sort = "WG_NUM ASC, EFF_DATE ASC"
        r_rsBaseSchedules.MoveFirst
    Else
        Exit Sub
    End If
    Do Until r_rsBaseSchedules.EOF
        If r_rsBaseSchedules!WG_NUM = v_lngWgNum And _
            r_rsBaseSchedules!EFF_DATE > v_datEffDate Then
            Exit Do
        End If
        r_rsBaseSchedules.MoveNext
    Loop
    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "FindNextBaseScheduleRecord"
End Sub


Private Function CheckBaseSchedulesEffStartEnd( _
                    ByRef r_rsBaseSchedules As ADODB.Recordset, _
                    ByRef r_rsWGAttributes As ADODB.Recordset) As Boolean
                    
On Error GoTo eh

    CheckBaseSchedulesEffStartEnd = False
    If Not RecordsExist(r_rsWGAttributes) Then
        CheckBaseSchedulesEffStartEnd = True
        Exit Function
    End If

    If RecordsExist(r_rsBaseSchedules) Then
        r_rsBaseSchedules.Sort = "WG_NUM ASC, EFF_DATE ASC"
        r_rsBaseSchedules.MoveFirst
    Else
        AppendWGSchedError m_rsErrors, _
            trErrbusValidationError, mstrCBUSINESSERR, _
            "Base schedule does not exist."
        CheckBaseSchedulesEffStartEnd = True
        Exit Function
    End If
    
    If r_rsBaseSchedules!EFF_DATE <> r_rsWGAttributes!EFF_START_DATE Then
        AppendWGSchedError m_rsErrors, _
            trErrbusValidationError, mstrCBUSINESSERR, _
            "First schedule interval Start Date(" & _
            Format(r_rsBaseSchedules!EFF_DATE, "mm/dd/yyyy") & _
            ") is not equal to workgroup Start Date(" & _
            Format(r_rsWGAttributes!EFF_START_DATE, "mm/dd/yyyy") & ")", _
            , , CStr(r_rsBaseSchedules!WG_NUM), r_rsBaseSchedules!EFF_DATE
        CheckBaseSchedulesEffStartEnd = True
    End If

    r_rsBaseSchedules.MoveLast
    If r_rsBaseSchedules!END_DATE <> #12/31/9999# Then
        AppendWGSchedError m_rsErrors, _
            trErrbusValidationError, mstrCBUSINESSERR, _
           "Last base schedule interval EFF_DATE is not equal to 12/31/9999"
        CheckBaseSchedulesEffStartEnd = True
    End If

    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "CheckBaseSchedulesEffStartEnd"
End Function


Private Function AnyWorkgroupAssignments( _
                ByVal v_lngWgNum As Long, _
                ByVal v_datStartDate As Date, _
                ByVal v_datEndDate As Date) As Boolean
    Dim l_strFilter As String
    Dim l_intIntersection As Integer

On Error GoTo eh
    AnyWorkgroupAssignments = False
'WGSCHED - Check for workgroup assignments later.
'    l_strFilter = "WG_NUM = " & v_lngWGNum
'    m_rsWGLineAssignments.Filter = l_strFilter
'    Do Until m_rsWGLineAssignments.EOF
'        l_intIntersection = DateIntersection(v_datStartDate, v_datEndDate, _
'                m_rsWGLineAssignments!EFF_DATE, m_rsWGLineAssignments!END_DATE)
'        If l_intIntersection >= 1 And l_intIntersection <= 4 Then
'            AnyLineAssignments = True
'            Exit Do
'        End If
'        m_rsWGLineAssignments.MoveNext
'    Loop
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "AnyWorkgroupAssignments"
End Function



Private Function ProcessWorkGroupDistribution( _
    ByVal v_rsDistributionRequests As ADODB.Recordset _
    ) As Boolean
    
    'This function processes the request for saving distribution pick list.
'There are only 2 rules in this routine.
'  1. It is a valid request
'  2. The distribution validates
'The distribution is not validated on deleted records.
'A blank distribution is placed in the distribution to be validated
'This is done so the index in the collection is the absolute position
'of the request recordset
    
    Dim l_colValidDistribution As Collection
    Dim l_blnValidRequest As Boolean
    Dim l_blnValidDistribution As Boolean
    Dim l_strRecStatus As String
    
    On Error GoTo eh
    
    
    If RecordsExist(v_rsDistributionRequests) Then
        'Continue
    Else
        ProcessWorkGroupDistribution = True
        Exit Function
    End If
    
    l_blnValidRequest = IsValidWorkGroupDistributionRequest(v_rsDistributionRequests, l_colValidDistribution)
    l_blnValidDistribution = IsValidWorkGroupDistribution(v_rsDistributionRequests, l_colValidDistribution)
    
    If l_blnValidRequest And l_blnValidDistribution Then
        'continue
    Else
        ProcessWorkGroupDistribution = False
        Exit Function
    End If
    
    ApplyDistributionRequests v_rsDistributionRequests, l_colValidDistribution
    
    ProcessWorkGroupDistribution = True
    Exit Function
 
eh:
    TrsRaiseError m_strCMODULENAME, "ProcessWorkGroupDistribution"
End Function


Private Function IsValidWorkGroupDistributionRequest( _
    ByVal v_rsDistributionRequests As ADODB.Recordset, _
    ByRef r_colDistribution As Collection _
    ) As Boolean

    Dim l_blnIsValid As Boolean
    Dim l_strRecStatus As String
    Dim l_strFilter As String
    
On Error GoTo eh
    
    l_blnIsValid = True
    
    Set r_colDistribution = New Collection
    v_rsDistributionRequests.MoveFirst
    While Not v_rsDistributionRequests.EOF
        l_strRecStatus = Trim(UCase(v_rsDistributionRequests!RECORD_STATUS))
        Select Case l_strRecStatus
            Case "I", "C"
                If ValidateKeyWorkGroup(v_rsDistributionRequests, TrType_WGM_ACCT) Then
                    If DistributionExist(v_rsDistributionRequests) Then
                        r_colDistribution.Add v_rsDistributionRequests.Fields("DISTRIBUTION_ID").Value
                    Else
                        l_blnIsValid = False
                        r_colDistribution.Add ""
                        AppendWGDistributionBRV v_rsDistributionRequests, 30000, mstrCBUSINESSERR, "Distribution is missing. Please enter and try again."
                    End If
                Else
                    l_blnIsValid = False
                End If
                
            Case "D"
                If ValidateKeyWorkGroup(v_rsDistributionRequests, TrType_WGM_ACCT) Then
                    r_colDistribution.Add ""
                Else
                    l_blnIsValid = False
                End If
                
            Case Else
                l_blnIsValid = False
                r_colDistribution.Add ""
                AppendWGDistributionBRV v_rsDistributionRequests, 30000, mstrCBUSINESSERR, "Invalid Record Status"
                
        End Select
        v_rsDistributionRequests.MoveNext
    Wend
    
    IsValidWorkGroupDistributionRequest = l_blnIsValid
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "IsValidWorkGroupDistribution"
    
        
End Function


Private Function AppendWGDistributionBRV( _
            ByVal v_rsWorkGroupDistribution As ADODB.Recordset, _
            ByVal v_lngErrorNumber As Long, _
            ByVal v_strErrorType As String, _
            ByVal v_strErrorDesc As String _
            ) As Boolean
            
    Dim l_bolError As Boolean
    Dim lngKeyWGNum As Long
    Dim l_strKeyDistribution As String
    Dim l_strKeyCreatedTime As String
    
    On Error GoTo AppendWGDistributionBRV_EH
    
    lngKeyWGNum = v_rsWorkGroupDistribution!KEY_WG_NUM
    l_strKeyDistribution = v_rsWorkGroupDistribution!KEY_DISTRIBUTION_ID
    If IsNull(v_rsWorkGroupDistribution!TIME_STAMP) Then
        l_strKeyCreatedTime = ""
    Else
        l_strKeyCreatedTime = v_rsWorkGroupDistribution!TIME_STAMP
    End If
    
    l_bolError = AppendBusinessError(m_rsErrors, _
            v_lngErrorNumber, _
            v_strErrorType, _
            v_strErrorDesc, _
            m_strCMODULENAME, _
            TrType_WGM_ACCT, , , "WG_NUM;DISTRIBUTION_ID", _
            str(lngKeyWGNum), _
            "", _
            "0", _
            "0", _
            l_strKeyDistribution, _
            "")

    AppendWGDistributionBRV = l_bolError
    Exit Function
    
AppendWGDistributionBRV_EH:
    TrsRaiseError m_strCMODULENAME, "AppendWGDistributionBRV"

End Function

Private Function ValidateKeyWorkGroup( _
         ByVal v_objRS As ADODB.Recordset, _
         ByVal v_intRequestType As Integer _
         ) As Boolean
            
'Answer True when valid, otherwise False.
    Dim l_blnIsValid As Boolean
    Dim l_strWGNum As String
    Dim l_datEffDate As Date
    Dim l_blnWriteError As Boolean
    Dim l_strErrDesc As String

    On Error GoTo eh
    
    l_strErrDesc = "Invalid request, please try again.  If problem continues contact your TRIS administrator to help you. "
    l_blnIsValid = True
    If Not IsNull(v_objRS!KEY_WG_NUM) Then
        l_strWGNum = v_objRS!KEY_WG_NUM
        If v_objRS!KEY_WG_NUM < 1 Then
         l_strErrDesc = l_strErrDesc & "KEY_WG_NUM must be > 0"
            l_blnIsValid = False
        End If
    Else
        l_strWGNum = "0"
        l_strErrDesc = l_strErrDesc & "KEY_WG_NUM must be > 0"
        l_blnIsValid = False
    End If
    
    If l_blnIsValid Then
        If Not IsNull(v_objRS!KEY_WG_NUM) And _
            Not IsNull(v_objRS!WG_NUM) Then
            l_strWGNum = v_objRS!WG_NUM
            If v_objRS!KEY_WG_NUM <> v_objRS!WG_NUM Then
                l_strErrDesc = l_strErrDesc & "WG_NUM must = KEY_WG_NUM"
                l_blnIsValid = False
            End If
        ElseIf IsNull(v_objRS!WG_NUM) Then
            l_strWGNum = "0"
            l_strErrDesc = l_strErrDesc & "WG_NUM cannot be Null"""
            l_blnIsValid = False
        End If
    End If
    
    If l_blnIsValid Then
        l_strWGNum = v_objRS!WG_NUM
        If m_lngWorkGroup <> v_objRS!WG_NUM Then
            l_strErrDesc = l_strErrDesc & "WG_NUM must match work group in the request"
            l_blnIsValid = False
        End If
    End If
    
    If Not l_blnIsValid Then
        Select Case v_intRequestType
            Case TrType_WGM
                l_blnWriteError = _
                AppendWGAttributesBRV(v_objRS, 30000, mstrCBUSINESSERR, l_strErrDesc)
            
            Case TrType_WGM_SCHED
                AppendWGSchedLineError m_rsErrors, v_objRS, _
                    trErrbusValidationError, mstrCBUSINESSERR, _
                    l_strErrDesc, , l_strWGNum
            
            Case TrType_WGM_ASGMT
                l_blnWriteError = _
                AppendWGAssignmentsBRV(v_objRS, 30000, mstrCBUSINESSERR, l_strErrDesc)
            
            Case TrType_WGM_ACCT
                l_blnWriteError = _
                AppendWGDistributionBRV(v_objRS, 30000, mstrCBUSINESSERR, l_strErrDesc)
            
            
            Case TrType_WGM_ACTIV
                l_blnWriteError = _
                AppendWGActivityBRV(v_objRS, 30000, mstrCBUSINESSERR, l_strErrDesc)
'
        End Select

    End If
    
    ValidateKeyWorkGroup = l_blnIsValid
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "ValidateKeyWorkGroup"
    
End Function

Private Function DistributionExist( _
    ByVal v_rsDistributionRequests As ADODB.Recordset _
    ) As Boolean
    
    Dim l_strErrDesc As String
    Dim l_blnWriteError As Boolean
    
    On Error GoTo eh

    l_strErrDesc = "Invalid request, please try again.  The DISTRIBUTION is missiing. If problem continues contact your TRIS administrator to help you. "
    If IsNull(v_rsDistributionRequests!DISTRIBUTION_ID) Or _
              IsEmpty(v_rsDistributionRequests!DISTRIBUTION_ID) Then
        DistributionExist = False
        l_blnWriteError = _
        AppendWGDistributionBRV(v_rsDistributionRequests, 30000, mstrCBUSINESSERR, l_strErrDesc)
    Else
        If Trim(v_rsDistributionRequests!DISTRIBUTION_ID) = "" Then
            l_strErrDesc = "The distribution is blank. Please enter it and try again. "
            DistributionExist = False
            l_blnWriteError = _
                AppendWGDistributionBRV(v_rsDistributionRequests, 30000, mstrCBUSINESSERR, l_strErrDesc)
        Else
            DistributionExist = True
        End If
    End If
    

    
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "DistributionExist"

End Function


Private Function IsValidWorkGroupDistribution( _
    ByVal v_rsDistributionRequests As ADODB.Recordset, _
    ByRef r_colDistribution As Collection _
    ) As Boolean
    
    Dim l_intI, l_intCount As Integer
    Dim l_strDistribution As String
    Dim l_strDistIdForValidation As String

    Dim l_blnFMCDefault As Boolean
    Dim l_blnACTVDefault As Boolean
    Dim l_blnWGDistributionIsValid As Boolean
    Dim l_strErrorMsg As String
    Dim l_strReturnMsg As String
    Dim l_blnWriteError As Boolean
    Dim l_colValidDistribution As Collection
    Dim l_strReasonForEdit As String
    Dim l_strTransType As String
    
    On Error GoTo eh
    

    Set l_colValidDistribution = New Collection
    l_blnWGDistributionIsValid = True
    l_intCount = r_colDistribution.Count
    For l_intI = 1 To l_intCount
        l_strDistribution = r_colDistribution.Item(l_intI)
        l_strTransType = Mid(l_strDistribution, 16, 1)  '//check for POET and skip validation
        
        If l_strDistribution = "" Then
            'skip
            l_colValidDistribution.Add ""
        ElseIf l_strTransType = "P" Then
            l_colValidDistribution.Add l_strDistribution
        Else
'            If IsServiceRequest(l_strDistribution) Then
'                l_strReasonForEdit = "PI"
'            Else
                l_strReasonForEdit = "FE"
'            End If
            
            l_strDistIdForValidation = GetDistIdForValidation(l_strDistribution, l_blnFMCDefault, l_blnACTVDefault)

            If Not ValidateDistID(l_strReasonForEdit, l_strDistIdForValidation, l_strReturnMsg) Then
                l_blnWGDistributionIsValid = False
                v_rsDistributionRequests.AbsolutePosition = l_intI
                l_strErrorMsg = "Distribution is not valid.  " & l_strReturnMsg
                l_blnWriteError = _
                  AppendWGDistributionBRV(v_rsDistributionRequests, 30000, mstrCBUSINESSERR, l_strErrorMsg)
            Else
                l_colValidDistribution.Add GetDistIdAfterValidation(l_strDistIdForValidation, l_blnFMCDefault, l_blnACTVDefault)
            End If
        End If
    Next l_intI
    
    Set r_colDistribution = l_colValidDistribution
    IsValidWorkGroupDistribution = l_blnWGDistributionIsValid
    
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "IsValidWorkGroupDistribution"
    
End Function

'//copied from busDistributionMgr.cbsDistributionMgr (cannot create reference to it)
Private Function IsServiceRequest(ByVal v_varDistributionId As Variant) As Boolean

' A Distribution ID is a service request
' if only the following 3 fields are entered
' FMC, Project, Activity
' The Feeder_ref (Memo) field is ignored
' The Project field must be entered to be a service request

    Dim l_strTempHold As Variant
    Dim l_intNumOfFields As Integer
    
On Error GoTo IsServiceRequest_EH
    
    Dim l_strDistributionId(11) As String
    Dim l_strEmptyFields As String
    
    If Not IsArray(v_varDistributionId) Then
        l_strDistributionId(0) = Left$(v_varDistributionId, 1) '(U)Utility
        l_strDistributionId(1) = Mid$(v_varDistributionId, 2, 2) '(BD) Business Division
        l_strDistributionId(2) = Mid$(v_varDistributionId, 4, 3) '(MAJ)FERC Major
        l_strDistributionId(3) = Mid$(v_varDistributionId, 7, 3) '(MIN)FERC Minor
        l_strDistributionId(4) = Mid$(v_varDistributionId, 10, 3) '(FMC) Fund Mgt. Center
        l_strDistributionId(5) = Mid$(v_varDistributionId, 13, 3) '(RMC) Resource Mgt. Center
        l_strDistributionId(6) = Mid$(v_varDistributionId, 16, 1) '(T) Trans Type
        l_strDistributionId(7) = Mid$(v_varDistributionId, 17, 5) '(PROJT) Project
        l_strDistributionId(8) = Mid$(v_varDistributionId, 22, 2) '(PR) Product
        l_strDistributionId(9) = Mid$(v_varDistributionId, 24, 4) '(ACTV) Activity
        l_strDistributionId(10) = Mid$(v_varDistributionId, 28, 2) '(RT) Resource Type
        l_strDistributionId(11) = Mid$(v_varDistributionId, 30, 10) '(FEEDER_REF) Memo
    Else
        l_strDistributionId(0) = v_varDistributionId(0) '(U)Utility
        l_strDistributionId(1) = v_varDistributionId(1) '(BD)  Business Division
        l_strDistributionId(2) = v_varDistributionId(2) '(MAJ)FERC Major
        l_strDistributionId(3) = v_varDistributionId(3) '(MIN)FERC Minor
        l_strDistributionId(4) = v_varDistributionId(4) '(FMC) Fund Mgt. Center
        l_strDistributionId(5) = v_varDistributionId(5) '(RMC) Resource Mgt. Center
        l_strDistributionId(6) = v_varDistributionId(6) '(T) Trans Type
        l_strDistributionId(7) = v_varDistributionId(7) '(PROJT) Project
        l_strDistributionId(8) = v_varDistributionId(8) '(PR) Product
        l_strDistributionId(9) = v_varDistributionId(9) '(ACTV) Activity
        l_strDistributionId(10) = v_varDistributionId(10) '(RT) Resource Type
        l_strDistributionId(11) = v_varDistributionId(11) '(FEEDER_REF) Memo
    End If
    
    l_strEmptyFields = l_strDistributionId(0) & _
                       l_strDistributionId(1) & _
                       l_strDistributionId(2) & _
                       l_strDistributionId(3) & _
                       l_strDistributionId(5) & _
                       l_strDistributionId(6) & _
                       l_strDistributionId(8) & _
                       l_strDistributionId(10)
    If Trim(l_strEmptyFields) <> "" Then
        IsServiceRequest = False
        Exit Function
    End If
    If Trim(l_strDistributionId(7)) <> "" Then
'Please not that FMC and Activity may be missing
        IsServiceRequest = True
    Else
        IsServiceRequest = False
    End If
    Exit Function
IsServiceRequest_EH:
    TrsRaiseError m_strCMODULENAME, "IsServiceRequest", , , , , True
    
End Function

Private Function GetDistIdForValidation( _
    ByVal v_strDistribution As String, _
    ByRef r_blnFMCDefault As Boolean, _
    ByRef r_blnACTVDefault As Boolean _
    ) As String
    
    Dim l_strFMC As String
    Dim l_strACTV As String
    Dim l_strDept As String
    Dim l_obj As dbDepartment.cdDepartment

    On Error GoTo eh
    
    r_blnFMCDefault = False
    r_blnACTVDefault = False
    l_strFMC = Mid$(v_strDistribution, 10, 3)
    l_strACTV = Mid$(v_strDistribution, 24, 4)
'    If Trim(l_strACTV) = "" Then
'        l_strACTV = "ADMN"
'        r_blnACTVDefault = True
'    End If
    If Trim(l_strFMC) = "" Then
        l_strDept = GetDepartment()
        Set l_obj = CreateObject("dbDepartment.cdDepartment")
        l_strFMC = l_obj.GetBudgetNumber(l_strDept)
        Set l_obj = Nothing
        r_blnFMCDefault = True
    End If

    GetDistIdForValidation = Mid$(v_strDistribution, 1, 9) & l_strFMC & _
        Mid$(v_strDistribution, 13, 11) & l_strACTV & Mid$(v_strDistribution, 28, 12)
        
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "GetDistIdForValidation"
    
End Function

Private Function ValidateDistID( _
    ByVal v_strReasonForEdit As String, _
    ByRef r_strDistID As String, _
    ByRef r_strReturnMsg As String _
    ) As Boolean
    
    Dim l_strReturnMsg As String
    
    Dim l_strCorporateCode As String
    Dim l_strDistributionIDEdits(221) As String
    Dim l_obj As dbDistributionMgr.cdDistributionMgr
    Dim l_blnSuccess As Boolean
    Dim l_strDistIdsReturned As Variant
    Dim l_strReturnCodes As Variant
    Dim l_strReturnMessages As Variant
    Dim l_strHighlightFields As Variant
    
    On Error GoTo eh
    
    r_strReturnMsg = ""
    l_strCorporateCode = GetCompanyCode()
    l_strDistributionIDEdits(0) = v_strReasonForEdit & l_strCorporateCode & r_strDistID
    Set l_obj = CreateObject("dbDistributionMgr.cdDistributionMgr")
    l_blnSuccess = l_obj.ValidateGLAccounting(l_strDistributionIDEdits, 1, l_strDistIdsReturned, l_strReturnCodes, l_strReturnMessages, l_strHighlightFields)
    If l_blnSuccess Then
        If Trim(l_strDistIdsReturned(0)) = "" Then
            ValidateDistID = False
            r_strReturnMsg = "Please try again. GL Distribution validation unavailable at this time."
            Exit Function
        End If
'    Else
'        r_strReturnMsg = "Call to AAVM Failed"
'        ValidateDistID = False
'        Exit Function
    End If
    If l_strReturnCodes(0) = 0 Then
        r_strDistID = l_strDistIdsReturned(0)
        r_strReturnMsg = l_strReturnMessages(0)
        ValidateDistID = True
    Else
        ValidateDistID = False
        r_strReturnMsg = l_strReturnMessages(0)
    End If
    Exit Function
    
    
    
eh:
    TrsRaiseError m_strCMODULENAME, "ValidateDistID"
    
    
End Function

Private Function GetCompanyCode()

    Dim l_strDept As String
    Dim l_obj As dbDepartment.cdDepartment
    
    On Error GoTo eh
    
    l_strDept = GetDepartment()
    Set l_obj = CreateObject("dbDepartment.cdDepartment")
    GetCompanyCode = l_obj.GetCompanyCode(l_strDept)
    Set l_obj = Nothing
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "GetCompanyCode"


End Function

Private Function GetDepartment()

    On Error GoTo eh
    '//receive an "Arguments are of the wrong type, are out of acceptable range, or are in conflict with one another" error
    '//on Windows 7 and Windows 2008
    '//m_rsWorkGroupAttributes.Filter = m_lngWorkGroup '//invalid filter, but WinXP/Windows 2003 did not complain
    
    m_rsWorkGroupAttributes.Filter = "WG_NUM = " & m_lngWorkGroup
    m_rsWorkGroupAttributes.MoveFirst
    GetDepartment = m_rsWorkGroupAttributes!DEPT_NUM
            
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "GetDepartment"
        
End Function
            
            
Private Function GetDistIdAfterValidation( _
    ByVal v_strValidatedWGDistribution As String, _
    ByVal v_blnFMCDefault As Boolean, _
    ByVal v_blnACTVDefault As Boolean _
    ) As String
    
    Dim l_strDistIdAfterValidation As String
    
    On Error GoTo eh
    
    l_strDistIdAfterValidation = v_strValidatedWGDistribution
    If v_blnFMCDefault Then
        l_strDistIdAfterValidation = Mid$(l_strDistIdAfterValidation, 1, 9) & _
            Space(3) & Mid$(l_strDistIdAfterValidation, 13, 27)
    End If
    
    If v_blnACTVDefault Then
        l_strDistIdAfterValidation = Mid$(l_strDistIdAfterValidation, 1, 23) & _
            Space(4) & Mid$(l_strDistIdAfterValidation, 28, 12)
    End If
    
    GetDistIdAfterValidation = l_strDistIdAfterValidation
    
                
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "GetDistIdAfterValidation"
    
    
End Function


Private Sub ApplyDistributionRequests( _
    ByVal v_rsWorkGroupDistribution As ADODB.Recordset, _
    ByVal v_colDistribution As Collection)
       
    Dim l_strRecStatus As String
    Dim l_blnRequestProcessed As Boolean
    
On Error GoTo eh

    v_rsWorkGroupDistribution.MoveFirst
    While Not v_rsWorkGroupDistribution.EOF
        l_strRecStatus = UCase(v_rsWorkGroupDistribution!RECORD_STATUS)
        Select Case l_strRecStatus
            Case "I"
                DistributionInsert v_rsWorkGroupDistribution, v_colDistribution
            Case "C"
                DistributionModify "C", v_rsWorkGroupDistribution, v_colDistribution
            Case "D", "Z"
                DistributionModify "D", v_rsWorkGroupDistribution, v_colDistribution
        End Select
    
        v_rsWorkGroupDistribution.MoveNext
    Wend

    Exit Sub
    
eh:
    
  TrsRaiseError m_strCMODULENAME, "ApplyDistributionRequests"

        
        
End Sub


Private Sub DistributionModify( _
    ByVal v_strRecStatus As String, _
    ByVal v_rsWorkGroupDistribution As ADODB.Recordset, _
    ByVal v_colDistribution As Collection)
    
    Dim l_strFilter As String
    Dim l_strDistribution As String
    Dim l_strTimeStamp As String
    Dim l_lngKeyWGNum As Long
    Dim l_intPos As Integer
    Dim l_strErrorDesc As String
    Dim l_colErrors As Collection
          
    On Error GoTo eh
    
    If v_strRecStatus = "C" Then
        l_strErrorDesc = "Invalid Update Request"
    Else
        l_strErrorDesc = "Invalid Delete Request"
    End If
    l_lngKeyWGNum = v_rsWorkGroupDistribution!KEY_WG_NUM
    l_strDistribution = v_rsWorkGroupDistribution!KEY_DISTRIBUTION_ID
    l_strTimeStamp = v_rsWorkGroupDistribution!TIME_STAMP
    l_strFilter = "WG_NUM = " & str(l_lngKeyWGNum) & _
            " AND DISTRIBUTION_ID = '" & l_strDistribution & "'" & _
            " AND TIME_STAMP = '" & l_strTimeStamp & "'"
    l_strFilter = MDAC25FilterFix(l_strFilter)
    m_rsWGDistribution.Filter = l_strFilter
    If Not RecordsExist(m_rsWGDistribution) Then
        m_rsWG_Assignments.Filter = ""
        Err.Raise vbObjectError + 30000, "busWorkGroupTrans.DistributionModify", l_strErrorDesc
        'Raise error
        'Invalid change request
        'Data has been updated by someone else prior to save
        'Please Try again.
        Exit Sub
    End If

    If v_strRecStatus = "C" Then
        l_intPos = v_rsWorkGroupDistribution.AbsolutePosition
        m_rsWGDistribution!DISTRIBUTION_ID = v_colDistribution.Item(l_intPos)
        m_rsWGDistribution!Description = v_rsWorkGroupDistribution!Description
        m_rsWGDistribution!LAST_MOD_EMP_NUM = CLng(m_strUserId)
    Else
        m_rsWGDistribution.Delete
    End If
    m_rsWGDistribution.Update
    
    Exit Sub
    
eh:
    Set l_colErrors = New Collection
    l_colErrors.Add "Work Group=" & CStr(m_lngWorkGroup)
    l_colErrors.Add "Key Work Group =" & CStr(l_lngKeyWGNum)
    l_colErrors.Add "Key Distribution= (" & l_strDistribution & ")"
    l_colErrors.Add "Timestamp= (" & l_strTimeStamp & ")"

    TrsRaiseError m_strCMODULENAME, "DistributionModify", , , , l_colErrors, True
    Set l_colErrors = Nothing

End Sub

Private Sub DistributionInsert( _
    ByVal v_rsWorkGroupDistribution As ADODB.Recordset, _
    ByVal v_colDistribution As Collection)
    
    Dim l_intPos As Integer
          
    On Error GoTo eh
    
    m_rsWGDistribution.AddNew
    m_rsWGDistribution!EMP_NUM = 0
    m_rsWGDistribution!WG_NUM = m_lngWorkGroup
    l_intPos = v_rsWorkGroupDistribution.AbsolutePosition
    m_rsWGDistribution!DISTRIBUTION_ID = v_colDistribution.Item(l_intPos)
    m_rsWGDistribution!Description = v_rsWorkGroupDistribution!Description
    m_rsWGDistribution!LAST_MOD_EMP_NUM = CLng(m_strUserId)
    m_rsWGDistribution.Update
 
    Exit Sub
    
eh:
    
  TrsRaiseError m_strCMODULENAME, "DistributionInsert"

End Sub


Private Function ProcessWorkGroupActivity( _
    ByVal v_rsActivityRequests As ADODB.Recordset _
    ) As Boolean
    
    'This function processes the request for saving Activity pick list.
'There are only 2 rules in this routine.
'  1. It is a valid request
'  2. The Activity validates
'The Activity is not validated on deleted records.
'A blank distribution is placed in the distribution to be validated
'This is done so the index in the collection is the absolute position
'of the request recordset
    
    Dim l_colActivity As Collection
    Dim l_blnValidRequest As Boolean
    Dim l_blnValidActivity As Boolean
    Dim l_strRecStatus As String
    
    On Error GoTo eh
    
''    m_rsWGActivities
    
    If RecordsExist(v_rsActivityRequests) Then
        'Continue
    Else
        ProcessWorkGroupActivity = True
        Exit Function
    End If
    
    l_blnValidRequest = IsValidWorkGroupActivityRequest(v_rsActivityRequests, l_colActivity)
    l_blnValidActivity = IsValidWorkGroupActivity(v_rsActivityRequests, l_colActivity)
    
    If l_blnValidRequest And l_blnValidActivity Then
        'continue
    Else
        ProcessWorkGroupActivity = False
        Exit Function
    End If
    
    ApplyActivityRequests v_rsActivityRequests
    
    ProcessWorkGroupActivity = True
    Exit Function
 
eh:
    TrsRaiseError m_strCMODULENAME, "ProcessWorkGroupActivity"
End Function


Private Function IsValidWorkGroupActivityRequest( _
    ByVal v_rsActivityRequests As ADODB.Recordset, _
    ByRef r_colDistribution As Collection _
    ) As Boolean

    Dim l_blnIsValid As Boolean
    Dim l_strRecStatus As String
    Dim l_strFilter As String
    
On Error GoTo eh
    
    l_blnIsValid = True
    
    Set r_colDistribution = New Collection
    v_rsActivityRequests.MoveFirst
    While Not v_rsActivityRequests.EOF
        l_strRecStatus = Trim(UCase(v_rsActivityRequests!RECORD_STATUS))
        Select Case l_strRecStatus
            Case "I"
                If ValidateKeyWorkGroup(v_rsActivityRequests, TrType_WGM_ACTIV) Then
                    If ActivityExist(v_rsActivityRequests) Then
                        r_colDistribution.Add v_rsActivityRequests.Fields("ACTIVITY_CD").Value
                    Else
                        l_blnIsValid = False
                        r_colDistribution.Add ""
                        AppendWGActivityBRV v_rsActivityRequests, 30000, mstrCBUSINESSERR, "Distribution is missing. Please enter and try again."
                    End If
                Else
                    l_blnIsValid = False
                End If
                
            Case "D"
                If ValidateKeyWorkGroup(v_rsActivityRequests, TrType_WGM_ACTIV) Then
                    r_colDistribution.Add ""
                Else
                    l_blnIsValid = False
                End If
                
            Case Else
                l_blnIsValid = False
                r_colDistribution.Add ""
                AppendWGActivityBRV v_rsActivityRequests, 30000, mstrCBUSINESSERR, "Invalid Record Status"
                
        End Select
        v_rsActivityRequests.MoveNext
    Wend
    
    IsValidWorkGroupActivityRequest = l_blnIsValid
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "IsValidWorkGroupActivityRequest"
    
        
End Function

Private Function AppendWGActivityBRV( _
            ByVal v_rsWorkGroupActivity As ADODB.Recordset, _
            ByVal v_lngErrorNumber As Long, _
            ByVal v_strErrorType As String, _
            ByVal v_strErrorDesc As String _
            ) As Boolean
            
    Dim l_bolError As Boolean
    Dim lngKeyWGNum As Long
    Dim l_strKeyActivity As String
    Dim l_strKeyCreatedTime As String
    
    On Error GoTo AppendWGActivityBRV_EH
    
    lngKeyWGNum = v_rsWorkGroupActivity!KEY_WG_NUM
    l_strKeyActivity = v_rsWorkGroupActivity!KEY_ACTIVITY_CD
    If IsNull(v_rsWorkGroupActivity!TIME_STAMP) Then
        l_strKeyCreatedTime = ""
    Else
        l_strKeyCreatedTime = v_rsWorkGroupActivity!TIME_STAMP
    End If
    
    l_bolError = AppendBusinessError(m_rsErrors, _
            v_lngErrorNumber, _
            v_strErrorType, _
            v_strErrorDesc, _
            m_strCMODULENAME, _
            TrType_WGM_ACTIV, , , "WG_NUM;ACTIVITY_CD", _
            str(lngKeyWGNum), _
            "", _
            "0", _
            "0", _
            "", _
            l_strKeyActivity)

    AppendWGActivityBRV = l_bolError
    Exit Function
    
AppendWGActivityBRV_EH:
    TrsRaiseError m_strCMODULENAME, "AppendWGActivityBRV"

End Function

Private Function ActivityExist( _
    ByVal v_rsActivityRequests As ADODB.Recordset _
    ) As Boolean
    
    Dim l_strErrDesc As String
    Dim l_blnWriteError As Boolean
    
    On Error GoTo eh

    l_strErrDesc = "Invalid request, please try again.  The Activity is missiing. If problem continues contact your TRIS administrator to help you. "
    If IsNull(v_rsActivityRequests!ACTIVITY_CD) Or _
              IsEmpty(v_rsActivityRequests!ACTIVITY_CD) Then
        ActivityExist = False
        l_blnWriteError = _
        AppendWGActivityBRV(v_rsActivityRequests, 30000, mstrCBUSINESSERR, l_strErrDesc)
    Else
        If Trim(v_rsActivityRequests!ACTIVITY_CD) = "" Then
            l_strErrDesc = "The activity is blank. Please enter it and try again. "
            ActivityExist = False
            l_blnWriteError = _
                AppendWGActivityBRV(v_rsActivityRequests, 30000, mstrCBUSINESSERR, l_strErrDesc)
        Else
            ActivityExist = True
        End If
    End If
    

    
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "ActivityExist"

End Function



Private Function IsValidWorkGroupActivity( _
    ByVal v_rsActivityRequests As ADODB.Recordset, _
    ByVal v_colActivity As Collection _
    ) As Boolean
    
    Dim l_intI, l_intCount As Integer
    Dim l_strActivity As String
    Dim l_strDistID As String
    Dim l_blnWGActivityIsValid As Boolean
    Dim l_strErrorMsg As String
    Dim l_strReturnMsg As String
    Dim l_blnWriteError As Boolean


    
    On Error GoTo eh
    
    l_blnWGActivityIsValid = True
    Set m_rsGLActivities = GetGLActivities(v_colActivity)

    l_intCount = v_colActivity.Count
    For l_intI = 1 To l_intCount
        l_strActivity = v_colActivity.Item(l_intI)
        If l_strActivity = "" Then
            'skip
        Else
            l_strDistID = Space(23) & l_strActivity & Space(12)
            If Not ValidateDistID("AV", l_strDistID, l_strReturnMsg) Then
                l_blnWGActivityIsValid = False
                l_strErrorMsg = "Activity is not valid.  " & l_strReturnMsg
                v_rsActivityRequests.AbsolutePosition = l_intI
                l_blnWriteError = _
                  AppendWGActivityBRV(v_rsActivityRequests, 30000, mstrCBUSINESSERR, l_strErrorMsg)
            Else
                ProcessGLActivity l_strActivity, l_strReturnMsg
            End If
        End If
    Next l_intI
    
    
    IsValidWorkGroupActivity = l_blnWGActivityIsValid
    
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "IsValidWorkGroupActivity"
    
End Function


Private Sub ApplyActivityRequests( _
    ByVal v_rsWorkGroupActivity As ADODB.Recordset)
    
       
    Dim l_strRecStatus As String
    
On Error GoTo eh

    v_rsWorkGroupActivity.MoveFirst
    While Not v_rsWorkGroupActivity.EOF
        l_strRecStatus = UCase(v_rsWorkGroupActivity!RECORD_STATUS)
        Select Case l_strRecStatus
            Case "I"
                ActivityInsert v_rsWorkGroupActivity
            Case "D", "Z"
                ActivityDelete v_rsWorkGroupActivity
        End Select
        v_rsWorkGroupActivity.MoveNext
    Wend

    Exit Sub
    
eh:
    
  TrsRaiseError m_strCMODULENAME, "ApplyActivityRequests"

        
        
End Sub

Private Sub ActivityDelete( _
    ByVal v_rsWorkGroupActivity As ADODB.Recordset)
    
    Dim l_strFilter As String
    Dim l_strActivity As String
    Dim l_strTimeStamp As String
    Dim l_lngKeyWGNum As Long
    Dim l_intPos As Integer
    Dim l_strErrorDesc As String
          
    On Error GoTo eh
    
    l_strErrorDesc = "Invalid Delete Request"

    l_lngKeyWGNum = v_rsWorkGroupActivity!KEY_WG_NUM
    l_strActivity = v_rsWorkGroupActivity!KEY_ACTIVITY_CD
    l_strTimeStamp = v_rsWorkGroupActivity!TIME_STAMP
    l_strFilter = "WG_NUM = " & str(l_lngKeyWGNum) & _
            " AND ACTIVITY_CD = '" & l_strActivity & "'" & _
            " AND TIME_STAMP = '" & l_strTimeStamp & "'"
    l_strFilter = MDAC25FilterFix(l_strFilter)
    m_rsWGActivities.Filter = l_strFilter
    If Not RecordsExist(m_rsWGActivities) Then
        m_rsWG_Assignments.Filter = ""
        Err.Raise vbObjectError + 30000, "busWorkGroupTrans.ActivityDelete", l_strErrorDesc
        'Raise error
        'Invalid change request
        'Data has been updated by someone else prior to save
        'Please Try again.
        Exit Sub
    End If

    m_rsWGActivities.Delete
    m_rsWGActivities.Update
    
    Exit Sub
    
eh:
    
  TrsRaiseError m_strCMODULENAME, "ActivityDelete"

End Sub


Private Sub ActivityInsert( _
    ByVal v_rsWorkGroupActivity As ADODB.Recordset)

    On Error GoTo eh
    
    m_rsWGActivities.AddNew
    m_rsWGActivities!ACTIVITY_CD = v_rsWorkGroupActivity!ACTIVITY_CD
    m_rsWGActivities!LVL_NUM = 50
    m_rsWGActivities!WG_NUM = m_lngWorkGroup
    m_rsWGActivities!LAST_MOD_EMP_NUM = CLng(m_strUserId)
    m_rsWGActivities.Update
 
    Exit Sub
    
eh:
    
  TrsRaiseError m_strCMODULENAME, "ActivityInsert"

End Sub

Private Function GetGLActivities( _
    ByVal v_colActivities As Collection _
    ) As ADODB.Recordset
    
    
    Dim l_intI, l_intCount As Integer
    Dim l_colActivities As Collection
    Dim l_strActivities As Variant
    Dim l_obj As dbWorkGroup.cdWorkGroup
    Dim l_strActivity As String
    
    On Error GoTo eh
    
    l_intCount = v_colActivities.Count
    Set l_colActivities = New Collection
    For l_intI = 1 To l_intCount
        l_strActivity = v_colActivities(l_intI)
        If l_strActivity = "" Then
        Else
            l_colActivities.Add l_strActivity
        End If
    Next l_intI
    
    If l_colActivities.Count = 0 Then
        Set GetGLActivities = Nothing
        Exit Function
    End If
    
    l_strActivities = ArrayFromCollection(l_colActivities, 0)
    
    Set l_obj = CreateObject("dbWorkGroup.cdWorkGroup")
    Set GetGLActivities = l_obj.GetGLActivities(l_strActivities)
    Set l_obj = Nothing
    
    Exit Function
    
    
eh:
    GetObjectContext.SetAbort
  TrsRaiseError m_strCMODULENAME, "GetGLActivities"
    
    
    
End Function

Private Sub ProcessGLActivity( _
    ByVal v_strActivity As String, _
    ByVal v_strReturnMsg As String)
    
    Dim l_strFilter As String
    
On Error GoTo eh

    m_rsGLActivities.Filter = "ACTIVITY_CD  = '" & v_strActivity & "'"

    If RecordsExist(m_rsGLActivities) Then
        '//the new accounting is not returning the activity description in the error field
        If Trim(v_strReturnMsg) <> "" Then
            m_rsGLActivities!ACTIVITY_DESC = v_strReturnMsg
        End If
    Else
        m_rsGLActivities.AddNew
        m_rsGLActivities!ACTIVITY_CD = v_strActivity
        m_rsGLActivities!ACTIVITY_DESC = v_strReturnMsg
    End If
    m_rsGLActivities.Update
    
    Exit Sub
    
eh:
    
  TrsRaiseError m_strCMODULENAME, "ProcessGLActivity"
    
    
    
    
End Sub

Private Function SaveWGDistribution() As Boolean
    Dim l_obj As dbWorkGroup.cdWorkGroupTrans

On Error GoTo eh
    m_rsWGDistribution.Filter = ""
    Set l_obj = CreateObject("dbWorkGroup.cdWorkGroupTrans")
    SaveWGDistribution = l_obj.SaveWGDistribution(m_rsWGDistribution)
    Set l_obj = Nothing
    Exit Function
    
eh:
    GetObjectContext.SetAbort
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "SaveWGDistribution"
End Function


Private Function SaveWGActivities() As Boolean
    Dim l_obj As dbWorkGroup.cdWorkGroupTrans

On Error GoTo eh
    m_rsWGActivities.Filter = ""
    Set l_obj = CreateObject("dbWorkGroup.cdWorkGroupTrans")
    SaveWGActivities = l_obj.SaveWGActivities(m_rsWGActivities)
    Set l_obj = Nothing
    Exit Function
    
eh:
    GetObjectContext.SetAbort
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "SaveWGActivities"
End Function


Private Function SaveGLActivities() As Boolean
    Dim l_obj As dbWorkGroup.cdWorkGroupTrans

On Error GoTo eh
    Set l_obj = CreateObject("dbWorkGroup.cdWorkGroupTrans")
    SaveGLActivities = l_obj.SaveGLActivities(m_rsGLActivities)
    Set l_obj = Nothing
    Exit Function
    
eh:
    GetObjectContext.SetAbort
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "SaveGLActivities"
End Function


Private Function GetWorkGroupDistribution( _
    ByVal v_lngWorkGroup As Long _
    ) As ADODB.Recordset
    
    Dim l_rs As ADODB.Recordset
    Dim l_obj As dbWorkGroup.cdWorkGroup

On Error GoTo eh


    Set l_obj = CreateObject("dbWorkGroup.cdWorkGroup")
    Set l_rs = l_obj.GetWorkGroupDistributionForUpdate(v_lngWorkGroup)
    Set GetWorkGroupDistribution = l_rs
    Exit Function
    
eh:
    GetObjectContext.SetAbort
  TrsRaiseError m_strCMODULENAME, "GetWorkGroupDistribution"
  
End Function


Private Function GetWorkGroupActivites( _
    ByVal v_lngWorkGroup As Long _
    ) As ADODB.Recordset
    
    Dim l_rs As ADODB.Recordset
    Dim l_obj As dbWorkGroup.cdWorkGroup

On Error GoTo eh


    Set l_obj = CreateObject("dbWorkGroup.cdWorkGroup")
    Set l_rs = l_obj.GetWorkGroupActivitiesForUpdate(v_lngWorkGroup)
    Set GetWorkGroupActivites = l_rs
    Exit Function
    
eh:
    GetObjectContext.SetAbort
  TrsRaiseError m_strCMODULENAME, "GetWorkGroupActivites"
    
End Function


Private Function GetPPMProject( _
    ByVal v_lngProject As Long _
    ) As ADODB.Recordset
    
 
    Dim l_rs As ADODB.Recordset
    Dim l_obj As dbCodeTables.cdCodeTables

On Error GoTo eh


    Set l_obj = CreateObject("dbCodeTables.cdCodeTables")
    Set l_rs = l_obj.GetProjectInfo(v_lngProject)
    Set GetPPMProject = l_rs
    Exit Function
    
eh:
    GetObjectContext.SetAbort
  TrsRaiseError m_strCMODULENAME, "GetPPMProject"
           
End Function


Private Function ProcessCascade( _
                    ByVal v_lTRISUser As Long, _
                    ByVal v_blnConfirmConflicts As Boolean, _
                    ByRef r_blnConfirmPending As Boolean, _
                    ByRef r_varCascadeUpdates As Variant, _
                    ByRef r_rsWGBaseScheduleUpdates As ADODB.Recordset, _
                    ByRef r_rsWGLineScheduleUpdates As ADODB.Recordset, _
                    ByRef r_rsWGAssignUpdates As ADODB.Recordset, _
                    ByRef r_rsWGAssignOtherUpdates As ADODB.Recordset, _
                    ByRef r_rsCascadeConflicts As ADODB.Recordset, _
                    ByRef r_rsExceptionSPA As ADODB.Recordset, _
                    ByRef r_rsAvailScheduleSPA As ADODB.Recordset, _
                    ByRef r_rsEmpSchedDatesSPA As ADODB.Recordset, _
                    ByRef r_rsErrors As ADODB.Recordset) _
                As Boolean

    Dim l_rsCascadeConflicts As ADODB.Recordset
    Dim l_rsExceptionSPA As ADODB.Recordset
    Dim l_rsAvailScheduleSPA As ADODB.Recordset
    Dim l_cbsCascade As busCascade.cbsCascade
    Dim l_varCascadeUpdates As Variant
    Dim l_rsCascadeErrors As ADODB.Recordset

    Dim l_colLineDateIntervals As Collection
    Dim l_colEmpDateIntervals As Collection

On Error GoTo eh

    r_varCascadeUpdates = Empty
    r_blnConfirmPending = False
    
    'still need to get the employees assigned to the workgroup/lines
' 10/7/02 RSG Moved to DetermineEmployeeIntervals to make available
' for employee trade existence check before cascade processing.
'    Set l_colLineDateIntervals = ConsolidatedCascadeWGLineIntervals
'    DetermineEmployeeAssignments l_colLineDateIntervals
'    Set l_colEmpDateIntervals = ConsolidatedCascadeEmployeeIntervals
'
'    Set r_rsEmpSchedDatesSPA = GetEmpSchedDates(l_colEmpDateIntervals)

    If r_rsEmpSchedDatesSPA Is Nothing Then
        ProcessCascade = True
        Exit Function
    End If

    Set l_cbsCascade = CreateObject("busCascade.cbsCascade")
    Set l_rsCascadeConflicts = l_cbsCascade.GetExceptionConflicts( _
                r_rsEmpSchedDatesSPA, trWorkGroupBaseScheduleLevel, _
                l_rsExceptionSPA, l_rsAvailScheduleSPA, _
                r_rsWGBaseScheduleUpdates, r_rsWGLineScheduleUpdates, _
                r_rsWGAssignUpdates, r_rsWGAssignOtherUpdates)

    If RecordsExist(l_rsCascadeConflicts) Then
        If v_blnConfirmConflicts Then
            Set r_rsCascadeConflicts = l_rsCascadeConflicts
            Set r_rsExceptionSPA = l_rsExceptionSPA
            Set r_rsAvailScheduleSPA = l_rsAvailScheduleSPA
'            In case inprocess with client
            ClearAllFilters r_rsWGBaseScheduleUpdates, r_rsWGLineScheduleUpdates, _
                    r_rsWGAssignUpdates, r_rsWGAssignOtherUpdates, _
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
    Set l_rsCascadeConflicts = Nothing
    Set l_rsExceptionSPA = Nothing
    Set l_rsAvailScheduleSPA = Nothing
    Set l_cbsCascade = Nothing
    Set l_rsCascadeErrors = Nothing
    TrsRaiseError m_strCMODULENAME, "ProcessCascade"
End Function


Private Function DetermineEmployeeAssignments( _
                    ByRef r_colWGLineDateIntervals As Collection)
' Determine the employee assignments for each of the workgroup/lines.
' Using a copy of workgroup assignments recordset retrieved from the
' database to ensure using all original assignments.  Assignment changes
' have already been captured during employee assignment processing.
' Find all intersecting employee workgroup assignments for each
' workgroup schedule changed (wg/line date interval with line number = 0).
' Find all intersecting employee line assignments for each workgroup/line
' schedule changed (wg/line date interval with line number > 0).

    Dim l_cpuLineDateInterval As cpuDateInterval
    Dim l_strFilter As String
    Dim l_intIntersection As Integer

On Error GoTo eh
    
    For Each l_cpuLineDateInterval In r_colWGLineDateIntervals
        If l_cpuLineDateInterval.LineNumber = 0 Then
            l_strFilter = _
                "WG_NUM = " & l_cpuLineDateInterval.WorkgroupNumber & _
                " AND EFF_END_DATE >= #" & Format(l_cpuLineDateInterval.StartDate, "mm/dd/yyyy") & "#" & _
                " AND EFF_START_DATE <= #" & Format(l_cpuLineDateInterval.EndDate, "mm/dd/yyyy") & "#"
            l_strFilter = MDAC25FilterFix(l_strFilter)
            m_rsWG_Assignments_Copy.Filter = l_strFilter
            If RecordsExist(m_rsWG_Assignments_Copy) Then
                Do Until m_rsWG_Assignments_Copy.EOF
                    l_intIntersection = DateIntersection( _
                        l_cpuLineDateInterval.StartDate, _
                        l_cpuLineDateInterval.EndDate, _
                        m_rsWG_Assignments_Copy!EFF_START_DATE, _
                        m_rsWG_Assignments_Copy!EFF_END_DATE)
                    Select Case l_intIntersection
                        Case 1
                            AppendEmployeeCascadeDates "C", _
                                m_rsWG_Assignments_Copy!EMP_NUM, _
                                m_rsWG_Assignments_Copy!EFF_START_DATE, _
                                m_rsWG_Assignments_Copy!EFF_END_DATE
                        Case 2
                            AppendEmployeeCascadeDates "C", _
                                m_rsWG_Assignments_Copy!EMP_NUM, _
                                l_cpuLineDateInterval.StartDate, _
                                l_cpuLineDateInterval.EndDate
                        Case 3
                            AppendEmployeeCascadeDates "C", _
                                m_rsWG_Assignments_Copy!EMP_NUM, _
                                m_rsWG_Assignments_Copy!EFF_START_DATE, _
                                l_cpuLineDateInterval.EndDate
                        Case 4
                            AppendEmployeeCascadeDates "C", _
                                m_rsWG_Assignments_Copy!EMP_NUM, _
                                l_cpuLineDateInterval.StartDate, _
                                m_rsWG_Assignments_Copy!EFF_END_DATE
                    End Select
                    m_rsWG_Assignments_Copy.MoveNext
                Loop
            End If
        End If
        
        If l_cpuLineDateInterval.LineNumber > 0 Then
            l_strFilter = _
                "WG_NUM = " & l_cpuLineDateInterval.WorkgroupNumber & _
                " AND LINE_NUM = " & l_cpuLineDateInterval.LineNumber & _
                " AND END_DATE >= #" & Format(l_cpuLineDateInterval.StartDate, "mm/dd/yyyy") & "#" & _
                " AND START_DATE <= #" & Format(l_cpuLineDateInterval.EndDate, "mm/dd/yyyy") & "#"
            l_strFilter = MDAC25FilterFix(l_strFilter)
            m_rsWGDemandLineAssignments.Filter = l_strFilter
            If RecordsExist(m_rsWGDemandLineAssignments) Then
                Do Until m_rsWGDemandLineAssignments.EOF
                    l_intIntersection = DateIntersection( _
                        l_cpuLineDateInterval.StartDate, _
                        l_cpuLineDateInterval.EndDate, _
                        m_rsWGDemandLineAssignments!START_DATE, _
                        m_rsWGDemandLineAssignments!END_DATE)
                    Select Case l_intIntersection
                        Case 1
                            AppendEmployeeCascadeDates "C", _
                                m_rsWGDemandLineAssignments!EMP_NUM, _
                                m_rsWGDemandLineAssignments!START_DATE, _
                                m_rsWGDemandLineAssignments!END_DATE
                        Case 2
                            AppendEmployeeCascadeDates "C", _
                                m_rsWGDemandLineAssignments!EMP_NUM, _
                                l_cpuLineDateInterval.StartDate, _
                                l_cpuLineDateInterval.EndDate
                        Case 3
                            AppendEmployeeCascadeDates "C", _
                                m_rsWGDemandLineAssignments!EMP_NUM, _
                                m_rsWGDemandLineAssignments!START_DATE, _
                                l_cpuLineDateInterval.EndDate
                        Case 4
                            AppendEmployeeCascadeDates "C", _
                                m_rsWGDemandLineAssignments!EMP_NUM, _
                                l_cpuLineDateInterval.StartDate, _
                                m_rsWGDemandLineAssignments!END_DATE
                    End Select
                    m_rsWGDemandLineAssignments.MoveNext
                Loop
            End If
        End If
    Next

    m_rsWG_Assignments_Copy.Filter = 0
    If Not m_rsWGDemandLineAssignments Is Nothing Then
        m_rsWGDemandLineAssignments.Filter = 0
    End If
    Set l_cpuLineDateInterval = Nothing
    Exit Function

eh:
    Set l_cpuLineDateInterval = Nothing
    TrsRaiseError m_strCMODULENAME, "DetermineEmployeeAssignments"
End Function


Private Function GetEmpSchedDates(ByRef r_colDateIntervals As Collection) As ADODB.Recordset
    Dim l_rsEmpDates As ADODB.Recordset
    Dim l_cpuDateInterval As cpuDateInterval

On Error GoTo eh
   
    If r_colDateIntervals.Count = 0 Then
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

    For Each l_cpuDateInterval In r_colDateIntervals
        l_rsEmpDates.AddNew
        l_rsEmpDates!EMP_NUM = l_cpuDateInterval.EmployeeNumber
        l_rsEmpDates!START_DATE = l_cpuDateInterval.StartDate
        l_rsEmpDates!END_DATE = l_cpuDateInterval.EndDate
        l_rsEmpDates.Update
    Next
    
    Set GetEmpSchedDates = l_rsEmpDates
    Set l_rsEmpDates = Nothing
    Set l_cpuDateInterval = Nothing
    Exit Function

eh:
    Set l_rsEmpDates = Nothing
    Set l_cpuDateInterval = Nothing
    TrsRaiseError m_strCMODULENAME, "GetEmpSchedDates"
End Function


Private Function ActualTimeExists( _
    ByRef r_rsEmpSchedDates As ADODB.Recordset, _
    ByRef r_rsErrors As ADODB.Recordset, _
    ByRef r_rsWGLineSched As ADODB.Recordset _
    ) As Boolean
'    Append error messages for any actual time found.  In some cases, this may
'    result in hundreds of messages.  If this becomes a problem we implement a
'    limit on the message append loop.
    Dim l_blnExists As Boolean
    Dim l_cdTimecard As dbTimeCard.cdTimeCard
    Dim l_rsActualTime As ADODB.Recordset
    Dim l_varEmployees As Variant
    Dim l_datStart As Date
    Dim l_datEnd As Date
    Dim l_strFilter As String
    Dim l_strErrMsg As String
    Dim l_lngEmpNum As Long
    Dim l_datChargeDate As Date
    Dim l_lngWgNum As Long
    Dim l_lngLineNum As Long

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
 'FRED
    ActualTimeExistsGetState r_rsWGLineSched
    Do Until r_rsEmpSchedDates.EOF
        l_strFilter = _
            "EMP_NUM = " & r_rsEmpSchedDates!EMP_NUM & _
            " AND CHARGE_DATE >= #" & Format(r_rsEmpSchedDates!START_DATE, "mm/dd/yyyy") & "#" & _
            " AND CHARGE_DATE <= #" & Format(r_rsEmpSchedDates!END_DATE, "mm/dd/yyyy") & "#"
        l_strFilter = MDAC25FilterFix(l_strFilter)
        l_rsActualTime.Filter = l_strFilter
        If RecordsExist(l_rsActualTime) Then
            l_lngEmpNum = l_rsActualTime!EMP_NUM
            l_datChargeDate = l_rsActualTime!CHARGE_DATE
            Do Until l_rsActualTime.EOF
                If Not IsEmployeeOnALine(l_lngEmpNum, l_datChargeDate, _
                        l_lngWgNum, l_lngLineNum) Then
                     l_blnExists = True
                    ActualTimeExistsMsg "E", r_rsEmpSchedDates, l_rsActualTime, r_rsErrors
                Else
                    If LineScheduleChanged(l_lngWgNum, l_lngLineNum, _
                            l_datChargeDate, l_datChargeDate) Then
                         l_blnExists = True
                        ActualTimeExistsMsg "E", r_rsEmpSchedDates, l_rsActualTime, r_rsErrors
                    Else
                        ActualTimeExistsMsg "W", r_rsEmpSchedDates, l_rsActualTime, r_rsErrors
                    End If
                End If
                l_rsActualTime.MoveNext
            Loop
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

' 10/7/02 RSG Replaced with LineTradeExists and EmployeeTradeExists,
' both are workgroup maintenance validations, not cascade.
'Private Function TradeExists( _
'                    ByRef r_rsEmpSchedDates As ADODB.Recordset, _
'                    ByRef r_rsErrors As ADODB.Recordset) As Boolean
'    Dim l_blnExists As Boolean
'    Dim l_cdSchedule As dbSchedule.cdSchedule
'    Dim l_rsEmpTrades As ADODB.Recordset
'    Dim l_varEmployees As Variant
'    Dim l_datStart As Date
'    Dim l_datEnd As Date
'    Dim l_strFilter As String
'    Dim l_strErrMsg As String
'
'On Error GoTo eh
'
'    l_blnExists = False
'    DetermineGetActualTimeScope r_rsEmpSchedDates, l_varEmployees, l_datStart, l_datEnd
'    If IsEmpty(l_varEmployees) Then
'        TradeExists = False
'        Exit Function
'    End If
'
'    Set l_cdSchedule = CreateObject("dbSchedule.cdSchedule")
'
'    Set l_rsEmpTrades = _
'        l_cdSchedule.GetEmployeeTrades(l_varEmployees, l_datStart, l_datEnd)
'    Set l_cdSchedule = Nothing
'
'    If l_rsEmpTrades.EOF Then
'        TradeExists = False
'        Set l_rsEmpTrades = Nothing
'        Exit Function
'    End If
'
'    Do Until r_rsEmpSchedDates.EOF
'        l_strFilter = _
'            "EMP_NUM = " & r_rsEmpSchedDates!EMP_NUM & _
'            " AND CHARGE_DATE >= #" & Format(r_rsEmpSchedDates!START_DATE, "mm/dd/yyyy") & "#" & _
'            " AND CHARGE_DATE <= #" & Format(r_rsEmpSchedDates!END_DATE, "mm/dd/yyyy") & "#"
'        l_strFilter = MDAC25FilterFix(l_strFilter)
'        l_rsEmpTrades.Filter = l_strFilter
'        If RecordsExist(l_rsEmpTrades) Then
'            Do Until l_rsEmpTrades.EOF
'                l_strErrMsg = _
'                    "Employee " & CStr(l_rsEmpTrades!EMP_NUM) & _
'                    "[" & Trim(l_rsEmpTrades!EMP_NAME) & "] has existing trade on " & _
'                    Format(l_rsEmpTrades!CHARGE_DATE, "mm/dd/yyyy") & "."
'                AppendBusinessError _
'                    r_rsErrors, _
'                    trErrbusTradeExists, _
'                    mstrCBUSINESSERR, _
'                    l_strErrMsg, _
'                    m_strCMODULENAME, _
'                    TrType_WGM_EMP_ACT, _
'                    , _
'                    , _
'                    CWGMAINTFIELDNAMES, _
'                    CStr(m_lngWorkGroup), _
'                    Format(r_rsEmpSchedDates!START_DATE, "mm/dd/yyyy"), _
'                    "0", _
'                    CStr(r_rsEmpSchedDates!EMP_NUM)
'                l_rsEmpTrades.MoveNext
'            Loop
'            l_blnExists = True
'        End If
'        r_rsEmpSchedDates.MoveNext
'    Loop
'    r_rsEmpSchedDates.MoveFirst
'
'    TradeExists = l_blnExists
'    Set l_rsEmpTrades = Nothing
'
'    Exit Function
'
'eh:
'    Set l_cdSchedule = Nothing
'    Set l_rsEmpTrades = Nothing
'    TrsRaiseError m_strCMODULENAME, "TradeExists"
'End Function


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


Private Sub AppendWGLineCascadeDates( _
                ByVal v_strAction As String, _
                ByVal v_lngWgNum As Long, _
                ByVal v_lngLineNum As Long, _
                ByVal v_datStartDate As Date, _
                ByVal v_datEndDate As Date)
' Simply add the date records, without attempting to adjust as added.
' Consolidate all the date intervals changed at the end of the process by
' workgroupNumber and lineNumber.
'
    Dim l_objDateInterval As cpuDateInterval

On Error GoTo eh
    Set l_objDateInterval = New cpuDateInterval
    l_objDateInterval.WorkgroupNumber = v_lngWgNum
    l_objDateInterval.LineNumber = v_lngLineNum
    l_objDateInterval.StartDate = v_datStartDate
    l_objDateInterval.EndDate = v_datEndDate
    GetCascadeWGLineIntervals.Add l_objDateInterval
    
    Set l_objDateInterval = Nothing
    Exit Sub
    
eh:
    Set l_objDateInterval = Nothing
    TrsRaiseError m_strCMODULENAME, "AppendWGLineCascadeDates"
End Sub


Private Sub AppendEmployeeCascadeDates( _
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
    GetCascadeEmployeeIntervals.Add l_objDateInterval
    
    Set l_objDateInterval = Nothing
    Exit Sub
    
eh:
    Set l_objDateInterval = Nothing
    TrsRaiseError m_strCMODULENAME, "AppendEmployeeCascadeDates"
End Sub


Private Function GetCascadeWGLineIntervals() As Collection
' Answer a collection of variant arrays that includes:
' wgNum, lineNum, startDate, endDate

On Error GoTo eh
    If m_colCascadeLineIntervals Is Nothing Then
        Set m_colCascadeLineIntervals = New Collection
    End If
    Set GetCascadeWGLineIntervals = m_colCascadeLineIntervals
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "GetCascadeWGLineIntervals"
End Function


Private Function GetCascadeEmployeeIntervals() As Collection
' Answer a collection of variant arrays that includes:
' empNum, startDate, endDate

On Error GoTo eh
    If m_colCascadeEmpIntervals Is Nothing Then
        Set m_colCascadeEmpIntervals = New Collection
    End If
    Set GetCascadeEmployeeIntervals = m_colCascadeEmpIntervals
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "GetCascadeEmployeeIntervals"
End Function


Private Function ConsolidatedCascadeWGLineIntervals() As Collection
' Consolidate/join all the raw date updates by workgroup/line number.
' These remaining date intervals will be used for cascade request.
    Dim l_setWgLines As cuSet
    Dim l_varWgLine As Variant
    Dim l_varWgLineArray As Variant
    Dim l_colRawIntervals As Collection
    Dim l_colConIntervals As Collection
    Dim l_objInterval As cpuDateInterval

On Error GoTo eh
    Set l_colConIntervals = New Collection
' Determine unique workgroup/lines to be processed.
    Set l_setWgLines = New cuSet
    For Each l_objInterval In GetCascadeWGLineIntervals
        l_setWgLines.Add _
            Join(Array(l_objInterval.WorkgroupNumber, l_objInterval.LineNumber), ";")
    Next

' For each unique workgroup/line,
'   copy the raw interval objects to a separate collection;
'   execute consolidatio/join process on filtered collection;
'   copy remaining consolidated interval objects to new collection.
    For Each l_varWgLine In l_setWgLines
        l_varWgLineArray = Split(l_varWgLine, ";")
        Set l_colRawIntervals = FilterCascadeWGLineIntervals(l_varWgLineArray(0), l_varWgLineArray(1))
        JoinIntervals l_colRawIntervals
        For Each l_objInterval In l_colRawIntervals
            l_colConIntervals.Add l_objInterval
        Next
    Next
    
    Set ConsolidatedCascadeWGLineIntervals = l_colConIntervals
    
    Set l_setWgLines = Nothing
    Set l_colRawIntervals = Nothing
    Set l_colConIntervals = Nothing
    Set l_objInterval = Nothing
    
    Exit Function

eh:
    Set l_setWgLines = Nothing
    Set l_colRawIntervals = Nothing
    Set l_colConIntervals = Nothing
    Set l_objInterval = Nothing
    TrsRaiseError m_strCMODULENAME, "ConsolidatedCascadeWGLineIntervals"
End Function


Private Function ConsolidatedCascadeEmployeeIntervals() As Collection
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
    For Each l_objInterval In GetCascadeEmployeeIntervals
        l_setEmp.Add l_objInterval.EmployeeNumber
    Next

'    For each unique employee,
'    copy the raw interval objects to a separate collection;
'    execute consolidation/join process on filtered collection;
'    copy remaining consolidated interval objects to new collection.
    For Each l_varEmp In l_setEmp
        Set l_colRawIntervals = FilterCascadeEmployeeIntervals(l_varEmp)
        JoinIntervals l_colRawIntervals
        For Each l_objInterval In l_colRawIntervals
            l_colConIntervals.Add l_objInterval
        Next
    Next
    
    Set ConsolidatedCascadeEmployeeIntervals = l_colConIntervals
    
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
    TrsRaiseError m_strCMODULENAME, "ConsolidatedCascadeEmployeeIntervals"
End Function


Private Function FilterCascadeWGLineIntervals( _
            ByVal v_lngWgNum As Long, _
            ByVal v_lngLineNum As Long) As Collection
' Answer a new collection with references that match workgroup and line number.
    Dim l_objInterval As cpuDateInterval
    Dim l_colFilteredIntervals As Collection

On Error GoTo eh
    Set l_colFilteredIntervals = New Collection
    For Each l_objInterval In GetCascadeWGLineIntervals
        If l_objInterval.WorkgroupNumber = v_lngWgNum And _
            l_objInterval.LineNumber = v_lngLineNum Then
            l_colFilteredIntervals.Add l_objInterval
        End If
    Next
    Set FilterCascadeWGLineIntervals = l_colFilteredIntervals
    Set l_colFilteredIntervals = Nothing
    Set l_objInterval = Nothing
    Exit Function

eh:
    Set l_colFilteredIntervals = Nothing
    Set l_objInterval = Nothing
    TrsRaiseError m_strCMODULENAME, "FilterCascadeWGLineIntervals"
End Function


Private Function FilterCascadeEmployeeIntervals( _
                    ByVal v_lngEmpNum As Long) _
                As Collection
' Answer a new collection with references that match employee number.
    Dim l_objInterval As cpuDateInterval
    Dim l_colFilteredIntervals As Collection

On Error GoTo eh
    Set l_colFilteredIntervals = New Collection
    For Each l_objInterval In GetCascadeEmployeeIntervals
        If l_objInterval.EmployeeNumber = v_lngEmpNum Then
            l_colFilteredIntervals.Add l_objInterval
        End If
    Next
    Set FilterCascadeEmployeeIntervals = l_colFilteredIntervals
    Set l_colFilteredIntervals = Nothing
    Set l_objInterval = Nothing
    Exit Function

eh:
    Set l_colFilteredIntervals = Nothing
    Set l_objInterval = Nothing
    TrsRaiseError m_strCMODULENAME, "FilterCascadeEmployeeIntervals"
End Function


Private Function JoinIntervals( _
            ByRef r_colIntervals As Collection)
'    Perform this for each workgroup/line or employee
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
'                                l_objInnerInterval.WorkgroupNumber & ":" & l_objInnerInterval.LineNumber & ":" & _
'                                l_objInnerInterval.StartDate & ":" & l_objInnerInterval.EndDate
                            l_colIntervals.Remove j
                            blnDone = False
                            Exit For
                        Case 2
'                            Debug.Print "Removing outer2: " & _
'                                l_objOuterInterval.WorkgroupNumber & ":" & l_objOuterInterval.LineNumber & ":" & _
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
'                                l_objInnerInterval.WorkgroupNumber & ":" & l_objInnerInterval.LineNumber & ":" & _
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

'Method HighManHours Not Needed in TRIS 5
'''''Private Function HighManHours() As Boolean
'''''
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
'''''    If m_rsWG_Assignments Is Nothing Then
'''''        HighManHours = True
'''''        Exit Function
'''''    End If
'''''
'''''    m_rsWG_Assignments.Filter = ""
'''''    If Not RecordsExist(m_rsWG_Assignments) Then
'''''        HighManHours = True
'''''        Exit Function
'''''    End If
'''''
'''''    Set l_obj = CreateObject("busOvertimeEqualizer.cbsOTEqualizer")
'''''    HighManHours = l_obj.HighmanHoursWorkGroupAssignments( _
'''''            m_rsWG_Assignments, m_lngUserNum)
'''''    Set l_obj = Nothing
'''''    Exit Function
'''''
'''''eh:
'''''    Set l_obj = Nothing
'''''    TrsRaiseError m_strCMODULENAME, "HighManHours"
'''''End Function


Private Function ScheduleChangePaymentExists( _
                    ByRef r_rsEmpSchedDates As ADODB.Recordset, _
                    ByRef r_rsErrors As ADODB.Recordset, _
                    ByRef r_rsWGBaseSched As ADODB.Recordset, _
                    ByRef r_rsWGLineSched As ADODB.Recordset, _
                    ByRef r_rsWGAssignments As ADODB.Recordset, _
                    ByRef r_rsWGOtherAssign As ADODB.Recordset _
                ) As Boolean
' Check for any schedule change payments during the scope of the workgroup schedule change.
' If a payment is found on an off-day(ORPHAN_IND= Y), just warn the user so they can take
' appropriate action, e.g. move the payment to the next scheduled work day.
' Answer TRUE if a payment exists for an off-day, otherwise FALSE.
    Dim l_blnExists As Boolean
    Dim l_cbsSchedChgPay As busSchedChgPay.cbsSchedChgPay
    Dim l_rsSchedChgPay As ADODB.Recordset
                    
    Dim l_rsWGBaseSched As ADODB.Recordset
    Dim l_rsWGLineSched As ADODB.Recordset
    Dim l_rsWGAssignments As ADODB.Recordset
    Dim l_rsWGOtherAssign As ADODB.Recordset
                    
On Error GoTo eh
                    
    l_blnExists = False
    If IsEmpty(r_rsEmpSchedDates) Then
        ScheduleChangePaymentExists = False
        Exit Function
    End If
    
    Set l_rsWGBaseSched = PendingRecords(r_rsWGBaseSched)
    Set l_rsWGLineSched = PendingRecords(r_rsWGLineSched)
    Set l_rsWGAssignments = PendingRecords(r_rsWGAssignments)
    Set l_rsWGOtherAssign = PendingRecords(r_rsWGOtherAssign)
                          
    Set l_cbsSchedChgPay = CreateObject("busSchedChgPay.cbsSchedChgPay")
    Set l_rsSchedChgPay = l_cbsSchedChgPay.GetSchedChgPayPending(r_rsEmpSchedDates, _
            l_rsWGBaseSched, l_rsWGLineSched, l_rsWGAssignments, l_rsWGOtherAssign)
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
                    m_lngWorkGroup, _
                    r_rsErrors
        End If
        l_rsSchedChgPay.MoveNext
    Loop

    ScheduleChangePaymentExists = l_blnExists
    Set l_rsSchedChgPay = Nothing
    Set l_rsWGBaseSched = Nothing
    Set l_rsWGLineSched = Nothing
    Set l_rsWGAssignments = Nothing
    Set l_rsWGOtherAssign = Nothing
    Exit Function
    
eh:
    Set l_cbsSchedChgPay = Nothing
    Set l_rsSchedChgPay = Nothing
    Set l_rsWGBaseSched = Nothing
    Set l_rsWGLineSched = Nothing
    Set l_rsWGAssignments = Nothing
    Set l_rsWGOtherAssign = Nothing
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
                    ByVal v_lngWgNum As Long, _
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
        TrType_WGM_SCHED, _
        , , _
        CDEMANDSCHEDFIELDNAMES, _
        CStr(v_lngWgNum), _
        Format(v_datChargeDate, "mm/dd/yyyy"), _
        "0", _
        CStr(v_lngEmpNum)

    Exit Function

eh:
    Set l_colDescArgs = Nothing
    TrsRaiseError m_strCMODULENAME, "AppendWarningScheduleChangePayment"
End Function

Private Function LineScheduleChanged( _
    ByVal v_lngWgNum As Long, _
    ByVal v_lngLineNum As Long, _
    ByVal v_datStart As Date, _
    ByVal v_datEnd As Date _
    ) As Boolean

    Dim l_strFilter As String
    Dim l_intSchedCd As Integer
    Dim l_intRotCd As Integer
    Dim l_intRotWeek As Integer
    Dim l_intOrigSchedCd As Integer
    Dim l_intOrigRotCd As Integer
    Dim l_intOrigRotWeek As Integer
    Dim l_datOrigStartDate As Date
    Dim l_datStartDate As Date
    Dim l_intNewRotWeek As Integer
    
    On Error GoTo eh
    
    l_strFilter = _
        "WG_NUM = " & str(v_lngWgNum) & _
        " AND LINE_NUM = " & str(v_lngLineNum) & _
        " AND END_DATE >= #" & Format(v_datStart, "mm/dd/yyyy") & "#" & _
        " AND START_DATE <= #" & Format(v_datEnd, "mm/dd/yyyy") & "#"
    l_strFilter = MDAC25FilterFix(l_strFilter)
    
    m_rsWGDemandLineSchedules.Filter = l_strFilter
    
    If Not RecordsExist(m_rsWGDemandLineSchedules) Then
        m_rsWGDemandLineSchedules.Filter = ""
        LineScheduleChanged = True
        Exit Function
    End If
    
    m_rsHoldWGDemandLineSchedules.Filter = l_strFilter
    If Not RecordsExist(m_rsHoldWGDemandLineSchedules) Then
        m_rsHoldWGDemandLineSchedules.Filter = ""
        LineScheduleChanged = True
        Exit Function
    End If
    
    l_intSchedCd = m_rsWGDemandLineSchedules!SCHED_CD
    l_intRotCd = m_rsWGDemandLineSchedules!ROT_CD
    l_intRotWeek = m_rsWGDemandLineSchedules!ROT_WEEK
    l_datStartDate = m_rsWGDemandLineSchedules!START_DATE
    
    l_intOrigSchedCd = m_rsHoldWGDemandLineSchedules!SCHED_CD
    l_intOrigRotCd = m_rsHoldWGDemandLineSchedules!ROT_CD
    l_intOrigRotWeek = m_rsHoldWGDemandLineSchedules!ROT_WEEK
    l_datOrigStartDate = m_rsHoldWGDemandLineSchedules!START_DATE
    
' Please note that SchedCd and RotCd are mutully exclusive
' and that edit has been performed prior to this code

'The demand schedule is based on schedule code only
    If l_intSchedCd <> 0 Then
        If l_intSchedCd = l_intOrigSchedCd Then
            LineScheduleChanged = False
        Else
            LineScheduleChanged = True
        End If
        m_rsWGDemandLineSchedules.Filter = ""
        m_rsHoldWGDemandLineSchedules.Filter = ""
        Exit Function
    End If
    
    
    If l_intRotCd <> l_intOrigRotCd Then
        LineScheduleChanged = True
        m_rsWGDemandLineSchedules.Filter = ""
        m_rsHoldWGDemandLineSchedules.Filter = ""
        Exit Function
    End If
    
'Now determine if the schedule will pick up were the other one left off
    
    l_intNewRotWeek = DetermineNewRotationWeek( _
                            l_datStartDate, _
                            l_intRotCd, _
                            l_datOrigStartDate, _
                            l_intOrigRotWeek)
    If l_intNewRotWeek = l_intRotWeek Then
        LineScheduleChanged = False
    Else
        LineScheduleChanged = True
    End If
    
    m_rsWGDemandLineSchedules.Filter = ""
    m_rsHoldWGDemandLineSchedules.Filter = ""

    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "LineScheduleChanged"
End Function




Private Sub ActualTimeExistsGetState( _
    ByVal v_rsDemandLineSchedules As ADODB.Recordset)
    
    Dim l_datMinEffDate As Date
    Dim l_datMaxEffDate As Date
    Dim l_varLineNums As Variant
    
    On Error GoTo eh
    
    If Not m_blnDemandWorkGroup Then
        Set m_rsHoldWGDemandLineSchedules = Nothing
        Set m_rsWGDemandLineAssignments = Nothing
        Exit Sub
    End If
    
    If RecordsExist(v_rsDemandLineSchedules) Then
        v_rsDemandLineSchedules.Filter = adFilterPendingRecords
    End If
    If Not RecordsExist(v_rsDemandLineSchedules) Then
        Set m_rsHoldWGDemandLineSchedules = Nothing
        Set m_rsWGDemandLineAssignments = Nothing
        If Not v_rsDemandLineSchedules Is Nothing Then
            v_rsDemandLineSchedules.Filter = ""
        End If
        Exit Sub
    End If
    
    ActualTimeExistsDetermineScope _
          v_rsDemandLineSchedules, l_datMinEffDate, l_datMaxEffDate, l_varLineNums
    
    Set m_rsHoldWGDemandLineSchedules = _
        GetDemandLineSchedules( _
            m_lngWorkGroup, l_datMinEffDate, l_datMaxEffDate, l_varLineNums)
    l_datMaxEffDate = DetermineGetAssignmentMaxEnd(m_rsHoldWGDemandLineSchedules)
    Set m_rsWGDemandLineAssignments = _
            GetDemandLineAssignments(m_lngWorkGroup, l_datMinEffDate, l_datMaxEffDate)

    v_rsDemandLineSchedules.Filter = ""
    
    Exit Sub

eh:

    TrsRaiseError m_strCMODULENAME, "ActualTimeExistsGetState"
End Sub

Private Sub ActualTimeExistsDetermineScope( _
    ByVal v_rsDemandLineSchedules As ADODB.Recordset, _
    ByRef r_datMinEffDate As Date, _
    ByRef r_datMaxEffDate As Date, _
    ByRef r_varLineNums As Variant)
    
    Dim l_datMinEffDate As Date
    Dim l_colLines As cuSet
    Dim l_datStart As Date
    Dim l_intLine As Integer

    On Error GoTo eh
    
    Set l_colLines = New cuSet
    l_datMinEffDate = CDate("9999-12-31")
    r_datMaxEffDate = CDate("9999-12-31")
    v_rsDemandLineSchedules.MoveFirst
    While Not v_rsDemandLineSchedules.EOF
        Select Case v_rsDemandLineSchedules.Status
            Case adRecNew
                l_datStart = v_rsDemandLineSchedules("START_DATE").Value
                If l_datMinEffDate > l_datStart Then
                    l_datMinEffDate = l_datStart
                End If
                l_intLine = v_rsDemandLineSchedules("LINE_NUM").Value
                l_colLines.Add l_intLine
            Case adRecModified
                l_datStart = v_rsDemandLineSchedules("START_DATE").Value
                If l_datMinEffDate > l_datStart Then
                    l_datMinEffDate = l_datStart
                End If
                l_intLine = v_rsDemandLineSchedules("LINE_NUM").Value
                l_colLines.Add l_intLine
                l_datStart = v_rsDemandLineSchedules("START_DATE").OriginalValue
                If l_datMinEffDate > l_datStart Then
                    l_datMinEffDate = l_datStart
                End If
                l_intLine = v_rsDemandLineSchedules("LINE_NUM").OriginalValue
                l_colLines.Add l_intLine
            Case adRecDeleted
                l_datStart = v_rsDemandLineSchedules("START_DATE").OriginalValue
                If l_datMinEffDate > l_datStart Then
                    l_datMinEffDate = l_datStart
                End If
                l_intLine = v_rsDemandLineSchedules("LINE_NUM").OriginalValue
                l_colLines.Add l_intLine
        End Select
        v_rsDemandLineSchedules.MoveNext
    Wend
    
    r_datMinEffDate = l_datMinEffDate
    r_varLineNums = ArrayFromCollection(l_colLines.Contents)
    
    Exit Sub

eh:

    TrsRaiseError m_strCMODULENAME, "ActualTimeExistsDetermineScope"
End Sub
        
Private Sub ActualTimeExistsMsg( _
    ByVal v_strErrorType As String, _
    ByVal v_rsEmpSchedDates As ADODB.Recordset, _
    ByVal v_rsActualTime As ADODB.Recordset, _
    ByRef r_rsErrors As ADODB.Recordset)
    
    Dim l_strErrMsg As String
    Dim l_lngErrorNum As Long
    Dim l_strSystemType As String
    
    On Error GoTo eh
    
    If v_strErrorType = "E" Then
        l_strErrMsg = _
            "Employee " & CStr(v_rsActualTime!EMP_NUM) & _
            "[" & Trim(v_rsActualTime!EMP_NAME) & "] has existing actual time on " & _
            Format(v_rsActualTime!CHARGE_DATE, "mm/dd/yyyy") & "." & _
            "  You must delete the time to change the schedule."
        l_lngErrorNum = trErrbusActualTimeExists
        l_strSystemType = mstrCBUSINESSERR
    Else
        l_strErrMsg = _
            "Please verify job number on actual time.  Employee " & CStr(v_rsActualTime!EMP_NUM) & _
            "[" & Trim(v_rsActualTime!EMP_NAME) & "] has existing actual time on " & _
            Format(v_rsActualTime!CHARGE_DATE, "mm/dd/yyyy") & "."
        l_lngErrorNum = 20000
        l_strSystemType = mstrCBUSINESSWAR
    End If
    
    AppendBusinessError _
        r_rsErrors, _
        l_lngErrorNum, _
        l_strSystemType, _
        l_strErrMsg, _
        m_strCMODULENAME, _
        TrType_WGM_EMP_ACT, _
        , _
        , _
        CWGMAINTFIELDNAMES, _
        CStr(m_lngWorkGroup), _
        Format(v_rsEmpSchedDates!START_DATE, "mm/dd/yyyy"), _
        "0", _
        CStr(v_rsEmpSchedDates!EMP_NUM)

    Exit Sub

eh:

    TrsRaiseError m_strCMODULENAME, "ActualTimeExistsDetermineScope"
End Sub

Private Function IsEmployeeOnALine( _
    ByVal v_lngEmpNum As Long, _
    ByVal v_datChargeDate As Date, _
    ByRef r_lngWGNum As Long, _
    ByRef r_lngLineNum As Long _
    ) As Boolean
    
    On Error GoTo eh
    
    Dim l_strFilter As String
    
    If Not RecordsExist(m_rsWGDemandLineAssignments) Then
        IsEmployeeOnALine = False
        Exit Function
    End If
    
    l_strFilter = "EMP_NUM = " & str(v_lngEmpNum) & _
        " AND START_DATE <= " & str(v_datChargeDate) & _
        " AND END_DATE >= " & str(v_datChargeDate)
    l_strFilter = MDAC25FilterFix(l_strFilter)
    
    m_rsWGDemandLineAssignments.Filter = l_strFilter
    
    If RecordsExist(m_rsWGDemandLineAssignments) Then
        r_lngWGNum = m_rsWGDemandLineAssignments!WG_NUM
        r_lngLineNum = m_rsWGDemandLineAssignments!LINE_NUM
        IsEmployeeOnALine = True
    Else
        r_lngWGNum = 0
        r_lngLineNum = 0
        IsEmployeeOnALine = False
    End If
    
    m_rsWGDemandLineAssignments.Filter = ""
    Exit Function

eh:

    TrsRaiseError m_strCMODULENAME, "IsEmployeeOnALine"
    
End Function


Private Function GetWGSequenceType_DOJM( _
    ByVal v_strDept As String _
    ) As String
    
    Dim l_strSeqType As String
    Dim l_rsDOJMSeqTypes As ADODB.Recordset
    
    On Error GoTo eh
    
    
    Set l_rsDOJMSeqTypes = GetDOJMSeqTypes()
    
    If Not RecordsExist(l_rsDOJMSeqTypes) Then
        GetWGSequenceType_DOJM = "DO"
        Exit Function
    End If
    
    l_rsDOJMSeqTypes.Filter = "VALUE_PART1 = '" & Trim(v_strDept) & "'"
    If Not RecordsExist(l_rsDOJMSeqTypes) Then
        GetWGSequenceType_DOJM = "DO"
        Exit Function
    End If
    
    GetWGSequenceType_DOJM = l_rsDOJMSeqTypes!VALUE_PART2
    
    Set l_rsDOJMSeqTypes = Nothing
    
    Exit Function

eh:
    Set l_rsDOJMSeqTypes = Nothing
    TrsRaiseError m_strCMODULENAME, "GetWGSequenceType_DOJM"
End Function


Private Function GetDOJMSeqTypes() As ADODB.Recordset

    Dim l_obj As dbRules.cdRules
    
'RULE0030

    On Error GoTo eh
    
    Set l_obj = CreateObject("dbRules.cdRules")
    Set GetDOJMSeqTypes = l_obj.GetRule(30)
    Set l_obj = Nothing
    
    Exit Function

eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "GetDOJMSeqTypes"
End Function


'''''Private Sub CheckForJobChanges()
'''''
''''''JIMB 11/14/2002
''''''Leaving this code in eventhough it is dead code before it is implemented
''''''Changed the design to do this in the check for actual time.
'''''
''''''PTS 1092
''''''Users want a BRV warning when the job number changes on a
'''''' line and the employee assigned to the line has actual time
'''''
'''''    Dim l_blnExists As Boolean
'''''    Dim l_rsActualTime As ADODB.Recordset
'''''    Dim l_rsEmpSchedDates As ADODB.Recordset
'''''    Dim l_datStart As Date
'''''    Dim l_datEnd As Date
'''''    Dim l_strFilter As String
'''''    Dim l_strErrMsg As String
'''''    Dim l_colLineDateIntervals As Collection
'''''    Dim l_varWorkgroups As Variant
'''''    Dim l_varEmployees As Variant
'''''    Dim l_varLines As Variant
'''''    Dim l_cdTimecard As dbTimeCard.cdTimeCard
'''''    Dim l_cpuDateInterval As cpuDateInterval
'''''
'''''On Error GoTo eh
'''''
'''''    l_blnExists = False
'''''    CheckForJobChangesApplySched
'''''    Set l_rsEmpSchedDates = DetermineEmployeeIntervals
'''''    DetermineGetActualTimeScope l_rsEmpSchedDates, l_varEmployees, l_datStart, l_datEnd
'''''    If IsEmpty(l_varEmployees) Then
'''''        Set l_rsEmpSchedDates = Nothing
'''''        Exit Sub
'''''    Else
'''''        l_rsEmpSchedDates.MoveFirst
'''''    End If
'''''
'''''    Set l_cdTimecard = CreateObject("dbTimeCard.cdTimeCard")
'''''
'''''    Set l_rsActualTime = _
'''''        l_cdTimecard.GetActualTimeCascadeExists(l_varEmployees, l_datStart, l_datEnd)
'''''    Set l_cdTimecard = Nothing
'''''
'''''    If Not RecordsExist(l_rsActualTime) Then
'''''        Set l_rsActualTime = Nothing
'''''        Set l_rsEmpSchedDates = Nothing
'''''        Exit Sub
'''''    End If
'''''
'''''    Do Until l_rsEmpSchedDates.EOF
'''''        l_strFilter = _
'''''            "EMP_NUM = " & l_rsEmpSchedDates!EMP_NUM & _
'''''            " AND CHARGE_DATE >= #" & Format(l_rsEmpSchedDates!START_DATE, "mm/dd/yyyy") & "#" & _
'''''            " AND CHARGE_DATE <= #" & Format(l_rsEmpSchedDates!END_DATE, "mm/dd/yyyy") & "#"
'''''        l_strFilter = MDAC25FilterFix(l_strFilter)
'''''        l_rsActualTime.Filter = l_strFilter
'''''        If RecordsExist(l_rsActualTime) Then
'''''            Do Until l_rsActualTime.EOF
'''''                l_strErrMsg = _
'''''                    "Employee " & CStr(l_rsActualTime!EMP_NUM) & _
'''''                    "[" & Trim(l_rsActualTime!EMP_NAME) & "] has existing actual time on " & _
'''''                    Format(l_rsActualTime!CHARGE_DATE, "mm/dd/yyyy") & "."
'''''                AppendBusinessError _
'''''                    m_rsErrors, _
'''''                    20000, _
'''''                    mstrCBUSINESSWAR, _
'''''                    l_strErrMsg, _
'''''                    m_strCMODULENAME, _
'''''                    TrType_WGM_EMP_ACT, _
'''''                    , _
'''''                    , _
'''''                    CWGMAINTFIELDNAMES, _
'''''                    CStr(m_lngWorkGroup), _
'''''                    Format(l_rsEmpSchedDates!START_DATE, "mm/dd/yyyy"), _
'''''                    "0", _
'''''                    CStr(l_rsEmpSchedDates!EMP_NUM)
'''''                l_rsActualTime.MoveNext
'''''            Loop
'''''            l_blnExists = True
'''''        End If
'''''        l_rsEmpSchedDates.MoveNext
'''''    Loop
'''''
'''''    Set l_rsActualTime = Nothing
'''''    Set l_rsEmpSchedDates = Nothing
'''''
'''''    Exit Sub
'''''
'''''eh:
'''''    Set l_rsActualTime = Nothing
'''''    Set l_rsEmpSchedDates = Nothing
'''''    TrsRaiseError m_strCMODULENAME, "CheckForJobChanges"
'''''End Sub
'''''
'''''Private Sub CheckForJobChangesApplySched()
'''''
''''''JIMB 11/14/2002
''''''Leaving this code in eventhough it is dead code before it is implemented
''''''Changed the design to do this in the check for actual time.
'''''
''''''This routine will apply the changes of the recordset to the
''''''Line changes collect.
'''''
'''''
'''''On Error GoTo eh:
'''''
'''''   m_rsWGDemandLineSchedules.MoveFirst
'''''    While Not m_rsWGDemandLineSchedules.EOF
'''''        Select Case m_rsWGDemandLineSchedules.Status
'''''            Case adRecNew
'''''                AppendWGLineCascadeDates "I", _
'''''                    m_rsWGDemandLineSchedules("WG_NUM").Value, _
'''''                    m_rsWGDemandLineSchedules("LINE_NUM").Value, _
'''''                    m_rsWGDemandLineSchedules("START_DATE").Value, _
'''''                    m_rsWGDemandLineSchedules("END_DATE").Value
'''''            Case adRecModified
'''''                If m_rsWGDemandLineSchedules("JOB_NUM").OriginalValue <> _
'''''                    m_rsWGDemandLineSchedules("JOB_NUM").Value Then
'''''                    AppendWGLineCascadeDates "D", _
'''''                        m_rsWGDemandLineSchedules("WG_NUM").OriginalValue, _
'''''                        m_rsWGDemandLineSchedules("LINE_NUM").OriginalValue, _
'''''                        m_rsWGDemandLineSchedules("START_DATE").OriginalValue, _
'''''                        m_rsWGDemandLineSchedules("END_DATE").OriginalValue
'''''                    AppendWGLineCascadeDates "I", _
'''''                        m_rsWGDemandLineSchedules("WG_NUM").Value, _
'''''                        m_rsWGDemandLineSchedules("LINE_NUM").Value, _
'''''                        m_rsWGDemandLineSchedules("START_DATE").Value, _
'''''                        m_rsWGDemandLineSchedules("END_DATE").Value
'''''                End If
'''''            Case adRecDeleted
'''''                AppendWGLineCascadeDates "D", _
'''''                    m_rsWGDemandLineSchedules("WG_NUM").OriginalValue, _
'''''                    m_rsWGDemandLineSchedules("LINE_NUM").OriginalValue, _
'''''                    m_rsWGDemandLineSchedules("START_DATE").OriginalValue, _
'''''                    m_rsWGDemandLineSchedules("END_DATE").OriginalValue
'''''        End Select
'''''        m_rsWGDemandLineSchedules.MoveNext
'''''    Wend
'''''    Exit Sub
'''''
'''''eh:
'''''
'''''    TrsRaiseError m_strCMODULENAME, "CheckForJobChangesApplySched"
'''''End Sub

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

