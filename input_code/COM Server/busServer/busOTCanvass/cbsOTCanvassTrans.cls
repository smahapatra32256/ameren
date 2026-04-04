VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "cbsOTCanvassTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements IObjectConstruct

'NOTES: PREM CREATED NEW DB METHODS FOR CERTAIN RETRIEVALS
'IN OCTOBER 2005
'ConstructCanvassSheet2, GetOvertimeCanvasses2
', GetOTOpportunity2, GetOTOppCanvassSheet2
', GetOTOppPhoneCalls2, GetOpportunityResponses2
', GetNonCanavassedOvertimeCanvasses2, GetCanvasses2
', GetOTOppCharges2, GetOTOppIneligibilities2
'THESE METHODS USE PARAMETERIZED QUERIES
'PREM HAS ALSO CREATED CreateOvertimeCanvass2 and
'ConstructCanvassSheet2 in the module.

Private Const m_strCMODULENAME As String = "cbsOTCanvassTrans"
Private Const m_lngMINOFFALLOWCONTINUOUS As Long = 60
'Private Const mINITCANVCALLTIME = "0001-01-01-00.00.00"
Private Const mINITCANVCALLTIMESTAMP = "0100-01-01-00.00.00.000000"

Private m_strUserID As String
Private m_lngUserNum As Long
Private m_lngTemplateNum As Long
Private m_lngOTOppNum As Long
Private m_rsOTOpportunity As ADODB.Recordset
Private m_rsCanvassSheet As ADODB.Recordset
Private m_rsTimeCharges As ADODB.Recordset
Private m_rsErrors As ADODB.Recordset
Private m_rsEmpSchedules As ADODB.Recordset
Private m_rsEmpOtherPay As ADODB.Recordset
Private m_rsOpportunityResponses As ADODB.Recordset
Private m_rsEmployees As ADODB.Recordset
Private m_rsAllExceptions As ADODB.Recordset
Private m_rsEmpWorkgroup As ADODB.Recordset
Private m_rsSchedOvertime As ADODB.Recordset
Private m_rsOvertimeResponses As ADODB.Recordset
Private m_rsSysDeletedTimeCharges As ADODB.Recordset
Private m_setChargeDates As utilTRISCommonLib.cuSet
Private m_rsEMPRVOTSchedule As ADODB.Recordset
Private m_rsJobTaskOTSchedule As ADODB.Recordset
Private m_rsEmpValidation As ADODB.Recordset
Private m_rsTTALocations As ADODB.Recordset
Private m_intOppLocataion As Integer

Private m_strErrorMessage As String

Private Type CanvassedShift
    ChargeDate As Date
    StartDateTime As Date
    EndDateTime As Date
    ShiftHours As Variant
    LunchStartTime As String
    LunchEndTime As String
    ExceptionCode As String
    ShiftCode As String
    TravelPaid As Boolean
End Type

'//-----Begin Getting custom settings from COM+ rather than config or registry-----//
'in Component Services, TRIS Components, select component, context menu property, activation tab, check enable object construction
'set Constructor String to ...
'CRACInterface=2;CRACInterfaceLog=0

Dim m_ConstructString As String
Dim m_colConstructStringProperty As Collection
Dim CRACInterface As Integer '1 - Legacy/COM+/In process;2 - WCF/Out of process
Dim CRACInterfaceLog As Integer '0 - off;1 - on

Private Sub Class_Initialize()
    CRACInterface = 1 'Legacy/COM+
    CRACInterfaceLog = 0 'Off
End Sub
   
Private Sub IObjectConstruct_Construct(ByVal pCtorObj As Object)
    Dim lConstructString As IObjectConstructString
    
    On Error GoTo eh
    Set lConstructString = pCtorObj
    
    m_ConstructString = lConstructString.ConstructString
       
    Set m_colConstructStringProperty = ParseContructString(m_ConstructString)
    
    PostConstruct
    '//App.LogEvent MODULENAME & " ObjectConstructString = " & m_ConstructString, vbLogEventTypeInformation
    
    Exit Sub
eh:
    TrsRaiseError m_strCMODULENAME, "IObjectConstruct_Construct", , , , DescArgsFrom(m_ConstructString), True
End Sub

Private Sub PostConstruct()
    On Error GoTo eh
    
    If Not m_colConstructStringProperty Is Nothing Then
        If DoesPropertyExist(m_colConstructStringProperty, "CRACInterface") Then
            If IsNumeric(m_colConstructStringProperty.Item("CRACInterface")) Then
                CRACInterface = CInt(m_colConstructStringProperty.Item("CRACInterface"))
            End If
        End If
    
        If DoesPropertyExist(m_colConstructStringProperty, "CRACInterfaceLog") Then
            If IsNumeric(m_colConstructStringProperty.Item("CRACInterfaceLog")) Then
                CRACInterfaceLog = CInt(m_colConstructStringProperty.Item("CRACInterfaceLog"))
            End If
        End If
    End If
    
    Exit Sub
eh:
    TrsRaiseError m_strCMODULENAME, "PostConstruct", , , , m_colConstructStringProperty, True
End Sub
    
Private Function DoesPropertyExist(ByVal colPropertyValue As Collection, ByVal Property As String) As Boolean
    Dim oValue As Variant
    
    On Error GoTo eh
    
    oValue = colPropertyValue.Item(Property)
    
    DoesPropertyExist = True

    Exit Function
eh:
    DoesPropertyExist = False
End Function

Private Function ParseContructString(ByVal ConstructString As String) As Collection
    Dim colConstructStringProperty As Collection
    Dim PropertyValueArray() As String
    Dim PropertyValue As Variant
    Dim PropertyValueArray2() As String

    On Error GoTo eh
        
    Set colConstructStringProperty = New Collection
    PropertyValueArray = Split(ConstructString, ";")
    
    For Each PropertyValue In PropertyValueArray
        PropertyValueArray2 = Split(PropertyValue, "=")
        If UBound(PropertyValueArray2) >= 1 Then
            colConstructStringProperty.Add PropertyValueArray2(1), PropertyValueArray2(0)
        End If
    Next
       
    Set ParseContructString = colConstructStringProperty
    
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "ParseContructString", , , , DescArgsFrom(ConstructString), True
End Function
'//-----End Getting custom settings from COM+ rather than config or registry-----//

'
' OVER_EXC_CD = CS - Callout standby
' This is time entered to pay 3 hour minimum.
' For now, user must enter EXCEPTION_CD when OVER_EXC_CD = CS.
' Later, change overtime rate determination to calculate, allowing
' timespan to overlap regular scheduled shift, i.e. time within schedule.
' CS time is not worked, so it is not put on the schedule.  Since CS is not
' worked, it is not included in (worked > 8 continuous) for day off.
'
'
'ERROR HANDLING
'
'When application error is detected.  Handle like a business error with trace.
'1.  Format an error message.
'2.  Append to BRV.
'3.  Raise business error.
'In Error handler:
'4.  Build error arguments collection.
'5.  Trace the error, requesting logging.
'6.  Re-raise error via TrsRaiseError.  This will only re-raise.
'Finally, in frontier error handler
'7.  Abort transaction
'8.  Return false and BRV recordset
'
' Example:
'
'    m_strErrorMessage = "ChangeTimeCharge - FindTimeCharge failed to find timecharge:" & _
'        " EMP_NUM = " & v_rsTimeChargeReq!KEY_EMP_NUM & _
'        ", CHARGE_DATE = " & Format(v_rsTimeChargeReq!KEY_CHARGE_DATE, "mm/dd/yyyy") & _
'        ", DATE_WORKED = " & Format(v_rsTimeChargeReq!KEY_DATE_WORKED, "mm/dd/yyyy") & _
'        ", START_TIME = " & v_rsTimeChargeReq!KEY_START_TIME
'        AppendTimeChargeRequestBRV v_rsTimeChargeReq, trErrbusOTCanvass, mstrCAPPLICATIONERR, m_strErrorMessage
'    Err.Raise trErrbusOTOppAppError, m_strCMODULENAME & ":ChangeTimeCharge", m_strErrorMessage
'
'eh:
'    Dim l_colErrInfo As Collection
'
'    If Err.Number = trErrbusOTOppAppError Then
'        Set l_colErrInfo = New Collection
'        l_colErrInfo.Add m_strErrorMessage
'        l_colErrInfo.Add ", m_rsTimeCharges.Filter=" & m_rsTimeCharges.Filter
'        l_colErrInfo.Add ", m_rsTimeCharges.RecordCount=" & m_rsTimeCharges.RecordCount
'        TrsTraceError m_strCMODULENAME, "ChangeTimeCharge", 0, False, Nothing, l_colErrInfo, True
'        Set l_colErrInfo = Nothing
'        TrsRaiseError m_strCMODULENAME, "ChangeTimeCharge"
'    Else
'        TrsRaiseError m_strCMODULENAME, "ChangeTimeCharge"
'    End If
'
'
'eh: (frontier method, eg SaveCanvass)
'    If IsBusinessRuleViolationError(Err.Number) Then
'        SaveCanvass = False
'        objCtx.SetAbort
'        GoTo xit
'    End If
'
'
' TRACING
'
'    Dim stTime As Long
'    Dim ndTime As Long
'    Dim dfTime As Long
'
'
'    #If TRTRACE Then
'        stTime = GetTickCount()
'    #End If
'
'    #If TRTRACE Then
'        ndTime = GetTickCount()
'        dfTime = ndTime - stTime
'        TraceAction App.ThreadID, "3", "GetCanvassState", Format(Now, "mm/dd/yyyy hh:mm:ss"), _
'            CStr(dfTime / 1000)
'    #End If
'


Public Function PostCanvass( _
                    ByVal v_lngOTOppNum As Long, _
                    ByVal v_strUserId As String, _
                    ByRef r_rsEmpSchedules As ADODB.Recordset, _
                    ByRef r_rsBRVs As ADODB.Recordset _
                ) As Boolean
     
    Dim objCtx As ObjectContext
    Dim l_strErrorDesc As String
    Dim l_rsOtherPayRequest As ADODB.Recordset
     
    Dim stTime As Long
    Dim ndTime As Long
    Dim dfTime As Long
     
On Error GoTo eh
    
    #If TRTRACE Then
        stTime = GetTickCount()
    #End If
    
    Set objCtx = GetObjectContext
    
    m_lngOTOppNum = v_lngOTOppNum
    m_strUserID = v_strUserId
    m_lngUserNum = ParseUserID(v_strUserId)
    Set m_rsErrors = CreateNewBusinessErrorSet()

    GetCanvassState "Post"
    GetCanvassedEmployeeScheduleState

    If Not PreValidatePostRequest Then
        objCtx.SetComplete
        GoTo xit
    End If

    m_rsOTOpportunity!OPP_STATUS = "P"
    m_rsOTOpportunity!LAST_MOD_EMP_NUM = m_lngUserNum
            
    Set l_rsOtherPayRequest = CreatePostOtherPayRequest(m_lngOTOppNum)
    
    If Not ProcessTimeEntries(l_rsOtherPayRequest) Then
        objCtx.SetAbort
        PostCanvass = False
        GoTo xit
    End If

    If Not OTBanksCharged() Then
        objCtx.SetAbort
        PostCanvass = False
        GoTo xit
    End If

    If Not SaveOpportunity() Then
        objCtx.SetAbort
        PostCanvass = False
        GoTo xit
    End If

'  Save refusal updates set by OT Equalizer
    If RecordsExist(m_rsTimeCharges) Then
        If Not SaveOppTimeCharges() Then
            PostCanvass = False
            objCtx.SetAbort
            GoTo xit
        End If
    End If
    
    BuildJobTaskInterface ("Post")
    
    If Not SaveJobTaskOTSchedule() Then
        l_strErrorDesc = "EMPRV is unavailable there EMPRV Job Task Charges cannot be saved"
        AppendBusinessError _
            m_rsErrors, _
            trErrdbSaveEMPRVJobTasks, _
            mstrCBUSINESSERR, _
            l_strErrorDesc, _
            m_strCMODULENAME, _
            TrType_OT, , , "OT_OPP_NUM", _
            CStr(m_lngOTOppNum)
        PostCanvass = False
        objCtx.SetAbort
        GoTo xit

    End If
    
    PostCanvass = True
    objCtx.SetComplete
    
xit:
    If Not m_rsErrors Is Nothing Then
        If m_rsErrors.RecordCount = 0 Then
            Set r_rsBRVs = Nothing
        Else
            Set r_rsBRVs = m_rsErrors
        End If
    Else
        Set r_rsBRVs = Nothing
    End If

    FreeGlobalRS

    #If TRTRACE Then
        ndTime = GetTickCount()
        dfTime = ndTime - stTime
        TraceAction App.ThreadID, "3", "PostCanvass", Format(Now, "mm/dd/yyyy hh:mm:ss"), _
            CStr(dfTime / 1000)
    #End If
    
    Set objCtx = Nothing
    Exit Function
    
eh:
    If IsBusinessRuleViolationError(Err.Number) Then
        If Err.Number = trErrAppCommitConflict Then
            AppendBusinessError _
                m_rsErrors, _
                trErrAppCommitConflict, _
                mstrCAPPLICATIONERR, _
                Err.Description, _
                m_strCMODULENAME, _
                TrType_OT, , , "OT_OPP_NUM", _
                CStr(m_lngOTOppNum)
        End If
        PostCanvass = False
        objCtx.SetAbort
        GoTo xit
    End If

    objCtx.SetAbort
    FreeGlobalRS
    Set objCtx = Nothing
    TrsRaiseError m_strCMODULENAME, "PostCanvass", , , , , True
    
End Function


Private Function ProcessTimeEntries(ByVal v_rsOtherPayRequest As ADODB.Recordset) As Boolean

    Dim objOTTimecard As busOTTimecard.cbsOTTimecard
    Dim objCtx As ObjectContext

    Dim stTime As Long
    Dim ndTime As Long
    Dim dfTime As Long

On Error GoTo eh

    #If TRTRACE Then
        stTime = GetTickCount()
    #End If

    If Not RecordsExist(m_rsTimeCharges) And Not RecordsExist(v_rsOtherPayRequest) Then
        ProcessTimeEntries = True
        GoTo xit
    End If

    Set objCtx = GetObjectContext

    Set objOTTimecard = CreateObject("busOTTimecard.cbsOTTimecard")
    ProcessTimeEntries = objOTTimecard.ProcessTimeEntries3( _
            m_rsTimeCharges, v_rsOtherPayRequest, m_rsErrors, CStr(m_lngUserNum))

    Set objOTTimecard = Nothing
    Set objCtx = Nothing

    
xit:
    #If TRTRACE Then
        ndTime = GetTickCount()
        dfTime = ndTime - stTime
        TraceAction App.ThreadID, "3", "ProcessTimeentries", Format(Now, "mm/dd/yyyy hh:mm:ss"), _
            CStr(dfTime / 1000)
    #End If

    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "ProcessTimeentries"

End Function


Public Function UnPostCanvass( _
                    ByVal v_lngOTOppNum As Long, _
                    ByVal v_strUserId As String, _
                    ByRef r_rsEmpSchedules As ADODB.Recordset, _
                    ByRef r_rsBRVs As ADODB.Recordset _
                ) As Boolean
     
    Dim objCtx As ObjectContext
    Dim l_strErrorDesc As String
     
On Error GoTo eh
    
    Set objCtx = GetObjectContext
    
    m_lngOTOppNum = v_lngOTOppNum
    m_strUserID = v_strUserId
    m_lngUserNum = ParseUserID(v_strUserId)
    Set m_rsErrors = CreateNewBusinessErrorSet()

    GetCanvassState "Post"

    If Not PreValidateUnPostRequest Then
        objCtx.SetComplete
        GoTo xit
    End If

    BuildJobTaskInterface ("UnPost")
    
    If Not SaveJobTaskOTSchedule() Then
        l_strErrorDesc = "EMPRV is unavailable there EMPRV Job Task Charges cannot be deleted"
        AppendBusinessError _
            m_rsErrors, _
            trErrdbSaveEMPRVJobTasks, _
            mstrCBUSINESSERR, _
            l_strErrorDesc, _
            m_strCMODULENAME, _
            TrType_OT, , , "OT_OPP_NUM", _
            CStr(m_lngOTOppNum)
        UnPostCanvass = False
        objCtx.SetAbort
        GoTo xit

    End If

    m_rsOTOpportunity!OPP_STATUS = "C"
    m_rsOTOpportunity!LAST_MOD_EMP_NUM = m_lngUserNum
    
    If Not DeleteTimeEntries(v_lngOTOppNum) Then
        UnPostCanvass = False
        objCtx.SetAbort
        GoTo xit
    End If

    If Not SaveOpportunity() Then
        UnPostCanvass = False
        objCtx.SetAbort
        GoTo xit
    End If

    If Not OTBanksUnCharged() Then
        UnPostCanvass = False
        objCtx.SetAbort
        GoTo xit
    End If

    DeleteGeneratedRefusalCharges
    If Not SaveOppTimeCharges() Then
        UnPostCanvass = False
        objCtx.SetAbort
        GoTo xit
    End If
    
    UnPostCanvass = True
    objCtx.SetComplete
    
xit:
    If Not m_rsErrors Is Nothing Then
        If m_rsErrors.RecordCount = 0 Then
            Set r_rsBRVs = Nothing
        Else
            Set r_rsBRVs = m_rsErrors
        End If
    Else
        Set r_rsBRVs = Nothing
    End If

    FreeGlobalRS
    Set objCtx = Nothing
    Exit Function
    
    
eh:
    If IsBusinessRuleViolationError(Err.Number) Then
        If Err.Number = trErrAppCommitConflict Then
            AppendBusinessError _
                m_rsErrors, _
                trErrAppCommitConflict, _
                mstrCAPPLICATIONERR, _
                Err.Description, _
                m_strCMODULENAME, _
                TrType_OT, , , "OT_OPP_NUM", _
                CStr(m_lngOTOppNum)
        End If
        UnPostCanvass = False
        objCtx.SetAbort
        GoTo xit
    End If

    objCtx.SetAbort
    FreeGlobalRS
    Set objCtx = Nothing
    TrsRaiseError m_strCMODULENAME, "UnPostCanvass", , , , , True
    
End Function


Private Function DeleteTimeEntries(ByVal v_lngOTOppNum As Long) As Boolean
     
    Dim objOTTimecard As busOTTimecard.cbsOTTimecard
    Dim l_colEmployees As Collection
    Dim l_aryEmployees As Variant
    Dim l_datStartDate As Date
    Dim l_datEndDate As Date
     
On Error GoTo eh

'    m_rsOTOpportunity!OPP_STATUS = "C"

    Set l_colEmployees = GetActTimeEmployees(m_rsTimeCharges)
    If l_colEmployees.Count = 0 Then
        DeleteTimeEntries = True
        GoTo xit
    Else
        l_aryEmployees = ArrayFromCollection(l_colEmployees)
    End If

'PTS1596
    l_datStartDate = m_rsOTOpportunity!CHARGE_DATE
    l_datEndDate = m_rsOTOpportunity!CHARGE_DATE
'    l_datStartDate = FindLowDateWorked(m_rsTimeCharges)
'    l_datEndDate = FindHighDateWorked(m_rsTimeCharges)
    
    Set objOTTimecard = CreateObject("busOTTimecard.cbsOTTimecard")
    DeleteTimeEntries = objOTTimecard.DeleteTimeEntries(v_lngOTOppNum, l_aryEmployees, l_datStartDate, l_datEndDate, m_rsErrors, CStr(m_lngUserNum))

xit:
    Set objOTTimecard = Nothing
    Exit Function
    
eh:
    Set objOTTimecard = Nothing
    TrsRaiseError m_strCMODULENAME, "DeleteTimeEntries"
    
End Function


Private Function GetPaidEmployees( _
                    ByVal v_rsTimeCharge As ADODB.Recordset) As Collection
    Dim l_objEmpList As cuSet

On Error GoTo eh
    Set l_objEmpList = New cuSet
    If RecordsExist(v_rsTimeCharge) Then
        v_rsTimeCharge.MoveFirst
        Do Until v_rsTimeCharge.EOF
            If IsOvertimePaid(v_rsTimeCharge!EXCEPTION_CD) Then
                l_objEmpList.Add v_rsTimeCharge.Fields("EMP_NUM").Value
            End If
            v_rsTimeCharge.MoveNext
        Loop
        v_rsTimeCharge.MoveFirst
    End If
    
    Set GetPaidEmployees = l_objEmpList.Contents
    Set l_objEmpList = Nothing
    Exit Function
    
eh:
    Set l_objEmpList = Nothing
    TrsRaiseError m_strCMODULENAME, "GetPaidEmployees"

End Function


Private Function GetAllChangedEmployees( _
                    ByVal v_rsTimeCharge As ADODB.Recordset) As Collection
    Dim l_objEmpList As cuSet

On Error GoTo eh
    Set l_objEmpList = New cuSet

    FilterPendingChanges v_rsTimeCharge
    Do Until v_rsTimeCharge.EOF
        Select Case v_rsTimeCharge.Status
            Case adRecNew
'                If IsOvertimePaid(v_rsTimeCharge!EXCEPTION_CD) Then
                    l_objEmpList.Add v_rsTimeCharge.Fields("EMP_NUM").Value
'                End If
            Case adRecModified
'                If IsOvertimePaid(v_rsTimeCharge!EXCEPTION_CD) Or _
                        IsOvertimePaid(v_rsTimeCharge.Fields("EXCEPTION_CD").OriginalValue) Then
                    l_objEmpList.Add v_rsTimeCharge.Fields("EMP_NUM").Value
'                End If
            Case adRecDeleted
'                If IsOvertimePaid(v_rsTimeCharge.Fields("EXCEPTION_CD").OriginalValue) Then
                    l_objEmpList.Add v_rsTimeCharge.Fields("EMP_NUM").OriginalValue
'                End If
        End Select
        v_rsTimeCharge.MoveNext
    Loop
    v_rsTimeCharge.Filter = adFilterNone

    
    Set GetAllChangedEmployees = l_objEmpList.Contents
    Set l_objEmpList = Nothing
    Exit Function
    
eh:
    Set l_objEmpList = Nothing
    TrsRaiseError m_strCMODULENAME, "GetAllChangedEmployees"

End Function


Public Function VoidCanvass( _
                    ByVal v_lngOTOppNum As Long, _
                    ByVal v_strUserId As String, _
                    ByRef r_rsEmpSchedules As ADODB.Recordset, _
                    ByRef r_rsBRVs As ADODB.Recordset _
                ) As Boolean
     
    Dim objCtx As ObjectContext
    Dim l_strErrorDesc As String
    Dim l_blnReturn As Boolean
     
On Error GoTo eh
    
    Set objCtx = GetObjectContext
    
    m_lngOTOppNum = v_lngOTOppNum
    m_strUserID = v_strUserId
    m_lngUserNum = ParseUserID(v_strUserId)
    Set m_rsErrors = CreateNewBusinessErrorSet()

    GetCanvassState "Void"

    If Not PreValidateVoidRequest Then
        objCtx.SetComplete
        GoTo xit
    End If

    BuildJobTaskInterface "Void"
    
    If Not SaveJobTaskOTSchedule() Then
        l_strErrorDesc = "EMPRV is unavailable there EMPRV Job Task Charges cannot be deleted"
        AppendBusinessError _
            m_rsErrors, _
            trErrdbSaveEMPRVJobTasks, _
            mstrCBUSINESSERR, _
            l_strErrorDesc, _
            m_strCMODULENAME, _
            TrType_OT, , , "OT_OPP_NUM", _
            CStr(m_lngOTOppNum)
        VoidCanvass = False
        objCtx.SetAbort
        GoTo xit

    End If
    
    If m_rsOTOpportunity!OPP_STATUS = "P" Then
        If Not DeleteTimeEntries(v_lngOTOppNum) Then
            VoidCanvass = False
            objCtx.SetAbort
            GoTo xit
        End If
    End If

    If m_rsOTOpportunity!OPP_STATUS = "P" Then
        If Not OTBanksUnCharged() Then
            VoidCanvass = False
            objCtx.SetAbort
            GoTo xit
        End If
    End If

    DeleteAllResponses

    m_rsOTOpportunity!OPP_STATUS = "V"
    m_rsOTOpportunity!LAST_MOD_EMP_NUM = m_lngUserNum
    
    If Not SaveOpportunity() Then
        VoidCanvass = False
        objCtx.SetAbort
        GoTo xit
    End If

    If Not SaveCanvassSheet() Then
        VoidCanvass = False
        objCtx.SetAbort
        GoTo xit
    End If
    
    If Not SaveOppTimeCharges() Then
        VoidCanvass = False
        objCtx.SetAbort
        GoTo xit
    End If
    
    If Not SaveScheduledOvertime() Then
        VoidCanvass = False
        objCtx.SetAbort
        GoTo xit
    End If
    
    VoidCanvass = True
    objCtx.SetComplete
    
xit:
    If Not m_rsErrors Is Nothing Then
        If m_rsErrors.RecordCount = 0 Then
            Set r_rsBRVs = Nothing
        Else
            Set r_rsBRVs = m_rsErrors
        End If
    Else
        Set r_rsBRVs = Nothing
    End If

    FreeGlobalRS
    Set objCtx = Nothing
    Exit Function
    
eh:
    If IsBusinessRuleViolationError(Err.Number) Then
        If Err.Number = trErrAppCommitConflict Then
            AppendBusinessError _
                m_rsErrors, _
                trErrAppCommitConflict, _
                mstrCAPPLICATIONERR, _
                Err.Description, _
                m_strCMODULENAME, _
                TrType_OT, , , "OT_OPP_NUM", _
                CStr(m_lngOTOppNum)
        End If
        VoidCanvass = False
        objCtx.SetAbort
        GoTo xit
    End If
    
    objCtx.SetAbort
    FreeGlobalRS
    Set objCtx = Nothing
    TrsRaiseError m_strCMODULENAME, "VoidCanvass", , , , , True
    
End Function


Private Function CreateTimecardRequest( _
                ByVal v_aryEmployees As Variant, _
                ByRef r_rsOvertimeEntries As ADODB.Recordset, _
                ByRef r_rsTimecardRequest As ADODB.Recordset)

    Dim i As Integer
    Dim l_lngEmpNum As Long

On Error GoTo eh

    Set r_rsTimecardRequest = CreateTimecardRequestRS

    For i = 0 To UBound(v_aryEmployees)
        l_lngEmpNum = v_aryEmployees(i)
        r_rsOvertimeEntries.MoveFirst
'what is the filter/sort on r_rsOvertimeEntries?
        Do Until r_rsOvertimeEntries.EOF
'            If r_rsOvertimeEntries!EMP_NUM = l_lngEmpNum And _
'                    IsOvertimePaid(r_rsOvertimeEntries!EXCEPTION_CD) Then
            If r_rsOvertimeEntries!EMP_NUM = l_lngEmpNum Then
                AppendTimecardRequest r_rsOvertimeEntries, r_rsTimecardRequest
            End If
            r_rsOvertimeEntries.MoveNext
        Loop
    Next

    r_rsOvertimeEntries.MoveFirst
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "CreateTimecardRequest"

End Function


Private Sub CreateTimecardRequest2( _
                ByVal v_aryEmployees As Variant, _
                ByRef r_rsOvertimeEntries As ADODB.Recordset, _
                ByRef r_rsTimecardRequest As ADODB.Recordset, _
                ByRef r_rsOtherPayRequest As ADODB.Recordset)

    Dim i As Integer
    Dim l_lngEmpNum As Long
    Dim l_datStartDateTime As Date
    Dim l_datEndDateTime As Date
    Dim l_blnEmpPaid As Boolean
    

On Error GoTo eh

    Set r_rsTimecardRequest = CreateTimecardRequestRS
    Set r_rsOtherPayRequest = CreateInitialOtherPayRequest

    For i = 0 To UBound(v_aryEmployees)
        l_blnEmpPaid = False
        l_lngEmpNum = v_aryEmployees(i)
        r_rsOvertimeEntries.Filter = "EMP_NUM = " & l_lngEmpNum
        Do Until r_rsOvertimeEntries.EOF
            If IsOvertimePaid(r_rsOvertimeEntries!EXCEPTION_CD) Then
                l_blnEmpPaid = True
            End If
            If IsOvertimePaid(r_rsOvertimeEntries!EXCEPTION_CD) Or Trim(r_rsOvertimeEntries!OVER_EXC_CD) = "OE" Then
                AppendTimecardRequest r_rsOvertimeEntries, r_rsTimecardRequest
            End If
            r_rsOvertimeEntries.MoveNext
        Loop
        If l_blnEmpPaid Then
            If EmployeeWorkedStartEndDateTimes(l_lngEmpNum, r_rsOvertimeEntries, l_datStartDateTime, l_datEndDateTime) Then
                If IsEntitledToTTA(l_lngEmpNum, l_datStartDateTime, l_datEndDateTime) Then
                    AppendTTAPaymentRequest l_lngEmpNum, m_lngOTOppNum, r_rsOtherPayRequest
                End If
            End If
        End If
    Next

    r_rsOvertimeEntries.Filter = adFilterNone

    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "CreateTimecardRequest2"

End Sub


Private Sub AppendTimecardRequest( _
                ByVal v_rsOvertimeEntries As ADODB.Recordset, _
                ByRef r_rsTimecardRequest As ADODB.Recordset)

'Is index time required for a delete?  If so, do I need to create a timecard
'request record since all will be deleted in ot timecard?


On Error GoTo eh

    With v_rsOvertimeEntries
        r_rsTimecardRequest.AddNew
        r_rsTimecardRequest!EMP_NUM = !EMP_NUM
        r_rsTimecardRequest!OT_OPP_NUM = !OT_OPP_NUM
        r_rsTimecardRequest!CHARGE_DATE = !CHARGE_DATE
        r_rsTimecardRequest!DATE_WORKED = !DATE_WORKED
        r_rsTimecardRequest!START_TIME = !START_TIME
        r_rsTimecardRequest!END_TIME = !END_TIME
        r_rsTimecardRequest!HOURS = !HOURS
        r_rsTimecardRequest!EXCEPTION_CD = !EXCEPTION_CD
        r_rsTimecardRequest!OVER_EXC_CD = !OVER_EXC_CD
        r_rsTimecardRequest!JOB_NUM = !JOB_NUM
        r_rsTimecardRequest!COMP_SHIFT_CD = !COMP_SHIFT_CD
        r_rsTimecardRequest!OT_REASON_CD = !OT_REASON_CD
        r_rsTimecardRequest!WG_NUM = !WG_NUM
        r_rsTimecardRequest!DISTRIBUTION_ID = !DISTRIBUTION_ID
        r_rsTimecardRequest!DISTRIBUTION_TYPE = !DISTRIBUTION_TYPE
        r_rsTimecardRequest!JOB_NBR = !JOB_NBR
        r_rsTimecardRequest!JOBTASK_NBR = !JOBTASK_NBR
        r_rsTimecardRequest!JOBTASK_STEP_NBR = !JOBTASK_STEP_NBR
        r_rsTimecardRequest!EMPRV_DATABASE_ID = !EMPRV_DATABASE_ID
        If !TRAVEL_PAY_HRS > 0 Then
            r_rsTimecardRequest!OT_TRAVEL_PAY_IND = "Y"
        End If
        r_rsTimecardRequest!KEY_EMP_NUM = !EMP_NUM
        r_rsTimecardRequest!KEY_CHARGE_DATE = !CHARGE_DATE
        r_rsTimecardRequest!KEY_START_TIME = !START_TIME
        r_rsTimecardRequest!KEY_INDEX_TIME = ""
        r_rsTimecardRequest!TIME_STAMP = !LAST_MOD_TIMESTAMP
        r_rsTimecardRequest!RECORD_STATUS = "I"
    End With
    
    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "AppendTimecardRequest"

End Sub


Private Function CreatePostOtherPayRequest( _
                    ByVal v_lngOTOppNum As Long) As ADODB.Recordset

    Dim l_colEmployees As Collection
    Dim l_aryEmployees As Variant
    Dim l_rsOtherPayRequest As ADODB.Recordset
    Dim l_datStartDateTime As Date
    Dim l_datEndDateTime As Date
    Dim l_lngEmpNum As Long
    Dim i As Integer

On Error GoTo eh

    Set l_colEmployees = GetPaidEmployees(m_rsTimeCharges)
    If l_colEmployees.Count = 0 Then
        GoTo xit
    Else
        l_aryEmployees = ArrayFromCollection(l_colEmployees)
    End If
    
    Set l_rsOtherPayRequest = CreateInitialOtherPayRequest

    For i = 0 To UBound(l_aryEmployees)
        l_lngEmpNum = l_aryEmployees(i)
        If EmployeeWorkedStartEndDateTimes(l_lngEmpNum, m_rsTimeCharges, l_datStartDateTime, l_datEndDateTime) Then
            If IsEntitledToTTA(l_lngEmpNum, l_datStartDateTime, l_datEndDateTime) Then
                AppendTTAPaymentRequest l_lngEmpNum, v_lngOTOppNum, l_rsOtherPayRequest
            End If
        End If
    Next
    
    
xit:
    Set CreatePostOtherPayRequest = l_rsOtherPayRequest
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "CreatePostOtherPayRequest"

End Function


'Private Function CreateSaveOtherPayRequest( _
'                    ByVal v_lngOTOppNum As Long) As ADODB.Recordset
'
'    Dim l_colEmployees As Collection
'    Dim l_aryEmployees As Variant
'    Dim l_rsOtherPayRequest As ADODB.Recordset
'    Dim l_datStartDateTime As Date
'    Dim l_datEndDateTime As Date
'    Dim l_lngEmpNum As Long
'    Dim i As Integer
'
'On Error GoTo eh
'
'    Set l_colEmployees = GetPaidEmployees(m_rsTimeCharges)
'    If l_colEmployees.Count = 0 Then
'        GoTo xit
'    Else
'        l_aryEmployees = ArrayFromCollection(l_colEmployees)
'    End If
'
'    Set l_rsOtherPayRequest = CreateInitialOtherPayRequest
'    OpportunityStartEndDateTimes l_datStartDateTime, l_datEndDateTime
'
'    For i = 0 To UBound(l_aryEmployees)
'        l_lngEmpNum = l_aryEmployees(i)
'        If IsEntitledToTTA(l_lngEmpNum, l_datStartDateTime, l_datEndDateTime) Then
'            AppendTTAPaymentRequest l_lngEmpNum, v_lngOTOppNum, l_rsOtherPayRequest
'        End If
'    Next
'
'
'xit:
'    Set CreateSaveOtherPayRequest = l_rsOtherPayRequest
'    Exit Function
'
'eh:
'    TrsRaiseError m_strCMODULENAME, "CreateSaveOtherPayRequest"
'
'End Function


Private Sub AppendTTAPaymentRequest( _
                ByVal v_lngEmpNum As Long, _
                ByVal v_lngOTOppNum As Long, _
                ByRef r_rsOtherPayRequest As ADODB.Recordset)
' Build TTA payment request for valid start/end location combinations.

    Dim l_datChgDate As Date
    Dim l_lngWgNum As Long
    Dim l_strSeqNum As String
    Dim l_lngStartLoc As Long
    Dim l_lngEndLoc As Long

On Error GoTo eh

    l_datChgDate = m_rsOTOpportunity!CHARGE_DATE
    l_lngWgNum = EmployeePrimaryWorkgroup(v_lngEmpNum)
    l_strSeqNum = NextOtherPaySequenceNum(v_lngEmpNum, l_datChgDate)
    l_lngStartLoc = EmployeeLocation(v_lngEmpNum, m_rsEmployees)
    l_lngEndLoc = m_rsOTOpportunity!LOC_CD

    If Not IsTTAStartEndLocationValid(l_lngStartLoc, l_lngEndLoc) Then
        GoTo xit
    End If

    r_rsOtherPayRequest.AddNew
    r_rsOtherPayRequest!EMP_NUM = v_lngEmpNum
    r_rsOtherPayRequest!OT_OPP_NUM = v_lngOTOppNum
    r_rsOtherPayRequest!CHARGE_DATE = l_datChgDate
    r_rsOtherPayRequest!DATE_WORKED = l_datChgDate
    r_rsOtherPayRequest!START_TIME = l_strSeqNum
    r_rsOtherPayRequest!END_TIME = l_strSeqNum
    r_rsOtherPayRequest!EXCEPTION_CD = "TTA"
    r_rsOtherPayRequest!START_LOC_CD = l_lngStartLoc
    r_rsOtherPayRequest!END_LOC_CD = l_lngEndLoc
    r_rsOtherPayRequest!WG_NUM = l_lngWgNum
    
    r_rsOtherPayRequest!KEY_EMP_NUM = v_lngEmpNum
    r_rsOtherPayRequest!KEY_CHARGE_DATE = l_datChgDate
    r_rsOtherPayRequest!KEY_START_TIME = l_strSeqNum
    r_rsOtherPayRequest!KEY_INDEX_TIME = ""
    r_rsOtherPayRequest!RECORD_STATUS = "I"

xit:
    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "AppendTTAPaymentRequest"
    
End Sub


Private Function IsTTAStartEndLocationValid( _
                    ByVal v_lngStartLoc As Long, _
                    ByVal v_lngEndLoc As Long) As Boolean
'Check for existence of valid start/end location for TTA.

On Error GoTo eh

    If v_lngStartLoc = v_lngEndLoc Then
        IsTTAStartEndLocationValid = False
        GoTo xit
    End If
    
    m_rsTTALocations.Filter = "START_LOC_CD = " & v_lngStartLoc & _
        " AND END_LOC_CD = " & v_lngEndLoc
    
    IsTTAStartEndLocationValid = Not m_rsTTALocations.EOF
    m_rsTTALocations.Filter = adFilterNone

xit:
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "IsTTAStartEndLocationValid"
    
End Function


Private Function NextOtherPaySequenceNum( _
                    ByVal v_lngEmpNum As Long, _
                    ByVal v_datChgDate As Date) As String
' Only creating single other pay entry for each employee, so only look at
' persistent other pay entries for last sequence number.  If create multiple
' TTA requests, e.g. multiple segments, not just roundtrip, must look at
' r_rsOtherPayRequest first.

    Dim l_lngLastSeqNum As Long
    Dim l_strNextSeqNum As String

On Error GoTo eh

    If m_rsEmpOtherPay Is Nothing Then
        l_strNextSeqNum = "99:01"
        GoTo xit
    Else
        m_rsEmpOtherPay.Filter = "EMP_NUM = " & v_lngEmpNum & _
                        " AND CHARGE_DATE = " & Format(v_datChgDate, "mm/dd/yyyy")
        m_rsEmpOtherPay.Sort = "START_TIME ASC"
    End If

    If Not RecordsExist(m_rsEmpOtherPay) Then
        l_strNextSeqNum = "99:01"
        GoTo xit
    End If
    
    m_rsEmpOtherPay.MoveLast
    l_lngLastSeqNum = Right(m_rsEmpOtherPay!START_TIME, 2)
    
    l_strNextSeqNum = "99:" & Right("00" & CStr(l_lngLastSeqNum + 1), 2)
    

xit:
    NextOtherPaySequenceNum = l_strNextSeqNum
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "NextOtherPaySequenceNum"
End Function


Private Function CreateTimecardRequestRS() As ADODB.Recordset

On Error GoTo eh

    Const l_lngCFIELDATTR = adFldUpdatable Or adFldIsNullable
    Dim l_adrsSaveTime As ADODB.Recordset

    Set l_adrsSaveTime = New ADODB.Recordset
    With l_adrsSaveTime
        .Fields.Append "EMP_NUM", adInteger, , l_lngCFIELDATTR
        .Fields.Append "CHARGE_DATE", adDate, , l_lngCFIELDATTR
        .Fields.Append "DATE_WORKED", adDate, , l_lngCFIELDATTR
        .Fields.Append "DISTRIBUTION_ID", adChar, 39, l_lngCFIELDATTR
        .Fields.Append "DISTRIBUTION_TYPE", adChar, 1, l_lngCFIELDATTR
        .Fields.Append "HOURS", adDecimal, , l_lngCFIELDATTR
            .Fields.Item("HOURS").Precision = 3
            .Fields.Item("HOURS").NumericScale = 1
        .Fields.Append "START_TIME", adChar, 5, l_lngCFIELDATTR
        .Fields.Append "END_TIME", adChar, 5, l_lngCFIELDATTR
        .Fields.Append "EXCEPTION_CD", adChar, 3, l_lngCFIELDATTR
        .Fields.Append "OVER_EXC_CD", adChar, 3, l_lngCFIELDATTR
        .Fields.Append "JOB_NUM", adChar, 6, l_lngCFIELDATTR
        .Fields.Append "OT_REASON_CD", adChar, 3, l_lngCFIELDATTR
        .Fields.Append "OT_OPP_NUM", adInteger, , l_lngCFIELDATTR
        .Fields.Append "OT_OPP_DESC", adVarChar, 60, l_lngCFIELDATTR
        .Fields.Append "OT_TRAVEL_PAY_IND", adChar, 1, l_lngCFIELDATTR
        .Fields.Append "OT_ADJUST_DESC", adChar, 30, l_lngCFIELDATTR
        .Fields.Append "WG_NUM", adInteger, , l_lngCFIELDATTR
        .Fields.Append "CASE_NUM", adInteger, , l_lngCFIELDATTR
        .Fields.Append "FAMILY_LEAVE_IND", adChar, 1, l_lngCFIELDATTR
        .Fields.Append "COMP_SHIFT_CD", adChar, 3, l_lngCFIELDATTR
        .Fields.Append "COMP_CD", adChar, 3, l_lngCFIELDATTR
        .Fields.Append "REL_CD", adChar, 3, l_lngCFIELDATTR
        .Fields.Append "SCHED_CHG_EXC_CD", adChar, 3, l_lngCFIELDATTR
        .Fields.Append "DISTR_DEPT_NUM", adChar, 3, l_lngCFIELDATTR
        .Fields.Append "APPR_IND", adChar, 1, l_lngCFIELDATTR
        .Fields.Append "TIME_STAMP", adChar, 26, l_lngCFIELDATTR
        .Fields.Append "JOB_NBR", adVarChar, 8, l_lngCFIELDATTR
        .Fields.Append "JOBTASK_NBR", adVarChar, 4, l_lngCFIELDATTR
        .Fields.Append "JOBTASK_STEP_NBR", adInteger, , l_lngCFIELDATTR
        .Fields.Append "EMPRV_DATABASE_ID", adVarChar, 8, l_lngCFIELDATTR

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
        .Fields.Append "KEY_EMP_NUM", adInteger, , l_lngCFIELDATTR
        .Fields.Append "KEY_CHARGE_DATE", adDate, , l_lngCFIELDATTR
        .Fields.Append "KEY_START_TIME", adChar, 5, l_lngCFIELDATTR
        .Fields.Append "KEY_INDEX_TIME", adChar, 26, l_lngCFIELDATTR

        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic

        .Open
    End With

    Set CreateTimecardRequestRS = l_adrsSaveTime
    Set l_adrsSaveTime = Nothing

    Exit Function

eh:
    Set l_adrsSaveTime = Nothing
    Set CreateTimecardRequestRS = Nothing

    TrsRaiseError m_strCMODULENAME, "CreateTimecardRequestRS"

End Function


Private Function CreateInitialOtherPayRequest() As ADODB.Recordset

    Const l_lngCFIELDATTR = adFldUpdatable Or adFldIsNullable
    Dim l_adrsSaveOtherPay As ADODB.Recordset

On Error GoTo eh

    Set l_adrsSaveOtherPay = New ADODB.Recordset

    With l_adrsSaveOtherPay
        .Fields.Append "EMP_NUM", adInteger, , l_lngCFIELDATTR
        .Fields.Append "CHARGE_DATE", adDate, , l_lngCFIELDATTR
        .Fields.Append "DATE_WORKED", adDate, , l_lngCFIELDATTR
        .Fields.Append "DISTRIBUTION_ID", adChar, 39, l_lngCFIELDATTR
        .Fields.Append "DISTRIBUTION_TYPE", adChar, 1, l_lngCFIELDATTR
        .Fields.Append "HOURS", adDecimal, , l_lngCFIELDATTR
            .Fields.Item("HOURS").Precision = 3
            .Fields.Item("HOURS").NumericScale = 1
        .Fields.Append "AMOUNT", adDecimal, , l_lngCFIELDATTR
            .Fields.Item("AMOUNT").Precision = 9
            .Fields.Item("AMOUNT").NumericScale = 2
        .Fields.Append "START_TIME", adChar, 5, l_lngCFIELDATTR
        .Fields.Append "END_TIME", adChar, 5, l_lngCFIELDATTR
        .Fields.Append "EXCEPTION_CD", adChar, 3, l_lngCFIELDATTR
        .Fields.Append "START_LOC_CD", adInteger, , l_lngCFIELDATTR
        .Fields.Append "END_LOC_CD", adInteger, , l_lngCFIELDATTR
        .Fields.Append "CHEM_MTL_TSTR_IND", adChar, 1, l_lngCFIELDATTR
        .Fields.Append "HALF_TRIP_IND", adChar, 1, l_lngCFIELDATTR
        .Fields.Append "JOB_NUM", adChar, 6, l_lngCFIELDATTR
        .Fields.Append "WG_NUM", adInteger, , l_lngCFIELDATTR
        .Fields.Append "COMP_SHIFT_CD", adChar, 3, l_lngCFIELDATTR
        .Fields.Append "DISTR_DEPT_NUM", adChar, 3, l_lngCFIELDATTR
        .Fields.Append "APPR_IND", adChar, 1, l_lngCFIELDATTR
        .Fields.Append "OT_OPP_NUM", adInteger, , l_lngCFIELDATTR

        .Fields.Append "LOC_CD", adInteger, , l_lngCFIELDATTR
        .Fields.Append "DEPT_NUM", adChar, 3, l_lngCFIELDATTR
        
        .Fields.Append "RECORD_STATUS", adChar, 1, adFldUpdatable
        .Fields.Append "SCHED_TIME_IND", adChar, 1, adFldUpdatable

        'KEY INFORMATION
        .Fields.Append "KEY_EMP_NUM", adInteger, , l_lngCFIELDATTR
        .Fields.Append "KEY_CHARGE_DATE", adDate, , l_lngCFIELDATTR
        .Fields.Append "KEY_START_TIME", adChar, 5, l_lngCFIELDATTR
        .Fields.Append "KEY_INDEX_TIME", adChar, 26, l_lngCFIELDATTR

        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic

        .Open
    End With

    Set CreateInitialOtherPayRequest = l_adrsSaveOtherPay
    Set l_adrsSaveOtherPay = Nothing

    Exit Function

eh:
    Set l_adrsSaveOtherPay = Nothing
    TrsRaiseError m_strCMODULENAME, "CreateInitialOtherPayRequest"

End Function


Private Function PreValidatePostRequest() As Boolean
    Dim l_strErrorDesc As String
    Dim l_objCodeTable As dbCodeTables.cdCodeTables
    Dim l_lngMissingTimecharge As Long
    Dim l_strRecentRespCd As String

On Error GoTo eh

    PreValidatePostRequest = True

    If Not RecordsExist(m_rsOTOpportunity) Then
        l_strErrorDesc = "The overtime opportunity, " & CStr(m_lngOTOppNum) _
            & ",  does not exist."
        AppendBusinessError _
            m_rsErrors, _
            trErrbusOTOppDoesNotExist, _
            mstrCAPPLICATIONERR, _
            l_strErrorDesc, _
            m_strCMODULENAME, _
            TrType_OT, , , "OT_OPP_NUM", _
            CStr(m_lngOTOppNum)
        PreValidatePostRequest = False
        GoTo xit
    End If

    If m_rsOTOpportunity!OPP_STATUS <> "C" Then
        l_strErrorDesc = "The overtime opportunity, " & CStr(m_lngOTOppNum) _
            & ", status, " & m_rsOTOpportunity!OPP_STATUS & " is not valid for posting."
        AppendBusinessError _
            m_rsErrors, _
            trErrbusOTOppPostNotValid, _
            mstrCBUSINESSERR, _
            l_strErrorDesc, _
            m_strCMODULENAME, _
            TrType_OT, , , "OT_OPP_NUM", _
            CStr(m_lngOTOppNum)
        PreValidatePostRequest = False
    End If
    
    If Not AnyCanvassResponsesExists() Then
        l_strErrorDesc = "The overtime opportunity, " & CStr(m_lngOTOppNum) _
            & ",  has no responses."
        AppendBusinessError _
            m_rsErrors, _
            trErrbusOTOppNoResponses, _
            mstrCBUSINESSERR, _
            l_strErrorDesc, _
            m_strCMODULENAME, _
            TrType_OT, , , "OT_OPP_NUM", _
            CStr(m_lngOTOppNum)
        PreValidatePostRequest = False
        GoTo xit
    End If
    
' Is it valid to have an opportunity with all refusals?
'    If Not RecordsExist(m_rsTimeCharges) Then
'        l_strErrorDesc = "The overtime opportunity, " & CStr(m_lngOTOppNum) _
'            & ",  has no time charges."
'        AppendBusinessError _
'            m_rsErrors, _
'            trErrbusOTOppNoTimeChgs, _
'            mstrCBUSINESSERR, _
'            l_strErrorDesc, _
'            m_strCMODULENAME, _
'            TrType_OT, , , "OT_OPP_NUM", _
'            CStr(m_lngOTOppNum)
'        PreValidatePostRequest = False
'        GoTo xit
'    End If
    
    m_rsCanvassSheet.MoveFirst
    Set l_objCodeTable = CreateObject("dbCodeTables.cdCodeTables")
    Do Until m_rsCanvassSheet.EOF
        l_strRecentRespCd = MostRecentResponseCode(m_rsCanvassSheet!CANVASS_RESP_CD, m_rsCanvassSheet!FORCE_RESP_CD)
        If l_strRecentRespCd <> "" Then
            If l_objCodeTable.OvertimeResponseWorkIndicator(l_strRecentRespCd) = "Y" Then
                m_rsTimeCharges.Filter = "EMP_NUM = " & m_rsCanvassSheet!EMP_NUM
                If Not RecordsExist(m_rsTimeCharges) Then
                    l_lngMissingTimecharge = l_lngMissingTimecharge + 1
                    l_strErrorDesc = "Employee, " & CStr(m_rsCanvassSheet!EMP_NUM) _
                        & ",  has a worked response but has no time charges."
                    AppendBusinessError _
                        m_rsErrors, _
                        trErrbusOTOppWorkedNoTimeChg, _
                        mstrCBUSINESSERR, _
                        l_strErrorDesc, _
                        m_strCMODULENAME, _
                        TrType_OT_R, , , "EMP_NUM", _
                        CStr(m_rsCanvassSheet!EMP_NUM)
                End If
            End If
        End If
        m_rsCanvassSheet.MoveNext
    Loop
    
    If l_lngMissingTimecharge > 0 Then
        PreValidatePostRequest = False
    End If
    
    m_rsTimeCharges.Filter = adFilterNone
    m_rsCanvassSheet.MoveFirst
    
    
xit:
    Set l_objCodeTable = Nothing
    Exit Function
    
eh:
    Set l_objCodeTable = Nothing
    TrsRaiseError m_strCMODULENAME, "PreValidatePostRequest"
    
End Function


Private Function PreValidateUnPostRequest() As Boolean
    Dim l_strErrorDesc As String
    Dim l_objCodeTable As dbCodeTables.cdCodeTables
    Dim l_lngMissingTimecharge As Long
    Dim l_strRecentRespCd As String

On Error GoTo eh

    PreValidateUnPostRequest = True

    If Not RecordsExist(m_rsOTOpportunity) Then
        l_strErrorDesc = "The overtime opportunity, " & CStr(m_lngOTOppNum) _
            & ",  does not exist."
        AppendBusinessError _
            m_rsErrors, _
            trErrbusOTOppDoesNotExist, _
            mstrCAPPLICATIONERR, _
            l_strErrorDesc, _
            m_strCMODULENAME, _
            TrType_OT, , , "OT_OPP_NUM", _
            CStr(m_lngOTOppNum)
        PreValidateUnPostRequest = False
        GoTo xit
    End If

    If m_rsOTOpportunity!OPP_STATUS <> "P" Then
        l_strErrorDesc = "The overtime opportunity, " & CStr(m_lngOTOppNum) _
            & ", status, " & m_rsOTOpportunity!OPP_STATUS & " is not valid for unposting."
        AppendBusinessError _
            m_rsErrors, _
            trErrbusOTOppPostNotValid, _
            mstrCBUSINESSERR, _
            l_strErrorDesc, _
            m_strCMODULENAME, _
            TrType_OT, , , "OT_OPP_NUM", _
            CStr(m_lngOTOppNum)
        PreValidateUnPostRequest = False
    End If
    
    If Not AnyCanvassResponsesExists() Then
        l_strErrorDesc = "The overtime opportunity, " & CStr(m_lngOTOppNum) _
            & ",  has no responses."
        AppendBusinessError _
            m_rsErrors, _
            trErrbusOTOppNoResponses, _
            mstrCBUSINESSERR, _
            l_strErrorDesc, _
            m_strCMODULENAME, _
            TrType_OT, , , "OT_OPP_NUM", _
            CStr(m_lngOTOppNum)
        PreValidateUnPostRequest = False
        GoTo xit
    End If
    
' Is it valid to have an opportunity with all refusals?
'    If Not RecordsExist(m_rsTimeCharges) Then
'        l_strErrorDesc = "The overtime opportunity, " & CStr(m_lngOTOppNum) _
'            & ",  has no time charges."
'        AppendBusinessError _
'            m_rsErrors, _
'            trErrbusOTOppNoTimeChgs, _
'            mstrCBUSINESSERR, _
'            l_strErrorDesc, _
'            m_strCMODULENAME, _
'            TrType_OT, , , "OT_OPP_NUM", _
'            CStr(m_lngOTOppNum)
'        PreValidateUnPostRequest = False
'        GoTo xit
'    End If
    
    m_rsCanvassSheet.MoveFirst
    Set l_objCodeTable = CreateObject("dbCodeTables.cdCodeTables")
    Do Until m_rsCanvassSheet.EOF
        l_strRecentRespCd = MostRecentResponseCode(m_rsCanvassSheet!CANVASS_RESP_CD, m_rsCanvassSheet!FORCE_RESP_CD)
        If l_strRecentRespCd <> "" Then
            If l_objCodeTable.OvertimeResponseWorkIndicator(l_strRecentRespCd) = "Y" Then
                m_rsTimeCharges.Filter = "EMP_NUM = " & m_rsCanvassSheet!EMP_NUM
                If Not RecordsExist(m_rsTimeCharges) Then
                    l_lngMissingTimecharge = l_lngMissingTimecharge + 1
                    l_strErrorDesc = "Employee, " & CStr(m_rsCanvassSheet!EMP_NUM) _
                        & ",  has a worked response but has no time charges."
                    AppendBusinessError _
                        m_rsErrors, _
                        trErrbusOTOppWorkedNoTimeChg, _
                        mstrCBUSINESSERR, _
                        l_strErrorDesc, _
                        m_strCMODULENAME, _
                        TrType_OT_R, , , "EMP_NUM", _
                        CStr(m_rsCanvassSheet!EMP_NUM)
                End If
            End If
        End If
        m_rsCanvassSheet.MoveNext
    Loop
    
    If l_lngMissingTimecharge > 0 Then
        PreValidateUnPostRequest = False
    End If
    
    m_rsTimeCharges.Filter = adFilterNone
    m_rsCanvassSheet.MoveFirst
    
    
xit:
    Set l_objCodeTable = Nothing
    Exit Function
    
eh:
    Set l_objCodeTable = Nothing
    TrsRaiseError m_strCMODULENAME, "PreValidateUnPostRequest"
    
End Function


Private Function PreValidateVoidRequest() As Boolean
    Dim l_strErrorDesc As String
'    Dim l_objCodeTable As dbCodeTables.cdCodeTables
'    Dim l_lngMissingTimecharge As Long
'    Dim l_strRecentRespCd As String

On Error GoTo eh

    PreValidateVoidRequest = True

    If Not RecordsExist(m_rsOTOpportunity) Then
        l_strErrorDesc = "The overtime opportunity, " & CStr(m_lngOTOppNum) _
            & ",  does not exist."
        AppendBusinessError _
            m_rsErrors, _
            trErrbusOTOppDoesNotExist, _
            mstrCAPPLICATIONERR, _
            l_strErrorDesc, _
            m_strCMODULENAME, _
            TrType_OT, , , "OT_OPP_NUM", _
            CStr(m_lngOTOppNum)
        PreValidateVoidRequest = False
        GoTo xit
    End If

    If m_rsOTOpportunity!OPP_STATUS = "V" Then
        l_strErrorDesc = "The overtime opportunity, " & CStr(m_lngOTOppNum) _
            & ", status, " & m_rsOTOpportunity!OPP_STATUS & " is not valid for voiding."
        AppendBusinessError _
            m_rsErrors, _
            trErrbusOTOppPostNotValid, _
            mstrCBUSINESSERR, _
            l_strErrorDesc, _
            m_strCMODULENAME, _
            TrType_OT, , , "OT_OPP_NUM", _
            CStr(m_lngOTOppNum)
        PreValidateVoidRequest = False
    End If
    
'    If Not AnyCanvassResponsesExists() Then
'        l_strErrorDesc = "The overtime opportunity, " & CStr(m_lngOTOppNum) _
'            & ",  has no responses."
'        AppendBusinessError _
'            m_rsErrors, _
'            trErrbusOTOppNoResponses, _
'            mstrCBUSINESSERR, _
'            l_strErrorDesc, _
'            m_strCMODULENAME, _
'            TrType_OT, , , "OT_OPP_NUM", _
'            CStr(m_lngOTOppNum)
'        PreValidateVoidRequest = False
'        GoTo xit
'    End If
    
' Is it valid to have an opportunity with all refusals?
'    If Not RecordsExist(m_rsTimeCharges) Then
'        l_strErrorDesc = "The overtime opportunity, " & CStr(m_lngOTOppNum) _
'            & ",  has no time charges."
'        AppendBusinessError _
'            m_rsErrors, _
'            trErrbusOTOppNoTimeChgs, _
'            mstrCBUSINESSERR, _
'            l_strErrorDesc, _
'            m_strCMODULENAME, _
'            TrType_OT, , , "OT_OPP_NUM", _
'            CStr(m_lngOTOppNum)
'        PreValidateUnPostRequest = False
'        GoTo xit
'    End If
'
'    m_rsCanvassSheet.MoveFirst
'    Set l_objCodeTable = CreateObject("dbCodeTables.cdCodeTables")
'    Do Until m_rsCanvassSheet.EOF
'        l_strRecentRespCd = MostRecentResponseCode(m_rsCanvassSheet!CANVASS_RESP_CD, m_rsCanvassSheet!FORCE_RESP_CD)
'        If l_strRecentRespCd <> "" Then
'            If l_objCodeTable.OvertimeResponseWorkIndicator(l_strRecentRespCd) = "Y" Then
'                m_rsTimeCharges.Filter = "EMP_NUM = " & m_rsCanvassSheet!EMP_NUM
'                If Not RecordsExist(m_rsTimeCharges) Then
'                    l_lngMissingTimecharge = l_lngMissingTimecharge + 1
'                    l_strErrorDesc = "Employee, " & CStr(m_rsCanvassSheet!EMP_NUM) _
'                        & ",  has a worked response but has no time charges."
'                    AppendBusinessError _
'                        m_rsErrors, _
'                        trErrbusOTOppWorkedNoTimeChg, _
'                        mstrCBUSINESSERR, _
'                        l_strErrorDesc, _
'                        m_strCMODULENAME, _
'                        TrType_OT_R, , , "EMP_NUM", _
'                        CStr(m_rsCanvassSheet!EMP_NUM)
'                End If
'            End If
'        End If
'        m_rsCanvassSheet.MoveNext
'    Loop
'
'    If l_lngMissingTimecharge > 0 Then
'        PreValidateVoidRequest = False
'    End If
    
    m_rsTimeCharges.Filter = adFilterNone
    m_rsCanvassSheet.MoveFirst
    
xit:
'    Set l_objCodeTable = Nothing
    Exit Function
    
eh:
'    Set l_objCodeTable = Nothing
    TrsRaiseError m_strCMODULENAME, "PreValidateVoidRequest"
    
End Function


Private Function MostRecentResponseCode( _
                    ByVal v_strCanvRespCd As String, _
                    ByVal v_strForceRespCd As String) As String

    Dim l_strForceRespCd As String

On Error GoTo eh

    l_strForceRespCd = Trim(v_strForceRespCd)
    If l_strForceRespCd <> "" Then
        MostRecentResponseCode = l_strForceRespCd
    Else
        MostRecentResponseCode = Trim(v_strCanvRespCd)
    End If

    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "MostRecentResponseCode"
    
End Function


Public Function SaveCanvass( _
                    ByVal v_lngOTOppNum As Long, _
                    ByVal v_rsCanvassSheet As ADODB.Recordset, _
                    ByVal v_rsOTOppTimeCharges As ADODB.Recordset, _
                    ByVal v_strUserId As String, _
                    ByRef r_rsBRVs As ADODB.Recordset _
                ) As Boolean

    Dim objCtx As ObjectContext
    Dim l_blnRespPreValidationError As Boolean
    Dim l_blnTimeChgPreValidationError As Boolean
    Dim l_rsTimecardRequest As ADODB.Recordset
    Dim l_rsOtherPayRequest As ADODB.Recordset
    
    Dim l_colEmployeesChanged As Collection
    Dim l_aryEmployees As Variant
    Dim objOTTimecard As busOTTimecard.cbsOTTimecard
    Dim l_datStartDate As Date
    Dim l_datEndDate As Date
    Dim l_strErrorDesc As String
    
    Dim l_datLowChargeDate As Date
    Dim l_datHighChargeDate As Date

     
On Error GoTo eh
        
    Set objCtx = GetObjectContext
       
    If Not RecordsExist(v_rsCanvassSheet) And _
        Not RecordsExist(v_rsOTOppTimeCharges) Then
        SaveCanvass = True
        objCtx.SetComplete
        GoTo xit
    End If
    
    m_lngOTOppNum = v_lngOTOppNum
    m_strUserID = v_strUserId
    m_lngUserNum = ParseUserID(v_strUserId)
    Set m_rsErrors = CreateNewBusinessErrorSet()
    Set m_rsEmpValidation = NewEmployeeValidationSet
    
    If Not PreValidate(v_rsCanvassSheet, v_rsOTOppTimeCharges) Then
        SaveCanvass = False
        objCtx.SetAbort
        GoTo xit
    End If
    
    GetCanvassState "Save"
    If Trim(m_rsOTOpportunity!OPP_STATUS) = "V" Then
        l_strErrorDesc = "The overtime opportunity, " & CStr(m_lngOTOppNum) _
            & ",  is void.  No further changes are allowed."
        AppendBusinessError _
            m_rsErrors, _
            trErrbusOTOppDoesNotExist, _
            mstrCAPPLICATIONERR, _
            l_strErrorDesc, _
            m_strCMODULENAME, _
            TrType_OT, , , "OT_OPP_NUM", _
            CStr(m_lngOTOppNum)
        SaveCanvass = False
        objCtx.SetAbort
        GoTo xit
    End If
    
    GetEmployeeScheduleState v_rsCanvassSheet, v_rsOTOppTimeCharges
    
    m_rsOTOpportunity!LAST_MOD_EMP_NUM = m_lngUserNum
    
    If RecordsExist(v_rsCanvassSheet) Then
        GetOpportunityResponses v_lngOTOppNum
        If Not ProcessResponses(v_rsCanvassSheet, v_rsOTOppTimeCharges) Then
            SaveCanvass = False
            objCtx.SetAbort
            GoTo xit
        End If
    End If
    
    If RecordsExist(v_rsOTOppTimeCharges) Then
        If Not ProcessRemainingTimeCharges(v_rsOTOppTimeCharges, v_rsCanvassSheet) Then
            SaveCanvass = False
            objCtx.SetAbort
            GoTo xit
        End If
    End If
    
    If Not PostValidate() Then
        SaveCanvass = False
        objCtx.SetAbort
        GoTo xit
    End If

    Dim HasErrors As Boolean
    
    '//check SOME registry flag setting
    '//utilRegAccessLib.cuRegistryAccess
    '//DoesKeyExist HKEY_LOCAL_MACHINE, "Software\Ameren\TRIS\Trace???"
    '//GetSetting(HKEY_LOCAL_MACHINE, "Software\Ameren\TRIS\Trace???", "SaveCRACRecordsets") '//registry value

    '//the call could come from OT Canvass, so try to be specific in the log name
    '//copied from busTimecard, the Post and UnPost will probably go there rather than here
    If CRACInterfaceLog = 1 Then
        Dim LogNamePrefix As String

        LogNamePrefix = "CRAC_OTSave" '//default ???
    End If

    '//see mADOExtensions, the data will be saved to %temp% in xml format, suffixes will be added to filenames to avoid duplicates
    If CRACInterfaceLog = 1 Then
        SaveRecordset m_rsTimeCharges, LogNamePrefix & "_TimeCharges", False '//save using any filter
        SaveRecordset m_rsSchedOvertime, LogNamePrefix & "_ScheduledOvertime", False '//save using any filter
    End If
    '//
    
    If CRACInterface = 2 Then
        '//use the New process to load the CRAC into a different process/machine (via WCF)
        Dim CRACIntegrationWCF As Ameren_TRIS_Client_COMIntegration_Library.ContractRuleAutomation
        Set CRACIntegrationWCF = New Ameren_TRIS_Client_COMIntegration_Library.ContractRuleAutomation
    
        CRACIntegrationWCF.ApplyContractRulesForCanvassedAndNonCanvassedOvertime m_rsTimeCharges, m_rsSchedOvertime, m_rsErrors, CLng(v_strUserId), HasErrors
        Set CRACIntegrationWCF = Nothing
    Else '//CRACInterface = 1
        '//use the Old process to load the CRAC into the COM+ process
        Dim CRACIntegration As Ameren_TRIS_Server_COMIntegration.CanvassedAndNonCanvassedOvertime
        Set CRACIntegration = New Ameren_TRIS_Server_COMIntegration.CanvassedAndNonCanvassedOvertime
        CRACIntegration.ApplyContractRules m_rsTimeCharges, m_rsSchedOvertime, m_rsErrors, CLng(v_strUserId), HasErrors
        Set CRACIntegration = Nothing
    End If
    
    '//see mADOExtensions, the data will be saved to %temp% in xml format, suffixes will be added to filenames to avoid duplicates
    If CRACInterfaceLog = 1 Then
        SaveRecordset m_rsErrors, LogNamePrefix & "_Brvs", False '//save using any filter
    End If
    '//
    
    If HasErrors Then
        SaveCanvass = False
        objCtx.SetAbort
        GoTo xit
    End If

    BuildEMPRVInterface
    
    If IsCanvassPosted Then
    'Create TimecardRequests based on Timecharges that have added/changed/deleted.
        Set l_colEmployeesChanged = GetAllChangedEmployees(m_rsTimeCharges)
        If l_colEmployeesChanged.Count > 0 Then
            l_aryEmployees = ArrayFromCollection(l_colEmployeesChanged)
            l_datStartDate = FindLowDateWorked(m_rsTimeCharges)
            l_datEndDate = FindHighDateWorked(m_rsTimeCharges)
            'Dec 2011 RSG - Fix problem when deleting response of posted canvass, system does not delete actual time.
            'Only fails to delete when OT entry deleted with date worked before charge date and OT deleted on Saturday night
            'before owl shift with charge date on Sunday.
            'If only change is delete, there may not be any m_rsTimeCharge rows, but GenDelete is set True so
            'OTTimecard.ProcessTimeEntries4 will use start and end date arguments to get employee state,
            'including timecard which queries based on charge date.
            'Ensure l_datStartDate is min between low date worked and low charge date
            'Ensure l_datEndDate is max between high date worked and high charge date
            l_datLowChargeDate = FindLowChargeDate(m_rsTimeCharges)
            l_datHighChargeDate = FindHighChargeDate(m_rsTimeCharges)
            l_datStartDate = MinDate(l_datStartDate, l_datLowChargeDate)
            l_datEndDate = MaxDate(l_datEndDate, l_datHighChargeDate)
            
'            CreateTimecardRequest l_aryEmployees, m_rsTimeCharges, l_rsTimecardRequest
            CreateTimecardRequest2 l_aryEmployees, m_rsTimeCharges, _
                        l_rsTimecardRequest, l_rsOtherPayRequest
            
            Set objOTTimecard = CreateObject("busOTTimecard.cbsOTTimecard")
            If Not objOTTimecard.ProcessTimeEntries4( _
                    l_rsTimecardRequest, l_rsOtherPayRequest, True, v_lngOTOppNum, l_aryEmployees, _
                    l_datStartDate, l_datEndDate, m_rsErrors, CStr(m_lngUserNum)) Then
                objCtx.SetAbort
                SaveCanvass = False
                GoTo xit
            End If
            
        End If
        
        If Not OTBanksCharged() Then
            objCtx.SetAbort
            SaveCanvass = False
            GoTo xit
        End If
    End If
    
    If Not SaveOpportunity Then
        SaveCanvass = False
        objCtx.SetAbort
        GoTo xit
    End If
        
    If RecordsExist(v_rsCanvassSheet) Then
        If Not SaveCanvassSheet() Then
            SaveCanvass = False
            objCtx.SetAbort
            GoTo xit
        End If
    End If
    
    If Not SaveOppTimeCharges() Then
        SaveCanvass = False
        objCtx.SetAbort
        GoTo xit
    End If
    
    If Not SaveScheduledOvertime() Then
        SaveCanvass = False
        objCtx.SetAbort
        GoTo xit
    End If
    
    If Not SaveEMPRVOTSchedule() Then
        SaveCanvass = False
        objCtx.SetAbort
        GoTo xit
    End If
    
    SaveCanvass = True
    objCtx.SetComplete
    
xit:
    If Not m_rsErrors Is Nothing Then
        If m_rsErrors.RecordCount = 0 Then
            Set r_rsBRVs = Nothing
        Else
            Set r_rsBRVs = m_rsErrors
        End If
    Else
        Set r_rsBRVs = Nothing
    End If
    Set objCtx = Nothing
    FreeGlobalRS

    Exit Function
    
    
eh:
    If IsBusinessRuleViolationError(Err.Number) Then
        If Err.Number = trErrAppCommitConflict Then
            AppendBusinessError _
                m_rsErrors, _
                trErrAppCommitConflict, _
                mstrCAPPLICATIONERR, _
                Err.Description, _
                m_strCMODULENAME, _
                TrType_OT, , , "OT_OPP_NUM", _
                CStr(m_lngOTOppNum)
        End If
        SaveCanvass = False
        objCtx.SetAbort
        GoTo xit
    End If
    
    FreeGlobalRS
    objCtx.SetAbort
    Set objCtx = Nothing
    TrsRaiseError m_strCMODULENAME, "SaveCanvass", , , , , True
    
End Function


Private Sub PreProcessTimeCharges( _
                    ByRef r_rsOTOppTimeCharges As ADODB.Recordset)


On Error GoTo eh

    r_rsOTOppTimeCharges.MoveFirst
    Do Until r_rsOTOppTimeCharges.EOF
        If r_rsOTOppTimeCharges!RECORD_STATUS <> "D" Then
            If TimeChargeRequestSpansMidnight(r_rsOTOppTimeCharges) Then
                SplitTimeChargeRequest r_rsOTOppTimeCharges
            End If
        End If
        r_rsOTOppTimeCharges.MoveNext
    Loop

    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "PreProcessTimeCharges"
    
End Sub


Private Function TimeChargeRequestSpansMidnight( _
                    ByRef r_rsOTOppTimeCharges As ADODB.Recordset) As Boolean

    If r_rsOTOppTimeCharges!END_TIME = "00:00" Then
        TimeChargeRequestSpansMidnight = False
    Else
        TimeChargeRequestSpansMidnight = _
            r_rsOTOppTimeCharges!START_TIME > r_rsOTOppTimeCharges!END_TIME
    End If

End Function


Private Sub SplitTimeChargeRequest( _
                ByRef r_rsOTOppTimeCharges As ADODB.Recordset)
' Split r_rsOTOppTimeCharges row on midnight.
' Clone r_rsOTOppTimeCharges to maintain position on existing row to be split
' and new row.
' Update original row and add new row.
' After split, reposition r_rsOTOppTimeCharges on original row.
' Set travel_pay_ind to "N" on new row if original was "Y" or "N", since
' only 1 timecharge should set travel.


    Dim l_rsTimeChgReqClone As ADODB.Recordset
    Dim l_strStart As String
    Dim l_strEnd As String
    Dim l_datWorked As Date
    Dim l_datWorkedPlus1 As Date
    
    Dim l_datStartDateTime As Date
    Dim l_datEndDateTime As Date
    Dim l_decHours As Variant

On Error GoTo eh

    Set l_rsTimeChgReqClone = r_rsOTOppTimeCharges.Clone
    l_rsTimeChgReqClone.Bookmark = r_rsOTOppTimeCharges.Bookmark
    
    l_strStart = "00:00"
    l_strEnd = r_rsOTOppTimeCharges!END_TIME
    l_datWorked = r_rsOTOppTimeCharges!DATE_WORKED
    l_datWorkedPlus1 = DateAdd("d", 1, l_datWorked)
    
    r_rsOTOppTimeCharges!END_TIME = "24:00"
    l_datStartDateTime = DateTimeFrom(l_datWorked, r_rsOTOppTimeCharges!START_TIME)
    l_datEndDateTime = DateTimeFrom(l_datWorked, r_rsOTOppTimeCharges!END_TIME)
    l_decHours = CDec(DateDiff("n", l_datStartDateTime, l_datEndDateTime) / 60)
    r_rsOTOppTimeCharges!HOURS = l_decHours
    
    l_datStartDateTime = DateTimeFrom(l_datWorkedPlus1, l_strStart)
    l_datEndDateTime = DateTimeFrom(l_datWorkedPlus1, l_strEnd)
    l_decHours = CDec(DateDiff("n", l_datStartDateTime, l_datEndDateTime) / 60)
    
    r_rsOTOppTimeCharges.AddNew
    r_rsOTOppTimeCharges!OT_OPP_NUM = l_rsTimeChgReqClone!OT_OPP_NUM
    r_rsOTOppTimeCharges!EMP_NUM = l_rsTimeChgReqClone!EMP_NUM
    r_rsOTOppTimeCharges!CHARGE_DATE = l_rsTimeChgReqClone!CHARGE_DATE
    r_rsOTOppTimeCharges!DATE_WORKED = l_datWorkedPlus1
    r_rsOTOppTimeCharges!START_TIME = l_strStart
    r_rsOTOppTimeCharges!END_TIME = l_strEnd
    r_rsOTOppTimeCharges!HOURS = l_decHours
    r_rsOTOppTimeCharges!EXCEPTION_CD = l_rsTimeChgReqClone!EXCEPTION_CD
    r_rsOTOppTimeCharges!OVER_EXC_CD = l_rsTimeChgReqClone!OVER_EXC_CD
    r_rsOTOppTimeCharges!JOB_NUM = l_rsTimeChgReqClone!JOB_NUM
    r_rsOTOppTimeCharges!COMP_SHIFT_CD = l_rsTimeChgReqClone!COMP_SHIFT_CD
    r_rsOTOppTimeCharges!COMP_CD = l_rsTimeChgReqClone!COMP_CD
    r_rsOTOppTimeCharges!OT_REASON_CD = l_rsTimeChgReqClone!OT_REASON_CD
    r_rsOTOppTimeCharges!WG_NUM = l_rsTimeChgReqClone!WG_NUM
    r_rsOTOppTimeCharges!DISTRIBUTION_ID = l_rsTimeChgReqClone!DISTRIBUTION_ID
    r_rsOTOppTimeCharges!DISTRIBUTION_TYPE = l_rsTimeChgReqClone!DISTRIBUTION_TYPE
    r_rsOTOppTimeCharges!JOB_NBR = l_rsTimeChgReqClone!JOB_NBR
    r_rsOTOppTimeCharges!JOBTASK_NBR = l_rsTimeChgReqClone!JOBTASK_NBR
    r_rsOTOppTimeCharges!JOBTASK_STEP_NBR = l_rsTimeChgReqClone!JOBTASK_STEP_NBR
    r_rsOTOppTimeCharges!EMPRV_DATABASE_ID = l_rsTimeChgReqClone!EMPRV_DATABASE_ID
    If Trim(l_rsTimeChgReqClone!TRAVEL_PAY_IND) = "" Then
        r_rsOTOppTimeCharges!TRAVEL_PAY_IND = ""
    Else
        r_rsOTOppTimeCharges!TRAVEL_PAY_IND = "N"
    End If
    
    r_rsOTOppTimeCharges!KEY_CHARGE_DATE = l_rsTimeChgReqClone!KEY_CHARGE_DATE
    r_rsOTOppTimeCharges!KEY_DATE_WORKED = l_datWorkedPlus1
    r_rsOTOppTimeCharges!KEY_START_TIME = l_strStart
    r_rsOTOppTimeCharges!KEY_OT_OPP_NUM = l_rsTimeChgReqClone!KEY_OT_OPP_NUM
    r_rsOTOppTimeCharges!KEY_EMP_NUM = l_rsTimeChgReqClone!KEY_EMP_NUM
    r_rsOTOppTimeCharges!RECORD_STATUS = "I"
    r_rsOTOppTimeCharges!CREATED_TIMESTAMP = ""
    r_rsOTOppTimeCharges!LAST_MOD_TIMESTAMP = ""

    r_rsOTOppTimeCharges.Bookmark = l_rsTimeChgReqClone.Bookmark
    
    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "SplitTimeChargeRequest"

End Sub


Private Function IsCanvassPosted() As Boolean

    IsCanvassPosted = Trim(m_rsOTOpportunity!OPP_STATUS) = "P"

End Function


Private Function PreValidate( _
                    ByVal v_rsCanvassSheet As ADODB.Recordset, _
                    ByVal v_rsOTOppTimeCharges As ADODB.Recordset) As Boolean

    Dim l_blnValid As Boolean

On Error GoTo eh

    l_blnValid = True
    
    If RecordsExist(v_rsCanvassSheet) Then
        Do Until v_rsCanvassSheet.EOF
            If Not PreValidateResponse(v_rsCanvassSheet) Then
                l_blnValid = False
            End If
            v_rsCanvassSheet.MoveNext
        Loop
        v_rsCanvassSheet.MoveFirst
    End If
    
    If Not l_blnValid Then
        GoTo xit
    End If

    If RecordsExist(v_rsOTOppTimeCharges) Then
        Do Until v_rsOTOppTimeCharges.EOF
            If Not PreValidateTimeCharge(v_rsOTOppTimeCharges) Then
                l_blnValid = False
            End If
            v_rsOTOppTimeCharges.MoveNext
        Loop
        If l_blnValid Then
            PreProcessTimeCharges v_rsOTOppTimeCharges
        End If
        v_rsOTOppTimeCharges.MoveFirst
    End If

    If Not l_blnValid Then
        GoTo xit
    End If


xit:
    PreValidate = l_blnValid
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "PreValidate"

End Function



Private Function PreValidateResponse( _
                    ByVal v_rsResp As ADODB.Recordset _
                ) As Boolean
     
    Dim l_strRecStatus As String
    Dim l_blnError As Boolean
    Dim l_strResp As String
    Dim l_strCallStart As String
    Dim l_strCallNotes As String
    Dim l_strErrorMessage As String
    Dim l_blnCanvassRespExists As Boolean
    
On Error GoTo eh
    
    If IsNull(v_rsResp!RECORD_STATUS) Or IsEmpty(v_rsResp!RECORD_STATUS) Then
        l_strErrorMessage = "Record Status is missing"
        AppendResponseRequestBRV v_rsResp, trErrbusOTCanvass, mstrCAPPLICATIONERR, l_strErrorMessage
        PreValidateResponse = False
        Exit Function
    Else
        l_strRecStatus = UCase(Trim(v_rsResp!RECORD_STATUS))
        If l_strRecStatus <> "C" Then
            l_strErrorMessage = "Record Status is not valid, change only check.  Input=" & l_strRecStatus
            AppendResponseRequestBRV v_rsResp, trErrbusOTCanvass, mstrCAPPLICATIONERR, l_strErrorMessage
            PreValidateResponse = False
            Exit Function
        End If
    End If
    
    If v_rsResp!KEY_OT_OPP_NUM <> m_lngOTOppNum Then
        l_strErrorMessage = "KEY_OT_OPP_NUM is not consistent with parameter. Request input=" & CStr(v_rsResp!KEY_OT_OPP_NUM) & _
            ", Parameter input=" & m_lngOTOppNum
        AppendResponseRequestBRV v_rsResp, trErrbusOTCanvass, mstrCAPPLICATIONERR, l_strErrorMessage
        PreValidateResponse = False
        Exit Function
    End If
    
    If Not ValidResponseKeyFields(v_rsResp) Then
        PreValidateResponse = False
        Exit Function
    End If
    

    If IsEmpty(v_rsResp!CANVASS_RESP_CD) Or IsNull(v_rsResp!CANVASS_RESP_CD) Then
        l_blnError = True
        l_strErrorMessage = "Canvass response is missing"
        AppendResponseRequestBRV v_rsResp, trErrbusOTCanvass, mstrCAPPLICATIONERR, l_strErrorMessage
    Else
        l_strResp = Trim(v_rsResp!CANVASS_RESP_CD)
        If Not IsValidResponse(l_strResp) Then
            l_blnError = True
            l_strErrorMessage = "Canvass response is not valid.  Input=" & l_strResp
            AppendResponseRequestBRV v_rsResp, trErrbusOTCanvass, mstrCAPPLICATIONERR, l_strErrorMessage
        ElseIf l_strResp <> "" Then
            l_blnCanvassRespExists = True
        End If
    End If
    
'The Call start does not have to be provided
    If IsEmpty(v_rsResp!CANVASS_CALL_TIME) Or IsNull(v_rsResp!CANVASS_CALL_TIME) Then
        v_rsResp!CANVASS_CALL_TIME = ""
    Else
        l_strCallStart = v_rsResp!CANVASS_CALL_TIME
        If Not IsAlphaNumeric(l_strCallStart) Then
             v_rsResp!CANVASS_CALL_TIME = ""
        End If
    End If
    
'    If IsEmpty(v_rsResp!FORCE_RESP_CD) Or IsNull(v_rsResp!FORCE_RESP_CD) Then
'        l_blnError = True
'        l_strErrorMessage = "Force response is missing"
'        AppendResponseRequestBRV v_rsResp, trErrbusOTCanvass, mstrCAPPLICATIONERR, l_strErrorMessage
'    Else
'        l_strResp = Trim(v_rsResp!FORCE_RESP_CD)
'        If l_strResp <> "" And Not l_blnCanvassRespExists Then
'            l_blnError = True
'            l_strErrorMessage = "Force response entered without canvass response.  Input=" & l_strResp
'            AppendResponseRequestBRV v_rsResp, trErrbusOTCanvass, mstrCAPPLICATIONERR, l_strErrorMessage
'        ElseIf Not IsValidResponse(l_strResp) Then
'            l_blnError = True
'            l_strErrorMessage = "Force response is not valid.  Input=" & l_strResp
'            AppendResponseRequestBRV v_rsResp, trErrbusOTCanvass, mstrCAPPLICATIONERR, l_strErrorMessage
'        End If
'    End If

    If IsEmpty(v_rsResp!FORCE_RESP_CD) Or IsNull(v_rsResp!FORCE_RESP_CD) Then
        l_blnError = True
        l_strErrorMessage = "Force response is missing"
        AppendResponseRequestBRV v_rsResp, trErrbusOTCanvass, mstrCAPPLICATIONERR, l_strErrorMessage
    Else
        l_strResp = Trim(v_rsResp!FORCE_RESP_CD)
        If Not IsValidResponse(l_strResp) Then
            l_blnError = True
            l_strErrorMessage = "Force response is not valid.  Input=" & l_strResp
            AppendResponseRequestBRV v_rsResp, trErrbusOTCanvass, mstrCAPPLICATIONERR, l_strErrorMessage
        End If
    End If
    
'The Call start does not have to be provided
    If IsEmpty(v_rsResp!FORCE_CALL_TIME) Or IsNull(v_rsResp!FORCE_CALL_TIME) Then
        v_rsResp!FORCE_CALL_TIME = ""
    Else
        l_strCallStart = v_rsResp!FORCE_CALL_TIME
        If Not IsAlphaNumeric(l_strCallStart) Then
             v_rsResp!FORCE_CALL_TIME = ""
        End If
    End If
    
'The Call Notes do not have to be provided
    If IsEmpty(v_rsResp!CALL_NOTES_TXT) Or IsNull(v_rsResp!CALL_NOTES_TXT) Then
        v_rsResp!CALL_NOTES_TXT = ""
    Else
        l_strCallNotes = v_rsResp!CALL_NOTES_TXT
        If Not IsAlphaNumeric(l_strCallNotes) Then
            v_rsResp!CALL_NOTES_TXT = ""
        End If
    End If
    
    PreValidateResponse = Not l_blnError
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "PreValidateResponse"

End Function


Private Function PostValidate() As Boolean
' Validate all records before saving.  Layer the validation to
' reduce confusion of too many error messages.

    Dim l_blnValid As Boolean
    
On Error GoTo eh
    
    If Not PostValidateProcess() Then
        l_blnValid = False
        GoTo xit
    End If
    
    If Not PostValidateResponses() Or _
                Not PostValidateTimeCharges() Then
        l_blnValid = False
        GoTo xit
    End If
        
    If Not PostValidateSchedules() Then
        l_blnValid = False
        GoTo xit
    End If
        
    l_blnValid = True


xit:
    PostValidate = l_blnValid
    Exit Function


eh:
    TrsRaiseError m_strCMODULENAME, "PostValidate"

End Function


Private Function PostValidateProcess() As Boolean

    Dim l_blnValid As Boolean

On Error GoTo eh

    l_blnValid = True
    If Not RecordsExist(m_rsEmpValidation) Then
        GoTo xit
    End If

    m_rsEmpValidation.Filter = adFilterNone
    m_rsEmpValidation.MoveFirst

    Do Until m_rsEmpValidation.EOF
        If m_rsEmpValidation!OTPROCESS = "Y" Then
            l_blnValid = False
            Exit Do
        End If
        m_rsEmpValidation.MoveNext
    Loop


xit:
    PostValidateProcess = l_blnValid
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "PostValidateProcess"
End Function


Private Function PostValidateResponses() As Boolean
' Answer True when all pending response changes are valid; otherwise False.

    Dim l_strErrorDesc As String
    Dim l_blnValid As Boolean
    Dim l_strRecentRespCd As String

On Error GoTo eh

    l_blnValid = True
    If m_rsCanvassSheet Is Nothing Then
        GoTo xit
    End If

    m_rsCanvassSheet.MoveFirst
    Do Until m_rsCanvassSheet.EOF
        If m_rsCanvassSheet.Status = adRecNew Or _
                m_rsCanvassSheet.Status = adRecModified Then
            l_strRecentRespCd = MostRecentResponseCode(m_rsCanvassSheet!CANVASS_RESP_CD, m_rsCanvassSheet!FORCE_RESP_CD)
            If l_strRecentRespCd <> "" Then
                If Not IsOvertimeWorkedResponse(l_strRecentRespCd) Then
                    m_rsTimeCharges.Filter = "EMP_NUM = " & m_rsCanvassSheet!EMP_NUM
                    Do Until m_rsTimeCharges.EOF
                        If IsOvertimeWorked(m_rsTimeCharges!EXCEPTION_CD, "") Then
                            l_blnValid = False
                            l_strErrorDesc = "Employee, " & CStr(m_rsCanvassSheet!EMP_NUM) _
                                & ",  has a refused response and worked time charge(s)."
                            AppendBusinessError _
                                m_rsErrors, _
                                trErrbusOTOppRefusedWorkedTimeChg, _
                                mstrCBUSINESSERR, _
                                l_strErrorDesc, _
                                m_strCMODULENAME, _
                                TrType_OT_R, , , "EMP_NUM", _
                                CStr(m_rsCanvassSheet!EMP_NUM)
                            Exit Do
                        End If
                        m_rsTimeCharges.MoveNext
                    Loop
                End If
            End If
        End If
        m_rsCanvassSheet.MoveNext
    Loop

    m_rsCanvassSheet.Filter = adFilterNone
    m_rsTimeCharges.Filter = adFilterNone
    
xit:
    PostValidateResponses = l_blnValid
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "PostValidateResponses"

End Function


Private Function PostValidateTimeCharges() As Boolean
' Answer True when all pending time charge changes are valid; otherwise False.

    Dim l_blnValid As Boolean
    Dim l_strRecentRespCd As String
    Dim l_intTravelCount As Integer
    Dim l_blnWorkedResp As Boolean
    Dim l_blnFirstMsg1 As Boolean
    Dim l_strErrorDesc As String
    Dim l_datChgDate As Date
    Dim l_strUnionCd As String
    Dim l_strEndTime As String
    Dim l_datDateWorked As Date
    Dim l_datEndTime As Date
    Dim l_datStartTime As Date
    Dim l_strStartTime As String
    Dim l_lngMinsBreak As Long
    Dim l_lngNumBreak As Long
    Dim l_strDeptNum As String

On Error GoTo eh

    l_blnValid = True
    If m_rsTimeCharges Is Nothing Or _
        m_rsCanvassSheet Is Nothing Then
        GoTo xit
    End If

    l_datChgDate = m_rsOTOpportunity!CHARGE_DATE

    FilterPendingChanges m_rsTimeCharges
    Do Until m_rsTimeCharges.EOF
        If m_rsTimeCharges.Status <> adRecDeleted Then
            If Not IsTimeChargeValid(m_rsTimeCharges) Then
                l_blnValid = False
                GoTo xit
            End If
        End If
        m_rsTimeCharges.MoveNext
    Loop

    m_rsTimeCharges.Filter = adFilterNone
    m_rsCanvassSheet.MoveFirst
    
    Do Until m_rsCanvassSheet.EOF
        l_strRecentRespCd = MostRecentResponseCode(m_rsCanvassSheet!CANVASS_RESP_CD, m_rsCanvassSheet!FORCE_RESP_CD)
        If l_strRecentRespCd <> "" Then
            m_rsTimeCharges.Filter = "EMP_NUM = " & m_rsCanvassSheet!EMP_NUM
            m_rsTimeCharges.Sort = "DATE_WORKED ASC, START_TIME ASC"
            l_strStartTime = ""
            l_intTravelCount = 0
            l_blnWorkedResp = IsOvertimeWorkedResponse(l_strRecentRespCd)
            l_blnFirstMsg1 = True
            Do Until m_rsTimeCharges.EOF
                If m_rsTimeCharges!TRAVEL_PAY_HRS > 0 Then
                    l_intTravelCount = l_intTravelCount + 1
                End If
                If Not l_blnWorkedResp And IsOvertimeWorked(m_rsTimeCharges!EXCEPTION_CD, "") _
                        And l_blnFirstMsg1 Then
                    l_blnValid = False
                    l_strErrorDesc = "Employee, " & CStr(m_rsCanvassSheet!EMP_NUM) _
                        & ",  has a refused response and worked time charge(s)."
                    AppendBusinessError _
                        m_rsErrors, _
                        trErrbusOTOppRefusedWorkedTimeChg, _
                        mstrCBUSINESSERR, _
                        l_strErrorDesc, _
                        m_strCMODULENAME, _
                        TrType_OT_R, , , "EMP_NUM", _
                        CStr(m_rsCanvassSheet!EMP_NUM)
                    l_blnFirstMsg1 = False
                End If
                
                If m_rsTimeCharges!CHARGE_DATE <> l_datChgDate Then
                    l_blnValid = False
                    l_strErrorDesc = "Employee, " & CStr(m_rsCanvassSheet!EMP_NUM) _
                        & ",  has a timecharge with charge date different than opportunity charge date."
                    AppendBusinessError _
                        m_rsErrors, _
                        trErrbusOTOppError, _
                        mstrCBUSINESSERR, _
                        l_strErrorDesc, _
                        m_strCMODULENAME, _
                        TrType_OT_TC, , , "EMP_NUM;CHARGE_DATE;DATE_WORKED;START_TIME;HOURS;EXCEPTION_CD", _
                        CStr(m_rsTimeCharges!EMP_NUM), _
                        Format(m_rsTimeCharges!CHARGE_DATE, "mm/dd/yyyy"), _
                        Format(m_rsTimeCharges!DATE_WORKED, "mm/dd/yyyy"), _
                        m_rsTimeCharges!START_TIME, _
                        m_rsTimeCharges!HOURS, _
                        m_rsTimeCharges!EXCEPTION_CD
                End If
                
                If m_rsTimeCharges!DATE_WORKED = l_datDateWorked And _
                        Trim(m_rsTimeCharges!START_TIME) = l_strStartTime Then
                    l_blnValid = False
                    l_strErrorDesc = "Employee, " & CStr(m_rsCanvassSheet!EMP_NUM) _
                        & ",  has a duplicate timecharge with same date/start time."
                    AppendBusinessError _
                        m_rsErrors, _
                        trErrbusOTOppError, _
                        mstrCBUSINESSERR, _
                        l_strErrorDesc, _
                        m_strCMODULENAME, _
                        TrType_OT_TC, , , "EMP_NUM;CHARGE_DATE;DATE_WORKED;START_TIME;HOURS;EXCEPTION_CD", _
                        CStr(m_rsTimeCharges!EMP_NUM), _
                        Format(m_rsTimeCharges!CHARGE_DATE, "mm/dd/yyyy"), _
                        Format(m_rsTimeCharges!DATE_WORKED, "mm/dd/yyyy"), _
                        m_rsTimeCharges!START_TIME, _
                        m_rsTimeCharges!HOURS, _
                        m_rsTimeCharges!EXCEPTION_CD
                End If
                l_datDateWorked = m_rsTimeCharges!DATE_WORKED
                l_strStartTime = Trim(m_rsTimeCharges!START_TIME)
                m_rsTimeCharges.MoveNext
            Loop
            
            If IsOvertimeWorkedResponse(l_strRecentRespCd) Then
                If l_intTravelCount > 1 Then
                    l_blnValid = False
                    m_strErrorMessage = "Employee, " & CStr(m_rsCanvassSheet!EMP_NUM) _
                            & ",  has more than 1 timecharge with travel time hours."
                    AppendBusinessError _
                        m_rsErrors, _
                        trErrbusOTOppRefusedWorkedTimeChg, _
                        mstrCBUSINESSERR, _
                        m_strErrorMessage, _
                        m_strCMODULENAME, _
                        TrType_OT_R, , , "EMP_NUM", _
                        CStr(m_rsCanvassSheet!EMP_NUM)
                End If
            Else
                If l_intTravelCount > 1 Then
                    l_blnValid = False
                    m_strErrorMessage = "Employee, " & CStr(m_rsCanvassSheet!EMP_NUM) _
                            & ",  has refused overtime and charged travel time more than once."
                    AppendBusinessError _
                        m_rsErrors, _
                        trErrbusOTOppRefusedWorkedTimeChg, _
                        mstrCBUSINESSERR, _
                        m_strErrorMessage, _
                        m_strCMODULENAME, _
                        TrType_OT_R, , , "EMP_NUM", _
                        CStr(m_rsCanvassSheet!EMP_NUM)
                End If
            End If
        End If
        m_rsCanvassSheet.MoveNext
    Loop

    If Not l_blnValid Then
        GoTo xit
    End If

'Check for breaks > 60 minutes.  Any overlapping time is allowed here and equal start
'and end times are expected.
    m_rsCanvassSheet.MoveFirst
    Do Until m_rsCanvassSheet.EOF
        l_strRecentRespCd = MostRecentResponseCode(m_rsCanvassSheet!CANVASS_RESP_CD, m_rsCanvassSheet!FORCE_RESP_CD)
        If l_strRecentRespCd <> "" Then
            l_strUnionCd = EmployeeUnion(m_rsCanvassSheet!EMP_NUM, l_datChgDate, m_rsEmployees)
            l_strDeptNum = EmployeeDepartment(m_rsCanvassSheet!EMP_NUM, l_datChgDate, m_rsEmployees)
            If DoesTimeChargeRuleApply(l_strUnionCd, l_strDeptNum, m_intOppLocataion) Then
                m_rsTimeCharges.Filter = "EMP_NUM = " & m_rsCanvassSheet!EMP_NUM
                m_rsTimeCharges.Sort = "DATE_WORKED ASC, START_TIME ASC"
                l_lngNumBreak = 0
                l_strEndTime = "-1"
                Do Until m_rsTimeCharges.EOF
                    If IsOvertimeWorked(m_rsTimeCharges!EXCEPTION_CD, "") Then
                        If l_strEndTime = "-1" Then     'seed with first time charge
                            l_strEndTime = m_rsTimeCharges!END_TIME
                            l_datDateWorked = m_rsTimeCharges!DATE_WORKED
                        Else
                            l_datEndTime = DateTimeFrom(l_datDateWorked, l_strEndTime, True)
                            l_datStartTime = DateTimeFrom( _
                                m_rsTimeCharges!DATE_WORKED, m_rsTimeCharges!START_TIME)
                            l_lngMinsBreak = DateDiff("n", l_datEndTime, l_datStartTime)
                            If l_lngMinsBreak > 60 Then
                                l_blnValid = False
                                l_strErrorDesc = "Employee, " & CStr(m_rsCanvassSheet!EMP_NUM) _
                                    & ",  has a timecharge with a break > 60 minutes."
                                AppendBusinessError _
                                    m_rsErrors, _
                                    trErrbusOTOppError, _
                                    mstrCBUSINESSERR, _
                                    l_strErrorDesc, _
                                    m_strCMODULENAME, _
                                    TrType_OT_TC, , , "EMP_NUM;CHARGE_DATE;DATE_WORKED;START_TIME;HOURS;EXCEPTION_CD", _
                                    CStr(m_rsTimeCharges!EMP_NUM), _
                                    Format(m_rsTimeCharges!CHARGE_DATE, "mm/dd/yyyy"), _
                                    Format(m_rsTimeCharges!DATE_WORKED, "mm/dd/yyyy"), _
                                    m_rsTimeCharges!START_TIME, _
                                    m_rsTimeCharges!HOURS, _
                                    m_rsTimeCharges!EXCEPTION_CD
                                Exit Do     'continue with next employee
                            Else
                                If l_lngMinsBreak > 0 Then
                                    l_lngNumBreak = l_lngNumBreak + 1
                                    If l_lngNumBreak > 1 Then
                                        l_blnValid = False
                                        l_strErrorDesc = "Employee, " & CStr(m_rsCanvassSheet!EMP_NUM) _
                                            & ",  has more than 1 break in shift."
                                        AppendBusinessError _
                                            m_rsErrors, _
                                            trErrbusOTOppError, _
                                            mstrCBUSINESSERR, _
                                            l_strErrorDesc, _
                                            m_strCMODULENAME, _
                                            TrType_OT_TC, , , "EMP_NUM;CHARGE_DATE;DATE_WORKED;START_TIME;HOURS;EXCEPTION_CD", _
                                            CStr(m_rsTimeCharges!EMP_NUM), _
                                            Format(m_rsTimeCharges!CHARGE_DATE, "mm/dd/yyyy"), _
                                            Format(m_rsTimeCharges!DATE_WORKED, "mm/dd/yyyy"), _
                                            m_rsTimeCharges!START_TIME, _
                                            m_rsTimeCharges!HOURS, _
                                            m_rsTimeCharges!EXCEPTION_CD
                                        Exit Do
                                    End If
                                End If
                                l_strEndTime = m_rsTimeCharges!END_TIME
                                l_datDateWorked = m_rsTimeCharges!DATE_WORKED
                            End If
                        End If
                    End If
                    m_rsTimeCharges.MoveNext
                Loop
            End If
        End If
        m_rsCanvassSheet.MoveNext
    Loop


    m_rsTimeCharges.Filter = adFilterNone
    m_rsTimeCharges.Sort = ""
    m_rsCanvassSheet.MoveFirst
    
xit:
    PostValidateTimeCharges = l_blnValid
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "PostValidateTimeCharges"

End Function


Private Function IsTimeChargeValid( _
                    ByVal v_rsTimeCharge As ADODB.Recordset) As Boolean
                    
    Dim l_blnValid As Boolean
    Dim l_strErrorDesc As String
                    
On Error GoTo eh

    l_blnValid = True
    If Not IsTimeValid(v_rsTimeCharge!START_TIME) Then
        l_blnValid = False
        l_strErrorDesc = "Timecharge START_TIME is not valid(" & v_rsTimeCharge!START_TIME & ")"
        AppendBusinessError _
            m_rsErrors, _
            trErrbusOTOppAppError, _
            mstrCBUSINESSERR, _
            l_strErrorDesc, _
            m_strCMODULENAME, _
            TrType_OT_TC, , , "EMP_NUM;CHARGE_DATE;DATE_WORKED;START_TIME;HOURS;EXCEPTION_CD", _
            CStr(v_rsTimeCharge!EMP_NUM), _
            Format(v_rsTimeCharge!CHARGE_DATE, "mm/dd/yyyy"), _
            Format(v_rsTimeCharge!DATE_WORKED, "mm/dd/yyyy"), _
            v_rsTimeCharge!START_TIME, _
            CStr(v_rsTimeCharge!HOURS), _
            v_rsTimeCharge!EXCEPTION_CD
    End If
    
    If Not IsTimeValid(v_rsTimeCharge!END_TIME) Then
        l_blnValid = False
        l_strErrorDesc = "Timecharge END_TIME is not valid(" & v_rsTimeCharge!END_TIME & ")"
        AppendBusinessError _
            m_rsErrors, _
            trErrbusOTOppAppError, _
            mstrCBUSINESSERR, _
            l_strErrorDesc, _
            m_strCMODULENAME, _
            TrType_OT_TC, , , "EMP_NUM;CHARGE_DATE;DATE_WORKED;START_TIME;HOURS;EXCEPTION_CD", _
            CStr(v_rsTimeCharge!EMP_NUM), _
            Format(v_rsTimeCharge!CHARGE_DATE, "mm/dd/yyyy"), _
            Format(v_rsTimeCharge!DATE_WORKED, "mm/dd/yyyy"), _
            v_rsTimeCharge!START_TIME, _
            CStr(v_rsTimeCharge!HOURS), _
            v_rsTimeCharge!EXCEPTION_CD
    End If
    
    If Not IsExceptionCodeValid(v_rsTimeCharge!EXCEPTION_CD) Then
        l_blnValid = False
        l_strErrorDesc = "Timecharge exception code is invalid(" & _
            v_rsTimeCharge!EXCEPTION_CD & ")"
        AppendBusinessError _
            m_rsErrors, _
            trErrbusOTOppAppError, _
            mstrCBUSINESSERR, _
            l_strErrorDesc, _
            m_strCMODULENAME, _
            TrType_OT_TC, , , "EMP_NUM;CHARGE_DATE;DATE_WORKED;START_TIME;HOURS;EXCEPTION_CD", _
            CStr(v_rsTimeCharge!EMP_NUM), _
            Format(v_rsTimeCharge!CHARGE_DATE, "mm/dd/yyyy"), _
            Format(v_rsTimeCharge!DATE_WORKED, "mm/dd/yyyy"), _
            v_rsTimeCharge!START_TIME, _
            CStr(v_rsTimeCharge!HOURS), _
            v_rsTimeCharge!EXCEPTION_CD
    End If
    
    If v_rsTimeCharge!TRAVEL_PAY_HRS <> 0# And v_rsTimeCharge!TRAVEL_PAY_HRS <> 0.7 Then
        l_blnValid = False
        l_strErrorDesc = "Timecharge travel pay hours is invalid(" & _
            CStr(v_rsTimeCharge!TRAVEL_PAY_HRS) & ")"
        AppendBusinessError _
            m_rsErrors, _
            trErrbusOTOppAppError, _
            mstrCBUSINESSERR, _
            l_strErrorDesc, _
            m_strCMODULENAME, _
            TrType_OT_TC, , , "EMP_NUM;CHARGE_DATE;DATE_WORKED;START_TIME;HOURS;EXCEPTION_CD", _
            CStr(v_rsTimeCharge!EMP_NUM), _
            Format(v_rsTimeCharge!CHARGE_DATE, "mm/dd/yyyy"), _
            Format(v_rsTimeCharge!DATE_WORKED, "mm/dd/yyyy"), _
            v_rsTimeCharge!START_TIME, _
            CStr(v_rsTimeCharge!HOURS), _
            v_rsTimeCharge!EXCEPTION_CD
    End If
    
xit:
    IsTimeChargeValid = l_blnValid
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "IsTimeChargeValid"
                    
End Function


Private Function PostValidateSchedules() As Boolean
' Answer True when all pending schedule changes are valid; otherwise False.

    Dim l_strErrorDesc As String
    Dim l_blnValid As Boolean
    Dim l_rsSchedOvertimeC As ADODB.Recordset
    Dim l_lngOTOppNum As Long
    Dim l_strExcCd As String

On Error GoTo eh

    l_blnValid = True
    
    If m_rsEmpSchedules Is Nothing Then
        GoTo xit
    End If
    
    If Not RecordsExist(m_rsSchedOvertime) Then
        GoTo xit
    Else
        Set l_rsSchedOvertimeC = m_rsSchedOvertime.Clone
    End If

    l_rsSchedOvertimeC.MoveFirst
    Do Until l_rsSchedOvertimeC.EOF
        If l_rsSchedOvertimeC.Status = adRecNew Or _
                l_rsSchedOvertimeC.Status = adRecModified Then
            If IsScheduledOvertimeValid(l_rsSchedOvertimeC) Then
                If IsScheduleConflict(m_rsEmpSchedules, l_rsSchedOvertimeC, l_lngOTOppNum) Then
                    l_blnValid = False
                    If l_lngOTOppNum = 0 Then
                        l_strErrorDesc = "Overtime conflicts with regular scheduled shift"
                    Else
                        l_strErrorDesc = "Overtime conflicts with scheduled overtime opportunity [" _
                                        & l_lngOTOppNum & "]"
                    End If
                    AppendPostValScheduleError trErrbusOTConflictRegShiftExc, l_strErrorDesc, _
                        l_rsSchedOvertimeC!EMP_NUM, l_rsSchedOvertimeC
                ElseIf OvertimeScheduledFollowing(m_rsEmpSchedules, l_rsSchedOvertimeC, l_lngOTOppNum, l_strExcCd) Then
                    l_strErrorDesc = "Overtime already scheduled following this shift at: " & l_strExcCd _
                        & ", opportunity number: " & l_lngOTOppNum & " .Adjust pay if needed."
                    AddTimeChargeError trWarOTEqualizationResult, mstrCBUSINESSWAR, l_strErrorDesc, _
                        l_rsSchedOvertimeC!EMP_NUM, l_rsSchedOvertimeC!CHARGE_DATE, l_rsSchedOvertimeC!Date, _
                        l_rsSchedOvertimeC!START_TIME, 0, l_rsSchedOvertimeC!EXCEPTION_CD
                End If
            Else
                l_blnValid = False
            End If
        End If
        l_rsSchedOvertimeC.MoveNext
    Loop
            
    m_rsEmpSchedules.Filter = adFilterNone

    
xit:
    Set l_rsSchedOvertimeC = Nothing
    PostValidateSchedules = l_blnValid
    Exit Function

eh:
    Set l_rsSchedOvertimeC = Nothing
    TrsRaiseError m_strCMODULENAME, "PostValidateSchedules"

End Function


Private Sub AppendPostValScheduleError( _
                ByVal v_lngErrorNumber As Long, _
                ByVal v_strErrorDescription As String, _
                ByVal v_lngEmpNum As Long, _
                ByVal v_rsSchedOvertime As ADODB.Recordset)
                
    
On Error GoTo eh
   
    If IsSystemGeneratedTimeCharge(v_lngEmpNum) Then
        If Not IsErrorAlreadyRecorded(v_lngEmpNum, "S") Then
            AddResponseError v_lngErrorNumber, mstrCBUSINESSERR, v_strErrorDescription, _
                v_lngEmpNum
        End If
    Else
        AddTimeChargeError v_lngErrorNumber, mstrCBUSINESSERR, v_strErrorDescription, _
            v_rsSchedOvertime!EMP_NUM, v_rsSchedOvertime!CHARGE_DATE, v_rsSchedOvertime!Date, _
            v_rsSchedOvertime!START_TIME, 0, v_rsSchedOvertime!EXCEPTION_CD
    End If

    SetEmployeeValidationError v_lngEmpNum, "S"

    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "AppendPostValScheduleError"
    
End Sub


Private Function IsScheduledOvertimeValid( _
                    ByVal v_rsSchedOvertime As ADODB.Recordset) As Boolean
                    
    Dim l_blnValid As Boolean
    Dim l_strErrorDesc As String
                    
On Error GoTo eh

    l_blnValid = True
    If Not IsTimeValid(v_rsSchedOvertime!START_TIME) Then
        l_blnValid = False
        l_strErrorDesc = "Employee, " & CStr(v_rsSchedOvertime!EMP_NUM) _
            & ",  START_TIME is not valid(" & v_rsSchedOvertime!START_TIME & ")"
        AppendBusinessError _
            m_rsErrors, _
            trErrbusOTOppAppError, _
            mstrCBUSINESSERR, _
            l_strErrorDesc, _
            m_strCMODULENAME, _
            TrType_OT_R, , , "EMP_NUM", _
            CStr(v_rsSchedOvertime!EMP_NUM)
    End If
    
    If Not IsTimeValid(v_rsSchedOvertime!END_TIME) Then
        l_blnValid = False
        l_strErrorDesc = "Employee, " & CStr(v_rsSchedOvertime!EMP_NUM) _
            & ",  END_TIME is not valid(" & v_rsSchedOvertime!END_TIME & ")"
        AppendBusinessError _
            m_rsErrors, _
            trErrbusOTOppAppError, _
            mstrCBUSINESSERR, _
            l_strErrorDesc, _
            m_strCMODULENAME, _
            TrType_OT_R, , , "EMP_NUM", _
            CStr(v_rsSchedOvertime!EMP_NUM)
    End If
    
    If Trim(v_rsSchedOvertime!EXCEPTION_CD) = "" Then
        l_blnValid = False
        l_strErrorDesc = "Employee, " & CStr(v_rsSchedOvertime!EMP_NUM) _
            & ",  EXCEPTION_CD is not valid(" & v_rsSchedOvertime!EXCEPTION_CD & ")"
        AppendBusinessError _
            m_rsErrors, _
            trErrbusOTOppAppError, _
            mstrCBUSINESSERR, _
            l_strErrorDesc, _
            m_strCMODULENAME, _
            TrType_OT_R, , , "EMP_NUM", _
            CStr(v_rsSchedOvertime!EMP_NUM)
    End If
    
    
xit:
    IsScheduledOvertimeValid = l_blnValid
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "IsScheduledOvertimeValid"
                    
End Function


Private Function IsScheduleConflict( _
                    ByVal v_rsEmpSchedules As ADODB.Recordset, _
                    ByVal v_rsSchedOvertime As ADODB.Recordset, _
                    ByRef r_lngOTOppNum As Long) As Boolean

On Error GoTo eh

    r_lngOTOppNum = 0
    v_rsEmpSchedules.Filter = "EMP_NUM = " & v_rsSchedOvertime!EMP_NUM
    Do Until v_rsEmpSchedules.EOF
    'RSG - needs work
    'Identify scheduled work time (regular or overtime),
    'Regular schedule with off exception
        If IsRegularScheduledShift(v_rsEmpSchedules) Then
            If TimePeriodsIntersect2( _
                    DateTimeFrom(v_rsSchedOvertime.Fields("DATE").Value, v_rsSchedOvertime!START_TIME), _
                    DateTimeFrom(v_rsSchedOvertime.Fields("DATE").Value, v_rsSchedOvertime!END_TIME, True), _
                    DateTimeFrom(v_rsEmpSchedules!DATE_WORKED, v_rsEmpSchedules!START_TIME), _
                    DateTimeFrom(v_rsEmpSchedules!DATE_WORKED, v_rsEmpSchedules!END_TIME, True)) Then
                IsScheduleConflict = True
                Exit Do
            End If
        ElseIf IsScheduledOvertime(v_rsEmpSchedules) Then
            If Not IsSameScheduledTime(v_rsSchedOvertime, v_rsEmpSchedules) Then
                If TimePeriodsIntersect2( _
                        DateTimeFrom(v_rsSchedOvertime.Fields("DATE").Value, v_rsSchedOvertime!START_TIME), _
                        DateTimeFrom(v_rsSchedOvertime.Fields("DATE").Value, v_rsSchedOvertime!END_TIME, True), _
                        DateTimeFrom(v_rsEmpSchedules!DATE_WORKED, v_rsEmpSchedules!START_TIME), _
                        DateTimeFrom(v_rsEmpSchedules!DATE_WORKED, v_rsEmpSchedules!END_TIME, True)) Then
                    IsScheduleConflict = True
                    r_lngOTOppNum = v_rsEmpSchedules!OT_OPP_NUM
                    Exit Do
                End If
            End If
        End If
        m_rsEmpSchedules.MoveNext
    Loop

    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "IsScheduleConflict"

End Function


Private Function OvertimeScheduledFollowing( _
                    ByVal v_rsEmpSchedules As ADODB.Recordset, _
                    ByVal v_rsSchedOvertime As ADODB.Recordset, _
                    ByRef r_lngOTOppNum As Long, _
                    ByRef r_strExcCd As String) As Boolean

On Error GoTo eh

    r_lngOTOppNum = 0
    r_strExcCd = ""
    v_rsEmpSchedules.Filter = "EMP_NUM = " & v_rsSchedOvertime!EMP_NUM
    Do Until v_rsEmpSchedules.EOF
        If IsScheduledOvertime(v_rsEmpSchedules) Then
            If v_rsSchedOvertime!OT_OPP_NUM <> v_rsEmpSchedules!OT_OPP_NUM And _
                    v_rsEmpSchedules!DATE_WORKED = v_rsSchedOvertime.Fields("DATE").Value Then
                If DatetimeCompare( _
                        DateTimeFrom(v_rsEmpSchedules!DATE_WORKED, v_rsEmpSchedules!START_TIME), _
                        "GE", DateTimeFrom(v_rsSchedOvertime.Fields("DATE").Value, v_rsSchedOvertime!START_TIME)) Then
                    OvertimeScheduledFollowing = True
                    r_lngOTOppNum = v_rsEmpSchedules!OT_OPP_NUM
                    r_strExcCd = v_rsEmpSchedules!EXCEPTION_CD
                    Exit Do
                End If
            End If
        End If
        m_rsEmpSchedules.MoveNext
    Loop

    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "OvertimeScheduledFollowing"

End Function


Private Function IsSameScheduledTime( _
                    ByVal v_rsSchedOvertime As ADODB.Recordset, _
                    ByVal v_rsEmpSchedule As ADODB.Recordset) As Boolean

On Error GoTo eh

    IsSameScheduledTime = _
        (v_rsSchedOvertime!OT_OPP_NUM = v_rsEmpSchedule!OT_OPP_NUM And _
         v_rsSchedOvertime.Fields("DATE").Value = v_rsEmpSchedule!DATE_WORKED And _
         v_rsSchedOvertime.Fields("START_TIME").Value = v_rsEmpSchedule!START_TIME)
    
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "IsSameScheduledTime"
    
End Function


Private Function IsExceptionCodeValid(ByVal v_strExcCd As String) As Boolean

    Dim l_strExcCd As String

On Error GoTo eh

    l_strExcCd = Trim(v_strExcCd)

    If l_strExcCd = "" Then
        IsExceptionCodeValid = False
        Exit Function
    End If
    
    m_rsAllExceptions.MoveFirst
    m_rsAllExceptions.Find "EXCEPTION_CD = '" & l_strExcCd & "'"
    IsExceptionCodeValid = Not m_rsAllExceptions.EOF
    
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "IsExceptionCodeValid"
    
End Function


Public Function CreateOvertimeCanvass( _
                    ByRef r_lngOTOppNum As Long, _
                    ByVal v_rsOTOppCriteria As ADODB.Recordset, _
                    ByVal v_strUserId As String, _
                    ByRef r_rsBRVs As ADODB.Recordset _
                ) As Boolean
    
 On Error GoTo eh
 
    Dim l_blnTemplateOK As Boolean
    Dim l_datOppChargeDate As Date
    Dim l_datCanvassDate As Date
    Dim l_rsEmpsCanvassDate As ADODB.Recordset
    Dim l_rsEmpsOppChargeDate As ADODB.Recordset
    Dim l_aryEmployees As Variant
    Dim l_lngTemplateNum As Long
    Dim l_lngOTOppNum As Long
    Dim l_intOTGroup As Integer
    
    Dim l_rsTemplateBase As ADODB.Recordset
    Dim l_rsSortOrder As ADODB.Recordset
    Dim l_rsForceSortOrder As ADODB.Recordset
    Dim l_rsGroupSortOrder As ADODB.Recordset
    Dim l_rsSortDefinitions As ADODB.Recordset
    Dim l_rsCanvassSheet As ADODB.Recordset
    Dim l_rsOpportunity As ADODB.Recordset
    Dim l_rsOTOppInelig As ADODB.Recordset
    Dim l_rsPhoneNumbers As ADODB.Recordset
    Dim l_intCanvassLoc As Integer
    Dim l_strOTReasonCd As String
    Dim objCtx As ObjectContext
    Dim l_intBankCD As Integer
    Dim l_intRefuseRule As Integer
    
    Set objCtx = GetObjectContext
    
    m_strUserID = v_strUserId
    l_datCanvassDate = Date
    Set m_rsErrors = CreateNewBusinessErrorSet()
    
    If Not PreValidateOppCriteria(v_rsOTOppCriteria, 0) Then
        CreateOvertimeCanvass = False
        objCtx.SetAbort
        GoTo xit
    End If

    l_strOTReasonCd = v_rsOTOppCriteria!OT_REASON_CD
    l_lngTemplateNum = v_rsOTOppCriteria!TEMPLATE_NUM
    
    If IsChemTesterCanvass(l_lngTemplateNum) Then
        l_datOppChargeDate = v_rsOTOppCriteria!START_DATE
    Else
        l_datOppChargeDate = v_rsOTOppCriteria!CHARGE_DATE
    End If
    l_intOTGroup = v_rsOTOppCriteria!OT_GROUP_NUM
    
    GetOvertimeTemplateData l_lngTemplateNum, l_rsTemplateBase, l_rsSortOrder, _
        l_rsForceSortOrder, l_rsGroupSortOrder, l_rsSortDefinitions
    
    If Not RecordsExist(l_rsTemplateBase) Then
        'writeBRV
        objCtx.SetAbort
        Set objCtx = Nothing
        Exit Function
    End If
    
'Inorder for an employee to appear on a canvass they must be in the overtime group
'When the overtime is going to be worked (l_datOppChargeDate)
    Set l_rsEmpsOppChargeDate = GetEmployeesInOTGroup(l_datOppChargeDate, l_intOTGroup)
    l_aryEmployees = DetermineEmployeesOnCanvass(l_rsEmpsOppChargeDate)
        
    If IsEmpty(l_aryEmployees) Then
        m_strErrorMessage = "There are no employees assigned to this overtime group on this charge date."
        AppendBusinessError _
            m_rsErrors, _
            trErrbusOTOppAppError, _
            mstrCBUSINESSERR, _
            m_strErrorMessage, _
            m_strCMODULENAME, _
            TrType_OT, , , "OT_OPP_NUM", _
            CStr(-1)
        objCtx.SetAbort
        Set objCtx = Nothing
        CreateOvertimeCanvass = False
        GoTo xit
    End If
    
    Set l_rsPhoneNumbers = Nothing
'    Set l_rsPhoneNumbers = GetPeopleSoftPhoneNumber(l_aryEmployees)
    Set l_rsCanvassSheet = ConstructCanvassSheet(l_aryEmployees, l_rsTemplateBase, _
        v_rsOTOppCriteria, l_intBankCD)
    SortCanvassSheet l_rsCanvassSheet, l_rsSortOrder, l_rsForceSortOrder, l_rsGroupSortOrder, l_rsSortDefinitions
    m_lngOTOppNum = GetNextOpportunityNumber()
    GetCanvassState

'What is l_aryLocations? RSG
'    If IsEmpty(l_aryLocations) Then
'        l_intCanvassLoc = l_rsTemplateBase!TEMPL_LOCATION
'    Else
'        l_intCanvassLoc = l_aryLocations(0)
'    End If
    l_intCanvassLoc = l_rsTemplateBase!TEMPL_LOCATION

    l_intRefuseRule = l_rsTemplateBase!REFUSE_RULE_NUM
    BuildOpportunity l_intCanvassLoc, l_intRefuseRule, v_rsOTOppCriteria
'    Set l_rsOTOppInelig = DetermineIneligibilities(l_lngOTOppNum, l_lngOTOppNum, l_aryEmployees)
    Set l_rsPhoneNumbers = GetPeopleSoftPhoneNumber(l_aryEmployees)
    BuildCanvassSheet l_strOTReasonCd, l_rsCanvassSheet, l_rsPhoneNumbers, l_intBankCD, l_rsOTOppInelig

    If v_rsOTOppCriteria!PRINT_ONLY_IND = "Y" Then
        m_rsOTOpportunity!OPP_STATUS = "V"
        m_rsOTOpportunity!MOD_REASON_DESC = "PRINT ONLY"
    End If

    If SaveOpportunity() Then
        'continue
    Else
        'raise error
        CreateOvertimeCanvass = False
        objCtx.SetAbort
        Set objCtx = Nothing
        Exit Function
    End If
    
    If SaveCanvassSheet() Then
            'continue
    Else
        'raise error
        CreateOvertimeCanvass = False
        objCtx.SetAbort
        Set objCtx = Nothing
        Exit Function
    End If
    
'    If SaveOppInelig() Then
'            'continue
'    Else
'        'raise error
'        CreateOvertimeCanvass = False
'        Exit Function
'    End If
'
    

'    Set r_rsBRVs = m_rsCanvassSheet
'    Set r_rsBRVs = Nothing
    
    r_lngOTOppNum = m_lngOTOppNum
    CreateOvertimeCanvass = True
    objCtx.SetComplete

xit:
    If Not m_rsErrors Is Nothing Then
        If m_rsErrors.RecordCount = 0 Then
            Set r_rsBRVs = Nothing
        Else
            Set r_rsBRVs = m_rsErrors
        End If
    Else
        Set r_rsBRVs = Nothing
    End If
    Set objCtx = Nothing
    FreeGlobalRS

    Exit Function
    
eh:
    FreeGlobalRS
    objCtx.SetAbort
    Set objCtx = Nothing
    TrsRaiseError m_strCMODULENAME, "CreateOvertimeCanvass", , , , , True
    
End Function


Private Function PreValidateOppCriteria( _
                    ByVal v_rsOTOppCriteria As ADODB.Recordset, _
                    ByVal v_lngOppType As Long _
                ) As Boolean
' v_lngOppType = 0 is canvass
' v_lngOppType = 1 is noncanvass

    Dim l_strErrorDesc As String
    Dim l_datStartTime As Date
    Dim l_datEndTime As Date
    Dim l_strLunchStart As String
    Dim l_strLunchEnd As String
    Dim l_datLunchStartTime As Date
    Dim l_datLunchEndTime As Date
    Dim l_datLunch As Date

On Error GoTo eh

    PreValidateOppCriteria = True

    If RecordsExist(v_rsOTOppCriteria) Then
        v_rsOTOppCriteria.MoveFirst
    Else
        l_strErrorDesc = "Opportunity criteria is missing"
        AppendBusinessError _
            m_rsErrors, _
            trErrbusOTOppAppError, _
            mstrCBUSINESSERR, _
            l_strErrorDesc, _
            m_strCMODULENAME, _
            TrType_OT, , , "OT_OPP_NUM", _
            CStr(-1)
        PreValidateOppCriteria = False
        GoTo xit
    End If

    If Not IsDate(v_rsOTOppCriteria!CHARGE_DATE) Then
        l_strErrorDesc = "CHARGE_DATE is not valid"
        AppendBusinessError _
            m_rsErrors, _
            trErrbusOTOppAppError, _
            mstrCBUSINESSERR, _
            l_strErrorDesc, _
            m_strCMODULENAME, _
            TrType_OT, , , "OT_OPP_NUM", _
            CStr(-1)
        PreValidateOppCriteria = False
        GoTo xit
    End If

    If Not IsDate(v_rsOTOppCriteria!START_DATE) Then
        l_strErrorDesc = "START_DATE is not valid"
        AppendBusinessError _
            m_rsErrors, _
            trErrbusOTOppAppError, _
            mstrCBUSINESSERR, _
            l_strErrorDesc, _
            m_strCMODULENAME, _
            TrType_OT, , , "OT_OPP_NUM", _
            CStr(-1)
        PreValidateOppCriteria = False
        GoTo xit
    End If

    If Not IsDate(v_rsOTOppCriteria!END_DATE) Then
        l_strErrorDesc = "END_DATE is not valid"
        AppendBusinessError _
            m_rsErrors, _
            trErrbusOTOppAppError, _
            mstrCBUSINESSERR, _
            l_strErrorDesc, _
            m_strCMODULENAME, _
            TrType_OT, , , "OT_OPP_NUM", _
            CStr(-1)
        PreValidateOppCriteria = False
        GoTo xit
    End If

'PTS1596
'DEC 2011 RSG - plus 4 before owl issue, START_DATE and END_DATE may not be same as CHARGE_DATE.
'This edit is correct, except for Callaway that uses canvassing to schedule and post unequalized overtime during outage.
'OT may be pre-shift owl (1900-2300) where charge date is next day, the charge date for adjacent ST owl shift.
'Keeping the edit since Callaway changes the date on the overtime entry to enter preshift plus 4.
    If v_lngOppType = 0 Then
        If v_rsOTOppCriteria!CHARGE_DATE <> v_rsOTOppCriteria!START_DATE And _
            v_rsOTOppCriteria!CHARGE_DATE <> v_rsOTOppCriteria!END_DATE Then
            l_strErrorDesc = "CHARGE_DATE must equal START_DATE or END_DATE"
            AppendBusinessError _
                m_rsErrors, _
                trErrbusOTOppAppError, _
                mstrCBUSINESSERR, _
                l_strErrorDesc, _
                m_strCMODULENAME, _
                TrType_OT, , , "OT_OPP_NUM", _
                CStr(-1)
            PreValidateOppCriteria = False
            GoTo xit
        End If
    Else
        If v_rsOTOppCriteria!CHARGE_DATE < v_rsOTOppCriteria!START_DATE - 1 Or _
            v_rsOTOppCriteria!CHARGE_DATE > v_rsOTOppCriteria!END_DATE + 1 Then
            l_strErrorDesc = "CHARGE_DATE is not valid range between START_DATE - 1 and END_DATE + 1"
            AppendBusinessError _
                m_rsErrors, _
                trErrbusOTOppAppError, _
                mstrCBUSINESSERR, _
                l_strErrorDesc, _
                m_strCMODULENAME, _
                TrType_OT, , , "OT_OPP_NUM", _
                CStr(-1)
            PreValidateOppCriteria = False
            GoTo xit
        End If
    End If

    If Not IsTimeValid(v_rsOTOppCriteria!START_TIME) Then
        l_strErrorDesc = "START_TIME is not valid"
        AppendBusinessError _
            m_rsErrors, _
            trErrbusOTOppAppError, _
            mstrCBUSINESSERR, _
            l_strErrorDesc, _
            m_strCMODULENAME, _
            TrType_OT, , , "OT_OPP_NUM", _
            CStr(-1)
        PreValidateOppCriteria = False
        GoTo xit
    End If

    If Not IsTimeValid(v_rsOTOppCriteria!END_TIME) Then
        l_strErrorDesc = "END_TIME is not valid"
        AppendBusinessError _
            m_rsErrors, _
            trErrbusOTOppAppError, _
            mstrCBUSINESSERR, _
            l_strErrorDesc, _
            m_strCMODULENAME, _
            TrType_OT, , , "OT_OPP_NUM", _
            CStr(-1)
        PreValidateOppCriteria = False
        GoTo xit
    End If

    l_datStartTime = DateTimeFrom(v_rsOTOppCriteria!START_DATE, v_rsOTOppCriteria!START_TIME)
    l_datEndTime = DateTimeFrom(v_rsOTOppCriteria!END_DATE, v_rsOTOppCriteria!END_TIME)
    If DatetimeCompare(l_datStartTime, "GE", l_datEndTime) Then
        l_strErrorDesc = "END_TIME must be after START_TIME"
        AppendBusinessError _
            m_rsErrors, _
            trErrbusOTOppAppError, _
            mstrCBUSINESSERR, _
            l_strErrorDesc, _
            m_strCMODULENAME, _
            TrType_OT, , , "OT_OPP_NUM", _
            CStr(-1)
        PreValidateOppCriteria = False
        GoTo xit
    End If
    
    If OpportunitySpansMidnight(v_rsOTOppCriteria) Then
        If v_rsOTOppCriteria!END_DATE <> v_rsOTOppCriteria!START_DATE + 1 Then
            l_strErrorDesc = "Opportunity spans midnight, END_DATE is not valid"
            AppendBusinessError _
                m_rsErrors, _
                trErrbusOTOppAppError, _
                mstrCBUSINESSERR, _
                l_strErrorDesc, _
                m_strCMODULENAME, _
                TrType_OT, , , "OT_OPP_NUM", _
                CStr(-1)
            PreValidateOppCriteria = False
            GoTo xit
        End If
    End If
    
    l_strLunchStart = Trim(v_rsOTOppCriteria!START_LUNCH)
    l_strLunchEnd = Trim(v_rsOTOppCriteria!END_LUNCH)
    If l_strLunchStart <> "" Or l_strLunchEnd <> "" Then
        If Not (l_strLunchStart <> "" And l_strLunchEnd <> "") Then
            l_strErrorDesc = "Both START_LUNCH and END_LUNCH are required"
            AppendBusinessError _
                m_rsErrors, _
                trErrbusOTOppAppError, _
                mstrCBUSINESSERR, _
                l_strErrorDesc, _
                m_strCMODULENAME, _
                TrType_OT, , , "OT_OPP_NUM", _
                CStr(-1)
            PreValidateOppCriteria = False
            GoTo xit
        End If
    End If
    
    If l_strLunchStart <> "" Then
        If Not IsTimeValid(l_strLunchStart) Then
            l_strErrorDesc = "START_LUNCH is not valid"
            AppendBusinessError _
                m_rsErrors, _
                trErrbusOTOppAppError, _
                mstrCBUSINESSERR, _
                l_strErrorDesc, _
                m_strCMODULENAME, _
                TrType_OT, , , "OT_OPP_NUM", _
                CStr(-1)
            PreValidateOppCriteria = False
            GoTo xit
        ElseIf Not IsTimeValid(l_strLunchEnd) Then
            l_strErrorDesc = "END_LUNCH is not valid"
            AppendBusinessError _
                m_rsErrors, _
                trErrbusOTOppAppError, _
                mstrCBUSINESSERR, _
                l_strErrorDesc, _
                m_strCMODULENAME, _
                TrType_OT, , , "OT_OPP_NUM", _
                CStr(-1)
            PreValidateOppCriteria = False
            GoTo xit
        End If
    End If
    
    If l_strLunchStart <> "" Then
        If v_rsOTOppCriteria!START_TIME > l_strLunchStart Then
            l_datLunch = DateAdd("d", 1, v_rsOTOppCriteria!START_DATE)
        Else
            l_datLunch = v_rsOTOppCriteria!START_DATE
        End If
        l_datLunchStartTime = DateTimeFrom(l_datLunch, l_strLunchStart)
        If v_rsOTOppCriteria!START_TIME > l_strLunchEnd Then
            l_datLunch = DateAdd("d", 1, v_rsOTOppCriteria!START_DATE)
        Else
            l_datLunch = v_rsOTOppCriteria!START_DATE
        End If
        l_datLunchEndTime = DateTimeFrom(l_datLunch, l_strLunchEnd)
        If DatetimeCompare(l_datLunchStartTime, "GE", l_datLunchEndTime) Then
            l_strErrorDesc = "END_LUNCH must be after START_LUNCH"
            AppendBusinessError _
                m_rsErrors, _
                trErrbusOTOppAppError, _
                mstrCBUSINESSERR, _
                l_strErrorDesc, _
                m_strCMODULENAME, _
                TrType_OT, , , "OT_OPP_NUM", _
                CStr(-1)
            PreValidateOppCriteria = False
            GoTo xit
        End If
    End If
    
    If l_strLunchStart <> "" Then
        If Not (DatetimeCompare(l_datLunchStartTime, "GT", l_datStartTime) And _
                DatetimeCompare(l_datLunchStartTime, "LT", l_datEndTime) And _
                DatetimeCompare(l_datLunchEndTime, "GT", l_datStartTime) And _
                DatetimeCompare(l_datLunchEndTime, "LT", l_datEndTime)) Then
            l_strErrorDesc = "Lunch period must be within shift"
            AppendBusinessError _
                m_rsErrors, _
                trErrbusOTOppAppError, _
                mstrCBUSINESSERR, _
                l_strErrorDesc, _
                m_strCMODULENAME, _
                TrType_OT, , , "OT_OPP_NUM", _
                CStr(-1)
            PreValidateOppCriteria = False
            GoTo xit
        End If
    End If

    If Not IsAlphaNumeric(v_rsOTOppCriteria!OT_OPP_DESC) Then
        l_strErrorDesc = "OT_OPP_DESC is not valid, no special characters allowed."
        AppendBusinessError _
            m_rsErrors, _
            trErrbusOTOppAppError, _
            mstrCBUSINESSERR, _
            l_strErrorDesc, _
            m_strCMODULENAME, _
            TrType_OT, , , "OT_OPP_NUM", _
            CStr(-1)
        PreValidateOppCriteria = False
        GoTo xit
    End If


xit:
    Exit Function


eh:
    TrsRaiseError m_strCMODULENAME, "PreValidateOppCriteria"

End Function


Private Function OpportunitySpansMidnight( _
                    ByVal v_rsOTOppCriteria As ADODB.Recordset) As Boolean

    If v_rsOTOppCriteria!END_TIME = "00:00" Then
        OpportunitySpansMidnight = True
    ElseIf v_rsOTOppCriteria!END_TIME < v_rsOTOppCriteria!START_TIME Then
        OpportunitySpansMidnight = True
    End If

    Exit Function

End Function


Private Sub GetOvertimeTemplateData( _
                ByVal v_lngTemplateNum As Long, _
                ByRef r_rsTemplateBase As ADODB.Recordset, _
                ByRef r_rsSortOrder As ADODB.Recordset, _
                ByRef r_rsForceSortOrder As ADODB.Recordset, _
                ByRef r_rsGroupSortOrder As ADODB.Recordset, _
                ByRef r_rsSortDefinitions As ADODB.Recordset)
    
    Dim l_obj As dbOTTemplate.cdOTTemplate
    
On Error GoTo eh
    
    Set l_obj = CreateObject("dbOTTemplate.cdOTTemplate")
    Set r_rsTemplateBase = l_obj.GetOvertimeTemplateBase(v_lngTemplateNum)
    Set r_rsSortOrder = l_obj.GetTemplateSort(v_lngTemplateNum, "C")
    Set r_rsForceSortOrder = l_obj.GetTemplateSort(v_lngTemplateNum, "F")
    Set r_rsGroupSortOrder = l_obj.GetTemplateSort(v_lngTemplateNum, "G")
    Set r_rsSortDefinitions = l_obj.GetAllAvailableSortGroupColumns()
    
    Set l_obj = Nothing
    Exit Sub

eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "GetOvertimeTemplateData"

End Sub


Private Function GetEmployeesInOTGroup( _
    ByVal v_datChargeDate As Date, _
    ByVal v_intOTGroup As Integer _
    ) As ADODB.Recordset
    
    Dim l_obj As dbOTCanvass.cdOTCanvass
    Dim l_rs As ADODB.Recordset
    Dim l_strEmployee As String
    Dim l_colEmployees As Collection
    
On Error GoTo eh

    
    Set l_obj = CreateObject("dbOTCanvass.cdOTCanvass")
    
    Set l_rs = l_obj.GetEmployeesInOTGroup( _
        v_datChargeDate, v_intOTGroup)
    Set l_obj = Nothing
    
    Set GetEmployeesInOTGroup = l_rs
    

    Exit Function
eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "GetEmployeesInOTGroup"
    
End Function
    
    
Private Function ConstructCanvassSheet( _
    ByVal v_aryEmployees As Variant, _
    ByVal v_rsTemplate As ADODB.Recordset, _
    ByVal v_rsOTOpp As ADODB.Recordset, _
    ByRef r_intBankCD As Integer _
    ) As ADODB.Recordset
    
    Dim l_obj As dbOTCanvass.cdOTCanvass
    Dim l_rsSheet As ADODB.Recordset
    Dim l_lngTemplateNum As Long
    Dim l_datOppChargeDate As Date
    Dim l_strCanvassType As String
    Dim l_strBank As String
    Dim l_intProjectID As Integer
    Dim l_intTemplLoc As Integer
    Dim l_intOTGroup As Integer
    Dim l_rsWGSort As ADODB.Recordset
    
On Error GoTo eh

    l_lngTemplateNum = v_rsTemplate!TEMPLATE_NUM
    If IsChemTesterCanvass(l_lngTemplateNum) Then
        l_datOppChargeDate = v_rsOTOpp!START_DATE
    Else
        l_datOppChargeDate = v_rsOTOpp!CHARGE_DATE
    End If
    l_intTemplLoc = v_rsTemplate!TEMPL_LOCATION
    l_intOTGroup = v_rsOTOpp!OT_GROUP_NUM
    GetCanvassBank l_intOTGroup, l_strBank, r_intBankCD
    
    Set l_obj = CreateObject("dbOTCanvass.cdOTCanvass")
    
    Set l_rsSheet = l_obj.ConstructCanvassSheet( _
        v_aryEmployees, l_datOppChargeDate, r_intBankCD)

    If RecordsExist(l_rsSheet) Then
        'continue
    Else
        'probably should raise error
        Set l_obj = Nothing
        Set ConstructCanvassSheet = Nothing
        Exit Function
    End If
    Set l_rsWGSort = l_obj.GetWorkGroupSortData(v_aryEmployees, l_intOTGroup, l_datOppChargeDate)
    UpdateSortColumns l_rsSheet, l_rsWGSort, l_intTemplLoc
    
    Set ConstructCanvassSheet = l_rsSheet
    Exit Function
eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "ConstructCanvassSheet"

End Function


Private Sub UpdateSortColumns( _
    ByRef r_rsSheet As ADODB.Recordset, _
    ByVal v_rsWGSort As ADODB.Recordset, _
    ByVal v_intTemplLoc As Integer)
    
    Dim l_intLocCD As Integer
    Dim l_lngEmpNum As Long
    Dim l_lngWgNum As Long
    
  On Error GoTo eh
  
    r_rsSheet.MoveFirst
    While Not r_rsSheet.EOF
        l_lngEmpNum = r_rsSheet!EMP_NUM
        If RecordsExist(v_rsWGSort) Then
            v_rsWGSort.Find "EMP_NUM = " & str(l_lngEmpNum)
            r_rsSheet!WG_SORT = v_rsWGSort!ORDER_OF_CALL
        End If
        l_intLocCD = r_rsSheet!LOC_CD
        If v_intTemplLoc = l_intLocCD Then
            r_rsSheet!LOC_SORT = l_intLocCD
        Else
            r_rsSheet!LOC_SORT = 2000
        End If
        r_rsSheet.MoveNext
    Wend
  
    r_rsSheet.MoveFirst
  Exit Sub
    
eh:
    
    TrsRaiseError m_strCMODULENAME, "UpdateSortColumns"
        

End Sub
    
    
Private Sub SortCanvassSheet( _
                ByRef r_rsCanvassSheet As ADODB.Recordset, _
                ByVal v_rsSortOrder As ADODB.Recordset, _
                ByVal v_rsForceSortOrder As ADODB.Recordset, _
                ByVal v_rsGroupSortOrder As ADODB.Recordset, _
                ByVal v_rsSortDefinitions As ADODB.Recordset)

    Dim l_strGroupSort As String
    Dim l_strCanvassSort As String
    Dim l_strForceSort As String
    Dim l_blnHasPlantSnrty As Boolean
    Dim l_blnGCMSEmpOnCanvass As Boolean
    Dim l_blnGCMSEmpIn1stOrder As Boolean
    Dim l_rsCS As ADODB.Recordset
    

On Error GoTo eh
    
    l_strGroupSort = BuildSortString(v_rsGroupSortOrder, v_rsSortDefinitions)
    If Trim(l_strGroupSort) = "" Then
        l_strGroupSort = "WG_SORT"
    Else
        l_strGroupSort = l_strGroupSort & ",WG_SORT"
    End If
    l_strCanvassSort = BuildSortString(v_rsSortOrder, v_rsSortDefinitions)
    l_strForceSort = BuildSortString(v_rsForceSortOrder, v_rsSortDefinitions)
    
    r_rsCanvassSheet.Sort = CanvassSheetSort(l_strGroupSort, l_strCanvassSort)
    r_rsCanvassSheet.MoveFirst
    While Not r_rsCanvassSheet.EOF
        r_rsCanvassSheet!ROW_NUM = r_rsCanvassSheet.AbsolutePosition
        r_rsCanvassSheet.MoveNext
    Wend
    
    r_rsCanvassSheet.Sort = CanvassSheetSort(l_strGroupSort, l_strForceSort)
    r_rsCanvassSheet.MoveFirst
    While Not r_rsCanvassSheet.EOF
        r_rsCanvassSheet!FORCE_ROW_NUM = r_rsCanvassSheet.AbsolutePosition
        r_rsCanvassSheet.MoveNext
    Wend
    
'Special sort for Weekend Canvasses with GC&MS employees
' has been commented out
'    DetermineGCMS v_rsSortDefinitions, v_rsSortOrder, v_rsForceSortOrder, r_rsCanvassSheet, l_blnHasPlantSnrty, l_blnGCMSEmpOnCanvass, l_blnGCMSEmpIn1stOrder
'
'    If l_blnGCMSEmpOnCanvass And l_blnHasPlantSnrty Then
'        r_rsCanvassSheet.MoveFirst
'        Set l_rsCS = CopyRecordset(r_rsCanvassSheet)
'        Set r_rsCanvassSheet = GCMSModifySort(l_rsCS, v_rsSortOrder, v_rsForceSortOrder, v_rsGroupSortOrder, v_rsSortDefinitions)
'    End If
        
    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "SortCanvassSheet"

End Sub


Private Function CanvassSheetSort( _
                    ByVal v_strGroupSort As String, _
                    ByVal v_strSortOrder As String _
                ) As String

    Dim l_strSort As String
    
On Error GoTo eh

    If v_strGroupSort = "" Then
        l_strSort = v_strSortOrder
    Else
        If v_strSortOrder = "" Then
            l_strSort = v_strGroupSort
        Else
            l_strSort = v_strGroupSort & ", " & v_strSortOrder
        End If
    End If

    CanvassSheetSort = l_strSort
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "CanvassSheetSort"
    
End Function


Private Function BuildSortString( _
    ByVal v_rsSortOrder As ADODB.Recordset, _
    ByVal v_rsSortDefinitions As ADODB.Recordset, _
    Optional l_blnRemovePlantSnrty As Boolean = False _
    ) As String
    
    Dim l_intSortCD As Integer
    Dim l_strSortCol As String
    Dim l_strSortDir As String
    Dim l_strSort As String

    On Error GoTo eh
    
    l_strSort = ""
    If RecordsExist(v_rsSortOrder) Then
        v_rsSortOrder.MoveFirst
        While Not v_rsSortOrder.EOF
            l_intSortCD = v_rsSortOrder!SORT_CD
            v_rsSortDefinitions.Filter = "SORT_CD = " & str(l_intSortCD)
            l_strSortCol = Trim(v_rsSortDefinitions!SORT_COL_NM)
            If v_rsSortDefinitions!SORT_DIRECTION_IND = "A" Then
                l_strSortDir = " ASC "
            Else
                l_strSortDir = " DESC "
            End If
            If l_strSortCol = "PLANT_SNRTY_QTY" And l_blnRemovePlantSnrty Then
            Else
                If l_strSort = "" Then
                    l_strSort = l_strSortCol & l_strSortDir
                Else
                    l_strSort = l_strSort & " , " & l_strSortCol & l_strSortDir
                End If
            End If
            v_rsSortOrder.MoveNext
        Wend
    End If

  BuildSortString = l_strSort
  
  Exit Function
    
eh:
    
    TrsRaiseError m_strCMODULENAME, "BuildSortString"
        

End Function


Private Sub BuildOpportunity( _
    ByVal v_intCanvassLoc As Integer, _
    ByVal v_intRefuseRule As Integer, _
    ByVal v_rsOTOppCriteria As ADODB.Recordset)
            
    Dim l_intI As Integer
    Dim l_strField As String
            
    On Error GoTo eh
    
    
    m_rsOTOpportunity.AddNew
    m_rsOTOpportunity!PROJECT_ID = 0
    For l_intI = 1 To v_rsOTOppCriteria.Fields.Count
        l_strField = Trim(v_rsOTOppCriteria.Fields(l_intI - 1).Name)
        Select Case l_strField
            Case "PRINT_ONLY_IND"
            'do nothing
                l_strField = ""
            Case "START_DATE"
                m_rsOTOpportunity("DATE_OF_OT") = v_rsOTOppCriteria(l_strField)
                l_strField = "DATE_OF_OT"
            Case "COMP_SHIFT_CD"
                m_rsOTOpportunity("CRAFT_JOB") = v_rsOTOppCriteria(l_strField)
                l_strField = "CRAFT_JOB"
            Case Else
                m_rsOTOpportunity(l_strField).Value = v_rsOTOppCriteria(l_strField).Value
        End Select
        If l_strField <> "" Then
        Select Case m_rsOTOpportunity(l_strField).Type
            Case adChar, adVarChar
                If Not IsAlphaNumeric(m_rsOTOpportunity(l_strField).Value) Then
                    Err.Raise 13
'                    MsgBox "field = " & l_strField
'                    m_rsOTOpportunity(l_strField).Value = ""
                End If
            Case adInteger, adDecimal, adNumeric, adSmallInt
'                m_rsOTOpportunity(l_strField).Value = 0
            Case adDate, adDBDate
'                m_rsOTOpportunity(l_strField).Value = "0100-01-01"
            End Select
        End If
    Next
    
    m_rsOTOpportunity!OT_OPP_NUM = m_lngOTOppNum
    m_rsOTOpportunity!LAST_MOD_EMP_NUM = CLng(m_strUserID)
    m_rsOTOpportunity!CREATED_BY_EMP_NUM = CLng(m_strUserID)
    m_rsOTOpportunity!LOC_CD = v_intCanvassLoc
    m_rsOTOpportunity!OPP_STATUS = "C"
    m_rsOTOpportunity!REFUSE_RULE_NUM = v_intRefuseRule
    m_rsOTOpportunity.Update

    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "BuildOpportunity"

End Sub


Private Sub BuildCanvassSheet( _
    ByVal v_strOTReasonCd As String, _
    ByVal v_rsSheet As ADODB.Recordset, _
    ByVal v_rsPhoneNumbers As ADODB.Recordset, _
    ByVal v_intBankCD As Integer, _
    ByVal v_rsInelig As ADODB.Recordset)
    
    Dim l_strPhone As String
    Dim l_lngEmpNum As Long
    Dim l_intI As Integer
    Dim l_strField As String
    
    On Error GoTo eh
    
    v_rsSheet.Sort = "ROW_NUM"
    v_rsSheet.MoveFirst
    While Not v_rsSheet.EOF
        m_rsCanvassSheet.AddNew
        For l_intI = 1 To v_rsSheet.Fields.Count
            l_strField = v_rsSheet.Fields(l_intI - 1).Name
            Select Case l_strField
                Case "DEPT_SNRTY_QTY", "CORP_SNRTY_QTY", "LOP_SNRTY_QTY", _
                    "JOB_SNRTY_QTY", "CMPY_SNRTY_QTY", _
                    "LOC_SORT", "WG_SORT", "DATE_OF_BIRTH", "DEPT_NUM", "EMP_NAME"
                    'skip
                Case Else
                    m_rsCanvassSheet(l_strField) = v_rsSheet(l_strField)
            End Select
        Next
        m_rsCanvassSheet!OT_OPP_NUM = m_lngOTOppNum
        m_rsCanvassSheet!LAST_MOD_EMP_NUM = CLng(m_strUserID)
        m_rsCanvassSheet!OT_REASON_CD = v_strOTReasonCd
        m_rsCanvassSheet!QUAL_SKILL_IND = "N"
        m_rsCanvassSheet!INELIG_IND = ""
        m_rsCanvassSheet!OT_BANK_CD = v_intBankCD
        l_lngEmpNum = m_rsCanvassSheet!EMP_NUM
        l_strPhone = ExtractPhone(l_lngEmpNum, v_rsPhoneNumbers)
        If l_strPhone <> "" Then
            m_rsCanvassSheet!HOME_PHONE_NUM = l_strPhone
        End If
        
        m_rsCanvassSheet!CANVASS_CALL_TIME = mINITCANVCALLTIMESTAMP
        m_rsCanvassSheet!FORCE_CALL_TIME = mINITCANVCALLTIMESTAMP
        m_rsCanvassSheet!HOURS_CHARGED = 0
        m_rsCanvassSheet!CREATED_EMP_NUM = CLng(m_strUserID)

        v_rsSheet.MoveNext
    Wend
    
    m_rsCanvassSheet.Update
     Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "BuildCanvassSheet"

End Sub


Private Function SaveOpportunity() As Boolean
    Dim l_obj As dbOTCanvass.cdOTCanvass

    Dim stTime As Long
    Dim ndTime As Long
    Dim dfTime As Long

On Error GoTo eh

    #If TRTRACE Then
        stTime = GetTickCount()
    #End If
    
    Set l_obj = CreateObject("dbOTCanvass.cdOTCanvass")
    SaveOpportunity = l_obj.SaveOpportunity(m_rsOTOpportunity)
    Set l_obj = Nothing
    
    #If TRTRACE Then
        ndTime = GetTickCount()
        dfTime = ndTime - stTime
        TraceAction App.ThreadID, "3", "SaveOpportunity", Format(Now, "mm/dd/yyyy hh:mm:ss"), _
            CStr(dfTime / 1000)
    #End If
    
    Exit Function
    
eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "SaveOpportunity"
    
End Function


Private Function SaveCanvassSheet() As Boolean
    Dim l_obj As dbOTCanvass.cdOTCanvass

On Error GoTo eh
    m_rsCanvassSheet.Filter = adFilterNone
    Set l_obj = CreateObject("dbOTCanvass.cdOTCanvass")
    SaveCanvassSheet = l_obj.SaveCanvassSheet(m_rsCanvassSheet)
    Set l_obj = Nothing
    Exit Function
    
eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "SaveCanvassSheet"
End Function


Private Function GetNextOpportunityNumber( _
    Optional l_strType As String = "CO") As Long

    Dim l_obj As utilDBDataServices.cuDBDataServices

On Error GoTo eh
    Set l_obj = CreateObject("utilDBDataServices.cuDBDataServices")
    GetNextOpportunityNumber = l_obj.GetNextSequenceNumber(l_strType)
    Set l_obj = Nothing
    Exit Function
    
eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "GetNextOpportunityNumber"

End Function


Private Sub GetCanvassState( _
    Optional v_strStateType As String = "Create")

    Dim l_obj As dbOTCanvass.cdOTCanvass
    Dim l_datChgDate As Date

    Dim stTime As Long
    Dim ndTime As Long
    Dim dfTime As Long
    
On Error GoTo eh

    #If TRTRACE Then
        stTime = GetTickCount()
    #End If

    Set l_obj = CreateObject("dbOTCanvass.cdOTCanvass")
    Set m_rsOTOpportunity = l_obj.GetOTOpportunityForUpdate(m_lngOTOppNum)
    If RecordsExist(m_rsOTOpportunity) Then
        l_datChgDate = m_rsOTOpportunity!CHARGE_DATE
        m_intOppLocataion = m_rsOTOpportunity!LOC_CD
    Else
        m_intOppLocataion = 0
        l_datChgDate = Date
    End If
    Set m_rsOvertimeResponses = GetOvertimeResponses
    If v_strStateType = "Create" Then
        Set m_rsCanvassSheet = l_obj.GetOTOppCanvassSheetForUpdate(m_lngOTOppNum)
    Else
        Set m_rsCanvassSheet = l_obj.GetOTOppCanvassSheetForUpdate(m_lngOTOppNum)
        Set m_rsTimeCharges = l_obj.GetOTOppChargesForUpdate(m_lngOTOppNum)
        Set m_rsSchedOvertime = l_obj.GetScheduledOvertimeForUpdate(m_lngOTOppNum)
        Set m_rsAllExceptions = l_obj.GetAllExceptions
        Set m_rsEmpWorkgroup = _
                l_obj.GetEmployeePrimaryLocation(m_lngOTOppNum, l_datChgDate)
        Set m_rsTTALocations = _
                l_obj.GetTTATravelLocations(CLng(DatePart("yyyy", l_datChgDate)))
    End If
    Set l_obj = Nothing
    
    #If TRTRACE Then
        ndTime = GetTickCount()
        dfTime = ndTime - stTime
        TraceAction App.ThreadID, "3", "GetCanvassState", Format(Now, "mm/dd/yyyy hh:mm:ss"), _
            CStr(dfTime / 1000)
    #End If
    
    Exit Sub
    
eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "GetCanvassState"


End Sub


Private Sub FreeGlobalRS()

        Set m_rsErrors = Nothing
        Set m_rsOpportunityResponses = Nothing
        Set m_rsOTOpportunity = Nothing
        Set m_rsCanvassSheet = Nothing
        Set m_rsTimeCharges = Nothing
        
        Set m_rsEmpSchedules = Nothing
        Set m_rsEmpOtherPay = Nothing
        Set m_rsEmployees = Nothing
        Set m_rsAllExceptions = Nothing
        Set m_rsEmpWorkgroup = Nothing
        Set m_rsSchedOvertime = Nothing
        Set m_rsOvertimeResponses = Nothing
        Set m_rsSysDeletedTimeCharges = Nothing
        Set m_setChargeDates = Nothing
        Set m_rsEMPRVOTSchedule = Nothing
        Set m_rsJobTaskOTSchedule = Nothing
        Set m_rsEmpValidation = Nothing
        Set m_rsTTALocations = Nothing

End Sub


Private Function DetermineEmployeesOnCanvass( _
                    ByVal v_rsEmpsOppChargeDate As ADODB.Recordset _
                ) As Variant
                
 Dim l_colEmployees As Collection
 Dim l_lngEmployee As Long
 
On Error GoTo eh

    If Not RecordsExist(v_rsEmpsOppChargeDate) Then
        DetermineEmployeesOnCanvass = Empty
        Exit Function
    End If
    
    Set l_colEmployees = New Collection
    v_rsEmpsOppChargeDate.MoveFirst
    While Not v_rsEmpsOppChargeDate.EOF
        l_lngEmployee = v_rsEmpsOppChargeDate!EMP_NUM
        l_colEmployees.Add l_lngEmployee
        v_rsEmpsOppChargeDate.MoveNext
    Wend
    
    DetermineEmployeesOnCanvass = ArrayFromCollection(l_colEmployees)
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "DetermineEmployeesOnCanvass"

End Function


Private Sub GetCanvassBank( _
    ByVal v_intOTGroup As Integer, _
    ByRef r_strBank As String, _
    ByRef r_intBankCD As Integer)
    
        
Dim l_rs As ADODB.Recordset
Dim l_obj As dbOTCanvass.cdOTCanvass
        
On Error GoTo eh
    
    Set l_obj = CreateObject("dbOTCanvass.cdOTCanvass")
    Set l_rs = l_obj.GetOTGroupCanvassBank(v_intOTGroup)
    Set l_obj = Nothing
    If RecordsExist(l_rs) Then
        r_strBank = Trim(l_rs!PROJ_OR_STAND_IND)
        r_intBankCD = l_rs!OT_BANK_CD
    Else
        r_strBank = "STAND"
        r_intBankCD = 1
    End If
    Set l_rs = Nothing
    
    Exit Sub
    
eh:

    TrsRaiseError m_strCMODULENAME, "GetCanvassBank"

    
End Sub


Private Function GetPeopleSoftPhoneNumber( _
    ByVal v_aryEmps As Variant) As ADODB.Recordset
    
   

    Dim l_obj As dbPeopleSoft.cdPeopleSoft
    Dim l_rs As ADODB.Recordset
    
    
    On Error GoTo eh


    Set l_obj = CreateObject("dbPeopleSoft.cdPeopleSoft")
    Set l_rs = l_obj.GetPeopleSoftPhoneNumber(v_aryEmps)
    Set GetPeopleSoftPhoneNumber = l_rs
    Set l_obj = Nothing
    
    Exit Function
eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "GetPeopleSoftPhoneNumber"

End Function

Private Function ExtractPhone( _
    ByVal v_lngEmployee As Long, _
    ByVal v_rsPhone As ADODB.Recordset _
    ) As String
    
Dim l_strPhone As String

On Error GoTo eh

    If Not RecordsExist(v_rsPhone) Then
        ExtractPhone = ""
        Exit Function
    End If
    
    v_rsPhone.Filter = "EMP_NUM = " & str(v_lngEmployee)
    If Not RecordsExist(v_rsPhone) Then
        l_strPhone = ""
    Else
        l_strPhone = Trim(v_rsPhone!HOME_PHONE_NUM)
    End If
    v_rsPhone.Filter = ""
    
    ExtractPhone = l_strPhone

    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "ExtractPhone"
    
End Function


Private Function ProcessResponses( _
                    ByVal v_rsResp As ADODB.Recordset, _
                    ByVal v_rsTimeChargeReq As ADODB.Recordset _
                ) As Boolean
    
    Dim l_strRecStatus As String
    Dim l_blnError As Boolean
    
On Error GoTo eh
    
    v_rsResp.MoveFirst
    Do Until v_rsResp.EOF
        If FindCanvassResponse(v_rsResp!KEY_EMP_NUM, m_rsCanvassSheet) Then
            l_strRecStatus = DetermineRequestAction(v_rsResp, m_rsCanvassSheet)
            Select Case l_strRecStatus
                Case "I"
                    AddResponse v_rsResp, v_rsTimeChargeReq
                Case "D"
                    DeleteResponse v_rsResp, v_rsTimeChargeReq
                Case "C"
                    ChangeResponse v_rsResp, v_rsTimeChargeReq
            End Select
        Else
        'BRVs were written in called module
            l_blnError = True
        End If
        v_rsResp.MoveNext
    Loop
    
    ProcessResponses = Not l_blnError
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "ProcessResponses"
    
End Function


'Private Function DetermineRequestType( _
'                    ByVal v_rsResp As ADODB.Recordset, _
'                    ByVal v_rsCanvassSheet As ADODB.Recordset, _
'                    ByRef r_strRespType As String) As String
''Answer "I", "D", or "C" based on the most recent response.  There may only be 2
''responses and canvass response is always before force.
'
''I am not sure r_strRespType is useful anymore?
'
'    Dim l_strOrigCanvRespCd As String
'    Dim l_strReqCanvRespCd As String
'    Dim l_strOrigForceRespCd As String
'    Dim l_strReqForceRespCd As String
'
'On Error GoTo eh
'
'    l_strOrigCanvRespCd = Trim(v_rsCanvassSheet!CANVASS_RESP_CD)
'    l_strReqCanvRespCd = Trim(v_rsResp!CANVASS_RESP_CD)
'    l_strOrigForceRespCd = Trim(v_rsCanvassSheet!FORCE_RESP_CD)
'    l_strReqForceRespCd = Trim(v_rsResp!FORCE_RESP_CD)
'
'    If l_strOrigForceRespCd <> l_strReqForceRespCd Then
'        r_strRespType = "FORCE_RESP_CD"
'        DetermineRequestType = _
'            DetermineRequestAction(l_strOrigRespCd, l_strReqRespCd)
'    Else
'        r_strRespType = "CANVASS_RESP_CD"
'        l_strOrigRespCd = Trim(v_rsCanvassSheet!CANVASS_RESP_CD)
'        l_strReqRespCd = Trim(v_rsResp!CANVASS_RESP_CD)
'        DetermineRequestType = _
'            DetermineRequestAction(l_strOrigRespCd, l_strReqRespCd)
'    End If
'
'    Exit Function
'
'eh:
'    TrsRaiseError m_strCMODULENAME, "DetermineRequestType"
'End Function


Private Function DetermineRequestAction( _
                    ByVal v_rsResp As ADODB.Recordset, _
                    ByVal v_rsCanvassSheet As ADODB.Recordset _
                ) As String
'Answer "I", "D", or "C".

    Dim l_strOrigCanvRespCd As String
    Dim l_strReqCanvRespCd As String
    Dim l_strOrigForceRespCd As String
    Dim l_strReqForceRespCd As String

On Error GoTo eh

    l_strOrigCanvRespCd = Trim(v_rsCanvassSheet!CANVASS_RESP_CD)
    l_strReqCanvRespCd = Trim(v_rsResp!CANVASS_RESP_CD)
    l_strOrigForceRespCd = Trim(v_rsCanvassSheet!FORCE_RESP_CD)
    l_strReqForceRespCd = Trim(v_rsResp!FORCE_RESP_CD)

    If (l_strOrigCanvRespCd = "" And l_strOrigForceRespCd = "") And _
            (l_strReqCanvRespCd <> "" Or l_strReqForceRespCd <> "") Then
        DetermineRequestAction = "I"
    ElseIf (l_strOrigCanvRespCd <> "" Or l_strOrigForceRespCd <> "") And _
            (l_strReqCanvRespCd = "" And l_strReqForceRespCd = "") Then
        DetermineRequestAction = "D"
    Else
        DetermineRequestAction = "C"
    End If

    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "DetermineRequestAction"
End Function


'Private Function DetermineRequestAction( _
'                    ByVal v_strOrigRespCd As String, _
'                    ByVal v_strReqRespCd As String) As String
''Answer "I", "D", or "C".
'
'On Error GoTo eh
'
'    If v_strOrigRespCd = "" And v_strReqRespCd <> "" Then
'        DetermineRequestAction = "I"
'    ElseIf v_strReqRespCd = "" And v_strOrigRespCd <> "" Then
'        DetermineRequestAction = "D"
'    Else
'        DetermineRequestAction = "C"
'    End If
'
'    Exit Function
'
'eh:
'    TrsRaiseError m_strCMODULENAME, "DetermineRequestAction"
'End Function


'Private Function PreValidateResponse( _
'                    ByVal v_rsResp As ADODB.Recordset _
'                ) As Boolean
'
'    Dim l_strRecStatus As String
'    Dim l_blnError As Boolean
'    Dim l_strCreatedTS As String
'    Dim l_strLastModTS As String
'    Dim l_strFilter As String
'    Dim l_strModFilter As String
'    Dim l_lngEmpNum As Long
'    Dim l_lngOTOppNum As Long
'    Dim l_strPersitentLastModTS As String
'    Dim l_strResp As String
'    Dim l_strCallStart As String
'    Dim l_strCallNotes As String
'    Dim l_strErrorMessage As String
'
'On Error GoTo eh
'
'    If IsNull(v_rsResp!RECORD_STATUS) Or IsEmpty(v_rsResp!RECORD_STATUS) Then
'        'Write BRV for Bad Record Status
'        l_strErrorMessage = "Record Status is missing"
'        AppendResponseEntryError v_rsResp, 30000, mstrCAPPLICATIONERR, l_strErrorMessage
'        PreValidateResponse = False
'        Exit Function
'    Else
'        If UCase(Trim(v_rsResp!RECORD_STATUS)) <> "C" Then
'        'Write BRV for Bad Record Status
'            l_strErrorMessage = "Record Status is not valid. Change only check"
'            AppendResponseEntryError v_rsResp, 30000, mstrCAPPLICATIONERR, l_strErrorMessage
'            PreValidateResponse = False
'            Exit Function
'        End If
'    End If
'
'    If Not ValidResponseKeyFields(v_rsResp) Then
'        'BRVs written in called procedure
'        PreValidateResponse = False
'        Exit Function
'    End If
'
'    l_strRecStatus = v_rsResp!RECORD_STATUS
'    l_lngOTOppNum = v_rsResp!OT_OPP_NUM
'    l_lngEmpNum = v_rsResp!EMP_NUM
'    l_strFilter = "OT_OPP_NUM = " & CStr(l_lngOTOppNum) & _
'            " AND EMP_NUM = " & CStr(l_lngEmpNum)
'    l_strCreatedTS = Trim(v_rsResp!CREATED_TIMESTAMP)
'    l_strLastModTS = Trim(v_rsResp!LAST_MOD_TIMESTAMP)
'    l_strModFilter = l_strFilter & _
'            " AND CREATED_TIMESTAMP = '" & l_strCreatedTS & "'"
'    l_strFilter = MDAC25FilterFix(l_strFilter)
'    l_strModFilter = MDAC25FilterFix(l_strModFilter)
'
''        Position canvass sheet on change request
'    m_rsCanvassSheet.Filter = l_strModFilter
'    If Not RecordsExist(m_rsCanvassSheet) Then
'        l_blnError = True
'        'write BRV Record not found for change request
'        l_strErrorMessage = "Please refresh your screen.  " & _
'                "You are trying to change a response that does not exist. " & _
'                "The response was probably deleted by someone else."
'        AppendResponseEntryError v_rsResp, 30000, mstrCBUSINESSERR, l_strErrorMessage
'    Else
'        If l_strLastModTS <> "0001-01-01-00.00.00.000000" Then
'            l_strPersitentLastModTS = Trim(m_rsCanvassSheet!TIME_STAMP)
'            If l_strPersitentLastModTS <> l_strLastModTS Then
'                l_blnError = True
'            'write BRV Data changed prior to the change request
'            l_strErrorMessage = "Please refresh your screen and verify the correctness of the response. " & _
'                    "You were trying to change the reponse, but someone else (" & _
'                    Trim(str(m_rsCanvassSheet!LAST_MOD_EMP_NUM)) & _
'                    ") changed this reponse during your request. "
'            AppendResponseEntryError v_rsResp, 30000, mstrCBUSINESSERR, l_strErrorMessage
'            End If
'        End If
'    End If
'
'    If IsEmpty(v_rsResp!CANVASS_RESP_CD) Or IsNull(v_rsResp!CANVASS_RESP_CD) Then
'        l_blnError = True
'        'write missing CANVASS_RESP_CD BRV
'        l_strErrorMessage = "Response is missing"
'            AppendResponseEntryError v_rsResp, 30000, mstrCAPPLICATIONERR, l_strErrorMessage
'    Else
'        l_strResp = v_rsResp!CANVASS_RESP_CD
'        If Not IsValidResponse(l_strResp) Then
'            l_blnError = True
'        'BRV was written in called module
'        End If
'    End If
'
''The Call start does not have to be provided
'    If IsEmpty(v_rsResp!CANVASS_CALL_TIME) Or IsNull(v_rsResp!CANVASS_CALL_TIME) Then
'        v_rsResp!CANVASS_CALL_TIME = ""
'    Else
'        l_strCallStart = v_rsResp!CANVASS_CALL_TIME
'        If Not IsAlphaNumeric(l_strCallStart) Then
'             v_rsResp!CANVASS_CALL_TIME = ""
'        End If
'    End If
'
''The Call Notes do not have to be provided
'    If IsEmpty(v_rsResp!CALL_NOTES_TXT) Or IsNull(v_rsResp!CALL_NOTES_TXT) Then
'        v_rsResp!CALL_NOTES_TXT = ""
'    Else
'        l_strCallNotes = v_rsResp!CALL_NOTES_TXT
'        If Not IsAlphaNumeric(l_strCallNotes) Then
'            v_rsResp!CALL_NOTES_TXT = ""
'        End If
'    End If
'
'    PreValidateResponse = Not l_blnError
'    Exit Function
'
'eh:
'    TrsRaiseError m_strCMODULENAME, "PreValidateResponse"
'
'End Function


Private Sub AddResponse( _
                ByVal v_rsResp As ADODB.Recordset, _
                ByVal v_rsTimeChargeReq As ADODB.Recordset)

    Dim l_lngEmpNum As Long
    Dim l_strRespCd As String
    Dim l_blnTimeChargeGen As Boolean
    Dim l_typCanvShift As CanvassedShift

On Error GoTo eh

    l_lngEmpNum = v_rsResp!EMP_NUM
    l_strRespCd = MostRecentResponseCode(v_rsResp!CANVASS_RESP_CD, v_rsResp!FORCE_RESP_CD)
    
    UpdateCanvassSheet v_rsResp
    DeleteEmployeeTimeCharges l_lngEmpNum, True     'Here for adding force resp
    ProcessTimeCharges l_lngEmpNum, v_rsTimeChargeReq
    If IsAcceptorResponse(l_strRespCd) Then
        If OvertimeRateCalculated(l_lngEmpNum, m_rsOTOpportunity!CHARGE_DATE) Then
            If Not TimeChargeRecordsExist(l_lngEmpNum) Then
                l_blnTimeChargeGen = True
                AddTimeChargeFromOpportunity l_lngEmpNum
            End If
            ProcessMissingComponents l_lngEmpNum, l_blnTimeChargeGen
        End If
    End If


xit:
    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "AddResponse"
End Sub


Private Sub ProcessMissingComponents( _
                    ByVal v_lngEmpNum As Long, _
                    ByVal v_blnTimeChgGen As Boolean)

    Dim l_typCanvShift As CanvassedShift
    
On Error GoTo eh

'may need to allow multiple shifts on a single opportunity.
'assuming only 1 for now.

    m_rsTimeCharges.Filter = "EMP_NUM =" & v_lngEmpNum
    m_rsTimeCharges.Sort = "DATE_WORKED ASC, START_TIME ASC"
'position on first new or modified time charge
    Do Until m_rsTimeCharges.EOF
        If m_rsTimeCharges.Status = adRecNew Or _
                m_rsTimeCharges.Status = adRecModified Then
            If IsOvertimeWorked(m_rsTimeCharges!EXCEPTION_CD, m_rsTimeCharges!OVER_EXC_CD) Or _
                    Trim(m_rsTimeCharges!EXCEPTION_CD) = "" Then
                Exit Do
            End If
        End If
        m_rsTimeCharges.MoveNext
    Loop
    If m_rsTimeCharges.EOF Then
        GoTo xit
    End If
    
    If v_blnTimeChgGen Then
        l_typCanvShift = NewCanvassedShiftFromOpportunity(m_rsOTOpportunity)
    Else
        l_typCanvShift = NewCanvassedShiftFromTimeCharges(v_lngEmpNum, m_rsTimeCharges)
    End If

    Do Until m_rsTimeCharges.EOF
        If m_rsTimeCharges.Status = adRecNew Or _
                m_rsTimeCharges.Status = adRecModified Then
            If IsOvertimeWorked(m_rsTimeCharges!EXCEPTION_CD, m_rsTimeCharges!OVER_EXC_CD) Or _
                    Trim(m_rsTimeCharges!EXCEPTION_CD) = "" Then
                DetermineMissingComponents l_typCanvShift, m_rsTimeCharges
                UpdateSchedulesWithMissingComponents m_rsTimeCharges
            End If
        End If
        m_rsTimeCharges.MoveNext
    Loop

xit:

    m_rsTimeCharges.Sort = ""
    m_rsTimeCharges.Filter = adFilterNone
    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "ProcessMissingComponents"

End Sub


Private Function ValidResponseKeyFields( _
                    ByVal v_rsResp As ADODB.Recordset _
                ) As Boolean

    Dim l_blnValidationError As Boolean
    Dim l_strErrorMessage As String

On Error GoTo eh

    If Not IsNumeric(v_rsResp!KEY_OT_OPP_NUM) Or _
            v_rsResp!KEY_OT_OPP_NUM = 0 Then
        l_strErrorMessage = "KEY_OT_OPP_NUM is not valid. Input=" & CStr(v_rsResp!KEY_OT_OPP_NUM)
        AppendResponseRequestBRV v_rsResp, trErrbusOTCanvass, mstrCAPPLICATIONERR, l_strErrorMessage
        l_blnValidationError = True
    End If
    
    If Not IsNumeric(v_rsResp!KEY_EMP_NUM) Or _
            v_rsResp!KEY_EMP_NUM = 0 Then
        l_strErrorMessage = "KEY_EMP_NUM is not valid. Input=" & CStr(v_rsResp!KEY_EMP_NUM)
        AppendResponseRequestBRV v_rsResp, trErrbusOTCanvass, mstrCAPPLICATIONERR, l_strErrorMessage
        l_blnValidationError = True
    End If

    If Not IsDB2Timestamp(v_rsResp!CREATED_TIMESTAMP) Then
        l_strErrorMessage = "CREATED_TIMESTAMP is not valid. Input=" & CStr(v_rsResp!CREATED_TIMESTAMP)
        AppendResponseRequestBRV v_rsResp, trErrbusOTCanvass, mstrCAPPLICATIONERR, l_strErrorMessage
        l_blnValidationError = True
    End If
    
    If Not IsDB2Timestamp(v_rsResp!LAST_MOD_TIMESTAMP) Then
        l_strErrorMessage = "LAST_MOD_TIMESTAMP is not valid. Input=" & CStr(v_rsResp!LAST_MOD_TIMESTAMP)
        AppendResponseRequestBRV v_rsResp, trErrbusOTCanvass, mstrCAPPLICATIONERR, l_strErrorMessage
        l_blnValidationError = True
    End If
  
    ValidResponseKeyFields = Not l_blnValidationError
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "ValidResponseKeyFields"
End Function



'*************************************************************************
'NON Canvassed OT

Public Function CreateNonCanvassOvertime( _
                    ByRef r_lngOTOppNum As Long, _
                    ByVal v_rsOTOppCriteria As ADODB.Recordset, _
                    ByVal v_rsEmpNums As ADODB.Recordset, _
                    ByVal v_strUserId As String, _
                    ByRef r_rsBRVs As ADODB.Recordset _
                ) As Boolean
        
 On Error GoTo eh
 
    Dim l_blnTemplateOK As Boolean
    Dim l_datOppChargeDate As Date
    Dim l_datCanvassDate As Date
    Dim l_rsEmpsCanvassDate As ADODB.Recordset
    Dim l_rsEmpsOppChargeDate As ADODB.Recordset
    Dim l_aryEmployees As Variant
    Dim l_lngTemplateNum As Long
    Dim l_lngOTOppNum As Long
    Dim l_intOTGroup As Integer
    
    Dim l_rsTemplateBase As ADODB.Recordset
    Dim l_rsSortOrder As ADODB.Recordset
    Dim l_rsForceSortOrder As ADODB.Recordset
    Dim l_rsGroupSortOrder As ADODB.Recordset
    Dim l_rsSortDefinitions As ADODB.Recordset
    Dim l_rsCanvassSheet As ADODB.Recordset
    Dim l_rsOpportunity As ADODB.Recordset
    Dim l_rsOTOppInelig As ADODB.Recordset
    Dim l_rsPhoneNumbers As ADODB.Recordset
    Dim l_intCanvassLoc As Integer
    Dim l_strOTReasonCd As String
    Dim objCtx As ObjectContext
    Dim l_intBankCD As Integer
    Dim l_intRefuseRule As Integer
    
    
    Set objCtx = GetObjectContext
    
    m_strUserID = v_strUserId
    Set m_rsErrors = CreateNewBusinessErrorSet()
    
'PTS1596
    If Not PreValidateOppCriteria(v_rsOTOppCriteria, 1) Then
        CreateNonCanvassOvertime = False
        objCtx.SetAbort
        GoTo xit
    End If
        
    l_strOTReasonCd = v_rsOTOppCriteria!OT_REASON_CD
    l_lngTemplateNum = v_rsOTOppCriteria!TEMPLATE_NUM
    l_datOppChargeDate = v_rsOTOppCriteria!CHARGE_DATE
    l_datCanvassDate = Date
    l_intOTGroup = v_rsOTOppCriteria!OT_GROUP_NUM
'    GetOvertimeTemplateData l_lngTemplateNum, l_rsTemplateBase, l_rsSortOrder, _
'        l_rsForceSortOrder, l_rsGroupSortOrder, l_rsSortDefinitions
'RSG
    GetNonCanvassOvertimeTemplateData l_lngTemplateNum, l_rsTemplateBase
    
    If Not RecordsExist(l_rsTemplateBase) Then
        'writeBRV
        objCtx.SetAbort
        Set objCtx = Nothing
        Exit Function
    End If
    
''Inorder for an employee to appear on a canvass they must be in the overtime group
''When the overtime is going to be worked (l_datOppChargeDate)
'' and when the overtime is being canvassed (l_datCanvassDate)
'    Set l_rsEmpsOppChargeDate = GetEmployeesInOTGroup(l_datOppChargeDate, l_intOTGroup)
'    If l_datOppChargeDate < l_datCanvassDate Then
'        Set l_rsEmpsCanvassDate = GetEmployeesInOTGroup(l_datCanvassDate, l_intOTGroup)
'    Else
'       Set l_rsEmpsCanvassDate = Nothing
'    End If
'    l_aryEmployees = DetermineEmployeesOnCanvass(l_rsEmpsOppChargeDate, l_rsEmpsCanvassDate)
    l_aryEmployees = ArrayOfFieldValues("EMP_NUM", v_rsEmpNums)
        
        
    If IsEmpty(l_aryEmployees) Then
        Exit Function ' No employees met criteria WriteBRV
    End If
    
    Set l_rsPhoneNumbers = Nothing
'    Set l_rsPhoneNumbers = GetPeopleSoftPhoneNumber(l_aryEmployees)

    Set l_rsCanvassSheet = ConstructNonCanvassSheet( _
                            l_aryEmployees, l_rsTemplateBase, _
                            v_rsOTOppCriteria, l_intBankCD)
    SortNonCanvassSheet v_rsEmpNums, l_rsCanvassSheet
    
    m_lngOTOppNum = GetNextOpportunityNumber("NO")
    GetCanvassState
    
    l_intCanvassLoc = l_rsTemplateBase!TEMPL_LOCATION
    l_intRefuseRule = l_rsTemplateBase!REFUSE_RULE_NUM
    
    
    BuildOpportunity l_intCanvassLoc, l_intRefuseRule, v_rsOTOppCriteria

    Set l_rsPhoneNumbers = GetPeopleSoftPhoneNumber(l_aryEmployees)
    
    BuildCanvassSheet l_strOTReasonCd, l_rsCanvassSheet, l_rsPhoneNumbers, l_intBankCD, l_rsOTOppInelig

    If SaveOpportunity() Then
        'continue
    Else
        'raise error
        CreateNonCanvassOvertime = False
        objCtx.SetAbort
        Set objCtx = Nothing
        Exit Function
    End If
    
    If SaveCanvassSheet() Then
            'continue
    Else
        'raise error
        CreateNonCanvassOvertime = False
        objCtx.SetAbort
        Set objCtx = Nothing
        Exit Function
    End If
    
'    If SaveOppInelig() Then
'            'continue
'    Else
'        'raise error
'        CreateOvertimeCanvass = False
'        Exit Function
'    End If
'
    
    r_lngOTOppNum = m_lngOTOppNum
    CreateNonCanvassOvertime = True
    objCtx.SetComplete

xit:
    If Not m_rsErrors Is Nothing Then
        If m_rsErrors.RecordCount = 0 Then
            Set r_rsBRVs = Nothing
        Else
            Set r_rsBRVs = m_rsErrors
        End If
    Else
        Set r_rsBRVs = Nothing
    End If
    Set objCtx = Nothing
    FreeGlobalRS

    Exit Function

    
eh:
    FreeGlobalRS
    objCtx.SetAbort
    Set objCtx = Nothing
    TrsRaiseError m_strCMODULENAME, "CreateNonCanvassOvertime", , , , , True
    
End Function


Private Sub GetNonCanvassOvertimeTemplateData( _
                ByVal v_lngTemplateNum As Long, _
                ByRef r_rsTemplateBase As ADODB.Recordset)
    
    Dim l_obj As dbOTTemplate.cdOTTemplate
   
On Error GoTo eh
    
    Set l_obj = CreateObject("dbOTTemplate.cdOTTemplate")
    Set r_rsTemplateBase = l_obj.GetOvertimeTemplateBase(v_lngTemplateNum)
    
    Set l_obj = Nothing
    Exit Sub

eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "GetNonCanvassOvertimeTemplateData"
    
End Sub


Private Function ConstructNonCanvassSheet( _
                    ByVal v_aryEmployees As Variant, _
                    ByVal v_rsTemplate As ADODB.Recordset, _
                    ByVal v_rsOTOpp As ADODB.Recordset, _
                    ByRef r_intBankCD As Integer _
                ) As ADODB.Recordset
    
    Dim l_obj As dbOTCanvass.cdOTCanvass
    Dim l_rsSheet As ADODB.Recordset
    Dim l_lngTemplateNum As Long
    Dim l_datOppChargeDate As Date
    Dim l_strCanvassType As String
    Dim l_strBank As String
    Dim l_intProjectID As Integer
    Dim l_intTemplLoc As Integer
    Dim l_intOTGroup As Integer
    Dim l_rsWGSort As ADODB.Recordset
    
On Error GoTo eh

    l_lngTemplateNum = v_rsTemplate!TEMPLATE_NUM
    l_datOppChargeDate = v_rsOTOpp!CHARGE_DATE
    l_intTemplLoc = v_rsTemplate!TEMPL_LOCATION
    l_intOTGroup = v_rsOTOpp!OT_GROUP_NUM
    GetCanvassBank l_intOTGroup, l_strBank, r_intBankCD
    
    Set l_obj = CreateObject("dbOTCanvass.cdOTCanvass")
    
    Set l_rsSheet = l_obj.ConstructCanvassSheet( _
                        v_aryEmployees, l_datOppChargeDate, r_intBankCD)
    
    If RecordsExist(l_rsSheet) Then
        'continue
    Else
        'probably should raise error
        Set l_obj = Nothing
        Set ConstructNonCanvassSheet = Nothing
        Exit Function
    End If
'    Set l_rsWGSort = l_obj.GetWorkGroupSortData(v_aryEmployees, l_intOTGroup, l_datOppChargeDate)
'    UpdateSortColumns l_rsSheet, l_rsWGSort, l_intTemplLoc
    
    Set ConstructNonCanvassSheet = l_rsSheet
    Exit Function
    
eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "ConstructNonCanvassSheet"
        
End Function


Private Sub SortNonCanvassSheet( _
            ByVal v_rsEmpNums As ADODB.Recordset, _
            ByRef r_rsCanvassSheet As ADODB.Recordset)
'Make the canvass sheet order the same as the original list of emp nums.

On Error GoTo eh
    
    v_rsEmpNums.MoveFirst
    Do Until v_rsEmpNums.EOF
        r_rsCanvassSheet.Filter = "EMP_NUM = " & v_rsEmpNums!EMP_NUM
        Do Until r_rsCanvassSheet.EOF
            r_rsCanvassSheet!ROW_NUM = v_rsEmpNums.AbsolutePosition
            r_rsCanvassSheet!FORCE_ROW_NUM = v_rsEmpNums.AbsolutePosition
            r_rsCanvassSheet.MoveNext
        Loop
        v_rsEmpNums.MoveNext
    Loop
    
    r_rsCanvassSheet.Filter = ""
    Exit Sub
    
eh:
    
    TrsRaiseError m_strCMODULENAME, "SortNonCanvassSheet"
        
End Sub

Private Sub DeleteAllResponses()


On Error GoTo eh

    If m_rsCanvassSheet Is Nothing Then
        Exit Sub
    End If

    m_rsCanvassSheet.Filter = adFilterNone
    Do Until m_rsCanvassSheet.EOF
        If Trim(m_rsCanvassSheet!CANVASS_RESP_CD) <> "" Or _
                Trim(m_rsCanvassSheet!FORCE_RESP_CD) <> "" Then
            m_rsCanvassSheet!CANVASS_RESP_CD = ""
            m_rsCanvassSheet!FORCE_RESP_CD = ""
            m_rsCanvassSheet!CANVASS_CALL_TIME = mINITCANVCALLTIMESTAMP
            m_rsCanvassSheet!FORCE_CALL_TIME = mINITCANVCALLTIMESTAMP
            m_rsCanvassSheet!LAST_MOD_EMP_NUM = m_lngUserNum
        End If
        m_rsCanvassSheet.MoveNext
    Loop

    DeleteEmployeeTimeCharges 0, True

    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "DeleteAllResponses"

End Sub


Private Sub DeleteResponse( _
                ByVal v_rsResp As ADODB.Recordset, _
                ByVal v_rsTimeChargeReq As ADODB.Recordset)
    
'
' If deleting a response, the only time charge requests should be
' manual deletions which should be ignored since the system will
' delete all existing timecharges.
'
    
    Dim l_lngEmpNum As Long


On Error GoTo eh

    l_lngEmpNum = v_rsResp!EMP_NUM
    
' Do not initialize call time in case response is deleted by mistake,
' only to be readded.
'    v_rsResp!CANVASS_CALL_TIME = mINITCANVCALLTIMESTAMP
'    v_rsResp!FORCE_CALL_TIME = mINITCANVCALLTIMESTAMP

    UpdateCanvassSheet v_rsResp
    
    DeleteEmployeeTimeCharges l_lngEmpNum, False
    
    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "DeleteResponse"
    
End Sub

'Private Sub ChangeForcedResponse( _
'                ByVal v_rsResp As ADODB.Recordset, _
'                ByVal v_rsTimeChargeReq As ADODB.Recordset)
'
'    ChangeResponse "FORCE_RESP_CD", v_rsResp, v_rsTimeChargeReq
'
'    Exit Sub
'
'eh:
'    TrsRaiseError m_strCMODULENAME, "ChangeForcedResponse"
'End Sub


Private Sub ChangeResponse( _
                ByVal v_rsResp As ADODB.Recordset, _
                ByVal v_rsTimeChargeReq As ADODB.Recordset)

    Dim l_strCallStart As String
    Dim l_blnOrigWorkedInd As Boolean
    Dim l_blnReqWorkedInd As Boolean
    Dim l_lngEmpNum As Long
    Dim l_blnDelAcceptance As Boolean
    Dim l_blnTimeChargeGen As Boolean
    Dim l_typCanvShift As CanvassedShift
    Dim l_strRespCd As String

On Error GoTo eh
    
    l_lngEmpNum = v_rsResp!EMP_NUM

    l_strRespCd = MostRecentResponseCode(m_rsCanvassSheet!CANVASS_RESP_CD, m_rsCanvassSheet!FORCE_RESP_CD)
    l_blnOrigWorkedInd = IsOvertimeWorkedResponse(l_strRespCd)
    
    l_strRespCd = MostRecentResponseCode(v_rsResp!CANVASS_RESP_CD, v_rsResp!FORCE_RESP_CD)
    l_blnReqWorkedInd = IsOvertimeWorkedResponse(l_strRespCd)

    UpdateCanvassSheet v_rsResp

    If l_blnOrigWorkedInd <> l_blnReqWorkedInd Then
        DeleteEmployeeTimeCharges l_lngEmpNum, l_blnOrigWorkedInd
    ElseIf HasCallTimeChanged(m_rsCanvassSheet) Then
        If OvertimeRateCalculated(m_rsCanvassSheet!EMP_NUM, m_rsOTOpportunity!CHARGE_DATE) Then
            MarkAllTimeChargesForMissingComponents m_rsCanvassSheet!EMP_NUM, m_rsTimeCharges
        End If
    End If
    
    ProcessTimeCharges l_lngEmpNum, v_rsTimeChargeReq
    If l_blnReqWorkedInd Then
        If OvertimeRateCalculated(l_lngEmpNum, m_rsOTOpportunity!CHARGE_DATE) Then
            If Not TimeChargeRecordsExist(l_lngEmpNum) Then
                l_blnTimeChargeGen = True
                AddTimeChargeFromOpportunity l_lngEmpNum
            End If
            ProcessMissingComponents l_lngEmpNum, l_blnTimeChargeGen
        End If
    End If


xit:
    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "ChangeResponse"
End Sub


Private Function HasCallTimeChanged(ByVal v_rsCanvSheet As ADODB.Recordset) As Boolean

On Error GoTo eh

    HasCallTimeChanged = _
        (Trim(v_rsCanvSheet.Fields("CANVASS_CALL_TIME").OriginalValue <> _
            Trim(v_rsCanvSheet.Fields("CANVASS_CALL_TIME").Value))) Or _
        (Trim(v_rsCanvSheet.Fields("FORCE_CALL_TIME").OriginalValue <> _
            Trim(v_rsCanvSheet.Fields("FORCE_CALL_TIME").Value)))
            
    Exit Function
        
eh:
    TrsRaiseError m_strCMODULENAME, "HasCallTimeChanged"

End Function


Private Sub UpdateCanvassSheet(ByVal v_rsResp As ADODB.Recordset)

    Dim l_strCallStart As String
    
On Error GoTo eh
    
    m_rsCanvassSheet!CANVASS_RESP_CD = v_rsResp!CANVASS_RESP_CD
    l_strCallStart = Trim(v_rsResp!CANVASS_CALL_TIME)
'    If l_strCallStart <> Trim(m_rsCanvassSheet!CANVASS_CALL_TIME) Then
    FormatCallStart l_strCallStart
    m_rsCanvassSheet!CANVASS_CALL_TIME = l_strCallStart
'    End If
    
    m_rsCanvassSheet!FORCE_RESP_CD = v_rsResp!FORCE_RESP_CD
    l_strCallStart = Trim(v_rsResp!FORCE_CALL_TIME)
'    If l_strCallStart <> Trim(m_rsCanvassSheet!FORCE_CALL_TIME) Then
    FormatCallStart l_strCallStart
    m_rsCanvassSheet!FORCE_CALL_TIME = l_strCallStart
'    End If
    
    m_rsCanvassSheet!CALL_NOTES_TXT = Trim(v_rsResp!CALL_NOTES_TXT)
    m_rsCanvassSheet!LAST_MOD_EMP_NUM = m_lngUserNum
    
    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "UpdateCanvassSheet"
    
End Sub


Private Sub DeleteGeneratedRefusalCharges( _
                Optional ByVal v_lngEmpNum As Long = 0)
'If v_lngEmpNum = 0, delete all timecharges.

    Dim l_rsTimeChargesC As ADODB.Recordset

On Error GoTo eh

    Set l_rsTimeChargesC = m_rsTimeCharges.Clone
    If v_lngEmpNum = 0 Then
        l_rsTimeChargesC.Filter = adFilterNone
    Else
        l_rsTimeChargesC.Filter = "EMP_NUM = " & v_lngEmpNum
    End If
    
    If RecordsExist(l_rsTimeChargesC) Then
        Set m_rsSysDeletedTimeCharges = NewSysDelTimeChgRecordset
    End If
    
    Do Until l_rsTimeChargesC.EOF
        If IsSystemGeneratedRefusal(l_rsTimeChargesC!EXCEPTION_CD, l_rsTimeChargesC!LAST_MOD_EMP_NUM) Then
            CopyTimeChargeKeyData l_rsTimeChargesC, m_rsSysDeletedTimeCharges
            m_rsTimeCharges.Bookmark = l_rsTimeChargesC.Bookmark
            m_rsTimeCharges.Delete
        End If
        l_rsTimeChargesC.MoveNext
    Loop

    Set l_rsTimeChargesC = Nothing
    
    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "DeleteGeneratedRefusalCharges"
    
End Sub


Private Sub DeleteEmployeeTimeCharges( _
                ByVal v_lngEmpNum As Long, _
                ByVal v_blnDelAcceptance As Boolean)
'If v_lngEmpNum = 0, delete all timecharges.

    Dim l_rsTimeChargesC As ADODB.Recordset

On Error GoTo eh

    Set l_rsTimeChargesC = m_rsTimeCharges.Clone
    If v_lngEmpNum > 0 Then
        l_rsTimeChargesC.Filter = "EMP_NUM = " & v_lngEmpNum
    End If
    
    If RecordsExist(l_rsTimeChargesC) Then
        Set m_rsSysDeletedTimeCharges = NewSysDelTimeChgRecordset
    End If
    
    Do Until l_rsTimeChargesC.EOF
        CopyTimeChargeKeyData l_rsTimeChargesC, m_rsSysDeletedTimeCharges
        
'RSG - needs work
' v_blnDelAcceptance no longer used
        
        If IsOvertimeWorked(l_rsTimeChargesC!EXCEPTION_CD, l_rsTimeChargesC!OVER_EXC_CD) Or _
                Trim(l_rsTimeChargesC!EXCEPTION_CD) = "" Then
'        If v_blnDelAcceptance Then
            DeleteOvertimeSchedule l_rsTimeChargesC!EMP_NUM, l_rsTimeChargesC!OT_OPP_NUM, _
                l_rsTimeChargesC!DATE_WORKED, l_rsTimeChargesC!START_TIME, m_rsSchedOvertime
            DeleteAdjustedScheduleOvertime l_rsTimeChargesC!EMP_NUM, l_rsTimeChargesC!OT_OPP_NUM, _
                l_rsTimeChargesC!DATE_WORKED, l_rsTimeChargesC!START_TIME, m_rsEmpSchedules
        End If
        
        m_rsTimeCharges.Bookmark = l_rsTimeChargesC.Bookmark
        m_rsTimeCharges.Delete
        l_rsTimeChargesC.MoveNext
    Loop

    Set l_rsTimeChargesC = Nothing
    
    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "DeleteEmployeeTimeCharges"
    
End Sub


Private Function NewSysDelTimeChgRecordset() As ADODB.Recordset

    Dim l_rs As ADODB.Recordset

On Error GoTo eh
   
    If Not m_rsSysDeletedTimeCharges Is Nothing Then
        Set NewSysDelTimeChgRecordset = m_rsSysDeletedTimeCharges
    End If
   
    Set l_rs = New Recordset
    With l_rs
        .Fields.Append "OT_OPP_NUM", adInteger, , mCFLDATTRUPDATE
        .Fields.Append "EMP_NUM", adInteger, mCFLDATTRUPDATE
        .Fields.Append "CHARGE_DATE", adDate, mCFLDATTRUPDATE
        .Fields.Append "DATE_WORKED", adDate, mCFLDATTRUPDATE
        .Fields.Append "START_TIME", adChar, 5, mCFLDATTRUPDATE
        .Fields.Append "EXCEPTION_CD", adChar, 3, mCFLDATTRUPDATE
        .Fields.Append "CREATED_TIMESTAMP", adChar, 26, mCFLDATTRUPDATE
        .Fields.Append "LAST_MOD_TIMESTAMP", adChar, 26, mCFLDATTRUPDATE

        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open
    End With
    
    Set NewSysDelTimeChgRecordset = l_rs
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "NewSysDelTimeChgRecordset"
    
End Function


Private Sub CopyTimeChargeKeyData( _
                ByVal v_rsTimeCharges As ADODB.Recordset, _
                ByRef r_rsSysDelTimeCharge As ADODB.Recordset)

On Error GoTo eh

    r_rsSysDelTimeCharge.AddNew
    r_rsSysDelTimeCharge!OT_OPP_NUM = v_rsTimeCharges!OT_OPP_NUM
    r_rsSysDelTimeCharge!EMP_NUM = v_rsTimeCharges!EMP_NUM
    r_rsSysDelTimeCharge!CHARGE_DATE = v_rsTimeCharges!CHARGE_DATE
    r_rsSysDelTimeCharge!DATE_WORKED = v_rsTimeCharges!DATE_WORKED
    r_rsSysDelTimeCharge!START_TIME = v_rsTimeCharges!START_TIME
    r_rsSysDelTimeCharge!EXCEPTION_CD = v_rsTimeCharges!EXCEPTION_CD
    r_rsSysDelTimeCharge!CREATED_TIMESTAMP = v_rsTimeCharges!CREATED_TIMESTAMP
    r_rsSysDelTimeCharge!LAST_MOD_TIMESTAMP = v_rsTimeCharges!LAST_MOD_TIMESTAMP
    r_rsSysDelTimeCharge.Update

    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "CopyTimeChargeKeyData"
    
End Sub


Private Function IsSystemDeletedTimeCharge( _
                    ByVal v_lngEmpNum As Long, _
                    ByVal v_datChargeDate As Date, _
                    ByVal v_datDateWorked As Date, _
                    ByVal v_strStartTime As String) As Boolean


On Error GoTo eh

    FindTimeCharge v_lngEmpNum, v_datChargeDate, v_datDateWorked, _
        v_strStartTime, m_rsSysDeletedTimeCharges

    IsSystemDeletedTimeCharge = Not m_rsSysDeletedTimeCharges.EOF

    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "IsSystemDeletedTimeCharge"
End Function


'Private Sub DeleteOvertimeSchedule( _
'                ByVal v_rsTimeChargesReq As ADODB.Recordset, _
'                ByRef r_rsSchedOvertime As ADODB.Recordset)
'
'On Error GoTo eh
'
'    r_rsSchedOvertime.Filter = "EMP_NUM = " & v_rsTimeChargesReq!KEY_EMP_NUM
'
'    Do Until r_rsSchedOvertime.EOF
'        If r_rsSchedOvertime!OT_OPP_NUM = v_rsTimeChargesReq!KEY_OT_OPP_NUM And _
'                r_rsSchedOvertime!Date = v_rsTimeChargesReq!KEY_DATE_WORKED And _
'                r_rsSchedOvertime!START_TIME = v_rsTimeChargesReq!KEY_START_TIME Then
'            r_rsSchedOvertime.Delete
'            Exit Do
'        End If
'        r_rsSchedOvertime.MoveNext
'    Loop
'    r_rsSchedOvertime.Filter = adFilterNone
'
'    Exit Sub
'
'eh:
'    TrsRaiseError m_strCMODULENAME, "DeleteOvertimeSchedule"
'
'End Sub


Private Sub DeleteOvertimeSchedule( _
                ByVal v_lngEmpNum As Long, _
                ByVal v_lngOTOppNum As Long, _
                ByVal v_datDateWorked As Date, _
                ByVal v_strStartTime As String, _
                ByRef r_rsSchedOvertime As ADODB.Recordset)

On Error GoTo eh

    r_rsSchedOvertime.Filter = "EMP_NUM = " & v_lngEmpNum

    Do Until r_rsSchedOvertime.EOF
        If r_rsSchedOvertime!OT_OPP_NUM = v_lngOTOppNum And _
                r_rsSchedOvertime!Date = v_datDateWorked And _
                r_rsSchedOvertime!START_TIME = v_strStartTime Then
            r_rsSchedOvertime.Delete
            Exit Do
        End If
        r_rsSchedOvertime.MoveNext
    Loop
    r_rsSchedOvertime.Filter = adFilterNone

    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "DeleteOvertimeSchedule"
    
End Sub


'Private Sub DeleteAdjustedScheduleOvertime( _
'                ByVal v_rsTimeChargesReq As ADODB.Recordset, _
'                ByRef r_rsEmpSchedules As ADODB.Recordset)
'
'On Error GoTo eh
'
'    r_rsEmpSchedules.Filter = "EMP_NUM = " & v_rsTimeChargesReq!KEY_EMP_NUM
'
'    Do Until r_rsEmpSchedules.EOF
'        If IsScheduledOvertime(r_rsEmpSchedules) Then
'            If r_rsEmpSchedules!OT_OPP_NUM = v_rsTimeChargesReq!KEY_OT_OPP_NUM And _
'                    r_rsEmpSchedules!DATE_WORKED = v_rsTimeChargesReq!KEY_DATE_WORKED And _
'                    r_rsEmpSchedules!START_TIME = v_rsTimeChargesReq!KEY_START_TIME Then
'                r_rsEmpSchedules.Delete
'                Exit Do
'            End If
'        End If
'        r_rsEmpSchedules.MoveNext
'    Loop
'    r_rsEmpSchedules.Filter = adFilterNone
'
'    Exit Sub
'
'eh:
'    TrsRaiseError m_strCMODULENAME, "DeleteAdjustedScheduleOvertime"
'
'End Sub


Private Sub DeleteAdjustedScheduleOvertime( _
                ByVal v_lngEmpNum As Long, _
                ByVal v_lngOTOppNum As Long, _
                ByVal v_datDateWorked As Date, _
                ByVal v_strStartTime As String, _
                ByRef r_rsEmpSchedules As ADODB.Recordset)

On Error GoTo eh

'    void request does not need to get the adjusted schedule
    If r_rsEmpSchedules Is Nothing Then
        Exit Sub
    End If

    r_rsEmpSchedules.Filter = "EMP_NUM = " & v_lngEmpNum

    Do Until r_rsEmpSchedules.EOF
        If IsScheduledOvertime(r_rsEmpSchedules) Then
            If r_rsEmpSchedules!OT_OPP_NUM = v_lngOTOppNum And _
                    r_rsEmpSchedules!DATE_WORKED = v_datDateWorked And _
                    r_rsEmpSchedules!START_TIME = v_strStartTime Then
                r_rsEmpSchedules.Delete
                Exit Do
            End If
        End If
        r_rsEmpSchedules.MoveNext
    Loop
    r_rsEmpSchedules.Filter = adFilterNone

    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "DeleteAdjustedScheduleOvertime"
    
End Sub



Private Sub FormatCallStart(ByRef r_strCallStart As String)

    Dim l_strCallStart As String
    
On Error GoTo eh
    
    l_strCallStart = Trim(r_strCallStart)
    
    If Not IsDB2Timestamp(l_strCallStart) Then
        If IsDate(l_strCallStart) Then
            l_strCallStart = Format(l_strCallStart, "YYYY-MM-DD-HH.MM.SS") & ".000000"
        Else
            l_strCallStart = mINITCANVCALLTIMESTAMP
        End If
    End If

    r_strCallStart = l_strCallStart
    
    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "FormatCallStart"

End Sub


Private Function IsValidResponse( _
                    ByVal v_strResp As String _
                ) As Boolean
    
    Dim l_obj As dbCodeTables.cdCodeTables
    Dim l_blnError As Boolean
    
    On Error GoTo eh
    
    If Trim(v_strResp) = "" Then
        IsValidResponse = True
        Exit Function
    End If
    
    Set l_obj = CreateObject("dbCodeTables.cdCodeTables")
    l_blnError = l_obj.OvertimeResponseIsValid(v_strResp)
    Set l_obj = Nothing
    If Not l_blnError Then
        'Write BRV that overtime response does not exist
'        AppendResponseEntryError m_rsErrors, 30000, mstrCAPPLICATIONERR, "Response does not exist in the system"
    Else
        If RecordsExist(m_rsOpportunityResponses) Then
            m_rsOpportunityResponses.Filter = "OT_RESP_CD = '" & Trim(v_strResp) & "'"
'            If Not RecordsExist(m_rsOpportunityResponses) Then
'                l_blnError = False
'                'Write BRV that overtime response is not valid for the opportunity
'                AppendResponseEntryError v_rsResp, 30000, mstrCAPPLICATIONERR, "Response is not valid for this canvass"
'            End If
            m_rsOpportunityResponses.Filter = ""
        End If
    End If
    
    IsValidResponse = l_blnError
    
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "IsValidResponse"
        
End Function


Private Sub GetOpportunityResponses(ByVal v_lngOTOppNum As Long)
'Get attributes of responses on template

    Dim l_obj As dbOTCanvass.cdOTCanvass

    On Error GoTo eh
    
    Set l_obj = CreateObject("dbOTCanvass.cdOTCanvass")
    Set m_rsOpportunityResponses = l_obj.GetOpportunityResponses(v_lngOTOppNum)
    Set l_obj = Nothing
    
    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "GetOpportunityResponses"
End Sub


Private Function GetOvertimeResponses() As ADODB.Recordset

    Dim l_obj As dbOTCanvass.cdOTCanvass
    
    On Error GoTo eh

    Set l_obj = CreateObject("dbOTCanvass.cdOTCanvass")
    Set GetOvertimeResponses = l_obj.GetOvertimeResponses()
    Exit Function

eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "GetOvertimeResponses"
End Function

'
'Private Function ProcessTimeCharges( _
'                    ByVal v_rsOTOppTimeCharges As ADODB.Recordset _
'                ) As Boolean
'
'    Dim l_strRecStatus As String
'    Dim l_blnError As Boolean
'
'On Error GoTo eh
'
'    v_rsOTOppTimeCharges.MoveFirst
'    While Not v_rsOTOppTimeCharges.EOF
'        If PreValidateTimeCharge(v_rsOTOppTimeCharges) Then
'            AddChargeDate v_rsOTOppTimeCharges!CHARGE_DATE
'            l_strRecStatus = v_rsOTOppTimeCharges!RECORD_STATUS
'            Select Case l_strRecStatus
'                Case "I"
'                    AddTimeCharge v_rsOTOppTimeCharges
'                Case "D"
'                    m_rsTimeCharges.Delete
'                Case "C"
'                    ChangeTimeCharge v_rsOTOppTimeCharges
'            End Select
'        Else
'            'BRVs were written in called module
'            l_blnError = True
'        End If
'        v_rsOTOppTimeCharges.MoveNext
'    Wend
'
'    ProcessTimeCharges = Not l_blnError
'    Exit Function
'
'eh:
'    TrsRaiseError m_strCMODULENAME, "ProcessTimeCharges"
'
'End Function


Private Function ProcessRemainingTimeCharges( _
                    ByVal v_rsOTOppTimeCharges As ADODB.Recordset, _
                    ByVal v_rsResp As ADODB.Recordset) As Boolean
'Process time charges that have not been processed with responses.

    Dim l_colRemainEmps As Collection
    Dim l_varEmpNum As Variant
    Dim l_blnReturn As Boolean
    Dim l_strRespCd As String

On Error GoTo eh

    l_blnReturn = True
    If Not RecordsExist(v_rsOTOppTimeCharges) Then
        GoTo xit
    End If

    If RecordsExist(v_rsResp) Then
        Set l_colRemainEmps = AllRemainingTimeChargeEmployees(v_rsOTOppTimeCharges, v_rsResp)
    Else
        Set l_colRemainEmps = AllTimeChargeEmployees(v_rsOTOppTimeCharges)
    End If

    For Each l_varEmpNum In l_colRemainEmps
        l_blnReturn = ProcessTimeCharges(CLng(l_varEmpNum), v_rsOTOppTimeCharges)
        If Not l_blnReturn Then
            Exit For
        End If
        l_strRespCd = GetResponseCode(CLng(l_varEmpNum))
        If IsAcceptorResponse(l_strRespCd) Then
            If OvertimeRateCalculated(CLng(l_varEmpNum), m_rsOTOpportunity!CHARGE_DATE) Then
                ProcessMissingComponents CLng(l_varEmpNum), False
            End If
        End If
    Next


xit:
    ProcessRemainingTimeCharges = l_blnReturn
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "ProcessRemainingTimeCharges"

End Function


Private Function GetCallTime( _
                    ByVal v_lngEmpNum As Long) As String
' Answer 26 char db2 timestamp

On Error GoTo eh
    
    If FindCanvassResponse(v_lngEmpNum, m_rsCanvassSheet) Then
        If Trim(m_rsCanvassSheet!FORCE_RESP_CD) <> "" Then
            GetCallTime = Trim(m_rsCanvassSheet!FORCE_CALL_TIME)
        Else
            GetCallTime = Trim(m_rsCanvassSheet!CANVASS_CALL_TIME)
        End If
    End If

    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "GetCallTime"

End Function


Private Function GetResponseCode( _
                    ByVal v_lngEmpNum As Long) As String

On Error GoTo eh
    
    If FindCanvassResponse(v_lngEmpNum, m_rsCanvassSheet) Then
        GetResponseCode = _
            MostRecentResponseCode(m_rsCanvassSheet!CANVASS_RESP_CD, m_rsCanvassSheet!FORCE_RESP_CD)
    End If

    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "GetResponseCode"

End Function


Private Function AllRemainingTimeChargeEmployees( _
                    ByVal v_rsOTOppTimeCharges As ADODB.Recordset, _
                    ByVal v_rsResp As ADODB.Recordset) As Collection

    Dim l_colRemainEmps As utilTRISCommonLib.cuSet
    Dim l_lngEmpNum As Long
    
On Error GoTo eh

    Set l_colRemainEmps = New cuSet
    
    v_rsOTOppTimeCharges.MoveFirst
    Do Until v_rsOTOppTimeCharges.EOF
        If l_lngEmpNum <> v_rsOTOppTimeCharges!KEY_EMP_NUM Then
            l_lngEmpNum = v_rsOTOppTimeCharges!KEY_EMP_NUM
            If Not FindCanvassResponse(l_lngEmpNum, v_rsResp) Then
                l_colRemainEmps.Add l_lngEmpNum
            End If
            
        End If
        v_rsOTOppTimeCharges.MoveNext
    Loop

    v_rsOTOppTimeCharges.MoveFirst
    Set AllRemainingTimeChargeEmployees = l_colRemainEmps.Contents
    Set l_colRemainEmps = Nothing
    Exit Function

eh:
    Set l_colRemainEmps = Nothing
    TrsRaiseError m_strCMODULENAME, "AllRemainingTimeChargeEmployees"

End Function


Private Function AllTimeChargeEmployees(ByVal v_rsOTOppTimeCharges As ADODB.Recordset) As Collection

    Dim l_colRemainEmps As utilTRISCommonLib.cuSet
    
On Error GoTo eh

    Set l_colRemainEmps = New cuSet
    
    v_rsOTOppTimeCharges.MoveFirst
    Do Until v_rsOTOppTimeCharges.EOF
        l_colRemainEmps.Add v_rsOTOppTimeCharges.Fields("KEY_EMP_NUM").Value
        v_rsOTOppTimeCharges.MoveNext
    Loop

    v_rsOTOppTimeCharges.MoveFirst
    Set AllTimeChargeEmployees = l_colRemainEmps.Contents
    
    Set l_colRemainEmps = Nothing
    Exit Function

eh:
    Set l_colRemainEmps = Nothing
    TrsRaiseError m_strCMODULENAME, "AllTimeChargeEmployees"

End Function


Private Function ProcessTimeCharges( _
                    ByVal v_lngEmp_Num, _
                    ByVal v_rsOTOppTimeCharges As ADODB.Recordset _
                ) As Boolean
    
    Dim l_strRecStatus As String
    Dim l_blnError As Boolean
    
On Error GoTo eh
    
    If v_rsOTOppTimeCharges Is Nothing Then
        ProcessTimeCharges = True
        Exit Function
    End If
    
    v_rsOTOppTimeCharges.Filter = "KEY_EMP_NUM = " & v_lngEmp_Num & _
        " AND RECORD_STATUS = 'D'"

    Do Until v_rsOTOppTimeCharges.EOF
        AddChargeDate v_rsOTOppTimeCharges!KEY_CHARGE_DATE
        DeleteTimeCharge v_rsOTOppTimeCharges
        v_rsOTOppTimeCharges.MoveNext
    Loop

    v_rsOTOppTimeCharges.Filter = "KEY_EMP_NUM = " & v_lngEmp_Num & _
        " AND RECORD_STATUS <> 'D'"
    v_rsOTOppTimeCharges.Sort = "DATE_WORKED ASC, START_TIME ASC"

    Do Until v_rsOTOppTimeCharges.EOF
        AddChargeDate v_rsOTOppTimeCharges!KEY_CHARGE_DATE
        l_strRecStatus = v_rsOTOppTimeCharges!RECORD_STATUS
        Select Case l_strRecStatus
            Case "I"
                AddTimeCharge v_rsOTOppTimeCharges
            Case "C"
                ChangeTimeCharge v_rsOTOppTimeCharges
        End Select
        v_rsOTOppTimeCharges.MoveNext
    Loop
    
    v_rsOTOppTimeCharges.Filter = adFilterNone
    ProcessTimeCharges = Not l_blnError
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "ProcessTimeCharges"
    
End Function


Private Function PreValidateTimeCharge( _
                    ByVal v_rsTimeChargeReq As ADODB.Recordset _
                ) As Boolean
    
    Dim l_strRecStatus As String
    Dim l_blnError As Boolean
    
On Error GoTo eh
    
    If IsNull(v_rsTimeChargeReq!RECORD_STATUS) Or IsEmpty(v_rsTimeChargeReq!RECORD_STATUS) Then
        m_strErrorMessage = "Record Status is missing."
        AppendTimeChargeRequestBRV v_rsTimeChargeReq, trErrbusOTCanvass, mstrCAPPLICATIONERR, m_strErrorMessage
        PreValidateTimeCharge = False
        Exit Function
    Else
        l_strRecStatus = UCase(Trim(v_rsTimeChargeReq!RECORD_STATUS))
        If Not IsStringInList(l_strRecStatus, Array("D", "I", "C")) Then
            m_strErrorMessage = "Record Status is not valid, must be D, I, or C. Input=" & l_strRecStatus
            AppendTimeChargeRequestBRV v_rsTimeChargeReq, trErrbusOTCanvass, mstrCAPPLICATIONERR, m_strErrorMessage
            PreValidateTimeCharge = False
            Exit Function
        End If
    End If

    If v_rsTimeChargeReq!KEY_OT_OPP_NUM <> m_lngOTOppNum Then
        m_strErrorMessage = "KEY_OT_OPP_NUM is not consistent with parameter. Request input=" & CStr(v_rsTimeChargeReq!KEY_OT_OPP_NUM) & _
            ", Parameter input=" & m_lngOTOppNum
        AppendTimeChargeRequestBRV v_rsTimeChargeReq, trErrbusOTCanvass, mstrCAPPLICATIONERR, m_strErrorMessage
        PreValidateTimeCharge = False
        Exit Function
    End If

    If Not ValidTimeChargeKeyFields(v_rsTimeChargeReq) Then
        PreValidateTimeCharge = False
        Exit Function
    End If

    If IsEmpty(v_rsTimeChargeReq!CHARGE_DATE) Or IsNull(v_rsTimeChargeReq!CHARGE_DATE) Then
        l_blnError = True
        m_strErrorMessage = "CHARGE_DATE is missing."
        AppendTimeChargeRequestBRV v_rsTimeChargeReq, trErrbusOTCanvass, mstrCAPPLICATIONERR, m_strErrorMessage
    Else
        If Not IsDate(v_rsTimeChargeReq!CHARGE_DATE) Then
            l_blnError = True
            m_strErrorMessage = "CHARGE_DATE is not a valid date."
            AppendTimeChargeRequestBRV v_rsTimeChargeReq, trErrbusOTCanvass, mstrCAPPLICATIONERR, m_strErrorMessage
        End If
    End If
    
    If l_strRecStatus = "I" Or l_strRecStatus = "C" Then
        If Not IsTimeValid(v_rsTimeChargeReq!START_TIME) Then
            l_blnError = True
            m_strErrorMessage = "START_TIME is not valid."
            AppendTimeChargeRequestBRV v_rsTimeChargeReq, trErrbusOTCanvass, mstrCAPPLICATIONERR, m_strErrorMessage
        End If
        If Not IsTimeValid(v_rsTimeChargeReq!END_TIME) Then
            l_blnError = True
            m_strErrorMessage = "END_TIME is not valid.."
            AppendTimeChargeRequestBRV v_rsTimeChargeReq, trErrbusOTCanvass, mstrCAPPLICATIONERR, m_strErrorMessage
        End If
        If IsOverrideThatAllowsOverlap(v_rsTimeChargeReq!OVER_EXC_CD) And _
                Trim(v_rsTimeChargeReq!EXCEPTION_CD) = "" Then
            l_blnError = True
            m_strErrorMessage = "EXCEPTION_CD is required for CS time."
            AppendTimeChargeRequestBRV v_rsTimeChargeReq, trErrbusOTOppAppError, mstrCBUSINESSERR, m_strErrorMessage
        End If
    End If
    
    PreValidateTimeCharge = Not l_blnError
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "PreValidateTimeCharge"

End Function

Private Function ValidTimeChargeKeyFields( _
                    ByVal v_rsTimeChargeReq As ADODB.Recordset _
                ) As Boolean

    Dim l_blnValidationError As Boolean

On Error GoTo eh

    If Not IsNumeric(v_rsTimeChargeReq!KEY_OT_OPP_NUM) Or _
            v_rsTimeChargeReq!KEY_OT_OPP_NUM = 0 Then
        m_strErrorMessage = "KEY_OT_OPP_NUM is not valid. Input=" & CStr(v_rsTimeChargeReq!KEY_OT_OPP_NUM)
        AppendTimeChargeRequestBRV v_rsTimeChargeReq, trErrbusOTCanvass, mstrCAPPLICATIONERR, m_strErrorMessage
        l_blnValidationError = True
    End If
    
    If Not IsNumeric(v_rsTimeChargeReq!KEY_EMP_NUM) Or _
            v_rsTimeChargeReq!KEY_EMP_NUM = 0 Then
        m_strErrorMessage = "KEY_EMP_NUM is not valid. Input=" & CStr(v_rsTimeChargeReq!KEY_EMP_NUM)
        AppendTimeChargeRequestBRV v_rsTimeChargeReq, trErrbusOTCanvass, mstrCAPPLICATIONERR, m_strErrorMessage
        l_blnValidationError = True
    End If

    If UCase(Trim(v_rsTimeChargeReq!RECORD_STATUS)) <> "I" Then
        If Not IsDB2Timestamp(v_rsTimeChargeReq!CREATED_TIMESTAMP) Then
            m_strErrorMessage = "CREATED_TIMESTAMP is not valid. Input=" & CStr(v_rsTimeChargeReq!CREATED_TIMESTAMP)
            AppendTimeChargeRequestBRV v_rsTimeChargeReq, trErrbusOTCanvass, mstrCAPPLICATIONERR, m_strErrorMessage
            l_blnValidationError = True
        End If
        
        If Not IsDB2Timestamp(v_rsTimeChargeReq!LAST_MOD_TIMESTAMP) Then
            m_strErrorMessage = "LAST_MOD_TIMESTAMP is not valid. Input=" & CStr(v_rsTimeChargeReq!LAST_MOD_TIMESTAMP)
            AppendTimeChargeRequestBRV v_rsTimeChargeReq, trErrbusOTCanvass, mstrCAPPLICATIONERR, m_strErrorMessage
            l_blnValidationError = True
        End If
    End If
  
    ValidTimeChargeKeyFields = Not l_blnValidationError
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "ValidTimeChargeKeyFields"
End Function


Private Sub AddTimeCharge(ByVal v_rsTimeCharge As ADODB.Recordset)
    
    
On Error GoTo eh
    
    m_rsTimeCharges.AddNew
    m_rsTimeCharges!OT_OPP_NUM = v_rsTimeCharge!OT_OPP_NUM
    m_rsTimeCharges!EMP_NUM = v_rsTimeCharge!EMP_NUM
    m_rsTimeCharges!CHARGE_DATE = v_rsTimeCharge!CHARGE_DATE
    m_rsTimeCharges!DATE_WORKED = v_rsTimeCharge!DATE_WORKED
    m_rsTimeCharges!START_TIME = v_rsTimeCharge!START_TIME
    m_rsTimeCharges!END_TIME = v_rsTimeCharge!END_TIME
    If v_rsTimeCharge!HOURS = 0 Then
        m_rsTimeCharges!HOURS = DateDiff("n", _
            DateTimeFrom(v_rsTimeCharge!DATE_WORKED, v_rsTimeCharge!START_TIME), _
            DateTimeFrom(v_rsTimeCharge!DATE_WORKED, v_rsTimeCharge!END_TIME, True)) / 60
    Else
        m_rsTimeCharges!HOURS = v_rsTimeCharge!HOURS
    End If
    m_rsTimeCharges!EXCEPTION_CD = v_rsTimeCharge!EXCEPTION_CD
    m_rsTimeCharges!OVER_EXC_CD = v_rsTimeCharge!OVER_EXC_CD
    m_rsTimeCharges!JOB_NUM = v_rsTimeCharge!JOB_NUM
    m_rsTimeCharges!COMP_SHIFT_CD = v_rsTimeCharge!COMP_SHIFT_CD
    m_rsTimeCharges!OT_REASON_CD = v_rsTimeCharge!OT_REASON_CD
    m_rsTimeCharges!WG_NUM = v_rsTimeCharge!WG_NUM
    m_rsTimeCharges!DISTRIBUTION_ID = v_rsTimeCharge!DISTRIBUTION_ID
    m_rsTimeCharges!DISTRIBUTION_TYPE = v_rsTimeCharge!DISTRIBUTION_TYPE
    m_rsTimeCharges!JOB_NBR = v_rsTimeCharge!JOB_NBR
    m_rsTimeCharges!JOBTASK_NBR = v_rsTimeCharge!JOBTASK_NBR
    m_rsTimeCharges!JOBTASK_STEP_NBR = v_rsTimeCharge!JOBTASK_STEP_NBR
    m_rsTimeCharges!EMPRV_DATABASE_ID = v_rsTimeCharge!EMPRV_DATABASE_ID
    
'    Refusal charges must have exception code entered. If not, assume
'    worked and must be determined.
    If IsOvertimeWorked(v_rsTimeCharge!EXCEPTION_CD, v_rsTimeCharge!OVER_EXC_CD) Or _
            Trim(v_rsTimeCharge!EXCEPTION_CD) = "" Then
        If Trim(v_rsTimeCharge!TRAVEL_PAY_IND) = "Y" Then
            m_rsTimeCharges!TRAVEL_PAY_HRS = 0.7
        ElseIf Trim(v_rsTimeCharge!TRAVEL_PAY_IND) = "N" Then
            m_rsTimeCharges!TRAVEL_PAY_HRS = 0#
        Else
'            m_rsTimeCharges!TRAVEL_PAY_HRS = CDec(-0.1)
            MarkForSystemDetermination m_rsTimeCharges
        End If
    Else
        If Trim(v_rsTimeCharge!TRAVEL_PAY_IND) = "Y" Then
            m_rsTimeCharges!TRAVEL_PAY_HRS = 0.7
        Else
            m_rsTimeCharges!TRAVEL_PAY_HRS = 0#
        End If
    End If
    
    m_rsTimeCharges!CREATED_EMP_NUM = m_lngUserNum
    m_rsTimeCharges!LAST_MOD_EMP_NUM = m_lngUserNum
    
    If IsOvertimeWorked(v_rsTimeCharge!EXCEPTION_CD, v_rsTimeCharge!OVER_EXC_CD) Or _
            Trim(v_rsTimeCharge!EXCEPTION_CD) = "" Then
        AddOvertimeSchedule m_rsTimeCharges
        AddAdjustedScheduleOvertime m_rsTimeCharges
        If OvertimeRateCalculated(m_rsTimeCharges!EMP_NUM, m_rsTimeCharges!CHARGE_DATE) Then
            MarkAllTimeChargesForMissingComponents m_rsTimeCharges!EMP_NUM, m_rsTimeCharges
        End If
    End If
    
    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "AddTimeCharge"
End Sub


Private Sub ChangeTimeCharge(ByVal v_rsTimeChargeReq As ADODB.Recordset)
    
    Dim l_typCanvShift As CanvassedShift
    Dim l_blnOrigWorkedInd As Boolean
    Dim l_blnReqWorkedInd As Boolean
    Dim l_blnUserChangedTravel As Boolean
    Dim l_blnUserChangedExcCd As Boolean
    
On Error GoTo eh
    
    FindTimeCharge _
            v_rsTimeChargeReq!KEY_EMP_NUM, _
            v_rsTimeChargeReq!KEY_CHARGE_DATE, _
            v_rsTimeChargeReq!KEY_DATE_WORKED, _
            v_rsTimeChargeReq!KEY_START_TIME, _
            m_rsTimeCharges
            
    If m_rsTimeCharges.EOF Then
        m_strErrorMessage = "ChangeTimeCharge - FindTimeCharge failed to find timecharge:" & _
            " EMP_NUM = " & v_rsTimeChargeReq!KEY_EMP_NUM & _
            ", CHARGE_DATE = " & Format(v_rsTimeChargeReq!KEY_CHARGE_DATE, "mm/dd/yyyy") & _
            ", DATE_WORKED = " & Format(v_rsTimeChargeReq!KEY_DATE_WORKED, "mm/dd/yyyy") & _
            ", START_TIME = " & v_rsTimeChargeReq!KEY_START_TIME
        AppendTimeChargeRequestBRV v_rsTimeChargeReq, trErrbusOTCanvass, mstrCAPPLICATIONERR, m_strErrorMessage
        Err.Raise trErrbusOTOppAppError, m_strCMODULENAME & ":ChangeTimeCharge", m_strErrorMessage
    End If
    
    If Trim(v_rsTimeChargeReq!EXCEPTION_CD) = "" And _
                Not IsOverrideThatAllowsOverlap(v_rsTimeChargeReq!OVER_EXC_CD) Then
        l_blnReqWorkedInd = True
    Else
        l_blnReqWorkedInd = IsOvertimeWorked(v_rsTimeChargeReq!EXCEPTION_CD, v_rsTimeChargeReq!OVER_EXC_CD)
    End If
    l_blnOrigWorkedInd = IsOvertimeWorked( _
        m_rsTimeCharges.Fields("EXCEPTION_CD").OriginalValue, m_rsTimeCharges.Fields("OVER_EXC_CD").OriginalValue)
    
    l_blnUserChangedExcCd = _
        m_rsTimeCharges.Fields("EXCEPTION_CD").OriginalValue <> _
            v_rsTimeChargeReq.Fields("EXCEPTION_CD")
            
    l_blnUserChangedTravel = _
        m_rsTimeCharges.Fields("TRAVEL_PAY_HRS").OriginalValue <> _
                TravelPayHoursFromIndicator(Trim(v_rsTimeChargeReq!TRAVEL_PAY_IND))
    
    UpdateTimeCharge v_rsTimeChargeReq, m_rsTimeCharges
    
    If Not l_blnOrigWorkedInd And l_blnReqWorkedInd Then
        If Not l_blnUserChangedTravel Then
            MarkForSystemDetermination m_rsTimeCharges
'            m_rsTimeCharges!TRAVEL_PAY_HRS = CDec(-0.1)
        End If
        AddOvertimeSchedule m_rsTimeCharges
        AddAdjustedScheduleOvertime m_rsTimeCharges
    ElseIf l_blnOrigWorkedInd And Not l_blnReqWorkedInd Then
'    Should this MarkAllTimeChargesForMissingComponents
'    eg. if this time chg had travel, should system determine travel based on remaining worked?
' Gregs opinion 10/11/04 - do not move travel time, or mark other time charges for redetermination.
' Maybe an informational or warning message here if this time charge had travel.
        If m_rsTimeCharges.Fields("TRAVEL_PAY_HRS").OriginalValue > 0 Then
            If IsOverrideThatAllowsOverlap(m_rsTimeCharges!OVER_EXC_CD) Then
                m_strErrorMessage = "Travel time has been discarded because time charge has been changed to callout/standby(CS)."
            Else
                m_strErrorMessage = "Travel time has been discarded because time charge has been changed to not worked."
            End If
            AppendTimeChargeBRV m_rsTimeCharges, trWarOTEqualizationResult, mstrCBUSINESSWAR, m_strErrorMessage
        End If
        m_rsTimeCharges!TRAVEL_PAY_HRS = 0#
        DeleteOvertimeSchedule v_rsTimeChargeReq!KEY_EMP_NUM, v_rsTimeChargeReq!KEY_OT_OPP_NUM, _
            v_rsTimeChargeReq!KEY_DATE_WORKED, v_rsTimeChargeReq!KEY_START_TIME, m_rsSchedOvertime
        DeleteAdjustedScheduleOvertime v_rsTimeChargeReq!KEY_EMP_NUM, v_rsTimeChargeReq!KEY_OT_OPP_NUM, _
            v_rsTimeChargeReq!KEY_DATE_WORKED, v_rsTimeChargeReq!KEY_START_TIME, m_rsEmpSchedules
    ElseIf l_blnOrigWorkedInd And l_blnReqWorkedInd Then
        ChangeScheduledOvertime v_rsTimeChargeReq, m_rsTimeCharges
        If ChangeRequiresComponentDetermination(m_rsTimeCharges) Then
            If OvertimeRateCalculated(m_rsTimeCharges!EMP_NUM, m_rsTimeCharges!CHARGE_DATE) Then
                MarkAllTimeChargesForMissingComponents m_rsTimeCharges!EMP_NUM, m_rsTimeCharges
            End If
        ElseIf Trim(v_rsTimeChargeReq!EXCEPTION_CD) = "" And Not l_blnUserChangedTravel Then
            MarkForSystemDetermination m_rsTimeCharges
        End If
    Else
        ' both refusals and no schedule to change
    End If

xit:
    Exit Sub

eh:
    Dim l_colErrInfo As Collection
    
    If Err.Number = trErrbusOTOppAppError Then
        Set l_colErrInfo = New Collection
        l_colErrInfo.Add m_strErrorMessage
        l_colErrInfo.Add ", m_rsTimeCharges.Filter=" & m_rsTimeCharges.Filter
        l_colErrInfo.Add ", m_rsTimeCharges.RecordCount=" & m_rsTimeCharges.RecordCount
        TrsTraceError m_strCMODULENAME, "ChangeTimeCharge", 0, False, Nothing, l_colErrInfo, True
        Set l_colErrInfo = Nothing
        Err.Raise Err.Number, Err.Source, Err.Description
    Else
        TrsRaiseError m_strCMODULENAME, "ChangeTimeCharge"
    End If
    
End Sub


Private Function ChangeRequiresComponentDetermination( _
                    ByVal v_rsTimeCharge As ADODB.Recordset) As Boolean
'If some fields have changed, but other fields that are determined, eg
' EXCEPTION_CD, and TRAVEL_PAY_HRS, have not changed; assume the user wants
' the system to determine any possible update to other time charges for this
' overtime shift.

On Error GoTo eh

    If ((v_rsTimeCharge.Fields("START_TIME").OriginalValue <> _
            v_rsTimeCharge.Fields("START_TIME").Value) Or _
        (v_rsTimeCharge.Fields("END_TIME").OriginalValue <> _
            v_rsTimeCharge.Fields("END_TIME").Value)) And _
        ((v_rsTimeCharge.Fields("EXCEPTION_CD").OriginalValue = _
            v_rsTimeCharge.Fields("EXCEPTION_CD").Value) And _
        (v_rsTimeCharge.Fields("TRAVEL_PAY_HRS").OriginalValue = _
            v_rsTimeCharge.Fields("TRAVEL_PAY_HRS").Value)) Then
        ChangeRequiresComponentDetermination = True
    End If
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "ChangeRequiresComponentDetermination"
    
End Function



Private Sub DeleteTimeCharge(ByVal v_rsTimeChargeReq As ADODB.Recordset)
    
    
On Error GoTo eh
    
    FindTimeCharge _
            v_rsTimeChargeReq!KEY_EMP_NUM, _
            v_rsTimeChargeReq!KEY_CHARGE_DATE, _
            v_rsTimeChargeReq!KEY_DATE_WORKED, _
            v_rsTimeChargeReq!KEY_START_TIME, _
            m_rsTimeCharges
            
    If m_rsTimeCharges.EOF Then
        If Not IsSystemDeletedTimeCharge( _
                            v_rsTimeChargeReq!KEY_EMP_NUM, _
                            v_rsTimeChargeReq!KEY_CHARGE_DATE, _
                            v_rsTimeChargeReq!KEY_DATE_WORKED, _
                            v_rsTimeChargeReq!KEY_START_TIME) Then
            m_strErrorMessage = "DeleteTimeCharge - FindTimeCharge failed to find timecharge:" & _
                " EMP_NUM = " & v_rsTimeChargeReq!KEY_EMP_NUM & _
                ", CHARGE_DATE = " & Format(v_rsTimeChargeReq!KEY_CHARGE_DATE, "mm/dd/yyyy") & _
                ", DATE_WORKED = " & Format(v_rsTimeChargeReq!KEY_DATE_WORKED, "mm/dd/yyyy") & _
                ", START_TIME = " & v_rsTimeChargeReq!KEY_START_TIME
                AppendTimeChargeRequestBRV v_rsTimeChargeReq, trErrbusOTCanvass, mstrCAPPLICATIONERR, m_strErrorMessage
            Err.Raise trErrbusOTOppAppError, m_strCMODULENAME & ":DeleteTimeCharge", m_strErrorMessage
        Else
            m_strErrorMessage = "Delete timecharge request ignored.  System deleted timecharge during response change/delete. (" & _
                v_rsTimeChargeReq!EMP_NUM & ", " & v_rsTimeChargeReq!KEY_DATE_WORKED & _
                ", " & v_rsTimeChargeReq!KEY_START_TIME & ")"
            AppendBusinessError m_rsErrors, trInfOTOppInfo, mstrCBUSINESSINF, _
                m_strErrorMessage, m_strCMODULENAME
            GoTo xit
        End If
    End If

    If IsOvertimeWorked(m_rsTimeCharges!EXCEPTION_CD, m_rsTimeCharges!OVER_EXC_CD) Or _
                Trim(m_rsTimeCharges!EXCEPTION_CD) = "" Then
        DeleteOvertimeSchedule v_rsTimeChargeReq!KEY_EMP_NUM, v_rsTimeChargeReq!KEY_OT_OPP_NUM, _
            v_rsTimeChargeReq!KEY_DATE_WORKED, v_rsTimeChargeReq!KEY_START_TIME, m_rsSchedOvertime
        DeleteAdjustedScheduleOvertime v_rsTimeChargeReq!KEY_EMP_NUM, v_rsTimeChargeReq!KEY_OT_OPP_NUM, _
            v_rsTimeChargeReq!KEY_DATE_WORKED, v_rsTimeChargeReq!KEY_START_TIME, m_rsEmpSchedules
        If OvertimeRateCalculated(m_rsTimeCharges!EMP_NUM, m_rsTimeCharges!CHARGE_DATE) Then
            MarkAllTimeChargesForMissingComponents m_rsTimeCharges!EMP_NUM, m_rsTimeCharges
        End If
    End If
        
    m_rsTimeCharges.Delete

xit:
    Exit Sub

eh:

    Dim l_colErrInfo As Collection

    If Err.Number = trErrbusOTOppAppError Then
        Set l_colErrInfo = New Collection
        l_colErrInfo.Add m_strErrorMessage
        l_colErrInfo.Add ", m_rsTimeCharges.Filter=" & m_rsTimeCharges.Filter
        l_colErrInfo.Add ", m_rsTimeCharges.RecordCount=" & m_rsTimeCharges.RecordCount
        TrsTraceError m_strCMODULENAME, "DeleteTimeCharge", 0, False, Nothing, l_colErrInfo, True
        Set l_colErrInfo = Nothing
        TrsRaiseError m_strCMODULENAME, "DeleteTimeCharge"
    Else
        TrsRaiseError m_strCMODULENAME, "DeleteTimeCharge"
    End If

End Sub


Private Sub FindTimeCharge( _
                ByVal v_lngEmpNum As Long, _
                ByVal v_datChargeDate As Date, _
                ByVal v_datDateWorked As Date, _
                ByVal v_strStartTime As String, _
                ByRef r_rsTimeCharge As ADODB.Recordset)

'Should CREATED_TIMESTAMP be included?

On Error GoTo eh

    If Not RecordsExist(r_rsTimeCharge) Then
        Exit Sub
    End If

    r_rsTimeCharge.MoveFirst
    Do Until r_rsTimeCharge.EOF
        If r_rsTimeCharge!EMP_NUM = v_lngEmpNum And _
            r_rsTimeCharge!CHARGE_DATE = v_datChargeDate And _
            r_rsTimeCharge!DATE_WORKED = v_datDateWorked And _
            r_rsTimeCharge!START_TIME = v_strStartTime Then
            Exit Do
        Else
            r_rsTimeCharge.MoveNext
        End If
    Loop

    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "FindTimeCharge"
                
End Sub


Private Sub UpdateSchedulesWithMissingComponents(ByVal v_rsTimeCharge As ADODB.Recordset)

    Dim l_intFind As Integer

On Error GoTo eh
    
    FindScheduledOvertime v_rsTimeCharge!OT_OPP_NUM, v_rsTimeCharge!EMP_NUM, _
        v_rsTimeCharge!DATE_WORKED, v_rsTimeCharge!START_TIME, m_rsSchedOvertime
        
    If m_rsSchedOvertime.EOF Then
        l_intFind = 1
        m_strErrorMessage = "UpdateSchedulesWithMissingComponents - FindScheduledOvertime failed to find schedule:" & _
            " EMP_NUM = " & v_rsTimeCharge!EMP_NUM & _
            ", DATE_WORKED = " & Format(v_rsTimeCharge!DATE_WORKED, "mm/dd/yyyy") & _
            ", START_TIME = " & v_rsTimeCharge!START_TIME
        AppendTimeChargeBRV v_rsTimeCharge, trErrbusOTCanvass, mstrCAPPLICATIONERR, m_strErrorMessage
        Err.Raise trErrbusOTOppAppError, m_strCMODULENAME & ":UpdateSchedulesWithMissingComponents", m_strErrorMessage
    Else
        UpdateScheduledOvertimePay v_rsTimeCharge, m_rsSchedOvertime
    End If
    
    FindAdjustedScheduleOvertime v_rsTimeCharge!OT_OPP_NUM, v_rsTimeCharge!EMP_NUM, _
        v_rsTimeCharge!DATE_WORKED, v_rsTimeCharge!START_TIME, m_rsEmpSchedules
        
    If m_rsEmpSchedules.EOF Then
        l_intFind = 2
        m_strErrorMessage = "UpdateSchedulesWithMissingComponents - FindAdjustedScheduleOvertime failed to find schedule:" & _
            " EMP_NUM = " & v_rsTimeCharge!EMP_NUM & _
            ", DATE_WORKED = " & Format(v_rsTimeCharge!DATE_WORKED, "mm/dd/yyyy") & _
            ", START_TIME = " & v_rsTimeCharge!START_TIME
        AppendTimeChargeBRV v_rsTimeCharge, trErrbusOTCanvass, mstrCAPPLICATIONERR, m_strErrorMessage
        Err.Raise trErrbusOTOppAppError, m_strCMODULENAME & ":UpdateSchedulesWithMissingComponents", m_strErrorMessage
    Else
        UpdateAdjustedScheduleOvertimePay v_rsTimeCharge, m_rsEmpSchedules
    End If
    
    
xit:

    Exit Sub

eh:
    Dim l_colErrInfo As Collection

    If Err.Number = trErrbusOTOppAppError Then
        Set l_colErrInfo = New Collection
        l_colErrInfo.Add m_strErrorMessage
        If l_intFind = 1 Then
            l_colErrInfo.Add ", m_rsSchedOvertime.Filter=" & m_rsSchedOvertime.Filter
            l_colErrInfo.Add ", m_rsSchedOvertime.RecordCount=" & m_rsSchedOvertime.RecordCount
        Else
            l_colErrInfo.Add ", m_rsEmpSchedules.Filter=" & m_rsEmpSchedules.Filter
            l_colErrInfo.Add ", m_rsEmpSchedules.RecordCount=" & m_rsEmpSchedules.RecordCount
        End If
        TrsTraceError m_strCMODULENAME, "UpdateSchedulesWithMissingComponents", 0, False, Nothing, l_colErrInfo, True
        Set l_colErrInfo = Nothing
        TrsRaiseError m_strCMODULENAME, "UpdateSchedulesWithMissingComponents"
    Else
        TrsRaiseError m_strCMODULENAME, "UpdateSchedulesWithMissingComponents"
    End If

End Sub


Private Sub UpdateSchedulesAfterSplit(ByVal v_rsTimeCharge As ADODB.Recordset)

    Dim l_intFind As Integer

On Error GoTo eh
    
    FindScheduledOvertime v_rsTimeCharge!OT_OPP_NUM, v_rsTimeCharge!EMP_NUM, _
        v_rsTimeCharge!DATE_WORKED, v_rsTimeCharge!START_TIME, m_rsSchedOvertime
        
    If m_rsSchedOvertime.EOF Then
        l_intFind = 1
        m_strErrorMessage = "UpdateSchedulesAfterSplit - FindScheduledOvertime failed to find schedule:" & _
            " EMP_NUM = " & v_rsTimeCharge!EMP_NUM & _
            ", DATE_WORKED = " & Format(v_rsTimeCharge!DATE_WORKED, "mm/dd/yyyy") & _
            ", START_TIME = " & v_rsTimeCharge!START_TIME
        AppendTimeChargeBRV v_rsTimeCharge, trErrbusOTCanvass, mstrCAPPLICATIONERR, m_strErrorMessage
        Err.Raise trErrbusOTOppAppError, m_strCMODULENAME & ":UpdateSchedulesAfterSplit", m_strErrorMessage
    Else
        m_rsSchedOvertime!END_TIME = v_rsTimeCharge!END_TIME
    End If
    
    FindAdjustedScheduleOvertime v_rsTimeCharge!OT_OPP_NUM, v_rsTimeCharge!EMP_NUM, _
        v_rsTimeCharge!DATE_WORKED, v_rsTimeCharge!START_TIME, m_rsEmpSchedules
        
    If m_rsEmpSchedules.EOF Then
        l_intFind = 2
        m_strErrorMessage = "UpdateSchedulesAfterSplit - FindAdjustedScheduleOvertime failed to find schedule:" & _
            " EMP_NUM = " & v_rsTimeCharge!EMP_NUM & _
            ", DATE_WORKED = " & Format(v_rsTimeCharge!DATE_WORKED, "mm/dd/yyyy") & _
            ", START_TIME = " & v_rsTimeCharge!START_TIME
        AppendTimeChargeBRV v_rsTimeCharge, trErrbusOTCanvass, mstrCAPPLICATIONERR, m_strErrorMessage
        Err.Raise trErrbusOTOppAppError, m_strCMODULENAME & ":UpdateSchedulesAfterSplit", m_strErrorMessage
    Else
        m_rsEmpSchedules!END_TIME = v_rsTimeCharge!END_TIME
        m_rsEmpSchedules!HOURS = v_rsTimeCharge!HOURS
    End If
    
    
xit:

    Exit Sub

eh:
    Dim l_colErrInfo As Collection

    If Err.Number = trErrbusOTOppAppError Then
        Set l_colErrInfo = New Collection
        l_colErrInfo.Add m_strErrorMessage
        If l_intFind = 1 Then
            l_colErrInfo.Add ", m_rsSchedOvertime.Filter=" & m_rsSchedOvertime.Filter
            l_colErrInfo.Add ", m_rsSchedOvertime.RecordCount=" & m_rsSchedOvertime.RecordCount
        Else
            l_colErrInfo.Add ", m_rsEmpSchedules.Filter=" & m_rsEmpSchedules.Filter
            l_colErrInfo.Add ", m_rsEmpSchedules.RecordCount=" & m_rsEmpSchedules.RecordCount
        End If
        TrsTraceError m_strCMODULENAME, "UpdateSchedulesAfterSplit", 0, False, Nothing, l_colErrInfo, True
        Set l_colErrInfo = Nothing
        TrsRaiseError m_strCMODULENAME, "UpdateSchedulesAfterSplit"
    Else
        TrsRaiseError m_strCMODULENAME, "UpdateSchedulesAfterSplit"
    End If

End Sub


Private Sub FindScheduledOvertime( _
                ByVal v_lngOTOppNum As Long, _
                ByVal v_lngEmpNum As Long, _
                ByVal v_datDateWorked As Date, _
                ByVal v_strStartTime As String, _
                ByRef r_rsScheduledOvertime As ADODB.Recordset)

'Should CREATED_TIMESTAMP be included?

On Error GoTo eh
                
    r_rsScheduledOvertime.MoveFirst
    Do Until r_rsScheduledOvertime.EOF
        If r_rsScheduledOvertime!EMP_NUM = v_lngEmpNum And _
            r_rsScheduledOvertime!Date = v_datDateWorked And _
            r_rsScheduledOvertime!START_TIME = v_strStartTime And _
            r_rsScheduledOvertime!OT_OPP_NUM = v_lngOTOppNum Then
            Exit Do
        Else
            r_rsScheduledOvertime.MoveNext
        End If
    Loop

    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "FindScheduledOvertime"
                
End Sub


Private Sub FindAdjustedScheduleOvertime( _
                ByVal v_lngOTOppNum As Long, _
                ByVal v_lngEmpNum As Long, _
                ByVal v_datDateWorked As Date, _
                ByVal v_strStartTime As String, _
                ByRef r_rsEmpSchedules As ADODB.Recordset)

On Error GoTo eh
                
    r_rsEmpSchedules.MoveFirst
    Do Until r_rsEmpSchedules.EOF
        If r_rsEmpSchedules!EMP_NUM = v_lngEmpNum And _
                r_rsEmpSchedules!DATE_WORKED = v_datDateWorked And _
                r_rsEmpSchedules!START_TIME = v_strStartTime And _
                r_rsEmpSchedules!OT_OPP_NUM = v_lngOTOppNum Then
            Exit Do
        Else
            r_rsEmpSchedules.MoveNext
        End If
    Loop

    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "FindAdjustedScheduleOvertime"
                
End Sub


Private Function FindCanvassResponse( _
                ByVal v_lngEmpNum As Long, _
                ByRef r_rsCanvassSheet As ADODB.Recordset) As Boolean

'Should CREATED_TIMESTAMP be included?

On Error GoTo eh
                
    FindCanvassResponse = False
    If r_rsCanvassSheet Is Nothing Then
        GoTo xit
    End If
    
    r_rsCanvassSheet.MoveFirst
    Do Until r_rsCanvassSheet.EOF
        If r_rsCanvassSheet!EMP_NUM = v_lngEmpNum Then
            FindCanvassResponse = True
            GoTo xit
        Else
            r_rsCanvassSheet.MoveNext
        End If
    Loop

xit:
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "FindCanvassResponse"
                
End Function


Private Sub UpdateTimeCharge( _
                ByVal v_rsTimeChargeReq As ADODB.Recordset, _
                ByRef r_rsTimeCharge As ADODB.Recordset)
    
On Error GoTo eh
    
    r_rsTimeCharge!CHARGE_DATE = v_rsTimeChargeReq!CHARGE_DATE
    r_rsTimeCharge!DATE_WORKED = v_rsTimeChargeReq!DATE_WORKED
    r_rsTimeCharge!START_TIME = v_rsTimeChargeReq!START_TIME
    r_rsTimeCharge!END_TIME = v_rsTimeChargeReq!END_TIME
'    r_rsTimeCharge!HOURS = v_rsTimeChargeReq!HOURS
    If v_rsTimeChargeReq!HOURS = 0 Then
        r_rsTimeCharge!HOURS = DateDiff("n", _
            DateTimeFrom(v_rsTimeChargeReq!DATE_WORKED, v_rsTimeChargeReq!START_TIME), _
            DateTimeFrom(v_rsTimeChargeReq!DATE_WORKED, v_rsTimeChargeReq!END_TIME, True)) / 60
    Else
        r_rsTimeCharge!HOURS = v_rsTimeChargeReq!HOURS
    End If

    r_rsTimeCharge!OVER_EXC_CD = v_rsTimeChargeReq!OVER_EXC_CD
    r_rsTimeCharge!JOB_NUM = v_rsTimeChargeReq!JOB_NUM
    r_rsTimeCharge!COMP_SHIFT_CD = v_rsTimeChargeReq!COMP_SHIFT_CD
    r_rsTimeCharge!OT_REASON_CD = v_rsTimeChargeReq!OT_REASON_CD
    r_rsTimeCharge!WG_NUM = v_rsTimeChargeReq!WG_NUM
    r_rsTimeCharge!DISTRIBUTION_ID = v_rsTimeChargeReq!DISTRIBUTION_ID
    r_rsTimeCharge!DISTRIBUTION_TYPE = v_rsTimeChargeReq!DISTRIBUTION_TYPE
    r_rsTimeCharge!JOB_NBR = v_rsTimeChargeReq!JOB_NBR
    r_rsTimeCharge!JOBTASK_NBR = v_rsTimeChargeReq!JOBTASK_NBR
    r_rsTimeCharge!JOBTASK_STEP_NBR = v_rsTimeChargeReq!JOBTASK_STEP_NBR
    r_rsTimeCharge!EMPRV_DATABASE_ID = v_rsTimeChargeReq!EMPRV_DATABASE_ID

    If SystemMarkedForDetermination(r_rsTimeCharge) Then
'        Override system determination when user changed value
        If r_rsTimeCharge.Fields("EXCEPTION_CD").OriginalValue <> _
                v_rsTimeChargeReq!EXCEPTION_CD Then
            r_rsTimeCharge!EXCEPTION_CD = v_rsTimeChargeReq!EXCEPTION_CD
        End If
    Else
        r_rsTimeCharge!EXCEPTION_CD = v_rsTimeChargeReq!EXCEPTION_CD
    End If

    If IsOvertimeRefused(r_rsTimeCharge!EXCEPTION_CD) Then
        If r_rsTimeCharge.Fields("TRAVEL_PAY_HRS").OriginalValue <> _
                    TravelPayHoursFromIndicator(Trim(v_rsTimeChargeReq!TRAVEL_PAY_IND)) Then
            If Trim(v_rsTimeChargeReq!TRAVEL_PAY_IND) = "Y" Then
                r_rsTimeCharge!TRAVEL_PAY_HRS = 0.7
            Else
                r_rsTimeCharge!TRAVEL_PAY_HRS = 0#
            End If
        End If
    Else
        If SystemMarkedForDetermination(r_rsTimeCharge) Then
    '        Override system determination when user changed value
            If r_rsTimeCharge.Fields("TRAVEL_PAY_HRS").OriginalValue <> _
                    TravelPayHoursFromIndicator(Trim(v_rsTimeChargeReq!TRAVEL_PAY_IND)) Then
                If Trim(v_rsTimeChargeReq!TRAVEL_PAY_IND) = "Y" Then
                    r_rsTimeCharge!TRAVEL_PAY_HRS = 0.7
                ElseIf Trim(v_rsTimeChargeReq!TRAVEL_PAY_IND) = "N" Then
                    r_rsTimeCharge!TRAVEL_PAY_HRS = 0#
                Else
                    r_rsTimeCharge!TRAVEL_PAY_HRS = 0#
                'should not be necessary since already this value based on SystemMarkedForDetermination
'                    r_rsTimeCharge!TRAVEL_PAY_HRS = CDec(-0.1)      'system
                End If
            End If
        Else
            If Trim(v_rsTimeChargeReq!TRAVEL_PAY_IND) = "Y" Then
                r_rsTimeCharge!TRAVEL_PAY_HRS = 0.7
            ElseIf Trim(v_rsTimeChargeReq!TRAVEL_PAY_IND) = "N" Then
                r_rsTimeCharge!TRAVEL_PAY_HRS = 0#
            Else
                MarkForSystemDetermination r_rsTimeCharge
'                r_rsTimeCharge!TRAVEL_PAY_HRS = CDec(-0.1)      'system
            End If
        End If
    End If

    r_rsTimeCharge!LAST_MOD_EMP_NUM = m_lngUserNum
    
    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "UpdateTimeCharge"
End Sub


Private Function SystemMarkedForDetermination( _
                    ByVal v_rsTimeCharge As ADODB.Recordset) As Boolean
' This should never be true for red time charge.

    SystemMarkedForDetermination = (v_rsTimeCharge!TRAVEL_PAY_HRS = CDec(-0.1))

End Function

Private Sub MarkForSystemDetermination( _
                    ByRef r_rsTimeCharge As ADODB.Recordset)

    If OvertimeRateCalculated(r_rsTimeCharge!EMP_NUM, r_rsTimeCharge!CHARGE_DATE) Then
        r_rsTimeCharge!TRAVEL_PAY_HRS = CDec(-0.1)
    End If

End Sub


Private Function TravelPayHoursFromIndicator(ByVal v_strTravPayInd) As Variant

    If v_strTravPayInd = "Y" Then
        TravelPayHoursFromIndicator = CDec(0.7)
    ElseIf v_strTravPayInd = "N" Then
        TravelPayHoursFromIndicator = CDec(0#)
    Else
        TravelPayHoursFromIndicator = CDec(-0.1)
    End If

End Function


Private Sub ChangeScheduledOvertime( _
                ByVal v_rsTimeChargeReq As ADODB.Recordset, _
                ByRef r_rsTimeCharge As ADODB.Recordset)
'Position on scheduled overtime with the request, then update based
'on the updated time charge.


On Error GoTo eh

    FindScheduledOvertime _
                v_rsTimeChargeReq!KEY_OT_OPP_NUM, _
                v_rsTimeChargeReq!KEY_EMP_NUM, _
                v_rsTimeChargeReq!KEY_DATE_WORKED, _
                v_rsTimeChargeReq!KEY_START_TIME, _
                m_rsSchedOvertime
    If m_rsSchedOvertime.EOF Then
        m_strErrorMessage = "ChangeScheduledOvertime - FindScheduledOvertime failed to find schedule:" & _
            " EMP_NUM = " & v_rsTimeChargeReq!KEY_EMP_NUM & _
            ", DATE_WORKED = " & Format(v_rsTimeChargeReq!KEY_DATE_WORKED, "mm/dd/yyyy") & _
            ", START_TIME = " & v_rsTimeChargeReq!KEY_START_TIME
        AppendTimeChargeRequestBRV v_rsTimeChargeReq, trErrbusOTCanvass, mstrCAPPLICATIONERR, m_strErrorMessage
        Err.Raise trErrbusOTOppAppError, m_strCMODULENAME & ":UpdateSchedulesAfterSplit", m_strErrorMessage
    End If
    
    UpdateScheduledOvertime r_rsTimeCharge, m_rsSchedOvertime
    
    ChangeAdjustedScheduleOvertime v_rsTimeChargeReq, r_rsTimeCharge
    
xit:
    Exit Sub

eh:

    Dim l_colErrInfo As Collection

    If Err.Number = trErrbusOTOppAppError Then
        Set l_colErrInfo = New Collection
        l_colErrInfo.Add m_strErrorMessage
        l_colErrInfo.Add ", m_rsSchedOvertime.Filter=" & m_rsSchedOvertime.Filter
        l_colErrInfo.Add ", m_rsSchedOvertime.RecordCount=" & m_rsSchedOvertime.RecordCount
        TrsTraceError m_strCMODULENAME, "ChangeScheduledOvertime", 0, False, Nothing, l_colErrInfo, True
        Set l_colErrInfo = Nothing
        TrsRaiseError m_strCMODULENAME, "ChangeScheduledOvertime"
    Else
        TrsRaiseError m_strCMODULENAME, "ChangeScheduledOvertime"
    End If

End Sub


Private Sub UpdateScheduledOvertime( _
                ByVal v_rsTimeChargeUpd As ADODB.Recordset, _
                ByRef r_rsSchedOvertime As ADODB.Recordset)
    
On Error GoTo eh
    
    r_rsSchedOvertime!Date = v_rsTimeChargeUpd!DATE_WORKED
    r_rsSchedOvertime!CHARGE_DATE = v_rsTimeChargeUpd!CHARGE_DATE
    r_rsSchedOvertime!START_TIME = v_rsTimeChargeUpd!START_TIME
    r_rsSchedOvertime!END_TIME = v_rsTimeChargeUpd!END_TIME
    r_rsSchedOvertime!EXCEPTION_CD = v_rsTimeChargeUpd!EXCEPTION_CD
    r_rsSchedOvertime!OVER_EXC_CD = v_rsTimeChargeUpd!OVER_EXC_CD
    r_rsSchedOvertime!JOB_NUM = v_rsTimeChargeUpd!JOB_NUM
    r_rsSchedOvertime!SHIFT_COMP = v_rsTimeChargeUpd!COMP_SHIFT_CD
    r_rsSchedOvertime!WG_NUM = v_rsTimeChargeUpd!WG_NUM
    r_rsSchedOvertime!TRAV_PAY_HRS = v_rsTimeChargeUpd!TRAVEL_PAY_HRS
    r_rsSchedOvertime!LAST_MOD_EMP_NUM = m_lngUserNum

    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "UpdateScheduledOvertime"
End Sub


Private Sub UpdateScheduledOvertimePay( _
                ByVal v_rsTimeCharge As ADODB.Recordset, _
                ByRef r_rsSchedOvertime As ADODB.Recordset)
' Update attributes that may have been determined.

On Error GoTo eh
    
    r_rsSchedOvertime!EXCEPTION_CD = v_rsTimeCharge!EXCEPTION_CD
    r_rsSchedOvertime!OVER_EXC_CD = v_rsTimeCharge!OVER_EXC_CD
    r_rsSchedOvertime!TRAV_PAY_HRS = v_rsTimeCharge!TRAVEL_PAY_HRS

    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "UpdateScheduledOvertimePay"
End Sub


Private Sub UpdateAdjustedScheduleOvertimePay( _
                ByVal v_rsTimeCharge As ADODB.Recordset, _
                ByRef r_rsEmpSchedules As ADODB.Recordset)
' Update attributes that may have been determined.
    
On Error GoTo eh

    r_rsEmpSchedules!EXCEPTION_CD = v_rsTimeCharge!EXCEPTION_CD
    r_rsEmpSchedules!OVER_EXC_CD = v_rsTimeCharge!OVER_EXC_CD
    r_rsEmpSchedules!TRAV_PAY_HOURS = v_rsTimeCharge!TRAVEL_PAY_HRS

    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "UpdateAdjustedScheduleOvertimePay"
End Sub


Private Sub ChangeAdjustedScheduleOvertime( _
                ByVal v_rsTimeChargeReq As ADODB.Recordset, _
                ByRef r_rsTimeCharge As ADODB.Recordset)

On Error GoTo eh

'    DeleteAdjustedScheduleOvertime v_rsTimeChargeReq, m_rsEmpSchedules
    DeleteAdjustedScheduleOvertime v_rsTimeChargeReq!KEY_EMP_NUM, v_rsTimeChargeReq!KEY_OT_OPP_NUM, _
        v_rsTimeChargeReq!KEY_DATE_WORKED, v_rsTimeChargeReq!KEY_START_TIME, m_rsEmpSchedules

    AddAdjustedScheduleOvertime r_rsTimeCharge
    
xit:
    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "ChangeAdjustedScheduleOvertime"
End Sub



'Private Function SaveOppResponses() As Boolean
'
'    Dim l_obj As dbOTCanvass.cdOTCanvass
'
'On Error GoTo eh
'    m_rsResponses.Filter = ""
'    Set l_obj = CreateObject("dbOTCanvass.cdOTCanvass")
'    SaveOppResponses = l_obj.SaveOppResponses(m_rsResponses)
'    Set l_obj = Nothing
'    Exit Function
'
'eh:
'    TrsRaiseError m_strCMODULENAME, "SaveOppResponses"
'
'End Function


Private Function SaveOppTimeCharges() As Boolean

    Dim l_obj As dbOTCanvass.cdOTCanvass

    Dim stTime As Long
    Dim ndTime As Long
    Dim dfTime As Long

On Error GoTo eh

    #If TRTRACE Then
        stTime = GetTickCount()
    #End If

    m_rsTimeCharges.Filter = adFilterNone
    Set l_obj = CreateObject("dbOTCanvass.cdOTCanvass")
    SaveOppTimeCharges = l_obj.SaveOppTimeCharges(m_rsTimeCharges)
    Set l_obj = Nothing
    
    #If TRTRACE Then
        ndTime = GetTickCount()
        dfTime = ndTime - stTime
        TraceAction App.ThreadID, "3", "SaveOppTimeCharges", Format(Now, "mm/dd/yyyy hh:mm:ss"), _
            CStr(dfTime / 1000)
    #End If
    
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "SaveOppTimeCharges"

End Function



Private Function SaveScheduledOvertime() As Boolean

    Dim l_obj As dbOTCanvass.cdOTCanvass

On Error GoTo eh

    m_rsSchedOvertime.Filter = adFilterNone
    Set l_obj = CreateObject("dbOTCanvass.cdOTCanvass")
    SaveScheduledOvertime = l_obj.SaveOppSchedule(m_rsSchedOvertime)
    Set l_obj = Nothing
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "SaveScheduledOvertime"

End Function


Private Sub BuildEMPRVInterface()
' For each employee that accepted overtime, add a record for each timecharge row.

    Dim l_strRecentRespCd As String

On Error GoTo eh

    Set m_rsEMPRVOTSchedule = NewEMPRVOTScheduleSet

    m_rsCanvassSheet.Filter = adFilterNone
    m_rsCanvassSheet.MoveFirst
    
    Do Until m_rsCanvassSheet.EOF
        l_strRecentRespCd = MostRecentResponseCode(m_rsCanvassSheet!CANVASS_RESP_CD, m_rsCanvassSheet!FORCE_RESP_CD)
        If l_strRecentRespCd <> "" Then
            If IsEmployeeEMPRV(m_rsCanvassSheet!EMP_NUM, m_rsOTOpportunity!CHARGE_DATE) Then
                If IsOvertimeWorkedResponse(l_strRecentRespCd) Then
                    m_rsTimeCharges.Filter = "EMP_NUM = " & m_rsCanvassSheet!EMP_NUM
                    Do Until m_rsTimeCharges.EOF
                        If IsOvertimeWorked(m_rsTimeCharges!EXCEPTION_CD, "") Then
                            AddEMPRVCanvassInsert m_rsTimeCharges, m_rsEMPRVOTSchedule
                        End If
                        m_rsTimeCharges.MoveNext
                    Loop
                End If
            End If
        End If
        m_rsCanvassSheet.MoveNext
    Loop

    If m_rsEMPRVOTSchedule.RecordCount > 0 Then
        AddEMPRVCanvassDelete m_rsEMPRVOTSchedule
    End If

    m_rsTimeCharges.Filter = adFilterNone
    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "BuildEMPRVInterface"

End Sub
Private Sub BuildJobTaskInterface(ByVal vstrCanvassState As String)
' For each employee that accepted overtime, add a record for each timecharge row charged to a jobtask.

    Dim l_strRecentRespCd As String

On Error GoTo eh

    Set m_rsJobTaskOTSchedule = NewJobTaskOTScheduleSet
    
    m_rsCanvassSheet.Filter = adFilterNone
    m_rsCanvassSheet.MoveFirst
    
    Do Until m_rsCanvassSheet.EOF
        l_strRecentRespCd = MostRecentResponseCode(m_rsCanvassSheet!CANVASS_RESP_CD, m_rsCanvassSheet!FORCE_RESP_CD)
        If l_strRecentRespCd <> "" Then
            If IsOvertimeWorkedResponse(l_strRecentRespCd) Then
                m_rsTimeCharges.Filter = "EMP_NUM = " & m_rsCanvassSheet!EMP_NUM
                Do Until m_rsTimeCharges.EOF
                    If IsOvertimeWorked(m_rsTimeCharges!EXCEPTION_CD, "") Then
                        If Trim(m_rsTimeCharges!JOB_NBR) <> "" Then
                            Select Case vstrCanvassState
                                Case "Post"
                                    AddJobTaskCanvassInsert m_rsTimeCharges, m_rsJobTaskOTSchedule
                                Case "Void"
                                    AddJobTaskCanvassDelete m_rsTimeCharges, m_rsJobTaskOTSchedule
                                Case "UnPost"
                                    AddJobTaskCanvassDelete m_rsTimeCharges, m_rsJobTaskOTSchedule
                            End Select
                        End If
                    End If
                    m_rsTimeCharges.MoveNext
                Loop
            End If
        End If
        m_rsCanvassSheet.MoveNext
    Loop

    m_rsTimeCharges.Filter = adFilterNone
    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "BuildJobTaskInterface"

End Sub


Private Function NewEMPRVOTScheduleSet() As ADODB.Recordset

    Dim l_rs As ADODB.Recordset

On Error GoTo eh

    Set l_rs = New ADODB.Recordset
    
    With l_rs
        .Fields.Append "OT_OPP_NUM", adInteger, , mCFLDATTRUPDATE
        .Fields.Append "EMP_NUM", adInteger, mCFLDATTRUPDATE
        .Fields.Append "CHARGE_DATE", adDate, mCFLDATTRUPDATE
        .Fields.Append "START_TIME", adChar, 5, mCFLDATTRUPDATE
        .Fields.Append "END_TIME", adChar, 5, mCFLDATTRUPDATE
        .Fields.Append "LOC_CD", adInteger, , mCFLDATTRUPDATE
        .Fields.Append "TRAN_CODE", adChar, 1, mCFLDATTRUPDATE
        .Fields.Append "SHIFT_CD", adChar, 1, mCFLDATTRUPDATE
        .Fields.Append "OT_OPP_DESC", adChar, 30, mCFLDATTRUPDATE

        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open
    End With

    Set NewEMPRVOTScheduleSet = l_rs
    
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "NewEMPRVOTScheduleSet"

End Function
Private Function NewJobTaskOTScheduleSet() As ADODB.Recordset

    Dim l_rs As ADODB.Recordset

On Error GoTo eh

    Set l_rs = New ADODB.Recordset
    
    With l_rs
        
        .Fields.Append "EMP_NUM", adInteger, 10, mCFLDATTRUPDATE
        .Fields.Append "DATE_WORKED", adDate, , mCFLDATTRUPDATE
        .Fields.Append "CHARGE_DATE", adDate, , mCFLDATTRUPDATE
        .Fields.Append "START_TIME", adChar, 5, mCFLDATTRUPDATE
        .Fields.Append "END_TIME", adChar, 5, mCFLDATTRUPDATE
        .Fields.Append "WG_NUM", adInteger, 10, mCFLDATTRUPDATE
        .Fields.Append "JOB_NBR", adVarChar, 8, mCFLDATTRUPDATE
        .Fields.Append "JOBTASK_NBR", adVarChar, 4, mCFLDATTRUPDATE
        .Fields.Append "JOBTASK_STEP_NBR", adInteger, , mCFLDATTRUPDATE
        .Fields.Append "EXCEPTION_CD", adChar, 3, mCFLDATTRUPDATE
        .Fields.Append "DISTRIBUTION_ID", adChar, 39, mCFLDATTRUPDATE
        .Fields.Append "HOURS", adDecimal, , mCFLDATTRUPDATE
        .Fields.Item("HOURS").Precision = 3
        .Fields.Item("HOURS").NumericScale = 1
        .Fields.Append "INDEX_TIME", adChar, 26, mCFLDATTRUPDATE
        .Fields.Append "EMPRV_DATABASE_ID", adVarChar, 8, mCFLDATTRUPDATE
        .Fields.Append "OT_FLG", adChar, 1, mCFLDATTRUPDATE
        .Fields.Append "CHG_IND", adChar, 1, mCFLDATTRUPDATE
        .Fields.Append "DEPT_NUM", adChar, 3, mCFLDATTRUPDATE

        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open
    End With

    Set NewJobTaskOTScheduleSet = l_rs
    
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "NewJobTaskOTScheduleSet"

End Function


Private Sub AddEMPRVCanvassInsert( _
                ByVal v_rsTimeCharge As ADODB.Recordset, _
                ByRef r_rsEMPRVSchedule As ADODB.Recordset)

On Error GoTo eh

    r_rsEMPRVSchedule.AddNew
    r_rsEMPRVSchedule!OT_OPP_NUM = m_rsOTOpportunity!OT_OPP_NUM
    r_rsEMPRVSchedule!CHARGE_DATE = m_rsOTOpportunity!CHARGE_DATE
    r_rsEMPRVSchedule!LOC_CD = m_rsOTOpportunity!LOC_CD
    r_rsEMPRVSchedule!SHIFT_CD = m_rsOTOpportunity!SHIFT_CD
    r_rsEMPRVSchedule!OT_OPP_DESC = m_rsOTOpportunity!OT_OPP_DESC
    
    r_rsEMPRVSchedule!EMP_NUM = v_rsTimeCharge!EMP_NUM
    r_rsEMPRVSchedule!START_TIME = v_rsTimeCharge!START_TIME
    r_rsEMPRVSchedule!END_TIME = v_rsTimeCharge!END_TIME

    r_rsEMPRVSchedule!TRAN_CODE = "I"
    
    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "AddEMPRVCanvassInsert"
    
End Sub
Private Sub AddJobTaskCanvassInsert( _
                ByVal v_rsTimeCharge As ADODB.Recordset, _
                ByRef r_rsJobTaskSchedule As ADODB.Recordset)

Dim l_objWorkGroup As dbWorkGroup.cdWorkGroup

On Error GoTo eh

'Do any of these fields need to come from m_rsotopportunity instead of v_rstimecharge???

    r_rsJobTaskSchedule.AddNew
    r_rsJobTaskSchedule!EMP_NUM = v_rsTimeCharge!EMP_NUM
    r_rsJobTaskSchedule!DATE_WORKED = v_rsTimeCharge!DATE_WORKED
    r_rsJobTaskSchedule!CHARGE_DATE = v_rsTimeCharge!CHARGE_DATE
    r_rsJobTaskSchedule!START_TIME = v_rsTimeCharge!START_TIME
    r_rsJobTaskSchedule!END_TIME = v_rsTimeCharge!END_TIME
    r_rsJobTaskSchedule!WG_NUM = v_rsTimeCharge!WG_NUM
    r_rsJobTaskSchedule!JOB_NBR = v_rsTimeCharge!JOB_NBR
    r_rsJobTaskSchedule!JOBTASK_NBR = v_rsTimeCharge!JOBTASK_NBR
    r_rsJobTaskSchedule!JOBTASK_STEP_NBR = v_rsTimeCharge!JOBTASK_STEP_NBR
    r_rsJobTaskSchedule!EXCEPTION_CD = v_rsTimeCharge!EXCEPTION_CD
    r_rsJobTaskSchedule!DISTRIBUTION_ID = v_rsTimeCharge!DISTRIBUTION_ID
    r_rsJobTaskSchedule!HOURS = v_rsTimeCharge!HOURS
    'r_rsJobTaskSchedule!INDEX_TIME = v_rsTimeCharge!INDEX_TIME
    r_rsJobTaskSchedule!EMPRV_DATABASE_ID = v_rsTimeCharge!EMPRV_DATABASE_ID
    If IsJobTaskOTExcep(v_rsTimeCharge!EXCEPTION_CD) Then
        r_rsJobTaskSchedule!OT_FLG = "Y"
    Else
        r_rsJobTaskSchedule!OT_FLG = "N"
    End If
    r_rsJobTaskSchedule!CHG_IND = "I"
    Set l_objWorkGroup = CreateObject("dbWorkGroup.cdWorkGroup")
    r_rsJobTaskSchedule!DEPT_NUM = l_objWorkGroup.GetDepartmentNumber(v_rsTimeCharge!WG_NUM)
    Set l_objWorkGroup = Nothing
    
    Exit Sub

eh:
    Set l_objWorkGroup = Nothing
    TrsRaiseError m_strCMODULENAME, "AddJobTaskCanvassInsert"
    
End Sub


Private Sub AddEMPRVCanvassDelete(ByRef r_rsEMPRVSchedule As ADODB.Recordset)

On Error GoTo eh

    r_rsEMPRVSchedule.AddNew
    r_rsEMPRVSchedule!OT_OPP_NUM = m_rsOTOpportunity!OT_OPP_NUM
    r_rsEMPRVSchedule!CHARGE_DATE = m_rsOTOpportunity!CHARGE_DATE
    r_rsEMPRVSchedule!LOC_CD = 0
    r_rsEMPRVSchedule!SHIFT_CD = ""
    r_rsEMPRVSchedule!OT_OPP_DESC = "CLEAR CANVASS"
    
    r_rsEMPRVSchedule!EMP_NUM = 0
    r_rsEMPRVSchedule!START_TIME = ""
    r_rsEMPRVSchedule!END_TIME = ""

    r_rsEMPRVSchedule!TRAN_CODE = "D"
    
    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "AddEMPRVCanvassDelete"
    
End Sub
Private Sub AddJobTaskCanvassDelete(ByVal v_rsTimeCharge As ADODB.Recordset, _
                                    ByRef r_rsJobTaskSchedule As ADODB.Recordset)

Dim l_objWorkGroup As dbWorkGroup.cdWorkGroup

On Error GoTo eh

    r_rsJobTaskSchedule.AddNew
    r_rsJobTaskSchedule!EMP_NUM = v_rsTimeCharge!EMP_NUM
    r_rsJobTaskSchedule!DATE_WORKED = v_rsTimeCharge!DATE_WORKED
    r_rsJobTaskSchedule!CHARGE_DATE = v_rsTimeCharge!CHARGE_DATE
    r_rsJobTaskSchedule!START_TIME = v_rsTimeCharge!START_TIME
    r_rsJobTaskSchedule!END_TIME = v_rsTimeCharge!END_TIME
    r_rsJobTaskSchedule!WG_NUM = v_rsTimeCharge!WG_NUM
    r_rsJobTaskSchedule!JOB_NBR = v_rsTimeCharge!JOB_NBR
    r_rsJobTaskSchedule!JOBTASK_NBR = v_rsTimeCharge!JOBTASK_NBR
    r_rsJobTaskSchedule!JOBTASK_STEP_NBR = v_rsTimeCharge!JOBTASK_STEP_NBR
    r_rsJobTaskSchedule!EXCEPTION_CD = v_rsTimeCharge!EXCEPTION_CD
    r_rsJobTaskSchedule!DISTRIBUTION_ID = v_rsTimeCharge!DISTRIBUTION_ID
    r_rsJobTaskSchedule!HOURS = v_rsTimeCharge!HOURS
    'r_rsJobTaskSchedule!INDEX_TIME = v_rsTimeCharge!INDEX_TIME
    r_rsJobTaskSchedule!EMPRV_DATABASE_ID = v_rsTimeCharge!EMPRV_DATABASE_ID
    If IsJobTaskOTExcep(v_rsTimeCharge!EXCEPTION_CD) Then
        r_rsJobTaskSchedule!OT_FLG = "Y"
    Else
        r_rsJobTaskSchedule!OT_FLG = "N"
    End If
    r_rsJobTaskSchedule!CHG_IND = "D"
    Set l_objWorkGroup = CreateObject("dbWorkGroup.cdWorkGroup")
    r_rsJobTaskSchedule!DEPT_NUM = l_objWorkGroup.GetDepartmentNumber(v_rsTimeCharge!WG_NUM)
    Set l_objWorkGroup = Nothing
    Exit Sub

eh:
    Set l_objWorkGroup = Nothing
    TrsRaiseError m_strCMODULENAME, "AddEMPRVCanvassDelete"
    
End Sub


Private Function SaveEMPRVOTSchedule() As Boolean

    Dim l_obj As dbOTCanvass.cdOTCanvass

On Error GoTo eh

    m_rsEMPRVOTSchedule.Filter = adFilterNone
    If Not RecordsExist(m_rsEMPRVOTSchedule) Then
        SaveEMPRVOTSchedule = True
        Exit Function
    End If

    Set l_obj = CreateObject("dbOTCanvass.cdOTCanvass")
    SaveEMPRVOTSchedule = l_obj.SaveEMPRVOTSchedule(m_rsEMPRVOTSchedule)
    Set l_obj = Nothing
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "SaveEMPRVOTSchedule"

End Function
Private Function SaveJobTaskOTSchedule() As Boolean

    Dim l_obj As dbOTCanvass.cdOTCanvass

On Error GoTo eh

    m_rsJobTaskOTSchedule.Filter = adFilterNone
    If Not RecordsExist(m_rsJobTaskOTSchedule) Then
        SaveJobTaskOTSchedule = True
        Exit Function
    End If

    Set l_obj = CreateObject("dbOTCanvass.cdOTCanvass")
    SaveJobTaskOTSchedule = l_obj.SaveJobTaskOTSchedule(m_rsJobTaskOTSchedule)
    Set l_obj = Nothing
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "SaveJobTaskOTSchedule"

End Function


'Private Sub AppendWarning( _
'            ByVal v_rsReq As ADODB.Recordset, _
'            ByVal v_lngErrorNumber As Long, _
'            Optional ByVal v_strErrorDescription As String, _
'            Optional ByRef vcolDescriptionArgs As Collection)
'    Dim l_strErrDesc As String
'
'    If v_strErrorDescription = "" Then
'        l_strErrDesc = TrsGetErrDescription(LongAppErr(v_lngErrorNumber), vcolDescriptionArgs)
'    Else
'        l_strErrDesc = v_strErrorDescription
'    End If
'    AppendEntryError v_rsReq, v_lngErrorNumber, mstrCBUSINESSWAR, l_strErrDesc
'
'End Sub


Private Sub AppendResponseEntryError( _
            ByVal v_rsReq As ADODB.Recordset, _
            ByVal v_lngErrorNumber As Long, _
            ByVal v_strErrorType As String, _
            ByVal v_strErrorDescription As String, _
            Optional ByVal v_strErrorSource As String)
            
            
            
    Dim l_bolError As Boolean
    Dim lngKeyOtOppNum As Long
    Dim lngKeyEmpNum As Long
    Dim l_strResponse As String

    On Error GoTo eh
    
    lngKeyEmpNum = v_rsReq!KEY_EMP_NUM
    lngKeyOtOppNum = v_rsReq!KEY_OT_OPP_NUM
    If IsNull(v_rsReq!CANVASS_RESP_CD) Or IsEmpty(v_rsReq!CANVASS_RESP_CD) Then
        l_strResponse = ""
    Else
        l_strResponse = Trim(v_rsReq!CANVASS_RESP_CD)
    End If
   
    
    l_bolError = AppendBusinessError(m_rsErrors, _
            v_lngErrorNumber, _
            v_strErrorType, _
            v_strErrorDescription, _
            m_strCMODULENAME, _
            10, , , "OT_OPP_NUM;EMP_NUM;CANVASS_RESP_CD", _
            str(lngKeyOtOppNum), _
            str(lngKeyEmpNum), _
            str(l_strResponse))

    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "AppendResponseEntryError"

End Sub


Private Sub AppendTimeChargeEntryError( _
            ByVal v_rsReq As ADODB.Recordset, _
            ByVal v_lngErrorNumber As Long, _
            ByVal v_strErrorType As String, _
            ByVal v_strErrorDescription As String, _
            Optional ByVal v_strErrorSource As String)
            
            
            
    Dim l_bolError As Boolean
    Dim lngKeyOtOppNum As Long
    Dim lngKeyEmpNum As Long
    Dim l_strCreatedTS As String
    Dim l_strLastModTS As String
    Dim l_datChargeDate As Date
    Dim l_datDateWorked As Date
    Dim l_strStartTime As String
    Dim l_strEndTime As String
    Dim l_strExcCd As String
    Dim l_strOverExcCd As String
    Dim l_strFieldNames As String

    On Error GoTo eh
    
    lngKeyEmpNum = v_rsReq!EMP_NUM
    lngKeyOtOppNum = v_rsReq!OT_OPP_NUM
    
    If IsNull(v_rsReq!CREATED_TIMESTAMP) Or IsEmpty(v_rsReq!CREATED_TIMESTAMP) Then
        l_strCreatedTS = ""
    Else
        l_strCreatedTS = v_rsReq!CREATED_TIMESTAMP
    End If
    
    If IsNull(v_rsReq!LAST_MOD_TIMESTAMP) Or IsEmpty(v_rsReq!LAST_MOD_TIMESTAMP) Then
        l_strLastModTS = ""
    Else
        l_strLastModTS = v_rsReq!LAST_MOD_TIMESTAMP
    End If
    
    If IsNull(v_rsReq!CHARGE_DATE) Or IsEmpty(v_rsReq!CHARGE_DATE) Then
        l_datChargeDate = Date
    Else
        l_datChargeDate = v_rsReq!CHARGE_DATE
    End If
    
    If IsNull(v_rsReq!DATE_WORKED) Or IsEmpty(v_rsReq!DATE_WORKED) Then
        l_datDateWorked = Date
    Else
        l_datDateWorked = v_rsReq!DATE_WORKED
    End If
    
    If IsNull(v_rsReq!START_TIME) Or IsEmpty(v_rsReq!START_TIME) Then
        l_strStartTime = ""
    Else
        l_strStartTime = v_rsReq!START_TIME
    End If
    If IsNull(v_rsReq!END_TIME) Or IsEmpty(v_rsReq!END_TIME) Then
        l_strEndTime = ""
    Else
        l_strEndTime = v_rsReq!END_TIME
    End If
    If IsNull(v_rsReq!EXCEPTION_CD) Or IsEmpty(v_rsReq!EXCEPTION_CD) Then
        l_strExcCd = ""
    Else
        l_strExcCd = v_rsReq!EXCEPTION_CD
    End If
    
    If IsNull(v_rsReq!OVER_EXC_CD) Or IsEmpty(v_rsReq!OVER_EXC_CD) Then
        l_strOverExcCd = ""
    Else
        l_strOverExcCd = v_rsReq!OVER_EXC_CD
    End If
    
    l_strFieldNames = "OT_OPP_NUM;EMP_NUM;CREATED_TIMESTAMP;LAST_MOD_TIMESTAMP;" & _
            "CHARGE_DATE;SHIFT_NUM;START_TIME;END_TIME;EXCEPTION_CD;OVER_EXC_CD"
    
    l_bolError = AppendBusinessError(m_rsErrors, _
            v_lngErrorNumber, _
            v_strErrorType, _
            v_strErrorDescription, _
            m_strCMODULENAME, _
            20, , , "EMP_NUM;CHARGE_DATE;SHIFT_NUM,START_TIME,END_TIME,EXCEPTION_CD, OVER_EXC_CD", _
            str(lngKeyOtOppNum), str(lngKeyEmpNum), _
            l_strCreatedTS, l_strLastModTS, _
            str(l_datChargeDate), str(l_datDateWorked), _
            l_strStartTime, l_strEndTime, l_strExcCd, l_strOverExcCd)

    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "AppendTimeChargeEntryError"

End Sub


Private Sub AppendTimeChargeRequestBRV( _
                ByVal v_rsTimeCharge As ADODB.Recordset, _
                ByVal v_lngErrorNumber As Long, _
                ByVal v_strErrorType As String, _
                ByVal v_strErrorDescription As String, _
                Optional ByVal v_strErrorSource As String)
            
    Dim l_lngEmpNum As Long
    Dim l_datChargeDate As Date
    Dim l_datDateWorked As Date
    Dim l_strStartTime As String
    Dim l_decHours As Variant
    Dim l_strExcCd As String
    
On Error GoTo eh
   
    If IsNull(v_rsTimeCharge!KEY_EMP_NUM) Then
        If Not IsNull(v_rsTimeCharge!EMP_NUM) Then
            l_lngEmpNum = v_rsTimeCharge!EMP_NUM
        End If
    Else
        l_lngEmpNum = v_rsTimeCharge!KEY_EMP_NUM
    End If
    
    If IsNull(v_rsTimeCharge!KEY_CHARGE_DATE) Then
        If Not IsNull(v_rsTimeCharge!CHARGE_DATE) Then
            l_datChargeDate = v_rsTimeCharge!CHARGE_DATE
        End If
    Else
        l_datChargeDate = v_rsTimeCharge!KEY_CHARGE_DATE
    End If
    
    If IsNull(v_rsTimeCharge!KEY_DATE_WORKED) Then
        If Not IsNull(v_rsTimeCharge!DATE_WORKED) Then
            l_datDateWorked = v_rsTimeCharge!DATE_WORKED
        End If
    Else
        l_datDateWorked = v_rsTimeCharge!KEY_DATE_WORKED
    End If
    
    If IsNull(v_rsTimeCharge!KEY_START_TIME) Then
        If Not IsNull(v_rsTimeCharge!START_TIME) Then
            l_strStartTime = v_rsTimeCharge!START_TIME
        End If
    Else
        l_strStartTime = v_rsTimeCharge!KEY_START_TIME
    End If
    
    If Not IsNull(v_rsTimeCharge!HOURS) Then
        l_decHours = v_rsTimeCharge!HOURS
    End If

    If Not IsNull(v_rsTimeCharge!EXCEPTION_CD) Then
        l_strExcCd = Trim(v_rsTimeCharge!EXCEPTION_CD)
    End If
    

    AppendBusinessError _
            m_rsErrors, _
            LongAppErr(v_lngErrorNumber), _
            v_strErrorType, _
            v_strErrorDescription, _
            m_strCMODULENAME, _
            TrType_OT_TC, , , "EMP_NUM;CHARGE_DATE;DATE_WORKED;START_TIME;HOURS;EXCEPTION_CD", _
            CStr(l_lngEmpNum), _
            Format(l_datChargeDate, "mm/dd/yyyy"), _
            Format(l_datDateWorked, "mm/dd/yyyy"), _
            l_strStartTime, _
            CStr(l_decHours), _
            l_strExcCd

    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "AppendTimeChargeRequestBRV"
    
End Sub


Private Sub AppendTimeChargeBRV( _
            ByVal v_rsTimeCharge As ADODB.Recordset, _
            ByVal v_lngErrorNumber As Long, _
            ByVal v_strErrorType As String, _
            ByVal v_strErrorDescription As String, _
            Optional ByVal v_strErrorSource As String)
            
    Dim l_lngEmpNum As Long
    Dim l_datChargeDate As Date
    Dim l_datDateWorked As Date
    Dim l_strStartTime As String
    Dim l_decHours As Variant
    Dim l_strExcCd As String
    
On Error GoTo eh
   
    l_lngEmpNum = v_rsTimeCharge!EMP_NUM
    l_datChargeDate = v_rsTimeCharge!CHARGE_DATE
    l_datDateWorked = v_rsTimeCharge!DATE_WORKED
    l_strStartTime = v_rsTimeCharge!START_TIME
    l_decHours = v_rsTimeCharge!HOURS
    l_strExcCd = Trim(v_rsTimeCharge!EXCEPTION_CD)

    AppendBusinessError _
            m_rsErrors, _
            LongAppErr(v_lngErrorNumber), _
            v_strErrorType, _
            v_strErrorDescription, _
            m_strCMODULENAME, _
            TrType_OT_TC, , , "EMP_NUM;CHARGE_DATE;DATE_WORKED;START_TIME;HOURS;EXCEPTION_CD", _
            CStr(l_lngEmpNum), _
            Format(l_datChargeDate, "mm/dd/yyyy"), _
            Format(l_datDateWorked, "mm/dd/yyyy"), _
            l_strStartTime, _
            CStr(l_decHours), _
            l_strExcCd

    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "AppendTimeChargeBRV"
    
End Sub


Private Sub AppendResponseRequestBRV( _
            ByVal v_rsResponseReq As ADODB.Recordset, _
            ByVal v_lngErrorNumber As Long, _
            ByVal v_strErrorType As String, _
            ByVal v_strErrorDescription As String, _
            Optional ByVal v_strErrorSource As String)
            
    Dim l_lngOTOppNum As Long
    Dim l_lngEmpNum As Long
    
On Error GoTo eh
   
    l_lngOTOppNum = v_rsResponseReq!KEY_OT_OPP_NUM
    l_lngEmpNum = v_rsResponseReq!KEY_EMP_NUM

    AppendBusinessError _
            m_rsErrors, _
            LongAppErr(v_lngErrorNumber), _
            v_strErrorType, _
            v_strErrorDescription, _
            m_strCMODULENAME, _
            TrType_OT_R, , , "EMP_NUM", _
            CStr(l_lngEmpNum)

    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "AppendResponseRequestBRV"
    
End Sub



'Private Function GenerateTimeCharges() As Boolean
''Returns true if a time charge was generated from a response.
''Generate timecharges for acceptors without any timecharges.
''Generate timecharges for refusers based on replacing man.
'
''    Dim l_strFilter As String
'    Dim l_blnTimeChargeGen As Boolean
'    Dim l_lngEmpNum As Long
'
'On Error GoTo eh
'
'    l_blnTimeChargeGen = False
''    m_rsResponses.Filter = adFilterNone
''    If Not RecordsExist(m_rsResponses) Then
''        GenerateTimeCharges = l_blnTimeChargeGen
''        Exit Function
''    End If
'
'    GetEmployeeScheduleState
'
''    Acceptors
'    m_rsCanvassSheet.Filter = adFilterPendingRecords
'    Do Until m_rsCanvassSheet.EOF
'        l_lngEmpNum = m_rsCanvassSheet!EMP_NUM
'        If IsAcceptorResponse(m_rsCanvassSheet!CANVASS_RESP_CD) And _
'                Trim(m_rsCanvassSheet.Fields("CANVASS_RESP_CD").OriginalValue) = "" Then
'            If Not TimeChargeRecordsExist(l_lngEmpNum) Then
'                l_blnTimeChargeGen = True
'                AddTimeChargeFromOpportunity l_lngEmpNum
'            End If
'        End If
'        m_rsCanvassSheet.MoveNext
'    Loop
'    m_rsCanvassSheet.Filter = adFilterNone
'
''First, work on getting new canvass. Then work on changes.
'
''    Refusers
''    m_rsResponses.Filter = adFilterNone
''    Do Until m_rsResponses.EOF
''        If Not IsAcceptorResponse(m_rsResponses!OT_RESP_CD) Then
''            ' process refusers
''            ' find replacing man
''            ' add all time charges of replacing man
''            ' do not add to schedule
''        End If
''        m_rsResponses.MoveNext
''    Loop
'
'
'    GenerateTimeCharges = l_blnTimeChargeGen
'    Exit Function
'
'eh:
'    TrsRaiseError m_strCMODULENAME, "GenerateTimeCharges"
'
'End Function


Private Function IsAcceptorResponse(ByVal v_strRespCd As String) As Boolean
    Dim l_obj As dbCodeTables.cdCodeTables

On Error GoTo eh

    Set l_obj = CreateObject("dbCodeTables.cdCodeTables")
    IsAcceptorResponse = (l_obj.OvertimeResponseWorkIndicator(v_strRespCd) = "Y")

    Set l_obj = Nothing
    Exit Function

eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "IsAcceptorResponse"
    
End Function


Private Function TimeChargeRecordsExist(ByVal v_lngEmpNum As Long) As Boolean

On Error GoTo eh

    m_rsTimeCharges.Filter = "EMP_NUM = " & v_lngEmpNum
    TimeChargeRecordsExist = Not m_rsTimeCharges.EOF
    
    m_rsTimeCharges.Filter = adFilterNone
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "TimeChargeRecordsExist"

End Function


Private Function EmployeeSchedulesAreRequired() As Boolean
'Get the schedules for all acceptors that have changed.

    Dim l_colUnionCodes As Collection
    Dim l_colChargeDates As Collection

On Error GoTo eh
 
    Set l_colChargeDates = m_setChargeDates.Contents
    If l_colChargeDates.Count = 0 Then
        l_colChargeDates.Add m_rsOTOpportunity.Fields("CHARGE_DATE").Value
    End If
    Set l_colUnionCodes = GetEmployeeUnions(m_rsEmployees, l_colChargeDates)
    If IncludesString("OA", l_colUnionCodes) Then
        EmployeeSchedulesAreRequired = True
    End If
    
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "EmployeeSchedulesAreRequired"

End Function


Private Sub GetEmployeeScheduleState( _
                ByVal v_rsCanvassResponsesReq As ADODB.Recordset, _
                ByVal v_rsTimeChargeReq As ADODB.Recordset)
'Get the schedules for all acceptors that have changed.

    Dim l_aryEmpNums As Variant
    Dim l_datTimeChgStart As Date
    Dim l_datTimeChgEnd As Date
    Dim l_datStart As Date
    Dim l_datEnd As Date
    Dim l_aryCanvassedEmpNums As Variant

On Error GoTo eh

    l_aryEmpNums = EmployeesWithUpdatedResponses(v_rsCanvassResponsesReq, v_rsTimeChargeReq)
    If IsEmpty(l_aryEmpNums) Then
        GoTo xit
    End If
    
    l_aryCanvassedEmpNums = EmployeesWithAnyResponse(m_rsCanvassSheet, v_rsCanvassResponsesReq)
    Set m_rsEmployees = GetEmployees(l_aryCanvassedEmpNums)

' RSG - needs work
' Get all schedules for now, as adding TTA.
'    AddChargeDate m_rsOTOpportunity.Fields("CHARGE_DATE").Value
'    If Not EmployeeSchedulesAreRequired Then
'        GoTo xit
'    End If

    
    FindHighLowTimeChargeDates l_datTimeChgStart, l_datTimeChgEnd, v_rsTimeChargeReq
    l_datStart = DateAdd("d", -1, l_datTimeChgStart)
    l_datEnd = DateAdd("d", 1, l_datTimeChgEnd)

    Set m_rsEmpSchedules = _
        GetEmployeeSchedules(l_aryEmpNums, l_datStart, l_datEnd)
    
    If IsCanvassPosted Then
        Set m_rsEmpOtherPay = _
            GetOtherPayEntries(l_aryCanvassedEmpNums, l_datStart, l_datEnd)
    End If
    
xit:
    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "GetEmployeeScheduleState"

End Sub


Private Sub GetCanvassedEmployeeScheduleState()
'Get the schedules for all acceptors.

    Dim l_aryEmpNums As Variant
    Dim l_datChgDate As Date
    Dim l_datStart As Date
    Dim l_datEnd As Date

    Dim stTime As Long
    Dim ndTime As Long
    Dim dfTime As Long

On Error GoTo eh

    #If TRTRACE Then
        stTime = GetTickCount()
    #End If
    
'    l_aryEmpNums = EmployeesWithAcceptedResponses(m_rsCanvassSheet)
    l_aryEmpNums = EmployeesWithAnyResponse(m_rsCanvassSheet, Nothing)
    
    If IsEmpty(l_aryEmpNums) Then
        GoTo xit
    End If
    
    Set m_rsEmployees = GetEmployees(l_aryEmpNums)
    
    l_datChgDate = m_rsOTOpportunity!CHARGE_DATE
    l_datStart = DateAdd("d", -1, l_datChgDate)
    l_datEnd = DateAdd("d", 1, l_datChgDate)

    Set m_rsEmpSchedules = _
        GetEmployeeSchedules(l_aryEmpNums, l_datStart, l_datEnd)
    
    Set m_rsEmpOtherPay = _
        GetOtherPayEntries(l_aryEmpNums, l_datStart, l_datEnd)
    
xit:

    #If TRTRACE Then
        ndTime = GetTickCount()
        dfTime = ndTime - stTime
        TraceAction App.ThreadID, "3", "GetCanvassedEmployeeScheduleState", Format(Now, "mm/dd/yyyy hh:mm:ss"), _
            CStr(dfTime / 1000)
    #End If

    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "GetCanvassedEmployeeScheduleState"

End Sub


'Private Function EmployeesWithoutTimecharges() As Variant
''Answer array of employee numbers.
'    Dim l_objSet As utilTRISCommonLib.cuSet
'
'On Error GoTo eh
'
'    m_rsCanvassSheet.Filter = adFilterNone
'    If m_rsCanvassSheet.EOF Then
'        EmployeesWithoutTimecharges = Empty
'        GoTo xit
'    End If
'
'    Set l_objSet = New cuSet
'    Do Until m_rsCanvassSheet.EOF
'        If Trim(m_rsCanvassSheet!CANVASS_RESP_CD) <> "" Then
'            m_rsTimeCharges.Filter = "EMP_NUM = " & m_rsCanvassSheet!EMP_NUM
'            If m_rsTimeCharges.EOF Then
'                l_objSet.Add m_rsCanvassSheet!EMP_NUM
'            End If
'        End If
'        m_rsCanvassSheet.MoveNext
'    Loop
'
'    EmployeesWithoutTimecharges = ArrayFromCollection(l_objSet.Contents)
'
'xit:
'    m_rsCanvassSheet.MoveFirst
'    m_rsTimeCharges.Filter = adFilterNone
'    Set l_objSet = Nothing
'    Exit Function
'
'eh:
'    Set l_objSet = Nothing
'    TrsRaiseError m_strCMODULENAME, "EmployeesWithoutTimecharges"
'
'End Function

'
'Private Function EmployeesWithUpdatedResponses() As Variant
''Answer array of employee numbers.
'    Dim l_objSet As utilTRISCommonLib.cuSet
'
'On Error GoTo eh
'
'    m_rsCanvassSheet.Filter = adFilterNone      'clear before filter pending
'    m_rsCanvassSheet.Filter = adFilterPendingRecords
'    If m_rsCanvassSheet.EOF Then
'        EmployeesWithUpdatedResponses = Empty
'        GoTo xit
'    End If
'
'    Set l_objSet = New cuSet
'    Do Until m_rsCanvassSheet.EOF
'        l_objSet.Add m_rsCanvassSheet.Fields("EMP_NUM").Value
'        m_rsCanvassSheet.MoveNext
'    Loop
'
'    EmployeesWithUpdatedResponses = ArrayFromCollection(l_objSet.Contents)
'
'xit:
'    m_rsCanvassSheet.Filter = adFilterNone
'    Set l_objSet = Nothing
'    Exit Function
'
'eh:
'    Set l_objSet = Nothing
'    TrsRaiseError m_strCMODULENAME, "EmployeesWithUpdatedResponses"
'
'End Function


Private Function EmployeesWithUpdatedResponses( _
                    ByVal v_rsCanvassResponseReq As ADODB.Recordset, _
                    ByVal v_rsTimeChargeReq As ADODB.Recordset) As Variant
'Answer array of employee numbers.
    Dim l_objSet As utilTRISCommonLib.cuSet

On Error GoTo eh

    Set l_objSet = New cuSet
    
    If RecordsExist(v_rsCanvassResponseReq) Then
        v_rsCanvassResponseReq.MoveFirst
        Do Until v_rsCanvassResponseReq.EOF
            l_objSet.Add v_rsCanvassResponseReq.Fields("KEY_EMP_NUM").Value
            v_rsCanvassResponseReq.MoveNext
        Loop
        v_rsCanvassResponseReq.MoveFirst
    End If
    
    If RecordsExist(v_rsTimeChargeReq) Then
        v_rsTimeChargeReq.MoveFirst
        Do Until v_rsTimeChargeReq.EOF
            l_objSet.Add v_rsTimeChargeReq.Fields("KEY_EMP_NUM").Value
            v_rsTimeChargeReq.MoveNext
        Loop
        v_rsTimeChargeReq.MoveFirst
    End If
    
    EmployeesWithUpdatedResponses = ArrayFromCollection(l_objSet.Contents)

xit:
    Set l_objSet = Nothing
    Exit Function

eh:
    Set l_objSet = Nothing
    TrsRaiseError m_strCMODULENAME, "EmployeesWithUpdatedResponses"

End Function


Private Function EmployeesWithAnyResponse( _
                    ByVal v_rsCanvassSheet As ADODB.Recordset, _
                    ByVal v_rsCanvassResponseReq As ADODB.Recordset) As Variant
'Answer array of employee numbers.
    Dim l_objSet As utilTRISCommonLib.cuSet
    Dim l_strRespCd As String

On Error GoTo eh

    Set l_objSet = New cuSet
    
    If RecordsExist(v_rsCanvassSheet) Then
        v_rsCanvassSheet.MoveFirst
        Do Until v_rsCanvassSheet.EOF
            l_strRespCd = _
                MostRecentResponseCode(v_rsCanvassSheet!CANVASS_RESP_CD, v_rsCanvassSheet!FORCE_RESP_CD)
            If l_strRespCd <> "" Then
                l_objSet.Add v_rsCanvassSheet.Fields("EMP_NUM").Value
            End If
            v_rsCanvassSheet.MoveNext
        Loop
        v_rsCanvassSheet.MoveFirst
    End If
    
    If RecordsExist(v_rsCanvassResponseReq) Then
        v_rsCanvassResponseReq.MoveFirst
        Do Until v_rsCanvassResponseReq.EOF
            l_objSet.Add v_rsCanvassResponseReq.Fields("KEY_EMP_NUM").Value
            v_rsCanvassResponseReq.MoveNext
        Loop
        v_rsCanvassResponseReq.MoveFirst
    End If
    
    EmployeesWithAnyResponse = ArrayFromCollection(l_objSet.Contents)

xit:
    Set l_objSet = Nothing
    Exit Function

eh:
    Set l_objSet = Nothing
    TrsRaiseError m_strCMODULENAME, "EmployeesWithAnyResponse"

End Function


Private Function EmployeesWithAcceptedResponses( _
                    ByVal v_rsCanvassSheet As ADODB.Recordset) As Variant
'Answer array of employee numbers.
    Dim l_objSet As utilTRISCommonLib.cuSet
    Dim l_strRespCd As String

On Error GoTo eh

    Set l_objSet = New cuSet
    
    If RecordsExist(v_rsCanvassSheet) Then
        v_rsCanvassSheet.MoveFirst
        Do Until v_rsCanvassSheet.EOF
            l_strRespCd = _
                MostRecentResponseCode(v_rsCanvassSheet!CANVASS_RESP_CD, v_rsCanvassSheet!FORCE_RESP_CD)
            If IsAcceptorResponse(l_strRespCd) Then
                l_objSet.Add v_rsCanvassSheet.Fields("EMP_NUM").Value
            End If
            v_rsCanvassSheet.MoveNext
        Loop
        v_rsCanvassSheet.MoveFirst
    End If
    
    EmployeesWithAcceptedResponses = ArrayFromCollection(l_objSet.Contents)

xit:
    Set l_objSet = Nothing
    Exit Function

eh:
    Set l_objSet = Nothing
    TrsRaiseError m_strCMODULENAME, "EmployeesWithAcceptedResponses"

End Function


Private Function AnyCanvassResponsesExists() As Boolean
'Answer array of employee numbers.
    Dim l_blnReturn As Boolean

On Error GoTo eh

    AnyCanvassResponsesExists = False
    
    m_rsCanvassSheet.Filter = adFilterNone
    
    Do Until m_rsCanvassSheet.EOF
        If Trim(m_rsCanvassSheet!CANVASS_RESP_CD) <> "" Then
            AnyCanvassResponsesExists = True
            Exit Do
        End If
        m_rsCanvassSheet.MoveNext
    Loop

xit:
    m_rsCanvassSheet.MoveFirst
    Exit Function

eh:

    TrsRaiseError m_strCMODULENAME, "AnyCanvassResponsesExists"

End Function



Private Function GetScheduledHolidays( _
                    ByVal v_aryEmployees As Variant, _
                    ByVal v_datStartDate As Date, _
                    ByVal v_datEndDate As Date _
                ) As ADODB.Recordset
         
    Dim l_obj As dbOTCanvass.cdOTCanvass
    
On Error GoTo eh
             
    Set l_obj = CreateObject("dbOTCanvass.cdOTCanvass")
    Set GetScheduledHolidays = _
        l_obj.GetEmployeeScheduledHolidays(v_aryEmployees, v_datStartDate, v_datEndDate)
    Set l_obj = Nothing
    
    Exit Function
    
eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "GetScheduledHolidays"
     
End Function


Private Function IsEntitledToTTA( _
                    ByVal v_lngEmpNum As Long, _
                    ByVal v_datStartDateTime As Date, _
                    ByVal v_datEndDateTime As Date) As Boolean
    
    Dim l_lngOtherOTOppNum As Long
    Dim l_datStartDateTime As Date
    Dim l_datEndDateTime As Date

On Error GoTo eh

    m_rsEmpSchedules.Filter = "EMP_NUM = " & v_lngEmpNum
    
    If Not IsEmployeeOnActivePPMProject(v_lngEmpNum, m_rsOTOpportunity!CHARGE_DATE) Then
        IsEntitledToTTA = False
        GoTo xit
    End If

    If TimespanContinuousRegularScheduledShift(v_datStartDateTime, v_datEndDateTime, 0, 0) Then
        IsEntitledToTTA = False
        GoTo xit
    End If

    If TimespanContinuousOvertimeShift(v_datStartDateTime, v_datEndDateTime, _
            m_lngOTOppNum, l_lngOtherOTOppNum, 0, 0) Then
        If PaidTTAForOvertime(v_lngEmpNum, l_lngOtherOTOppNum) Then
            IsEntitledToTTA = False
            GoTo xit
        ElseIf EmployeeScheduledStartEndDateTimes(v_lngEmpNum, l_lngOtherOTOppNum, _
                    l_datStartDateTime, l_datEndDateTime) Then
            If TimespanContinuousRegularScheduledShift(l_datStartDateTime, l_datEndDateTime, 0, 0) Then
                IsEntitledToTTA = False
                GoTo xit
            End If
        End If
    End If
    
    IsEntitledToTTA = True

xit:
    m_rsEmpSchedules.Filter = adFilterNone
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "IsEntitledToTTA"

End Function


Private Function IsEmployeeOnActivePPMProject( _
                    ByVal v_lngEmpNum As Long, _
                    ByVal v_datChgDate As Date) As Boolean

    Dim l_blnReturn As Boolean

On Error GoTo eh

    l_blnReturn = False
    m_rsEmpWorkgroup.Filter = "EMP_NUM =" & v_lngEmpNum
    If m_rsEmpWorkgroup.EOF Then
        GoTo xit
    End If

    If m_rsEmpWorkgroup!PROJECT_ID > 0 Then
        If Trim(m_rsEmpWorkgroup!PPM_IND = "Y") And _
                m_rsEmpWorkgroup!PROJ_START_DATE <= v_datChgDate And _
                m_rsEmpWorkgroup!PROJ_END_DATE >= v_datChgDate Then
            l_blnReturn = True
        End If
    End If

xit:
    IsEmployeeOnActivePPMProject = l_blnReturn
    m_rsEmpWorkgroup.Filter = adFilterNone
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "IsEmployeeOnActivePPMProject"

End Function


Private Function IsEmployeeEMPRV( _
                    ByVal v_lngEmpNum As Long, _
                    ByVal v_datChgDate As Date) As Boolean
    
    Dim l_strIntToCd As String

On Error GoTo eh
   
    m_rsEmployees.Filter = "EMP_NUM = " & v_lngEmpNum
    If v_datChgDate >= m_rsEmployees!DEPT_EFF_DATE Then
        l_strIntToCd = Trim(m_rsEmployees!INTERFACE_TO_CD)
    Else
        l_strIntToCd = Trim(m_rsEmployees!PREV_INTERFC_TO_CD)
    End If
    
    IsEmployeeEMPRV = l_strIntToCd = "E"
    m_rsEmployees.Filter = adFilterNone

    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "IsEmployeeEMPRV"
    
End Function
Private Function IsJobTaskEmployee( _
    ByVal v_lngEmpNum As Long, _
    ByVal v_datChgDate As Date) As Boolean
    
    Dim l_obj As busDepartment.cbsDepartment
    Dim l_strDept As String
    
    On Error GoTo eh
    
    m_rsEmployees.Filter = "EMP_NUM = " & v_lngEmpNum
    If v_datChgDate >= m_rsEmployees!DEPT_EFF_DATE Then
        l_strDept = Trim(m_rsEmployees!DEPT_NUM)
    Else
        l_strDept = Trim(m_rsEmployees!PREV_DEPT_NUM)
    End If
    Set l_obj = CreateObject("busDepartment.cbsDepartment")
    
    IsJobTaskEmployee = l_obj.IsJobTaskDepartmentNumber(l_strDept)
    m_rsEmployees.Filter = adFilterNone
    
    Exit Function
    
eh:
     TrsRaiseError m_strCMODULENAME, "IsEMPRVEmployee"
    
End Function


Private Function EmployeePrimaryWorkgroup( _
                    ByVal v_lngEmpNum As Long) As Long


On Error GoTo eh

    m_rsEmpWorkgroup.Filter = "EMP_NUM = " & v_lngEmpNum
    If m_rsEmpWorkgroup.EOF Then
        EmployeePrimaryWorkgroup = 0
    Else
        EmployeePrimaryWorkgroup = m_rsEmpWorkgroup!WG_NUM
    End If

    m_rsEmpWorkgroup.Filter = adFilterNone
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "EmployeePrimaryWorkgroup"

End Function


Private Function PaidTTAForOvertime( _
                    ByVal v_lngEmpNum As Long, _
                    ByVal v_lngOTOppNum As Long) As Boolean


On Error GoTo eh

    PaidTTAForOvertime = False
    If m_rsEmpOtherPay Is Nothing Then
        GoTo xit
    Else
        m_rsEmpOtherPay.Filter = "EMP_NUM = " & v_lngEmpNum
    End If

    If RecordsExist(m_rsEmpOtherPay) Then
        Do Until m_rsEmpOtherPay.EOF
            If m_rsEmpOtherPay!OT_OPP_NUM = v_lngOTOppNum And _
                    m_rsEmpOtherPay!EXCEPTION_CD = "TTA" Then
                PaidTTAForOvertime = True
                GoTo xit
            End If
            m_rsEmpOtherPay.MoveNext
        Loop
    End If


xit:
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "PaidTTAForOvertime"
End Function


Private Sub DetermineMissingComponents( _
                    ByRef r_typCanvType As CanvassedShift, _
                    ByRef r_rsTimeCharge As ADODB.Recordset)

    Dim l_strUnionCd As String

On Error GoTo eh

'    Probably take this out in order to provide informational message when
'    exception code provided deviates from pay rule implemented
'    If Trim(r_rsTimeCharge!EXCEPTION_CD) <> "" Then
'        GoTo xit
'    End If
    
'    Probably take this out in order to provide informational message when
'    overriden travel pay deviates from pay rule implemented
'    If r_rsTimeCharge!TRAVEL_PAY_HRS >= 0 Then
'        GoTo xit
'    End If

'    If Trim(r_rsTimeCharge!EXCEPTION_CD) <> "" And _
'            r_rsTimeCharge!TRAVEL_PAY_HRS >= 0 Then
'        GoTo xit
'    End If

    l_strUnionCd = EmployeeUnion(r_rsTimeCharge!EMP_NUM, r_rsTimeCharge!CHARGE_DATE, m_rsEmployees)
    
    Select Case l_strUnionCd
        Case "OA"
            DetermineOvertimeRateOA r_typCanvType, r_rsTimeCharge
            DetermineTravelTimeOA r_typCanvType, r_rsTimeCharge
        Case "FA", "IA", "JA", "JB", "TA", "TB"
            If Trim(r_rsTimeCharge!EXCEPTION_CD) = "" Then
                r_rsTimeCharge!EXCEPTION_CD = "OT4"
            End If
            r_rsTimeCharge!TRAVEL_PAY_HRS = CDec(0#)
        Case "SO"
            If Trim(r_rsTimeCharge!EXCEPTION_CD) = "" Then
                r_rsTimeCharge!EXCEPTION_CD = "OT2"
            End If
            r_rsTimeCharge!TRAVEL_PAY_HRS = CDec(0#)
        Case Else
            r_rsTimeCharge!TRAVEL_PAY_HRS = CDec(0#)
            GoTo xit
    End Select

xit:
    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "DetermineMissingComponents"
    
End Sub


Private Function OvertimeRateCalculated( _
                    ByVal v_lngEmpNum As Long, _
                    ByVal v_datChargeDate As Date) As Boolean

    Dim l_strUnionCd As String

On Error GoTo eh

    l_strUnionCd = EmployeeUnion(v_lngEmpNum, v_datChargeDate, m_rsEmployees)
            
    Select Case l_strUnionCd
        Case "OA", "FA", "IA", "JA", "JB", "TA", "TB", "SO"
            OvertimeRateCalculated = True
        Case Else
            OvertimeRateCalculated = False
    End Select

xit:
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "OvertimeRateCalculated"
    
End Function


Private Sub DetermineOvertimeRateOA( _
                    ByRef r_typCanvShift As CanvassedShift, _
                    ByRef r_rsTimeCharge As ADODB.Recordset)

On Error GoTo eh

    If Trim(r_rsTimeCharge!EXCEPTION_CD) <> "" Then
        GoTo xit
    End If

    r_rsTimeCharge!EXCEPTION_CD = _
        DetermineOvertimeExceptionCodeOA(r_typCanvShift, r_rsTimeCharge)

xit:
    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "DetermineOvertimeRateOA"
                    
End Sub


Private Function DetermineOvertimeExceptionCodeOA( _
                    ByRef r_typCanvShift As CanvassedShift, _
                    ByRef r_rsTimeCharge As ADODB.Recordset) As String

    Dim l_strOvertimeRate As String
    Dim l_decPrevSchedHours As Variant
    Dim l_decCanvassHours As Variant
    Dim l_decOvertimeHours As Variant
    Dim l_decSplitHours As Variant
    Dim l_varBookmark As Variant
    Dim l_blnCallout As Boolean
    Dim l_lngEmpNum As Long
    Dim l_strErrorMessage As String
    Dim l_strExcCd As String
    Dim l_lngOtherOTOppNum As Long

On Error GoTo eh
     
    l_lngEmpNum = r_rsTimeCharge!EMP_NUM
    m_rsEmpSchedules.Filter = "EMP_NUM = " & l_lngEmpNum
    
    #If TRTRACE Then
        TraceActionMessage "3", "DetermineOvertimeExceptionCodeOA", _
            "Shift Start: " & r_typCanvShift.StartDateTime & " End: " & _
            r_typCanvShift.EndDateTime & _
            vbCrLf & " Timecharge: " & FormatTimeCharge(r_rsTimeCharge), True
    #End If
    
    If TimeOutsideRegularScheduledHours( _
                r_typCanvShift.StartDateTime, _
                r_typCanvShift.EndDateTime) Then
        l_strExcCd = "OT3"
        #If TRTRACE Then
            TraceActionMessage "3", "DetermineOvertimeExceptionCodeOA", _
                "TimeOutsideRegularScheduledHours = True, " & _
                "Exception code = " & l_strExcCd
        #End If
    Else
    
        #If TRTRACE Then
            TraceActionMessage "3", "DetermineOvertimeExceptionCodeOA", _
                "TimeOutsideRegularScheduledHours = False"
        #End If

        l_strErrorMessage = "Overtime conflicts with regular scheduled shift"
        AppendProcessTimeChargeError trErrbusOTConflictRegShiftExc, l_strErrorMessage, _
            r_rsTimeCharge!EMP_NUM, r_typCanvShift, r_rsTimeCharge
        GoTo xit
    End If
    
    If IsCalloutOvertime(r_rsTimeCharge!EMP_NUM, r_typCanvShift) Then
        l_blnCallout = True
        l_strExcCd = "OT5"
        #If TRTRACE Then
            TraceActionMessage "3", "DetermineOvertimeExceptionCodeOA", _
                "IsCalloutOvertime = " & l_blnCallout & _
                ", Exception code = " & l_strExcCd
        #End If
                
        GoTo xit
    Else    'prearranged
        l_blnCallout = False
        l_strExcCd = "OT3"
        #If TRTRACE Then
            TraceActionMessage "3", "DetermineOvertimeExceptionCodeOA", _
                "IsCalloutOvertime = " & l_blnCallout & _
                ", Exception code = " & l_strExcCd
        #End If
    End If
    
    If TimespanContinuousRegularScheduledShift( _
            r_typCanvShift.StartDateTime, r_typCanvShift.EndDateTime) Then
        l_strExcCd = "OT5"
        #If TRTRACE Then
            TraceActionMessage "3", "DetermineOvertimeExceptionCodeOA", _
                "TimespanContinuousRegularScheduledShift = True, " & _
                    "Exception code = " & l_strExcCd
        #End If
        GoTo xit
    End If
        
    
    If IsHolidayScheduled(r_rsTimeCharge!CHARGE_DATE) Then
        l_strExcCd = "OT5"
        #If TRTRACE Then
            TraceActionMessage "3", "DetermineOvertimeExceptionCodeOA", _
                "IsHolidayScheduled = True, (" & r_rsTimeCharge!CHARGE_DATE & "), " & _
                ", Exception code = " & l_strExcCd
        #End If
        GoTo xit
    End If
        
'    l_decCanvassHours = CDec(0)
'    If IsDayOff(r_rsTimeCharge!CHARGE_DATE) Then
'        l_decPrevSchedHours = _
'            HoursPreviouslyScheduledContinuousOvertime( _
'                r_rsTimeCharge!EMP_NUM, _
'                r_rsTimeCharge!CHARGE_DATE, _
'                r_rsTimeCharge!DATE_WORKED, _
'                r_rsTimeCharge!START_TIME, _
'                r_rsTimeCharge!END_TIME)
'
'        l_decOvertimeHours = l_decPrevSchedHours + l_decCanvassHours
'        If l_decOvertimeHours + r_rsTimeCharge!HOURS <= 8 Then
'            l_strExcCd = "OT3"
'            #If TRTRACE Then
'                TraceActionMessage "3", "DetermineOvertimeExceptionCodeOA", _
'                    "IsDayOff = True,(" & r_rsTimeCharge!CHARGE_DATE & "), " & _
'                    ", SchedHours <=8, Scheduled Hours = " & l_decPrevSchedHours & _
'                    ", Exception code = " & l_strExcCd
'            #End If
'        ElseIf l_decOvertimeHours >= 8 Then
'            l_strExcCd = "OT5"
'            #If TRTRACE Then
'                TraceActionMessage "3", "DetermineOvertimeExceptionCodeOA", _
'                    ", IsDayOff = True,(" & r_rsTimeCharge!CHARGE_DATE & "), " & _
'                    ", SchedHours >=8, Scheduled Hours = " & l_decPrevSchedHours & _
'                    ", Exception code = " & l_strExcCd
'            #End If
'        Else
'            l_decSplitHours = 8 - l_decOvertimeHours
'            l_varBookmark = r_rsTimeCharge.Bookmark
'            l_strExcCd = "OT3"
'            #If TRTRACE Then
'                TraceActionMessage "3", "DetermineOvertimeExceptionCodeOA", _
'                    "IsDayOff = True,(" & r_rsTimeCharge!CHARGE_DATE & "), " & _
'                    ", Scheduled Hours = " & l_decPrevSchedHours & _
'                    ", Split Hours = " & l_decSplitHours & _
'                    ", Exception code = " & l_strExcCd
'            #End If
'            SplitTimeCharge l_decSplitHours, r_rsTimeCharge
'            r_rsTimeCharge.Bookmark = l_varBookmark
'        End If
'    End If

    l_decCanvassHours = CDec(0)
    If IsHolidayOff(r_rsTimeCharge!CHARGE_DATE) Then
        l_decPrevSchedHours = _
            HoursPreviouslyScheduledOvertime( _
                r_rsTimeCharge!EMP_NUM, _
                r_rsTimeCharge!CHARGE_DATE, _
                r_rsTimeCharge!DATE_WORKED, _
                r_rsTimeCharge!START_TIME, _
                r_rsTimeCharge!END_TIME)

        l_decOvertimeHours = l_decPrevSchedHours + l_decCanvassHours
        If l_decOvertimeHours + r_rsTimeCharge!HOURS <= 8 Then
            l_strExcCd = "OT3"
            #If TRTRACE Then
                TraceActionMessage "3", "DetermineOvertimeExceptionCodeOA", _
                    "IsHolidayOff = True,(" & r_rsTimeCharge!CHARGE_DATE & "), " & _
                    ", SchedHours <=8, Scheduled Hours = " & l_decPrevSchedHours & _
                    ", Exception code = " & l_strExcCd
            #End If
        ElseIf l_decOvertimeHours >= 8 Then
            l_strExcCd = "OT5"
            #If TRTRACE Then
                TraceActionMessage "3", "DetermineOvertimeExceptionCodeOA", _
                    ", IsHolidayOff = True,(" & r_rsTimeCharge!CHARGE_DATE & "), " & _
                    ", SchedHours >=8, Scheduled Hours = " & l_decPrevSchedHours & _
                    ", Exception code = " & l_strExcCd
            #End If
        Else
            l_decSplitHours = 8 - l_decOvertimeHours
            l_varBookmark = r_rsTimeCharge.Bookmark
            l_strExcCd = "OT3"
            #If TRTRACE Then
                TraceActionMessage "3", "DetermineOvertimeExceptionCodeOA", _
                    "IsHolidayOff = True,(" & r_rsTimeCharge!CHARGE_DATE & "), " & _
                    ", Scheduled Hours = " & l_decPrevSchedHours & _
                    ", Split Hours = " & l_decSplitHours & _
                    ", Exception code = " & l_strExcCd
            #End If
            SplitTimeCharge l_decSplitHours, r_rsTimeCharge
            r_rsTimeCharge.Bookmark = l_varBookmark
        End If
    End If
    
' More than 8 continuous overtime
' If already entitled to highest pay rate, skip this.

    If l_strExcCd = "OT5" Then
        GoTo xit
    End If

    If TimespanIsPaidLunchOvertimeShift( _
            r_rsTimeCharge!EMP_NUM, _
            DateTimeFrom(r_rsTimeCharge!DATE_WORKED, r_rsTimeCharge!START_TIME), _
            DateTimeFrom(r_rsTimeCharge!DATE_WORKED, r_rsTimeCharge!END_TIME, True), _
            m_lngOTOppNum, l_lngOtherOTOppNum) Then
        If IsEmployeePaidDoubleForOvertimeShift(r_rsTimeCharge!EMP_NUM, l_lngOtherOTOppNum) Then
            l_strExcCd = "OT5"
            #If TRTRACE Then
                TraceActionMessage "3", "DetermineOvertimeExceptionCodeOA", _
                    ", IsEmployeePaidDoubleForOvertimeShift = True,(" & r_rsTimeCharge!CHARGE_DATE & "), " & _
                    ", Exception code = " & l_strExcCd
            #End If
            GoTo xit
        Else
            l_decPrevSchedHours = _
                ContinuousScheduledOvertimeHours( _
                    r_rsTimeCharge!EMP_NUM, _
                    r_rsTimeCharge!DATE_WORKED, _
                    r_rsTimeCharge!START_TIME, _
                    r_rsTimeCharge!END_TIME, _
                    m_lngOTOppNum, _
                    r_typCanvShift.StartDateTime, _
                    r_typCanvShift.EndDateTime)
        End If
    Else
        l_decPrevSchedHours = _
            ContinuousScheduledOvertimeHours( _
                r_rsTimeCharge!EMP_NUM, _
                r_rsTimeCharge!DATE_WORKED, _
                r_rsTimeCharge!START_TIME, _
                r_rsTimeCharge!END_TIME, _
                m_lngOTOppNum, _
                r_typCanvShift.StartDateTime, _
                r_typCanvShift.EndDateTime)
    End If

    l_decOvertimeHours = l_decPrevSchedHours
    If l_decOvertimeHours + r_rsTimeCharge!HOURS <= 8 Then
        l_strExcCd = "OT3"
        #If TRTRACE Then
            TraceActionMessage "3", "DetermineOvertimeExceptionCodeOA", _
                "ContinuousOvertime = True,(" & r_rsTimeCharge!CHARGE_DATE & "), " & _
                ", SchedHours <=8, Scheduled Hours = " & l_decPrevSchedHours & _
                ", Exception code = " & l_strExcCd
        #End If
    ElseIf l_decOvertimeHours >= 8 Then
        l_strExcCd = "OT5"
        #If TRTRACE Then
            TraceActionMessage "3", "DetermineOvertimeExceptionCodeOA", _
                ", ContinuousOvertime = True,(" & r_rsTimeCharge!CHARGE_DATE & "), " & _
                ", SchedHours >=8, Scheduled Hours = " & l_decPrevSchedHours & _
                ", Exception code = " & l_strExcCd
        #End If
    Else
        l_decSplitHours = 8 - l_decOvertimeHours
        l_varBookmark = r_rsTimeCharge.Bookmark
        l_strExcCd = "OT3"
        #If TRTRACE Then
            TraceActionMessage "3", "DetermineOvertimeExceptionCodeOA", _
                "ContinuousOvertime = True,(" & r_rsTimeCharge!CHARGE_DATE & "), " & _
                ", Scheduled Hours = " & l_decPrevSchedHours & _
                ", Split Hours = " & l_decSplitHours & _
                ", Exception code = " & l_strExcCd
        #End If
        SplitTimeCharge l_decSplitHours, r_rsTimeCharge
        r_rsTimeCharge.Bookmark = l_varBookmark
    End If
    
    
    
xit:
    DetermineOvertimeExceptionCodeOA = l_strExcCd
    m_rsEmpSchedules.Filter = adFilterNone
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "DetermineOvertimeExceptionCodeOA"
    
End Function


Private Sub DetermineTravelTimeOA( _
                    ByRef r_typCanvShift As CanvassedShift, _
                    ByRef r_rsTimeCharge As ADODB.Recordset)

On Error GoTo eh
    
    If r_rsTimeCharge!TRAVEL_PAY_HRS >= 0 Then
        GoTo xit
    End If
    
    r_rsTimeCharge!TRAVEL_PAY_HRS = _
        DetermineTravelTimeHoursOA(r_typCanvShift, r_rsTimeCharge)
    If r_rsTimeCharge!TRAVEL_PAY_HRS > 0 And Not r_typCanvShift.TravelPaid Then
        r_typCanvShift.TravelPaid = True
    End If
    
    
xit:
    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "DetermineTravelTimeOA"
    
End Sub


Private Function DetermineTravelTimeHoursOA( _
                    ByRef r_typCanvShift As CanvassedShift, _
                    ByRef r_rsTimeCharge As ADODB.Recordset) As Variant
'Answer a decimal.  Possible values are 0.0 and 0.7.

    Dim l_datPrevRegEndTime As Date
    Dim l_blnCallout As Boolean
    Dim l_lngEmpNum As Long
    Dim l_decTravelHours As Variant
    Dim l_blnTradedShift As Boolean
    Dim l_lngIntDiff As Long
    Dim l_strErrorMessage As String
    Dim l_blnConnectedToPriorSTShift As Boolean
    Dim l_blnConnectedToNextSTShift As Boolean
    Dim l_datNextRegStartTime As Date

On Error GoTo eh
     
    l_decTravelHours = CDec(0#)
    
    l_lngEmpNum = r_rsTimeCharge!EMP_NUM
    m_rsEmpSchedules.Filter = "EMP_NUM = " & l_lngEmpNum
    
    #If TRTRACE Then
        TraceActionMessage "3", "DetermineTravelTimeHoursOA", _
            "Shift Start: " & r_typCanvShift.StartDateTime & " End: " & _
            r_typCanvShift.EndDateTime & _
            vbCrLf & " Timecharge: " & FormatTimeCharge(r_rsTimeCharge), True
    #End If
    
    If Not TimeOutsideRegularScheduledHours( _
                r_typCanvShift.StartDateTime, _
                r_typCanvShift.EndDateTime) Then

        #If TRTRACE Then
            TraceActionMessage "3", "DetermineTravelTimeHoursOA", _
                m_strErrorMessage
        #End If
  
        m_strErrorMessage = "Overtime conflicts with regular scheduled shift"
        AppendProcessTimeChargeError trErrbusOTConflictRegShiftExc, m_strErrorMessage, _
            r_rsTimeCharge!EMP_NUM, r_typCanvShift, r_rsTimeCharge

        l_decTravelHours = CDec(0#)
        GoTo xit
    End If
    
    If IsCalloutOvertime(r_rsTimeCharge!EMP_NUM, r_typCanvShift) Then
        l_blnCallout = True
        
        If TimespanContinuousRegularScheduledShift( _
                r_typCanvShift.StartDateTime, r_typCanvShift.EndDateTime, 1) Then
            If r_typCanvShift.ShiftHours < 8 Then
                l_decTravelHours = _
                    TravelTimeHoursThisTimeCharge(r_typCanvShift, r_rsTimeCharge)
                #If TRTRACE Then
                    TraceActionMessage "3", "DetermineTravelTimeHoursOA", _
                    "Callout is continuous, not double-over, entitled 0.7 travel for this shift" & _
                    "Paid this time charge=" & CStr(l_decTravelHours)
                #End If
            End If
        Else
            l_decTravelHours = _
                TravelTimeHoursThisTimeCharge(r_typCanvShift, r_rsTimeCharge)
            #If TRTRACE Then
                TraceActionMessage "3", "DetermineTravelTimeHoursOA", _
                "Callout is not continuous, entitled 0.7 travel for this shift" & _
                "Paid this time charge=" & CStr(l_decTravelHours)
            #End If
        End If
        
        GoTo xit
    End If
        
'If it is not Callout then it is Prearranged
    l_blnCallout = False
'    MsgBox "Prearranged", , "jimb -" & str(l_lngEmpNum)
    
        
'Is the break between the regular shift and following prearranged overtime < 8 hours
'Do not pay travel if previous regular shift was a trade.
    If Not TimespanContinuousRegularScheduledShift( _
        r_typCanvShift.StartDateTime, r_typCanvShift.EndDateTime) Then
'        MsgBox "not TimespanContinuousRegularScheduledShift", , "jimb"
        'There is no shift before and no shift after the opportunity
        'Checking to see if less than an 8 hour gap in previous shift
        l_datPrevRegEndTime = PreviousRegularShiftEndTimeBefore( _
            l_lngEmpNum, r_typCanvShift.StartDateTime, r_typCanvShift.EndDateTime, l_blnTradedShift)
        l_lngIntDiff = DateDiff("n", l_datPrevRegEndTime, r_typCanvShift.StartDateTime)
'        MsgBox "l_lngIntDiff = " & str(l_lngIntDiff), , "jimb"
        If l_lngIntDiff < 480 And Not l_blnTradedShift Then
'        MsgBox "pay travel", , "jimb"
            l_decTravelHours = _
                TravelTimeHoursThisTimeCharge(r_typCanvShift, r_rsTimeCharge)
            #If TRTRACE Then
                TraceActionMessage "3", "DetermineTravelTimeHoursOA", _
                "Prearranged is not continuous and break < 8 hours and not trade, travel 0.7 this shift, " & _
                "Paid this time charge=" & CStr(l_decTravelHours) & _
                ".  Prior ST shift endtime = " & l_datPrevRegEndTime
            #End If
        Else
'            MsgBox "DO NOT pay travel", , "jimb"
            #If TRTRACE Then
                TraceActionMessage "3", "DetermineTravelTimeHoursOA", _
                "Prearranged is not continuous, but no travel paid this shift, " & _
                "Minutes between previous regular shift = " & CStr(l_lngIntDiff) & _
                ".  Is previous shift a trade = " & CStr(l_blnTradedShift) & "."
            #End If
        End If
        GoTo xit
    End If

'There is a ST shift contiguous either before or after the OT shift
'if the pre-arranged OT is for less than 8 hours and they have less than an 8 hour
'break on either side, then the get travel hours
    If r_typCanvShift.ShiftHours >= 8 Then
        #If TRTRACE Then
            TraceActionMessage "3", "DetermineTravelTimeHoursOA", _
            "Prearranged is continuous and for a full shift, so no travel paid this shift,"
        #End If
        GoTo xit
    End If
        
    l_blnConnectedToNextSTShift = TimespanContinuousRegularScheduledShift( _
        r_typCanvShift.StartDateTime, r_typCanvShift.EndDateTime, 1)
    l_blnConnectedToPriorSTShift = TimespanContinuousRegularScheduledShift( _
        r_typCanvShift.StartDateTime, r_typCanvShift.EndDateTime, 2)
    
'    If l_blnConnectedToNextSTShift Then
'        MsgBox "ConnectedToNextSTShift", , "jimb -" & str(l_lngEmpNum)
'    End If
    
'    If l_blnConnectedToPriorSTShift Then
'        MsgBox "ConnectedToPriorSTShift", , "jimb -" & str(l_lngEmpNum)
'    End If
     
    If l_blnConnectedToNextSTShift And l_blnConnectedToPriorSTShift Then
    'should not happen, but will not give travel pay
        #If TRTRACE Then
            TraceActionMessage "3", "DetermineTravelTimeHoursOA", _
            "Prearranged is continuous to 2 ST shifts, so no travel paid this shift,"
        #End If
        GoTo xit
    End If
    If l_blnConnectedToNextSTShift Then
        l_datPrevRegEndTime = PreviousRegularShiftEndTimeBefore( _
            l_lngEmpNum, r_typCanvShift.StartDateTime, _
            r_typCanvShift.EndDateTime, l_blnTradedShift)
            l_lngIntDiff = DateDiff("n", l_datPrevRegEndTime, r_typCanvShift.StartDateTime)
        If l_lngIntDiff < 480 And Not l_blnTradedShift Then
            l_decTravelHours = _
                TravelTimeHoursThisTimeCharge(r_typCanvShift, r_rsTimeCharge)
            #If TRTRACE Then
                TraceActionMessage "3", "DetermineTravelTimeHoursOA", _
                "Prearranged shift is for less than 8 hours and is continuous to the next shift.  " & _
                "The break after the prior shift is < 8 hours and not trade, travel 0.7 this shift, " & _
                "Paid this time charge=" & CStr(l_decTravelHours) & _
                ".  Prior ST shift endtime = " & l_datPrevRegEndTime
            #End If
        Else
            #If TRTRACE Then
                TraceActionMessage "3", "DetermineTravelTimeHoursOA", _
                "Prearranged shift is for less than 8 hours and is continuous to the next shift.  " & _
                "The break after the prior shift is >= 8 hours or was a trade, so no travel paid this shift, " & _
                "Minutes between previous regular shift = " & CStr(l_lngIntDiff) & _
                ".  Is previous shift a trade = " & CStr(l_blnTradedShift) & "."
            #End If
        End If
        GoTo xit
    End If
'''''    l_datNextRegStartTime = NextRegularShiftStartTimeAfter( _
'''''    l_lngEmpNum, r_typCanvShift.StartDateTime, _
'''''    r_typCanvShift.EndDateTime, l_blnTradedShift)
''''''    MsgBox "l_datNextRegStartTime = " & Format(l_datNextRegStartTime, "mm/dd/yyyy HH:MM"), , _
'''''    "r_typCanvShift.EndDateTime = " & Format(r_typCanvShift.EndDateTime, "mm/dd/yyyy HH:MM")
'''''    l_lngIntDiff = DateDiff("n", r_typCanvShift.EndDateTime, l_datNextRegStartTime)
''''''    MsgBox "l_lngIntDiff = " & str(l_lngIntDiff), , "JIMB"
'''''    If l_lngIntDiff < 480 And Not l_blnTradedShift Then
'''''        l_decTravelHours = _
'''''            TravelTimeHoursThisTimeCharge(r_typCanvShift, r_rsTimeCharge)
'''''        #If TRTRACE Then
'''''            TraceActionMessage "3", "DetermineTravelTimeHoursOA", _
'''''            "Prearranged shift is for less than 8 hours and is continuous to the next shift.  " & _
'''''            "The break after the prior shift is < 8 hours and not trade, travel 0.7 this shift, " & _
'''''            "Paid this time charge=" & CStr(l_decTravelHours) & _
'''''            ".  Prior ST shift endtime = " & l_datPrevRegEndTime
'''''        #End If
'''''    Else
'''''        #If TRTRACE Then
'''''            TraceActionMessage "3", "DetermineTravelTimeHoursOA", _
'''''            "Prearranged shift is for less than 8 hours and is continuous to the next shift.  " & _
'''''            "The break after the prior shift is >= 8 hours or was a trade, so no travel paid this shift, " & _
'''''            "Minutes between previous regular shift = " & CStr(l_lngIntDiff) & _
'''''            ".  Is previous shift a trade = " & CStr(l_blnTradedShift) & "."
'''''        #End If
'''''    End If
    
    
xit:
    DetermineTravelTimeHoursOA = l_decTravelHours
    m_rsEmpSchedules.Filter = adFilterNone
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "DetermineTravelTimeHoursOA"
    
End Function


Private Function TravelTimeHoursThisTimeCharge( _
                ByRef r_typCanvShift As CanvassedShift, _
                ByRef r_rsTimeCharge As ADODB.Recordset) As Variant
'Only pay travel time if not already paid for this shift.
    Dim l_decTravelHours As Variant

On Error GoTo eh

    If r_typCanvShift.TravelPaid Then
        l_decTravelHours = CDec(0#)
    Else
        l_decTravelHours = CDec(0.7)
    End If

    TravelTimeHoursThisTimeCharge = l_decTravelHours

    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "TravelTimeHoursThisTimeCharge"
    
End Function


Private Sub SplitTimeCharge( _
                ByVal v_decSplitHours As Variant, _
                ByRef r_rsTimeCharge As ADODB.Recordset)

    Dim l_rsTimeChargeC As ADODB.Recordset
    Dim l_strStartTime As String
    Dim l_strNewStartTime As String
    Dim l_strEndTime As String
    Dim l_strNewEndTime As String
    Dim l_decHours As Variant
    Dim l_decNewHours As Variant

On Error GoTo eh

    l_strStartTime = r_rsTimeCharge!START_TIME
    l_strEndTime = AddMinutes(l_strStartTime, v_decSplitHours * 60)
    l_decHours = HoursDiff(l_strStartTime, l_strEndTime)
    l_strNewStartTime = l_strEndTime
    l_decNewHours = HoursDiff(l_strNewStartTime, r_rsTimeCharge!END_TIME)
    l_strNewEndTime = r_rsTimeCharge!END_TIME
    
    r_rsTimeCharge!END_TIME = l_strEndTime
    r_rsTimeCharge!HOURS = l_decHours
    
    If IsOvertimeWorked(r_rsTimeCharge!EXCEPTION_CD, r_rsTimeCharge!OVER_EXC_CD) Or _
            Trim(r_rsTimeCharge!EXCEPTION_CD) = "" Then
        UpdateSchedulesAfterSplit r_rsTimeCharge
    End If
    
    Set l_rsTimeChargeC = r_rsTimeCharge.Clone
    l_rsTimeChargeC.Bookmark = r_rsTimeCharge.Bookmark
    
    r_rsTimeCharge.AddNew
    r_rsTimeCharge!OT_OPP_NUM = l_rsTimeChargeC!OT_OPP_NUM
    r_rsTimeCharge!EMP_NUM = l_rsTimeChargeC!EMP_NUM
    r_rsTimeCharge!CHARGE_DATE = m_rsOTOpportunity!CHARGE_DATE
    r_rsTimeCharge!DATE_WORKED = l_rsTimeChargeC!DATE_WORKED
    r_rsTimeCharge!START_TIME = l_strNewStartTime
    r_rsTimeCharge!END_TIME = l_strNewEndTime
    r_rsTimeCharge!HOURS = l_decNewHours
    r_rsTimeCharge!EXCEPTION_CD = ""
    r_rsTimeCharge!OVER_EXC_CD = l_rsTimeChargeC!OVER_EXC_CD
    r_rsTimeCharge!JOB_NUM = l_rsTimeChargeC!JOB_NUM
    r_rsTimeCharge!COMP_SHIFT_CD = l_rsTimeChargeC!COMP_SHIFT_CD
    r_rsTimeCharge!OT_REASON_CD = l_rsTimeChargeC!OT_REASON_CD
    r_rsTimeCharge!WG_NUM = l_rsTimeChargeC!WG_NUM
    r_rsTimeCharge!DISTRIBUTION_ID = l_rsTimeChargeC!DISTRIBUTION_ID
    r_rsTimeCharge!DISTRIBUTION_TYPE = l_rsTimeChargeC!DISTRIBUTION_TYPE
    r_rsTimeCharge!JOB_NBR = l_rsTimeChargeC!JOB_NBR
    r_rsTimeCharge!JOBTASK_NBR = l_rsTimeChargeC!JOBTASK_NBR
    r_rsTimeCharge!JOBTASK_STEP_NBR = l_rsTimeChargeC!JOBTASK_STEP_NBR
    r_rsTimeCharge!EMPRV_DATABASE_ID = l_rsTimeChargeC!EMPRV_DATABASE_ID
    r_rsTimeCharge!TRAVEL_PAY_HRS = l_rsTimeChargeC!TRAVEL_PAY_HRS
    r_rsTimeCharge!CREATED_EMP_NUM = l_rsTimeChargeC!CREATED_EMP_NUM
    r_rsTimeCharge!LAST_MOD_EMP_NUM = l_rsTimeChargeC!LAST_MOD_EMP_NUM

    If IsOvertimeWorked(r_rsTimeCharge!EXCEPTION_CD, r_rsTimeCharge!OVER_EXC_CD) Or _
            Trim(r_rsTimeCharge!EXCEPTION_CD) = "" Then
        AddOvertimeSchedule r_rsTimeCharge
        AddAdjustedScheduleOvertime r_rsTimeCharge
    End If

    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "SplitTimeCharge"

End Sub


Private Function IsCalloutOvertime(ByVal v_lngEmpNum, ByRef r_typCanvShift As CanvassedShift) As Boolean
'Callout is defined as asking someone, outside of schedule, to report to work
' in less than 8 hours.  An opportunity maybe designated as callout, meaning that
' the overtime is to be worked in less than 8 hours, but an employee maybe canvassed
' while on duty or at home.  If called at home, it is callout.  If canvassed
' while on duty, it is prearranged.

On Error GoTo eh

    If IsEmployeeOnDutyPriorToOvertimeShift(v_lngEmpNum, r_typCanvShift) Then
        IsCalloutOvertime = False
    ElseIf IsEmployeeCalledToReportWithin8Hours(v_lngEmpNum, r_typCanvShift) Then
        IsCalloutOvertime = True
    Else
        IsCalloutOvertime = False
    End If
        
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "IsCalloutOvertime"
    
End Function


Private Function IsDayOff(ByVal v_datChargeDate As Date) As Boolean

    IsDayOff = Not IsScheduledWorkDay(v_datChargeDate)

End Function


Private Function IsScheduledWorkDay(ByVal v_datChargeDate As Date) As Boolean
'There may be upto 2 shifts per charge date.  Answer True if any shift found
'has a start time.

On Error GoTo eh

    IsScheduledWorkDay = False
    
    m_rsEmpSchedules.MoveFirst
    Do Until m_rsEmpSchedules.EOF
        If m_rsEmpSchedules!CHARGE_DATE = v_datChargeDate Then
            If Trim(m_rsEmpSchedules!SHIFT_START_TIME) <> "" And _
                    Not IsHolidayRest(m_rsEmpSchedules!EXCEPTION_CD) Then
                IsScheduledWorkDay = True
                Exit Do
            End If
        End If
        m_rsEmpSchedules.MoveNext
    Loop

    m_rsEmpSchedules.MoveFirst

    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "IsScheduledWorkDay"
    
End Function


Private Function IsHolidayOff(ByVal v_datChargeDate As Date) As Boolean
'There may be upto 2 shifts per charge date.  Answer True if any shift found
'has a holiday at rest.

On Error GoTo eh

    IsHolidayOff = False
    m_rsEmpSchedules.MoveFirst
    Do Until m_rsEmpSchedules.EOF
        If m_rsEmpSchedules!CHARGE_DATE = v_datChargeDate Then
            If IsHolidayRest(m_rsEmpSchedules!EXCEPTION_CD) Then
                IsHolidayOff = True
                Exit Do
            End If
        End If
        m_rsEmpSchedules.MoveNext
    Loop

    m_rsEmpSchedules.MoveFirst

    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "IsHolidayOff"
    
End Function


Private Function IsHolidayScheduled(ByVal v_datChargeDate As Date) As Boolean
'There may be upto 2 shifts per charge date.  Answer True if any shift found
'has a holiday scheduled to work.

On Error GoTo eh

    IsHolidayScheduled = False
    m_rsEmpSchedules.MoveFirst
    Do Until m_rsEmpSchedules.EOF
        If m_rsEmpSchedules!CHARGE_DATE = v_datChargeDate Then
            If IsHolidayWork(m_rsEmpSchedules!EXCEPTION_CD) Then
                IsHolidayScheduled = True
                Exit Do
            End If
        End If
        m_rsEmpSchedules.MoveNext
    Loop

    m_rsEmpSchedules.MoveFirst

    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "IsHolidayScheduled"
    
End Function


'
'
'Private Function IsHoliday(ByVal v_datChargeDate As Date) As Boolean
''There may be upto 2 shifts per charge date.  Answer True if any shift found
''has a start time.
'
'On Error GoTo eh
'
'    IsHoliday = False
'    m_rsEmpSchedules.MoveFirst
'    Do Until m_rsEmpSchedules.EOF
'        If m_rsEmpSchedules!CHARGE_DATE = v_datChargeDate Then
'            If IsHolidayOff(m_rsEmpSchedules!EXCEPTION_CD) Then
'                IsHoliday = True
'                Exit Do
'            End If
'        End If
'        m_rsEmpSchedules.MoveNext
'    Loop
'
'    m_rsEmpSchedules.MoveFirst
'
'    Exit Function
'
'eh:
'    TrsRaiseError m_strCMODULENAME, "IsHoliday"
'
'End Function


Private Function IsSystemGeneratedRefusal( _
                    ByVal v_strExc_Cd As String, _
                    ByVal v_lngEmpNum As Long) As Boolean

On Error GoTo eh

    IsSystemGeneratedRefusal = UCase(Left(v_strExc_Cd, 2)) = "RF" And _
                                v_lngEmpNum = 999999999
    
    Exit Function

eh:
        TrsRaiseError m_strCMODULENAME, "IsSystemGeneratedRefusal"

End Function


Private Function IsOvertimeRefused(ByVal v_strExc_Cd As String) As Boolean

On Error GoTo eh

    IsOvertimeRefused = GetWorkRestInd(v_strExc_Cd) = "R"
    
    Exit Function

eh:
        TrsRaiseError m_strCMODULENAME, "IsOvertimeRefused"

End Function


Private Function IsOvertimeWorked( _
            ByVal v_strExc_Cd As String, ByVal v_strOverExc_Cd As String) As Boolean


On Error GoTo eh

    If IsOverrideThatAllowsOverlap(v_strOverExc_Cd) Then
        IsOvertimeWorked = False
    Else
        IsOvertimeWorked = GetWorkRestInd(v_strExc_Cd) = "W"
    End If

    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "IsOvertimeWorked"

End Function


Private Function IsOvertimePaid(ByVal v_strExc_Cd As String) As Boolean

On Error GoTo eh

    IsOvertimePaid = GetWorkRestInd(v_strExc_Cd) = "W"

    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "IsOvertimePaid"

End Function


Private Function IsHolidayRest(ByVal v_strExc_Cd As String) As Boolean

On Error GoTo eh

    IsHolidayRest = GetWorkRestInd(v_strExc_Cd, "H") = "R"
    
    Exit Function

eh:
        TrsRaiseError m_strCMODULENAME, "IsHolidayRest"

End Function


Private Function IsHolidayWork(ByVal v_strExc_Cd As String) As Boolean

On Error GoTo eh

    If Trim(v_strExc_Cd) = "" Then
        IsHolidayWork = False
    Else
        IsHolidayWork = GetWorkRestInd(v_strExc_Cd, "H") = "W"
    End If
    
    Exit Function

eh:
        TrsRaiseError m_strCMODULENAME, "IsHolidayWork"

End Function


Private Function GetWorkRestInd(ByVal v_strExc_Cd As String, _
                                Optional v_strGroupInd As String) As String

On Error GoTo eh

    If v_strGroupInd = "" Then
        m_rsAllExceptions.MoveFirst
    Else
        m_rsAllExceptions.Filter = "GROUP_IND = '" & Trim(v_strGroupInd) & "'"
    End If
    
    m_rsAllExceptions.Find "EXCEPTION_CD = '" & Trim(v_strExc_Cd) & "'"
    If m_rsAllExceptions.EOF Then
        GetWorkRestInd = ""
    Else
        GetWorkRestInd = Trim(m_rsAllExceptions!WORK_REST_IND)
    End If

    m_rsAllExceptions.Filter = adFilterNone
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "GetWorkRestInd"

End Function


Private Function GetSymbolCode(ByVal v_strExc_Cd As String) As String

On Error GoTo eh

    m_rsAllExceptions.MoveFirst
    m_rsAllExceptions.Find "EXCEPTION_CD = '" & Trim(v_strExc_Cd) & "'"
    If m_rsAllExceptions.EOF Then
        GetSymbolCode = ""
    Else
        GetSymbolCode = Trim(m_rsAllExceptions!SYMBOL_CD)
    End If

    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "GetSymbolCode"

End Function


Private Function IsOvertimeWorkedResponse(ByVal v_strRespCd As String) As Boolean

On Error GoTo eh

    IsOvertimeWorkedResponse = UCase(GetOvertimeWorkedInd(v_strRespCd)) = "Y"

    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "IsOvertimeWorkedResponse"

End Function


Private Function GetOvertimeWorkedInd(ByVal v_strRespCd As String) As String

On Error GoTo eh

    m_rsOvertimeResponses.MoveFirst
    
    m_rsOvertimeResponses.Find "OT_RESP_CD = '" & v_strRespCd & "'"
    If m_rsOvertimeResponses.EOF Then
        GetOvertimeWorkedInd = ""
    Else
        GetOvertimeWorkedInd = Trim(m_rsOvertimeResponses!OT_WORKED_IND)
    End If

    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "GetOvertimeWorkedInd"

End Function

Private Function TimeOutsideRegularScheduledHours( _
                    ByVal v_datStartDateTime As Date, _
                    ByVal v_datEndDateTime As Date) As Boolean

    Dim l_datDateLunchStart As Date
    Dim l_datDateLunchEnd As Date
    Dim l_datShiftStart As Date
    Dim l_datLunchStart As Date
    Dim l_datLunchEnd As Date
    Dim l_datShiftEnd As Date

On Error GoTo eh

    TimeOutsideRegularScheduledHours = True
    
    m_rsEmpSchedules.MoveFirst
    Do Until m_rsEmpSchedules.EOF
        If IsRegularScheduledShift(m_rsEmpSchedules) Then
            If IsScheduledLunch(m_rsEmpSchedules) Then
                l_datShiftStart = DateTimeFrom(m_rsEmpSchedules!SHIFT_START_DATE, m_rsEmpSchedules!SHIFT_START_TIME)
                l_datShiftEnd = DateTimeFrom(m_rsEmpSchedules!SHIFT_END_DATE, m_rsEmpSchedules!SHIFT_END_TIME)
                If m_rsEmpSchedules!SHIFT_START_TIME > m_rsEmpSchedules!SHIFT_LUNCH_START_TIME Then
                    l_datDateLunchStart = DateAdd("d", 1, m_rsEmpSchedules!SHIFT_START_DATE)
                Else
                    l_datDateLunchStart = m_rsEmpSchedules!SHIFT_START_DATE
                End If
                If m_rsEmpSchedules!SHIFT_START_TIME > m_rsEmpSchedules!SHIFT_LUNCH_END_TIME Then
                    l_datDateLunchEnd = DateAdd("d", 1, m_rsEmpSchedules!SHIFT_START_DATE)
                Else
                    l_datDateLunchEnd = m_rsEmpSchedules!SHIFT_START_DATE
                End If
                l_datLunchStart = DateTimeFrom(l_datDateLunchStart, m_rsEmpSchedules!SHIFT_LUNCH_START_TIME, True)
                l_datLunchEnd = DateTimeFrom(l_datDateLunchEnd, m_rsEmpSchedules!SHIFT_LUNCH_END_TIME)
            
                If TimePeriodsIntersect2( _
                        v_datStartDateTime, v_datEndDateTime, _
                        l_datShiftStart, l_datLunchStart) Or _
                    TimePeriodsIntersect2( _
                        v_datStartDateTime, v_datEndDateTime, _
                        l_datLunchEnd, l_datShiftEnd) Then
                    TimeOutsideRegularScheduledHours = False
                    Exit Do
                End If
            Else
                If TimePeriodsIntersect2( _
                        v_datStartDateTime, v_datEndDateTime, _
                        DateTimeFrom(m_rsEmpSchedules!SHIFT_START_DATE, m_rsEmpSchedules!SHIFT_START_TIME), _
                        DateTimeFrom(m_rsEmpSchedules!SHIFT_END_DATE, m_rsEmpSchedules!SHIFT_END_TIME)) Then
                    TimeOutsideRegularScheduledHours = False
                    Exit Do
                End If
            End If
        End If
        m_rsEmpSchedules.MoveNext
    Loop

    m_rsEmpSchedules.MoveFirst
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "TimeOutsideRegularScheduledHours"

End Function


Private Function IsScheduledLunch(ByVal v_rsEmpSched As ADODB.Recordset) As Boolean
    
    IsScheduledLunch = _
        Trim(v_rsEmpSched!SHIFT_LUNCH_START_TIME) <> "" And _
        Trim(v_rsEmpSched!SHIFT_LUNCH_END_TIME) <> ""
        
End Function


Private Function TimeContinuousRegularScheduledShift( _
                    ByVal v_datDateWorked As Date, _
                    ByVal v_strStartTime As String, _
                    ByVal v_rsTimeCharge As ADODB.Recordset) _
                As Boolean
'This time charge may be a component of a composite shift.  Determine so-called
'shift, i.e. group of continuous time charges and check it for continuous with
'regular scheduled shift.  Up to a 60 minute break is allowed and still considered
'continuous.
    Dim l_datShiftStart As Date
    Dim l_datShiftEnd As Date
    Dim l_rsTimeChargeC As ADODB.Recordset

On Error GoTo eh

    TimeContinuousRegularScheduledShift = False
    
    Set l_rsTimeChargeC = v_rsTimeCharge.Clone
'    l_rsTimeChargeC.Filter = "EMP_NUM = " & v_lngEmpNum
    PositionOnTimeCharge l_rsTimeChargeC, v_datDateWorked, v_strStartTime
    DetermineContinuousStartEndTimes l_rsTimeChargeC, l_datShiftStart, l_datShiftEnd
    Set l_rsTimeChargeC = Nothing
    
'    this depends on all timecharges added to timecharge recordset.
'    problem if determining rate as added to time charge because it will not
'    know if preceding a reg shift.
'    Must know the shift to determine rate.
    
    m_rsEmpSchedules.MoveFirst
    Do Until m_rsEmpSchedules.EOF
        If m_rsEmpSchedules!SHIFT_HOURS > 0 Then
            If TimespanContinous( _
                    l_datShiftStart, l_datShiftEnd, _
                    DateTimeFrom(m_rsEmpSchedules!SHIFT_START_DATE, m_rsEmpSchedules!SHIFT_START_TIME), _
                    DateTimeFrom(m_rsEmpSchedules!SHIFT_END_DATE, m_rsEmpSchedules!SHIFT_END_TIME), _
                    m_lngMINOFFALLOWCONTINUOUS) Then
                TimeContinuousRegularScheduledShift = True
                Exit Do
            End If
        End If
        m_rsEmpSchedules.MoveNext
    Loop

    m_rsEmpSchedules.MoveFirst
    Exit Function

eh:
    Set l_rsTimeChargeC = Nothing
    TrsRaiseError m_strCMODULENAME, "TimeContinuousRegularScheduledShift"
    
End Function


Private Function TimespanContinuousRegularScheduledShift( _
                    ByVal v_datStartDateTime As Date, _
                    ByVal v_datEndDateTime As Date, _
                    Optional ByVal v_lngPrePostInd As Long = 0, _
                    Optional ByVal v_lngMinBreak As Long = m_lngMINOFFALLOWCONTINUOUS) _
                As Boolean

    Dim l_datDateLunchStart As Date
    Dim l_datDateLunchEnd As Date
    Dim l_datLunchStart As Date
    Dim l_datLunchEnd As Date

On Error GoTo eh

    TimespanContinuousRegularScheduledShift = False
        
    m_rsEmpSchedules.MoveFirst
    Do Until m_rsEmpSchedules.EOF
        If IsRegularScheduledShift(m_rsEmpSchedules) Then
            If TimespanContinous( _
                    v_datStartDateTime, v_datEndDateTime, _
                    DateTimeFrom(m_rsEmpSchedules!SHIFT_START_DATE, m_rsEmpSchedules!SHIFT_START_TIME), _
                    DateTimeFrom(m_rsEmpSchedules!SHIFT_END_DATE, m_rsEmpSchedules!SHIFT_END_TIME), _
                    v_lngMinBreak) Then
                If v_lngPrePostInd = 0 Then         'pre or post shift
                    TimespanContinuousRegularScheduledShift = True
                    Exit Do
                ElseIf v_lngPrePostInd = 1 Then      'preshift only
                    If DatetimeCompare(v_datStartDateTime, "LT", DateTimeFrom(m_rsEmpSchedules!SHIFT_START_DATE, m_rsEmpSchedules!SHIFT_START_TIME)) Then
'                    If v_datStartDateTime < DateTimeFrom(m_rsEmpSchedules!SHIFT_START_DATE, m_rsEmpSchedules!SHIFT_START_TIME) Then
                        TimespanContinuousRegularScheduledShift = True
                        Exit Do
                    End If
                ElseIf v_lngPrePostInd = 2 Then      'postshift only
                    If DatetimeCompare(v_datStartDateTime, "GT", DateTimeFrom(m_rsEmpSchedules!SHIFT_START_DATE, m_rsEmpSchedules!SHIFT_START_TIME)) Then
'                    If v_datStartDateTime > DateTimeFrom(m_rsEmpSchedules!SHIFT_START_DATE, m_rsEmpSchedules!SHIFT_START_TIME) Then
                        TimespanContinuousRegularScheduledShift = True
                        Exit Do
                    End If
                Else
                    TimespanContinuousRegularScheduledShift = True
                    Exit Do
                End If
                
            Else
                If IsScheduledLunch(m_rsEmpSchedules) Then
                    If m_rsEmpSchedules!SHIFT_START_TIME > m_rsEmpSchedules!SHIFT_LUNCH_START_TIME Then
                        l_datDateLunchStart = DateAdd("d", 1, m_rsEmpSchedules!SHIFT_START_DATE)
                    Else
                        l_datDateLunchStart = m_rsEmpSchedules!SHIFT_START_DATE
                    End If
                    If m_rsEmpSchedules!SHIFT_START_TIME > m_rsEmpSchedules!SHIFT_LUNCH_END_TIME Then
                        l_datDateLunchEnd = DateAdd("d", 1, m_rsEmpSchedules!SHIFT_START_DATE)
                    Else
                        l_datDateLunchEnd = m_rsEmpSchedules!SHIFT_START_DATE
                    End If
                    l_datLunchStart = DateTimeFrom(l_datDateLunchStart, m_rsEmpSchedules!SHIFT_LUNCH_START_TIME, True)
                    l_datLunchEnd = DateTimeFrom(l_datDateLunchEnd, m_rsEmpSchedules!SHIFT_LUNCH_END_TIME)
                    If TimePeriodsIntersect2(v_datStartDateTime, v_datEndDateTime, _
                            l_datLunchStart, l_datLunchEnd) Then
                        TimespanContinuousRegularScheduledShift = True
                        Exit Do
                    End If
                End If
            End If
        End If
        m_rsEmpSchedules.MoveNext
    Loop

    m_rsEmpSchedules.MoveFirst
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "TimespanContinuousRegularScheduledShift"
    
End Function


Private Function TimespanContinuousOvertimeShift( _
                    ByVal v_datStartDateTime As Date, _
                    ByVal v_datEndDateTime As Date, _
                    ByVal v_lngThisOTOppNum As Long, _
                    ByRef r_lngOtherOTOppNum As Long, _
                    Optional ByVal v_lngPrePostInd As Long = 0, _
                    Optional ByVal v_lngMinBreak As Long = m_lngMINOFFALLOWCONTINUOUS) _
                As Boolean

' Answer the overtime opportunity number if this timespan is continuous
' with another overtime shift; otherwise zero.
'
' v_lngPrePostInd values (0,1,2)
' 0 = Pre or Post shift
' 1 = Preshift - Is the input timespan before continuous schedule found
' 2 = Postshift - Is the input timespan after continuous schedule found
' When navigating backward from a given timespan, specify 2.
' When navigating forward specify 1.

    Dim l_datSchedStartDateTime As Date
    Dim l_datSchedEndDateTime As Date

On Error GoTo eh

    TimespanContinuousOvertimeShift = False
    r_lngOtherOTOppNum = 0

    m_rsEmpSchedules.MoveFirst
    Do Until m_rsEmpSchedules.EOF
        If IsScheduledOvertime(m_rsEmpSchedules) Then
            If m_rsEmpSchedules!OT_OPP_NUM <> v_lngThisOTOppNum Then
                l_datSchedStartDateTime = DateTimeFrom(m_rsEmpSchedules!DATE_WORKED, m_rsEmpSchedules!START_TIME)
                l_datSchedEndDateTime = DateTimeFrom(m_rsEmpSchedules!DATE_WORKED, m_rsEmpSchedules!END_TIME, True)
                If TimespanContinous( _
                        v_datStartDateTime, v_datEndDateTime, _
                        l_datSchedStartDateTime, l_datSchedEndDateTime, v_lngMinBreak) Then
                    If v_lngPrePostInd = 0 Then         'pre or post shift
                        TimespanContinuousOvertimeShift = True
                        r_lngOtherOTOppNum = m_rsEmpSchedules!OT_OPP_NUM
                        Exit Do
                    ElseIf v_lngPrePostInd = 1 Then      'preshift only
                        If DatetimeCompare(v_datStartDateTime, "LT", l_datSchedStartDateTime) Then
'                        If v_datStartDateTime < l_datSchedStartDateTime Then
                            TimespanContinuousOvertimeShift = True
                            r_lngOtherOTOppNum = m_rsEmpSchedules!OT_OPP_NUM
                            Exit Do
                        End If
                    ElseIf v_lngPrePostInd = 2 Then      'postshift only
                        If DatetimeCompare(v_datStartDateTime, "GT", l_datSchedStartDateTime) Then
'                        If v_datStartDateTime > l_datSchedStartDateTime Then
                            TimespanContinuousOvertimeShift = True
                            r_lngOtherOTOppNum = m_rsEmpSchedules!OT_OPP_NUM
                            Exit Do
                        End If
                    Else
                        TimespanContinuousOvertimeShift = True
                        r_lngOtherOTOppNum = m_rsEmpSchedules!OT_OPP_NUM
                        Exit Do
                    End If
                End If
            End If
        End If
        m_rsEmpSchedules.MoveNext
    Loop

    m_rsEmpSchedules.MoveFirst
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "TimespanContinuousOvertimeShift"
    
End Function


Private Function TimespanIsPaidLunchOvertimeShift( _
                    ByVal v_lngEmpNum As Long, _
                    ByVal v_datStartTime As Date, _
                    ByVal v_datEndTime As Date, _
                    ByVal v_lngThisOTOppNum As Long, _
                    ByRef r_lngOtherOTOppNum As Long) _
                As Boolean
' Answer True if this timespan is the lunch period of another overtime
' opportunity.

    Dim l_rsEmpSchedC As ADODB.Recordset
    Dim l_datLunchStartTime As Date
    Dim l_datLunchEndTime As Date
    Dim l_lngOppNum As Long

On Error GoTo eh

    TimespanIsPaidLunchOvertimeShift = False
    r_lngOtherOTOppNum = 0

    Set l_rsEmpSchedC = m_rsEmpSchedules.Clone
    l_rsEmpSchedC.Filter = "EMP_NUM = " & v_lngEmpNum
    l_rsEmpSchedC.Sort = "DATE_WORKED ASC, START_TIME ASC"

    If Not RecordsExist(l_rsEmpSchedC) Then
        GoTo xit
    End If

    Do Until l_rsEmpSchedC.EOF
        If IsScheduledOvertime(l_rsEmpSchedC) Then
            If l_rsEmpSchedC!OT_OPP_NUM <> v_lngThisOTOppNum And _
                    l_rsEmpSchedC!OT_OPP_NUM <> l_lngOppNum Then
                l_lngOppNum = l_rsEmpSchedC!OT_OPP_NUM  'only check once per opp
                If EmployeeScheduledLunchStartEndDateTimes( _
                    v_lngEmpNum, l_rsEmpSchedC!OT_OPP_NUM, _
                    l_datLunchStartTime, l_datLunchEndTime) Then
                    If DatetimeCompare(v_datStartTime, "EQ", l_datLunchStartTime) And _
                            DatetimeCompare(v_datEndTime, "EQ", l_datLunchEndTime) Then
                        r_lngOtherOTOppNum = l_rsEmpSchedC!OT_OPP_NUM
                        TimespanIsPaidLunchOvertimeShift = True
                        Exit Do
                    End If
                End If
            End If
        End If
        l_rsEmpSchedC.MoveNext
    Loop


xit:
    Set l_rsEmpSchedC = Nothing
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "TimespanIsPaidLunchOvertimeShift"
    
End Function


Private Function TimespanContinuousScheduledShift( _
                    ByVal v_datStartDateTime As Date, _
                    ByVal v_datEndDateTime As Date, _
                    ByRef r_lngOTOppNum As Long, _
                    Optional ByVal v_lngRegOTInd As Long = 1, _
                    Optional ByVal v_lngPrePostInd As Long = 0, _
                    Optional ByVal v_lngMinBreak As Long = m_lngMINOFFALLOWCONTINUOUS) _
                As Boolean

' v_lngRegOTInd values (0,1,2) 0 = Regular or Overtime; 1 = Regular; 2 = Overtime
' v_lngPrePostInd values (0,1,2) 0 = Pre or Post shift; 1 = Preshift; 2 = Postshift

On Error GoTo eh

    TimespanContinuousScheduledShift = False
    
    If RecordsExist(m_rsEmpSchedules) Then
        m_rsEmpSchedules.MoveFirst
    Else
        Exit Function
    End If
    
    Do Until m_rsEmpSchedules.EOF
        If IsShiftSelected(v_lngRegOTInd, m_rsEmpSchedules) Then
            If TimespanContinous( _
                    v_datStartDateTime, v_datEndDateTime, _
                    DateTimeFrom(m_rsEmpSchedules!SHIFT_START_DATE, m_rsEmpSchedules!SHIFT_START_TIME), _
                    DateTimeFrom(m_rsEmpSchedules!SHIFT_END_DATE, m_rsEmpSchedules!SHIFT_END_TIME), _
                    v_lngMinBreak) Then
                r_lngOTOppNum = m_rsEmpSchedules!OT_OPP_NUM
                If v_lngPrePostInd = 0 Then         'pre or post shift
                    TimespanContinuousScheduledShift = True
                    Exit Do
                ElseIf v_lngPrePostInd = 1 Then      'preshift only
                    If DatetimeCompare(v_datStartDateTime, "LT", DateTimeFrom(m_rsEmpSchedules!SHIFT_START_DATE, m_rsEmpSchedules!SHIFT_START_TIME)) Then
'                    If v_datStartDateTime < DateTimeFrom(m_rsEmpSchedules!SHIFT_START_DATE, m_rsEmpSchedules!SHIFT_START_TIME) Then
                        TimespanContinuousScheduledShift = True
                        Exit Do
                    End If
                ElseIf v_lngPrePostInd = 2 Then      'postshift only
                    If DatetimeCompare(v_datStartDateTime, "GT", DateTimeFrom(m_rsEmpSchedules!SHIFT_START_DATE, m_rsEmpSchedules!SHIFT_START_TIME)) Then
'                    If v_datStartDateTime > DateTimeFrom(m_rsEmpSchedules!SHIFT_START_DATE, m_rsEmpSchedules!SHIFT_START_TIME) Then
                        TimespanContinuousScheduledShift = True
                        Exit Do
                    End If
                Else
                    TimespanContinuousScheduledShift = True
                    Exit Do
                End If
            End If
        End If
        m_rsEmpSchedules.MoveNext
    Loop

    m_rsEmpSchedules.MoveFirst
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "TimespanContinuousScheduledShift"
    
End Function


Private Function IsShiftSelected( _
                    ByVal v_lngShiftInd As Long, _
                    ByVal v_rsEmpSched As ADODB.Recordset) As Boolean

    Select Case v_lngShiftInd
        Case 0
            IsShiftSelected = IsRegularScheduledShift(v_rsEmpSched) Or _
                                IsScheduledOvertime(v_rsEmpSched)
        Case 1
            IsShiftSelected = IsRegularScheduledShift(v_rsEmpSched)
        Case 2
            IsShiftSelected = IsScheduledOvertime(v_rsEmpSched)
    End Select


End Function



Private Function PreviousRegularShiftEndTimeBefore( _
                    ByVal v_lngEmpNum As Long, _
                    ByVal v_datStartDateTime As Date, _
                    ByVal v_datEndDateTime As Date, _
                    ByRef r_blnIsTradedShift As Boolean) _
                As Date
'If regular shift not found, just answer 1 day prior to v_datStartDateTime.

    Dim l_rsEmpSchedC As ADODB.Recordset

On Error GoTo eh

    r_blnIsTradedShift = False
    
    Set l_rsEmpSchedC = m_rsEmpSchedules.Clone
    l_rsEmpSchedC.Filter = "EMP_NUM = " & v_lngEmpNum
    l_rsEmpSchedC.Sort = "DATE_WORKED DESC, START_TIME DESC"

    PositionRegularScheduleBeforeDate v_datStartDateTime, l_rsEmpSchedC

    If l_rsEmpSchedC.EOF Then
        PreviousRegularShiftEndTimeBefore = DateAdd("d", -1, v_datStartDateTime)
    Else
        PreviousRegularShiftEndTimeBefore = _
            DateTimeFrom(l_rsEmpSchedC!SHIFT_END_DATE, l_rsEmpSchedC!SHIFT_END_TIME)
        r_blnIsTradedShift = IsTradedWorkShift(Trim(l_rsEmpSchedC!SYSTEM_EXC_CD))
    End If

    Set l_rsEmpSchedC = Nothing
    Exit Function

eh:
    Set l_rsEmpSchedC = Nothing
    TrsRaiseError m_strCMODULENAME, "PreviousRegularShiftEndTimeBefore"
    
End Function

Private Function IsTradedWorkShift(ByVal v_strSysExcCd As String) As Boolean

    IsTradedWorkShift = _
        v_strSysExcCd = "TSR" Or _
        v_strSysExcCd = "TSA" Or _
        v_strSysExcCd = "TWS"
        
End Function


Private Function HoursPreviouslyScheduledOvertime( _
                    ByVal v_lngEmpNum As Long, _
                    ByVal v_datChargeDate As Date, _
                    ByVal v_datDateWorked As Date, _
                    ByVal v_strStartTime As String, _
                    ByVal v_strEndTime As String) _
                As Variant
'Answer the number hours scheduled upto v_datDateWorked/v_strStartTime,
' this charge date only.

    Dim l_rsEmpSchedC As ADODB.Recordset
    Dim l_datStartDateTime As Date
    Dim l_datSchedEndDateTime As Date
    Dim l_decSchedOTHours As Variant

On Error GoTo eh

    l_decSchedOTHours = CDec(0)
    l_datStartDateTime = DateTimeFrom(v_datDateWorked, v_strStartTime)

    Set l_rsEmpSchedC = m_rsEmpSchedules.Clone
    l_rsEmpSchedC.Filter = "EMP_NUM = " & v_lngEmpNum & " AND CHARGE_DATE = #" & v_datChargeDate & "#"
    l_rsEmpSchedC.Sort = "DATE_WORKED ASC, START_TIME ASC"
    
    Do Until l_rsEmpSchedC.EOF
        If IsScheduledOvertime(l_rsEmpSchedC) Then
            l_datSchedEndDateTime = DateTimeFrom(l_rsEmpSchedC!DATE_WORKED, l_rsEmpSchedC!END_TIME, True)
'            If l_datSchedEndDateTime <= l_datStartDateTime Then
            If DatetimeCompare(l_datSchedEndDateTime, "LE", l_datStartDateTime) Then
                l_decSchedOTHours = l_decSchedOTHours + CDec(l_rsEmpSchedC!HOURS)
            Else
                Exit Do
            End If
        End If
        l_rsEmpSchedC.MoveNext
    Loop

    HoursPreviouslyScheduledOvertime = l_decSchedOTHours
    Set l_rsEmpSchedC = Nothing

xit:
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "HoursPreviouslyScheduledOvertime"
    
End Function


Private Function PrescheduledOvertimeAfterThisOpportunity( _
                    ByVal v_lngOppNum As Long, _
                    ByVal v_lngEmpNum As Long, _
                    ByVal v_datChargeDate As Date, _
                    ByVal v_datDateWorked As Date, _
                    ByVal v_strStartTime As String, _
                    ByVal v_strEndTime As String) _
                As Long
'Answer overtime opportunity number if any preschedule overtime is found
'following this opportunity either on the same charge date or continuous.

    Dim l_rsEmpSchedC As ADODB.Recordset
    Dim l_datStartDateTime As Date
    Dim l_datEndDateTime As Date
    Dim l_datStartDateTimeSched As Date
    Dim l_datEndDateTimeSched As Date
    
On Error GoTo eh
    
    If m_rsEmpSchedules Is Nothing Then
        Exit Function
    End If
    
    l_datStartDateTime = DateTimeFrom(v_datDateWorked, v_strStartTime)
    l_datEndDateTime = DateTimeFrom(v_datDateWorked, v_strEndTime)

    Set l_rsEmpSchedC = m_rsEmpSchedules.Clone
    l_rsEmpSchedC.Filter = "EMP_NUM = " & v_lngEmpNum
    l_rsEmpSchedC.Sort = "DATE_WORKED ASC, START_TIME ASC"

    Do Until l_rsEmpSchedC.EOF
        If IsScheduledOvertime(l_rsEmpSchedC) Then
            If l_rsEmpSchedC!OT_OPP_NUM <> v_lngOppNum Then
                l_datStartDateTimeSched = DateTimeFrom(l_rsEmpSchedC!DATE_WORKED, l_rsEmpSchedC!START_TIME)
'                If l_datStartDateTimeSched >= l_datEndDateTime Then
                If DatetimeCompare(l_datStartDateTimeSched, "GE", l_datEndDateTime) Then
                    If TimespanContinous(l_datStartDateTime, l_datEndDateTime, _
                                        l_datStartDateTimeSched, _
                                        DateTimeFrom(l_rsEmpSchedC!DATE_WORKED, l_rsEmpSchedC!END_TIME, True), _
                                        m_lngMINOFFALLOWCONTINUOUS) Then
                        PrescheduledOvertimeAfterThisOpportunity = l_rsEmpSchedC!OT_OPP_NUM
                        Exit Do
                    ElseIf l_rsEmpSchedC!CHARGE_DATE = v_datChargeDate Then
                        PrescheduledOvertimeAfterThisOpportunity = l_rsEmpSchedC!OT_OPP_NUM
                        Exit Do
                    End If
                End If
            End If
        End If
        l_rsEmpSchedC.MoveNext
    Loop
                
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "PrescheduledOvertimeAfterThisOpportunity"
End Function


Private Function HoursPreviouslyScheduledContinuousOvertime( _
                    ByVal v_lngEmpNum As Long, _
                    ByVal v_datChargeDate As Date, _
                    ByVal v_datDateWorked As Date, _
                    ByVal v_strStartTime As String, _
                    ByVal v_strEndTime As String) _
                As Variant
'Answer the number hours scheduled upto v_datDateWorked/v_strStartTime

    Dim l_datDateWorked As Date
    Dim l_strStartTime As String
    Dim l_strEndTime As String
    Dim l_rsEmpSchedC As ADODB.Recordset
    Dim l_decSchedOTHours As Variant

On Error GoTo eh

    l_datDateWorked = v_datDateWorked
    l_strStartTime = v_strStartTime
    l_strEndTime = v_strEndTime
    l_decSchedOTHours = CDec(0)

    Set l_rsEmpSchedC = m_rsEmpSchedules.Clone
    l_rsEmpSchedC.Filter = "EMP_NUM = " & v_lngEmpNum
    l_rsEmpSchedC.Sort = "DATE_WORKED ASC, START_TIME ASC"
    
'Sum overtime hours for charge date, except this opportunity.
    Do Until l_rsEmpSchedC.EOF
        If IsScheduledOvertime(l_rsEmpSchedC) Then
            If l_rsEmpSchedC!CHARGE_DATE = v_datChargeDate And _
                    l_rsEmpSchedC!OT_OPP_NUM <> m_lngOTOppNum Then
                l_decSchedOTHours = l_decSchedOTHours + CDec(l_rsEmpSchedC!HOURS)
            End If
        End If
        l_rsEmpSchedC.MoveNext
    Loop
        
    l_rsEmpSchedC.Sort = "DATE_WORKED DESC, START_TIME DESC"
    l_rsEmpSchedC.MoveFirst
    
    PositionScheduleBefore v_datDateWorked, v_strStartTime, l_rsEmpSchedC
    
'Sum prior overtime hours for this opportunity.
    Do Until l_rsEmpSchedC.EOF
        If IsScheduledOvertime(l_rsEmpSchedC) Then
            If TimespanContinous(DateTimeFrom(l_datDateWorked, l_strStartTime), _
                                DateTimeFrom(l_datDateWorked, l_strEndTime), _
                                DateTimeFrom(l_rsEmpSchedC!DATE_WORKED, l_rsEmpSchedC!START_TIME), _
                                DateTimeFrom(l_rsEmpSchedC!DATE_WORKED, l_rsEmpSchedC!END_TIME, True), _
                                m_lngMINOFFALLOWCONTINUOUS) Then
                l_datDateWorked = l_rsEmpSchedC!DATE_WORKED
                l_strStartTime = l_rsEmpSchedC!START_TIME
                l_strEndTime = l_rsEmpSchedC!END_TIME
                If l_rsEmpSchedC!OT_OPP_NUM = m_lngOTOppNum Then
                    l_decSchedOTHours = l_decSchedOTHours + CDec(l_rsEmpSchedC!HOURS)
                End If
            Else
                Exit Do
            End If
        End If
        l_rsEmpSchedC.MoveNext
    Loop

    HoursPreviouslyScheduledContinuousOvertime = l_decSchedOTHours
    Set l_rsEmpSchedC = Nothing
    Exit Function

eh:
    Set l_rsEmpSchedC = Nothing
    TrsRaiseError m_strCMODULENAME, "HoursPreviouslyScheduledContinuousOvertime"
    
End Function


Private Function ContinuousScheduledOvertimeHours( _
                    ByVal v_lngEmpNum As Long, _
                    ByVal v_datDateWorked As Date, _
                    ByVal v_strStartTime As String, _
                    ByVal v_strEndTime As String, _
                    ByVal v_lngOTOppNum As Long, _
                    ByVal v_datStartShift As Date, _
                    ByVal v_datEndShift As Date) As Variant
' Sum all continuous overtime hours for opportunities other than this one,
' regardless of charge date.  Then add prior overtime hours for this opportunity.
'
' Employees are due double time any time they work more than 8 hours continuous
' on a day off.  However, not all continuous hours need to be worked on the off day.
' E.g. Regular shift OWL, prearranged OT PM, prearranged OWL next day, first day off.
'

    Dim l_datDateWorked As Date
    Dim l_strStartTime As String
    Dim l_strEndTime As String
    Dim l_rsEmpSchedC As ADODB.Recordset
    
    Dim l_decSchedOTHours As Variant
    Dim l_blnContinuous As Boolean
    Dim l_lngOTOppNum As Long
    Dim l_lngNextOTOppNum As Long
    Dim l_datStartTime As Date
    Dim l_datEndTime As Date
    
On Error GoTo eh

    l_blnContinuous = True
    l_decSchedOTHours = CDec(0)
    l_datDateWorked = v_datDateWorked
    l_strStartTime = v_strStartTime
    l_strEndTime = v_strEndTime
    l_lngOTOppNum = v_lngOTOppNum
    
    l_datStartTime = v_datStartShift
    l_datEndTime = v_datEndShift

' Check preshift
' Determine continuous based on entire shift because, a lunch break is
' continuous while any break between shifts is not continuous.
' Paid lunch opportunity on different canvass is continuous with different canvass.
' Therefore this will return hours of the different opportunity + any previous continuous.
    Do While l_blnContinuous
        l_blnContinuous = TimespanContinuousOvertimeShift(l_datStartTime, l_datEndTime, _
                l_lngOTOppNum, l_lngNextOTOppNum, 2, 0)
        If l_blnContinuous Then
            l_decSchedOTHours = l_decSchedOTHours + _
                EmployeeScheduledOvertimeHours(v_lngEmpNum, l_lngNextOTOppNum, _
                    l_datStartTime, l_datEndTime)
            l_lngOTOppNum = l_lngNextOTOppNum
            l_lngNextOTOppNum = 0
        End If
    Loop

' 2/11/05 RSG - Only consider hours scheduled before this time charge.
'' Check postshift
'    l_blnContinuous = True
'    l_lngOTOppNum = v_lngOTOppNum
'    l_lngNextOTOppNum = 0
'    l_datStartTime = v_datStartShift
'    l_datEndTime = v_datEndShift
'    Do While l_blnContinuous
'        l_blnContinuous = TimespanContinuousOvertimeShift(l_datStartTime, l_datEndTime, _
'                l_lngOTOppNum, l_lngNextOTOppNum, 1, 0)
'        If l_blnContinuous Then
'            l_decSchedOTHours = l_decSchedOTHours + _
'                EmployeeScheduledOvertimeHours(v_lngEmpNum, l_lngNextOTOppNum, _
'                    l_datStartTime, l_datEndTime)
'            l_lngOTOppNum = l_lngNextOTOppNum
'            l_lngNextOTOppNum = 0
'        End If
'    Loop

'Sum prior overtime hours for this opportunity.
    Set l_rsEmpSchedC = m_rsEmpSchedules.Clone
    l_rsEmpSchedC.Filter = "EMP_NUM = " & v_lngEmpNum & " AND OT_OPP_NUM = " & v_lngOTOppNum
    
    If Not RecordsExist(l_rsEmpSchedC) Then
        GoTo xit
    End If
    
    l_rsEmpSchedC.Sort = "DATE_WORKED DESC, START_TIME DESC"
    PositionScheduleBefore v_datDateWorked, v_strStartTime, l_rsEmpSchedC
    
    Do Until l_rsEmpSchedC.EOF
        If IsScheduledOvertime(l_rsEmpSchedC) Then
            If TimespanContinous(DateTimeFrom(l_datDateWorked, l_strStartTime), _
                                DateTimeFrom(l_datDateWorked, l_strEndTime), _
                                DateTimeFrom(l_rsEmpSchedC!DATE_WORKED, l_rsEmpSchedC!START_TIME), _
                                DateTimeFrom(l_rsEmpSchedC!DATE_WORKED, l_rsEmpSchedC!END_TIME, True), _
                                m_lngMINOFFALLOWCONTINUOUS) Then
                l_datDateWorked = l_rsEmpSchedC!DATE_WORKED
                l_strEndTime = l_rsEmpSchedC!END_TIME
                l_strStartTime = l_rsEmpSchedC!START_TIME
                l_decSchedOTHours = l_decSchedOTHours + CDec(l_rsEmpSchedC!HOURS)
            Else
                Exit Do
            End If
        End If
        l_rsEmpSchedC.MoveNext
    Loop


xit:
    ContinuousScheduledOvertimeHours = l_decSchedOTHours
    Set l_rsEmpSchedC = Nothing
    Exit Function

eh:
    Set l_rsEmpSchedC = Nothing
    TrsRaiseError m_strCMODULENAME, "ContinuousScheduledOvertimeHours"
    
End Function


Private Function EmployeeScheduledOvertimeHours( _
                ByVal v_lngEmpNum As Long, _
                ByVal v_lngOTOppNum As Long, _
                ByRef r_datStartShift As Date, _
                ByRef r_datEndShift As Date) As Variant
' Answer the hours scheduled overtime for v_lngOTOppNum.

    Dim l_rsEmpSchedC As ADODB.Recordset
    Dim l_decCanvOTHours As Variant

On Error GoTo eh

    l_decCanvOTHours = CDec(0)
    Set l_rsEmpSchedC = m_rsEmpSchedules.Clone
    l_rsEmpSchedC.Filter = "EMP_NUM = " & v_lngEmpNum & " AND OT_OPP_NUM = " & v_lngOTOppNum
    l_rsEmpSchedC.Sort = "DATE_WORKED ASC, START_TIME ASC"

    Do Until l_rsEmpSchedC.EOF
        If IsOvertimeWorked(l_rsEmpSchedC!EXCEPTION_CD, l_rsEmpSchedC!OVER_EXC_CD) Then
            If l_decCanvOTHours = 0 Then    'first
                r_datStartShift = DateTimeFrom(l_rsEmpSchedC!DATE_WORKED, l_rsEmpSchedC!START_TIME)
            End If
            l_decCanvOTHours = l_decCanvOTHours + CDec(l_rsEmpSchedC!HOURS)
            'last worked will be last end time
            r_datEndShift = DateTimeFrom(l_rsEmpSchedC!DATE_WORKED, l_rsEmpSchedC!END_TIME)
        End If
        l_rsEmpSchedC.MoveNext
    Loop

xit:
    EmployeeScheduledOvertimeHours = l_decCanvOTHours
    Set l_rsEmpSchedC = Nothing
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "EmployeeScheduledOvertimeHours"
    
End Function


Private Function HoursOvertimeThisCanvass( _
                    ByVal v_lngEmpNum As Long, _
                    ByVal v_datChargeDate As Date, _
                    ByVal v_datDateWorked As Date, _
                    ByVal v_strStartTime As String, _
                    ByVal v_strEndTime As String, _
                    ByVal v_rsTimeCharges As ADODB.Recordset) _
                As Variant
'Answer the number hours overtime scheduled on this canvass
' upto v_datDateWorked/v_strStartTime

    Dim l_rsTimeChargesC As ADODB.Recordset
    Dim l_decCanvOTHours As Variant
    Dim l_datStartDateTime As Date

On Error GoTo eh

    l_decCanvOTHours = CDec(0)
    l_datStartDateTime = DateTimeFrom(v_datDateWorked, v_strStartTime)

    Set l_rsTimeChargesC = v_rsTimeCharges.Clone
    l_rsTimeChargesC.Filter = "EMP_NUM = " & v_lngEmpNum
    l_rsTimeChargesC.Sort = "DATE_WORKED DESC, START_TIME DESC"
        
    Do Until l_rsTimeChargesC.EOF
        If l_rsTimeChargesC!CHARGE_DATE = v_datChargeDate And _
            DatetimeCompare(l_datStartDateTime, "GT", DateTimeFrom(l_rsTimeChargesC!DATE_WORKED, l_rsTimeChargesC!START_TIME)) Then
'            (l_datStartDateTime > DateTimeFrom(l_rsTimeChargesC!DATE_WORKED, l_rsTimeChargesC!START_TIME)) Then
            If IsOvertimeWorked(l_rsTimeChargesC!EXCEPTION_CD, l_rsTimeChargesC!OVER_EXC_CD) Then
                l_decCanvOTHours = l_decCanvOTHours + CDec(l_rsTimeChargesC!HOURS)
            Else
                Exit Do
            End If
        End If
        l_rsTimeChargesC.MoveNext
    Loop

    HoursOvertimeThisCanvass = l_decCanvOTHours
    Set l_rsTimeChargesC = Nothing
    Exit Function

eh:
    Set l_rsTimeChargesC = Nothing
    TrsRaiseError m_strCMODULENAME, "HoursOvertimeThisCanvass"
    
End Function


Private Function IsEmployeeOnDutyPriorToOvertimeShift( _
                    ByVal v_lngEmpNum As Long, _
                    ByRef r_typCanvShift As CanvassedShift) _
                As Boolean

    Dim l_rsEmpSchedC As ADODB.Recordset
    Dim l_datStartDateTime As Date
    Dim l_datEndDateTime As Date

On Error GoTo eh

    IsEmployeeOnDutyPriorToOvertimeShift = False

    Set l_rsEmpSchedC = m_rsEmpSchedules.Clone
    l_rsEmpSchedC.Filter = "EMP_NUM = " & v_lngEmpNum
    l_rsEmpSchedC.Sort = "DATE_WORKED ASC, START_TIME ASC"
    
    Do Until l_rsEmpSchedC.EOF
        If IsScheduledWorktime(l_rsEmpSchedC) Then
            l_datStartDateTime = DateTimeFrom(l_rsEmpSchedC!DATE_WORKED, l_rsEmpSchedC!START_TIME)
            If DatetimeCompare(l_datStartDateTime, "LT", r_typCanvShift.StartDateTime) Then
'            If l_datStartDateTime < r_typCanvShift.StartDateTime Then
                l_datEndDateTime = DateTimeFrom(l_rsEmpSchedC!DATE_WORKED, l_rsEmpSchedC!END_TIME, True)
                If TimespanContinous(l_datStartDateTime, l_datEndDateTime, _
                                    r_typCanvShift.StartDateTime, _
                                    r_typCanvShift.EndDateTime, _
                                    0) Then
                    IsEmployeeOnDutyPriorToOvertimeShift = True
                    Exit Do
                End If
            Else
                Exit Do
            End If
        End If
        l_rsEmpSchedC.MoveNext
    Loop
                
    Set l_rsEmpSchedC = Nothing
    Exit Function

eh:
    Set l_rsEmpSchedC = Nothing
    TrsRaiseError m_strCMODULENAME, "IsEmployeeOnDutyPriorToOvertimeShift"
    
End Function


Private Function IsEmployeeCalledToReportWithin8Hours( _
                    ByVal v_lngEmpNum As Long, _
                    ByRef r_typCanvShift As CanvassedShift) _
                As Boolean
' Answer True if the last call start time is within 8 hours of shift start.

    Dim l_strCallStartTimestamp As String
    Dim l_datCallStart As Date

On Error GoTo eh

    l_strCallStartTimestamp = GetCallTime(v_lngEmpNum)
    If Not l_strCallStartTimestamp = mINITCANVCALLTIMESTAMP And _
            Not l_strCallStartTimestamp = "" Then
        l_datCallStart = DateTimeFromDB2Timestamp(l_strCallStartTimestamp)
        IsEmployeeCalledToReportWithin8Hours = _
            Abs(DateDiff("n", l_datCallStart, r_typCanvShift.StartDateTime)) < 480
    End If
    
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "IsEmployeeCalledToReportWithin8Hours"
    
End Function


Private Function IsScheduledOvertime(ByVal v_rsEmpSched As ADODB.Recordset) As Boolean

' EXC_CD = HR is also SHIFT_HOURS = 0 and START_TIME <> "" e.g. 08:00 16:00.
' RSG - Can this simply look at OT_OPP_NUM?

    IsScheduledOvertime = _
        (v_rsEmpSched!SHIFT_HOURS = 0 And Trim(v_rsEmpSched!START_TIME) <> "") And _
            v_rsEmpSched!OT_OPP_NUM > 0

'    IsScheduledOvertime = _
'        (v_rsEmpSched!SHIFT_HOURS = 0 And Trim(v_rsEmpSched!START_TIME) <> "")

End Function


Private Function IsRegularScheduledShift(ByVal v_rsEmpSched As ADODB.Recordset) As Boolean

    IsRegularScheduledShift = v_rsEmpSched!SHIFT_HOURS > 0 And _
                Not IsScheduledHolidayOrVacation(v_rsEmpSched!EXCEPTION_CD) And _
                Not IsSpecialOverrideCode(v_rsEmpSched!OVER_EXC_CD) And _
                Trim(v_rsEmpSched!SYSTEM_EXC_CD) <> "SLT"

End Function


Private Function IsScheduledWorktime(ByVal v_rsEmpSched As ADODB.Recordset) As Boolean

'    IsScheduledWorktime = (Trim(v_rsEmpSched!START_TIME) <> "" And _
'                                Trim(v_rsEmpSched!SYSTEM_EXC_CD) <> "SLT")

    IsScheduledWorktime = False
    If (Trim(v_rsEmpSched!START_TIME) <> "" And _
            Trim(v_rsEmpSched!SYSTEM_EXC_CD) <> "SLT") Then
        If IsActualWork(v_rsEmpSched!EXCEPTION_CD) Then
            IsScheduledWorktime = True
        End If
    End If

End Function

Private Function IsActualWork(ByVal v_strExc_Cd As String) As Boolean

On Error GoTo eh

    If Trim(v_strExc_Cd) = "" Then
        IsActualWork = True
    Else
        IsActualWork = GetWorkRestInd(v_strExc_Cd) = "W"
    End If
    
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "IsActualWork"

End Function


Private Function IsScheduledHolidayOrVacation(ByVal v_strExc_Cd As String) As Boolean

    Dim l_strSymbolCd As String

On Error GoTo eh

    If Trim(v_strExc_Cd) = "" Then
        IsScheduledHolidayOrVacation = False
    Else
        l_strSymbolCd = GetSymbolCode(v_strExc_Cd)
        IsScheduledHolidayOrVacation = (l_strSymbolCd = "H" Or l_strSymbolCd = "V")
    End If
    
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "IsScheduledHolidayOrVacation"

End Function


Private Sub PositionScheduleBefore( _
                    ByVal v_datDateWorked As Date, _
                    ByVal v_strStartTime As String, _
                    ByRef r_rsEmpSched As ADODB.Recordset)
'Position r_rsEmpSched on first timespan found(in descending order), that is
'before the input datetime.

    Dim l_datInput As Date

On Error GoTo eh
    l_datInput = DateTimeFrom(v_datDateWorked, v_strStartTime)

    PositionScheduleBeforeDate l_datInput, r_rsEmpSched

    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "PositionScheduleBefore"
End Sub


Private Sub PositionScheduleBeforeDate( _
                    ByVal v_datStartDateTime As Date, _
                    ByRef r_rsEmpSched As ADODB.Recordset)
'Position r_rsEmpSched on first timespan found(in descending order), that is
'before the input datetime.

    Dim l_datSchedule As Date
    Dim l_datInput As Date

On Error GoTo eh

    l_datInput = v_datStartDateTime

    Do Until r_rsEmpSched.EOF
        If Trim(r_rsEmpSched!END_TIME) <> "" Then
            l_datSchedule = DateTimeFrom(r_rsEmpSched!DATE_WORKED, r_rsEmpSched!END_TIME, True)
            If DatetimeCompare(l_datSchedule, "LE", l_datInput) Then
'            If l_datSchedule <= l_datInput Then
                Exit Do
            End If
        End If
        r_rsEmpSched.MoveNext
    Loop

    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "PositionScheduleBeforeDate"
End Sub


Private Sub PositionRegularScheduleBeforeDate( _
                    ByVal v_datStartDateTime As Date, _
                    ByRef r_rsEmpSched As ADODB.Recordset)
'Position r_rsEmpSched on first timespan found(in descending order), that is
'before the input datetime.

    Dim l_datSchedule As Date

On Error GoTo eh

    Do Until r_rsEmpSched.EOF
        If IsRegularScheduledShift(r_rsEmpSched) Then
            l_datSchedule = DateTimeFrom(r_rsEmpSched!DATE_WORKED, r_rsEmpSched!END_TIME, True)
'            If l_datSchedule <= v_datStartDateTime Then
            If DatetimeCompare(l_datSchedule, "LE", v_datStartDateTime) Then
                Exit Do
            End If
        End If
        r_rsEmpSched.MoveNext
    Loop

    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "PositionRegularScheduleBeforeDate"
End Sub


Private Function WorkedMoreThanEightThisDate(ByVal v_rsTimeCharge As ADODB.Recordset) As Boolean
'Answer True if this time charge is > 8 hours, or


'UNTESTED

    Dim l_rsTimeChargeC As ADODB.Recordset
    Dim l_datDateWorked As Date
    Dim l_strStartTime As String
    Dim l_strEndTime As String
    Dim l_datDateWorkedStart As String

On Error GoTo eh

    If v_rsTimeCharge!HOURS > 8 Then
        WorkedMoreThanEightThisDate = True
        GoTo xit
    End If

    Set l_rsTimeChargeC = v_rsTimeCharge.Clone
    l_rsTimeChargeC.Bookmark = v_rsTimeCharge.Bookmark
    
'    l_datChargeDate = l_rsTimeChargeC!CHARGE_DATE
'    l_varBookmark = v_rsTimeCharge.Bookmark
    l_rsTimeChargeC.Sort = "DATE_WORKED ASC, START_TIME ASC"
    
    'move backward to find first start time and date worked of continuous time charges.
    l_rsTimeChargeC.MovePrevious
    Do Until l_rsTimeChargeC.BOF
        If TimespanContinous(DateTimeFrom(l_datDateWorked, l_strStartTime), _
                            DateTimeFrom(l_datDateWorked, l_strEndTime), _
                            DateTimeFrom(l_rsTimeChargeC!DATE_WORKED, l_rsTimeChargeC!START_TIME), _
                            DateTimeFrom(l_rsTimeChargeC!DATE_WORKED, l_rsTimeChargeC!END_TIME, True), _
                            m_lngMINOFFALLOWCONTINUOUS) Then
            l_strStartTime = l_rsTimeChargeC!START_TIME
            l_datDateWorkedStart = l_rsTimeChargeC!DATE_WORKED
            l_rsTimeChargeC.MovePrevious
        Else
            Exit Do
        End If
    Loop


xit:
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "WorkedMoreThanEightThisDate"
    
End Function


Private Sub DetermineContinuousStartEndTimes( _
                ByVal v_rsTimeCharge As ADODB.Recordset, _
                ByRef r_datStartShift As Date, _
                ByRef r_datEndShift As Date)
' Split charges, ie. refusals should only be entered at the beginning
' or ending of a shift.  If entered in the middle, it will break the
' continuity of the shift.
'
' CS time is not worked and does not count toward any summed work time.
' Eg. does not count toward working > 8 continuous on day off, due double time.

    Dim l_varBookmark As Variant
    Dim l_strOrigSort As String
    Dim l_strStartTime As String
    Dim l_strEndTime As String
    Dim l_datDateWorkedStart As Date
    Dim l_datDateWorkedEnd As Date

On Error GoTo eh

    l_strStartTime = v_rsTimeCharge!START_TIME
    l_strEndTime = v_rsTimeCharge!END_TIME
    l_datDateWorkedStart = v_rsTimeCharge!DATE_WORKED
    l_datDateWorkedEnd = v_rsTimeCharge!DATE_WORKED
    l_varBookmark = v_rsTimeCharge.Bookmark
    l_strOrigSort = v_rsTimeCharge.Sort
    v_rsTimeCharge.Sort = "DATE_WORKED ASC, START_TIME ASC"

' restore position
    v_rsTimeCharge.Bookmark = l_varBookmark
    'move backward to find first start time and date worked of continuous time charges.
    v_rsTimeCharge.MovePrevious
    Do Until v_rsTimeCharge.BOF
        If TimespanContinous(DateTimeFrom(l_datDateWorkedStart, l_strStartTime), _
                            DateTimeFrom(l_datDateWorkedEnd, l_strEndTime), _
                            DateTimeFrom(v_rsTimeCharge!DATE_WORKED, v_rsTimeCharge!START_TIME), _
                            DateTimeFrom(v_rsTimeCharge!DATE_WORKED, v_rsTimeCharge!END_TIME, True), _
                            m_lngMINOFFALLOWCONTINUOUS) Then
            If Not IsOvertimeRefused(v_rsTimeCharge!EXCEPTION_CD) And _
                    Not IsOverrideThatAllowsOverlap(v_rsTimeCharge!OVER_EXC_CD) Then
                l_strStartTime = v_rsTimeCharge!START_TIME
                l_datDateWorkedStart = v_rsTimeCharge!DATE_WORKED
            Else
                Exit Do
            End If
        Else
            Exit Do
        End If
        v_rsTimeCharge.MovePrevious
    Loop
    
    'move forward to find last end time and date worked of continuous time charges.
    v_rsTimeCharge.Bookmark = l_varBookmark
    v_rsTimeCharge.MoveNext
    Do Until v_rsTimeCharge.EOF
        If TimespanContinous(DateTimeFrom(l_datDateWorkedStart, l_strStartTime), _
                            DateTimeFrom(l_datDateWorkedEnd, l_strEndTime), _
                            DateTimeFrom(v_rsTimeCharge!DATE_WORKED, v_rsTimeCharge!START_TIME), _
                            DateTimeFrom(v_rsTimeCharge!DATE_WORKED, v_rsTimeCharge!END_TIME, True), _
                            m_lngMINOFFALLOWCONTINUOUS) Then
            If Not IsOvertimeRefused(v_rsTimeCharge!EXCEPTION_CD) And _
                    Not IsOverrideThatAllowsOverlap(v_rsTimeCharge!OVER_EXC_CD) Then
                l_strEndTime = v_rsTimeCharge!END_TIME
                l_datDateWorkedEnd = v_rsTimeCharge!DATE_WORKED
            Else
                Exit Do
            End If
        Else
            Exit Do
        End If
        v_rsTimeCharge.MoveNext
    Loop
    
    r_datStartShift = DateTimeFrom(l_datDateWorkedStart, l_strStartTime)
    r_datEndShift = DateTimeFrom(l_datDateWorkedEnd, l_strEndTime)
    v_rsTimeCharge.Sort = l_strOrigSort
    v_rsTimeCharge.Bookmark = l_varBookmark

    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "DetermineContinuousStartEndTimes"
    
End Sub


Private Function EmployeeWorkedStartEndDateTimes( _
                ByVal v_lngEmpNum As Long, _
                ByVal v_rsTimeCharge As ADODB.Recordset, _
                ByRef r_datStartShift As Date, _
                ByRef r_datEndShift As Date) As Boolean
' Answer True if employee has any worked time; otherwise false.
' If there is worked time, return the first worked start datetime,
' and last worked end datetime.
'
' Only call this after missing components are determined.

    Dim l_rsTimeChargeC As ADODB.Recordset

On Error GoTo eh

    Set l_rsTimeChargeC = v_rsTimeCharge.Clone
    l_rsTimeChargeC.Filter = "EMP_NUM = " & v_lngEmpNum
    l_rsTimeChargeC.Sort = "DATE_WORKED ASC, START_TIME ASC"

    If Not RecordsExist(l_rsTimeChargeC) Then
        EmployeeWorkedStartEndDateTimes = False
        GoTo xit
    End If

'Find first worked timecharge
    Do Until l_rsTimeChargeC.EOF
        If IsOvertimeWorked(l_rsTimeChargeC!EXCEPTION_CD, l_rsTimeChargeC!OVER_EXC_CD) Then
            r_datStartShift = DateTimeFrom(l_rsTimeChargeC!DATE_WORKED, l_rsTimeChargeC!START_TIME)
            EmployeeWorkedStartEndDateTimes = True
            Exit Do
        End If
        l_rsTimeChargeC.MoveNext
    Loop

    If l_rsTimeChargeC.EOF Then
        EmployeeWorkedStartEndDateTimes = False
        GoTo xit
    End If

'Find last worked timecharge
    l_rsTimeChargeC.MoveLast
    Do Until l_rsTimeChargeC.BOF
        If IsOvertimeWorked(l_rsTimeChargeC!EXCEPTION_CD, l_rsTimeChargeC!OVER_EXC_CD) Then
            r_datEndShift = DateTimeFrom(l_rsTimeChargeC!DATE_WORKED, l_rsTimeChargeC!END_TIME, True)
            EmployeeWorkedStartEndDateTimes = True
            Exit Do
        End If
        l_rsTimeChargeC.MovePrevious
    Loop

xit:
    Set l_rsTimeChargeC = Nothing
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "EmployeeWorkedStartEndDateTimes"
    
End Function


Private Function EmployeeScheduledStartEndDateTimes( _
                ByVal v_lngEmpNum As Long, _
                ByVal v_lngOTOppNum As Long, _
                ByRef r_datStartShift As Date, _
                ByRef r_datEndShift As Date) As Boolean
' Answer True if employee has any scheduled time for opportunity; otherwise false.
' If there is scheduled time, return the first worked start datetime,
' and last worked end datetime.

    Dim l_rsEmpSchedC As ADODB.Recordset

On Error GoTo eh

    Set l_rsEmpSchedC = m_rsEmpSchedules.Clone
    l_rsEmpSchedC.Filter = "EMP_NUM = " & v_lngEmpNum & " AND OT_OPP_NUM = " & v_lngOTOppNum
    l_rsEmpSchedC.Sort = "DATE_WORKED ASC, START_TIME ASC"

    If Not RecordsExist(l_rsEmpSchedC) Then
        EmployeeScheduledStartEndDateTimes = False
        GoTo xit
    End If

'Find first scheduled work
    Do Until l_rsEmpSchedC.EOF
        If IsOvertimeWorked(l_rsEmpSchedC!EXCEPTION_CD, l_rsEmpSchedC!OVER_EXC_CD) Then
            r_datStartShift = DateTimeFrom(l_rsEmpSchedC!DATE_WORKED, l_rsEmpSchedC!START_TIME)
            EmployeeScheduledStartEndDateTimes = True
            Exit Do
        End If
        l_rsEmpSchedC.MoveNext
    Loop

    If l_rsEmpSchedC.EOF Then
        EmployeeScheduledStartEndDateTimes = False
        GoTo xit
    End If

'Find last scheduled work
    l_rsEmpSchedC.MoveLast
    Do Until l_rsEmpSchedC.BOF
        If IsOvertimeWorked(l_rsEmpSchedC!EXCEPTION_CD, l_rsEmpSchedC!OVER_EXC_CD) Then
            r_datEndShift = DateTimeFrom(l_rsEmpSchedC!DATE_WORKED, l_rsEmpSchedC!END_TIME, True)
            EmployeeScheduledStartEndDateTimes = True
            Exit Do
        End If
        l_rsEmpSchedC.MovePrevious
    Loop


xit:
    Set l_rsEmpSchedC = Nothing
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "EmployeeScheduledStartEndDateTimes"
    
End Function


Private Function EmployeeScheduledLunchStartEndDateTimes( _
                ByVal v_lngEmpNum As Long, _
                ByVal v_lngOTOppNum As Long, _
                ByRef r_datLunchStart As Date, _
                ByRef r_datLunchEnd As Date) As Boolean
' Answer True if employee has any scheduled lunch time for opportunity; otherwise false.
' Assuming the first break found is lunch.

    Dim l_rsEmpSchedC As ADODB.Recordset
    Dim l_datStartTime As Date
    Dim l_datEndTime As Date
    Dim l_datPrevEndTime As Date
    Dim l_blnFirst As Boolean
    

On Error GoTo eh

    Set l_rsEmpSchedC = m_rsEmpSchedules.Clone
    l_rsEmpSchedC.Filter = "EMP_NUM = " & v_lngEmpNum & " AND OT_OPP_NUM = " & v_lngOTOppNum
    l_rsEmpSchedC.Sort = "DATE_WORKED ASC, START_TIME ASC"

    EmployeeScheduledLunchStartEndDateTimes = False
    If Not RecordsExist(l_rsEmpSchedC) Then
        GoTo xit
    End If

    Do Until l_rsEmpSchedC.EOF
        If Trim(l_rsEmpSchedC!EXCEPTION_CD) = "SLT" Then
            r_datLunchStart = l_datStartTime = DateTimeFrom(l_rsEmpSchedC!DATE_WORKED, l_rsEmpSchedC!START_TIME)
            r_datLunchEnd = DateTimeFrom(l_rsEmpSchedC!DATE_WORKED, l_rsEmpSchedC!END_TIME, True)
            EmployeeScheduledLunchStartEndDateTimes = True
            Exit Do
        End If
        l_rsEmpSchedC.MoveNext
    Loop
    If Not l_rsEmpSchedC.EOF Then
        GoTo xit
    End If

    l_rsEmpSchedC.MoveFirst
    l_blnFirst = True
    Do Until l_rsEmpSchedC.EOF
        If IsOvertimeWorked(l_rsEmpSchedC!EXCEPTION_CD, l_rsEmpSchedC!OVER_EXC_CD) Then
            l_datStartTime = DateTimeFrom(l_rsEmpSchedC!DATE_WORKED, l_rsEmpSchedC!START_TIME)
            l_datEndTime = DateTimeFrom(l_rsEmpSchedC!DATE_WORKED, l_rsEmpSchedC!END_TIME, True)
            If l_blnFirst Then      'seed
                l_blnFirst = False
                l_datPrevEndTime = l_datStartTime
            End If
            If DatetimeCompare(l_datPrevEndTime, "NE", l_datStartTime) Then
                r_datLunchStart = l_datPrevEndTime
                r_datLunchEnd = l_datStartTime
                EmployeeScheduledLunchStartEndDateTimes = True
                Exit Do
            Else
                l_datPrevEndTime = l_datEndTime
            End If
        End If
        l_rsEmpSchedC.MoveNext
    Loop


xit:
    Set l_rsEmpSchedC = Nothing
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "EmployeeScheduledLunchStartEndDateTimes"
    
End Function


Private Function IsEmployeePaidDoubleForOvertimeShift( _
                ByVal v_lngEmpNum As Long, _
                ByVal v_lngOTOppNum As Long) As Boolean
' Answer True if all worked time scheduled at double time; otherwise false.

    Dim l_rsEmpSchedC As ADODB.Recordset
    Dim l_cuSet As utilTRISCommonLib.cuSet
    Dim l_strExcCd As String

On Error GoTo eh

    Set l_rsEmpSchedC = m_rsEmpSchedules.Clone
    l_rsEmpSchedC.Filter = "EMP_NUM = " & v_lngEmpNum & " AND OT_OPP_NUM = " & v_lngOTOppNum
    l_rsEmpSchedC.Sort = "DATE_WORKED ASC, START_TIME ASC"

    IsEmployeePaidDoubleForOvertimeShift = False
    If Not RecordsExist(l_rsEmpSchedC) Then
        GoTo xit
    End If

    Set l_cuSet = New cuSet
    Do Until l_rsEmpSchedC.EOF
        If IsOvertimeWorked(l_rsEmpSchedC!EXCEPTION_CD, l_rsEmpSchedC!OVER_EXC_CD) Then
            l_cuSet.Add Trim(l_rsEmpSchedC.Fields("EXCEPTION_CD").Value)
        End If
        l_rsEmpSchedC.MoveNext
    Loop

    If l_cuSet.Count = 1 Then
        If l_cuSet.Item(1) = "OT5" Then
            IsEmployeePaidDoubleForOvertimeShift = True
        End If
    End If

xit:
    Set l_cuSet = Nothing
    Set l_rsEmpSchedC = Nothing
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "IsEmployeePaidDoubleForOvertimeShift"
    
End Function


Private Function TimespanContinous( _
                    ByVal v_datStart1 As Date, _
                    ByVal v_datEnd1 As Date, _
                    ByVal v_datStart2 As Date, _
                    ByVal v_datEnd2 As Date, _
                    Optional ByVal v_lngMinBreak As Long = 0) _
                As Boolean
                    
On Error GoTo eh

    If DatetimeCompare(v_datStart1, "GE", v_datStart2) Then
        If DatetimeCompare(DateAdd("n", v_lngMinBreak, v_datEnd2), "GE", v_datStart1) Then
            TimespanContinous = True
        Else
            TimespanContinous = False
        End If
    Else
        If DatetimeCompare(DateAdd("n", v_lngMinBreak, v_datEnd1), "GE", v_datStart2) Then
            TimespanContinous = True
        Else
            TimespanContinous = False
        End If
    End If

    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "TimespanContinous"

End Function


Private Function DatetimeCompare( _
                    ByVal v_datDate1 As Date, _
                    ByVal v_strOperator As String, _
                    ByVal v_datDate2 As Date) As Boolean

    Dim l_bln As Boolean

On Error GoTo eh

    Select Case v_strOperator
        Case "EQ"
            l_bln = DateDiff("s", v_datDate1, v_datDate2) = 0
        Case "NE"
            l_bln = Abs(DateDiff("s", v_datDate1, v_datDate2)) > 0
        Case "LT"
            l_bln = DateDiff("s", v_datDate1, v_datDate2) > 0
        Case "LE"
            l_bln = DateDiff("s", v_datDate1, v_datDate2) >= 0
        Case "GT"
            l_bln = DateDiff("s", v_datDate1, v_datDate2) < 0
        Case "GE"
            l_bln = DateDiff("s", v_datDate1, v_datDate2) <= 0
        Case Else
            Err.Raise 13
    End Select

    DatetimeCompare = l_bln
    Exit Function
                    
eh:
    TrsRaiseError m_strCMODULENAME, "DatetimeCompare"
End Function


Private Sub PositionOnTimeCharge( _
                ByVal v_rsTimeCharges As ADODB.Recordset, _
                ByVal v_datDateWorked As Date, _
                ByVal v_strStartTime As String)
                
On Error GoTo eh

    v_rsTimeCharges.MoveFirst
    Do Until v_rsTimeCharges.EOF
        If v_rsTimeCharges!DATE_WORKED = v_datDateWorked And _
            v_rsTimeCharges!START_TIME = v_strStartTime Then
            Exit Do
        Else
            v_rsTimeCharges.MoveNext
        End If
    Loop

    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "PositionOnTimeCharge"

End Sub


Private Function EmployeeExceptionEarningsCode( _
                    ByVal v_lngEmpNum As Long, _
                    ByVal v_datChargeDate As Date, _
                    ByVal v_strExcCd As String, _
                    ByVal v_rsEmployees As ADODB.Recordset) As String

    Dim l_strUnionCd As String
    Dim l_strExecConInd As String
    Dim l_strEarningsCode As String
    Dim l_objCodeTables As dbCodeTables.cdCodeTables

On Error GoTo eh

    l_strUnionCd = EmployeeUnion(v_lngEmpNum, v_datChargeDate, v_rsEmployees)
    
    Set l_objCodeTables = CreateObject("dbCodeTables.cdCodeTables")
    
    l_strExecConInd = _
        l_objCodeTables.UnionExecutiveContractIndicator(l_strUnionCd)
    
    l_strEarningsCode = _
        l_objCodeTables.ExceptionEarningsCode( _
            v_strExcCd, _
            l_strExecConInd)
            
    Set l_objCodeTables = Nothing
    EmployeeExceptionEarningsCode = Trim(l_strEarningsCode)

    Exit Function

eh:
    Set l_objCodeTables = Nothing
    TrsRaiseError m_strCMODULENAME, "EmployeeExceptionEarningsCode"
    
End Function


Private Function EmployeeUnion( _
                    ByVal v_lngEmpNum As Long, _
                    ByVal v_datChargeDate As Date, _
                    ByVal v_rsEmployees As ADODB.Recordset) _
                As String
    
    Dim l_obj As busEmployee.cbsEmployee

On Error GoTo eh

    v_rsEmployees.Filter = "EMP_NUM = " & v_lngEmpNum
    Set l_obj = CreateObject("busEmployee.cbsEmployee")
    EmployeeUnion = _
        l_obj.EmployeeUnion( _
            v_rsEmployees!JOB_EFF_DATE, _
            v_rsEmployees!UNION_CD, _
            v_rsEmployees!PRIOR_UNION_CD, _
            v_datChargeDate)

    v_rsEmployees.Filter = adFilterNone
    Set l_obj = Nothing
    Exit Function

eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "EmployeeUnion"

End Function


Private Function EmployeeLocation( _
                    ByVal v_lngEmpNum As Long, _
                    ByVal v_rsEmployees As ADODB.Recordset) _
                As Long
    
On Error GoTo eh

    v_rsEmployees.Filter = "EMP_NUM = " & v_lngEmpNum
    If v_rsEmployees.EOF Then
        EmployeeLocation = 0
    Else
        EmployeeLocation = v_rsEmployees!LOC_CD
    End If

    v_rsEmployees.Filter = adFilterNone
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "EmployeeLocation"

End Function


Private Function GetEmployees( _
                    ByVal v_aryEmployees As Variant) _
                As ADODB.Recordset
    Dim l_obj As dbEmployee.cdEmployee

On Error GoTo eh
    
    Set l_obj = CreateObject("dbEmployee.cdEmployee")
    Set GetEmployees = l_obj.GetEmployee(v_aryEmployees)
    Set l_obj = Nothing
    Exit Function
    
eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "GetEmployees"
    
End Function


Private Function GetEmployeeSchedules( _
                    ByVal v_aryEmployees As Variant, _
                    ByVal v_datStart As Date, _
                    ByVal v_datEnd As Date) _
                As ADODB.Recordset

    Dim l_obj As busSchedule.cbsSchedule

On Error GoTo eh
    Set l_obj = CreateObject("busSchedule.cbsSchedule")
'    Set GetEmployeeSchedules = l_obj.GetAdjustedBaseSchedule( _
'        v_aryEmployees, v_datStart, v_datEnd, Nothing, , trAdjBaseSchedWithExceptionsAndOvertime)
        
    Set GetEmployeeSchedules = l_obj.GetScheduleForOTCanvassing( _
        v_aryEmployees, v_datStart, v_datEnd)
        
    Set l_obj = Nothing
    Exit Function

eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "GetEmployeeSchedules"

End Function


Private Function GetOtherPayEntries( _
                    ByVal v_aryEmployees As Variant, _
                    ByVal startDt As Date, _
                    ByVal endDt As Date) As ADODB.Recordset
    Dim ocdTimeCard As dbTimeCard.cdTimeCard

On Error GoTo eh
    
    Set ocdTimeCard = _
        CreateObject("dbTimeCard.cdTimeCard")
    Set GetOtherPayEntries = _
        ocdTimeCard.Get_OtherPay(v_aryEmployees, startDt, endDt)
    Set ocdTimeCard = Nothing
    Exit Function
    
eh:
    Set ocdTimeCard = Nothing
    TrsRaiseError m_strCMODULENAME, "GetOtherPayEntries"
    
End Function


Private Function GetEmployeeUnions( _
                    ByVal v_rsEmployees As ADODB.Recordset, _
                    ByVal v_colChargeDates As Collection) _
                As Collection
    Dim l_objEmployee As busEmployee.cbsEmployee
    Dim l_setUnionCodes As utilTRISCommonLib.cuSet
    Dim l_strUnionCd As String
    Dim l_varChargeDate As Variant

On Error GoTo eh

    Set l_setUnionCodes = New cuSet
    Set l_objEmployee = CreateObject("busEmployee.cbsEmployee")
 
    v_rsEmployees.MoveFirst
    Do Until v_rsEmployees.EOF
        For Each l_varChargeDate In v_colChargeDates
            l_strUnionCd = _
                l_objEmployee.EmployeeUnion( _
                    v_rsEmployees!JOB_EFF_DATE, _
                    v_rsEmployees!UNION_CD, _
                    v_rsEmployees!PRIOR_UNION_CD, _
                    CDate(l_varChargeDate))
            l_setUnionCodes.Add l_strUnionCd
        Next
        v_rsEmployees.MoveNext
    Loop

    Set GetEmployeeUnions = l_setUnionCodes.Contents
    Exit Function

eh:
    Set l_objEmployee = Nothing
    Set l_setUnionCodes = Nothing
    TrsRaiseError m_strCMODULENAME, "GetEmployeeUnions"
    
End Function


Private Sub FindHighLowTimeChargeDates( _
                ByRef r_LowDate As Date, _
                ByRef r_HighDate As Date, _
                ByVal v_rsTimeChargeReq As ADODB.Recordset)

On Error GoTo eh

    r_LowDate = FindLowDate(v_rsTimeChargeReq)
    r_HighDate = FindHighDate(v_rsTimeChargeReq)
    
    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "FindHighLowTimeChargeDates"

End Sub


Private Function FindLowDate( _
                    ByVal v_rsTimeCharges As ADODB.Recordset) As Date
                    
    Dim l_datDate As Date

On Error GoTo eh
    
    If Not RecordsExist(v_rsTimeCharges) Then
        FindLowDate = m_rsOTOpportunity!DATE_OF_OT
        Exit Function
    End If
    
    l_datDate = #12/31/9999#
    v_rsTimeCharges.MoveFirst
    Do Until v_rsTimeCharges.EOF
        If Not IsEmpty(v_rsTimeCharges.Fields("CHARGE_DATE").Value) Then
            l_datDate = MinDate(l_datDate, v_rsTimeCharges!CHARGE_DATE)
        End If
        l_datDate = MinDate(l_datDate, v_rsTimeCharges!KEY_CHARGE_DATE)
        v_rsTimeCharges.MoveNext
    Loop
    
    v_rsTimeCharges.MoveFirst
    FindLowDate = l_datDate
    Exit Function
    
eh:
    
    TrsRaiseError m_strCMODULENAME, "FindLowDate"
        
End Function


Private Function FindHighDate( _
                    ByVal v_rsTimeCharges As ADODB.Recordset) As Date
                    
    Dim l_datDate As Date

On Error GoTo eh
    
    If Not RecordsExist(v_rsTimeCharges) Then
        FindHighDate = m_rsOTOpportunity!DATE_OF_OT
        Exit Function
    End If
    
    l_datDate = #1/1/1901#
    v_rsTimeCharges.MoveFirst
    Do Until v_rsTimeCharges.EOF
        If Not IsEmpty(v_rsTimeCharges.Fields("CHARGE_DATE").Value) Then
            l_datDate = MaxDate(l_datDate, v_rsTimeCharges!CHARGE_DATE)
        End If
        l_datDate = MaxDate(l_datDate, v_rsTimeCharges!KEY_CHARGE_DATE)
        v_rsTimeCharges.MoveNext
    Loop
    
    v_rsTimeCharges.MoveFirst
    FindHighDate = l_datDate
    Exit Function
    
eh:
    
    TrsRaiseError m_strCMODULENAME, "FindHighDate"
        
End Function


Private Function FindLowDateWorked( _
                    ByVal v_rsTimeCharge As ADODB.Recordset) As Date
' Scan all time charges, including deleted rows, to find low date.
                    
    Dim l_datDate As Date

On Error GoTo eh
    
    l_datDate = #12/31/9999#
    
    If RecordsExist(v_rsTimeCharge) Then
        v_rsTimeCharge.MoveFirst
        Do Until v_rsTimeCharge.EOF
            l_datDate = MinDate(l_datDate, v_rsTimeCharge!DATE_WORKED)
            v_rsTimeCharge.MoveNext
        Loop
    End If
    
    FilterPendingChanges v_rsTimeCharge
    If RecordsExist(v_rsTimeCharge) Then
        Do Until v_rsTimeCharge.EOF
            If v_rsTimeCharge.Status = adRecDeleted Or v_rsTimeCharge.Status = adRecModified Then
                l_datDate = MinDate(l_datDate, v_rsTimeCharge.Fields("DATE_WORKED").OriginalValue)
            End If
            v_rsTimeCharge.MoveNext
        Loop
    End If
    v_rsTimeCharge.Filter = adFilterNone
    
    If RecordsExist(v_rsTimeCharge) Then
        v_rsTimeCharge.MoveFirst
    End If

    FindLowDateWorked = l_datDate
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "FindLowDateWorked"
        
End Function


Private Function FindLowChargeDate( _
                    ByVal v_rsTimeCharge As ADODB.Recordset) As Date
' Scan all time charges, including deleted rows, to find low date.
                    
    Dim l_datDate As Date

On Error GoTo eh
    
    l_datDate = #12/31/9999#
    
    If RecordsExist(v_rsTimeCharge) Then
        v_rsTimeCharge.MoveFirst
        Do Until v_rsTimeCharge.EOF
            l_datDate = MinDate(l_datDate, v_rsTimeCharge!CHARGE_DATE)
            v_rsTimeCharge.MoveNext
        Loop
    End If
    
    FilterPendingChanges v_rsTimeCharge
    If RecordsExist(v_rsTimeCharge) Then
        Do Until v_rsTimeCharge.EOF
            If v_rsTimeCharge.Status = adRecDeleted Or v_rsTimeCharge.Status = adRecModified Then
                l_datDate = MinDate(l_datDate, v_rsTimeCharge.Fields("CHARGE_DATE").OriginalValue)
            End If
            v_rsTimeCharge.MoveNext
        Loop
    End If
    v_rsTimeCharge.Filter = adFilterNone
    
    If RecordsExist(v_rsTimeCharge) Then
        v_rsTimeCharge.MoveFirst
    End If

    FindLowChargeDate = l_datDate
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "FindLowChargeDate"
        
End Function



Private Function FindHighDateWorked( _
                    ByVal v_rsTimeCharge As ADODB.Recordset) As Date
                    
' Scan all time charges, including deleted rows, to find high date.
                    
    Dim l_datDate As Date

On Error GoTo eh
    
    l_datDate = #1/1/1901#
    
    If RecordsExist(v_rsTimeCharge) Then
        v_rsTimeCharge.MoveFirst
        Do Until v_rsTimeCharge.EOF
            l_datDate = MaxDate(l_datDate, v_rsTimeCharge!DATE_WORKED)
            v_rsTimeCharge.MoveNext
        Loop
    End If
    
    FilterPendingChanges v_rsTimeCharge
    If RecordsExist(v_rsTimeCharge) Then
        Do Until v_rsTimeCharge.EOF
            If v_rsTimeCharge.Status = adRecDeleted Or v_rsTimeCharge.Status = adRecModified Then
                l_datDate = MaxDate(l_datDate, v_rsTimeCharge.Fields("DATE_WORKED").OriginalValue)
            End If
            v_rsTimeCharge.MoveNext
        Loop
    End If
    v_rsTimeCharge.Filter = adFilterNone
    
    If RecordsExist(v_rsTimeCharge) Then
        v_rsTimeCharge.MoveFirst
    End If
    
    FindHighDateWorked = l_datDate
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "FindHighDateWorked"
        
End Function


Private Function FindHighChargeDate( _
                    ByVal v_rsTimeCharge As ADODB.Recordset) As Date
                    
' Scan all time charges, including deleted rows, to find high date.
                    
    Dim l_datDate As Date

On Error GoTo eh
    
    l_datDate = #1/1/1901#
    
    If RecordsExist(v_rsTimeCharge) Then
        v_rsTimeCharge.MoveFirst
        Do Until v_rsTimeCharge.EOF
            l_datDate = MaxDate(l_datDate, v_rsTimeCharge!CHARGE_DATE)
            v_rsTimeCharge.MoveNext
        Loop
    End If
    
    FilterPendingChanges v_rsTimeCharge
    If RecordsExist(v_rsTimeCharge) Then
        Do Until v_rsTimeCharge.EOF
            If v_rsTimeCharge.Status = adRecDeleted Or v_rsTimeCharge.Status = adRecModified Then
                l_datDate = MaxDate(l_datDate, v_rsTimeCharge.Fields("CHARGE_DATE").OriginalValue)
            End If
            v_rsTimeCharge.MoveNext
        Loop
    End If
    v_rsTimeCharge.Filter = adFilterNone
    
    If RecordsExist(v_rsTimeCharge) Then
        v_rsTimeCharge.MoveFirst
    End If
    
    FindHighChargeDate = l_datDate
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "FindHighChargeDate"
        
End Function



Private Sub AddChargeDate(ByVal v_datChgDate As Date)

On Error GoTo eh

    If m_setChargeDates Is Nothing Then
        Set m_setChargeDates = New cuSet
    End If

    m_setChargeDates.Add v_datChgDate

    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "AddChargeDate"
    
End Sub

Private Function GetDistinctChargeDates( _
                    ByVal v_rsTimeCharges As ADODB.Recordset) As Collection
                    
    Dim l_datDate As Date
    Dim l_setChargeDates As utilTRISCommonLib.cuSet

On Error GoTo eh

    v_rsTimeCharges.Filter = adFilterPendingRecords
    Set l_setChargeDates = New cuSet
    
    Do Until v_rsTimeCharges.EOF
        l_setChargeDates.Add v_rsTimeCharges.Fields("CHARGE_DATE").Value
        v_rsTimeCharges.MoveNext
    Loop
    
    Set GetDistinctChargeDates = l_setChargeDates.Contents
    v_rsTimeCharges.Filter = adFilterNone
    Set l_setChargeDates = Nothing
    Exit Function
    
eh:
    Set l_setChargeDates = Nothing
    TrsRaiseError m_strCMODULENAME, "GetDistinctChargeDates"
        
End Function


Private Sub AddTimeChargeFromOpportunity(ByVal v_lngEmpNum As Long)
'Times may be 24:00 or 00:00.  Adjusted date depending on whether start or end time
'and crosses midnight.

    Dim l_strStartTime As String
    Dim l_strEndTime As String
    Dim l_strStartLunch As String
    Dim l_strEndLunch As String
    Dim l_datDateWorked As Date
    Dim l_datChargeDate As Date
    Dim l_strOvertimeRate As String
    
    Dim l_typCanvShift As CanvassedShift
'    Dim l_rsTempTimeCharges As ADODB.Recordset
    
On Error GoTo eh
    
    SetSystemGeneratedTimeCharge v_lngEmpNum
    
    l_datDateWorked = m_rsOTOpportunity!DATE_OF_OT
    l_strStartTime = m_rsOTOpportunity!START_TIME
    l_strEndTime = m_rsOTOpportunity!END_TIME
    l_strStartLunch = Trim(m_rsOTOpportunity!START_LUNCH)
    l_strEndLunch = Trim(m_rsOTOpportunity!END_LUNCH)
    
    l_typCanvShift = NewCanvassedShiftFromOpportunity(m_rsOTOpportunity)
    
    If l_strStartLunch <> "" Then
        If l_strStartTime > l_strStartLunch Then
            AddTimeChargeRow v_lngEmpNum, l_datDateWorked, l_strStartTime, "24:00", l_typCanvShift
            If l_strStartLunch = "00:00" Then
                AddTimeChargeRow v_lngEmpNum, DateAdd("d", 1, l_datDateWorked), l_strEndLunch, l_strEndTime, l_typCanvShift
            Else
                AddTimeChargeRow v_lngEmpNum, DateAdd("d", 1, l_datDateWorked), "00:00", l_strStartLunch, l_typCanvShift
                AddTimeChargeRow v_lngEmpNum, DateAdd("d", 1, l_datDateWorked), l_strEndLunch, l_strEndTime, l_typCanvShift
            End If
        Else
            AddTimeChargeRow v_lngEmpNum, l_datDateWorked, l_strStartTime, l_strStartLunch, l_typCanvShift
            If l_strStartTime > l_strEndLunch Then
                AddTimeChargeRow v_lngEmpNum, DateAdd("d", 1, l_datDateWorked), l_strEndLunch, l_strEndTime, l_typCanvShift
            ElseIf l_strEndLunch = "24:00" Then
                AddTimeChargeRow v_lngEmpNum, DateAdd("d", 1, l_datDateWorked), "00:00", l_strEndTime, l_typCanvShift
            ElseIf l_strStartTime > l_strEndTime Then
                AddTimeChargeRow v_lngEmpNum, l_datDateWorked, l_strEndLunch, "24:00", l_typCanvShift
                If l_strEndTime <> "00:00" Then
                    AddTimeChargeRow v_lngEmpNum, DateAdd("d", 1, l_datDateWorked), "00:00", l_strEndTime, l_typCanvShift
                End If
            Else
                AddTimeChargeRow v_lngEmpNum, l_datDateWorked, l_strEndLunch, l_strEndTime, l_typCanvShift
            End If
        End If
    ElseIf l_strStartTime > l_strEndTime Then
        AddTimeChargeRow v_lngEmpNum, l_datDateWorked, l_strStartTime, "24:00", l_typCanvShift
        If l_strEndTime <> "00:00" Then
            AddTimeChargeRow v_lngEmpNum, DateAdd("d", 1, l_datDateWorked), "00:00", l_strEndTime, l_typCanvShift
        End If
    Else
        AddTimeChargeRow v_lngEmpNum, l_datDateWorked, l_strStartTime, l_strEndTime, l_typCanvShift
    End If
    
    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "AddTimeChargeFromOpportunity"
    
End Sub


Private Sub AddOvertimeSchedule(ByVal v_rsTimeCharge As ADODB.Recordset)
    
On Error GoTo eh
    
    If IsOvertimeRefused(v_rsTimeCharge!EXCEPTION_CD) Or _
        IsOverrideThatAllowsOverlap(v_rsTimeCharge!OVER_EXC_CD) Or _
            m_rsSchedOvertime Is Nothing Then
        Exit Sub
    End If
    
    m_rsSchedOvertime.AddNew
    m_rsSchedOvertime!EMP_NUM = v_rsTimeCharge!EMP_NUM
    m_rsSchedOvertime!Date = v_rsTimeCharge!DATE_WORKED
    m_rsSchedOvertime!SCHED_CD = 0
    m_rsSchedOvertime!SCHED_TYPE = 120
    m_rsSchedOvertime!START_TIME = v_rsTimeCharge!START_TIME
    m_rsSchedOvertime!END_TIME = v_rsTimeCharge!END_TIME
    m_rsSchedOvertime!EXCEPTION_CD = v_rsTimeCharge!EXCEPTION_CD
    m_rsSchedOvertime!WG_NUM = v_rsTimeCharge!WG_NUM
    m_rsSchedOvertime!JOB_NUM = v_rsTimeCharge!JOB_NUM
'    m_rsSchedOvertime!COMP_SHIFT_CD = v_rsTimeCharge!COMP_SHIFT_CD
    m_rsSchedOvertime!RELIEF_COMP = ""
    m_rsSchedOvertime!OT_OPP_NUM = v_rsTimeCharge!OT_OPP_NUM
    m_rsSchedOvertime!CASE_NUM = 0
    m_rsSchedOvertime!LOC_CD = m_rsOTOpportunity!LOC_CD
    m_rsSchedOvertime!COMP_CD = ""
    m_rsSchedOvertime!SHIFT_COMP = ""
    m_rsSchedOvertime!OVER_EXC_CD = v_rsTimeCharge!OVER_EXC_CD
    m_rsSchedOvertime!FMLA_IND = ""
    m_rsSchedOvertime!TRAV_PAY_HRS = v_rsTimeCharge!TRAVEL_PAY_HRS
    m_rsSchedOvertime!CHARGE_DATE = v_rsTimeCharge!CHARGE_DATE
    m_rsSchedOvertime!SUBSYS_OWNER = "E"
    m_rsSchedOvertime!LAST_MOD_EMP_NUM = m_lngUserNum
    m_rsSchedOvertime!PENDING_STATUS = ""
    m_rsSchedOvertime!LINE_NUM = 0
    m_rsSchedOvertime!REQ_IND = ""
    m_rsSchedOvertime!PREV_SCHED_CD = 0
    
    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "AddOvertimeSchedule"
    
End Sub


Private Sub AddAdjustedScheduleOvertime(ByVal v_rsTimeCharge As ADODB.Recordset)
    
On Error GoTo eh
    
    If IsOvertimeRefused(v_rsTimeCharge!EXCEPTION_CD) Or _
        IsOverrideThatAllowsOverlap(v_rsTimeCharge!OVER_EXC_CD) Or _
            m_rsEmpSchedules Is Nothing Then
        Exit Sub
    End If
    
    m_rsEmpSchedules.AddNew
    m_rsEmpSchedules!EMP_NUM = v_rsTimeCharge!EMP_NUM
    m_rsEmpSchedules!CHARGE_DATE = v_rsTimeCharge!CHARGE_DATE
    m_rsEmpSchedules!DATE_WORKED = v_rsTimeCharge!DATE_WORKED
    m_rsEmpSchedules!SHIFT_START_TIME = ""
    m_rsEmpSchedules!SHIFT_START_DATE = v_rsTimeCharge!DATE_WORKED
    m_rsEmpSchedules!SHIFT_END_TIME = ""
    If v_rsTimeCharge!START_TIME > v_rsTimeCharge!END_TIME Then
        m_rsEmpSchedules!SHIFT_END_DATE = DateAdd("d", 1, v_rsTimeCharge!DATE_WORKED)
    Else
        m_rsEmpSchedules!SHIFT_END_DATE = v_rsTimeCharge!DATE_WORKED
    End If
    m_rsEmpSchedules!SHIFT_HOURS = 0
    m_rsEmpSchedules!START_TIME = v_rsTimeCharge!START_TIME
    m_rsEmpSchedules!END_TIME = v_rsTimeCharge!END_TIME
    m_rsEmpSchedules!HOURS = v_rsTimeCharge!HOURS
    m_rsEmpSchedules!EXCEPTION_CD = v_rsTimeCharge!EXCEPTION_CD
    m_rsEmpSchedules!OVER_EXC_CD = v_rsTimeCharge!OVER_EXC_CD
    m_rsEmpSchedules!SCHED_JOB_NUM = v_rsTimeCharge!JOB_NUM
    m_rsEmpSchedules!EMP_JOB_NUM = ""
    m_rsEmpSchedules!COMP_SHIFT_CD = v_rsTimeCharge!COMP_SHIFT_CD
    m_rsEmpSchedules!SYS_COMP_SHIFT_CD = ""
    m_rsEmpSchedules!WG_NUM = v_rsTimeCharge!WG_NUM
    m_rsEmpSchedules!CASE_NUM = 0
    m_rsEmpSchedules!FMLA_IND = ""
    m_rsEmpSchedules!COMP_CD = ""
    m_rsEmpSchedules!REL_CD = ""
    m_rsEmpSchedules!SYSTEM_EXC_CD = ""
    m_rsEmpSchedules!Shift_Code = ""
    m_rsEmpSchedules!OT_OPP_NUM = v_rsTimeCharge!OT_OPP_NUM
    m_rsEmpSchedules!TRAV_PAY_HOURS = v_rsTimeCharge!TRAVEL_PAY_HRS
    m_rsEmpSchedules!LINE_NUM = 0
    m_rsEmpSchedules!SHIFT_NUM = 0
    m_rsEmpSchedules!SHIFT_LUNCH_START_TIME = ""
    m_rsEmpSchedules!SHIFT_LUNCH_END_TIME = ""
    
    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "AddAdjustedScheduleOvertime"
    
End Sub


Private Function NewCanvassedShiftFromOpportunity( _
                    ByVal v_rsOTOpportunity As ADODB.Recordset) _
                As CanvassedShift

    Dim l_typ As CanvassedShift
    Dim l_datEnd As Date
    Dim l_datLunch As Date
    Dim l_datLunchStart As Date
    Dim l_datLunchEnd As Date
    Dim l_varHoursBeforeLunch As Variant
    Dim l_varHoursAfterLunch As Variant
    
On Error GoTo eh

    l_typ.ChargeDate = v_rsOTOpportunity!CHARGE_DATE
    l_typ.StartDateTime = DateTimeFrom(v_rsOTOpportunity!DATE_OF_OT, m_rsOTOpportunity!START_TIME)
    If m_rsOTOpportunity!START_TIME > m_rsOTOpportunity!END_TIME Then
        l_datEnd = DateAdd("d", 1, v_rsOTOpportunity!DATE_OF_OT)
    Else
        l_datEnd = v_rsOTOpportunity!DATE_OF_OT
    End If
    l_typ.EndDateTime = DateTimeFrom(l_datEnd, m_rsOTOpportunity!END_TIME)
    l_typ.LunchStartTime = Trim(m_rsOTOpportunity!START_LUNCH)
    l_typ.LunchEndTime = Trim(m_rsOTOpportunity!END_LUNCH)
    If l_typ.LunchStartTime <> "" Then
        If m_rsOTOpportunity!START_TIME > m_rsOTOpportunity!START_LUNCH Then
            l_datLunch = DateAdd("d", 1, v_rsOTOpportunity!DATE_OF_OT)
        Else
            l_datLunch = v_rsOTOpportunity!DATE_OF_OT
        End If
        l_datLunchStart = DateTimeFrom(l_datLunch, m_rsOTOpportunity!START_LUNCH)
        If m_rsOTOpportunity!START_TIME > m_rsOTOpportunity!END_LUNCH Then
            l_datLunch = DateAdd("d", 1, v_rsOTOpportunity!DATE_OF_OT)
        Else
            l_datLunch = v_rsOTOpportunity!DATE_OF_OT
        End If
        l_datLunchEnd = DateTimeFrom(l_datLunch, m_rsOTOpportunity!END_LUNCH)
        l_varHoursBeforeLunch = CDec(DateDiff("n", l_typ.StartDateTime, l_datLunchStart) / 60)
        l_varHoursAfterLunch = CDec(DateDiff("n", l_datLunchEnd, l_typ.EndDateTime) / 60)
        l_typ.ShiftHours = l_varHoursBeforeLunch + l_varHoursAfterLunch
    Else
        l_typ.ShiftHours = CDec(DateDiff("n", l_typ.StartDateTime, l_typ.EndDateTime) / 60)
    End If
    l_typ.ShiftCode = v_rsOTOpportunity!SHIFT_CD
    l_typ.TravelPaid = False

    NewCanvassedShiftFromOpportunity = l_typ
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "NewCanvassedShiftFromOpportunity"
End Function


Private Function OpportunityStartEndDateTimes( _
                    ByRef r_datStart As Date, ByRef r_datEnd As Date)

    Dim l_datEnd As Date

On Error GoTo eh

    r_datStart = DateTimeFrom(m_rsOTOpportunity!DATE_OF_OT, m_rsOTOpportunity!START_TIME)
    If m_rsOTOpportunity!START_TIME > m_rsOTOpportunity!END_TIME Then
        l_datEnd = DateAdd("d", 1, m_rsOTOpportunity!DATE_OF_OT)
    Else
        l_datEnd = m_rsOTOpportunity!DATE_OF_OT
    End If
    r_datEnd = DateTimeFrom(l_datEnd, m_rsOTOpportunity!END_TIME)

    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "OpportunityStartEndDateTimes"

End Function


Private Function NewCanvassedShiftFromTimeCharges( _
                    ByVal v_lngEmpNum As Long, _
                    ByVal v_rsTimeCharge As ADODB.Recordset) _
                As CanvassedShift
' Construct a shift, i.e. a continuous work period, based on the current time charge.
' A shift occurs on a charge date.

' Is a refusal time charge for an acceptor considered part of a shift?
'
    Dim l_rsTimeChargeC As ADODB.Recordset
'    Dim l_varBookmark As Variant
    Dim l_typ As CanvassedShift
    Dim l_datStartShift As Date
    Dim l_datEndShift As Date
    Dim l_decShiftHours As Variant
    Dim l_blnTravelPaid As Boolean
    
On Error GoTo eh

    Set l_rsTimeChargeC = v_rsTimeCharge.Clone
    l_rsTimeChargeC.Filter = "EMP_NUM = " & v_lngEmpNum
    l_rsTimeChargeC.Sort = "DATE_WORKED ASC, START_TIME ASC"
    l_rsTimeChargeC.Bookmark = v_rsTimeCharge.Bookmark

'Assuming all timecharges must have same charge date.  True?

    DetermineContinuousStartEndTimes l_rsTimeChargeC, l_datStartShift, l_datEndShift
    l_typ.StartDateTime = l_datStartShift
    l_typ.EndDateTime = l_datEndShift
    l_typ.ChargeDate = m_rsOTOpportunity!CHARGE_DATE
    l_typ.ShiftCode = ""
    
'    Sum shift hours
    l_rsTimeChargeC.MoveFirst
    Do Until l_rsTimeChargeC.EOF
        If TimePeriodsSubset( _
                DateTimeFrom(l_rsTimeChargeC!DATE_WORKED, l_rsTimeChargeC!START_TIME), _
                DateTimeFrom(l_rsTimeChargeC!DATE_WORKED, l_rsTimeChargeC!END_TIME, True), _
                l_datStartShift, l_datEndShift) Then
            l_decShiftHours = CDec(l_decShiftHours) + CDec(l_rsTimeChargeC!HOURS)
        End If
        If l_rsTimeChargeC!TRAVEL_PAY_HRS > 0 Then
            l_blnTravelPaid = True
        End If
        l_rsTimeChargeC.MoveNext
    Loop
    l_typ.ShiftHours = l_decShiftHours
    l_typ.TravelPaid = l_blnTravelPaid

    Set l_rsTimeChargeC = Nothing
    NewCanvassedShiftFromTimeCharges = l_typ
    Exit Function

eh:
    Set l_rsTimeChargeC = Nothing
    TrsRaiseError m_strCMODULENAME, "NewCanvassedShiftFromTimeCharges"
    
End Function


Private Sub MarkAllTimeChargesForMissingComponents( _
                    ByVal v_lngEmpNum As Long, _
                    ByVal v_rsTimeCharge As ADODB.Recordset)
' Do not mark refusal time charges for determination.  Now that a refusal will
' break continuity, it should not be a subset of shift start and end.

    Dim l_rsTimeChargeC As ADODB.Recordset
    Dim l_datStartShift As Date
    Dim l_datEndShift As Date
    Dim l_blnTravelPaid As Boolean
    
On Error GoTo eh

    Set l_rsTimeChargeC = v_rsTimeCharge.Clone
    l_rsTimeChargeC.Filter = "EMP_NUM = " & v_lngEmpNum
    
    If Not RecordsExist(l_rsTimeChargeC) Then
        GoTo xit
    End If
    
    l_rsTimeChargeC.Sort = "DATE_WORKED ASC, START_TIME ASC"

'All timecharges for an opportunity will have same charge date.

    DetermineContinuousStartEndTimes l_rsTimeChargeC, l_datStartShift, l_datEndShift
    
    l_rsTimeChargeC.MoveFirst
    Do Until l_rsTimeChargeC.EOF
        If TimePeriodsSubset( _
                DateTimeFrom(l_rsTimeChargeC!DATE_WORKED, l_rsTimeChargeC!START_TIME), _
                DateTimeFrom(l_rsTimeChargeC!DATE_WORKED, l_rsTimeChargeC!END_TIME, True), _
                l_datStartShift, l_datEndShift) Then
'            Only mark for determination if the user did not change value
            If Not IsOvertimeRefused(l_rsTimeChargeC!EXCEPTION_CD) Then
                If IsEmpty(l_rsTimeChargeC.Fields("EXCEPTION_CD").OriginalValue) Then
                    If IsEmpty(l_rsTimeChargeC.Fields("EXCEPTION_CD").Value) Then
                        l_rsTimeChargeC!EXCEPTION_CD = ""
                    End If
                Else
                    If (l_rsTimeChargeC.Fields("EXCEPTION_CD").OriginalValue = _
                            l_rsTimeChargeC.Fields("EXCEPTION_CD").Value) Then
                        l_rsTimeChargeC!EXCEPTION_CD = ""
                    End If
                End If
                If IsEmpty(l_rsTimeChargeC.Fields("TRAVEL_PAY_HRS").OriginalValue) Then
                    If IsEmpty(l_rsTimeChargeC.Fields("TRAVEL_PAY_HRS").Value) Then
                        l_rsTimeChargeC!TRAVEL_PAY_HRS = CDec(-0.1)
                    End If
                Else
                    If (l_rsTimeChargeC.Fields("TRAVEL_PAY_HRS").OriginalValue = _
                            l_rsTimeChargeC.Fields("TRAVEL_PAY_HRS").Value) Then
                        l_rsTimeChargeC!TRAVEL_PAY_HRS = CDec(-0.1)
                    End If
                End If
            End If
        End If
        l_rsTimeChargeC.MoveNext
    Loop


xit:
    Set l_rsTimeChargeC = Nothing
    Exit Sub

eh:
    Set l_rsTimeChargeC = Nothing
    TrsRaiseError m_strCMODULENAME, "MarkAllTimeChargesForMissingComponents"
    
End Sub


Private Sub AddTimeChargeRow( _
                ByVal v_lngEmpNum As Long, _
                ByVal v_datDateWorked As Date, _
                ByVal v_strStartTime As String, _
                ByVal v_strEndTime As String, _
                ByRef r_typCanvType As CanvassedShift)
' This method is only executed for accepted responses where timecharges are generated.
    
    Dim l_datStartDateTime As Date
    Dim l_datEndDateTime As Date
    Dim l_decHours As Variant
    
On Error GoTo eh
    
    l_datStartDateTime = DateTimeFrom(v_datDateWorked, v_strStartTime)
    l_datEndDateTime = DateTimeFrom(v_datDateWorked, v_strEndTime)
    l_decHours = CDec(DateDiff("n", l_datStartDateTime, l_datEndDateTime) / 60)
    
    m_rsTimeCharges.AddNew
    m_rsTimeCharges!OT_OPP_NUM = m_lngOTOppNum
    m_rsTimeCharges!EMP_NUM = v_lngEmpNum
    m_rsTimeCharges!CHARGE_DATE = m_rsOTOpportunity!CHARGE_DATE
    m_rsTimeCharges!DATE_WORKED = v_datDateWorked
    m_rsTimeCharges!START_TIME = v_strStartTime
    m_rsTimeCharges!END_TIME = v_strEndTime
    m_rsTimeCharges!HOURS = l_decHours
    m_rsTimeCharges!EXCEPTION_CD = ""           'system
    m_rsTimeCharges!OVER_EXC_CD = ""
    m_rsTimeCharges!JOB_NUM = ""
    m_rsTimeCharges!COMP_SHIFT_CD = Left(m_rsOTOpportunity!CRAFT_JOB, 1)
    m_rsTimeCharges!OT_REASON_CD = m_rsOTOpportunity!OT_REASON_CD
    m_rsTimeCharges!WG_NUM = 0
    m_rsTimeCharges!DISTRIBUTION_ID = ""
    m_rsTimeCharges!DISTRIBUTION_TYPE = ""
    m_rsTimeCharges!JOB_NBR = ""
    m_rsTimeCharges!JOBTASK_NBR = ""
    m_rsTimeCharges!JOBTASK_STEP_NBR = 0
    m_rsTimeCharges!EMPRV_DATABASE_ID = ""
    m_rsTimeCharges!TRAVEL_PAY_HRS = CDec(-0.1)     'system
    m_rsTimeCharges!CREATED_EMP_NUM = m_lngUserNum
    m_rsTimeCharges!LAST_MOD_EMP_NUM = m_lngUserNum

    AddOvertimeSchedule m_rsTimeCharges
    AddAdjustedScheduleOvertime m_rsTimeCharges
    
    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "AddTimeChargeRow"
    
End Sub


Private Sub TraceActionMessage( _
                ByVal v_strLevel As String, _
                ByVal v_strActionLabel As String, _
                ByVal v_strMessage As String, _
                Optional ByVal v_blnNewLine As Boolean = False)

    TraceAction App.ThreadID, v_strLevel, v_strActionLabel, m_strUserID, vbCrLf & v_strMessage, , v_blnNewLine

End Sub


Private Function FormatTimeCharge(ByVal v_rsTimeCharge As ADODB.Recordset) As String

    FormatTimeCharge = _
        CStr(v_rsTimeCharge!OT_OPP_NUM) & vbTab & _
        CStr(v_rsTimeCharge!EMP_NUM) & vbTab & _
        Format(v_rsTimeCharge!CHARGE_DATE, "yyyy-mm-dd") & vbTab & _
        Format(v_rsTimeCharge!DATE_WORKED, "yyyy-mm-dd") & vbTab & _
        v_rsTimeCharge!START_TIME & vbTab & _
        v_rsTimeCharge!END_TIME & vbTab & _
        CStr(v_rsTimeCharge!HOURS) & vbTab & _
        v_rsTimeCharge!EXCEPTION_CD & vbTab

End Function


Private Function OTBanksCharged() As Boolean

    Dim l_obj As busOvertimeEqualizer.cbsOTEqualizer

    Dim stTime As Long
    Dim ndTime As Long
    Dim dfTime As Long

On Error GoTo eh

    #If TRTRACE Then
        stTime = GetTickCount()
    #End If

    Set l_obj = CreateObject("busOvertimeEqualizer.cbsOTEqualizer")

    OTBanksCharged = l_obj.ChargeOvertimeBanksForOpportunity( _
            m_rsOTOpportunity, m_rsCanvassSheet, Nothing, Nothing, _
            m_strUserID, m_rsTimeCharges, m_rsErrors)
    Set l_obj = Nothing

    #If TRTRACE Then
        ndTime = GetTickCount()
        dfTime = ndTime - stTime
        TraceAction App.ThreadID, "3", "OTBanksCharged", Format(Now, "mm/dd/yyyy hh:mm:ss"), _
            CStr(dfTime / 1000)
    #End If

    Exit Function

eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "OTBanksCharged"

End Function


Private Function OTBanksUnCharged() As Boolean

    Dim l_obj As busOvertimeEqualizer.cbsOTEqualizer

On Error GoTo eh

    Set l_obj = CreateObject("busOvertimeEqualizer.cbsOTEqualizer")

    m_rsOTOpportunity.Filter = ""
    m_rsOTOpportunity.MoveFirst

    OTBanksUnCharged = l_obj.UnChargeOvertimeBanksForOpportunity( _
            m_rsOTOpportunity, m_strUserID, m_rsErrors)
    Set l_obj = Nothing
    Exit Function

eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "OTBanksUnCharged"

End Function


Private Sub SetSystemGeneratedTimeCharge(ByVal v_lngEmpNum)


On Error GoTo eh

    If FindEmployeeValidation(v_lngEmpNum) Then
        m_rsEmpValidation!GENTIMECHARGE = "Y"
    End If
    
    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "SetSystemGeneratedTimeCharge"
End Sub


Private Function IsSystemGeneratedTimeCharge(ByVal v_lngEmpNum) As Boolean

On Error GoTo eh

    If FindEmployeeValidation(v_lngEmpNum) Then
        IsSystemGeneratedTimeCharge = m_rsEmpValidation!GENTIMECHARGE = "Y"
    End If
    
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "IsSystemGeneratedTimeCharge"
End Function


Private Function IsErrorAlreadyRecorded(ByVal v_lngEmpNum, ByVal v_strType As String) As Boolean

    Dim l_strValStat As String

On Error GoTo eh

    l_strValStat = GetEmployeeValidationStatus(v_lngEmpNum, v_strType)
    
    IsErrorAlreadyRecorded = l_strValStat = "Y"
    
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "IsErrorAlreadyRecorded"
End Function


Private Function GetEmployeeValidationStatus(ByVal v_lngEmpNum As Long, ByVal v_strType As String) As String

On Error GoTo eh

    If FindEmployeeValidation(v_lngEmpNum) Then
        Select Case v_strType
            Case "P"
                GetEmployeeValidationStatus = m_rsEmpValidation!OTPROCESS
            Case "D"
                GetEmployeeValidationStatus = m_rsEmpValidation!DETERMINEEXCCD
            Case "R"
                GetEmployeeValidationStatus = m_rsEmpValidation!RESPONSE
            Case "T"
                GetEmployeeValidationStatus = m_rsEmpValidation!TIMECHARGE
            Case "S"
                GetEmployeeValidationStatus = m_rsEmpValidation!SCHEDULE
        End Select
    End If
    
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "GetEmployeeValidationStatus"
End Function


Private Sub SetEmployeeValidationError(ByVal v_lngEmpNum As Long, ByVal v_strType As String)

On Error GoTo eh

    If FindEmployeeValidation(v_lngEmpNum) Then
        Select Case v_strType
            Case "P"
                m_rsEmpValidation!OTPROCESS = "Y"
            Case "D"
                m_rsEmpValidation!DETERMINEEXCCD = "Y"
            Case "R"
                m_rsEmpValidation!RESPONSE = "Y"
            Case "T"
                m_rsEmpValidation!TIMECHARGE = "Y"
            Case "S"
                m_rsEmpValidation!SCHEDULE = "Y"
        End Select
    End If
    
    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "SetEmployeeValidationError"
End Sub


Private Function FindEmployeeValidation(ByVal v_lngEmpNum) As Boolean

On Error GoTo eh

    If RecordsExist(m_rsEmpValidation) Then
        m_rsEmpValidation.MoveFirst
    Else
        AddEmployeeValidation v_lngEmpNum
        FindEmployeeValidation = True
        Exit Function
    End If

    m_rsEmpValidation.Find "EMP_NUM = " & v_lngEmpNum
    FindEmployeeValidation = Not m_rsEmpValidation.EOF
    
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "FindEmployeeValidation"
End Function


Private Sub AddEmployeeValidation(ByVal v_lngEmpNum As Long)

On Error GoTo eh

    m_rsEmpValidation.AddNew
    m_rsEmpValidation!EMP_NUM = v_lngEmpNum
    m_rsEmpValidation!GENTIMECHARGE = "N"
    m_rsEmpValidation!OTPROCESS = "N"
    m_rsEmpValidation!DETERMINEEXCCD = "N"
    m_rsEmpValidation!RESPONSE = "N"
    m_rsEmpValidation!TIMECHARGE = "N"
    m_rsEmpValidation!SCHEDULE = "N"
    
    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "AddEmployeeValidation"

End Sub


Private Function NewEmployeeValidationSet() As ADODB.Recordset
' Keep track of type of errors encountered.  Used to avoid validation error messages
' that result from errors during process.

    Dim l_rs As ADODB.Recordset

On Error GoTo eh

    Set l_rs = New ADODB.Recordset
    
    With l_rs
        .Fields.Append "EMP_NUM", adInteger, mCFLDATTRUPDATE
        .Fields.Append "GENTIMECHARGE", adChar, 1, mCFLDATTRUPDATE
        .Fields.Append "OTPROCESS", adChar, 1, mCFLDATTRUPDATE
        .Fields.Append "DETERMINEEXCCD", adChar, 1, mCFLDATTRUPDATE
        .Fields.Append "RESPONSE", adChar, 1, mCFLDATTRUPDATE
        .Fields.Append "TIMECHARGE", adChar, 1, mCFLDATTRUPDATE
        .Fields.Append "SCHEDULE", adChar, 1, mCFLDATTRUPDATE

        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open
    End With

    Set NewEmployeeValidationSet = l_rs
    
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "NewEmployeeValidationSet"
    
End Function


Private Sub AppendProcessTimeChargeError( _
                ByVal v_lngErrorNumber As Long, _
                ByVal v_strErrorDescription As String, _
                ByVal v_lngEmpNum As Long, _
                ByRef r_CanvassedShift As CanvassedShift, _
                ByVal v_rsTimeCharge As ADODB.Recordset)
                
    
On Error GoTo eh
   
    SetEmployeeValidationError v_lngEmpNum, "P"
   
    If v_lngErrorNumber = trErrbusOTConflictRegShiftExc Then
'            v_lngErrorNumber = trErrbusOTConflictRegShiftTravel Then
        If IsSystemGeneratedTimeCharge(v_lngEmpNum) Then
            If Not IsErrorAlreadyRecorded(v_lngEmpNum, "D") Then
                AddResponseError v_lngErrorNumber, mstrCBUSINESSERR, v_strErrorDescription, _
                    v_lngEmpNum
            End If
        Else
'            AddTimeChargeError v_lngErrorNumber, mstrCBUSINESSERR, v_strErrorDescription, _
'                v_rsTimeCharge
            AddTimeChargeError v_lngErrorNumber, mstrCBUSINESSERR, v_strErrorDescription, _
                v_rsTimeCharge!EMP_NUM, v_rsTimeCharge!CHARGE_DATE, v_rsTimeCharge!DATE_WORKED, _
                v_rsTimeCharge!START_TIME, v_rsTimeCharge!HOURS, v_rsTimeCharge!EXCEPTION_CD
        End If
        SetEmployeeValidationError v_lngEmpNum, "D"
    End If

    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "AppendProcessTimeChargeError"
    
End Sub


Private Sub AddResponseError( _
                ByVal v_lngErrorNumber As Long, _
                ByVal v_strErrorType As String, _
                ByVal v_strErrorDescription As String, _
                ByVal v_lngEmpNum As Long)
    
On Error GoTo eh

    If Not ResponseErrorAlreadyRecorded(v_lngErrorNumber, v_lngEmpNum) Then
        AppendBusinessError _
                m_rsErrors, _
                LongAppErr(v_lngErrorNumber), _
                v_strErrorType, _
                v_strErrorDescription, _
                m_strCMODULENAME, _
                TrType_OT_R, , , "EMP_NUM", _
                CStr(v_lngEmpNum)
    End If

    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "AddResponseError"
    
End Sub


Private Sub AddTimeChargeError( _
                ByVal v_lngErrorNumber As Long, _
                ByVal v_strErrorType As String, _
                ByVal v_strErrorDescription As String, _
                ByVal v_lngEmpNum As Long, _
                ByVal v_datChargeDate As Date, _
                ByVal v_datDateWorked As Date, _
                ByVal v_strStartTime As String, _
                ByVal v_decHours As Variant, _
                ByVal v_strExcCd As String)
    

On Error GoTo eh

    If Not TimeChargeErrorAlreadyRecorded(v_lngErrorNumber, v_lngEmpNum, v_datChargeDate, v_strStartTime) Then
        AppendBusinessError _
                m_rsErrors, _
                LongAppErr(v_lngErrorNumber), _
                v_strErrorType, _
                v_strErrorDescription, _
                m_strCMODULENAME, _
                TrType_OT_TC, , , "EMP_NUM;CHARGE_DATE;DATE_WORKED;START_TIME;HOURS;EXCEPTION_CD", _
                CStr(v_lngEmpNum), _
                Format(v_datChargeDate, "mm/dd/yyyy"), _
                Format(v_datDateWorked, "mm/dd/yyyy"), _
                v_strStartTime, _
                CStr(v_decHours), _
                v_strExcCd
    End If

    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "AddTimeChargeError"
    
End Sub


Private Function TimeChargeErrorAlreadyRecorded( _
                    ByVal v_lngErrorNumber As Long, _
                    ByVal v_lngEmpNum As Long, _
                    ByVal v_datChargeDate As Date, _
                    ByVal v_strStartTime As String) As Boolean


    If RecordsExist(m_rsErrors) Then
        m_rsErrors.MoveFirst
    End If

    Do Until m_rsErrors.EOF
        If m_rsErrors!ERR_NUMBER = LongAppErr(v_lngErrorNumber) Then
            If Trim(m_rsErrors!KEY_FIELD1) = CStr(v_lngEmpNum) And _
                    Trim(m_rsErrors!KEY_FIELD2) = Format(v_datChargeDate, "mm/dd/yyyy") And _
                    Trim(m_rsErrors!KEY_FIELD4) = v_strStartTime Then
                TimeChargeErrorAlreadyRecorded = True
            End If
        End If
        m_rsErrors.MoveNext
    Loop

        
End Function


Private Function ResponseErrorAlreadyRecorded( _
                    ByVal v_lngErrorNumber As Long, _
                    ByVal v_lngEmpNum As Long) As Boolean


    If RecordsExist(m_rsErrors) Then
        m_rsErrors.MoveFirst
    End If

    Do Until m_rsErrors.EOF
        If m_rsErrors!ERR_NUMBER = v_lngErrorNumber Then
            If Trim(m_rsErrors!KEY_FIELD1) = CStr(v_lngEmpNum) Then
                ResponseErrorAlreadyRecorded = True
                Exit Do
            End If
        End If
        m_rsErrors.MoveNext
    Loop

        
End Function


Private Sub TraceTemplate()

    Dim stTime As Long
    Dim ndTime As Long
    Dim dfTime As Long


    #If TRTRACE Then
        stTime = GetTickCount()
    #End If

'... app code

    #If TRTRACE Then
        ndTime = GetTickCount()
        dfTime = ndTime - stTime
        TraceAction App.ThreadID, "3", "GetCanvassState", Format(Now, "mm/dd/yyyy hh:mm:ss"), _
            CStr(dfTime / 1000)
    #End If

End Sub




Private Function IsOverrideThatAllowsOverlap( _
    ByVal v_strOverExc As String) As Boolean

    On Error GoTo eh

    Select Case Trim(v_strOverExc)
        Case "CS", "FR", "MRP", "LOT"
            IsOverrideThatAllowsOverlap = True
        Case Else
            IsOverrideThatAllowsOverlap = False
    End Select

    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "IsOverrideThatAllowsOverlap"

End Function


Private Sub DetermineGCMS( _
    ByVal v_rsSortDefinitions As ADODB.Recordset, _
    ByVal v_rsSortOrder As ADODB.Recordset, _
    ByVal v_rsForceSortOrder As ADODB.Recordset, _
    ByRef r_rsCanvassSheet As ADODB.Recordset, _
    ByRef r_blnHasPlantSnrty As Boolean, _
    ByRef r_blnGCMSEmpOnCanvass As Boolean, _
    ByRef r_blnGCMSEmpIn1stOrder As Boolean)

    Dim l_strFilter As String
    Dim l_blnHasPlantSnrty As Boolean
    Dim l_strSortCol As String
    Dim l_intSortCD As Integer
    
    
    On Error GoTo eh
    
    l_strFilter = "DEPT_NUM = '180' AND WG_SORT = 1"
    l_strFilter = MDAC25FilterFix(l_strFilter)
    r_rsCanvassSheet.Filter = l_strFilter
    If RecordsExist(r_rsCanvassSheet) Then
        r_blnGCMSEmpIn1stOrder = True
    Else
        r_blnGCMSEmpIn1stOrder = False
    End If
    
    l_strFilter = "DEPT_NUM = '180'"
    l_strFilter = MDAC25FilterFix(l_strFilter)
    r_rsCanvassSheet.Filter = l_strFilter
    If RecordsExist(r_rsCanvassSheet) Then
        r_blnGCMSEmpOnCanvass = True
    Else
        r_blnGCMSEmpOnCanvass = False
    End If
    
    r_rsCanvassSheet.Filter = ""
    
    l_blnHasPlantSnrty = False
    v_rsSortOrder.Filter = ""
    If RecordsExist(v_rsSortOrder) Then
        v_rsSortOrder.MoveFirst
        While Not v_rsSortOrder.EOF
            l_intSortCD = v_rsSortOrder!SORT_CD
            v_rsSortDefinitions.Filter = "SORT_CD = " & str(l_intSortCD)
            l_strSortCol = Trim(v_rsSortDefinitions!SORT_COL_NM)
            If Trim(l_strSortCol) = "PLANT_SNRTY_QTY" Then
                l_blnHasPlantSnrty = True
            End If
            v_rsSortOrder.MoveNext
        Wend
    End If
    
    If l_blnHasPlantSnrty Then
        r_blnHasPlantSnrty = l_blnHasPlantSnrty
        Exit Sub
    End If
    
    v_rsForceSortOrder.Filter = ""
    If RecordsExist(v_rsForceSortOrder) Then
        v_rsForceSortOrder.MoveFirst
        While Not v_rsForceSortOrder.EOF
            l_intSortCD = v_rsForceSortOrder!SORT_CD
            v_rsSortDefinitions.Filter = "SORT_CD = " & str(l_intSortCD)
            l_strSortCol = Trim(v_rsSortDefinitions!SORT_COL_NM)
            If Trim(l_strSortCol) = "PLANT_SNRTY_QTY" Then
                l_blnHasPlantSnrty = True
            End If
            v_rsForceSortOrder.MoveNext
        Wend
    End If
    
    r_blnHasPlantSnrty = l_blnHasPlantSnrty
      
    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "DetermineGCMS"

End Sub

Private Function GCMSModifySort( _
    ByRef r_rsCanvassSheet As ADODB.Recordset, _
    ByVal v_rsSortOrder As ADODB.Recordset, _
    ByVal v_rsForceSortOrder As ADODB.Recordset, _
    ByVal v_rsGroupSortOrder As ADODB.Recordset, _
    ByVal v_rsSortDefinitions As ADODB.Recordset) _
    As ADODB.Recordset

    Dim l_intEmpsOnCanvass As Integer
    Dim l_intEmpsInOrder As Integer
    Dim l_colOrderOfCalls As Collection
    Dim l_intCount As Integer
    Dim l_intWGSort As Integer
    Dim l_intI As Integer
    Dim l_rsEmpsCallOrder As ADODB.Recordset
    Dim l_rsCanvassSheet As ADODB.Recordset
    Dim l_strSort As String
    Dim l_strForceSort As String
    Dim l_blnPlantSnrtyInSort As Boolean
    Dim l_intSortPos As Integer
    Dim l_strGroupSort As String
    
    On Error GoTo eh
    
    
    r_rsCanvassSheet.Filter = ""
    If RecordsExist(r_rsCanvassSheet) Then
'        l_intEmpsOnCanvass = r_rsCanvassSheet.RecordCount
        l_intEmpsOnCanvass = 0
    Else
        Set GCMSModifySort = Nothing
        Exit Function
    End If
    
    l_strGroupSort = BuildSortString(v_rsGroupSortOrder, v_rsSortDefinitions)
    If Trim(l_strGroupSort) = "" Then
        l_strGroupSort = "WG_SORT"
    Else
        l_strGroupSort = l_strGroupSort & ",WG_SORT"
    End If
    Set l_colOrderOfCalls = GetOrderOfCalls(r_rsCanvassSheet)
    l_intCount = l_colOrderOfCalls.Count
    For l_intI = 1 To l_intCount
        l_intWGSort = l_colOrderOfCalls(l_intI)
        r_rsCanvassSheet.Filter = "WG_SORT = " & str(l_intWGSort)
        l_intEmpsInOrder = r_rsCanvassSheet.RecordCount
        Set l_rsEmpsCallOrder = CopyRecordset(r_rsCanvassSheet)
        l_rsEmpsCallOrder.Filter = "DEPT_NUM = '180'"
        If RecordsExist(l_rsEmpsCallOrder) Then
            l_strSort = BuildSortString(v_rsSortOrder, v_rsSortDefinitions, True)
            l_strForceSort = BuildSortString(v_rsForceSortOrder, v_rsSortDefinitions, True)
        Else
            l_strSort = BuildSortString(v_rsSortOrder, v_rsSortDefinitions)
            l_strForceSort = BuildSortString(v_rsForceSortOrder, v_rsSortDefinitions)
        End If
        l_strSort = l_strGroupSort & "," & l_strSort
        l_strForceSort = l_strGroupSort & "," & l_strForceSort
        l_rsEmpsCallOrder.Filter = ""
        l_intEmpsInOrder = l_rsEmpsCallOrder.RecordCount
        l_rsEmpsCallOrder.MoveFirst
        l_rsEmpsCallOrder.Sort = l_strSort
        l_rsEmpsCallOrder.MoveFirst
        While Not l_rsEmpsCallOrder.EOF
            l_intSortPos = l_rsEmpsCallOrder.AbsolutePosition + l_intEmpsOnCanvass
            l_rsEmpsCallOrder!ROW_NUM = l_intSortPos
            l_rsEmpsCallOrder.MoveNext
        Wend
        
        l_rsEmpsCallOrder.MoveFirst
        l_rsEmpsCallOrder.Sort = l_strForceSort
        l_rsEmpsCallOrder.MoveFirst
        While Not l_rsEmpsCallOrder.EOF
            l_intSortPos = l_rsEmpsCallOrder.AbsolutePosition + l_intEmpsOnCanvass
            l_rsEmpsCallOrder!FORCE_ROW_NUM = l_intSortPos
            l_rsEmpsCallOrder.MoveNext
        Wend
        l_rsEmpsCallOrder.MoveFirst
        l_intEmpsOnCanvass = l_intEmpsOnCanvass + l_intEmpsInOrder
        If l_rsCanvassSheet Is Nothing Then
            Set l_rsCanvassSheet = CopyRecordset(l_rsEmpsCallOrder)
        Else
            AppendRecordset l_rsEmpsCallOrder, l_rsCanvassSheet
        End If
    Next l_intI

    Set GCMSModifySort = l_rsCanvassSheet
    
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "GCMSModifySort"

End Function

Private Function GetOrderOfCalls( _
    ByVal v_rsCanvassSheet As ADODB.Recordset _
    ) As Collection
    
    Dim l_colOrdersofCall As Collection
    Dim l_intPrevWgSort As Integer
    Dim l_intWGSort As Integer
    
    On Error GoTo eh
    
    Set l_colOrdersofCall = New Collection
    v_rsCanvassSheet.Sort = "WG_SORT"
    v_rsCanvassSheet.MoveFirst
    While Not v_rsCanvassSheet.EOF
        l_intWGSort = v_rsCanvassSheet!WG_SORT
        If l_intWGSort <> l_intPrevWgSort Then
            l_colOrdersofCall.Add l_intWGSort
            l_intPrevWgSort = l_intWGSort
        End If
        v_rsCanvassSheet.MoveNext
    Wend
        
    v_rsCanvassSheet.Sort = ""
    v_rsCanvassSheet.Filter = ""
    
    Set GetOrderOfCalls = l_colOrdersofCall
    
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "GetOrderOfCalls"

End Function
'Begin Prem TRIS 5.6
Public Function CreateOvertimeCanvass2(ByRef rlOpportunity As Long _
                , ByVal vadrsCriteria As ADODB.Recordset _
                , ByVal vsUser As String _
                , ByRef radrsBRV As ADODB.Recordset _
                ) As Boolean
                
'JIMB
' 08/25/2009 -- Currently not used
' CreateOvertimeCanvass is the method called from the client
' fixed the bugs in ConstructCanvassSheet2 in cbsOTCanvasstrans
' and in ConstructCanvassSheet2 in cdOTCanvass
'
On Error GoTo CreateOvertimeCanvass2_EH
'RS
Dim ladrsTemplateBase As ADODB.Recordset
Dim ladrsSortOrder As ADODB.Recordset
Dim ladrsForceSortOrder As ADODB.Recordset
Dim ladrsGroupSortOrder As ADODB.Recordset
Dim ladrsSortDefinitions As ADODB.Recordset
Dim ladrsCanvassSheet As ADODB.Recordset
Dim ladrsOTOppInelig As ADODB.Recordset
Dim ladrsPhoneNumbers As ADODB.Recordset

'Other
Dim ldtOppChargeDate As Date, ldtCanvassDate As Date
Dim laryEmployees As Variant, llTemplateNum As Long
Dim llOTOppNum As Long, liOTGroup As Integer
Dim liCanvassLoc As Integer, lsOTReasonCd As String
Dim liBankCD As Integer, liRefuseRule As Integer

    CreateOvertimeCanvass2 = False

    'Initialize
    m_strUserID = vsUser
    ldtCanvassDate = Date

    Set m_rsErrors = CreateNewBusinessErrorSet()

    If Not (PreValidateOppCriteria(vadrsCriteria, 0)) Then
        GetObjectContext.SetAbort
        GoTo xit
    End If

    lsOTReasonCd = vadrsCriteria!OT_REASON_CD
    llTemplateNum = vadrsCriteria!TEMPLATE_NUM
    ldtOppChargeDate = vadrsCriteria!CHARGE_DATE
    liOTGroup = vadrsCriteria!OT_GROUP_NUM

    GetOvertimeTemplateData llTemplateNum _
        , ladrsTemplateBase _
        , ladrsSortOrder _
        , ladrsForceSortOrder _
        , ladrsGroupSortOrder _
        , ladrsSortDefinitions
    If Not (RecordsExist(ladrsTemplateBase)) Then
        'writeBRV
        GetObjectContext.SetAbort
        Exit Function
    End If

    Set ladrsCanvassSheet = ConstructCanvassSheet2(ldtOppChargeDate _
                            , liOTGroup _
                            , ladrsTemplateBase!TEMPL_LOCATION _
                            , liBankCD)
    SortCanvassSheet ladrsCanvassSheet _
        , ladrsSortOrder _
        , ladrsForceSortOrder _
        , ladrsGroupSortOrder _
        , ladrsSortDefinitions
    m_lngOTOppNum = GetNextOpportunityNumber()

    GetCanvassState

    liCanvassLoc = ladrsTemplateBase!TEMPL_LOCATION
    liRefuseRule = ladrsTemplateBase!REFUSE_RULE_NUM
    BuildOpportunity liCanvassLoc _
        , liRefuseRule _
        , vadrsCriteria

    Set ladrsPhoneNumbers = GetPeopleSoftPhoneNumber(laryEmployees)
    BuildCanvassSheet lsOTReasonCd _
        , ladrsCanvassSheet _
        , ladrsPhoneNumbers _
        , liBankCD _
        , ladrsOTOppInelig

    If (vadrsCriteria!PRINT_ONLY_IND = "Y") Then
        m_rsOTOpportunity!OPP_STATUS = "V"
        m_rsOTOpportunity!MOD_REASON_DESC = "PRINT ONLY"
    End If

    If (SaveOpportunity) Then
        'continue
    Else
        'raise error
        GetObjectContext.SetAbort
        Exit Function
    End If

    If (SaveCanvassSheet) Then
        'continue
    Else
        'raise error
        GetObjectContext.SetAbort
        Exit Function
    End If

    rlOpportunity = m_lngOTOppNum

    CreateOvertimeCanvass2 = True
    GetObjectContext.SetComplete
xit:
    If Not (m_rsErrors Is Nothing) Then
        If (m_rsErrors.RecordCount = 0) Then
            Set radrsBRV = Nothing
        Else
            Set radrsBRV = m_rsErrors
        End If
    Else
        Set radrsBRV = Nothing
    End If

    FreeGlobalRS

    Exit Function
CreateOvertimeCanvass2_EH:
    FreeGlobalRS

    GetObjectContext.SetAbort
    TrsRaiseError m_strCMODULENAME _
        , "CreateOvertimeCanvass2" _
        , , , , , True

End Function
Private Function ConstructCanvassSheet2(ByVal vdtChargeDate As Date _
                , ByVal viOTGroup As Integer _
                , ByVal viTemplateLocation As Integer _
                , ByRef riBankCD As Integer _
                ) As ADODB.Recordset
On Error GoTo ConstructCanvassSheet2_EH
'DB
Dim lobjDB As dbOTCanvass.cdOTCanvass
'RS
Dim ladrsSheet As ADODB.Recordset
Dim ladrsWGSort As ADODB.Recordset
'Other
Dim ldtOppChargeDate As Date
Dim lsCanvassType As String, lsBank As String
Dim liOTGroup As Integer, llArrEmployee As Variant

    Set ConstructCanvassSheet2 = Nothing

    GetCanvassBank viOTGroup _
        , lsBank _
        , riBankCD

    Set lobjDB = CreateObject("dbOTCanvass.cdOTCanvass")
    Set ladrsSheet = lobjDB.ConstructCanvassSheet2(viOTGroup _
                        , riBankCD _
                        , vdtChargeDate)
    If Not (RecordsExist(ladrsSheet)) Then
        'probably should raise error
        Set lobjDB = Nothing
        Exit Function
    End If

    llArrEmployee = DetermineEmployeesOnCanvass(ladrsSheet)
    Set ladrsWGSort = lobjDB.GetWorkGroupSortData( _
                        llArrEmployee _
                        , liOTGroup _
                        , vdtChargeDate)

    UpdateSortColumns ladrsSheet _
        , ladrsWGSort _
        , viTemplateLocation

    Set ConstructCanvassSheet2 = ladrsSheet
    Exit Function
ConstructCanvassSheet2_EH:
    Set lobjDB = Nothing
    TrsRaiseError m_strCMODULENAME, "ConstructCanvassSheet2"

End Function
'End Prem TRIS 5.6


Private Function NextRegularShiftStartTimeAfter( _
                    ByVal v_lngEmpNum As Long, _
                    ByVal v_datStartDateTime As Date, _
                    ByVal v_datEndDateTime As Date, _
                    ByRef r_blnIsTradedShift As Boolean) _
                As Date
'If regular shift not found, just answer 1 day prior to v_datStartDateTime.

    Dim l_rsEmpSchedC As ADODB.Recordset

On Error GoTo eh

    r_blnIsTradedShift = False
    
    Set l_rsEmpSchedC = m_rsEmpSchedules.Clone
    l_rsEmpSchedC.Filter = "EMP_NUM = " & v_lngEmpNum
    l_rsEmpSchedC.Sort = "DATE_WORKED DESC, START_TIME DESC"

    PositionRegularScheduleAfterDate v_datEndDateTime, l_rsEmpSchedC

    If l_rsEmpSchedC.EOF Then
        NextRegularShiftStartTimeAfter = DateAdd("d", -1, v_datStartDateTime)
    Else
        NextRegularShiftStartTimeAfter = _
            DateTimeFrom(l_rsEmpSchedC!SHIFT_START_DATE, l_rsEmpSchedC!SHIFT_START_TIME)
        r_blnIsTradedShift = IsTradedWorkShift(Trim(l_rsEmpSchedC!SYSTEM_EXC_CD))
    End If

    Set l_rsEmpSchedC = Nothing
    Exit Function

eh:
    Set l_rsEmpSchedC = Nothing
    TrsRaiseError m_strCMODULENAME, "NextRegularShiftStartTimeAfter"
    
End Function

Private Sub PositionRegularScheduleAfterDate( _
                    ByVal v_datEndDateTime As Date, _
                    ByRef r_rsEmpSched As ADODB.Recordset)
'Position r_rsEmpSched on first timespan found(in descending order), that is
'before the input datetime.

    Dim l_datSchedule As Date

On Error GoTo eh

    Do Until r_rsEmpSched.EOF
        If IsRegularScheduledShift(r_rsEmpSched) Then
            l_datSchedule = DateTimeFrom(r_rsEmpSched!DATE_WORKED, r_rsEmpSched!START_TIME)
'            If l_datSchedule <= v_datStartDateTime Then
            If DatetimeCompare(l_datSchedule, "GE", v_datEndDateTime) Then
                Exit Do
            End If
        End If
        r_rsEmpSched.MoveNext
    Loop

    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "PositionRegularScheduleAfterDate"
End Sub

Private Function IsJobTaskOTExcep( _
    ByVal v_strException As String _
    ) As Boolean
    
    Dim l_obj As dbRules.cdRules
    
    On Error GoTo eh
    
    Set l_obj = CreateObject("dbRules.cdRules")
    IsJobTaskOTExcep = l_obj.DoesRuleApply(143, trExceptionCode, v_strException)
    Exit Function
    
eh:
     TrsRaiseError m_strCMODULENAME, "IsJobTaskOTExcep"
    
End Function

Private Function GetActTimeEmployees( _
                    ByVal v_rsTimeCharge As ADODB.Recordset) As Collection
    Dim l_objEmpList As cuSet

On Error GoTo eh
    Set l_objEmpList = New cuSet
    If RecordsExist(v_rsTimeCharge) Then
        v_rsTimeCharge.MoveFirst
        Do Until v_rsTimeCharge.EOF
            l_objEmpList.Add v_rsTimeCharge.Fields("EMP_NUM").Value
            v_rsTimeCharge.MoveNext
        Loop
        v_rsTimeCharge.MoveFirst
    End If
    
    Set GetActTimeEmployees = l_objEmpList.Contents
    Set l_objEmpList = Nothing
    Exit Function
    
eh:
    Set l_objEmpList = Nothing
    TrsRaiseError m_strCMODULENAME, "GetPaidEmployees"

End Function

Private Function TimeChargeRuleAppliesByUnion( _
    ByVal v_strUnion As String _
    ) As Boolean
    
    Dim l_obj As dbRules.cdRules
    
    On Error GoTo eh
    
    Set l_obj = CreateObject("dbRules.cdRules")
    TimeChargeRuleAppliesByUnion = l_obj.DoesRuleApply(225, trUnionCode, v_strUnion)
    Exit Function
    
eh:
     TrsRaiseError m_strCMODULENAME, "TimeChargeRuleAppliesByUnion"
    
End Function

Private Function TimeChargeRuleAppliesByDepartment( _
    ByVal v_strDept As String _
    ) As Boolean
    
    Dim l_obj As dbRules.cdRules
    
    On Error GoTo eh
    
    Set l_obj = CreateObject("dbRules.cdRules")
    TimeChargeRuleAppliesByDepartment = l_obj.DoesRuleApply(226, trDepartmentNumber, v_strDept)
    Exit Function
    
eh:
     TrsRaiseError m_strCMODULENAME, "TimeChargeRuleAppliesByDepartment"
    
End Function

Private Function IsExecptionToTimeChargeRuleByLocation( _
    ByVal v_intLocation As Integer _
    ) As Boolean
    
    Dim l_obj As dbRules.cdRules
    
    On Error GoTo eh
    
    Set l_obj = CreateObject("dbRules.cdRules")
    IsExecptionToTimeChargeRuleByLocation = l_obj.DoesRuleApply(223, trLocationCode, v_intLocation)
    Exit Function
    
eh:
     TrsRaiseError m_strCMODULENAME, "IsExecptionToTimeChargeRuleByLocation"
    
End Function

Private Function IsExecptionToTimeChargeRuleByDepartment( _
    ByVal v_strDept As String _
    ) As Boolean
    
    Dim l_obj As dbRules.cdRules
    
    On Error GoTo eh
    
    Set l_obj = CreateObject("dbRules.cdRules")
    IsExecptionToTimeChargeRuleByDepartment = l_obj.DoesRuleApply(224, trDepartmentNumber, v_strDept)
    Exit Function
    
eh:
     TrsRaiseError m_strCMODULENAME, "IsExecptionToTimeChargeRuleByDepartment"
    
End Function

Private Function EmployeeDepartment( _
                    ByVal v_lngEmpNum As Long, _
                    ByVal v_datChargeDate As Date, _
                    ByVal v_rsEmployees As ADODB.Recordset) _
                As String
    
    Dim l_obj As busEmployee.cbsEmployee

On Error GoTo eh

    v_rsEmployees.Filter = "EMP_NUM = " & v_lngEmpNum
    Set l_obj = CreateObject("busEmployee.cbsEmployee")
    EmployeeDepartment = _
        l_obj.EmployeeDepartment( _
            v_rsEmployees!DEPT_EFF_DATE, _
            v_rsEmployees!DEPT_NUM, _
            v_rsEmployees!PREV_DEPT_NUM, _
            v_datChargeDate)

    v_rsEmployees.Filter = adFilterNone
    Set l_obj = Nothing
    Exit Function

eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "EmployeeDepartment"

End Function

Private Function DoesTimeChargeRuleApply( _
    ByVal v_strUnion As String, _
    ByVal v_strDept As String, _
    ByVal v_intLoc As Integer _
    ) As Boolean
    
    Dim l_blnApplies As Boolean
    
    On Error GoTo eh
    
    l_blnApplies = False
    If TimeChargeRuleAppliesByUnion(v_strUnion) Or TimeChargeRuleAppliesByDepartment(v_strDept) Then
        If IsExecptionToTimeChargeRuleByLocation(v_intLoc) Or IsExecptionToTimeChargeRuleByDepartment(v_strDept) Then
            l_blnApplies = False
        Else
            l_blnApplies = True
        End If
    Else
        l_blnApplies = False
    End If

    DoesTimeChargeRuleApply = l_blnApplies
    Exit Function
    
eh:
     TrsRaiseError m_strCMODULENAME, "DoesTimeChargeRuleApply"
    
End Function

Private Function IsChemTesterCanvass(ByVal v_TemplateNum As Long) As Boolean

    Dim l_obj As dbRules.cdRules

    On Error GoTo eh
    
    Set l_obj = CreateObject("dbRules.cdRules")
    IsChemTesterCanvass = l_obj.DoesRuleApply(263, trLocationCode, v_TemplateNum)
    
    Exit Function
    
eh:
     TrsRaiseError m_strCMODULENAME, "IsChemTesterCanvass"
End Function

Private Function IsSpecialOverrideCode( _
    ByVal v_strOverExc As String) As Boolean

    Dim l_obj As dbRules.cdRules

    On Error GoTo eh
    
    Set l_obj = CreateObject("dbRules.cdRules")
    IsSpecialOverrideCode = l_obj.DoesRuleApply(280, trExceptionCode, v_strOverExc)


    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "IsSpecialOverrideCode"

End Function

