VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "cbsEmployeeTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'AMEREN Project: TRIS Conversion Phase 3
'February 2004
'Author: Prem M. Chander
'Module Description: This class will provide the user the ability
'to create/modify employee skills and qualifications

'COMMENTS: OCT, 2005 - CHANGE GETOBJECTCONTEXT.CREATEINSTANCE
'TO CREATEOBJECT -- PREM

Option Explicit

Private Const ADOFIELDATTR = adFldUpdatable Or adFldIsNullable
Private Const MODULENAME As String = "busEmployee.cbsEmployeeTrans"

Private Const trCOMPANY As Integer = 1
Private Const trDEPARTMENT As Integer = 2
Private Const trJOB As Integer = 3
Private Const trQUALIFICATIONS As Integer = 4
Private Const trSPECIALSKILL As Integer = 5
Private Const trUNION As Integer = 6
Private Const trOTBANKCODE As Integer = 7
Private Const trLOCATION As Integer = 8
Private Const trIRAREA As Integer = 9
Private Const trIRDIVISION As Integer = 10
Private Const trCRAFT As Integer = 11
Private Const trEXCEPTION As Integer = 12
Private Const trEMPLOYEE As Integer = 13
Private Const trOTGROUP As Integer = 14
Private Const trWORKGROUP As Integer = 15

'JIMB added these for Volunteer processing
Private m_rsVolunteerList As ADODB.Recordset
Private m_rsVolunteerChanges As ADODB.Recordset
'This method will save employee skills and qualifications entries
Public Function ProcessEmpQualsSkillsEntries( _
                    ByVal vadrsQAndSInformation As ADODB.Recordset _
                    , ByRef radrsErrors As ADODB.Recordset _
                    , ByVal vlUserEmployeeNumber As Long _
                    ) As Boolean
On Error GoTo ProcessEmpQualsSkillsEntries_EH

Dim lbEmptyRequest As Boolean, lsFilterString As String
Dim lobjdbEmployee As dbEmployee.cdEmployeeTrans
Dim llRecordIndex As Long

    'Check for request existence
    lbEmptyRequest = True
    If Not (vadrsQAndSInformation Is Nothing) Then
        If Not (vadrsQAndSInformation.BOF And _
            vadrsQAndSInformation.EOF) Then
            lbEmptyRequest = False
        End If
    End If

    'Prepare BRV Recordset
    Set radrsErrors = CreateQAndSBRV

    If (lbEmptyRequest) Then
        AppendQAndSBRV radrsErrors _
            , trErrbusEmployeeQualsSkills _
            , mstrCBUSINESSERR _
            , "There are no Employee Qualification And Special Skills entries to process." _
            , MODULENAME & ": ProcessEmpQualsSkillsEntries"
        GoTo REPORTBRV
    Else
        With vadrsQAndSInformation
            .Filter = adFilterNone
            If Not (.EOF) Then
                .MoveFirst
                If Not (ValidateQAndSEntries( _
                    vadrsQAndSInformation _
                    , radrsErrors _
                    , vlUserEmployeeNumber)) Then
                    GoTo REPORTBRV
                End If

                .Filter = adFilterNone
                .MoveFirst
            End If
        End With
    End If

    Set lobjdbEmployee = CreateObject( _
                            "dbEmployee.cdEmployeeTrans")
    'Set lobjdbEmployee = New dbEmployee.cdEmployee
    If Not (lobjdbEmployee.dbProcessQualsAndSkillsSave( _
            vadrsQAndSInformation _
            , vlUserEmployeeNumber)) Then
        Set lobjdbEmployee = Nothing
        Err.Raise trErrbusEmployeeQualsSkills
    End If
    Set lobjdbEmployee = Nothing

    Set vadrsQAndSInformation = Nothing
    Set radrsErrors = Nothing

    GetObjectContext().SetComplete

    ProcessEmpQualsSkillsEntries = True

    Set lobjdbEmployee = Nothing

    Exit Function
ProcessEmpQualsSkillsEntries_EH:
    'Abort Transaction
    GetObjectContext().SetAbort

    Set lobjdbEmployee = Nothing
    Set vadrsQAndSInformation = Nothing
    Set radrsErrors = Nothing

    ProcessEmpQualsSkillsEntries = False

    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "ProcessEmpQualsSkillsEntries"
            Case trErrdbInsIntoEmpQualSkill, trErrdbUpdEmpQualSkill _
                , trErrdbDelFromEmpQualSkill, trErrdbSaveEmpQualSkill _
                , trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "ProcessEmpQualsSkillsEntries" _
                    , Err.Number _
                    , , , , _
                    True
            Case Else
                TrsRaiseError MODULENAME _
                    , "ProcessEmpQualsSkillsEntries" _
                    , trErrbusEmployeeQualsSkills _
                    , , , , _
                    True
    End Select

    Exit Function

REPORTBRV:
    Set lobjdbEmployee = Nothing
    Set vadrsQAndSInformation = Nothing

    If Not (radrsErrors Is Nothing) Then
        With radrsErrors
            If Not (.EOF And .BOF) Then
                .MoveFirst
                .Filter = adFilterNone

                'Check for errors
'MDAC25 FILTER FIX
                lsFilterString = "ERR_TYPE = 'BE'"
                .Filter = stdmTRISCommon.MDAC25FilterFix(lsFilterString)
                lsFilterString = ""
                'If found abort else complete
                If Not (.EOF) Then
                    'Abort Transaction
                    GetObjectContext().SetAbort
                    ProcessEmpQualsSkillsEntries = False
                Else
                    'Complete Transaction
                    GetObjectContext().SetComplete
                    ProcessEmpQualsSkillsEntries = True
                End If

                .Filter = adFilterNone
            Else
                'Complete Transaction
                GetObjectContext().SetComplete
                ProcessEmpQualsSkillsEntries = True
            End If
        End With
    Else
        'Complete Transaction
        GetObjectContext().SetComplete
        ProcessEmpQualsSkillsEntries = True
    End If

    Exit Function

End Function
Private Function CreateQAndSBRV() As ADODB.Recordset
On Error GoTo CreateQAndSBRV_EH
Dim ladrs As New ADODB.Recordset

    With ladrs
        .Fields.Append "EMP_NUM", adInteger, , ADOFIELDATTR
        .Fields.Append "QUAL_CD", adChar, 3, ADOFIELDATTR
        .Fields.Append "SPEC_SKILL_CD", adChar, 3, ADOFIELDATTR
        .Fields.Append "EFF_START_DATE", adDate, , ADOFIELDATTR
        .Fields.Append "EFF_END_DATE", adDate, , ADOFIELDATTR

        .Fields.Append "REC_SOURCE", adInteger, , ADOFIELDATTR
        .Fields.Append "ERR_NUMBER", adInteger, , ADOFIELDATTR
        .Fields.Append "ERR_TYPE", adChar, 2, ADOFIELDATTR
        .Fields.Append "ERR_DESCRIPTION", adChar, 300, ADOFIELDATTR
        .Fields.Append "ERR_SOURCE", adChar, 1000, ADOFIELDATTR

        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open
    End With

    Set CreateQAndSBRV = ladrs
    Set ladrs = Nothing

    Exit Function
CreateQAndSBRV_EH:
    Set ladrs = Nothing

    TrsRaiseError MODULENAME _
        , "CreateQAndSBRV"

End Function
Private Function ValidateQAndSEntries( _
                    ByRef radrsRSToValidate As ADODB.Recordset _
                    , ByRef radrsBRV As ADODB.Recordset _
                    , ByVal vlUserEmployeeNumber As Long _
                    ) As Boolean
On Error GoTo ValidateQAndSEntries_EH
Dim llRecordIndex As Long, lsFilterString As String

    ValidateQAndSEntries = False

    If Not (radrsRSToValidate Is Nothing) Then
        If Not (radrsRSToValidate.EOF And _
            radrsRSToValidate.BOF) Then
            radrsRSToValidate.Filter = adFilterNone
            radrsRSToValidate.MoveFirst
        Else
            Exit Function
        End If
    Else
        Exit Function
    End If

    If Not (stdmTRISCommon.CheckForValidLong(vlUserEmployeeNumber _
        , True)) Then
        AppendQAndSBRV radrsBRV _
            , trErrbusEmployeeQualsSkills _
            , mstrCBUSINESSERR _
            , "Invalid TRIS User." _
            , MODULENAME & ": ValidateQAndSEntries"
    End If

    With radrsRSToValidate
        For llRecordIndex = 1 To .RecordCount
            Select Case (UCase(Trim(!RECORD_STATUS)))
                Case "I", "C"
                    If Not (stdmTRISCommon.CheckForValidLong(!EMP_NUM _
                        , True)) Then
                        AppendQAndSBRV radrsBRV _
                            , trErrbusEmployeeQualsSkills _
                            , mstrCBUSINESSERR _
                            , "Invalid Employee." _
                            , MODULENAME & ": ValidateQAndSEntries" _
                            , _
                            , !QUAL_CD _
                            , !SPEC_SKILL_CD _
                            , !EFF_START_DATE _
                            , !EFF_END_DATE
                    Else
                        If Not (CheckForItemExistenceInDB(trEMPLOYEE, , !EMP_NUM)) Then
                            AppendQAndSBRV radrsBRV _
                                , trErrbusEmployeeQualsSkills _
                                , mstrCBUSINESSERR _
                                , "Invalid Employee. Employee " & Trim(!EMP_NUM) & " does not exist." _
                                , MODULENAME & ": ValidateQAndSEntries" _
                                , !EMP_NUM _
                                , !QUAL_CD _
                                , !SPEC_SKILL_CD _
                                , !EFF_START_DATE _
                                , !EFF_END_DATE
                        End If
                    End If

                    If Not (stdmTRISCommon.CheckForValidString(!QUAL_CD)) Then
                        AppendQAndSBRV radrsBRV _
                            , trErrbusEmployeeQualsSkills _
                            , mstrCBUSINESSERR _
                            , "Invalid Qualification." _
                            , MODULENAME & ": ValidateQAndSEntries" _
                            , !EMP_NUM _
                            , _
                            , !SPEC_SKILL_CD _
                            , !EFF_START_DATE _
                            , !EFF_END_DATE
                    Else
                        If Not (CheckForItemExistenceInDB(trQUALIFICATIONS, !QUAL_CD)) Then
                            AppendQAndSBRV radrsBRV _
                                , trErrbusEmployeeQualsSkills _
                                , mstrCBUSINESSERR _
                                , "Invalid Qualification. Qualification Code " & Trim(!QUAL_CD) & " does not exist." _
                                , MODULENAME & ": ValidateQAndSEntries" _
                                , !EMP_NUM _
                                , !QUAL_CD _
                                , !SPEC_SKILL_CD _
                                , !EFF_START_DATE _
                                , !EFF_END_DATE
                        End If
                    End If

                    If (stdmTRISCommon.CheckForValidString(!SPEC_SKILL_CD)) Then
                        If Not (CheckForItemExistenceInDB(trSPECIALSKILL _
                            , !SPEC_SKILL_CD, , , !QUAL_CD)) Then
                            AppendQAndSBRV radrsBRV _
                                , trErrbusEmployeeQualsSkills _
                                , mstrCBUSINESSERR _
                                , "Invalid Special Skill. Special Skill code " & Trim(!SPEC_SKILL_CD) & " does not exist." _
                                , MODULENAME & ": ValidateQAndSEntries" _
                                , !EMP_NUM _
                                , !QUAL_CD _
                                , !SPEC_SKILL_CD _
                                , !EFF_START_DATE _
                                , !EFF_END_DATE
                        End If
                    Else
                        !SPEC_SKILL_CD = ""
                        .Update
                    End If

                    If Not (CheckForValidDate(!EFF_START_DATE)) Then
                        AppendQAndSBRV radrsBRV _
                            , trErrbusEmployeeQualsSkills _
                            , mstrCBUSINESSERR _
                            , "Invalid Start Date." _
                            , MODULENAME & ": ValidateQAndSEntries" _
                            , !EMP_NUM _
                            , !QUAL_CD _
                            , !SPEC_SKILL_CD _
                            , _
                            , !EFF_END_DATE
                    End If

                    If (IsNull(!EFF_END_DATE)) Then
                        !EFF_END_DATE = #12/31/9999#
                        .Update
                    Else
                        If Not (IsDate(!EFF_END_DATE)) Then
                            AppendQAndSBRV radrsBRV _
                                , trErrbusEmployeeQualsSkills _
                                , mstrCBUSINESSERR _
                                , "Invalid End Date." _
                                , MODULENAME & ": ValidateQAndSEntries" _
                                , !EMP_NUM _
                                , !QUAL_CD _
                                , !SPEC_SKILL_CD _
                                , !EFF_START_DATE
                        End If
                    End If
            End Select

            Select Case (UCase(Trim(!RECORD_STATUS)))
                Case "C", "D"
                    If Not (stdmTRISCommon.CheckForValidLong(!KEY_EMP_NUM, True)) Then
                        AppendQAndSBRV radrsBRV _
                            , trErrbusEmployeeQualsSkills _
                            , mstrCBUSINESSERR _
                            , "Invalid Delete/Change request. Employee Number is invalid." _
                            , MODULENAME & ": ValidateQAndSEntries" _
                            , !EMP_NUM _
                            , !QUAL_CD _
                            , !SPEC_SKILL_CD _
                            , !EFF_START_DATE _
                            , !EFF_END_DATE
                    End If

                    If Not (stdmTRISCommon.CheckForValidString(!KEY_QUAL_CD)) Then
                        AppendQAndSBRV radrsBRV _
                            , trErrbusEmployeeQualsSkills _
                            , mstrCBUSINESSERR _
                            , "Invalid Delete/Change request. Qualification code is invalid." _
                            , MODULENAME & ": ValidateQAndSEntries" _
                            , !EMP_NUM _
                            , !QUAL_CD _
                            , !SPEC_SKILL_CD _
                            , !EFF_START_DATE _
                            , !EFF_END_DATE
                    End If

                    If Not (CheckForValidDate(!KEY_EFF_START_DATE)) Then
                        AppendQAndSBRV radrsBRV _
                            , trErrbusEmployeeQualsSkills _
                            , mstrCBUSINESSERR _
                            , "Invalid Delete/Change request. Effective Start Date is invalid." _
                            , MODULENAME & ": ValidateQAndSEntries" _
                            , !EMP_NUM _
                            , !QUAL_CD _
                            , !SPEC_SKILL_CD _
                            , !EFF_START_DATE _
                            , !EFF_END_DATE
                    End If

                    If Not (stdmTRISCommon.CheckForValidString(!TIME_STAMP _
                        , , , True, 26)) Then
                        AppendQAndSBRV radrsBRV _
                            , trErrbusEmployeeQualsSkills _
                            , mstrCBUSINESSERR _
                            , "Invalid Delete/Change request. Time Stamp is invalid." _
                            , MODULENAME & ": ValidateQAndSEntries" _
                            , !EMP_NUM _
                            , !QUAL_CD _
                            , !SPEC_SKILL_CD _
                            , !EFF_START_DATE _
                            , !EFF_END_DATE
                    End If
            End Select

            .MoveNext
        Next llRecordIndex
    End With

    radrsRSToValidate.Filter = adFilterNone
    radrsRSToValidate.MoveFirst

    ValidateQAndSEntries = True

    If Not (radrsBRV Is Nothing) Then
        With radrsBRV
            If Not (.BOF And .EOF) Then
                .Filter = adFilterNone
                .MoveFirst
                lsFilterString = "ERR_TYPE = 'BE' OR ERR_TYPE = 'AE' OR ERR_TYPE = 'AW'"
                .Filter = stdmTRISCommon.MDAC25FilterFix(lsFilterString)
                lsFilterString = ""
                If Not (.EOF) Then
                    ValidateQAndSEntries = False
                End If

                .Filter = adFilterNone
            End If
        End With
    End If

    Exit Function
ValidateQAndSEntries_EH:
    ValidateQAndSEntries = False

    TrsRaiseError MODULENAME _
            , "ValidateQAndSEntries"

End Function
Private Sub AppendQAndSBRV(ByRef radrsBRV As ADODB.Recordset _
                , ByVal vlErrorNumber As Long _
                , ByVal vsErrorType As String _
                , ByVal vsErrorDescription As String _
                , ByVal vsErrorSource As String _
                , Optional ByVal vlEmployeeNumber As Long = 0 _
                , Optional ByVal vsQualificationCode As String = "" _
                , Optional ByVal vsSpecialSkillCode As String = "" _
                , Optional ByVal vdtEffectiveStartDate As Date = #1/1/100# _
                , Optional ByVal vdtEffectiveEndDate As Date = #1/1/100# _
                , Optional ByVal viRecordSource As Integer = 0 _
                )
On Error GoTo AppendBRV_EH

    With radrsBRV
        If (.State <> adStateOpen) Then
            Exit Sub
        End If

        .AddNew

        !EMP_NUM = vlEmployeeNumber
        !QUAL_CD = vsQualificationCode
        !SPEC_SKILL_CD = vsSpecialSkillCode
        !EFF_START_DATE = vdtEffectiveStartDate
        !EFF_END_DATE = vdtEffectiveEndDate

        !ERR_NUMBER = vlErrorNumber - vbObjectError
        !ERR_TYPE = Trim(vsErrorType)
        !REC_SOURCE = viRecordSource
        !ERR_DESCRIPTION = Trim(vsErrorDescription)
        !ERR_SOURCE = Trim(vsErrorSource)

        .Update
    End With

    Exit Sub
AppendBRV_EH:
    TrsRaiseError MODULENAME _
        , "AppendBRV"

End Sub
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'This method will save employee seniority and frozen seniority entries
Public Function ProcessSeniorityEntries( _
                    ByVal vadrsSeniorityInformation As ADODB.Recordset _
                    , ByVal vadrsFrozenSeniorityInformation As ADODB.Recordset _
                    , ByRef radrsErrors As ADODB.Recordset _
                    , ByVal vlSeniorityEmployeeNumber As Long _
                    , ByVal vlUserEmployeeNumber As Long _
                    ) As Boolean
On Error GoTo ProcessSeniorityEntries_EH
Dim lbEmptySeniorityRequest As Boolean, lbEmptyFrozenSeniorityRequest As Boolean
Dim lsFilterString As String, llRecordIndex As Long
'For validation/processing
Dim lobjdbEmployee As dbEmployee.cdEmployeeTrans
Dim lobjdbEmployee2 As dbEmployee.cdEmployee
Dim lobjbusEmployee As busEmployee.cbsEmployee
'Dim lobjRules As dbRules.cdRules
'Employee Type (Genco Vs. UEC, AMS etc.)
Dim ladrsEmployeeInformation As ADODB.Recordset, lbIsAGencoEmployee As Boolean

    'Check for request existence
    lbEmptySeniorityRequest = True
    If Not (vadrsSeniorityInformation Is Nothing) Then
        If Not (vadrsSeniorityInformation.BOF And _
            vadrsSeniorityInformation.EOF) Then
            lbEmptySeniorityRequest = False
        End If
    End If

    lbEmptyFrozenSeniorityRequest = True
    If Not (vadrsFrozenSeniorityInformation Is Nothing) Then
        If Not (vadrsFrozenSeniorityInformation.BOF And _
            vadrsFrozenSeniorityInformation.EOF) Then
            lbEmptyFrozenSeniorityRequest = False
        End If
    End If

    'Prepare BRV Recordset
    Set radrsErrors = CreateSenioritiesBRV

    If (lbEmptySeniorityRequest And _
        lbEmptyFrozenSeniorityRequest) Then
        AppendSeniorityBRV radrsErrors _
            , trErrbusEmployeeSeniority _
            , mstrCBUSINESSERR _
            , "There are no Employee Seniority/Frozen Seniority entries to process." _
            , MODULENAME & ": ProcessSeniorityEntries"
        GoTo REPORTBRV
    Else
        If Not (CheckForValidLong(vlUserEmployeeNumber, True)) Then
            AppendSeniorityBRV radrsErrors _
                , trErrbusEmployeeSeniority _
                , mstrCBUSINESSERR _
                , "Invalid TRIS User." _
                , MODULENAME & ": ValidateSeniorityEntries"
            GoTo REPORTBRV
        End If

        'Note: The reason this has been moved out of the Validate Function
        'is because of the IsAGencoEmployee Boolean variable required
        'for further processing
        If Not (CheckForValidLong(vlSeniorityEmployeeNumber, True)) Then
            AppendSeniorityBRV radrsErrors _
                , trErrbusEmployeeSeniority _
                , mstrCBUSINESSERR _
                , "Invalid Employee Number." _
                , MODULENAME & ": ValidateDataAccessEntries"
            GoTo REPORTBRV
        End If

        If Not (lbEmptySeniorityRequest) Then
            'Get basic employee information
            Set lobjdbEmployee2 = CreateObject("dbEmployee.cdEmployee")
            'Set lobjdbEmployee = New dbEmployee.cdEmployee
            Set ladrsEmployeeInformation = lobjdbEmployee2.GetBasicEmployeeInformation(vlSeniorityEmployeeNumber)
            Set lobjdbEmployee2 = Nothing

            'check to see if the information exists
            If (mADOExtensions.RecordsExist(ladrsEmployeeInformation)) Then
                With ladrsEmployeeInformation
                    'Check to see if the Employee's Company Code follows
                    'AmerenCIPS rules (CIP, GEN, IHC, AED, IMS)
                    'the Rule Code is 11, the Business Type is Company
                    'the value is returned by Employee Company
                    'lbIsAGencoEmployee = False
'TODO: What should the charge date be?
                    Set lobjbusEmployee = CreateObject("busEmployee.cbsEmployee")
                    Dim lsUnionCd As String
                    lsUnionCd = lobjbusEmployee.EmployeeUnion(!JOB_EFF_DATE, !UNION_CD, !PRIOR_UNION_CD)
                    lbIsAGencoEmployee = lobjbusEmployee.IsAGencoEmployee(EmployeeCompany(!DEPT_EFF_DATE, !DEPT_NUM, !PREV_DEPT_NUM), _
                                                                        lsUnionCd, !PAYROLL_GROUP_CD, Date)
                    Set lobjbusEmployee = Nothing
'                    Set lobjRules = CreateObject("dbRules.cdRules")
'                    lbIsAGencoEmployee = _
'                        lobjRules.DoesRuleApply(11 _
'                                            , trCompanyCode _
'                                            , EmployeeCompany(!DEPT_EFF_DATE _
'                                                , !DEPT_NUM _
'                                                , !PREV_DEPT_NUM))
'                    Set lobjRules = Nothing
                End With
                Set ladrsEmployeeInformation = Nothing
            End If

            With vadrsSeniorityInformation
                .Filter = adFilterNone
                If Not (.EOF) Then
                    .MoveFirst
                    If Not (ValidateSeniorityEntries( _
                        vadrsSeniorityInformation _
                        , radrsErrors _
                        , lbIsAGencoEmployee)) Then
                        GoTo REPORTBRV
                    End If

                    .Filter = adFilterNone
                    .MoveFirst
                End If
            End With
        End If

        If Not (lbEmptyFrozenSeniorityRequest) Then
            With vadrsFrozenSeniorityInformation
                .Filter = adFilterNone
                If Not (.EOF) Then
                    .MoveFirst
                    If Not (ValidateFrozenSeniorityEntries( _
                        vadrsFrozenSeniorityInformation _
                        , radrsErrors)) Then
                        GoTo REPORTBRV
                    End If

                    .Filter = adFilterNone
                    .MoveFirst
                End If
            End With
        End If
    End If
    
    If Not (lbEmptySeniorityRequest) Then
        Set lobjdbEmployee = CreateObject( _
                                "dbEmployee.cdEmployeeTrans")
        'Set lobjdbEmployee = New dbEmployee.cdEmployee
        If Not (lobjdbEmployee.dbProcessSeniorityEntries( _
                vadrsSeniorityInformation _
                , lbIsAGencoEmployee _
                , vlSeniorityEmployeeNumber _
                , vlUserEmployeeNumber)) Then
            Set lobjdbEmployee = Nothing
            Err.Raise trErrbusEmployeeSeniority
        End If

        Set lobjdbEmployee = Nothing
    End If
    
    If Not (lbEmptyFrozenSeniorityRequest) Then
        Set lobjdbEmployee = CreateObject( _
                                "dbEmployee.cdEmployeeTrans")
        'Set lobjdbEmployee = New dbEmployee.cdEmployee
        If Not (lobjdbEmployee.dbProcessFrozenSeniorityEntries( _
                vadrsFrozenSeniorityInformation _
                , vlSeniorityEmployeeNumber _
                , vlUserEmployeeNumber)) Then
            Set lobjdbEmployee = Nothing
            Err.Raise trErrbusEmployeeSeniority
        End If

        Set lobjdbEmployee = Nothing
    End If

    Set vadrsSeniorityInformation = Nothing
    Set vadrsFrozenSeniorityInformation = Nothing
'    Set radrsErrors = Nothing

    GetObjectContext().SetComplete

    ProcessSeniorityEntries = True

    Set lobjdbEmployee = Nothing

    Exit Function
ProcessSeniorityEntries_EH:
    'Abort Transaction
    GetObjectContext().SetAbort

    Set lobjdbEmployee = Nothing
    Set lobjbusEmployee = Nothing
    'Set lobjRules = Nothing

    Set ladrsEmployeeInformation = Nothing
    Set vadrsSeniorityInformation = Nothing
    Set vadrsFrozenSeniorityInformation = Nothing
    Set radrsErrors = Nothing

    ProcessSeniorityEntries = False

    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "ProcessSeniorityEntries"
            Case trErrdbInsIntoEmpSnrty, trErrdbUpdEmpSnrty _
                , trErrdbDelFromEmpSnrty, trErrdbSaveEmpEmpSnrty _
                , trErrdbInsIntoEmpFrznSnrty, trErrdbUpdEmpFrznSnrty _
                , trErrdbDelFromEmpFrznSnrty, trErrdbSaveEmpFrznSnrty _
                , trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "ProcessSeniorityEntries" _
                    , Err.Number _
                    , , , , _
                    True
            Case Else
                TrsRaiseError MODULENAME _
                    , "ProcessSeniorityEntries" _
                    , trErrbusEmployeeSeniority _
                    , , , , _
                    True
    End Select

    Exit Function

REPORTBRV:
    Set ladrsEmployeeInformation = Nothing
    Set lobjdbEmployee = Nothing
    Set vadrsSeniorityInformation = Nothing
    Set vadrsFrozenSeniorityInformation = Nothing

    If Not (radrsErrors Is Nothing) Then
        With radrsErrors
            If Not (.EOF And .BOF) Then
                .MoveFirst
                .Filter = adFilterNone

                'Check for errors
'MDAC25 FILTER FIX
                lsFilterString = "ERR_TYPE = 'BE'"
                .Filter = stdmTRISCommon.MDAC25FilterFix(lsFilterString)
                lsFilterString = ""
                'If found abort else complete
                If Not (.EOF) Then
                    'Abort Transaction
                    GetObjectContext().SetAbort
                    ProcessSeniorityEntries = False
                Else
                    'Complete Transaction
                    GetObjectContext().SetComplete
                    ProcessSeniorityEntries = True
                End If

                .Filter = adFilterNone
            Else
                'Complete Transaction
                GetObjectContext().SetComplete
                ProcessSeniorityEntries = True
            End If
        End With
    Else
        'Complete Transaction
        GetObjectContext().SetComplete
        ProcessSeniorityEntries = True
    End If

    Exit Function

End Function
Private Function CreateSenioritiesBRV() As ADODB.Recordset
On Error GoTo CreateSenioritiesBRV_EH
Dim ladrs As New ADODB.Recordset

    With ladrs
        .Fields.Append "SNRTY_REC_TYPE_CD", adChar, 10, ADOFIELDATTR
        .Fields.Append "UNION_CD", adChar, 3, ADOFIELDATTR
        .Fields.Append "DEPT_NUM", adChar, 3, ADOFIELDATTR
        .Fields.Append "JOB_NUM", adChar, 6, ADOFIELDATTR

        .Fields.Append "REC_SOURCE", adInteger, , ADOFIELDATTR
        .Fields.Append "ERR_NUMBER", adInteger, , ADOFIELDATTR
        .Fields.Append "ERR_TYPE", adChar, 2, ADOFIELDATTR
        .Fields.Append "ERR_DESCRIPTION", adChar, 300, ADOFIELDATTR
        .Fields.Append "ERR_SOURCE", adChar, 1000, ADOFIELDATTR

        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open
    End With

    Set CreateSenioritiesBRV = ladrs
    Set ladrs = Nothing

    Exit Function
CreateSenioritiesBRV_EH:
    Set ladrs = Nothing

    TrsRaiseError MODULENAME _
        , "CreateSenioritiesBRV"

End Function
Private Function ValidateSeniorityEntries( _
                    ByRef radrsRSToValidate As ADODB.Recordset _
                    , ByRef radrsBRV As ADODB.Recordset _
                    , ByVal vbIsAGencoEmployee As Boolean _
                    ) As Boolean
On Error GoTo ValidateSeniorityEntries_EH
'Validate
Dim llRecordIndex As Long, lsFilterString As String
Dim llMonths As Long, llDays As Long

    ValidateSeniorityEntries = False

    If Not (radrsRSToValidate Is Nothing) Then
        If Not (radrsRSToValidate.EOF And _
            radrsRSToValidate.BOF) Then
            radrsRSToValidate.Filter = adFilterNone
            radrsRSToValidate.MoveFirst
        Else
            Exit Function
        End If
    Else
        Exit Function
    End If

    With radrsRSToValidate
        For llRecordIndex = 1 To .RecordCount
            Select Case (UCase(Trim(!RECORD_STATUS)))
                Case "I", "C"
                    If Not (CheckForValidString(!UNION_CD)) Then
                        AppendSeniorityBRV radrsBRV _
                            , trErrbusEmployeeSeniority _
                            , mstrCBUSINESSERR _
                            , "Invalid Union." _
                            , MODULENAME & ": ValidateSeniorityEntries" _
                            , _
                            , !DEPT_NUM _
                            , !JOB_NUM
                    Else
                        If Not (CheckForItemExistenceInDB(trUNION, !UNION_CD)) Then
                            AppendSeniorityBRV radrsBRV _
                                , trErrbusEmployeeSeniority _
                                , mstrCBUSINESSERR _
                                , "Invalid Union. Union " & Trim(!UNION_CD) & " does not exist." _
                                , MODULENAME & ": ValidateSeniorityEntries" _
                                , !UNION_CD _
                                , !DEPT_NUM _
                                , !JOB_NUM
                        End If
                    End If

                    If Not (CheckForValidString(!DEPT_NUM)) Then
                        AppendSeniorityBRV radrsBRV _
                            , trErrbusEmployeeSeniority _
                            , mstrCBUSINESSERR _
                            , "Invalid Department." _
                            , MODULENAME & ": ValidateSeniorityEntries" _
                            , !UNION_CD _
                            , _
                            , !JOB_NUM
                    Else
                        If Not (CheckForItemExistenceInDB(trDEPARTMENT, !DEPT_NUM)) Then
                            AppendSeniorityBRV radrsBRV _
                                , trErrbusEmployeeSeniority _
                                , mstrCBUSINESSERR _
                                , "Invalid Department. Department " & Trim(!DEPT_NUM) & " does not exist." _
                                , MODULENAME & ": ValidateSeniorityEntries" _
                                , !UNION_CD _
                                , !DEPT_NUM _
                                , !JOB_NUM
                        End If
                    End If

                    If Not (CheckForValidString(!JOB_NUM)) Then
                        AppendSeniorityBRV radrsBRV _
                            , trErrbusEmployeeSeniority _
                            , mstrCBUSINESSERR _
                            , "Invalid Job." _
                            , MODULENAME & ": ValidateSeniorityEntries" _
                            , !UNION_CD _
                            , !DEPT_NUM
                    Else
                        If Not (CheckForItemExistenceInDB(trJOB, !JOB_NUM)) Then
                            AppendSeniorityBRV radrsBRV _
                                , trErrbusEmployeeSeniority _
                                , mstrCBUSINESSERR _
                                , "Invalid Job. Job " & Trim(!JOB_NUM) & " does not exist." _
                                , MODULENAME & ": ValidateSeniorityEntries" _
                                , !UNION_CD _
                                , !DEPT_NUM _
                                , !JOB_NUM
                        End If
                    End If

                    If Trim(!SNRTY_PRTY_CD) = "" Then
                        !SNRTY_PRTY_CD = "00"
                        
                        AppendSeniorityBRV radrsBRV _
                            , trErrbusEmployeeSeniority _
                            , mstrCBUSINESSWAR _
                            , "Priority code has been changed from blanks to '00'." _
                            , MODULENAME & ": ValidateSeniorityEntries" _
                            , !UNION_CD _
                            , !DEPT_NUM _
                            , !JOB_NUM
                            
                            .Update
                    End If

                    'Check if it is a Genco Employee
                    If (vbIsAGencoEmployee) Then
                        'The reason the null and isdate are checked before
                        'calling the function in the else part is beacause
                        'VB sometimes substitutes a default date or throws an
                        'error if an invalid date is passed to a function with
                        'a date parameter
                        If Not (CheckForValidDate(!CORP_SNRTY_EFF_DT)) Then
                            AppendSeniorityBRV radrsBRV _
                                , trErrbusEmployeeSeniority _
                                , mstrCBUSINESSERR _
                                , "Invalid Corporate Effective Date." _
                                , MODULENAME & ": ValidateSeniorityEntries" _
                                , !UNION_CD _
                                , !DEPT_NUM _
                                , !JOB_NUM
                        Else
                            If Not (IsAValidSeniorityEffectiveDate(!CORP_SNRTY_EFF_DT)) Then
                                AppendSeniorityBRV radrsBRV _
                                    , trErrbusEmployeeSeniority _
                                    , mstrCBUSINESSERR _
                                    , "Invalid Corporate Effective Date. Effective date must be less than or equal to the current date." _
                                    , MODULENAME & ": ValidateSeniorityEntries" _
                                    , !UNION_CD _
                                    , !DEPT_NUM _
                                    , !JOB_NUM
                            Else
                                'Determine the number of months and days
                                llMonths = 0
                                llDays = 0
                                DetermineDaysAndMonths !CORP_SNRTY_EFF_DT _
                                    , llMonths _
                                    , llDays
                                !CORP_MONTHS = llMonths
                                !CORP_DAYS = llDays

                                .Update
                            End If
                        End If

                        If Not (CheckForValidDate(!CMPY_SNRTY_EFF_DT)) Then
                            AppendSeniorityBRV radrsBRV _
                                , trErrbusEmployeeSeniority _
                                , mstrCBUSINESSERR _
                                , "Invalid Company Effective Date." _
                                , MODULENAME & ": ValidateSeniorityEntries" _
                                , !UNION_CD _
                                , !DEPT_NUM _
                                , !JOB_NUM
                        Else
                            If Not (IsAValidSeniorityEffectiveDate(!CMPY_SNRTY_EFF_DT)) Then
                                AppendSeniorityBRV radrsBRV _
                                    , trErrbusEmployeeSeniority _
                                    , mstrCBUSINESSERR _
                                    , "Invalid Company Effective Date. Effective date must be less than or equal to the current date." _
                                    , MODULENAME & ": ValidateSeniorityEntries" _
                                    , !UNION_CD _
                                    , !DEPT_NUM _
                                    , !JOB_NUM
                            Else
                                'Determine the number of months and days
                                llMonths = 0
                                llDays = 0
                                DetermineDaysAndMonths !CMPY_SNRTY_EFF_DT _
                                    , llMonths _
                                    , llDays
                                !CMPY_MONTHS = llMonths
                                !CMPY_DAYS = llDays
                                .Update
                            End If
                        End If

                        If Not (CheckForValidDate(!PLANT_SNRTY_EFF_DT)) Then
                            AppendSeniorityBRV radrsBRV _
                                , trErrbusEmployeeSeniority _
                                , mstrCBUSINESSERR _
                                , "Invalid Plant Effective Date." _
                                , MODULENAME & ": ValidateSeniorityEntries" _
                                , !UNION_CD _
                                , !DEPT_NUM _
                                , !JOB_NUM
                        Else
                            If Not (IsAValidSeniorityEffectiveDate(!PLANT_SNRTY_EFF_DT)) Then
                                AppendSeniorityBRV radrsBRV _
                                    , trErrbusEmployeeSeniority _
                                    , mstrCBUSINESSERR _
                                    , "Invalid Plant Effective Date. Effective date must be less than or equal to the current date." _
                                    , MODULENAME & ": ValidateSeniorityEntries" _
                                    , !UNION_CD _
                                    , !DEPT_NUM _
                                    , !JOB_NUM
                            Else
                                'Determine the number of months and days
                                llMonths = 0
                                llDays = 0
                                DetermineDaysAndMonths !PLANT_SNRTY_EFF_DT _
                                    , llMonths _
                                    , llDays
                                !PLANT_MONTHS = llMonths
                                !PLANT_DAYS = llDays
                                .Update
                            End If
                        End If

                        If Not (CheckForValidDate(!DEPT_SNRTY_EFF_DT)) Then
                            AppendSeniorityBRV radrsBRV _
                                , trErrbusEmployeeSeniority _
                                , mstrCBUSINESSERR _
                                , "Invalid Department Effective Date." _
                                , MODULENAME & ": ValidateSeniorityEntries" _
                                , !UNION_CD _
                                , !DEPT_NUM _
                                , !JOB_NUM
                        Else
                            If Not (IsAValidSeniorityEffectiveDate(!DEPT_SNRTY_EFF_DT)) Then
                                AppendSeniorityBRV radrsBRV _
                                    , trErrbusEmployeeSeniority _
                                    , mstrCBUSINESSERR _
                                    , "Invalid Department Effective Date. Effective date must be less than or equal to the current date." _
                                    , MODULENAME & ": ValidateSeniorityEntries" _
                                    , !UNION_CD _
                                    , !DEPT_NUM _
                                    , !JOB_NUM
                            Else
                                'Determine the number of months and days
                                llMonths = 0
                                llDays = 0
                                DetermineDaysAndMonths !DEPT_SNRTY_EFF_DT _
                                    , llMonths _
                                    , llDays
                                !DEPT_MONTHS = llMonths
                                !DEPT_DAYS = llDays
                                .Update
                            End If
                        End If

                        If Not (CheckForValidDate(!JOB_SNRTY_EFF_DT)) Then
                            AppendSeniorityBRV radrsBRV _
                                , trErrbusEmployeeSeniority _
                                , mstrCBUSINESSERR _
                                , "Invalid Job Effective Date." _
                                , MODULENAME & ": ValidateSeniorityEntries" _
                                , !UNION_CD _
                                , !DEPT_NUM _
                                , !JOB_NUM
                        Else
                            If Not (IsAValidSeniorityEffectiveDate(!JOB_SNRTY_EFF_DT)) Then
                                AppendSeniorityBRV radrsBRV _
                                    , trErrbusEmployeeSeniority _
                                    , mstrCBUSINESSERR _
                                    , "Invalid Job Effective Date. Effective date must be less than or equal to the current date." _
                                    , MODULENAME & ": ValidateSeniorityEntries" _
                                    , !UNION_CD _
                                    , !DEPT_NUM _
                                    , !JOB_NUM
                            Else
                                'Determine the number of months and days
                                llMonths = 0
                                llDays = 0
                                DetermineDaysAndMonths !JOB_SNRTY_EFF_DT _
                                    , llMonths _
                                    , llDays
                                !JOB_MONTHS = llMonths
                                !JOB_DAYS = llDays
                                .Update
                            End If
                        End If
                    Else    'NON -GENCO
                        If (IsValidNumberOfDays(!CORP_DAYS) And _
                            IsValidNumberOfMonths(!CORP_MONTHS)) Then
                            !CORP_SNRTY_UP_DT = SetUpdateDateTo(!CORP_MONTHS _
                                                    , !CORP_DAYS)
                        Else
                            AppendSeniorityBRV radrsBRV _
                                , trErrbusEmployeeSeniority _
                                , mstrCBUSINESSERR _
                                , "Invalid Corporate Months and/or Days." _
                                , MODULENAME & ": ValidateSeniorityEntries" _
                                , !UNION_CD _
                                , !DEPT_NUM _
                                , !JOB_NUM
                        End If

                        If (IsValidNumberOfDays(!CMPY_DAYS) And _
                            IsValidNumberOfMonths(!CMPY_MONTHS)) Then
                            !CMPY_SNRTY_UP_DT = SetUpdateDateTo(!CMPY_MONTHS _
                                                    , !CMPY_DAYS)
                        Else
                            AppendSeniorityBRV radrsBRV _
                                , trErrbusEmployeeSeniority _
                                , mstrCBUSINESSERR _
                                , "Invalid Company Months and/or Days." _
                                , MODULENAME & ": ValidateSeniorityEntries" _
                                , !UNION_CD _
                                , !DEPT_NUM _
                                , !JOB_NUM
                        End If

                        If (IsValidNumberOfDays(!PLANT_DAYS) And _
                            IsValidNumberOfMonths(!PLANT_MONTHS)) Then
                            !PLANT_SNRTY_UP_DT = SetUpdateDateTo(!PLANT_MONTHS _
                                                    , !PLANT_DAYS)
                        Else
                            AppendSeniorityBRV radrsBRV _
                                , trErrbusEmployeeSeniority _
                                , mstrCBUSINESSERR _
                                , "Invalid Plant Months and/or Days." _
                                , MODULENAME & ": ValidateSeniorityEntries" _
                                , !UNION_CD _
                                , !DEPT_NUM _
                                , !JOB_NUM
                        End If

                        If (IsValidNumberOfDays(!UNION_DAYS) And _
                            IsValidNumberOfMonths(!UNION_MONTHS)) Then
                            !UNION_SNRTY_UP_DT = SetUpdateDateTo(!UNION_MONTHS _
                                                    , !UNION_DAYS)
                        Else
                            AppendSeniorityBRV radrsBRV _
                                , trErrbusEmployeeSeniority _
                                , mstrCBUSINESSERR _
                                , "Invalid Union Months and/or Days." _
                                , MODULENAME & ": ValidateSeniorityEntries" _
                                , !UNION_CD _
                                , !DEPT_NUM _
                                , !JOB_NUM
                        End If

                        If (IsValidNumberOfDays(!DEPT_DAYS) And _
                            IsValidNumberOfMonths(!DEPT_MONTHS)) Then
                            !DEPT_SNRTY_UP_DT = SetUpdateDateTo(!DEPT_MONTHS _
                                                    , !DEPT_DAYS)
                        Else
                            AppendSeniorityBRV radrsBRV _
                                , trErrbusEmployeeSeniority _
                                , mstrCBUSINESSERR _
                                , "Invalid Department Months and/or Days." _
                                , MODULENAME & ": ValidateSeniorityEntries" _
                                , !UNION_CD _
                                , !DEPT_NUM _
                                , !JOB_NUM
                        End If

                        If (IsValidNumberOfDays(!JOB_DAYS) And _
                            IsValidNumberOfMonths(!JOB_MONTHS)) Then
                            !JOB_SNRTY_UP_DT = SetUpdateDateTo(!JOB_MONTHS _
                                                    , !JOB_DAYS)
                        Else
                            AppendSeniorityBRV radrsBRV _
                                , trErrbusEmployeeSeniority _
                                , mstrCBUSINESSERR _
                                , "Invalid Job Months and/or Days." _
                                , MODULENAME & ": ValidateSeniorityEntries" _
                                , !UNION_CD _
                                , !DEPT_NUM _
                                , !JOB_NUM
                        End If

                        If (IsValidNumberOfDays(!LOP_DAYS) And _
                            IsValidNumberOfMonths(!LOP_MONTHS)) Then
                            !LOP_SNRTY_UP_DT = SetUpdateDateTo(!LOP_MONTHS _
                                                    , !LOP_DAYS)
                        Else
                            AppendSeniorityBRV radrsBRV _
                                , trErrbusEmployeeSeniority _
                                , mstrCBUSINESSERR _
                                , "Invalid Line of Promotion Months and/or Days." _
                                , MODULENAME & ": ValidateSeniorityEntries" _
                                , !UNION_CD _
                                , !DEPT_NUM _
                                , !JOB_NUM
                        End If
                        
                        .Update
                    End If
            End Select

            Select Case (UCase(Trim(!RECORD_STATUS)))
                Case "C", "D"
                    If Not (CheckForValidString(!KEY_UNION_CD)) Then
                        AppendSeniorityBRV radrsBRV _
                            , trErrbusEmployeeSeniority _
                            , mstrCBUSINESSERR _
                            , "Invalid Delete/Change request. Union Code is invalid." _
                            , MODULENAME & ": ValidateSeniorityEntries" _
                            , !UNION_CD _
                            , !DEPT_NUM _
                            , !JOB_NUM
                    End If

                    If Not (CheckForValidString(!KEY_DEPT_NUM)) Then
                        AppendSeniorityBRV radrsBRV _
                            , trErrbusEmployeeSeniority _
                            , mstrCBUSINESSERR _
                            , "Invalid Delete/Change request. Department Number is invalid." _
                            , MODULENAME & ": ValidateSeniorityEntries" _
                            , !UNION_CD _
                            , !DEPT_NUM _
                            , !JOB_NUM
                    End If

                    If Not (CheckForValidString(!KEY_JOB_NUM)) Then
                        AppendSeniorityBRV radrsBRV _
                            , trErrbusEmployeeSeniority _
                            , mstrCBUSINESSERR _
                            , "Invalid Delete/Change request. Job Number is invalid." _
                            , MODULENAME & ": ValidateSeniorityEntries" _
                            , !UNION_CD _
                            , !DEPT_NUM _
                            , !JOB_NUM
                    End If

                    If Not (CheckForValidString(!TIME_STAMP _
                        , , , True, 26)) Then
                        AppendSeniorityBRV radrsBRV _
                            , trErrbusEmployeeSeniority _
                            , mstrCBUSINESSERR _
                            , "Invalid Delete/Change request. Time Stamp is invalid." _
                            , MODULENAME & ": ValidateSeniorityEntries" _
                            , !UNION_CD _
                            , !DEPT_NUM _
                            , !JOB_NUM
                    End If
            End Select

            .MoveNext
        Next llRecordIndex
    End With

    radrsRSToValidate.Filter = adFilterNone
    radrsRSToValidate.MoveFirst

    ValidateSeniorityEntries = True

    If Not (radrsBRV Is Nothing) Then
        With radrsBRV
            If Not (.BOF And .EOF) Then
                .Filter = adFilterNone
                .MoveFirst
                lsFilterString = "ERR_TYPE = 'BE' OR ERR_TYPE = 'AE' OR ERR_TYPE = 'AW'"
                .Filter = stdmTRISCommon.MDAC25FilterFix(lsFilterString)
                lsFilterString = ""
                If Not (.EOF) Then
                    ValidateSeniorityEntries = False
                End If

                .Filter = adFilterNone
            End If
        End With
    End If

    Exit Function
ValidateSeniorityEntries_EH:
    ValidateSeniorityEntries = False

    TrsRaiseError MODULENAME _
            , "ValidateSeniorityEntries"

End Function
Private Function ValidateFrozenSeniorityEntries( _
                    ByRef radrsRSToValidate As ADODB.Recordset _
                    , ByRef radrsBRV As ADODB.Recordset _
                    ) As Boolean
On Error GoTo ValidateFrozenSeniorityEntries_EH
'Validate
Dim llRecordIndex As Long, lsFilterString As String
Dim llMonths As Long, llDays As Long

    ValidateFrozenSeniorityEntries = False

    If Not (radrsRSToValidate Is Nothing) Then
        If Not (radrsRSToValidate.EOF And _
            radrsRSToValidate.BOF) Then
            radrsRSToValidate.Filter = adFilterNone
            radrsRSToValidate.MoveFirst
        Else
            Exit Function
        End If
    Else
        Exit Function
    End If

    With radrsRSToValidate
        For llRecordIndex = 1 To .RecordCount
            Select Case (UCase(Trim(!RECORD_STATUS)))
                Case "I", "C"
                    If Not (CheckForValidString(!SNRTY_REC_TYPE_CD)) Then
                        AppendFrozenSeniorityBRV radrsBRV _
                            , trErrbusEmployeeFrozenSeniority _
                            , mstrCBUSINESSERR _
                            , "Invalid Union." _
                            , MODULENAME & ": ValidateFrozenSeniorityEntries"
                    Else
                        Select Case (UCase(Trim(!SNRTY_REC_TYPE_CD)))
                            Case "COMPANY"
                                If Not (CheckForValidString(!CMPY_CD)) Then
                                    AppendFrozenSeniorityBRV radrsBRV _
                                        , trErrbusEmployeeFrozenSeniority _
                                        , mstrCBUSINESSERR _
                                        , "Invalid Company." _
                                        , MODULENAME & ": ValidateFrozenSeniorityEntries" _
                                        , !SNRTY_REC_TYPE_CD
                                Else
                                    If Not (CheckForItemExistenceInDB(trCOMPANY, !CMPY_CD)) Then
                                        AppendFrozenSeniorityBRV radrsBRV _
                                            , trErrbusEmployeeFrozenSeniority _
                                            , mstrCBUSINESSERR _
                                            , "Invalid Company. Company " & Trim(!CMPY_CD) & " does not exist." _
                                            , MODULENAME & ": ValidateFrozenSeniorityEntries" _
                                            , !SNRTY_REC_TYPE_CD
                                    End If
                                End If

                            Case "UNION"
                                If Not (CheckForValidString(!UNION_CD)) Then
                                    AppendFrozenSeniorityBRV radrsBRV _
                                        , trErrbusEmployeeFrozenSeniority _
                                        , mstrCBUSINESSERR _
                                        , "Invalid Union." _
                                        , MODULENAME & ": ValidateFrozenSeniorityEntries" _
                                        , !SNRTY_REC_TYPE_CD
                                Else
                                    If Not (CheckForItemExistenceInDB(trUNION, !UNION_CD)) Then
                                        AppendFrozenSeniorityBRV radrsBRV _
                                            , trErrbusEmployeeFrozenSeniority _
                                            , mstrCBUSINESSERR _
                                            , "Invalid Union. Union " & Trim(!UNION_CD) & " does not exist." _
                                            , MODULENAME & ": ValidateFrozenSeniorityEntries" _
                                            , !SNRTY_REC_TYPE_CD
                                    End If
                                End If

                            Case "JOB"
                                If Not (CheckForValidString(!JOB_NUM)) Then
                                    AppendFrozenSeniorityBRV radrsBRV _
                                        , trErrbusEmployeeFrozenSeniority _
                                        , mstrCBUSINESSERR _
                                        , "Invalid Job." _
                                        , MODULENAME & ": ValidateFrozenSeniorityEntries" _
                                        , !SNRTY_REC_TYPE_CD
                                Else
                                    If Not (CheckForItemExistenceInDB(trJOB, !JOB_NUM)) Then
                                        AppendFrozenSeniorityBRV radrsBRV _
                                            , trErrbusEmployeeFrozenSeniority _
                                            , mstrCBUSINESSERR _
                                            , "Invalid Job. Job " & Trim(!JOB_NUM) & " does not exist." _
                                            , MODULENAME & ": ValidateFrozenSeniorityEntries" _
                                            , !SNRTY_REC_TYPE_CD
                                    End If
                                End If

                            Case "PLANT", "LOP", "DEPARTMENT"
                                If Not (CheckForValidString(!UNION_CD)) Then
                                    AppendFrozenSeniorityBRV radrsBRV _
                                        , trErrbusEmployeeFrozenSeniority _
                                        , mstrCBUSINESSERR _
                                        , "Invalid Union." _
                                        , MODULENAME & ": ValidateFrozenSeniorityEntries" _
                                        , !SNRTY_REC_TYPE_CD
                                Else
                                    If Not (CheckForItemExistenceInDB(trUNION, !UNION_CD)) Then
                                        AppendFrozenSeniorityBRV radrsBRV _
                                            , trErrbusEmployeeFrozenSeniority _
                                            , mstrCBUSINESSERR _
                                            , "Invalid Union. Union " & Trim(!UNION_CD) & " does not exist." _
                                            , MODULENAME & ": ValidateFrozenSeniorityEntries" _
                                            , !SNRTY_REC_TYPE_CD
                                    End If
                                End If

                                If Not (CheckForValidString(!IR_AREA_CD)) Then
                                    AppendFrozenSeniorityBRV radrsBRV _
                                        , trErrbusEmployeeFrozenSeniority _
                                        , mstrCBUSINESSERR _
                                        , "Invalid Industrial Relations Area." _
                                        , MODULENAME & ": ValidateFrozenSeniorityEntries" _
                                        , !SNRTY_REC_TYPE_CD
                                Else
                                    If Not (CheckForItemExistenceInDB(trIRAREA, !IR_AREA_CD)) Then
                                        AppendFrozenSeniorityBRV radrsBRV _
                                            , trErrbusEmployeeFrozenSeniority _
                                            , mstrCBUSINESSERR _
                                            , "Invalid Industrial Relations Area. Industrial Relations Area " & Trim(!IR_AREA_CD) & " does not exist." _
                                            , MODULENAME & ": ValidateFrozenSeniorityEntries" _
                                            , !SNRTY_REC_TYPE_CD
                                    End If
                                End If

                                If Not (CheckForValidString(!IR_DIV_CD)) Then
                                    AppendFrozenSeniorityBRV radrsBRV _
                                        , trErrbusEmployeeFrozenSeniority _
                                        , mstrCBUSINESSERR _
                                        , "Invalid Industrial Relations Division." _
                                        , MODULENAME & ": ValidateFrozenSeniorityEntries" _
                                        , !SNRTY_REC_TYPE_CD
                                Else
                                    If Not (CheckForItemExistenceInDB(trIRDIVISION, !IR_DIV_CD)) Then
                                        AppendFrozenSeniorityBRV radrsBRV _
                                            , trErrbusEmployeeFrozenSeniority _
                                            , mstrCBUSINESSERR _
                                            , "Invalid Industrial Relations Division. Industrial Relations Division " & Trim(!IR_DIV_CD) & " does not exist." _
                                            , MODULENAME & ": ValidateFrozenSeniorityEntries" _
                                            , !SNRTY_REC_TYPE_CD
                                    End If
                                End If
                        End Select

                        If Not (CheckForValidDate(!SNRTY_FRZN_DT)) Then
                            !SNRTY_FRZN_DT = Date
                            .Update
                        Else
                            If Not (IsAValidSeniorityEffectiveDate(!SNRTY_FRZN_DT)) Then
                                AppendFrozenSeniorityBRV radrsBRV _
                                    , trErrbusEmployeeFrozenSeniority _
                                    , mstrCBUSINESSERR _
                                    , "Invalid Freeze Date. Freeze date must be less than or equal to the current date." _
                                    , MODULENAME & ": ValidateFrozenSeniorityEntries" _
                                    , !SNRTY_REC_TYPE_CD
                            End If
                        End If

                        If Not (IsValidNumberOfDays(!FRZN_SNRTY_DAYS) Or _
                            IsValidNumberOfMonths(!FRZN_SNRTY_MONTHS)) Then
                            AppendFrozenSeniorityBRV radrsBRV _
                                , trErrbusEmployeeFrozenSeniority _
                                , mstrCBUSINESSERR _
                                , "Invalid Frozen Seniority Months and/or Days." _
                                , MODULENAME & ": ValidateFrozenSeniorityEntries" _
                                , !SNRTY_REC_TYPE_CD
                        End If
                    End If
            End Select

            Select Case (UCase(Trim(!RECORD_STATUS)))
                Case "C", "D"
                    If Not (CheckForValidString(!KEY_SNRTY_REC_TYPE_CD)) Then
                        AppendFrozenSeniorityBRV radrsBRV _
                            , trErrbusEmployeeFrozenSeniority _
                            , mstrCBUSINESSERR _
                            , "Invalid Delete/Change request. Frozen Seniority Type is invalid." _
                            , MODULENAME & ": ValidateFrozenSeniorityEntries" _
                            , !SNRTY_REC_TYPE_CD
                    End If

                    If Not (CheckForValidString(!TIME_STAMP _
                        , , , True, 26)) Then
                        AppendFrozenSeniorityBRV radrsBRV _
                            , trErrbusEmployeeFrozenSeniority _
                            , mstrCBUSINESSERR _
                            , "Invalid Delete/Change request. Time Stamp is invalid." _
                            , MODULENAME & ": ValidateFrozenSeniorityEntries" _
                            , !SNRTY_REC_TYPE_CD
                    End If
            End Select

            Select Case (UCase(Trim(!SNRTY_REC_TYPE_CD)))
                Case "COMPANY"
                    !SNRTY_REC_TYPE_CD = "CMP"
                    .Update
                Case "PLANT"
                    !SNRTY_REC_TYPE_CD = "PLT"
                    .Update
                Case "UNION"
                    !SNRTY_REC_TYPE_CD = "UNI"
                    .Update
                Case "DEPARTMENT"
                    !SNRTY_REC_TYPE_CD = "DPT"
                    .Update
                Case "LOP"
                    !SNRTY_REC_TYPE_CD = "LOP"
                    .Update
                Case "JOB"
                    !SNRTY_REC_TYPE_CD = "JOB"
                    .Update
            End Select

            Select Case (UCase(Trim(!KEY_SNRTY_REC_TYPE_CD)))
                Case "COMPANY"
                    !KEY_SNRTY_REC_TYPE_CD = "CMP"
                    .Update
                Case "PLANT"
                    !KEY_SNRTY_REC_TYPE_CD = "PLT"
                    .Update
                Case "UNION"
                    !KEY_SNRTY_REC_TYPE_CD = "UNI"
                    .Update
                Case "DEPARTMENT"
                    !KEY_SNRTY_REC_TYPE_CD = "DPT"
                    .Update
                Case "LOP"
                    !KEY_SNRTY_REC_TYPE_CD = "LOP"
                    .Update
                Case "JOB"
                    !KEY_SNRTY_REC_TYPE_CD = "JOB"
                    .Update
            End Select

            .MoveNext
        Next llRecordIndex
    End With

    radrsRSToValidate.Filter = adFilterNone
    radrsRSToValidate.MoveFirst

    ValidateFrozenSeniorityEntries = True

    If Not (radrsBRV Is Nothing) Then
        With radrsBRV
            If Not (.BOF And .EOF) Then
                .Filter = adFilterNone
                .MoveFirst
                lsFilterString = "ERR_TYPE = 'BE' OR ERR_TYPE = 'AE' OR ERR_TYPE = 'AW'"
                .Filter = stdmTRISCommon.MDAC25FilterFix(lsFilterString)
                lsFilterString = ""
                If Not (.EOF) Then
                    ValidateFrozenSeniorityEntries = False
                End If

                .Filter = adFilterNone
            End If
        End With
    End If

    Exit Function
ValidateFrozenSeniorityEntries_EH:
    ValidateFrozenSeniorityEntries = False

    TrsRaiseError MODULENAME _
            , "ValidateFrozenSeniorityEntries"

End Function
Private Sub DetermineDaysAndMonths(ByVal vdtEffectiveDate As Date _
                , ByRef vlMonths As Long _
                , ByRef vlDays As Long)
On Error GoTo DetermineDaysAndMonths_EH
Dim llMonthOfEffectiveDate As Long, llDayOfEffectiveDate
Dim llYearOfEffectiveDate As Long
Dim ldtToday As Date, llMonthOfToday As Long
Dim llDayOfToday As Long, llYearOfToday As Long

    'Hopefully the date does not change
    '(no midnight calls please!!!)
    'Hence the variable
    ldtToday = Date

    If (ldtToday = vdtEffectiveDate) Then
        vlMonths = 0
        vlDays = 0
    Else
        llMonthOfToday = Month(ldtToday)
        llDayOfToday = Day(ldtToday)
        llYearOfToday = Year(ldtToday)

        llYearOfEffectiveDate = Year(vdtEffectiveDate)
        llMonthOfEffectiveDate = Month(vdtEffectiveDate)
        llDayOfEffectiveDate = Day(vdtEffectiveDate)

        If (llDayOfEffectiveDate < llDayOfToday) Or _
            (llDayOfEffectiveDate = llDayOfToday) Then
            vlDays = (llDayOfToday - llDayOfEffectiveDate)
        Else
            If (llDayOfEffectiveDate > llDayOfToday) Then
                vlDays = ((llDayOfToday + 30) - llDayOfEffectiveDate)
                llMonthOfToday = llMonthOfToday - 1
            End If
        End If

        If ((llYearOfToday - llYearOfEffectiveDate) > 0) Then
            vlMonths = ((llYearOfToday - llYearOfEffectiveDate) * 12 _
                            + (llMonthOfToday - llMonthOfEffectiveDate))
        Else
            vlMonths = (llMonthOfToday - llMonthOfEffectiveDate)
        End If
    End If

    Exit Sub
DetermineDaysAndMonths_EH:
    TrsRaiseError MODULENAME _
        , "DetermineDaysAndMonths"

End Sub
Private Function SetUpdateDateTo(ByVal vlMonths As Long _
                    , ByVal vlDays As Long _
                    ) As Date
On Error GoTo SetUpdateDateTo_EH

    If ((vlDays > 0 And vlDays <= 29) Or _
        vlMonths > 0) Then
        SetUpdateDateTo = Date
    Else
        SetUpdateDateTo = #1/1/100#
    End If

    Exit Function
SetUpdateDateTo_EH:
    TrsRaiseError MODULENAME _
        , "SetUpdateDateTo"

End Function
Private Function IsValidNumberOfDays(ByVal vlDays As Long _
                    ) As Boolean
On Error GoTo IsValidNumberOfDays_EH

    IsValidNumberOfDays = False
    'The domain value for Days includes any number between 0 and 29
    If (vlDays >= 0 And vlDays <= 29) Then
        IsValidNumberOfDays = True
    End If

    Exit Function
IsValidNumberOfDays_EH:
    TrsRaiseError MODULENAME _
        , "IsValidNumberOfDays"

End Function
Private Function IsValidNumberOfMonths(ByVal vlMonths As Long _
                    ) As Boolean
On Error GoTo IsValidNumberOfMonths_EH

    IsValidNumberOfMonths = False
    'The domain value for Days includes any number between 0 and 29
    If (vlMonths >= 0 And vlMonths <= 999) Then
        IsValidNumberOfMonths = True
    End If

    Exit Function
IsValidNumberOfMonths_EH:
    TrsRaiseError MODULENAME _
        , "IsValidNumberOfMonths"

End Function
Private Function IsAValidSeniorityEffectiveDate(ByVal vdtEffectiveDate As Date _
                    ) As Boolean
On Error GoTo IsAValidSeniorityEffectiveDate_EH

    IsAValidSeniorityEffectiveDate = True

    If (vdtEffectiveDate > Date) Then
        IsAValidSeniorityEffectiveDate = False
    End If

    Exit Function
IsAValidSeniorityEffectiveDate_EH:
    TrsRaiseError MODULENAME _
        , "IsAValidSeniorityEffectiveDate"

End Function
Private Sub AppendSeniorityBRV(ByRef radrsBRV As ADODB.Recordset _
                , ByVal vlErrorNumber As Long _
                , ByVal vsErrorType As String _
                , ByVal vsErrorDescription As String _
                , ByVal vsErrorSource As String _
                , Optional ByVal vsUnionCode As String = "" _
                , Optional ByVal vsDepartmentNumber As String = "" _
                , Optional ByVal vsJobNumber As String = "" _
                , Optional ByVal viRecordSource As Integer = 0 _
                )
On Error GoTo AppendSeniorityBRV_EH

    With radrsBRV
        If (.State <> adStateOpen) Then
            Exit Sub
        End If

        .AddNew

        !UNION_CD = vsUnionCode
        !DEPT_NUM = vsDepartmentNumber
        !JOB_NUM = vsJobNumber

        !ERR_NUMBER = vlErrorNumber - vbObjectError
        !ERR_TYPE = Trim(vsErrorType)
        !REC_SOURCE = viRecordSource
        !ERR_DESCRIPTION = Trim(vsErrorDescription)
        !ERR_SOURCE = Trim(vsErrorSource)

        .Update
    End With

    Exit Sub
AppendSeniorityBRV_EH:
    TrsRaiseError MODULENAME _
        , "AppendSeniorityBRV"

End Sub
Private Sub AppendFrozenSeniorityBRV(ByRef radrsBRV As ADODB.Recordset _
                , ByVal vlErrorNumber As Long _
                , ByVal vsErrorType As String _
                , ByVal vsErrorDescription As String _
                , ByVal vsErrorSource As String _
                , Optional ByVal vsFrozenSeniorityType As String = "" _
                , Optional ByVal viRecordSource As Integer = 0 _
                )
On Error GoTo AppendFrozenSeniorityBRV_EH

    With radrsBRV
        If (.State <> adStateOpen) Then
            Exit Sub
        End If

        .AddNew

        !SNRTY_REC_TYPE_CD = vsFrozenSeniorityType

        !ERR_NUMBER = vlErrorNumber - vbObjectError
        !ERR_TYPE = Trim(vsErrorType)
        !REC_SOURCE = viRecordSource
        !ERR_DESCRIPTION = Trim(vsErrorDescription)
        !ERR_SOURCE = Trim(vsErrorSource)

        .Update
    End With

    Exit Sub
AppendFrozenSeniorityBRV_EH:
    TrsRaiseError MODULENAME _
        , "AppendFrozenSeniorityBRV"

End Sub
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'The following code has been REPLICATED
'It was borrowed from cbsEmployee
Public Function EmployeeCompany( _
                    ByVal v_datDeptEffDate As Date _
                    , ByVal v_strDeptNum As String _
                    , ByVal v_strPriorDeptNum As String _
                    , Optional ByVal v_strCompanyCode As String = "NA" _
                    , Optional ByVal v_strPriorCompanyCode As String = "NA" _
                    , Optional ByVal v_strChgDate As Date = #1/1/100# _
                    ) As String
On Error GoTo eh
Dim l_datChgDate As Date

    l_datChgDate = EffectiveDate(v_strChgDate)
    EmployeeCompany = CompanyCode( _
                        l_datChgDate, v_strDeptNum _
                        , v_strPriorDeptNum, v_datDeptEffDate _
                        , v_strCompanyCode, v_strPriorCompanyCode)
    Exit Function
eh:
    TrsRaiseError MODULENAME _
        , "EmployeeCompany" _
        , , , , _
        , True

End Function
Private Function EffectiveDate(ByVal v_datChgDate As Date _
                    ) As Date
'    If v_datChgDate is the default date, answer current date;
'    otherwise answer v_datChgDate
On Error GoTo eh

    If (v_datChgDate = #1/1/100#) Then
        EffectiveDate = Date
    Else
        EffectiveDate = v_datChgDate
    End If

    Exit Function
eh:
    TrsRaiseError MODULENAME _
        , "EffectiveDate"

End Function
Private Function CompanyCode( _
            ByVal v_datChgDate As Date _
            , ByVal v_strDept As String _
            , ByVal v_strPriorDept As String _
            , ByVal v_datDeptEffDate As Date _
            , ByVal v_strCompanyCode As String _
            , ByVal v_strPriorCompanyCode As String _
            ) As String
On Error GoTo eh

Dim l_blnUseCurrent As Boolean

    If (EffectiveDate(v_datChgDate) >= v_datDeptEffDate) Then
        l_blnUseCurrent = True
    Else
        If (Trim(v_strPriorDept) = "") Then
            l_blnUseCurrent = True
        Else
            l_blnUseCurrent = False
        End If
    End If

    If (l_blnUseCurrent) Then
        CompanyCode = GetCompanyCode( _
                        v_strDept, v_strCompanyCode)
    Else
        CompanyCode = GetCompanyCode( _
                        v_strPriorDept, v_strPriorCompanyCode)
    End If
    
    Exit Function
eh:
    TrsRaiseError MODULENAME _
        , "CompanyCode"

End Function
Private Function GetCompanyCode( _
            ByVal v_strDept As String _
            , ByVal v_strCompanyCode As String _
            ) As String
On Error GoTo eh
Dim l_obj As dbDepartment.cdDepartment

    If (v_strCompanyCode = "NA") Then
        Set l_obj = CreateObject( _
                        "dbDepartment.cdDepartment")
        GetCompanyCode = l_obj.GetCompanyCode( _
                            v_strDept)
        Set l_obj = Nothing
    Else
        GetCompanyCode = v_strCompanyCode
    End If

    Exit Function
eh:
    Set l_obj = Nothing
    TrsRaiseError MODULENAME _
        , "GetCompanyCode"

End Function
'The above code has been REPLICATED
'It was borrowed from cbsEmployee
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'This method will save employee canvass job entries
Public Function ProcessEmpCanvassJobEntries( _
                    ByVal vadrsCanvassJobInformation As ADODB.Recordset _
                    , ByRef radrsErrors As ADODB.Recordset _
                    , ByVal vlUserEmployeeNumber As Long _
                    ) As Boolean
On Error GoTo ProcessEmpCanvassJobEntries_EH
Dim lbEmptyRequest As Boolean, lsFilterString As String
Dim lobjdbEmployee As dbEmployee.cdEmployeeTrans
Dim llRecordIndex As Long

    'Check for request existence
    lbEmptyRequest = True
    If Not (vadrsCanvassJobInformation Is Nothing) Then
        If Not (vadrsCanvassJobInformation.BOF And _
            vadrsCanvassJobInformation.EOF) Then
            lbEmptyRequest = False
        End If
    End If

    'Prepare BRV Recordset
    Set radrsErrors = CreateCanvassJobBRV

    If (lbEmptyRequest) Then
        AppendCanvassJobBRV radrsErrors _
            , trErrbusEmployeeCanvassJob _
            , mstrCBUSINESSERR _
            , "There are no Employee Canvass Job entries to process." _
            , MODULENAME & ": ProcessEmpCanvassJobEntries"
        GoTo REPORTBRV
    Else
        With vadrsCanvassJobInformation
            .Filter = adFilterNone
            If Not (.EOF) Then
                .MoveFirst
                If Not (ValidateCanvassJobEntries( _
                    vadrsCanvassJobInformation _
                    , radrsErrors _
                    , vlUserEmployeeNumber)) Then
                    GoTo REPORTBRV
                End If

                .Filter = adFilterNone
                .MoveFirst
            End If
        End With
    End If

    Set lobjdbEmployee = CreateObject( _
                            "dbEmployee.cdEmployeeTrans")
    'Set lobjdbEmployee = New dbEmployee.cdEmployee
    If Not (lobjdbEmployee.dbProcessCanvassJobEntries( _
            vadrsCanvassJobInformation _
            , vlUserEmployeeNumber)) Then
        Set lobjdbEmployee = Nothing
        Err.Raise trErrbusEmployeeCanvassJob
    End If
    Set lobjdbEmployee = Nothing

    Set vadrsCanvassJobInformation = Nothing
    Set radrsErrors = Nothing

    GetObjectContext().SetComplete

    ProcessEmpCanvassJobEntries = True

    Set lobjdbEmployee = Nothing

    Exit Function
ProcessEmpCanvassJobEntries_EH:
    'Abort Transaction
    GetObjectContext().SetAbort
    Set lobjdbEmployee = Nothing

    Set vadrsCanvassJobInformation = Nothing
    Set radrsErrors = Nothing

    ProcessEmpCanvassJobEntries = False

    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "ProcessEmpCanvassJobEntries"
            Case trErrdbInsIntoEmpCanvJob, trErrdbUpdEmpCanvJob _
                , trErrdbDelFromEmpCanvJob, trErrdbSaveEmpCanvJob _
                , trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "ProcessEmpCanvassJobEntries" _
                    , Err.Number _
                    , , , , _
                    True
            Case Else
                TrsRaiseError MODULENAME _
                    , "ProcessEmpCanvassJobEntries" _
                    , trErrbusEmployeeCanvassJob _
                    , , , , _
                    True
    End Select

    Exit Function

REPORTBRV:
    Set lobjdbEmployee = Nothing
    Set vadrsCanvassJobInformation = Nothing

    If Not (radrsErrors Is Nothing) Then
        With radrsErrors
            If Not (.EOF And .BOF) Then
                .MoveFirst
                .Filter = adFilterNone

                'Check for errors
'MDAC25 FILTER FIX
                lsFilterString = "ERR_TYPE = 'BE'"
                .Filter = stdmTRISCommon.MDAC25FilterFix(lsFilterString)
                lsFilterString = ""
                'If found abort else complete
                If Not (.EOF) Then
                    'Abort Transaction
                    GetObjectContext().SetAbort
                    ProcessEmpCanvassJobEntries = False
                Else
                    'Complete Transaction
                    GetObjectContext().SetComplete
                    ProcessEmpCanvassJobEntries = True
                End If

                .Filter = adFilterNone
            Else
                'Complete Transaction
                GetObjectContext().SetComplete
                ProcessEmpCanvassJobEntries = True
            End If
        End With
    Else
        'Complete Transaction
        GetObjectContext().SetComplete
        ProcessEmpCanvassJobEntries = True
    End If

    Exit Function

End Function
Private Function CreateCanvassJobBRV() As ADODB.Recordset
On Error GoTo CreateCanvassJobBRV_EH
Dim ladrs As New ADODB.Recordset

    With ladrs
        .Fields.Append "EMP_NUM", adInteger, , ADOFIELDATTR
        .Fields.Append "JOB_NUM", adChar, 6, ADOFIELDATTR
        .Fields.Append "CRAFT_NUM", adChar, 3, ADOFIELDATTR
        .Fields.Append "EFF_START_DATE", adDate, , ADOFIELDATTR
        .Fields.Append "EFF_END_DATE", adDate, , ADOFIELDATTR

        .Fields.Append "REC_SOURCE", adInteger, , ADOFIELDATTR
        .Fields.Append "ERR_NUMBER", adInteger, , ADOFIELDATTR
        .Fields.Append "ERR_TYPE", adChar, 2, ADOFIELDATTR
        .Fields.Append "ERR_DESCRIPTION", adChar, 300, ADOFIELDATTR
        .Fields.Append "ERR_SOURCE", adChar, 1000, ADOFIELDATTR

        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open
    End With

    Set CreateCanvassJobBRV = ladrs
    Set ladrs = Nothing

    Exit Function
CreateCanvassJobBRV_EH:
    Set ladrs = Nothing

    TrsRaiseError MODULENAME _
        , "CreateCanvassJobBRV"

End Function
Private Function ValidateCanvassJobEntries( _
                    ByRef radrsRSToValidate As ADODB.Recordset _
                    , ByRef radrsBRV As ADODB.Recordset _
                    , ByVal vlUserEmployeeNumber As Long _
                    ) As Boolean
On Error GoTo ValidateCanvassJobEntries_EH
Dim llRecordIndex As Long, lsFilterString As String

    ValidateCanvassJobEntries = False

    If Not (radrsRSToValidate Is Nothing) Then
        If Not (radrsRSToValidate.EOF And _
            radrsRSToValidate.BOF) Then
            radrsRSToValidate.Filter = adFilterNone
            radrsRSToValidate.MoveFirst
        Else
            Exit Function
        End If
    Else
        Exit Function
    End If

    If Not (CheckForValidLong(vlUserEmployeeNumber, True)) Then
        AppendCanvassJobBRV radrsBRV _
            , trErrbusEmployeeCanvassJob _
            , mstrCBUSINESSERR _
            , "Invalid TRIS User." _
            , MODULENAME & ": ValidateCanvassJobEntries"
    End If

    With radrsRSToValidate
        For llRecordIndex = 1 To .RecordCount
            Select Case (UCase(Trim(!RECORD_STATUS)))
                Case "I", "C"
                    If Not (CheckForValidLong(!EMP_NUM, True)) Then
                        AppendCanvassJobBRV radrsBRV _
                            , trErrbusEmployeeCanvassJob _
                            , mstrCBUSINESSERR _
                            , "Invalid Employee." _
                            , MODULENAME & ": ValidateCanvassJobEntries" _
                            , _
                            , !JOB_NUM _
                            , !CRAFT_NUM _
                            , !EFF_START_DATE _
                            , !EFF_END_DATE
                    Else
                        If Not (CheckForItemExistenceInDB(trEMPLOYEE _
                            , , !EMP_NUM)) Then
                            AppendCanvassJobBRV radrsBRV _
                                , trErrbusEmployeeCanvassJob _
                                , mstrCBUSINESSERR _
                                , "Invalid Employee. Employee " & Trim(!EMP_NUM) & " does not exist." _
                                , MODULENAME & ": ValidateCanvassJobEntries" _
                                , !EMP_NUM _
                                , !JOB_NUM _
                                , !CRAFT_NUM _
                                , !EFF_START_DATE _
                                , !EFF_END_DATE
                        End If
                    End If

                    If Not (CheckForValidString(!JOB_NUM)) Then
                        AppendCanvassJobBRV radrsBRV _
                            , trErrbusEmployeeCanvassJob _
                            , mstrCBUSINESSERR _
                            , "Invalid Job." _
                            , MODULENAME & ": ValidateCanvassJobEntries" _
                            , !EMP_NUM _
                            , _
                            , !CRAFT_NUM _
                            , !EFF_START_DATE _
                            , !EFF_END_DATE
                    Else
                        If Not (CheckForItemExistenceInDB(trJOB _
                            , !JOB_NUM)) Then
                            AppendCanvassJobBRV radrsBRV _
                                , trErrbusEmployeeCanvassJob _
                                , mstrCBUSINESSERR _
                                , "Invalid Job. Job number " & Trim(!JOB_NUM) & " does not exist." _
                                , MODULENAME & ": ValidateCanvassJobEntries" _
                                , !EMP_NUM _
                                , !JOB_NUM _
                                , !CRAFT_NUM _
                                , !EFF_START_DATE _
                                , !EFF_END_DATE
                        End If
                    End If

                    If (CheckForValidString(!CRAFT_NUM)) Then
                        If Not (CheckForItemExistenceInDB(trCRAFT, !CRAFT_NUM)) Then
                            AppendCanvassJobBRV radrsBRV _
                                , trErrbusEmployeeCanvassJob _
                                , mstrCBUSINESSERR _
                                , "Invalid Craft. Craft number " & Trim(!CRAFT_NUM) & " does not exist." _
                                , MODULENAME & ": ValidateCanvassJobEntries" _
                                , !EMP_NUM _
                                , !JOB_NUM _
                                , _
                                , !EFF_START_DATE _
                                , !EFF_END_DATE
                        End If
                    Else
                        !CRAFT_NUM = ""
                        .Update
                    End If

                    If Not (CheckForValidDate(!EFF_START_DATE)) Then
                        AppendCanvassJobBRV radrsBRV _
                            , trErrbusEmployeeCanvassJob _
                            , mstrCBUSINESSERR _
                            , "Invalid Start Date." _
                            , MODULENAME & ": ValidateCanvassJobEntries" _
                            , !EMP_NUM _
                            , !JOB_NUM _
                            , !CRAFT_NUM _
                            , _
                            , !EFF_END_DATE
                    End If

                    If (IsNull(!EFF_END_DATE)) Then
                        !EFF_END_DATE = #12/31/9999#
                        .Update
                    Else
                        If Not (IsDate(!EFF_END_DATE)) Then
                            AppendCanvassJobBRV radrsBRV _
                                , trErrbusEmployeeCanvassJob _
                                , mstrCBUSINESSERR _
                                , "Invalid End Date." _
                                , MODULENAME & ": ValidateCanvassJobEntries" _
                                , !EMP_NUM _
                                , !JOB_NUM _
                                , !CRAFT_NUM _
                                , !EFF_START_DATE _
                                , !EFF_END_DATE
                        End If
                    End If
            End Select

            Select Case (UCase(Trim(!RECORD_STATUS)))
                Case "C", "D"
                    If Not (CheckForValidLong(!KEY_EMP_NUM, True)) Then
                        AppendCanvassJobBRV radrsBRV _
                            , trErrbusEmployeeCanvassJob _
                            , mstrCBUSINESSERR _
                            , "Invalid Delete/Change request. Employee Number is invalid." _
                            , MODULENAME & ": ValidateCanvassJobEntries" _
                            , !EMP_NUM _
                            , !JOB_NUM _
                            , !CRAFT_NUM _
                            , !EFF_START_DATE _
                            , !EFF_END_DATE
                    End If

                    If Not (CheckForValidString(!KEY_JOB_NUM)) Then
                        AppendCanvassJobBRV radrsBRV _
                            , trErrbusEmployeeCanvassJob _
                            , mstrCBUSINESSERR _
                            , "Invalid Delete/Change request. Job number is invalid." _
                            , MODULENAME & ": ValidateCanvassJobEntries" _
                            , !EMP_NUM _
                            , !JOB_NUM _
                            , !CRAFT_NUM _
                            , !EFF_START_DATE _
                            , !EFF_END_DATE
                    End If

                    If Not (CheckForValidDate(!KEY_EFF_START_DATE)) Then
                        AppendCanvassJobBRV radrsBRV _
                            , trErrbusEmployeeCanvassJob _
                            , mstrCBUSINESSERR _
                            , "Invalid Delete/Change request. Effective Start Date is invalid." _
                            , MODULENAME & ": ValidateCanvassJobEntries" _
                            , !EMP_NUM _
                            , !JOB_NUM _
                            , !CRAFT_NUM _
                            , !EFF_START_DATE _
                            , !EFF_END_DATE
                    End If

                    If Not (CheckForValidString(!CREATED_TIMESTAMP, , , True, 26)) Then
                        AppendCanvassJobBRV radrsBRV _
                            , trErrbusEmployeeCanvassJob _
                            , mstrCBUSINESSERR _
                            , "Invalid Delete/Change request. Time Stamp is invalid." _
                            , MODULENAME & ": ValidateCanvassJobEntries" _
                            , !EMP_NUM _
                            , !JOB_NUM _
                            , !CRAFT_NUM _
                            , !EFF_START_DATE _
                            , !EFF_END_DATE
                    End If
            End Select

            .MoveNext
        Next llRecordIndex
    End With

    radrsRSToValidate.Filter = adFilterNone
    radrsRSToValidate.MoveFirst

    ValidateCanvassJobEntries = True

    If Not (radrsBRV Is Nothing) Then
        With radrsBRV
            If Not (.BOF And .EOF) Then
                .Filter = adFilterNone
                .MoveFirst
                lsFilterString = "ERR_TYPE = 'BE' OR ERR_TYPE = 'AE' OR ERR_TYPE = 'AW'"
                .Filter = stdmTRISCommon.MDAC25FilterFix(lsFilterString)
                lsFilterString = ""
                If Not (.EOF) Then
                    ValidateCanvassJobEntries = False
                End If

                .Filter = adFilterNone
            End If
        End With
    End If

    Exit Function
ValidateCanvassJobEntries_EH:
    ValidateCanvassJobEntries = False

    TrsRaiseError MODULENAME _
            , "ValidateCanvassJobEntries"

End Function
Private Sub AppendCanvassJobBRV(ByRef radrsBRV As ADODB.Recordset _
                , ByVal vlErrorNumber As Long _
                , ByVal vsErrorType As String _
                , ByVal vsErrorDescription As String _
                , ByVal vsErrorSource As String _
                , Optional ByVal vlEmployeeNumber As Long = 0 _
                , Optional ByVal vsJobNumber As String = "" _
                , Optional ByVal vsCraftNumber As String = "" _
                , Optional ByVal vdtEffectiveStartDate As Date = #1/1/100# _
                , Optional ByVal vdtEffectiveEndDate As Date = #1/1/100# _
                , Optional ByVal viRecordSource As Integer = 0)
On Error GoTo AppendBRV_EH

    With radrsBRV
        If (.State <> adStateOpen) Then
            Exit Sub
        End If

        .AddNew

        !EMP_NUM = vlEmployeeNumber
        !JOB_NUM = vsJobNumber
        !CRAFT_NUM = vsCraftNumber
        !EFF_START_DATE = vdtEffectiveStartDate
        !EFF_END_DATE = vdtEffectiveEndDate

        !ERR_NUMBER = vlErrorNumber - vbObjectError
        !ERR_TYPE = Trim(vsErrorType)
        !REC_SOURCE = viRecordSource
        !ERR_DESCRIPTION = Trim(vsErrorDescription)
        !ERR_SOURCE = Trim(vsErrorSource)

        .Update
    End With

    Exit Sub
AppendBRV_EH:
    TrsRaiseError MODULENAME _
        , "AppendBRV"

End Sub
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'This method will save employee ineligibility entries
Public Function ProcessEmpIneligibilityEntries( _
                    ByVal vadrsIneligibilityInformation As ADODB.Recordset _
                    , ByRef radrsErrors As ADODB.Recordset _
                    , ByVal vlUserEmployeeNumber As Long _
                    ) As Boolean
On Error GoTo ProcessEmpIneligibilityEntries_EH
Dim lbEmptyRequest As Boolean, lsFilterString As String
Dim lobjdbEmployee As dbEmployee.cdEmployeeTrans
Dim llRecordIndex As Long

    'Check for request existence
    lbEmptyRequest = True
    If Not (vadrsIneligibilityInformation Is Nothing) Then
        If Not (vadrsIneligibilityInformation.BOF And _
            vadrsIneligibilityInformation.EOF) Then
            lbEmptyRequest = False
        End If
    End If

    'Prepare BRV Recordset
    Set radrsErrors = CreateIneligibilityBRV

    If (lbEmptyRequest) Then
        AppendIneligibilityBRV radrsErrors _
            , trErrbusEmployeeIneligibility _
            , mstrCBUSINESSERR _
            , "There are no Employee Ineligibility entries to process." _
            , MODULENAME & ": ProcessEmpIneligibilityEntries"
        GoTo REPORTBRV
    Else
        With vadrsIneligibilityInformation
            .Filter = adFilterNone
            If Not (.EOF) Then
                .MoveFirst
                If Not (ValidateIneligibilityEntries( _
                    vadrsIneligibilityInformation _
                    , radrsErrors _
                    , vlUserEmployeeNumber)) Then
                    GoTo REPORTBRV
                End If

                .Filter = adFilterNone
                .MoveFirst
            End If
        End With
    End If

    Set lobjdbEmployee = CreateObject( _
                            "dbEmployee.cdEmployeeTrans")
    'Set lobjdbEmployee = New dbEmployee.cdEmployee
    If Not (lobjdbEmployee.dbProcessIneligibilityEntries( _
            vadrsIneligibilityInformation _
            , vlUserEmployeeNumber)) Then
        Set lobjdbEmployee = Nothing
        Err.Raise trErrbusEmployeeIneligibility
    End If
    Set lobjdbEmployee = Nothing

    Set vadrsIneligibilityInformation = Nothing
    Set radrsErrors = Nothing

    GetObjectContext().SetComplete

    ProcessEmpIneligibilityEntries = True

    Set lobjdbEmployee = Nothing

    Exit Function
ProcessEmpIneligibilityEntries_EH:
    'Abort Transaction
    GetObjectContext().SetAbort
    Set lobjdbEmployee = Nothing

    Set vadrsIneligibilityInformation = Nothing
    Set radrsErrors = Nothing

    ProcessEmpIneligibilityEntries = False

    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "ProcessEmpIneligibilityEntries"
            Case trErrdbInsIntoEmpIneligibility, trErrdbUpdEmpIneligibility _
                , trErrdbDelFromEmpIneligibility, trErrdbSaveEmpIneligibility _
                , trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "ProcessEmpIneligibilityEntries" _
                    , Err.Number _
                    , , , , _
                    True
            Case Else
                TrsRaiseError MODULENAME _
                    , "ProcessEmpIneligibilityEntries" _
                    , trErrbusEmployeeIneligibility _
                    , , , , _
                    True
    End Select

    Exit Function

REPORTBRV:
    Set lobjdbEmployee = Nothing
    Set vadrsIneligibilityInformation = Nothing

    If Not (radrsErrors Is Nothing) Then
        With radrsErrors
            If Not (.EOF And .BOF) Then
                .MoveFirst
                .Filter = adFilterNone

                'Check for errors
'MDAC25 FILTER FIX
                lsFilterString = "ERR_TYPE = 'BE'"
                .Filter = stdmTRISCommon.MDAC25FilterFix(lsFilterString)
                lsFilterString = ""
                'If found abort else complete
                If Not (.EOF) Then
                    'Abort Transaction
                    GetObjectContext().SetAbort
                    ProcessEmpIneligibilityEntries = False
                Else
                    'Complete Transaction
                    GetObjectContext().SetComplete
                    ProcessEmpIneligibilityEntries = True
                End If

                .Filter = adFilterNone
            Else
                'Complete Transaction
                GetObjectContext().SetComplete
                ProcessEmpIneligibilityEntries = True
            End If
        End With
    Else
        'Complete Transaction
        GetObjectContext().SetComplete
        ProcessEmpIneligibilityEntries = True
    End If

    Exit Function

End Function
Private Function CreateIneligibilityBRV() As ADODB.Recordset
On Error GoTo CreateIneligibilityBRV_EH
Dim ladrs As New ADODB.Recordset

    With ladrs
        .Fields.Append "EMP_NUM", adInteger, , ADOFIELDATTR
        .Fields.Append "EXCEPTION_CD", adChar, 3, ADOFIELDATTR
        .Fields.Append "EFF_START_DATE", adDate, , ADOFIELDATTR
        .Fields.Append "EFF_END_DATE", adDate, , ADOFIELDATTR
        .Fields.Append "EFF_START_TIME", adChar, 5, ADOFIELDATTR
        .Fields.Append "EFF_END_TIME", adChar, 5, ADOFIELDATTR

        .Fields.Append "REC_SOURCE", adInteger, , ADOFIELDATTR
        .Fields.Append "ERR_NUMBER", adInteger, , ADOFIELDATTR
        .Fields.Append "ERR_TYPE", adChar, 2, ADOFIELDATTR
        .Fields.Append "ERR_DESCRIPTION", adChar, 300, ADOFIELDATTR
        .Fields.Append "ERR_SOURCE", adChar, 1000, ADOFIELDATTR

        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open
    End With

    Set CreateIneligibilityBRV = ladrs
    Set ladrs = Nothing

    Exit Function
CreateIneligibilityBRV_EH:
    Set ladrs = Nothing

    TrsRaiseError MODULENAME _
        , "CreateIneligibilityBRV"

End Function
Private Function ValidateIneligibilityEntries( _
                    ByRef radrsRSToValidate As ADODB.Recordset _
                    , ByRef radrsBRV As ADODB.Recordset _
                    , ByVal vlUserEmployeeNumber As Long _
                    ) As Boolean
On Error GoTo ValidateIneligibilityEntries_EH
Dim llRecordIndex As Long, lsFilterString As String

    ValidateIneligibilityEntries = False

    If Not (radrsRSToValidate Is Nothing) Then
        If Not (radrsRSToValidate.EOF And _
            radrsRSToValidate.BOF) Then
            radrsRSToValidate.Filter = adFilterNone
            radrsRSToValidate.MoveFirst
        Else
            Exit Function
        End If
    Else
        Exit Function
    End If

    If Not (CheckForValidLong(vlUserEmployeeNumber, True)) Then
        AppendIneligibilityBRV radrsBRV _
            , trErrbusEmployeeIneligibility _
            , mstrCBUSINESSERR _
            , "Invalid TRIS User." _
            , MODULENAME & ": ValidateIneligibilityEntries"
    End If

    With radrsRSToValidate
        For llRecordIndex = 1 To .RecordCount
            Select Case (UCase(Trim(!RECORD_STATUS)))
                Case "I", "C"
                    If Not (CheckForValidLong(!EMP_NUM, True)) Then
                        AppendIneligibilityBRV radrsBRV _
                            , trErrbusEmployeeIneligibility _
                            , mstrCBUSINESSERR _
                            , "Invalid Employee." _
                            , MODULENAME & ": ValidateIneligibilityEntries" _
                            , _
                            , !EXCEPTION_CD _
                            , !EFF_START_DATE _
                            , !EFF_START_TIME _
                            , !EFF_END_DATE _
                            , !EFF_END_TIME
                    Else
                        If Not (CheckForItemExistenceInDB(trEMPLOYEE, , !EMP_NUM)) Then
                            AppendIneligibilityBRV radrsBRV _
                                , trErrbusEmployeeIneligibility _
                                , mstrCBUSINESSERR _
                                , "Invalid Employee. Employee " & Trim(!EMP_NUM) & " does not exist." _
                                , MODULENAME & ": ValidateIneligibilityEntries" _
                                , !EMP_NUM _
                                , !EXCEPTION_CD _
                                , !EFF_START_DATE _
                                , !EFF_START_TIME _
                                , !EFF_END_DATE _
                                , !EFF_END_TIME
                        End If
                    End If

                    If Not (CheckForValidString(!EXCEPTION_CD)) Then
                        AppendIneligibilityBRV radrsBRV _
                            , trErrbusEmployeeIneligibility _
                            , mstrCBUSINESSERR _
                            , "Invalid Exception." _
                            , MODULENAME & ": ValidateIneligibilityEntries" _
                            , !EMP_NUM _
                            , _
                            , !EFF_START_DATE _
                            , !EFF_START_TIME _
                            , !EFF_END_DATE _
                            , !EFF_END_TIME
                    Else
                        If Not (CheckForItemExistenceInDB(trEXCEPTION, !EXCEPTION_CD)) Then
                            AppendIneligibilityBRV radrsBRV _
                                , trErrbusEmployeeIneligibility _
                                , mstrCBUSINESSERR _
                                , "Invalid Exception. Exception Code " & Trim(!EXCEPTION_CD) & " does not exist." _
                                , MODULENAME & ": ValidateIneligibilityEntries" _
                                , !EMP_NUM _
                                , !EXCEPTION_CD _
                                , !EFF_START_DATE _
                                , !EFF_START_TIME _
                                , !EFF_END_DATE _
                                , !EFF_END_TIME
                        End If
                    End If

                    If Not (CheckForValidDate(!EFF_START_DATE)) Then
                        AppendIneligibilityBRV radrsBRV _
                            , trErrbusEmployeeIneligibility _
                            , mstrCBUSINESSERR _
                            , "Invalid Start Date." _
                            , MODULENAME & ": ValidateIneligibilityEntries" _
                            , !EMP_NUM _
                            , !EXCEPTION_CD _
                            , _
                            , !EFF_START_TIME _
                            , !EFF_END_DATE _
                            , !EFF_END_TIME
                    End If

                    If Not (CheckForValidDate(!EFF_START_TIME)) Then
                        AppendIneligibilityBRV radrsBRV _
                            , trErrbusEmployeeIneligibility _
                            , mstrCBUSINESSERR _
                            , "Invalid Start Time." _
                            , MODULENAME & ": ValidateIneligibilityEntries" _
                            , !EMP_NUM _
                            , !EXCEPTION_CD _
                            , !EFF_START_DATE _
                            , _
                            , !EFF_END_DATE _
                            , !EFF_END_TIME
                    End If

                    If Not (CheckForValidDate(!EFF_END_TIME)) Then
                        AppendIneligibilityBRV radrsBRV _
                            , trErrbusEmployeeIneligibility _
                            , mstrCBUSINESSERR _
                            , "Invalid End Time." _
                            , MODULENAME & ": ValidateIneligibilityEntries" _
                            , !EMP_NUM _
                            , !EXCEPTION_CD _
                            , !EFF_START_DATE _
                            , !EFF_START_TIME _
                            , !EFF_END_DATE
                    End If

                    If (IsNull(!EFF_END_DATE)) Then
                        !EFF_END_DATE = #12/31/9999#
                        .Update
                    Else
                        If Not (IsDate(!EFF_END_DATE)) Then
                            AppendIneligibilityBRV radrsBRV _
                                , trErrbusEmployeeIneligibility _
                                , mstrCBUSINESSERR _
                                , "Invalid End Date." _
                                , MODULENAME & ": ValidateIneligibilityEntries" _
                                , !EMP_NUM _
                                , !EXCEPTION_CD _
                                , !EFF_START_DATE _
                                , !EFF_START_TIME _
                                , _
                                , !EFF_END_TIME
                        End If
                    End If
            End Select

            Select Case (UCase(Trim(!RECORD_STATUS)))
                Case "C", "D"
                    If Not (CheckForValidLong(!KEY_EMP_NUM)) Then
                        AppendIneligibilityBRV radrsBRV _
                            , trErrbusEmployeeIneligibility _
                            , mstrCBUSINESSERR _
                            , "Invalid Delete/Change request. Employee Number is invalid." _
                            , MODULENAME & ": ValidateIneligibilityEntries" _
                            , !EMP_NUM _
                            , !EXCEPTION_CD _
                            , !EFF_START_DATE _
                            , !EFF_START_TIME _
                            , !EFF_END_DATE _
                            , !EFF_END_TIME
                    End If

                    If Not (CheckForValidString(!KEY_EXCEPTION_CD)) Then
                        AppendIneligibilityBRV radrsBRV _
                            , trErrbusEmployeeIneligibility _
                            , mstrCBUSINESSERR _
                            , "Invalid Delete/Change request. Job number is invalid." _
                            , MODULENAME & ": ValidateIneligibilityEntries" _
                            , !EMP_NUM _
                            , !EXCEPTION_CD _
                            , !EFF_START_DATE _
                            , !EFF_START_TIME _
                            , !EFF_END_DATE _
                            , !EFF_END_TIME
                    End If

                    If Not (CheckForValidDate(!KEY_EFF_START_DATE)) Then
                        AppendIneligibilityBRV radrsBRV _
                            , trErrbusEmployeeIneligibility _
                            , mstrCBUSINESSERR _
                            , "Invalid Delete/Change request. Effective Start Date is invalid." _
                            , MODULENAME & ": ValidateIneligibilityEntries" _
                            , !EMP_NUM _
                            , !EXCEPTION_CD _
                            , !EFF_START_DATE _
                            , !EFF_START_TIME _
                            , !EFF_END_DATE _
                            , !EFF_END_TIME
                    End If

                    If Not (CheckForValidDate(!KEY_EFF_START_TIME)) Then
                        AppendIneligibilityBRV radrsBRV _
                            , trErrbusEmployeeIneligibility _
                            , mstrCBUSINESSERR _
                            , "Invalid Delete/Change request. Effective Start Time is invalid." _
                            , MODULENAME & ": ValidateIneligibilityEntries" _
                            , !EMP_NUM _
                            , !EXCEPTION_CD _
                            , !EFF_START_DATE _
                            , !EFF_START_TIME _
                            , !EFF_END_DATE _
                            , !EFF_END_TIME
                    End If

                    If Not (CheckForValidString(!TIME_STAMP _
                        , , , True, 26)) Then
                        AppendIneligibilityBRV radrsBRV _
                            , trErrbusEmployeeIneligibility _
                            , mstrCBUSINESSERR _
                            , "Invalid Delete/Change request. Time Stamp is invalid." _
                            , MODULENAME & ": ValidateIneligibilityEntries" _
                            , !EMP_NUM _
                            , !EXCEPTION_CD _
                            , !EFF_START_DATE _
                            , !EFF_START_TIME _
                            , !EFF_END_DATE _
                            , !EFF_END_TIME
                    End If
            End Select

            .MoveNext
        Next llRecordIndex
    End With

    radrsRSToValidate.Filter = adFilterNone
    radrsRSToValidate.MoveFirst

    ValidateIneligibilityEntries = True

    If Not (radrsBRV Is Nothing) Then
        With radrsBRV
            If Not (.BOF And .EOF) Then
                .Filter = adFilterNone
                .MoveFirst
                lsFilterString = "ERR_TYPE = 'BE' OR ERR_TYPE = 'AE' OR ERR_TYPE = 'AW'"
                .Filter = stdmTRISCommon.MDAC25FilterFix(lsFilterString)
                lsFilterString = ""
                If Not (.EOF) Then
                    ValidateIneligibilityEntries = False
                End If

                .Filter = adFilterNone
            End If
        End With
    End If

    Exit Function
ValidateIneligibilityEntries_EH:
    ValidateIneligibilityEntries = False

    TrsRaiseError MODULENAME _
            , "ValidateIneligibilityEntries"

End Function
Private Sub AppendIneligibilityBRV(ByRef radrsBRV As ADODB.Recordset _
                , ByVal vlErrorNumber As Long _
                , ByVal vsErrorType As String _
                , ByVal vsErrorDescription As String _
                , ByVal vsErrorSource As String _
                , Optional ByVal vlEmployeeNumber As Long = 0 _
                , Optional ByVal vsExceptionCode As String = "" _
                , Optional ByVal vdtEffectiveStartDate As Date = #1/1/100# _
                , Optional ByVal vsEffectiveStartTime As String = "" _
                , Optional ByVal vdtEffectiveEndDate As Date = #1/1/100# _
                , Optional ByVal vsEffectiveEndTime As String = "" _
                , Optional ByVal viRecordSource As Integer = 0 _
                )
On Error GoTo AppendBRV_EH

    With radrsBRV
        If (.State <> adStateOpen) Then
            Exit Sub
        End If

        .AddNew

        !EMP_NUM = vlEmployeeNumber
        !EXCEPTION_CD = vsExceptionCode
        !EFF_START_DATE = vdtEffectiveStartDate
        !EFF_START_TIME = vsEffectiveStartTime
        !EFF_END_DATE = vdtEffectiveEndDate
        !EFF_END_TIME = vsEffectiveEndTime

        !ERR_NUMBER = vlErrorNumber - vbObjectError
        !ERR_TYPE = Trim(vsErrorType)
        !REC_SOURCE = viRecordSource
        !ERR_DESCRIPTION = Trim(vsErrorDescription)
        !ERR_SOURCE = Trim(vsErrorSource)

        .Update
    End With

    Exit Sub
AppendBRV_EH:
    TrsRaiseError MODULENAME _
        , "AppendBRV"

End Sub
'This method will save volunteer entries
Public Function ProcessVolunteerEntries( _
                    ByVal vadrsVolunteerInformation As ADODB.Recordset _
                    , ByRef radrsErrors As ADODB.Recordset _
                    , ByVal vlUserEmployeeNumber As Long _
                    ) As Boolean
On Error GoTo ProcessVolunteerEntries_EH
Dim lbEmptyRequest As Boolean, lsFilterString As String
Dim lobjdbEmployee As dbEmployee.cdEmployeeTrans
Dim llRecordIndex As Long

    'Check for request existence
    lbEmptyRequest = True
    If Not (vadrsVolunteerInformation Is Nothing) Then
        If Not (vadrsVolunteerInformation.BOF And _
            vadrsVolunteerInformation.EOF) Then
            lbEmptyRequest = False
        End If
    End If

    'Prepare BRV Recordset
    Set radrsErrors = CreateVolunteerBRV

    If (lbEmptyRequest) Then
        AppendVolunteerBRV radrsErrors _
            , trErrbusVolunteer _
            , mstrCBUSINESSERR _
            , "There are no Volunteer entries to process." _
            , MODULENAME & ": ProcessVolunteerEntries"
        GoTo REPORTBRV
    Else
        With vadrsVolunteerInformation
            .Filter = adFilterNone
            If Not (.EOF) Then
                .MoveFirst
                If Not (ValidateVolunteerEntries( _
                    vadrsVolunteerInformation _
                    , radrsErrors _
                    , vlUserEmployeeNumber)) Then
                    GoTo REPORTBRV
                End If

                .Filter = adFilterNone
                .MoveFirst
            End If
        End With
    End If

    UpdateVolunteerList vadrsVolunteerInformation

    Set lobjdbEmployee = CreateObject( _
                            "dbEmployee.cdEmployeeTrans")
    'Set lobjdbEmployee = New dbEmployee.cdEmployee
    
    If Not (lobjdbEmployee.SaveVolunteerListChanges( _
            m_rsVolunteerChanges _
            , vlUserEmployeeNumber)) Then
        Set lobjdbEmployee = Nothing
        Err.Raise trErrbusVolunteer
    End If

    Set lobjdbEmployee = Nothing

    Set vadrsVolunteerInformation = Nothing
    Set radrsErrors = Nothing
    Set m_rsVolunteerChanges = Nothing
    
    GetObjectContext().SetComplete

    ProcessVolunteerEntries = True

    Set lobjdbEmployee = Nothing

    Exit Function
ProcessVolunteerEntries_EH:
    'Abort Transaction
    GetObjectContext().SetAbort
    Set lobjdbEmployee = Nothing

    Set vadrsVolunteerInformation = Nothing
    Set radrsErrors = Nothing
    Set m_rsVolunteerChanges = Nothing

    ProcessVolunteerEntries = False

    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "ProcessVolunteerEntries"
            Case trErrdbInsIntoVolunteer, trErrdbDelFromVolunteer _
                , trErrdbSaveVolunteer, trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "ProcessVolunteerEntries" _
                    , Err.Number _
                    , , , , _
                    True
            Case Else
                TrsRaiseError MODULENAME _
                    , "ProcessVolunteerEntries" _
                    , trErrbusVolunteer _
                    , , , , _
                    True
    End Select

    Exit Function

REPORTBRV:
    Set lobjdbEmployee = Nothing
    Set vadrsVolunteerInformation = Nothing

    If Not (radrsErrors Is Nothing) Then
        With radrsErrors
            If Not (.EOF And .BOF) Then
                .MoveFirst
                .Filter = adFilterNone

                'Check for errors
'MDAC25 FILTER FIX
                lsFilterString = "ERR_TYPE = 'BE'"
                .Filter = stdmTRISCommon.MDAC25FilterFix(lsFilterString)
                lsFilterString = ""
                'If found abort else complete
                If Not (.EOF) Then
                    'Abort Transaction
                    GetObjectContext().SetAbort
                    ProcessVolunteerEntries = False
                Else
                    'Complete Transaction
                    GetObjectContext().SetComplete
                    ProcessVolunteerEntries = True
                End If

                .Filter = adFilterNone
            Else
                'Complete Transaction
                GetObjectContext().SetComplete
                ProcessVolunteerEntries = True
            End If
        End With
    Else
        'Complete Transaction
        GetObjectContext().SetComplete
        ProcessVolunteerEntries = True
    End If

    Exit Function

End Function
Private Function CreateVolunteerBRV() As ADODB.Recordset
On Error GoTo CreateVolunteerBRV_EH
Dim ladrs As New ADODB.Recordset

    With ladrs
        .Fields.Append "EMP_NUM", adInteger, , ADOFIELDATTR

        .Fields.Append "REC_SOURCE", adInteger, , ADOFIELDATTR
        .Fields.Append "ERR_NUMBER", adInteger, , ADOFIELDATTR
        .Fields.Append "ERR_TYPE", adChar, 2, ADOFIELDATTR
        .Fields.Append "ERR_DESCRIPTION", adChar, 300, ADOFIELDATTR
        .Fields.Append "ERR_SOURCE", adChar, 1000, ADOFIELDATTR

        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open
    End With

    Set CreateVolunteerBRV = ladrs
    Set ladrs = Nothing

    Exit Function

CreateVolunteerBRV_EH:
    Set ladrs = Nothing

    TrsRaiseError MODULENAME _
        , "CreateVolunteerBRV"

End Function
Private Function ValidateVolunteerEntries( _
                    ByRef radrsRSToValidate As ADODB.Recordset _
                    , ByRef radrsBRV As ADODB.Recordset _
                    , ByVal vlUserEmployeeNumber As Long _
                    ) As Boolean
On Error GoTo ValidateVolunteerEntries_EH
Dim llRecordIndex As Long, lsFilterString As String
Dim lobjdbEmployee As dbEmployee.cdEmployee
Dim lsEmployees As String, ladrs As ADODB.Recordset

    ValidateVolunteerEntries = False

    If Not (radrsRSToValidate Is Nothing) Then
        If Not (radrsRSToValidate.EOF And _
            radrsRSToValidate.BOF) Then
            radrsRSToValidate.Filter = adFilterNone
            radrsRSToValidate.MoveFirst
        Else
            Exit Function
        End If
    Else
        Exit Function
    End If

    If Not (CheckForValidLong(vlUserEmployeeNumber, True)) Then
        AppendVolunteerBRV radrsBRV _
            , trErrbusVolunteer _
            , mstrCBUSINESSERR _
            , "Invalid TRIS User." _
            , MODULENAME & ": ValidateVolunteerEntries"
    End If

    With radrsRSToValidate
        For llRecordIndex = 1 To .RecordCount
            If Not (CheckForValidLong(!EMP_NUM, True)) Then
                AppendVolunteerBRV radrsBRV _
                    , trErrbusVolunteer _
                    , mstrCBUSINESSERR _
                    , "Invalid Employee." _
                    , MODULENAME & ": ValidateVolunteerEntries"
            Else
                lsEmployees = lsEmployees & _
                                CStr(!EMP_NUM) & ","
            End If

            .MoveNext
        Next llRecordIndex

        If (Trim(lsEmployees) <> "") Then
            If (Right(lsEmployees, 1) = ",") Then
                lsEmployees = Left(lsEmployees _
                                , Len(lsEmployees) - 1)
            End If
            Set lobjdbEmployee = CreateObject( _
                                "dbEmployee.cdEmployee")
            'Set lobjdbEmployee = New dbEmployee.cdEmployee
            Set ladrs = lobjdbEmployee.AreValidEmployees(lsEmployees)
            Set lobjdbEmployee = Nothing

            If Not (mADOExtensions.RecordsExist(ladrs)) Then
                AppendVolunteerBRV radrsBRV _
                    , trErrbusVolunteer _
                    , mstrCBUSINESSERR _
                    , "Invalid Employee(s). " _
                    , MODULENAME & ": ValidateVolunteerEntries" _
                    , !EMP_NUM
            Else
                'If the number of records returned is equal to
                'the number of records to be validated then
                'Note: Volunteer List contains unique employees
                If (ladrs.RecordCount <> .RecordCount) Then
                    AppendVolunteerBRV radrsBRV _
                        , trErrbusVolunteer _
                        , mstrCBUSINESSERR _
                        , "Invalid Employee(s). " _
                        , MODULENAME & ": ValidateVolunteerEntries" _
                        , !EMP_NUM
                End If
            End If
            Set ladrs = Nothing
        End If

        .MoveFirst
        For llRecordIndex = 1 To .RecordCount
            Select Case (UCase(Trim(!VLNTR_STATUS)))
                Case "YES"
                    'basically setting the Action Domain Value
                    !VLNTR_STATUS = "A"
                    .Update
                Case "NO"
                    !VLNTR_STATUS = "D"
                    .Update
                Case Else
                    AppendVolunteerBRV radrsBRV _
                        , trErrbusVolunteer _
                        , mstrCBUSINESSERR _
                        , "Invalid Volunteer Status." _
                        , MODULENAME & ": ValidateVolunteerEntries"
            End Select

            .MoveNext
        Next llRecordIndex
    End With

    radrsRSToValidate.Filter = adFilterNone
    radrsRSToValidate.MoveFirst

    ValidateVolunteerEntries = True

    If Not (radrsBRV Is Nothing) Then
        With radrsBRV
            If Not (.BOF And .EOF) Then
                .Filter = adFilterNone
                .MoveFirst
                lsFilterString = "ERR_TYPE = 'BE' OR ERR_TYPE = 'AE' OR ERR_TYPE = 'AW'"
                .Filter = stdmTRISCommon.MDAC25FilterFix(lsFilterString)
                lsFilterString = ""
                If Not (.EOF) Then
                    ValidateVolunteerEntries = False
                End If

                .Filter = adFilterNone
            End If
        End With
    End If

    Exit Function
ValidateVolunteerEntries_EH:
    GetObjectContext().SetAbort
    Set lobjdbEmployee = Nothing
    Set ladrs = Nothing
    ValidateVolunteerEntries = False

    TrsRaiseError MODULENAME _
            , "ValidateVolunteerEntries"

End Function
Private Sub AppendVolunteerBRV(ByRef radrsBRV As ADODB.Recordset _
                , ByVal vlErrorNumber As Long _
                , ByVal vsErrorType As String _
                , ByVal vsErrorDescription As String _
                , ByVal vsErrorSource As String _
                , Optional ByVal vlEmployeeNumber As Long = 0 _
                , Optional ByVal viRecordSource As Integer = 0 _
                )
On Error GoTo AppendBRV_EH

    With radrsBRV
        If (.State <> adStateOpen) Then
            Exit Sub
        End If

        .AddNew

        !EMP_NUM = vlEmployeeNumber

        !ERR_NUMBER = vlErrorNumber - vbObjectError
        !ERR_TYPE = Trim(vsErrorType)
        !REC_SOURCE = viRecordSource
        !ERR_DESCRIPTION = Trim(vsErrorDescription)
        !ERR_SOURCE = Trim(vsErrorSource)

        .Update
    End With

    Exit Sub
AppendBRV_EH:
    TrsRaiseError MODULENAME _
        , "AppendBRV"

End Sub
Public Function ProcessEmployeeECWorkReasons(ByVal vadrsEmployeeECWorkReasons As ADODB.Recordset _
                    , ByRef radrsErrors As ADODB.Recordset _
                    , ByVal vlUserEmployeeNumber As Long _
                    ) As Boolean
TrsRaiseError MODULENAME, "ProcessEmployeeECWorkReasons", trErrObsoleteMethod, , , , True
'On Error GoTo ProcessEmployeeECWorkReasons_EH
'Dim lbEmptyRequest As Boolean, lsFilterString As String
'Dim lobjdbEmployee As dbEmployee.cdEmployeeTrans
'Dim llRecordIndex As Long
'
'    'Check for request existence
'    lbEmptyRequest = True
'    If Not (vadrsEmployeeECWorkReasons Is Nothing) Then
'        If Not (vadrsEmployeeECWorkReasons.BOF And _
'            vadrsEmployeeECWorkReasons.EOF) Then
'            lbEmptyRequest = False
'        End If
'    End If
'
'    'Prepare BRV Recordset
'    Set radrsErrors = CreateEECWRBRV
'
'    If (lbEmptyRequest) Then
'        AppendEECWRBRV radrsErrors _
'            , trErrbusECWorkReasons _
'            , mstrCBUSINESSERR _
'            , "There are no E&C Employee Work Reason entries to process." _
'            , MODULENAME & ": ProcessEmployeeECWorkReasons"
'        GoTo REPORTBRV
'    Else
'        With vadrsEmployeeECWorkReasons
'            .Filter = adFilterNone
'            If Not (.EOF) Then
'                .MoveFirst
'                If Not (ValidateEmployeeECWorkReasons( _
'                    vadrsEmployeeECWorkReasons _
'                    , radrsErrors _
'                    , vlUserEmployeeNumber)) Then
'                    GoTo REPORTBRV
'                End If
'
'                .Filter = adFilterNone
'                .MoveFirst
'            End If
'        End With
'    End If
'
'    Set lobjdbEmployee = CreateObject( _
'                            "dbEmployee.cdEmployeeTrans")
'    'Set lobjdbEmployee = New dbEmployee.cdEmployee
'    If Not (lobjdbEmployee.dbProcessECWorkReasonEntries( _
'            vadrsEmployeeECWorkReasons _
'            , vlUserEmployeeNumber)) Then
'        Set lobjdbEmployee = Nothing
'        Err.Raise trErrbusECWorkReasons
'    End If
'    Set lobjdbEmployee = Nothing
'
'    Set vadrsEmployeeECWorkReasons = Nothing
'    Set radrsErrors = Nothing
'
'    GetObjectContext().SetComplete
'
'    ProcessEmployeeECWorkReasons = True
'
'    Set lobjdbEmployee = Nothing
'
'    Exit Function
'ProcessEmployeeECWorkReasons_EH:
'    'Abort Transaction
'    GetObjectContext().SetAbort
'
'    Set lobjdbEmployee = Nothing
'
'    Set vadrsEmployeeECWorkReasons = Nothing
'    Set radrsErrors = Nothing
'
'    ProcessEmployeeECWorkReasons = False
'
'    Select Case (Err.Number)
'            Case trErrutOpenConnection
'                TrsRaiseError MODULENAME _
'                    , "ProcessEmployeeECWorkReasons"
'            Case trErrdbDelFromECWorkReasons, trErrdbInsIntoECWorkReasons _
'                , trErrdbUpdECWorkReasons, trErrdbSaveECWorkReasons _
'                , trErrdb_ExecuteSQL_F
'                TrsRaiseError MODULENAME _
'                    , "ProcessEmployeeECWorkReasons" _
'                    , Err.Number _
'                    , , , , _
'                    True
'            Case Else
'                TrsRaiseError MODULENAME _
'                    , "ProcessEmployeeECWorkReasons" _
'                    , trErrbusECWorkReasons _
'                    , , , , _
'                    True
'    End Select
'
'    Exit Function
'
'REPORTBRV:
'    Set lobjdbEmployee = Nothing
'    Set vadrsEmployeeECWorkReasons = Nothing
'
'    If Not (radrsErrors Is Nothing) Then
'        With radrsErrors
'            If Not (.EOF And .BOF) Then
'                .MoveFirst
'                .Filter = adFilterNone
'
'                'Check for errors
''MDAC25 FILTER FIX
'                lsFilterString = "ERR_TYPE = 'BE'"
'                .Filter = stdmTRISCommon.MDAC25FilterFix(lsFilterString)
'                lsFilterString = ""
'                'If found abort else complete
'                If Not (.EOF) Then
'                    'Abort Transaction
'                    GetObjectContext().SetAbort
'                    ProcessEmployeeECWorkReasons = False
'                Else
'                    'Complete Transaction
'                    GetObjectContext().SetComplete
'                    ProcessEmployeeECWorkReasons = True
'                End If
'
'                .Filter = adFilterNone
'            Else
'                'Complete Transaction
'                GetObjectContext().SetComplete
'                ProcessEmployeeECWorkReasons = True
'            End If
'        End With
'    Else
'        'Complete Transaction
'        GetObjectContext().SetComplete
'        ProcessEmployeeECWorkReasons = True
'    End If
'
'    Exit Function
End Function
'Private Function CreateEECWRBRV() As ADODB.Recordset
'On Error GoTo CreateEECWRBRV_EH
'
'Dim ladrs As New ADODB.Recordset
'
'    With ladrs
'        .Fields.Append "EMP_NUM", adInteger, , ADOFIELDATTR
'
'        .Fields.Append "REC_SOURCE", adInteger, , ADOFIELDATTR
'        .Fields.Append "ERR_NUMBER", adInteger, , ADOFIELDATTR
'        .Fields.Append "ERR_TYPE", adChar, 2, ADOFIELDATTR
'        .Fields.Append "ERR_DESCRIPTION", adChar, 300, ADOFIELDATTR
'        .Fields.Append "ERR_SOURCE", adChar, 1000, ADOFIELDATTR
'
'        .CursorLocation = adUseClient
'        .CursorType = adOpenStatic
'        .LockType = adLockBatchOptimistic
'        .Open
'    End With
'
'    Set CreateEECWRBRV = ladrs
'    Set ladrs = Nothing
'
'    Exit Function
'
'CreateEECWRBRV_EH:
'    Set ladrs = Nothing
'
'    TrsRaiseError MODULENAME _
'        , "CreateEECWRBRV"
'
'End Function
'Private Function ValidateEmployeeECWorkReasons( _
'                    ByRef radrsRSToValidate As ADODB.Recordset _
'                    , ByRef radrsBRV As ADODB.Recordset _
'                    , ByVal vlUserEmployeeNumber As Long _
'                    ) As Boolean
'On Error GoTo ValidateEmployeeECWorkReasons_EH
'Dim llRecordIndex As Long, lsFilterString As String
'Dim lsEmployees As String
'Dim lobjdbEmployee As dbEmployee.cdEmployee
'Dim ladrs As ADODB.Recordset
'
'    ValidateEmployeeECWorkReasons = False
'
'    If Not (radrsRSToValidate Is Nothing) Then
'        If Not (radrsRSToValidate.EOF And _
'            radrsRSToValidate.BOF) Then
'            radrsRSToValidate.Filter = adFilterNone
'            radrsRSToValidate.MoveFirst
'        Else
'            Exit Function
'        End If
'    Else
'        Exit Function
'    End If
'
'    If Not (CheckForValidLong(vlUserEmployeeNumber, True)) Then
'        AppendEECWRBRV radrsBRV _
'            , trErrbusECWorkReasons _
'            , mstrCBUSINESSERR _
'            , "Invalid TRIS User." _
'            , MODULENAME & ": ValidateEmployeeECWorkReasons"
'    End If
'
'    With radrsRSToValidate
'        For llRecordIndex = 1 To .RecordCount
'            If Not (CheckForValidLong(!EMP_NUM, True)) Then
'                AppendEECWRBRV radrsBRV _
'                    , trErrbusECWorkReasons _
'                    , mstrCBUSINESSERR _
'                    , "Invalid Employee." _
'                    , MODULENAME & ": ValidateEmployeeECWorkReasons"
'            Else
'                lsEmployees = lsEmployees & _
'                                CStr(!EMP_NUM) & ","
'            End If
'
'            .MoveNext
'        Next llRecordIndex
'
'        If (Trim(lsEmployees) <> "") Then
'            If (Right(lsEmployees, 1) = ",") Then
'                lsEmployees = Left(lsEmployees _
'                                , Len(lsEmployees) - 1)
'            End If
'            Set lobjdbEmployee = CreateObject( _
'                                "dbEmployee.cdEmployee")
'            'Set lobjdbEmployee = New dbEmployee.cdEmployee
'            Set ladrs = lobjdbEmployee.AreValidEmployees(lsEmployees)
'            Set lobjdbEmployee = Nothing
'
'            If Not (mADOExtensions.RecordsExist(ladrs)) Then
'                AppendEECWRBRV radrsBRV _
'                    , trErrbusECWorkReasons _
'                    , mstrCBUSINESSERR _
'                    , "Invalid Employee(s). " _
'                    , MODULENAME & ": ValidateEmployeeECWorkReasons"
'            End If
'
'            Set ladrs = Nothing
'        End If
'
'        .MoveFirst
'        For llRecordIndex = 1 To .RecordCount
'            Select Case (UCase(Trim(!RECORD_STATUS)))
'                Case "I", "C"
'                    If Not (CheckForValidString(Trim(!WORK_REASON) _
'                        , True, 2)) Then
'                        AppendEECWRBRV radrsBRV _
'                            , trErrbusECWorkReasons _
'                            , mstrCBUSINESSERR _
'                            , "Invalid E&C Work Reason. " _
'                            , MODULENAME & ": ValidateEmployeeECWorkReasons" _
'                            , !EMP_NUM
'                    End If
'            End Select
'
'            Select Case (UCase(Trim(!RECORD_STATUS)))
'                Case "C", "D"
'                    If Not (CheckForValidLong(!KEY_EMP_NUM, True)) Then
'                        AppendEECWRBRV radrsBRV _
'                            , trErrbusECWorkReasons _
'                            , mstrCBUSINESSERR _
'                            , "Invalid Delete/Change request. Employee Number is invalid." _
'                            , MODULENAME & ": ValidateEmployeeECWorkReasons" _
'                            , !EMP_NUM
'                    End If
'
'                    If Not (CheckForValidString(!TIME_STAMP _
'                        , , , True, 26)) Then
'                        AppendEECWRBRV radrsBRV _
'                            , trErrbusECWorkReasons _
'                            , mstrCBUSINESSERR _
'                            , "Invalid Delete/Change request. Time Stamp is invalid." _
'                            , MODULENAME & ": ValidateEmployeeECWorkReasons" _
'                            , !EMP_NUM
'                    End If
'            End Select
'
'            .MoveNext
'        Next llRecordIndex
'    End With
'
'    radrsRSToValidate.Filter = adFilterNone
'    radrsRSToValidate.MoveFirst
'
'    ValidateEmployeeECWorkReasons = True
'
'    If Not (radrsBRV Is Nothing) Then
'        With radrsBRV
'            If Not (.BOF And .EOF) Then
'                .Filter = adFilterNone
'                .MoveFirst
'                lsFilterString = "ERR_TYPE = 'BE' OR ERR_TYPE = 'AE' OR ERR_TYPE = 'AW'"
'                .Filter = stdmTRISCommon.MDAC25FilterFix(lsFilterString)
'                lsFilterString = ""
'                If Not (.EOF) Then
'                    ValidateEmployeeECWorkReasons = False
'                End If
'
'                .Filter = adFilterNone
'            End If
'        End With
'    End If
'
'    Exit Function
'ValidateEmployeeECWorkReasons_EH:
'    GetObjectContext().SetAbort
'    Set lobjdbEmployee = Nothing
'    Set ladrs = Nothing
'    ValidateEmployeeECWorkReasons = False
'
'    TrsRaiseError MODULENAME _
'            , "ValidateEmployeeECWorkReasons"
'
'End Function
'Private Sub AppendEECWRBRV(ByRef radrsBRV As ADODB.Recordset _
'                , ByVal vlErrorNumber As Long _
'                , ByVal vsErrorType As String _
'                , ByVal vsErrorDescription As String _
'                , ByVal vsErrorSource As String _
'                , Optional ByVal vlEmployeeNumber As Long = 0 _
'                , Optional ByVal viRecordSource As Integer = 0 _
'                )
'On Error GoTo AppendEECWRBRV_EH
'
'    With radrsBRV
'        If (.State <> adStateOpen) Then
'            Exit Sub
'        End If
'
'        .AddNew
'
'        !EMP_NUM = vlEmployeeNumber
'
'        !ERR_NUMBER = vlErrorNumber - vbObjectError
'        !ERR_TYPE = Trim(vsErrorType)
'        !REC_SOURCE = viRecordSource
'        !ERR_DESCRIPTION = Trim(vsErrorDescription)
'        !ERR_SOURCE = Trim(vsErrorSource)
'
'        .Update
'    End With
'
'    Exit Sub
'AppendEECWRBRV_EH:
'    TrsRaiseError MODULENAME _
'        , "AppendEECWRBRV"
'
'End Sub
Private Function CheckForItemExistenceInDB(ByVal viItemType As Integer _
                    , Optional ByVal vsStringItem As String = "" _
                    , Optional ByVal vlLongItem As Long = 0 _
                    , Optional ByVal viIntegerItem As Integer = 0 _
                    , Optional ByVal vsAssociatedItem As String = "" _
                    ) As Boolean
On Error GoTo CheckForItemExistenceInDB_EH
Dim lobjdbEmployee As dbEmployee.cdEmployee
Dim lobjdbOTGroup As dbOTGroup.cdOTGroup
Dim lobjdbCodeTables As dbCodeTables.cdCodeTables
Dim lobjdbWorkgroup As dbWorkGroup.cdWorkGroup

    Select Case (viItemType)
        Case trCOMPANY, trDEPARTMENT _
            , trJOB, trQUALIFICATIONS _
            , trSPECIALSKILL, trUNION _
            , trOTBANKCODE, trLOCATION _
            , trIRAREA, trIRDIVISION _
            , trCRAFT, trEXCEPTION
            Set lobjdbCodeTables = CreateObject( _
                                    "dbCodeTables.cdCodeTables")
    End Select

    Select Case (viItemType)
        Case trCOMPANY
            CheckForItemExistenceInDB = _
                lobjdbCodeTables.IsAValidCompany(Trim(vsStringItem))

        Case trCRAFT
            CheckForItemExistenceInDB = _
                lobjdbCodeTables.IsAValidCraft(Trim(vsStringItem))

        Case trDEPARTMENT
            CheckForItemExistenceInDB = _
                lobjdbCodeTables.IsAValidDepartment(Trim(vsStringItem))

        Case trEMPLOYEE
            Set lobjdbEmployee = CreateObject( _
                                    "dbEmployee.cdEmployee")
            CheckForItemExistenceInDB = _
                lobjdbEmployee.IsValidEmployee(vlLongItem)
            Set lobjdbEmployee = Nothing

        Case trEXCEPTION
            CheckForItemExistenceInDB = _
                lobjdbCodeTables.ExceptionCodeIsValid(Trim(vsStringItem))

        Case trIRAREA
            CheckForItemExistenceInDB = _
                lobjdbCodeTables.IsAValidIndustrialRelationsArea(Trim(vsStringItem))

        Case trIRDIVISION
            CheckForItemExistenceInDB = _
                lobjdbCodeTables.IsAValidIndustrialRelationsDivision(Trim(vsStringItem))

        Case trJOB
            CheckForItemExistenceInDB = _
                lobjdbCodeTables.JobNumberIsValid(Trim(vsStringItem))

        Case trQUALIFICATIONS
            CheckForItemExistenceInDB = _
                lobjdbCodeTables.IsAValidQualificationCode(Trim(vsStringItem))

        Case trSPECIALSKILL
            CheckForItemExistenceInDB = _
                lobjdbCodeTables.IsAValidSpecialSkillCode( _
                    Trim(vsStringItem), Trim(vsAssociatedItem))

        Case trUNION
            CheckForItemExistenceInDB = _
                lobjdbCodeTables.UnionCodeIsValid(Trim(vsStringItem))

        Case trOTGROUP
            Set lobjdbOTGroup = CreateObject( _
                                    "dbOTGroup.cdOTGroup")
            CheckForItemExistenceInDB = _
                lobjdbOTGroup.IsAValidOvertimeGroup(vlLongItem)
            Set lobjdbOTGroup = Nothing

        Case trOTBANKCODE
            CheckForItemExistenceInDB = _
                lobjdbCodeTables.DoesBankExist(viIntegerItem)

        Case trLOCATION
            CheckForItemExistenceInDB = _
                lobjdbCodeTables.LocationCodeIsValid(CStr(viIntegerItem))

        Case trWORKGROUP
            Set lobjdbWorkgroup = CreateObject( _
                                    "dbWorkGroup.cdWorkGroup")
            CheckForItemExistenceInDB = _
                lobjdbWorkgroup.IsValidWorkgroup(vlLongItem, Date)
            Set lobjdbWorkgroup = Nothing
    End Select

    Select Case (viItemType)
        Case trCOMPANY, trDEPARTMENT _
            , trJOB, trQUALIFICATIONS _
            , trSPECIALSKILL, trUNION _
            , trOTBANKCODE, trLOCATION _
            , trIRAREA, trIRDIVISION _
            , trCRAFT, trEXCEPTION
            Set lobjdbCodeTables = Nothing
    End Select

    Exit Function
CheckForItemExistenceInDB_EH:
    GetObjectContext().SetAbort
    Set lobjdbOTGroup = Nothing
    Set lobjdbCodeTables = Nothing
    Set lobjdbWorkgroup = Nothing
    Set lobjdbEmployee = Nothing

    TrsRaiseError MODULENAME _
        , "CheckForItemExistenceInDB"

End Function
'END PREM
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!


Private Sub GetVolunteerState()
    
    Dim l_obj As dbEmployee.cdEmployee
    Dim l_obj2 As dbEmployee.cdEmployeeTrans
    
    On Error GoTo eh
  
    Set l_obj = CreateObject( _
            "dbEmployee.cdEmployee")
            
    Set m_rsVolunteerList = _
            l_obj.GetVolunteerInformation
    
    Set l_obj = Nothing
    
    Set l_obj2 = CreateObject( _
            "dbEmployee.cdEmployeeTrans")
    Set m_rsVolunteerChanges = _
            l_obj2.GetVolunteerChangeList
            
    Set l_obj2 = Nothing
    
    Exit Sub
eh:

    Set l_obj = Nothing
    Set l_obj2 = Nothing

    TrsRaiseError MODULENAME _
        , "GetVolunteerState"

    
End Sub

Private Function IsEmployeeAVolunteer( _
    ByVal v_lngEmpNum As Long, _
    ByVal v_rs As ADODB.Recordset _
    ) As Boolean
    
    Dim l_strFilter As String
    
    On Error GoTo eh
    
    l_strFilter = "EMP_NUM = " & str(v_lngEmpNum) _
                & " AND VLNTR_STATUS = 'YES' "
                
    l_strFilter = MDAC25FilterFix(l_strFilter)
    
    v_rs.Filter = l_strFilter
    
    If RecordsExist(v_rs) Then
        IsEmployeeAVolunteer = True
    Else
        IsEmployeeAVolunteer = False
    End If
       
    v_rs.Filter = ""
    
    Exit Function
eh:

    TrsRaiseError MODULENAME _
        , "IsEmployeeAVolunteer"
    

End Function

Private Sub UpdateVolunteerList( _
    ByVal v_rsReq As ADODB.Recordset)

    Dim l_blnVolunteer As Boolean
    Dim l_blnOnChangeList As Boolean
    
    On Error GoTo eh

    GetVolunteerState

    v_rsReq.MoveFirst
    While Not v_rsReq.EOF
        l_blnVolunteer = IsEmployeeAVolunteer(v_rsReq!EMP_NUM, m_rsVolunteerList)
        m_rsVolunteerChanges.Filter = "EMP_NUM = " & CStr(v_rsReq!EMP_NUM)
        l_blnOnChangeList = RecordsExist(m_rsVolunteerChanges)
        If Trim(v_rsReq!VLNTR_STATUS) = "A" Then
            If l_blnVolunteer Then
                If l_blnOnChangeList Then
                    m_rsVolunteerChanges.Delete
                End If
            Else
                If Not l_blnOnChangeList Then
                    m_rsVolunteerChanges.AddNew
                    m_rsVolunteerChanges!EMP_NUM = v_rsReq!EMP_NUM
                    m_rsVolunteerChanges!ACTN_CD = "A"
                    m_rsVolunteerChanges!STUS_EFF_DT = Date
                End If
            End If
        Else
            If l_blnVolunteer Then
                If Not l_blnOnChangeList Then
                    m_rsVolunteerChanges.AddNew
                    m_rsVolunteerChanges!EMP_NUM = v_rsReq!EMP_NUM
                    m_rsVolunteerChanges!ACTN_CD = "D"
                    m_rsVolunteerChanges!STUS_EFF_DT = Date
                End If
            Else
                If l_blnOnChangeList Then
                    m_rsVolunteerChanges.Delete
                End If
            End If
        End If
        
        m_rsVolunteerChanges.Filter = ""
        v_rsReq.MoveNext
    
    Wend

    
    Set m_rsVolunteerList = Nothing
    
    Exit Sub
    
eh:
    Set m_rsVolunteerChanges = Nothing
    Set m_rsVolunteerList = Nothing
    TrsRaiseError MODULENAME _
        , "UpdateVolunteerList"
     
End Sub

