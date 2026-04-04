VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "cbsCodeTablesTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'AMEREN Project: TRIS Conversion Phase 3
'February 2004
'Author: Prem M. Chander
'Module Description: This class will provide the user the ability
'to maintain (Insert, Update and Delete) Overtime Code Table information

'COMMENTS: OCT, 2005 - CHANGED GETOBJECTCONTEXT.CREATEINSTANCE
'TO CREATEOBJECT -- PREM

Option Explicit

Private Const ADOFIELDATTR = adFldUpdatable Or adFldIsNullable
Private Const MODULENAME As String = "busCodeTables.cbsCodeTablesTrans"

Private Const trCRAFTS As Integer = 1
Private Const trQUALIFICATIONS As Integer = 2
Private Const trSPECIALSKILLS As Integer = 3
Private Const trOTREASONS As Integer = 4
Private Const trOTRESPONSES As Integer = 5
Private Const trASSOCIATECRAFTTOJOB As Integer = 6
Private Const trOTBANKHEADERS As Integer = 7
Private Const trDEPARTMENT As Integer = 8
Private Const trJOB As Integer = 9
Private Const trQUALUSED As Integer = 10
Private Const trSKILLUSED As Integer = 11
Private Const trCRAFTUSED As Integer = 12
Private Const trOTREASONUSED As Integer = 13
Private Const trOTRESPONSEUSED As Integer = 14
Private Const trOTBANKUSED As Integer = 15
Private Function CreateBRV(ByVal vsProcessID As String _
            ) As ADODB.Recordset
On Error GoTo CreateBRV_EH

Dim ladrs As New ADODB.Recordset

    With ladrs
        Select Case (UCase(Trim(vsProcessID)))
            Case "OTCDTBL"
                .Fields.Append "JOB_NUM", adChar, 6, ADOFIELDATTR
                .Fields.Append "CRAFT_NUM", adChar, 3, ADOFIELDATTR
                .Fields.Append "QUAL_CD", adChar, 3, ADOFIELDATTR
                .Fields.Append "SPEC_SKILL_CD", adChar, 3, ADOFIELDATTR
                .Fields.Append "OT_REASON_CD", adChar, 3, ADOFIELDATTR
                .Fields.Append "OT_RESP_CD", adChar, 5, ADOFIELDATTR
                .Fields.Append "OT_BANK_CD", adInteger, , ADOFIELDATTR

            Case "PROMLINES"
                .Fields.Append "DEPT_NUM", adChar, 3, ADOFIELDATTR
                .Fields.Append "HR_LOC", adChar, 1, ADOFIELDATTR
                .Fields.Append "JOB_NUM", adChar, 6, ADOFIELDATTR
                .Fields.Append "IR_AREA_CD", adChar, 2, ADOFIELDATTR
                .Fields.Append "IR_DIV_CD", adChar, 2, ADOFIELDATTR
                .Fields.Append "LOP_CURR_LEV_IND", adChar, 2, ADOFIELDATTR
                .Fields.Append "LOP_NEXT_LEV_IND", adChar, 2, ADOFIELDATTR
        End Select

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

    Set CreateBRV = ladrs
    Set ladrs = Nothing

    Exit Function

CreateBRV_EH:
    Set ladrs = Nothing
    Set CreateBRV = Nothing

    TrsRaiseError MODULENAME _
        , "CreateBRV"

End Function
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'This method will save OT Code Table entries
Public Function ProcessOTCodeTableEntries( _
                    ByVal vadrsOTCodeTableInformation As ADODB.Recordset _
                    , ByRef radrsErrors As ADODB.Recordset _
                    , ByVal vlUserEmployeeNumber As Long _
                    ) As Boolean
On Error GoTo ProcessOTCodeTableEntries_EH
Dim lbEmptyRequest As Boolean, lsFilterString As String
Dim lobjdbCodeTables As dbCodeTables.cdCodeTablesTrans
Dim llIndex As Long, liCodeTableType As Integer

    'Check for request existence
    lbEmptyRequest = True
    If Not (vadrsOTCodeTableInformation Is Nothing) Then
        If Not (vadrsOTCodeTableInformation.BOF And _
            vadrsOTCodeTableInformation.EOF) Then
            lbEmptyRequest = False
        End If
    End If

    'Prepare BRV Recordset
    Set radrsErrors = CreateBRV("OTCDTBL")

    If (lbEmptyRequest) Then
        AppendOTCDTBLBRV radrsErrors _
            , trErrbusOTCodeTables _
            , mstrCBUSINESSERR _
            , "There are no Code Table entries to process." _
            , MODULENAME & ": ProcessOTCodeTableEntries"
        GoTo REPORTBRV
    Else
        With vadrsOTCodeTableInformation
            'This is done so that this module can behave
            'generically, this is really not a good way though
            'it will work as long as the fields are ordered correctly.
            For llIndex = 0 To .Fields.Count - 1
                Select Case (UCase(Trim(.Fields(llIndex).Name)))
                        Case "CRAFT_NUM"
                            liCodeTableType = trCRAFTS
                            Exit For
                        Case "QUAL_CD"
                            liCodeTableType = trQUALIFICATIONS
                            Exit For
                        Case "SPEC_SKILL_CD"
                            liCodeTableType = trSPECIALSKILLS
                            Exit For
                        Case "OT_REASON_CD"
                            liCodeTableType = trOTREASONS
                            Exit For
                        Case "OT_RESP_CD"
                            liCodeTableType = trOTRESPONSES
                            Exit For
                        Case "JOB_NUM"
                            liCodeTableType = trASSOCIATECRAFTTOJOB
                            Exit For
                        Case "OT_BANK_CD"
                            liCodeTableType = trOTBANKHEADERS
                            Exit For
                End Select
            Next llIndex

            If (liCodeTableType = trCRAFTS) Then
                'NOTE: JOURN_APPR is defined as Type Character and having Length 1
                .Filter = adFilterNone
                'MDAC25 FILTER FIX
                lsFilterString = "JOURN_APPR = 'JOURNEY'"
                .Filter = stdmTRISCommon.MDAC25FilterFix(lsFilterString)
                lsFilterString = ""
                If Not (.EOF) Then
                    .MoveFirst
                    For llIndex = 1 To .RecordCount
                        !JOURN_APPR = "J"
                        .Update
                        .MoveNext
                    Next llIndex
                End If

                .Filter = adFilterNone
                'MDAC25 FILTER FIX
                lsFilterString = "JOURN_APPR = 'APPRENTICE'"
                .Filter = stdmTRISCommon.MDAC25FilterFix(lsFilterString)
                lsFilterString = ""
                If Not (.EOF) Then
                    .MoveFirst
                    For llIndex = 1 To .RecordCount
                        !JOURN_APPR = "A"
                        .Update
                        .MoveNext
                    Next llIndex
                End If
                .Filter = adFilterNone
            End If

            'Domain values for OT_CHRG_IND include C and N
            If (liCodeTableType = trOTRESPONSES) Then
                .Filter = adFilterNone
                'MDAC25 FILTER FIX
                lsFilterString = "OT_CHRG_IND = 'Y'"
                .Filter = stdmTRISCommon.MDAC25FilterFix(lsFilterString)
                lsFilterString = ""
                If Not (.EOF) Then
                    .MoveFirst
                    For llIndex = 1 To .RecordCount
                        !OT_CHRG_IND = "C"
                        .Update
                        .MoveNext
                    Next llIndex
                End If
                .Filter = adFilterNone
            End If

            'Validate
            If Not (.EOF) Then
                .MoveFirst
                If Not (ValidateOTCodeTableEntries( _
                    vadrsOTCodeTableInformation _
                    , radrsErrors _
                    , vlUserEmployeeNumber _
                    , liCodeTableType)) Then
                    GoTo REPORTBRV
                End If

                .Filter = adFilterNone
                .MoveFirst
            End If
        End With
    End If

    'Process SQL
    Set lobjdbCodeTables = CreateObject( _
                                "dbCodeTables.cdCodeTablesTrans")
    'Set lobjdbCodeTables = New dbCodeTables.cdCodeTablesTrans
    Select Case (liCodeTableType)
            Case trCRAFTS
                If Not (lobjdbCodeTables.dbProcessCraftSave( _
                        vadrsOTCodeTableInformation _
                        , vlUserEmployeeNumber)) Then
                    Set lobjdbCodeTables = Nothing

                    Err.Raise trErrbusOTCodeTables
                End If

            Case trQUALIFICATIONS
                If Not (lobjdbCodeTables.dbProcessQualificationsSave( _
                        vadrsOTCodeTableInformation _
                        , vlUserEmployeeNumber)) Then
                    Set lobjdbCodeTables = Nothing

                    Err.Raise trErrbusOTCodeTables
                End If

            Case trSPECIALSKILLS
                If Not (lobjdbCodeTables.dbProcessSpecialSkillsSave( _
                        vadrsOTCodeTableInformation _
                        , vlUserEmployeeNumber)) Then
                    Set lobjdbCodeTables = Nothing

                    Err.Raise trErrbusOTCodeTables
                End If

            Case trOTREASONS
                If Not (lobjdbCodeTables.dbProcessOTReasonsSave( _
                        vadrsOTCodeTableInformation _
                        , vlUserEmployeeNumber)) Then
                    Set lobjdbCodeTables = Nothing

                    Err.Raise trErrbusOTCodeTables
                End If

            Case trOTRESPONSES
                If Not (lobjdbCodeTables.dbProcessOTResponsesSave( _
                        vadrsOTCodeTableInformation _
                        , vlUserEmployeeNumber)) Then
                    Set lobjdbCodeTables = Nothing

                    Err.Raise trErrbusOTCodeTables
                End If

            Case trASSOCIATECRAFTTOJOB
                If Not (lobjdbCodeTables.dbProcessAssociateCraftToJob( _
                        vadrsOTCodeTableInformation _
                        , vlUserEmployeeNumber)) Then
                    Set lobjdbCodeTables = Nothing

                    Err.Raise trErrbusOTCodeTables
                End If

            Case trOTBANKHEADERS
                If Not (lobjdbCodeTables.dbProcessOTBankHeaderSave( _
                        vadrsOTCodeTableInformation _
                        , vlUserEmployeeNumber)) Then
                    Set lobjdbCodeTables = Nothing

                    Err.Raise trErrbusOTCodeTables
                End If
    End Select

    'Destroy
    Set lobjdbCodeTables = Nothing
    Set vadrsOTCodeTableInformation = Nothing
    Set radrsErrors = Nothing

    'Return
    ProcessOTCodeTableEntries = True

    'Complete
    GetObjectContext().SetComplete

    Exit Function
ProcessOTCodeTableEntries_EH:
    'Abort Transaction
    GetObjectContext().SetAbort

    Set vadrsOTCodeTableInformation = Nothing
    Set radrsErrors = Nothing
    Set lobjdbCodeTables = Nothing

    ProcessOTCodeTableEntries = False

    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "ProcessOTCodeTableEntries"
            Case trErrdbInsIntoOTCDTBLCraft, trErrdbUpdOTCDTBLCraft _
                , trErrdbDelFromOTCDTBLCraft, trErrdbOTCDTBLSaveCraft _
                , trErrdbInsIntoOTCDTBLQuals, trErrdbUpdOTCDTBLQuals _
                , trErrdbDelFromOTCDTBLQuals, trErrdbOTCDTBLSaveQuals _
                , trErrdbInsIntoOTCDTBLSkills, trErrdbUpdOTCDTBLSkills _
                , trErrdbDelFromOTCDTBLSkills, trErrdbOTCDTBLSaveSkills _
                , trErrdbInsIntoOTCDTBLReasons, trErrdbUpdOTCDTBLReasons _
                , trErrdbDelFromOTCDTBLReasons, trErrdbOTCDTBLSaveReasons _
                , trErrdbInsIntoOTCDTBLResponses, trErrdbUpdOTCDTBLResponses _
                , trErrdbDelFromOTCDTBLResponses, trErrdbOTCDTBLSaveResponses _
                , trErrdbUpdOTCDTBLCraftJob, trErrdbOTCDTBLSaveCraftJob _
                , trErrdbDelFromOTBankCode, trErrdbInsIntoOTBankCode _
                , trErrdbUpdOTBankCode, trErrdbSaveOTBankCode _
                , trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "ProcessOTCodeTableEntries" _
                    , Err.Number _
                    , , , , _
                    True
            Case Else
                TrsRaiseError MODULENAME _
                    , "ProcessOTCodeTableEntries" _
                    , trErrbusOTCodeTables _
                    , , , , _
                    True
    End Select

    Exit Function

REPORTBRV:
    Set lobjdbCodeTables = Nothing
    Set vadrsOTCodeTableInformation = Nothing

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
                    ProcessOTCodeTableEntries = False
                Else
                    'Complete Transaction
                    GetObjectContext().SetComplete
                    ProcessOTCodeTableEntries = True
                End If

                .Filter = adFilterNone
            Else
                'Complete Transaction
                GetObjectContext().SetComplete
                ProcessOTCodeTableEntries = True
            End If
        End With
    Else
        'Complete Transaction
        GetObjectContext().SetComplete
        ProcessOTCodeTableEntries = True
    End If

    Exit Function

End Function
Private Function ValidateOTCodeTableEntries( _
                    ByRef radrsRSToValidate As ADODB.Recordset _
                    , ByRef radrsBRV As ADODB.Recordset _
                    , ByVal vlUserEmployeeNumber As Long _
                    , ByVal viOTCodeTableType As Integer _
                    ) As Boolean
On Error GoTo ValidateOTCodeTableEntries_EH
Dim llRecordIndex As Long, lsFilterString As String

    ValidateOTCodeTableEntries = False

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
        AppendOTCDTBLBRV radrsBRV _
            , trErrbusOTCodeTables _
            , mstrCBUSINESSERR _
            , "Invalid TRIS User." _
            , MODULENAME & ": ValidateOTCodeTableEntries"
    End If

    With radrsRSToValidate
        For llRecordIndex = 1 To .RecordCount
            Select Case (viOTCodeTableType)
                    Case trCRAFTS
                        Select Case (UCase(Trim(!RECORD_STATUS)))
                            Case "I", "C"
                                If Not (CheckForValidString(!CRAFT_NUM)) Then
                                    AppendOTCDTBLBRV radrsBRV _
                                        , trErrbusOTCodeTables _
                                        , mstrCBUSINESSERR _
                                        , "Invalid Craft." _
                                        , MODULENAME & ": ValidateOTCodeTableEntries" _
                                        , , , , , , , , 1
                                End If

                                If Not (CheckForValidString(!CRAFT_DESC, True, 30)) Then
                                    AppendOTCDTBLBRV radrsBRV _
                                        , trErrbusOTCodeTables _
                                        , mstrCBUSINESSERR _
                                        , "Invalid Craft Description." _
                                        , MODULENAME & ": ValidateOTCodeTableEntries" _
                                        , _
                                        , !CRAFT_NUM _
                                        , , , , , , 1
                                Else
                                    !CRAFT_DESC = stdmTRISCommon.ParseString(!CRAFT_DESC)
                                    .Update
                                End If

                                If Not (CheckForValidString(!CRAFT_GROUP, True, 30)) Then
                                    AppendOTCDTBLBRV radrsBRV _
                                        , trErrbusOTCodeTables _
                                        , mstrCBUSINESSINF _
                                        , "Craft Group was empty." _
                                        , MODULENAME & ": ValidateOTCodeTableEntries" _
                                        , _
                                        , !CRAFT_NUM _
                                        , , , , , , 1
                                Else
                                    !CRAFT_DESC = stdmTRISCommon.ParseString(!CRAFT_DESC)
                                    .Update
                                End If
                        End Select

                        Select Case (UCase(Trim(!RECORD_STATUS)))
                            Case "C", "D"
                                If Not (CheckForValidString(!KEY_CRAFT_NUM)) Then
                                    AppendOTCDTBLBRV radrsBRV _
                                        , trErrbusOTCodeTables _
                                        , mstrCBUSINESSERR _
                                        , "Invalid Delete/Change request. Craft is invalid." _
                                        , MODULENAME & ": ValidateOTCodeTableEntries" _
                                        , _
                                        , !CRAFT_NUM _
                                        , , , , , , 1
                                Else
                                    If (CheckForValidCodeTableItem(trCRAFTUSED, !KEY_CRAFT_NUM)) Then
                                        AppendOTCDTBLBRV radrsBRV _
                                            , trErrbusOTCodeTables _
                                            , mstrCBUSINESSERR _
                                            , "Invalid Delete/Change request. Craft is currently being used in the application and cannot be changed or deleted." _
                                            , MODULENAME & ": ValidateOTCodeTableEntries" _
                                            , _
                                            , !CRAFT_NUM _
                                            , , , , , , 1
                                    End If
                                End If
                        End Select

                    Case trQUALIFICATIONS
                        Select Case (UCase(Trim(!RECORD_STATUS)))
                            Case "I", "C"
                                If Not (CheckForValidString(!QUAL_CD)) Then
                                    AppendOTCDTBLBRV radrsBRV _
                                        , trErrbusOTCodeTables _
                                        , mstrCBUSINESSERR _
                                        , "Invalid Qualification." _
                                        , MODULENAME & ": ValidateOTCodeTableEntries" _
                                        , , , , , , , , 2
                                End If

                                If Not (CheckForValidString(!QUAL_DESC, True, 30)) Then
                                    AppendOTCDTBLBRV radrsBRV _
                                        , trErrbusOTCodeTables _
                                        , mstrCBUSINESSERR _
                                        , "Invalid Qualification Description." _
                                        , MODULENAME & ": ValidateOTCodeTableEntries" _
                                        , , _
                                        , !QUAL_CD _
                                        , , , , , 2
                                Else
                                    !QUAL_DESC = stdmTRISCommon.ParseString(!QUAL_DESC)
                                    .Update
                                End If
                        End Select

                        Select Case (UCase(Trim(!RECORD_STATUS)))
                            Case "C", "D"
                                If Not (CheckForValidString(!KEY_QUAL_CD)) Then
                                    AppendOTCDTBLBRV radrsBRV _
                                        , trErrbusOTCodeTables _
                                        , mstrCBUSINESSERR _
                                        , "Invalid Delete/Change request. Qualification is invalid." _
                                        , MODULENAME & ": ValidateOTCodeTableEntries" _
                                        , , _
                                        , !QUAL_CD _
                                        , , , , , 2
                                Else
                                    If (CheckForValidCodeTableItem(trQUALUSED, !KEY_QUAL_CD)) Then
                                        AppendOTCDTBLBRV radrsBRV _
                                            , trErrbusOTCodeTables _
                                            , mstrCBUSINESSERR _
                                            , "Invalid Delete/Change request. Qualification is currently being used in the application and cannot be changed or deleted." _
                                            , MODULENAME & ": ValidateOTCodeTableEntries" _
                                            , , _
                                            , !QUAL_CD _
                                            , , , , , 2
                                    End If
                                End If
                        End Select

                    Case trSPECIALSKILLS
                        Select Case (UCase(Trim(!RECORD_STATUS)))
                            Case "I", "C"
                                If Not (CheckForValidString(!SPEC_SKILL_CD)) Then
                                    AppendOTCDTBLBRV radrsBRV _
                                        , trErrbusOTCodeTables _
                                        , mstrCBUSINESSERR _
                                        , "Invalid Special Skill." _
                                        , MODULENAME & ": ValidateOTCodeTableEntries" _
                                        , , , , , , , , 3
                                End If

                                If Not (CheckForValidString(!SPEC_SKILL_DESC, True, 30)) Then
                                    AppendOTCDTBLBRV radrsBRV _
                                        , trErrbusOTCodeTables _
                                        , mstrCBUSINESSERR _
                                        , "Invalid Special Skill Description." _
                                        , MODULENAME & ": ValidateOTCodeTableEntries" _
                                        , , _
                                        , !QUAL_CD _
                                        , !SPEC_SKILL_CD _
                                        , , , , 3
                                Else
                                    !SPEC_SKILL_DESC = stdmTRISCommon.ParseString(!SPEC_SKILL_DESC)
                                    .Update
                                End If

                                If Not (CheckForValidString(!QUAL_CD)) Then
                                    AppendOTCDTBLBRV radrsBRV _
                                        , trErrbusOTCodeTables _
                                        , mstrCBUSINESSERR _
                                        , "Invalid Qualifications associated to the Special Skill." _
                                        , MODULENAME & ": ValidateOTCodeTableEntries" _
                                        , , , _
                                        , !SPEC_SKILL_CD _
                                        , , , , 3
                                Else
                                    If Not (CheckForValidCodeTableItem(trQUALIFICATIONS, !QUAL_CD)) Then
                                        AppendOTCDTBLBRV radrsBRV _
                                            , trErrbusOTCodeTables _
                                            , mstrCBUSINESSERR _
                                            , "Invalid Qualifications Code. The associated Qualification " & _
                                                    Trim(!QUAL_CD) & " does not exist." _
                                            , MODULENAME & ": ValidateOTCodeTableEntries" _
                                            , , _
                                            , !QUAL_CD _
                                            , !SPEC_SKILL_CD _
                                            , , , , 3
                                    End If
                                End If
                        End Select

                        Select Case (UCase(Trim(!RECORD_STATUS)))
                            Case "C", "D"
                                If Not (CheckForValidString(!KEY_SPEC_SKILL_CD)) Then
                                    AppendOTCDTBLBRV radrsBRV _
                                        , trErrbusOTCodeTables _
                                        , mstrCBUSINESSERR _
                                        , "Invalid Delete/Change request. Special Skill is invalid." _
                                        , MODULENAME & ": ValidateOTCodeTableEntries" _
                                        , , _
                                        , !QUAL_CD _
                                        , !SPEC_SKILL_CD _
                                        , , , , 3
                                Else
                                    If (CheckForValidCodeTableItem(trSKILLUSED _
                                        , !KEY_SPEC_SKILL_CD _
                                        , , _
                                        , !KEY_QUAL_CD)) Then
                                        AppendOTCDTBLBRV radrsBRV _
                                            , trErrbusOTCodeTables _
                                            , mstrCBUSINESSERR _
                                            , "Invalid Delete/Change request. Special Skill is currently being used in the application and cannot be changed or deleted." _
                                            , MODULENAME & ": ValidateOTCodeTableEntries" _
                                            , , _
                                            , !QUAL_CD _
                                            , !SPEC_SKILL_CD _
                                            , , , , 3
                                    End If
                                End If

                                If Not (CheckForValidString(!KEY_QUAL_CD)) Then
                                    AppendOTCDTBLBRV radrsBRV _
                                        , trErrbusOTCodeTables _
                                        , mstrCBUSINESSERR _
                                        , "Invalid Delete/Change request. Qualification associated to the Special Skills code is invalid." _
                                        , MODULENAME & ": ValidateOTCodeTableEntries" _
                                        , , _
                                        , !QUAL_CD _
                                        , !SPEC_SKILL_CD _
                                        , , , , 3
                                End If
                        End Select

                    Case trOTREASONS
                        Select Case (UCase(Trim(!RECORD_STATUS)))
                            Case "I", "C"
                                If Not (CheckForValidString(!OT_REASON_CD)) Then
                                    AppendOTCDTBLBRV radrsBRV _
                                        , trErrbusOTCodeTables _
                                        , mstrCBUSINESSERR _
                                        , "Invalid Overtime Reason." _
                                        , MODULENAME & ": ValidateOTCodeTableEntries" _
                                        , , , , , , , , 4
                                End If

                                If Not (CheckForValidString(!OT_REASON_DESC, True, 30)) Then
                                    AppendOTCDTBLBRV radrsBRV _
                                        , trErrbusOTCodeTables _
                                        , mstrCBUSINESSERR _
                                        , "Invalid Overtime Reason Description." _
                                        , MODULENAME & ": ValidateOTCodeTableEntries" _
                                        , , , , _
                                        , !OT_REASON_CD _
                                        , , , 4
                                Else
                                    !OT_REASON_DESC = stdmTRISCommon.ParseString(!OT_REASON_DESC)
                                    .Update
                                End If
                        End Select

                        Select Case (UCase(Trim(!RECORD_STATUS)))
                            Case "C", "D"
                                If Not (CheckForValidString(!KEY_OT_REASON_CD)) Then
                                    AppendOTCDTBLBRV radrsBRV _
                                        , trErrbusOTCodeTables _
                                        , mstrCBUSINESSERR _
                                        , "Invalid Delete/Change request. Overtime Reason is invalid." _
                                        , MODULENAME & ": ValidateOTCodeTableEntries" _
                                        , , , , _
                                        , !OT_REASON_CD _
                                        , , , 4
                                Else
                                    If (CheckForValidCodeTableItem(trOTREASONUSED, !KEY_OT_REASON_CD)) Then
                                        AppendOTCDTBLBRV radrsBRV _
                                            , trErrbusOTCodeTables _
                                            , mstrCBUSINESSERR _
                                            , "Invalid Delete/Change request. Overtime reason is currently being used in the application and cannot be changed or deleted." _
                                            , MODULENAME & ": ValidateOTCodeTableEntries" _
                                            , , , , _
                                            , !OT_REASON_CD _
                                            , , , 4
                                    End If
                                End If
                        End Select

                    Case trOTRESPONSES
                        Select Case (UCase(Trim(!RECORD_STATUS)))
                            Case "I", "C"
                                If Not (CheckForValidString(!OT_RESP_CD)) Then
                                    AppendOTCDTBLBRV radrsBRV _
                                        , trErrbusOTCodeTables _
                                        , mstrCBUSINESSERR _
                                        , "Invalid Overtime Response." _
                                        , MODULENAME & ": ValidateOTCodeTableEntries" _
                                        , , , , , , , , 5
                                End If

                                If Not (CheckForValidString(!RESPONSE_DESC, True, 30)) Then
                                    AppendOTCDTBLBRV radrsBRV _
                                        , trErrbusOTCodeTables _
                                        , mstrCBUSINESSERR _
                                        , "Invalid Overtime Response Description." _
                                        , MODULENAME & ": ValidateOTCodeTableEntries" _
                                        , , , , , _
                                        , !OT_RESP_CD _
                                        , , 5
                                Else
                                    !RESPONSE_DESC = stdmTRISCommon.ParseString(!RESPONSE_DESC)
                                    .Update
                                End If
                        End Select

                        Select Case (UCase(Trim(!RECORD_STATUS)))
                            Case "C", "D"
                                If Not (CheckForValidString(!KEY_OT_RESP_CD)) Then
                                    AppendOTCDTBLBRV radrsBRV _
                                        , trErrbusOTCodeTables _
                                        , mstrCBUSINESSERR _
                                        , "Invalid Delete/Change request. Overtime Response is invalid." _
                                        , MODULENAME & ": ValidateOTCodeTableEntries" _
                                        , , , , , _
                                        , !OT_RESP_CD _
                                        , , 5
                                Else
                                    If (CheckForValidCodeTableItem(trOTRESPONSEUSED, !KEY_OT_RESP_CD)) Then
                                        AppendOTCDTBLBRV radrsBRV _
                                            , trErrbusOTCodeTables _
                                            , mstrCBUSINESSERR _
                                            , "Invalid Delete/Change request. Overtime reason is currently being used in the application and cannot be changed or deleted." _
                                            , MODULENAME & ": ValidateOTCodeTableEntries" _
                                            , , , , , _
                                            , !OT_RESP_CD _
                                            , , 5
                                    End If
                                End If
                        End Select

                    Case trASSOCIATECRAFTTOJOB
                        If (UCase(Trim(!RECORD_STATUS)) = "C") Then
                            If Not (CheckForValidString(!KEY_JOB_NUM)) Then
                                AppendOTCDTBLBRV radrsBRV _
                                    , trErrbusOTCodeTables _
                                    , mstrCBUSINESSERR _
                                    , "Invalid Change request. Job is invalid." _
                                    , MODULENAME & ": ValidateOTCodeTableEntries" _
                                    , _
                                    , !CRAFT_NUM _
                                    , , , , , , 6
                            End If

                            If (CheckForValidString(!CRAFT_NUM)) Then
                                If Not (CheckForValidCodeTableItem(trCRAFTS, !CRAFT_NUM)) Then
                                    AppendOTCDTBLBRV radrsBRV _
                                        , trErrbusOTCodeTables _
                                        , mstrCBUSINESSERR _
                                        , "Invalid Change request. Craft is invalid." _
                                        , MODULENAME & ": ValidateOTCodeTableEntries" _
                                        , !KEY_JOB_NUM _
                                        , !CRAFT_NUM _
                                        , , , , , , 6
                                End If
                            End If
                        End If

                    Case trOTBANKHEADERS
                        Select Case (UCase(Trim(!RECORD_STATUS)))
                            Case "C", "D"
                                If Not (CheckForValidInteger(!OT_BANK_CD)) Then
                                    AppendOTCDTBLBRV radrsBRV _
                                        , trErrbusOTCodeTables _
                                        , mstrCBUSINESSERR _
                                        , "Invalid Overtime Bank." _
                                        , MODULENAME & ": ValidateOTCodeTableEntries" _
                                        , , , , , , , , 7
                                Else
                                    If Not (CheckForValidCodeTableItem(trOTBANKHEADERS, , , !OT_BANK_CD)) Then
                                        AppendOTCDTBLBRV radrsBRV _
                                            , trErrbusOTCodeTables _
                                            , mstrCBUSINESSERR _
                                            , "Overtime Bank " & CStr(!OT_BANK_CD) & " already does not exists." _
                                            , MODULENAME & ": ValidateOTCodeTableEntries" _
                                            , , , , , , , , 7
                                    End If
                                End If
                        End Select

                        Select Case (UCase(Trim(!RECORD_STATUS)))
                            Case "I", "C"
                                If Not (CheckForValidString(!OT_BANK_DESC, True, 30)) Then
                                    AppendOTCDTBLBRV radrsBRV _
                                        , trErrbusOTCodeTables _
                                        , mstrCBUSINESSERR _
                                        , "Invalid Overtime Bank Description." _
                                        , MODULENAME & ": ValidateOTCodeTableEntries" _
                                        , , , , , , _
                                        , !OT_BANK_CD _
                                        , 7
                                Else
                                    !OT_BANK_DESC = stdmTRISCommon.ParseString(!OT_BANK_DESC)
                                    .Update
                                End If
                        End Select

                        Select Case (UCase(Trim(!RECORD_STATUS)))
                            Case "C", "D"
                                If Not (CheckForValidInteger(!KEY_OT_BANK_CD)) Then
                                    AppendOTCDTBLBRV radrsBRV _
                                        , trErrbusOTCodeTables _
                                        , mstrCBUSINESSERR _
                                        , "Invalid Delete/Change request. Overtime Bank is invalid." _
                                        , MODULENAME & ": ValidateOTCodeTableEntries" _
                                        , , , , , , _
                                        , !KEY_OT_BANK_CD _
                                        , 7
                                Else
                                    If (CheckForValidCodeTableItem(trOTBANKUSED, , , !KEY_OT_BANK_CD)) Then
                                        AppendOTCDTBLBRV radrsBRV _
                                            , trErrbusOTCodeTables _
                                            , mstrCBUSINESSERR _
                                            , "Invalid Delete/Change request. Overtime bank is currently being used in the application and cannot be changed or deleted." _
                                            , MODULENAME & ": ValidateOTCodeTableEntries" _
                                            , , , , , , _
                                            , !KEY_OT_BANK_CD _
                                            , 7
                                    End If
                                End If

                                If Not (CheckForValidString(!CREATED_TIMESTAMP _
                                        , , , True, 26)) Then
                                    AppendOTCDTBLBRV radrsBRV _
                                        , trErrbusOTCodeTables _
                                        , mstrCBUSINESSERR _
                                        , "Invalid Delete/Change request. Time Stamp is invalid." _
                                        , MODULENAME & ": ValidateOTCodeTableEntries" _
                                        , , , , , , _
                                        , !KEY_OT_BANK_CD _
                                        , 7
                                End If
                    End Select
            End Select

            If (viOTCodeTableType <> trOTBANKHEADERS) Then
                Select Case (UCase(Trim(!RECORD_STATUS)))
                    Case "D", "C"
                        If Not (CheckForValidString(!TIME_STAMP _
                                , , , True, 26)) Then
                            AppendOTCDTBLBRV radrsBRV _
                                , trErrbusOTCodeTables _
                                , mstrCBUSINESSERR _
                                , "Invalid Delete/Change request. Time Stamp is invalid." _
                                , MODULENAME & ": ValidateOTCodeTableEntries" _
                                , !KEY_JOB_NUM _
                                , !KEY_CRAFT_NUM _
                                , !KEY_QUAL_CD _
                                , !KEY_SPEC_SKILL_CD _
                                , !KEY_OT_REASON_CD _
                                , !KEY_OT_RESP_CD _
                                , !KEY_OT_BANK_CD _
                                , viOTCodeTableType
                        End If
                End Select
            End If

            .MoveNext
        Next llRecordIndex
    End With

    radrsRSToValidate.Filter = adFilterNone
    radrsRSToValidate.MoveFirst

    ValidateOTCodeTableEntries = True

    If Not (radrsBRV Is Nothing) Then
        With radrsBRV
            If Not (.BOF And .EOF) Then
                .Filter = adFilterNone
                .MoveFirst
                lsFilterString = "ERR_TYPE = 'BE' OR ERR_TYPE = 'AE' OR ERR_TYPE = 'AW'"
                .Filter = stdmTRISCommon.MDAC25FilterFix(lsFilterString)
                lsFilterString = ""
                If Not (.EOF) Then
                    ValidateOTCodeTableEntries = False
                End If

                .Filter = adFilterNone
            End If
        End With
    End If

    Exit Function
ValidateOTCodeTableEntries_EH:
    ValidateOTCodeTableEntries = False

    TrsRaiseError MODULENAME _
            , "ValidateOTCodeTableEntries"

End Function
Private Sub AppendOTCDTBLBRV(ByRef radrsBRV As ADODB.Recordset _
                , ByVal vlErrorNumber As Long _
                , ByVal vsErrorType As String _
                , ByVal vsErrorDescription As String _
                , ByVal vsErrorSource As String _
                , Optional ByVal vsJobNumber As String = "" _
                , Optional ByVal vsCraftNumber As String = "" _
                , Optional ByVal vsQualificationCode As String = "" _
                , Optional ByVal vsSpecialSkillCode As String = "" _
                , Optional ByVal vsOTReasonCode As String = "" _
                , Optional ByVal vsOTResponseCode As String = "" _
                , Optional ByVal vlOTBankCode As Integer = 0 _
                , Optional ByVal viRecordSource As Integer = 0 _
                )
On Error GoTo AppendOTCDTBLBRV_EH

    With radrsBRV
        If (.State <> adStateOpen) Then
            Exit Sub
        End If

        .AddNew
        !JOB_NUM = vsJobNumber
        !CRAFT_NUM = vsCraftNumber
        !QUAL_CD = vsQualificationCode
        !SPEC_SKILL_CD = vsSpecialSkillCode
        !OT_REASON_CD = vsOTReasonCode
        !OT_RESP_CD = vsOTResponseCode
        !OT_BANK_CD = vlOTBankCode

        !ERR_NUMBER = vlErrorNumber - vbObjectError
        !ERR_TYPE = Trim(vsErrorType)
        !REC_SOURCE = viRecordSource
        Select Case (viRecordSource)
                Case 1
                    !ERR_DESCRIPTION = "Craft: " & Trim(vsErrorDescription)
                Case 2
                    !ERR_DESCRIPTION = "Qualification: " & Trim(vsErrorDescription)
                Case 3
                    !ERR_DESCRIPTION = "Special Skills: " & Trim(vsErrorDescription)
                Case 4
                    !ERR_DESCRIPTION = "Overtime Reason: " & Trim(vsErrorDescription)
                Case 5
                    !ERR_DESCRIPTION = "Overtime Response: " & Trim(vsErrorDescription)
                Case 6
                    !ERR_DESCRIPTION = "Associate Craft to Job: " & Trim(vsErrorDescription)
                Case 7
                    !ERR_DESCRIPTION = "Overtime Bank Definition: " & Trim(vsErrorDescription)
                Case Else
                    !ERR_DESCRIPTION = Trim(vsErrorDescription)
        End Select
        !ERR_SOURCE = Trim(vsErrorSource)

        .Update
    End With

    Exit Sub
AppendOTCDTBLBRV_EH:
    TrsRaiseError MODULENAME _
        , "AppendOTCDTBLBRV"

End Sub
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'This method will save Promotional Lines entries
Public Function ProcessPromotionalLinesEntries( _
                    ByVal vadrsPromotionalLines As ADODB.Recordset _
                    , ByRef radrsErrors As ADODB.Recordset _
                    , ByVal vlUserEmployeeNumber As Long _
                    ) As Boolean
On Error GoTo ProcessPromotionalLinesEntries_EH
Dim lbEmptyRequest As Boolean, lsFilterString As String
Dim lobjdbCodeTables As dbCodeTables.cdCodeTablesTrans
Dim llIndex As Long

    'Check for request existence
    lbEmptyRequest = True
    If Not (vadrsPromotionalLines Is Nothing) Then
        If Not (vadrsPromotionalLines.BOF And _
            vadrsPromotionalLines.EOF) Then
            lbEmptyRequest = False
        End If
    End If

    'Prepare BRV Recordset
    Set radrsErrors = CreateBRV("PROMLINES")

    If (lbEmptyRequest) Then
        AppendPROMLINESBRV radrsErrors _
            , trErrbusPromLines _
            , mstrCBUSINESSERR _
            , "There are no Promotional Lines entries to process." _
            , MODULENAME & ": ProcessPromotionalLinesEntries"
        GoTo REPORTBRV
    Else
        With vadrsPromotionalLines
            .Filter = adFilterNone

            If Not (.EOF) Then
                .MoveFirst
                'Validate
                If Not (ValidatePromotionalLinesEntries( _
                    vadrsPromotionalLines _
                    , radrsErrors _
                    , vlUserEmployeeNumber)) Then
                    GoTo REPORTBRV
                End If

                .Filter = adFilterNone
                .MoveFirst
            End If
        End With
    End If

    Set lobjdbCodeTables = CreateObject( _
                                "dbCodeTables.cdCodeTablesTrans")
    'Set lobjdbCodeTables = New dbCodeTables.cdCodeTablesTrans
    If Not (lobjdbCodeTables.dbProcessPromotionalLines( _
            vadrsPromotionalLines _
            , vlUserEmployeeNumber)) Then
        Set lobjdbCodeTables = Nothing

        Err.Raise trErrbusPromLines
    End If

    'Destroy
    Set lobjdbCodeTables = Nothing
    Set vadrsPromotionalLines = Nothing
    Set radrsErrors = Nothing

    'Return
    ProcessPromotionalLinesEntries = True

    'Complete
    GetObjectContext().SetComplete

    Exit Function
ProcessPromotionalLinesEntries_EH:
    'Abort Transaction
    GetObjectContext().SetAbort

    Set vadrsPromotionalLines = Nothing
    Set radrsErrors = Nothing
    Set lobjdbCodeTables = Nothing

    ProcessPromotionalLinesEntries = False

    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "ProcessPromotionalLinesEntries"
            Case trErrdbInsIntoPromLines, trErrdbUpdPromLines _
                , trErrdbDelFromPromLines, trErrdbSavePromLines _
                , trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "ProcessPromotionalLinesEntries" _
                    , Err.Number _
                    , , , , _
                    True
            Case Else
                TrsRaiseError MODULENAME _
                    , "ProcessPromotionalLinesEntries" _
                    , trErrbusPromLines _
                    , , , , _
                    True
    End Select

    Exit Function

REPORTBRV:
    Set lobjdbCodeTables = Nothing
    Set vadrsPromotionalLines = Nothing

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
                    ProcessPromotionalLinesEntries = False
                Else
                    'Complete Transaction
                    GetObjectContext().SetComplete
                    ProcessPromotionalLinesEntries = True
                End If

                .Filter = adFilterNone
            Else
                'Complete Transaction
                GetObjectContext().SetComplete
                ProcessPromotionalLinesEntries = True
            End If
        End With
    Else
        'Complete Transaction
        GetObjectContext().SetComplete
        ProcessPromotionalLinesEntries = True
    End If

    Exit Function

End Function
Private Function ValidatePromotionalLinesEntries( _
                    ByRef radrsRSToValidate As ADODB.Recordset _
                    , ByRef radrsBRV As ADODB.Recordset _
                    , ByVal vlUserEmployeeNumber As Long _
                    ) As Boolean
On Error GoTo ValidatePromotionalLinesEntries_EH

Dim llRecordIndex As Long, lsFilterString As String

    ValidatePromotionalLinesEntries = False

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

    If Not (CheckForValidLong(vlUserEmployeeNumber)) Then
        AppendPROMLINESBRV radrsBRV _
            , trErrbusPromLines _
            , mstrCBUSINESSERR _
            , "Invalid TRIS User." _
            , MODULENAME & ": ValidatePromotionalLinesEntries"
    End If

    With radrsRSToValidate
        For llRecordIndex = 1 To .RecordCount
            Select Case (UCase(Trim(!RECORD_STATUS)))
                Case "I", "C"
                    If Not (stdmTRISCommon.CheckForValidString(!DEPT_NUM)) Then
                        AppendPROMLINESBRV radrsBRV _
                            , trErrbusPromLines _
                            , mstrCBUSINESSERR _
                            , "Invalid Department." _
                            , MODULENAME & ": ValidatePromotionalLinesEntries" _
                            , _
                            , !HR_LOC _
                            , !JOB_NUM _
                            , !IR_AREA_CD _
                            , !IR_DIV_CD _
                            , !LOP_CURR_LEV_IND _
                            , !LOP_NEXT_LEV_IND
                    Else
                        If Not (CheckForValidCodeTableItem(trDEPARTMENT, !DEPT_NUM)) Then
                            AppendPROMLINESBRV radrsBRV _
                                , trErrbusPromLines _
                                , mstrCBUSINESSERR _
                                , "Invalid Department. Department Number " & Trim(!DEPT_NUM) & " does not exist." _
                                , MODULENAME & ": ValidatePromotionalLinesEntries" _
                                , !DEPT_NUM _
                                , !HR_LOC _
                                , !JOB_NUM _
                                , !IR_AREA_CD _
                                , !IR_DIV_CD _
                                , !LOP_CURR_LEV_IND _
                                , !LOP_NEXT_LEV_IND
                        End If
                    End If

                    If Not (stdmTRISCommon.CheckForValidString(!JOB_NUM)) Then
                        AppendPROMLINESBRV radrsBRV _
                            , trErrbusPromLines _
                            , mstrCBUSINESSERR _
                            , "Invalid Job." _
                            , MODULENAME & ": ValidatePromotionalLinesEntries" _
                            , !DEPT_NUM _
                            , !HR_LOC _
                            , _
                            , !IR_AREA_CD _
                            , !IR_DIV_CD _
                            , !LOP_CURR_LEV_IND _
                            , !LOP_NEXT_LEV_IND
                    Else
                        If Not (CheckForValidCodeTableItem(trJOB, !JOB_NUM)) Then
                            AppendPROMLINESBRV radrsBRV _
                                , trErrbusPromLines _
                                , mstrCBUSINESSERR _
                                , "Invalid Job. Job Number " & Trim(!JOB_NUM) & " does not exist." _
                                , MODULENAME & ": ValidatePromotionalLinesEntries" _
                                , !DEPT_NUM _
                                , !HR_LOC _
                                , !JOB_NUM _
                                , !IR_AREA_CD _
                                , !IR_DIV_CD _
                                , !LOP_CURR_LEV_IND _
                                , !LOP_NEXT_LEV_IND
                        End If
                    End If

                    If Not (stdmTRISCommon.CheckForValidString(!IR_AREA_CD)) Then
                        AppendPROMLINESBRV radrsBRV _
                            , trErrbusPromLines _
                            , mstrCBUSINESSERR _
                            , "Invalid Industrial Relations Area Code." _
                            , MODULENAME & ": ValidatePromotionalLinesEntries" _
                            , !DEPT_NUM _
                            , !HR_LOC _
                            , !JOB_NUM _
                            , _
                            , !IR_DIV_CD _
                            , !LOP_CURR_LEV_IND _
                            , !LOP_NEXT_LEV_IND
                    End If

                    If Not (stdmTRISCommon.CheckForValidString(!IR_DIV_CD)) Then
                        AppendPROMLINESBRV radrsBRV _
                            , trErrbusPromLines _
                            , mstrCBUSINESSERR _
                            , "Invalid Industrial Relations Division Code." _
                            , MODULENAME & ": ValidatePromotionalLinesEntries" _
                            , !DEPT_NUM _
                            , !HR_LOC _
                            , !JOB_NUM _
                            , !IR_AREA_CD _
                            , _
                            , !LOP_CURR_LEV_IND _
                            , !LOP_NEXT_LEV_IND
                    End If

                    If Not (stdmTRISCommon.CheckForValidString(!LOP_CURR_LEV_IND)) Then
                        AppendPROMLINESBRV radrsBRV _
                            , trErrbusPromLines _
                            , mstrCBUSINESSERR _
                            , "Invalid Current Line of Promotion Level Indicator." _
                            , MODULENAME & ": ValidatePromotionalLinesEntries" _
                            , !DEPT_NUM _
                            , !HR_LOC _
                            , !JOB_NUM _
                            , !IR_AREA_CD _
                            , !IR_DIV_CD _
                            , _
                            , !LOP_NEXT_LEV_IND
                    End If

                    If Not (stdmTRISCommon.CheckForValidString(!LOP_NEXT_LEV_IND)) Then
                        AppendPROMLINESBRV radrsBRV _
                            , trErrbusPromLines _
                            , mstrCBUSINESSERR _
                            , "Invalid Next Line of Promotion Level Indicator." _
                            , MODULENAME & ": ValidatePromotionalLinesEntries" _
                            , !DEPT_NUM _
                            , !HR_LOC _
                            , !JOB_NUM _
                            , !IR_AREA_CD _
                            , !IR_DIV_CD _
                            , !LOP_CURR_LEV_IND
                    End If
            End Select

            Select Case (UCase(Trim(!RECORD_STATUS)))
                Case "C", "D"
                    If Not (stdmTRISCommon.CheckForValidString(!KEY_DEPT_NUM)) Then
                        AppendPROMLINESBRV radrsBRV _
                            , trErrbusPromLines _
                            , mstrCBUSINESSERR _
                            , "Invalid Delete/Change request. Department Number is invalid." _
                            , MODULENAME & ": ValidatePromotionalLinesEntries" _
                            , !DEPT_NUM _
                            , !HR_LOC _
                            , !JOB_NUM _
                            , !IR_AREA_CD _
                            , !IR_DIV_CD _
                            , !LOP_CURR_LEV_IND _
                            , !LOP_NEXT_LEV_IND
                    End If

                    If Not (stdmTRISCommon.CheckForValidString(!KEY_JOB_NUM)) Then
                        AppendPROMLINESBRV radrsBRV _
                            , trErrbusPromLines _
                            , mstrCBUSINESSERR _
                            , "Invalid Delete/Change request. Job Number is invalid." _
                            , MODULENAME & ": ValidatePromotionalLinesEntries" _
                            , !DEPT_NUM _
                            , !HR_LOC _
                            , !JOB_NUM _
                            , !IR_AREA_CD _
                            , !IR_DIV_CD _
                            , !LOP_CURR_LEV_IND _
                            , !LOP_NEXT_LEV_IND
                    End If

                    If Not (stdmTRISCommon.CheckForValidString(!KEY_IR_AREA_CD)) Then
                        AppendPROMLINESBRV radrsBRV _
                            , trErrbusPromLines _
                            , mstrCBUSINESSERR _
                            , "Invalid Delete/Change request. Industrial Relations Area Code is invalid." _
                            , MODULENAME & ": ValidatePromotionalLinesEntries" _
                            , !DEPT_NUM _
                            , !HR_LOC _
                            , !JOB_NUM _
                            , !IR_AREA_CD _
                            , !IR_DIV_CD _
                            , !LOP_CURR_LEV_IND _
                            , !LOP_NEXT_LEV_IND
                    End If

                    If Not (stdmTRISCommon.CheckForValidString(!KEY_IR_DIV_CD)) Then
                        AppendPROMLINESBRV radrsBRV _
                            , trErrbusPromLines _
                            , mstrCBUSINESSERR _
                            , "Invalid Delete/Change request. Industrial Relations Division Code is invalid." _
                            , MODULENAME & ": ValidatePromotionalLinesEntries" _
                            , !DEPT_NUM _
                            , !HR_LOC _
                            , !JOB_NUM _
                            , !IR_AREA_CD _
                            , !IR_DIV_CD _
                            , !LOP_CURR_LEV_IND _
                            , !LOP_NEXT_LEV_IND
                    End If

                    If Not (stdmTRISCommon.CheckForValidString(!KEY_LOP_CURR_LEV_IND)) Then
                        AppendPROMLINESBRV radrsBRV _
                            , trErrbusPromLines _
                            , mstrCBUSINESSERR _
                            , "Invalid Delete/Change request. Current Line of Promotion Level Indicator is invalid." _
                            , MODULENAME & ": ValidatePromotionalLinesEntries" _
                            , !DEPT_NUM _
                            , !HR_LOC _
                            , !JOB_NUM _
                            , !IR_AREA_CD _
                            , !IR_DIV_CD _
                            , !LOP_CURR_LEV_IND _
                            , !LOP_NEXT_LEV_IND
                    End If

                    If Not (stdmTRISCommon.CheckForValidString(!TIME_STAMP _
                            , , , True, 26)) Then
                        AppendPROMLINESBRV radrsBRV _
                            , trErrbusPromLines _
                            , mstrCBUSINESSERR _
                            , "Invalid Delete/Change request. Time Stamp is invalid." _
                            , MODULENAME & ": ValidatePromotionalLinesEntries" _
                            , !DEPT_NUM _
                            , !HR_LOC _
                            , !JOB_NUM _
                            , !IR_AREA_CD _
                            , !IR_DIV_CD _
                            , !LOP_CURR_LEV_IND _
                            , !LOP_NEXT_LEV_IND
                    End If
            End Select

            .MoveNext
        Next llRecordIndex
    End With

    radrsRSToValidate.Filter = adFilterNone
    radrsRSToValidate.MoveFirst

    ValidatePromotionalLinesEntries = True

    If Not (radrsBRV Is Nothing) Then
        With radrsBRV
            If Not (.BOF And .EOF) Then
                .Filter = adFilterNone
                .MoveFirst
                lsFilterString = "ERR_TYPE = 'BE' OR ERR_TYPE = 'AE' OR ERR_TYPE = 'AW'"
                .Filter = stdmTRISCommon.MDAC25FilterFix(lsFilterString)
                lsFilterString = ""
                If Not (.EOF) Then
                    ValidatePromotionalLinesEntries = False
                End If

                .Filter = adFilterNone
            End If
        End With
    End If

    Exit Function
ValidatePromotionalLinesEntries_EH:
    ValidatePromotionalLinesEntries = False

    TrsRaiseError MODULENAME _
            , "ValidatePromotionalLinesEntries"

End Function
Private Sub AppendPROMLINESBRV(ByRef radrsBRV As ADODB.Recordset _
                , ByVal vlErrorNumber As Long _
                , ByVal vsErrorType As String _
                , ByVal vsErrorDescription As String _
                , ByVal vsErrorSource As String _
                , Optional ByVal vsDepartmentNumber As String = "" _
                , Optional ByVal vsHRLocation As String = "" _
                , Optional ByVal vsJobNumber As String = "" _
                , Optional ByVal vsIRAreaCode As String = "" _
                , Optional ByVal vsIRDivisionCode As String = "" _
                , Optional ByVal vsCurrentLOP As String = "" _
                , Optional ByVal vsNextLOP As String = "" _
                , Optional ByVal viRecordSource As Integer = 0 _
                )
On Error GoTo AppendPROMLINESBRV_EH

    With radrsBRV
        If (.State <> adStateOpen) Then
            Exit Sub
        End If

        .AddNew

        !DEPT_NUM = vsDepartmentNumber
        !HR_LOC = vsHRLocation
        !JOB_NUM = vsJobNumber
        !IR_AREA_CD = vsIRAreaCode
        !IR_DIV_CD = vsIRDivisionCode
        !LOP_CURR_LEV_IND = vsCurrentLOP
        !LOP_NEXT_LEV_IND = vsNextLOP

        !ERR_NUMBER = vlErrorNumber - vbObjectError
        !ERR_TYPE = Trim(vsErrorType)
        !REC_SOURCE = viRecordSource
        !ERR_DESCRIPTION = Trim(vsErrorDescription)
        !ERR_SOURCE = Trim(vsErrorSource)

        .Update
    End With

    Exit Sub
AppendPROMLINESBRV_EH:
    TrsRaiseError MODULENAME _
        , "AppendPROMLINESBRV"

End Sub
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Private Function CheckForValidCodeTableItem(ByVal viCodeType As Integer _
                    , Optional ByVal vsStringItem As String = "" _
                    , Optional ByVal vlLongItem As Long = 0 _
                    , Optional ByVal viIntegerItem As Integer = 0 _
                    , Optional ByVal vsAssociatedItem As String = "" _
                    ) As Boolean
On Error GoTo CheckForValidCodeTableItem_EH
Dim lobjdbCodeTables As dbCodeTables.cdCodeTables

    Set lobjdbCodeTables = CreateObject( _
                    "dbCodeTables.cdCodeTables")

    Select Case (viCodeType)
        Case trCRAFTS
            CheckForValidCodeTableItem = _
                lobjdbCodeTables.IsAValidCraft(Trim(vsStringItem))

        Case trQUALIFICATIONS
            CheckForValidCodeTableItem = _
                lobjdbCodeTables.IsAValidQualificationCode(Trim(vsStringItem))

        Case trOTBANKHEADERS
            CheckForValidCodeTableItem = _
                lobjdbCodeTables.DoesBankExist(viIntegerItem)

        Case trDEPARTMENT
            CheckForValidCodeTableItem = _
                lobjdbCodeTables.IsAValidDepartment(Trim(vsStringItem))

        Case trJOB
            CheckForValidCodeTableItem = _
                lobjdbCodeTables.JobNumberIsValid(Trim(vsStringItem))

        Case trQUALUSED
            CheckForValidCodeTableItem = _
                lobjdbCodeTables.IsAnAssignedQualification(Trim(vsStringItem))

        Case trSKILLUSED
            CheckForValidCodeTableItem = _
                lobjdbCodeTables.IsAnAssignedSkill(Trim(vsStringItem) _
                    , Trim(vsAssociatedItem))

        Case trCRAFTUSED
            CheckForValidCodeTableItem = _
                lobjdbCodeTables.IsCraftUsedinTRIS(Trim(vsStringItem) _
                    , Trim(vsAssociatedItem))

        Case trOTREASONUSED
            CheckForValidCodeTableItem = _
                lobjdbCodeTables.IsOTReasonUsedinTRIS(Trim(vsStringItem))

        Case trOTRESPONSEUSED
            CheckForValidCodeTableItem = _
                lobjdbCodeTables.IsOTResponseUsedinTRIS(Trim(vsStringItem))

        Case trOTBANKUSED
            CheckForValidCodeTableItem = _
                lobjdbCodeTables.IsOTBankUsedinTRIS(viIntegerItem)
    End Select

    Set lobjdbCodeTables = Nothing

    Exit Function
CheckForValidCodeTableItem_EH:
    GetObjectContext().SetAbort
    Set lobjdbCodeTables = Nothing

    TrsRaiseError MODULENAME _
        , "CheckForValidCodeTableItem"

End Function
