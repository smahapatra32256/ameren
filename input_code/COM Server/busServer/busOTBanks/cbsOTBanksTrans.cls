VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "cbsOTBanksTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'AMEREN Project: TRIS Conversion Phase 3
'May 2004
'Author: Prem M. Chander
'Module Description: This class will provide the user the ability
'to create/modify Employee Overtime Banks

'COMMENTS: OCT, 2005 - CHANGE GETOBJECTCONTEXT.CREATEINSTANCE
'TO CREATEOBJECT -- PREM

Option Explicit

Private Const ADOFIELDATTR = adFldUpdatable Or adFldIsNullable
Private Const MODULENAME As String = "busOTBanks.cbsOTBanksTrans"
Private Const trEMPLOYEE As Integer = 1
Private Const trOTBANKCODE As Integer = 2
Public Function ProcessEmployeeOvertimeBanks( _
                    ByVal vadrsEmployeeOTBanks As ADODB.Recordset _
                    , ByRef radrsErrors As ADODB.Recordset _
                    , ByVal vlUserEmployeeNumber As Long _
                    ) As Boolean
On Error GoTo ProcessEmployeeOvertimeBanks_EH
Dim lbEmptyRequest As Boolean, llRecordIndex As Long
Dim lobjdbOTBanks As dbOTBanks.cdOTBanks
Dim ladrsOTBankAuditInformation As ADODB.Recordset

    'Check for request existence
    lbEmptyRequest = True
    If (RecordsExist(vadrsEmployeeOTBanks)) Then
        lbEmptyRequest = False
    End If

    'Prepare BRV Recordset
    Set radrsErrors = CreateOTBanksBRV

    If (lbEmptyRequest) Then
        AppendOTBanksBRV radrsErrors _
            , trErrbusEmployeeOTBanks _
            , mstrCBUSINESSERR _
            , "There are no Employee Overtime Bank entries to process." _
            , MODULENAME & ": ProcessEmployeeOvertimeBanks"
        GoTo REPORTBRV
    Else
        'Create Audit Recordset
        Set ladrsOTBankAuditInformation = CreateOTBanksAuditRecordset

        With vadrsEmployeeOTBanks
            .Filter = adFilterNone
            If Not (.EOF) Then
                .MoveFirst
                If Not (ValidateEmployeeOTBanksEntries( _
                    vadrsEmployeeOTBanks _
                    , ladrsOTBankAuditInformation _
                    , radrsErrors _
                    , vlUserEmployeeNumber)) Then
                    GoTo REPORTBRV
                End If

                .Filter = adFilterNone
                .MoveFirst
            End If
        End With
    End If

    Set lobjdbOTBanks = CreateObject( _
                            "dbOTBanks.cdOTBanks")
    If Not (lobjdbOTBanks.dbProcessEmployeeOvertimeBanks( _
            vadrsEmployeeOTBanks _
            , ladrsOTBankAuditInformation _
            , vlUserEmployeeNumber)) Then
        Set lobjdbOTBanks = Nothing

        Err.Raise trErrbusEmployeeOTBanks
    End If

    Set lobjdbOTBanks = Nothing
    Set ladrsOTBankAuditInformation = Nothing

    GetObjectContext().SetComplete

    ProcessEmployeeOvertimeBanks = True

    Set lobjdbOTBanks = Nothing

    Exit Function
ProcessEmployeeOvertimeBanks_EH:
    GetObjectContext().SetAbort

    Set lobjdbOTBanks = Nothing
    Set ladrsOTBankAuditInformation = Nothing

    Select Case (Err.Number)
        Case trErrutOpenConnection
            TrsRaiseError MODULENAME _
                , "ProcessEmployeeOvertimeBanks"
        Case trErrdbDelFromOTBanks, trErrdbInsIntoOTBanks _
            , trErrdbUpdOTBanks, trErrdbSaveOTBanks _
            , trErrdbInsIntoOTBanksAudit, trErrdb_ExecuteSQL_F
            TrsRaiseError MODULENAME _
                , "ProcessEmployeeOvertimeBanks" _
                , Err.Number _
                , , , , _
                True
        Case Else
            TrsRaiseError MODULENAME _
                , "ProcessEmployeeOvertimeBanks" _
                , trErrbusEmployeeOTBanks _
                , , , , _
                True
    End Select

    Exit Function

REPORTBRV:
    Set lobjdbOTBanks = Nothing
    Set ladrsOTBankAuditInformation = Nothing

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
                    ProcessEmployeeOvertimeBanks = False
                Else
                    'Complete Transaction
                    GetObjectContext().SetComplete
                    ProcessEmployeeOvertimeBanks = True
                End If

                .Filter = adFilterNone
            Else
                'Complete Transaction
                GetObjectContext().SetComplete
                ProcessEmployeeOvertimeBanks = True
            End If
        End With
    Else
        'Complete Transaction
        GetObjectContext().SetComplete
        ProcessEmployeeOvertimeBanks = True
    End If

    Exit Function

End Function
Private Function CreateOTBanksBRV() As ADODB.Recordset
On Error GoTo CreateOTBanksBRV_EH

Dim ladrs As New ADODB.Recordset

    With ladrs
        .Fields.Append "EMP_NUM", adInteger, , ADOFIELDATTR
        .Fields.Append "OT_BANK_CD", adInteger, , ADOFIELDATTR
        .Fields.Append "TOT_OT_HRS", adDecimal, , ADOFIELDATTR
        .Fields.Item("TOT_OT_HRS").Precision = 7
        .Fields.Item("TOT_OT_HRS").NumericScale = 2
        .Fields.Append "TOT_OT_HRS_REF", adDecimal, , ADOFIELDATTR
        .Fields.Item("TOT_OT_HRS_REF").Precision = 7
        .Fields.Item("TOT_OT_HRS_REF").NumericScale = 2

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

    Set CreateOTBanksBRV = ladrs
    Set ladrs = Nothing
    Exit Function

CreateOTBanksBRV_EH:
    Set ladrs = Nothing
    Set CreateOTBanksBRV = Nothing

    TrsRaiseError MODULENAME _
        , "CreateOTBanksBRV"

End Function
Private Function ValidateEmployeeOTBanksEntries( _
                    ByRef radrsRSToValidate As ADODB.Recordset _
                    , ByRef radrsOTBankAuditInformation As ADODB.Recordset _
                    , ByRef radrsBRV As ADODB.Recordset _
                    , ByVal vlUserEmployeeNumber As Long _
                    ) As Boolean
On Error GoTo ValidateEmployeeOTBanksEntries_EH
Dim lobjdbOTBanks As dbOTBanks.cdOTBanks
Dim ladrsOriginalOTBanks As ADODB.Recordset
Dim ladrs As ADODB.Recordset
Dim llRecordIndex As Long, lbValidProject As Boolean
Dim ldblTotalOTHours As Double, ldblRefusedHours As Double
Dim lsWhereClause As String, ldblAdjustedHours As Double

    ValidateEmployeeOTBanksEntries = False

    If Not (RecordsExist(radrsRSToValidate)) Then
        Exit Function
    End If

    'Get Original Banks Information
    lsWhereClause = PrepareEmployeeWhereClauseForSelect(radrsRSToValidate)
    If (Trim(lsWhereClause) <> "") Then
        Set lobjdbOTBanks = CreateObject( _
                            "dbOTBanks.cdOTBanks")
        Set ladrsOriginalOTBanks = _
            lobjdbOTBanks.GetEmployeeOvertimeBanks2(lsWhereClause)
        Set lobjdbOTBanks = Nothing
    End If

    If Not (stdmTRISCommon.CheckForValidLong(vlUserEmployeeNumber, True)) Then
        AppendOTBanksBRV radrsBRV _
            , trErrbusEmployeeOTBanks _
            , mstrCBUSINESSERR _
            , "Invalid TRIS User." _
            , MODULENAME & ": ValidateEmployeeOTBanksEntries"
    End If

    With radrsRSToValidate
        .MoveFirst

        For llRecordIndex = 1 To .RecordCount
            Select Case (UCase(Trim(!RECORD_STATUS)))
                Case "I", "C"
                    If Not (stdmTRISCommon.CheckForValidLong(!EMP_NUM, True)) Then
                        AppendOTBanksBRV radrsBRV _
                            , trErrbusEmployeeOTBanks _
                            , mstrCBUSINESSERR _
                            , "Invalid Employee." _
                            , MODULENAME & ": ValidateEmployeeOTBanksEntries" _
                            , _
                            , !OT_BANK_CD _
                            , !TOT_OT_HRS _
                            , !TOT_OT_HRS_REF
                    Else
                        If Not (CheckForItemExistenceInDB(trEMPLOYEE, , !EMP_NUM)) Then
                            AppendOTBanksBRV radrsBRV _
                                , trErrbusEmployeeOTBanks _
                                , mstrCBUSINESSERR _
                                , "Invalid Employee. Employee " & Trim(!EMP_NUM) & " does not exist." _
                                , MODULENAME & ": ValidateEmployeeOTBanksEntries" _
                                , !EMP_NUM _
                                , !OT_BANK_CD _
                                , !TOT_OT_HRS _
                                , !TOT_OT_HRS_REF
                        End If
                    End If

                    If Not (stdmTRISCommon.CheckForValidInteger(!OT_BANK_CD)) Then
                        AppendOTBanksBRV radrsBRV _
                            , trErrbusEmployeeOTBanks _
                            , mstrCBUSINESSERR _
                            , "Invalid Bank." _
                            , MODULENAME & ": ValidateEmployeeOTBanksEntries" _
                            , !EMP_NUM _
                            , !OT_BANK_CD _
                            , !TOT_OT_HRS _
                            , !TOT_OT_HRS_REF
                    Else
                        If Not (CheckForItemExistenceInDB(trOTBANKCODE _
                            , , , !OT_BANK_CD)) Then
                            AppendOTBanksBRV radrsBRV _
                                , trErrbusEmployeeOTBanks _
                                , mstrCBUSINESSERR _
                                , "Invalid Bank. Overtime bank " & CStr(!OT_BANK_CD) & " does not exist." _
                                , MODULENAME & ": ValidateEmployeeOTBanksEntries" _
                                , !EMP_NUM _
                                , !OT_BANK_CD _
                                , !TOT_OT_HRS _
                                , !TOT_OT_HRS_REF
                        End If
                    End If

                    If (!TOT_OT_HRS < 0) Then
                        !TOT_OT_HRS = Abs(!TOT_OT_HRS)
                        .Update
                    End If
                    If (!TOT_OT_HRS_REF < 0) Then
                        !TOT_OT_HRS_REF = Abs(!TOT_OT_HRS_REF)
                        .Update
                    End If

                    'Prepare Audit Information
                    If (!RECORD_STATUS = "I") Then
                        If Not (radrsOTBankAuditInformation Is Nothing) Then
                            'Accepted
                            AddOTBanksAuditRecord radrsOTBankAuditInformation _
                                , !EMP_NUM _
                                , Trim(!REASON) _
                                , (!TOT_OT_HRS - !TOT_OT_HRS_REF) _
                                , !OT_BANK_CD _
                                , "A"
                            'Refused
                            AddOTBanksAuditRecord radrsOTBankAuditInformation _
                                , !EMP_NUM _
                                , Trim(!REASON) _
                                , !TOT_OT_HRS_REF _
                                , !OT_BANK_CD _
                                , "R"
                        End If
                    End If
            End Select

            Select Case (UCase(Trim(!RECORD_STATUS)))
                Case "C", "D"
                    ldblTotalOTHours = 0#
                    ldblRefusedHours = 0#

                    Select Case (UCase(Trim(!RECORD_STATUS)))
                        Case "C"
                            ldblTotalOTHours = !TOT_OT_HRS
                            ldblRefusedHours = !TOT_OT_HRS_REF
                        Case "D"
                            ldblTotalOTHours = !KEY_TOT_OT_HRS
                            ldblRefusedHours = !KEY_TOT_OT_HRS_REF
                    End Select

                    If Not (stdmTRISCommon.CheckForValidString(!REASON, True, 30)) Then
                        AppendOTBanksBRV radrsBRV _
                            , trErrbusEmployeeOTBanks _
                            , mstrCBUSINESSERR _
                            , "Invalid Delete/Change request. Invalid reason for overtime banks adjustment." _
                            , MODULENAME & ": ValidateEmployeeOTBanksEntries" _
                            , !KEY_EMP_NUM _
                            , !KEY_OT_BANK_CD _
                            , ldblTotalOTHours _
                            , ldblRefusedHours
                    Else
                        !REASON = stdmTRISCommon.ParseString(!REASON)
                        .Update
                    End If

                    If Not (stdmTRISCommon.CheckForValidLong(!KEY_EMP_NUM, True)) Then
                        AppendOTBanksBRV radrsBRV _
                            , trErrbusEmployeeOTBanks _
                            , mstrCBUSINESSERR _
                            , "Invalid Delete/Change request. Invalid Employee." _
                            , MODULENAME & ": ValidateEmployeeOTBanksEntries" _
                            , !KEY_EMP_NUM _
                            , !KEY_OT_BANK_CD _
                            , ldblTotalOTHours _
                            , ldblRefusedHours
                    End If

                    If Not (stdmTRISCommon.CheckForValidInteger(!KEY_OT_BANK_CD, True)) Then
                        AppendOTBanksBRV radrsBRV _
                            , trErrbusEmployeeOTBanks _
                            , mstrCBUSINESSERR _
                            , "Invalid Delete/Change request. Invalid bank." _
                            , MODULENAME & ": ValidateEmployeeOTBanksEntries" _
                            , !KEY_EMP_NUM _
                            , _
                            , ldblTotalOTHours _
                            , ldblRefusedHours
                    End If

                    If Not (stdmTRISCommon.CheckForValidString(!TIME_STAMP _
                            , , , True, 26)) Then
                        AppendOTBanksBRV radrsBRV _
                            , trErrbusEmployeeOTBanks _
                            , mstrCBUSINESSERR _
                            , "Invalid Delete/Change request. Invalid Time Stamp." _
                            , MODULENAME & ": ValidateEmployeeOTBanksEntries" _
                            , !KEY_EMP_NUM _
                            , !KEY_OT_BANK_CD _
                            , ldblTotalOTHours _
                            , ldblRefusedHours
                    End If

                    'Prepare Audit Information
                    If Not (radrsOTBankAuditInformation Is Nothing) Then
                        Select Case (UCase(Trim(!RECORD_STATUS)))
                            Case "C"
                                If (RecordsExist(ladrsOriginalOTBanks)) Then
                                    ladrsOriginalOTBanks.Filter = adFilterNone
                                    If Not (ladrsOriginalOTBanks.EOF) Then
                                        ladrsOriginalOTBanks.MoveFirst

                                        ladrsOriginalOTBanks.Filter = MDAC25FilterFix("EMP_NUM = " & CStr(!KEY_EMP_NUM) & _
                                                                " AND OT_BANK_CD = " & CStr(!KEY_OT_BANK_CD))
                                        If Not (ladrsOriginalOTBanks.EOF) Then
                                            ladrsOriginalOTBanks.MoveFirst

                                            'Accepted
                                            ldblAdjustedHours = DetermineAdjustedHours( _
                                                (ladrsOriginalOTBanks!TOT_OT_HRS - ladrsOriginalOTBanks!TOT_OT_HRS_REF) _
                                                , (!TOT_OT_HRS - !TOT_OT_HRS_REF))
                                            If ((!TOT_OT_HRS - !TOT_OT_HRS_REF) > _
                                                (ladrsOriginalOTBanks!TOT_OT_HRS - ladrsOriginalOTBanks!TOT_OT_HRS_REF)) Then
                                                ldblAdjustedHours = Abs(ldblAdjustedHours)
                                            Else
                                                If ((!TOT_OT_HRS - !TOT_OT_HRS_REF) < _
                                                    (ladrsOriginalOTBanks!TOT_OT_HRS - ladrsOriginalOTBanks!TOT_OT_HRS_REF)) Then
                                                    ldblAdjustedHours = ldblAdjustedHours * -1
                                                End If
                                            End If
                                            If (ldblAdjustedHours <> 0) Then
                                                AddOTBanksAuditRecord radrsOTBankAuditInformation _
                                                    , !EMP_NUM _
                                                    , Trim(!REASON) _
                                                    , ldblAdjustedHours _
                                                    , !OT_BANK_CD _
                                                    , "A"
                                            End If

                                            'Refused
                                            ldblAdjustedHours = DetermineAdjustedHours(ladrsOriginalOTBanks!TOT_OT_HRS_REF _
                                                , !TOT_OT_HRS_REF)
                                            If (!TOT_OT_HRS_REF > ladrsOriginalOTBanks!TOT_OT_HRS_REF) Then
                                                ldblAdjustedHours = Abs(ldblAdjustedHours)
                                            Else
                                                If (!TOT_OT_HRS_REF < ladrsOriginalOTBanks!TOT_OT_HRS_REF) Then
                                                    ldblAdjustedHours = ldblAdjustedHours * -1
                                                End If
                                            End If
                                            If (ldblAdjustedHours <> 0) Then
                                                AddOTBanksAuditRecord radrsOTBankAuditInformation _
                                                    , !EMP_NUM _
                                                    , Trim(!REASON) _
                                                    , ldblAdjustedHours _
                                                    , !OT_BANK_CD _
                                                    , "R"
                                            End If
                                        End If
                                    End If
                                    ladrsOriginalOTBanks.Filter = adFilterNone
                                End If

                           Case "D"
                                If (RecordsExist(ladrsOriginalOTBanks)) Then
                                    ladrsOriginalOTBanks.Filter = adFilterNone
                                    If Not (ladrsOriginalOTBanks.EOF) Then
                                        ladrsOriginalOTBanks.MoveFirst

                                        ladrsOriginalOTBanks.Filter = MDAC25FilterFix("EMP_NUM = " & CStr(!KEY_EMP_NUM) & _
                                                " AND OT_BANK_CD = " & CStr(!KEY_OT_BANK_CD))
                                        If Not (ladrsOriginalOTBanks.EOF) Then
                                            ladrsOriginalOTBanks.MoveFirst

                                            'Accepted
                                            AddOTBanksAuditRecord radrsOTBankAuditInformation _
                                                , !EMP_NUM _
                                                , Trim(!REASON) _
                                                , (ladrsOriginalOTBanks!TOT_OT_HRS - ladrsOriginalOTBanks!TOT_OT_HRS_REF) _
                                                , !OT_BANK_CD _
                                                , "A"

                                            'Refused
                                            AddOTBanksAuditRecord radrsOTBankAuditInformation _
                                                , !EMP_NUM _
                                                , Trim(!REASON) _
                                                , ladrsOriginalOTBanks!TOT_OT_HRS_REF _
                                                , !OT_BANK_CD _
                                                , "R"
                                        End If
                                    End If
                                    ladrsOriginalOTBanks.Filter = adFilterNone
                                End If
                        End Select
                    End If
            End Select

            .MoveNext
        Next llRecordIndex
    End With

    Set ladrsOriginalOTBanks = Nothing
    Set ladrs = Nothing

    radrsRSToValidate.Filter = adFilterNone
    radrsRSToValidate.MoveFirst

    ValidateEmployeeOTBanksEntries = True

    If Not (radrsBRV Is Nothing) Then
        With radrsBRV
            If Not (.BOF And .EOF) Then
                .Filter = adFilterNone
                .Filter = MDAC25FilterFix("ERR_TYPE = 'BE' OR ERR_TYPE = 'AE' OR ERR_TYPE = 'AW'")
                If Not (.EOF) Then
                    ValidateEmployeeOTBanksEntries = False
                End If

                .Filter = adFilterNone
            End If
        End With
    End If

    Exit Function
ValidateEmployeeOTBanksEntries_EH:
    Set lobjdbOTBanks = Nothing

    Set ladrsOriginalOTBanks = Nothing
    Set ladrs = Nothing

    TrsRaiseError MODULENAME _
            , "ValidateEmployeeOTBanksEntries"

End Function
Private Sub AppendOTBanksBRV(ByRef radrsBRV As ADODB.Recordset _
                , ByVal vlErrorNumber As Long _
                , ByVal vsErrorType As String _
                , ByVal vsErrorDescription As String _
                , ByVal vsErrorSource As String _
                , Optional ByVal vlEmployeeNumber As Long = 0 _
                , Optional ByVal viBankCode As Integer = 0 _
                , Optional ByVal vdblTotalHours As Double = 0# _
                , Optional ByVal vdblRefusedHours As Double = 0# _
                , Optional ByVal viRecordSource As Integer = 0)
On Error GoTo AppendOTBanksBRV_EH

    With radrsBRV
        If (.State <> adStateOpen) Then
            Exit Sub
        End If

        .AddNew

        !EMP_NUM = vlEmployeeNumber
        !OT_BANK_CD = viBankCode
        !TOT_OT_HRS = vdblTotalHours
        !TOT_OT_HRS_REF = vdblRefusedHours

        !ERR_NUMBER = vlErrorNumber - vbObjectError
        !ERR_TYPE = Trim(vsErrorType)
        !REC_SOURCE = viRecordSource
        !ERR_DESCRIPTION = Trim(vsErrorDescription)
        !ERR_SOURCE = Trim(vsErrorSource)

        .Update
    End With

    Exit Sub
AppendOTBanksBRV_EH:
    TrsRaiseError MODULENAME _
        , "AppendOTBanksBRV"

End Sub
Private Function PrepareEmployeeWhereClauseForSelect(ByVal vadrsEmployees As ADODB.Recordset _
                ) As String
On Error GoTo PrepareEmployeeWhereClauseForSelect_EH
Dim lsWhereClause As String, llPreviousEmployee As Long
Dim llIndex As Long

    'Check for existence and prepare the where clause
    If (RecordsExist(vadrsEmployees)) Then
        With vadrsEmployees
            .MoveFirst

            lsWhereClause = ""
            llPreviousEmployee = 0

            For llIndex = 1 To .RecordCount
                If (!EMP_NUM <> llPreviousEmployee) Then
                    lsWhereClause = lsWhereClause & _
                                ", " & CStr(!EMP_NUM)
                End If
                llPreviousEmployee = !EMP_NUM

                .MoveNext
            Next llIndex

            .MoveFirst
        End With
    End If

    If (Trim(lsWhereClause) <> "") Then
        If (Left(lsWhereClause, 1) = ",") Then
            'Replace with a blank
            lsWhereClause = Replace(lsWhereClause, "," _
                                , "", 1, 1)
        End If
    End If

    If (Trim(lsWhereClause) <> "") Then
        PrepareEmployeeWhereClauseForSelect = lsWhereClause
    Else
        PrepareEmployeeWhereClauseForSelect = ""
    End If

    Exit Function
PrepareEmployeeWhereClauseForSelect_EH:
    TrsRaiseError MODULENAME _
        , "PrepareEmployeeWhereClauseForSelect"

End Function
'Hours Adjusted is the difference between original Accepted/Refused Hours
'and modified Accepted/Refused Hours
Private Function CreateOTBanksAuditRecordset() As ADODB.Recordset
On Error GoTo CreateOTBanksAuditRecordset_EH
Dim ladrs As ADODB.Recordset

    Set ladrs = New ADODB.Recordset
    With ladrs
      .Fields.Append "EMP_NUM", adInteger, , ADOFIELDATTR
      .Fields.Append "REASON", adChar, 30, ADOFIELDATTR
      .Fields.Append "HRS_ADJ", adDecimal, , ADOFIELDATTR
      .Fields.Item("HRS_ADJ").Precision = 7
      .Fields.Item("HRS_ADJ").NumericScale = 2
      .Fields.Append "OT_BANK_CD", adInteger, , ADOFIELDATTR
      .Fields.Append "ACC_REF_IND", adChar, 1, ADOFIELDATTR
      .Fields.Append "OT_OPP_NUM", adInteger, , ADOFIELDATTR

      .CursorLocation = adUseClient
      .CursorType = adOpenStatic
      .LockType = adLockBatchOptimistic
      .Open
    End With

    Set CreateOTBanksAuditRecordset = ladrs

    Set ladrs = Nothing

    Exit Function
CreateOTBanksAuditRecordset_EH:
    Set ladrs = Nothing
    TrsRaiseError MODULENAME _
        , "CreateOTBanksAuditRecordset"
End Function
Private Sub AddOTBanksAuditRecord(ByRef radrsAudit As ADODB.Recordset _
                , ByVal vlEmployeeNumber As Long _
                , ByVal vsReason As String _
                , ByVal vdblAdjustedHours As Double _
                , ByVal viBankCode As Integer _
                , ByVal vsAcceptedRefusedIndicator As String _
                , Optional ByVal vlOTOpportunityNumber As Long = 0)
On Error GoTo AddOTBanksAuditRecord_EH

    If Not (radrsAudit Is Nothing) Then
        With radrsAudit
            .AddNew
            !EMP_NUM = vlEmployeeNumber
            !REASON = Trim(vsReason)
            !HRS_ADJ = vdblAdjustedHours
            !OT_BANK_CD = viBankCode
            !ACC_REF_IND = Trim(vsAcceptedRefusedIndicator)
            !OT_OPP_NUM = vlOTOpportunityNumber
            .Update
        End With
    End If

    Exit Sub
AddOTBanksAuditRecord_EH:
    TrsRaiseError MODULENAME _
        , "AddOTBanksAuditRecord"

End Sub
'Note this is done only for Accepted and Refused, not Total Hours.
'Do not call this method during Insert New/Delete Bank, call it
'during Update only.
Private Function DetermineAdjustedHours(ByVal vdblOriginalHours As Double _
                    , ByVal vdblNewHours As Double _
                    ) As Double
On Error GoTo DetermineAdjustedHours_EH

    DetermineAdjustedHours = vdblOriginalHours - vdblNewHours

    Exit Function
DetermineAdjustedHours_EH:
    TrsRaiseError MODULENAME _
        , "DetermineAdjustedHours"

End Function
Private Function CheckForItemExistenceInDB(ByVal viItemType As Integer _
                    , Optional ByVal vsStringItem As String = "" _
                    , Optional ByVal vlLongItem As Long = 0 _
                    , Optional ByVal viIntegerItem As Integer = 0 _
                    ) As Boolean
On Error GoTo CheckForItemExistenceInDB_EH
Dim lobjdbCodeTables As dbCodeTables.cdCodeTables
Dim lobjdbEmployee As dbEmployee.cdEmployee

    Select Case (viItemType)
        Case trEMPLOYEE
            Set lobjdbEmployee = CreateObject( _
                                    "dbEmployee.cdEmployee")
            CheckForItemExistenceInDB = _
                lobjdbEmployee.IsValidEmployee(vlLongItem)
            Set lobjdbEmployee = Nothing

        Case trOTBANKCODE
            Set lobjdbCodeTables = CreateObject( _
                                    "dbCodeTables.cdCodeTables")
            CheckForItemExistenceInDB = _
                lobjdbCodeTables.DoesBankExist(viIntegerItem)
            Set lobjdbCodeTables = Nothing
    End Select

    Set lobjdbEmployee = Nothing
    Set lobjdbCodeTables = Nothing

    Exit Function
CheckForItemExistenceInDB_EH:
    Set lobjdbEmployee = Nothing
    Set lobjdbCodeTables = Nothing

    TrsRaiseError MODULENAME _
        , "CheckForItemExistenceInDB"

End Function
'END PREM TRIS CONVERSION PHASE 3
