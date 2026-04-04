VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "cbsRequestFirstApprTran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Project:   TRIS Conversion Phase 1
'Author:    Prem M. Chander
'Date:      November 2001 (Firstcut)
'Description: This class component will reside in a Server Package
'under Microsoft Transaction Server (MTS). This component will be called by
'the client to processes save requests. It is transactional
'This module will call its associated data access component
'to retrieve/modify pertinant information from the database.

Option Explicit
'Constants
Private Const MODULENAME As String = "busAdjustments.cbsRequestFirstApprTran"
Private Const ADOFIELDATTR = adFldUpdatable Or adFldIsNullable
Private Const EMPLOYEE As Integer = 1
Private Const DEPARTMENT As Integer = 2
'ADJUSTMENT REQUEST SAVE
Public Function ProcessAdjustmentRequest( _
                ByVal v_adrsAdjRequest As ADODB.Recordset _
                , ByRef r_adrsErrors As ADODB.Recordset _
                , ByVal v_lUserEmployeeNumber As Long _
                ) As Boolean
On Error GoTo ProcessAdjustmentRequest_EH
Dim lobjDataAccess As dbAdjustments.cdRequestFirstAppr
'MDAC25 FILTER FIX
Dim lsFilterString As String

    ProcessAdjustmentRequest = False

    'Create/Define BRV
    Set r_adrsErrors = DefineBRV

    'Check for Recordset Existence
    If Not (RecordsExist(v_adrsAdjRequest)) Then
        AppendBRV r_adrsErrors _
            , 0 _
            , Date _
            , "" _
            , v_lUserEmployeeNumber _
            , trErrAdjustEntry_F _
            , mstrCBUSINESSERR _
            , "No Adjustment entries to process" _
            , MODULENAME & ": ProcessAdjustmentRequest"

        GoTo REPORTBRV
    End If

    With v_adrsAdjRequest
        .Filter = adFilterNone

        'Validate Request for any processing (DIC)
        If Not (.EOF) Then
            .MoveFirst

            If Not (ValidateReqFirstAppEntries(v_adrsAdjRequest _
                , r_adrsErrors _
                , v_lUserEmployeeNumber _
                , "REQ")) Then
                GoTo REPORTBRV
            End If

            'MDAC25 FILTER FIX
            lsFilterString = "RECORD_STATUS = 'D'"
            .Filter = stdmTRISCommon.MDAC25FilterFix(lsFilterString)
            lsFilterString = ""
            If Not (.EOF) Then
                .MoveFirst

                'Construct the Data Access Object
                Set lobjDataAccess = CreateObject( _
                            "dbAdjustments.cdRequestFirstAppr")
                If Not (lobjDataAccess.db_DeleteReqFirstApp( _
                        v_adrsAdjRequest _
                        , v_lUserEmployeeNumber)) Then
                    Set lobjDataAccess = Nothing

                    Err.Raise trErrdbDelAdj
                End If

                Set lobjDataAccess = Nothing
            End If

            .Filter = adFilterNone
            'MDAC25 FILTER FIX
            lsFilterString = "RECORD_STATUS = 'I'"
            .Filter = stdmTRISCommon.MDAC25FilterFix(lsFilterString)
            lsFilterString = ""
            If Not (.EOF) Then
                .MoveFirst

                'Construct the Data Access Object
                Set lobjDataAccess = CreateObject( _
                            "dbAdjustments.cdRequestFirstAppr")
                If Not (lobjDataAccess.db_InsertRequest( _
                        v_adrsAdjRequest _
                        , v_lUserEmployeeNumber)) Then
                    Set lobjDataAccess = Nothing

                    Err.Raise trErrdbInsReqAdj
                End If

                Set lobjDataAccess = Nothing
            End If

            .Filter = adFilterNone
            'MDAC25 FILTER FIX
            lsFilterString = "RECORD_STATUS = 'C'"
            .Filter = stdmTRISCommon.MDAC25FilterFix(lsFilterString)
            lsFilterString = ""
            If Not (.EOF) Then
                .MoveFirst

                'Construct the Data Access Object
                Set lobjDataAccess = CreateObject( _
                            "dbAdjustments.cdRequestFirstAppr")
                If Not (lobjDataAccess.db_UpdateRequest( _
                        v_adrsAdjRequest _
                        , v_lUserEmployeeNumber)) Then
                    Set lobjDataAccess = Nothing

                    Err.Raise trErrdbUpdReqAdj
                End If

                Set lobjDataAccess = Nothing
            End If
        End If

        .Filter = adFilterNone
    End With

    Set v_adrsAdjRequest = Nothing
    Set lobjDataAccess = Nothing

    GetObjectContext().SetComplete

    ProcessAdjustmentRequest = True
    Exit Function
ProcessAdjustmentRequest_EH:
    'Abort Transaction
    GetObjectContext().SetAbort

    Set v_adrsAdjRequest = Nothing
    Set r_adrsErrors = Nothing
    Set lobjDataAccess = Nothing

    Select Case Err.Number
        Case trErrutOpenConnection
            TrsRaiseError MODULENAME _
                , "ProcessAdjustmentRequest"

        Case trErrdbInsReqAdj, trErrdbUpdReqAdj _
            , trErrdbUpdFirstAppAdj, trErrdbDelAdj _
            , trErrdb_ExecuteSQL_F
            TrsRaiseError MODULENAME _
                , "ProcessAdjustmentRequest" _
                , Err.Number _
                , , , , _
                True

        Case Else
            TrsRaiseError MODULENAME _
                , "ProcessAdjustmentRequest" _
                , trErrProcessRequestAdjust_F _
                , , , , _
                True
    End Select

    Exit Function
REPORTBRV:
    Set lobjDataAccess = Nothing

    If (RecordsExist(r_adrsErrors)) Then
        With r_adrsErrors
            .Filter = adFilterNone

            'Check for errors
            'MDAC25 FILTER FIX
            lsFilterString = "ERR_TYPE = 'BE'"
            .Filter = stdmTRISCommon.MDAC25FilterFix(lsFilterString)
            lsFilterString = ""

            'If found, abort else complete
            If Not (.EOF) Then
                'Abort Transaction
                GetObjectContext().SetAbort
                ProcessAdjustmentRequest = False
            Else
                'Complete Transaction
                GetObjectContext().SetComplete
                ProcessAdjustmentRequest = True
            End If

            .Filter = adFilterNone
        End With
    Else
        'Complete Transaction
        GetObjectContext().SetComplete
        ProcessAdjustmentRequest = True
    End If

    Set v_adrsAdjRequest = Nothing
    Exit Function

End Function
'ADJUSTMENT FIRST APPROVAL
Public Function ProcessAdjustmentFirstApproval( _
                ByVal v_adrsAdjRequest As ADODB.Recordset _
                , ByRef r_adrsErrors As ADODB.Recordset _
                , ByVal v_lUserEmployeeNumber As Long _
                ) As Boolean
On Error GoTo ProcessAdjustmentFirstApproval_EH
Dim lobjDataAccess As dbAdjustments.cdRequestFirstAppr
Dim lsFilterString As String, llRecordIndex As Long

    ProcessAdjustmentFirstApproval = False

    Set r_adrsErrors = DefineBRV

    'Check for Recordset Existence
    If Not (RecordsExist(v_adrsAdjRequest)) Then
        AppendBRV r_adrsErrors _
            , 0 _
            , Date _
            , "" _
            , v_lUserEmployeeNumber _
            , trErrAdjustEntry_F _
            , mstrCBUSINESSERR _
            , "No Adjustment entries to process" _
            , MODULENAME & ": ProcessAdjustmentRequest"
        GoTo REPORTBRV
    End If

    With v_adrsAdjRequest
        .Filter = adFilterNone

        'Validate Request for any processing (DIC)
        If Not (.EOF) Then
            .MoveFirst

            If Not (ValidateReqFirstAppEntries(v_adrsAdjRequest _
                    , r_adrsErrors _
                    , v_lUserEmployeeNumber _
                    , "FIRAPP")) Then
                GoTo REPORTBRV
            End If

            lsFilterString = "RECORD_STATUS = 'D'"
            .Filter = stdmTRISCommon.MDAC25FilterFix(lsFilterString)
            lsFilterString = ""
            If Not (.EOF) Then
                .MoveFirst

                'Construct the Data Access Object
                Set lobjDataAccess = CreateObject( _
                            "dbAdjustments.cdRequestFirstAppr")
                If Not (lobjDataAccess.db_DeleteReqFirstApp( _
                            v_adrsAdjRequest _
                            , v_lUserEmployeeNumber)) Then
                    Set lobjDataAccess = Nothing

                    Err.Raise trErrdbDelAdj
                End If

                Set lobjDataAccess = Nothing
            End If

            .Filter = adFilterNone
            lsFilterString = "RECORD_STATUS = 'C'"
            .Filter = stdmTRISCommon.MDAC25FilterFix(lsFilterString)
            lsFilterString = ""
            If Not (.EOF) Then
                .MoveFirst

                'Construct the Data Access Object
                Set lobjDataAccess = CreateObject( _
                            "dbAdjustments.cdRequestFirstAppr")
                If Not (lobjDataAccess.db_UpdateWithFirstApproval( _
                        v_adrsAdjRequest _
                        , v_lUserEmployeeNumber)) Then
                    Set lobjDataAccess = Nothing

                    Err.Raise trErrdbUpdFirstAppAdj
                End If

                Set lobjDataAccess = Nothing
            End If
        End If

        .Filter = adFilterNone
    End With

    Set v_adrsAdjRequest = Nothing
    Set lobjDataAccess = Nothing

    GetObjectContext().SetComplete

    ProcessAdjustmentFirstApproval = True

    Exit Function
ProcessAdjustmentFirstApproval_EH:
    'Abort Transaction
    GetObjectContext().SetAbort

    Set v_adrsAdjRequest = Nothing
    Set r_adrsErrors = Nothing
    Set lobjDataAccess = Nothing

    Select Case Err.Number
        Case trErrutOpenConnection
            TrsRaiseError MODULENAME _
                , "ProcessAdjustmentFirstApproval"

        Case trErrdbInsReqAdj, trErrdbUpdReqAdj _
            , trErrdbUpdFirstAppAdj, trErrdbDelAdj _
            , trErrdb_ExecuteSQL_F
            TrsRaiseError MODULENAME _
                , "ProcessAdjustmentFirstApproval" _
                , Err.Number, , , , True

        Case Else
            TrsRaiseError MODULENAME _
                , "ProcessAdjustmentFirstApproval" _
                , trErrProcessFirstAppAdjust_F _
                , , , , True
    End Select

    Exit Function
REPORTBRV:
    Set lobjDataAccess = Nothing

    If (RecordsExist(r_adrsErrors)) Then
        With r_adrsErrors
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
                ProcessAdjustmentFirstApproval = False
            Else
                'Complete Transaction
                GetObjectContext().SetComplete
                ProcessAdjustmentFirstApproval = True
            End If
        End With
    Else
        'Complete Transaction
        GetObjectContext().SetComplete
        ProcessAdjustmentFirstApproval = True
    End If

    Set v_adrsAdjRequest = Nothing

    Exit Function

End Function
Private Function ValidateReqFirstAppEntries( _
                ByRef radrsAdjEntries As ADODB.Recordset _
                , ByRef radrsBRV As ADODB.Recordset _
                , ByVal vlUserEmployeeNumber As Long _
                , ByVal vsCallerID As String _
                ) As Boolean
On Error GoTo ValidateReqFirstAppEntries_EH
Dim lsOrigReason As String, lvArrOfEmps As Variant
Dim ladrsEmployee As ADODB.Recordset
Dim lsFilterString As String, llRecordIndex As Long
Dim lsParsedReason As String

    With radrsAdjEntries
        .Filter = adFilterNone

        lsFilterString = "RECORD_STATUS = 'I' OR RECORD_STATUS = 'C'"
        .Filter = stdmTRISCommon.MDAC25FilterFix(lsFilterString)
        lsFilterString = ""

        If Not (.EOF) Then
            .MoveFirst

            ReDim lvArrOfEmps(.RecordCount - 1)
            For llRecordIndex = 1 To .RecordCount
                lvArrOfEmps(llRecordIndex - 1) = !EMP_NUM

                .MoveNext
            Next llRecordIndex
        End If

        If Not (IsEmpty(lvArrOfEmps)) Then
            Set ladrsEmployee = GetEmployeeInfo(lvArrOfEmps)
        End If
    End With

    With radrsAdjEntries
        .Filter = adFilterNone

        .MoveFirst
        For llRecordIndex = 1 To .RecordCount
            lsOrigReason = ""

            Select Case (!RECORD_STATUS)
                Case "D"    'DELETE
                    If Not (CheckForValidLong(!EMP_NUM, True)) Then
                        AppendBRV radrsBRV _
                            , 0 _
                            , !CREATED_DATE _
                            , !DEPT_NUM _
                            , vlUserEmployeeNumber _
                            , trErrAdjustEntry_F _
                            , mstrCBUSINESSERR _
                            , "Invalid employee." _
                            , MODULENAME & ": ValidateReqFirstAppEntries"
                    Else
                        If Not (CheckForItemExistenceInDB(EMPLOYEE, , !EMP_NUM)) Then
                            AppendBRV radrsBRV _
                                , !EMP_NUM _
                                , !CREATED_DATE _
                                , !DEPT_NUM _
                                , vlUserEmployeeNumber _
                                , trErrAdjustEntry_F _
                                , mstrCBUSINESSERR _
                                , "Invalid employee. Employee " & CStr(!EMP_NUM) & _
                                    " does not exist." _
                                , MODULENAME & ": ValidateReqFirstAppEntries"
                        End If
                    End If

                    If Not (CheckForValidString(!DEPT_NUM, True, 3)) Then
                        AppendBRV radrsBRV _
                            , !EMP_NUM _
                            , !CREATED_DATE _
                            , "" _
                            , vlUserEmployeeNumber _
                            , trErrAdjustEntry_F _
                            , mstrCBUSINESSERR _
                            , "Invalid department." _
                            , MODULENAME & ": ValidateReqFirstAppEntries"
                    Else
                        If Not (CheckForItemExistenceInDB(DEPARTMENT, !DEPT_NUM)) Then
                            AppendBRV radrsBRV _
                                , !EMP_NUM _
                                , !CREATED_DATE _
                                , !DEPT_NUM _
                                , vlUserEmployeeNumber _
                                , trErrAdjustEntry_F _
                                , mstrCBUSINESSERR _
                                , "Invalid department. Department " & !DEPT_NUM & _
                                    " does not exist." _
                                , MODULENAME & ": ValidateReqFirstAppEntries"
                        End If
                    End If

                    If Not (CheckForValidDate(!CREATED_DATE)) Then
                        AppendBRV radrsBRV _
                            , !EMP_NUM _
                            , "" _
                            , !DEPT_NUM _
                            , vlUserEmployeeNumber _
                            , trErrAdjustEntry_F _
                            , mstrCBUSINESSERR _
                            , "Invalid request date." _
                            , MODULENAME & ": ValidateReqFirstAppEntries"
                    End If

                    If Not (CheckForValidLong(!CREATED_EMP_NUM, True)) Then
                        AppendBRV radrsBRV _
                            , !EMP_NUM _
                            , !CREATED_DATE _
                            , !DEPT_NUM _
                            , vlUserEmployeeNumber _
                            , trErrAdjustEntry_F _
                            , mstrCBUSINESSERR _
                            , "Invalid request by employee." _
                            , MODULENAME & ": ValidateReqFirstAppEntries"
                    Else
                        If Not (CheckForItemExistenceInDB(EMPLOYEE, , !CREATED_EMP_NUM)) Then
                            AppendBRV radrsBRV _
                                , !EMP_NUM _
                                , !CREATED_DATE _
                                , !DEPT_NUM _
                                , vlUserEmployeeNumber _
                                , trErrAdjustEntry_F _
                                , mstrCBUSINESSERR _
                                , "Invalid request by employee. Employee " & CStr(!CREATED_EMP_NUM) & _
                                    " does not exist." _
                                , MODULENAME & ": ValidateReqFirstAppEntries"
                        End If
                    End If

                    If Not (CheckForValidString(!CREATED_TIMESTAMP, , , True, 26)) Then
                        AppendBRV radrsBRV _
                            , !EMP_NUM _
                            , !CREATED_DATE _
                            , !DEPT_NUM _
                            , vlUserEmployeeNumber _
                            , trErrAdjustEntry_F _
                            , mstrCBUSINESSERR _
                            , "Invalid time stamp." _
                            , MODULENAME & ": ValidateReqFirstAppEntries"
                    End If

                Case "I"    'INSERT
                    If Not (CheckForValidLong(!EMP_NUM, True)) Then
                        AppendBRV radrsBRV _
                            , 0 _
                            , !CREATED_DATE _
                            , !DEPT_NUM _
                            , vlUserEmployeeNumber _
                            , trErrAdjustEntry_F _
                            , mstrCBUSINESSERR _
                            , "Invalid employee." _
                            , MODULENAME & ": ValidateReqFirstAppEntries"
                    Else
                        If Not (CheckForItemExistenceInDB(EMPLOYEE, , !EMP_NUM)) Then
                            AppendBRV radrsBRV _
                                , !EMP_NUM _
                                , !CREATED_DATE _
                                , !DEPT_NUM _
                                , vlUserEmployeeNumber _
                                , trErrAdjustEntry_F _
                                , mstrCBUSINESSERR _
                                , "Invalid employee. Employee " & CStr(!EMP_NUM) & _
                                    " does not exist." _
                                , MODULENAME & ": ValidateReqFirstAppEntries"
                        End If
                    End If

                    If Not (CheckForValidString(!DEPT_NUM, True, 3)) Then
                        AppendBRV radrsBRV _
                            , !EMP_NUM _
                            , !CREATED_DATE _
                            , "" _
                            , vlUserEmployeeNumber _
                            , trErrAdjustEntry_F _
                            , mstrCBUSINESSERR _
                            , "Invalid department." _
                            , MODULENAME & ": ValidateReqFirstAppEntries"
                    Else
                        If Not (CheckForItemExistenceInDB(DEPARTMENT, !DEPT_NUM)) Then
                            AppendBRV radrsBRV _
                                , !EMP_NUM _
                                , !CREATED_DATE _
                                , !DEPT_NUM _
                                , vlUserEmployeeNumber _
                                , trErrAdjustEntry_F _
                                , mstrCBUSINESSERR _
                                , "Invalid department. Department " & !DEPT_NUM & _
                                    " does not exist." _
                                , MODULENAME & ": ValidateReqFirstAppEntries"
                        End If
                    End If

                    If Not (CheckForValidLong(vlUserEmployeeNumber, True)) Then
                        AppendBRV radrsBRV _
                            , !EMP_NUM _
                            , !CREATED_DATE _
                            , !DEPT_NUM _
                            , vlUserEmployeeNumber _
                            , trErrAdjustEntry_F _
                            , mstrCBUSINESSERR _
                            , "Invalid TRIS User." _
                            , MODULENAME & ": ValidateReqFirstAppEntries"
                    Else
                        If Not (CheckForItemExistenceInDB(EMPLOYEE, , vlUserEmployeeNumber)) Then
                            AppendBRV radrsBRV _
                                , !EMP_NUM _
                                , !CREATED_DATE _
                                , !DEPT_NUM _
                                , vlUserEmployeeNumber _
                                , trErrAdjustEntry_F _
                                , mstrCBUSINESSERR _
                                , "Invalid TRIS User. User " & CStr(vlUserEmployeeNumber) & _
                                    " does not exist." _
                                , MODULENAME & ": ValidateReqFirstAppEntries"
                        End If
                    End If

                    If (!EMP_NUM > 0) Then
                        If (RecordsExist(ladrsEmployee)) Then
                            ladrsEmployee.MoveFirst
                            ladrsEmployee.Find "EMP_NUM = " & !EMP_NUM

                            If Not (ladrsEmployee.EOF) Then
                                If (Trim(ladrsEmployee!EXEC_CONTRACT_IND) = "E") Then
                                    If (!SPECIAL_CHECK_IND = "Y") Then
                                        !SPECIAL_CHECK_IND = "N"
                                        .Update

                                        AppendBRV radrsBRV _
                                            , !EMP_NUM _
                                            , !CREATED_DATE _
                                            , !DEPT_NUM _
                                            , vlUserEmployeeNumber _
                                            , trErrAdjustEntry_F _
                                            , mstrCBUSINESSWAR _
                                            , "Special check is not applicable to executive employees." & _
                                                " Special check has been removed for this request." _
                                            , MODULENAME
                                    End If
                                End If
                            Else
                                Select Case (vsCallerID)
                                    Case "REQ"
                                        Err.Raise trErrProcessRequestAdjust_F
                                    Case "FIRST"
                                        Err.Raise trErrProcessFirstAppAdjust_F
                                End Select
                            End If
                        End If
                    End If

                    If Not (CheckForValidString(!SPECIAL_CHECK_IND, True, 1)) Then
                        !SPECIAL_CHECK_IND = "N"
                        .Update
                    End If

                    'Check for the Length of the Reason
                    If (CheckForValidString(!REQ_REASON)) Then
                        lsOrigReason = Trim(!REQ_REASON)
                        lsParsedReason = Trim(ParseString(Trim(!REQ_REASON)))

                        'Get the Length of the Reason
                        If (Len(lsParsedReason) > 255) Then
                            'Send a warning to the user stating that anything
                            'above 255 will be truncated.
                            !REQ_REASON = Left(lsParsedReason, 255)
                            .Update

                            AppendBRV radrsBRV _
                                , !EMP_NUM _
                                , !CREATED_DATE _
                                , !DEPT_NUM _
                                , vlUserEmployeeNumber _
                                , trErrAdjustEntry_F _
                                , mstrCBUSINESSWAR _
                                , "Adjustment request reason is more than 255 characters long." & _
                                    " Adjustment request reason was truncated to 255 characters for this request." _
                                , MODULENAME
                        Else
                            !REQ_REASON = lsParsedReason
                            .Update
                        End If

                        If (Trim(lsOrigReason) <> Trim(!REQ_REASON)) Then
                            AppendBRV radrsBRV _
                                , !EMP_NUM _
                                , !CREATED_DATE _
                                , !DEPT_NUM _
                                , vlUserEmployeeNumber _
                                , trErrAdjustEntry_F _
                                , mstrCBUSINESSWAR _
                                , "Special characters (Tab, New Line, Semi Colon " & _
                                    ", Tilda, etc.) were found and removed from request reason." _
                                , MODULENAME
                        End If
                    Else
                        AppendBRV radrsBRV _
                            , !EMP_NUM _
                            , !CREATED_DATE _
                            , !DEPT_NUM _
                            , vlUserEmployeeNumber _
                            , trErrAdjustEntry_F _
                            , mstrCBUSINESSERR _
                            , "Invalid adjustment request reason." _
                            , MODULENAME & ": ValidateReqFirstAppEntries"
                    End If

                Case "C"    'CHANGE/UPDATE
                    If Not (CheckForValidLong(!EMP_NUM, True)) Then
                        AppendBRV radrsBRV _
                            , 0 _
                            , !CREATED_DATE _
                            , !DEPT_NUM _
                            , vlUserEmployeeNumber _
                            , trErrAdjustEntry_F _
                            , mstrCBUSINESSERR _
                            , "Invalid employee." _
                            , MODULENAME & ": ValidateReqFirstAppEntries"
                    Else
                        If Not (CheckForItemExistenceInDB(EMPLOYEE, , !EMP_NUM)) Then
                            AppendBRV radrsBRV _
                                , !EMP_NUM _
                                , !CREATED_DATE _
                                , !DEPT_NUM _
                                , vlUserEmployeeNumber _
                                , trErrAdjustEntry_F _
                                , mstrCBUSINESSERR _
                                , "Invalid employee. Employee " & CStr(!EMP_NUM) & _
                                    " does not exist." _
                                , MODULENAME & ": ValidateReqFirstAppEntries"
                        End If
                    End If

                    If Not (CheckForValidString(!DEPT_NUM, True, 3)) Then
                        AppendBRV radrsBRV _
                            , !EMP_NUM _
                            , !CREATED_DATE _
                            , "" _
                            , vlUserEmployeeNumber _
                            , trErrAdjustEntry_F _
                            , mstrCBUSINESSERR _
                            , "Invalid department." _
                            , MODULENAME & ": ValidateReqFirstAppEntries"
                    Else
                        If Not (CheckForItemExistenceInDB(DEPARTMENT, !DEPT_NUM)) Then
                            AppendBRV radrsBRV _
                                , !EMP_NUM _
                                , !CREATED_DATE _
                                , !DEPT_NUM _
                                , vlUserEmployeeNumber _
                                , trErrAdjustEntry_F _
                                , mstrCBUSINESSERR _
                                , "Invalid department. Department " & !DEPT_NUM & _
                                    " does not exist." _
                                , MODULENAME & ": ValidateReqFirstAppEntries"
                        End If
                    End If

                    If Not (CheckForValidDate(!CREATED_DATE)) Then
                        AppendBRV radrsBRV _
                            , !EMP_NUM _
                            , "" _
                            , !DEPT_NUM _
                            , vlUserEmployeeNumber _
                            , trErrAdjustEntry_F _
                            , mstrCBUSINESSERR _
                            , "Invalid request date" _
                            , MODULENAME & ": ValidateReqFirstAppEntries"
                    End If

                    If Not (CheckForValidLong(!CREATED_EMP_NUM, True)) Then
                        AppendBRV radrsBRV _
                            , 0 _
                            , !CREATED_DATE _
                            , !DEPT_NUM _
                            , vlUserEmployeeNumber _
                            , trErrAdjustEntry_F _
                            , mstrCBUSINESSERR _
                            , "Invalid request by employee." _
                            , MODULENAME & ": ValidateReqFirstAppEntries"
                    Else
                        If Not (CheckForItemExistenceInDB(EMPLOYEE, , !CREATED_EMP_NUM)) Then
                            AppendBRV radrsBRV _
                                , !EMP_NUM _
                                , !CREATED_DATE _
                                , !DEPT_NUM _
                                , vlUserEmployeeNumber _
                                , trErrAdjustEntry_F _
                                , mstrCBUSINESSERR _
                                , "Invalid request by employee. Employee " & CStr(!CREATED_EMP_NUM) & _
                                    " does not exist." _
                                , MODULENAME & ": ValidateReqFirstAppEntries"
                        End If
                    End If

                    If Not (CheckForValidLong(vlUserEmployeeNumber, True)) Then
                        AppendBRV radrsBRV _
                            , !EMP_NUM _
                            , !CREATED_DATE _
                            , !DEPT_NUM _
                            , vlUserEmployeeNumber _
                            , trErrAdjustEntry_F _
                            , mstrCBUSINESSERR _
                            , "Invalid TRIS User." _
                            , MODULENAME & ": ValidateReqFirstAppEntries"
                    Else
                        If Not (CheckForItemExistenceInDB(EMPLOYEE, , vlUserEmployeeNumber)) Then
                            AppendBRV radrsBRV _
                                , !EMP_NUM _
                                , !CREATED_DATE _
                                , !DEPT_NUM _
                                , vlUserEmployeeNumber _
                                , trErrAdjustEntry_F _
                                , mstrCBUSINESSERR _
                                , "Invalid TRIS User. User " & CStr(vlUserEmployeeNumber) & _
                                    " does not exist." _
                                , MODULENAME & ": ValidateReqFirstAppEntries"
                        End If
                    End If

                    If (!EMP_NUM > 0) Then
                        If (RecordsExist(ladrsEmployee)) Then
                            ladrsEmployee.MoveFirst
                            ladrsEmployee.Find "EMP_NUM = " & !EMP_NUM

                            If Not (ladrsEmployee.EOF) Then
                                If (Trim(ladrsEmployee!EXEC_CONTRACT_IND) = "E") Then
                                    If (!SPECIAL_CHECK_IND = "Y") Then
                                        !SPECIAL_CHECK_IND = "N"
                                        .Update

                                        AppendBRV radrsBRV _
                                            , !EMP_NUM _
                                            , !CREATED_DATE _
                                            , !DEPT_NUM _
                                            , vlUserEmployeeNumber _
                                            , trErrAdjustEntry_F _
                                            , mstrCBUSINESSWAR _
                                            , "Special check is not applicable to executive employees." & _
                                                " Special check has been removed for this request." _
                                            , MODULENAME
                                    End If
                                End If
                            Else
                                Select Case (vsCallerID)
                                    Case "REQ"
                                        Err.Raise trErrProcessRequestAdjust_F
                                    Case "FIRST"
                                        Err.Raise trErrProcessFirstAppAdjust_F
                                End Select
                            End If
                        End If
                    End If

                    If Not (CheckForValidString(!SPECIAL_CHECK_IND, True, 1)) Then
                        !SPECIAL_CHECK_IND = "N"
                        .Update
                    End If

                    'Check for the Length of the Reason
                    If (CheckForValidString(!REQ_REASON)) Then
                        lsOrigReason = Trim(!REQ_REASON)
                        lsParsedReason = Trim(ParseString(Trim(!REQ_REASON)))

                        'Get the Length of the Reason
                        If (Len(lsParsedReason) > 255) Then
                            'Send a warning to the user stating that anything
                            'above 255 will be truncated.
                            !REQ_REASON = Left(lsParsedReason, 255)
                            .Update

                            AppendBRV radrsBRV _
                                , !EMP_NUM _
                                , !CREATED_DATE _
                                , !DEPT_NUM _
                                , vlUserEmployeeNumber _
                                , trErrAdjustEntry_F _
                                , mstrCBUSINESSWAR _
                                , "Adjustment request reason is more than 255 characters long." & _
                                    " Adjustment request reason was truncated to 255 characters for this request." _
                                , MODULENAME
                        Else
                            !REQ_REASON = lsParsedReason
                            .Update
                        End If

                        If (Trim(lsOrigReason) <> Trim(!REQ_REASON)) Then
                            AppendBRV radrsBRV _
                                , !EMP_NUM _
                                , !CREATED_DATE _
                                , !DEPT_NUM _
                                , vlUserEmployeeNumber _
                                , trErrAdjustEntry_F _
                                , mstrCBUSINESSWAR _
                                , "Special characters (Tab, New Line, Semi Colon " & _
                                    ", Tilda, etc.) were found and removed from request reason." _
                                , MODULENAME
                        End If
                    Else
                        AppendBRV radrsBRV _
                            , !EMP_NUM _
                            , !CREATED_DATE _
                            , !DEPT_NUM _
                            , vlUserEmployeeNumber _
                            , trErrAdjustEntry_F _
                            , mstrCBUSINESSERR _
                            , "Invalid adjustment request reason." _
                            , MODULENAME & ": ValidateReqFirstAppEntries"
                    End If
                    
                    Select Case (Trim(!APPROVED_IND))
                        Case "Y"
                            If (!APPROVED_EMP_NUM = 0) Then
                                !APPROVED_EMP_NUM = vlUserEmployeeNumber
                            End If

                        Case "N"
                            If (!APPROVED_EMP_NUM > 0) Then
                                !APPROVED_EMP_NUM = 0
                            End If
                    End Select

                    If Not (CheckForValidString(!CREATED_TIMESTAMP, , , True, 26)) Then
                        AppendBRV radrsBRV _
                            , !EMP_NUM _
                            , !CREATED_DATE _
                            , !DEPT_NUM _
                            , vlUserEmployeeNumber _
                            , trErrAdjustEntry_F _
                            , mstrCBUSINESSERR _
                            , "Invalid time stamp." _
                            , MODULENAME & ": ValidateReqFirstAppEntries"
                    End If
            End Select

            .MoveNext
        Next llRecordIndex
    End With

    ValidateReqFirstAppEntries = True

    If (RecordsExist(radrsBRV)) Then
        With radrsBRV
            .Filter = adFilterNone

            'MDAC25 FILTER FIX
            lsFilterString = "ERR_TYPE = 'BE' OR ERR_TYPE = 'AE' OR ERR_TYPE = 'AW'"
            .Filter = stdmTRISCommon.MDAC25FilterFix(lsFilterString)
            lsFilterString = ""
            If Not (.EOF) Then
                ValidateReqFirstAppEntries = False
            End If

            .Filter = adFilterNone
        End With
    End If

    Set ladrsEmployee = Nothing

    Exit Function
ValidateReqFirstAppEntries_EH:
    Set ladrsEmployee = Nothing
    Set radrsBRV = Nothing

    TrsRaiseError MODULENAME _
        , "ValidateReqFirstAppEntries"

End Function
Private Function DefineBRV() As ADODB.Recordset
On Error GoTo DefineBRV_EH
Dim ladrs As New ADODB.Recordset

    With ladrs
        .Fields.Append "EMP_NUM", adInteger, , adFldUpdatable
        .Fields.Append "CREATED_DATE", adDate, , adFldUpdatable
        .Fields.Append "DEPT_NUM", adChar, 3, adFldUpdatable
        .Fields.Append "CREATED_EMP_NUM", adInteger, , adFldUpdatable

        .Fields.Append "ERR_NUMBER", adInteger, , adFldUpdatable
        .Fields.Append "ERR_TYPE", adChar, 2, adFldUpdatable
        .Fields.Append "ERR_DESCRIPTION", adChar, 300, adFldUpdatable
        .Fields.Append "ERR_SOURCE", adChar, 1000, adFldUpdatable

        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open
    End With

    Set DefineBRV = ladrs
    Set ladrs = Nothing

    Exit Function
DefineBRV_EH:
    Set ladrs = Nothing

    TrsRaiseError MODULENAME _
        , "DefineBRV"

End Function
Private Sub AppendBRV(ByRef radrsBRV As ADODB.Recordset _
            , ByVal vlEmployeeNum As Long _
            , ByVal vdtReqDate As Date _
            , ByVal vsDeptNum As String _
            , ByVal vlRequestByEmployee As Long _
            , ByVal vlErrorNumber As Long _
            , ByVal vsErrorType As String _
            , ByVal vsErrorDescription As String _
            , ByVal vsErrorSource As String)
On Error GoTo AppendBRV_EH

    With radrsBRV
        .AddNew

        !EMP_NUM = vlEmployeeNum
        !CREATED_DATE = vdtReqDate
        !DEPT_NUM = vsDeptNum
        !CREATED_EMP_NUM = vlRequestByEmployee
        !ERR_NUMBER = vlErrorNumber - vbObjectError
        !ERR_TYPE = Trim(vsErrorType)
        !ERR_DESCRIPTION = Trim(vsErrorDescription)
        !ERR_SOURCE = Trim(vsErrorSource)

        .Update
    End With

    Exit Sub
AppendBRV_EH:
    TrsRaiseError MODULENAME _
        , "AppendBRV"

End Sub
Private Function GetEmployeeInfo(ByVal vvArrOfEmployees As Variant _
                    ) As ADODB.Recordset
On Error GoTo GetEmployeeInfo_EH
Dim lobjEmployee As dbEmployee.cdEmployee

    Set GetEmployeeInfo = Nothing

    Set lobjEmployee = CreateObject( _
                    "dbEmployee.cdEmployee")
    Set GetEmployeeInfo = lobjEmployee.GetEmployee( _
                            vvArrOfEmployees)
    Set lobjEmployee = Nothing

    'Destroy Objects
    Set lobjEmployee = Nothing

    Exit Function
GetEmployeeInfo_EH:
    Set lobjEmployee = Nothing

    TrsRaiseError MODULENAME _
        , "GetEmployeeInfo"

End Function
Private Function CheckForItemExistenceInDB( _
                    ByVal viItemType As Integer _
                    , Optional ByVal vsStringItem As String = "" _
                    , Optional ByVal vlLongItem As Long = 0 _
                    ) As Boolean
On Error GoTo CheckForItemExistenceInDB_EH
Dim lobjEmployee As dbEmployee.cdEmployee
Dim lobjCodeTable As dbCodeTables.cdCodeTables

    Select Case (viItemType)
        Case EMPLOYEE
            Set lobjEmployee = CreateObject( _
                                    "dbEmployee.cdEmployee")
            CheckForItemExistenceInDB = _
                lobjEmployee.IsValidEmployee(vlLongItem)
            Set lobjEmployee = Nothing

        Case DEPARTMENT
            Set lobjCodeTable = CreateObject( _
                                    "dbCodeTables.cdCodeTables")
            CheckForItemExistenceInDB = _
                lobjCodeTable.IsAValidDepartment(vsStringItem)
            Set lobjCodeTable = Nothing
    End Select

    Set lobjEmployee = Nothing
    Set lobjCodeTable = Nothing

    Exit Function
CheckForItemExistenceInDB_EH:
    Set lobjEmployee = Nothing
    Set lobjCodeTable = Nothing

    TrsRaiseError MODULENAME _
        , "CheckForItemExistenceInDB"

End Function
