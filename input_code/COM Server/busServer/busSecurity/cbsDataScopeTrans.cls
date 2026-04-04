VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "cbsDataScopeTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'AMEREN Project: TRIS Conversion Phase 3
'February 2004
'Author: Prem M. Chander
'Module Description: This class will provide the user the ability
'to create/modify employee datascope

Option Explicit

Private Const ADOFIELDATTR = adFldUpdatable Or adFldIsNullable
Private Const MODULENAME As String = "busSecurity.cbsDataScopeTrans"
'This method will save datascope entries
Public Function ProcessSecurityProfileEntries( _
                    ByVal vadrsDataAccessInformation As ADODB.Recordset _
                    , ByVal vadrsSecurityLevelInformation As ADODB.Recordset _
                    , ByVal vlDataAccessEmployeeNumber As Long _
                    , ByRef radrsErrors As ADODB.Recordset _
                    , ByVal vlUserEmployeeNumber As Long _
                    ) As Boolean
On Error GoTo ProcessSecurityProfileEntries_EH
Dim lbEmptyRequest As Boolean, lbEmptySecurityLevel As Boolean
Dim lobjDataAccess As dbSecurity.cdDataScope
Dim lobjDataScope As dbSecurity.cdBuiltDataScope
Dim lobjLevelMaintenance As dbSecurity.cdLevelMaintenance
Dim llRecordIndex As Long

    lbEmptySecurityLevel = True
    If Not (vadrsSecurityLevelInformation Is Nothing) Then
        If Not (vadrsSecurityLevelInformation.BOF And _
            vadrsSecurityLevelInformation.EOF) Then
            lbEmptySecurityLevel = False
        End If
    End If

    'Check for request existence
    lbEmptyRequest = True
    If Not (vadrsDataAccessInformation Is Nothing) Then
        If Not (vadrsDataAccessInformation.BOF And _
            vadrsDataAccessInformation.EOF) Then
            lbEmptyRequest = False
        End If
    End If

    'Prepare BRV Recordset
    Set radrsErrors = CreateBRV

    If (lbEmptyRequest And _
        lbEmptySecurityLevel) Then
        AppendBRV radrsErrors _
            , trErrbusSecurityDataScope _
            , mstrCBUSINESSERR _
            , "There are no Security Profile entries to process." _
            , MODULENAME & ": ProcessSecurityProfileEntries"
        GoTo REPORTBRV
    End If

    If Not (lbEmptyRequest) Then
        With vadrsDataAccessInformation
'NOTE: APPR_LEVEL_IND is a Integer field in the database
'(10 - DIVISION, 20 - FUNCTION, 30 - DEPARTMENT
'40 - LOCATION, 50- WORKGROUP, 60 - EMPLOYEE)
'It has been declared as Character so that the if the attibute domain values
'change then only the server needs to be changed
            .Filter = adFilterNone
            'MDAC25 FILTER FIX
            .Filter = MDAC25FilterFix("APPR_LEVEL_IND = 'DIVISION'")
            If Not (.EOF) Then
                .MoveFirst

                For llRecordIndex = 1 To .RecordCount
                    !APPR_LEVEL_IND = "10"
                    .Update

                    .MoveNext
                Next llRecordIndex
            End If

            .Filter = adFilterNone
            'MDAC25 FILTER FIX
            .Filter = MDAC25FilterFix("APPR_LEVEL_IND = 'FUNCTION'")
            If Not (.EOF) Then
                .MoveFirst

                For llRecordIndex = 1 To .RecordCount
                    !APPR_LEVEL_IND = "20"

                    .Update
                    .MoveNext
                Next llRecordIndex
            End If

            .Filter = adFilterNone
            'MDAC25 FILTER FIX
            .Filter = MDAC25FilterFix("APPR_LEVEL_IND = 'DEPARTMENT'")
            If Not (.EOF) Then
                .MoveFirst

                For llRecordIndex = 1 To .RecordCount
                    !APPR_LEVEL_IND = "30"
                    .Update

                    .MoveNext
                Next llRecordIndex
            End If

            .Filter = adFilterNone
            'MDAC25 FILTER FIX
            .Filter = MDAC25FilterFix("APPR_LEVEL_IND = 'WORKGROUP'")
            If Not (.EOF) Then
                .MoveFirst

                For llRecordIndex = 1 To .RecordCount
                    !APPR_LEVEL_IND = "50"
                    .Update

                    .MoveNext
                Next llRecordIndex
            End If

            .Filter = adFilterNone
            'MDAC25 FILTER FIX
            .Filter = MDAC25FilterFix("APPR_LEVEL_IND = 'EMPLOYEE'")
            If Not (.EOF) Then
                .MoveFirst

                For llRecordIndex = 1 To .RecordCount
                    !APPR_LEVEL_IND = "60"
                    .Update

                    .MoveNext
                Next llRecordIndex
            End If

            .Filter = adFilterNone
            'MDAC25 FILTER FIX
            .Filter = MDAC25FilterFix("ACCESS_LEVEL = 'APPROVAL'")
            If Not (.EOF) Then
                .MoveFirst

                For llRecordIndex = 1 To .RecordCount
                    !ACCESS_LEVEL = "A"
                    .Update

                    .MoveNext
                Next llRecordIndex
            End If

            .Filter = adFilterNone
            .Filter = MDAC25FilterFix("ACCESS_LEVEL = 'ENTRY'")
            If Not (.EOF) Then
                .MoveFirst

                For llRecordIndex = 1 To .RecordCount
                    !ACCESS_LEVEL = "E"
                    .Update

                    .MoveNext
                Next llRecordIndex
            End If

            .Filter = adFilterNone
            .Filter = MDAC25FilterFix("ACCESS_LEVEL = 'VIEW'")
            If Not (.EOF) Then
                .MoveFirst

                For llRecordIndex = 1 To .RecordCount
                    !ACCESS_LEVEL = "V"
                    .Update

                    .MoveNext
                Next llRecordIndex
            End If

            .Filter = adFilterNone
            If Not (.EOF) Then
                .MoveFirst
                If Not (ValidateDataAccessEntries( _
                    vadrsDataAccessInformation _
                    , vlDataAccessEmployeeNumber _
                    , radrsErrors _
                    , vlUserEmployeeNumber)) Then
                    GoTo REPORTBRV
                End If

                .Filter = adFilterNone
                .MoveFirst
            End If
        End With

        Set lobjDataAccess = CreateObject( _
                                "dbSecurity.cdDataScope")
        If Not (lobjDataAccess.dbProcessSave( _
                vadrsDataAccessInformation _
                , vlDataAccessEmployeeNumber _
                , vlUserEmployeeNumber)) Then
            Set lobjDataAccess = Nothing
        End If
        
        '//-----Invalidate DataScope
        Set lobjDataScope = CreateObject("dbSecurity.cdBuiltDataScope")
        
        vadrsDataAccessInformation.Filter = "APPR_LEVEL_IND = '10'"
        
        If Not vadrsDataAccessInformation.EOF Then
            lobjDataScope.InvalidateDataScopeDivByEffEmp vlDataAccessEmployeeNumber, True
        Else
            vadrsDataAccessInformation.Filter = "APPR_LEVEL_IND = '20'"
            
            If Not vadrsDataAccessInformation.EOF Then
                lobjDataScope.InvalidateDataScopeFunByEffEmp vlDataAccessEmployeeNumber, True
            Else
                vadrsDataAccessInformation.Filter = "APPR_LEVEL_IND = '30'"
                
                If Not vadrsDataAccessInformation.EOF Then
                    lobjDataScope.InvalidateDataScopeDeptByEffEmp vlDataAccessEmployeeNumber, True
                Else
                    vadrsDataAccessInformation.Filter = "APPR_LEVEL_IND = '50'"
                    
                    If Not vadrsDataAccessInformation.EOF Then
                        lobjDataScope.InvalidateDataScopeWGByEffEmp vlDataAccessEmployeeNumber, True
                    Else
                        vadrsDataAccessInformation.Filter = "APPR_LEVEL_IND = '60'"
                        
                        If Not vadrsDataAccessInformation.EOF Then
                            lobjDataScope.InvalidateDataScopeEmpByEffEmp vlDataAccessEmployeeNumber, True
                        End If
                    End If
                End If
            End If
        End If
        
        vadrsDataAccessInformation.Filter = adFilterNone
        Set lobjDataScope = Nothing
    End If

    If Not (lbEmptySecurityLevel) Then
        If Not (ValidateSecurityLevel( _
                vadrsSecurityLevelInformation _
                , radrsErrors _
                , vlUserEmployeeNumber)) Then
            GoTo REPORTBRV
        End If

        Set lobjLevelMaintenance = CreateObject( _
                                "dbSecurity.cdLevelMaintenance")
        If Not (lobjLevelMaintenance.dbProcessSecurityLevelForEmployee( _
                vadrsSecurityLevelInformation _
                , vlUserEmployeeNumber)) Then
            Set lobjLevelMaintenance = Nothing
            Err.Raise trErrbusSecurityDataScope
        End If
    End If

    Set lobjDataAccess = Nothing
    Set lobjLevelMaintenance = Nothing

    Set vadrsDataAccessInformation = Nothing
    Set vadrsSecurityLevelInformation = Nothing
    Set radrsErrors = Nothing

    ProcessSecurityProfileEntries = True

    GetObjectContext().SetComplete

    Exit Function
ProcessSecurityProfileEntries_EH:
    'Abort Transaction
    GetObjectContext().SetAbort

    Set lobjDataAccess = Nothing
    Set lobjLevelMaintenance = Nothing

    Set vadrsDataAccessInformation = Nothing
    Set vadrsSecurityLevelInformation = Nothing
    Set radrsErrors = Nothing

    ProcessSecurityProfileEntries = False

    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "ProcessSecurityProfileEntries"
            Case trErrdbInsIntoApproval, trErrdbUpdApproval _
                , trErrdbDelFromApproval, trErrdbSaveApproval _
                , trErrdb_UpdEmpLevelAssignment, trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "ProcessSecurityProfileEntries" _
                    , Err.Number _
                    , , , , _
                    True
            Case Else
                TrsRaiseError MODULENAME _
                    , "ProcessSecurityProfileEntries" _
                    , trErrbusSecurityDataScope _
                    , , , , _
                    True
    End Select

    Exit Function

REPORTBRV:
    Set lobjDataAccess = Nothing
    Set lobjLevelMaintenance = Nothing

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
                    ProcessSecurityProfileEntries = False
                Else
                    'Complete Transaction
                    GetObjectContext().SetComplete
                    ProcessSecurityProfileEntries = True
                End If

                .Filter = adFilterNone
            Else
                'Complete Transaction
                GetObjectContext().SetComplete
                ProcessSecurityProfileEntries = True
            End If
        End With
    Else
        'Complete Transaction
        GetObjectContext().SetComplete
        ProcessSecurityProfileEntries = True
    End If

    Set vadrsDataAccessInformation = Nothing
    Set vadrsSecurityLevelInformation = Nothing

    Exit Function

End Function
Private Function CreateBRV() As ADODB.Recordset
On Error GoTo CreateBRV_EH
Dim ladrs As New ADODB.Recordset

    With ladrs
        .Fields.Append "APPR_LEVEL_IND", adInteger, , ADOFIELDATTR
        .Fields.Append "GROUP_ID", adInteger, , ADOFIELDATTR
        .Fields.Append "DEPT_NUM", adChar, 3, ADOFIELDATTR
        .Fields.Append "SEC_LEVEL", adInteger, , ADOFIELDATTR

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
Private Function ValidateDataAccessEntries( _
                    ByRef radrsRSToValidate As ADODB.Recordset _
                    , ByVal vlDataAccessEmployeeNumber As Long _
                    , ByRef radrsBRV As ADODB.Recordset _
                    , ByVal vlUserEmployeeNumber As Long _
                    ) As Boolean
On Error GoTo ValidateDataAccessEntries_EH
Dim llRecordIndex As Long
Dim lobjDB As dbSecurity.cdFunSecurity
Dim ladrsUser As ADODB.Recordset, ladrsEmployee As ADODB.Recordset
Dim llUserSecurityLevel As Long, llEmployeeSecurityLevel As Long

    ValidateDataAccessEntries = False

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

    If Not (CheckForValidLong(vlDataAccessEmployeeNumber, True)) Then
        AppendBRV radrsBRV _
            , trErrbusSecurityDataScope _
            , mstrCBUSINESSERR _
            , "Invalid Employee." _
            , MODULENAME & ": ValidateDataAccessEntries"
        GoTo SKIPTOEND
    End If

    If Not (CheckForValidLong(vlUserEmployeeNumber, True)) Then
        AppendBRV radrsBRV _
            , trErrbusSecurityDataScope _
            , mstrCBUSINESSERR _
            , "Invalid TRIS User." _
            , MODULENAME & ": ValidateDataAccessEntries"
        GoTo SKIPTOEND
    End If

    Set lobjDB = CreateObject("dbSecurity.cdFunSecurity")
    Set ladrsUser = lobjDB.db_CheckEmployee(vlUserEmployeeNumber)
    Set ladrsEmployee = lobjDB.db_CheckEmployee(vlDataAccessEmployeeNumber)
    Set lobjDB = Nothing

    llUserSecurityLevel = -1
    If Not (ladrsUser Is Nothing) Then
        If Not (ladrsUser.EOF And ladrsUser.BOF) Then
            ladrsUser.MoveFirst
            llUserSecurityLevel = ladrsUser!SEC_LEVEL
        End If
    End If
    Set ladrsUser = Nothing

    llEmployeeSecurityLevel = -1
    If Not (ladrsEmployee Is Nothing) Then
        If Not (ladrsEmployee.EOF And ladrsEmployee.BOF) Then
            ladrsEmployee.MoveFirst
            llEmployeeSecurityLevel = ladrsEmployee!SEC_LEVEL
        End If
    End If
    Set ladrsEmployee = Nothing

    If (llUserSecurityLevel = -1) Then
        AppendBRV radrsBRV _
            , trErrbusSecurityDataScope _
            , mstrCBUSINESSERR _
            , "Invalid TRIS User." _
            , MODULENAME & ": ValidateDataAccessEntries"
        GoTo SKIPTOEND
    End If

    If (llEmployeeSecurityLevel = -1) Then
        AppendBRV radrsBRV _
            , trErrbusSecurityDataScope _
            , mstrCBUSINESSERR _
            , "Invalid Employee." _
            , MODULENAME & ": ValidateDataAccessEntries"
        GoTo SKIPTOEND
    End If

    Select Case (llEmployeeSecurityLevel)
        Case 9, 10, 11, 981 'Admin levels
            If (llUserSecurityLevel <> 10) Then
                AppendBRV radrsBRV _
                    , trErrbusSecurityDataScope _
                    , mstrCBUSINESSERR _
                    , StringFormat("You are not authorized to modify the security profile (Level = {0}) for this employee {1}" & _
                    ", please submit an IdM TRIS Admin request.", llEmployeeSecurityLevel, CStr(vlDataAccessEmployeeNumber)) _
                    , MODULENAME & ": ValidateDataAccessEntries"
                GoTo SKIPTOEND
            End If

        Case 403 'Central Payroll Timekeepers
            Select Case (llUserSecurityLevel)
                Case 9, 10
                    '
                Case Else
                    AppendBRV radrsBRV _
                        , trErrbusSecurityDataScope _
                        , mstrCBUSINESSERR _
                        , "You are not authorized to modify the security profile for this employee " & _
                        CStr(vlDataAccessEmployeeNumber) & ", please contact the TRIS Team." _
                        , MODULENAME & ": ValidateDataAccessEntries"
                    GoTo SKIPTOEND
            End Select
    End Select

    With radrsRSToValidate
        For llRecordIndex = 1 To .RecordCount
            If Not (CheckForValidLong(!GROUP_ID, True)) Then
                Select Case (Trim(!APPR_LEVEL_IND))
                    Case "10"
                        AppendBRV radrsBRV _
                            , trErrbusSecurityDataScope _
                            , mstrCBUSINESSERR _
                            , "Invalid Division Number." _
                            , MODULENAME & ": ValidateDataAccessEntries" _
                            , 10, , , _
                            , 1
                    Case "20"
                        AppendBRV radrsBRV _
                            , trErrbusSecurityDataScope _
                            , mstrCBUSINESSERR _
                            , "Invalid Function Number." _
                            , MODULENAME & ": ValidateDataAccessEntries" _
                            , 20, , , _
                            , 2
                    Case "50"
                        AppendBRV radrsBRV _
                            , trErrbusSecurityDataScope _
                            , mstrCBUSINESSERR _
                            , "Invalid Workgroup Number." _
                            , MODULENAME & ": ValidateDataAccessEntries" _
                            , 50, , , _
                            , 4
                    Case "60"
                        AppendBRV radrsBRV _
                            , trErrbusSecurityDataScope _
                            , mstrCBUSINESSERR _
                            , "Invalid Employee Number." _
                            , MODULENAME & ": ValidateDataAccessEntries" _
                            , 60, , , _
                            , 5
                End Select
            End If

            If (Trim(!APPR_LEVEL_IND) = "30") Then
                If Not (CheckForValidString(!DEPT_NUM)) Then
                    AppendBRV radrsBRV _
                        , trErrbusSecurityDataScope _
                        , mstrCBUSINESSERR _
                        , "Invalid Department Number." _
                        , MODULENAME & ": ValidateDataAccessEntries" _
                        , 30, , , _
                        , 3
                End If
            End If

            Select Case (!RECORD_STATUS)
                Case "D", "C"
                    If Not (CheckForValidLong(!KEY_GROUP_ID, True)) Then
                        Select Case (Trim(!APPR_LEVEL_IND))
                            Case "10"
                                AppendBRV radrsBRV _
                                    , trErrbusSecurityDataScope _
                                    , mstrCBUSINESSERR _
                                    , "Invalid delete/update request. Division number is invalid." _
                                    , MODULENAME & ": ValidateDataAccessEntries" _
                                    , 10, !KEY_GROUP_ID, , _
                                    , 1
                            Case "20"
                                AppendBRV radrsBRV _
                                    , trErrbusSecurityDataScope _
                                    , mstrCBUSINESSERR _
                                    , "Invalid delete/update request. Function number is invalid." _
                                    , MODULENAME & ": ValidateDataAccessEntries" _
                                    , 20, !KEY_GROUP_ID, , _
                                    , 2
                            Case "50"
                                AppendBRV radrsBRV _
                                    , trErrbusSecurityDataScope _
                                    , mstrCBUSINESSERR _
                                    , "Invalid delete/update request. Workgroup number is invalid." _
                                    , MODULENAME & ": ValidateDataAccessEntries" _
                                    , 50, !KEY_GROUP_ID, , _
                                    , 4
                            Case "60"
                                AppendBRV radrsBRV _
                                    , trErrbusSecurityDataScope _
                                    , mstrCBUSINESSERR _
                                    , "Invalid delete/update request. Employee number is invalid." _
                                    , MODULENAME & ": ValidateDataAccessEntries" _
                                    , 60, !KEY_GROUP_ID, , _
                                    , 5
                        End Select
                    End If

                    If (Trim(!APPR_LEVEL_IND) = "30") Then
                        If Not (CheckForValidString(!KEY_DEPT_NUM)) Then
                            AppendBRV radrsBRV _
                                , trErrbusSecurityDataScope _
                                , mstrCBUSINESSERR _
                                , "Invalid delete/update request. Department number is invalid." _
                                , MODULENAME & ": ValidateDataAccessEntries" _
                                , 30, , !KEY_DEPT_NUM, _
                                , 5
                        End If
                    End If

                    If Not (CheckForValidString(!APP_TIME_STAMP, , , True, 26)) Then
                        AppendBRV radrsBRV _
                            , trErrbusSecurityDataScope _
                            , mstrCBUSINESSERR _
                            , "Invalid delete/update request. Time Stamp is invalid." _
                            , MODULENAME & ": ValidateDataAccessEntries" _
                            , !APPR_LEVEL_IND, !KEY_GROUP_ID, !KEY_DEPT_NUM
                    End If
            End Select

            .MoveNext
        Next llRecordIndex
    End With

SKIPTOEND:
    radrsRSToValidate.Filter = adFilterNone
    radrsRSToValidate.MoveFirst

    ValidateDataAccessEntries = True

    If Not (radrsBRV Is Nothing) Then
        With radrsBRV
            If Not (.BOF And .EOF) Then
                .Filter = adFilterNone
                .Filter = MDAC25FilterFix("ERR_TYPE = 'BE' OR ERR_TYPE = 'AE' OR ERR_TYPE = 'AW'")
                If Not (.EOF) Then
                    ValidateDataAccessEntries = False
                End If

                .Filter = adFilterNone
            End If
        End With
    End If

    Set lobjDB = Nothing
    Set ladrsEmployee = Nothing
    Set ladrsUser = Nothing

    Exit Function
ValidateDataAccessEntries_EH:
    Set lobjDB = Nothing
    Set ladrsEmployee = Nothing
    Set ladrsUser = Nothing

    ValidateDataAccessEntries = False

    TrsRaiseError MODULENAME _
            , "ValidateDataAccessEntries"

End Function
Private Function ValidateSecurityLevel( _
                     ByRef radrsRSToValidate As ADODB.Recordset _
                    , ByRef radrsBRV As ADODB.Recordset _
                    , ByVal vlUserEmployeeNumber As Long _
                    ) As Boolean
On Error GoTo ValidateSecurityLevel_EH
    Dim llRecordIndex As Long
    Dim lobjLevelMaintenance As dbSecurity.cdLevelMaintenance
    Dim lobjFunctionalSecurity As dbSecurity.cdFunSecurity
    Dim ladrs As ADODB.Recordset, llUserSecurityLevel As Long
    Dim llEmployeeSecurityLevel As Long
    Dim lobjEnvironment As dbEnvironment.cdEnvironment
    Dim lbTRISDeveloper As Boolean

    ValidateSecurityLevel = False

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

    llUserSecurityLevel = -1
    If Not (CheckForValidLong(vlUserEmployeeNumber, True)) Then
        AppendBRV radrsBRV _
            , trErrbusSecurityDataScope _
            , mstrCBUSINESSERR _
            , "Invalid TRIS User." _
            , MODULENAME & ": ValidateSecurityLevel"
    Else
        Set lobjFunctionalSecurity = CreateObject( _
                                "dbSecurity.cdFunSecurity")
        Set ladrs = lobjFunctionalSecurity.db_CheckEmployee(vlUserEmployeeNumber)
        Set lobjFunctionalSecurity = Nothing

        If Not (ladrs Is Nothing) Then
            If Not (ladrs.EOF And ladrs.BOF) Then
                ladrs.MoveFirst
                llUserSecurityLevel = ladrs!SEC_LEVEL
            End If
        End If
    End If

    With radrsRSToValidate
        For llRecordIndex = 1 To .RecordCount
            llEmployeeSecurityLevel = -1

            If Not (CheckForValidLong(!EMP_NUM, True)) Then
                AppendBRV radrsBRV _
                    , trErrbusSecurityDataScope _
                    , mstrCBUSINESSERR _
                    , "Invalid Employee." _
                    , MODULENAME & ": ValidateSecurityLevel" _
                    , , , , !SEC_LEVEL
            Else
                Set lobjFunctionalSecurity = CreateObject( _
                                        "dbSecurity.cdFunSecurity")
                Set ladrs = lobjFunctionalSecurity.db_CheckEmployee(!EMP_NUM)
                Set lobjFunctionalSecurity = Nothing

                If Not (ladrs Is Nothing) Then
                    If (ladrs.EOF And ladrs.BOF) Then
                        AppendBRV radrsBRV _
                            , trErrbusSecurityDataScope _
                            , mstrCBUSINESSERR _
                            , "Invalid Employee. Employee " & CStr(!EMP_NUM) & " does not exist." _
                            , MODULENAME & ": ValidateSecurityLevel" _
                            , , , , !SEC_LEVEL
                    Else
                        llEmployeeSecurityLevel = ladrs!SEC_LEVEL
                    End If
                Else
                    AppendBRV radrsBRV _
                        , trErrbusSecurityDataScope _
                        , mstrCBUSINESSERR _
                        , "Invalid Employee. Employee " & CStr(!EMP_NUM) & " does not exist." _
                        , MODULENAME & ": ValidateSecurityLevel" _
                        , , , , !SEC_LEVEL
                End If
            End If
            Set ladrs = Nothing

            Select Case (llEmployeeSecurityLevel)
                Case 9, 10, 11, 981 'Admin levels
                    If (llUserSecurityLevel <> 10) Then
                        AppendBRV radrsBRV _
                            , trErrbusSecurityDataScope _
                            , mstrCBUSINESSERR _
                            , StringFormat("You are not authorized to modify the security profile (Level = {0}) for this employee {1} " & _
                            ", please contact the TRIS Team.", llEmployeeSecurityLevel, CStr(!EMP_NUM)) _
                            , MODULENAME & ": ValidateSecurityLevel" _
                            , , , , !SEC_LEVEL
                    End If

                Case 403 'Central Payroll Timekeepers
                    Select Case (llUserSecurityLevel)
                        Case 9, 10
                            '
                        Case Else
                            AppendBRV radrsBRV _
                                , trErrbusSecurityDataScope _
                                , mstrCBUSINESSERR _
                                , "You are not authorized to modify the security profile for this employee " & _
                                CStr(!EMP_NUM) & ", please contact the TRIS Team." _
                                , MODULENAME & ": ValidateSecurityLevel" _
                                , , , , !SEC_LEVEL
                    End Select
            End Select

            If Not (IsNumeric(!SEC_LEVEL)) Or _
                (!SEC_LEVEL = -1) Then
                AppendBRV radrsBRV _
                    , trErrbusSecurityDataScope _
                    , mstrCBUSINESSERR _
                    , "Invalid Security Level." _
                    , MODULENAME & ": ValidateSecurityLevel"
            Else
                'Determine system admin/TRIS Developer
                Set lobjEnvironment = CreateObject("dbEnvironment.cdEnvironment")
                lbTRISDeveloper = lobjEnvironment.IsDeveloper(vlUserEmployeeNumber)
                Set lobjEnvironment = Nothing

                Set lobjLevelMaintenance = CreateObject("dbSecurity.cdLevelMaintenance")
                Set ladrs = lobjLevelMaintenance.db_GetLevels(GETSECLVLINFOBYLEVEL, !SEC_LEVEL, , lbTRISDeveloper)
                Set lobjLevelMaintenance = Nothing

                'TODO: This could have used RecordsExist() and ONE AppendBRV...
                If Not (ladrs Is Nothing) Then
                    If (ladrs.EOF And ladrs.BOF) Then
                        AppendBRV radrsBRV _
                            , trErrbusSecurityDataScope _
                            , mstrCBUSINESSERR _
                            , "Invalid Security Level. Security Level " & CStr(!SEC_LEVEL) & " does not exist." _
                            , MODULENAME & ": ValidateSecurityLevel" _
                            , , , , !SEC_LEVEL
                    End If
                    Set ladrs = Nothing
                Else
                    AppendBRV radrsBRV _
                        , trErrbusSecurityDataScope _
                        , mstrCBUSINESSERR _
                        , "Invalid Security Level. Security Level " & CStr(!SEC_LEVEL) & " does not exist." _
                        , MODULENAME & ": ValidateSecurityLevel" _
                        , , , , !SEC_LEVEL
                End If
            End If
            Set ladrs = Nothing

            If (Trim(!RECORD_STATUS) = "D") Then
                If Not (CheckForValidString(!TIME_STAMP, , , True, 26)) Then
                    AppendBRV radrsBRV _
                        , trErrbusSecurityDataScope _
                        , mstrCBUSINESSERR _
                        , "Invalid Time Stamp." _
                        , MODULENAME & ": ValidateSecurityLevel" _
                        , , , , !SEC_LEVEL
                End If
            End If
            .MoveNext
        Next llRecordIndex
    End With

    Set ladrs = Nothing
    Set lobjLevelMaintenance = Nothing

    ValidateSecurityLevel = True

    If Not (radrsBRV Is Nothing) Then
        With radrsBRV
            If Not (.BOF And .EOF) Then
                .Filter = adFilterNone
                .Filter = MDAC25FilterFix("ERR_TYPE = 'BE' OR ERR_TYPE = 'AE' OR ERR_TYPE = 'AW'")
                If Not (.EOF) Then
                    ValidateSecurityLevel = False
                End If

                .Filter = adFilterNone
            End If
        End With
    End If

    Exit Function
ValidateSecurityLevel_EH:
    Set lobjFunctionalSecurity = Nothing
    Set lobjLevelMaintenance = Nothing
    Set lobjEnvironment = Nothing

    Set ladrs = Nothing

    ValidateSecurityLevel = False

    TrsRaiseError MODULENAME, "ValidateSecurityLevel"
End Function

Private Sub AppendBRV(ByRef radrsBRV As ADODB.Recordset _
                , ByVal vlErrorNumber As Long _
                , ByVal vsErrorType As String _
                , ByVal vsErrorDescription As String _
                , ByVal vsErrorSource As String _
                , Optional ByVal vlApprovalLevel As Long = 0 _
                , Optional ByVal vlGroupID As Long = 0 _
                , Optional ByVal vsDepartmentNumber As String = "" _
                , Optional ByVal vlSecurityLevel As Long = -1 _
                , Optional ByVal viRecordSource As Integer = 0 _
                )
On Error GoTo AppendBRV_EH

    With radrsBRV
        If (.State <> adStateOpen) Then
            Exit Sub
        End If

        .AddNew

        !APPR_LEVEL_IND = vlApprovalLevel
        !GROUP_ID = vlGroupID
        !DEPT_NUM = vsDepartmentNumber
        !SEC_LEVEL = vlSecurityLevel

        !ERR_NUMBER = vlErrorNumber - vbObjectError
        !ERR_TYPE = Trim(vsErrorType)
        !REC_SOURCE = viRecordSource
        Select Case (viRecordSource)
                Case 1
                    !ERR_DESCRIPTION = "Division access error: " & Trim(vsErrorDescription)
                Case 2
                    !ERR_DESCRIPTION = "Function access error: " & Trim(vsErrorDescription)
                Case 3
                    !ERR_DESCRIPTION = "Department access error: " & Trim(vsErrorDescription)
                Case 4
                    !ERR_DESCRIPTION = "Workgroup access error: " & Trim(vsErrorDescription)
                Case 5
                    !ERR_DESCRIPTION = "Employee access error: " & Trim(vsErrorDescription)
                Case Else
                    !ERR_DESCRIPTION = Trim(vsErrorDescription)
        End Select
        !ERR_SOURCE = Trim(vsErrorSource)

        .Update
    End With

    Exit Sub
AppendBRV_EH:
    TrsRaiseError MODULENAME _
        , "AppendBRV"

End Sub
