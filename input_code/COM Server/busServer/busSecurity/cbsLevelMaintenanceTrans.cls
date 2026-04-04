VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "cbsLevelMaintenanceTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'AMEREN Project: TRIS Conversion Phase 2
'April 2002
'Author: Prem M. Chander
'Module Description: This class will provide the user the ability
'to create/modify security level information
Option Explicit

Private Const ADOFIELDATTR = adFldUpdatable Or adFldIsNullable
Private Const MODULENAME As String = "busSecurity.cbsLevelMaintenanceTrans"
Public Function ProcessSecurityLevelInformation( _
                    ByRef r_lSecurityLevel As Long _
                    , ByVal v_adrsSecurityLevelInfo As ADODB.Recordset _
                    , ByVal v_adrsEmployeeAssignments As ADODB.Recordset _
                    , ByVal v_adrsBusFuncAssignments As ADODB.Recordset _
                    , ByRef r_adrsErrors As ADODB.Recordset _
                    , ByVal v_lUserEmployeeNumber As Long _
                    ) As Boolean
On Error GoTo ProcessSecurityLevelInformation_EH
Dim lobjDataAccess As dbSecurity.cdLevelMaintenance
Dim lbEmptySecLvlInfo As Boolean, lbEmptyEmpInfo As Boolean
Dim lbEmptyBusFuncInfo As Boolean, lbEmployeesExistForLevel As Boolean
Dim ladrsEmployees As ADODB.Recordset

    Set r_adrsErrors = DefineSecLvlInfoBRV

    lbEmptySecLvlInfo = True
    If Not (v_adrsSecurityLevelInfo Is Nothing) Then
        If Not (v_adrsSecurityLevelInfo.BOF And _
            v_adrsSecurityLevelInfo.EOF) Then
            lbEmptySecLvlInfo = False
        End If
    End If

    lbEmptyEmpInfo = True
    If Not (v_adrsEmployeeAssignments Is Nothing) Then
        If Not (v_adrsEmployeeAssignments.EOF And _
            v_adrsEmployeeAssignments.BOF) Then
            lbEmptyEmpInfo = False
        End If
    End If

    lbEmptyBusFuncInfo = True
    If Not (v_adrsBusFuncAssignments Is Nothing) Then
        If Not (v_adrsBusFuncAssignments.EOF And _
            v_adrsBusFuncAssignments.BOF) Then
            lbEmptyBusFuncInfo = False
        End If
    End If

    If (lbEmptySecLvlInfo And _
        lbEmptyEmpInfo And _
        lbEmptyBusFuncInfo) Then
        AppendSecLvlInfoBRV r_adrsErrors _
            , trErrSecurityLvlInfo_F _
            , mstrCBUSINESSERR _
            , "There are no Security Level Information entries to process" _
            , MODULENAME & ": ProcessSecurityLevelInformation"
        GoTo REPORTBRV
    End If

    If Not (CheckForValidLong(v_lUserEmployeeNumber, True)) Then
        AppendSecLvlInfoBRV _
            r_adrsErrors _
            , trErrSecurityLvlInfo_F _
            , mstrCBUSINESSERR _
            , "Invalid TRIS User." _
            , MODULENAME & ": ProcessSecurityLevelInformation"
        GoTo REPORTBRV
    End If

    If Not (lbEmptySecLvlInfo) Then
        With v_adrsSecurityLevelInfo
            .Filter = adFilterNone

            'Validate Request for any processing (DIC)
            If Not (.EOF) Then
                .MoveFirst

                If Not (ValidateSecurityLevelInfo( _
                    v_adrsSecurityLevelInfo _
                    , r_adrsErrors)) Then
                    GoTo REPORTBRV
                End If

                .Filter = adFilterNone
                .MoveFirst
            End If
        End With
    End If

    If Not (lbEmptyEmpInfo) Then
        With v_adrsEmployeeAssignments
            .Filter = adFilterNone

            'Validate Request for any processing (DIC)
            If Not (.EOF) Then
                .MoveFirst

                If Not (ValidateEmployeeAssignments( _
                    v_adrsEmployeeAssignments _
                    , r_adrsErrors _
                    , v_lUserEmployeeNumber)) Then
                    GoTo REPORTBRV
                End If

                .Filter = adFilterNone
                .MoveFirst
            End If
        End With
    End If

    If Not (lbEmptyBusFuncInfo) Then
        With v_adrsBusFuncAssignments
            .Filter = adFilterNone

            'Validate Request for any processing (DIC)
            If Not (.EOF) Then
                .MoveFirst

                If Not (ValidateBusFuncAssignments( _
                    v_adrsBusFuncAssignments _
                    , r_adrsErrors)) Then
                    GoTo REPORTBRV
                End If

                .Filter = adFilterNone
                .MoveFirst
            End If
        End With
    End If

    Set lobjDataAccess = CreateObject( _
                            "dbSecurity.cdLevelMaintenance")
    If Not (lobjDataAccess.db_ProcessSave(r_lSecurityLevel _
            , v_adrsSecurityLevelInfo _
            , v_adrsEmployeeAssignments _
            , v_adrsBusFuncAssignments _
            , v_lUserEmployeeNumber)) Then
        Set lobjDataAccess = Nothing
        Err.Raise trErrSecurityLvlInfo_F
    End If
    Set lobjDataAccess = Nothing

    Set v_adrsSecurityLevelInfo = Nothing
    Set v_adrsEmployeeAssignments = Nothing
    Set v_adrsBusFuncAssignments = Nothing
    Set r_adrsErrors = Nothing

    GetObjectContext().SetComplete
    ProcessSecurityLevelInformation = True

    Exit Function
ProcessSecurityLevelInformation_EH:
    'Abort Transaction
    GetObjectContext().SetAbort

    Set v_adrsSecurityLevelInfo = Nothing
    Set v_adrsEmployeeAssignments = Nothing
    Set v_adrsBusFuncAssignments = Nothing
    Set r_adrsErrors = Nothing

    Set lobjDataAccess = Nothing

    ProcessSecurityLevelInformation = False

    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "ProcessSecurityLevelInformation"
            Case trErrdb_InsSecLevel, trErrdb_UpdSecLevel _
                , trErrdb_DelSecLevel, trErrdb_UpdEmpLevelAssignment _
                , trErrdb_BusFuncLevelAssignment, trErrdb_ProcessSave_F _
                , trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "ProcessSecurityLevelInformation" _
                    , Err.Number _
                    , , , , _
                    True
            Case Else
                TrsRaiseError MODULENAME _
                    , "ProcessSecurityLevelInformation" _
                    , trErrSecurityLvlInfo_F _
                    , , , , _
                    True
    End Select

    Exit Function

REPORTBRV:
    Set lobjDataAccess = Nothing

    If Not (r_adrsErrors Is Nothing) Then
        With r_adrsErrors
            If Not (.EOF And .BOF) Then
                .Filter = adFilterNone

                'Check for errors
'MDAC25 FILTER FIX
                .Filter = MDAC25FilterFix("ERR_TYPE = 'BE'")
                'If found abort else complete
                If Not (.EOF) Then
                    'Abort Transaction
                    GetObjectContext().SetAbort
                    ProcessSecurityLevelInformation = False
                Else
                    'Complete Transaction
                    GetObjectContext().SetComplete
                    ProcessSecurityLevelInformation = True
                End If

                .Filter = adFilterNone
            Else
                'Complete Transaction
                GetObjectContext().SetComplete
                ProcessSecurityLevelInformation = True
            End If
        End With
    Else
        'Complete Transaction
        GetObjectContext().SetComplete
        ProcessSecurityLevelInformation = True
    End If

    Set v_adrsSecurityLevelInfo = Nothing
    Set v_adrsEmployeeAssignments = Nothing
    Set v_adrsBusFuncAssignments = Nothing
    Set lobjDataAccess = Nothing

    Exit Function

End Function
Private Function DefineSecLvlInfoBRV() As ADODB.Recordset
On Error GoTo DefineSecLvlInfoBRV_EH

Dim ladrsBRV As New ADODB.Recordset

    With ladrsBRV
        .Fields.Append "SEC_LEVEL", adInteger, , ADOFIELDATTR
        .Fields.Append "EMP_NUM", adInteger, , ADOFIELDATTR
        .Fields.Append "BUS_FUNC_KEY", adChar, 9, ADOFIELDATTR

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

    Set DefineSecLvlInfoBRV = ladrsBRV
    Set ladrsBRV = Nothing

    Exit Function

DefineSecLvlInfoBRV_EH:
    Set ladrsBRV = Nothing
    Set DefineSecLvlInfoBRV = Nothing

    TrsRaiseError MODULENAME _
        , "DefineSecLvlInfoBRV"

End Function
Private Sub AppendSecLvlInfoBRV(ByRef r_adrsBRV As ADODB.Recordset _
                , ByVal v_lErrorNumber As Long _
                , ByVal v_sErrorType As String _
                , ByVal v_sErrorDescription As String _
                , ByVal v_sErrorSource As String _
                , Optional ByVal v_lSecurityLevel As Long = 0 _
                , Optional ByVal v_lEmployeeNumber As Long = 0 _
                , Optional ByVal v_sBusFuncKey As String = "" _
                , Optional ByVal v_iRecordSource As Integer = 0 _
                )
On Error GoTo AppendSecLvlInfoBRV_EH

    With r_adrsBRV
        If (.State <> adStateOpen) Then
            Exit Sub
        End If

        .AddNew

        !SEC_LEVEL = v_lSecurityLevel
        !EMP_NUM = v_lEmployeeNumber
        !BUS_FUNC_KEY = v_sBusFuncKey

        !ERR_NUMBER = v_lErrorNumber - vbObjectError
        !ERR_TYPE = Trim(v_sErrorType)
        !REC_SOURCE = Trim(v_iRecordSource)
        Select Case (Trim(v_iRecordSource))
                Case 1
                    !ERR_DESCRIPTION = "Security Level entry error: " & Trim(v_sErrorDescription)
                Case 2
                    !ERR_DESCRIPTION = "Employee Assignment entry error: " & Trim(v_sErrorDescription)
                Case 3
                    !ERR_DESCRIPTION = "Business Function Assignment entry error: " & Trim(v_sErrorDescription)
                Case Else
                    !ERR_DESCRIPTION = Trim(v_sErrorDescription)
        End Select
        !ERR_SOURCE = Trim(v_sErrorSource)

        .Update
    End With

    Exit Sub
AppendSecLvlInfoBRV_EH:
    TrsRaiseError MODULENAME _
        , "AppendSecLvlInfoBRV"

End Sub
Private Function ValidateSecurityLevelInfo( _
                    ByRef r_adrsRSToProcess As ADODB.Recordset _
                    , ByRef r_adrsBRV As ADODB.Recordset _
                    ) As Boolean
On Error GoTo ValidateSecurityLevelInfo_EH
Dim llRecordIndex As Long
Dim lsOrigDesc As String, lsParsedDesc As String

    ValidateSecurityLevelInfo = False

    If Not (r_adrsRSToProcess Is Nothing) Then
        If Not (r_adrsRSToProcess.EOF And _
            r_adrsRSToProcess.BOF) Then
            r_adrsRSToProcess.Filter = adFilterNone
            r_adrsRSToProcess.MoveFirst
        Else
            Exit Function
        End If
    Else
        Exit Function
    End If

    With r_adrsRSToProcess
        For llRecordIndex = 1 To .RecordCount
            Select Case (!RECORD_STATUS)
                Case "D"    'DELETE
                    If Not (CheckForValidLong(!SEC_LEVEL)) Then
                        AppendSecLvlInfoBRV _
                            r_adrsBRV _
                            , trErrSecurityLvlInfo_F _
                            , mstrCBUSINESSERR _
                            , "Invalid Security Level." _
                            , MODULENAME & ": ValidateSecurityLevelInfo" _
                            , _
                            , _
                            , _
                            , 1
                    End If

                    If Not (CheckForValidString(!DIVISION_NUM)) Then
                        AppendSecLvlInfoBRV _
                            r_adrsBRV _
                            , trErrSecurityLvlInfo_F _
                            , mstrCBUSINESSERR _
                            , "Invalid Division." _
                            , MODULENAME & ": ValidateSecurityLevelInfo" _
                            , !SEC_LEVEL _
                            , _
                            , _
                            , 1
                    End If

                    If Not (CheckForValidString(!TIME_STAMP, , , True, 26)) Then
                        AppendSecLvlInfoBRV _
                            r_adrsBRV _
                            , trErrSecurityLvlInfo_F _
                            , mstrCBUSINESSERR _
                            , "Invalid Time Stamp." _
                            , MODULENAME & ": ValidateSecurityLevelInfo" _
                            , !SEC_LEVEL _
                            , _
                            , _
                            , 1
                    End If

                Case "I"    'INSERT
                    If Not (CheckForValidLong(!SEC_LEVEL)) Then
                        AppendSecLvlInfoBRV _
                            r_adrsBRV _
                            , trErrSecurityLvlInfo_F _
                            , mstrCBUSINESSERR _
                            , "Invalid Security Level." _
                            , MODULENAME & ": ValidateSecurityLevelInfo" _
                            , _
                            , _
                            , _
                            , 1
                    End If

                    If Not (CheckForValidString(!SEC_LEVEL_DESC, True, 30)) Then
                        AppendSecLvlInfoBRV _
                            r_adrsBRV _
                            , trErrSecurityLvlInfo_F _
                            , mstrCBUSINESSERR _
                            , "Invalid Security Level Description." _
                            , MODULENAME & ": ValidateSecurityLevelInfo" _
                            , !SEC_LEVEL _
                            , _
                            , _
                            , 1
                    Else
                        lsOrigDesc = Trim(!SEC_LEVEL_DESC)
                        lsParsedDesc = ParseString(Trim(!SEC_LEVEL_DESC))
                        If (Len(lsParsedDesc) > 30) Then
                            lsParsedDesc = Left(lsParsedDesc, 30)
                        End If

                        If (lsParsedDesc <> lsOrigDesc) Then
                            'update the attribute in the record
                            !SEC_LEVEL_DESC = lsParsedDesc
                            .Update

                            AppendSecLvlInfoBRV _
                                r_adrsBRV _
                                , trErrSecurityLvlInfo_F _
                                , mstrCBUSINESSWAR _
                                , "Special characters (New Line, Semi Colon " & _
                                ", Colon, Single Quote, Double Quote etc.) were found and removed in Security Level Description." _
                                , MODULENAME & ": ValidateSecurityLevelInfo" _
                                , !SEC_LEVEL _
                                , _
                                , _
                                , 1
                        End If
                    End If

                    If Not (CheckForValidString(!DIVISION_NUM)) Then
                        AppendSecLvlInfoBRV _
                            r_adrsBRV _
                            , trErrSecurityLvlInfo_F _
                            , mstrCBUSINESSERR _
                            , "Invalid Division." _
                            , MODULENAME & ": ValidateSecurityLevelInfo" _
                            , !SEC_LEVEL _
                            , _
                            , _
                            , 1
                    End If

                Case "C"    'CHANGE
                    If Not (CheckForValidLong(!SEC_LEVEL)) Then
                        AppendSecLvlInfoBRV _
                            r_adrsBRV _
                            , trErrSecurityLvlInfo_F _
                            , mstrCBUSINESSERR _
                            , "Invalid Security Level." _
                            , MODULENAME & ": ValidateSecurityLevelInfo" _
                            , _
                            , _
                            , _
                            , 1
                    End If

                    If Not (CheckForValidString(!SEC_LEVEL_DESC, True, 30)) Then
                        AppendSecLvlInfoBRV _
                            r_adrsBRV _
                            , trErrSecurityLvlInfo_F _
                            , mstrCBUSINESSERR _
                            , "Invalid Security Level Description" _
                            , MODULENAME & ": ValidateSecurityLevelInfo" _
                            , !SEC_LEVEL _
                            , _
                            , _
                            , 1
                    Else
                        lsOrigDesc = Trim(!SEC_LEVEL_DESC)
                        lsParsedDesc = ParseString(Trim(!SEC_LEVEL_DESC))
                        If (Len(lsParsedDesc) > 30) Then
                            lsParsedDesc = Left(lsParsedDesc, 30)
                        End If

                        If (lsParsedDesc <> lsOrigDesc) Then
                            'update the attribute in the record
                            !SEC_LEVEL_DESC = lsParsedDesc
                            .Update

                            AppendSecLvlInfoBRV _
                                r_adrsBRV _
                                , trErrSecurityLvlInfo_F _
                                , mstrCBUSINESSWAR _
                                , "Special characters (New Line, Semi Colon " & _
                                ", Colon, Single Quote, Double Quote etc.) were found and removed in Security Level Description." _
                                , MODULENAME & ": ValidateSecurityLevelInfo" _
                                , !SEC_LEVEL _
                                , _
                                , _
                                , 1
                        End If
                    End If

                    If Not (CheckForValidString(!DIVISION_NUM)) Then
                        AppendSecLvlInfoBRV _
                            r_adrsBRV _
                            , trErrSecurityLvlInfo_F _
                            , mstrCBUSINESSERR _
                            , "Invalid Division." _
                            , MODULENAME & ": ValidateSecurityLevelInfo" _
                            , !SEC_LEVEL _
                            , _
                            , _
                            , 1
                    End If

                    If Not (CheckForValidString(!TIME_STAMP, , , True, 26)) Then
                        AppendSecLvlInfoBRV _
                            r_adrsBRV _
                            , trErrSecurityLvlInfo_F _
                            , mstrCBUSINESSERR _
                            , "Invalid Time Stamp." _
                            , MODULENAME & ": ValidateSecurityLevelInfo" _
                            , !SEC_LEVEL _
                            , _
                            , _
                            , 1
                    End If
            End Select

            .MoveNext
        Next llRecordIndex
    End With

    r_adrsRSToProcess.Filter = adFilterNone
    r_adrsRSToProcess.MoveFirst

    ValidateSecurityLevelInfo = True

    If Not (r_adrsBRV Is Nothing) Then
        With r_adrsBRV
            If Not (.BOF And .EOF) Then
                .Filter = adFilterNone
                .Filter = MDAC25FilterFix("ERR_TYPE = 'BE' OR ERR_TYPE = 'AE' OR ERR_TYPE = 'AW'")
                If Not (.EOF) Then
                    ValidateSecurityLevelInfo = False
                End If

                .Filter = adFilterNone
            End If
        End With
    End If
    Exit Function

ValidateSecurityLevelInfo_EH:
    Set r_adrsBRV = Nothing
    ValidateSecurityLevelInfo = False

    TrsRaiseError MODULENAME _
            , "ValidateSecurityLevelInfo"

End Function
Private Function ValidateEmployeeAssignments( _
                    ByVal v_adrsRSToProcess As ADODB.Recordset _
                    , ByRef r_adrsBRV As ADODB.Recordset _
                    , ByVal vlUserEmployeeNumber As Long) As Boolean
On Error GoTo ValidateEmployeeAssignments_EH
Dim llRecordIndex As Long
Dim lobjDB As dbSecurity.cdFunSecurity
Dim ladrs As ADODB.Recordset
Dim llUserSecurityLevel As Long, llEmployeeSecurityLevel As Long

    ValidateEmployeeAssignments = False

    If Not (v_adrsRSToProcess Is Nothing) Then
        If Not (v_adrsRSToProcess.EOF And _
            v_adrsRSToProcess.BOF) Then
            v_adrsRSToProcess.Filter = adFilterNone
            v_adrsRSToProcess.MoveFirst
        Else
            Exit Function
        End If
    Else
        Exit Function
    End If

    Set lobjDB = CreateObject("dbSecurity.cdFunSecurity")
    Set ladrs = lobjDB.db_CheckEmployee(vlUserEmployeeNumber)
    Set lobjDB = Nothing

    llUserSecurityLevel = -1
    If Not (ladrs Is Nothing) Then
        If Not (ladrs.EOF And ladrs.BOF) Then
            ladrs.MoveFirst
            llUserSecurityLevel = ladrs!SEC_LEVEL
        End If
    End If
    Set ladrs = Nothing

    With v_adrsRSToProcess
        For llRecordIndex = 1 To .RecordCount
            If Not (CheckForValidLong(!SEC_LEVEL)) Then
                AppendSecLvlInfoBRV _
                    r_adrsBRV _
                    , trErrSecurityLvlInfo_F _
                    , mstrCBUSINESSERR _
                    , "Invalid Security Level." _
                    , MODULENAME & ": ValidateEmployeeAssignments" _
                    , !SEC_LEVEL _
                    , !EMP_NUM _
                    , _
                    , 2
            End If

            If Not (CheckForValidLong(!EMP_NUM, True)) Then
                AppendSecLvlInfoBRV _
                    r_adrsBRV _
                    , trErrSecurityLvlInfo_F _
                    , mstrCBUSINESSERR _
                    , "Invalid Employee assignment." _
                    , MODULENAME & ": ValidateEmployeeAssignments" _
                    , !SEC_LEVEL _
                    , !EMP_NUM _
                    , _
                    , 2
            Else
                Set lobjDB = CreateObject("dbSecurity.cdFunSecurity")
                Set ladrs = lobjDB.db_CheckEmployee(!EMP_NUM)
                Set lobjDB = Nothing

                llEmployeeSecurityLevel = -1
                If Not (ladrs Is Nothing) Then
                    If Not (ladrs.EOF And ladrs.BOF) Then
                        ladrs.MoveFirst
                        llEmployeeSecurityLevel = ladrs!SEC_LEVEL
                    Else
                        AppendSecLvlInfoBRV _
                            r_adrsBRV _
                            , trErrSecurityLvlInfo_F _
                            , mstrCBUSINESSERR _
                            , "Invalid Employee assignment. Employee " & CStr(!EMP_NUM) & " does not exist." _
                            , MODULENAME & ": ValidateEmployeeAssignments" _
                            , !SEC_LEVEL _
                            , !EMP_NUM _
                            , _
                            , 2
                    End If
                Else
                    AppendSecLvlInfoBRV _
                        r_adrsBRV _
                        , trErrSecurityLvlInfo_F _
                        , mstrCBUSINESSERR _
                        , "Invalid Employee assignment. Employee " & CStr(!EMP_NUM) & " does not exist." _
                        , MODULENAME & ": ValidateEmployeeAssignments" _
                        , !SEC_LEVEL _
                        , !EMP_NUM _
                        , _
                        , 2
                End If
            End If
            Set ladrs = Nothing

            Select Case (llEmployeeSecurityLevel)
                Case 9, 10, 11, 981 'Admin levels
                    If (llUserSecurityLevel <> 10) Then
                        AppendSecLvlInfoBRV _
                            r_adrsBRV _
                            , trErrSecurityLvlInfo_F _
                            , mstrCBUSINESSERR _
                            , "You are not authorized to modify security level for Employee " & _
                            CStr(!EMP_NUM) & ", please contact the TRIS Team." _
                            , MODULENAME & ": ValidateEmployeeAssignments" _
                            , !SEC_LEVEL _
                            , !EMP_NUM _
                            , _
                            , 2
                    End If

                Case 403 'Central Payroll Timekeepers
                    Select Case (llUserSecurityLevel)
                        Case 9, 10
                            '
                        Case Else
                            AppendSecLvlInfoBRV _
                                r_adrsBRV _
                                , trErrSecurityLvlInfo_F _
                                , mstrCBUSINESSERR _
                                , "You are not authorized to modify security level for Employee " & _
                                CStr(!EMP_NUM) & ", please contact the TRIS Team." _
                                , MODULENAME & ": ValidateEmployeeAssignments" _
                                , !SEC_LEVEL _
                                , !EMP_NUM _
                                , _
                                , 2
                    End Select
            End Select

            If (!RECORD_STATUS = "D") Then
                If Not (CheckForValidString(!TIME_STAMP, , , True, 26)) Then
                    AppendSecLvlInfoBRV _
                        r_adrsBRV _
                        , trErrSecurityLvlInfo_F _
                        , mstrCBUSINESSERR _
                        , "Invalid Time Stamp." _
                        , MODULENAME & ": ValidateEmployeeAssignments" _
                        , !SEC_LEVEL _
                        , !EMP_NUM _
                        , _
                        , 2
                End If
            End If

            .MoveNext
        Next llRecordIndex
    End With

    v_adrsRSToProcess.Filter = adFilterNone
    v_adrsRSToProcess.MoveFirst

    ValidateEmployeeAssignments = True

    If Not (r_adrsBRV Is Nothing) Then
        With r_adrsBRV
            If Not (.BOF And .EOF) Then
                .Filter = adFilterNone
                .Filter = MDAC25FilterFix("ERR_TYPE = 'BE' OR ERR_TYPE = 'AE' OR ERR_TYPE = 'AW'")
                If Not (.EOF) Then
                    ValidateEmployeeAssignments = False
                End If

                .Filter = adFilterNone
            End If
        End With
    End If
    
    Set lobjDB = Nothing
    Set ladrs = Nothing
    Exit Function
ValidateEmployeeAssignments_EH:
    Set lobjDB = Nothing

    Set r_adrsBRV = Nothing
    Set ladrs = Nothing

    ValidateEmployeeAssignments = False

    TrsRaiseError MODULENAME _
            , "ValidateEmployeeAssignments"

End Function
Private Function ValidateBusFuncAssignments( _
                    ByVal v_adrsRSToProcess As ADODB.Recordset _
                    , ByRef r_adrsBRV As ADODB.Recordset _
                    ) As Boolean
On Error GoTo ValidateBusFuncAssignments_EH
Dim llRecordIndex As Long

    ValidateBusFuncAssignments = False

    If Not (v_adrsRSToProcess Is Nothing) Then
        If Not (v_adrsRSToProcess.EOF And _
            v_adrsRSToProcess.BOF) Then
            v_adrsRSToProcess.Filter = adFilterNone
            v_adrsRSToProcess.MoveFirst
        Else
            Exit Function
        End If
    Else
        Exit Function
    End If

    With v_adrsRSToProcess
        For llRecordIndex = 1 To .RecordCount
            Select Case (!RECORD_STATUS)
                Case "D"    'DELETE
                    If Not (CheckForValidLong(!SEC_LEVEL)) Then
                        AppendSecLvlInfoBRV _
                            r_adrsBRV _
                            , trErrSecurityLvlInfo_F _
                            , mstrCBUSINESSERR _
                            , "Invalid Security Level." _
                            , MODULENAME & ": ValidateBusFuncAssignments" _
                            , !SEC_LEVEL _
                            , !EMP_NUM _
                            , Trim(!BUS_FUNC_KEY) _
                            , 3
                    End If

                    If Not (CheckForValidString(!BUS_FUNC_KEY)) Then
                        AppendSecLvlInfoBRV _
                            r_adrsBRV _
                            , trErrSecurityLvlInfo_F _
                            , mstrCBUSINESSERR _
                            , "Invalid Business Function assignment." _
                            , MODULENAME & ": ValidateBusFuncAssignments" _
                            , !SEC_LEVEL _
                            , !EMP_NUM _
                            , Trim(!BUS_FUNC_KEY) _
                            , 3
                    End If

                    If Not (CheckForValidString(!CREATED_TIMESTAMP, , , True, 26)) Then
                        AppendSecLvlInfoBRV _
                            r_adrsBRV _
                            , trErrSecurityLvlInfo_F _
                            , mstrCBUSINESSERR _
                            , "Invalid Time Stamp." _
                            , MODULENAME & ": ValidateBusFuncAssignments" _
                            , !SEC_LEVEL _
                            , !EMP_NUM _
                            , Trim(!BUS_FUNC_KEY) _
                            , 3
                    End If

                Case "I"
                    If Not (CheckForValidLong(!SEC_LEVEL)) Then
                        AppendSecLvlInfoBRV _
                            r_adrsBRV _
                            , trErrSecurityLvlInfo_F _
                            , mstrCBUSINESSERR _
                            , "Invalid Security Level." _
                            , MODULENAME & ": ValidateBusFuncAssignments" _
                            , !SEC_LEVEL _
                            , !EMP_NUM _
                            , Trim(!BUS_FUNC_KEY) _
                            , 3
                    End If

                    If Not (CheckForValidString(!BUS_FUNC_KEY)) Then
                        AppendSecLvlInfoBRV _
                            r_adrsBRV _
                            , trErrSecurityLvlInfo_F _
                            , mstrCBUSINESSERR _
                            , "Invalid Business Function assignment." _
                            , MODULENAME & ": ValidateBusFuncAssignments" _
                            , !SEC_LEVEL _
                            , !EMP_NUM _
                            , Trim(!BUS_FUNC_KEY) _
                            , 3
                    End If
            End Select

            .MoveNext
        Next llRecordIndex
    End With

    v_adrsRSToProcess.Filter = adFilterNone
    v_adrsRSToProcess.MoveFirst

    ValidateBusFuncAssignments = True

    If Not (r_adrsBRV Is Nothing) Then
        With r_adrsBRV
            If Not (.BOF And .EOF) Then
                .Filter = adFilterNone
                .Filter = MDAC25FilterFix("ERR_TYPE = 'BE' OR ERR_TYPE = 'AE' OR ERR_TYPE = 'AW'")
                If Not (.EOF) Then
                    ValidateBusFuncAssignments = False
                End If

                .Filter = adFilterNone
            End If
        End With
    End If

    Exit Function
ValidateBusFuncAssignments_EH:
    Set r_adrsBRV = Nothing
    ValidateBusFuncAssignments = False

    TrsRaiseError MODULENAME _
            , "ValidateBusFuncAssignments"

End Function
