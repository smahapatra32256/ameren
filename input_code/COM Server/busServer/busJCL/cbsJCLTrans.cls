VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "cbsJCLTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'AMEREN Project: TRIS Conversion Phase 3
'October 2004
'Author:    Prem M. Chander
'Advisor:   John R. Stricklin
'Module Description: This class will provide the user the ability
'to insert TRIS 16 bit report request information into
'TTR_JCL%% tables. It may also call TR427

Option Explicit

Private Const MODULENAME = "busJCL.cbsJCLTrans"
Private Const ADOFIELDATTR = adFldUpdatable Or adFldIsNullable

Private Const PROCEDURE = 1
Private Const PRINTER = 2
Private Const PRINTERS = 3
Private Const EMPLOYEE = 4
Private Const EMPLOYEES = 5
Private Const OPPORTUNITY = 6
Private Const QUALIFICATION = 7
Private Const SKILL = 8
Private Const LOCATION = 9
Private Const BUDGET = 10
Private Const DEPARTMENT = 11
Private Const WORKGROUP = 12
Private Const UNION = 13
Private Const JOB = 14
Private Const OTBANK = 15
Public Function SubmitReport(ByVal vsProcedureName As String _
                    , ByVal vadrsParameterList As ADODB.Recordset _
                    , ByVal vadrsPrinters As ADODB.Recordset _
                    , ByRef radrsBRV As ADODB.Recordset _
                    , ByVal vlTRISUserNumber As Long _
                    , Optional ByVal viCopies As Integer = 1 _
                    ) As Long
On Error GoTo SubmitReport_EH
Dim lobjDBJCL As dbJCL.cdJCL
Dim lbReturn As Boolean, lsFilterString As String
Dim ladrsJCLParameterList As ADODB.Recordset

    SubmitReport = -1

    Set radrsBRV = CreateBRV

    If (Trim(vsProcedureName) = "") Then
        AppendBRV radrsBRV _
            , trErrbusJCLRequest _
            , mstrCBUSINESSERR _
            , "Invalid report procedure name. Procedure name has not been provided." _
            , MODULENAME & ": SubmitReport"
        GoTo REPORTBRV
    End If

    If Not (mADOExtensions.RecordsExist(vadrsParameterList)) Then
        AppendBRV radrsBRV _
            , trErrbusJCLRequest _
            , mstrCBUSINESSERR _
            , "Invalid report paramters. Report input criteria has not been provided." _
            , MODULENAME & ": SubmitReport"
        GoTo REPORTBRV
    End If

    If Not (mADOExtensions.RecordsExist(vadrsPrinters)) Then
        AppendBRV radrsBRV _
            , trErrbusJCLRequest _
            , mstrCBUSINESSERR _
            , "Invalid report printer. Printer(s) to print the report on has not been provided." _
            , MODULENAME & ": SubmitReport"
        GoTo REPORTBRV
    End If

    'Perform Validations
    Set ladrsJCLParameterList = CreateJCLParameterList

    If Not (ValidateRequest(vsProcedureName _
            , vadrsPrinters _
            , vadrsParameterList _
            , ladrsJCLParameterList _
            , radrsBRV)) Then
        GoTo REPORTBRV
    End If

    'Call dbJCL method
    Set lobjDBJCL = CreateObject("dbJCL.cdJCL")
    SubmitReport = lobjDBJCL.dbProcessJCLRequest(vsProcedureName _
                    , ladrsJCLParameterList _
                    , vadrsPrinters _
                    , vlTRISUserNumber _
                    , viCopies)
    Set lobjDBJCL = Nothing

    Set ladrsJCLParameterList = Nothing

    GetObjectContext.SetComplete
    Exit Function
SubmitReport_EH:
    GetObjectContext.SetAbort
    Set lobjDBJCL = Nothing
    Set ladrsJCLParameterList = Nothing

    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "SubmitReport"
            Case trErrdbGetPrinters, trErrdbCheckStatus _
                , trErrdb427, trErrdbJCLRequest _
                , trErrdbInsIntoJCLRQT, trErrdbInsIntoJCLDEST _
                , trErrdbInsIntoJCLDATA, trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "ProcessOTCodeTableEntries" _
                    , Err.Number _
                    , , , , _
                    True
            Case Else
                TrsRaiseError MODULENAME _
                    , "ProcessOTCodeTableEntries" _
                    , trErrbusJCLRequest _
                    , , , , _
                    True
    End Select

    Exit Function
REPORTBRV:
    Set lobjDBJCL = Nothing
    Set ladrsJCLParameterList = Nothing

    If Not (radrsBRV Is Nothing) Then
        With radrsBRV
            If Not (.EOF And .BOF) Then
                .MoveFirst
                .Filter = adFilterNone

                'Check for errors
'MDAC25 FILTER FIX
                lsFilterString = "ERR_TYPE = 'BE'"
                .Filter = MDAC25FilterFix(lsFilterString)
                lsFilterString = ""
                'If found abort else complete
                If Not (.EOF) Then
                    'Abort Transaction
                    GetObjectContext().SetAbort
                    SubmitReport = -1
                Else
                    'Complete Transaction
                    GetObjectContext().SetComplete
                    SubmitReport = -1
                End If

                .Filter = adFilterNone
            Else
                'Complete Transaction
                GetObjectContext().SetComplete
                SubmitReport = -1
            End If
        End With
    Else
        'Complete Transaction
        GetObjectContext().SetComplete
        SubmitReport = -1
    End If

    Exit Function

End Function
Private Function CreateBRV() As ADODB.Recordset
On Error GoTo CreateBRV_EH

Dim ladrs As New ADODB.Recordset

    With ladrs
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
Private Sub AppendBRV(ByRef radrsBRV As ADODB.Recordset _
                , ByVal vlErrorNumber As Long _
                , ByVal vsErrorType As String _
                , ByVal vsErrorDescription As String _
                , ByVal vsErrorSource As String)
On Error GoTo AppendBRV_EH

    With radrsBRV
        If (.State <> adStateOpen) Then
            Exit Sub
        End If

        .AddNew

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
Private Function ValidateRequest(ByVal vsProcedureName As String _
                    , ByVal vadrsPrinters As ADODB.Recordset _
                    , ByVal vadrsParameterList As ADODB.Recordset _
                    , ByRef radrsJCLParameterList As ADODB.Recordset _
                    , ByRef radrsBRV As ADODB.Recordset _
                    ) As Boolean
On Error GoTo ValidateRequest_EH
Dim llIndex As Long, lsMultiple As String
Dim llCounter As Long, lsFilterString As String
Dim lbAddRecord As Boolean, llHoldIndex As Long
Dim lsQuals As String, lsSkills As String
Dim lbAddToParm As Boolean, ldtStartDate As Date
Dim lobjCnMgr As util_DBConnectMgr.clsuDBConnectMgr
Dim lsDB2SubSystem As String, lbCheck As Boolean

    'Printers
    'NOTE: PRINTER_ID IS NOT THE SAME AS PRINTER_ID ON THE CPF_TABLE
    If (mADOExtensions.RecordsExist(vadrsPrinters)) Then
        With vadrsPrinters
            .MoveFirst

            Set lobjCnMgr = CreateObject( _
                            "util_DBConnectMgr.clsuDBConnectMgr")
            lsDB2SubSystem = lobjCnMgr.GetDataSourceName
            Set lobjCnMgr = Nothing

            If (InStr(1, lsDB2SubSystem, "LUW") > 0) Then
                lsDB2SubSystem = Trim(Left(lsDB2SubSystem, 7))
            Else
                lsDB2SubSystem = Trim(Left(lsDB2SubSystem, 4))
            End If

            lbCheck = True
            Select Case (lsDB2SubSystem)
                Case "DB2P", "DB2Q" _
                    , "DB2PLUW", "DB2QLUW" _
                    , "DB2ALUW", "DB2BLUW"
                    '
                Case Else
                    lbCheck = False
            End Select

            Select Case (.RecordCount)
                Case 1
                    If (lbCheck) Then   'DB2P AND DB2Q ONLY
                        If Not (CheckForItemExistenceInDB(PRINTER _
                                , Trim(!PRINTER_ID))) Then
                            AppendBRV radrsBRV _
                                , trErrbusJCLRequest _
                                , mstrCBUSINESSERR _
                                , "Invalid report printer. Printer " & _
                                    Trim(!PRINTER_ID) & " does not exist." _
                                , MODULENAME & ": ValidateRequest"
                        End If
                    Else 'DO NOT CHECK PRINTER EXISTENCE FOR TEST REGION
                        'UPDATE THE PRINTER ID SO THAT IT DOES NOT CONTAIN
                        '. IN THE TEST REGIONS E.G. CAL.U88 = CALU88 IN TEST REGION
                        !PRINTER_ID = Replace(!PRINTER_ID, ".", "", 1)
                        .Update
                    End If

                Case Is > 1
                    If (lbCheck) Then
                        'Prepare printers
                        lsMultiple = ""
                        For llIndex = 1 To .RecordCount
                            lsMultiple = lsMultiple & _
                                    "'" & Trim(!PRINTER_ID) & "', "

                            .MoveNext
                        Next llIndex

                        lsMultiple = Trim(lsMultiple)
                        If (Right(lsMultiple, 1) = ",") Then
                            lsMultiple = Left(lsMultiple, Len(lsMultiple) - 1)
                        End If

                        If (Trim(lsMultiple) <> "") Then
                            If Not (CheckForItemExistenceInDB(PRINTERS _
                                    , Trim(lsMultiple), , , .RecordCount)) Then
                                AppendBRV radrsBRV _
                                    , trErrbusJCLRequest _
                                    , mstrCBUSINESSERR _
                                    , "Invalid report printers. Provided printers " & _
                                        Trim(lsMultiple) & " do not exist." _
                                    , MODULENAME & ": ValidateRequest"
                            End If
                        End If

                        lsMultiple = ""
                    Else 'DO NOT CHECK PRINTER EXISTENCE FOR TEST REGION
                        'UPDATE THE PRINTER IDS SO THAT THEY DO NOT CONTAIN
                        '. IN THE TEST REGIONS E.G. CAL.U88 = CALU88 IN TEST REGION
                        For llIndex = 1 To .RecordCount
                            !PRINTER_ID = Replace(!PRINTER_ID, ".", "", 1)
                            .Update

                            .MoveNext
                        Next llIndex
                    End If
            End Select
        End With
    End If

    'Check criteria
    'NOTE: FOR FOCUS REPORTS USE COMMA AFTER A DATA_VALUE
    Select Case (Trim(vsProcedureName))
        Case "TR33117"  'Weekly WG Schedule
            With vadrsParameterList
                .MoveFirst

                llCounter = 0
                lbAddRecord = True

                If Not (CheckForValidLong(!WG_NUM, True)) Then
                    AppendBRV radrsBRV _
                        , trErrbusJCLRequest _
                        , mstrCBUSINESSERR _
                        , "Invalid workgroup." _
                        , MODULENAME & ": ValidateRequest"

                    lbAddRecord = False
                Else
                    If Not (CheckForItemExistenceInDB(WORKGROUP, _
                            , !WG_NUM)) Then
                        AppendBRV radrsBRV _
                            , trErrbusJCLRequest _
                            , mstrCBUSINESSERR _
                            , "Invalid workgroup. Workgroup " & CStr(!WG_NUM) & _
                                " does not exist." _
                            , MODULENAME & ": ValidateRequest"
                        lbAddRecord = False
                    End If
                End If

                If Not (CheckForValidDate(!START_DATE)) Then
                    AppendBRV radrsBRV _
                        , trErrbusJCLRequest _
                        , mstrCBUSINESSERR _
                        , "Invalid start date." _
                        , MODULENAME & ": ValidateRequest"

                    lbAddRecord = False
                End If

                If Not (CheckForValidDate(!END_DATE)) Then
                    AppendBRV radrsBRV _
                        , trErrbusJCLRequest _
                        , mstrCBUSINESSERR _
                        , "Invalid end date." _
                        , MODULENAME & ": ValidateRequest"

                    lbAddRecord = False
                End If

                If (lbAddRecord) Then
                    If (!END_DATE < !START_DATE) Then
                        AppendBRV radrsBRV _
                            , trErrbusJCLRequest _
                            , mstrCBUSINESSERR _
                            , "Invalid start and end dates. Start date cannot be greater than the end date." _
                            , MODULENAME & ": ValidateRequest"

                        lbAddRecord = False
                    End If
                End If

                If (lbAddRecord) Then
                    If (DateDiff("d", !START_DATE, !END_DATE) > 6) Then
                        AppendBRV radrsBRV _
                            , trErrbusJCLRequest _
                            , mstrCBUSINESSERR _
                            , "Invalid start and end dates. Start and end dates cannot range for more than a week." _
                            , MODULENAME & ": ValidateRequest"

                        lbAddRecord = False
                    End If
                End If

                If (lbAddRecord) Then
                    With radrsJCLParameterList
                        .AddNew
                        If (llCounter = 0) Then
                            !DATA_SEQ_NUM = 0
                        Else
                            llCounter = llCounter + 1
                            !DATA_SEQ_NUM = llCounter
                        End If

                        lsMultiple = Format(vadrsParameterList!START_DATE, "mmddyy") & _
                            Format(vadrsParameterList!END_DATE, "mmddyy")

                        Select Case (Len(CStr(vadrsParameterList!WG_NUM)))
                            Case 1
                                lsMultiple = lsMultiple & _
                                        "000000" & CStr(vadrsParameterList!WG_NUM)
                            Case 2
                                lsMultiple = lsMultiple & _
                                        "00000" & CStr(vadrsParameterList!WG_NUM)
                            Case 3
                                lsMultiple = lsMultiple & _
                                        "0000" & CStr(vadrsParameterList!WG_NUM)
                            Case 4
                                lsMultiple = lsMultiple & _
                                        "000" & CStr(vadrsParameterList!WG_NUM)
                            Case 5
                                lsMultiple = lsMultiple & _
                                        "00" & CStr(vadrsParameterList!WG_NUM)
                            Case 6
                                lsMultiple = lsMultiple & _
                                        "0" & CStr(vadrsParameterList!WG_NUM)
                            Case 7
                                lsMultiple = lsMultiple & _
                                        CStr(vadrsParameterList!WG_NUM)
                        End Select

                        !DATA_VALUE = Left(lsMultiple & Space(80), 80)

                        .Update

                        lsMultiple = ""
                    End With
                End If
            End With

        Case "TR33118"  'Employee Schedule (TTR_JCL -- INSERTS GOOD)
            With vadrsParameterList
                .MoveFirst

                llCounter = 0
                lbAddRecord = True

                If Not (CheckForValidString(!MONTH_YEAR, , , True, 6)) Then
                    AppendBRV radrsBRV _
                        , trErrbusJCLRequest _
                        , mstrCBUSINESSERR _
                        , "Invalid month and year. (" & !MONTH_YEAR & ")" _
                        , MODULENAME & ": ValidateRequest"

                    lbAddRecord = False
                End If

                If Not (CheckForValidLong(!EMP_NUM, True)) Then
                    AppendBRV radrsBRV _
                        , trErrbusJCLRequest _
                        , mstrCBUSINESSERR _
                        , "Invalid employee." _
                        , MODULENAME & ": ValidateRequest"

                    lbAddRecord = False
                Else
                    If Not (CheckForItemExistenceInDB(EMPLOYEE, _
                            , !EMP_NUM)) Then
                        AppendBRV radrsBRV _
                            , trErrbusJCLRequest _
                            , mstrCBUSINESSERR _
                            , "Invalid employee. Employee " & CStr(!EMP_NUM) & _
                                " does not exist." _
                            , MODULENAME & ": ValidateRequest"
                        lbAddRecord = False
                    End If
                End If

                If (lbAddRecord) Then
                    With radrsJCLParameterList
                        .AddNew
                        If (llCounter = 0) Then
                            !DATA_SEQ_NUM = 0
                        Else
                            llCounter = llCounter + 1
                            !DATA_SEQ_NUM = llCounter
                        End If
                        !DATA_VALUE = vadrsParameterList!MONTH_YEAR & _
                                    PadLeft(CStr(vadrsParameterList!EMP_NUM), 9, "0")
                        .Update
                    End With
                End If
            End With

        Case "TR33144"  'Trade (TTR_JCL -- INSERTS GOOD)
            llCounter = 0
            lbAddRecord = True

            With vadrsParameterList
                .MoveFirst

                If Not (CheckForValidDate(!START_DATE)) Then
                    AppendBRV radrsBRV _
                        , trErrbusJCLRequest _
                        , mstrCBUSINESSERR _
                        , "Invalid start date." _
                        , MODULENAME & ": ValidateRequest"

                    lbAddRecord = False
                End If

                If Not (CheckForValidDate(!END_DATE)) Then
                    AppendBRV radrsBRV _
                        , trErrbusJCLRequest _
                        , mstrCBUSINESSERR _
                        , "Invalid end date." _
                        , MODULENAME & ": ValidateRequest"

                    lbAddRecord = False
                End If

                If (lbAddRecord) Then
                    If (!END_DATE < !START_DATE) Then
                        AppendBRV radrsBRV _
                            , trErrbusJCLRequest _
                            , mstrCBUSINESSERR _
                            , "Invalid start and end dates. Start date cannot be greater than the end date." _
                            , MODULENAME & ": ValidateRequest"

                        lbAddRecord = False
                    End If
                End If

                If (lbAddRecord) Then
                    With radrsJCLParameterList
                        .AddNew
                        If (llCounter = 0) Then
                            !DATA_SEQ_NUM = 0
                        Else
                            llCounter = llCounter + 1
                            !DATA_SEQ_NUM = llCounter
                        End If
                        !DATA_VALUE = Format(vadrsParameterList!START_DATE _
                                        , "yyyy-mm-dd") & _
                                    Format(vadrsParameterList!END_DATE _
                                        , "yyyy-mm-dd")
                        .Update
                    End With
                End If

                .Filter = MDAC25FilterFix("EMP_NUM > 0")
                If Not (.EOF) Then
                    For llIndex = 1 To .RecordCount
                        lbAddRecord = True

                        If Not (CheckForItemExistenceInDB(EMPLOYEE, , !EMP_NUM)) Then
                            AppendBRV radrsBRV _
                                , trErrbusJCLRequest _
                                , mstrCBUSINESSERR _
                                , "Invalid employee. Employee " & CStr(!EMP_NUM) & _
                                    " does not exist." _
                                , MODULENAME & ": ValidateRequest"

                            lbAddRecord = False
                        End If

                        If (lbAddRecord) Then
                            If (llIndex <= 5) Then
                                Select Case (Len(CStr(!EMP_NUM)))
                                    Case 1
                                        lsMultiple = "00000000" & _
                                            CStr(!EMP_NUM)
                                    Case 2
                                        lsMultiple = "0000000" & _
                                            CStr(!EMP_NUM)
                                    Case 3
                                        lsMultiple = "000000" & _
                                            CStr(!EMP_NUM)
                                    Case 4
                                        lsMultiple = "00000" & _
                                            CStr(!EMP_NUM)
                                    Case 5
                                        lsMultiple = "0000" & _
                                            CStr(!EMP_NUM)
                                    Case 6
                                        lsMultiple = "000" & _
                                            CStr(!EMP_NUM)
                                    Case 7
                                        lsMultiple = "00" & _
                                            CStr(!EMP_NUM)
                                    Case 8
                                        lsMultiple = "0" & _
                                            CStr(!EMP_NUM)
                                    Case 9
                                        lsMultiple = CStr(!EMP_NUM)
                                End Select

                                If (llIndex = 5) Or _
                                    (llIndex = vadrsParameterList.RecordCount) Then
                                    lsMultiple = lsMultiple & _
                                        "000000000000000000000000000000000000000000000"

                                    With radrsJCLParameterList
                                        .AddNew
                                        llCounter = llCounter + 1
                                        !DATA_SEQ_NUM = llCounter
                                        !DATA_VALUE = Left(lsMultiple, 45)
                                        .Update
                                    End With
                                    If (llIndex = vadrsParameterList.RecordCount) Then
                                        With radrsJCLParameterList
                                            .AddNew
                                            !DATA_SEQ_NUM = llCounter + 1
                                            !DATA_VALUE = "000000000000000000000000000000000000000000000"
                                            .Update
                                        End With
                                    End If
                                End If
                            Else
                                Select Case (Len(CStr(!EMP_NUM)))
                                    Case 1
                                        lsMultiple = "00000000" & _
                                            CStr(!EMP_NUM)
                                    Case 2
                                        lsMultiple = "0000000" & _
                                            CStr(!EMP_NUM)
                                    Case 3
                                        lsMultiple = "000000" & _
                                            CStr(!EMP_NUM)
                                    Case 4
                                        lsMultiple = "00000" & _
                                            CStr(!EMP_NUM)
                                    Case 5
                                        lsMultiple = "0000" & _
                                            CStr(!EMP_NUM)
                                    Case 6
                                        lsMultiple = "000" & _
                                            CStr(!EMP_NUM)
                                    Case 7
                                        lsMultiple = "00" & _
                                            CStr(!EMP_NUM)
                                    Case 8
                                        lsMultiple = "0" & _
                                            CStr(!EMP_NUM)
                                    Case 9
                                        lsMultiple = CStr(!EMP_NUM)
                                End Select

                                If (llIndex = .RecordCount) Then
                                    lsMultiple = lsMultiple & _
                                        "000000000000000000000000000000000000000000000"

                                    With radrsJCLParameterList
                                        .AddNew
                                        !DATA_SEQ_NUM = llCounter + 1
                                        !DATA_VALUE = Left(lsMultiple, 45)
                                        .Update
                                    End With
                                End If
                            End If
                        End If

                        .MoveNext
                    Next llIndex
                End If

                .Filter = adFilterNone
                lsMultiple = ""
            End With

        Case "TR33120"  'Overtime Canvass   (TTR_JCL -- INSERTS GOOD)
            With vadrsParameterList
                .MoveFirst

                llCounter = 0
                lbAddRecord = True

                If Not (CheckForValidLong(!OT_OPP_NUM, True)) Then
                    AppendBRV radrsBRV _
                        , trErrbusJCLRequest _
                        , mstrCBUSINESSERR _
                        , "Invalid opportunity." _
                        , MODULENAME & ": ValidateRequest"

                    lbAddRecord = False
                Else
                    If Not (CheckForItemExistenceInDB(OPPORTUNITY, , !OT_OPP_NUM)) Then
                        AppendBRV radrsBRV _
                            , trErrbusJCLRequest _
                            , mstrCBUSINESSERR _
                            , "Invalid opportunity. Opportunity " & _
                                CStr(!OT_OPP_NUM) & " does not exist." _
                            , MODULENAME & ": ValidateRequest"

                        lbAddRecord = False
                    End If
                End If

                If (lbAddRecord) Then
                    lsMultiple = !PRINT_PHONE & _
                                !PRINT_FORCE & _
                                !PRINT_QUAL_SKILL

                    Select Case (Len(CStr(!OT_OPP_NUM)))
                        Case "0"
                            lsMultiple = "0000000" & lsMultiple
                        Case "1"
                            lsMultiple = "000000" & _
                                        CStr(!OT_OPP_NUM) & _
                                        lsMultiple
                        Case "2"
                            lsMultiple = "00000" & _
                                        CStr(!OT_OPP_NUM) & _
                                        lsMultiple
                        Case "3"
                            lsMultiple = "0000" & _
                                        CStr(!OT_OPP_NUM) & _
                                        lsMultiple
                        Case "4"
                            lsMultiple = "000" & _
                                        CStr(!OT_OPP_NUM) & _
                                        lsMultiple
                        Case "5"
                            lsMultiple = "00" & _
                                        CStr(!OT_OPP_NUM) & _
                                        lsMultiple
                        Case "6"
                            lsMultiple = "0" & _
                                        CStr(!OT_OPP_NUM) & _
                                        lsMultiple
                        Case "7"
                            lsMultiple = CStr(!OT_OPP_NUM) & _
                                        lsMultiple
                    End Select

                    lsQuals = ""
                    .Filter = MDAC25FilterFix("QUAL_CD <> ''")
                    If Not (.EOF) Then
                        .MoveFirst

                        For llIndex = 1 To .RecordCount
                            lbAddToParm = True

                            If Not (CheckForItemExistenceInDB(QUALIFICATION, !QUAL_CD)) Then
                                AppendBRV radrsBRV _
                                    , trErrbusJCLRequest _
                                    , mstrCBUSINESSERR _
                                    , "Invalid qualification. Qualification " & _
                                        CStr(!QUAL_CD) & " does not exist." _
                                    , MODULENAME & ": ValidateRequest"

                                lbAddToParm = False
                            End If

                            If (lbAddToParm) Then
                                lsQuals = lsQuals & !QUAL_CD
                            End If

                            .MoveNext
                        Next llIndex
                    Else
                        lsQuals = Space(15)
                    End If

                    lsSkills = ""
                    .Filter = adFilterNone
                    .Filter = MDAC25FilterFix("SPEC_SKILL_CD <> ''")
                    If Not (.EOF) Then
                        .MoveFirst

                        For llIndex = 1 To .RecordCount
                            lbAddToParm = True

                            If Not (CheckForItemExistenceInDB(SKILL, !SPEC_SKILL_CD)) Then
                                AppendBRV radrsBRV _
                                    , trErrbusJCLRequest _
                                    , mstrCBUSINESSERR _
                                    , "Invalid opportunity. Opportunity " & _
                                        CStr(!SPEC_SKILL_CD) & " does not exist." _
                                    , MODULENAME & ": ValidateRequest"

                                lbAddToParm = False
                            End If

                            If (lbAddToParm) Then
                                lsSkills = lsSkills & !SPEC_SKILL_CD
                            End If

                            .MoveNext
                        Next llIndex
                    Else
                        lsSkills = Space(15)
                    End If

                    If (Trim(lsQuals) <> "") Then
                        lsQuals = lsQuals & Space(15 - Len(lsQuals))
                    End If

                    If (Trim(lsSkills) <> "") Then
                        lsSkills = lsSkills & Space(15 - Len(lsSkills))
                    End If

                    radrsJCLParameterList.AddNew
                    If (llCounter = 0) Then
                        radrsJCLParameterList!DATA_SEQ_NUM = 0
                    Else
                        llCounter = llCounter + 1
                        radrsJCLParameterList!DATA_SEQ_NUM = llCounter
                    End If
                    radrsJCLParameterList!DATA_VALUE = lsMultiple & lsQuals & lsSkills
                    radrsJCLParameterList.Update

                    .Filter = adFilterNone
                End If
            End With

            lsMultiple = ""
            lsQuals = ""
            lsSkills = ""

        Case "TR33135"  'Overtime GENCO
            With vadrsParameterList
                .MoveFirst

                llCounter = 0
                lbAddRecord = True

                If Not (CheckForValidDate(!CANVASS_DATE)) Then
                    AppendBRV radrsBRV _
                        , trErrbusJCLRequest _
                        , mstrCBUSINESSERR _
                        , "Invalid canvass date." _
                        , MODULENAME & ": ValidateRequest"

                    lbAddRecord = False
                End If

                If Not (CheckForValidLong(!LOC_CD, True)) Then
                    AppendBRV radrsBRV _
                        , trErrbusJCLRequest _
                        , mstrCBUSINESSERR _
                        , "Invalid location." _
                        , MODULENAME & ": ValidateRequest"

                    lbAddRecord = False
                Else
                    If Not (CheckForItemExistenceInDB(LOCATION, , !LOC_CD)) Then
                        AppendBRV radrsBRV _
                            , trErrbusJCLRequest _
                            , mstrCBUSINESSERR _
                            , "Invalid location. Location " & _
                                CStr(!LOC_CD) & " does not exist." _
                            , MODULENAME & ": ValidateRequest"

                        lbAddRecord = False
                    End If
                End If
            End With

            If (lbAddRecord) Then
                With radrsJCLParameterList
                    .AddNew
                    If (llCounter = 0) Then
                        !DATA_SEQ_NUM = 0
                    Else
                        llCounter = llCounter + 1
                        !DATA_SEQ_NUM = llCounter
                    End If

                    Select Case (Len(CStr(vadrsParameterList!LOC_CD)))
                        Case 0
                            !DATA_VALUE = Format(vadrsParameterList!CANVASS_DATE _
                                                    , "yyyy-mm-dd") & _
                                        "0000" & _
                                        CStr(vadrsParameterList!NUM_OF_COPIES) & _
                                        CStr(vadrsParameterList!PAGE_BREAK)
                        Case 1
                            !DATA_VALUE = Format(vadrsParameterList!CANVASS_DATE _
                                                    , "yyyy-mm-dd") & _
                                        "000" & CStr(vadrsParameterList!LOC_CD) & _
                                        CStr(vadrsParameterList!NUM_OF_COPIES) & _
                                        CStr(vadrsParameterList!PAGE_BREAK)
                        Case 2
                            !DATA_VALUE = Format(vadrsParameterList!CANVASS_DATE _
                                                    , "yyyy-mm-dd") & _
                                        "00" & CStr(vadrsParameterList!LOC_CD) & _
                                        CStr(vadrsParameterList!NUM_OF_COPIES) & _
                                        CStr(vadrsParameterList!PAGE_BREAK)
                        Case 3
                            !DATA_VALUE = Format(vadrsParameterList!CANVASS_DATE _
                                                    , "yyyy-mm-dd") & _
                                        "0" & CStr(vadrsParameterList!LOC_CD) & _
                                        CStr(vadrsParameterList!NUM_OF_COPIES) & _
                                        CStr(vadrsParameterList!PAGE_BREAK)
                        Case 4
                            !DATA_VALUE = Format(vadrsParameterList!CANVASS_DATE _
                                                    , "yyyy-mm-dd") & _
                                        CStr(vadrsParameterList!LOC_CD) & _
                                        CStr(vadrsParameterList!NUM_OF_COPIES) & _
                                        CStr(vadrsParameterList!PAGE_BREAK)
                    End Select
                    .Update

                    .AddNew
                    llCounter = llCounter + 1
                    !DATA_SEQ_NUM = llCounter
                    !DATA_VALUE = "FIN"
                    .Update
                End With
            End If

        Case "TR33121"  'Overtime After Post
            With vadrsParameterList
                .MoveFirst

                llCounter = 0
                lbAddRecord = True

                If Not (CheckForValidLong(!OT_OPP_NUM, True)) Then
                    AppendBRV radrsBRV _
                        , trErrbusJCLRequest _
                        , mstrCBUSINESSERR _
                        , "Invalid opportunity." _
                        , MODULENAME & ": ValidateRequest"

                    lbAddRecord = False
                Else
                    If Not (CheckForItemExistenceInDB(OPPORTUNITY, , !OT_OPP_NUM)) Then
                        AppendBRV radrsBRV _
                            , trErrbusJCLRequest _
                            , mstrCBUSINESSERR _
                            , "Invalid opportunity. Opportunity " & _
                                CStr(!OT_OPP_NUM) & " does not exist." _
                            , MODULENAME & ": ValidateRequest"

                        lbAddRecord = False
                    End If
                End If
            End With

            If (lbAddRecord) Then
                With radrsJCLParameterList
                    .AddNew

                    If (llCounter = 0) Then
                        !DATA_SEQ_NUM = 0
                    Else
                        llCounter = llCounter + 1
                        !DATA_SEQ_NUM = llCounter
                    End If

                    Select Case (Len(CStr(vadrsParameterList!OT_OPP_NUM)))
                        Case "0"
                            !DATA_VALUE = "0000000"
                        Case "1"
                            !DATA_VALUE = "000000" & _
                                        CStr(vadrsParameterList!OT_OPP_NUM)
                        Case "2"
                            !DATA_VALUE = "00000" & _
                                        CStr(vadrsParameterList!OT_OPP_NUM)
                        Case "3"
                            !DATA_VALUE = "0000" & _
                                        CStr(vadrsParameterList!OT_OPP_NUM)
                        Case "4"
                            !DATA_VALUE = "000" & _
                                        CStr(vadrsParameterList!OT_OPP_NUM)
                        Case "5"
                            !DATA_VALUE = "00" & _
                                        CStr(vadrsParameterList!OT_OPP_NUM)
                        Case "6"
                            !DATA_VALUE = "0" & _
                                        CStr(vadrsParameterList!OT_OPP_NUM)
                        Case "7"
                            !DATA_VALUE = CStr(vadrsParameterList!OT_OPP_NUM)
                    End Select

                    .Update
                End With
            End If

        Case "TR33125"  'Overtime Reason
            With vadrsParameterList
                .MoveFirst

                llCounter = 0
                lbAddRecord = True

                If Not (CheckForValidString(!MONTH_YEAR, , , True, 6)) Then
                    AppendBRV radrsBRV _
                        , trErrbusJCLRequest _
                        , mstrCBUSINESSERR _
                        , "Invalid month and year. (" & !MONTH_YEAR & ")" _
                        , MODULENAME & ": ValidateRequest"

                    lbAddRecord = False
                End If

                If Not (CheckForValidString(!BUDGET_NUM, True, 3)) And _
                    Not (CheckForValidString(!DEPT_NUM, True, 3)) And _
                    Not (CheckForValidLong(!WG_NUM, True)) Then
                    AppendBRV radrsBRV _
                        , trErrbusJCLRequest _
                        , mstrCBUSINESSERR _
                        , "Invalid budget, department or workgroup." _
                        , MODULENAME & ": ValidateRequest"

                    lbAddRecord = False
                Else
                    If (Trim(!BUDGET_NUM) <> "") Then
                        If Not (CheckForItemExistenceInDB(BUDGET, !BUDGET_NUM)) Then
                            AppendBRV radrsBRV _
                                , trErrbusJCLRequest _
                                , mstrCBUSINESSERR _
                                , "Invalid budget. Budget " & !BUDGET_NUM & _
                                    " does not exist." _
                                , MODULENAME & ": ValidateRequest"

                            lbAddRecord = False
                        End If
                    End If

                    If (Trim(!DEPT_NUM) <> "") Then
                        If Not (CheckForItemExistenceInDB(DEPARTMENT, !DEPT_NUM)) Then
                            AppendBRV radrsBRV _
                                , trErrbusJCLRequest _
                                , mstrCBUSINESSERR _
                                , "Invalid department. Department " & !DEPT_NUM & _
                                    " does not exist." _
                                , MODULENAME & ": ValidateRequest"

                            lbAddRecord = False
                        End If
                    End If
                End If

                If (lbAddRecord) Then
                    vadrsParameterList.MoveFirst

                    With radrsJCLParameterList
                        .AddNew
                        If (llCounter = 0) Then
                            !DATA_SEQ_NUM = 0
                        Else
                            llCounter = llCounter + 1
                            !DATA_SEQ_NUM = llCounter
                        End If

                        lsMultiple = vadrsParameterList!MONTH_YEAR

                        If (Trim(vadrsParameterList!BUDGET_NUM) = "") Then
                            lsMultiple = lsMultiple & "   "
                        Else
                            Select Case (Len(Trim(vadrsParameterList!BUDGET_NUM)))
                                Case 1
                                    lsMultiple = lsMultiple & "  " & vadrsParameterList!BUDGET_NUM
                                Case 2
                                    lsMultiple = lsMultiple & " " & vadrsParameterList!BUDGET_NUM
                                Case 3
                                    lsMultiple = lsMultiple & vadrsParameterList!BUDGET_NUM
                            End Select
                        End If

                        If (Trim(vadrsParameterList!DEPT_NUM) = "") Then
                            lsMultiple = lsMultiple & "000"
                        Else
                            Select Case (Len(Trim(vadrsParameterList!DEPT_NUM)))
                                Case 1
                                    lsMultiple = lsMultiple & "00" & vadrsParameterList!DEPT_NUM
                                Case 2
                                    lsMultiple = lsMultiple & "0" & vadrsParameterList!DEPT_NUM
                                Case 3
                                    lsMultiple = lsMultiple & vadrsParameterList!DEPT_NUM
                            End Select
                        End If

                        lsMultiple = lsMultiple & Space(80)

                        !DATA_VALUE = Left(lsMultiple, 80)
                        '!DATA_VALUE = vadrsParameterList!MONTH_YEAR & _
                                    vadrsParameterList!BUDGET_NUM & _
                                    vadrsParameterList!DEPT_NUM
                        .Update
                    End With
                End If

                lsMultiple = ""
                .Filter = MDAC25FilterFix("WG_NUM > 0")
                If Not (.EOF) Then
                    .MoveFirst

                    For llIndex = 1 To .RecordCount
                        lbAddRecord = True

                        If Not (CheckForItemExistenceInDB(WORKGROUP, , !WG_NUM)) Then
                            AppendBRV radrsBRV _
                                , trErrbusJCLRequest _
                                , mstrCBUSINESSERR _
                                , "Invalid workgroup. Workgroup " & _
                                    CStr(!WG_NUM) & " does not exist." _
                                , MODULENAME & ": ValidateRequest"

                            lbAddRecord = False
                        End If

                        If (lbAddRecord) Then
                            Select Case (Len(CStr(!WG_NUM)))
                                Case 0
                                    lsMultiple = lsMultiple & "0000000"
                                Case 1
                                    lsMultiple = lsMultiple & _
                                                "000000" & _
                                                CStr(!WG_NUM)
                                Case 2
                                    lsMultiple = lsMultiple & _
                                                "00000" & _
                                                CStr(!WG_NUM)
                                Case 3
                                    lsMultiple = lsMultiple & _
                                                "0000" & _
                                                CStr(!WG_NUM)
                                Case 4
                                    lsMultiple = lsMultiple & _
                                                "000" & _
                                                CStr(!WG_NUM)
                                Case 5
                                    lsMultiple = lsMultiple & _
                                                "00" & _
                                                CStr(!WG_NUM)
                                Case 6
                                    lsMultiple = lsMultiple & _
                                                "0" & _
                                                CStr(!WG_NUM)
                                Case 7
                                    lsMultiple = lsMultiple & CStr(!WG_NUM)
                            End Select
                        End If

                        .MoveNext
                    Next llIndex

                    lsMultiple = lsMultiple & "0000000000000000000000000000000000000000000000000000000000000000000000"

                    With radrsJCLParameterList
                        .AddNew
                        !DATA_SEQ_NUM = llCounter + 1
                        !DATA_VALUE = Left(lsMultiple, 80)
                        .Update
                    End With
                Else
                    radrsJCLParameterList.AddNew
                    radrsJCLParameterList!DATA_SEQ_NUM = llCounter + 1
                    radrsJCLParameterList!DATA_VALUE = "0000000000000000000000000000000000000000000000000000000000000000000000"
                    radrsJCLParameterList.Update
                End If

                .Filter = adFilterNone
                lsMultiple = ""
            End With

        Case "TR33122"  'Overtime Equalization
            With vadrsParameterList
                .MoveFirst

                llCounter = 0
                lbAddRecord = True

                If (CheckForValidString(!DEPT_NUM)) Then
                    If Not (CheckForItemExistenceInDB(DEPARTMENT, !DEPT_NUM)) Then
                        AppendBRV radrsBRV _
                            , trErrbusJCLRequest _
                            , mstrCBUSINESSERR _
                            , "Invalid department. Department " & !DEPT_NUM & _
                                " does not exist." _
                            , MODULENAME & ": ValidateRequest"
                        lbAddRecord = False
                    End If
                End If

                If (lbAddRecord) Then
                    radrsJCLParameterList.AddNew
                    If (llCounter = 0) Then
                        radrsJCLParameterList!DATA_SEQ_NUM = 0
                    Else
                        llCounter = llCounter + 1
                        radrsJCLParameterList!DATA_SEQ_NUM = llCounter
                    End If

                    radrsJCLParameterList!DATA_VALUE = "EX TRROCR03 DEPT=" & _
                                        !DEPT_NUM & ","
                    radrsJCLParameterList.Update
                Else
                    If (llCounter = 0) Then
                        radrsJCLParameterList!DATA_SEQ_NUM = 0
                    Else
                        llCounter = llCounter + 1
                        radrsJCLParameterList!DATA_SEQ_NUM = llCounter
                    End If

                    radrsJCLParameterList!DATA_VALUE = "EX TRROCR03 DEPT=,"
                    radrsJCLParameterList.Update
                End If

                .Filter = adFilterNone
                .Filter = MDAC25FilterFix("WG_NUM > 0")
                If Not (.EOF) Then
                    For llIndex = 1 To .RecordCount
                        If Not (CheckForItemExistenceInDB(WORKGROUP, , !WG_NUM)) Then
                            AppendBRV radrsBRV _
                                , trErrbusJCLRequest _
                                , mstrCBUSINESSERR _
                                , "Invalid workgroup. Workgroup " & CStr(!WG_NUM) & _
                                    " does not exist." _
                                , MODULENAME & ": ValidateRequest"
                            lbAddRecord = False
                        End If

                        If (lbAddRecord) Then
                            radrsJCLParameterList.AddNew
                            llCounter = llCounter + 1
                            radrsJCLParameterList!DATA_SEQ_NUM = llCounter

                            radrsJCLParameterList!DATA_VALUE = "WGNUM" & CStr(llIndex - 1) & _
                                        "=" & CStr(!WG_NUM) & ","
                            radrsJCLParameterList.Update
                        End If

                        .MoveNext
                    Next llIndex
                End If

                For llIndex = .RecordCount To 9
                    radrsJCLParameterList.AddNew
                    llCounter = llCounter + 1
                    radrsJCLParameterList!DATA_SEQ_NUM = llCounter
                    radrsJCLParameterList!DATA_VALUE = "WGNUM" & CStr(llIndex) & _
                                    "=0,"
                    radrsJCLParameterList.Update
                Next llIndex

                .Filter = adFilterNone
                .Filter = MDAC25FilterFix("JOB_NUM <> ''")
                If Not (.EOF) Then
                    For llIndex = 1 To .RecordCount
                        If Not (CheckForItemExistenceInDB(JOB, !JOB_NUM)) Then
                            AppendBRV radrsBRV _
                                , trErrbusJCLRequest _
                                , mstrCBUSINESSERR _
                                , "Invalid job. Job " & !JOB_NUM & _
                                    " does not exist." _
                                , MODULENAME & ": ValidateRequest"
                            lbAddRecord = False
                        End If

                        If (lbAddRecord) Then
                            radrsJCLParameterList.AddNew
                            llCounter = llCounter + 1
                            radrsJCLParameterList!DATA_SEQ_NUM = llCounter

                            radrsJCLParameterList!DATA_VALUE = "JOB" & CStr(llIndex) & _
                                        "=" & !JOB_NUM & ","
                            radrsJCLParameterList.Update
                        End If

                        .MoveNext
                    Next llIndex
                End If

                For llIndex = (.RecordCount + 1) To 7
                    radrsJCLParameterList.AddNew
                    llCounter = llCounter + 1
                    radrsJCLParameterList!DATA_SEQ_NUM = llCounter
                    radrsJCLParameterList!DATA_VALUE = "JOB" & CStr(llIndex) & _
                                    "=,"
                    radrsJCLParameterList.Update
                Next llIndex

                .Filter = adFilterNone

                lbAddRecord = True
                radrsJCLParameterList.AddNew
                llCounter = llCounter + 1
                radrsJCLParameterList!DATA_SEQ_NUM = llCounter
                'MAY PASS 0 TO DENOTE ALL BANKS
                Select Case (!OT_BANK_CD)
                    Case 0
                        radrsJCLParameterList!DATA_VALUE = "BANK=0,"

                    Case Is > 0
                        If (CheckForValidLong(!OT_BANK_CD, True)) Then
                            If Not (CheckForItemExistenceInDB(OTBANK, , !OT_BANK_CD)) Then
                                AppendBRV radrsBRV _
                                    , trErrbusJCLRequest _
                                    , mstrCBUSINESSERR _
                                    , "Invalid overtime bank. Overtime bank " & CStr(!OT_BANK_CD) & _
                                        " does not exist." _
                                    , MODULENAME & ": ValidateRequest"

                                lbAddRecord = False
                            End If
                        End If

                        If (lbAddRecord) Then
                            radrsJCLParameterList!DATA_VALUE = "BANK=" & CStr(!OT_BANK_CD) & ","
                        Else
                            radrsJCLParameterList!DATA_VALUE = "BANK= ,"
                        End If

                    Case Else
                        radrsJCLParameterList!DATA_VALUE = "BANK= ,"
                End Select
                radrsJCLParameterList.Update

                .Filter = adFilterNone
                .Filter = MDAC25FilterFix("SORT_BY <> ''")
                If Not (.EOF) Then
                    For llIndex = 1 To .RecordCount
                        radrsJCLParameterList.AddNew
                        llCounter = llCounter + 1
                        radrsJCLParameterList!DATA_SEQ_NUM = llCounter
                        radrsJCLParameterList!DATA_VALUE = "BY_PARM" & CStr(llIndex) & _
                                    "=" & Trim(!SORT_BY) & ","
                        radrsJCLParameterList.Update

                        .MoveNext
                    Next llIndex
                End If

                For llIndex = (.RecordCount + 1) To 17
                    radrsJCLParameterList.AddNew
                    llCounter = llCounter + 1
                    radrsJCLParameterList!DATA_SEQ_NUM = llCounter
                    radrsJCLParameterList!DATA_VALUE = "BY_PARM" & CStr(llIndex) & _
                                    "=-*,"
                    radrsJCLParameterList.Update
                Next llIndex

                radrsJCLParameterList.AddNew
                llCounter = llCounter + 1
                radrsJCLParameterList!DATA_SEQ_NUM = llCounter
                radrsJCLParameterList!DATA_VALUE = "FIN"
                radrsJCLParameterList.Update

                .Filter = adFilterNone
            End With

        Case "TR33128"  'Overtime Bank Adjustment
            With vadrsParameterList
                .MoveFirst

                lbAddRecord = True
                llCounter = 0

                If Not (CheckForValidDate(!START_DATE)) Then
                    AppendBRV radrsBRV _
                        , trErrbusJCLRequest _
                        , mstrCBUSINESSERR _
                        , "Invalid start date." _
                        , MODULENAME & ": ValidateRequest"

                    lbAddRecord = False
                End If

                If Not (CheckForValidDate(!END_DATE)) Then
                    AppendBRV radrsBRV _
                        , trErrbusJCLRequest _
                        , mstrCBUSINESSERR _
                        , "Invalid end date." _
                        , MODULENAME & ": ValidateRequest"

                    lbAddRecord = False
                End If

                If (lbAddRecord) Then
                    If (!END_DATE < !START_DATE) Then
                        AppendBRV radrsBRV _
                            , trErrbusJCLRequest _
                            , mstrCBUSINESSERR _
                            , "Invalid start and end dates. Start date cannot be greater than the end date." _
                            , MODULENAME & ": ValidateRequest"

                        lbAddRecord = False
                    End If
                End If

                If (lbAddRecord) Then
                    radrsJCLParameterList.AddNew
                    If (llCounter = 0) Then
                        radrsJCLParameterList!DATA_SEQ_NUM = 0
                    Else
                        llCounter = llCounter + 1
                        radrsJCLParameterList!DATA_SEQ_NUM = llCounter
                    End If
                    radrsJCLParameterList!DATA_VALUE = "EX TRROTBKA  SDATE=" & _
                                        Format(!START_DATE, "yyyymmdd") & ","
                    radrsJCLParameterList.Update

                    radrsJCLParameterList.AddNew
                    llCounter = llCounter + 1
                    radrsJCLParameterList!DATA_SEQ_NUM = llCounter
                    radrsJCLParameterList!DATA_VALUE = "EDATE=" & _
                                        Format(!END_DATE, "yyyymmdd") & ","
                    radrsJCLParameterList.Update
                End If

                lbAddRecord = True
                If (CheckForValidLong(!EMP_NUM)) Then
                    If Not (CheckForItemExistenceInDB(EMPLOYEE, , !EMP_NUM)) Then
                        AppendBRV radrsBRV _
                            , trErrbusJCLRequest _
                            , mstrCBUSINESSERR _
                            , "Invalid employee. Employee " & CStr(!EMP_NUM) & _
                                " does not exist." _
                            , MODULENAME & ": ValidateRequest"

                        lbAddRecord = False
                    End If
                End If

                If (lbAddRecord) Then
                    radrsJCLParameterList.AddNew
                    llCounter = llCounter + 1
                    radrsJCLParameterList!DATA_SEQ_NUM = llCounter
                    radrsJCLParameterList!DATA_VALUE = "EMPNUM=" & CStr(!EMP_NUM) & ","
                    radrsJCLParameterList.Update
                End If

                lbAddRecord = True
                'MAY PASS 0 TO DENOTE ALL BANKS
                If (!OT_BANK_CD > 0) Then
                    If (CheckForValidLong(!OT_BANK_CD, True)) Then
                        If Not (CheckForItemExistenceInDB(OTBANK, , !OT_BANK_CD)) Then
                            AppendBRV radrsBRV _
                                , trErrbusJCLRequest _
                                , mstrCBUSINESSERR _
                                , "Invalid overtime bank. Overtime bank " & CStr(!OT_BANK_CD) & _
                                    " does not exist." _
                                , MODULENAME & ": ValidateRequest"

                            lbAddRecord = False
                        End If
                    End If
                End If

                If (lbAddRecord) Then
                    radrsJCLParameterList.AddNew
                    llCounter = llCounter + 1
                    radrsJCLParameterList!DATA_SEQ_NUM = llCounter
                    radrsJCLParameterList!DATA_VALUE = "BTYPE=" & CStr(!OT_BANK_CD) & ","
                    radrsJCLParameterList.Update
                End If

                lbAddRecord = True
                If (CheckForValidLong(!OT_OPP_NUM)) Then
                    If Not (CheckForItemExistenceInDB(OPPORTUNITY, , !OT_OPP_NUM)) Then
                        AppendBRV radrsBRV _
                            , trErrbusJCLRequest _
                            , mstrCBUSINESSERR _
                            , "Invalid overtime opportunity. Opportunity " & CStr(!OT_OPP_NUM) & _
                                " does not exist." _
                            , MODULENAME & ": ValidateRequest"

                        lbAddRecord = False
                    End If
                End If

                If (lbAddRecord) Then
                    radrsJCLParameterList.AddNew
                    llCounter = llCounter + 1
                    radrsJCLParameterList!DATA_SEQ_NUM = llCounter
                    radrsJCLParameterList!DATA_VALUE = "OPPNUM=" & CStr(!OT_OPP_NUM) & ","
                    radrsJCLParameterList.Update
                End If

                radrsJCLParameterList.AddNew
                llCounter = llCounter + 1
                radrsJCLParameterList!DATA_SEQ_NUM = llCounter
                If (!OT_BANK_CD = 0) Then
                    radrsJCLParameterList!DATA_VALUE = "COMP1=GE,"
                Else
                    radrsJCLParameterList!DATA_VALUE = "COMP1=EQ,"
                End If
                radrsJCLParameterList.Update

                radrsJCLParameterList.AddNew
                llCounter = llCounter + 1
                radrsJCLParameterList!DATA_SEQ_NUM = llCounter
                If (!OT_OPP_NUM = 0) Then
                    radrsJCLParameterList!DATA_VALUE = "COMP2=GE,"
                Else
                    radrsJCLParameterList!DATA_VALUE = "COMP2=EQ,"
                End If
                radrsJCLParameterList.Update

                radrsJCLParameterList.AddNew
                llCounter = llCounter + 1
                radrsJCLParameterList!DATA_SEQ_NUM = llCounter
                radrsJCLParameterList!DATA_VALUE = "FIN"
                radrsJCLParameterList.Update
            End With

        Case "TR33132"  'Overtime Pre - Overtime
            With vadrsParameterList
                .MoveFirst

                llCounter = 0
                lbAddRecord = True

                If Not (CheckForValidDate(!OT_DATE)) Then
                    AppendBRV radrsBRV _
                        , trErrbusJCLRequest _
                        , mstrCBUSINESSERR _
                        , "Invalid overtime standing date." _
                        , MODULENAME & ": ValidateRequest"

                    lbAddRecord = False
                End If

                If (CheckForValidString(!DEPT_NUM, True, 3)) Then
                    If Not (CheckForItemExistenceInDB(DEPARTMENT, !DEPT_NUM)) Then
                        AppendBRV radrsBRV _
                            , trErrbusJCLRequest _
                            , mstrCBUSINESSERR _
                            , "Invalid department. Department " & !DEPT_NUM & _
                                " does not exist." _
                            , MODULENAME & ": ValidateRequest"

                        lbAddRecord = False
                    End If
                End If

                If (lbAddRecord) Then
                    With radrsJCLParameterList
                        .AddNew
                        If (llCounter = 0) Then
                            !DATA_SEQ_NUM = 0
                        Else
                            llCounter = llCounter + 1
                            !DATA_SEQ_NUM = llCounter
                        End If
                        !DATA_VALUE = "EX TRPREOPP OTDATE=" & _
                                        Format(vadrsParameterList!OT_DATE, "mmddyyyy") & ","
                        .Update

                        .AddNew
                        llCounter = llCounter + 1
                        !DATA_SEQ_NUM = llCounter
                        !DATA_VALUE = "PHONE=" & vadrsParameterList!PRINT_PHONE & ","
                        .Update

                        .AddNew
                        llCounter = llCounter + 1
                        !DATA_SEQ_NUM = llCounter
                        !DATA_VALUE = "QUALSK=" & vadrsParameterList!PRINT_QUALS_SKILLS & ","
                        .Update

                        .AddNew
                        llCounter = llCounter + 1
                        !DATA_SEQ_NUM = llCounter
                        !DATA_VALUE = "DEPT=" & vadrsParameterList!DEPT_NUM & ","
                        .Update
                    End With
                End If

                .Filter = MDAC25FilterFix("WG_NUM > 0")
                If Not (.EOF) Then
                    .MoveFirst

                    For llIndex = 1 To .RecordCount
                        lbAddRecord = True

                        If Not (CheckForItemExistenceInDB(WORKGROUP, , !WG_NUM)) Then
                            AppendBRV radrsBRV _
                                , trErrbusJCLRequest _
                                , mstrCBUSINESSERR _
                                , "Invalid workgroup. Workgroup " & CStr(!WG_NUM) & _
                                    " does not exist." _
                                , MODULENAME & ": ValidateRequest"

                            lbAddRecord = False
                        End If

                        If (lbAddRecord) Then
                            radrsJCLParameterList.AddNew
                            llCounter = llCounter + 1
                            radrsJCLParameterList!DATA_SEQ_NUM = llCounter
                            radrsJCLParameterList!DATA_VALUE = "WGNUM" & CStr(llIndex - 1) & _
                                            "=" & CStr(!WG_NUM) & ","
                            radrsJCLParameterList.Update
                        End If

                        .MoveNext
                    Next llIndex
                End If

                For llIndex = (.RecordCount) To 9
                    radrsJCLParameterList.AddNew
                    llCounter = llCounter + 1
                    radrsJCLParameterList!DATA_SEQ_NUM = llCounter
                    radrsJCLParameterList!DATA_VALUE = "WGNUM" & CStr(llIndex) & _
                                    "=0000000" & ","
                    radrsJCLParameterList.Update
                Next llIndex

                .Filter = adFilterNone

                .Filter = MDAC25FilterFix("SORT_BY <> ''")
                If Not (.EOF) Then
                    .MoveFirst

                    For llIndex = 1 To .RecordCount
                        lbAddRecord = True

                        radrsJCLParameterList.AddNew
                        llCounter = llCounter + 1
                        radrsJCLParameterList!DATA_SEQ_NUM = llCounter
                        radrsJCLParameterList!DATA_VALUE = "BY_PARM" & CStr(llIndex) & _
                                        "=" & !SORT_BY & ","
                        radrsJCLParameterList.Update

                        .MoveNext
                    Next llIndex

                    llHoldIndex = .RecordCount + 1

                    For llIndex = llHoldIndex To 11
                        radrsJCLParameterList.AddNew
                        llCounter = llCounter + 1
                        radrsJCLParameterList!DATA_SEQ_NUM = llCounter
                        radrsJCLParameterList!DATA_VALUE = "BY_PARM" & CStr(llIndex) & _
                                        "= -*,"
                        radrsJCLParameterList.Update
                    Next llIndex
                End If

                .Filter = adFilterNone
            End With

            If (lbAddRecord) Then
                radrsJCLParameterList.AddNew
                llCounter = llCounter + 1
                radrsJCLParameterList!DATA_SEQ_NUM = llCounter
                radrsJCLParameterList!DATA_VALUE = "FIN"
                radrsJCLParameterList.Update
            End If

        Case "TR33145"  'Vacation Summary
            With vadrsParameterList
                .MoveFirst

                llCounter = 0
                lbAddRecord = True
                
                .Filter = MDAC25FilterFix("DEPT_NUM <> ''")
                If Not (.EOF) Then
                    For llIndex = 1 To .RecordCount
                        If (CheckForValidString(!DEPT_NUM)) Then
                            If Not (CheckForItemExistenceInDB(DEPARTMENT, !DEPT_NUM)) Then
                                AppendBRV radrsBRV _
                                    , trErrbusJCLRequest _
                                    , mstrCBUSINESSERR _
                                    , "Invalid department. Department " & !DEPT_NUM & _
                                        " does not exist." _
                                    , MODULENAME & ": ValidateRequest"
                                lbAddRecord = False
                            End If
                        End If

                        If (lbAddRecord) And _
                            (llIndex = 1) Then
                            radrsJCLParameterList.AddNew
                            If (llCounter = 0) Then
                                radrsJCLParameterList!DATA_SEQ_NUM = 0
                            Else
                                llCounter = llCounter + 1
                                radrsJCLParameterList!DATA_SEQ_NUM = llCounter
                            End If

                            radrsJCLParameterList!DATA_VALUE = "EX TRVACSUM DEPT_NUM1=" & _
                                                !DEPT_NUM & ","
                            radrsJCLParameterList.Update
                        Else
                            If (lbAddRecord) Then
                                radrsJCLParameterList.AddNew
                                llCounter = llCounter + 1
                                radrsJCLParameterList!DATA_SEQ_NUM = llCounter
                                radrsJCLParameterList!DATA_VALUE = "DEPT_NUM" & CStr(llIndex) & _
                                                "=" & !DEPT_NUM & ","
                                radrsJCLParameterList.Update
                            End If
                        End If

                        .MoveNext
                    Next llIndex

                    For llIndex = (.RecordCount + 1) To 4
                        radrsJCLParameterList.AddNew
                        llCounter = llCounter + 1
                        radrsJCLParameterList!DATA_SEQ_NUM = llCounter
                        radrsJCLParameterList!DATA_VALUE = "DEPT_NUM" & CStr(llIndex) & _
                                        "=,"
                        radrsJCLParameterList.Update
                    Next llIndex
                Else
                    radrsJCLParameterList.AddNew
                    If (llCounter = 0) Then
                        radrsJCLParameterList!DATA_SEQ_NUM = 0
                    Else
                        llCounter = llCounter + 1
                        radrsJCLParameterList!DATA_SEQ_NUM = llCounter
                    End If

                    radrsJCLParameterList!DATA_VALUE = "EX TRVACSUM DEPT_NUM1=,DEPT_NUM2=," & _
                                        "DEPT_NUM3=,DEPT_NUM4=,"
                    radrsJCLParameterList.Update
                End If

                .Filter = adFilterNone

                radrsJCLParameterList.AddNew
                llCounter = llCounter + 1
                radrsJCLParameterList!DATA_SEQ_NUM = llCounter
                radrsJCLParameterList!DATA_VALUE = "HEADER=" & !PRINT_HEADER & ","
                radrsJCLParameterList.Update

                .Filter = adFilterNone

                .Filter = MDAC25FilterFix("LOC_CD > 0")
                If Not (.EOF) Then
                    For llIndex = 1 To .RecordCount
                        If Not (CheckForItemExistenceInDB(LOCATION, , , !LOC_CD)) Then
                            AppendBRV radrsBRV _
                                , trErrbusJCLRequest _
                                , mstrCBUSINESSERR _
                                , "Invalid location. Location " & CStr(!LOC_CD) & _
                                    " does not exist." _
                                , MODULENAME & ": ValidateRequest"
                            lbAddRecord = False
                        End If

                        If (lbAddRecord) Then
                            radrsJCLParameterList.AddNew
                            llCounter = llCounter + 1
                            radrsJCLParameterList!DATA_SEQ_NUM = llCounter

                            radrsJCLParameterList!DATA_VALUE = "LOC_CD" & CStr(llIndex) & _
                                        "=" & CStr(!LOC_CD) & ","
                            radrsJCLParameterList.Update
                        End If

                        .MoveNext
                    Next llIndex
                End If

                For llIndex = (.RecordCount + 1) To 3
                    radrsJCLParameterList.AddNew
                    llCounter = llCounter + 1
                    radrsJCLParameterList!DATA_SEQ_NUM = llCounter
                    radrsJCLParameterList!DATA_VALUE = "LOC_CD" & CStr(llIndex) & _
                                    "=0,"
                    radrsJCLParameterList.Update
                Next llIndex

                .Filter = adFilterNone
                .Filter = MDAC25FilterFix("EMP_NUM > 0")
                If Not (.EOF) Then
                    For llIndex = 1 To .RecordCount
                        If Not (CheckForItemExistenceInDB(EMPLOYEE, , !EMP_NUM)) Then
                            AppendBRV radrsBRV _
                                , trErrbusJCLRequest _
                                , mstrCBUSINESSERR _
                                , "Invalid employee. Employee " & CStr(!EMP_NUM) & _
                                    " does not exist." _
                                , MODULENAME & ": ValidateRequest"
                            lbAddRecord = False
                        End If

                        If (lbAddRecord) Then
                            radrsJCLParameterList.AddNew
                            llCounter = llCounter + 1
                            radrsJCLParameterList!DATA_SEQ_NUM = llCounter

                            radrsJCLParameterList!DATA_VALUE = "EMP_NUM" & CStr(llIndex) & _
                                        "=" & CStr(!EMP_NUM) & ","
                            radrsJCLParameterList.Update
                        End If

                        .MoveNext
                    Next llIndex
                End If

                For llIndex = (.RecordCount + 1) To 15
                    radrsJCLParameterList.AddNew
                    llCounter = llCounter + 1
                    radrsJCLParameterList!DATA_SEQ_NUM = llCounter
                    radrsJCLParameterList!DATA_VALUE = "EMP_NUM" & CStr(llIndex) & _
                                    "=0,"
                    radrsJCLParameterList.Update
                Next llIndex

                .Filter = adFilterNone
                .Filter = MDAC25FilterFix("WG_NUM > 0")
                If Not (.EOF) Then
                    For llIndex = 1 To .RecordCount
                        If Not (CheckForItemExistenceInDB(WORKGROUP, , !WG_NUM)) Then
                            AppendBRV radrsBRV _
                                , trErrbusJCLRequest _
                                , mstrCBUSINESSERR _
                                , "Invalid workgroup. Workgroup " & CStr(!WG_NUM) & _
                                    " does not exist." _
                                , MODULENAME & ": ValidateRequest"
                            lbAddRecord = False
                        End If

                        If (lbAddRecord) Then
                            radrsJCLParameterList.AddNew
                            llCounter = llCounter + 1
                            radrsJCLParameterList!DATA_SEQ_NUM = llCounter

                            radrsJCLParameterList!DATA_VALUE = "WG_NUM" & CStr(llIndex) & _
                                        "=" & CStr(!WG_NUM) & ","
                            radrsJCLParameterList.Update
                        End If

                        .MoveNext
                    Next llIndex
                End If

                For llIndex = (.RecordCount + 1) To 20
                    radrsJCLParameterList.AddNew
                    llCounter = llCounter + 1
                    radrsJCLParameterList!DATA_SEQ_NUM = llCounter
                    radrsJCLParameterList!DATA_VALUE = "WG_NUM" & CStr(llIndex) & _
                                    "=0,"
                    radrsJCLParameterList.Update
                Next llIndex

                .Filter = adFilterNone
                .Filter = MDAC25FilterFix("JOB_NUM <> ''")
                If Not (.EOF) Then
                    For llIndex = 1 To .RecordCount
                        If Not (CheckForItemExistenceInDB(JOB, !JOB_NUM)) Then
                            AppendBRV radrsBRV _
                                , trErrbusJCLRequest _
                                , mstrCBUSINESSERR _
                                , "Invalid job. Job " & !JOB_NUM & _
                                    " does not exist." _
                                , MODULENAME & ": ValidateRequest"
                            lbAddRecord = False
                        End If

                        If (lbAddRecord) Then
                            radrsJCLParameterList.AddNew
                            llCounter = llCounter + 1
                            radrsJCLParameterList!DATA_SEQ_NUM = llCounter

                            radrsJCLParameterList!DATA_VALUE = "JOB_NUM" & CStr(llIndex) & _
                                        "=" & !JOB_NUM & ","
                            radrsJCLParameterList.Update
                        End If

                        .MoveNext
                    Next llIndex
                End If

                For llIndex = (.RecordCount + 1) To 10
                    radrsJCLParameterList.AddNew
                    llCounter = llCounter + 1
                    radrsJCLParameterList!DATA_SEQ_NUM = llCounter
                    radrsJCLParameterList!DATA_VALUE = "JOB_NUM" & CStr(llIndex) & _
                                    "=,"
                    radrsJCLParameterList.Update
                Next llIndex

                .Filter = adFilterNone

                radrsJCLParameterList.AddNew
                llCounter = llCounter + 1
                radrsJCLParameterList!DATA_SEQ_NUM = llCounter
                radrsJCLParameterList!DATA_VALUE = "TYPE=" & !REPORT_TYPE & ","
                radrsJCLParameterList.Update

                ldtStartDate = DateSerial(!Year, 1, 1)
                If (IsDate(ldtStartDate)) Then
                    If (ldtStartDate < #1/1/1994#) Then
                        ldtStartDate = DateSerial(Year(Date), 1, 1)
                    End If
                End If

                Select Case (Weekday(ldtStartDate))
                    Case vbSunday
                    Case vbMonday
                        ldtStartDate = DateAdd("d", -1, ldtStartDate)
                    Case vbTuesday
                        ldtStartDate = DateAdd("d", -2, ldtStartDate)
                    Case vbWednesday
                        ldtStartDate = DateAdd("d", -3, ldtStartDate)
                    Case vbThursday
                        ldtStartDate = DateAdd("d", -4, ldtStartDate)
                    Case vbFriday
                        ldtStartDate = DateAdd("d", -5, ldtStartDate)
                    Case vbSaturday
                        ldtStartDate = DateAdd("d", -6, ldtStartDate)
                End Select

                radrsJCLParameterList.AddNew
                llCounter = llCounter + 1
                radrsJCLParameterList!DATA_SEQ_NUM = llCounter
                radrsJCLParameterList!DATA_VALUE = "BEGDTE=" & _
                                Format(ldtStartDate, "yyyymmdd") & ","
                radrsJCLParameterList.Update

                radrsJCLParameterList.AddNew
                llCounter = llCounter + 1
                radrsJCLParameterList!DATA_SEQ_NUM = llCounter
                radrsJCLParameterList!DATA_VALUE = "YEAR=" & _
                                CStr(!Year) & ","
                radrsJCLParameterList.Update

                .Filter = adFilterNone
                .Filter = MDAC25FilterFix("SORT_BY <> ''")
                If Not (.EOF) Then
                    For llIndex = 1 To .RecordCount
                        radrsJCLParameterList.AddNew
                        llCounter = llCounter + 1
                        radrsJCLParameterList!DATA_SEQ_NUM = llCounter
                        radrsJCLParameterList!DATA_VALUE = "BYPARM" & CStr(llIndex) & _
                                    "=" & Trim(!SORT_BY) & ","
                        radrsJCLParameterList.Update

                        .MoveNext
                    Next llIndex
                End If

                For llIndex = (.RecordCount + 1) To 10
                    radrsJCLParameterList.AddNew
                    llCounter = llCounter + 1
                    radrsJCLParameterList!DATA_SEQ_NUM = llCounter
                    radrsJCLParameterList!DATA_VALUE = "BYPARM" & CStr(llIndex) & _
                                    "=,"
                    radrsJCLParameterList.Update
                Next llIndex

                radrsJCLParameterList.AddNew
                llCounter = llCounter + 1
                radrsJCLParameterList!DATA_SEQ_NUM = llCounter
                radrsJCLParameterList!DATA_VALUE = "FIN"
                radrsJCLParameterList.Update

                .Filter = adFilterNone
            End With

        Case "TR33142"  'Seniority
            'Max Limit: 3 locs, 10 depts, 10 WGs, 5 jobs, 3 unions, 4 sort items
            With vadrsParameterList
                .MoveFirst

                llCounter = 0
                lsMultiple = ""

                .Filter = MDAC25FilterFix("DEPT_NUM <> ''")
                If Not (.EOF) Then
                    For llIndex = 1 To .RecordCount
                        Select Case (Len(Trim(!DEPT_NUM)))
                            Case 0
                                lsMultiple = lsMultiple & "000"
                            Case 1
                                lsMultiple = lsMultiple & _
                                                "00" & _
                                                Trim(!DEPT_NUM)
                            Case 2
                                lsMultiple = lsMultiple & _
                                                "0" & _
                                                Trim(!DEPT_NUM)
                            Case 3
                                lsMultiple = lsMultiple & _
                                                Trim(!DEPT_NUM)
                        End Select

                        .MoveNext
                    Next llIndex
                Else
                    lsMultiple = lsMultiple & _
                        "00000000000000000000000000000000000000000000000000000000000000000000000000000000"
                End If

                lsMultiple = Left(lsMultiple & _
                    "00000000000000000000000000000000000000000000000000000000000000000000000000000000" _
                    , 80)

                radrsJCLParameterList.AddNew
                If (llCounter = 0) Then
                    radrsJCLParameterList!DATA_SEQ_NUM = 0
                Else
                    llCounter = llCounter + 1
                    radrsJCLParameterList!DATA_SEQ_NUM = llCounter
                End If
                radrsJCLParameterList!DATA_VALUE = lsMultiple
                radrsJCLParameterList.Update

                .Filter = adFilterNone
                lsMultiple = ""
                .Filter = MDAC25FilterFix("LOC_CD > 0")
                If Not (.EOF) Then
                    For llIndex = 1 To .RecordCount
                        Select Case (Len(CStr(!LOC_CD)))
                            Case 0
                                lsMultiple = lsMultiple & "000"
                            Case 1
                                lsMultiple = lsMultiple & _
                                                "00" & _
                                                CStr(!LOC_CD)
                            Case 2
                                lsMultiple = lsMultiple & _
                                                "0" & _
                                                CStr(!LOC_CD)
                            Case 3, 4, 5, 6, 7
                                lsMultiple = lsMultiple & _
                                                CStr(!LOC_CD)
                        End Select

                        .MoveNext
                    Next llIndex
                Else
                    lsMultiple = lsMultiple & _
                        "00000000000000000000000000000000000000000000000000000000000000000000000000000000"
                End If

                lsMultiple = Left(lsMultiple & _
                    "00000000000000000000000000000000000000000000000000000000000000000000000000000000" _
                    , 80)

                radrsJCLParameterList.AddNew
                llCounter = llCounter + 1
                radrsJCLParameterList!DATA_SEQ_NUM = llCounter
                radrsJCLParameterList!DATA_VALUE = lsMultiple
                radrsJCLParameterList.Update

                .Filter = adFilterNone
                lsMultiple = ""
                .Filter = MDAC25FilterFix("UNION_CD <> ''")
                If Not (.EOF) Then
                    For llIndex = 1 To .RecordCount
                        Select Case (Len(Trim(!UNION_CD)))
                            Case 0
                                lsMultiple = lsMultiple & "00"
                            Case 1
                                lsMultiple = lsMultiple & _
                                                "0" & _
                                                Trim(!UNION_CD)
                            Case 2
                                lsMultiple = lsMultiple & !UNION_CD
                        End Select

                        .MoveNext
                    Next llIndex
                Else
                    lsMultiple = lsMultiple & _
                        "00000000000000000000000000000000000000000000000000000000000000000000000000000000"
                End If

                lsMultiple = Left(lsMultiple & _
                    "00000000000000000000000000000000000000000000000000000000000000000000000000000000" _
                    , 80)

                radrsJCLParameterList.AddNew
                llCounter = llCounter + 1
                radrsJCLParameterList!DATA_SEQ_NUM = llCounter
                radrsJCLParameterList!DATA_VALUE = lsMultiple
                radrsJCLParameterList.Update


                .Filter = adFilterNone
                lsMultiple = ""
                .Filter = MDAC25FilterFix("JOB_NUM <> ''")
                If Not (.EOF) Then
                    For llIndex = 1 To .RecordCount
                        Select Case (Len(Trim(!JOB_NUM)))
                            Case 0
                                lsMultiple = lsMultiple & "000000"
                            Case 1
                                lsMultiple = lsMultiple & _
                                                "00000" & _
                                                Trim(!JOB_NUM)
                            Case 2
                                lsMultiple = lsMultiple & _
                                                "0000" & _
                                                Trim(!JOB_NUM)
                            Case 3
                                lsMultiple = lsMultiple & _
                                                "000" & _
                                                Trim(!JOB_NUM)
                            Case 4
                                lsMultiple = lsMultiple & _
                                                "00" & _
                                                Trim(!JOB_NUM)
                            Case 5
                                lsMultiple = lsMultiple & _
                                                "0" & _
                                                Trim(!JOB_NUM)
                            Case 6
                                lsMultiple = lsMultiple & _
                                                !JOB_NUM
                        End Select

                        .MoveNext
                    Next llIndex
                Else
                    lsMultiple = lsMultiple & _
                        "00000000000000000000000000000000000000000000000000000000000000000000000000000000"
                End If

                lsMultiple = Left(lsMultiple & _
                    "00000000000000000000000000000000000000000000000000000000000000000000000000000000" _
                    , 80)

                radrsJCLParameterList.AddNew
                llCounter = llCounter + 1
                radrsJCLParameterList!DATA_SEQ_NUM = llCounter
                radrsJCLParameterList!DATA_VALUE = lsMultiple
                radrsJCLParameterList.Update


                .Filter = adFilterNone
                lsMultiple = ""
                lbAddRecord = True
                .Filter = MDAC25FilterFix("WG_NUM > 0")
                If Not (.EOF) Then
                    For llIndex = 1 To .RecordCount
                        If (llIndex = 6) Then
                            lsMultiple = Left(lsMultiple & _
                                "00000000000000000000000000000000000000000000000000000000000000000000000000000000" _
                                , 80)

                            radrsJCLParameterList.AddNew
                            llCounter = llCounter + 1
                            radrsJCLParameterList!DATA_SEQ_NUM = llCounter
                            radrsJCLParameterList!DATA_VALUE = lsMultiple
                            radrsJCLParameterList.Update

                            lsMultiple = ""
                            lbAddRecord = False
                        End If

                        Select Case (Len(CStr(!WG_NUM)))
                            Case 0
                                lsMultiple = lsMultiple & "000000000"
                            Case 1
                                lsMultiple = lsMultiple & _
                                                "00000000" & _
                                                CStr(!WG_NUM)
                            Case 2
                                lsMultiple = lsMultiple & _
                                                "0000000" & _
                                                CStr(!WG_NUM)
                            Case 3
                                lsMultiple = lsMultiple & _
                                                "000000" & _
                                                CStr(!WG_NUM)
                            Case 4
                                lsMultiple = lsMultiple & _
                                                "00000" & _
                                                CStr(!WG_NUM)
                            Case 5
                                lsMultiple = lsMultiple & _
                                                "0000" & _
                                                CStr(!WG_NUM)
                            Case 6
                                lsMultiple = lsMultiple & _
                                                "000" & _
                                                CStr(!WG_NUM)
                            Case 7
                                lsMultiple = lsMultiple & _
                                                "00" & _
                                                CStr(!WG_NUM)
                            Case 8
                                lsMultiple = lsMultiple & _
                                                "0" & _
                                                CStr(!WG_NUM)
                            Case 9
                                lsMultiple = lsMultiple & _
                                                CStr(!WG_NUM)
                        End Select

                        .MoveNext
                    Next llIndex
                Else
                    lsMultiple = lsMultiple & _
                        "00000000000000000000000000000000000000000000000000000000000000000000000000000000"
                End If

                If (lbAddRecord) Then
                    lsMultiple = Left(lsMultiple & _
                        "00000000000000000000000000000000000000000000000000000000000000000000000000000000" _
                        , 80)

                    radrsJCLParameterList.AddNew
                    llCounter = llCounter + 1
                    radrsJCLParameterList!DATA_SEQ_NUM = llCounter
                    radrsJCLParameterList!DATA_VALUE = lsMultiple
                    radrsJCLParameterList.Update

                    lsMultiple = Left(lsMultiple & _
                        "00000000000000000000000000000000000000000000000000000000000000000000000000000000" _
                        , 80)

                    radrsJCLParameterList.AddNew
                    llCounter = llCounter + 1
                    radrsJCLParameterList!DATA_SEQ_NUM = llCounter
                    radrsJCLParameterList!DATA_VALUE = lsMultiple
                    radrsJCLParameterList.Update
                Else
                    lsMultiple = Left(lsMultiple & _
                        "00000000000000000000000000000000000000000000000000000000000000000000000000000000" _
                        , 80)

                    radrsJCLParameterList.AddNew
                    llCounter = llCounter + 1
                    radrsJCLParameterList!DATA_SEQ_NUM = llCounter
                    radrsJCLParameterList!DATA_VALUE = lsMultiple
                    radrsJCLParameterList.Update
                End If

                .Filter = adFilterNone
                lsMultiple = ""
                .Filter = MDAC25FilterFix("SORT_NUM > 0")
                If Not (.EOF) Then
                    For llIndex = 1 To .RecordCount
                        Select Case (Len(CStr(!SORT_NUM)))
                            Case 0
                                lsMultiple = lsMultiple & "00"
                            Case 1
                                lsMultiple = lsMultiple & _
                                                "0" & _
                                                CStr(!SORT_NUM)
                            Case 2
                                lsMultiple = lsMultiple & CStr(!SORT_NUM)
                        End Select

                        .MoveNext
                    Next llIndex
                Else
                    lsMultiple = lsMultiple & _
                        "00000000000000000000000000000000000000000000000000000000000000000000000000000000"
                End If

                lsMultiple = Left(lsMultiple & _
                    "00000000000000000000000000000000000000000000000000000000000000000000000000000000" _
                    , 80)

                radrsJCLParameterList.AddNew
                llCounter = llCounter + 1
                radrsJCLParameterList!DATA_SEQ_NUM = llCounter
                radrsJCLParameterList!DATA_VALUE = lsMultiple
                radrsJCLParameterList.Update

                .Filter = adFilterNone
                lsMultiple = ""
            End With

        Case "TR33143"  'Year of Seniority
            With vadrsParameterList
                .MoveFirst

                llCounter = 0
                lbAddRecord = True

                If Not (CheckForValidString(!UNION_CD, , , True, 2)) Then
                    AppendBRV radrsBRV _
                        , trErrbusJCLRequest _
                        , mstrCBUSINESSERR _
                        , "Invalid union." _
                        , MODULENAME & ": ValidateRequest"

                    lbAddRecord = False
                Else
                    If Not (CheckForItemExistenceInDB(UNION, !UNION_CD)) Then
                        AppendBRV radrsBRV _
                            , trErrbusJCLRequest _
                            , mstrCBUSINESSERR _
                            , "Invalid union. Union " & !UNION_CD & _
                                " does not exist." _
                            , MODULENAME & ": ValidateRequest"

                        lbAddRecord = False
                    End If
                End If

                If Not (CheckForValidInteger(!YEARS, True)) Then
                    AppendBRV radrsBRV _
                        , trErrbusJCLRequest _
                        , mstrCBUSINESSERR _
                        , "Invalid number of years." _
                        , MODULENAME & ": ValidateRequest"

                    lbAddRecord = False
                End If
            End With

            If (lbAddRecord) Then
                With radrsJCLParameterList
                    .AddNew
                    If (llCounter = 0) Then
                        !DATA_SEQ_NUM = 0
                    Else
                        llCounter = llCounter + 1
                        !DATA_SEQ_NUM = llCounter
                    End If
                    !DATA_VALUE = "EX TRSNR5YR UNION=" & _
                                    vadrsParameterList!UNION_CD & ","
                    .Update

                    .AddNew
                    llCounter = llCounter + 1
                    !DATA_SEQ_NUM = llCounter
                    !DATA_VALUE = "YEARS=" & vadrsParameterList!YEARS & ","
                    .Update

                    .AddNew
                    llCounter = llCounter + 1
                    !DATA_SEQ_NUM = llCounter
                    !DATA_VALUE = "FIN"
                    .Update
                End With
            End If
    End Select

    If (mADOExtensions.RecordsExist(vadrsPrinters)) Then
        vadrsPrinters.MoveFirst
    End If
    If (mADOExtensions.RecordsExist(vadrsParameterList)) Then
        vadrsParameterList.MoveFirst
    End If
    If (mADOExtensions.RecordsExist(radrsJCLParameterList)) Then
        radrsJCLParameterList.MoveFirst
    End If

    ValidateRequest = True

    If (mADOExtensions.RecordsExist(radrsBRV)) Then
        With radrsBRV
            .Filter = adFilterNone
            .MoveFirst

            lsFilterString = "ERR_TYPE = 'BE' OR ERR_TYPE = 'AE' OR ERR_TYPE = 'AW'"
            .Filter = MDAC25FilterFix(lsFilterString)
            lsFilterString = ""
            If Not (.EOF) Then
                ValidateRequest = False
            End If

            .Filter = adFilterNone
            .MoveFirst
        End With
    End If

    Exit Function
ValidateRequest_EH:
    TrsRaiseError MODULENAME _
        , "ValidateRequest"

End Function
Private Function CreateJCLParameterList() As ADODB.Recordset
On Error GoTo CreateJCLParameterList_EH
Dim ladrs As ADODB.Recordset

    Set ladrs = New ADODB.Recordset
    With ladrs
        .Fields.Append "DATA_SEQ_NUM", adInteger, , ADOFIELDATTR
        .Fields.Append "DATA_VALUE", adChar, 80, ADOFIELDATTR

        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open
    End With

    Set CreateJCLParameterList = ladrs

    Set ladrs = Nothing
    Exit Function
CreateJCLParameterList_EH:
    Set ladrs = Nothing
    TrsRaiseError MODULENAME _
        , "CreateJCLParameterList"

End Function
'NOTE: I decided to be silly and wrote IsAValid and AreValid methods
'AreValid methods are capable of handling one or more items - TONG
Private Function CheckForItemExistenceInDB(ByVal viItemType As Integer _
                    , Optional ByVal vsStringItem As String = "" _
                    , Optional ByVal vlLongItem As Long = 0 _
                    , Optional ByVal viIntegerItem As Integer = 0 _
                    , Optional ByVal viCheckSum As Integer = 0 _
                    ) As Boolean
On Error GoTo CheckForItemExistenceInDB_EH
Dim lobjDBJCL As dbJCL.cdJCL
Dim lobjdbRules As dbRules.cdRules
Dim lobjEmployee As dbEmployee.cdEmployee
Dim lobjCodeTable As dbCodeTables.cdCodeTables
Dim lobjWorkgroup As dbWorkGroup.cdWorkGroup
Dim ladrs As ADODB.Recordset

    Select Case (viItemType)
        Case PROCEDURE
            Set lobjdbRules = CreateObject( _
                                    "dbRules.cdRules")
            CheckForItemExistenceInDB = _
                lobjdbRules.DoesRuleApply(108, 22, vsStringItem)
            Set lobjdbRules = Nothing

        Case PRINTER
            CheckForItemExistenceInDB = True
            Set lobjDBJCL = CreateObject( _
                                    "dbJCL.cdJCL")
            CheckForItemExistenceInDB = _
                lobjDBJCL.IsAValidPrinter(vsStringItem)
            Set lobjDBJCL = Nothing

        Case PRINTERS
            CheckForItemExistenceInDB = True
            Set lobjDBJCL = CreateObject( _
                                    "dbJCL.cdJCL")
            If (lobjDBJCL.AreValidPrinters(vsStringItem) = _
                viCheckSum) Then
                CheckForItemExistenceInDB = True
            Else
                CheckForItemExistenceInDB = False
            End If

        Case EMPLOYEE
            Set lobjEmployee = CreateObject( _
                                    "dbEmployee.cdEmployee")
            CheckForItemExistenceInDB = _
                lobjEmployee.IsValidEmployee(vlLongItem)
            Set lobjEmployee = Nothing

        Case EMPLOYEES
            Set lobjEmployee = CreateObject( _
                                    "dbEmployee.cdEmployee")
            Set ladrs = _
                lobjEmployee.AreValidEmployees(vsStringItem)
            Set lobjEmployee = Nothing

            CheckForItemExistenceInDB = False
            If (mADOExtensions.RecordsExist(ladrs)) Then
                If (ladrs.RecordCount = viCheckSum) Then
                    CheckForItemExistenceInDB = True
                End If
            End If

        Case OPPORTUNITY
            Set lobjCodeTable = CreateObject( _
                                    "dbCodeTables.cdCodeTables")
            CheckForItemExistenceInDB = _
                lobjCodeTable.IsAValidOpportunity(vlLongItem)
            Set lobjCodeTable = Nothing

            CheckForItemExistenceInDB = True

        Case QUALIFICATION
            Set lobjCodeTable = CreateObject( _
                                    "dbCodeTables.cdCodeTables")
            CheckForItemExistenceInDB = _
                lobjCodeTable.IsAValidQualificationCode(vsStringItem)
            Set lobjCodeTable = Nothing

        Case SKILL
            Set lobjCodeTable = CreateObject( _
                                    "dbCodeTables.cdCodeTables")
            CheckForItemExistenceInDB = _
                lobjCodeTable.IsAValidSpecialSkillCode2(vsStringItem)
            Set lobjCodeTable = Nothing

        Case LOCATION
            Set lobjCodeTable = CreateObject( _
                                    "dbCodeTables.cdCodeTables")
            CheckForItemExistenceInDB = _
                lobjCodeTable.LocationCodeIsValid(CStr(vlLongItem))
            Set lobjCodeTable = Nothing

        Case BUDGET
            Set lobjCodeTable = CreateObject( _
                                    "dbCodeTables.cdCodeTables")
            CheckForItemExistenceInDB = _
                lobjCodeTable.IsAValidBudget(vsStringItem)
            Set lobjCodeTable = Nothing

        Case DEPARTMENT
            Set lobjCodeTable = CreateObject( _
                                    "dbCodeTables.cdCodeTables")
            CheckForItemExistenceInDB = _
                lobjCodeTable.IsAValidDepartment(vsStringItem)
            Set lobjCodeTable = Nothing

        Case WORKGROUP
            Set lobjWorkgroup = CreateObject( _
                                    "dbWorkgroup.cdWorkgroup")
            CheckForItemExistenceInDB = _
                lobjWorkgroup.IsValidWorkgroup(vlLongItem, Date)
            Set lobjWorkgroup = Nothing

        Case UNION
            Set lobjCodeTable = CreateObject( _
                                    "dbCodeTables.cdCodeTables")
            CheckForItemExistenceInDB = _
                lobjCodeTable.UnionCodeIsValid(vsStringItem)
            Set lobjCodeTable = Nothing

        Case JOB
            Set lobjCodeTable = CreateObject( _
                                    "dbCodeTables.cdCodeTables")
            CheckForItemExistenceInDB = _
                lobjCodeTable.JobNumberIsValid(vsStringItem)
            Set lobjCodeTable = Nothing

        Case OTBANK
            Set lobjCodeTable = CreateObject( _
                                    "dbCodeTables.cdCodeTables")
            CheckForItemExistenceInDB = _
                lobjCodeTable.DoesBankExist(vlLongItem)
            Set lobjCodeTable = Nothing
    End Select

    Set lobjDBJCL = Nothing
    Set lobjEmployee = Nothing
    Set lobjCodeTable = Nothing
    Set lobjWorkgroup = Nothing
    Set lobjdbRules = Nothing
    Set ladrs = Nothing

    Exit Function
CheckForItemExistenceInDB_EH:
    Set lobjDBJCL = Nothing
    Set lobjEmployee = Nothing
    Set lobjCodeTable = Nothing
    Set lobjWorkgroup = Nothing
    Set lobjdbRules = Nothing
    Set ladrs = Nothing

    TrsRaiseError MODULENAME _
        , "CheckForItemExistenceInDB"

End Function

''use procedure TR501
''&PROCNAME = TR33142
''&DESTOUT  = MSS1
'''DEPT   -- 6 DIGITS (0)
''&DATA001  = 00000000000000000000000000000000000000000000000000000000000000000000000000000000
'''LOC    - 2 DIGITS (0)
''&DATA002  = 00000000000000000000000000000000000000000000000000000000000000000000000000000000
'''UNION  - 2 CHARS (0)
''&DATA003  = 00000000000000000000000000000000000000000000000000000000000000000000000000000000
'''JOB    - 6 DIGITS (0)
''&DATA004  = 00000000000000000000000000000000000000000000000000000000000000000000000000000000
'''WG -- 9 DIGITS (0)
''&DATA005  = 00001167400000000000000000000000000000000000000000000000000000000000000000000000
'''WG -- 9 DIGITS (0)
''&DATA006  = 00000000000000000000000000000000000000000000000000000000000000000000000000000000
'''SORT -- 2 DIGITS (0)
''&DATA007  = 02000000000000000000000000000000000000000000000000000000000000000000000000000000

''Use Procedure TR501
''&PROCNAME = TR33145
''&DESTOUT = MSS1
''&DATA001 = EX TRVACSUM DEPT_NUM1=,DEPT_NUM2=,DEPT_NUM3=,DEPT_NUM4=,
''&DATA002 = HEADER=N,LOC_CD1=000000,LOC_CD2=000000,LOC_CD3=000000,
''&DATA003 = EMP_NUM1=41802,EMP_NUM2=00000,EMP_NUM3=00000,
''&DATA004 = EMP_NUM4=00000,EMP_NUM5=00000,EMP_NUM6=00000,
''&DATA005 = EMP_NUM7=00000,EMP_NUM8=00000,EMP_NUM9=00000,
''&DATA006 = EMP_NUM10=000000000,EMP_NUM11=000000000,EMP_NUM12=000000000,
''&DATA007 = EMP_NUM13=000000000,EMP_NUM14=000000000,EMP_NUM15=000000000,
''&DATA008 = WG_NUM1=000000000,WG_NUM2=000000000,WG_NUM3=000000000,
''&DATA009 = WG_NUM4=000000000,WG_NUM5=000000000,WG_NUM6=000000000,
''&DATA010 = WG_NUM7=000000000,WG_NUM8=000000000,WG_NUM9=000000000,
''&DATA011 = WG_NUM10=000000000,WG_NUM11=000000000,WG_NUM12=000000000,
''&DATA012 = WG_NUM13=000000000,WG_NUM14=000000000,WG_NUM15=000000000,
''&DATA013 = WG_NUM16=000000000,WG_NUM17=000000000,WG_NUM18=000000000,
''&DATA014 = WG_NUM19=000000000,WG_NUM20=000000000,
''&DATA015 = JOB_NUM1=,JOB_NUM2=,JOB_NUM3=,JOB_NUM4=,
''&DATA016 = JOB_NUM5=,JOB_NUM6=,JOB_NUM7=,JOB_NUM8=,
''&DATA017 = JOB_NUM9=,JOB_NUM10=,
''&DATA018 = TYPE=Y,BEGDTE=20031228,YEAR=2004,
''&DATA019 = BYPARM1=BY EMP_NAME NOPRINT,BYPARM2=,
''&DATA020 = BYPARM3=,BYPARM4=,
''&DATA021 = BYPARM5=,BYPARM6=,
''&DATA022 = BYPARM7=,BYPARM8=,
''&DATA023 = BYPARM9=,BYPARM10=,
''&DATA024 = FIN

''Use Procedure TR501
''&PROCNAME = TR33145
''&DESTOUT = MSS1
''&DATA001 = EX TRVACSUM DEPT_NUM1=,DEPT_NUM2=,DEPT_NUM3=,DEPT_NUM4=,
''&DATA002 = HEADER=N,LOC_CD1=000000,LOC_CD2=000000,LOC_CD3=000000,
''&DATA003 = EMP_NUM1=41802,EMP_NUM2=00000,EMP_NUM3=00000,
''&DATA004 = EMP_NUM4=00000,EMP_NUM5=00000,EMP_NUM6=00000,
''&DATA005 = EMP_NUM7=00000,EMP_NUM8=00000,EMP_NUM9=00000,
''&DATA006 = EMP_NUM10=000000000,EMP_NUM11=000000000,EMP_NUM12=000000000,
''&DATA007 = EMP_NUM13=000000000,EMP_NUM14=000000000,EMP_NUM15=000000000,
''&DATA008 = WG_NUM1=000000000,WG_NUM2=000000000,WG_NUM3=000000000,
''&DATA009 = WG_NUM4=000000000,WG_NUM5=000000000,WG_NUM6=000000000,
''&DATA010 = WG_NUM7=000000000,WG_NUM8=000000000,WG_NUM9=000000000,
''&DATA011 = WG_NUM10=000000000,WG_NUM11=000000000,WG_NUM12=000000000,
''&DATA012 = WG_NUM13=000000000,WG_NUM14=000000000,WG_NUM15=000000000,
''&DATA013 = WG_NUM16=000000000,WG_NUM17=000000000,WG_NUM18=000000000,
''&DATA014 = WG_NUM19=000000000,WG_NUM20=000000000,
''&DATA015 = JOB_NUM1=,JOB_NUM2=,JOB_NUM3=,JOB_NUM4=,
''&DATA016 = JOB_NUM5=,JOB_NUM6=,JOB_NUM7=,JOB_NUM8=,
''&DATA017 = JOB_NUM9=,JOB_NUM10=,
''&DATA018 = TYPE=Y,BEGDTE=20031228,YEAR=2004,
''&DATA019 = BYPARM1=BY EMP_NAME NOPRINT,BYPARM2=,
''&DATA020 = BYPARM3=,BYPARM4=,
''&DATA021 = BYPARM5=,BYPARM6=,
''&DATA022 = BYPARM7=,BYPARM8=,
''&DATA023 = BYPARM9=,BYPARM10=,
''&DATA024 = FIN

''Use Procedure TR501
''&PROCNAME = TR33128
''&DESTOUT = MSS1
''&DATA001 = EX TRROTBKA  SDATE=20041001,EDATE=20041013,EMPNUM=41802,BTYPE=      ,
''&DATA002 = OPPNUM=0,BSTAT= ,COMP1=GE,COMP2=GE,COMP3=GE,
''&DATA003 = FIN

''use procedure TR501
''&PROCNAME = TR33122
''&DESTOUT = MSS1
''&DATA001 = EX TRROCR03 DEPT=567,WGNUM0=0000000,
''&DATA002 = WGNUM1=0000000,WGNUM2=0000000,WGNUM3=0000000,
''&DATA003 = WGNUM4=0000000,WGNUM5=0000000,WGNUM6=0000000,
''&DATA004 = WGNUM7=0000000,WGNUM8=0000000,WGNUM9=0000000,
''&DATA005 = JOB1=,JOB2=,JOB3=,JOB4=,
''&DATA006 = JOB5=,JOB6=,JOB7=,
''&DATA007 = BANK= ALL, TEMP= B,
''&DATA008 = BY_PARM1  = BY LOWEST TOT_OT_HRS,
''&DATA009 = BY_PARM2  = -*,
''&DATA010 = BY_PARM3  = -*,
''&DATA011 = BY_PARM4  = -*,
''&DATA012 = BY_PARM5  = -*,
''&DATA013 = BY_PARM6  = -*,
''&DATA014 = BY_PARM7  = -*,
''&DATA015 = BY_PARM8  = -*,
''&DATA016 = BY_PARM9  = -*,
''&DATA017 = BY_PARM10 = -*,
''&DATA018 = BY_PARM11 = -*,
''&DATA019 = BY_PARM12 = -*,
''&DATA020 = BY_PARM13 = -*,
''&DATA021 = BY_PARM14 = -*,
''&DATA022 = BY_PARM15 = -*,
''&DATA023 = BY_PARM16 = -*,
''&DATA024 = BY_PARM17 = -*,
''&DATA025 = FIN
