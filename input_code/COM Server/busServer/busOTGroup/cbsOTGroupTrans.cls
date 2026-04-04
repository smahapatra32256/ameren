VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "cbsOTGroupTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'AMEREN Project: TRIS Conversion Phase 3
'June 2004
'Author: Prem M. Chander
'Module Description: This class will provide the user the ability
'to create/modify/delete Overtime groups

'COMMENTS: OCT, 2005 - CHANGE GETOBJECTCONTEXT.CREATEINSTANCE
'TO CREATEOBJECT -- PREM

Option Explicit

Private Const MODULENAME = "busOTGroup.cbsOTGroupTrans"
Private Const ADOFIELDATTR = adFldUpdatable Or adFldIsNullable

Private Const trOTGROUP As Integer = 1
Private Const trOTBANKCODE As Integer = 2
Private Const trLOCATION As Integer = 3
Private Const trWORKGROUP As Integer = 4
Private Const trOTWG As Integer = 5
Private Const trOTGROUPOPPORTUNITY = 6
Public Function ProcessOvertimeGroupEntries( _
                ByRef rlOTGroupNumber As Long _
                , ByVal vadrsOTGroupBase As ADODB.Recordset _
                , ByVal vadrsOTGroupBanks As ADODB.Recordset _
                , ByVal vadrsOTGroupWorkgroups As ADODB.Recordset _
                , ByRef radrsErrors As ADODB.Recordset _
                , ByVal vlUserEmployeeNumber As Long _
                ) As Boolean
On Error GoTo ProcessOvertimeGroupEntries_EH
Dim lbEmptyRequest As Boolean
Dim llRecordIndex As Long, lobjdbOTGroup As dbOTGroup.cdOTGroup

    'Prepare BRV Recordset
    Set radrsErrors = CreateOTGroupBRV

    If Not (mADOExtensions.RecordsExist(vadrsOTGroupBase)) And _
        Not (mADOExtensions.RecordsExist(vadrsOTGroupBanks)) And _
        Not (mADOExtensions.RecordsExist(vadrsOTGroupWorkgroups)) Then
        AppendOTGroupBRV radrsErrors _
            , trErrbusOvertimeGroup _
            , mstrCBUSINESSERR _
            , "There are no Overtime Group entries to process." _
            , MODULENAME & ": ProcessOvertimeGroupEntries"

        GoTo REPORTBRV
    End If

    If Not (ValidateOTGroupInformation(rlOTGroupNumber _
            , vadrsOTGroupBase _
            , vadrsOTGroupBanks _
            , vadrsOTGroupWorkgroups _
            , radrsErrors _
            , vlUserEmployeeNumber)) Then
        GoTo REPORTBRV
    End If

    Set lobjdbOTGroup = CreateObject( _
                            "dbOTGroup.cdOTGroup")
    If Not (lobjdbOTGroup.dbProcessOvertimeGroup( _
            rlOTGroupNumber _
            , vadrsOTGroupBase _
            , vadrsOTGroupBanks _
            , vadrsOTGroupWorkgroups _
            , vlUserEmployeeNumber)) Then
        Set lobjdbOTGroup = Nothing

        Err.Raise trErrbusOvertimeGroup
    End If
    Set lobjdbOTGroup = Nothing

    GetObjectContext().SetComplete

    ProcessOvertimeGroupEntries = True

    Exit Function
ProcessOvertimeGroupEntries_EH:
    GetObjectContext().SetAbort

    Set lobjdbOTGroup = Nothing

    Select Case (Err.Number)
        Case trErrutOpenConnection
            TrsRaiseError MODULENAME _
                , "ProcessOvertimeGroupEntries"
        Case trErrdbDelFromOTGroup, trErrdbInsIntoOTGroup _
            , trErrdbUpdOTGroup, trErrdbDelFromOTGroupBank _
            , trErrdbInsIntoOTGroupBank, trErrdbDelFromOTGroupWG _
            , trErrdbInsIntoOTGroupWG, trErrdbUpdOTGroupWG _
            , trErrdbSaveOTGroup, trErrdb_ExecuteSQL_F
            TrsRaiseError MODULENAME _
                , "ProcessOvertimeGroupEntries" _
                , Err.Number _
                , , , , _
                True
        Case Else
            TrsRaiseError MODULENAME _
                , "ProcessOvertimeGroupEntries" _
                , trErrbusOvertimeGroup _
                , , , , _
                True
    End Select

    Exit Function

REPORTBRV:
    Set lobjdbOTGroup = Nothing

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
                    ProcessOvertimeGroupEntries = False
                Else
                    'Complete Transaction
                    GetObjectContext().SetComplete
                    ProcessOvertimeGroupEntries = True
                End If

                .Filter = adFilterNone
            Else
                'Complete Transaction
                GetObjectContext().SetComplete
                ProcessOvertimeGroupEntries = True
            End If
        End With
    Else
        'Complete Transaction
        GetObjectContext().SetComplete
        ProcessOvertimeGroupEntries = True
    End If

    Exit Function

End Function
Public Function ProcessOTGroupDelete(ByVal vlOTGroupNumber As Long _
                , ByRef radrsErrors As ADODB.Recordset _
                , ByVal vlUserEmployeeNumber As Long _
                ) As Boolean
On Error GoTo ProcessOTGroupDelete_EH
Dim lobjdbOTGroup As dbOTGroup.cdOTGroup

    'Prepare BRV Recordset
    Set radrsErrors = CreateOTGroupBRV

    If Not (ValidateOTGroupDelete(vlOTGroupNumber _
            , radrsErrors _
            , vlUserEmployeeNumber)) Then
        GoTo REPORTBRV
    End If

    Set lobjdbOTGroup = CreateObject( _
                            "dbOTGroup.cdOTGroup")
    If Not (lobjdbOTGroup.ProcessOTGroupDelete(vlOTGroupNumber _
            , vlUserEmployeeNumber)) Then
        Set lobjdbOTGroup = Nothing

        Err.Raise trErrbusOvertimeGroup
    End If
    Set lobjdbOTGroup = Nothing

    GetObjectContext().SetComplete

    ProcessOTGroupDelete = True

    Exit Function
ProcessOTGroupDelete_EH:
    GetObjectContext().SetAbort

    Set lobjdbOTGroup = Nothing

    Select Case (Err.Number)
        Case trErrutOpenConnection
            TrsRaiseError MODULENAME _
                , "ProcessOTGroupDelete"
        Case trErrdbDelFromOTGroup, trErrdbDelFromOTGroupBank _
            , trErrdbDelFromOTGroupWG, trErrdbSaveOTGroup _
            , trErrdb_ExecuteSQL_F
            TrsRaiseError MODULENAME _
                , "ProcessOTGroupDelete" _
                , Err.Number _
                , , , , _
                True
        Case Else
            TrsRaiseError MODULENAME _
                , "ProcessOTGroupDelete" _
                , trErrbusOvertimeGroup _
                , , , , _
                True
    End Select

    Exit Function
REPORTBRV:
    Set lobjdbOTGroup = Nothing

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
                    ProcessOTGroupDelete = False
                Else
                    'Complete Transaction
                    GetObjectContext().SetComplete
                    ProcessOTGroupDelete = True
                End If

                .Filter = adFilterNone
            Else
                'Complete Transaction
                GetObjectContext().SetComplete
                ProcessOTGroupDelete = True
            End If
        End With
    Else
        'Complete Transaction
        GetObjectContext().SetComplete
        ProcessOTGroupDelete = True
    End If

    Exit Function

End Function
Private Function CreateOTGroupBRV() As ADODB.Recordset
On Error GoTo CreateOTGroupBRV_EH

Dim ladrs As New ADODB.Recordset

    With ladrs
        .Fields.Append "OT_GROUP_NUM", adInteger, , ADOFIELDATTR
        .Fields.Append "OT_BANK_CD", adInteger, , ADOFIELDATTR
        .Fields.Append "WG_NUM", adInteger, , ADOFIELDATTR

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

    Set CreateOTGroupBRV = ladrs
    Set ladrs = Nothing
    Exit Function

CreateOTGroupBRV_EH:
    Set ladrs = Nothing
    Set CreateOTGroupBRV = Nothing

    TrsRaiseError MODULENAME _
        , "CreateOTGroupBRV"

End Function
Private Function ValidateOTGroupInformation( _
                    ByVal vlOTGroupNumber As Long _
                    , ByRef radrsOTGroupBase As ADODB.Recordset _
                    , ByRef radrsOTGroupBanks As ADODB.Recordset _
                    , ByRef radrsOTGroupWorkgroups As ADODB.Recordset _
                    , ByRef radrsBRV As ADODB.Recordset _
                    , ByVal vlUserEmployeeNumber As Long _
                    ) As Boolean
On Error GoTo ValidateOTGroupInformation_EH
Dim llIndex As Long

    If Not (stdmTRISCommon.CheckForValidLong(vlUserEmployeeNumber, True)) Then
        AppendOTGroupBRV radrsBRV _
            , trErrbusOvertimeGroup _
            , mstrCBUSINESSERR _
            , "Invalid TRIS User." _
            , MODULENAME & ": ValidateOTGroupInformation"
    End If

    If (stdmTRISCommon.CheckForValidLong(vlOTGroupNumber, True)) Then
        If Not (CheckForItemExistenceInDB(trOTGROUP _
            , , vlOTGroupNumber)) Then
            AppendOTGroupBRV radrsBRV _
                , trErrbusOvertimeGroup _
                , mstrCBUSINESSERR _
                , "Invalid Overtime Group." _
                , MODULENAME & ": ValidateOTGroupInformation" _
                , vlOTGroupNumber
        End If
    End If

    If (mADOExtensions.RecordsExist(radrsOTGroupBase)) Then
        With radrsOTGroupBase
            .Filter = adFilterNone
            If Not (.EOF) Then
                .MoveFirst

                Select Case (!RECORD_STATUS)
                    Case "I", "C"
                        If Not (stdmTRISCommon.CheckForValidString(!OT_GROUP_DESC _
                            , True, 30)) Then
                            AppendOTGroupBRV radrsBRV _
                                , trErrbusOvertimeGroup _
                                , mstrCBUSINESSERR _
                                , "Invalid Overtime Group description." _
                                , MODULENAME & ": ValidateOTGroupInformation" _
                                , vlOTGroupNumber
                        Else
                            !OT_GROUP_DESC = stdmTRISCommon.ParseString(!OT_GROUP_DESC)
                            .Update
                        End If

                        If Not (stdmTRISCommon.CheckForValidInteger(!OT_STAND_BANK_CD _
                            , True)) Then
                            AppendOTGroupBRV radrsBRV _
                                , trErrbusOvertimeGroup _
                                , mstrCBUSINESSERR _
                                , "Invalid Overtime Group standing bank." _
                                , MODULENAME & ": ValidateOTGroupInformation" _
                                , vlOTGroupNumber
                        Else
                            If Not (CheckForItemExistenceInDB(trOTBANKCODE _
                                , , , !OT_STAND_BANK_CD)) Then
                                AppendOTGroupBRV radrsBRV _
                                    , trErrbusOvertimeGroup _
                                    , mstrCBUSINESSERR _
                                    , "Invalid Overtime standing bank." _
                                    , MODULENAME & ": ValidateOTGroupInformation" _
                                    , vlOTGroupNumber
                            End If
                        End If

                        If (stdmTRISCommon.CheckForValidInteger(!LOC_CD _
                            , True)) Then
                            If Not (CheckForItemExistenceInDB(trLOCATION _
                                , , , !LOC_CD)) Then
                                AppendOTGroupBRV radrsBRV _
                                    , trErrbusOvertimeGroup _
                                    , mstrCBUSINESSERR _
                                    , "Invalid Overtime Group." _
                                    , MODULENAME & ": ValidateOTGroupInformation" _
                                    , vlOTGroupNumber
                            End If
                        Else
                            AppendOTGroupBRV radrsBRV _
                                , trErrbusOvertimeGroup _
                                , mstrCBUSINESSERR _
                                , "Invalid location." _
                                , MODULENAME & ": ValidateOTGroupInformation" _
                                , vlOTGroupNumber
                        End If

                    Case "D"
'PREM PERFORM CHECK FOR OT_GROUP ASSOCIATION TO OPEN CANVASS
                        If Not (stdmTRISCommon.CheckForValidInteger(!KEY_OT_GROUP_NUM _
                            , True)) Then
                            AppendOTGroupBRV radrsBRV _
                                , trErrbusOvertimeGroup _
                                , mstrCBUSINESSERR _
                                , "Invalid overtime group." _
                                , MODULENAME & ": ValidateOTGroupInformation" _
                                , vlOTGroupNumber
                        Else
                            If (CheckForItemExistenceInDB(trOTGROUPOPPORTUNITY _
                                , , !KEY_OT_GROUP_NUM)) Then
                                AppendOTGroupBRV radrsBRV _
                                    , trErrbusOvertimeGroup _
                                    , mstrCBUSINESSERR _
                                    , "Overtime group is associated to an open canvass and cannot be deleted." _
                                    , MODULENAME & ": ValidateOTGroupInformation" _
                                    , vlOTGroupNumber
                            End If
                        End If

                        If Not (stdmTRISCommon.CheckForValidString(!CREATED_TIMESTAMP _
                            , , , True, 26)) Then
                            AppendOTGroupBRV radrsBRV _
                                , trErrbusOvertimeGroup _
                                , mstrCBUSINESSERR _
                                , "Invalid time stamp." _
                                , MODULENAME & ": ValidateOTGroupInformation" _
                                , !KEY_OT_GROUP_NUM
                        End If
                End Select
            End If
        End With
    End If

    If (mADOExtensions.RecordsExist(radrsOTGroupBase)) Then
        radrsOTGroupBase.Filter = adFilterNone
        radrsOTGroupBase.MoveFirst
    End If

    If (mADOExtensions.RecordsExist(radrsOTGroupBanks)) Then
        With radrsOTGroupBanks
            .Filter = adFilterNone
            If Not (.EOF) Then
                .MoveFirst

                For llIndex = 1 To .RecordCount
                    Select Case (!RECORD_STATUS)
                        Case "I"
                            If Not (stdmTRISCommon.CheckForValidInteger(!OT_BANK_CD _
                                , True)) Then
                                AppendOTGroupBRV radrsBRV _
                                    , trErrbusOvertimeGroup _
                                    , mstrCBUSINESSERR _
                                    , "Invalid Overtime Group standing bank." _
                                    , MODULENAME & ": ValidateOTGroupInformation" _
                                    , vlOTGroupNumber
                            Else
                                If Not (CheckForItemExistenceInDB(trOTBANKCODE _
                                    , , , !OT_BANK_CD)) Then
                                    AppendOTGroupBRV radrsBRV _
                                        , trErrbusOvertimeGroup _
                                        , mstrCBUSINESSERR _
                                        , "Invalid Overtime standing bank." _
                                        , MODULENAME & ": ValidateOTGroupInformation" _
                                        , vlOTGroupNumber _
                                        , !OT_BANK_CD
                                End If
                            End If

                        Case "D"
                            If Not (stdmTRISCommon.CheckForValidInteger(!KEY_OT_BANK_CD _
                                , True)) Then
                                AppendOTGroupBRV radrsBRV _
                                    , trErrbusOvertimeGroup _
                                    , mstrCBUSINESSERR _
                                    , "Invalid Overtime Group standing bank." _
                                    , MODULENAME & ": ValidateOTGroupInformation" _
                                    , vlOTGroupNumber
                            End If

                            If Not (stdmTRISCommon.CheckForValidString(!CREATED_TIMESTAMP _
                                , , , True, 26)) Then
                                AppendOTGroupBRV radrsBRV _
                                    , trErrbusOvertimeGroup _
                                    , mstrCBUSINESSERR _
                                    , "Invalid time stamp." _
                                    , MODULENAME & ": ValidateOTGroupInformation" _
                                    , vlOTGroupNumber _
                                    , !KEY_OT_BANK_CD
                            End If
                    End Select

                    .MoveNext
                Next llIndex
            End If
        End With
    End If

    If (mADOExtensions.RecordsExist(radrsOTGroupBanks)) Then
        radrsOTGroupBanks.Filter = adFilterNone
        radrsOTGroupBanks.MoveFirst
    End If

    If (mADOExtensions.RecordsExist(radrsOTGroupWorkgroups)) Then
        With radrsOTGroupWorkgroups
            .Filter = adFilterNone
            If Not (.EOF) Then
                .MoveFirst

                For llIndex = 1 To .RecordCount
                    Select Case (!RECORD_STATUS)
                        Case "I", "C"
                            If (stdmTRISCommon.CheckForValidLong(!WG_NUM _
                                , True)) Then
                                If Not (CheckForItemExistenceInDB(trWORKGROUP _
                                    , , !WG_NUM)) Then
                                    AppendOTGroupBRV radrsBRV _
                                        , trErrbusOvertimeGroup _
                                        , mstrCBUSINESSERR _
                                        , "Invalid workgroup." _
                                        , MODULENAME & ": ValidateOTGroupInformation" _
                                        , vlOTGroupNumber _
                                        , _
                                        , !WG_NUM
'PER USER DEMO ON AUGUST 19, 2004
'''                                Else
'''                                    If (!RECORD_STATUS = "I") Then
'''                                        If (CheckForItemExistenceInDB(trOTWG _
'''                                                , , !WG_NUM)) Then
'''                                            AppendOTGroupBRV radrsBRV _
'''                                                , trErrbusOvertimeGroup _
'''                                                , mstrCBUSINESSERR _
'''                                                , "Workgroup is already associated to another overtime group and cannot be assigned." _
'''                                                , MODULENAME & ": ValidateOTGroupInformation" _
'''                                                , vlOTGroupNumber _
'''                                                , _
'''                                                , !WG_NUM
'''                                        End If
'''                                    End If
                                End If
                            Else
                                AppendOTGroupBRV radrsBRV _
                                    , trErrbusOvertimeGroup _
                                    , mstrCBUSINESSERR _
                                    , "Invalid workgroup." _
                                    , MODULENAME & ": ValidateOTGroupInformation" _
                                    , vlOTGroupNumber
                            End If

                            If Not (stdmTRISCommon.CheckForValidInteger(!ORDER_OF_CALL _
                                , True)) Then
                                AppendOTGroupBRV radrsBRV _
                                    , trErrbusOvertimeGroup _
                                    , mstrCBUSINESSERR _
                                    , "Invalid order of call." _
                                    , MODULENAME & ": ValidateOTGroupInformation" _
                                    , vlOTGroupNumber _
                                    , _
                                    , !WG_NUM
                            End If
                    End Select

                    Select Case (!RECORD_STATUS)
                        Case "C", "D"
                            If Not (stdmTRISCommon.CheckForValidLong(!KEY_WG_NUM _
                                , True)) Then
                                AppendOTGroupBRV radrsBRV _
                                    , trErrbusOvertimeGroup _
                                    , mstrCBUSINESSERR _
                                    , "Invalid workgroup." _
                                    , MODULENAME & ": ValidateOTGroupInformation" _
                                    , vlOTGroupNumber
                            End If
                            If Not (stdmTRISCommon.CheckForValidString(!CREATED_TIMESTAMP _
                                , , , True, 26)) Then
                                AppendOTGroupBRV radrsBRV _
                                    , trErrbusOvertimeGroup _
                                    , mstrCBUSINESSERR _
                                    , "Invalid time stamp." _
                                    , MODULENAME & ": ValidateOTGroupInformation" _
                                    , vlOTGroupNumber _
                                    , _
                                    , !KEY_WG_NUM
                            End If
                    End Select

                    .MoveNext
                Next llIndex
            End If
        End With
    End If

    If (mADOExtensions.RecordsExist(radrsOTGroupWorkgroups)) Then
        radrsOTGroupWorkgroups.Filter = adFilterNone
        radrsOTGroupWorkgroups.MoveFirst
    End If

    ValidateOTGroupInformation = True

    If (mADOExtensions.RecordsExist(radrsBRV)) Then
        With radrsBRV
            .Filter = adFilterNone
            .MoveFirst

            .Filter = MDAC25FilterFix("ERR_TYPE = 'BE' OR ERR_TYPE = 'AE' OR ERR_TYPE = 'AW'")
            If Not (.EOF) Then
                ValidateOTGroupInformation = False
            End If

            .Filter = adFilterNone
        End With
    End If

    Exit Function
ValidateOTGroupInformation_EH:
    TrsRaiseError MODULENAME _
            , "ValidateOTGroupInformation"

End Function
Private Function ValidateOTGroupDelete(ByVal vlOTGroupNumber As Long _
                    , ByRef radrsBRV As ADODB.Recordset _
                    , ByVal vlUserEmployeeNumber As Long _
                    ) As Boolean
On Error GoTo ValidateOTGroupDelete_EH

    If Not (stdmTRISCommon.CheckForValidLong(vlUserEmployeeNumber, True)) Then
        AppendOTGroupBRV radrsBRV _
            , trErrbusOvertimeGroup _
            , mstrCBUSINESSERR _
            , "Invalid TRIS User." _
            , MODULENAME & ": ValidateOTGroupDelete"
    End If

    If (stdmTRISCommon.CheckForValidLong(vlOTGroupNumber, True)) Then
        If Not (CheckForItemExistenceInDB(trOTGROUP _
            , , vlOTGroupNumber)) Then
            AppendOTGroupBRV radrsBRV _
                , trErrbusOvertimeGroup _
                , mstrCBUSINESSERR _
                , "Invalid Overtime Group." _
                , MODULENAME & ": ValidateOTGroupDelete" _
                , vlOTGroupNumber
        Else
            If (CheckForItemExistenceInDB(trOTGROUPOPPORTUNITY _
                , , vlOTGroupNumber)) Then
                AppendOTGroupBRV radrsBRV _
                    , trErrbusOvertimeGroup _
                    , mstrCBUSINESSERR _
                    , "Overtime group is associated to an open canvass and cannot be deleted." _
                    , MODULENAME & ": ValidateOTGroupDelete" _
                    , vlOTGroupNumber
            End If
        End If
    End If

    ValidateOTGroupDelete = True

    If (mADOExtensions.RecordsExist(radrsBRV)) Then
        With radrsBRV
            .Filter = adFilterNone
            .MoveFirst

            .Filter = MDAC25FilterFix("ERR_TYPE = 'BE' OR ERR_TYPE = 'AE' OR ERR_TYPE = 'AW'")
            If Not (.EOF) Then
                ValidateOTGroupDelete = False
            End If

            .Filter = adFilterNone
        End With
    End If

    Exit Function
ValidateOTGroupDelete_EH:
    TrsRaiseError MODULENAME _
            , "ValidateOTGroupDelete"

End Function
Private Sub AppendOTGroupBRV(ByRef radrsBRV As ADODB.Recordset _
                , ByVal vlErrorNumber As Long _
                , ByVal vsErrorType As String _
                , ByVal vsErrorDescription As String _
                , ByVal vsErrorSource As String _
                , Optional ByVal vlOTGroupNumber As Long = 0 _
                , Optional ByVal viOTBankCode As Integer = 0 _
                , Optional ByVal vlWorkgroupNumber As Long = 0 _
                , Optional ByVal viRecordSource As Integer = 0)
On Error GoTo AppendOTGroupBRV_EH

    With radrsBRV
        If (.State <> adStateOpen) Then
            Exit Sub
        End If

        .AddNew

        !OT_GROUP_NUM = vlOTGroupNumber
        !OT_BANK_CD = viOTBankCode
        !WG_NUM = vlWorkgroupNumber

        !ERR_NUMBER = vlErrorNumber - vbObjectError
        !ERR_TYPE = Trim(vsErrorType)
        !REC_SOURCE = viRecordSource
        !ERR_DESCRIPTION = Trim(vsErrorDescription)
        !ERR_SOURCE = Trim(vsErrorSource)

        .Update
    End With

    Exit Sub
AppendOTGroupBRV_EH:
    TrsRaiseError MODULENAME _
        , "AppendOTGroupBRV"

End Sub
Private Function CheckForItemExistenceInDB(ByVal viItemType As Integer _
                    , Optional ByVal vsStringItem As String = "" _
                    , Optional ByVal vlLongItem As Long = 0 _
                    , Optional ByVal viIntegerItem As Integer = 0 _
                    ) As Boolean
On Error GoTo CheckForItemExistenceInDB_EH
Dim lobjdbOTGroup As dbOTGroup.cdOTGroup
Dim lobjdbCodeTables As dbCodeTables.cdCodeTables
Dim lobjWorkgroup As dbWorkGroup.cdWorkGroup
Dim ladrsOTWG As ADODB.Recordset

    Select Case (viItemType)
        Case trOTGROUP
            Set lobjdbOTGroup = CreateObject( _
                                    "dbOTGroup.cdOTGroup")
            CheckForItemExistenceInDB = _
                lobjdbOTGroup.IsAValidOvertimeGroup(vlLongItem)
            Set lobjdbOTGroup = Nothing

        Case trOTBANKCODE
            Set lobjdbCodeTables = CreateObject( _
                                    "dbCodeTables.cdCodeTables")
            CheckForItemExistenceInDB = _
                lobjdbCodeTables.DoesBankExist(viIntegerItem)
            Set lobjdbCodeTables = Nothing

        Case trLOCATION
            Set lobjdbCodeTables = CreateObject( _
                                    "dbCodeTables.cdCodeTables")
            CheckForItemExistenceInDB = _
                lobjdbCodeTables.LocationCodeIsValid(CStr(viIntegerItem))
            Set lobjdbCodeTables = Nothing

        Case trWORKGROUP
            Set lobjWorkgroup = CreateObject( _
                                    "dbWorkGroup.cdWorkGroup")
            CheckForItemExistenceInDB = _
                lobjWorkgroup.IsValidWorkgroup(vlLongItem, Date)
            Set lobjWorkgroup = Nothing

        Case trOTWG
            Set lobjdbOTGroup = CreateObject( _
                                    "dbOTGroup.cdOTGroup")
            CheckForItemExistenceInDB = _
                lobjdbOTGroup.IsWorkgroupAssociatedtoOTGroup(vlLongItem)
            Set lobjdbOTGroup = Nothing

        Case trOTGROUPOPPORTUNITY
            Set lobjdbOTGroup = CreateObject( _
                                    "dbOTGroup.cdOTGroup")
            CheckForItemExistenceInDB = _
                lobjdbOTGroup.IsOTGroupAssociatedToOpenOpportunity(vlLongItem)
            Set lobjdbOTGroup = Nothing
    End Select

    Exit Function
CheckForItemExistenceInDB_EH:
    Set lobjdbOTGroup = Nothing
    Set lobjdbCodeTables = Nothing
    Set lobjWorkgroup = Nothing

    TrsRaiseError MODULENAME _
        , "CheckForItemExistenceInDB"

End Function
