VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "cbsOTBanks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'AMEREN Project: TRIS Conversion Phase 3
'May 2004
'Author: Prem M. Chander
'Module Description: This class will provide the user the ability
'to get Employee Overtime Banks

Option Explicit

Private Const ADOFIELDATTR = adFldUpdatable Or adFldIsNullable
Private Const MODULENAME = "busOTBanks.cbsOTBanks"
'BEGIN PREM TRIS CONVERSION PHASE 3
Public Function GetEmployeeOvertimeBanks( _
                ByVal vadrsEmployees As ADODB.Recordset _
                , ByVal viOTBankCode As Integer _
                ) As ADODB.Recordset
On Error GoTo GetEmployeeOvertimeBanks_EH
Dim lsWhereClause As String, llIndex As Long
Dim lobjOTBanks As dbOTBanks.cdOTBanks

    'Check for existence and prepare the where clause
    If (RecordsExist(vadrsEmployees)) Then
        With vadrsEmployees
            .MoveFirst

            lsWhereClause = ""
            For llIndex = 1 To .RecordCount
                lsWhereClause = lsWhereClause & _
                                ", " & CStr(!EMP_NUM)
                .MoveNext
            Next llIndex
        End With
    End If

    If (Trim(lsWhereClause) <> "") Then
        If (Left(lsWhereClause, 1) = ",") Then
            'Replace with a blank
            lsWhereClause = Replace(lsWhereClause, "," _
                                , "", 1, 1)
        End If
    Else
        Set GetEmployeeOvertimeBanks = Nothing
        Exit Function
    End If

    'Get OT Banks
    Set lobjOTBanks = CreateObject( _
                        "dbOTBanks.cdOTBanks")
    Set GetEmployeeOvertimeBanks = lobjOTBanks.GetEmployeeOvertimeBanks( _
                        lsWhereClause, viOTBankCode)
    Set lobjOTBanks = Nothing

    Exit Function
GetEmployeeOvertimeBanks_EH:
    Set lobjOTBanks = Nothing

    TrsRaiseError MODULENAME _
        , "GetEmployeeOvertimeBanks" _
        , , , , _
        , True

End Function
Public Function GetEmployeeOvertimeBanks2( _
                ByVal vadrsEmployees As ADODB.Recordset _
                ) As ADODB.Recordset
On Error GoTo GetEmployeeOvertimeBanks2_EH
Dim lsWhereClause As String, llIndex As Long
Dim lobjOTBanks As dbOTBanks.cdOTBanks

    'Check for existence and prepare the where clause
    If (RecordsExist(vadrsEmployees)) Then
        With vadrsEmployees
            .MoveFirst

            lsWhereClause = ""
            For llIndex = 1 To .RecordCount
                lsWhereClause = lsWhereClause & _
                                ", " & CStr(!EMP_NUM)
                .MoveNext
            Next llIndex
        End With
    End If

    If (Trim(lsWhereClause) <> "") Then
        If (Left(lsWhereClause, 1) = ",") Then
            'Replace with a blank
            lsWhereClause = Replace(lsWhereClause, "," _
                                , "", 1, 1)
        End If
    Else
        Set GetEmployeeOvertimeBanks2 = Nothing
        Exit Function
    End If

    'Get OT Banks
    Set lobjOTBanks = CreateObject( _
                        "dbOTBanks.cdOTBanks")
    Set GetEmployeeOvertimeBanks2 = lobjOTBanks.GetEmployeeOvertimeBanks2( _
                                        lsWhereClause)
    Set lobjOTBanks = Nothing

    Exit Function
GetEmployeeOvertimeBanks2_EH:
    Set lobjOTBanks = Nothing

    TrsRaiseError MODULENAME _
        , "GetEmployeeOvertimeBanks2" _
        , , , , _
        , True

End Function
Public Function GetEmployeeOvertimeBanksForRezeroing( _
                ByVal vadrsEmployees As ADODB.Recordset _
                ) As ADODB.Recordset
On Error GoTo GetEmployeeOvertimeBanksForRezeroing_EH
Dim lsWhereClause As String, llIndex As Long
Dim lobjOTBanks As dbOTBanks.cdOTBanks

    'Check for existence and prepare the where clause
    If (RecordsExist(vadrsEmployees)) Then
        With vadrsEmployees
            .MoveFirst

            lsWhereClause = ""
            For llIndex = 1 To .RecordCount
                lsWhereClause = lsWhereClause & _
                                ", " & CStr(!EMP_NUM)
                .MoveNext
            Next llIndex
        End With
    End If

    If (Trim(lsWhereClause) <> "") Then
        If (Left(lsWhereClause, 1) = ",") Then
            'Replace with a blank
            lsWhereClause = Replace(lsWhereClause, "," _
                                , "", 1, 1)
        End If
    Else
        Set GetEmployeeOvertimeBanksForRezeroing = Nothing
        Exit Function
    End If

    'Get OT Banks
    Set lobjOTBanks = CreateObject( _
                        "dbOTBanks.cdOTBanks")
    Set GetEmployeeOvertimeBanksForRezeroing = _
        lobjOTBanks.GetEmployeeOvertimeBanksForRezeroing( _
                        lsWhereClause)
    Set lobjOTBanks = Nothing

    Exit Function
GetEmployeeOvertimeBanksForRezeroing_EH:
    Set lobjOTBanks = Nothing

    TrsRaiseError MODULENAME _
        , "GetEmployeeOvertimeBanksForRezeroing" _
        , , , , _
        , True

End Function
Public Function GetOTGroupEmployeesStandingBank( _
                ByVal vlOTGroupNumber As Long _
                ) As ADODB.Recordset
On Error GoTo GetOTGroupEmployeesStandingBank_EH
Dim lobjOTBanks As dbOTBanks.cdOTBanks

    'Get OT Banks
    Set lobjOTBanks = CreateObject( _
                        "dbOTBanks.cdOTBanks")
    Set GetOTGroupEmployeesStandingBank = lobjOTBanks.GetOTGroupEmployeesStandingBank( _
                        vlOTGroupNumber)
    Set lobjOTBanks = Nothing

    Exit Function
GetOTGroupEmployeesStandingBank_EH:
    Set lobjOTBanks = Nothing

    TrsRaiseError MODULENAME _
        , "GetOTGroupEmployeesStandingBank" _
        , , , , _
        , True

End Function
Public Function GetBanksForRezeroingByOTGroup( _
                ByVal vlOTGroupNumber As Long _
                ) As ADODB.Recordset
On Error GoTo GetBanksForRezeroingByOTGroup_EH
Dim lobjOTGroup As dbOTGroup.cdOTGroup
Dim lobjOTBanks As dbOTBanks.cdOTBanks
Dim ladrs As ADODB.Recordset, ladrsVlntr As ADODB.Recordset
Dim liOTGroupStandBank As Integer, liOTGroupLocation As Integer
Dim lbSetStandingBankToVlntr As Boolean

    Set GetBanksForRezeroingByOTGroup = Nothing
    lbSetStandingBankToVlntr = False

    'Get OT Group Information
    Set lobjOTGroup = CreateObject( _
                        "dbOTGroup.cdOTGroup")
    Set ladrs = lobjOTGroup.GetOvertimeGroupBase(vlOTGroupNumber)
    Set lobjOTGroup = Nothing

    If (RecordsExist(ladrs)) Then
        liOTGroupLocation = ladrs!LOC_CD
        liOTGroupStandBank = ladrs!OT_STAND_BANK_CD
        If (liOTGroupLocation <= 0) Or _
            (liOTGroupStandBank <= 0) Then
            Err.Raise trErrdbGetOTGroup
        End If
    Else
        Err.Raise trErrdbGetOTGroup
    End If
    Set ladrs = Nothing

    Set lobjOTBanks = CreateObject( _
                        "dbOTBanks.cdOTBanks")

    Select Case (liOTGroupLocation)
        Case 121
            Select Case (liOTGroupStandBank)
                Case 2 'Volunteer
                    'Get Standard Permanent Bank
                    Set ladrs = lobjOTBanks.GetBanksForRezeroByGroupAndBank(1 _
                                    , vlOTGroupNumber)
                    lbSetStandingBankToVlntr = True

                Case Is <> 2 'Not Volunteer
                    'Get Standing Bank
                    Set ladrs = lobjOTBanks.GetBanksForRezeroByGroupAndBank( _
                                    liOTGroupStandBank _
                                    , vlOTGroupNumber)
            End Select

            'Get Volunteer Information
            Set ladrsVlntr = lobjOTBanks.GetBanksForRezeroByGroupAndBank(2 _
                                , vlOTGroupNumber)
            Set lobjOTBanks = Nothing

        Case Else
            Set ladrs = lobjOTBanks.GetBanksForRezeroByGroupAndBank( _
                            liOTGroupStandBank _
                            , vlOTGroupNumber)
    End Select
    Set lobjOTBanks = Nothing

    Set GetBanksForRezeroingByOTGroup = AddAndSetBanks(ladrs _
                                            , ladrsVlntr _
                                            , lbSetStandingBankToVlntr)

    Set ladrs = Nothing
    Set ladrsVlntr = Nothing

    Exit Function
GetBanksForRezeroingByOTGroup_EH:
    Set lobjOTGroup = Nothing
    Set lobjOTBanks = Nothing
    Set ladrs = Nothing
    Set ladrsVlntr = Nothing

    TrsRaiseError MODULENAME _
        , "GetOTGroupEmployeesStandingBank" _
        , , , , _
        , True

End Function
Private Function AddAndSetBanks(ByVal vadrsBanks As ADODB.Recordset _
                    , ByVal vadrsVolunteer As ADODB.Recordset _
                    , ByVal vbVolunteer As Boolean _
                    ) As ADODB.Recordset
On Error GoTo AddAndSetBanks_EH
Dim ladrs As ADODB.Recordset, llRecordIndex As Long

    Set ladrs = DefineRezeroBanks

    If (RecordsExist(vadrsBanks)) Then
        With vadrsBanks
            For llRecordIndex = 1 To .RecordCount
                ladrs.AddNew

                ladrs!EMP_NUM = !EMP_NUM
                ladrs!EMP_NAME = !EMP_NAME
                ladrs!TOT_OT_HRS = !TOT_OT_HRS
                ladrs!TOT_OT_HRS_ACC = !TOT_OT_HRS_ACC
                ladrs!TOT_OT_HRS_REF = !TOT_OT_HRS_REF
                ladrs!TOT_VLNTR_HRS = 0
                ladrs!TOT_VLNTR_HRS_ACC = 0
                ladrs!TOT_VLNTR_HRS_REF = 0
                ladrs!LAST_MOD_EMP_NUM = !LAST_MOD_EMP_NUM
                ladrs!TIME_STAMP = !TIME_STAMP
                ladrs!VLNTR_LAST_MOD_EMP = 0
                ladrs!VLNTR_TIME_STAMP = ""
                ladrs!OT_BANK_CD = !OT_BANK_CD
                ladrs!OT_BANK_DESC = !OT_BANK_DESC
                If Not (vbVolunteer) Then
                    ladrs!OT_STAND_BANK_CD = !OT_BANK_CD
                    ladrs!OT_STAND_BANK_DESC = !OT_BANK_DESC
                Else
                    ladrs!OT_STAND_BANK_CD = 0
                    ladrs!OT_STAND_BANK_DESC = ""
                End If

                ladrs!VLNTR_BANK_CD = 0
                ladrs!VLNTR_BANK_DESC = ""

                ladrs.Update

                .MoveNext
            Next llRecordIndex
        End With
    End If

    If (RecordsExist(vadrsVolunteer)) Then
        With vadrsVolunteer
            .Filter = adFilterNone

            .Filter = MDAC25FilterFix("TOT_OT_HRS > 0")
            If Not (.EOF) Then
                .MoveFirst

                For llRecordIndex = 1 To .RecordCount
                    If (RecordsExist(ladrs)) Then
                        ladrs.MoveFirst

                        ladrs.Find "EMP_NUM = " & CStr(!EMP_NUM)
                        If Not (ladrs.EOF) Then
                            ladrs!TOT_VLNTR_HRS = !TOT_OT_HRS
                            ladrs!TOT_VLNTR_HRS_ACC = !TOT_OT_HRS_ACC
                            ladrs!TOT_VLNTR_HRS_REF = !TOT_OT_HRS_REF
                            ladrs!VLNTR_BANK_CD = !OT_BANK_CD
                            ladrs!VLNTR_BANK_DESC = !OT_BANK_DESC
                            ladrs!VLNTR_LAST_MOD_EMP = !LAST_MOD_EMP_NUM
                            ladrs!VLNTR_TIME_STAMP = !TIME_STAMP

                            If (vbVolunteer) Then
                                ladrs!OT_STAND_BANK_CD = !OT_BANK_CD
                                ladrs!OT_STAND_BANK_DESC = !OT_BANK_DESC
                            End If
                            ladrs.Update
                        End If
                    Else
                        ladrs.AddNew

                        ladrs!EMP_NUM = !EMP_NUM
                        ladrs!EMP_NAME = !EMP_NAME
                        ladrs!TOT_OT_HRS = 0
                        ladrs!TOT_OT_HRS_ACC = 0
                        ladrs!TOT_OT_HRS_REF = 0
                        ladrs!TOT_VLNTR_HRS = !TOT_OT_HRS
                        ladrs!TOT_VLNTR_HRS_ACC = !TOT_OT_HRS_ACC
                        ladrs!TOT_VLNTR_HRS_REF = !TOT_OT_HRS_REF
                        ladrs!LAST_MOD_EMP_NUM = !LAST_MOD_EMP_NUM
                        ladrs!TIME_STAMP = !TIME_STAMP
                        ladrs!OT_BANK_CD = 0
                        ladrs!OT_BANK_DESC = ""
                        ladrs!VLNTR_BANK_CD = !OT_BANK_CD
                        ladrs!VLNTR_BANK_DESC = !OT_BANK_DESC
                        ladrs!VLNTR_LAST_MOD_EMP = !LAST_MOD_EMP_NUM
                        ladrs!VLNTR_TIME_STAMP = !TIME_STAMP
                        If (vbVolunteer) Then
                            ladrs!OT_STAND_BANK_CD = !OT_BANK_CD
                            ladrs!OT_STAND_BANK_DESC = !OT_BANK_DESC
                        End If

                        ladrs.Update
                    End If

                    .MoveNext
                Next llRecordIndex
            End If

            vadrsVolunteer.MoveFirst

            If (vbVolunteer) And _
                (RecordsExist(ladrs)) Then
                With ladrs
                    .MoveFirst

                    For llRecordIndex = 1 To .RecordCount
                        !OT_STAND_BANK_CD = vadrsVolunteer!OT_BANK_CD
                        !OT_STAND_BANK_DESC = vadrsVolunteer!OT_BANK_DESC

                        .Update

                        .MoveNext
                    Next llRecordIndex

                    .MoveFirst
                End With
            End If
        End With
    End If

    Set AddAndSetBanks = ladrs
    Set ladrs = Nothing

    Exit Function
AddAndSetBanks_EH:
    Set ladrs = Nothing

    TrsRaiseError MODULENAME, "AddAndSetBanks"
End Function
Private Function DefineRezeroBanks() As ADODB.Recordset
On Error GoTo DefineRezeroBanks_EH
Dim ladrs As ADODB.Recordset

    Set DefineRezeroBanks = Nothing

    Set ladrs = New ADODB.Recordset
    With ladrs
        .Fields.Append "EMP_NUM", adInteger, , ADOFIELDATTR
        .Fields.Append "EMP_NAME", adChar, 30, ADOFIELDATTR
        .Fields.Append "TOT_OT_HRS", adDecimal, , ADOFIELDATTR
        .Fields.Item("TOT_OT_HRS").Precision = 7
        .Fields.Item("TOT_OT_HRS").NumericScale = 2
        .Fields.Append "TOT_OT_HRS_ACC", adDecimal, , ADOFIELDATTR
        .Fields.Item("TOT_OT_HRS_ACC").Precision = 7
        .Fields.Item("TOT_OT_HRS_ACC").NumericScale = 2
        .Fields.Append "TOT_OT_HRS_REF", adDecimal, , ADOFIELDATTR
        .Fields.Item("TOT_OT_HRS_REF").Precision = 7
        .Fields.Item("TOT_OT_HRS_REF").NumericScale = 2
        .Fields.Append "TOT_VLNTR_HRS", adDecimal, , ADOFIELDATTR
        .Fields.Item("TOT_VLNTR_HRS").Precision = 7
        .Fields.Item("TOT_VLNTR_HRS").NumericScale = 2
        .Fields.Append "TOT_VLNTR_HRS_ACC", adDecimal, , ADOFIELDATTR
        .Fields.Item("TOT_VLNTR_HRS_ACC").Precision = 7
        .Fields.Item("TOT_VLNTR_HRS_ACC").NumericScale = 2
        .Fields.Append "TOT_VLNTR_HRS_REF", adDecimal, , ADOFIELDATTR
        .Fields.Item("TOT_VLNTR_HRS_REF").Precision = 7
        .Fields.Item("TOT_VLNTR_HRS_REF").NumericScale = 2

        .Fields.Append "LAST_MOD_EMP_NUM", adInteger, , ADOFIELDATTR
        .Fields.Append "VLNTR_LAST_MOD_EMP", adInteger, , ADOFIELDATTR
        .Fields.Append "TIME_STAMP", adChar, 26, ADOFIELDATTR
        .Fields.Append "VLNTR_TIME_STAMP", adChar, 26, ADOFIELDATTR

        .Fields.Append "OT_BANK_CD", adInteger, , ADOFIELDATTR
        .Fields.Append "OT_BANK_DESC", adChar, 30, ADOFIELDATTR
        .Fields.Append "VLNTR_BANK_CD", adInteger, , ADOFIELDATTR
        .Fields.Append "VLNTR_BANK_DESC", adChar, 30, ADOFIELDATTR
        .Fields.Append "OT_STAND_BANK_CD", adInteger, , ADOFIELDATTR
        .Fields.Append "OT_STAND_BANK_DESC", adChar, 30, ADOFIELDATTR

        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic

        .Open
    End With

    Set DefineRezeroBanks = ladrs

    Set ladrs = Nothing

    Exit Function
DefineRezeroBanks_EH:
    Set ladrs = Nothing

    TrsRaiseError MODULENAME, "DefineRezeroBanks"
End Function
'END PREM TRIS CONVERSION PHASE 3
