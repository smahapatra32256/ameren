VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "cbsEAM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Project:   Post TRIS 5.5 Employee Absence Management
'Author:    Prem Chander
'Date:      August 2005
'Description:

Option Explicit
Private Const MODULENAME As String = "busEAM.cbsEAM"
Private Const ADOFIELDATTR = adFldUpdatable Or adFldIsNullable
'BEGIN PREM
Public Sub GetAbsenceEmployeeInformation(ByVal vlEmployeeNumber As Long _
        , ByVal vsTRISUser As String _
        , ByRef radrsEmployeeInformation As ADODB.Recordset)
TrsRaiseError MODULENAME, "GetAbsenceEmployeeInformation", trErrObsoleteMethod, , , , True
'On Error GoTo GetAbsenceEmployeeInformation_EH
'Dim lobjdbEAM As dbEAM.cdEAM
'
'    Set lobjdbEAM = CreateObject("dbEAM.cdEAM")
'    Set radrsEmployeeInformation = lobjdbEAM.GetAbsenceEmployeeInformation( _
'                                    vlEmployeeNumber)
'    Set lobjdbEAM = Nothing
'
'    Exit Sub
'GetAbsenceEmployeeInformation_EH:
'    Set lobjdbEAM = Nothing
'
'    TrsRaiseError MODULENAME _
'        , "GetAbsenceEmployeeInformation" _
'        , , , , _
'        , True
End Sub
Public Function GetEmployeeAbsences(ByVal vlEmployeeNumber As Long _
                , ByVal vsTRISUser As String _
                , ByVal vlTRISUserNumber As Long _
                , Optional ByVal vdtStartDate As Date = #1/1/1900# _
                , Optional ByVal vdtEndDate As Date = #12/31/9999# _
                ) As ADODB.Recordset
On Error GoTo GetEmployeeAbsences_EH
Dim lobjdb As dbEAM.cdEAM

''    'THIS IS A VERY CRUDE WAY...
''    If (vdtStartDate = #1/1/1900#) Then
''        vdtStartDate = DateAdd("d", -30, Date)
''    End If

    Set lobjdb = CreateObject("dbEAM.cdEAM")
    Set GetEmployeeAbsences = lobjdb.GetEmployeeAbsences(vlEmployeeNumber _
                            , vdtStartDate _
                            , vdtEndDate _
                            , vlTRISUserNumber)
    Set lobjdb = Nothing

    Exit Function
GetEmployeeAbsences_EH:
    Set lobjdb = Nothing

    TrsRaiseError MODULENAME _
        , "GetEmployeeAbsences" _
        , , , , _
        , True

End Function
Public Function GetEAMEmployeeInformation(ByVal vlEmployeeNumber As Long _
                    , ByVal vsTRISUser As String _
                    , ByVal vlTRISUserNumber As Long _
                    ) As ADODB.Recordset
TrsRaiseError MODULENAME, "GetEAMEmployeeInformation", trErrObsoleteMethod, , , , True
'On Error GoTo GetEAMEmployeeInformation_EH
'Dim lobjCollPeriod As busCollectionPeriod.cbsCollectionPeriod
'Dim lobjEmployee As busEmployee.cbsEmployee
'Dim lobjdbRules As dbRules.cdRules
'Dim lobjPS As dbPeopleSoft.cdPeopleSoft
'
'Dim ladrsRules As ADODB.Recordset
'Dim ladrs As ADODB.Recordset
'Dim liTimeIncrement As Integer, lbExecutive As Boolean
'
'Dim ldtLastProcessedDate As Date, lsHomePhone As String
'Dim liLocation As Integer, lsDepartment As String
'Dim lsPreviousDept As String, ldtDeptEffectiveDate As Date
'Dim lsUnion As String, lsExecContractIndicator As String
'Dim lsCompany As String, lsMgmtBenefitsPkg As String
'Dim lsUnionBenefitsPkg As String, lsEmpCmpy As Boolean
'
'    Set GetEAMEmployeeInformation = Nothing
'
'    ldtLastProcessedDate = #1/1/100#
'    Set lobjCollPeriod = CreateObject("busCollectionPeriod.cbsCollectionPeriod")
'    Set ladrs = lobjCollPeriod.GetCollectionPeriodStatusByEmployee( _
'                    vlEmployeeNumber, CDate("12/30/1899"), 0)
'    If (RecordsExist(ladrs)) Then
'        ldtLastProcessedDate = lobjCollPeriod.GetLatestProcessedPayPeriodEndDate( _
'                                    ladrs!PAYROLL_GROUP_CD)
'    End If
'    Set lobjCollPeriod = Nothing
'    Set ladrs = Nothing
'
'    Set lobjPS = CreateObject("dbPeopleSoft.cdPeopleSoft")
'    Set ladrs = lobjPS.GetPeopleSoftPhoneNumber(vlEmployeeNumber)
'    Set lobjPS = Nothing
'
'    lsHomePhone = ""
'    If (RecordsExist(ladrs)) Then
'        lsHomePhone = Trim(ladrs!HOME_PHONE_NUM)
'    End If
'    Set ladrs = Nothing
'
'    liTimeIncrement = 6
'
'    Set lobjEmployee = CreateObject("busEmployee.cbsEmployee")
'    Set ladrs = lobjEmployee.GetBasicEmployeeInformation( _
'                                    vlEmployeeNumber)
'    If (RecordsExist(ladrs)) Then
'        lbExecutive = False
'        If (ladrs!EXEC_CONTRACT_IND = "E" Or _
'            ladrs!EXEC_CONTRACT_IND = "e") Then
'            lbExecutive = True
'        End If
'        liTimeIncrement = lobjEmployee.PartialHourMinutes(ladrs!DEPT_NUM _
'                            , lbExecutive _
'                            , ladrs!CMPY_CD)
'    End If
'    Set lobjEmployee = Nothing
'
'    liLocation = 0
'    lsUnionBenefitsPkg = ""
'    lsMgmtBenefitsPkg = ""
'    lsCompany = ""
'    lsExecContractIndicator = ""
'    lsDepartment = ""
'    lsPreviousDept = ""
'    ldtDeptEffectiveDate = #1/1/100#
'    lsUnion = ""
'
'    If (RecordsExist(ladrs)) Then
'        With ladrs
'            .MoveFirst
'            liLocation = !LOC_CD
'            If (Trim(lsHomePhone) = "") Then
'                lsHomePhone = !HOME_PHONE_NUM
'            End If
'            lsUnionBenefitsPkg = !UNION_BEN_PKG_IND
'            lsMgmtBenefitsPkg = !MGMT_BEN_PKG_IND
'            lsCompany = !CMPY_CD
'            lsExecContractIndicator = !EXEC_CONTRACT_IND
'            lsDepartment = !DEPT_NUM
'            lsPreviousDept = !PREV_DEPT_NUM
'            ldtDeptEffectiveDate = !DEPT_EFF_DATE
'            lsUnion = !UNION_CD
'        End With
'    End If
'    Set lobjEmployee = Nothing
'    Set ladrs = Nothing
'
'    'Get Auto Determine Case Rules
'    Set lobjdbRules = CreateObject("dbRules.cdRules")
'    Set ladrsRules = lobjdbRules.GetRule(302)
'    Set lobjdbRules = Nothing
'
'    Set ladrs = New ADODB.Recordset
'    With ladrs
'        .Fields.Append "CMPY_CD", adChar, 3, ADOFIELDATTR
'        .Fields.Append "DEPT_NUM", adChar, 3, ADOFIELDATTR
'        .Fields.Append "PREV_DEPT_NUM", adChar, 3, ADOFIELDATTR
'        .Fields.Append "DEPT_EFF_DATE", adDate, , ADOFIELDATTR
'        .Fields.Append "EXEC_CONTRACT_IND", adChar, 1, ADOFIELDATTR
'        .Fields.Append "HOME_PHONE_NUM", adChar, 15, ADOFIELDATTR
'        .Fields.Append "LOC_CD", adInteger, , ADOFIELDATTR
'        .Fields.Append "MGMT_BEN_PKG_IND", adChar, 4, ADOFIELDATTR
'        .Fields.Append "UNION_BEN_PKG_IND", adChar, 4, ADOFIELDATTR
'        .Fields.Append "UNION_CD", adChar, 3, ADOFIELDATTR
'
'        .Fields.Append "ILLINOIS_CMPY", adChar, 1, ADOFIELDATTR
'        .Fields.Append "AUTO_CASE", adChar, 1, ADOFIELDATTR
'        .Fields.Append "AUTO_CASE_TYPE", adChar, 1, ADOFIELDATTR
'        .Fields.Append "AUTO_CASE_MM_DD_IND", adChar, 1, ADOFIELDATTR
'        .Fields.Append "AUTO_CASE_DURATION", adInteger, , ADOFIELDATTR
'        .Fields.Append "LAST_PP_PROC_DATE", adDate, , ADOFIELDATTR
'        .Fields.Append "PS_URL", adVarChar, 4096, ADOFIELDATTR
'        .Fields.Append "FORM457_URL", adVarChar, 4096, ADOFIELDATTR
'        .Fields.Append "SL_POLICY_TXT", adVarChar, 4096, ADOFIELDATTR
'        .Fields.Append "SL_POLICY_FROM", adVarChar, 4096, ADOFIELDATTR
'        .Fields.Append "TIME_INCREMENT", adInteger, , ADOFIELDATTR
'        .Fields.Append "SUPERVISOR_NUM", adInteger, , ADOFIELDATTR
'        .Fields.Append "SUPERVISOR_NAME", adVarChar, 30, ADOFIELDATTR
'        .Fields.Append "LOC_DESC", adVarChar, 30, ADOFIELDATTR
'
'        .CursorLocation = adUseClient
'        .CursorType = adOpenStatic
'        .LockType = adLockBatchOptimistic
'        .Open
'
'        .AddNew
'
'        !CMPY_CD = lsCompany
'        !DEPT_NUM = lsDepartment
'        !PREV_DEPT_NUM = lsPreviousDept
'        !DEPT_EFF_DATE = ldtDeptEffectiveDate
'        Select Case (Trim(lsExecContractIndicator))
'            Case "E", "e"
'                !EXEC_CONTRACT_IND = "E"
'            Case "C", "c"
'                !EXEC_CONTRACT_IND = "C"
'        End Select
'        !HOME_PHONE_NUM = lsHomePhone
'        !LOC_CD = liLocation
'        !MGMT_BEN_PKG_IND = lsMgmtBenefitsPkg
'        !UNION_BEN_PKG_IND = lsUnionBenefitsPkg
'        !UNION_CD = lsUnion
'
'        !ILLINOIS_CMPY = "N"
'        !SL_POLICY_TXT = "(The Sick Leave policy may be discussed with the Employee on first follow-up call and as necessary thereafter.)" & vbCrLf & "# " & Chr(34) & "...Employee on sick leave, whether for personal illness or injury, are generally required to remain at their address of record unless hospitalized, keeping a medical appointment, or obtaining medication." & vbCrLf & "...Special cases which require that an employee on sick leave be away from his/her address of record for prescribed physical activity, therapy, etc. should be discussed with the employee's supervisor in advance." & vbCrLf & "# ...Violations of these requirements will be considered as abuse of sick leave and subject to disciplinary action, up to and including discharge." & Chr(34)
'        !SL_POLICY_FROM = "From letter by D.A. Slagley dated 05/23/1989, Personal Activities While On Sick Leave"
'        Select Case (Trim(!CMPY_CD))
'            Case "ARG", "CIL", "CIS", "MV1", "MV3" _
'                , "AED", "CIC", "CIP", "GEN", "GMC" _
'                , "IHC", "IMS", "IPC"
'                !ILLINOIS_CMPY = "Y"
'                !SL_POLICY_TXT = "Employees on sick leave, whether for personal illness or injury, are generally required to remain at their address of record unless hospitalized, keeping a medical appointment or obtaining medication."
'                !SL_POLICY_FROM = ""
'        End Select
'
'        !AUTO_CASE = "N"
'        If (Trim(lsUnion) <> "") Then
'            If (RecordsExist(ladrsRules)) Then
'                ladrsRules.MoveFirst
'
'                ladrsRules.Find "RG_VALUE_CD LIKE '" & lsUnion & "%'"
'                If Not (ladrsRules.EOF) Then
'                    !AUTO_CASE = "Y"
'
'                    If (InStr(1, ladrsRules!RG_VALUE_CD, lsUnion & " .0.mm") > 0) Then
'                        !AUTO_CASE_TYPE = "S"
'                        !AUTO_CASE_MM_DD_IND = "M"
'                        !AUTO_CASE_DURATION = CInt(Right(Trim(ladrsRules!RG_VALUE_CD), 2))
'                    End If
'                    If (InStr(1, ladrsRules!RG_VALUE_CD, lsUnion & " .0.dd") > 0) Then
'                        !AUTO_CASE_TYPE = "S"
'                        !AUTO_CASE_MM_DD_IND = "D"
'                        !AUTO_CASE_DURATION = CInt(Right(Trim(ladrsRules!RG_VALUE_CD), 2))
'                    End If
'                    If (InStr(1, ladrsRules!RG_VALUE_CD, lsUnion & " .1.mm") > 0) Then
'                        !AUTO_CASE_TYPE = "C"
'                        !AUTO_CASE_MM_DD_IND = "M"
'                        !AUTO_CASE_DURATION = CInt(Right(Trim(ladrsRules!RG_VALUE_CD), 2))
'                    End If
'                    If (InStr(1, ladrsRules!RG_VALUE_CD, lsUnion & " .1.dd") > 0) Then
'                        !AUTO_CASE_TYPE = "C"
'                        !AUTO_CASE_MM_DD_IND = "D"
'                        !AUTO_CASE_DURATION = CInt(Right(Trim(ladrsRules!RG_VALUE_CD), 2))
'                    End If
'                End If
'            End If
'        End If
'
'        !LAST_PP_PROC_DATE = ldtLastProcessedDate
'        !PS_URL = "https://hrprod1:7002/"
'        !FORM457_URL = "https://sunp9.ameren.com:80/idc/groups/public/@public015506/documents/document_other/gob019622.pdf"
'        !TIME_INCREMENT = liTimeIncrement
'
'        .Update
'    End With
'
'    Set lobjCollPeriod = Nothing
'    Set lobjEmployee = Nothing
'    Set lobjdbRules = Nothing
'    Set ladrsRules = Nothing
'
'    Set GetEAMEmployeeInformation = ladrs
'
'    Set ladrs = Nothing
'
'    Exit Function
'GetEAMEmployeeInformation_EH:
'    Set lobjCollPeriod = Nothing
'    Set lobjEmployee = Nothing
'    Set lobjdbRules = Nothing
'    Set lobjPS = Nothing
'
'    Set ladrsRules = Nothing
'    Set ladrs = Nothing
'
'    TrsRaiseError MODULENAME _
'        , "GetEAMEmployeeInformation" _
'        , , , , _
'        , True

End Function
Public Sub GetAbsenceInformation(ByVal vlAbsenceNumber As Long _
                , ByVal vsTRISUser As String _
                , ByRef radrsAbsenceDate As ADODB.Recordset _
                , ByRef radrsInitialAbsence As ADODB.Recordset _
                , ByRef radrsAbsenceProvider As ADODB.Recordset _
                , ByRef radrsAbsenceContact As ADODB.Recordset _
                , ByRef radrsAbsenceReturnToWork As ADODB.Recordset _
                , ByRef radrsAbsenceContactCallLog As ADODB.Recordset _
                , ByRef radrsScheduledExceptions As ADODB.Recordset _
                , ByRef radrsAbsenceTimeDetails As ADODB.Recordset _
                , ByRef radrsCaseNumbers As ADODB.Recordset)
On Error GoTo GetAbsenceInformation_EH
Dim lobjdbEAM As dbEAM.cdEAM
Dim lobjSchedule As busSchedule.cbsSchedule
Dim lobjEmployee As busEmployee.cbsEmployee
Dim lobjSchedExc As busSchedException.cbsSchedExc

Dim ladrsSchedule As ADODB.Recordset
Dim ladrsDates As ADODB.Recordset
Dim llEmployeeNumber As Long, ldtReturnDate As Date, ldtAbsInitEndDate As Date
Dim llIndex As Long

    Set lobjdbEAM = CreateObject("dbEAM.cdEAM")

    Set radrsAbsenceDate = lobjdbEAM.GetAbsenceDatesByAbsence(vlAbsenceNumber)
    Set radrsInitialAbsence = lobjdbEAM.GetInitialAbsenceByAbsence(vlAbsenceNumber)
    Set radrsAbsenceProvider = lobjdbEAM.GetAbsenceProviderByAbsence(vlAbsenceNumber)
    Set radrsAbsenceContact = lobjdbEAM.GetContactByAbsence(vlAbsenceNumber)
    Set radrsAbsenceReturnToWork = lobjdbEAM.GetReturnToWorkByAbsence(vlAbsenceNumber)
    Set radrsAbsenceContactCallLog = lobjdbEAM.GetContactCallLogByAbsence(vlAbsenceNumber)

    Set lobjdbEAM = Nothing

    If (RecordsExist(radrsInitialAbsence)) Then
        radrsInitialAbsence.MoveFirst

        llEmployeeNumber = radrsInitialAbsence!EMP_NUM
        ldtReturnDate = radrsInitialAbsence!RTN_TO_WRK_DT
        If (ldtReturnDate < radrsInitialAbsence!ABS_ST_DT) Then
            ldtReturnDate = DateAdd("d", 1, radrsInitialAbsence!ABS_ST_DT)
        End If

        Set lobjEmployee = CreateObject("busEmployee.cbsEmployee")
'        Set radrsCaseNumbers = lobjEmployee.GetCaseGroupsAndNumbers( _
'                                llEmployeeNumber _
'                                , radrsInitialAbsence!ABS_ST_DT _
'                                , ldtReturnDate)
        Set radrsCaseNumbers = Nothing '//this is intentional as it is not used anymore
        Set lobjEmployee = Nothing
       
        Set lobjSchedExc = CreateObject("busSchedException.cbsSchedExc")
        Set radrsScheduledExceptions = lobjSchedExc.GetScheduledExceptions( _
                                llEmployeeNumber _
                                , radrsInitialAbsence!ABS_ST_DT _
                                , ldtReturnDate)

'        '//exclude any exceptions not tied to absence, either past the return to work date or exception code based
'        '//TODO: Should this be using the absence end date or the return to work date (either way it needs to be date worked rather than charge date)
'        '//the time is in this format "00.00.00"
'        ldtAbsInitEndDate = DateTimeFrom(radrsInitialAbsence!ABS_END_DT, Format(radrsInitialAbsence!ABS_END_TM, "hh:mm"))
'        If RecordsExist(radrsScheduledExceptions) Then
'            radrsScheduledExceptions.MoveLast
'            Do While Not radrsScheduledExceptions.BOF
'                If DateTimeFrom(radrsScheduledExceptions!CHARGE_DATE, radrsScheduledExceptions!START_TIME) >= ldtAbsInitEndDate Then
''                    If Trim(radrsScheduledExceptions!EXCEPTION_CD) = "H" Then '//TODO: This could include other exceptions
'                        radrsScheduledExceptions.Delete
''                    End If
'                Else
'                    Exit Do
'                End If
'                radrsScheduledExceptions.MovePrevious
'            Loop
'            radrsScheduledExceptions.MoveFirst
'        End If
        
        Set radrsAbsenceTimeDetails = lobjSchedExc.GetViewCalendarDetails( _
                                llEmployeeNumber _
                                , radrsInitialAbsence!ABS_ST_DT _
                                , ldtReturnDate)
        Set lobjSchedExc = Nothing
        '//exclude any exceptions not tied to absence, either past the return to work date or exception code based
        '//TODO: Should this be using the absence end date or the return to work date (either way it needs to be date worked rather than charge date)
        '//the time is in this format "00.00.00"
        ldtAbsInitEndDate = DateTimeFrom(radrsInitialAbsence!ABS_END_DT, Format(radrsInitialAbsence!ABS_END_TM, "hh:mm"), True)
        If RecordsExist(radrsAbsenceTimeDetails) Then
            radrsAbsenceTimeDetails.MoveLast
            Do While Not radrsAbsenceTimeDetails.BOF
                If radrsAbsenceTimeDetails!Hours > 0 And radrsAbsenceTimeDetails!TIME_TYPE = "2" Then '//reported time
                    If DateTimeFrom(radrsAbsenceTimeDetails!CHARGE_DATE, radrsAbsenceTimeDetails!START_TIME) >= ldtAbsInitEndDate Then
                        If Trim(radrsAbsenceTimeDetails!EXCEPTION_CD) = "ST" Then '//TODO: This could include other exception codes
                            radrsAbsenceTimeDetails.Delete
                        End If
                    Else
                        Exit Do
                    End If
                End If
                radrsAbsenceTimeDetails.MovePrevious
            Loop
            radrsAbsenceTimeDetails.MoveFirst
        End If

    End If

    If (RecordsExist(radrsAbsenceDate) And _
        llEmployeeNumber > 0) Then
        Set ladrsDates = New ADODB.Recordset
        With ladrsDates
            .Fields.Append "EMP_NUM", adInteger, , ADOFIELDATTR
            .Fields.Append "START_DATE", adDate, , ADOFIELDATTR
            .Fields.Append "END_DATE", adDate, , ADOFIELDATTR

            .CursorLocation = adUseClient
            .CursorType = adOpenStatic
            .LockType = adLockBatchOptimistic
            .Open
        End With

        With radrsAbsenceDate
            .Sort = "ABS_DATE"
            .MoveFirst

            For llIndex = 1 To .RecordCount
                ladrsDates.AddNew
                ladrsDates!EMP_NUM = llEmployeeNumber
                ladrsDates!START_DATE = !ABS_DATE
                ladrsDates!END_DATE = !ABS_DATE
                ladrsDates.Update

                .MoveNext
            Next llIndex

            .MoveFirst
        End With

        ladrsDates.AddNew
        ladrsDates!EMP_NUM = llEmployeeNumber
        ladrsDates!START_DATE = ldtReturnDate
        ladrsDates!END_DATE = ldtReturnDate
        ladrsDates.Update

        Set lobjSchedule = CreateObject("busSchedule.cbsSchedule")
        Set ladrsSchedule = lobjSchedule.GABEmployeeSchedOnSpecificDates( _
                            ladrsDates _
                            , Nothing _
                            , False _
                            , trAdjBaseSchedWithExceptionsAndOvertime _
                            , trAllScheduleLevels)
        Set lobjSchedule = Nothing

        Set ladrsDates = Nothing

        If (RecordsExist(radrsAbsenceTimeDetails)) Then
            radrsAbsenceDate.Sort = "ABS_DATE"

            With radrsAbsenceTimeDetails
                .MoveFirst

                For llIndex = 1 To .RecordCount
                    If (ldtReturnDate <> !CHARGE_DATE) Then
                        radrsAbsenceDate.MoveFirst

                        radrsAbsenceDate.Find "ABS_DATE = #" & _
                            !CHARGE_DATE & "#"
                        If (radrsAbsenceDate.EOF) Then
                            .Delete adAffectCurrent
                            .Update
                        End If

                        radrsAbsenceDate.MoveFirst
                    End If

                    .MoveNext
                Next llIndex

                .MoveFirst
            End With
        End If
    End If

    If (RecordsExist(ladrsSchedule) And _
        RecordsExist(radrsInitialAbsence) And _
        RecordsExist(radrsAbsenceDate)) Then
        With radrsInitialAbsence
            ladrsSchedule.Filter = adFilterNone
            ladrsSchedule.Filter = "CHARGE_DATE = #" & !RTN_TO_WRK_DT & _
                                    "# AND SHIFT_NUM = 1"
            If Not (ladrsSchedule.EOF) Then
                If (ladrsSchedule!SHIFT_HOURS > 0) Then
                    If (Format(Trim(!RTN_TO_WRK_TM), "hh:mm") >= Trim(ladrsSchedule!SHIFT_START_TIME)) And _
                        (Format(Trim(!RTN_TO_WRK_TM), "hh:mm") <= Trim(ladrsSchedule!SHIFT_END_TIME)) Then
                        Select Case (Trim(ladrsSchedule!SHIFT_CODE))
                            Case "A"
                                !RTN_TO_WRK_SFT = "AM"
                            Case "P"
                                !RTN_TO_WRK_SFT = "PM"
                            Case "O"
                                !RTN_TO_WRK_SFT = "OWL"
                            Case "D"
                                !RTN_TO_WRK_SFT = "DAY"
                            Case "N"
                                !RTN_TO_WRK_SFT = "NIGHT"
                        End Select

                        .Update
                    End If
                End If
            End If

            ladrsSchedule.Filter = adFilterNone
            ladrsSchedule.Filter = "CHARGE_DATE = #" & !RTN_TO_WRK_DT & _
                                    "# AND SHIFT_NUM = 2"
            If Not (ladrsSchedule.EOF) Then
                If (ladrsSchedule!SHIFT_HOURS > 0) Then
                    If (Format(Trim(!RTN_TO_WRK_TM), "hh:mm") >= Trim(ladrsSchedule!SHIFT_START_TIME)) And _
                        (Format(Trim(!RTN_TO_WRK_TM), "hh:mm") <= Trim(ladrsSchedule!SHIFT_END_TIME)) Then
                        Select Case (Trim(ladrsSchedule!SHIFT_CODE))
                            Case "A"
                                !RTN_TO_WRK_SFT = "AM"
                            Case "P"
                                !RTN_TO_WRK_SFT = "PM"
                            Case "O"
                                !RTN_TO_WRK_SFT = "OWL"
                            Case "D"
                                !RTN_TO_WRK_SFT = "DAY"
                            Case "N"
                                !RTN_TO_WRK_SFT = "NIGHT"
                        End Select

                        .Update
                    End If
                End If
            End If
        End With

        With radrsAbsenceDate
            For llIndex = 1 To .RecordCount
                ladrsSchedule.Filter = adFilterNone
                ladrsSchedule.Filter = "CHARGE_DATE = #" & _
                                    !ABS_DATE & "#" & _
                                    " AND SHIFT_NUM = 1"
                If Not (ladrsSchedule.EOF) Then
                    !SHIFT_1_SCHED = ladrsSchedule!SCHED_CD

                    Select Case (Trim(ladrsSchedule!SHIFT_CODE))
                        Case "A"
                            !SHIFT_1_SHIFT = "AM"
                        Case "P"
                            !SHIFT_1_SHIFT = "PM"
                        Case "O"
                            !SHIFT_1_SHIFT = "OWL"
                        Case "D"
                            !SHIFT_1_SHIFT = "DAY"
                        Case "N"
                            !SHIFT_1_SHIFT = "NIGHT"
                    End Select

                    !SHIFT_1_START = ladrsSchedule!SHIFT_START_TIME
                    !SHIFT_1_END = ladrsSchedule!SHIFT_END_TIME

                    .Update
                End If

                ladrsSchedule.Filter = adFilterNone
                ladrsSchedule.Filter = "CHARGE_DATE = #" & _
                                    !ABS_DATE & "#" & _
                                    " AND SHIFT_NUM = 2"
                If Not (ladrsSchedule.EOF) Then
                    !SHIFT_2_SCHED = ladrsSchedule!SCHED_CD

                    Select Case (Trim(ladrsSchedule!SHIFT_CODE))
                        Case "A"
                            !SHIFT_2_SHIFT = "AM"
                        Case "P"
                            !SHIFT_2_SHIFT = "PM"
                        Case "O"
                            !SHIFT_2_SHIFT = "OWL"
                        Case "D"
                            !SHIFT_2_SHIFT = "DAY"
                        Case "N"
                            !SHIFT_2_SHIFT = "NIGHT"
                    End Select

                    !SHIFT_2_START = ladrsSchedule!SHIFT_START_TIME
                    !SHIFT_2_END = ladrsSchedule!SHIFT_END_TIME

                    .Update
                End If

                .MoveNext
            Next llIndex

            .MoveFirst
        End With
    End If

    Set ladrsSchedule = Nothing

    Exit Sub
GetAbsenceInformation_EH:
    Set lobjdbEAM = Nothing
    Set lobjEmployee = Nothing
    Set lobjSchedule = Nothing
    Set lobjSchedExc = Nothing

    Set ladrsSchedule = Nothing
    Set ladrsDates = Nothing

    TrsRaiseError MODULENAME _
        , "GetAbsenceInformation" _
        , , , , _
        , True

End Sub
'Private Function CreateGABRequestRecordset() As ADODB.Recordset
'On Error GoTo CreateGABRequestRecordset_EH
'Dim ladrs As ADODB.Recordset
'
'    Set CreateGABRequestRecordset = Nothing
'
'    Set ladrs = New ADODB.Recordset
'    With ladrs
'        .Fields.Append "EMP_NUM", adInteger, , ADOFIELDATTR
'        .Fields.Append "START_DATE", adDate, , ADOFIELDATTR
'        .Fields.Append "END_DATE", adDate, , ADOFIELDATTR
'        .CursorLocation = adUseClient
'        .CursorType = adOpenStatic
'        .LockType = adLockBatchOptimistic
'        .Open
'    End With
'
'    Set CreateGABRequestRecordset = ladrs
'
'    Set ladrs = Nothing
'
'    Exit Function
'CreateGABRequestRecordset_EH:
'    Set ladrs = Nothing
'
'    TrsRaiseError MODULENAME, "CreateGABRequestRecordset", , , , , True
'End Function
Public Function GetEmployeeProtfolios(ByVal vlEmployeeNumber As Long _
                , ByVal vsTRISUser As String _
                , ByVal vlTRISUserNumber As Long _
                , Optional ByVal vdtStartDate As Date = #1/1/1900# _
                , Optional ByVal vdtEndDate As Date = #12/31/9999# _
                ) As ADODB.Recordset
On Error GoTo GetEmployeeProtfolios_EH
Dim lobjdb As dbEAM.cdEAM

    Set GetEmployeeProtfolios = Nothing

    Set lobjdb = CreateObject("dbEAM.cdEAM")
    Set GetEmployeeProtfolios = lobjdb.GetEmployeePortfolios(vlEmployeeNumber _
                                , vdtStartDate _
                                , vdtEndDate _
                                , vlTRISUserNumber)
    Set lobjdb = Nothing

    Exit Function
GetEmployeeProtfolios_EH:
    Set lobjdb = Nothing

    TrsRaiseError MODULENAME _
        , "GetEmployeeProtfolios" _
        , , , , _
        , True

End Function
Public Sub GetProtfolioInformation(ByVal vlProtfolioNumber As Long _
                , ByVal vsTRISUser As String _
                , ByVal vlTRISUserNumber As Long _
                , ByRef radrsPortfolio As ADODB.Recordset _
                , ByRef radrsPortfolioDetail As ADODB.Recordset _
                , ByRef radrsAbsenceDate As ADODB.Recordset _
                , ByRef radrsInitialAbsence As ADODB.Recordset _
                , ByRef radrsAbsenceProvider As ADODB.Recordset _
                , ByRef radrsAbsenceContact As ADODB.Recordset _
                , ByRef radrsAbsenceReturnToWork As ADODB.Recordset _
                , ByRef radrsAbsenceTimeDetail As ADODB.Recordset _
                , Optional ByRef radrsAbsenceContactCallLog As ADODB.Recordset = Nothing)
On Error GoTo GetProtfolioInformation_EH
Dim lobjdb As dbEAM.cdEAM
Dim lobjSchedExc As busSchedException.cbsSchedExc
Dim ldtStartDate As Date, ldtEndDate As Date
Dim llEmployeeNumber As Long, lbTRISAdmin As Boolean

    Set lobjdb = CreateObject("dbEAM.cdEAM")

    Set radrsPortfolio = lobjdb.GetPortfolio(vlProtfolioNumber)
    Set radrsPortfolioDetail = lobjdb.GetPortfolioAbsences(vlProtfolioNumber)
    Set radrsAbsenceDate = lobjdb.GetPortfolioAbsenceDates(vlProtfolioNumber)
    Set radrsInitialAbsence = lobjdb.GetPortfolioInitialAbsence(vlProtfolioNumber)
    Set radrsAbsenceProvider = lobjdb.GetPortfolioAbsenceProviders(vlProtfolioNumber)
    Set radrsAbsenceContact = lobjdb.GetPortfolioAbsenceContacts(vlProtfolioNumber)
    Set radrsAbsenceReturnToWork = lobjdb.GetPortfolioAbsenceRtnToWork(vlProtfolioNumber)
    Set radrsAbsenceContactCallLog = lobjdb.GetPortfolioAbsenceCallLog(vlProtfolioNumber)

    Set lobjdb = Nothing

    If (RecordsExist(radrsAbsenceDate)) And _
        (RecordsExist(radrsInitialAbsence)) Then
        radrsInitialAbsence.MoveFirst
        llEmployeeNumber = radrsInitialAbsence!EMP_NUM

        With radrsAbsenceDate
            .MoveFirst
            ldtStartDate = !ABS_DATE
            .MoveLast
            ldtEndDate = !ABS_DATE

            .MoveFirst
        End With

        Set lobjSchedExc = CreateObject("busSchedException.cbsSchedExc")
        Set radrsAbsenceTimeDetail = lobjSchedExc.GetViewCalendarDetails( _
                                llEmployeeNumber _
                                , ldtStartDate _
                                , ldtEndDate)
        Set lobjSchedExc = Nothing
    End If

    Exit Sub
GetProtfolioInformation_EH:
    Set lobjdb = Nothing

    TrsRaiseError MODULENAME _
        , "GetProtfolioInformation" _
        , , , , _
        , True

End Sub
Public Function GetListOfProviders() As ADODB.Recordset
On Error GoTo GetListOfProviders_EH
Dim lobjdb As dbEAM.cdEAM

    Set GetListOfProviders = Nothing
    
    Set lobjdb = CreateObject("dbEAM.cdEAM")
    Set GetListOfProviders = lobjdb.GetListOfProviders
    Set lobjdb = Nothing

    Exit Function
GetListOfProviders_EH:
    Set lobjdb = Nothing

    TrsRaiseError MODULENAME _
        , "GetListOfProviders" _
        , , , , _
        , True

End Function

'calling this failed because HOME_PHONE_NUM does not exist in db.GetContactsDueByEmployees
'//---------------------------------------------------------------------------------------------------------//
'//not used, obsolete
'//---------------------------------------------------------------------------------------------------------//
Public Function GetContactsByEmployees(ByVal vlTRISUserEmployeeNumber As Long _
                , ByVal vsTRISUser As String _
                , ByVal vadrsEmployees As ADODB.Recordset _
                ) As ADODB.Recordset
TrsRaiseError MODULENAME, "GetContactsByEmployees", trErrObsoleteMethod, , , , True
'On Error GoTo GetContactsByEmployees_EH
'Dim lobjdb As dbEAM.cdEAM
'Dim lobjPS As dbPeopleSoft.cdPeopleSoft
'Dim ladrs As ADODB.Recordset, llArrEmps() As Long
'Dim llOuterIndex As Long, lsWhereClause As String
'Dim llInnerIndex As Long
'
'    Set GetContactsByEmployees = Nothing
'
'    lsWhereClause = ""
'
'    If (RecordsExist(vadrsEmployees)) Then
'        With vadrsEmployees
'            .MoveFirst
'
'            ReDim llArrEmps(.RecordCount - 1)
'
'            For llOuterIndex = 1 To .RecordCount
'                llArrEmps(llOuterIndex - 1) = !EMP_NUM
'
'                If (llOuterIndex = 1) Then
'                    lsWhereClause = CStr(!EMP_NUM)
'                Else
'                    lsWhereClause = lsWhereClause & ", " & CStr(!EMP_NUM)
'                End If
'
'                .MoveNext
'            Next llOuterIndex
'
'            .MoveFirst
'        End With
'
'        If (Trim(lsWhereClause) <> "") Then
'            Set lobjdb = CreateObject("dbEAM.cdEAM")
'            Set GetContactsByEmployees = lobjdb.GetContactsDueByEmployees(lsWhereClause)
'            Set lobjdb = Nothing
'
'            If (RecordsExist(GetContactsByEmployees)) Then
'                With GetContactsByEmployees
'                    .Filter = adFilterNone
'                    .Filter = "ILLINOIS_CMPY = 'N'"
'                    If Not (.EOF) Then
'                        .MoveFirst
'
'                        For llOuterIndex = 1 To .RecordCount
'                            !SL_POLICY_TXT = "(The Sick Leave policy may be discussed with the Employee on first follow-up call and as necessary thereafter.)" & vbCrLf & "# " & Chr(34) & "...Employee on sick leave, whether for personal illness or injury, are generally required to remain at their address of record unless hospitalized, keeping a medical appointment, or obtaining medication." & vbCrLf & "...Special cases which require that an employee on sick leave be away from his/her address of record for prescribed physical activity, therapy, etc. should be discussed with the employee's supervisor in advance." & vbCrLf & "# ...Violations of these requirements will be considered as abuse of sick leave and subject to disciplinary action, up to and including discharge." & Chr(34)
'                            !SL_POLICY_FROM = "From letter by D.A. Slagley dated 05/23/1989, Personal Activities While On Sick Leave"
'                            .Update
'
'                            .MoveNext
'                        Next llOuterIndex
'                    End If
'
'                    .Filter = adFilterNone
'                    .Filter = "ILLINOIS_CMPY = 'Y'"
'                    If Not (.EOF) Then
'                        .MoveFirst
'
'                        For llOuterIndex = 1 To .RecordCount
'                            !SL_POLICY_TXT = "Employees on sick leave, whether for personal illness or injury, are generally required to remain at their address of record unless hospitalized, keeping a medical appointment or obtaining medication."
'                            !SL_POLICY_FROM = ""
'                            .Update
'
'                            .MoveNext
'                        Next llOuterIndex
'                    End If
'
'                    .Filter = adFilterNone
'
'                    Set lobjPS = CreateObject("dbPeopleSoft.cdPeopleSoft")
'                    Set ladrs = lobjPS.GetPeopleSoftPhoneNumber(llArrEmps)
'                    Set lobjPS = Nothing
'
'                    Erase llArrEmps
'
'                    If (RecordsExist(ladrs)) Then
'                        ladrs.MoveFirst
'
'                        For llOuterIndex = 1 To ladrs.RecordCount
'                            .Filter = adFilterNone
'
'                            .Filter = "EMP_NUM = " & CStr(!EMP_NUM)
'                            If Not (.EOF) Then
'                                .MoveFirst
'
'                                For llInnerIndex = 1 To .RecordCount
'                                    !HOME_PHONE_NUM = ladrs!HOME_PHONE_NUM
'                                    .Update
'
'                                    .MoveNext
'                                Next llInnerIndex
'                            End If
'
'                            ladrs.MoveNext
'                        Next llOuterIndex
'                    End If
'
'                    Set ladrs = Nothing
'
'                    .Filter = adFilterNone
'                    .MoveFirst
'                End With
'            End If
'        End If
'    End If
'
'    Exit Function
'GetContactsByEmployees_EH:
'    Set lobjdb = Nothing
'    Set lobjPS = Nothing
'
'    Erase llArrEmps
'
'    Set ladrs = Nothing
'
'    TrsRaiseError MODULENAME _
'        , "GetContactsByEmployees" _
'        , , , , _
'        , True

End Function
Public Function GetReturnToWorkByEmployees(ByVal vlTRISUserEmployeeNumber As Long _
                , ByVal vsTRISUser As String _
                , ByVal vadrsEmployees As ADODB.Recordset _
                ) As ADODB.Recordset
On Error GoTo GetReturnToWorkByEmployees_EH
Dim lobjdb As dbEAM.cdEAM
Dim llIndex As Long, lsWhereClause As String

    Set GetReturnToWorkByEmployees = Nothing

    lsWhereClause = ""

    If (RecordsExist(vadrsEmployees)) Then
        With vadrsEmployees
            .MoveFirst

            For llIndex = 1 To .RecordCount
                If (llIndex = 1) Then
                    lsWhereClause = CStr(!EMP_NUM)
                Else
                    lsWhereClause = lsWhereClause & ", " & CStr(!EMP_NUM)
                End If

                .MoveNext
            Next llIndex

            .MoveFirst
        End With

        If (Trim(lsWhereClause) <> "") Then
            Set lobjdb = CreateObject("dbEAM.cdEAM")
            Set GetReturnToWorkByEmployees = lobjdb.GetReturnToWorksDueByEmployees(lsWhereClause)
            Set lobjdb = Nothing
        End If
    End If

    Exit Function
GetReturnToWorkByEmployees_EH:
    Set lobjdb = Nothing

    TrsRaiseError MODULENAME _
        , "GetReturnToWorkByEmployees" _
        , , , , _
        , True

End Function
Public Function GetFMLAAbsenceByEmployees(ByVal vlTRISUserEmployeeNumber As Long _
                , ByVal vsTRISUser As String _
                , ByVal vadrsEmployees As ADODB.Recordset _
                ) As ADODB.Recordset
On Error GoTo GetFMLAAbsenceByEmployees_EH
Dim lobjdb As dbEAM.cdEAM
Dim llIndex As Long, lsWhereClause As String

    Set GetFMLAAbsenceByEmployees = Nothing

    lsWhereClause = ""

    If (RecordsExist(vadrsEmployees)) Then
        With vadrsEmployees
            .MoveFirst

            For llIndex = 1 To .RecordCount
                If (llIndex = 1) Then
                    lsWhereClause = CStr(!EMP_NUM)
                Else
                    lsWhereClause = lsWhereClause & ", " & CStr(!EMP_NUM)
                End If

                .MoveNext
            Next llIndex

            .MoveFirst
        End With

        If (Trim(lsWhereClause) <> "") Then
            Set lobjdb = CreateObject("dbEAM.cdEAM")
            Set GetFMLAAbsenceByEmployees = lobjdb.GetFMLAAbsences(lsWhereClause)
            Set lobjdb = Nothing
        End If
    End If

    Exit Function
GetFMLAAbsenceByEmployees_EH:
    Set lobjdb = Nothing

    TrsRaiseError MODULENAME _
        , "GetFMLAAbsenceByEmployees" _
        , , , , _
        , True
End Function
Public Function GetContactsByScope(ByVal vlTRISUserEmployeeNumber As Long _
                , ByVal vsTRISUser As String _
                , ByVal vbActiveInformationOnly As Boolean _
                ) As ADODB.Recordset
On Error GoTo GetContactsByScope_EH
Dim lobjdb As dbEAM.cdEAM, llOuterIndex As Long

    Set GetContactsByScope = Nothing

    Set lobjdb = CreateObject("dbEAM.cdEAM")
    Set GetContactsByScope = lobjdb.GetContactsDueByScope2(vlTRISUserEmployeeNumber)
    Set lobjdb = Nothing

    If (RecordsExist(GetContactsByScope)) Then
        With GetContactsByScope
            .Filter = adFilterNone
            .Filter = "ILLINOIS_CMPY = 'N'"
            If Not (.EOF) Then
                .MoveFirst

                For llOuterIndex = 1 To .RecordCount
                    !SL_POLICY_TXT = "(The Sick Leave policy may be discussed with the Employee on first follow-up call and as necessary thereafter.)" & vbCrLf & "# " & Chr(34) & "...Employee on sick leave, whether for personal illness or injury, are generally required to remain at their address of record unless hospitalized, keeping a medical appointment, or obtaining medication." & vbCrLf & "...Special cases which require that an employee on sick leave be away from his/her address of record for prescribed physical activity, therapy, etc. should be discussed with the employee's supervisor in advance." & vbCrLf & "# ...Violations of these requirements will be considered as abuse of sick leave and subject to disciplinary action, up to and including discharge." & Chr(34)
                    !SL_POLICY_FROM = "From letter by D.A. Slagley dated 05/23/1989, Personal Activities While On Sick Leave"
                    .Update

                    .MoveNext
                Next llOuterIndex
            End If

            .Filter = adFilterNone
            .Filter = "ILLINOIS_CMPY = 'Y'"
            If Not (.EOF) Then
                .MoveFirst

                For llOuterIndex = 1 To .RecordCount
                    !SL_POLICY_TXT = "Employees on sick leave, whether for personal illness or injury, are generally required to remain at their address of record unless hospitalized, keeping a medical appointment or obtaining medication."
                    !SL_POLICY_FROM = ""
                    .Update

                    .MoveNext
                Next llOuterIndex
            End If

            .Filter = adFilterNone
            .MoveFirst
        End With
    End If

    Exit Function
GetContactsByScope_EH:
    Set lobjdb = Nothing

    TrsRaiseError MODULENAME _
        , "GetContactsByScope" _
        , , , , _
        , True
End Function
Public Function GetReturnToWorkByScope(ByVal vlTRISUserEmployeeNumber As Long _
                , ByVal vsTRISUser As String _
                , ByVal vbActiveInformationOnly As Boolean _
                ) As ADODB.Recordset
On Error GoTo GetReturnToWorkByScope_EH
Dim lobjdb As dbEAM.cdEAM

    Set GetReturnToWorkByScope = Nothing

    Set lobjdb = CreateObject("dbEAM.cdEAM")
    Set GetReturnToWorkByScope = lobjdb.GetReturnToWorksDueByScope2(vlTRISUserEmployeeNumber)
    Set lobjdb = Nothing

    Exit Function
GetReturnToWorkByScope_EH:
    Set lobjdb = Nothing

    TrsRaiseError MODULENAME _
        , "GetReturnToWorkByScope" _
        , , , , _
        , True

End Function
Public Function GetFMLAAbsenceByScope(ByVal vlTRISUserEmployeeNumber As Long _
                , ByVal vsTRISUser As String _
                ) As ADODB.Recordset
On Error GoTo GetFMLAAbsenceByScope_EH
Dim lobjdb As dbEAM.cdEAM

    Set GetFMLAAbsenceByScope = Nothing

    Exit Function
GetFMLAAbsenceByScope_EH:
    Set lobjdb = Nothing

    TrsRaiseError MODULENAME _
        , "GetFMLAAbsenceByScope" _
        , , , , _
        , True

End Function
Public Function GetEAMEmployeeInformation2(ByVal vlEmployeeNumber As Long _
                    , ByVal vsTRISUser As String _
                    , ByVal vlTRISUserNumber As Long _
                    , ByRef radrsRTWExemptExceptions As ADODB.Recordset _
                    ) As ADODB.Recordset
TrsRaiseError MODULENAME, "GetEAMEmployeeInformation2", trErrObsoleteMethod, , , , True
'On Error GoTo GetEAMEmployeeInformation2_EH
'Dim lobjCollPeriod As busCollectionPeriod.cbsCollectionPeriod
'Dim lobjEmployee As busEmployee.cbsEmployee
'Dim lobjdbRules As dbRules.cdRules
'Dim lobjPS As dbPeopleSoft.cdPeopleSoft
'Dim lobjdbEAM As dbEAM.cdEAM
'
'Dim ladrsRules As ADODB.Recordset
'Dim ladrs As ADODB.Recordset
'
'Dim liTimeIncrement As Integer, lbExecutive As Boolean
'Dim ldtLastProcessedDate As Date, lsHomePhone As String
'Dim liLocation As Integer, lsDepartment As String
'Dim lsPreviousDept As String, ldtDeptEffectiveDate As Date
'Dim lsUnion As String, lsExecContractIndicator As String
'Dim lsCompany As String, lsMgmtBenefitsPkg As String
'Dim lsUnionBenefitsPkg As String, lsEmpCmpy As Boolean
'Dim llSupervisorNumber As Long, lsSupervisorName As String
'Dim liEAMEmployeeLocation As Integer, lsEAMEmployeeLocationDesc As String
'
'    Set GetEAMEmployeeInformation2 = Nothing
'
'    ldtLastProcessedDate = #1/1/100#
'    Set lobjCollPeriod = CreateObject("busCollectionPeriod.cbsCollectionPeriod")
'    Set ladrs = lobjCollPeriod.GetCollectionPeriodStatusByEmployee( _
'                    vlEmployeeNumber, CDate("12/30/1899"), 0)
'    If (RecordsExist(ladrs)) Then
'        ldtLastProcessedDate = lobjCollPeriod.GetLatestProcessedPayPeriodEndDate( _
'                                    ladrs!PAYROLL_GROUP_CD)
'    End If
'    Set lobjCollPeriod = Nothing
'    Set ladrs = Nothing
'
'    Set lobjdbEAM = CreateObject("dbEAM.cdEAM")
'    Set ladrs = lobjdbEAM.GetAbsenceEmployeeInformation(vlEmployeeNumber)
'    Set lobjdbEAM = Nothing
'
'    llSupervisorNumber = 0
'    lsSupervisorName = ""
'    liEAMEmployeeLocation = 0
'    lsEAMEmployeeLocationDesc = ""
'    If (RecordsExist(ladrs)) Then
'        With ladrs
'            .MoveFirst
'            llSupervisorNumber = !SUPERVISOR_NUM
'            lsSupervisorName = Trim(!SUPERVISOR_NAME)
'            liEAMEmployeeLocation = !LOC_CD
'            lsEAMEmployeeLocationDesc = Trim(!LOC_DESC)
'        End With
'    End If
'    Set ladrs = Nothing
'
'
'    Set lobjPS = CreateObject("dbPeopleSoft.cdPeopleSoft")
'    Set ladrs = lobjPS.GetPeopleSoftPhoneNumber(vlEmployeeNumber)
'    Set lobjPS = Nothing
'
'    lsHomePhone = ""
'    If (RecordsExist(ladrs)) Then
'        lsHomePhone = Trim(ladrs!HOME_PHONE_NUM)
'    End If
'    Set ladrs = Nothing
'
'    liTimeIncrement = 6
'
'    Set lobjEmployee = CreateObject("busEmployee.cbsEmployee")
'    Set ladrs = lobjEmployee.GetBasicEmployeeInformation( _
'                                    vlEmployeeNumber)
'    If (RecordsExist(ladrs)) Then
'        lbExecutive = False
'        If (ladrs!EXEC_CONTRACT_IND = "E" Or _
'            ladrs!EXEC_CONTRACT_IND = "e") Then
'            lbExecutive = True
'        End If
'        liTimeIncrement = lobjEmployee.PartialHourMinutes(ladrs!DEPT_NUM _
'                            , lbExecutive _
'                            , ladrs!CMPY_CD)
'    End If
'    Set lobjEmployee = Nothing
'
'    liLocation = 0
'    lsUnionBenefitsPkg = ""
'    lsMgmtBenefitsPkg = ""
'    lsCompany = ""
'    lsExecContractIndicator = ""
'    lsDepartment = ""
'    lsPreviousDept = ""
'    ldtDeptEffectiveDate = #1/1/100#
'    lsUnion = ""
'
'    If (RecordsExist(ladrs)) Then
'        With ladrs
'            .MoveFirst
'            liLocation = !LOC_CD
'            If (Trim(lsHomePhone) = "") Then
'                lsHomePhone = !HOME_PHONE_NUM
'            End If
'            lsUnionBenefitsPkg = !UNION_BEN_PKG_IND
'            lsMgmtBenefitsPkg = !MGMT_BEN_PKG_IND
'            lsCompany = !CMPY_CD
'            lsExecContractIndicator = !EXEC_CONTRACT_IND
'            lsDepartment = !DEPT_NUM
'            lsPreviousDept = !PREV_DEPT_NUM
'            ldtDeptEffectiveDate = !DEPT_EFF_DATE
'            lsUnion = !UNION_CD
'        End With
'    End If
'    Set lobjEmployee = Nothing
'    Set ladrs = Nothing
'
'    'Get Auto Determine Case Rules
'    Set lobjdbRules = CreateObject("dbRules.cdRules")
'    Set ladrsRules = lobjdbRules.GetRule(302)
'    Set lobjdbRules = Nothing
'
'    Set ladrs = New ADODB.Recordset
'    With ladrs
'        .Fields.Append "CMPY_CD", adChar, 3, ADOFIELDATTR
'        .Fields.Append "DEPT_NUM", adChar, 3, ADOFIELDATTR
'        .Fields.Append "PREV_DEPT_NUM", adChar, 3, ADOFIELDATTR
'        .Fields.Append "DEPT_EFF_DATE", adDate, , ADOFIELDATTR
'        .Fields.Append "EXEC_CONTRACT_IND", adChar, 1, ADOFIELDATTR
'        .Fields.Append "HOME_PHONE_NUM", adChar, 15, ADOFIELDATTR
'        .Fields.Append "LOC_CD", adInteger, , ADOFIELDATTR
'        .Fields.Append "MGMT_BEN_PKG_IND", adChar, 4, ADOFIELDATTR
'        .Fields.Append "UNION_BEN_PKG_IND", adChar, 4, ADOFIELDATTR
'        .Fields.Append "UNION_CD", adChar, 3, ADOFIELDATTR
'
'        .Fields.Append "ILLINOIS_CMPY", adChar, 1, ADOFIELDATTR
'        .Fields.Append "AUTO_CASE", adChar, 1, ADOFIELDATTR
'        .Fields.Append "AUTO_CASE_TYPE", adChar, 1, ADOFIELDATTR
'        .Fields.Append "AUTO_CASE_MM_DD_IND", adChar, 1, ADOFIELDATTR
'        .Fields.Append "AUTO_CASE_DURATION", adInteger, , ADOFIELDATTR
'        .Fields.Append "LAST_PP_PROC_DATE", adDate, , ADOFIELDATTR
'        .Fields.Append "PS_URL", adVarChar, 4096, ADOFIELDATTR
'        .Fields.Append "FORM457_URL", adVarChar, 4096, ADOFIELDATTR
'        .Fields.Append "SL_POLICY_TXT", adVarChar, 4096, ADOFIELDATTR
'        .Fields.Append "SL_POLICY_FROM", adVarChar, 4096, ADOFIELDATTR
'        .Fields.Append "TIME_INCREMENT", adInteger, , ADOFIELDATTR
'        .Fields.Append "SUPERVISOR_NUM", adInteger, , ADOFIELDATTR
'        .Fields.Append "SUPERVISOR_NAME", adVarChar, 30, ADOFIELDATTR
'        .Fields.Append "LOC_DESC", adVarChar, 30, ADOFIELDATTR
'
'        .CursorLocation = adUseClient
'        .CursorType = adOpenStatic
'        .LockType = adLockBatchOptimistic
'        .Open
'
'        .AddNew
'
'        !CMPY_CD = lsCompany
'        !DEPT_NUM = lsDepartment
'        !PREV_DEPT_NUM = lsPreviousDept
'        !DEPT_EFF_DATE = ldtDeptEffectiveDate
'        Select Case (Trim(lsExecContractIndicator))
'            Case "E", "e"
'                !EXEC_CONTRACT_IND = "E"
'            Case "C", "c"
'                !EXEC_CONTRACT_IND = "C"
'        End Select
'        !HOME_PHONE_NUM = lsHomePhone
'        !LOC_CD = liLocation
'        !MGMT_BEN_PKG_IND = lsMgmtBenefitsPkg
'        !UNION_BEN_PKG_IND = lsUnionBenefitsPkg
'        !UNION_CD = lsUnion
'
'        !ILLINOIS_CMPY = "N"
'        !SL_POLICY_TXT = "The Sick Leave policy may be discussed with the Employee on first follow-up call and as necessary thereafter.)" & vbCrLf & "# " & Chr(34) & "...Employee on sick leave, whether for personal illness or injury, are generally required to remain at their address of record unless hospitalized, keeping a medical appointment, or obtaining medication." & vbCrLf & "...Special cases which require that an employee on sick leave be away from his/her address of record for prescribed physical activity, therapy, etc. should be discussed with the employee's supervisor in advance." & vbCrLf & "# ...Violations of these requirements will be considered as abuse of sick leave and subject to disciplinary action, up to and including discharge." & Chr(34)
'        !SL_POLICY_FROM = "From letter by D.A. Slagley dated 05/23/1989, Personal Activities While On Sick Leave"
'        Select Case (Trim(!CMPY_CD))
'            Case "ARG", "CIL", "CIS", "MV1", "MV3" _
'                , "AED", "CIC", "CIP", "GEN", "GMC" _
'                , "IHC", "IMS", "IPC"
'                !ILLINOIS_CMPY = "Y"
'                !SL_POLICY_TXT = "Employees on sick leave, whether for personal illness or injury, are generally required to remain at their address of record unless hospitalized, keeping a medical appointment or obtaining medication."
'                !SL_POLICY_FROM = ""
'        End Select
'
'        !AUTO_CASE = "N"
'        If (Trim(lsUnion) <> "") Then
'            If (RecordsExist(ladrsRules)) Then
'                ladrsRules.MoveFirst
'
'                ladrsRules.Find "RG_VALUE_CD LIKE '" & Trim(lsUnion) & "%'"
'                If Not (ladrsRules.EOF) Then
'                    !AUTO_CASE = "Y"
'
'                    If (InStr(1, ladrsRules!RG_VALUE_CD, Trim(lsUnion) & " .0.mm") > 0) Then
'                        !AUTO_CASE_TYPE = "S"
'                        !AUTO_CASE_MM_DD_IND = "M"
'                        !AUTO_CASE_DURATION = CInt(Right(Trim(ladrsRules!RG_VALUE_CD), 2))
'                    End If
'                    If (InStr(1, ladrsRules!RG_VALUE_CD, Trim(lsUnion) & " .0.dd") > 0) Then
'                        !AUTO_CASE_TYPE = "S"
'                        !AUTO_CASE_MM_DD_IND = "D"
'                        !AUTO_CASE_DURATION = CInt(Right(Trim(ladrsRules!RG_VALUE_CD), 2))
'                    End If
'                    If (InStr(1, ladrsRules!RG_VALUE_CD, Trim(lsUnion) & " .1.mm") > 0) Then
'                        !AUTO_CASE_TYPE = "C"
'                        !AUTO_CASE_MM_DD_IND = "M"
'                        !AUTO_CASE_DURATION = CInt(Right(Trim(ladrsRules!RG_VALUE_CD), 2))
'                    End If
'                    If (InStr(1, ladrsRules!RG_VALUE_CD, Trim(lsUnion) & " .1.dd") > 0) Then
'                        !AUTO_CASE_TYPE = "C"
'                        !AUTO_CASE_MM_DD_IND = "D"
'                        !AUTO_CASE_DURATION = CInt(Right(Trim(ladrsRules!RG_VALUE_CD), 2))
'                    End If
'                End If
'            End If
'        End If
'
'        !LAST_PP_PROC_DATE = ldtLastProcessedDate
'        !PS_URL = "https://hrprod1:7002/"
'        !FORM457_URL = "https://sunp9.ameren.com:80/idc/groups/public/@public015506/documents/document_other/gob019622.pdf"
'        !TIME_INCREMENT = liTimeIncrement
'
'        .Update
'    End With
'
'    'Get RTW Exempt Exceptions
'    Set lobjdbRules = CreateObject("dbRules.cdRules")
'    Set radrsRTWExemptExceptions = lobjdbRules.GetRule(304)
'    Set lobjdbRules = Nothing
'
'    Set lobjCollPeriod = Nothing
'    Set lobjEmployee = Nothing
'    Set lobjdbRules = Nothing
'    Set ladrsRules = Nothing
'    Set lobjdbEAM = Nothing
'
'    Set GetEAMEmployeeInformation2 = ladrs
'
'    Set ladrs = Nothing
'
'    Exit Function
'GetEAMEmployeeInformation2_EH:
'    Set lobjCollPeriod = Nothing
'    Set lobjEmployee = Nothing
'    Set lobjdbRules = Nothing
'    Set lobjPS = Nothing
'    Set lobjdbEAM = Nothing
'
'    Set ladrsRules = Nothing
'    Set ladrs = Nothing
'
'    TrsRaiseError MODULENAME _
'        , "GetEAMEmployeeInformation2" _
'        , , , , _
'        , True

End Function
Public Function GetDailyEmployeeAbsences(ByVal vlTRISUserNumber As Long _
                    , ByVal vsTRISUser As String _
                    , ByVal vbActiveInformationOnly As Boolean _
                    ) As ADODB.Recordset
On Error GoTo GetDailyEmployeeAbsences_EH
Dim lobjdbEAM As dbEAM.cdEAM

    Set GetDailyEmployeeAbsences = Nothing

    Set lobjdbEAM = CreateObject("dbEAM.cdEAM")
    Set GetDailyEmployeeAbsences = lobjdbEAM.GetDailyEmployeeAbsences2(vlTRISUserNumber)
    Set lobjdbEAM = Nothing

    Exit Function
GetDailyEmployeeAbsences_EH:
    Set lobjdbEAM = Nothing

    TrsRaiseError MODULENAME, "GetDailyEmployeeAbsences", , , , , True
End Function
'END PREM

'Begin Prem TRIS 5.9
Public Function GetEAMEmployeeInformation3(ByVal vlEmployeeNumber As Long _
                    , ByVal vsTRISUser As String _
                    , ByVal vlTRISUserNumber As Long _
                    , ByRef radrsRTWExemptExceptions As ADODB.Recordset _
                    , ByRef radrsWCExceptions As ADODB.Recordset _
                    ) As ADODB.Recordset
On Error GoTo GetEAMEmployeeInformation3_EH
Dim lobjdbRules As dbRules.cdRules
Dim lobjPS As dbPeopleSoft.cdPeopleSoft
Dim lobjdbEAM As dbEAM.cdEAM

Dim ladrsRules As ADODB.Recordset
Dim ladrs As ADODB.Recordset

Dim liTimeIncrement As Integer, lbExecutive As Boolean
Dim ldtLastProcessedDate As Date, lsHomePhone As String
Dim liLocation As Integer, lsDepartment As String
Dim lsPreviousDept As String, ldtDeptEffectiveDate As Date
Dim lsUnion As String, lsExecContractIndicator As String
Dim lsCompany As String, lsMgmtBenefitsPkg As String
Dim lsUnionBenefitsPkg As String, lsEmpCmpy As Boolean
Dim llSupervisorNumber As Long, lsSupervisorName As String
Dim liEAMEmployeeLocation As Integer, lsEAMEmployeeLocationDesc As String

    Set GetEAMEmployeeInformation3 = Nothing

    'Get RTW Exempt Exceptions
    Set lobjdbRules = CreateObject("dbRules.cdRules")
    Set radrsRTWExemptExceptions = lobjdbRules.GetRule(304)
    Set lobjdbRules = Nothing

    'Get Workman's Compensation Exceptions
    Set lobjdbRules = CreateObject("dbRules.cdRules")
    Set radrsWCExceptions = lobjdbRules.GetRule(308)
    Set lobjdbRules = Nothing

    'Get PeopleSoft Phone
    Set lobjPS = CreateObject("dbPeopleSoft.cdPeopleSoft")
    Set ladrs = lobjPS.GetPeopleSoftPhoneNumber(vlEmployeeNumber)
    Set lobjPS = Nothing
    'Extract phone number
    lsHomePhone = ""
    If (RecordsExist(ladrs)) Then
        lsHomePhone = Trim(ladrs!HOME_PHONE_NUM)
    End If
    Set ladrs = Nothing

    'Get Employee Information
    Set lobjdbEAM = CreateObject("dbEAM.cdEAM")
    Set ladrs = lobjdbEAM.GetAbsenceEmployeeInformation3(vlEmployeeNumber)
    Set lobjdbEAM = Nothing
    'Update Phone Number
    If (RecordsExist(ladrs)) Then
        With ladrs
            .MoveFirst

            If (Trim(lsHomePhone) <> "") Then
                !HOME_PHONE_NUM = lsHomePhone
            End If

            If (Trim(!ILLINOIS_CMPY) = "Y") Then
                !SL_POLICY_TXT = "The Sick Leave policy may be discussed with the Employee on first contact and as necessary thereafter." & _
                    vbCrLf & "# " & Chr(34) & _
                    "Employees on sick leave, whether for personal illness or injury, are generally required to remain at their address of record unless hospitalized, keeping a medical appointment or obtaining medication."
                !SL_POLICY_FROM = ""
            Else
                !SL_POLICY_TXT = "The Sick Leave policy may be discussed with the Employee on first contact and as necessary thereafter." & _
                    vbCrLf & "# " & Chr(34) & _
                    "...Employee on sick leave, whether for personal illness or injury, are generally required to remain at their address of record unless hospitalized, keeping a medical appointment, or obtaining medication." & _
                    vbCrLf & _
                    "...Special cases which require that an employee on sick leave be away from his/her address of record for prescribed physical activity, therapy, etc. should be discussed with the employee's supervisor in advance." & _
                    vbCrLf & _
                    "# ...Violations of these requirements will be considered as abuse of sick leave and subject to disciplinary action, up to and including discharge." & _
                    Chr(34)
                !SL_POLICY_FROM = "From letter by D.A. Slagley dated 05/23/1989, Personal Activities While On Sick Leave"
            End If

            .Update
        End With
    End If

    Set GetEAMEmployeeInformation3 = ladrs
    Set ladrs = Nothing

    Exit Function
GetEAMEmployeeInformation3_EH:
    Set lobjdbRules = Nothing
    Set lobjPS = Nothing
    Set lobjdbEAM = Nothing

    Set ladrs = Nothing

    TrsRaiseError MODULENAME _
        , "GetEAMEmployeeInformation3" _
        , , , , _
        , True

End Function
'obsolete
'Public Sub GetAbsenceInformation2(ByVal vlAbsenceNumber As Long _
'                , ByVal vsTRISUser As String _
'                , ByRef radrsAbsenceDate As ADODB.Recordset _
'                , ByRef radrsInitialAbsence As ADODB.Recordset _
'                , ByRef radrsAbsenceProvider As ADODB.Recordset _
'                , ByRef radrsAbsenceContact As ADODB.Recordset _
'                , ByRef radrsAbsenceReturnToWork As ADODB.Recordset _
'                , ByRef radrsAbsenceContactCallLog As ADODB.Recordset _
'                , ByRef radrsWorkmanCompensation As ADODB.Recordset _
'                , ByRef radrsScheduledExceptions As ADODB.Recordset _
'                , ByRef radrsAbsenceTimeDetails As ADODB.Recordset _
'                , ByRef radrsCaseNumbers As ADODB.Recordset)
'On Error GoTo GetAbsenceInformation2_EH
'Dim lobjdbEAM As dbEAM.cdEAM
'Dim lobjSchedule As busSchedule.cbsSchedule
'Dim lobjEmployee As busEmployee.cbsEmployee
'Dim lobjSchedExc As busSchedException.cbsSchedExc
'
'Dim ladrsSchedule As ADODB.Recordset
'Dim ladrsDates As ADODB.Recordset
'Dim llEmployeeNumber As Long, ldtReturnDate As Date
'Dim llIndex As Long
'
'    Set lobjdbEAM = CreateObject("dbEAM.cdEAM")
'
'    Set radrsAbsenceDate = lobjdbEAM.GetAbsenceDatesByAbsence(vlAbsenceNumber)
'    Set radrsInitialAbsence = lobjdbEAM.GetInitialAbsenceByAbsence(vlAbsenceNumber)
'    Set radrsAbsenceProvider = lobjdbEAM.GetAbsenceProviderByAbsence(vlAbsenceNumber)
'    Set radrsAbsenceContact = lobjdbEAM.GetContactByAbsence(vlAbsenceNumber)
'    Set radrsAbsenceReturnToWork = lobjdbEAM.GetReturnToWorkByAbsence(vlAbsenceNumber)
'    Set radrsAbsenceContactCallLog = lobjdbEAM.GetContactCallLogByAbsence(vlAbsenceNumber)
'    Set radrsWorkmanCompensation = lobjdbEAM.GetWorkmanCompensation(vlAbsenceNumber)
'
'    Set lobjdbEAM = Nothing
'
'    If (RecordsExist(radrsInitialAbsence)) Then
'        radrsInitialAbsence.MoveFirst
'
'        llEmployeeNumber = radrsInitialAbsence!EMP_NUM
'        ldtReturnDate = radrsInitialAbsence!RTN_TO_WRK_DT
'        If (ldtReturnDate < radrsInitialAbsence!ABS_ST_DT) Then
'            ldtReturnDate = DateAdd("d", 1, radrsInitialAbsence!ABS_ST_DT)
'        End If
'
'        Set lobjEmployee = CreateObject("busEmployee.cbsEmployee")
'        Set radrsCaseNumbers = lobjEmployee.GetCaseGroupsAndNumbers( _
'                                llEmployeeNumber _
'                                , radrsInitialAbsence!ABS_ST_DT _
'                                , ldtReturnDate)
'        Set lobjEmployee = Nothing
'
'        Set lobjSchedExc = CreateObject("busSchedException.cbsSchedExc")
'        Set radrsScheduledExceptions = lobjSchedExc.GetScheduledExceptions( _
'                                llEmployeeNumber _
'                                , radrsInitialAbsence!ABS_ST_DT _
'                                , ldtReturnDate)
'
'        Set radrsAbsenceTimeDetails = lobjSchedExc.GetViewCalendarDetails( _
'                                llEmployeeNumber _
'                                , radrsInitialAbsence!ABS_ST_DT _
'                                , ldtReturnDate)
'        Set lobjSchedExc = Nothing
'    End If
'
'    If (RecordsExist(radrsAbsenceDate) And _
'        llEmployeeNumber > 0) Then
'        Set ladrsDates = New ADODB.Recordset
'        With ladrsDates
'            .Fields.Append "EMP_NUM", adInteger, , ADOFIELDATTR
'            .Fields.Append "START_DATE", adDate, , ADOFIELDATTR
'            .Fields.Append "END_DATE", adDate, , ADOFIELDATTR
'
'            .CursorLocation = adUseClient
'            .CursorType = adOpenStatic
'            .LockType = adLockBatchOptimistic
'            .Open
'        End With
'
'        With radrsAbsenceDate
'            .Sort = "ABS_DATE"
'            .MoveFirst
'
'            For llIndex = 1 To .RecordCount
'                ladrsDates.AddNew
'                ladrsDates!EMP_NUM = llEmployeeNumber
'                ladrsDates!START_DATE = !ABS_DATE
'                ladrsDates!END_DATE = !ABS_DATE
'                ladrsDates.Update
'
'                .MoveNext
'            Next llIndex
'
'            .MoveFirst
'        End With
'
'        ladrsDates.AddNew
'        ladrsDates!EMP_NUM = llEmployeeNumber
'        ladrsDates!START_DATE = ldtReturnDate
'        ladrsDates!END_DATE = ldtReturnDate
'        ladrsDates.Update
'
'        Set lobjSchedule = CreateObject("busSchedule.cbsSchedule")
'        Set ladrsSchedule = lobjSchedule.GABEmployeeSchedOnSpecificDates( _
'                            ladrsDates _
'                            , Nothing _
'                            , False _
'                            , trAdjBaseSchedWithExceptionsAndOvertime _
'                            , trAllScheduleLevels)
'        Set lobjSchedule = Nothing
'
'        Set ladrsDates = Nothing
'
'        If (RecordsExist(radrsAbsenceTimeDetails)) Then
'            radrsAbsenceDate.Sort = "ABS_DATE"
'
'            With radrsAbsenceTimeDetails
'                .MoveFirst
'
'                For llIndex = 1 To .RecordCount
'                    If (ldtReturnDate <> !CHARGE_DATE) Then
'                        radrsAbsenceDate.MoveFirst
'
'                        radrsAbsenceDate.Find "ABS_DATE = #" & _
'                            !CHARGE_DATE & "#"
'                        If (radrsAbsenceDate.EOF) Then
'                            .Delete adAffectCurrent
'                            .Update
'                        End If
'
'                        radrsAbsenceDate.MoveFirst
'                    End If
'
'                    .MoveNext
'                Next llIndex
'
'                .MoveFirst
'            End With
'        End If
'    End If
'
'    If (RecordsExist(ladrsSchedule) And _
'        RecordsExist(radrsInitialAbsence) And _
'        RecordsExist(radrsAbsenceDate)) Then
'        With radrsInitialAbsence
'            ladrsSchedule.Filter = adFilterNone
'            ladrsSchedule.Filter = "CHARGE_DATE = #" & !RTN_TO_WRK_DT & _
'                                    "# AND SHIFT_NUM = 1"
'            If Not (ladrsSchedule.EOF) Then
'                If (ladrsSchedule!SHIFT_HOURS > 0) Then
'                    If (Format(Trim(!RTN_TO_WRK_TM), "hh:mm") >= Trim(ladrsSchedule!SHIFT_START_TIME)) And _
'                        (Format(Trim(!RTN_TO_WRK_TM), "hh:mm") <= Trim(ladrsSchedule!SHIFT_END_TIME)) Then
'                        Select Case (Trim(ladrsSchedule!SHIFT_CODE))
'                            Case "A"
'                                !RTN_TO_WRK_SFT = "AM"
'                            Case "P"
'                                !RTN_TO_WRK_SFT = "PM"
'                            Case "O"
'                                !RTN_TO_WRK_SFT = "OWL"
'                            Case "D"
'                                !RTN_TO_WRK_SFT = "DAY"
'                            Case "N"
'                                !RTN_TO_WRK_SFT = "NIGHT"
'                        End Select
'
'                        .Update
'                    End If
'                End If
'            End If
'
'            ladrsSchedule.Filter = adFilterNone
'            ladrsSchedule.Filter = "CHARGE_DATE = #" & !RTN_TO_WRK_DT & _
'                                    "# AND SHIFT_NUM = 2"
'            If Not (ladrsSchedule.EOF) Then
'                If (ladrsSchedule!SHIFT_HOURS > 0) Then
'                    If (Format(Trim(!RTN_TO_WRK_TM), "hh:mm") >= Trim(ladrsSchedule!SHIFT_START_TIME)) And _
'                        (Format(Trim(!RTN_TO_WRK_TM), "hh:mm") <= Trim(ladrsSchedule!SHIFT_END_TIME)) Then
'                        Select Case (Trim(ladrsSchedule!SHIFT_CODE))
'                            Case "A"
'                                !RTN_TO_WRK_SFT = "AM"
'                            Case "P"
'                                !RTN_TO_WRK_SFT = "PM"
'                            Case "O"
'                                !RTN_TO_WRK_SFT = "OWL"
'                            Case "D"
'                                !RTN_TO_WRK_SFT = "DAY"
'                            Case "N"
'                                !RTN_TO_WRK_SFT = "NIGHT"
'                        End Select
'
'                        .Update
'                    End If
'                End If
'            End If
'        End With
'
'        With radrsAbsenceDate
'            For llIndex = 1 To .RecordCount
'                ladrsSchedule.Filter = adFilterNone
'                ladrsSchedule.Filter = "CHARGE_DATE = #" & _
'                                    !ABS_DATE & "#" & _
'                                    " AND SHIFT_NUM = 1"
'                If Not (ladrsSchedule.EOF) Then
'                    !SHIFT_1_SCHED = ladrsSchedule!SCHED_CD
'
'                    Select Case (Trim(ladrsSchedule!SHIFT_CODE))
'                        Case "A"
'                            !SHIFT_1_SHIFT = "AM"
'                        Case "P"
'                            !SHIFT_1_SHIFT = "PM"
'                        Case "O"
'                            !SHIFT_1_SHIFT = "OWL"
'                        Case "D"
'                            !SHIFT_1_SHIFT = "DAY"
'                        Case "N"
'                            !SHIFT_1_SHIFT = "NIGHT"
'                    End Select
'
'                    !SHIFT_1_START = ladrsSchedule!SHIFT_START_TIME
'                    !SHIFT_1_END = ladrsSchedule!SHIFT_END_TIME
'
'                    .Update
'                End If
'
'                ladrsSchedule.Filter = adFilterNone
'                ladrsSchedule.Filter = "CHARGE_DATE = #" & _
'                                    !ABS_DATE & "#" & _
'                                    " AND SHIFT_NUM = 2"
'                If Not (ladrsSchedule.EOF) Then
'                    !SHIFT_2_SCHED = ladrsSchedule!SCHED_CD
'
'                    Select Case (Trim(ladrsSchedule!SHIFT_CODE))
'                        Case "A"
'                            !SHIFT_2_SHIFT = "AM"
'                        Case "P"
'                            !SHIFT_2_SHIFT = "PM"
'                        Case "O"
'                            !SHIFT_2_SHIFT = "OWL"
'                        Case "D"
'                            !SHIFT_2_SHIFT = "DAY"
'                        Case "N"
'                            !SHIFT_2_SHIFT = "NIGHT"
'                    End Select
'
'                    !SHIFT_2_START = ladrsSchedule!SHIFT_START_TIME
'                    !SHIFT_2_END = ladrsSchedule!SHIFT_END_TIME
'
'                    .Update
'                End If
'
'                .MoveNext
'            Next llIndex
'
'            .MoveFirst
'        End With
'    End If
'
'    Set ladrsSchedule = Nothing
'
'    Exit Sub
'GetAbsenceInformation2_EH:
'    Set lobjdbEAM = Nothing
'    Set lobjEmployee = Nothing
'    Set lobjSchedule = Nothing
'    Set lobjSchedExc = Nothing
'
'    Set ladrsSchedule = Nothing
'    Set ladrsDates = Nothing
'
'    TrsRaiseError MODULENAME _
'        , "GetAbsenceInformation2" _
'        , , , , _
'        , True
'
'End Sub
'End Prem TRIS 5.9
