VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "cdEAMTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Project:   Post TRIS 5.5 Employee Absence Management
'Author:    Prem Chander
'Date:      August 2005
'Description:

Option Explicit
Private Const MODULENAME As String = "dbEAM.cdEAMTrans"
Public Function dbProcessEAMRequest(ByVal vlEmployeeNumber As Long _
                , ByVal vadrsAbsenceDates As ADODB.Recordset _
                , ByVal vadrsInitialAbsence As ADODB.Recordset _
                , ByVal vadrsAbsenceProvider As ADODB.Recordset _
                , ByVal vadrsAbsenceContact As ADODB.Recordset _
                , ByVal vadrsAbsenceReturnToWork As ADODB.Recordset _
                , ByVal vadrsAbsenceCallLog As ADODB.Recordset _
                , ByVal vlTRISUserNumber As Long _
                , ByVal vsTRISUserID As String _
                , ByRef rlAbsenceNumber As Long _
                , ByRef radrsBRV As ADODB.Recordset _
                , Optional ByVal vadrsPortfolio As ADODB.Recordset = Nothing _
                , Optional ByVal vadrsPortfolioDetail As ADODB.Recordset = Nothing _
                ) As Boolean
On Error GoTo dbProcessEAMRequest_EH
'ADODB Connection Object
Dim ladcnConnection As ADODB.Connection
'Recordset
Dim llRecordIndex As Long, llFilterIndex As Long
'SQL
Dim lsSQL As String
'SEQ_NUM
Dim lobjUtility As utilDBDataServices.cuDBDataServices
Dim llAbsenceNumber As Long, llProviderNumber As Long
Dim llContactNumber As Long
'Special
Dim lsOverride As String, llPortfolioNumber As Long
', ByVal vadrsAbsenceExceptions As ADODB.Recordset

    dbProcessEAMRequest = False

    lsSQL = ""

    If (RecordsExist(vadrsInitialAbsence)) Then
        With vadrsInitialAbsence
            .Filter = adFilterNone
            .Filter = "RECORD_STATUS = 'D'"
            If Not (.EOF) Then
                .MoveFirst

                For llRecordIndex = 1 To .RecordCount
                    lsSQL = lsSQL & _
                        " DELETE FROM METHDBA.TTR_ABS_PRTFL_DTL" & _
                        " WHERE ABS_NUM = " & CStr(!ABS_NUM) & ";"

                    lsSQL = lsSQL & _
                        " DELETE FROM METHDBA.TTR_ABS_DATE" & _
                        " WHERE ABS_NUM = " & CStr(!ABS_NUM) & ";"

                    lsSQL = lsSQL & _
                        " DELETE FROM METHDBA.TTR_ABS_RET_WRK" & _
                        " WHERE ABS_NUM = " & CStr(!ABS_NUM) & ";"

                    lsSQL = lsSQL & _
                        " DELETE FROM METHDBA.TTR_ABS_CALL_LOG" & _
                        " WHERE ABS_NUM = " & CStr(!ABS_NUM) & ";"

                    lsSQL = lsSQL & _
                        " DELETE FROM METHDBA.TTR_ABS_CONTACT" & _
                        " WHERE ABS_NUM = " & CStr(!ABS_NUM) & ";"

                    lsSQL = lsSQL & _
                        " DELETE FROM METHDBA.TTR_ABS_PROVIDER" & _
                        " WHERE ABS_NUM = " & CStr(!ABS_NUM) & ";"

                    lsSQL = lsSQL & _
                        " DELETE FROM METHDBA.TTR_ABS_INIT" & _
                        " WHERE ABS_NUM = " & CStr(!ABS_NUM) & ";"

                    .MoveNext
                Next llRecordIndex

                .Filter = adFilterNone
                GoTo EXITNORMALLY
            End If

            .Filter = adFilterNone
            .Filter = "RECORD_STATUS = 'I'"
            If Not (.EOF) Then
                .MoveFirst

                For llRecordIndex = 1 To .RecordCount
                    Set lobjUtility = CreateObject( _
                                        "utilDBDataServices.cuDBDataServices")
                    llAbsenceNumber = lobjUtility.GetNextSequenceNumber("EA")
                    Set lobjUtility = Nothing

                    AppendBusinessError radrsBRV _
                        , trInfEAM _
                        , mstrCBUSINESSINF _
                        , "A new absence " & CStr(llAbsenceNumber) & " was created." _
                        , "dbProcessEAMRequest"

                    If (llRecordIndex = 1) Then
                        rlAbsenceNumber = llAbsenceNumber
                    End If

                    'NOTE: It is "Null" not Empty String or VBNULL
                    'Reason for this special treatment: FOREIGN KEY FTRABI4

                    lsOverride = "Null"
                    If Not (IsNull(!OVER_EXC_CD)) Then
                        If (Trim(!OVER_EXC_CD) <> "") Then
                            lsOverride = "'" & Trim(!OVER_EXC_CD) & "'"
                        End If
                    End If

                    lsSQL = lsSQL & _
                        " INSERT INTO METHDBA.TTR_ABS_INIT" & _
                        " (ABS_NUM, EMP_NUM" & _
                        ", ABS_ST_TS, ABS_END_TS" & _
                        ", ABS_LOC_CD, ABS_CALLIN_TS" & _
                        ", EMP_PHONE, EMP_LOC_CD" & _
                        ", EXCEPTION_CD, OVER_EXC_CD" & _
                        ", RPTNG_EMP_NUM, RTN_TO_WRK_TS" & _
                        ", FMLA_REASON_CD, ABS_PRIM_CONT_EMP" & _
                        ", CREATED_EMP_NUM, CREATED_TIMESTAMP" & _
                        ", LAST_MOD_EMP_NUM, LAST_MOD_TIMESTAMP" & _
                        ", ABS_RPTR_TYPE, ABS_RPTR_TXT" & _
                        ", ADDL_TXT, ABS_RPT_TS" & _
                        ", RTW_NOT_REQD_IND, ABS_NOTIFY_WG_NUM" & _
                        ", PRVDR_NM, PRVDR_PHN_NBR" & _
                        ", FOLW_UP_DT)" & _
                        " VALUES"
                    lsSQL = lsSQL & _
                        " (" & CStr(llAbsenceNumber) & ", " & CStr(vlEmployeeNumber) & _
                        ", '" & Format(!ABS_ST_DT & " " & !ABS_ST_TM, "YYYY-MM-DD-hh.mm.ss.000000") & "'" & _
                        ", '" & Format(!ABS_END_DT & " " & !ABS_END_TM, "YYYY-MM-DD-hh.mm.ss.000000") & "'" & _
                        ", " & CStr(!ABS_LOC_CD) & ", '" & Format(!ABS_CALLIN_DT & " " & !ABS_CALLIN_TM, "YYYY-MM-DD-hh.mm.ss.000000") & "'" & _
                        ", '" & Trim(!EMP_PHONE) & "', " & CStr(!EMP_LOC_CD) & _
                        ", '" & Trim(!EXCEPTION_CD) & "', " & lsOverride & _
                        ", " & CStr(!RPTNG_EMP_NUM) & _
                        ", '" & Format(!RTN_TO_WRK_DT & " " & !RTN_TO_WRK_TM, "YYYY-MM-DD-hh.mm.ss.000000") & "'" & _
                        ", '" & Trim(!FMLA_REASON) & "', " & CStr(!ABS_PRIM_CONT_EMP) & _
                        ", " & CStr(vlTRISUserNumber) & ", CURRENT TIMESTAMP" & _
                        ", " & CStr(vlTRISUserNumber) & ", CURRENT TIMESTAMP" & _
                        ", '" & Trim(!ABS_RPTR_TYPE) & "', '" & Trim(!ABS_RPTR_TXT) & "'" & _
                        ", '" & Trim(!ADDL_TXT) & "', '" & Format(!ABS_RPT_TS, "YYYY-MM-DD-hh.mm.ss.000000") & "'" & _
                        ", '" & Trim(!RTW_NOT_REQD_IND) & "', " & CStr(!ABS_NOTIFY_WG_NUM) & _
                        ", '" & Trim(!PRVDR_NM) & "', '" & Trim(!PRVDR_PHN_NBR) & "'" & _
                        ", '" & Format(Trim(!FOLW_UP_DT), "MM/DD/YYYY") & "');"

                    If (RecordsExist(vadrsAbsenceDates)) Then
                        With vadrsAbsenceDates
                            .Filter = adFilterNone
                            .Filter = "ABS_NUM = " & _
                                    CStr(vadrsInitialAbsence!ABS_NUM)
                            If Not (.EOF) Then
                                .MoveFirst

                                For llFilterIndex = 1 To .RecordCount
                                    !ABS_NUM = llAbsenceNumber
                                    .Update

                                    .MoveNext
                                Next llFilterIndex
                            End If

                            .Filter = adFilterNone
                        End With
                    End If

                    If (RecordsExist(vadrsAbsenceProvider)) Then
                        With vadrsAbsenceProvider
                            .Filter = adFilterNone
                            .Filter = "ABS_NUM = " & _
                                    CStr(vadrsInitialAbsence!ABS_NUM)
                            If Not (.EOF) Then
                                .MoveFirst

                                For llFilterIndex = 1 To .RecordCount
                                    !ABS_NUM = llAbsenceNumber
                                    .Update

                                    .MoveNext
                                Next llFilterIndex
                            End If

                            .Filter = adFilterNone
                        End With
                    End If

                    If (RecordsExist(vadrsAbsenceContact)) Then
                        With vadrsAbsenceContact
                            .Filter = adFilterNone
                            .Filter = "ABS_NUM = " & _
                                    CStr(vadrsInitialAbsence!ABS_NUM)
                            If Not (.EOF) Then
                                .MoveFirst

                                For llFilterIndex = 1 To .RecordCount
                                    !ABS_NUM = llAbsenceNumber
                                    .Update

                                    .MoveNext
                                Next llFilterIndex
                            End If

                            .Filter = adFilterNone
                        End With
                    End If

                    If (RecordsExist(vadrsAbsenceReturnToWork)) Then
                        With vadrsAbsenceReturnToWork
                            .Filter = adFilterNone
                            .Filter = "ABS_NUM = " & _
                                    CStr(vadrsInitialAbsence!ABS_NUM)
                            If Not (.EOF) Then
                                .MoveFirst

                                For llFilterIndex = 1 To .RecordCount
                                    !ABS_NUM = llAbsenceNumber
                                    .Update

                                    .MoveNext
                                Next llFilterIndex
                            End If

                            .Filter = adFilterNone
                        End With
                    End If

                    If (RecordsExist(vadrsPortfolioDetail)) Then
                        With vadrsPortfolioDetail
                            .Filter = adFilterNone
                            .Filter = "ABS_NUM = " & _
                                    CStr(vadrsInitialAbsence!ABS_NUM)
                            If Not (.EOF) Then
                                .MoveFirst

                                For llFilterIndex = 1 To .RecordCount
                                    !ABS_NUM = llAbsenceNumber
                                    .Update

                                    .MoveNext
                                Next llFilterIndex
                            End If

                            .Filter = adFilterNone
                        End With
                    End If

                    If (RecordsExist(vadrsAbsenceCallLog)) Then
                        With vadrsAbsenceCallLog
                            .Filter = adFilterNone
                            .Filter = "ABS_NUM = " & _
                                    CStr(vadrsInitialAbsence!ABS_NUM) & _
                                    " AND RECORD_STATUS = 'I'"
                            If Not (.EOF) Then
                                .MoveFirst

                                For llFilterIndex = 1 To .RecordCount
                                    !ABS_NUM = llAbsenceNumber
                                    !CALL_NUM = llFilterIndex
                                    .Update

                                    .MoveNext
                                Next llFilterIndex
                            End If

                            .Filter = adFilterNone
                        End With
                    End If

                    .MoveNext
                Next llRecordIndex
            End If

            .Filter = adFilterNone
            .Filter = "RECORD_STATUS = 'C'"
            If Not (.EOF) Then
                .MoveFirst

                For llRecordIndex = 1 To .RecordCount
                    lsOverride = "Null"
                    If Not (IsNull(!OVER_EXC_CD)) Then
                        If (Trim(!OVER_EXC_CD) <> "") Then
                            lsOverride = "'" & Trim(!OVER_EXC_CD) & "'"
                        End If
                    End If

                    lsSQL = lsSQL & _
                        " UPDATE    METHDBA.TTR_ABS_INIT" & _
                        " SET       ABS_ST_TS = '" & Format(!ABS_ST_DT & " " & !ABS_ST_TM, "YYYY-MM-DD-hh.mm.ss.000000") & "'" & _
                        "           , ABS_END_TS = '" & Format(!ABS_END_DT & " " & !ABS_END_TM, "YYYY-MM-DD-hh.mm.ss.000000") & "'" & _
                        "           , ABS_CALLIN_TS = '" & Format(!ABS_CALLIN_DT & " " & !ABS_CALLIN_TM, "YYYY-MM-DD-hh.mm.ss.000000") & "'" & _
                        "           , EMP_PHONE = '" & Trim(!EMP_PHONE) & "'" & _
                        "           , EMP_LOC_CD = " & CStr(!EMP_LOC_CD) & _
                        "           , EXCEPTION_CD = '" & Trim(!EXCEPTION_CD) & "'" & _
                        "           , OVER_EXC_CD = " & lsOverride & _
                        "           , RPTNG_EMP_NUM = " & CStr(!RPTNG_EMP_NUM) & _
                        "           , RTN_TO_WRK_TS = '" & Format(!RTN_TO_WRK_DT & " " & !RTN_TO_WRK_TM, "YYYY-MM-DD-hh.mm.ss.000000") & "'" & _
                        "           , FMLA_REASON_CD = '" & Trim(!FMLA_REASON) & "'" & _
                        "           , ABS_PRIM_CONT_EMP = " & CStr(!ABS_PRIM_CONT_EMP) & _
                        "           , LAST_MOD_EMP_NUM = " & CStr(vlTRISUserNumber) & _
                        "           , LAST_MOD_TIMESTAMP = CURRENT TIMESTAMP" & _
                        "           , ABS_RPTR_TYPE = '" & Trim(!ABS_RPTR_TYPE) & "'" & _
                        "           , ABS_RPTR_TXT = '" & Trim(!ABS_RPTR_TXT) & "'" & _
                        "           , ADDL_TXT = '" & Trim(!ADDL_TXT) & "'" & _
                        "           , FOLW_UP_DT = '" & Format(Trim(!FOLW_UP_DT), "MM/DD/YYYY") & "'" & _
                        "           , RTW_NOT_REQD_IND = '" & Trim(!RTW_NOT_REQD_IND) & "'" & _
                        "           , ABS_NOTIFY_WG_NUM = " & CStr(!ABS_NOTIFY_WG_NUM) & _
                        "           , PRVDR_NM = '" & Trim(!PRVDR_NM) & "'" & _
                        "           , PRVDR_PHN_NBR = '" & Trim(!PRVDR_PHN_NBR) & "'" & _
                        " WHERE     ABS_NUM = " & CStr(!ABS_NUM) & _
                        " AND       CREATED_TIMESTAMP = '" & Trim(!CREATED_TIMESTAMP) & "';"

                    .MoveNext
                Next llRecordIndex
            End If

            .Filter = adFilterNone
        End With
    End If

    If (RecordsExist(vadrsAbsenceDates)) Then
        With vadrsAbsenceDates
            .Filter = adFilterNone
            .Filter = "RECORD_STATUS = 'D'"
            If Not (.EOF) Then
                .MoveFirst

                For llRecordIndex = 1 To .RecordCount
                    lsSQL = lsSQL & _
                        " DELETE FROM METHDBA.TTR_ABS_DATE" & _
                        " WHERE ABS_NUM = " & CStr(!ABS_NUM) & _
                        " AND   ABS_DATE = '" & CStr(!ABS_DATE) & "';"

                    .MoveNext
                Next llRecordIndex
            End If

            .Filter = adFilterNone
            .Filter = "RECORD_STATUS = 'I'"
            If Not (.EOF) Then
                .MoveFirst

                For llRecordIndex = 1 To .RecordCount
                    lsSQL = lsSQL & _
                        " INSERT INTO METHDBA.TTR_ABS_DATE" & _
                        " (ABS_NUM, ABS_DATE" & _
                        ", CREATED_EMP_NUM, CREATED_TIMESTAMP" & _
                        ", LAST_MOD_EMP_NUM, LAST_MOD_TIMESTAMP)" & _
                        " VALUES" & _
                        " (" & CStr(!ABS_NUM) & ", '" & CStr(!ABS_DATE) & "'" & _
                        ", " & CStr(vlTRISUserNumber) & ", CURRENT TIMESTAMP" & _
                        ", " & CStr(vlTRISUserNumber) & ", CURRENT TIMESTAMP);"

                    .MoveNext
                Next llRecordIndex
            End If

            .Filter = adFilterNone
        End With
    End If

    If (RecordsExist(vadrsAbsenceProvider)) Then
        With vadrsAbsenceProvider
            .Filter = adFilterNone
            .Filter = "RECORD_STATUS = 'D'"
            If Not (.EOF) Then
                .MoveFirst

                For llRecordIndex = 1 To .RecordCount
                    lsSQL = lsSQL & _
                        " DELETE FROM METHDBA.TTR_ABS_PROVIDER" & _
                        " WHERE ABS_NUM = " & CStr(!ABS_NUM) & _
                        " AND   PRVDR_NBR = " & CStr(!PRVDR_NBR) & _
                        " AND   CREATED_TIMESTAMP = '" & Trim(!CREATED_TIMESTAMP) & "';"

                    .MoveNext
                Next llRecordIndex
            End If

            .Filter = adFilterNone
            .Filter = "RECORD_STATUS = 'I'"
            If Not (.EOF) Then
                .MoveFirst

                For llRecordIndex = 1 To .RecordCount
                    If (!PRVDR_NBR = -1) Then
                        Set lobjUtility = CreateObject( _
                                            "utilDBDataServices.cuDBDataServices")
                        llProviderNumber = lobjUtility.GetNextSequenceNumber("AP")
                        Set lobjUtility = Nothing

                        !PRVDR_NBR = llProviderNumber
                        .Update

                        lsSQL = lsSQL & _
                            " INSERT INTO METHDBA.TTR_PROVIDER" & _
                            " (PRVDR_NBR, PRVDR_NM" & _
                            ", LN1_ADDR, LN2_ADDR" & _
                            ", CTY_NM, ST_CD" & _
                            ", ZIP_CD, PHN_NBR" & _
                            ", CLS_TXT, TYP_CD)" & _
                            " VALUES" & _
                            " (" & CStr(!PRVDR_NBR) & ", '" & Trim(!PRVDR_NM) & "'" & _
                            ", '" & Trim(!LN1_ADDR) & "', '" & Trim(!LN2_ADDR) & "'" & _
                            ", '" & Trim(!CTY_NM) & "', '" & Trim(!ST_CD) & "'" & _
                            ", '" & Trim(!ZIP_CD) & "', '" & Trim(!PHN_NBR) & "'" & _
                            ", '" & Trim(!CLS_TXT) & "', '" & Trim(!TYP_DESC) & "');"
                    End If

                    lsSQL = lsSQL & _
                        " INSERT INTO METHDBA.TTR_ABS_PROVIDER" & _
                        " (ABS_NUM, PRVDR_NBR" & _
                        ", VST_ST_ADM_TS, VST_END_DGCE_TS" & _
                        ", ADDL_TXT, RLSE_IND" & _
                        ", CREATED_EMP_NUM, CREATED_TIMESTAMP" & _
                        ", LAST_MOD_EMP_NUM, LAST_MOD_TIMESTAMP)" & _
                        " VALUES" & _
                        " (" & CStr(!ABS_NUM) & ", " & CStr(!PRVDR_NBR) & _
                        ", '" & Format(!VST_ST_ADM_DT & " " & !VST_ST_ADM_TM, "YYYY-MM-DD-hh.mm.ss.000000") & "'" & _
                        ", '" & Format(!VST_END_DGCE_DT & " " & !VST_END_DGCE_TM, "YYYY-MM-DD-hh.mm.ss.000000") & "'" & _
                        ", '" & Trim(!ADDL_TXT) & "', '" & Trim(!RLSE_IND) & "'" & _
                        ", " & CStr(vlTRISUserNumber) & ", CURRENT TIMESTAMP" & _
                        ", " & CStr(vlTRISUserNumber) & ", CURRENT TIMESTAMP);"

                    .MoveNext
                Next llRecordIndex
            End If

            .Filter = adFilterNone
            .Filter = "RECORD_STATUS = 'C'"
            If Not (.EOF) Then
                .MoveFirst

                For llRecordIndex = 1 To .RecordCount
                    lsSQL = lsSQL & _
                        " UPDATE METHDBA.TTR_PROVIDER" & _
                        " SET   PRVDR_NM = '" & Trim(!PRVDR_NM) & "'" & _
                        "       , LN1_ADDR = '" & Trim(!LN1_ADDR) & "'" & _
                        "       , LN2_ADDR = '" & Trim(!LN2_ADDR) & "'" & _
                        "       , CTY_NM = '" & Trim(!CTY_NM) & "'" & _
                        "       , ST_CD = '" & Trim(!ST_CD) & "'" & _
                        "       , ZIP_CD = '" & Trim(!ZIP_CD) & "'" & _
                        "       , PHN_NBR = '" & Trim(!PHN_NBR) & "'" & _
                        "       , CLS_TXT = '" & Trim(!CLS_TXT) & "'" & _
                        "       , TYP_CD = '" & Trim(!TYP_DESC) & "'" & _
                        " WHERE PRVDR_NBR = " & CStr(!PRVDR_NBR) & ";"

                    lsSQL = lsSQL & _
                        " UPDATE METHDBA.TTR_ABS_PROVIDER" & _
                        " SET   PRVDR_NBR = " & CStr(!PRVDR_NBR) & _
                        "       , VST_ST_ADM_TS = '" & Format(!VST_ST_ADM_DT & " " & !VST_ST_ADM_TM, "YYYY-MM-DD-hh.mm.ss.000000") & "'" & _
                        "       , VST_END_DGCE_TS = '" & Format(!VST_END_DGCE_DT & " " & !VST_END_DGCE_TM, "YYYY-MM-DD-hh.mm.ss.000000") & "'" & _
                        "       , RLSE_IND = '" & Trim(!RLSE_IND) & "'" & _
                        "       , ADDL_TXT = '" & Trim(!ADDL_TXT) & "'" & _
                        "       , LAST_MOD_EMP_NUM = " & CStr(vlTRISUserNumber) & _
                        "       , LAST_MOD_TIMESTAMP = CURRENT TIMESTAMP" & _
                        " WHERE ABS_NUM = " & CStr(!ABS_NUM) & _
                        " AND   CREATED_TIMESTAMP = '" & Trim(!CREATED_TIMESTAMP) & "';"

                    .MoveNext
                Next llRecordIndex
            End If

            .Filter = adFilterNone
        End With
    End If

    If (RecordsExist(vadrsAbsenceContact)) Then
        With vadrsAbsenceContact
            .Filter = adFilterNone
            .Filter = "RECORD_STATUS = 'D'"
            If Not (.EOF) Then
                .MoveFirst

                For llRecordIndex = 1 To .RecordCount
                    lsSQL = lsSQL & _
                        " DELETE FROM METHDBA.TTR_ABS_CONTACT" & _
                        " WHERE ABS_NUM = " & CStr(!ABS_NUM) & _
                        " AND   ABS_CONT_NUM = " & CStr(!ABS_CONT_NUM) & _
                        " AND   CREATED_TIMESTAMP = '" & Trim(!CREATED_TIMESTAMP) & "';"

                    .MoveNext
                Next llRecordIndex
            End If

            .Filter = adFilterNone
            .Filter = "RECORD_STATUS = 'I'"
            If Not (.EOF) Then
                .MoveFirst

                For llRecordIndex = 1 To .RecordCount
                    If (!ABS_CONT_NUM = -1) Then
                        Set lobjUtility = CreateObject( _
                                            "utilDBDataServices.cuDBDataServices")
                        llContactNumber = lobjUtility.GetNextSequenceNumber("AC")
                        Set lobjUtility = Nothing

                        !ABS_CONT_NUM = llContactNumber
                        .Update
                    End If

                    lsSQL = lsSQL & _
                        " INSERT INTO METHDBA.TTR_ABS_CONTACT" & _
                        " (ABS_NUM, ABS_CONT_NUM" & _
                        ", ABS_CONT_TS, USR_INIT_IND" & _
                        ", TALK_TO_EMP_IND, TALK_TO_EMP_TXT" & _
                        ", EMP_TO_DR_IND, EMP_TO_DR_TXT" & _
                        ", OTH_TRTMNT_IND, OTH_TRTMNT_TXT" & _
                        ", ADVS_SL_IND, ADVS_SL_TXT" & _
                        ", PHONE_NUM, RPTNG_EMP_NUM" & _
                        ", ADDL_TXT, PRVDR_PHN_NUM" & _
                        ", CREATED_EMP_NUM, CREATED_TIMESTAMP" & _
                        ", LAST_MOD_EMP_NUM, LAST_MOD_TIMESTAMP)" & _
                        " VALUES" & _
                        " (" & CStr(!ABS_NUM) & ", " & CStr(!ABS_CONT_NUM) & _
                        ", '" & Format(!ABS_CONT_DT & " " & !ABS_CONT_TM, "YYYY-MM-DD-hh.mm.ss.000000") & "', '" & Trim(!USR_INIT_IND) & "'" & _
                        ", '" & Trim(!TALK_TO_EMP_IND) & "', '" & Trim(!TALK_TO_EMP_TXT) & "'" & _
                        ", '" & Trim(!EMP_TO_DR_IND) & "', '" & Trim(!EMP_TO_DR_TXT) & "'" & _
                        ", '" & Trim(!OTH_TRTMNT_IND) & "', '" & Trim(!OTH_TRTMNT_TXT) & "'" & _
                        ", '" & Trim(!ADVS_SL_IND) & "', '" & Trim(!ADVS_SL_TXT) & "'" & _
                        ", '" & Trim(!PHONE_NUM) & "', " & CStr(!RPTNG_EMP_NUM) & _
                        ", '" & Trim(!ADDL_TXT) & "', '" & Trim(!PRVDR_PHN_NUM) & "'" & _
                        ", " & CStr(vlTRISUserNumber) & ", CURRENT TIMESTAMP" & _
                        ", " & CStr(vlTRISUserNumber) & ", CURRENT TIMESTAMP);"

                    .MoveNext
                Next llRecordIndex
            End If

            .Filter = adFilterNone
            .Filter = "RECORD_STATUS = 'C'"
            If Not (.EOF) Then
                .MoveFirst

                For llRecordIndex = 1 To .RecordCount
                    
                    lsSQL = lsSQL & _
                        " UPDATE METHDBA.TTR_ABS_CONTACT" & _
                        " SET   ABS_CONT_TS = '" & Format(!ABS_CONT_DT & " " & !ABS_CONT_TM, "YYYY-MM-DD-hh.mm.ss.000000") & "'" & _
                        "       , USR_INIT_IND = '" & Trim(!USR_INIT_IND) & "'" & _
                        "       , TALK_TO_EMP_IND = '" & Trim(!TALK_TO_EMP_IND) & "'" & _
                        "       , TALK_TO_EMP_TXT = '" & Trim(!TALK_TO_EMP_TXT) & "'" & _
                        "       , EMP_TO_DR_IND = '" & Trim(!EMP_TO_DR_IND) & "'" & _
                        "       , EMP_TO_DR_TXT = '" & Trim(!EMP_TO_DR_TXT) & "'" & _
                        "       , OTH_TRTMNT_IND = '" & Trim(!OTH_TRTMNT_IND) & "'" & _
                        "       , OTH_TRTMNT_TXT = '" & Trim(!OTH_TRTMNT_TXT) & "'" & _
                        "       , ADVS_SL_IND = '" & Trim(!ADVS_SL_IND) & "'" & _
                        "       , ADVS_SL_TXT = '" & Trim(!ADVS_SL_TXT) & "'" & _
                        "       , PHONE_NUM = '" & Trim(!PHONE_NUM) & "'" & _
                        "       , RPTNG_EMP_NUM = " & CStr(!RPTNG_EMP_NUM) & _
                        "       , ADDL_TXT = '" & Trim(!ADDL_TXT) & "'" & _
                        "       , PRVDR_PHN_NUM = '" & Trim(!PRVDR_PHN_NUM) & "'" & _
                        "       , LAST_MOD_EMP_NUM = " & CStr(vlTRISUserNumber) & _
                        "       , LAST_MOD_TIMESTAMP = CURRENT TIMESTAMP" & _
                        " WHERE ABS_NUM = " & CStr(!ABS_NUM) & _
                        " AND   ABS_CONT_NUM = " & CStr(!ABS_CONT_NUM) & _
                        " AND   CREATED_TIMESTAMP = '" & Trim(!CREATED_TIMESTAMP) & "';"

                    .MoveNext
                Next llRecordIndex
            End If

            .Filter = adFilterNone
        End With
    End If

    If (RecordsExist(vadrsAbsenceReturnToWork)) Then
        With vadrsAbsenceReturnToWork
            .Filter = adFilterNone
            .Filter = "RECORD_STATUS = 'D'"
            If Not (.EOF) Then
                .MoveFirst

                For llRecordIndex = 1 To .RecordCount
                    lsSQL = lsSQL & _
                        " DELETE FROM METHDBA.TTR_ABS_RET_WRK" & _
                        " WHERE ABS_NUM = " & CStr(!ABS_NUM) & _
                        " AND   CREATED_TIMESTAMP = '" & Trim(!CREATED_TIMESTAMP) & "';"

                    .MoveNext
                Next llRecordIndex
            End If

            .Filter = adFilterNone
            .Filter = "RECORD_STATUS = 'I'"
            If Not (.EOF) Then
                .MoveFirst

                For llRecordIndex = 1 To .RecordCount
                    lsSQL = lsSQL & _
                        " INSERT INTO METHDBA.TTR_ABS_RET_WRK" & _
                        " (ABS_NUM, CONFRM_RSN_IND" & _
                        ", ABS_REASON_TXT" & _
                        ", MEET_EMP_IND, MEET_EMP_TXT" & _
                        ", PRFM_DUTY_IND, PRFM_DUTY_TXT" & _
                        ", SIDE_EFF_IND, SIDE_EFF_TXT" & _
                        ", ADDL_TXT, RTN_TO_WRK_TS" & _
                        ", RPT_DT, RPTNG_EMP_NUM" & _
                        ", CREATED_EMP_NUM, CREATED_TIMESTAMP" & _
                        ", LAST_MOD_EMP_NUM, LAST_MOD_TIMESTAMP)" & _
                        " VALUES" & _
                        " (" & CStr(!ABS_NUM) & ", '" & Trim(!CONFRM_RSN_IND) & "'" & _
                        ", '" & Trim(!ABS_REASON_TXT) & "'" & _
                        ", '" & Trim(!MEET_EMP_IND) & "', '" & Trim(!MEET_EMP_TXT) & "'" & _
                        ", '" & Trim(!PRFM_DUTY_IND) & "', '" & Trim(!PRFM_DUTY_TXT) & "'" & _
                        ", '" & Trim(!SIDE_EFF_IND) & "', '" & Trim(!SIDE_EFF_TXT) & "'" & _
                        ", '" & Trim(!ADDL_TXT) & "', '" & Format(!RTN_TO_WRK_DT & " " & !RTN_TO_WRK_TM, "YYYY-MM-DD-hh.mm.ss.000000") & "'" & _
                        ", '" & Format(Trim(!RPT_DT), "MM/DD/YYYY") & "', " & CStr(!RPTNG_EMP_NUM) & _
                        ", " & CStr(vlTRISUserNumber) & ", CURRENT TIMESTAMP" & _
                        ", " & CStr(vlTRISUserNumber) & ", CURRENT TIMESTAMP);"

                    .MoveNext
                Next llRecordIndex
            End If

            .Filter = adFilterNone
            .Filter = "RECORD_STATUS = 'C'"
            If Not (.EOF) Then
                .MoveFirst

                For llRecordIndex = 1 To .RecordCount
                    '"       , EMPL_CURR_STAT_TXT = '" & Trim(!EMPL_CURR_STAT_TXT) & "'"
                    lsSQL = lsSQL & _
                        " UPDATE METHDBA.TTR_ABS_RET_WRK" & _
                        " SET   CONFRM_RSN_IND = '" & Trim(!CONFRM_RSN_IND) & "'" & _
                        "       , ABS_REASON_TXT = '" & Trim(!ABS_REASON_TXT) & "'" & _
                        "       , MEET_EMP_IND = '" & Trim(!MEET_EMP_IND) & "'" & _
                        "       , MEET_EMP_TXT = '" & Trim(!MEET_EMP_TXT) & "'" & _
                        "       , PRFM_DUTY_IND = '" & Trim(!PRFM_DUTY_IND) & "'" & _
                        "       , PRFM_DUTY_TXT = '" & Trim(!PRFM_DUTY_TXT) & "'" & _
                        "       , SIDE_EFF_IND = '" & Trim(!SIDE_EFF_IND) & "'" & _
                        "       , SIDE_EFF_TXT = '" & Trim(!SIDE_EFF_TXT) & "'" & _
                        "       , ADDL_TXT = '" & Trim(!ADDL_TXT) & "'" & _
                        "       , RTN_TO_WRK_TS = '" & Format(!RTN_TO_WRK_DT & " " & !RTN_TO_WRK_TM, "YYYY-MM-DD-hh.mm.ss.000000") & "'" & _
                        "       , RPT_DT = '" & Format(Trim(!RPT_DT), "MM/DD/YYYY") & "'" & _
                        "       , RPTNG_EMP_NUM = " & CStr(!RPTNG_EMP_NUM) & _
                        "       , LAST_MOD_EMP_NUM = " & CStr(vlTRISUserNumber) & _
                        "       , LAST_MOD_TIMESTAMP = CURRENT TIMESTAMP" & _
                        " WHERE ABS_NUM = " & CStr(!ABS_NUM) & _
                        " AND   CREATED_TIMESTAMP = '" & Trim(!CREATED_TIMESTAMP) & "';"

                    .MoveNext
                Next llRecordIndex
            End If

            .Filter = adFilterNone
        End With
    End If

    If (RecordsExist(vadrsAbsenceCallLog)) Then
        With vadrsAbsenceCallLog
            .Filter = adFilterNone
            .Filter = "RECORD_STATUS = 'I'"
            If Not (.EOF) Then
                For llRecordIndex = 1 To .RecordCount
                    lsSQL = lsSQL & _
                        " INSERT INTO METHDBA.TTR_ABS_CALL_LOG" & _
                        " (ABS_NUM, CALL_NUM" & _
                        ", CALL_ST_TS, CALL_END_TS" & _
                        ", PHONE_NUM, CALLER_EMP_NUM" & _
                        ", CREATED_TIMESTAMP)" & _
                        " VALUES " & _
                        " (" & CStr(!ABS_NUM) & ", " & CStr(!CALL_NUM) & _
                        ", '" & Trim(!CALL_ST_TS) & "', '" & Trim(!CALL_END_TS) & "'" & _
                        ", '" & Trim(!PHONE_NUM) & "', " & CStr(vlTRISUserNumber) & _
                        ", CURRENT TIMESTAMP);"

                    .MoveNext
                Next llRecordIndex
            End If
        End With
    End If

    If (RecordsExist(vadrsPortfolio)) Then
        With vadrsPortfolio
            .Filter = adFilterNone
            .Filter = "RECORD_STATUS = 'I'"
            If Not (.EOF) Then
                If (!PRTFL_NUM = -1) Then
                    Set lobjUtility = CreateObject( _
                                        "utilDBDataServices.cuDBDataServices")
                    llPortfolioNumber = lobjUtility.GetNextSequenceNumber("EP")
                    Set lobjUtility = Nothing

                    AppendBusinessError radrsBRV _
                        , trInfEAPM _
                        , mstrCBUSINESSINF _
                        , "A new Portfolio " & CStr(llPortfolioNumber) & " was created." _
                        , "dbProcessEAMRequest"

                    If (RecordsExist(vadrsPortfolioDetail)) Then
                        vadrsPortfolioDetail.MoveFirst

                        For llRecordIndex = 1 To vadrsPortfolioDetail.RecordCount
                            If (vadrsPortfolioDetail!PRTFL_NUM = !PRTFL_NUM) Then
                                vadrsPortfolioDetail!PRTFL_NUM = llPortfolioNumber
                                vadrsPortfolioDetail.Update
                            End If

                            vadrsPortfolioDetail.MoveNext
                        Next llRecordIndex

                        vadrsPortfolioDetail.MoveFirst
                    End If

                    lsSQL = lsSQL & _
                        " INSERT INTO METHDBA.TTR_ABS_PORTFOLIO" & _
                        " (PRTFL_NUM, PRTFL_DESC" & _
                        " , SHARE_IND" & _
                        " , CREATED_EMP_NUM, CREATED_TIMESTAMP" & _
                        " , LAST_MOD_EMP_NUM, LAST_MOD_TIMESTAMP)" & _
                        " VALUES" & _
                        " (" & llPortfolioNumber & ", '" & Trim(!PRTFL_DESC) & "'" & _
                        " , '" & Trim(!SHARE_IND) & "'" & _
                        " , " & CStr(vlTRISUserNumber) & ", CURRENT TIMESTAMP" & _
                        " , " & CStr(vlTRISUserNumber) & ", CURRENT TIMESTAMP);"
                End If
            End If
        End With
    End If

    If (RecordsExist(vadrsPortfolioDetail)) Then
        With vadrsPortfolioDetail
            .Filter = adFilterNone
            .Filter = "RECORD_STATUS = 'I'"
            If Not (.EOF) Then
                For llRecordIndex = 1 To .RecordCount
                    lsSQL = lsSQL & _
                        " INSERT INTO METHDBA.TTR_ABS_PRTFL_DTL" & _
                        " (PRTFL_NUM, ABS_NUM" & _
                        " , CREATED_EMP_NUM, CREATED_TIMESTAMP" & _
                        " , LAST_MOD_EMP_NUM, LAST_MOD_TIMESTAMP)" & _
                        " VALUES" & _
                        " (" & CStr(!PRTFL_NUM) & ", " & CStr(!ABS_NUM) & _
                        " , " & CStr(vlTRISUserNumber) & ", CURRENT TIMESTAMP" & _
                        " , " & CStr(vlTRISUserNumber) & ", CURRENT TIMESTAMP);"

                    .MoveNext
                Next llRecordIndex
            End If
        End With
    End If

EXITNORMALLY:
    If (Trim(lsSQL) <> "") Then
        'Get Connection
        Set ladcnConnection = OpenConnection()
        'Call Routine to Execute SQL
        ExecuteOTHERSQL lsSQL, ladcnConnection
    End If
    'Close and destroy
    If Not (ladcnConnection Is Nothing) Then
        ladcnConnection.Close
        Set ladcnConnection = Nothing
    End If

    Set lobjUtility = Nothing

    dbProcessEAMRequest = True

    GetObjectContext.SetComplete

    Exit Function
dbProcessEAMRequest_EH:
    GetObjectContext.SetAbort

    Set lobjUtility = Nothing
    Set ladcnConnection = Nothing

    Set radrsBRV = Nothing

    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "dbProcessEAMRequest"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "dbProcessEAMRequest" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "dbProcessEAMRequest" _
                    , trErrdbSaveEAM _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
Public Function dbProcessDeleteAbsence(ByVal vadrsAbsence As ADODB.Recordset _
                ) As Boolean
On Error GoTo dbProcessDeleteAbsence_EH
Dim ladcnConnection As ADODB.Connection
Dim llRecordIndex As Long, lsSQL As String

    dbProcessDeleteAbsence = False

    If (RecordsExist(vadrsAbsence)) Then
        With vadrsAbsence
            .MoveFirst

            For llRecordIndex = 1 To .RecordCount
'ON HOLD INIT ABS. MGMT
'                lsSQL = lsSQL & _
'                    " DELETE FROM METHDBA.TTR_ABS_WC" & _
'                    " WHERE ABS_NUM = " & CStr(!ABS_NUM) & ";"

                lsSQL = lsSQL & _
                    " DELETE FROM METHDBA.TTR_ABS_PRTFL_DTL" & _
                    " WHERE ABS_NUM = " & CStr(!ABS_NUM) & ";"

                lsSQL = lsSQL & _
                    " DELETE FROM METHDBA.TTR_ABS_DATE" & _
                    " WHERE ABS_NUM = " & CStr(!ABS_NUM) & ";"

                lsSQL = lsSQL & _
                    " DELETE FROM METHDBA.TTR_ABS_RET_WRK" & _
                    " WHERE ABS_NUM = " & CStr(!ABS_NUM) & ";"

                lsSQL = lsSQL & _
                    " DELETE FROM METHDBA.TTR_ABS_CALL_LOG" & _
                    " WHERE ABS_NUM = " & CStr(!ABS_NUM) & ";"

                lsSQL = lsSQL & _
                    " DELETE FROM METHDBA.TTR_ABS_CONTACT" & _
                    " WHERE ABS_NUM = " & CStr(!ABS_NUM) & ";"

                lsSQL = lsSQL & _
                    " DELETE FROM METHDBA.TTR_ABS_PROVIDER" & _
                    " WHERE ABS_NUM = " & CStr(!ABS_NUM) & ";"

                lsSQL = lsSQL & _
                    " DELETE FROM METHDBA.TTR_ABS_INIT" & _
                    " WHERE ABS_NUM = " & CStr(!ABS_NUM) & ";"

                .MoveNext
            Next llRecordIndex

            .MoveFirst
        End With
    End If

    If (Trim(lsSQL) <> "") Then
        'Get Connection
        Set ladcnConnection = OpenConnection()
        'Call Routine to Execute SQL
        ExecuteOTHERSQL lsSQL, ladcnConnection
    End If
    'Close and destroy
    If Not (ladcnConnection Is Nothing) Then
        ladcnConnection.Close
        Set ladcnConnection = Nothing
    End If

    dbProcessDeleteAbsence = True

    GetObjectContext.SetComplete

    Exit Function
dbProcessDeleteAbsence_EH:
    GetObjectContext.SetAbort

    Set ladcnConnection = Nothing

    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "dbProcessDeleteAbsence"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "dbProcessDeleteAbsence" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "dbProcessDeleteAbsence" _
                    , trErrdbSaveEAM _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select
End Function
Public Function dbProcessPorfolioRequest(ByVal vadrsPortfolio As ADODB.Recordset _
                , ByVal vadrsAbsences As ADODB.Recordset _
                , ByVal vlTRISUserNumber As Long _
                , ByVal vsTRISUserID As String _
                , ByRef rlPortfolioNumber As Long _
                , ByRef radrsBRV As ADODB.Recordset _
                ) As Boolean
On Error GoTo dbProcessPorfolioRequest_EH
'ADODB Connection Object
Dim ladcnConnection As ADODB.Connection
'Recordset
Dim llRecordIndex As Long, llPortfolioNumber As Long
'SQL
Dim lsSQL As String
'SEQ_NUM
Dim lobjUtility As utilDBDataServices.cuDBDataServices

    dbProcessPorfolioRequest = False

    lsSQL = ""

    If (RecordsExist(vadrsPortfolio)) Then
        With vadrsPortfolio
            .Filter = adFilterNone
            .Filter = "RECORD_STATUS = 'D'"
            If Not (.EOF) Then
                lsSQL = lsSQL & _
                    " DELETE FROM METHDBA.TTR_ABS_PRTFL_DTL" & _
                    " WHERE PRTFL_NUM = " & CStr(!PRTFL_NUM) & ";"

                lsSQL = lsSQL & _
                    " DELETE FROM METHDBA.TTR_ABS_PORTFOLIO" & _
                    " WHERE PRTFL_NUM = " & CStr(!PRTFL_NUM) & ";"

                GoTo EXITNORMALLY
            End If

            .Filter = adFilterNone
            .Filter = "RECORD_STATUS = 'I'"
            If Not (.EOF) Then
                If (rlPortfolioNumber = -1) Then
                    Set lobjUtility = CreateObject( _
                                        "utilDBDataServices.cuDBDataServices")
                    llPortfolioNumber = lobjUtility.GetNextSequenceNumber("EP")
                    Set lobjUtility = Nothing

                    AppendBusinessError radrsBRV _
                        , trInfEAPM _
                        , mstrCBUSINESSINF _
                        , "A new Portfolio " & CStr(llPortfolioNumber) & " was created." _
                        , "dbProcessPorfolioRequest"

                    If (RecordsExist(vadrsAbsences)) Then
                        vadrsAbsences.MoveFirst

                        For llRecordIndex = 1 To vadrsAbsences.RecordCount
                            If (vadrsAbsences!PRTFL_NUM = rlPortfolioNumber) Then
                                vadrsAbsences!PRTFL_NUM = llPortfolioNumber
                                vadrsAbsences.Update
                            End If

                            vadrsAbsences.MoveNext
                        Next llRecordIndex

                        vadrsAbsences.MoveFirst
                    End If

                    rlPortfolioNumber = llPortfolioNumber

                    lsSQL = lsSQL & _
                        " INSERT INTO METHDBA.TTR_ABS_PORTFOLIO" & _
                        " (PRTFL_NUM, PRTFL_DESC" & _
                        " , SHARE_IND" & _
                        " , CREATED_EMP_NUM, CREATED_TIMESTAMP" & _
                        " , LAST_MOD_EMP_NUM, LAST_MOD_TIMESTAMP)" & _
                        " VALUES" & _
                        " (" & llPortfolioNumber & ", '" & Trim(!PRTFL_DESC) & "'" & _
                        " , '" & Trim(!SHARE_IND) & "'" & _
                        " , " & CStr(vlTRISUserNumber) & ", CURRENT TIMESTAMP" & _
                        " , " & CStr(vlTRISUserNumber) & ", CURRENT TIMESTAMP);"
                End If
            End If

            .Filter = adFilterNone
            .Filter = "RECORD_STATUS = 'C'"
            If Not (.EOF) Then
                lsSQL = lsSQL & _
                    " UPDATE    METHDBA.TTR_ABS_PORTFOLIO" & _
                    " SET       PRTFL_DESC = '" & Trim(!PRTFL_DESC) & "'" & _
                    "           , SHARE_IND = '" & Trim(!SHARE_IND) & "'" & _
                    "           , LAST_MOD_EMP_NUM = " & CStr(vlTRISUserNumber) & _
                    "           , LAST_MOD_TIMESTAMP = CURRENT TIMESTAMP" & _
                    " WHERE     PRTFL_NUM = " & CStr(!PRTFL_NUM) & _
                    " AND       CREATED_TIMESTAMP = '" & Trim(!CREATED_TIMESTAMP) & "';"
            End If

            .Filter = adFilterNone
        End With
    End If

    If (RecordsExist(vadrsAbsences)) Then
        With vadrsAbsences
            .MoveFirst

            .Filter = adFilterNone
            .Filter = "RECORD_STATUS = 'D'"
            If Not (.EOF) Then
                For llRecordIndex = 1 To .RecordCount
                    lsSQL = lsSQL & _
                        " DELETE    FROM METHDBA.TTR_ABS_PRTFL_DTL" & _
                        " WHERE     PRTFL_NUM = " & CStr(!PRTFL_NUM) & _
                        " AND       ABS_NUM = " & CStr(!ABS_NUM) & ";"

                    .MoveNext
                Next llRecordIndex
            End If

            .Filter = adFilterNone
            .Filter = "RECORD_STATUS = 'I'"
            If Not (.EOF) Then
                For llRecordIndex = 1 To .RecordCount
                    lsSQL = lsSQL & _
                        " INSERT INTO METHDBA.TTR_ABS_PRTFL_DTL" & _
                        " (PRTFL_NUM, ABS_NUM" & _
                        " , CREATED_EMP_NUM, CREATED_TIMESTAMP" & _
                        " , LAST_MOD_EMP_NUM, LAST_MOD_TIMESTAMP)" & _
                        " VALUES" & _
                        " (" & CStr(!PRTFL_NUM) & ", " & CStr(!ABS_NUM) & _
                        " , " & CStr(vlTRISUserNumber) & ", CURRENT TIMESTAMP" & _
                        " , " & CStr(vlTRISUserNumber) & ", CURRENT TIMESTAMP);"

                    .MoveNext
                Next llRecordIndex
            End If
        End With
    End If

EXITNORMALLY:
    If (Trim(lsSQL) <> "") Then
        'Get Connection
        Set ladcnConnection = OpenConnection()
        'Call Routine to Execute SQL
        ExecuteOTHERSQL lsSQL, ladcnConnection
    End If
    'Close and destroy
    If Not (ladcnConnection Is Nothing) Then
        ladcnConnection.Close
        Set ladcnConnection = Nothing
    End If

    Set lobjUtility = Nothing

    dbProcessPorfolioRequest = True

    GetObjectContext.SetComplete

    Exit Function
dbProcessPorfolioRequest_EH:
    GetObjectContext.SetAbort

    Set lobjUtility = Nothing
    Set ladcnConnection = Nothing

    Set radrsBRV = Nothing

    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "dbProcessPorfolioRequest"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "dbProcessPorfolioRequest" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "dbProcessPorfolioRequest" _
                    , trErrdbSaveAbsencePortfolio _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
Public Function dbProcessDeletePortfolio(ByVal vlPortfolioNumber As Long _
                ) As Boolean
On Error GoTo dbProcessDeletePortfolio_EH
Dim ladcnConnection As ADODB.Connection
Dim llRecordIndex As Long, lsSQL As String

    dbProcessDeletePortfolio = False

    lsSQL = lsSQL & _
        " DELETE FROM METHDBA.TTR_ABS_PRTFL_DTL" & _
        " WHERE PRTFL_NUM = " & CStr(vlPortfolioNumber) & ";"

    lsSQL = lsSQL & _
        " DELETE FROM METHDBA.TTR_ABS_PORTFOLIO" & _
        " WHERE PRTFL_NUM = " & CStr(vlPortfolioNumber) & ";"

    If (Trim(lsSQL) <> "") Then
        'Get Connection
        Set ladcnConnection = OpenConnection()
        'Call Routine to Execute SQL
        ExecuteOTHERSQL lsSQL, ladcnConnection
    End If
    'Close and destroy
    If Not (ladcnConnection Is Nothing) Then
        ladcnConnection.Close
        Set ladcnConnection = Nothing
    End If

    dbProcessDeletePortfolio = True

    GetObjectContext.SetComplete

    Exit Function
dbProcessDeletePortfolio_EH:
    GetObjectContext.SetAbort

    Set ladcnConnection = Nothing

    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "dbProcessDeletePortfolio"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "dbProcessDeletePortfolio" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "dbProcessDeletePortfolio" _
                    , trErrdbSaveAbsencePortfolio _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
Public Function dbProcessContactRequest(ByVal vadrsContacts As ADODB.Recordset _
                , ByVal vlTRISUserNumber As Long _
                , ByVal vsTRISUserID As String _
                ) As Boolean
On Error GoTo dbProcessContactRequest_EH
'ADODB Connection Object
Dim ladcnConnection As ADODB.Connection
'Recordset
Dim llRecordIndex As Long, llContactNumber As Long
'SQL
Dim lsSQL As String
'SEQ_NUM
Dim lobjUtility As utilDBDataServices.cuDBDataServices

    dbProcessContactRequest = False

    If (RecordsExist(vadrsContacts)) Then
        With vadrsContacts
            .Filter = adFilterNone
            .Filter = "RECORD_STATUS = 'I'"

            If Not (.EOF) Then
                lsSQL = ""

                For llRecordIndex = 1 To .RecordCount
                    If (!ABS_CONT_NUM = -1) Then
                        Set lobjUtility = CreateObject( _
                                            "utilDBDataServices.cuDBDataServices")
                        llContactNumber = lobjUtility.GetNextSequenceNumber("AC")
                        Set lobjUtility = Nothing

                        lsSQL = lsSQL & _
                            " INSERT INTO METHDBA.TTR_ABS_CONTACT" & _
                            " (ABS_NUM, ABS_CONT_NUM" & _
                            ", ABS_CONT_TS, USR_INIT_IND" & _
                            ", TALK_TO_EMP_IND, TALK_TO_EMP_TXT" & _
                            ", EMP_TO_DR_IND, EMP_TO_DR_TXT" & _
                            ", OTH_TRTMNT_IND, OTH_TRTMNT_TXT" & _
                            ", ADVS_SL_IND, ADVS_SL_TXT" & _
                            ", PHONE_NUM, RPTNG_EMP_NUM" & _
                            ", ADDL_TXT, PRVDR_PHN_NUM" & _
                            ", CREATED_EMP_NUM, CREATED_TIMESTAMP" & _
                            ", LAST_MOD_EMP_NUM, LAST_MOD_TIMESTAMP)" & _
                            " VALUES" & _
                            " (" & CStr(!ABS_NUM) & ", " & CStr(llContactNumber) & _
                            ", '" & Format(!ABS_CONT_DT & " " & !ABS_CONT_TM, "YYYY-MM-DD-hh.mm.ss.000000") & "', '" & Trim(!USR_INIT_IND) & "'" & _
                            ", '" & Trim(!TALK_TO_EMP_IND) & "', '" & Trim(!TALK_TO_EMP_TXT) & "'" & _
                            ", '" & Trim(!EMP_TO_DR_IND) & "', '" & Trim(!EMP_TO_DR_TXT) & "'" & _
                            ", '" & Trim(!OTH_TRTMNT_IND) & "', '" & Trim(!OTH_TRTMNT_TXT) & "'" & _
                            ", '" & Trim(!ADVS_SL_IND) & "', '" & Trim(!ADVS_SL_TXT) & "'" & _
                            ", '" & Trim(!PHONE_NUM) & "', " & CStr(!RPTNG_EMP_NUM) & _
                            ", '" & Trim(!ADDL_TXT) & "', '" & Trim(!PRVDR_PHN_NUM) & "'" & _
                            ", " & CStr(vlTRISUserNumber) & ", CURRENT TIMESTAMP" & _
                            ", " & CStr(vlTRISUserNumber) & ", CURRENT TIMESTAMP);"

                        .MoveNext
                    End If
                Next llRecordIndex
            End If
        End With
    End If

    If (Trim(lsSQL) <> "") Then
        'Get Connection
        Set ladcnConnection = OpenConnection()
        'Call Routine to Execute SQL
        ExecuteOTHERSQL lsSQL, ladcnConnection
    End If
    'Close and destroy
    If Not (ladcnConnection Is Nothing) Then
        ladcnConnection.Close
        Set ladcnConnection = Nothing
    End If

    Set lobjUtility = Nothing

    dbProcessContactRequest = True

    GetObjectContext.SetComplete

    Exit Function
dbProcessContactRequest_EH:
    GetObjectContext.SetAbort

    Set lobjUtility = Nothing
    Set ladcnConnection = Nothing

    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "dbProcessContactRequest"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "dbProcessContactRequest" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "dbProcessContactRequest" _
                    , trErrdbSaveAbsenceContact _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
Public Function dbProcessRtnToWorkRequest(ByVal vadrsRtnToWrk As ADODB.Recordset _
                , ByVal vlTRISUserNumber As Long _
                , ByVal vsTRISUserID As String _
                ) As Boolean
On Error GoTo dbProcessRtnToWorkRequest_EH
'ADODB Connection Object
Dim ladcnConnection As ADODB.Connection
'Recordset
Dim llRecordIndex As Long, llFilterIndex As Long
'SQL
Dim lsSQL As String

    dbProcessRtnToWorkRequest = False

    If (RecordsExist(vadrsRtnToWrk)) Then
        With vadrsRtnToWrk
            .Filter = adFilterNone
            .Filter = "RECORD_STATUS = 'I'"
            If Not (.EOF) Then
                .MoveFirst

                For llRecordIndex = 1 To .RecordCount
                    lsSQL = lsSQL & _
                        " INSERT INTO METHDBA.TTR_ABS_RET_WRK" & _
                        " (ABS_NUM, CONFRM_RSN_IND" & _
                        ", ABS_REASON_TXT" & _
                        ", MEET_EMP_IND, MEET_EMP_TXT" & _
                        ", PRFM_DUTY_IND, PRFM_DUTY_TXT" & _
                        ", SIDE_EFF_IND, SIDE_EFF_TXT" & _
                        ", ADDL_TXT, RTN_TO_WRK_TS" & _
                        ", RPT_DT, RPTNG_EMP_NUM" & _
                        ", CREATED_EMP_NUM, CREATED_TIMESTAMP" & _
                        ", LAST_MOD_EMP_NUM, LAST_MOD_TIMESTAMP)" & _
                        " VALUES" & _
                        " (" & CStr(!ABS_NUM) & ", '" & Trim(!CONFRM_RSN_IND) & "'" & _
                        ", '" & Trim(!ABS_REASON_TXT) & "'" & _
                        ", '" & Trim(!MEET_EMP_IND) & "', '" & Trim(!MEET_EMP_TXT) & "'" & _
                        ", '" & Trim(!PRFM_DUTY_IND) & "', '" & Trim(!PRFM_DUTY_TXT) & "'" & _
                        ", '" & Trim(!SIDE_EFF_IND) & "', '" & Trim(!SIDE_EFF_TXT) & "'" & _
                        ", '" & Trim(!ADDL_TXT) & "', '" & Format(!RTN_TO_WRK_DT & " " & !RTN_TO_WRK_TM, "YYYY-MM-DD-hh.mm.ss.000000") & "'" & _
                        ", '" & Format(Trim(!RPT_DT), "MM/DD/YYYY") & "', " & CStr(!RPTNG_EMP_NUM) & _
                        ", " & CStr(vlTRISUserNumber) & ", CURRENT TIMESTAMP" & _
                        ", " & CStr(vlTRISUserNumber) & ", CURRENT TIMESTAMP);"

                    .MoveNext
                Next llRecordIndex
            End If
        End With
    End If

    If (Trim(lsSQL) <> "") Then
        'Get Connection
        Set ladcnConnection = OpenConnection()
        'Call Routine to Execute SQL
        ExecuteOTHERSQL lsSQL, ladcnConnection
    End If
    'Close and destroy
    If Not (ladcnConnection Is Nothing) Then
        ladcnConnection.Close
        Set ladcnConnection = Nothing
    End If

    dbProcessRtnToWorkRequest = True

    GetObjectContext.SetComplete

    Exit Function
dbProcessRtnToWorkRequest_EH:
    GetObjectContext.SetAbort

    Set ladcnConnection = Nothing

    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "dbProcessRtnToWorkRequest"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "dbProcessRtnToWorkRequest" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "dbProcessRtnToWorkRequest" _
                    , trErrdbSaveAbsenceReturnToWork _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
Public Function dbCallStart(ByVal vlAbsenceNumber As Long _
                , ByVal vsPhoneNumber As String _
                , ByVal vlTRISUserNumber As Long _
                ) As ADODB.Recordset
On Error GoTo dbCallStart_EH
'ADODB Connection Object
Dim ladcnConnection As ADODB.Connection
'SQL
Dim lsSQL As String
'RS
Dim ladrs As ADODB.Recordset, ladrsParms As ADODB.Recordset

    Set dbCallStart = Nothing

    lsSQL = "SELECT COALESCE(MAX(CALL_NUM) + 1, 1) AS CALL_NUM" & _
        " FROM  METHDBA.TTR_ABS_CALL_LOG" & _
        " WHERE ABS_NUM = ?" & _
        " FOR FETCH ONLY WITH UR;"
    Set ladrsParms = CreateParameterRS
    With ladrsParms
        .AddNew
        !PARM_ORDER = 1
        !PARM_NAME = ""
        !PARM_TYPE = adInteger
        !PARM_DIRECTION = adParamInput
        !PARM_SIZE = 0
        !PARM_VALUE = vlAbsenceNumber
        .Update
    End With

    Set ladrs = NewExecuteSQL(lsSQL _
        , ladrsParms _
        , adOpenForwardOnly _
        , adLockReadOnly _
        , adCmdText)

     If (RecordsExist(ladrs)) Then
        lsSQL = "INSERT INTO METHDBA.TTR_ABS_CALL_LOG" & _
            " (ABS_NUM, CALL_NUM" & _
            ", CALL_ST_TS, CALL_END_TS" & _
            ", PHONE_NUM, CALLER_EMP_NUM" & _
            ", CREATED_TIMESTAMP)" & _
            " VALUES " & _
            " (?, ?" & _
            ", CURRENT TIMESTAMP, CURRENT TIMESTAMP" & _
            ", ?, ?" & _
            ", CURRENT TIMESTAMP);"

        With ladrsParms
            .AddNew
            !PARM_ORDER = 2
            !PARM_NAME = ""
            !PARM_TYPE = adInteger
            !PARM_DIRECTION = adParamInput
            !PARM_SIZE = 0
            !PARM_VALUE = ladrs!CALL_NUM
            .Update

            .AddNew
            !PARM_ORDER = 3
            !PARM_NAME = ""
            !PARM_TYPE = adChar
            !PARM_DIRECTION = adParamInput
            !PARM_SIZE = 10
            !PARM_VALUE = Trim(vsPhoneNumber)
            .Update
        
            .AddNew
            !PARM_ORDER = 4
            !PARM_NAME = ""
            !PARM_TYPE = adInteger
            !PARM_DIRECTION = adParamInput
            !PARM_SIZE = 0
            !PARM_VALUE = vlTRISUserNumber
            .Update
        End With

        Set ladcnConnection = OpenConnection()
        'Call Routine to Execute SQL
        NewExecuteOtherSQL lsSQL _
            , ladrsParms _
            , ladcnConnection _
            , adCmdText

        'Close and destroy
        If Not (ladcnConnection Is Nothing) Then
            ladcnConnection.Close
        End If
        Set ladcnConnection = Nothing

        With ladrsParms
            .MoveFirst
            .Find "PARM_ORDER = 3"
            If Not (.EOF) Then
                .Delete adAffectCurrent
            End If

            .MoveFirst
            .Find "PARM_ORDER = 4"
            If Not (.EOF) Then
                .Delete adAffectCurrent
            End If

            .MoveFirst
        End With

        lsSQL = "SELECT ABS_NUM, CALL_NUM" & _
            "       , CHAR(CALL_ST_TS) AS CALL_ST_TS, CHAR(CALL_END_TS) AS CALL_END_TS" & _
            "       , PHONE_NUM, CALLER_EMP_NUM" & _
            "       , CHAR(CREATED_TIMESTAMP) AS CREATED_TIMESTAMP" & _
            " FROM  METHDBA.TTR_ABS_CALL_LOG" & _
            " WHERE ABS_NUM = ?" & _
            " AND   CALL_NUM = ?" & _
            " FOR FETCH ONLY WITH UR;"

        Set dbCallStart = NewExecuteSQL(lsSQL _
                        , ladrsParms _
                        , adOpenStatic _
                        , adLockReadOnly _
                        , adCmdText)
    End If

    Set ladcnConnection = Nothing
    Set ladrsParms = Nothing
    Set ladrs = Nothing

    GetObjectContext.SetComplete

    Exit Function
dbCallStart_EH:
    GetObjectContext.SetAbort

    Set ladrsParms = Nothing
    Set ladrs = Nothing
    Set ladcnConnection = Nothing

    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "dbCallStart"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "dbCallStart" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "dbCallStart" _
                    , trErrdbSaveAbsenceContactCallLog _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select
    
End Function
Public Function dbCallEnd(ByVal vadrsCallStart As ADODB.Recordset _
                ) As String
On Error GoTo dbCallEnd_EH
'ADODB Connection Object
Dim ladcnConnection As ADODB.Connection
'SQL
Dim lsSQL As String
'RS
Dim ladrs As ADODB.Recordset, ladrsParms As ADODB.Recordset

    dbCallEnd = ""

    If (RecordsExist(vadrsCallStart)) Then
        lsSQL = "UPDATE METHDBA.TTR_ABS_CALL_LOG" & _
            " SET   CALL_END_TS = CURRENT TIMESTAMP" & _
            " WHERE ABS_NUM = ?" & _
            " AND   CALL_NUM = ?;"

        Set ladrsParms = CreateParameterRS
        With ladrsParms
            .AddNew
            !PARM_ORDER = 1
            !PARM_NAME = ""
            !PARM_TYPE = adInteger
            !PARM_DIRECTION = adParamInput
            !PARM_SIZE = 0
            !PARM_VALUE = vadrsCallStart!ABS_NUM
            .Update

            .AddNew
            !PARM_ORDER = 2
            !PARM_NAME = ""
            !PARM_TYPE = adInteger
            !PARM_DIRECTION = adParamInput
            !PARM_SIZE = 0
            !PARM_VALUE = vadrsCallStart!CALL_NUM
            .Update
        End With

        'Get Connection
        Set ladcnConnection = OpenConnection()
        'Call Routine to Execute SQL
        NewExecuteOtherSQL lsSQL _
            , ladrsParms _
            , ladcnConnection _
            , adCmdText

        'Close and destroy
        If Not (ladcnConnection Is Nothing) Then
            ladcnConnection.Close
        End If
        Set ladcnConnection = Nothing

        lsSQL = "SELECT ABS_NUM, CALL_NUM" & _
            "       , CHAR(CALL_ST_TS) AS CALL_ST_TS, CHAR(CALL_END_TS) AS CALL_END_TS" & _
            "       , PHONE_NUM, CALLER_EMP_NUM" & _
            "       , CHAR(CREATED_TIMESTAMP) AS CREATED_TIMESTAMP" & _
            " FROM  METHDBA.TTR_ABS_CALL_LOG" & _
            " WHERE ABS_NUM = ?" & _
            " AND   CALL_NUM = ?" & _
            " FOR FETCH ONLY WITH UR;"

        Set ladrs = NewExecuteSQL(lsSQL _
                        , ladrsParms _
                        , adOpenForwardOnly _
                        , adLockReadOnly _
                        , adCmdText)
        Set ladrsParms = Nothing

        If (RecordsExist(ladrs)) Then
            dbCallEnd = ladrs!CALL_END_TS
        End If
        Set ladrs = Nothing
    End If

    GetObjectContext.SetComplete

    Exit Function
dbCallEnd_EH:
    GetObjectContext.SetAbort

    Set ladcnConnection = Nothing
    Set ladrsParms = Nothing
    Set ladrs = Nothing

    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "dbCallEnd"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "dbCallEnd" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "dbCallEnd" _
                    , trErrdbSaveAbsenceContactCallLog _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select
    
End Function
'Begin Prem TRIS 5.9
'Public Function dbProcessEAMRequest2(ByVal vlEmployeeNumber As Long _
'                , ByVal vadrsAbsenceDates As ADODB.Recordset _
'                , ByVal vadrsInitialAbsence As ADODB.Recordset _
'                , ByVal vadrsAbsenceProvider As ADODB.Recordset _
'                , ByVal vadrsAbsenceContact As ADODB.Recordset _
'                , ByVal vadrsAbsenceReturnToWork As ADODB.Recordset _
'                , ByVal vadrsAbsenceCallLog As ADODB.Recordset _
'                , ByVal vadrsWorkmanCompensation As ADODB.Recordset _
'                , ByVal vlTRISUserNumber As Long _
'                , ByVal vsTRISUserID As String _
'                , ByRef rlAbsenceNumber As Long _
'                , ByRef radrsBRV As ADODB.Recordset _
'                , Optional ByVal vadrsPortfolio As ADODB.Recordset = Nothing _
'                , Optional ByVal vadrsPortfolioDetail As ADODB.Recordset = Nothing _
'                ) As Boolean
'On Error GoTo dbProcessEAMRequest2_EH
''ADODB Connection Object
'Dim ladcnConnection As ADODB.Connection
''Recordset
'Dim llRecordIndex As Long, llFilterIndex As Long
''SQL
'Dim lsSQL As String
''SEQ_NUM
'Dim lobjUtility As utilDBDataServices.cuDBDataServices
'Dim llAbsenceNumber As Long, llProviderNumber As Long
'Dim llContactNumber As Long
''Special
'Dim lsOverride As String, llPortfolioNumber As Long
'', ByVal vadrsAbsenceExceptions As ADODB.Recordset
'
'    dbProcessEAMRequest2 = False
'
'    lsSQL = ""
'
'    If (RecordsExist(vadrsInitialAbsence)) Then
'        With vadrsInitialAbsence
'            .Filter = adFilterNone
'            .Filter = "RECORD_STATUS = 'D'"
'            If Not (.EOF) Then
'                .MoveFirst
'
'                For llRecordIndex = 1 To .RecordCount
'                    lsSQL = lsSQL & _
'                        " DELETE FROM METHDBA.TTR_ABS_PRTFL_DTL" & _
'                        " WHERE ABS_NUM = " & CStr(!ABS_NUM) & ";"
'
'                    lsSQL = lsSQL & _
'                        " DELETE FROM METHDBA.TTR_ABS_DATE" & _
'                        " WHERE ABS_NUM = " & CStr(!ABS_NUM) & ";"
'
'                    lsSQL = lsSQL & _
'                        " DELETE FROM METHDBA.TTR_ABS_RET_WRK" & _
'                        " WHERE ABS_NUM = " & CStr(!ABS_NUM) & ";"
'
'                    lsSQL = lsSQL & _
'                        " DELETE FROM METHDBA.TTR_ABS_CALL_LOG" & _
'                        " WHERE ABS_NUM = " & CStr(!ABS_NUM) & ";"
'
'                    lsSQL = lsSQL & _
'                        " DELETE FROM METHDBA.TTR_ABS_CONTACT" & _
'                        " WHERE ABS_NUM = " & CStr(!ABS_NUM) & ";"
'
'                    lsSQL = lsSQL & _
'                        " DELETE FROM METHDBA.TTR_ABS_PROVIDER" & _
'                        " WHERE ABS_NUM = " & CStr(!ABS_NUM) & ";"
'
'                    lsSQL = lsSQL & _
'                        " DELETE FROM METHDBA.TTR_ABS_INIT" & _
'                        " WHERE ABS_NUM = " & CStr(!ABS_NUM) & ";"
'
'                    .MoveNext
'                Next llRecordIndex
'
'                .Filter = adFilterNone
'                GoTo EXITNORMALLY
'            End If
'
'            .Filter = adFilterNone
'            .Filter = "RECORD_STATUS = 'I'"
'            If Not (.EOF) Then
'                .MoveFirst
'
'                For llRecordIndex = 1 To .RecordCount
'                    Set lobjUtility = CreateObject( _
'                                        "utilDBDataServices.cuDBDataServices")
'                    llAbsenceNumber = lobjUtility.GetNextSequenceNumber("EA")
'                    Set lobjUtility = Nothing
'
'                    AppendBusinessError radrsBRV _
'                        , trInfEAM _
'                        , mstrCBUSINESSINF _
'                        , "A new absence " & CStr(llAbsenceNumber) & " was created." _
'                        , "dbProcessEAMRequest2"
'
'                    If (llRecordIndex = 1) Then
'                        rlAbsenceNumber = llAbsenceNumber
'                    End If
'
'                    'NOTE: It is "Null" not Empty String or VBNULL
'                    'Reason for this special treatment: FOREIGN KEY FTRABI4
'
'                    lsOverride = "Null"
'                    If Not (IsNull(!OVER_EXC_CD)) Then
'                        If (Trim(!OVER_EXC_CD) <> "") Then
'                            lsOverride = "'" & Trim(!OVER_EXC_CD) & "'"
'                        End If
'                    End If
'
'                    lsSQL = lsSQL & _
'                        " INSERT INTO METHDBA.TTR_ABS_INIT" & _
'                        " (ABS_NUM, EMP_NUM" & _
'                        ", ABS_ST_TS, ABS_END_TS" & _
'                        ", ABS_LOC_CD, ABS_CALLIN_TS" & _
'                        ", EMP_PHONE, EMP_LOC_CD" & _
'                        ", EXCEPTION_CD, OVER_EXC_CD" & _
'                        ", RPTNG_EMP_NUM, RTN_TO_WRK_TS" & _
'                        ", FMLA_REASON_CD, ABS_PRIM_CONT_EMP" & _
'                        ", CREATED_EMP_NUM, CREATED_TIMESTAMP" & _
'                        ", LAST_MOD_EMP_NUM, LAST_MOD_TIMESTAMP" & _
'                        ", ABS_RPTR_TYPE, ABS_RPTR_TXT" & _
'                        ", ADDL_TXT, ABS_RPT_TS" & _
'                        ", RTW_NOT_REQD_IND, ABS_NOTIFY_WG_NUM" & _
'                        ", PRVDR_NM, PRVDR_PHN_NBR" & _
'                        ", FOLW_UP_DT)" & _
'                        " VALUES"
'                    lsSQL = lsSQL & _
'                        " (" & CStr(llAbsenceNumber) & ", " & CStr(vlEmployeeNumber) & _
'                        ", '" & Format(!ABS_ST_DT & " " & !ABS_ST_TM, "YYYY-MM-DD-hh.mm.ss.000000") & "'" & _
'                        ", '" & Format(!ABS_END_DT & " " & !ABS_END_TM, "YYYY-MM-DD-hh.mm.ss.000000") & "'" & _
'                        ", " & CStr(!ABS_LOC_CD) & ", '" & Format(!ABS_CALLIN_DT & " " & !ABS_CALLIN_TM, "YYYY-MM-DD-hh.mm.ss.000000") & "'" & _
'                        ", '" & Trim(!EMP_PHONE) & "', " & CStr(!EMP_LOC_CD) & _
'                        ", '" & Trim(!EXCEPTION_CD) & "', " & lsOverride & _
'                        ", " & CStr(!RPTNG_EMP_NUM) & _
'                        ", '" & Format(!RTN_TO_WRK_DT & " " & !RTN_TO_WRK_TM, "YYYY-MM-DD-hh.mm.ss.000000") & "'" & _
'                        ", '" & Trim(!FMLA_REASON) & "', " & CStr(!ABS_PRIM_CONT_EMP) & _
'                        ", " & CStr(vlTRISUserNumber) & ", CURRENT TIMESTAMP" & _
'                        ", " & CStr(vlTRISUserNumber) & ", CURRENT TIMESTAMP" & _
'                        ", '" & Trim(!ABS_RPTR_TYPE) & "', '" & Trim(!ABS_RPTR_TXT) & "'" & _
'                        ", '" & Trim(!ADDL_TXT) & "', '" & Format(!ABS_RPT_TS, "YYYY-MM-DD-hh.mm.ss.000000") & "'" & _
'                        ", '" & Trim(!RTW_NOT_REQD_IND) & "', " & CStr(!ABS_NOTIFY_WG_NUM) & _
'                        ", '" & Trim(!PRVDR_NM) & "', '" & Trim(!PRVDR_PHN_NBR) & "'" & _
'                        ", '" & Format(Trim(!FOLW_UP_DT), "MM/DD/YYYY") & "');"
'
'                    If (RecordsExist(vadrsAbsenceDates)) Then
'                        With vadrsAbsenceDates
'                            .Filter = adFilterNone
'                            .Filter = "ABS_NUM = " & _
'                                    CStr(vadrsInitialAbsence!ABS_NUM)
'                            If Not (.EOF) Then
'                                .MoveFirst
'
'                                For llFilterIndex = 1 To .RecordCount
'                                    !ABS_NUM = llAbsenceNumber
'                                    .Update
'
'                                    .MoveNext
'                                Next llFilterIndex
'                            End If
'
'                            .Filter = adFilterNone
'                        End With
'                    End If
'
'                    If (RecordsExist(vadrsAbsenceProvider)) Then
'                        With vadrsAbsenceProvider
'                            .Filter = adFilterNone
'                            .Filter = "ABS_NUM = " & _
'                                    CStr(vadrsInitialAbsence!ABS_NUM)
'                            If Not (.EOF) Then
'                                .MoveFirst
'
'                                For llFilterIndex = 1 To .RecordCount
'                                    !ABS_NUM = llAbsenceNumber
'                                    .Update
'
'                                    .MoveNext
'                                Next llFilterIndex
'                            End If
'
'                            .Filter = adFilterNone
'                        End With
'                    End If
'
'                    If (RecordsExist(vadrsAbsenceContact)) Then
'                        With vadrsAbsenceContact
'                            .Filter = adFilterNone
'                            .Filter = "ABS_NUM = " & _
'                                    CStr(vadrsInitialAbsence!ABS_NUM)
'                            If Not (.EOF) Then
'                                .MoveFirst
'
'                                For llFilterIndex = 1 To .RecordCount
'                                    !ABS_NUM = llAbsenceNumber
'                                    .Update
'
'                                    .MoveNext
'                                Next llFilterIndex
'                            End If
'
'                            .Filter = adFilterNone
'                        End With
'                    End If
'
'                    If (RecordsExist(vadrsAbsenceReturnToWork)) Then
'                        With vadrsAbsenceReturnToWork
'                            .Filter = adFilterNone
'                            .Filter = "ABS_NUM = " & _
'                                    CStr(vadrsInitialAbsence!ABS_NUM)
'                            If Not (.EOF) Then
'                                .MoveFirst
'
'                                For llFilterIndex = 1 To .RecordCount
'                                    !ABS_NUM = llAbsenceNumber
'                                    .Update
'
'                                    .MoveNext
'                                Next llFilterIndex
'                            End If
'
'                            .Filter = adFilterNone
'                        End With
'                    End If
'
'                    If (RecordsExist(vadrsWorkmanCompensation)) Then
'                        With vadrsWorkmanCompensation
'                            .Filter = adFilterNone
'                            .Filter = "ABS_NUM = " & _
'                                    CStr(vadrsInitialAbsence!ABS_NUM)
'                            If Not (.EOF) Then
'                                .MoveFirst
'
'                                For llFilterIndex = 1 To .RecordCount
'                                    !ABS_NUM = llAbsenceNumber
'                                    .Update
'
'                                    .MoveNext
'                                Next llFilterIndex
'                            End If
'
'                            .Filter = adFilterNone
'                        End With
'                    End If
'
'                    If (RecordsExist(vadrsPortfolioDetail)) Then
'                        With vadrsPortfolioDetail
'                            .Filter = adFilterNone
'                            .Filter = "ABS_NUM = " & _
'                                    CStr(vadrsInitialAbsence!ABS_NUM)
'                            If Not (.EOF) Then
'                                .MoveFirst
'
'                                For llFilterIndex = 1 To .RecordCount
'                                    !ABS_NUM = llAbsenceNumber
'                                    .Update
'
'                                    .MoveNext
'                                Next llFilterIndex
'                            End If
'
'                            .Filter = adFilterNone
'                        End With
'                    End If
'
'                    If (RecordsExist(vadrsAbsenceCallLog)) Then
'                        With vadrsAbsenceCallLog
'                            .Filter = adFilterNone
'                            .Filter = "ABS_NUM = " & _
'                                    CStr(vadrsInitialAbsence!ABS_NUM) & _
'                                    " AND RECORD_STATUS = 'I'"
'                            If Not (.EOF) Then
'                                .MoveFirst
'
'                                For llFilterIndex = 1 To .RecordCount
'                                    !ABS_NUM = llAbsenceNumber
'                                    !CALL_NUM = llFilterIndex
'                                    .Update
'
'                                    .MoveNext
'                                Next llFilterIndex
'                            End If
'
'                            .Filter = adFilterNone
'                        End With
'                    End If
'
'                    .MoveNext
'                Next llRecordIndex
'            End If
'
'            .Filter = adFilterNone
'            .Filter = "RECORD_STATUS = 'C'"
'            If Not (.EOF) Then
'                .MoveFirst
'
'                For llRecordIndex = 1 To .RecordCount
'                    lsOverride = "Null"
'                    If Not (IsNull(!OVER_EXC_CD)) Then
'                        If (Trim(!OVER_EXC_CD) <> "") Then
'                            lsOverride = "'" & Trim(!OVER_EXC_CD) & "'"
'                        End If
'                    End If
'
'                    lsSQL = lsSQL & _
'                        " UPDATE    METHDBA.TTR_ABS_INIT" & _
'                        " SET       ABS_ST_TS = '" & Format(!ABS_ST_DT & " " & !ABS_ST_TM, "YYYY-MM-DD-hh.mm.ss.000000") & "'" & _
'                        "           , ABS_END_TS = '" & Format(!ABS_END_DT & " " & !ABS_END_TM, "YYYY-MM-DD-hh.mm.ss.000000") & "'" & _
'                        "           , ABS_CALLIN_TS = '" & Format(!ABS_CALLIN_DT & " " & !ABS_CALLIN_TM, "YYYY-MM-DD-hh.mm.ss.000000") & "'" & _
'                        "           , EMP_PHONE = '" & Trim(!EMP_PHONE) & "'" & _
'                        "           , EMP_LOC_CD = " & CStr(!EMP_LOC_CD) & _
'                        "           , EXCEPTION_CD = '" & Trim(!EXCEPTION_CD) & "'" & _
'                        "           , OVER_EXC_CD = " & lsOverride & _
'                        "           , RPTNG_EMP_NUM = " & CStr(!RPTNG_EMP_NUM) & _
'                        "           , RTN_TO_WRK_TS = '" & Format(!RTN_TO_WRK_DT & " " & !RTN_TO_WRK_TM, "YYYY-MM-DD-hh.mm.ss.000000") & "'" & _
'                        "           , FMLA_REASON_CD = '" & Trim(!FMLA_REASON) & "'" & _
'                        "           , ABS_PRIM_CONT_EMP = " & CStr(!ABS_PRIM_CONT_EMP) & _
'                        "           , LAST_MOD_EMP_NUM = " & CStr(vlTRISUserNumber) & _
'                        "           , LAST_MOD_TIMESTAMP = CURRENT TIMESTAMP" & _
'                        "           , ABS_RPTR_TYPE = '" & Trim(!ABS_RPTR_TYPE) & "'" & _
'                        "           , ABS_RPTR_TXT = '" & Trim(!ABS_RPTR_TXT) & "'" & _
'                        "           , ADDL_TXT = '" & Trim(!ADDL_TXT) & "'" & _
'                        "           , FOLW_UP_DT = '" & Format(Trim(!FOLW_UP_DT), "MM/DD/YYYY") & "'" & _
'                        "           , RTW_NOT_REQD_IND = '" & Trim(!RTW_NOT_REQD_IND) & "'" & _
'                        "           , ABS_NOTIFY_WG_NUM = " & CStr(!ABS_NOTIFY_WG_NUM) & _
'                        "           , PRVDR_NM = '" & Trim(!PRVDR_NM) & "'" & _
'                        "           , PRVDR_PHN_NBR = '" & Trim(!PRVDR_PHN_NBR) & "'" & _
'                        " WHERE     ABS_NUM = " & CStr(!ABS_NUM) & _
'                        " AND       CREATED_TIMESTAMP = '" & Trim(!CREATED_TIMESTAMP) & "';"
'
'                    .MoveNext
'                Next llRecordIndex
'            End If
'
'            .Filter = adFilterNone
'        End With
'    End If
'
'    If (RecordsExist(vadrsAbsenceDates)) Then
'        With vadrsAbsenceDates
'            .Filter = adFilterNone
'            .Filter = "RECORD_STATUS = 'D'"
'            If Not (.EOF) Then
'                .MoveFirst
'
'                For llRecordIndex = 1 To .RecordCount
'                    lsSQL = lsSQL & _
'                        " DELETE FROM METHDBA.TTR_ABS_DATE" & _
'                        " WHERE ABS_NUM = " & CStr(!ABS_NUM) & _
'                        " AND   ABS_DATE = '" & CStr(!ABS_DATE) & "';"
'
'                    .MoveNext
'                Next llRecordIndex
'            End If
'
'            .Filter = adFilterNone
'            .Filter = "RECORD_STATUS = 'I'"
'            If Not (.EOF) Then
'                .MoveFirst
'
'                For llRecordIndex = 1 To .RecordCount
'                    lsSQL = lsSQL & _
'                        " INSERT INTO METHDBA.TTR_ABS_DATE" & _
'                        " (ABS_NUM, ABS_DATE" & _
'                        ", CREATED_EMP_NUM, CREATED_TIMESTAMP" & _
'                        ", LAST_MOD_EMP_NUM, LAST_MOD_TIMESTAMP)" & _
'                        " VALUES" & _
'                        " (" & CStr(!ABS_NUM) & ", '" & CStr(!ABS_DATE) & "'" & _
'                        ", " & CStr(vlTRISUserNumber) & ", CURRENT TIMESTAMP" & _
'                        ", " & CStr(vlTRISUserNumber) & ", CURRENT TIMESTAMP);"
'
'                    .MoveNext
'                Next llRecordIndex
'            End If
'
'            .Filter = adFilterNone
'        End With
'    End If
'
'    If (RecordsExist(vadrsAbsenceProvider)) Then
'        With vadrsAbsenceProvider
'            .Filter = adFilterNone
'            .Filter = "RECORD_STATUS = 'D'"
'            If Not (.EOF) Then
'                .MoveFirst
'
'                For llRecordIndex = 1 To .RecordCount
'                    lsSQL = lsSQL & _
'                        " DELETE FROM METHDBA.TTR_ABS_PROVIDER" & _
'                        " WHERE ABS_NUM = " & CStr(!ABS_NUM) & _
'                        " AND   PRVDR_NBR = " & CStr(!PRVDR_NBR) & _
'                        " AND   CREATED_TIMESTAMP = '" & Trim(!CREATED_TIMESTAMP) & "';"
'
'                    .MoveNext
'                Next llRecordIndex
'            End If
'
'            .Filter = adFilterNone
'            .Filter = "RECORD_STATUS = 'I'"
'            If Not (.EOF) Then
'                .MoveFirst
'
'                For llRecordIndex = 1 To .RecordCount
'                    If (!PRVDR_NBR = -1) Then
'                        Set lobjUtility = CreateObject( _
'                                            "utilDBDataServices.cuDBDataServices")
'                        llProviderNumber = lobjUtility.GetNextSequenceNumber("AP")
'                        Set lobjUtility = Nothing
'
'                        !PRVDR_NBR = llProviderNumber
'                        .Update
'
'                        lsSQL = lsSQL & _
'                            " INSERT INTO METHDBA.TTR_PROVIDER" & _
'                            " (PRVDR_NBR, PRVDR_NM" & _
'                            ", LN1_ADDR, LN2_ADDR" & _
'                            ", CTY_NM, ST_CD" & _
'                            ", ZIP_CD, PHN_NBR" & _
'                            ", CLS_TXT, TYP_CD)" & _
'                            " VALUES" & _
'                            " (" & CStr(!PRVDR_NBR) & ", '" & Trim(!PRVDR_NM) & "'" & _
'                            ", '" & Trim(!LN1_ADDR) & "', '" & Trim(!LN2_ADDR) & "'" & _
'                            ", '" & Trim(!CTY_NM) & "', '" & Trim(!ST_CD) & "'" & _
'                            ", '" & Trim(!ZIP_CD) & "', '" & Trim(!PHN_NBR) & "'" & _
'                            ", '" & Trim(!CLS_TXT) & "', '" & Trim(!TYP_DESC) & "');"
'                    End If
'
'                    lsSQL = lsSQL & _
'                        " INSERT INTO METHDBA.TTR_ABS_PROVIDER" & _
'                        " (ABS_NUM, PRVDR_NBR" & _
'                        ", VST_ST_ADM_TS, VST_END_DGCE_TS" & _
'                        ", ADDL_TXT, RLSE_IND" & _
'                        ", CREATED_EMP_NUM, CREATED_TIMESTAMP" & _
'                        ", LAST_MOD_EMP_NUM, LAST_MOD_TIMESTAMP)" & _
'                        " VALUES" & _
'                        " (" & CStr(!ABS_NUM) & ", " & CStr(!PRVDR_NBR) & _
'                        ", '" & Format(!VST_ST_ADM_DT & " " & !VST_ST_ADM_TM, "YYYY-MM-DD-hh.mm.ss.000000") & "'" & _
'                        ", '" & Format(!VST_END_DGCE_DT & " " & !VST_END_DGCE_TM, "YYYY-MM-DD-hh.mm.ss.000000") & "'" & _
'                        ", '" & Trim(!ADDL_TXT) & "', '" & Trim(!RLSE_IND) & "'" & _
'                        ", " & CStr(vlTRISUserNumber) & ", CURRENT TIMESTAMP" & _
'                        ", " & CStr(vlTRISUserNumber) & ", CURRENT TIMESTAMP);"
'
'                    .MoveNext
'                Next llRecordIndex
'            End If
'
'            .Filter = adFilterNone
'            .Filter = "RECORD_STATUS = 'C'"
'            If Not (.EOF) Then
'                .MoveFirst
'
'                For llRecordIndex = 1 To .RecordCount
'                    lsSQL = lsSQL & _
'                        " UPDATE METHDBA.TTR_PROVIDER" & _
'                        " SET   PRVDR_NM = '" & Trim(!PRVDR_NM) & "'" & _
'                        "       , LN1_ADDR = '" & Trim(!LN1_ADDR) & "'" & _
'                        "       , LN2_ADDR = '" & Trim(!LN2_ADDR) & "'" & _
'                        "       , CTY_NM = '" & Trim(!CTY_NM) & "'" & _
'                        "       , ST_CD = '" & Trim(!ST_CD) & "'" & _
'                        "       , ZIP_CD = '" & Trim(!ZIP_CD) & "'" & _
'                        "       , PHN_NBR = '" & Trim(!PHN_NBR) & "'" & _
'                        "       , CLS_TXT = '" & Trim(!CLS_TXT) & "'" & _
'                        "       , TYP_CD = '" & Trim(!TYP_DESC) & "'" & _
'                        " WHERE PRVDR_NBR = " & CStr(!PRVDR_NBR) & ";"
'
'                    lsSQL = lsSQL & _
'                        " UPDATE METHDBA.TTR_ABS_PROVIDER" & _
'                        " SET   PRVDR_NBR = " & CStr(!PRVDR_NBR) & _
'                        "       , VST_ST_ADM_TS = '" & Format(!VST_ST_ADM_DT & " " & !VST_ST_ADM_TM, "YYYY-MM-DD-hh.mm.ss.000000") & "'" & _
'                        "       , VST_END_DGCE_TS = '" & Format(!VST_END_DGCE_DT & " " & !VST_END_DGCE_TM, "YYYY-MM-DD-hh.mm.ss.000000") & "'" & _
'                        "       , RLSE_IND = '" & Trim(!RLSE_IND) & "'" & _
'                        "       , ADDL_TXT = '" & Trim(!ADDL_TXT) & "'" & _
'                        "       , LAST_MOD_EMP_NUM = " & CStr(vlTRISUserNumber) & _
'                        "       , LAST_MOD_TIMESTAMP = CURRENT TIMESTAMP" & _
'                        " WHERE ABS_NUM = " & CStr(!ABS_NUM) & _
'                        " AND   CREATED_TIMESTAMP = '" & Trim(!CREATED_TIMESTAMP) & "';"
'
'                    .MoveNext
'                Next llRecordIndex
'            End If
'
'            .Filter = adFilterNone
'        End With
'    End If
'
'    If (RecordsExist(vadrsAbsenceContact)) Then
'        With vadrsAbsenceContact
'            .Filter = adFilterNone
'            .Filter = "RECORD_STATUS = 'D'"
'            If Not (.EOF) Then
'                .MoveFirst
'
'                For llRecordIndex = 1 To .RecordCount
'                    lsSQL = lsSQL & _
'                        " DELETE FROM METHDBA.TTR_ABS_CONTACT" & _
'                        " WHERE ABS_NUM = " & CStr(!ABS_NUM) & _
'                        " AND   ABS_CONT_NUM = " & CStr(!ABS_CONT_NUM) & _
'                        " AND   CREATED_TIMESTAMP = '" & Trim(!CREATED_TIMESTAMP) & "';"
'
'                    .MoveNext
'                Next llRecordIndex
'            End If
'
'            .Filter = adFilterNone
'            .Filter = "RECORD_STATUS = 'I'"
'            If Not (.EOF) Then
'                .MoveFirst
'
'                For llRecordIndex = 1 To .RecordCount
'                    If (!ABS_CONT_NUM = -1) Then
'                        Set lobjUtility = CreateObject( _
'                                            "utilDBDataServices.cuDBDataServices")
'                        llContactNumber = lobjUtility.GetNextSequenceNumber("AC")
'                        Set lobjUtility = Nothing
'
'                        !ABS_CONT_NUM = llContactNumber
'                        .Update
'                    End If
'
'                    lsSQL = lsSQL & _
'                        " INSERT INTO METHDBA.TTR_ABS_CONTACT" & _
'                        " (ABS_NUM, ABS_CONT_NUM" & _
'                        ", ABS_CONT_TS, USR_INIT_IND" & _
'                        ", TALK_TO_EMP_IND, TALK_TO_EMP_TXT" & _
'                        ", EMP_TO_DR_IND, EMP_TO_DR_TXT" & _
'                        ", OTH_TRTMNT_IND, OTH_TRTMNT_TXT" & _
'                        ", ADVS_SL_IND, ADVS_SL_TXT" & _
'                        ", PHONE_NUM, RPTNG_EMP_NUM" & _
'                        ", ADDL_TXT, PRVDR_PHN_NUM" & _
'                        ", CREATED_EMP_NUM, CREATED_TIMESTAMP" & _
'                        ", LAST_MOD_EMP_NUM, LAST_MOD_TIMESTAMP)" & _
'                        " VALUES" & _
'                        " (" & CStr(!ABS_NUM) & ", " & CStr(!ABS_CONT_NUM) & _
'                        ", '" & Format(!ABS_CONT_DT & " " & !ABS_CONT_TM, "YYYY-MM-DD-hh.mm.ss.000000") & "', '" & Trim(!USR_INIT_IND) & "'" & _
'                        ", '" & Trim(!TALK_TO_EMP_IND) & "', '" & Trim(!TALK_TO_EMP_TXT) & "'" & _
'                        ", '" & Trim(!EMP_TO_DR_IND) & "', '" & Trim(!EMP_TO_DR_TXT) & "'" & _
'                        ", '" & Trim(!OTH_TRTMNT_IND) & "', '" & Trim(!OTH_TRTMNT_TXT) & "'" & _
'                        ", '" & Trim(!ADVS_SL_IND) & "', '" & Trim(!ADVS_SL_TXT) & "'" & _
'                        ", '" & Trim(!PHONE_NUM) & "', " & CStr(!RPTNG_EMP_NUM) & _
'                        ", '" & Trim(!ADDL_TXT) & "', '" & Trim(!PRVDR_PHN_NUM) & "'" & _
'                        ", " & CStr(vlTRISUserNumber) & ", CURRENT TIMESTAMP" & _
'                        ", " & CStr(vlTRISUserNumber) & ", CURRENT TIMESTAMP);"
'
'                    .MoveNext
'                Next llRecordIndex
'            End If
'
'            .Filter = adFilterNone
'            .Filter = "RECORD_STATUS = 'C'"
'            If Not (.EOF) Then
'                .MoveFirst
'
'                For llRecordIndex = 1 To .RecordCount
'
'                    lsSQL = lsSQL & _
'                        " UPDATE METHDBA.TTR_ABS_CONTACT" & _
'                        " SET   ABS_CONT_TS = '" & Format(!ABS_CONT_DT & " " & !ABS_CONT_TM, "YYYY-MM-DD-hh.mm.ss.000000") & "'" & _
'                        "       , USR_INIT_IND = '" & Trim(!USR_INIT_IND) & "'" & _
'                        "       , TALK_TO_EMP_IND = '" & Trim(!TALK_TO_EMP_IND) & "'" & _
'                        "       , TALK_TO_EMP_TXT = '" & Trim(!TALK_TO_EMP_TXT) & "'" & _
'                        "       , EMP_TO_DR_IND = '" & Trim(!EMP_TO_DR_IND) & "'" & _
'                        "       , EMP_TO_DR_TXT = '" & Trim(!EMP_TO_DR_TXT) & "'" & _
'                        "       , OTH_TRTMNT_IND = '" & Trim(!OTH_TRTMNT_IND) & "'" & _
'                        "       , OTH_TRTMNT_TXT = '" & Trim(!OTH_TRTMNT_TXT) & "'" & _
'                        "       , ADVS_SL_IND = '" & Trim(!ADVS_SL_IND) & "'" & _
'                        "       , ADVS_SL_TXT = '" & Trim(!ADVS_SL_TXT) & "'" & _
'                        "       , PHONE_NUM = '" & Trim(!PHONE_NUM) & "'" & _
'                        "       , RPTNG_EMP_NUM = " & CStr(!RPTNG_EMP_NUM) & _
'                        "       , ADDL_TXT = '" & Trim(!ADDL_TXT) & "'" & _
'                        "       , PRVDR_PHN_NUM = '" & Trim(!PRVDR_PHN_NUM) & "'" & _
'                        "       , LAST_MOD_EMP_NUM = " & CStr(vlTRISUserNumber) & _
'                        "       , LAST_MOD_TIMESTAMP = CURRENT TIMESTAMP" & _
'                        " WHERE ABS_NUM = " & CStr(!ABS_NUM) & _
'                        " AND   ABS_CONT_NUM = " & CStr(!ABS_CONT_NUM) & _
'                        " AND   CREATED_TIMESTAMP = '" & Trim(!CREATED_TIMESTAMP) & "';"
'
'                    .MoveNext
'                Next llRecordIndex
'            End If
'
'            .Filter = adFilterNone
'        End With
'    End If
'
'    If (RecordsExist(vadrsAbsenceReturnToWork)) Then
'        With vadrsAbsenceReturnToWork
'            .Filter = adFilterNone
'            .Filter = "RECORD_STATUS = 'D'"
'            If Not (.EOF) Then
'                .MoveFirst
'
'                For llRecordIndex = 1 To .RecordCount
'                    lsSQL = lsSQL & _
'                        " DELETE FROM METHDBA.TTR_ABS_RET_WRK" & _
'                        " WHERE ABS_NUM = " & CStr(!ABS_NUM) & _
'                        " AND   CREATED_TIMESTAMP = '" & Trim(!CREATED_TIMESTAMP) & "';"
'
'                    .MoveNext
'                Next llRecordIndex
'            End If
'
'            .Filter = adFilterNone
'            .Filter = "RECORD_STATUS = 'I'"
'            If Not (.EOF) Then
'                .MoveFirst
'
'                For llRecordIndex = 1 To .RecordCount
'                    lsSQL = lsSQL & _
'                        " INSERT INTO METHDBA.TTR_ABS_RET_WRK" & _
'                        " (ABS_NUM, CONFRM_RSN_IND" & _
'                        ", ABS_REASON_TXT" & _
'                        ", MEET_EMP_IND, MEET_EMP_TXT" & _
'                        ", PRFM_DUTY_IND, PRFM_DUTY_TXT" & _
'                        ", SIDE_EFF_IND, SIDE_EFF_TXT" & _
'                        ", ADDL_TXT, RTN_TO_WRK_TS" & _
'                        ", RPT_DT, RPTNG_EMP_NUM" & _
'                        ", CREATED_EMP_NUM, CREATED_TIMESTAMP" & _
'                        ", LAST_MOD_EMP_NUM, LAST_MOD_TIMESTAMP)" & _
'                        " VALUES" & _
'                        " (" & CStr(!ABS_NUM) & ", '" & Trim(!CONFRM_RSN_IND) & "'" & _
'                        ", '" & Trim(!ABS_REASON_TXT) & "'" & _
'                        ", '" & Trim(!MEET_EMP_IND) & "', '" & Trim(!MEET_EMP_TXT) & "'" & _
'                        ", '" & Trim(!PRFM_DUTY_IND) & "', '" & Trim(!PRFM_DUTY_TXT) & "'" & _
'                        ", '" & Trim(!SIDE_EFF_IND) & "', '" & Trim(!SIDE_EFF_TXT) & "'" & _
'                        ", '" & Trim(!ADDL_TXT) & "', '" & Format(!RTN_TO_WRK_DT & " " & !RTN_TO_WRK_TM, "YYYY-MM-DD-hh.mm.ss.000000") & "'" & _
'                        ", '" & Format(Trim(!RPT_DT), "MM/DD/YYYY") & "', " & CStr(!RPTNG_EMP_NUM) & _
'                        ", " & CStr(vlTRISUserNumber) & ", CURRENT TIMESTAMP" & _
'                        ", " & CStr(vlTRISUserNumber) & ", CURRENT TIMESTAMP);"
'
'                    .MoveNext
'                Next llRecordIndex
'            End If
'
'            .Filter = adFilterNone
'            .Filter = "RECORD_STATUS = 'C'"
'            If Not (.EOF) Then
'                .MoveFirst
'
'                For llRecordIndex = 1 To .RecordCount
'                    lsSQL = lsSQL & _
'                        " UPDATE METHDBA.TTR_ABS_RET_WRK" & _
'                        " SET   CONFRM_RSN_IND = '" & Trim(!CONFRM_RSN_IND) & "'" & _
'                        "       , ABS_REASON_TXT = '" & Trim(!ABS_REASON_TXT) & "'" & _
'                        "       , MEET_EMP_IND = '" & Trim(!MEET_EMP_IND) & "'" & _
'                        "       , MEET_EMP_TXT = '" & Trim(!MEET_EMP_TXT) & "'" & _
'                        "       , PRFM_DUTY_IND = '" & Trim(!PRFM_DUTY_IND) & "'" & _
'                        "       , PRFM_DUTY_TXT = '" & Trim(!PRFM_DUTY_TXT) & "'" & _
'                        "       , SIDE_EFF_IND = '" & Trim(!SIDE_EFF_IND) & "'" & _
'                        "       , SIDE_EFF_TXT = '" & Trim(!SIDE_EFF_TXT) & "'" & _
'                        "       , ADDL_TXT = '" & Trim(!ADDL_TXT) & "'" & _
'                        "       , RTN_TO_WRK_TS = '" & Format(!RTN_TO_WRK_DT & " " & !RTN_TO_WRK_TM, "YYYY-MM-DD-hh.mm.ss.000000") & "'" & _
'                        "       , RPT_DT = '" & Format(Trim(!RPT_DT), "MM/DD/YYYY") & "'" & _
'                        "       , RPTNG_EMP_NUM = " & CStr(!RPTNG_EMP_NUM) & _
'                        "       , LAST_MOD_EMP_NUM = " & CStr(vlTRISUserNumber) & _
'                        "       , LAST_MOD_TIMESTAMP = CURRENT TIMESTAMP" & _
'                        " WHERE ABS_NUM = " & CStr(!ABS_NUM) & _
'                        " AND   CREATED_TIMESTAMP = '" & Trim(!CREATED_TIMESTAMP) & "';"
'
'                    .MoveNext
'                Next llRecordIndex
'            End If
'
'            .Filter = adFilterNone
'        End With
'    End If
'
'    If (RecordsExist(vadrsAbsenceCallLog)) Then
'        With vadrsAbsenceCallLog
'            .Filter = adFilterNone
'            .Filter = "RECORD_STATUS = 'I'"
'            If Not (.EOF) Then
'                For llRecordIndex = 1 To .RecordCount
'                    lsSQL = lsSQL & _
'                        " INSERT INTO METHDBA.TTR_ABS_CALL_LOG" & _
'                        " (ABS_NUM, CALL_NUM" & _
'                        ", CALL_ST_TS, CALL_END_TS" & _
'                        ", PHONE_NUM, CALLER_EMP_NUM" & _
'                        ", CREATED_TIMESTAMP)" & _
'                        " VALUES " & _
'                        " (" & CStr(!ABS_NUM) & ", " & CStr(!CALL_NUM) & _
'                        ", '" & Trim(!CALL_ST_TS) & "', '" & Trim(!CALL_END_TS) & "'" & _
'                        ", '" & Trim(!PHONE_NUM) & "', " & CStr(vlTRISUserNumber) & _
'                        ", CURRENT TIMESTAMP);"
'
'                    .MoveNext
'                Next llRecordIndex
'            End If
'        End With
'    End If
'
'    If (RecordsExist(vadrsWorkmanCompensation)) Then
'        With vadrsWorkmanCompensation
'            .Filter = adFilterNone
'
'            .Filter = "RECORD_STATUS = 'D'"
'            If Not (.EOF) Then
'                .MoveFirst
'
'                For llRecordIndex = 1 To .RecordCount
'                    lsSQL = lsSQL & _
'                        " DELETE FROM METHDBA.TTR_ABS_WC" & _
'                        " WHERE ABS_NUM = " & CStr(!ABS_NUM) & _
'                        " AND   CREATED_TIMESTAMP = '" & Trim(!CREATED_TIMESTAMP) & "';"
'
'                    .MoveNext
'                Next llRecordIndex
'            End If
'
'            .Filter = adFilterNone
'            .Filter = "RECORD_STATUS = 'I'"
'            If Not (.EOF) Then
'                .MoveFirst
'
'                For llRecordIndex = 1 To .RecordCount
'                    lsSQL = lsSQL & _
'                        " INSERT INTO METHDBA.TTR_ABS_WC" & _
'                        "(ABS_NUM, INCD_RPTD_IND" & _
'                        ", INCD_OCCURRED_DTM, INJ_LOC_NM" & _
'                        ", INJD_BODY_PART_TXT, WITNESSES_NM" & _
'                        ", CO_DCTR_VISIT_IND" & _
'                        ", CREATED_EMP_NUM, CREATED_TIMESTAMP" & _
'                        ", LAST_MOD_EMP_NUM, LAST_MOD_TIMESTAMP)" & _
'                        " VALUES " & _
'                        "(" & CStr(!ABS_NUM) & ", '" & !INCD_RPTD_IND & "'" & _
'                        ", '" & !INCD_OCCURRED_DTM & "', '" & !INJ_LOC_NM & "'" & _
'                        ", '" & !INJD_BODY_PART_TXT & "', '" & !WITNESSES_NM & "'" & _
'                        ", '" & !CO_DCTR_VISIT_IND & "'" & _
'                        ", " & CStr(vlTRISUserNumber) & ", CURRENT TIMESTAMP" & _
'                        ", " & CStr(vlTRISUserNumber) & ", CURRENT TIMESTAMP);"
'
'                    .MoveNext
'                Next llRecordIndex
'            End If
'
'            .Filter = adFilterNone
'            .Filter = "RECORD_STATUS = 'C'"
'            If Not (.EOF) Then
'                .MoveFirst
'
'                For llRecordIndex = 1 To .RecordCount
'                    lsSQL = lsSQL & _
'                        " UPDATE METHDBA.TTR_ABS_WC" & _
'                        " SET   INCD_RPTD_IND = '" & !INCD_RPTD_IND & "'" & _
'                        "       , INCD_OCCURRED_DTM = '" & !INCD_OCCURRED_DTM & "'" & _
'                        "       , INJD_BODY_PART_TXT = '" & !INJD_BODY_PART_TXT & "'" & _
'                        "       , WITNESSES_NM = '" & !WITNESSES_NM & "'" & _
'                        "       , CO_DCTR_VISIT_IND = '" & !CO_DCTR_VISIT_IND & "'" & _
'                        "       , LAST_MOD_EMP_NUM = " & CStr(vlTRISUserNumber) & _
'                        "       , LAST_MOD_TIMESTAMP = CURRENT TIMESTAMP" & _
'                        " WHERE ABS_NUM = " & CStr(!ABS_NUM) & _
'                        " AND   CREATED_TIMESTAMP = '" & Trim(!CREATED_TIMESTAMP) & "';"
'
'                    .MoveNext
'                Next llRecordIndex
'            End If
'        End With
'    End If
'
'    If (RecordsExist(vadrsPortfolio)) Then
'        With vadrsPortfolio
'            .Filter = adFilterNone
'            .Filter = "RECORD_STATUS = 'I'"
'            If Not (.EOF) Then
'                If (!PRTFL_NUM = -1) Then
'                    Set lobjUtility = CreateObject( _
'                                        "utilDBDataServices.cuDBDataServices")
'                    llPortfolioNumber = lobjUtility.GetNextSequenceNumber("EP")
'                    Set lobjUtility = Nothing
'
'                    AppendBusinessError radrsBRV _
'                        , trInfEAPM _
'                        , mstrCBUSINESSINF _
'                        , "A new Portfolio " & CStr(llPortfolioNumber) & " was created." _
'                        , "dbProcessEAMRequest2"
'
'                    If (RecordsExist(vadrsPortfolioDetail)) Then
'                        vadrsPortfolioDetail.MoveFirst
'
'                        For llRecordIndex = 1 To vadrsPortfolioDetail.RecordCount
'                            If (vadrsPortfolioDetail!PRTFL_NUM = !PRTFL_NUM) Then
'                                vadrsPortfolioDetail!PRTFL_NUM = llPortfolioNumber
'                                vadrsPortfolioDetail.Update
'                            End If
'
'                            vadrsPortfolioDetail.MoveNext
'                        Next llRecordIndex
'
'                        vadrsPortfolioDetail.MoveFirst
'                    End If
'
'                    lsSQL = lsSQL & _
'                        " INSERT INTO METHDBA.TTR_ABS_PORTFOLIO" & _
'                        " (PRTFL_NUM, PRTFL_DESC" & _
'                        " , SHARE_IND" & _
'                        " , CREATED_EMP_NUM, CREATED_TIMESTAMP" & _
'                        " , LAST_MOD_EMP_NUM, LAST_MOD_TIMESTAMP)" & _
'                        " VALUES" & _
'                        " (" & llPortfolioNumber & ", '" & Trim(!PRTFL_DESC) & "'" & _
'                        " , '" & Trim(!SHARE_IND) & "'" & _
'                        " , " & CStr(vlTRISUserNumber) & ", CURRENT TIMESTAMP" & _
'                        " , " & CStr(vlTRISUserNumber) & ", CURRENT TIMESTAMP);"
'                End If
'            End If
'        End With
'    End If
'
'    If (RecordsExist(vadrsPortfolioDetail)) Then
'        With vadrsPortfolioDetail
'            .Filter = adFilterNone
'            .Filter = "RECORD_STATUS = 'I'"
'            If Not (.EOF) Then
'                For llRecordIndex = 1 To .RecordCount
'                    lsSQL = lsSQL & _
'                        " INSERT INTO METHDBA.TTR_ABS_PRTFL_DTL" & _
'                        " (PRTFL_NUM, ABS_NUM" & _
'                        " , CREATED_EMP_NUM, CREATED_TIMESTAMP" & _
'                        " , LAST_MOD_EMP_NUM, LAST_MOD_TIMESTAMP)" & _
'                        " VALUES" & _
'                        " (" & CStr(!PRTFL_NUM) & ", " & CStr(!ABS_NUM) & _
'                        " , " & CStr(vlTRISUserNumber) & ", CURRENT TIMESTAMP" & _
'                        " , " & CStr(vlTRISUserNumber) & ", CURRENT TIMESTAMP);"
'
'                    .MoveNext
'                Next llRecordIndex
'            End If
'        End With
'    End If
'
'EXITNORMALLY:
'    If (Trim(lsSQL) <> "") Then
'        'Get Connection
'        Set ladcnConnection = OpenConnection()
'        'Call Routine to Execute SQL
'        ExecuteOTHERSQL lsSQL, ladcnConnection
'    End If
'    'Close and destroy
'    If Not (ladcnConnection Is Nothing) Then
'        ladcnConnection.Close
'        Set ladcnConnection = Nothing
'    End If
'
'    Set lobjUtility = Nothing
'
'    dbProcessEAMRequest2 = True
'
'    GetObjectContext.SetComplete
'
'    Exit Function
'dbProcessEAMRequest2_EH:
'    GetObjectContext.SetAbort
'
'    Set lobjUtility = Nothing
'    Set ladcnConnection = Nothing
'
'    Set radrsBRV = Nothing
'
'    Select Case (Err.Number)
'            Case trErrutOpenConnection
'                TrsRaiseError MODULENAME _
'                    , "dbProcessEAMRequest2"
'            Case trErrdb_ExecuteSQL_F
'                TrsRaiseError MODULENAME _
'                    , "dbProcessEAMRequest2" _
'                    , trErrdb_ExecuteSQL_F _
'                    , , _
'                    , DescArgsFrom(lsSQL) _
'                    , True
'            Case Else
'                TrsRaiseError MODULENAME _
'                    , "dbProcessEAMRequest2" _
'                    , trErrdbSaveEAM _
'                    , , _
'                    , DescArgsFrom(lsSQL) _
'                    , True
'    End Select
'
'End Function
'End Prem TRIS 5.9
'    trErrdbGetEAM = 30338 + vbObjectError
'    trErrdbSaveEAM = 30339 + vbObjectError
'    trErrdbGetAbsencePortfolio = 30340 + vbObjectError
'    trErrdbSaveAbsencePortfolio = 30341 + vbObjectError
'    trErrdbGetAbsenceInitial = 30342 + vbObjectError
'    trErrdbGetAbsenceProvider = 30343 + vbObjectError
'    trErrdbGetAbsenceContact = 30344 + vbObjectError
'    trErrdbSaveAbsenceContact = 30345 + vbObjectError
'    trErrdbGetAbsenceReturnToWork = 30346 + vbObjectError
'    trErrdbSaveAbsenceReturnToWork = 30347 + vbObjectError
'    trErrdbGetAbsenceContactCallLog = 30348 + vbObjectError
'    trErrdbSaveAbsenceContactCallLog = 30349 + vbObjectError
'    trErrdbEAMEmployee = 30350 + vbObjectError
'    trErrdbGetAbsenceDates = 30351 + vbObjectError
'    trErrdbGetProviderCodeTable = 30352 + vbObjectError
