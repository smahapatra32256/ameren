VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "cdEmployeeTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const m_strCMODULENAME As String = "cdEmployeeTrans"
Private Const m_strC1SPACESTRING As String = " "
Public Function dbProcessQualsAndSkillsSave( _
                ByVal vadrsEmpQAndSInformation As ADODB.Recordset _
                , ByVal vlUserEmployeeNumber As Long _
                ) As Boolean
On Error GoTo dbProcessQualsAndSkillsSave_EH
'ADODB Connection Object
Dim ladcnConnection As ADODB.Connection
'Recordset
Dim llRecordIndex As Long
'SQL
Dim lsSQL As String

    dbProcessQualsAndSkillsSave = False

    'Check for existence!
    If Not (vadrsEmpQAndSInformation Is Nothing) Then
        With vadrsEmpQAndSInformation
            If Not (.EOF And .BOF) Then
                .Filter = adFilterNone

                'Check for deletes, if present process them first
                .Filter = MDAC25FilterFix("RECORD_STATUS = 'D'")
                If Not (.EOF) Then
                    .MoveFirst

                    For llRecordIndex = 1 To .RecordCount
                        lsSQL = lsSQL & _
                                " DELETE    FROM    METHDBA.TTR_EMP_QUAL_SKILL" & _
                                " WHERE     EMP_NUM = " & CStr(!KEY_EMP_NUM) & _
                                " AND       QUAL_CD = '" & Trim(!KEY_QUAL_CD) & "'" & _
                                " AND       EFF_START_DATE = '" & !KEY_EFF_START_DATE & "'" & _
                                " AND       TIME_STAMP = '" & Trim(!TIME_STAMP) & "'"
                        If (IsDate(!KEY_EFF_END_DATE)) Then
                            lsSQL = lsSQL & _
                                " AND       EFF_END_DATE = '" & !KEY_EFF_END_DATE & "'"
                        End If
                        If Not (IsNull(!KEY_SPEC_SKILL_CD) And _
                                Trim(!KEY_SPEC_SKILL_CD) = "") Then
                            lsSQL = lsSQL & _
                                " AND       SPEC_SKILL_CD = '" & !KEY_SPEC_SKILL_CD & "'"
                        End If

                        lsSQL = lsSQL & ";"

                        .MoveNext
                    Next llRecordIndex
                End If

                .Filter = adFilterNone
                'Check for inserts, if present process
                .Filter = MDAC25FilterFix("RECORD_STATUS = 'I'")
                If Not (.EOF) Then
                    .MoveFirst

                    For llRecordIndex = 1 To .RecordCount
                        lsSQL = lsSQL & _
                                " INSERT    INTO METHDBA.TTR_EMP_QUAL_SKILL" & _
                                " (EMP_NUM, QUAL_CD, SPEC_SKILL_CD" & _
                                " , EFF_START_DATE, EFF_END_DATE" & _
                                " , TIME_STAMP, LAST_MOD_EMP_NUM)" & _
                                " VALUES " & _
                                " (" & !EMP_NUM & ", '" & UCase(Trim(!QUAL_CD)) & "', '" & UCase(Trim(!SPEC_SKILL_CD)) & "'" & _
                                " , '" & !EFF_START_DATE & "', '" & !EFF_END_DATE & "'" & _
                                " , CURRENT TIMESTAMP, " & CStr(vlUserEmployeeNumber) & ");"

                        .MoveNext
                    Next llRecordIndex
                End If

                .Filter = adFilterNone
                'Check for changes
                .Filter = MDAC25FilterFix("RECORD_STATUS = 'C'")
                If Not (.EOF) Then
                    .MoveFirst

                    For llRecordIndex = 1 To .RecordCount
                        lsSQL = lsSQL & _
                                " UPDATE    METHDBA.TTR_EMP_QUAL_SKILL" & _
                                " SET       QUAL_CD = '" & UCase(Trim(!QUAL_CD)) & "'" & _
                                "           , SPEC_SKILL_CD = '" & UCase(Trim(!SPEC_SKILL_CD)) & "'" & _
                                "           , EFF_START_DATE = '" & Trim(!EFF_START_DATE) & "'" & _
                                "           , EFF_END_DATE = '" & Trim(!EFF_END_DATE) & "'" & _
                                "           , TIME_STAMP = CURRENT TIMESTAMP " & _
                                "           , LAST_MOD_EMP_NUM = " & CStr(vlUserEmployeeNumber) & _
                                " WHERE     EMP_NUM = " & (!KEY_EMP_NUM) & _
                                " AND       QUAL_CD = '" & Trim(!KEY_QUAL_CD) & "'" & _
                                " AND       EFF_START_DATE = '" & Trim(!KEY_EFF_START_DATE) & "'" & _
                                " AND       EFF_END_DATE = '" & Trim(!KEY_EFF_END_DATE) & "'" & _
                                " AND       TIME_STAMP = '" & Trim(!TIME_STAMP) & "'"

                        If Not (IsNull(!KEY_SPEC_SKILL_CD) And _
                                Trim(!KEY_SPEC_SKILL_CD) = "") Then
                            lsSQL = lsSQL & _
                                " AND       SPEC_SKILL_CD = '" & !KEY_SPEC_SKILL_CD & "'"
                        End If

                        lsSQL = lsSQL & ";"

                        .MoveNext
                    Next llRecordIndex
                End If

                .Filter = adFilterNone
            End If
        End With
    End If

    If (Trim(lsSQL) <> "") Then
        'Get Connection
        Set ladcnConnection = OpenConnection()
        'Call Routine to Execute SQL
        ExecuteOTHERSQL lsSQL, ladcnConnection
    End If

    'close and destroy
    If Not (ladcnConnection Is Nothing) Then
        ladcnConnection.Close
        Set ladcnConnection = Nothing
    End If

    dbProcessQualsAndSkillsSave = True

    'Complete Transaction
    GetObjectContext.SetComplete

    Exit Function

dbProcessQualsAndSkillsSave_EH:
    'Abort Transaction
    GetObjectContext.SetAbort

    Set ladcnConnection = Nothing

    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError m_strCMODULENAME _
                    , "dbProcessQualsAndSkillsSave"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError m_strCMODULENAME _
                    , "dbProcessQualsAndSkillsSave" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case trErrdbInsIntoEmpQualSkill, trErrdbUpdEmpQualSkill _
                , trErrdbDelFromEmpQualSkill
                TrsRaiseError m_strCMODULENAME _
                    , "dbProcessQualsAndSkillsSave" _
                    , Err.Number _
                    , , , _
                    , True
            Case Else
                TrsRaiseError m_strCMODULENAME _
                    , "dbProcessQualsAndSkillsSave" _
                    , trErrdbSaveEmpQualSkill _
                    , , , _
                    , True
    End Select

End Function
Public Function dbProcessSeniorityEntries(ByVal vadrsSeniorityInformation As ADODB.Recordset _
                    , ByVal vbIsAGencoEmployee As Boolean _
                    , ByVal vlSeniorityEmployeeNumber As Long _
                    , ByVal vlUserEmployeeNumber As Long _
                    ) As Boolean
On Error GoTo dbProcessSeniorityEntries_EH
'ADODB Connection Object
Dim ladcnConnection As ADODB.Connection
'Recordset
Dim llRecordIndex As Long
'SQL
Dim lsSQL As String

Dim llCorpQty As Long, llCmpyQty As Long
Dim llPlantQty As Long, llUnionQty As Long
Dim llDeptQty As Long, llJobQty As Long
Dim llLOPQty As Long, lsValidDB2CorpDate As String
'Oh! IBM Vs. Microsoft
Dim lsValidDB2CmpyDate As String, lsValidDB2PlantDate As String
Dim lsValidDB2UnionDate As String, lsValidDB2DeptDate As String
Dim lsValidDB2JobDate As String, lsValidDB2LOPDate As String

    dbProcessSeniorityEntries = False

    'Check for existence!
    If Not (vadrsSeniorityInformation Is Nothing) Then
        With vadrsSeniorityInformation
            If Not (.EOF And .BOF) Then
                .Filter = adFilterNone

                'Check for deletes, if present process them first
                .Filter = MDAC25FilterFix("RECORD_STATUS = 'D'")
                If Not (.EOF) Then
                    .MoveFirst

                    For llRecordIndex = 1 To .RecordCount
                        If (vbIsAGencoEmployee) Then
                            lsSQL = lsSQL & _
                                    " DELETE FROM  METHDBA.TTR_CIPS_SNRTY_DT" & _
                                    " WHERE EMP_NUM = " & vlSeniorityEmployeeNumber & _
                                    " AND   UNION_CD = '" & Trim(!KEY_UNION_CD) & "'" & _
                                    " AND   DEPT_NUM = '" & Trim(!KEY_DEPT_NUM) & "'" & _
                                    " AND   JOB_NUM = '" & Trim(!KEY_JOB_NUM) & "';"
                        End If

                        '//-----for auditing
                        lsSQL = lsSQL & _
                                " UPDATE METHDBA.TTR_SENIORITY" & _
                                " SET LAST_MOD_EMP_NUM = " & CStr(vlUserEmployeeNumber) & _
                                " WHERE EMP_NUM = " & vlSeniorityEmployeeNumber & _
                                " AND   UNION_CD = '" & Trim(!KEY_UNION_CD) & "'" & _
                                " AND   DEPT_NUM = '" & Trim(!KEY_DEPT_NUM) & "'" & _
                                " AND   JOB_NUM = '" & Trim(!KEY_JOB_NUM) & "'" & _
                                " AND   TIME_STAMP = '" & Trim(!TIME_STAMP) & "';"
                                
                        lsSQL = lsSQL & _
                                " DELETE FROM   METHDBA.TTR_SENIORITY" & _
                                " WHERE EMP_NUM = " & vlSeniorityEmployeeNumber & _
                                " AND   UNION_CD = '" & Trim(!KEY_UNION_CD) & "'" & _
                                " AND   DEPT_NUM = '" & Trim(!KEY_DEPT_NUM) & "'" & _
                                " AND   JOB_NUM = '" & Trim(!KEY_JOB_NUM) & "'" & _
                                " AND   TIME_STAMP = '" & Trim(!TIME_STAMP) & "';"

                        .MoveNext
                    Next llRecordIndex
                End If

                'Check for inserts
                .Filter = adFilterNone
                .Filter = MDAC25FilterFix("RECORD_STATUS = 'I'")
                If Not (.EOF) Then
                    .MoveFirst

                    For llRecordIndex = 1 To .RecordCount
                        'Calculate the Quantity Field Value!!!!!!
                        llCorpQty = (!CORP_MONTHS * 100) + !CORP_DAYS
                        llCmpyQty = (!CMPY_MONTHS * 100) + !CMPY_DAYS
                        llPlantQty = (!PLANT_MONTHS * 100) + !PLANT_DAYS
                        llUnionQty = (!UNION_MONTHS * 100) + !UNION_DAYS
                        llDeptQty = (!DEPT_MONTHS * 100) + !DEPT_DAYS
                        llJobQty = (!JOB_MONTHS * 100) + !JOB_DAYS
                        llLOPQty = (!LOP_MONTHS * 100) + !LOP_DAYS

                        lsValidDB2CorpDate = GetValidDB2Date(!CORP_SNRTY_UP_DT)
                        lsValidDB2CmpyDate = GetValidDB2Date(!CMPY_SNRTY_UP_DT)
                        lsValidDB2PlantDate = GetValidDB2Date(!PLANT_SNRTY_UP_DT)
                        lsValidDB2UnionDate = GetValidDB2Date(!UNION_SNRTY_UP_DT)
                        lsValidDB2DeptDate = GetValidDB2Date(!DEPT_SNRTY_UP_DT)
                        lsValidDB2JobDate = GetValidDB2Date(!JOB_SNRTY_UP_DT)
                        lsValidDB2LOPDate = GetValidDB2Date(!LOP_SNRTY_UP_DT)

                        lsSQL = lsSQL & _
                            " INSERT    INTO METHDBA.TTR_SENIORITY" & _
                            " (EMP_NUM, UNION_CD, DEPT_NUM, JOB_NUM" & _
                            " , SNRTY_PRTY_CD, PPM_FL" & _
                            " , LTD_IND, MLTRY_LOA_IND, SPCL_ATTN_FL" & _
                            " , CORP_SNRTY_QTY, CORP_SNRTY_UP_DT" & _
                            " , CMPY_SNRTY_QTY, CMPY_SNRTY_UP_DT" & _
                            " , PLANT_SNRTY_QTY, PLANT_SNRTY_UP_DT" & _
                            " , UNION_SNRTY_QTY, UNION_SNRTY_UP_DT" & _
                            " , DEPT_SNRTY_QTY, DEPT_SNRTY_UP_DT" & _
                            " , JOB_SNRTY_QTY, JOB_SNRTY_UP_DT" & _
                            " , LOP_SNRTY_QTY, LOP_SNRTY_UP_DT" & _
                            " , LAST_MOD_EMP_NUM, TIME_STAMP) VALUES "
                        lsSQL = lsSQL & _
                            " (" & CStr(vlSeniorityEmployeeNumber) & ", '" & Trim(!UNION_CD) & "'" & _
                            " , '" & Trim(!DEPT_NUM) & "', '" & Trim(!JOB_NUM) & "'" & _
                            " , '" & Trim(!SNRTY_PRTY_CD) & "', '" & Trim(!PPM_FL) & "'" & _
                            " , '" & Trim(!LTD_IND) & "', '" & Trim(!MLTRY_LOA_IND) & "'" & _
                            " , '" & Trim(!SPCL_ATTN_FL) & _
                            "', " & CStr(llCorpQty) & " , '" & lsValidDB2CorpDate & _
                            "', " & CStr(llCmpyQty) & " , '" & lsValidDB2CmpyDate & _
                            "', " & CStr(llPlantQty) & " , '" & lsValidDB2PlantDate & _
                            "', " & CStr(llUnionQty) & " , '" & lsValidDB2UnionDate & _
                            "', " & CStr(llDeptQty) & " , '" & lsValidDB2DeptDate & _
                            "', " & CStr(llJobQty) & " , '" & lsValidDB2JobDate & _
                            "', " & CStr(llLOPQty) & " , '" & lsValidDB2LOPDate & _
                            "', " & CStr(vlUserEmployeeNumber) & " , CURRENT TIMESTAMP);"

                        If (vbIsAGencoEmployee) Then
                            lsValidDB2CorpDate = GetValidDB2Date(!CORP_SNRTY_EFF_DT)
                            lsValidDB2CmpyDate = GetValidDB2Date(!CMPY_SNRTY_EFF_DT)
                            lsValidDB2PlantDate = GetValidDB2Date(!PLANT_SNRTY_EFF_DT)
                            lsValidDB2DeptDate = GetValidDB2Date(!DEPT_SNRTY_EFF_DT)
                            lsValidDB2JobDate = GetValidDB2Date(!JOB_SNRTY_EFF_DT)

                            lsSQL = lsSQL & _
                                    "INSERT    INTO METHDBA.TTR_CIPS_SNRTY_DT" & _
                                    " (EMP_NUM, UNION_CD" & _
                                    " , DEPT_NUM, JOB_NUM" & _
                                    " , CORP_SNRTY_EFF_DT, CMPY_SNRTY_EFF_DT" & _
                                    " , PLANT_SNRTY_EFF_DT, DEPT_SNRTY_EFF_DT" & _
                                    " , JOB_SNRTY_EFF_DT)" & _
                                    " VALUES " & _
                                    " (" & CStr(vlSeniorityEmployeeNumber) & ", '" & Trim(!UNION_CD) & "'" & _
                                    " , '" & Trim(!DEPT_NUM) & "', '" & Trim(!JOB_NUM) & "'" & _
                                    " , '" & lsValidDB2CorpDate & "', '" & lsValidDB2CmpyDate & "'" & _
                                    " , '" & lsValidDB2PlantDate & "', '" & lsValidDB2DeptDate & "'" & _
                                    " , '" & lsValidDB2JobDate & "');"
                        End If

                        .MoveNext
                    Next llRecordIndex
                End If

                'Check for changes
                .Filter = adFilterNone
                .Filter = MDAC25FilterFix("RECORD_STATUS = 'C'")
                If Not (.EOF) Then
                    .MoveFirst

                    For llRecordIndex = 1 To .RecordCount
                        'Calculate the Quantity Field Value!!!!!!
                        llCorpQty = (!CORP_MONTHS * 100) + !CORP_DAYS
                        llCmpyQty = (!CMPY_MONTHS * 100) + !CMPY_DAYS
                        llPlantQty = (!PLANT_MONTHS * 100) + !PLANT_DAYS
                        llUnionQty = (!UNION_MONTHS * 100) + !UNION_DAYS
                        llDeptQty = (!DEPT_MONTHS * 100) + !DEPT_DAYS
                        llJobQty = (!JOB_MONTHS * 100) + !JOB_DAYS
                        llLOPQty = (!LOP_MONTHS * 100) + !LOP_DAYS

                        lsValidDB2CorpDate = GetValidDB2Date(!CORP_SNRTY_UP_DT)
                        lsValidDB2CmpyDate = GetValidDB2Date(!CMPY_SNRTY_UP_DT)
                        lsValidDB2PlantDate = GetValidDB2Date(!PLANT_SNRTY_UP_DT)
                        lsValidDB2UnionDate = GetValidDB2Date(!UNION_SNRTY_UP_DT)
                        lsValidDB2DeptDate = GetValidDB2Date(!DEPT_SNRTY_UP_DT)
                        lsValidDB2JobDate = GetValidDB2Date(!JOB_SNRTY_UP_DT)
                        lsValidDB2LOPDate = GetValidDB2Date(!LOP_SNRTY_UP_DT)

                        lsSQL = lsSQL & _
                            " UPDATE    METHDBA.TTR_SENIORITY" & _
                            " SET   SNRTY_PRTY_CD = '" & Trim(!SNRTY_PRTY_CD) & "'" & _
                            "       , PPM_FL = '" & Trim(!PPM_FL) & "'" & _
                            "       , LTD_IND = '" & Trim(!LTD_IND) & "'" & _
                            "       , MLTRY_LOA_IND = '" & Trim(!MLTRY_LOA_IND) & "'" & _
                            "       , SPCL_ATTN_FL = '" & Trim(!SPCL_ATTN_FL) & "'" & _
                            "       , CORP_SNRTY_QTY = " & CStr(llCorpQty) & _
                            "       , CORP_SNRTY_UP_DT = '" & lsValidDB2CorpDate & "'" & _
                            "       , CMPY_SNRTY_QTY = " & CStr(llCmpyQty) & _
                            "       , CMPY_SNRTY_UP_DT = '" & lsValidDB2CmpyDate & "'" & _
                            "       , UNION_SNRTY_QTY = " & CStr(llUnionQty) & _
                            "       , UNION_SNRTY_UP_DT = '" & lsValidDB2UnionDate & "'" & _
                            "       , PLANT_SNRTY_QTY = " & CStr(llPlantQty) & _
                            "       , PLANT_SNRTY_UP_DT = '" & lsValidDB2PlantDate & "'" & _
                            "       , DEPT_SNRTY_QTY = " & CStr(llDeptQty) & _
                            "       , DEPT_SNRTY_UP_DT = '" & lsValidDB2DeptDate & "'" & _
                            "       , JOB_SNRTY_QTY = " & CStr(llJobQty) & _
                            "       , JOB_SNRTY_UP_DT = '" & lsValidDB2JobDate & "'" & _
                            "       , LOP_SNRTY_QTY = " & CStr(llLOPQty) & _
                            "       , LOP_SNRTY_UP_DT = '" & lsValidDB2LOPDate & "'" & _
                            "       , LAST_MOD_EMP_NUM = " & CStr(vlUserEmployeeNumber) & _
                            "       , TIME_STAMP = CURRENT TIMESTAMP"
                            'TOO MANY LINES OF CONTINUATION!!!
                        lsSQL = lsSQL & _
                            " WHERE EMP_NUM = " & CStr(vlSeniorityEmployeeNumber) & _
                            " AND   UNION_CD = '" & Trim(!KEY_UNION_CD) & "'" & _
                            " AND   DEPT_NUM = '" & Trim(!KEY_DEPT_NUM) & "'" & _
                            " AND   JOB_NUM = '" & Trim(!KEY_JOB_NUM) & "'" & _
                            " AND   TIME_STAMP = '" & Trim(!TIME_STAMP) & "';"

                        If (vbIsAGencoEmployee) Then
                            lsValidDB2CorpDate = GetValidDB2Date(!CORP_SNRTY_EFF_DT)
                            lsValidDB2CmpyDate = GetValidDB2Date(!CMPY_SNRTY_EFF_DT)
                            lsValidDB2PlantDate = GetValidDB2Date(!PLANT_SNRTY_EFF_DT)
                            lsValidDB2DeptDate = GetValidDB2Date(!DEPT_SNRTY_EFF_DT)
                            lsValidDB2JobDate = GetValidDB2Date(!JOB_SNRTY_EFF_DT)

                            lsSQL = lsSQL & _
                                " UPDATE    METHDBA.TTR_CIPS_SNRTY_DT" & _
                                " SET   CORP_SNRTY_EFF_DT = '" & lsValidDB2CorpDate & "'" & _
                                "       , CMPY_SNRTY_EFF_DT = '" & lsValidDB2CmpyDate & "'" & _
                                "       , PLANT_SNRTY_EFF_DT = '" & lsValidDB2PlantDate & "'" & _
                                "       , DEPT_SNRTY_EFF_DT = '" & lsValidDB2DeptDate & "'" & _
                                "       , JOB_SNRTY_EFF_DT = '" & lsValidDB2JobDate & "'" & _
                                " WHERE EMP_NUM = " & CStr(vlSeniorityEmployeeNumber) & _
                                " AND   UNION_CD = '" & Trim(!KEY_UNION_CD) & "'" & _
                                " AND   DEPT_NUM = '" & Trim(!KEY_DEPT_NUM) & "'" & _
                                " AND   JOB_NUM = '" & Trim(!KEY_JOB_NUM) & "';"
                        End If

                        .MoveNext
                    Next llRecordIndex
                End If

                .Filter = adFilterNone
            End If
        End With
    End If

    If (Trim(lsSQL) <> "") Then
        'Get Connection
        Set ladcnConnection = OpenConnection()
        'Execute
        ExecuteOTHERSQL lsSQL, ladcnConnection
    End If

    'close and destroy
    If Not (ladcnConnection Is Nothing) Then
        ladcnConnection.Close
        Set ladcnConnection = Nothing
    End If

    dbProcessSeniorityEntries = True

    'Complete Transaction
    GetObjectContext.SetComplete

    Exit Function
dbProcessSeniorityEntries_EH:
    'Abort Transaction
    GetObjectContext.SetAbort

    Set ladcnConnection = Nothing

    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError m_strCMODULENAME _
                    , "dbProcessSeniorityEntries"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError m_strCMODULENAME _
                    , "dbProcessSeniorityEntries" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case trErrdbInsIntoEmpSnrty, trErrdbUpdEmpSnrty _
                , trErrdbDelFromEmpSnrty
                TrsRaiseError m_strCMODULENAME _
                    , "dbProcessSeniorityEntries" _
                    , Err.Number _
                    , , , _
                    , True
            Case Else
                TrsRaiseError m_strCMODULENAME _
                    , "dbProcessSeniorityEntries" _
                    , trErrdbSaveEmpEmpSnrty _
                    , , , _
                    , True
    End Select

End Function
'Oh! IBM Vs. Microsoft
Private Function GetValidDB2Date(ByVal vdtADODBDate As Date _
                    ) As String
On Error GoTo GetValidDB2Date_EH

    If (vdtADODBDate = #1/1/100#) Then
        GetValidDB2Date = "01/01/0001"
    Else
        GetValidDB2Date = Format(vdtADODBDate, "mm/dd/yyyy")
    End If

    Exit Function
GetValidDB2Date_EH:
    TrsRaiseError m_strCMODULENAME _
        , "GetValidDB2Date" _
        , trErrdbSaveEmpEmpSnrty _
        , , , _
        , True

End Function
Public Function dbProcessFrozenSeniorityEntries(ByVal vadrsFrozenSeniorityInformation As ADODB.Recordset _
                    , ByVal vlSeniorityEmployeeNumber As Long _
                    , ByVal vlUserEmployeeNumber As Long _
                    ) As Boolean
On Error GoTo dbProcessFrozenSeniorityEntries_EH
'ADODB Connection Object
Dim ladcnConnection As ADODB.Connection
'Recordset
Dim llRecordIndex As Long
'SQL
Dim lsSQL As String, llQty As Long

    dbProcessFrozenSeniorityEntries = False

    'Check for existence!
    If Not (vadrsFrozenSeniorityInformation Is Nothing) Then
        With vadrsFrozenSeniorityInformation
            If Not (.EOF And .BOF) Then
                .Filter = adFilterNone

                'Check for deletes, if present process them first
                .Filter = MDAC25FilterFix("RECORD_STATUS = 'D'")
                If Not (.EOF) Then
                    .MoveFirst

                    For llRecordIndex = 1 To .RecordCount
                        lsSQL = lsSQL & _
                                " DELETE FROM  METHDBA.TTR_FRZN_SNRTY " & _
                                " WHERE EMP_NUM = " & vlSeniorityEmployeeNumber & _
                                " AND   SNRTY_REC_TYPE_CD = '" & Trim(!KEY_SNRTY_REC_TYPE_CD) & "'" & _
                                " AND   TIME_STAMP = '" & Trim(!TIME_STAMP) & "';"

                        .MoveNext
                    Next llRecordIndex
                End If

                'Check for inserts
                .Filter = adFilterNone
                .Filter = MDAC25FilterFix("RECORD_STATUS = 'I'")
                If Not (.EOF) Then
                    .MoveFirst

                    For llRecordIndex = 1 To .RecordCount
                        'Calculate the Quantity Field Value!!!!!!
                        llQty = (!FRZN_SNRTY_MONTHS * 100) + !FRZN_SNRTY_DAYS

                        lsSQL = lsSQL & _
                            " INSERT    INTO METHDBA.TTR_FRZN_SNRTY" & _
                            " (EMP_NUM, SNRTY_REC_TYPE_CD" & _
                            " , CMPY_CD, UNION_CD" & _
                            " , DEPT_NUM, JOB_NUM" & _
                            " , IR_AREA_CD, IR_DIV_CD" & _
                            " , SNRTY_FRZN_DT, FRZN_SNRTY_REC_QTY" & _
                            " , LAST_MOD_EMP_NUM, TIME_STAMP) VALUES " & _
                            " (" & CStr(vlSeniorityEmployeeNumber) & ", '" & Trim(!SNRTY_REC_TYPE_CD) & "'" & _
                            " , '" & Trim(!CMPY_CD) & "', '" & Trim(!UNION_CD) & "'" & _
                            " , '" & Trim(!DEPT_NUM) & "', '" & Trim(!JOB_NUM) & "'" & _
                            " , '" & Trim(!IR_AREA_CD) & "', '" & Trim(!IR_DIV_CD) & "'" & _
                            " , '" & Trim(!SNRTY_FRZN_DT) & "', " & CStr(llQty) & _
                            " , " & CStr(vlUserEmployeeNumber) & " , CURRENT TIMESTAMP);"

                        .MoveNext
                    Next llRecordIndex
                End If

                'Check for inserts, if present process them first
                .Filter = adFilterNone
                .Filter = MDAC25FilterFix("RECORD_STATUS = 'C'")
                If Not (.EOF) Then
                    .MoveFirst

                    For llRecordIndex = 1 To .RecordCount
                        'Calculate the Quantity Field Value!!!!!!
                        llQty = (!FRZN_SNRTY_MONTHS * 100) + !FRZN_SNRTY_DAYS

                        lsSQL = lsSQL & _
                            " UPDATE    METHDBA.TTR_FRZN_SNRTY" & _
                            " SET   CMPY_CD = '" & Trim(!CMPY_CD) & "'" & _
                            "       , UNION_CD = '" & Trim(!UNION_CD) & "'" & _
                            "       , DEPT_NUM = '" & Trim(!DEPT_NUM) & "'" & _
                            "       , JOB_NUM = '" & Trim(!JOB_NUM) & "'" & _
                            "       , IR_AREA_CD = '" & Trim(!IR_AREA_CD) & "'" & _
                            "       , IR_DIV_CD = '" & Trim(!IR_DIV_CD) & "'" & _
                            "       , SNRTY_FRZN_DT = '" & Trim(!SNRTY_FRZN_DT) & "'" & _
                            "       , FRZN_SNRTY_REC_QTY = " & CStr(llQty) & _
                            "       , LAST_MOD_EMP_NUM = " & CStr(vlUserEmployeeNumber) & _
                            "       , TIME_STAMP = CURRENT TIMESTAMP " & _
                            " WHERE EMP_NUM = " & CStr(vlSeniorityEmployeeNumber) & _
                            " AND   SNRTY_REC_TYPE_CD = '" & Trim(!KEY_SNRTY_REC_TYPE_CD) & "'" & _
                            " AND   TIME_STAMP = '" & Trim(!TIME_STAMP) & "';"

                        .MoveNext
                    Next llRecordIndex
                End If

                .Filter = adFilterNone
            End If
        End With
    End If

    If (Trim(lsSQL) <> "") Then
        'Get Connection
        Set ladcnConnection = OpenConnection()
        'Call Routine to Execute SQL
        ExecuteOTHERSQL lsSQL, ladcnConnection
    End If

    'close and destroy
    If Not (ladcnConnection Is Nothing) Then
        ladcnConnection.Close
        Set ladcnConnection = Nothing
    End If

    dbProcessFrozenSeniorityEntries = True

    'Complete Transaction
    GetObjectContext.SetComplete

    Exit Function
dbProcessFrozenSeniorityEntries_EH:
    'Abort Transaction
    GetObjectContext.SetAbort

    Set ladcnConnection = Nothing

    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError m_strCMODULENAME _
                    , "dbProcessFrozenSeniorityEntries"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError m_strCMODULENAME _
                    , "dbProcessFrozenSeniorityEntries" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case trErrdbInsIntoEmpFrznSnrty, trErrdbUpdEmpFrznSnrty _
                , trErrdbDelFromEmpFrznSnrty
                TrsRaiseError m_strCMODULENAME _
                    , "dbProcessFrozenSeniorityEntries" _
                    , Err.Number _
                    , , , _
                    , True
            Case Else
                TrsRaiseError m_strCMODULENAME _
                    , "dbProcessFrozenSeniorityEntries" _
                    , trErrdbSaveEmpFrznSnrty _
                    , , , _
                    , True
    End Select

End Function
Public Function dbProcessCanvassJobEntries(ByVal vadrsCanvassJobInformation As ADODB.Recordset _
                    , ByVal vlUserEmployeeNumber As Long _
                    ) As Boolean
On Error GoTo dbProcessCanvassJobEntries_EH
'ADODB Connection Object
Dim ladcnConnection As ADODB.Connection
'Recordset
Dim llRecordIndex As Long
'SQL
Dim lsSQL As String, llQty As Long

    dbProcessCanvassJobEntries = False

    'Check for existence!
    If Not (vadrsCanvassJobInformation Is Nothing) Then
        With vadrsCanvassJobInformation
            If Not (.EOF And .BOF) Then
                .Filter = adFilterNone

                'Check for deletes, if present process them first
                .Filter = MDAC25FilterFix("RECORD_STATUS = 'D'")
                If Not (.EOF) Then
                    .MoveFirst

                    For llRecordIndex = 1 To .RecordCount
                        lsSQL = lsSQL & _
                            " DELETE FROM  METHDBA.TTR_OT_EMP_OVRD " & _
                            " WHERE EMP_NUM = " & CStr(!KEY_EMP_NUM) & _
                            " AND   JOB_NUM = '" & Trim(!KEY_JOB_NUM) & "'" & _
                            " AND   EFF_START_DATE = '" & Trim(!KEY_EFF_START_DATE) & "'" & _
                            " AND   CREATED_TIMESTAMP = '" & Trim(!CREATED_TIMESTAMP) & "';"

                        .MoveNext
                    Next llRecordIndex
                End If

                'Check for inserts
                .Filter = adFilterNone
                .Filter = MDAC25FilterFix("RECORD_STATUS = 'I'")
                If Not (.EOF) Then
                    .MoveFirst

                    For llRecordIndex = 1 To .RecordCount
                        lsSQL = lsSQL & _
                            " INSERT    INTO METHDBA.TTR_OT_EMP_OVRD" & _
                            " (EMP_NUM, JOB_NUM, CRAFT_NUM" & _
                            " , EFF_START_DATE, EFF_END_DATE" & _
                            " , CREATED_EMP_NUM, CREATED_TIMESTAMP" & _
                            " , LAST_MOD_EMP_NUM , LAST_MOD_TIMESTAMP)" & _
                            " VALUES " & _
                            " (" & CStr(!EMP_NUM) & ", '" & Trim(!JOB_NUM) & "'" & _
                            " , '" & Trim(!CRAFT_NUM) & "'" & _
                            " , '" & Trim(!EFF_START_DATE) & "', '" & Trim(!EFF_END_DATE) & "'" & _
                            " , " & CStr(vlUserEmployeeNumber) & ", CURRENT TIMESTAMP " & _
                            " , " & CStr(vlUserEmployeeNumber) & ", CURRENT TIMESTAMP);"

                        .MoveNext
                    Next llRecordIndex
                End If

                'Check for changes
                .Filter = adFilterNone
                .Filter = MDAC25FilterFix("RECORD_STATUS = 'C'")
                If Not (.EOF) Then
                    .MoveFirst

                    For llRecordIndex = 1 To .RecordCount
                        lsSQL = lsSQL & _
                            " UPDATE    METHDBA.TTR_OT_EMP_OVRD" & _
                            " SET   JOB_NUM = '" & Trim(!JOB_NUM) & "'" & _
                            "       , CRAFT_NUM = '" & Trim(!CRAFT_NUM) & "'" & _
                            "       , EFF_END_DATE = '" & Trim(!EFF_END_DATE) & "'" & _
                            "       , LAST_MOD_EMP_NUM = " & CStr(vlUserEmployeeNumber) & _
                            "       , LAST_MOD_TIMESTAMP = CURRENT TIMESTAMP " & _
                            " WHERE EMP_NUM = " & CStr(!KEY_EMP_NUM) & _
                            " AND   JOB_NUM = '" & Trim(!KEY_JOB_NUM) & "'" & _
                            " AND   EFF_START_DATE = '" & Trim(!KEY_EFF_START_DATE) & "'" & _
                            " AND   CREATED_TIMESTAMP = '" & Trim(!CREATED_TIMESTAMP) & "';"

                        .MoveNext
                    Next llRecordIndex
                End If

                .Filter = adFilterNone
            End If
        End With
    End If

    If (Trim(lsSQL) <> "") Then
        'Get Connection
        Set ladcnConnection = OpenConnection()
        'Call Routine to Execute SQL
        ExecuteOTHERSQL lsSQL, ladcnConnection
    End If

    'close and destroy
    If Not (ladcnConnection Is Nothing) Then
        ladcnConnection.Close
        Set ladcnConnection = Nothing
    End If

    dbProcessCanvassJobEntries = True

    'Complete Transaction
    GetObjectContext.SetComplete

    Exit Function

dbProcessCanvassJobEntries_EH:
    'Abort Transaction
    GetObjectContext.SetAbort

    Set ladcnConnection = Nothing
    dbProcessCanvassJobEntries = False

    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError m_strCMODULENAME _
                    , "dbProcessCanvassJobEntries"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError m_strCMODULENAME _
                    , "dbProcessCanvassJobEntries" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case trErrdbInsIntoEmpCanvJob, trErrdbUpdEmpCanvJob _
                , trErrdbDelFromEmpCanvJob
                TrsRaiseError m_strCMODULENAME _
                    , "dbProcessCanvassJobEntries" _
                    , Err.Number _
                    , , , _
                    , True
            Case Else
                TrsRaiseError m_strCMODULENAME _
                    , "dbProcessCanvassJobEntries" _
                    , trErrdbSaveEmpCanvJob _
                    , , , _
                    , True
    End Select

End Function
Public Function dbProcessIneligibilityEntries(ByVal vadrsIneligibilityInformation As ADODB.Recordset _
                    , ByVal vlUserEmployeeNumber As Long _
                    ) As Boolean
On Error GoTo dbProcessIneligibilityEntries_EH
'ADODB Connection Object
Dim ladcnConnection As ADODB.Connection
'Recordset
Dim llRecordIndex As Long
'SQL
Dim lsSQL As String

    dbProcessIneligibilityEntries = False

    'Check for existence!
    If Not (vadrsIneligibilityInformation Is Nothing) Then
        With vadrsIneligibilityInformation
            If Not (.EOF And .BOF) Then
                .Filter = adFilterNone

                'Check for deletes, if present process them first
                .Filter = MDAC25FilterFix("RECORD_STATUS = 'D'")
                If Not (.EOF) Then
                    .MoveFirst

                    For llRecordIndex = 1 To .RecordCount
                        lsSQL = lsSQL & _
                            " DELETE FROM  METHDBA.TTR_INELIGIBILITY" & _
                            " WHERE EMP_NUM = " & CStr(!KEY_EMP_NUM) & _
                            " AND   EXCEPTION_CD = '" & Trim(!KEY_EXCEPTION_CD) & "'" & _
                            " AND   EFF_START_DATE = '" & Trim(!KEY_EFF_START_DATE) & "'" & _
                            " AND   EFF_START_TIME = '" & Trim(!KEY_EFF_START_TIME) & "'" & _
                            " AND   TIME_STAMP = '" & Trim(!TIME_STAMP) & "';"

                        .MoveNext
                    Next llRecordIndex
                End If

                'Check for inserts
                .Filter = adFilterNone
                .Filter = MDAC25FilterFix("RECORD_STATUS = 'I'")
                If Not (.EOF) Then
                    .MoveFirst

                    For llRecordIndex = 1 To .RecordCount
                        lsSQL = lsSQL & _
                            " INSERT    INTO METHDBA.TTR_INELIGIBILITY" & _
                            " (EMP_NUM, EXCEPTION_CD" & _
                            " , EFF_START_DATE, EFF_START_TIME" & _
                            " , EFF_END_DATE, EFF_END_TIME" & _
                            " , LAST_MOD_EMP_NUM, TIME_STAMP)" & _
                            " VALUES " & _
                            " (" & CStr(!EMP_NUM) & ", '" & Trim(!EXCEPTION_CD) & "'" & _
                            " , '" & Trim(!EFF_START_DATE) & "', '" & Trim(!EFF_START_TIME) & "'" & _
                            " , '" & Trim(!EFF_END_DATE) & "', '" & Trim(!EFF_END_TIME) & "'" & _
                            " , " & CStr(vlUserEmployeeNumber) & ", CURRENT TIMESTAMP);"

                        .MoveNext
                    Next llRecordIndex
                End If

                'Check for inserts, if present process them first
                .Filter = adFilterNone
                .Filter = MDAC25FilterFix("RECORD_STATUS = 'C'")
                If Not (.EOF) Then
                    .MoveFirst

                    For llRecordIndex = 1 To .RecordCount
                        lsSQL = lsSQL & _
                            " UPDATE    METHDBA.TTR_INELIGIBILITY" & _
                            " SET   EFF_END_DATE = '" & Trim(!EFF_END_DATE) & "'" & _
                            "       , EFF_END_TIME = '" & Trim(!EFF_END_TIME) & "'" & _
                            "       , LAST_MOD_EMP_NUM = " & CStr(vlUserEmployeeNumber) & _
                            "       , TIME_STAMP = CURRENT TIMESTAMP " & _
                            " WHERE EMP_NUM = " & CStr(!KEY_EMP_NUM) & _
                            " AND   EFF_START_DATE = '" & Trim(!KEY_EFF_START_DATE) & "'" & _
                            " AND   EXCEPTION_CD = '" & Trim(!KEY_EXCEPTION_CD) & "'" & _
                            " AND   EFF_START_TIME = '" & Trim(!KEY_EFF_START_TIME) & "'" & _
                            " AND   TIME_STAMP = '" & Trim(!TIME_STAMP) & "';"

                        .MoveNext
                    Next llRecordIndex
                End If

                .Filter = adFilterNone
            End If
        End With
    End If

    If (Trim(lsSQL) <> "") Then
        'Get Connection
        Set ladcnConnection = OpenConnection()
        'Call Routine to Execute SQL
        ExecuteOTHERSQL lsSQL, ladcnConnection
    End If

    'close and destroy
    If Not (ladcnConnection Is Nothing) Then
        ladcnConnection.Close
        Set ladcnConnection = Nothing
    End If

    dbProcessIneligibilityEntries = True

    'Complete Transaction
    GetObjectContext.SetComplete

    Exit Function
dbProcessIneligibilityEntries_EH:
    'Abort Transaction
    GetObjectContext.SetAbort

    Set ladcnConnection = Nothing

    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError m_strCMODULENAME _
                    , "dbProcessIneligibilityEntries"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError m_strCMODULENAME _
                    , "dbProcessIneligibilityEntries" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case trErrdbInsIntoEmpIneligibility, trErrdbUpdEmpIneligibility _
                , trErrdbDelFromEmpIneligibility
                TrsRaiseError m_strCMODULENAME _
                    , "dbProcessIneligibilityEntries" _
                    , Err.Number _
                    , , , _
                    , True
            Case Else
                TrsRaiseError m_strCMODULENAME _
                    , "dbProcessIneligibilityEntries" _
                    , trErrdbSaveEmpIneligibility _
                    , , , _
                    , True
    End Select

End Function
Public Function dbProcessVolunteerEntries(ByVal vadrsVolunteerInformation As ADODB.Recordset _
                    , ByVal vlUserEmployeeNumber As Long _
                    ) As Boolean
On Error GoTo dbProcessVolunteerEntries_EH
'ADODB Connection Object
Dim ladcnConnection As ADODB.Connection
'Recordset
Dim llRecordIndex As Long
'SQL
Dim lsSQL As String

    dbProcessVolunteerEntries = False

    'Check for existence!
    If Not (vadrsVolunteerInformation Is Nothing) Then
        With vadrsVolunteerInformation
            If Not (.EOF And .BOF) Then
                .Filter = adFilterNone

                If Not (.EOF) Then
                    'Delete All
                    lsSQL = lsSQL & _
                        " DELETE FROM  METHDBA.TTR_VOLUNTEER;"

                    'Loop through and insert
                    .MoveFirst
                    For llRecordIndex = 1 To .RecordCount
                        lsSQL = lsSQL & _
                            " INSERT    INTO METHDBA.TTR_VOLUNTEER" & _
                            " (EMP_NUM, STUS_EFF_DT" & _
                            " , ACTN_CD, TIME_STAMP)" & _
                            " VALUES " & _
                            " (" & CStr(!EMP_NUM) & ", CURRENT DATE" & _
                            " , '" & Trim(!VLNTR_STATUS) & "', CURRENT TIMESTAMP);"

                        .MoveNext
                    Next llRecordIndex
                End If

                .Filter = adFilterNone
            End If
        End With
    End If

    If (Trim(lsSQL) <> "") Then
        'Get Connection
        Set ladcnConnection = OpenConnection()
        'Call Routine to Execute SQL
        ExecuteOTHERSQL lsSQL, ladcnConnection
    End If

    'close and destroy
    If Not (ladcnConnection Is Nothing) Then
        ladcnConnection.Close
        Set ladcnConnection = Nothing
    End If

    dbProcessVolunteerEntries = True

    'Complete Transaction
    GetObjectContext.SetComplete

    Exit Function
dbProcessVolunteerEntries_EH:
    'Abort Transaction
    GetObjectContext.SetAbort

    Set ladcnConnection = Nothing

    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError m_strCMODULENAME _
                    , "dbProcessVolunteerEntries"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError m_strCMODULENAME _
                    , "dbProcessVolunteerEntries" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case trErrdbInsIntoVolunteer, trErrdbDelFromVolunteer
                TrsRaiseError m_strCMODULENAME _
                    , "dbProcessVolunteerEntries" _
                    , Err.Number _
                    , , , _
                    , True
            Case Else
                TrsRaiseError m_strCMODULENAME _
                    , "dbProcessVolunteerEntries" _
                    , trErrdbSaveVolunteer _
                    , , , _
                    , True
    End Select

End Function
Public Function dbProcessECWorkReasonEntries(ByVal vadrsEmployeeECWorkReasons As ADODB.Recordset _
                    , ByVal vlUserEmployeeNumber As Long _
                    ) As Boolean
TrsRaiseError m_strCMODULENAME, "dbProcessECWorkReasonEntries", trErrObsoleteMethod, , , , True
'On Error GoTo dbProcessECWorkReasonEntries_EH
''ADODB Connection Object
'Dim ladcnConnection As ADODB.Connection
''Recordset
'Dim llRecordIndex As Long
''SQL
'Dim lsSQL As String
'
'    dbProcessECWorkReasonEntries = False
'
'    'Check for existence!
'    If Not (vadrsEmployeeECWorkReasons Is Nothing) Then
'        With vadrsEmployeeECWorkReasons
'            If Not (.EOF And .BOF) Then
'                .Filter = adFilterNone
'
'                'Check for deletes, if present process them first
'                .Filter = MDAC25FilterFix("RECORD_STATUS = 'D'")
'                If Not (.EOF) Then
'                    .MoveFirst
'
'                    For llRecordIndex = 1 To .RecordCount
'                        lsSQL = lsSQL & _
'                            " DELETE FROM  METHDBA.TTR_EC_WORK_REAS" & _
'                            " WHERE EMP_NUM = " & CStr(!KEY_EMP_NUM) & _
'                            " AND   TIME_STAMP = '" & Trim(!TIME_STAMP) & "';"
'
'                        .MoveNext
'                    Next llRecordIndex
'                End If
'
'                'Check for inserts
'                .Filter = adFilterNone
'                .Filter = MDAC25FilterFix("RECORD_STATUS = 'I'")
'                If Not (.EOF) Then
'                    .MoveFirst
'
'                    For llRecordIndex = 1 To .RecordCount
'                        lsSQL = lsSQL & _
'                            " INSERT    INTO METHDBA.TTR_EC_WORK_REAS" & _
'                            " (EMP_NUM, WORK_REASON" & _
'                            " , LAST_MOD_EMP_NUM, TIME_STAMP)" & _
'                            " VALUES " & _
'                            " (" & CStr(!EMP_NUM) & ", '" & Trim(!WORK_REASON) & "'" & _
'                            " , " & CStr(vlUserEmployeeNumber) & ", CURRENT TIMESTAMP);"
'
'                        .MoveNext
'                    Next llRecordIndex
'                End If
'
'                'Check for changes
'                .Filter = adFilterNone
'                .Filter = MDAC25FilterFix("RECORD_STATUS = 'C'")
'                If Not (.EOF) Then
'                    .MoveFirst
'
'                    For llRecordIndex = 1 To .RecordCount
'                        lsSQL = lsSQL & _
'                            " UPDATE    METHDBA.TTR_EC_WORK_REAS" & _
'                            " SET   WORK_REASON = '" & Trim(!WORK_REASON) & "'" & _
'                            "       , LAST_MOD_EMP_NUM = " & CStr(vlUserEmployeeNumber) & _
'                            "       , TIME_STAMP = CURRENT TIMESTAMP " & _
'                            " WHERE EMP_NUM = " & CStr(!KEY_EMP_NUM) & _
'                            " AND   TIME_STAMP = '" & Trim(!TIME_STAMP) & "';"
'
'                        .MoveNext
'                    Next llRecordIndex
'                End If
'
'                .Filter = adFilterNone
'            End If
'        End With
'    End If
'
'    If (Trim(lsSQL) <> "") Then
'        'Get Connection
'        Set ladcnConnection = OpenConnection()
'        'Call Routine to Execute SQL
'        ExecuteOTHERSQL lsSQL, ladcnConnection
'    End If
'
'    'close and destroy
'    If Not (ladcnConnection Is Nothing) Then
'        ladcnConnection.Close
'        Set ladcnConnection = Nothing
'    End If
'
'    dbProcessECWorkReasonEntries = True
'
'    'Complete Transaction
'    GetObjectContext.SetComplete
'
'    Exit Function
'dbProcessECWorkReasonEntries_EH:
'    'Abort Transaction
'    GetObjectContext.SetAbort
'
'    Set ladcnConnection = Nothing
'
'    Select Case (Err.Number)
'            Case trErrutOpenConnection
'                TrsRaiseError m_strCMODULENAME _
'                    , "dbProcessECWorkReasonEntries"
'            Case trErrdb_ExecuteSQL_F
'                TrsRaiseError m_strCMODULENAME _
'                    , "dbProcessECWorkReasonEntries" _
'                    , trErrdb_ExecuteSQL_F _
'                    , , _
'                    , DescArgsFrom(lsSQL) _
'                    , True
'            Case trErrdbInsIntoECWorkReasons, trErrdbUpdECWorkReasons _
'                , trErrdbDelFromECWorkReasons
'                TrsRaiseError m_strCMODULENAME _
'                    , "dbProcessECWorkReasonEntries" _
'                    , Err.Number _
'                    , , , _
'                    , True
'            Case Else
'                TrsRaiseError m_strCMODULENAME _
'                    , "dbProcessECWorkReasonEntries" _
'                    , trErrdbSaveECWorkReasons _
'                    , , , _
'                    , True
'    End Select

End Function
'END PREM

Public Function SaveVolunteerListChanges( _
    ByVal v_rsListChanges As ADODB.Recordset _
    , ByVal vlUserEmployeeNumber As Long _
                    ) As Boolean
            
    Dim l_objCN As ADODB.Connection
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_varIgnore(0) As String
    Dim l_aryPrimaryKey(0) As String
    Dim l_colErrorDesc As Collection
    Dim l_lngRecsAffected As Long

    On Error GoTo eh
           

    v_rsListChanges.Filter = adFilterPendingRecords
    If Not RecordsExist(v_rsListChanges) Then
        v_rsListChanges.Filter = ""
        GetObjectContext.SetComplete
        SaveVolunteerListChanges = True
        Exit Function
    End If
    
        Set l_objCN = New ADODB.Connection
    Set l_objCM = New clsuDBConnectMgr
    
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open

    l_aryPrimaryKey(0) = "EMP_NUM"
    l_varIgnore(0) = "TIME_STAMP"
    v_rsListChanges.MoveFirst
    l_strSQL = ""
    Do Until v_rsListChanges.EOF
        Select Case v_rsListChanges.Status
            Case adRecModified
                l_strSQL = UpdateStatement(v_rsListChanges, "METHDBA.TTR_VOLUNTEER", l_aryPrimaryKey, l_varIgnore, True)
            Case adRecNew
                l_strSQL = InsertStatement(v_rsListChanges, "METHDBA.TTR_VOLUNTEER", l_varIgnore)
            Case adRecDeleted
                l_strSQL = DeleteStatement(v_rsListChanges, "METHDBA.TTR_VOLUNTEER", l_aryPrimaryKey)
            Case Else
                l_strSQL = ""
        End Select
        If l_strSQL <> "" Then
            l_objCN.Execute l_strSQL, l_lngRecsAffected
            If l_lngRecsAffected = -1 Or l_lngRecsAffected = 0 Then
                 Err.Raise trErrAppCommitConflict
            End If
        End If
        v_rsListChanges.MoveNext
    Loop
    
    GetObjectContext.SetComplete
    l_objCN.Close
    Set l_objCN = Nothing
    SaveVolunteerListChanges = True
    
    Exit Function
        
eh:
     GetObjectContext.SetAbort

    Set l_objCN = Nothing
    SaveVolunteerListChanges = False

    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError m_strCMODULENAME _
                    , "SaveVolunteerListChanges"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError m_strCMODULENAME _
                    , "SaveVolunteerListChanges" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(l_strSQL) _
                    , True
            Case trErrdbInsIntoVolunteer, trErrdbDelFromVolunteer
                TrsRaiseError m_strCMODULENAME _
                    , "SaveVolunteerListChanges" _
                    , Err.Number _
                    , , , _
                    , True
            Case Else
                TrsRaiseError m_strCMODULENAME _
                    , "SaveVolunteerListChanges" _
                    , trErrdbSaveVolunteer _
                    , , , _
                    , True
    End Select
    

End Function


Public Function GetVolunteerChangeList() As ADODB.Recordset
On Error GoTo eh
'SQL Statement Container
Dim lsSQL As String

    lsSQL = " SELECT  EMP_NUM, STUS_EFF_DT, ACTN_CD " & _
        "           , CHAR(TIME_STAMP) AS TIME_STAMP " & _
        " FROM      METHDBA.TTR_VOLUNTEER " & _
        " ORDER BY  EMP_NUM "

    'Call ExecuteSQL and retrieve information
    Set GetVolunteerChangeList = ExecuteSQL(lsSQL _
                            , adOpenStatic _
                            , adLockBatchOptimistic)

    'Complete Transaction
    GetObjectContext.SetComplete

    Exit Function
eh:
    'Abort Transaction
    GetObjectContext.SetAbort

    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError m_strCMODULENAME _
                    , "GetVolunteerChangeList"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError m_strCMODULENAME _
                    , "GetVolunteerChangeList" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case Else
                TrsRaiseError m_strCMODULENAME _
                    , "GetVolunteerChangeList" _
                    , trErrdbGetVolunteer _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
    End Select

End Function
