VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "cdCodeTablesTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const m_strCMODULENAME As String = "cdCodeTablesTrans"
Public Function dbProcessCraftSave( _
            ByVal vadrsOTCodeTableEntries As ADODB.Recordset _
            , ByVal vlUserEmployeeNumber As Long _
            ) As Boolean
On Error GoTo dbProcessCraftSave_EH
'ADODB Connection Object
Dim ladcnConnection As ADODB.Connection
'Recordset
Dim llRecordIndex As Long
'SQL
Dim lsSQL As String

    dbProcessCraftSave = False
    lsSQL = ""

    'Check for existence!
    If Not (vadrsOTCodeTableEntries Is Nothing) Then
        With vadrsOTCodeTableEntries
            If Not (.EOF And .BOF) Then
                .Filter = adFilterNone

                If Not (.EOF) Then
                    'Check for deletes, if present process them first
                    .Filter = "RECORD_STATUS = 'D'"
                    If Not (.EOF) Then
                        .MoveFirst

                        For llRecordIndex = 1 To .RecordCount
                            lsSQL = lsSQL & _
                                    " DELETE    FROM    METHDBA.TTR_CRAFTS " & _
                                    " WHERE     CRAFT_NUM = '" & Trim(!KEY_CRAFT_NUM) & "'" & _
                                    " AND       TIME_STAMP = '" & Trim(!TIME_STAMP) & "';"

                            .MoveNext
                        Next llRecordIndex
                    End If
                
                    .Filter = adFilterNone
                    'Check for inserts, if present process them first
                    .Filter = "RECORD_STATUS = 'I'"
                    If Not (.EOF) Then
                        .MoveFirst

                        For llRecordIndex = 1 To .RecordCount
                            lsSQL = lsSQL & _
                                    " INSERT    INTO METHDBA.TTR_CRAFTS " & _
                                    " (CRAFT_NUM, CRAFT_DESC, CRAFT_GROUP " & _
                                    " , JOURN_APPR, TIME_STAMP, LAST_MOD_EMP_NUM) " & _
                                    " VALUES " & _
                                    " ('" & Trim(!CRAFT_NUM) & "', '" & Trim(!CRAFT_DESC) & "', '" & Trim(!CRAFT_GROUP) & "'" & _
                                    " , '" & Trim(!JOURN_APPR) & "', CURRENT TIMESTAMP, " & CStr(vlUserEmployeeNumber) & ");"

                            .MoveNext
                        Next llRecordIndex
                    End If

                    .Filter = adFilterNone
                    'Check for deletes, if present process them first
                    .Filter = "RECORD_STATUS = 'C'"
                    If Not (.EOF) Then
                        .MoveFirst

                        For llRecordIndex = 1 To .RecordCount
                            lsSQL = lsSQL & _
                                    " UPDATE    METHDBA.TTR_CRAFTS " & _
                                    " SET       CRAFT_DESC = '" & Trim(!CRAFT_DESC) & "'" & _
                                    "           , CRAFT_GROUP = '" & Trim(!CRAFT_GROUP) & "'" & _
                                    "           , JOURN_APPR = '" & Trim(!JOURN_APPR) & "'" & _
                                    "           , TIME_STAMP = CURRENT TIMESTAMP " & _
                                    "           , LAST_MOD_EMP_NUM = " & CStr(vlUserEmployeeNumber) & _
                                    " WHERE     CRAFT_NUM = '" & Trim(!KEY_CRAFT_NUM) & "'" & _
                                    " AND       TIME_STAMP = '" & Trim(!TIME_STAMP) & "';"

                            .MoveNext
                        Next llRecordIndex
                    End If

                    .Filter = adFilterNone
                End If
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

    dbProcessCraftSave = True

    'Complete Transaction
    GetObjectContext.SetComplete

    Exit Function
dbProcessCraftSave_EH:
    'Abort Transaction
    GetObjectContext.SetAbort

    Set ladcnConnection = Nothing
    dbProcessCraftSave = False

    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError m_strCMODULENAME _
                    , "dbProcessCraftSave"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError m_strCMODULENAME _
                    , "dbProcessCraftSave" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
'            Case trErrdbInsIntoOTCDTBLCraft, trErrdbUpdOTCDTBLCraft _
'                , trErrdbDelFromOTCDTBLCraft
'                TrsRaiseError m_strCMODULENAME _
'                    , "dbProcessCraftSave" _
'                    , Err.Number _
'                    , , , _
'                    , True
            Case Else
                TrsRaiseError m_strCMODULENAME _
                    , "dbProcessCraftSave" _
                    , trErrdbOTCDTBLSaveCraft _
                    , , , _
                    , True
    End Select

End Function
Public Function dbProcessQualificationsSave( _
            ByVal vadrsOTCodeTableEntries As ADODB.Recordset _
            , ByVal vlUserEmployeeNumber As Long _
            ) As Boolean
On Error GoTo dbProcessQualificationsSave_EH
'ADODB Connection Object
Dim ladcnConnection As ADODB.Connection
'Recordset
Dim llRecordIndex As Long
'SQL
Dim lsSQL As String

    dbProcessQualificationsSave = False

    lsSQL = ""

    'Check for existence!
    If Not (vadrsOTCodeTableEntries Is Nothing) Then
        With vadrsOTCodeTableEntries
            If Not (.EOF And .BOF) Then
                .Filter = adFilterNone

                If Not (.EOF) Then
                    'Check for deletes, if present process them first
                    .Filter = "RECORD_STATUS = 'D'"
                    If Not (.EOF) Then
                        .MoveFirst

                        For llRecordIndex = 1 To .RecordCount
                            lsSQL = lsSQL & _
                                    " DELETE    FROM    METHDBA.TTR_QUALIFICATIONS " & _
                                    " WHERE     QUAL_CD = '" & Trim(!KEY_QUAL_CD) & "'" & _
                                    " AND       TIME_STAMP = '" & Trim(!TIME_STAMP) & "';"

                            .MoveNext
                        Next llRecordIndex
                    End If
                
                    .Filter = adFilterNone
                    'Check for inserts, if present process them first
                    .Filter = "RECORD_STATUS = 'I'"
                    If Not (.EOF) Then
                        .MoveFirst

                        For llRecordIndex = 1 To .RecordCount
                            lsSQL = lsSQL & _
                                    " INSERT    INTO METHDBA.TTR_QUALIFICATIONS " & _
                                    " (QUAL_CD, QUAL_DESC, TIME_STAMP, LAST_MOD_EMP_NUM) " & _
                                    " VALUES " & _
                                    " ('" & Trim(!QUAL_CD) & "', '" & Trim(!QUAL_DESC) & "'" & _
                                    " , CURRENT TIMESTAMP, " & CStr(vlUserEmployeeNumber) & ");"

                            .MoveNext
                        Next llRecordIndex
                    End If

                    .Filter = adFilterNone
                    'Check for deletes, if present process them first
                    .Filter = "RECORD_STATUS = 'C'"

                    If Not (.EOF) Then
                        .MoveFirst

                        For llRecordIndex = 1 To .RecordCount
                            lsSQL = lsSQL & _
                                    " UPDATE    METHDBA.TTR_QUALIFICATIONS " & _
                                    " SET       QUAL_DESC = '" & Trim(!QUAL_DESC) & "'" & _
                                    "           , TIME_STAMP = CURRENT TIMESTAMP " & _
                                    "           , LAST_MOD_EMP_NUM = " & CStr(vlUserEmployeeNumber) & _
                                    " WHERE     QUAL_CD = '" & Trim(!KEY_QUAL_CD) & "'" & _
                                    " AND       TIME_STAMP = '" & Trim(!TIME_STAMP) & "'"

                            .MoveNext
                        Next llRecordIndex
                    End If

                    .Filter = adFilterNone
                End If
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

    dbProcessQualificationsSave = True

    'Complete Transaction
    GetObjectContext.SetComplete

    Exit Function
dbProcessQualificationsSave_EH:
    'Abort Transaction
    GetObjectContext.SetAbort

    Set ladcnConnection = Nothing

    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError m_strCMODULENAME _
                    , "dbProcessQualificationsSave"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError m_strCMODULENAME _
                    , "dbProcessQualificationsSave" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
'            Case trErrdbInsIntoOTCDTBLQuals, trErrdbUpdOTCDTBLQuals _
'                , trErrdbDelFromOTCDTBLQuals
'                TrsRaiseError m_strCMODULENAME _
'                    , "dbProcessQualificationsSave" _
'                    , Err.Number _
'                    , , , _
'                    , True
            Case Else
                TrsRaiseError m_strCMODULENAME _
                    , "dbProcessQualificationsSave" _
                    , trErrdbOTCDTBLSaveQuals _
                    , , , _
                    , True
    End Select

End Function
Public Function dbProcessSpecialSkillsSave( _
            ByVal vadrsOTCodeTableEntries As ADODB.Recordset _
            , ByVal vlUserEmployeeNumber As Long _
            ) As Boolean
On Error GoTo dbProcessSpecialSkillsSave_EH
'ADODB Connection Object
Dim ladcnConnection As ADODB.Connection
'Recordset
Dim llRecordIndex As Long
'SQL
Dim lsSQL As String

    dbProcessSpecialSkillsSave = False
    lsSQL = ""

    'Check for existence!
    If Not (vadrsOTCodeTableEntries Is Nothing) Then
        With vadrsOTCodeTableEntries
            If Not (.EOF And .BOF) Then
                .Filter = adFilterNone

                If Not (.EOF) Then
                    'Check for deletes, if present process them first
                    .Filter = "RECORD_STATUS = 'D'"
                    If Not (.EOF) Then
                        .MoveFirst

                        For llRecordIndex = 1 To .RecordCount
                            lsSQL = lsSQL & _
                                    " DELETE    FROM    METHDBA.TTR_SPEC_SKILLS " & _
                                    " WHERE     SPEC_SKILL_CD = '" & Trim(!KEY_SPEC_SKILL_CD) & "'" & _
                                    " AND       QUAL_CD = '" & Trim(!KEY_QUAL_CD) & "'" & _
                                    " AND       TIME_STAMP = '" & Trim(!TIME_STAMP) & "';"

                            .MoveNext
                        Next llRecordIndex
                    End If

                    .Filter = adFilterNone
                    'Check for inserts, if present process them first
                    .Filter = "RECORD_STATUS = 'I'"
                    If Not (.EOF) Then
                        .MoveFirst

                        For llRecordIndex = 1 To .RecordCount
                            lsSQL = lsSQL & _
                                    " INSERT    INTO METHDBA.TTR_SPEC_SKILLS " & _
                                    " (SPEC_SKILL_CD, SPEC_SKILL_DESC, QUAL_CD " & _
                                    " , TIME_STAMP, LAST_MOD_EMP_NUM) " & _
                                    " VALUES " & _
                                    " ('" & Trim(!SPEC_SKILL_CD) & "', '" & Trim(!SPEC_SKILL_DESC) & "', '" & Trim(!QUAL_CD) & "'" & _
                                    " , CURRENT TIMESTAMP, " & CStr(vlUserEmployeeNumber) & ");"

                            .MoveNext
                        Next llRecordIndex
                    End If

                    .Filter = adFilterNone
                    'Check for deletes, if present process them first
                    .Filter = "RECORD_STATUS = 'C'"
                    If Not (.EOF) Then
                        .MoveFirst

                        For llRecordIndex = 1 To .RecordCount
                            lsSQL = lsSQL & _
                                    " UPDATE    METHDBA.TTR_SPEC_SKILLS " & _
                                    " SET       SPEC_SKILL_DESC = '" & Trim(!SPEC_SKILL_DESC) & "'" & _
                                    "           , TIME_STAMP = CURRENT TIMESTAMP " & _
                                    "           , LAST_MOD_EMP_NUM = " & CStr(vlUserEmployeeNumber) & _
                                    " WHERE     SPEC_SKILL_CD = '" & Trim(!KEY_SPEC_SKILL_CD) & "'" & _
                                    " AND       QUAL_CD = '" & Trim(!KEY_QUAL_CD) & "'" & _
                                    " AND       TIME_STAMP = '" & Trim(!TIME_STAMP) & "';"

                            .MoveNext
                        Next llRecordIndex
                    End If

                    .Filter = adFilterNone
                End If
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

    dbProcessSpecialSkillsSave = True

    'Complete Transaction
    GetObjectContext.SetComplete

    Exit Function
dbProcessSpecialSkillsSave_EH:
    'Abort Transaction
    GetObjectContext.SetAbort

    Set ladcnConnection = Nothing

    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError m_strCMODULENAME _
                    , "dbProcessSpecialSkillsSave"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError m_strCMODULENAME _
                    , "dbProcessSpecialSkillsSave" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
'            Case trErrdbInsIntoOTCDTBLSkills, trErrdbUpdOTCDTBLSkills _
'                , trErrdbDelFromOTCDTBLSkills
'                TrsRaiseError m_strCMODULENAME _
'                    , "dbProcessSpecialSkillsSave" _
'                    , Err.Number _
'                    , , , _
'                    , True
            Case Else
                TrsRaiseError m_strCMODULENAME _
                    , "dbProcessSpecialSkillsSave" _
                    , trErrdbOTCDTBLSaveSkills _
                    , , , _
                    , True
    End Select

End Function
Public Function dbProcessOTReasonsSave( _
            ByVal vadrsOTCodeTableEntries As ADODB.Recordset _
            , ByVal vlUserEmployeeNumber As Long _
            ) As Boolean
On Error GoTo dbProcessOTReasonsSave_EH
'ADODB Connection Object
Dim ladcnConnection As ADODB.Connection
'Recordset
Dim llRecordIndex As Long
'SQL
Dim lsSQL As String

    dbProcessOTReasonsSave = False
    lsSQL = ""

    'Check for existence!
    If Not (vadrsOTCodeTableEntries Is Nothing) Then
        With vadrsOTCodeTableEntries
            If Not (.EOF And .BOF) Then
                .Filter = adFilterNone

                If Not (.EOF) Then
                    'Check for deletes, if present process them first
                    .Filter = "RECORD_STATUS = 'D'"
                    If Not (.EOF) Then
                        .MoveFirst

                        For llRecordIndex = 1 To .RecordCount
                            lsSQL = lsSQL & _
                                    " DELETE    FROM    METHDBA.TTR_OT_REASONS " & _
                                    " WHERE     OT_REASON_CD = '" & Trim(!KEY_OT_REASON_CD) & "'" & _
                                    " AND       TIME_STAMP = '" & Trim(!TIME_STAMP) & "';"

                            .MoveNext
                        Next llRecordIndex
                    End If

                    .Filter = adFilterNone
                    'Check for inserts, if present process them first
                    .Filter = "RECORD_STATUS = 'I'"
                    If Not (.EOF) Then
                        .MoveFirst

                        For llRecordIndex = 1 To .RecordCount
                            lsSQL = lsSQL & _
                                    " INSERT    INTO METHDBA.TTR_OT_REASONS " & _
                                    " (OT_REASON_CD, OT_REASON_DESC " & _
                                    " , TIME_STAMP, LAST_MOD_EMP_NUM) " & _
                                    " VALUES " & _
                                    " ('" & Trim(!OT_REASON_CD) & "', '" & Trim(!OT_REASON_DESC) & "'" & _
                                    " , CURRENT TIMESTAMP, " & CStr(vlUserEmployeeNumber) & ");"

                            .MoveNext
                        Next llRecordIndex
                    End If

                    .Filter = adFilterNone
                    'Check for deletes, if present process them first
                    .Filter = "RECORD_STATUS = 'C'"
                    If Not (.EOF) Then
                        .MoveFirst

                        For llRecordIndex = 1 To .RecordCount
                            lsSQL = lsSQL & _
                                    " UPDATE    METHDBA.TTR_OT_REASONS " & _
                                    " SET       OT_REASON_DESC = '" & Trim(!OT_REASON_DESC) & "'" & _
                                    "           , TIME_STAMP = CURRENT TIMESTAMP " & _
                                    "           , LAST_MOD_EMP_NUM = " & CStr(vlUserEmployeeNumber) & _
                                    " WHERE     OT_REASON_CD = '" & Trim(!KEY_OT_REASON_CD) & "'" & _
                                    " AND       TIME_STAMP = '" & Trim(!TIME_STAMP) & "';"

                            .MoveNext
                        Next llRecordIndex
                    End If

                    .Filter = adFilterNone
                End If
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

    dbProcessOTReasonsSave = True

    'Complete Transaction
    GetObjectContext.SetComplete

    Exit Function
dbProcessOTReasonsSave_EH:
    'Abort Transaction
    GetObjectContext.SetAbort

    Set ladcnConnection = Nothing

    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError m_strCMODULENAME _
                    , "dbProcessOTReasonsSave"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError m_strCMODULENAME _
                    , "dbProcessOTReasonsSave" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
'            Case trErrdbInsIntoOTCDTBLReasons, trErrdbUpdOTCDTBLReasons _
'                , trErrdbDelFromOTCDTBLReasons
'                TrsRaiseError m_strCMODULENAME _
'                    , "dbProcessOTReasonsSave" _
'                    , Err.Number _
'                    , , , _
'                    , True
            Case Else
                TrsRaiseError m_strCMODULENAME _
                    , "dbProcessOTReasonsSave" _
                    , trErrdbOTCDTBLSaveReasons _
                    , , , _
                    , True
    End Select

End Function
Public Function dbProcessOTResponsesSave( _
            ByVal vadrsOTCodeTableEntries As ADODB.Recordset _
            , ByVal vlUserEmployeeNumber As Long _
            ) As Boolean
On Error GoTo dbProcessOTResponsesSave_EH
'ADODB Connection Object
Dim ladcnConnection As ADODB.Connection
'Recordset
Dim llRecordIndex As Long
'SQL
Dim lsSQL As String

    dbProcessOTResponsesSave = False
    lsSQL = ""

    'Check for existence!
    If Not (vadrsOTCodeTableEntries Is Nothing) Then
        With vadrsOTCodeTableEntries
            If Not (.EOF And .BOF) Then
                .Filter = adFilterNone

                If Not (.EOF) Then
                    'Check for deletes, if present process them first
                    .Filter = "RECORD_STATUS = 'D'"
                    If Not (.EOF) Then
                        .MoveFirst

                        For llRecordIndex = 1 To .RecordCount
                            lsSQL = lsSQL & _
                                    " DELETE    FROM    METHDBA.TTR_OT_RESPONSES " & _
                                    " WHERE     OT_RESP_CD = '" & Trim(!KEY_OT_RESP_CD) & "'" & _
                                    " AND       TIME_STAMP = '" & Trim(!TIME_STAMP) & "';"

                            .MoveNext
                        Next llRecordIndex
                    End If

                    .Filter = adFilterNone
                    'Check for inserts, if present process them first
                    .Filter = "RECORD_STATUS = 'I'"
                    If Not (.EOF) Then
                        .MoveFirst

                        For llRecordIndex = 1 To .RecordCount
                            lsSQL = lsSQL & _
                                    " INSERT    INTO METHDBA.TTR_OT_RESPONSES " & _
                                    " (OT_RESP_CD, RESPONSE_DESC " & _
                                    " , OT_WORKED_IND, OT_CHRG_IND " & _
                                    " , OT_CONTACTED_IND " & _
                                    " , TIME_STAMP, LAST_MOD_EMP_NUM) " & _
                                    " VALUES " & _
                                    " ('" & Trim(!OT_RESP_CD) & "', '" & Trim(!RESPONSE_DESC) & "'" & _
                                    " , '" & Trim(!OT_WORKED_IND) & "', '" & Trim(!OT_CHRG_IND) & "'" & _
                                    " , '" & Trim(!OT_CONTACTED_IND) & "'" & _
                                    " , CURRENT TIMESTAMP, " & CStr(vlUserEmployeeNumber) & ");"

                            .MoveNext
                        Next llRecordIndex
                    End If

                    .Filter = adFilterNone
                    'Check for deletes, if present process them first
                    .Filter = "RECORD_STATUS = 'C'"
                    If Not (.EOF) Then
                        .MoveFirst

                        For llRecordIndex = 1 To .RecordCount
                            lsSQL = lsSQL & _
                                    " UPDATE    METHDBA.TTR_OT_RESPONSES " & _
                                    " SET       RESPONSE_DESC = '" & Trim(!RESPONSE_DESC) & "'" & _
                                    "           , OT_WORKED_IND = '" & Trim(!OT_WORKED_IND) & "'" & _
                                    "           , OT_CHRG_IND = '" & Trim(!OT_CHRG_IND) & "'" & _
                                    "           , OT_CONTACTED_IND = '" & Trim(!OT_CONTACTED_IND) & "'" & _
                                    "           , TIME_STAMP = CURRENT TIMESTAMP " & _
                                    "           , LAST_MOD_EMP_NUM = " & CStr(vlUserEmployeeNumber) & _
                                    " WHERE     OT_RESP_CD = '" & Trim(!KEY_OT_RESP_CD) & "'" & _
                                    " AND       TIME_STAMP = '" & Trim(!TIME_STAMP) & "';"

                            .MoveNext
                        Next llRecordIndex
                    End If

                    .Filter = adFilterNone
                End If
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

    dbProcessOTResponsesSave = True

    'Complete Transaction
    GetObjectContext.SetComplete

    Exit Function
dbProcessOTResponsesSave_EH:
    'Abort Transaction
    GetObjectContext.SetAbort

    Set ladcnConnection = Nothing

    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError m_strCMODULENAME _
                    , "dbProcessOTResponsesSave"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError m_strCMODULENAME _
                    , "dbProcessOTResponsesSave" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
''            Case trErrdbInsIntoOTCDTBLResponses, trErrdbUpdOTCDTBLResponses _
''                , trErrdbDelFromOTCDTBLResponses
''                TrsRaiseError m_strCMODULENAME _
''                    , "dbProcessOTResponsesSave" _
''                    , Err.Number _
''                    , , , _
''                    , True
            Case Else
                TrsRaiseError m_strCMODULENAME _
                    , "dbProcessOTResponsesSave" _
                    , trErrdbOTCDTBLSaveResponses _
                    , , , _
                    , True
    End Select

End Function
Public Function dbProcessAssociateCraftToJob( _
            ByVal vadrsOTCodeTableEntries As ADODB.Recordset _
            , ByVal vlUserEmployeeNumber As Long _
            ) As Boolean
On Error GoTo dbProcessAssociateCraftToJob_EH
'ADODB Connection Object
Dim ladcnConnection As ADODB.Connection
'Recordset
Dim llRecordIndex As Long
'SQL
Dim lsSQL As String

    dbProcessAssociateCraftToJob = False
    lsSQL = ""

    'Check for existence!
    If Not (vadrsOTCodeTableEntries Is Nothing) Then
        With vadrsOTCodeTableEntries
            If Not (.EOF And .BOF) Then
                .Filter = adFilterNone

                If Not (.EOF) Then
                    'Check for deletes, if present process them first
                    .Filter = "RECORD_STATUS = 'C'"
                    If Not (.EOF) Then
                        .MoveFirst

                        For llRecordIndex = 1 To .RecordCount
                            lsSQL = lsSQL & _
                                    " UPDATE    METHDBA.TTR_JOBS " & _
                                    " SET       CRAFT_NUM = '" & Trim(!CRAFT_NUM) & "'" & _
                                    "           , TIME_STAMP = CURRENT TIMESTAMP " & _
                                    "           , LAST_MOD_EMP_NUM = " & CStr(vlUserEmployeeNumber) & _
                                    " WHERE     JOB_NUM = '" & Trim(!KEY_JOB_NUM) & "'" & _
                                    " AND       TIME_STAMP = '" & Trim(!TIME_STAMP) & "';"

                            .MoveNext
                        Next llRecordIndex
                    End If

                    .Filter = adFilterNone
                End If
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

    dbProcessAssociateCraftToJob = True

    'Complete Transaction
    GetObjectContext.SetComplete

    Exit Function
dbProcessAssociateCraftToJob_EH:
    'Abort Transaction
    GetObjectContext.SetAbort

    Set ladcnConnection = Nothing

    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError m_strCMODULENAME _
                    , "dbProcessAssociateCraftToJob"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError m_strCMODULENAME _
                    , "dbProcessAssociateCraftToJob" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
            Case trErrdbUpdOTCDTBLCraftJob
                TrsRaiseError m_strCMODULENAME _
                    , "dbProcessAssociateCraftToJob" _
                    , Err.Number _
                    , , , _
                    , True
            Case Else
                TrsRaiseError m_strCMODULENAME _
                    , "dbProcessAssociateCraftToJob" _
                    , trErrdbOTCDTBLSaveCraftJob _
                    , , , _
                    , True
    End Select

End Function
Public Function dbProcessPromotionalLines( _
            ByVal vadrsPromotionalLinesEntries As ADODB.Recordset _
            , ByVal vlUserEmployeeNumber As Long _
            ) As Boolean
On Error GoTo dbProcessPromotionalLines_EH
'ADODB Connection Object
Dim ladcnConnection As ADODB.Connection
'Recordset
Dim llRecordIndex As Long
'SQL
Dim lsSQL As String

    dbProcessPromotionalLines = False
    lsSQL = ""

    'Check for existence!
    If Not (vadrsPromotionalLinesEntries Is Nothing) Then
        With vadrsPromotionalLinesEntries
            If Not (.EOF And .BOF) Then
                .Filter = adFilterNone

                If Not (.EOF) Then
                    'Check for deletes, if present process them first
                    .Filter = "RECORD_STATUS = 'D'"
                    If Not (.EOF) Then
                        .MoveFirst

                        For llRecordIndex = 1 To .RecordCount
                            lsSQL = lsSQL & _
                                    " DELETE    FROM    METHDBA.TTR_PROM_LINES " & _
                                    " WHERE     DEPT_NUM = '" & Trim(!KEY_DEPT_NUM) & "'" & _
                                    " AND       HR_LOC = '" & Trim(!KEY_HR_LOC) & "'" & _
                                    " AND       JOB_NUM = '" & Trim(!KEY_JOB_NUM) & "'" & _
                                    " AND       IR_AREA_CD = '" & Trim(!KEY_IR_AREA_CD) & "'" & _
                                    " AND       IR_DIV_CD = '" & Trim(!KEY_IR_DIV_CD) & "'" & _
                                    " AND       LOP_CURR_LEV_IND = '" & Trim(!KEY_LOP_CURR_LEV_IND) & "'" & _
                                    " AND       TIME_STAMP = '" & Trim(!TIME_STAMP) & "';"

                            .MoveNext
                        Next llRecordIndex
                    End If

                    .Filter = adFilterNone
                    'Check for inserts, if present process them first
                    .Filter = "RECORD_STATUS = 'I'"
                    If Not (.EOF) Then
                        .MoveFirst

                        For llRecordIndex = 1 To .RecordCount
                            lsSQL = lsSQL & _
                                    " INSERT    INTO METHDBA.TTR_PROM_LINES " & _
                                    " (DEPT_NUM, HR_LOC, JOB_NUM " & _
                                    " , IR_AREA_CD, IR_DIV_CD " & _
                                    " , LOP_CURR_LEV_IND, LOP_NEXT_LEV_IND " & _
                                    " , TIME_STAMP, LAST_MOD_EMP_NUM) " & _
                                    " VALUES " & _
                                    " ('" & Trim(!DEPT_NUM) & "', '" & Trim(!HR_LOC) & "', '" & Trim(!JOB_NUM) & "'" & _
                                    " , '" & Trim(!IR_AREA_CD) & "', '" & Trim(!IR_DIV_CD) & "'" & _
                                    " , '" & Trim(!LOP_CURR_LEV_IND) & "', '" & Trim(!LOP_NEXT_LEV_IND) & "'" & _
                                    " , CURRENT TIMESTAMP, " & CStr(vlUserEmployeeNumber) & ");"

                            .MoveNext
                        Next llRecordIndex
                    End If

                    .Filter = adFilterNone
                    'Check for deletes, if present process them first
                    .Filter = "RECORD_STATUS = 'C'"
                    If Not (.EOF) Then
                        .MoveFirst

                        For llRecordIndex = 1 To .RecordCount
                            lsSQL = lsSQL & _
                                    " UPDATE    METHDBA.TTR_PROM_LINES " & _
                                    " SET       IR_AREA_CD = '" & Trim(!IR_AREA_CD) & "'" & _
                                    "           , IR_DIV_CD = '" & Trim(!IR_DIV_CD) & "'" & _
                                    "           , LOP_CURR_LEV_IND = '" & Trim(!LOP_CURR_LEV_IND) & "'" & _
                                    "           , LOP_NEXT_LEV_IND = '" & Trim(!LOP_NEXT_LEV_IND) & "'" & _
                                    "           , TIME_STAMP = CURRENT TIMESTAMP " & _
                                    "           , LAST_MOD_EMP_NUM = " & CStr(vlUserEmployeeNumber) & _
                                    " WHERE     DEPT_NUM = '" & Trim(!KEY_DEPT_NUM) & "'" & _
                                    " AND       HR_LOC = '" & Trim(!KEY_HR_LOC) & "'" & _
                                    " AND       JOB_NUM = '" & Trim(!KEY_JOB_NUM) & "'" & _
                                    " AND       IR_AREA_CD = '" & Trim(!KEY_IR_AREA_CD) & "'" & _
                                    " AND       IR_DIV_CD = '" & Trim(!KEY_IR_DIV_CD) & "'" & _
                                    " AND       LOP_CURR_LEV_IND = '" & Trim(!KEY_LOP_CURR_LEV_IND) & "'" & _
                                    " AND       TIME_STAMP = '" & Trim(!TIME_STAMP) & "';"

                            .MoveNext
                        Next llRecordIndex
                    End If

                    .Filter = adFilterNone
                End If
            End If
        End With
    End If

    'Get Connection
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

    dbProcessPromotionalLines = True

    'Complete Transaction
    GetObjectContext.SetComplete

    Exit Function
dbProcessPromotionalLines_EH:
    'Abort Transaction
    GetObjectContext.SetAbort

    Set ladcnConnection = Nothing

    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError m_strCMODULENAME _
                    , "dbProcessPromotionalLines"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError m_strCMODULENAME _
                    , "dbProcessPromotionalLines" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
''            Case trErrdbInsIntoPromLines, trErrdbUpdPromLines _
''                , trErrdbDelFromPromLines
''                TrsRaiseError m_strCMODULENAME _
''                    , "dbProcessPromotionalLines" _
''                    , Err.Number _
''                    , , , _
''                    , True
            Case Else
                TrsRaiseError m_strCMODULENAME _
                    , "dbProcessPromotionalLines" _
                    , trErrdbSavePromLines _
                    , , , _
                    , True
    End Select

End Function
Public Function dbProcessOTBankHeaderSave( _
            ByVal vadrsOTCodeTableEntries As ADODB.Recordset _
            , ByVal vlUserEmployeeNumber As Long _
            ) As Boolean
On Error GoTo dbProcessOTBankHeaderSave_EH
'ADODB Connection Object
Dim ladcnConnection As ADODB.Connection
'Recordset
Dim llRecordIndex As Long
'SQL
Dim lsSQL As String, llOTBankCode As Long
'SEQ_NUM
Dim lobjUtility As utilDBDataServices.cuDBDataServices

    dbProcessOTBankHeaderSave = False
    lsSQL = ""

    'Check for existence!
    If Not (vadrsOTCodeTableEntries Is Nothing) Then
        With vadrsOTCodeTableEntries
            If Not (.EOF And .BOF) Then
                .Filter = adFilterNone

                If Not (.EOF) Then
                    'Check for deletes, if present process them first
                    .Filter = "RECORD_STATUS = 'D'"
                    If Not (.EOF) Then
                        .MoveFirst

                        For llRecordIndex = 1 To .RecordCount
                            lsSQL = lsSQL & _
                                    " DELETE    FROM    METHDBA.TTR_OT_BANK_CD " & _
                                    " WHERE     OT_BANK_CD = " & !OT_BANK_CD & _
                                    " AND       CREATED_TIMESTAMP = '" & Trim(!CREATED_TIMESTAMP) & "';"

                            .MoveNext
                        Next llRecordIndex
                    End If

                    .Filter = adFilterNone
                    'Check for inserts, if present process them first
                    .Filter = "RECORD_STATUS = 'I'"
                    If Not (.EOF) Then
                        .MoveFirst

                        For llRecordIndex = 1 To .RecordCount
                            Set lobjUtility = CreateObject( _
                                                "utilDBDataServices.cuDBDataServices")
                            llOTBankCode = lobjUtility.GetNextSequenceNumber("OB")
                            Set lobjUtility = Nothing
'PREM PLEASE REMOVE PROJ_OR_STAND_IND, TEMP_ASSIGN_IND AFTER
'DATA CONVERSION.
                            lsSQL = lsSQL & _
                                    " INSERT    INTO METHDBA.TTR_OT_BANK_CD " & _
                                    " (OT_BANK_CD, OT_BANK_DESC " & _
                                    " , CREATED_TIMESTAMP, CREATED_EMP_NUM " & _
                                    " , LAST_MOD_TIMESTAMP, LAST_MOD_EMP_NUM " & _
                                    " , PROJ_OR_STAND_IND, TEMP_ASSIGN_IND) " & _
                                    " VALUES " & _
                                    " (" & CStr(llOTBankCode) & ", '" & Trim(!OT_BANK_DESC) & "'" & _
                                    " , CURRENT TIMESTAMP, " & CStr(vlUserEmployeeNumber) & _
                                    " , CURRENT TIMESTAMP, " & CStr(vlUserEmployeeNumber) & _
                                    " , 'TRIS5', 'T');"

                            .MoveNext
                        Next llRecordIndex
                    End If

                    .Filter = adFilterNone
                    'Check for deletes, if present process them first
                    .Filter = "RECORD_STATUS = 'C'"
                    If Not (.EOF) Then
                        .MoveFirst

                        For llRecordIndex = 1 To .RecordCount
                            lsSQL = lsSQL & _
                                    " UPDATE    METHDBA.TTR_OT_BANK_CD " & _
                                    " SET       OT_BANK_DESC = '" & Trim(!OT_BANK_DESC) & "'" & _
                                    "           , LAST_MOD_TIMESTAMP = CURRENT TIMESTAMP " & _
                                    "           , LAST_MOD_EMP_NUM = " & CStr(vlUserEmployeeNumber) & _
                                    " WHERE     OT_BANK_CD = " & !OT_BANK_CD & _
                                    " AND       CREATED_TIMESTAMP = '" & Trim(!CREATED_TIMESTAMP) & "';"

                            .MoveNext
                        Next llRecordIndex
                    End If

                    .Filter = adFilterNone
                End If
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

    dbProcessOTBankHeaderSave = True

    'Complete Transaction
    GetObjectContext.SetComplete

    Exit Function
dbProcessOTBankHeaderSave_EH:
    'Abort Transaction
    GetObjectContext.SetAbort

    Set lobjUtility = Nothing
    Set ladcnConnection = Nothing

    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError m_strCMODULENAME _
                    , "dbProcessOTBankHeaderSave"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError m_strCMODULENAME _
                    , "dbProcessOTBankHeaderSave" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(lsSQL) _
                    , True
'            Case trErrdbDelFromOTBankCode, trErrdbInsIntoOTBankCode _
'                , trErrdbUpdOTBankCode
'                TrsRaiseError m_strCMODULENAME _
'                    , "dbProcessOTBankHeaderSave" _
'                    , Err.Number _
'                    , , , _
'                    , True
            Case Else
                TrsRaiseError m_strCMODULENAME _
                    , "dbProcessOTBankHeaderSave" _
                    , trErrdbSaveOTBankCode _
                    , , , _
                    , True
    End Select

End Function
