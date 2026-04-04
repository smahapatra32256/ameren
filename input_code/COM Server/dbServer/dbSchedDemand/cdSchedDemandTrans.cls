VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "cdSchedDemandTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const m_strCMODULENAME As String = "cdSchedDemandTrans"
Private Const m_datCOPTIONALSTARTDATE As Date = #1/1/100#
Private Const m_datCOPTIONALENDDATE As Date = #12/31/9999#
'//-----used the following for testing
Private Const m_strConnectStringReg As String = "SHADOWDIRECT"
'DB2TEST_UDB
'SHADOWDIRECT


Public Function SaveAssignments(ByVal v_rsAssignmentRequest As ADODB.Recordset, ByVal v_lTRISUser As Long) As Boolean
'//, ByVal v_rsJobPosRequest As ADODB.Recordset, ByVal v_rs21stWatchRequest As ADODB.Recordset
    Dim cn As ADODB.Connection
    Dim rc As Boolean, bError As Boolean
On Error GoTo eh

    Set cn = OpenConnection

    With v_rsAssignmentRequest
        bError = False
        If Not (v_rsAssignmentRequest Is Nothing) Then
            '//-----The following code was moved to DeleteAssignments so that it is done prior to validation-----//
        '  Perform Deletes
            .Filter = adFilterNone
            .MoveFirst
            '//-----clean up job pos and 21st Watch-----//
            .Filter = MDAC25FilterFix("RECORD_STATUS = 'D'")
             Do While Not .EOF
                rc = DeleteAssignment(v_rsAssignmentRequest, cn)
                If Not rc Then bError = True
                .MoveNext
            Loop

        '  Perform Inserts
            .Filter = adFilterNone
            .MoveFirst
            '//-----an insert with emp_num=0 means to unassign, so skip it here (AdjustForOverlap should do the rest)-----//
            '//.Filter = MDAC25FilterFix("RECORD_STATUS = 'I' AND EMP_NUM > 0")
            .Filter = MDAC25FilterFix("RECORD_STATUS = 'I'") '//-----check for valid empnum in InsertAssignment
            Do While Not .EOF
                rc = InsertAssignment(v_rsAssignmentRequest, v_lTRISUser, cn)
                If Not rc Then bError = True
                .MoveNext
            Loop

        '  Perform Changes
            .Filter = adFilterNone
            .MoveFirst
            .Filter = MDAC25FilterFix("RECORD_STATUS = 'C'")
            Do While Not .EOF
                rc = UpdateAssignment(v_rsAssignmentRequest, v_lTRISUser, cn)
                If Not rc Then bError = True
                .MoveNext
            Loop
        .Filter = adFilterNone
        End If
    End With

    cn.Close

    Set cn = Nothing
    GetObjectContext.SetComplete

    SaveAssignments = Not bError

    Exit Function
eh:
    Dim adoError As ADODB.Error
    Dim sError As String
    
    GetObjectContext.SetAbort
    
    If Not v_rsAssignmentRequest Is Nothing Then
        sError = v_rsAssignmentRequest.GetString
    End If
    
    TrsRaiseError m_strCMODULENAME, "SaveAssignments", , True, , DescArgsFrom(sError), True
End Function
    
Private Function DeleteAssignment(v_rsAssignmentRequest As ADODB.Recordset, cn As ADODB.Connection) As Boolean
    Dim l_strSQL As String
    Dim l_lngRecsAffected As Long

    On Error GoTo eh

    With v_rsAssignmentRequest
        l_strSQL = "DELETE FROM METHDBA.TTR_LINE_ASGNM"
        l_strSQL = l_strSQL & " WHERE WG_NUM = " & !KEY_WG_NUM
        l_strSQL = l_strSQL & " AND LINE_NUM = " & !KEY_LINE_NUM
        l_strSQL = l_strSQL & " AND START_DATE = " & dbFormat(!KEY_START_DATE)
        l_strSQL = l_strSQL & " AND END_DATE = " & dbFormat(!KEY_END_DATE)
        l_strSQL = l_strSQL & " AND EMP_NUM = " & dbFormat(!KEY_EMP_NUM)
        l_strSQL = l_strSQL & " AND LAST_MOD_TIMESTAMP = " & dbFormat(!LAST_MOD_TIMESTAMP)
    End With
       
    If l_strSQL <> "" Then
       cn.Execute l_strSQL, l_lngRecsAffected
       If l_lngRecsAffected = -1 Or l_lngRecsAffected = 0 Then
            Err.Raise trErrAppCommitConflict
       End If
    End If
    
    DeleteAssignment = True
Exit Function
eh:

    Select Case Err.Number
        Case trErrAppCommitConflict
            Err.Source = _
                Join( _
                    Array( _
                        "1", _
                        CStr(v_rsAssignmentRequest!KEY_WG_NUM), _
                        CStr(v_rsAssignmentRequest!KEY_LINE_NUM), _
                        Format(v_rsAssignmentRequest!KEY_START_DATE, "mm/dd/yyyy"), _
                        Format(v_rsAssignmentRequest!KEY_END_DATE, "mm/dd/yyyy"), _
                        CStr(v_rsAssignmentRequest!KEY_EMP_NUM)) _
                    , ";")
            TrsRaiseError m_strCMODULENAME, "DeleteAssignment", Err.Number
        Case Else
            TrsRaiseError m_strCMODULENAME, "DeleteAssignment", , True, , DescArgsFrom(l_strSQL), True
    End Select
End Function

Private Function InsertAssignment(v_rsAssignmentRequest As ADODB.Recordset, ByVal v_lTRISUser As Long, cn As ADODB.Connection) As Boolean
    Dim l_strSQL As String

On Error GoTo eh

    With v_rsAssignmentRequest
        '//-----a request of EmpNum=0 is really an unassign to anyone on the line,
        '//-----which would be done in busSchedDemand by finding the overlaps in the scope
        '//-----??? somehow an EmpNum of 0 was inserted, so the filter was simplied and an extra condition added here ???-----//
        If !EMP_NUM > 0 And Not IsEmpty(!EMP_NUM) And Not IsNull(!EMP_NUM) Then
            l_strSQL = "INSERT INTO METHDBA.TTR_LINE_ASGNM"
            l_strSQL = l_strSQL & " (WG_NUM"
            l_strSQL = l_strSQL & ", LINE_NUM"
            l_strSQL = l_strSQL & ", START_DATE"
            l_strSQL = l_strSQL & ", END_DATE"
            l_strSQL = l_strSQL & ", EMP_NUM"
            If Not (IsNull(!RELIEF_COMP) Or IsEmpty(!RELIEF_COMP)) Then '//if null let db2 assign the default
                l_strSQL = l_strSQL & ", RELIEF_COMP"
            End If
            l_strSQL = l_strSQL & ", CREATED_EMP_NUM"
            l_strSQL = l_strSQL & ", LAST_MOD_EMP_NUM"
            l_strSQL = l_strSQL & ", CREATED_TIMESTAMP"
            l_strSQL = l_strSQL & ", LAST_MOD_TIMESTAMP)"
            l_strSQL = l_strSQL & " VALUES(" & !WG_NUM
            l_strSQL = l_strSQL & ", " & !LINE_NUM
            l_strSQL = l_strSQL & ", " & dbFormat(!START_DATE)
            l_strSQL = l_strSQL & ", " & dbFormat(!END_DATE)
            l_strSQL = l_strSQL & ", " & !EMP_NUM
            If Not (IsNull(!RELIEF_COMP) Or IsEmpty(!RELIEF_COMP)) Then '//if null let db2 assign the default
                l_strSQL = l_strSQL & ", " & dbFormat(!RELIEF_COMP)
            End If
            
            If v_lTRISUser >= 0 Then
                '//-----because an Update ("C") request changes to a delete/insert,
                '//----- need to carry the created columns over-----//
                '//-----also needed for conversion of 21st Watch which will preserve the last_mod columns-----//
                If IsNull(!CREATED_EMP_NUM) Or !CREATED_EMP_NUM = 0 Then
                    l_strSQL = l_strSQL & ", " & v_lTRISUser
                Else
                    l_strSQL = l_strSQL & ", " & !CREATED_EMP_NUM '//-----not a new insert, but re-inserted-----//
                End If
                
                l_strSQL = l_strSQL & ", " & v_lTRISUser
                
                If IsNull(!CREATED_TIMESTAMP) Or Trim(!CREATED_TIMESTAMP) = "" Then
                    l_strSQL = l_strSQL & ", CURRENT TIMESTAMP "
                Else
                    l_strSQL = l_strSQL & ", " & dbFormat(!CREATED_TIMESTAMP) '//-----not a new insert, but re-inserted-----//
                End If
                
                l_strSQL = l_strSQL & ", CURRENT TIMESTAMP )"
            Else '//-----for conversion of legacy data (ttr_base_hours_exc) 21st Watch-----//
                l_strSQL = l_strSQL & ", " & !CREATED_EMP_NUM
                l_strSQL = l_strSQL & ", 999999999" '//& !LAST_MOD_EMP_NUM
                l_strSQL = l_strSQL & ", " & dbFormat(!CREATED_TIMESTAMP)
                l_strSQL = l_strSQL & ", CURRENT TIMESTAMP )" '// & dbFormat(!LAST_MOD_TIMESTAMP) & " )"
            End If

            If l_strSQL <> "" Then
               cn.Execute l_strSQL
            End If
        End If
    End With
    InsertAssignment = True
Exit Function
eh:
    Dim adoError As ADODB.Error
    GetObjectContext.SetAbort
    For Each adoError In cn.Errors
         If adoError.NativeError = -803 Then '//-----db2 duplicate key-----//
            Err.Raise trErrbusSchedDemandDuplicateLineAsgnm
            Exit Function
         Else
            TrsRaiseError m_strCMODULENAME, "InsertAssignment", , True, , DescArgsFrom(l_strSQL), True
            Exit Function
         End If
    Next

    InsertAssignment = False
    TrsRaiseError m_strCMODULENAME, "InsertAssignment", , True, , DescArgsFrom(l_strSQL), True
End Function
Private Function UpdateAssignment(v_rsAssignmentRequest As ADODB.Recordset, ByVal v_lTRISUser As Long, cn As ADODB.Connection) As Boolean
    Dim l_strSQL As String
    Dim l_lngRecsAffected As Long

On Error GoTo eh

    With v_rsAssignmentRequest
        l_strSQL = "UPDATE METHDBA.TTR_LINE_ASGNM SET"
        l_strSQL = l_strSQL & " WG_NUM = " & !WG_NUM
        l_strSQL = l_strSQL & ", LINE_NUM = " & !LINE_NUM
        l_strSQL = l_strSQL & ", START_DATE = " & dbFormat(!START_DATE)
        l_strSQL = l_strSQL & ", END_DATE = " & dbFormat(!END_DATE)
        l_strSQL = l_strSQL & ", EMP_NUM = " & !EMP_NUM
        If Not (IsNull(!RELIEF_COMP) Or IsEmpty(!RELIEF_COMP)) Then '//if null don't update
            l_strSQL = l_strSQL & ", RELIEF_COMP = " & dbFormat(!RELIEF_COMP)
        End If
        l_strSQL = l_strSQL & ", LAST_MOD_EMP_NUM = " & v_lTRISUser
        l_strSQL = l_strSQL & ", LAST_MOD_TIMESTAMP = CURRENT TIMESTAMP"
        l_strSQL = l_strSQL & " WHERE WG_NUM = " & !KEY_WG_NUM
        l_strSQL = l_strSQL & " AND LINE_NUM = " & !KEY_LINE_NUM
        l_strSQL = l_strSQL & " AND START_DATE = " & dbFormat(!KEY_START_DATE)
        l_strSQL = l_strSQL & " AND END_DATE = " & dbFormat(!KEY_END_DATE)
        l_strSQL = l_strSQL & " AND EMP_NUM = " & dbFormat(!KEY_EMP_NUM)
        l_strSQL = l_strSQL & " AND LAST_MOD_TIMESTAMP = " & dbFormat(!LAST_MOD_TIMESTAMP)
    End With

    If l_strSQL <> "" Then
       cn.Execute l_strSQL, l_lngRecsAffected
       If l_lngRecsAffected = -1 Or l_lngRecsAffected = 0 Then
            Err.Raise trErrAppCommitConflict
       End If
    End If

    UpdateAssignment = True
Exit Function
eh:
    GetObjectContext.SetAbort
    
    UpdateAssignment = False
   
    Select Case Err.Number
        Case trErrAppCommitConflict
            Err.Source = _
                Join( _
                    Array( _
                        "1", _
                        CStr(v_rsAssignmentRequest!KEY_WG_NUM), _
                        CStr(v_rsAssignmentRequest!KEY_LINE_NUM), _
                        Format(v_rsAssignmentRequest!KEY_START_DATE, "mm/dd/yyyy"), _
                        Format(v_rsAssignmentRequest!KEY_END_DATE, "mm/dd/yyyy"), _
                        CStr(v_rsAssignmentRequest!KEY_EMP_NUM)) _
                    , ";")
            TrsRaiseError m_strCMODULENAME, "UpdateAssignment", Err.Number
        Case Else
            TrsRaiseError m_strCMODULENAME, "UpdateAssignment", , True, , DescArgsFrom(l_strSQL), True
    End Select
    
End Function
Private Function OpenConnection() As ADODB.Connection
On Error GoTo OpenConnection_EH
Dim cm As New util_DBConnectMgr.clsuDBConnectMgr
Dim cn As ADODB.Connection

    Set cn = New ADODB.Connection
    With cn
        .CursorLocation = adUseClient       'Must use client to support RecordCount
        .ConnectionString = cm.GetConnectionString(m_strConnectStringReg)
        .Open
    End With

    Set OpenConnection = cn
    Set cm = Nothing
    Set cn = Nothing

    Exit Function
OpenConnection_EH:
    Set cm = Nothing
    Set cn = Nothing

    TrsRaiseError m_strCMODULENAME, "OpenConnection"
End Function
Private Function dbFormat(ByVal v_varData As Variant) As String
    Dim l_strFormated As String, fieldType As Integer
    
    On Error GoTo eh
    
    fieldType = VarType(v_varData)
    Select Case fieldType
        Case vbString
            l_strFormated = "'" & Replace(v_varData, "'", "''") & "'"
        Case vbInteger, vbLong
            l_strFormated = CStr(v_varData)
        Case vbDate
            l_strFormated = "'" & Format(v_varData, "mm/dd/yyyy") & "'"
        Case Else
            l_strFormated = Format(v_varData)
    End Select
    dbFormat = l_strFormated
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "dbFormat", , , , DescArgsFrom(CStr(v_varData)), True
End Function
'BEGIN PREM TRIS CONVERSION PHASE 3
Public Function dbSaveJobPosition(ByVal vadrsJobPosition As ADODB.Recordset _
                , ByVal vlTRISUserNumber As Long _
                ) As Boolean
On Error GoTo dbSaveJobPosition_EH
Dim lsSQL As String, llIndex As Long
Dim ladcnConnection As ADODB.Connection
Dim llRecordsAffected As Long

    dbSaveJobPosition = False

    If Not (RecordsExist(vadrsJobPosition)) Then
        GetObjectContext.SetComplete
        Exit Function
    End If

    With vadrsJobPosition
        .Filter = adFilterNone

        'DELETES
        .Filter = MDAC25FilterFix("RECORD_STATUS = 'D'")
        If Not (.EOF) Then
            .MoveFirst

            For llIndex = 1 To .RecordCount
                lsSQL = lsSQL & _
                    " DELETE FROM    METHDBA.TTR_DEMAND_JOB_POS" & _
                    " WHERE WG_NUM = " & CStr(!KEY_WG_NUM) & _
                    " AND   LINE_NUM = " & CStr(!LINE_NUM) & _
                    " AND   CHARGE_DATE = '" & CStr(!KEY_CHARGE_DATE) & "'" & _
                    " AND   TIME_STAMP = '" & Trim(!TIME_STAMP) & "';"

                .MoveNext
            Next llIndex
        End If

        'INSERTS
        .Filter = adFilterNone
        .Filter = MDAC25FilterFix("RECORD_STATUS = 'I'")
        If Not (.EOF) Then
            .MoveFirst

            For llIndex = 1 To .RecordCount
                lsSQL = lsSQL & _
                    " INSERT INTO    METHDBA.TTR_DEMAND_JOB_POS" & _
                    " (WG_NUM, LINE_NUM" & _
                    " , CHARGE_DATE, JOB_POSITION" & _
                    " , LAST_MOD_EMP_NUM, TIME_STAMP)" & _
                    " VALUES " & _
                    " (" & CStr(!WG_NUM) & ", " & CStr(!LINE_NUM) & _
                    " , '" & CStr(!CHARGE_DATE) & "', '" & Trim(!JOB_POSITION) & "'" & _
                    " , " & CStr(vlTRISUserNumber) & ", CURRENT TIMESTAMP);"

                .MoveNext
            Next llIndex
        End If

        'UPDATE
        .Filter = adFilterNone
        .Filter = MDAC25FilterFix("RECORD_STATUS = 'C'")
        If Not (.EOF) Then
            .MoveFirst

            For llIndex = 1 To .RecordCount
                lsSQL = lsSQL & _
                    " UPDATE METHDBA.TTR_DEMAND_JOB_POS" & _
                    " SET    JOB_POSITION = '" & Trim(!JOB_POSITION) & "'" & _
                    "       , LAST_MOD_EMP_NUM = " & CStr(vlTRISUserNumber) & _
                    "       , TIME_STAMP = CURRENT TIMESTAMP" & _
                    " WHERE WG_NUM = " & CStr(!KEY_WG_NUM) & _
                    " AND   LINE_NUM = " & CStr(!LINE_NUM) & _
                    " AND   CHARGE_DATE = '" & CStr(!KEY_CHARGE_DATE) & "'" & _
                    " AND   TIME_STAMP = '" & Trim(!TIME_STAMP) & "';"

                .MoveNext
            Next llIndex
        End If

        .Filter = adFilterNone
    End With

    If (Trim(lsSQL) <> "") Then
        Set ladcnConnection = OpenConnection

        ladcnConnection.Execute lsSQL, , adCmdText
    End If

    Set ladcnConnection = Nothing

    dbSaveJobPosition = True

    GetObjectContext.SetComplete
    Exit Function
dbSaveJobPosition_EH:
    GetObjectContext.SetAbort

    Set ladcnConnection = Nothing

    Select Case (Err.Number)
        Case trErrdbDelDemandJobPositions, trErrdbInsDemandJobPositions _
            , trErrdbUpdDemandJobPositions
            TrsRaiseError m_strCMODULENAME, "dbSaveJobPosition" _
                , , True, _
                , DescArgsFrom(lsSQL) _
                , True
        Case Else
            TrsRaiseError m_strCMODULENAME, "dbSaveJobPosition" _
                , trErrdbSaveDemandJobPositions _
                , , , _
                , True
    End Select
End Function
'END PREM TRIS CONVERSION PHASE 3
