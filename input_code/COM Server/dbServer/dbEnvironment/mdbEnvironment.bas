Attribute VB_Name = "mdbEnvironment"
Option Explicit

Private Const MODULENAME As String = "mdbEnvironment"
Private Const m_strCMODULENAME As String = MODULENAME
Private Const m_lngDefaultKeyOPT As Long = -999991
Public Const m_lngSerializationChunkSize As Long = 10240

Public Function DoesUserPreferenceExist(ByVal v_lngEmpNum As Long, _
                                        ByVal v_strTypeCd As String) As Boolean
                                    
    On Error GoTo eh
    
    Dim l_sSQLStatement As String
    Dim rs As ADODB.Recordset
    
    'Set SQL SELECT statment
    l_sSQLStatement = _
        "SELECT COUNT(*)" & _
        " FROM METHDBA.TTR_USER_PREF " & _
        " WHERE EMP_NUM = " & v_lngEmpNum & _
        "   AND TYPE_CD = '" & Left(Trim(v_strTypeCd), 3) & "'"
        
    Set rs = ExecuteSQL(l_sSQLStatement, adOpenStatic, adLockReadOnly)
    
    If RecordsExist(rs) Then
        If CLng(rs.Fields(0).Value) > 0 Then
            DoesUserPreferenceExist = True
        Else
            DoesUserPreferenceExist = False
        End If
    Else
        DoesUserPreferenceExist = False
    End If
                  
    Exit Function
eh:
    'Abort Transaction
'    GetObjectContext.SetAbort
    
    'Call the standard module function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME, "DoesUserPreferenceExist"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "DoesUserPreferenceExist" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(l_sSQLStatement) _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "DoesUserPreferenceExist" _
                    , trErrdb_TRISAvailable_F _
                    , , , _
                    , True
    End Select
End Function

Public Function GetUserPreference(ByVal v_lngEmpNum As Long, _
                                    ByVal v_strTypeCd As String) As ADODB.Recordset
                                    
    On Error GoTo eh
    
    Dim l_sSQLStatement As String
    Dim rs As ADODB.Recordset
    
    'Set SQL SELECT statment
    l_sSQLStatement = _
        "SELECT SETTINGS, FORMAT_CD" & _
        " FROM METHDBA.TTR_USER_PREF " & _
        " WHERE EMP_NUM = " & v_lngEmpNum & _
        "   AND TYPE_CD = '" & Left(Trim(v_strTypeCd), 3) & "'" & _
        " ORDER BY SEQ_NUM"
        
    Set rs = ExecuteSQL(l_sSQLStatement, adOpenStatic, adLockReadOnly)
    
    If RecordsExist(rs) Then
        '//create multiple records for each database row (settings column is a db2 blob)
        '//EMP_NUM, TYPE_CD, FORMAT_CD, SEQ_NUM, ROW_ID, SETTINGS, CREATED_EMP_NUM, LAST_MOD_EMP_NUM, CREATED_TIMESTAMP, LAST_MOD_TIMESTAMP
        Set GetUserPreference = CreateRSFromBlob(rs)
        
        '//increase the size of the value column for report settings (do once)
        If v_strTypeCd = "PST" Then
            '//Section, Key, Value, ValueType, LAST_MOD_TIMESTAMP
            If GetUserPreference.Fields("Value").DefinedSize = 32768 Or _
                GetUserPreference.Fields("Value").DefinedSize = 262144 Or _
                GetUserPreference.Fields("Value").DefinedSize = 524288 Then '//smaller size 32K or 256K, value is also a blob/LongVarBinary
                Dim rsNewPref As ADODB.Recordset
                
                Set rsNewPref = CreateUserPreferenceRS(v_strTypeCd) '//larger size 256K
                
                '//TODO: Could use GetChunk/AppendChunk to move data in the LongVarBinary column
                AppendRecordset GetUserPreference, rsNewPref
                Set GetUserPreference = rsNewPref
            End If
        End If
    Else
        Set GetUserPreference = CreateUserPreferenceRS(v_strTypeCd)
    End If
                  
    Exit Function
eh:
    'Abort Transaction
'    GetObjectContext.SetAbort
    
    'Call the standard module function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME, "GetUserPreference"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "GetUserPreference" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(l_sSQLStatement) _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "GetUserPreference" _
                    , trErrdb_TRISAvailable_F _
                    , , , _
                    , True
    End Select
End Function

'//-----could access db for defaults GetUserPreference(-999991, v_strTypeCd)
Public Function CreateUserPreferenceRS(ByVal v_strTypeCd As String) As ADODB.Recordset
    On Error GoTo eh
    
    Select Case v_strTypeCd
    Case "OPT"
        Set CreateUserPreferenceRS = GetUserPreference(m_lngDefaultKeyOPT, v_strTypeCd)
    Case "JOB"
        Set CreateUserPreferenceRS = CreateJobListRS
    Case "SCH"
        Set CreateUserPreferenceRS = CreateScheduleListRS
    Case "ROT"
        Set CreateUserPreferenceRS = CreateRotationListRS
    Case "EXC"
        Set CreateUserPreferenceRS = CreateExceptionListRS
    Case "OTR"
        Set CreateUserPreferenceRS = CreateOTReasonListRS
    Case "WG"
        Set CreateUserPreferenceRS = CreateWorkGroupListRS
    Case "PST"
        Set CreateUserPreferenceRS = CreateTRISPersistenceRS
    End Select
    
    Exit Function
eh:
    'Abort Transaction
'    GetObjectContext.SetAbort
    
    TrsRaiseError MODULENAME _
        , "CreateUserPreferenceRS" _
        , _
        , , , _
        , True
End Function

Private Function CreateWorkGroupListRS() As ADODB.Recordset
    Dim rs As New ADODB.Recordset

    With rs
        .Fields.Append "WG_NUM", adInteger, , adFldUpdatable
        .Fields.Append "CREW_NUM", adChar, 3, adFldUpdatable
        .Fields.Append "WG_DESC", adChar, 30, adFldUpdatable
        .Fields.Append "DEMAND_IND", adChar, 1, adFldUpdatable
        .Fields.Append "ACCESS_SOURCE", adVarChar, 10, adFldUpdatable
        .Fields.Append "APPR_LEVEL_IND", adInteger, , adFldUpdatable
        .Fields.Append "APP_OR_ENT", adVarChar, 10, adFldUpdatable
        .Fields.Append "APP_ENT_AUTHORITY", adChar, 1, adFldUpdatable
        .Fields.Append "EFF_START_DATE", adDBDate, , adFldUpdatable
        .Fields.Append "EFF_END_DATE", adDBDate, , adFldUpdatable
        .Fields.Append "WG_SUPERVISOR", adInteger, , adFldUpdatable
        .Fields.Append "WG_SUPERVISOR_NAME", adChar, 30, adFldUpdatable
        .Fields.Append "PROJECT_ID", adInteger, , adFldUpdatable
        .Fields.Append "DEPT_NUM", adChar, 3, adFldUpdatable
        .Fields.Append "LOC_CD", adSmallInt, , adFldUpdatable
        .Fields.Append "PLANT_CD", adChar, 1, adFldUpdatable
        .Fields.Append "CREATED_EMP_NUM", adInteger, , adFldUpdatable
        .Fields.Append "CREATED_TIMESTAMP", adDBTimeStamp, , adFldUpdatable
        .Fields.Append "LAST_MOD_EMP_NUM", adInteger, , adFldUpdatable
        .Fields.Append "LAST_MOD_TIMESTAMP", adDBTimeStamp, , adFldUpdatable
        .Fields.Append "HR_LOC", adChar, 1, adFldUpdatable
        .Fields.Append "ASSOC_DIV_NUM", adChar, 3, adFldUpdatable
        .Fields.Append "ASSOC_FUNC_NUM", adChar, 3, adFldUpdatable
        .Fields.Append "ASSOC_DEPT_NUM", adChar, 3, adFldUpdatable
        .Fields.Append "APP_TIME_STAMP", adDBTimeStamp, , adFldUpdatable
        .Fields.Append "APP_MOD_EMP_NUM", adInteger, , adFldUpdatable

        .CursorLocation = adUseClient
        .CursorType = adOpenDynamic
        .LockType = adLockBatchOptimistic
        .Open
    End With

    Set CreateWorkGroupListRS = rs
    Set rs = Nothing
    Exit Function
End Function

Private Function CreateScheduleListRS() As ADODB.Recordset
    Dim rs As New ADODB.Recordset

    With rs
        .Fields.Append "SCHED_CD", adInteger, , adFldUpdatable
        .Fields.Append "SCHED_DESC", adChar, 30, adFldUpdatable
        .Fields.Append "SUN_SHIFT_CD", adChar, 1, adFldUpdatable
        .Fields.Append "MON_SHIFT_CD", adChar, 1, adFldUpdatable
        .Fields.Append "TUE_SHIFT_CD", adChar, 1, adFldUpdatable
        .Fields.Append "WED_SHIFT_CD", adChar, 1, adFldUpdatable
        .Fields.Append "THU_SHIFT_CD", adChar, 1, adFldUpdatable
        .Fields.Append "FRI_SHIFT_CD", adChar, 1, adFldUpdatable
        .Fields.Append "SAT_SHIFT_CD", adChar, 1, adFldUpdatable

        .CursorLocation = adUseClient
        .CursorType = adOpenDynamic
        .LockType = adLockBatchOptimistic
        .Open
    End With

    Set CreateScheduleListRS = rs
    Set rs = Nothing
    Exit Function
End Function

Private Function CreateRotationListRS() As ADODB.Recordset
    Dim rs As New ADODB.Recordset

    With rs
        .Fields.Append "ROT_CD", adInteger, , adFldUpdatable
        .Fields.Append "ROT_DESC", adChar, 60, adFldUpdatable

        .CursorLocation = adUseClient
        .CursorType = adOpenDynamic
        .LockType = adLockBatchOptimistic
        .Open
    End With

    Set CreateRotationListRS = rs
    Set rs = Nothing
    Exit Function
End Function

Private Function CreateOTReasonListRS() As ADODB.Recordset
    Dim rs As New ADODB.Recordset

    With rs
        .Fields.Append "OT_REASON_CD", adVarChar, 3, adFldUpdatable
        .Fields.Append "OT_REASON_DESC", adChar, 30, adFldUpdatable
        .Fields.Append "LAST_MOD_EMP_NUM", adInteger, , adFldUpdatable
        .Fields.Append "TIME_STAMP", adDBTimeStamp, , adFldUpdatable

        .CursorLocation = adUseClient
        .CursorType = adOpenDynamic
        .LockType = adLockBatchOptimistic
        .Open
    End With

    Set CreateOTReasonListRS = rs
    Set rs = Nothing
    Exit Function
End Function

Private Function CreateJobListRS() As ADODB.Recordset
    Dim rs As New ADODB.Recordset

    With rs
        .Fields.Append "JOB_NUM", adChar, 6, adFldUpdatable
        .Fields.Append "JOB_DESC", adChar, 30, adFldUpdatable
        .Fields.Append "CRAFT_NUM", adChar, 3, adFldUpdatable
        .Fields.Append "OT_TYPE", adChar, 1, adFldUpdatable

        .CursorLocation = adUseClient
        .CursorType = adOpenDynamic
        .LockType = adLockBatchOptimistic
        .Open
    End With

    Set CreateJobListRS = rs
    Set rs = Nothing
    Exit Function
End Function

Private Function CreateExceptionListRS() As ADODB.Recordset
    Dim rs As New ADODB.Recordset

    With rs
        .Fields.Append "EXCEPTION_CD", adVarChar, 3, adFldUpdatable
        .Fields.Append "EXCEPTION_DESC", adChar, 30, adFldUpdatable
        .Fields.Append "EXCEPTION_TYPE", adChar, 1, adFldUpdatable
        .Fields.Append "CONTRACT_EARN_CD", adChar, 3, adFldUpdatable
        .Fields.Append "EXEC_EARN_CD", adChar, 3, adFldUpdatable
        .Fields.Append "SYMBOL_TIME_IND", adChar, 1, adFldUpdatable
        .Fields.Append "EXCEPTION_BANK_CD", adChar, 2, adFldUpdatable
        .Fields.Append "INELIGIBILITY_IND", adChar, 1, adFldUpdatable
        .Fields.Append "OTHERPAY_IND", adChar, 1, adFldUpdatable
        .Fields.Append "CASE_GROUP", adChar, 4, adFldUpdatable
        .Fields.Append "SCHED_IND", adChar, 1, adFldUpdatable
        .Fields.Append "FMLA_IND", adVarChar, 1, adFldUpdatable
        .Fields.Append "WORK_REST_IND", adChar, 1, adFldUpdatable
        .Fields.Append "SCHED_CHG_IND", adChar, 1, adFldUpdatable
        .Fields.Append "ACCT_NUM", adChar, 39, adFldUpdatable
        .Fields.Append "SYMBOL_CD", adChar, 3, adFldUpdatable

        .CursorLocation = adUseClient
        .CursorType = adOpenDynamic
        .LockType = adLockBatchOptimistic
        .Open
    End With

    Set CreateExceptionListRS = rs
    Set rs = Nothing
    Exit Function
End Function

Private Function CreateTRISPersistenceRS() As ADODB.Recordset
    Dim rs As New ADODB.Recordset

    With rs
'        .Fields.Append "AppName", adVarChar, 10, adFldUpdatable 'Ameren
        .Fields.Append "Section", adVarChar, 255, adFldUpdatable 'TRIS\Time\Collection\Day
        .Fields.Append "Key", adVarChar, 50, adFldUpdatable 'Automatic Refresh
        .Fields.Append "Value", adLongVarBinary, 1048576, adFldUpdatable '1M=1048576, 512K=524288, 256K=262144, "Y" or "N"
        .Fields.Append "ValueType", adVarChar, 30, adFldUpdatable '"Y" or "N"
'        .Fields.Append "CREATED_EMP_NUM", adInteger, , adFldUpdatable
'        .Fields.Append "LAST_MOD_EMP_NUM", adInteger, , adFldUpdatable
'        .Fields.Append "CREATED_TIMESTAMP", adDBTimeStamp, , adFldUpdatable
        .Fields.Append "LAST_MOD_TIMESTAMP", adDBTimeStamp, , adFldUpdatable

        .CursorLocation = adUseClient
        .CursorType = adOpenDynamic
        .LockType = adLockBatchOptimistic
        .Open
    End With

    Set CreateTRISPersistenceRS = rs
    Set rs = Nothing
    Exit Function
End Function

':TODO store these in the db with a key of emp_num=????? (99999)
' or derive column attributes using data dictionary
Private Function CreateTRISOptionsRS() As ADODB.Recordset
    Dim rs As New ADODB.Recordset

    With rs
        .Fields.Append "AppName", adVarChar, 10, adFldUpdatable 'Ameren
        .Fields.Append "Section", adVarChar, 255, adFldUpdatable 'TRIS\Time\Collection\Day
        .Fields.Append "Key", adVarChar, 50, adFldUpdatable 'Automatic Refresh
        .Fields.Append "Setting", adVarChar, 50, adFldUpdatable '"Y" or "N"
'        .Fields.Append "CREATED_EMP_NUM", adInteger, , adFldUpdatable
'        .Fields.Append "LAST_MOD_EMP_NUM", adInteger, , adFldUpdatable
'        .Fields.Append "CREATED_TIMESTAMP", adDBTimeStamp, , adFldUpdatable
        .Fields.Append "LAST_MOD_TIMESTAMP", adDBTimeStamp, , adFldUpdatable

        .CursorLocation = adUseClient
        .CursorType = adOpenDynamic
        .LockType = adLockBatchOptimistic
        .Open
    End With

    Set CreateTRISOptionsRS = rs
    Set rs = Nothing
    Exit Function
End Function

Private Function CreateRSFromBlob(ByVal rsBlob As ADODB.Recordset) As ADODB.Recordset

    On Error GoTo eh
    Dim sBlob As String, sFormatCd As String
    Dim rs As ADODB.Recordset
    
    If RecordsExist(rsBlob) Then
        rsBlob.MoveFirst
        
        Do While Not rsBlob.EOF
            sBlob = sBlob & rsBlob.Fields("SETTINGS").Value
            sFormatCd = rsBlob.Fields("FORMAT_CD").Value
        
            rsBlob.MoveNext
        Loop
        
        Select Case sFormatCd
        Case "RSX"
            Set rs = DeSerializeRS(sBlob)
        Case "BIN"
            Set rs = DeSerializeRS2(TRISCompressionType_eNone, ADOSerializationType_eADTG, sBlob)
        Case "ZIP"
            Set rs = DeSerializeRS2(TRISCompressionType_eGZip, ADOSerializationType_eADTG, sBlob)
        End Select
    End If
    
    Set CreateRSFromBlob = rs
    
    Exit Function
eh:
    'Abort Transaction
'    GetObjectContext.SetAbort
    
    TrsRaiseError MODULENAME _
        , "CreateRSFromBlob" _
        , _
        , , , _
        , True
End Function

Private Function DeSerializeRS(ByVal sSerializedRecordset As String) As ADODB.Recordset
    Dim InputStream As ADODB.Stream
    Dim rs As ADODB.Recordset
    Dim iChunkLength As Long
    Dim sSerializedRS As String
    
    iChunkLength = m_lngSerializationChunkSize
    
    Set InputStream = CreateObject("ADODB.Stream")
    InputStream.Open
    
    InputStream.WriteText sSerializedRecordset, ADODB.StreamWriteEnum.adWriteChar
    InputStream.Position = 0

    Set rs = CreateObject("ADODB.Recordset")
    rs.Open InputStream

    '//-----Clean Up-----//
    InputStream.Close
    
    Set DeSerializeRS = rs
'    GetObjectContext.SetComplete
    Set InputStream = Nothing
    Set rs = Nothing
    Exit Function
eh:
    'Abort Transaction
'    GetObjectContext.SetAbort
    
    Set InputStream = Nothing
    Set rs = Nothing
    
    TrsRaiseError MODULENAME _
        , "DeSerializeRS" _
        , _
        , , , _
        , True
End Function

Private Function DeSerializeRS2(ByVal CompressionType As TRISServerUtilityCOM.TRISCompressionType, _
                                ByVal SerializationType As TRISServerUtilityCOM.ADOSerializationType, _
                                ByVal sSerializedRecordset As String) As ADODB.Recordset
    Dim oUtility As TRISServerUtilityCOM.Utility
    
    Set oUtility = CreateObject("TRISServerUtilityCOM.Utility")
    
    Set DeSerializeRS2 = oUtility.StringToRecordset(CompressionType, SerializationType, sSerializedRecordset)

    Set oUtility = Nothing
'    GetObjectContext.SetComplete
    Exit Function
eh:
    'Abort Transaction
'    GetObjectContext.SetAbort
    
    Set oUtility = Nothing
    
    TrsRaiseError MODULENAME _
        , "DeSerializeRS2" _
        , _
        , , , _
        , True
End Function


