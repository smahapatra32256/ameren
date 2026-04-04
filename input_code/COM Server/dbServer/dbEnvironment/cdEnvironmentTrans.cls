VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cdEnvironmentTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const MODULENAME As String = "cdFunSecurityTrans"
Private Const m_strCMODULENAME As String = MODULENAME

'//-----------------------------------------------------------------------------------------------------------------------//
'//-----this class has been changed to Non-Transactional, cbsEnvironmentTrans2 is now the transactional class
'//-----------------------------------------------------------------------------------------------------------------------//

'Private Const m_lngBLOBSize As Long = 1048576 '1024 * 1024
'Private Const m_lngSerializationChunkSize As Long = 10240
'Private Const m_lngDefaultKeyOPT As Long = -999991
'Public Enum PreferenceTypeCode
'    UserPreference = 0
'    JobList
'    ScheduleList
'    RotationList
'    ExceptionList
'    OTReasonList
'    WorkGroupList
'End Enum

'for testing with db2QLUW
'ExecuteSQL2
'OpenConnection2
'NewExecuteOtherSQL2

Public Function GetUserPreference(ByVal v_lngEmpNum As Long, _
                                    ByVal v_strTypeCd As String) As ADODB.Recordset
                                    
    On Error GoTo eh
    
    Set GetUserPreference = mdbEnvironment.GetUserPreference(v_lngEmpNum, v_strTypeCd)
    
'    Dim l_sSQLStatement As String
'    Dim rs As ADODB.Recordset
'
'    'Set SQL SELECT statment
'    l_sSQLStatement = _
'        "SELECT SETTINGS, FORMAT_CD" & _
'        " FROM METHDBA.TTR_USER_PREF " & _
'        " WHERE EMP_NUM = " & v_lngEmpNum & _
'        "   AND TYPE_CD = '" & Left(Trim(v_strTypeCd), 3) & "'" & _
'        " ORDER BY SEQ_NUM"
'
'    Set rs = ExecuteSQL(l_sSQLStatement, adOpenStatic, adLockReadOnly)
'
'    If RecordsExist(rs) Then
'        '//create multiple records for each database row (settings column is a db2 blob)
'        '//EMP_NUM, TYPE_CD, FORMAT_CD, SEQ_NUM, ROW_ID, SETTINGS, CREATED_EMP_NUM, LAST_MOD_EMP_NUM, CREATED_TIMESTAMP, LAST_MOD_TIMESTAMP
'        Set GetUserPreference = CreateRSFromBlob(rs)
'
'        '//increase the size of the value column for report settings (do once)
'        If v_strTypeCd = "PST" Then
'            '//Section, Key, Value, ValueType, LAST_MOD_TIMESTAMP
'            If GetUserPreference.Fields("Value").DefinedSize = 32768 Or _
'                GetUserPreference.Fields("Value").DefinedSize = 262144 Or _
'                GetUserPreference.Fields("Value").DefinedSize = 524288 Then '//smaller size 32K or 256K, value is also a blob/LongVarBinary
'                Dim rsNewPref As ADODB.Recordset
'
'                Set rsNewPref = CreateUserPreferenceRS(v_strTypeCd) '//larger size 256K
'
'                '//TODO: Could use GetChunk/AppendChunk to move data in the LongVarBinary column
'                AppendRecordset GetUserPreference, rsNewPref
'                Set GetUserPreference = rsNewPref
'            End If
'        End If
'    Else
'        Set GetUserPreference = CreateUserPreferenceRS(v_strTypeCd)
'    End If
                  
    Exit Function
eh:
    'Abort Transaction
'    GetObjectContext.SetAbort
    
'    'Call the standard module function
'    'Raise the appropriate error message depending on the error number
'    Select Case (Err.Number)
'            Case trErrutOpenConnection
'                TrsRaiseError MODULENAME, "GetUserPreference"
'            Case trErrdb_ExecuteSQL_F
'                TrsRaiseError MODULENAME _
'                    , "GetUserPreference" _
'                    , trErrdb_ExecuteSQL_F _
'                    , , _
'                    , DescArgsFrom(l_sSQLStatement) _
'                    , True
'            Case Else
                TrsRaiseError MODULENAME _
                    , "GetUserPreference" _
                    , trErrdb_TRISAvailable_F _
                    , , , _
                    , True
'    End Select
End Function

'Private Function GetUserPreferenceState(ByVal v_lngEmpNum As Long, _
'                                    ByVal v_strTypeCd As String) As ADODB.Recordset
'
'    On Error GoTo eh
'
'    Dim l_sSQLStatement As String
'    Dim rs As ADODB.Recordset, rs2 As ADODB.Recordset
'    Dim i As Long, j As Long
'
'    'Set SQL SELECT statment
'    l_sSQLStatement = _
'        "SELECT EMP_NUM, TYPE_CD, FORMAT_CD, SEQ_NUM, SETTINGS, " & _
'        " CREATED_EMP_NUM, Char(CREATED_TIMESTAMP) as CREATED_TIMESTAMP, " & _
'        " LAST_MOD_EMP_NUM, Char(LAST_MOD_TIMESTAMP) as LAST_MOD_TIMESTAMP, " & _
'        " EMP_NUM as KEY_EMP_NUM, TYPE_CD as KEY_TYPE_CD, SEQ_NUM as KEY_SEQ_NUM, 'D' AS RECORD_STATUS " & _
'        " FROM METHDBA.TTR_USER_PREF " & _
'        " WHERE EMP_NUM = " & v_lngEmpNum & _
'        "   AND TYPE_CD = '" & Left(Trim(v_strTypeCd), 3) & "'" & _
'        " ORDER BY SEQ_NUM"
'
'    Set rs = ExecuteSQL(l_sSQLStatement, adOpenStatic, adLockBatchOptimistic)
'    Set rs2 = CreateUserPreferenceStateRS
'
'    AppendRecordset rs, rs2
'''    '//-----copy
''    For j = 1 To rs.RecordCount
''        rs2.AddNew
''        For i = 1 To rs.Fields.Count
''            If rs2.Fields(i - 1).Type <> adLongVarBinary Then
''                rs2.Fields(i - 1).Value = rs.Fields(i - 1).Value
''            Else
''                rs2.Fields(i - 1).AppendChunk rs.Fields(i - 1).Value
''            End If
''        Next
''        rs2.Update
''        rs.MoveNext
''    Next
'
'    Set GetUserPreferenceState = rs2
'
'    Exit Function
'eh:
'    'Abort Transaction
'    GetObjectContext.SetAbort
'
'    'Call the standard module function
'    'Raise the appropriate error message depending on the error number
'    Select Case (Err.Number)
'            Case trErrutOpenConnection
'                TrsRaiseError MODULENAME, "GetUserPreferenceState"
'            Case trErrdb_ExecuteSQL_F
'                TrsRaiseError MODULENAME _
'                    , "GetUserPreferenceState" _
'                    , trErrdb_ExecuteSQL_F _
'                    , , _
'                    , DescArgsFrom(l_sSQLStatement) _
'                    , True
'            Case Else
'                TrsRaiseError MODULENAME _
'                    , "GetUserPreferenceState" _
'                    , trErrdb_TRISAvailable_F _
'                    , , , _
'                    , True
'    End Select
'End Function

''need the blob column created with an type of adLongVarBinary (attribute = adFldLong) in order to use AppendChunk, GetChunk
'Private Function CreateUserPreferenceStateRS() As ADODB.Recordset
'    Dim rs As New ADODB.Recordset
'
'    With rs
'        .Fields.Append "EMP_NUM", adInteger, , adFldUpdatable
'        .Fields.Append "TYPE_CD", adChar, 3, adFldUpdatable
'        .Fields.Append "FORMAT_CD", adChar, 3, adFldUpdatable
'        .Fields.Append "SEQ_NUM", adSmallInt, , adFldUpdatable
'        .Fields.Append "SETTINGS", adLongVarBinary, 10240, adFldUpdatable
'        .Fields.Append "CREATED_EMP_NUM", adInteger, , adFldUpdatable
'        .Fields.Append "LAST_MOD_EMP_NUM", adInteger, , adFldUpdatable
'        .Fields.Append "CREATED_TIMESTAMP", adDBTimeStamp, , adFldUpdatable
'        .Fields.Append "LAST_MOD_TIMESTAMP", adDBTimeStamp, , adFldUpdatable
'        .Fields.Append "KEY_EMP_NUM", adInteger, , adFldUpdatable
'        .Fields.Append "KEY_TYPE_CD", adChar, 3, adFldUpdatable
'        .Fields.Append "KEY_SEQ_NUM", adSmallInt, , adFldUpdatable
'        .Fields.Append "RECORD_STATUS", adVarChar, 1, adFldUpdatable
'
'        .CursorLocation = adUseClient
'        .CursorType = adOpenStatic
'        .LockType = adLockBatchOptimistic
'        .Open
'    End With
'
'    Set CreateUserPreferenceStateRS = rs
'    Set rs = Nothing
'    Exit Function
'End Function

''//-----use with OPT
'Public Function GetUserPreferenceDefault(ByVal v_strTypeCd As String) As ADODB.Recordset
'
'    On Error GoTo eh
'
'    Dim l_sSQLStatement As String
'    Dim rs As ADODB.Recordset
'
'    'Set SQL SELECT statment
'    l_sSQLStatement = _
'        "SELECT SETTINGS" & _
'        " FROM METHDBA.TTR_USER_PREF " & _
'        " WHERE EMP_NUM = " & m_lngDefaultKey & _
'        "   AND TYPE_CD = '" & Left(Trim(v_strTypeCd), 3) & "'" & _
'        " ORDER BY SEQ_NUM"
'
'    Set rs = ExecuteSQL(l_sSQLStatement, adOpenStatic, adLockReadOnly)
'
'    If RecordsExist(rs) Then
'        Set GetUserPreference = CreateRSFromBlob(rs)
'    Else
'        Set GetUserPreference = CreateUserPreferenceRS(v_strTypeCd)
'    End If
'
'    Exit Function
'eh:
'    'Abort Transaction
'    GetObjectContext.SetAbort
'
'    'Call the standard module function
'    'Raise the appropriate error message depending on the error number
'    Select Case (Err.Number)
'            Case trErrutOpenConnection
'                TrsRaiseError MODULENAME, "GetUserPreferenceDefault"
'            Case trErrdb_ExecuteSQL_F
'                TrsRaiseError MODULENAME _
'                    , "GetUserPreferenceDefault" _
'                    , trErrdb_ExecuteSQL_F _
'                    , , _
'                    , DescArgsFrom(l_sSQLStatement) _
'                    , True
'            Case Else
'                TrsRaiseError MODULENAME _
'                    , "GetUserPreferenceDefault" _
'                    , trErrdb_TRISAvailable_F _
'                    , , , _
'                    , True
'    End Select
'End Function

'//TODO handle recordsets over 1M (1048576) chars
'//moved to cdEnvironmentTrans2
Public Function SetUserPreference(ByVal v_lngEmpNum As Long, _
                                    ByVal v_strTypeCd As String, _
                                    ByVal v_rsPreference As ADODB.Recordset) As Boolean

'    On Error GoTo eh
    
    TrsRaiseError MODULENAME, "SetUserPreference", trErrObsoleteMethod, , , , True
'    Dim l_sSQLStatement As String
'    Dim ladcn As ADODB.Connection
'    Dim rc As Long
'    Dim sSerializedPreference As String, sFormatCd As String
'    Dim rsParms As ADODB.Recordset
'
'    If v_rsPreference Is Nothing Then
'        Set v_rsPreference = CreateUserPreferenceRS(v_strTypeCd)
'    End If
'
''    sSerializedPreference = SerializeRS(v_rsPreference)
''    sFormatCd = "RSX"
'    'sSerializedPreference = SerializeRS2(TRISCompressionType_eNone, ADOSerializationType_eADTG, v_rsPreference)
'    'sFormatCd = "BIN"
'    sSerializedPreference = SerializeRS2(TRISCompressionType_eGZip, ADOSerializationType_eADTG, v_rsPreference)
'    sFormatCd = "ZIP"
'
'    If DoesUserPreferenceExist(v_lngEmpNum, v_strTypeCd) Then
'        l_sSQLStatement = _
'            "UPDATE METHDBA.TTR_USER_PREF " & _
'            " SET EMP_NUM = ? " & _
'            " , TYPE_CD = ? " & _
'            " , FORMAT_CD = ? " & _
'            " , SEQ_NUM = ? " & _
'            " , SETTINGS = ? " & _
'            " , LAST_MOD_EMP_NUM = ?" & _
'            " , LAST_MOD_TIMESTAMP = CURRENT TIMESTAMP" & _
'            " WHERE EMP_NUM = ? " & _
'            " AND TYPE_CD = ? " & _
'            " AND SEQ_NUM = ?"
'        Set rsParms = GetParmsForSQL("Update", v_lngEmpNum, Left(Trim(v_strTypeCd), 3), _
'                                    sFormatCd, 0, sSerializedPreference, _
'                                    v_lngEmpNum, v_lngEmpNum, _
'                                    v_lngEmpNum, Left(Trim(v_strTypeCd), 3), 0)
'    Else
'        l_sSQLStatement = _
'            "Insert Into METHDBA.TTR_USER_PREF" & _
'            " (EMP_NUM, TYPE_CD, FORMAT_CD, SEQ_NUM, SETTINGS, CREATED_EMP_NUM, CREATED_TIMESTAMP, LAST_MOD_EMP_NUM, LAST_MOD_TIMESTAMP) " & _
'            " Values( ? " & _
'            " , ? " & _
'            " , ? " & _
'            " , ? " & _
'            " , ? " & _
'            " , ? " & _
'            " , CURRENT TIMESTAMP" & _
'            " , ? " & _
'            " , CURRENT TIMESTAMP )"
'
'        Set rsParms = GetParmsForSQL("Insert", v_lngEmpNum, Left(Trim(v_strTypeCd), 3), _
'                                    sFormatCd, 0, sSerializedPreference, _
'                                    v_lngEmpNum, v_lngEmpNum, _
'                                    0, "", 0)
'    End If
'
'    Set ladcn = OpenConnection
'    rc = NewExecuteOtherSQL(l_sSQLStatement, rsParms, ladcn)
'
'    ladcn.Close
'    Set ladcn = Nothing
'
'    SetUserPreference = True
'    GetObjectContext.SetComplete
'    Exit Function
'eh:
'    SetUserPreference = False
    'Abort Transaction
'    GetObjectContext.SetAbort

    'Call the standard module function
    'Raise the appropriate error message depending on the error number
'    Select Case (Err.Number)
'            Case trErrutOpenConnection
'                TrsRaiseError MODULENAME, "SetUserPreference"
'            Case trErrdb_ExecuteSQL_F
'                TrsRaiseError MODULENAME _
'                    , "SetUserPreference" _
'                    , trErrdb_ExecuteSQL_F _
'                    , , _
'                    , DescArgsFrom(l_sSQLStatement) _
'                    , True
'            Case Else
'                TrsRaiseError MODULENAME _
'                    , "SetUserPreference" _
'                    , trErrdb_TRISAvailable_F _
'                    , , , _
'                    , True
'    End Select
End Function

'Public Function SetUserPreference(ByVal v_lngEmpNum As Long, _
'                                    ByVal v_strTypeCd As String, _
'                                    ByVal v_rsPreference As ADODB.Recordset) As Boolean
'
'    On Error GoTo eh
'    Dim l_sSQLStatement As String
'    Dim ladcn As ADODB.Connection
'    Dim rc As Long
'    Dim sSerializedPreference As String
'    Dim rsParms As ADODB.Recordset
'    Dim rsState As ADODB.Recordset, intSeqNum As Integer
'
'    If v_rsPreference Is Nothing Then
'        Set v_rsPreference = CreateUserPreferenceRS(v_strTypeCd)
'    End If
'
'    sSerializedPreference = SerializeRS(v_rsPreference)
'
'    Set rsState = GetUserPreferenceState(v_lngEmpNum, v_strTypeCd)
'
'    '//-----break into 10k chunks
'    For intSeqNum = 0 To Int(Len(sSerializedPreference) / m_lngBLOBSize)
'        rsState.Filter = "SEQ_NUM = " & intSeqNum
'        If Not rsState.EOF Then
'            rsState.Fields("RECORD_STATUS").Value = "C"
'            rsState.Fields("EMP_NUM").Value = v_lngEmpNum
'            rsState.Fields("TYPE_CD").Value = v_strTypeCd
'            rsState.Fields("FORMAT_CD").Value = "RSX"
'            rsState.Fields("SEQ_NUM").Value = intSeqNum
''having trouble assigning values to settings column
'            rsState.Fields("SETTINGS").AppendChunk (Mid(sSerializedPreference, intSeqNum * m_lngBLOBSize + 1, m_lngBLOBSize))
'            rsState.Fields("LAST_MOD_EMP_NUM").Value = v_lngEmpNum
''            rsState.Fields("KEY_EMP_NUM").Value = v_lngEmpNum
''            rsState.Fields("KEY_TYPE_CD").Value = v_strTypeCd
''            rsState.Fields("KEY_SEQ_NUM").Value = intSeqNum
'            rsState.Update
'        Else
'            rsState.AddNew
'            rsState.Fields("RECORD_STATUS").Value = "I"
'            rsState.Fields("EMP_NUM").Value = v_lngEmpNum
'            rsState.Fields("TYPE_CD").Value = v_strTypeCd
'            rsState.Fields("FORMAT_CD").Value = "RSX"
'            rsState.Fields("SEQ_NUM").Value = intSeqNum
'            rsState.Fields("SETTINGS").AppendChunk (Mid(sSerializedPreference, intSeqNum * m_lngBLOBSize + 1, m_lngBLOBSize))
'            rsState.Fields("CREATED_EMP_NUM").Value = v_lngEmpNum
'            rsState.Fields("LAST_MOD_EMP_NUM").Value = v_lngEmpNum
'            rsState.Fields("KEY_EMP_NUM").Value = v_lngEmpNum
'            rsState.Fields("KEY_TYPE_CD").Value = v_strTypeCd
'            rsState.Fields("KEY_SEQ_NUM").Value = intSeqNum
'            rsState.Update
'        End If
'    Next intSeqNum
'
'    Set ladcn = OpenConnection
'
'    rsState.Filter = adFilterNone
'    If RecordsExist(rsState) Then
'        rsState.MoveFirst
'        Do While Not rsState.EOF
'    '        intSeqNum = CInt(rsState.Fields("SEQ_NUM").Value)
'            Select Case rsState.Fields("RECORD_STATUS").Value
'            Case "C"
'                l_sSQLStatement = _
'                    "UPDATE METHDBA.TTR_USER_PREF " & _
'                    " SET EMP_NUM = ? " & _
'                    " , TYPE_CD = ? " & _
'                    " , FORMAT_CD = ? " & _
'                    " , SEQ_NUM = ? " & _
'                    " , SETTINGS = ? " & _
'                    " , LAST_MOD_EMP_NUM = ?" & _
'                    " , LAST_MOD_TIMESTAMP = CURRENT TIMESTAMP" & _
'                    " WHERE EMP_NUM = ? " & _
'                    " AND TYPE_CD = ? " & _
'                    " AND SEQ_NUM = ?"
'                Set rsParms = GetParmsForSQL("Update", _
'                                            rsState.Fields("EMP_NUM").Value, _
'                                            Left(Trim(rsState.Fields("TYPE_CD").Value), 3), _
'                                            rsState.Fields("FORMAT_CD").Value, _
'                                            rsState.Fields("SEQ_NUM").Value, _
'                                            rsState.Fields("SETTINGS").Value, _
'                                            0, _
'                                            rsState.Fields("EMP_NUM").Value, _
'                                            rsState.Fields("KEY_EMP_NUM").Value, _
'                                            Left(Trim(rsState.Fields("KEY_TYPE_CD").Value), 3), _
'                                            rsState.Fields("KEY_SEQ_NUM").Value)
'            Case "I"
'                l_sSQLStatement = _
'                    "Insert Into METHDBA.TTR_USER_PREF" & _
'                    " (EMP_NUM, TYPE_CD, FORMAT_CD, SEQ_NUM, SETTINGS, CREATED_EMP_NUM, CREATED_TIMESTAMP, LAST_MOD_EMP_NUM, LAST_MOD_TIMESTAMP) " & _
'                    " Values( ? " & _
'                    " , ? " & _
'                    " , ? " & _
'                    " , ? " & _
'                    " , ? " & _
'                    " , ? " & _
'                    " , CURRENT TIMESTAMP" & _
'                    " , ? " & _
'                    " , CURRENT TIMESTAMP )"
'                Set rsParms = GetParmsForSQL("Insert", _
'                                            rsState.Fields("EMP_NUM").Value, _
'                                            Left(Trim(rsState.Fields("TYPE_CD").Value), 3), _
'                                            rsState.Fields("FORMAT_CD").Value, _
'                                            rsState.Fields("SEQ_NUM").Value, _
'                                            CStr(rsState.Fields("SETTINGS").Value), _
'                                            rsState.Fields("EMP_NUM").Value, _
'                                            rsState.Fields("EMP_NUM").Value, _
'                                            0, "", 0)
'            Case "D"
'                l_sSQLStatement = _
'                    "DELETE METHDBA.TTR_USER_PREF" & _
'                    " WHERE EMP_NUM = ? " & _
'                    " AND TYPE_CD = ? " & _
'                    " AND SEQ_NUM = ?"
'                Set rsParms = GetParmsForSQL("Delete", 0, "", "", 0, "", 0, 0, _
'                                            rsState.Fields("KEY_EMP_NUM").Value, _
'                                            Left(Trim(rsState.Fields("KEY_TYPE_CD").Value), 3), _
'                                            rsState.Fields("KEY_SEQ_NUM").Value)
'            End Select
'
'            rc = NewExecuteOtherSQL(l_sSQLStatement, rsParms, ladcn)
'
'            rsState.MoveNext
'        Loop
'    End If
'
'    ladcn.Close
'    Set ladcn = Nothing
'
'    GetObjectContext.SetComplete
'    Exit Function
'eh:
'    'Abort Transaction
'    GetObjectContext.SetAbort
'
'    'Call the standard module function
'    'Raise the appropriate error message depending on the error number
'    Select Case (Err.Number)
'            Case trErrutOpenConnection
'                TrsRaiseError MODULENAME, "SetUserPreference"
'            Case trErrdb_ExecuteSQL_F
'                TrsRaiseError MODULENAME _
'                    , "SetUserPreference" _
'                    , trErrdb_ExecuteSQL_F _
'                    , , _
'                    , DescArgsFrom(l_sSQLStatement) _
'                    , True
'            Case Else
'                TrsRaiseError MODULENAME _
'                    , "SetUserPreference" _
'                    , trErrdb_TRISAvailable_F _
'                    , , , _
'                    , True
'    End Select
'End Function

'Private Function PreferenceTypeCodeToString(ByVal v_eTypeCd As PreferenceTypeCode) As String
'    Select Case v_eTypeCd
'    Case UserPreference
'        PreferenceTypeCodeToString = "PRF"
'    Case JobList
'        PreferenceTypeCodeToString = "JOB"
'    Case ScheduleList
'        PreferenceTypeCodeToString = "SCH"
'    Case RotationList
'        PreferenceTypeCodeToString = "ROT"
'    Case ExceptionList
'        PreferenceTypeCodeToString = "EXC"
'    Case OTReasonList
'        PreferenceTypeCodeToString = "OTR"
'    Case WorkGroupList
'        PreferenceTypeCodeToString = "WG"
'    End Select
'End Function

Public Function DoesUserPreferenceExist(ByVal v_lngEmpNum As Long, _
                                        ByVal v_strTypeCd As String) As Boolean
                                    
    On Error GoTo eh
    
    DoesUserPreferenceExist = mdbEnvironment.DoesUserPreferenceExist(v_lngEmpNum, v_strTypeCd)
'    Dim l_sSQLStatement As String
'    Dim rs As ADODB.Recordset
'
'    'Set SQL SELECT statment
'    l_sSQLStatement = _
'        "SELECT COUNT(*)" & _
'        " FROM METHDBA.TTR_USER_PREF " & _
'        " WHERE EMP_NUM = " & v_lngEmpNum & _
'        "   AND TYPE_CD = '" & Left(Trim(v_strTypeCd), 3) & "'"
'
'    Set rs = ExecuteSQL(l_sSQLStatement, adOpenStatic, adLockReadOnly)
'
'    If RecordsExist(rs) Then
'        If CLng(rs.Fields(0).Value) > 0 Then
'            DoesUserPreferenceExist = True
'        Else
'            DoesUserPreferenceExist = False
'        End If
'    Else
'        DoesUserPreferenceExist = False
'    End If
                  
    Exit Function
eh:
'    'Abort Transaction
'    GetObjectContext.SetAbort
'
'    'Call the standard module function
'    'Raise the appropriate error message depending on the error number
'    Select Case (Err.Number)
'            Case trErrutOpenConnection
'                TrsRaiseError MODULENAME, "DoesUserPreferenceExist"
'            Case trErrdb_ExecuteSQL_F
'                TrsRaiseError MODULENAME _
'                    , "DoesUserPreferenceExist" _
'                    , trErrdb_ExecuteSQL_F _
'                    , , _
'                    , DescArgsFrom(l_sSQLStatement) _
'                    , True
'            Case Else
                TrsRaiseError MODULENAME _
                    , "DoesUserPreferenceExist" _
                    , trErrdb_TRISAvailable_F _
                    , , , _
                    , True
'    End Select
End Function


