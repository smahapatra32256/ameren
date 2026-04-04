VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "cdEnvironmentTrans2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const MODULENAME As String = "cdEnvironmentTrans2"
Private Const m_strCMODULENAME As String = MODULENAME
Private Const m_lngBLOBSize As Long = 1048576 '1024 * 1024
'Private Const m_lngSerializationChunkSize As Long = 10240
'//TODO handle recordsets over 1M (1048576) chars

Public Function SetUserPreference(ByVal v_lngEmpNum As Long, _
                                    ByVal v_strTypeCd As String, _
                                    ByVal v_rsPreference As ADODB.Recordset) As Boolean

    On Error GoTo eh
    Dim l_sSQLStatement As String
    Dim ladcn As ADODB.Connection
    Dim rc As Long
    Dim sSerializedPreference As String, sFormatCd As String
    Dim rsParms As ADODB.Recordset

    If v_rsPreference Is Nothing Then
        Set v_rsPreference = CreateUserPreferenceRS(v_strTypeCd)
    End If

'    sSerializedPreference = SerializeRS(v_rsPreference)
'    sFormatCd = "RSX"
    'sSerializedPreference = SerializeRS2(TRISCompressionType_eNone, ADOSerializationType_eADTG, v_rsPreference)
    'sFormatCd = "BIN"
    sSerializedPreference = SerializeRS2(TRISCompressionType_eGZip, ADOSerializationType_eADTG, v_rsPreference)
    sFormatCd = "ZIP"
    
    If mdbEnvironment.DoesUserPreferenceExist(v_lngEmpNum, v_strTypeCd) Then
        l_sSQLStatement = _
            "UPDATE METHDBA.TTR_USER_PREF " & _
            " SET EMP_NUM = ? " & _
            " , TYPE_CD = ? " & _
            " , FORMAT_CD = ? " & _
            " , SEQ_NUM = ? " & _
            " , SETTINGS = ? " & _
            " , LAST_MOD_EMP_NUM = ?" & _
            " , LAST_MOD_TIMESTAMP = CURRENT TIMESTAMP" & _
            " WHERE EMP_NUM = ? " & _
            " AND TYPE_CD = ? " & _
            " AND SEQ_NUM = ?"
        Set rsParms = GetParmsForSQL("Update", v_lngEmpNum, Left(Trim(v_strTypeCd), 3), _
                                    sFormatCd, 0, sSerializedPreference, _
                                    v_lngEmpNum, v_lngEmpNum, _
                                    v_lngEmpNum, Left(Trim(v_strTypeCd), 3), 0)
    Else
        l_sSQLStatement = _
            "Insert Into METHDBA.TTR_USER_PREF" & _
            " (EMP_NUM, TYPE_CD, FORMAT_CD, SEQ_NUM, SETTINGS, CREATED_EMP_NUM, CREATED_TIMESTAMP, LAST_MOD_EMP_NUM, LAST_MOD_TIMESTAMP) " & _
            " Values( ? " & _
            " , ? " & _
            " , ? " & _
            " , ? " & _
            " , ? " & _
            " , ? " & _
            " , CURRENT TIMESTAMP" & _
            " , ? " & _
            " , CURRENT TIMESTAMP )"

        Set rsParms = GetParmsForSQL("Insert", v_lngEmpNum, Left(Trim(v_strTypeCd), 3), _
                                    sFormatCd, 0, sSerializedPreference, _
                                    v_lngEmpNum, v_lngEmpNum, _
                                    0, "", 0)
    End If

    Set ladcn = OpenConnection
    rc = NewExecuteOtherSQL(l_sSQLStatement, rsParms, ladcn)

    ladcn.Close
    Set ladcn = Nothing

    SetUserPreference = True
'    GetObjectContext.SetComplete
    Exit Function
eh:
    SetUserPreference = False
    'Abort Transaction
'    GetObjectContext.SetAbort

    'Call the standard module function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME, "SetUserPreference"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "SetUserPreference" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(l_sSQLStatement) _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "SetUserPreference" _
                    , trErrdb_TRISAvailable_F _
                    , , , _
                    , True
    End Select
End Function

Private Function SerializeRS(ByVal v_rsPreference As ADODB.Recordset) As String
    Dim OutputStream As ADODB.Stream
    Dim iChunkLength As Long
    Dim sSerialized As String
    
    iChunkLength = mdbEnvironment.m_lngSerializationChunkSize
    
    Set OutputStream = CreateObject("ADODB.Stream")
'    If RecordsExist(v_rsPreference) Then
        v_rsPreference.Save OutputStream, adPersistXML
        Do While Not OutputStream.EOS
            sSerialized = sSerialized & OutputStream.ReadText(iChunkLength)
        Loop
'    End If
    
    SerializeRS = sSerialized
    
    Set OutputStream = Nothing
'    GetObjectContext.SetComplete
    Exit Function
eh:
    'Abort Transaction
'    GetObjectContext.SetAbort
    
    Set OutputStream = Nothing
    
    TrsRaiseError MODULENAME _
        , "SerializeRS" _
        , _
        , , , _
        , True
End Function

'//-----try utilizing the .NET Framework for the tuff stuff
Private Function SerializeRS2(ByVal CompressionType As TRISServerUtilityCOM.TRISCompressionType, _
                                ByVal SerializationType As TRISServerUtilityCOM.ADOSerializationType, _
                                ByVal v_rsPreference As ADODB.Recordset _
                                ) As String
    Dim oUtility As TRISServerUtilityCOM.Utility
    
    Set oUtility = CreateObject("TRISServerUtilityCOM.Utility")
    
    SerializeRS2 = oUtility.RecordsetToString(CompressionType, SerializationType, v_rsPreference)

    Set oUtility = Nothing
'    GetObjectContext.SetComplete
    Exit Function
eh:
    'Abort Transaction
'    GetObjectContext.SetAbort
    
    Set oUtility = Nothing
    
    TrsRaiseError MODULENAME _
        , "SerializeRS2" _
        , _
        , , , _
        , True
End Function

'" (EMP_NUM, TYPE_CD, FORMAT_CD, SEQ_NUM, SETTINGS, CREATED_EMP_NUM, LAST_MOD_EMP_NUM, ) "
Private Function GetParmsForSQL(ByVal v_strMode As String, _
                                ByVal v_lngEmpNum As Long, _
                                ByVal v_strTypeCd As String, _
                                ByVal v_strFormatCd As String, _
                                ByVal v_intSeqNum As Integer, _
                                ByVal v_strSettings As String, _
                                ByVal v_lngCreatedEmpNum As Long, _
                                ByVal v_lngLastModEmpNum As Long, _
                                ByVal v_lngKeyEmpNum As Long, _
                                ByVal v_strKeyTypeCd As String, _
                                ByVal v_intKeySeqNum As Integer) As ADODB.Recordset
    Dim rsParms As ADODB.Recordset
    
    Set rsParms = CreateParameterRS
    With rsParms
        If v_strMode <> "Delete" Then
            .AddNew
            !PARM_ORDER = 1
            !PARM_NAME = ""
            !PARM_TYPE = adInteger
            !PARM_DIRECTION = adParamInput
            !PARM_SIZE = 0
            !PARM_VALUE = v_lngEmpNum
            .Update
            
            .AddNew
            !PARM_ORDER = 2
            !PARM_NAME = ""
            !PARM_TYPE = adChar
            !PARM_DIRECTION = adParamInput
            !PARM_SIZE = 3
            !PARM_VALUE = v_strTypeCd
            .Update
        
            .AddNew
            !PARM_ORDER = 3
            !PARM_NAME = ""
            !PARM_TYPE = adChar
            !PARM_DIRECTION = adParamInput
            !PARM_SIZE = 3
            !PARM_VALUE = v_strFormatCd
            .Update
        
            .AddNew
            !PARM_ORDER = 4
            !PARM_NAME = ""
            !PARM_TYPE = adSmallInt
            !PARM_DIRECTION = adParamInput
            !PARM_SIZE = 0
            !PARM_VALUE = v_intSeqNum
            .Update
            
'ADODB.DataTypeEnum.adLongVarBinary works when db is 1 M and parm size is 1048576
'ADODB.DataTypeEnum.adVarBinary - error creating paramater 3421 Application uses a value of the wrong type for the current operation.
'ADODB.DataTypeEnum.adLongVarChar - error on execute SQL0301N -2147217887

            .AddNew
            !PARM_ORDER = 5
            !PARM_NAME = ""
            !PARM_TYPE = ADODB.DataTypeEnum.adLongVarBinary 'ADODB.DataTypeEnum.adLongVarBinary
            !PARM_DIRECTION = adParamInput
            !PARM_SIZE = m_lngBLOBSize  '1048576 'm_lngBLOBSize
            !PARM_VALUE = v_strSettings
            .Update
            
            If v_strMode = "Insert" Then
                .AddNew
                !PARM_ORDER = 6
                !PARM_NAME = ""
                !PARM_TYPE = adInteger
                !PARM_DIRECTION = adParamInput
                !PARM_SIZE = 0
                !PARM_VALUE = v_lngCreatedEmpNum
                .Update
            End If
            
            .AddNew
            !PARM_ORDER = 7
            !PARM_NAME = ""
            !PARM_TYPE = adInteger
            !PARM_DIRECTION = adParamInput
            !PARM_SIZE = 0
            !PARM_VALUE = v_lngLastModEmpNum
            .Update

        End If
            
        If v_strMode = "Update" Or v_strMode = "Delete" Then
            .AddNew
            !PARM_ORDER = 8
            !PARM_NAME = ""
            !PARM_TYPE = adInteger
            !PARM_DIRECTION = adParamInput
            !PARM_SIZE = 0
            !PARM_VALUE = v_lngKeyEmpNum
            .Update
            
            .AddNew
            !PARM_ORDER = 9
            !PARM_NAME = ""
            !PARM_TYPE = adChar
            !PARM_DIRECTION = adParamInput
            !PARM_SIZE = 3
            !PARM_VALUE = v_strKeyTypeCd
            .Update
            
            .AddNew
            !PARM_ORDER = 10
            !PARM_NAME = ""
            !PARM_TYPE = adSmallInt
            !PARM_DIRECTION = adParamInput
            !PARM_SIZE = 0
            !PARM_VALUE = v_intKeySeqNum
            .Update
        End If
    End With

    Set GetParmsForSQL = rsParms
End Function

