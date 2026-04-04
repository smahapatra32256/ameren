VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "cbsSchedRotationTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const mstrCMODULENAME As String = "busSchedRotationTrans"
Private Const mstrCZERO As String = "0"
Private Const mstrCVARFIELDNAMES_RTH As String = "ROT_CD"
Private Const mstrCVARFIELDNAMES_RTD As String = "ROT_CD;ROT_WEEK"
Private Const m_lngEMPNOCONVERSIONMINUSONE As Long = -1
Private Const m_lngEMPNOCONVERSIONID As Long = 99999  'conversion default last mod id
Private m_strUcaseYorN As String
Private m_lngCreatedUserNo As Long
Private m_lngUserId As Long
Private m_lngErrorNo As Long
Private m_lngRecordId As Long
Private m_intNextRotationCode As Integer
Private Enum TRSTranOption
    trDelete = 1
    trInsert = 2
    trChange = 3
End Enum



'public call to record header and detail rotation
Public Function RecordRotation(ByRef r_rsScheduleRotationRequest As ADODB.Recordset, _
                               ByRef r_rsScheduleRotationDetailRequest As ADODB.Recordset, _
                               ByVal v_lngUserId As Long, _
                               ByRef r_rsErrors As ADODB.Recordset _
                                    ) As Boolean

On Error GoTo RecordRotation_EH
    Dim dc As cuSet
    Dim l_lngIndex As Long
    Dim l_strRotCditem As String

    RecordRotation = True
    
    If Val(v_lngUserId) > 0 Then             ' is the user id valid
        If IsValidEmployee(v_lngUserId) Then
            m_lngUserId = v_lngUserId
            m_lngCreatedUserNo = v_lngUserId
        Else
            Err.Raise trErrboEmpNotFound
        End If
    Else
        If Val(v_lngUserId) = m_lngEMPNOCONVERSIONMINUSONE Then 'processing conversion only
            m_lngUserId = m_lngEMPNOCONVERSIONID 'use default user id for create and last mod emp num
            m_lngCreatedUserNo = r_rsScheduleRotationRequest!CREATED_EMP_NUM
        Else
            Err.Raise trErrdbInvalidZerolngEmpNum
        End If
    End If
    
    If RecordsExist(r_rsScheduleRotationRequest) Then
        If Not ValidateRotationHeaderEntry(r_rsScheduleRotationRequest, r_rsErrors) Then
           GetObjectContext.SetAbort
           RecordRotation = False
           Exit Function
        End If
        If Not IsBUSValidRotationHeader(r_rsScheduleRotationRequest, r_rsErrors) Then
           GetObjectContext.SetAbort
           RecordRotation = False
           Exit Function
        End If
        If HasAnyRotationHeaderfieldChanged(r_rsScheduleRotationRequest, r_rsErrors) Then
            If SaveRotationHeader(r_rsScheduleRotationRequest) = True Then
            Else
                GetObjectContext.SetAbort
                RecordRotation = False
                Exit Function
            End If
        Else
        End If
        If IsClearedFilterRecordsetEmpty(r_rsScheduleRotationRequest) Then
        End If
    End If

    If RecordsExist(r_rsScheduleRotationDetailRequest) Then
        If Not ValidateRotationDetailEntry(r_rsScheduleRotationDetailRequest, r_rsErrors) Then
           GetObjectContext.SetAbort
           RecordRotation = False
           Exit Function
        End If
        If Not IsBUSValidRotationDetail(r_rsScheduleRotationDetailRequest, r_rsErrors) Then
           GetObjectContext.SetAbort
           RecordRotation = False
           Exit Function
        End If
        If HasAnyRotationDetailfieldChanged(r_rsScheduleRotationDetailRequest, r_rsErrors) Then
            If SaveRotationDetail(r_rsScheduleRotationDetailRequest) = True Then
            Else
                GetObjectContext.SetAbort
                RecordRotation = False
                Exit Function
            End If
        Else
        End If
        If IsClearedFilterRecordsetEmpty(r_rsScheduleRotationDetailRequest) Then
        End If
    End If
    
    If Val(v_lngUserId) = m_lngEMPNOCONVERSIONMINUSONE Then 'processing conversion only - do not call spam
    Else
        Set dc = New cuSet
        If IsClearedFilterRecordsetEmpty(r_rsScheduleRotationRequest) Then
        Else
            'get distinct list then filter and sort each, with getrows
            r_rsScheduleRotationRequest.MoveFirst
            Do Until r_rsScheduleRotationRequest.EOF
                dc.Add Trim(r_rsScheduleRotationRequest!ROT_CD)
                r_rsScheduleRotationRequest.MoveNext
            Loop
            r_rsScheduleRotationRequest.MoveFirst
        End If
        If IsClearedFilterRecordsetEmpty(r_rsScheduleRotationDetailRequest) Then
        Else
            'store ROT_CD distinct list in collection component cuSet
            r_rsScheduleRotationDetailRequest.MoveFirst
            Do Until r_rsScheduleRotationDetailRequest.EOF
                dc.Add Trim(r_rsScheduleRotationDetailRequest!ROT_CD)
                r_rsScheduleRotationDetailRequest.MoveNext
            Loop
            r_rsScheduleRotationDetailRequest.MoveFirst
        End If
        For l_lngIndex = 1 To dc.Count
            l_strRotCditem = CStr(dc.Item(l_lngIndex))
            CodeTableDBMgr.InvalidateRotationCache l_strRotCditem  'mark these ROT_CD as invalid in SPAM cache
        Next
        Set dc = Nothing
    End If
    
    GetObjectContext.SetComplete
    
Exit Function

RecordRotation_EH:

    Set dc = Nothing
    RecordRotation = False
    GetObjectContext.SetAbort
    Select Case Err.Number
        Case trErrbusSchedRotDuplicateHeader
            m_lngErrorNo = trErrbusSchedRotDuplicateHeader
            Err.Clear
            m_lngRecordId = Val(r_rsScheduleRotationRequest!RECORD_ID)
            AppendRotationError r_rsErrors, TrType_RTH, m_lngRecordId, "ROT_CD", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), CStr(r_rsScheduleRotationRequest!ROT_CD)
            Exit Function
        Case trErrbusSchedRotDuplicate
            m_lngErrorNo = trErrbusSchedRotDuplicate
            Err.Clear
            'trErrbusSchedRotDuplicate "cannot create duplicate detail record"
            m_lngRecordId = Val(r_rsScheduleRotationDetailRequest!RECORD_ID)
            AppendRotationError r_rsErrors, TrType_RTD, m_lngRecordId, "ROT_WEEK", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), CStr(r_rsScheduleRotationDetailRequest!ROT_CD), CStr(r_rsScheduleRotationDetailRequest!ROT_WEEK)
            Exit Function
        Case trErrboEmpNotFound
            TrsRaiseError mstrCMODULENAME, "RecordRotation", trErrboEmpNotFound, , , , True
        Case trErrdbInvalidZerolngEmpNum
            TrsRaiseError mstrCMODULENAME, "RecordRotation", trErrdbInvalidZerolngEmpNum, , , , True
        Case Else
            TrsRaiseError mstrCMODULENAME, "RecordRotation", , , , , True
    
    End Select



End Function
'saving rotation header
Private Function SaveRotationHeader(ByVal v_rsScheduleRotationRequest As ADODB.Recordset) As Boolean

On Error GoTo SaveRotationHeader_EH

    Dim l_dbSchedRotation As dbSchedRotation.cdSchedRotation

    Set l_dbSchedRotation = CreateObject("dbSchedRotation.cdSchedRotation")

    If l_dbSchedRotation.SaveScheduleRotation(v_rsScheduleRotationRequest) Then
        SaveRotationHeader = True
    Else
        SaveRotationHeader = False
    End If

    Set l_dbSchedRotation = Nothing

Exit Function

SaveRotationHeader_EH:

    Set l_dbSchedRotation = Nothing
    
    TrsRaiseError mstrCMODULENAME, "SaveRotationHeader", , , , , True

End Function
'saving rotation detail
Private Function SaveRotationDetail(ByVal v_rsScheduleRotationDetailRequest As ADODB.Recordset) As Boolean

    Dim l_dbSchedRotation As dbSchedRotation.cdSchedRotation


On Error GoTo SaveRotationDetail_EH


    Set l_dbSchedRotation = CreateObject("dbSchedRotation.cdSchedRotation")

    If l_dbSchedRotation.SaveRotationDetail(v_rsScheduleRotationDetailRequest) Then
        SaveRotationDetail = True
    Else
        SaveRotationDetail = False
    End If

    Set l_dbSchedRotation = Nothing

Exit Function
SaveRotationDetail_EH:

    Set l_dbSchedRotation = Nothing
    
    TrsRaiseError mstrCMODULENAME, "SaveRotationDetail", , , , , True

End Function

'business rules for rotation detail
Private Function IsBUSValidRotationDetail(ByRef r_rsScheduleRotationDetailRequest As ADODB.Recordset, _
                                          ByRef r_rsErrors As ADODB.Recordset) As Boolean

On Error GoTo IsBUSValidRotationDetail_EH

    Dim l_objDBSchedDetail As ADODB.Recordset
    Dim l_objRSchedHeader As ADODB.Recordset
    Dim l_blnDoRotWeekContiguous As Boolean

    IsBUSValidRotationDetail = True
    l_blnDoRotWeekContiguous = False

    If Not IsBUSInsertDetail(r_rsScheduleRotationDetailRequest, r_rsErrors) Then
       GetObjectContext.SetAbort
       IsBUSValidRotationDetail = False
       Exit Function
    End If
        
    If Not IsBUSChangeDetail(r_rsScheduleRotationDetailRequest, r_rsErrors) Then
       GetObjectContext.SetAbort
       IsBUSValidRotationDetail = False
       Exit Function
    Else
       l_blnDoRotWeekContiguous = True
    End If
        
    If Not IsBUSDeleteDetail(r_rsScheduleRotationDetailRequest, r_rsErrors) Then
       GetObjectContext.SetAbort
       IsBUSValidRotationDetail = False
       Exit Function
    Else
       l_blnDoRotWeekContiguous = True
    End If
    
    If Not IsBUSSplitUpdateDetail(r_rsScheduleRotationDetailRequest, r_rsErrors) Then
       GetObjectContext.SetAbort
       IsBUSValidRotationDetail = False
       Exit Function
    Else
       l_blnDoRotWeekContiguous = True
    End If

    If l_blnDoRotWeekContiguous Then
        If Not IsBUSNetRotWeekContiguous(r_rsScheduleRotationDetailRequest, r_rsErrors) Then
           GetObjectContext.SetAbort
           IsBUSValidRotationDetail = False
           Exit Function
        End If
    Else
    End If

Exit Function

IsBUSValidRotationDetail_EH:

    TrsRaiseError mstrCMODULENAME, "IsBUSValidRotationDetail"

End Function

'Business rules for insert

Private Function IsBUSInsertDetail(ByRef r_rsScheduleRotationDetailRequest As ADODB.Recordset, _
                                   ByRef r_rsErrors As ADODB.Recordset) As Boolean

On Error GoTo IsBUSInsertDetail_EH

    Dim l_objDBSchedDetail As ADODB.Recordset
    Dim l_objRSchedHeader As ADODB.Recordset
    Dim l_dbSchedRotation As dbSchedRotation.cdSchedRotation
    Dim l_strFilter As String
    
    IsBUSInsertDetail = True
    
    If IsClearedFilterRecordsetEmpty(r_rsScheduleRotationDetailRequest) Then
        Exit Function
    Else
    End If
    
    l_strFilter = "RECORD_STATUS = 'I'"
    If IsRecordsInFilterRecordSet(l_strFilter, r_rsScheduleRotationDetailRequest) Then
        If m_intNextRotationCode > 0 Then
        Else
            If IsRotationCodeNotAssigned(r_rsScheduleRotationDetailRequest!ROT_CD) = True Then
            Else
        '       trErrbusSchedRotHeadAssigned = "Header record is currently assigned"
                m_lngErrorNo = trErrbusSchedRotHeadAssigned
                AppendRotationError r_rsErrors, TrType_RTH, m_lngRecordId, "ROT_CD", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), CStr(r_rsScheduleRotationDetailRequest!KEY_ROT_CD), CStr(r_rsScheduleRotationDetailRequest!ROT_WEEK), CStr(r_rsScheduleRotationDetailRequest!CREATED_TIMESTAMP)
                IsBUSInsertDetail = False
                Exit Function
            End If
        End If
        Set l_dbSchedRotation = CreateObject("dbSchedRotation.cdSchedRotation")
        Set l_objRSchedHeader = l_dbSchedRotation.GetRotationSchedHeader(r_rsScheduleRotationDetailRequest!ROT_CD)
        Set l_dbSchedRotation = Nothing
        If l_objRSchedHeader Is Nothing Then
'            trErrbusSchedRotNoHeader = "cannot process detail without a parent header"
            m_lngErrorNo = trErrbusSchedRotNoHeader
            AppendRotationError r_rsErrors, TrType_RTD, m_lngRecordId, "ROT_CD", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), CStr(r_rsScheduleRotationDetailRequest!KEY_ROT_CD), CStr(r_rsScheduleRotationDetailRequest!ROT_WEEK), CStr(r_rsScheduleRotationDetailRequest!CREATED_TIMESTAMP)
            IsBUSInsertDetail = False
            Exit Function
        Else
        End If
        Do Until r_rsScheduleRotationDetailRequest.EOF
            If r_rsScheduleRotationDetailRequest!ROT_WEEK = 0 Then
    '           trErrbusSchedRotInvalidZero = "Rotation week cannot be equal to 0"
                m_lngErrorNo = trErrbusSchedRotInvalidZero
                AppendRotationError r_rsErrors, TrType_RTD, m_lngRecordId, "ROT_WEEK", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), CStr(r_rsScheduleRotationDetailRequest!KEY_ROT_CD), CStr(r_rsScheduleRotationDetailRequest!ROT_WEEK), CStr(r_rsScheduleRotationDetailRequest!CREATED_TIMESTAMP)
                IsBUSInsertDetail = False
                Exit Function
            Else
                r_rsScheduleRotationDetailRequest!last_mod_emp_num = m_lngUserId
                r_rsScheduleRotationDetailRequest!CREATED_EMP_NUM = m_lngCreatedUserNo
                If IsValidSchedCd(r_rsScheduleRotationDetailRequest) Then
                Else
                   m_lngErrorNo = trErrbusSchedCodeInvalid
                   AppendRotationError r_rsErrors, TrType_RTD, m_lngRecordId, "SCHED_CD", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), CStr(r_rsScheduleRotationDetailRequest!KEY_ROT_CD), CStr(r_rsScheduleRotationDetailRequest!ROT_WEEK), CStr(r_rsScheduleRotationDetailRequest!CREATED_TIMESTAMP)
                   IsBUSInsertDetail = False
                   Exit Function
                End If
                r_rsScheduleRotationDetailRequest.Update
            End If
            r_rsScheduleRotationDetailRequest.MoveNext
        Loop
    Else
        Exit Function
    End If
    
    If IsClearedFilterRecordsetEmpty(r_rsScheduleRotationDetailRequest) Then
    End If
    
    Set l_objRSchedHeader = Nothing
    Set l_objDBSchedDetail = Nothing


Exit Function

IsBUSInsertDetail_EH:

    Set l_objRSchedHeader = Nothing
    Set l_objDBSchedDetail = Nothing
    Set l_dbSchedRotation = Nothing
    
    TrsRaiseError mstrCMODULENAME, "IsBUSInsertDetail"

End Function


'Business rules for change

Private Function IsBUSChangeDetail(ByRef r_rsScheduleRotationDetailRequest As ADODB.Recordset, _
                                   ByRef r_rsErrors As ADODB.Recordset) As Boolean

On Error GoTo IsBUSChangeDetail_EH

    Dim l_objDBSchedDetail As ADODB.Recordset
    Dim l_objRSchedHeader As ADODB.Recordset
    Dim l_dbSchedRotation As dbSchedRotation.cdSchedRotation
    Dim l_strFilter As String
    
    IsBUSChangeDetail = True

    If IsClearedFilterRecordsetEmpty(r_rsScheduleRotationDetailRequest) Then
        Exit Function
    Else
    End If
    
    Set l_dbSchedRotation = CreateObject("dbSchedRotation.cdSchedRotation")
    
    l_strFilter = "RECORD_STATUS = 'C'"
    If IsRecordsInFilterRecordSet(l_strFilter, r_rsScheduleRotationDetailRequest) Then
        If IsRotationCodeNotAssigned(r_rsScheduleRotationDetailRequest!KEY_ROT_CD) = True Then
        Else
'            trErrbusSchedRotHeadAssigned = "Header record is currently assigned"
            m_lngErrorNo = trErrbusSchedRotHeadAssigned
            AppendRotationError r_rsErrors, TrType_RTH, m_lngRecordId, "ROT_CD", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), CStr(r_rsScheduleRotationDetailRequest!KEY_ROT_CD)
            IsBUSChangeDetail = False
            Exit Function
        End If
        Set l_objRSchedHeader = l_dbSchedRotation.GetRotationSchedHeader(r_rsScheduleRotationDetailRequest!KEY_ROT_CD)  'used for nondemand indicator
        If l_objRSchedHeader Is Nothing Then
'            trErrbusSchedRotNoHeader = "cannot process detail without a parent header"
            m_lngErrorNo = trErrbusSchedRotNoHeader
            AppendRotationError r_rsErrors, TrType_RTD, m_lngRecordId, "ROT_CD", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), CStr(r_rsScheduleRotationDetailRequest!KEY_ROT_CD), CStr(r_rsScheduleRotationDetailRequest!KEY_ROT_WEEK)
            IsBUSChangeDetail = False
            Exit Function
        Else
        End If
        Do Until r_rsScheduleRotationDetailRequest.EOF
            r_rsScheduleRotationDetailRequest!last_mod_emp_num = m_lngUserId
            If IsValidSchedCd(r_rsScheduleRotationDetailRequest) Then
            Else
               m_lngErrorNo = trErrbusSchedCodeInvalid
               AppendRotationError r_rsErrors, TrType_RTD, m_lngRecordId, "SCHED_CD", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), CStr(r_rsScheduleRotationDetailRequest!KEY_ROT_CD), CStr(r_rsScheduleRotationDetailRequest!KEY_ROT_WEEK)
               IsBUSChangeDetail = False
               Exit Function
            End If
            Set l_objDBSchedDetail = l_dbSchedRotation.GetRotationSchedDetail(r_rsScheduleRotationDetailRequest!KEY_ROT_CD, r_rsScheduleRotationDetailRequest!KEY_ROT_WEEK)
            If l_objDBSchedDetail Is Nothing Then
        '       trErrbusSchedRotDetailNotFound = "Rotation Detail - Record not found    "
                m_lngErrorNo = trErrbusSchedRotDetailNotFound
                AppendRotationError r_rsErrors, TrType_RTD, m_lngRecordId, "ROT_CD", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), CStr(r_rsScheduleRotationDetailRequest!KEY_ROT_CD), CStr(r_rsScheduleRotationDetailRequest!KEY_ROT_WEEK)
                IsBUSChangeDetail = False
                Exit Function
            Else
            End If
            If l_objDBSchedDetail!ROT_WEEK = r_rsScheduleRotationDetailRequest!ROT_WEEK Then
            Else
                r_rsScheduleRotationDetailRequest!RECORD_STATUS = "U" 'build insert and delete records
                'build insert with new key r_rsScheduleRotationDetailRequest!KEY_ROT_WEEK
                'build delete with old key l_objDBSchedDetail!ROT_WEEK
            End If
            r_rsScheduleRotationDetailRequest.Update
            r_rsScheduleRotationDetailRequest.MoveNext
        Loop
    Else
        Exit Function
    End If
    
    If IsClearedFilterRecordsetEmpty(r_rsScheduleRotationDetailRequest) Then
    End If
    
    Set l_objRSchedHeader = Nothing
    Set l_objDBSchedDetail = Nothing
    Set l_dbSchedRotation = Nothing

Exit Function

IsBUSChangeDetail_EH:

    Set l_objRSchedHeader = Nothing
    Set l_objDBSchedDetail = Nothing
    Set l_dbSchedRotation = Nothing

    TrsRaiseError mstrCMODULENAME, "IsBUSChangeDetail"

End Function


'Business rules for delete

Private Function IsBUSDeleteDetail(ByRef r_rsScheduleRotationDetailRequest As ADODB.Recordset, _
                                   ByRef r_rsErrors As ADODB.Recordset) As Boolean

On Error GoTo IsBUSDeleteDetail_EH

    Dim l_strFilter As String
    
    IsBUSDeleteDetail = True
    
    If IsClearedFilterRecordsetEmpty(r_rsScheduleRotationDetailRequest) Then
        Exit Function
    Else
    End If
    
    l_strFilter = "RECORD_STATUS = 'D'"
    If IsRecordsInFilterRecordSet(l_strFilter, r_rsScheduleRotationDetailRequest) Then
        Do Until r_rsScheduleRotationDetailRequest.EOF
            If IsRotationCodeNotAssigned(r_rsScheduleRotationDetailRequest!KEY_ROT_CD) = True Then
            Else
    '            trErrbusSchedRotHeadAssigned = "Header record is currently assigned"
                m_lngErrorNo = trErrbusSchedRotHeadAssigned
                AppendRotationError r_rsErrors, TrType_RTH, m_lngRecordId, "ROT_CD", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), CStr(r_rsScheduleRotationDetailRequest!KEY_ROT_CD)
                IsBUSDeleteDetail = False
                Exit Function
            End If
            r_rsScheduleRotationDetailRequest.MoveNext
        Loop
    Else
        Exit Function
    End If
    
    If IsClearedFilterRecordsetEmpty(r_rsScheduleRotationDetailRequest) Then
    End If
    

Exit Function

IsBUSDeleteDetail_EH:

    TrsRaiseError mstrCMODULENAME, "IsBUSDeleteDetail"

End Function



'Business rules for change

Private Function IsBUSSplitUpdateDetail(ByRef r_rsScheduleRotationDetailRequest As ADODB.Recordset, _
                                        ByRef r_rsErrors As ADODB.Recordset) As Boolean

On Error GoTo IsBUSSplitUpdateDetail_EH

    Dim l_intKeyRotCode As Integer
    Dim l_intKeyRotWeek As Integer
    Dim l_intRotCode As Integer
    Dim l_intRotWeek As Integer
    Dim l_intSchedCode As Integer
    Dim l_intRecordId As Integer
    Dim l_strCreatedTimestamp As String
    Dim l_strFilter As String

    IsBUSSplitUpdateDetail = True
    
    If IsClearedFilterRecordsetEmpty(r_rsScheduleRotationDetailRequest) Then
        Exit Function
    Else
    End If
    
    l_strFilter = "RECORD_STATUS = 'U'"
    If IsRecordsInFilterRecordSet(l_strFilter, r_rsScheduleRotationDetailRequest) Then
        'convert "U" KEY records to delete and adds for detail change of Rot code and Rot week keys
        Do Until r_rsScheduleRotationDetailRequest.EOF
            l_intKeyRotCode = Val(r_rsScheduleRotationDetailRequest!KEY_ROT_CD)
            l_intKeyRotWeek = Val(r_rsScheduleRotationDetailRequest!KEY_ROT_WEEK)
            l_intRotCode = Val(r_rsScheduleRotationDetailRequest!ROT_CD)
            l_intRotWeek = Val(r_rsScheduleRotationDetailRequest!ROT_WEEK)
            l_intSchedCode = Val(r_rsScheduleRotationDetailRequest!SCHED_CD)
            l_intRecordId = Val(r_rsScheduleRotationDetailRequest!RECORD_ID)
            l_strCreatedTimestamp = r_rsScheduleRotationDetailRequest!CREATED_TIMESTAMP
            r_rsScheduleRotationDetailRequest!RECORD_ID = l_intRecordId
            r_rsScheduleRotationDetailRequest!KEY_ROT_CD = l_intKeyRotCode
            r_rsScheduleRotationDetailRequest!KEY_ROT_WEEK = l_intKeyRotWeek
            r_rsScheduleRotationDetailRequest!RECORD_STATUS = "D" 'build insert and delete records
            r_rsScheduleRotationDetailRequest.Update
           
            r_rsScheduleRotationDetailRequest.AddNew
            r_rsScheduleRotationDetailRequest!RECORD_ID = l_intRecordId
            r_rsScheduleRotationDetailRequest!ROT_CD = l_intRotCode
            r_rsScheduleRotationDetailRequest!ROT_WEEK = l_intRotWeek
            r_rsScheduleRotationDetailRequest!SCHED_CD = l_intSchedCode
            r_rsScheduleRotationDetailRequest!last_mod_emp_num = m_lngUserId
            r_rsScheduleRotationDetailRequest!CREATED_EMP_NUM = m_lngCreatedUserNo
            r_rsScheduleRotationDetailRequest!CREATED_TIMESTAMP = l_strCreatedTimestamp
            r_rsScheduleRotationDetailRequest!RECORD_STATUS = "I" 'build insert and delete records
            r_rsScheduleRotationDetailRequest.Update
            
            If r_rsScheduleRotationDetailRequest.EOF _
            Or r_rsScheduleRotationDetailRequest.BOF Then 'since the filter result is changed to "D" and "I" from "C" or modified "U"
            Else
                r_rsScheduleRotationDetailRequest.MoveFirst
            End If
        Loop
    Else
        Exit Function
    End If
    
    
    If IsClearedFilterRecordsetEmpty(r_rsScheduleRotationDetailRequest) Then
    End If
    

Exit Function

IsBUSSplitUpdateDetail_EH:

    TrsRaiseError mstrCMODULENAME, "IsBUSSplitUpdateDetail"

End Function




'Business rules for inserts, change and deletes
Private Function IsBUSValidRotationHeader(ByRef r_rsScheduleRotationHeaderRequest As ADODB.Recordset, _
                                          ByRef r_rsErrors As ADODB.Recordset) As Boolean

On Error GoTo IsBUSValidRotationHeader_EH
    
    IsBUSValidRotationHeader = True
    
    If Not IsBUSInsertHeader(r_rsScheduleRotationHeaderRequest, r_rsErrors) Then
       GetObjectContext.SetAbort
       IsBUSValidRotationHeader = False
       Exit Function
    End If

    If Not IsBUSChangeHeader(r_rsScheduleRotationHeaderRequest, r_rsErrors) Then
       GetObjectContext.SetAbort
       IsBUSValidRotationHeader = False
       Exit Function
    End If

    If Not IsBUSDeleteHeader(r_rsScheduleRotationHeaderRequest, r_rsErrors) Then
       GetObjectContext.SetAbort
       IsBUSValidRotationHeader = False
       Exit Function
    End If

Exit Function

IsBUSValidRotationHeader_EH:
    
    TrsRaiseError mstrCMODULENAME, "IsBUSValidRotationHeader"

End Function

'Business rules for insert
Private Function IsBUSInsertHeader(ByRef r_rsScheduleRotationHeaderRequest As ADODB.Recordset, _
                                   ByRef r_rsErrors As ADODB.Recordset) As Boolean

On Error GoTo IsBUSInsertHeader_EH
    
    Dim l_objDBchedHeader As ADODB.Recordset
    Dim l_strFilter As String

    IsBUSInsertHeader = True
    
    If IsClearedFilterRecordsetEmpty(r_rsScheduleRotationHeaderRequest) Then
        Exit Function
    Else
    End If
    
    l_strFilter = "RECORD_STATUS = 'I'"
    If IsRecordsInFilterRecordSet(l_strFilter, r_rsScheduleRotationHeaderRequest) Then
        Do Until r_rsScheduleRotationHeaderRequest.EOF
            r_rsScheduleRotationHeaderRequest!CREATED_EMP_NUM = m_lngCreatedUserNo
            r_rsScheduleRotationHeaderRequest!last_mod_emp_num = m_lngUserId
            r_rsScheduleRotationHeaderRequest.Update
            r_rsScheduleRotationHeaderRequest.MoveNext
        Loop
    Else
        Exit Function
    End If
    

    If IsClearedFilterRecordsetEmpty(r_rsScheduleRotationHeaderRequest) Then
    End If
    Set l_objDBchedHeader = Nothing


Exit Function
IsBUSInsertHeader_EH:
    Set l_objDBchedHeader = Nothing
    TrsRaiseError mstrCMODULENAME, "IsBUSInsertHeader"

End Function

'Business rules for change
Private Function IsBUSChangeHeader(ByRef r_rsScheduleRotationHeaderRequest As ADODB.Recordset, _
                                   ByRef r_rsErrors As ADODB.Recordset) As Boolean

On Error GoTo IsBUSChangeHeader_EH
    
    Dim l_objDBchedHeader As ADODB.Recordset
    Dim l_dbSchedRotation As dbSchedRotation.cdSchedRotation
    Dim l_strFilter As String

    Set l_dbSchedRotation = CreateObject("dbSchedRotation.cdSchedRotation")

    IsBUSChangeHeader = True
    
    If IsClearedFilterRecordsetEmpty(r_rsScheduleRotationHeaderRequest) Then
        Exit Function
    Else
    End If
    
    l_strFilter = "RECORD_STATUS = 'C'"
    If IsRecordsInFilterRecordSet(l_strFilter, r_rsScheduleRotationHeaderRequest) Then
        Do Until r_rsScheduleRotationHeaderRequest.EOF
            If IsRotationCodeNotAssigned(r_rsScheduleRotationHeaderRequest!KEY_ROT_CD) = True Then
            Else
    '            trErrbusSchedRotHeadAssigned = "Header record is currently assigned"
                m_lngErrorNo = trErrbusSchedRotHeadAssigned
                AppendRotationError r_rsErrors, TrType_RTH, m_lngRecordId, "ROT_CD", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), CStr(r_rsScheduleRotationHeaderRequest!KEY_ROT_CD)
                IsBUSChangeHeader = False
                Exit Function
            End If
            Set l_objDBchedHeader = l_dbSchedRotation.GetRotationSchedHeader(r_rsScheduleRotationHeaderRequest!KEY_ROT_CD)
            If l_objDBchedHeader Is Nothing Then
    '            trErrbusSchedRotHeaderNotFound = "Rotation Header - Record does not exist -    "
                m_lngErrorNo = trErrbusSchedRotHeaderNotFound
                AppendRotationError r_rsErrors, TrType_RTH, m_lngRecordId, "ROT_CD", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), CStr(r_rsScheduleRotationHeaderRequest!ROT_CD)
                IsBUSChangeHeader = False
                Exit Function
            Else
            End If
            
            r_rsScheduleRotationHeaderRequest!last_mod_emp_num = m_lngUserId
            r_rsScheduleRotationHeaderRequest.Update
            r_rsScheduleRotationHeaderRequest.MoveNext
        Loop
    Else
        Exit Function
    End If
    

    If IsClearedFilterRecordsetEmpty(r_rsScheduleRotationHeaderRequest) Then
    End If
    Set l_objDBchedHeader = Nothing
    Set l_dbSchedRotation = Nothing

Exit Function
IsBUSChangeHeader_EH:
    Set l_objDBchedHeader = Nothing
    Set l_dbSchedRotation = Nothing
    TrsRaiseError mstrCMODULENAME, "IsBUSChangeHeader"

End Function


'Business rules for delete
Private Function IsBUSDeleteHeader(ByRef r_rsScheduleRotationHeaderRequest As ADODB.Recordset, _
                                   ByRef r_rsErrors As ADODB.Recordset) As Boolean

On Error GoTo IsBUSDeleteHeader_EH
    
    Dim l_objDBchedHeader As ADODB.Recordset
    Dim l_strFilter As String

    IsBUSDeleteHeader = True
    
    If IsClearedFilterRecordsetEmpty(r_rsScheduleRotationHeaderRequest) Then
        Exit Function
    Else
    End If
    
    
    l_strFilter = "RECORD_STATUS = 'D'"
    If IsRecordsInFilterRecordSet(l_strFilter, r_rsScheduleRotationHeaderRequest) Then
        Do Until r_rsScheduleRotationHeaderRequest.EOF
            If IsRotationCodeNotAssigned(r_rsScheduleRotationHeaderRequest!KEY_ROT_CD) = True Then
            Else
    '            trErrbusSchedRotHeadAssigned = "Header record is currently assigned"
                m_lngErrorNo = trErrbusSchedRotHeadAssigned
                AppendRotationError r_rsErrors, TrType_RTH, m_lngRecordId, "ROT_CD", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), CStr(r_rsScheduleRotationHeaderRequest!KEY_ROT_CD)
                IsBUSDeleteHeader = False
                Exit Function
            End If
            r_rsScheduleRotationHeaderRequest.MoveNext
        Loop
    Else
        Exit Function
    End If
    
    
    If IsClearedFilterRecordsetEmpty(r_rsScheduleRotationHeaderRequest) Then
    End If
    Set l_objDBchedHeader = Nothing


Exit Function
IsBUSDeleteHeader_EH:
    Set l_objDBchedHeader = Nothing
    TrsRaiseError mstrCMODULENAME, "IsBUSDeleteHeader"

End Function


'prevalidate rotation header
Private Function ValidateRotationHeaderEntry( _
                    ByRef r_rsScheduleRotationHeaderRequest As ADODB.Recordset, _
                    ByRef r_rsErrors As ADODB.Recordset _
                    ) As Boolean
    
On Error GoTo eh
    
    Dim l_blnValidationError As Boolean
    
    If IsClearedFilterRecordsetEmpty(r_rsScheduleRotationHeaderRequest) Then
        ValidateRotationHeaderEntry = True
        Exit Function
    Else
    End If
    

    Do Until r_rsScheduleRotationHeaderRequest.EOF
        If Not IsNull(r_rsScheduleRotationHeaderRequest!RECORD_STATUS) Then
            If Not IsNull(r_rsScheduleRotationHeaderRequest!RECORD_ID) Then 'store record_id for append use+
                m_lngRecordId = Val(r_rsScheduleRotationHeaderRequest!RECORD_ID)
            Else
                m_lngRecordId = 0
            End If
            If Trim(r_rsScheduleRotationHeaderRequest!RECORD_STATUS) = "D" Then  'delete header
                If Not IsNull(r_rsScheduleRotationHeaderRequest!KEY_ROT_CD) Then
                    If Val(r_rsScheduleRotationHeaderRequest!KEY_ROT_CD) > 0 _
                        And IsNumeric(r_rsScheduleRotationHeaderRequest!KEY_ROT_CD) Then
                    Else
                        'trErrbusSchedRotCodeIsZero
                        m_lngErrorNo = trErrbusSchedRotCodeIsZero
                        AppendRotationError r_rsErrors, TrType_RTH, m_lngRecordId, "ROT_CD", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), CStr(r_rsScheduleRotationHeaderRequest!KEY_ROT_CD)
                        l_blnValidationError = True
                    End If
                Else
                    'trErrbusSchedRotCodeNull
                    m_lngErrorNo = trErrbusSchedRotCodeNull
                    AppendRotationError r_rsErrors, TrType_RTH, m_lngRecordId, "ROT_CD", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), CStr(r_rsScheduleRotationHeaderRequest!KEY_ROT_CD)
                    l_blnValidationError = True
                End If
            Else
                If IsNull(r_rsScheduleRotationHeaderRequest!ROT_DESC) Then  'applies for inserts and changes
                    'trErrbusSchedRotDescNull
                    m_lngErrorNo = trErrbusSchedRotDescNull
                    AppendRotationError r_rsErrors, TrType_RTH, m_lngRecordId, "ROT_DESC", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), CStr(r_rsScheduleRotationHeaderRequest!KEY_ROT_CD)
                    l_blnValidationError = True
                Else
                    If Trim(r_rsScheduleRotationHeaderRequest!ROT_DESC) = "" Then
                        'trErrbusSchedRotDescBlanks
                        m_lngErrorNo = trErrbusSchedRotDescBlanks
                        AppendRotationError r_rsErrors, TrType_RTH, m_lngRecordId, "ROT_DESC", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), CStr(r_rsScheduleRotationHeaderRequest!KEY_ROT_CD)
                        l_blnValidationError = True
                    Else
                        Dim l_strRotDesc As String
                        l_strRotDesc = Trim(r_rsScheduleRotationHeaderRequest!ROT_DESC)
                        r_rsScheduleRotationHeaderRequest!ROT_DESC = Left(Replace(l_strRotDesc, "'", "''"), 50)
                        r_rsScheduleRotationHeaderRequest.Update
                    End If
                End If
                If Trim(r_rsScheduleRotationHeaderRequest!RECORD_STATUS) = "C" Then   'change header
                    If Not IsNull(r_rsScheduleRotationHeaderRequest!KEY_ROT_CD) Then
                        If Val(r_rsScheduleRotationHeaderRequest!KEY_ROT_CD) > 0 _
                            And IsNumeric(r_rsScheduleRotationHeaderRequest!KEY_ROT_CD) Then
                        Else
                            'trErrbusSchedRotCodeIsZero
                            m_lngErrorNo = trErrbusSchedRotCodeIsZero
                            AppendRotationError r_rsErrors, TrType_RTH, m_lngRecordId, "ROT_CD", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), CStr(r_rsScheduleRotationHeaderRequest!KEY_ROT_CD)
                            l_blnValidationError = True
                        End If
                    Else
                        'trErrbusSchedRotCodeNull
                        m_lngErrorNo = trErrbusSchedRotCodeNull
                        AppendRotationError r_rsErrors, TrType_RTH, m_lngRecordId, "ROT_CD", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), CStr(r_rsScheduleRotationHeaderRequest!KEY_ROT_CD)
                        l_blnValidationError = True
                    End If
                Else ' INSERT Header
                    If l_blnValidationError = True Then
                    Else  'assigned next rotation code
                        If Not IsNull(r_rsScheduleRotationHeaderRequest!ROT_CD) Then
                            If Val(r_rsScheduleRotationHeaderRequest!ROT_CD) < 1 Then
                                m_intNextRotationCode = GetNextRotationCode
                                r_rsScheduleRotationHeaderRequest!ROT_CD = m_intNextRotationCode
                                r_rsScheduleRotationHeaderRequest.Update
                            End If
                        Else
                            m_intNextRotationCode = GetNextRotationCode
                            r_rsScheduleRotationHeaderRequest!ROT_CD = m_intNextRotationCode
                            r_rsScheduleRotationHeaderRequest.Update
                        End If
                    End If
                End If
            End If
        Else
            'trErrbusSchedRotHeadRecStatNull
            m_lngErrorNo = trErrbusSchedRotHeadRecStatNull
            AppendRotationError r_rsErrors, TrType_RTH, m_lngRecordId, "RECORD_STATUS", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), CStr(r_rsScheduleRotationHeaderRequest!ROT_CD)
            l_blnValidationError = True
        End If
        r_rsScheduleRotationHeaderRequest.MoveNext
    Loop

    If IsClearedFilterRecordsetEmpty(r_rsScheduleRotationHeaderRequest) Then
    End If

    ValidateRotationHeaderEntry = Not l_blnValidationError

Exit Function

eh:
    TrsRaiseError mstrCMODULENAME, "ValidateRotationHeaderEntry"

End Function


'prevalidate rotation detail
Private Function ValidateRotationDetailEntry( _
                    ByRef r_rsScheduleRotationDetailRequest As ADODB.Recordset, _
                    ByRef r_rsErrors As ADODB.Recordset _
                    ) As Boolean
    
On Error GoTo eh

    Dim l_blnValidationError As Boolean

    If IsClearedFilterRecordsetEmpty(r_rsScheduleRotationDetailRequest) Then
        ValidateRotationDetailEntry = True
        Exit Function
    Else
    End If

    Do Until r_rsScheduleRotationDetailRequest.EOF
        If Not IsNull(r_rsScheduleRotationDetailRequest!RECORD_STATUS) Then
            If Not IsNull(r_rsScheduleRotationDetailRequest!RECORD_ID) Then
                m_lngRecordId = Val(r_rsScheduleRotationDetailRequest!RECORD_ID)
            Else
                m_lngRecordId = 0
            End If
            If Trim(r_rsScheduleRotationDetailRequest!RECORD_STATUS) = "D" Then
                If Not IsNull(r_rsScheduleRotationDetailRequest!KEY_ROT_CD) Then
                    If Val(r_rsScheduleRotationDetailRequest!KEY_ROT_CD) > 0 _
                        And IsNumeric(r_rsScheduleRotationDetailRequest!KEY_ROT_CD) Then
                    Else
                        ' TrsGetErrDescription(LongAppErr(trErrbusSchedRotCodeInvalid))
                        m_lngErrorNo = trErrbusSchedRotCodeInvalid
                        AppendRotationError r_rsErrors, TrType_RTD, m_lngRecordId, "ROT_CD", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), CStr(r_rsScheduleRotationDetailRequest!KEY_ROT_CD), CStr(r_rsScheduleRotationDetailRequest!KEY_ROT_WEEK)
                        l_blnValidationError = True
                    End If
                Else
                    ' TrsGetErrDescription(LongAppErr(trErrbusSchedRotDetailCodeNull))
                    m_lngErrorNo = trErrbusSchedRotDetailCodeNull
                    AppendRotationError r_rsErrors, TrType_RTD, m_lngRecordId, "ROT_CD", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), CStr(r_rsScheduleRotationDetailRequest!KEY_ROT_CD), CStr(r_rsScheduleRotationDetailRequest!KEY_ROT_WEEK)
                    l_blnValidationError = True
                End If
                If Not IsNull(r_rsScheduleRotationDetailRequest!KEY_ROT_WEEK) Then
                    If Val(r_rsScheduleRotationDetailRequest!KEY_ROT_WEEK) > 0 _
                        And IsNumeric(r_rsScheduleRotationDetailRequest!KEY_ROT_WEEK) Then
                    Else
                        ' TrsGetErrDescription(LongAppErr(trErrbusSchedRotOrderInvalid))
                        m_lngErrorNo = trErrbusSchedRotOrderInvalid
                        AppendRotationError r_rsErrors, TrType_RTD, m_lngRecordId, "ROT_WEEK", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), CStr(r_rsScheduleRotationDetailRequest!KEY_ROT_CD), CStr(r_rsScheduleRotationDetailRequest!KEY_ROT_WEEK)
                        l_blnValidationError = True
                    End If
                Else
                    ' TrsGetErrDescription(LongAppErr(trErrbusSchedRotOrderNull))
                    m_lngErrorNo = trErrbusSchedRotOrderNull
                    AppendRotationError r_rsErrors, TrType_RTD, m_lngRecordId, "ROT_WEEK", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), CStr(r_rsScheduleRotationDetailRequest!KEY_ROT_CD), CStr(r_rsScheduleRotationDetailRequest!KEY_ROT_WEEK)
                    l_blnValidationError = True
                End If
            Else
                If Trim(r_rsScheduleRotationDetailRequest!RECORD_STATUS) = "C" Then
                    If Not IsNull(r_rsScheduleRotationDetailRequest!KEY_ROT_CD) Then
                        If Val(r_rsScheduleRotationDetailRequest!KEY_ROT_CD) > 0 _
                            And IsNumeric(r_rsScheduleRotationDetailRequest!KEY_ROT_CD) Then
                        Else
                            ' TrsGetErrDescription(LongAppErr(trErrbusSchedRotCodeInvalid))
                            m_lngErrorNo = trErrbusSchedRotCodeInvalid
                            AppendRotationError r_rsErrors, TrType_RTD, m_lngRecordId, "ROT_CD", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), CStr(r_rsScheduleRotationDetailRequest!KEY_ROT_CD), CStr(r_rsScheduleRotationDetailRequest!KEY_ROT_WEEK)
                            l_blnValidationError = True
                        End If
                    Else
                        ' TrsGetErrDescription(LongAppErr(trErrbusSchedRotDetailCodeNull))
                        m_lngErrorNo = trErrbusSchedRotDetailCodeNull
                        AppendRotationError r_rsErrors, TrType_RTD, m_lngRecordId, "ROT_CD", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), CStr(r_rsScheduleRotationDetailRequest!KEY_ROT_CD), CStr(r_rsScheduleRotationDetailRequest!KEY_ROT_WEEK)
                        l_blnValidationError = True
                    End If
                    If Not IsNull(r_rsScheduleRotationDetailRequest!KEY_ROT_WEEK) Then
                        If Val(r_rsScheduleRotationDetailRequest!KEY_ROT_WEEK) > 0 _
                            And IsNumeric(r_rsScheduleRotationDetailRequest!KEY_ROT_WEEK) Then
                        Else
                            ' TrsGetErrDescription(LongAppErr(trErrbusSchedRotOrderInvalid))
                            m_lngErrorNo = trErrbusSchedRotOrderInvalid
                            AppendRotationError r_rsErrors, TrType_RTD, m_lngRecordId, "ROT_WEEK", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), CStr(r_rsScheduleRotationDetailRequest!KEY_ROT_CD), CStr(r_rsScheduleRotationDetailRequest!KEY_ROT_WEEK)
                            l_blnValidationError = True
                        End If
                    Else
                        ' TrsGetErrDescription(LongAppErr(trErrbusSchedRotOrderNull))
                        m_lngErrorNo = trErrbusSchedRotOrderNull
                        AppendRotationError r_rsErrors, TrType_RTD, m_lngRecordId, "ROT_WEEK", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), CStr(r_rsScheduleRotationDetailRequest!KEY_ROT_CD), CStr(r_rsScheduleRotationDetailRequest!KEY_ROT_WEEK)
                        l_blnValidationError = True
                    End If
                    
                    If Not IsNull(r_rsScheduleRotationDetailRequest!ROT_WEEK) Then
                        If Val(r_rsScheduleRotationDetailRequest!ROT_WEEK) > 0 _
                            And IsNumeric(r_rsScheduleRotationDetailRequest!ROT_WEEK) Then
                        Else
                            ' TrsGetErrDescription(LongAppErr(trErrbusSchedRotOrderInvalid))
                            m_lngErrorNo = trErrbusSchedRotOrderInvalid
                            AppendRotationError r_rsErrors, TrType_RTD, m_lngRecordId, "ROT_WEEK", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), CStr(r_rsScheduleRotationDetailRequest!KEY_ROT_CD), CStr(r_rsScheduleRotationDetailRequest!KEY_ROT_WEEK)
                            l_blnValidationError = True
                        End If
                    Else
                        ' TrsGetErrDescription(LongAppErr(trErrbusSchedRotOrderNull))
                        m_lngErrorNo = trErrbusSchedRotOrderNull
                        AppendRotationError r_rsErrors, TrType_RTD, m_lngRecordId, "ROT_WEEK", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), CStr(r_rsScheduleRotationDetailRequest!KEY_ROT_CD), CStr(r_rsScheduleRotationDetailRequest!KEY_ROT_WEEK)
                        l_blnValidationError = True
                    End If
                    If Not IsNull(r_rsScheduleRotationDetailRequest!SCHED_CD) Then
                        If Val(r_rsScheduleRotationDetailRequest!SCHED_CD) > 0 _
                            And IsNumeric(r_rsScheduleRotationDetailRequest!SCHED_CD) Then
                        Else
                            ' TrsGetErrDescription(LongAppErr(trErrbusSchedCodeIsZero))
                            m_lngErrorNo = trErrbusSchedCodeIsZero
                            AppendRotationError r_rsErrors, TrType_RTD, m_lngRecordId, "SCHED_CD", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), CStr(r_rsScheduleRotationDetailRequest!KEY_ROT_CD), CStr(r_rsScheduleRotationDetailRequest!KEY_ROT_WEEK)
                            l_blnValidationError = True
                        End If
                    Else
                        ' TrsGetErrDescription(LongAppErr(trErrbusSchedCodeNull))
                        m_lngErrorNo = trErrbusSchedCodeNull
                        AppendRotationError r_rsErrors, TrType_RTD, m_lngRecordId, "SCHED_CD", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), CStr(r_rsScheduleRotationDetailRequest!KEY_ROT_CD), CStr(r_rsScheduleRotationDetailRequest!KEY_ROT_WEEK)
                        l_blnValidationError = True
                    End If
                Else
                    If Not IsNull(r_rsScheduleRotationDetailRequest!ROT_CD) Then
                        If Val(r_rsScheduleRotationDetailRequest!ROT_CD) < 1 Then
                           If m_intNextRotationCode > 0 Then
                              r_rsScheduleRotationDetailRequest!ROT_CD = m_intNextRotationCode
                              r_rsScheduleRotationDetailRequest.Update
                           Else
                                ' TrsGetErrDescription(LongAppErr(trErrbusSchedRotCodeInvalid))
                                m_lngErrorNo = trErrbusSchedRotCodeInvalid
                                AppendRotationError r_rsErrors, TrType_RTD, m_lngRecordId, "ROT_CD", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), CStr(r_rsScheduleRotationDetailRequest!ROT_CD), CStr(r_rsScheduleRotationDetailRequest!ROT_WEEK)
                                l_blnValidationError = True
                           End If
                        End If
                    Else
                        If m_intNextRotationCode > 0 Then
                              r_rsScheduleRotationDetailRequest!ROT_CD = m_intNextRotationCode
                              r_rsScheduleRotationDetailRequest.Update
                        Else
                            ' TrsGetErrDescription(LongAppErr(trErrbusSchedRotDetailCodeNull))
                            m_lngErrorNo = trErrbusSchedRotDetailCodeNull
                            AppendRotationError r_rsErrors, TrType_RTD, m_lngRecordId, "ROT_CD", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), CStr(r_rsScheduleRotationDetailRequest!ROT_CD), CStr(r_rsScheduleRotationDetailRequest!ROT_WEEK)
                            l_blnValidationError = True
                        End If
                    End If
                    If Not IsNull(r_rsScheduleRotationDetailRequest!ROT_WEEK) Then
                        If Val(r_rsScheduleRotationDetailRequest!ROT_WEEK) > 0 _
                            And IsNumeric(r_rsScheduleRotationDetailRequest!ROT_WEEK) Then
                        Else
                            ' TrsGetErrDescription(LongAppErr(trErrbusSchedRotOrderInvalid))
                            m_lngErrorNo = trErrbusSchedRotOrderInvalid
                            AppendRotationError r_rsErrors, TrType_RTD, m_lngRecordId, "ROT_WEEK", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), CStr(r_rsScheduleRotationDetailRequest!ROT_CD), CStr(r_rsScheduleRotationDetailRequest!ROT_WEEK)
                            l_blnValidationError = True
                        End If
                    Else
                        ' TrsGetErrDescription(LongAppErr(trErrbusSchedRotOrderNull))
                        m_lngErrorNo = trErrbusSchedRotOrderNull
                        AppendRotationError r_rsErrors, TrType_RTD, m_lngRecordId, "ROT_WEEK", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), CStr(r_rsScheduleRotationDetailRequest!ROT_CD), CStr(r_rsScheduleRotationDetailRequest!ROT_WEEK)
                        l_blnValidationError = True
                    End If
                    If Not IsNull(r_rsScheduleRotationDetailRequest!SCHED_CD) Then
                        If Val(r_rsScheduleRotationDetailRequest!SCHED_CD) > 0 _
                            And IsNumeric(r_rsScheduleRotationDetailRequest!SCHED_CD) Then
                        Else
                            ' TrsGetErrDescription(LongAppErr(trErrbusSchedCodeIsZero))
                            m_lngErrorNo = trErrbusSchedCodeIsZero
                            AppendRotationError r_rsErrors, TrType_RTD, m_lngRecordId, "SCHED_CD", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), CStr(r_rsScheduleRotationDetailRequest!ROT_CD), CStr(r_rsScheduleRotationDetailRequest!ROT_WEEK)
                            l_blnValidationError = True
                        End If
                    Else
                        ' TrsGetErrDescription(LongAppErr(trErrbusSchedCodeNull))
                        m_lngErrorNo = trErrbusSchedCodeNull
                        AppendRotationError r_rsErrors, TrType_RTD, m_lngRecordId, "SCHED_CD", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), CStr(r_rsScheduleRotationDetailRequest!ROT_CD), CStr(r_rsScheduleRotationDetailRequest!ROT_WEEK)
                        l_blnValidationError = True
                    End If
                End If
            End If
        Else
            ' TrsGetErrDescription(LongAppErr(trErrbusSchedRotRecStatusNull))
            m_lngErrorNo = trErrbusSchedRotRecStatusNull
            AppendRotationError r_rsErrors, TrType_RTD, m_lngRecordId, "RECORD_STATUS", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), CStr(r_rsScheduleRotationDetailRequest!KEY_ROT_CD), CStr(r_rsScheduleRotationDetailRequest!KEY_ROT_WEEK)
            l_blnValidationError = True
        End If
        r_rsScheduleRotationDetailRequest.MoveNext
    Loop

    If IsClearedFilterRecordsetEmpty(r_rsScheduleRotationDetailRequest) Then
    End If
    
    ValidateRotationDetailEntry = Not l_blnValidationError

Exit Function

eh:
    TrsRaiseError mstrCMODULENAME, "ValidateRotationDetailEntry"

End Function
'check for existing rotation assignments
Private Function IsRotationCodeNotAssigned _
                (ByVal v_intRotationCode As Integer _
                ) As Boolean

On Error GoTo IsRotationCodeNotAssigned_EH

    Dim l_dbSchedRotation As dbSchedRotation.cdSchedRotation

    Set l_dbSchedRotation = CreateObject("dbSchedRotation.cdSchedRotation")

    IsRotationCodeNotAssigned = l_dbSchedRotation.IsRotationCodeNotAssigned(v_intRotationCode)

    Set l_dbSchedRotation = Nothing
    GetObjectContext.SetComplete

Exit Function

IsRotationCodeNotAssigned_EH:
    Set l_dbSchedRotation = Nothing
    GetObjectContext.SetAbort
    TrsRaiseError mstrCMODULENAME, "IsRotationCodeNotAssigned"

End Function

''''Comment out for future use
''''Public call to save for rotation assignment
'''Public Function RecordRotationAsgnmt(ByVal v_rsRotationAsgnmt As ADODB.Recordset) As Boolean
'''
'''
'''On Error GoTo RecordRotationAsgnmt_EH
'''
'''    RecordRotationAsgnmt = True
'''
'''    If RecordsExist(v_rsRotationAsgnmt) Then
'''       If SaveRotationAsgnmt(v_rsRotationAsgnmt) = True Then
'''       Else
'''           GetObjectContext.SetAbort
'''           RecordRotationAsgnmt = False
'''           Exit Function
'''       End If
'''    End If
'''
'''    GetObjectContext.SetComplete
'''
'''Exit Function
'''
'''RecordRotationAsgnmt_EH:
'''
'''    RecordRotationAsgnmt = False
'''    GetObjectContext.SetAbort
'''    Select Case Err.Number
'''        Case trErrbusSchedRotAsgnmtDup
'''            RecordRotationAsgnmt = True 'duplicate error on inserts are ignored
'''            Err.Clear
'''            Exit Function
'''        Case Else
'''            TrsRaiseError mstrCMODULENAME, "RecordRotationAsgnmt", , , , , True
'''
'''    End Select
'''
'''
'''
'''End Function
'Comment out for future use
''''call data service rotation assignment save
'''Private Function SaveRotationAsgnmt(ByVal v_rsScheduleRotationAsgnmtRequest As ADODB.Recordset) As Boolean
'''
'''
'''On Error GoTo SaveRotationAsgnmt_EH
'''
'''    Dim l_dbSchedRotationAsgnmt As dbSchedRotation.cdSchedRotation
'''
'''    Set l_dbSchedRotationAsgnmt = CreateObject("dbSchedRotation.cdSchedRotation")
'''
'''    If l_dbSchedRotationAsgnmt.SaveRotationAsgnmt(v_rsScheduleRotationAsgnmtRequest) Then
'''        SaveRotationAsgnmt = True
'''    Else
'''        SaveRotationAsgnmt = False
'''    End If
'''
'''    Set l_dbSchedRotationAsgnmt = Nothing
'''
'''Exit Function
'''SaveRotationAsgnmt_EH:
'''
'''    Set l_dbSchedRotationAsgnmt = Nothing
'''
'''    TrsRaiseError mstrCMODULENAME, "SaveRotationAsgnmt", , , , , True
'''
'''End Function
'''
'''
'Comment out for future use
''''create rotation assignment recordset use for save of rotation assignment
'''Public Function CreateNewRSRotationAsgnmt() As ADODB.Recordset
''''
'''
'''On Error GoTo eh
'''
'''    Dim l_rsRotationAsgnmt As ADODB.Recordset
'''
'''    Set l_rsRotationAsgnmt = New ADODB.Recordset
'''    With l_rsRotationAsgnmt
'''        .Fields.Append "ROT_CD", adInteger, , adFldUpdatable
'''        .Fields.Append "LEVEL_NUM", adInteger, , adFldUpdatable
'''        .Fields.Append "EMP_NUM", adInteger, , adFldUpdatable
'''        .Fields.Append "WG_NUM", adInteger, , adFldUpdatable
'''        .Fields.Append "TIME_STAMP", adChar, 26, adFldUpdatable
'''        .Fields.Append "LAST_MOD_EMP_NUM", adInteger, , adFldUpdatable
'''        .Fields.Append "KEY_ROT_CD", adInteger, , adFldUpdatable
'''        .Fields.Append "KEY_LEVEL_NUM", adInteger, , adFldUpdatable
'''        .Fields.Append "KEY_EMP_NUM", adInteger, , adFldUpdatable
'''        .Fields.Append "KEY_WG_NUM", adInteger, , adFldUpdatable
'''        .Fields.Append "RECORD_STATUS", adChar, 1, adFldUpdatable
'''        .Open
'''    End With
'''
'''    Set CreateNewRSRotationAsgnmt = l_rsRotationAsgnmt
'''Exit Function
'''
'''eh:
'''    TrsRaiseError mstrCMODULENAME, "CreateNewRSRotationAsgnmt", , , , , True
'''
'''End Function


'validate for correct schedule code
Private Function IsValidSchedCd(ByRef r_rsRotationRequest As ADODB.Recordset) As Boolean

On Error GoTo IsValidSchedCd_EH
    
    If CodeTableDBMgr.ValidateScheduleCode(Trim(r_rsRotationRequest("SCHED_CD"))) Then
       IsValidSchedCd = True
    Else
       IsValidSchedCd = False
    End If


Exit Function

IsValidSchedCd_EH:

    TrsRaiseError mstrCMODULENAME, "IsValidSchedCd"

End Function

''create interface with buscodetables
'
'Private Function CodeTableDataManager() As busCodeTables.cbsCodeTables
'
'On Error GoTo eh
'
'    Set CodeTableDataManager = _
'        CreateObject("busCodeTables.cbsCodeTables")
'
'Exit Function
'
'eh:
'    TrsRaiseError mstrCMODULENAME, "CodeTableDataManager"
'
'End Function


'create interface with dbcodetables

Private Function CodeTableDBMgr() As dbCodeTables.cdCodeTables
    
On Error GoTo eh
    
    Set CodeTableDBMgr = _
        CreateObject("dbCodeTables.cdCodeTables")
    
Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "CodeTableDBMgr"

End Function

'checks for valid employee number
Private Function IsValidEmployee _
                (ByVal v_lngEmployee As Long _
                ) As Boolean

On Error GoTo IsValidEmployee_EH

    Dim l_dbEmployee As busEmployee.cbsEmployee

    Set l_dbEmployee = CreateObject("busEmployee.cbsEmployee")

    IsValidEmployee = l_dbEmployee.IsValidEmployee(v_lngEmployee)
        
    Set l_dbEmployee = Nothing
    
    GetObjectContext.SetComplete

Exit Function

IsValidEmployee_EH:
    Set l_dbEmployee = Nothing
    GetObjectContext.SetAbort
    TrsRaiseError mstrCMODULENAME, "IsValidEmployee"

End Function

'provide the next available rotation code for the new header record
Private Function GetNextRotationCode() As Integer
    
On Error GoTo GetNextRotationCode_EH

    Dim l_dbSchedRotation As dbSchedRotation.cdSchedRotation

    Set l_dbSchedRotation = CreateObject("dbSchedRotation.cdSchedRotation")

    GetNextRotationCode = l_dbSchedRotation.GetNextRotationCode

    Set l_dbSchedRotation = Nothing
    GetObjectContext.SetComplete

Exit Function

GetNextRotationCode_EH:
    Set l_dbSchedRotation = Nothing
    GetObjectContext.SetAbort
    TrsRaiseError mstrCMODULENAME, "GetNextRotationCode"

End Function

'use to detect any needed updates
Private Function HasAnyRotationDetailfieldChanged( _
    ByVal v_rsScheduleRotationDetailRequest As ADODB.Recordset, _
    ByRef r_rsErrors As ADODB.Recordset) As Boolean
        
On Error GoTo HasAnyRotationDetailfieldChanged_EH

    Dim l_objDBSchedDetail As ADODB.Recordset
    Dim l_dbSchedRotation As dbSchedRotation.cdSchedRotation
    Dim l_strFilter As String
    
    If IsClearedFilterRecordsetEmpty(v_rsScheduleRotationDetailRequest) Then
        HasAnyRotationDetailfieldChanged = False
        Exit Function
    Else
    End If
    
    
    l_strFilter = "RECORD_STATUS = 'I'" & _
              " OR RECORD_STATUS = 'D'"

    If IsRecordsInFilterRecordSet(l_strFilter, v_rsScheduleRotationDetailRequest) Then
        HasAnyRotationDetailfieldChanged = True
        Exit Function
    Else
    End If
    
    
    Set l_dbSchedRotation = CreateObject("dbSchedRotation.cdSchedRotation")

    l_strFilter = "RECORD_STATUS = 'C'"
    If IsRecordsInFilterRecordSet(l_strFilter, v_rsScheduleRotationDetailRequest) Then
        Do Until v_rsScheduleRotationDetailRequest.EOF
            Set l_objDBSchedDetail = l_dbSchedRotation.GetRotationSchedDetail(v_rsScheduleRotationDetailRequest!KEY_ROT_CD, v_rsScheduleRotationDetailRequest!KEY_ROT_WEEK)
            If l_objDBSchedDetail Is Nothing Then
        '       trErrbusSchedRotDetailNotFound = "Rotation Detail - Record does not exist    "
                m_lngErrorNo = trErrbusSchedRotDetailNotFound
                AppendRotationError r_rsErrors, TrType_RTD, m_lngRecordId, "ROT_CD", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), CStr(v_rsScheduleRotationDetailRequest!KEY_ROT_CD), CStr(v_rsScheduleRotationDetailRequest!KEY_ROT_WEEK)
                HasAnyRotationDetailfieldChanged = False
                Exit Function
            Else
            End If
            If l_objDBSchedDetail!ROT_WEEK = v_rsScheduleRotationDetailRequest!ROT_WEEK Then
            Else
                HasAnyRotationDetailfieldChanged = True
                Exit Function
            End If
            If l_objDBSchedDetail!SCHED_CD = v_rsScheduleRotationDetailRequest!SCHED_CD Then
            Else
                HasAnyRotationDetailfieldChanged = True
                Exit Function
            End If
            v_rsScheduleRotationDetailRequest.MoveNext
        Loop
    Else
        Exit Function
    End If
    
    
    If IsClearedFilterRecordsetEmpty(v_rsScheduleRotationDetailRequest) Then
    End If
    
    Set l_dbSchedRotation = Nothing
    
Exit Function

HasAnyRotationDetailfieldChanged_EH:

    Set l_objDBSchedDetail = Nothing
    Set l_dbSchedRotation = Nothing
    TrsRaiseError mstrCMODULENAME, "HasAnyRotationDetailfieldChanged"

End Function


'use to detect any needed updates
Private Function HasAnyRotationHeaderfieldChanged( _
    ByVal v_rsScheduleRotationHeaderRequest As ADODB.Recordset, _
    ByRef r_rsErrors As ADODB.Recordset) As Boolean
        
On Error GoTo HasAnyRotationHeaderfieldChanged_EH

    Dim l_objDBchedHeader As ADODB.Recordset
    Dim l_dbSchedRotation As dbSchedRotation.cdSchedRotation
    Dim l_strFilter As String
    
    If IsClearedFilterRecordsetEmpty(v_rsScheduleRotationHeaderRequest) Then
        HasAnyRotationHeaderfieldChanged = False
        Exit Function
    Else
    End If
    

    l_strFilter = "RECORD_STATUS = 'I'" & _
              " OR RECORD_STATUS = 'D'"

    If IsRecordsInFilterRecordSet(l_strFilter, v_rsScheduleRotationHeaderRequest) Then
        HasAnyRotationHeaderfieldChanged = True
        Exit Function
    Else
    End If
    
    
    Set l_dbSchedRotation = CreateObject("dbSchedRotation.cdSchedRotation")
        
    
    l_strFilter = "RECORD_STATUS = 'C'"
    If IsRecordsInFilterRecordSet(l_strFilter, v_rsScheduleRotationHeaderRequest) Then
        Do Until v_rsScheduleRotationHeaderRequest.EOF
            Set l_objDBchedHeader = l_dbSchedRotation.GetRotationSchedHeader(v_rsScheduleRotationHeaderRequest!KEY_ROT_CD)
            If l_objDBchedHeader Is Nothing Then
    '            l_strRotationError = "Rotation Header - Record does not exist -    "
                m_lngErrorNo = trErrbusSchedRotHeaderNotFound
                AppendRotationError r_rsErrors, TrType_RTH, m_lngRecordId, "ROT_CD", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), CStr(v_rsScheduleRotationHeaderRequest!ROT_CD)
                HasAnyRotationHeaderfieldChanged = False
                Exit Function
            Else
            End If
            If Trim(l_objDBchedHeader!ROT_DESC) = Trim(v_rsScheduleRotationHeaderRequest!ROT_DESC) Then
            Else
                HasAnyRotationHeaderfieldChanged = True
                Exit Function
            End If
            v_rsScheduleRotationHeaderRequest.MoveNext
        Loop
    Else
        Exit Function
    End If
    
    
    If IsClearedFilterRecordsetEmpty(v_rsScheduleRotationHeaderRequest) Then
    End If
    
    Set l_dbSchedRotation = Nothing
    
Exit Function

HasAnyRotationHeaderfieldChanged_EH:
    Set l_objDBchedHeader = Nothing
    Set l_dbSchedRotation = Nothing
    
    TrsRaiseError mstrCMODULENAME, "HasAnyRotationHeaderfieldChanged"

End Function
'
'use for building BRVs
Private Sub AppendRotationError( _
                ByRef r_rsErrors As ADODB.Recordset, _
                ByVal v_intErrorType As TrTypeBRV, _
                ByVal v_lngRecordId As Long, _
                ByVal v_strErrorField As String, _
                ByVal v_strErrorMsg As String, _
                Optional ByVal v_strKeyField1 As String, _
                Optional ByVal v_strKeyField2 As String, _
                Optional ByVal v_strKeyField3 As String, _
                Optional ByVal v_strKeyField4 As String, _
                Optional ByVal v_strKeyField5 As String, _
                Optional ByVal v_strKeyField6 As String, _
                Optional ByVal v_strKeyField7 As String, _
                Optional ByVal v_strKeyField8 As String, _
                Optional ByVal v_strKeyField9 As String, _
                Optional ByVal v_strKeyField10 As String)

On Error GoTo AppendRotationError_EH
    
    If v_intErrorType = TrType_RTH Then
         AppendBusinessError r_rsErrors, LongAppErr(m_lngErrorNo), mstrCBUSINESSERR, v_strErrorMsg, mstrCMODULENAME, v_intErrorType, v_lngRecordId, v_strErrorField, mstrCVARFIELDNAMES_RTH, v_strKeyField1, mstrCZERO, v_strKeyField3, v_strKeyField4, v_strKeyField5, v_strKeyField6, v_strKeyField7, v_strKeyField8, v_strKeyField9, v_strKeyField10
    Else
        If v_intErrorType = TrType_RTD Then
            AppendBusinessError r_rsErrors, LongAppErr(m_lngErrorNo), mstrCBUSINESSERR, v_strErrorMsg, mstrCMODULENAME, v_intErrorType, v_lngRecordId, v_strErrorField, mstrCVARFIELDNAMES_RTD, v_strKeyField1, v_strKeyField2, v_strKeyField3, v_strKeyField4, v_strKeyField5, v_strKeyField6, v_strKeyField7, v_strKeyField8, v_strKeyField9, v_strKeyField10
        End If
    End If

Exit Sub

AppendRotationError_EH:
    GetObjectContext.SetAbort
    TrsRaiseError mstrCMODULENAME, "AppendRotationError"

End Sub





Public Function IsRotDetPatternExists( _
                    ByRef r_rsScheduleRotationRequest As ADODB.Recordset, _
                    ByRef r_rsScheduleRotationDetailRequest As ADODB.Recordset _
                ) As Boolean

On Error GoTo IsRotDetPatternExists_EH
    
    Dim l_rsRotDetPattern As ADODB.Recordset
    Dim l_lngRotCd As Long

    Dim l_dbSchedRotation As dbSchedRotation.cdSchedRotation
    Dim l_strFilter As String

    If IsClearedFilterRecordsetEmpty(r_rsScheduleRotationRequest) Then
    End If
    
    
    If IsClearedFilterRecordsetEmpty(r_rsScheduleRotationDetailRequest) Then
    End If
    
    l_strFilter = "RECORD_STATUS = 'I'"
    
    If IsRecordsInFilterRecordSet(l_strFilter, r_rsScheduleRotationDetailRequest) Then
    Else
        Exit Function
    End If
    
    If IsRecordsInFilterRecordSet(l_strFilter, r_rsScheduleRotationRequest) Then
    Else
        Exit Function
    End If
    
    Set l_dbSchedRotation = CreateObject("dbSchedRotation.cdSchedRotation")
    Set l_rsRotDetPattern = l_dbSchedRotation.GetRotDetPatterns(r_rsScheduleRotationDetailRequest)
    If RecordsExist(l_rsRotDetPattern) Then
        l_lngRotCd = l_rsRotDetPattern!ROT_CD
        r_rsScheduleRotationRequest.MoveFirst
        r_rsScheduleRotationDetailRequest.MoveFirst
        Do Until r_rsScheduleRotationRequest.EOF
            r_rsScheduleRotationRequest!ROT_CD = l_lngRotCd
            r_rsScheduleRotationRequest.Update
            r_rsScheduleRotationRequest.MoveNext
        Loop
        Do Until r_rsScheduleRotationDetailRequest.EOF
            r_rsScheduleRotationDetailRequest!ROT_CD = l_lngRotCd
            r_rsScheduleRotationDetailRequest.Update
            r_rsScheduleRotationDetailRequest.MoveNext
        Loop
        IsRotDetPatternExists = True
    Else
        IsRotDetPatternExists = False
    End If
    Set l_dbSchedRotation = Nothing
    
    If IsClearedFilterRecordsetEmpty(r_rsScheduleRotationRequest) Then
    End If
   
    If IsClearedFilterRecordsetEmpty(r_rsScheduleRotationDetailRequest) Then
    End If
   
    
    GetObjectContext.SetComplete

Exit Function

IsRotDetPatternExists_EH:

    Set l_dbSchedRotation = Nothing
    GetObjectContext.SetAbort
    TrsRaiseError mstrCMODULENAME, "IsRotDetPatternExists"

End Function


'    contiguous construction


'Business rules for change
Private Function IsBUSNetRotWeekContiguous(ByRef r_rsScheduleRotationDetailRequest As ADODB.Recordset, _
                                        ByRef r_rsErrors As ADODB.Recordset) As Boolean

On Error GoTo IsBUSNetRotWeekContiguous_EH

    Dim l_intContiguousRotWeek As Integer
    Dim l_intKeyRotCode As Integer
    Dim l_intKeyRotWeek As Integer
    Dim l_intRotCode As Integer
    Dim l_intRotWeek As Integer
    Dim l_intSchedCode As Integer
    Dim l_intRecordId As Integer
    Dim l_objDBSchedDetail As ADODB.Recordset
    Dim l_dbSchedRotation As dbSchedRotation.cdSchedRotation
    Dim l_strFilter As String

    IsBUSNetRotWeekContiguous = True
   
    l_strFilter = "RECORD_STATUS = 'I'" & _
              " OR RECORD_STATUS = 'D'"  'there can only be "I" or "D" by the previous split process for key changed

    If IsRecordsInFilterRecordSet(l_strFilter, r_rsScheduleRotationDetailRequest) Then
    Else
        Exit Function
    End If
    
    'get all the ROT_WEEK detail records for ROT_CODE from Database
    Set l_dbSchedRotation = CreateObject("dbSchedRotation.cdSchedRotation")
    If r_rsScheduleRotationDetailRequest!RECORD_STATUS = "I" Then 'if new get rot code else get key
        Set l_objDBSchedDetail = l_dbSchedRotation.GetRotationSchedDetail(r_rsScheduleRotationDetailRequest!ROT_CD)
    Else
        Set l_objDBSchedDetail = l_dbSchedRotation.GetRotationSchedDetail(r_rsScheduleRotationDetailRequest!KEY_ROT_CD)
    End If
    If RecordsExist(l_objDBSchedDetail) Then
    Else
        Set l_objDBSchedDetail = CreateDetailRecordSet
    End If
    
    'Update or add to db recordset with "D" deleted requests and "C" changed request ("C" field changes without key changed)
    'any delete requests in the db file fot rot_week and schedule code?
    l_strFilter = "RECORD_STATUS = 'C'" & _
              " OR RECORD_STATUS = 'D'"
        
    If IsRecordsInFilterRecordSet(l_strFilter, r_rsScheduleRotationDetailRequest) Then
        Do Until r_rsScheduleRotationDetailRequest.EOF
            l_strFilter = "ROT_WEEK = '" & r_rsScheduleRotationDetailRequest!KEY_ROT_WEEK & "'" & " AND RECORD_STATUS = ''"
            If IsRecordsInFilterRecordSet(l_strFilter, l_objDBSchedDetail) Then 'db matched req keys
               l_objDBSchedDetail!RECORD_ID = r_rsScheduleRotationDetailRequest!RECORD_ID
               l_objDBSchedDetail!KEY_ROT_CD = l_objDBSchedDetail!ROT_CD
               l_objDBSchedDetail!KEY_ROT_WEEK = l_objDBSchedDetail!ROT_WEEK
               l_objDBSchedDetail!ROT_CD = r_rsScheduleRotationDetailRequest!ROT_CD
               l_objDBSchedDetail!ROT_WEEK = r_rsScheduleRotationDetailRequest!ROT_WEEK
               l_objDBSchedDetail!SCHED_CD = r_rsScheduleRotationDetailRequest!SCHED_CD
               l_objDBSchedDetail!last_mod_emp_num = r_rsScheduleRotationDetailRequest!last_mod_emp_num
               l_objDBSchedDetail!CREATED_EMP_NUM = r_rsScheduleRotationDetailRequest!CREATED_EMP_NUM
               l_objDBSchedDetail!RECORD_STATUS = r_rsScheduleRotationDetailRequest!RECORD_STATUS
               l_objDBSchedDetail.Update
            Else
            End If
            r_rsScheduleRotationDetailRequest.MoveNext
        Loop
    Else
    End If
    
    
    'Update or add to db recordset with "I" insert requests
    l_strFilter = "RECORD_STATUS = 'I'"
    If IsRecordsInFilterRecordSet(l_strFilter, r_rsScheduleRotationDetailRequest) Then
        Do Until r_rsScheduleRotationDetailRequest.EOF
            l_objDBSchedDetail.AddNew
            l_objDBSchedDetail!RECORD_ID = r_rsScheduleRotationDetailRequest!RECORD_ID
            l_objDBSchedDetail!ROT_CD = r_rsScheduleRotationDetailRequest!ROT_CD
            l_objDBSchedDetail!ROT_WEEK = r_rsScheduleRotationDetailRequest!ROT_WEEK
            l_objDBSchedDetail!SCHED_CD = r_rsScheduleRotationDetailRequest!SCHED_CD
            l_objDBSchedDetail!last_mod_emp_num = r_rsScheduleRotationDetailRequest!last_mod_emp_num
            l_objDBSchedDetail!CREATED_EMP_NUM = r_rsScheduleRotationDetailRequest!CREATED_EMP_NUM
            l_objDBSchedDetail!CREATED_TIMESTAMP = r_rsScheduleRotationDetailRequest!CREATED_TIMESTAMP
            l_objDBSchedDetail!RECORD_STATUS = r_rsScheduleRotationDetailRequest!RECORD_STATUS 'build insert and delete records
            l_objDBSchedDetail.Update
            r_rsScheduleRotationDetailRequest.MoveNext
        Loop
    Else
    End If
    
    If IsBUSSORTWeekContiguous(l_objDBSchedDetail, r_rsErrors) Then
        Set r_rsScheduleRotationDetailRequest = l_objDBSchedDetail
    Else
        GetObjectContext.SetAbort
        IsBUSNetRotWeekContiguous = False
        Exit Function
    End If

    If IsClearedFilterRecordsetEmpty(r_rsScheduleRotationDetailRequest) Then
    End If
    
    Set l_objDBSchedDetail = Nothing

Exit Function

IsBUSNetRotWeekContiguous_EH:
    
    Set l_objDBSchedDetail = Nothing

    TrsRaiseError mstrCMODULENAME, "IsBUSNetRotWeekContiguous"

End Function


'Business rules for change

Private Function IsBUSSORTWeekContiguous(ByRef r_objDBSchedDetail As ADODB.Recordset, _
                                         ByRef r_rsErrors As ADODB.Recordset _
                                        ) As Boolean

On Error GoTo IsBUSSORTWeekContiguous_EH

    Dim l_intContiguousRotWeek As Integer
    Dim l_intKeyRotCode As Integer
    Dim l_intKeyRotWeek As Integer
    Dim l_intRotCode As Integer
    Dim l_intRotWeek As Integer
    Dim l_intSchedCode As Integer
    Dim l_intRecordId As Integer
    Dim l_dbSchedRotation As dbSchedRotation.cdSchedRotation
    Dim l_strFilter As String
    Dim l_DBSchedDetail_CLONE As ADODB.Recordset
    Dim l_intBRWRotCode As Integer
    Dim l_intBRWRotWeek As Integer

    IsBUSSORTWeekContiguous = True
    
    If IsClearedFilterRecordsetEmpty(r_objDBSchedDetail) Then
        Exit Function
    Else
        r_objDBSchedDetail.Sort = "ROT_WEEK ASC" 'sort file
    End If
    
    
    l_strFilter = "RECORD_STATUS <> 'D'"
    If IsRecordsInFilterRecordSet(l_strFilter, r_objDBSchedDetail) Then
    Else
        Exit Function
    End If

    
    l_intContiguousRotWeek = 1

    Do Until r_objDBSchedDetail.EOF
        If l_intContiguousRotWeek = r_objDBSchedDetail!ROT_WEEK Then
        Else
            Exit Do
        End If
        l_intContiguousRotWeek = l_intContiguousRotWeek + 1
        r_objDBSchedDetail.MoveNext
    Loop
    'insert and blanks are here resequence them all
    Do Until r_objDBSchedDetail.EOF
        If r_objDBSchedDetail!RECORD_STATUS = "" Then
            r_objDBSchedDetail!RECORD_STATUS = "U" 'build insert and delete records
            r_objDBSchedDetail!KEY_ROT_CD = r_objDBSchedDetail!ROT_CD 'build insert and delete records
            r_objDBSchedDetail!KEY_ROT_WEEK = r_objDBSchedDetail!ROT_WEEK 'build insert and delete records
            r_objDBSchedDetail.Update
        Else
        End If
        r_objDBSchedDetail.MoveNext
    Loop

    If Not IsBUSSplitUpdateDetail(r_objDBSchedDetail, r_rsErrors) Then
       GetObjectContext.SetAbort
       IsBUSSORTWeekContiguous = False
       Exit Function
    End If
    
'check for duplicates
'

    Set l_DBSchedDetail_CLONE = r_objDBSchedDetail.Clone
    If IsClearedFilterRecordsetEmpty(l_DBSchedDetail_CLONE) Then
    End If
    
    l_strFilter = "RECORD_STATUS = 'I'"
    If IsRecordsInFilterRecordSet(l_strFilter, l_DBSchedDetail_CLONE) Then
        Do Until l_DBSchedDetail_CLONE.EOF
            l_strFilter = "ROT_WEEK = '" & l_DBSchedDetail_CLONE!ROT_WEEK & "' AND RECORD_STATUS = '" & l_DBSchedDetail_CLONE!RECORD_STATUS & "'"
            If IsRecordsInFilterRecordSet(l_strFilter, r_objDBSchedDetail) Then 'db matched req keys
                If r_objDBSchedDetail.RecordCount > 1 Then
                    'duplicate error exist user brv here
                    m_lngErrorNo = trErrbusSchedRotDuplicate
                    'trErrbusSchedRotDuplicate "cannot create duplicate detail record"
                    m_lngRecordId = Val(l_DBSchedDetail_CLONE!RECORD_ID)
                    AppendRotationError r_rsErrors, TrType_RTD, m_lngRecordId, "ROT_WEEK", TrsGetErrDescription(LongAppErr(m_lngErrorNo)), CStr(l_DBSchedDetail_CLONE!ROT_CD), CStr(l_DBSchedDetail_CLONE!ROT_WEEK), CStr(l_DBSchedDetail_CLONE!CREATED_TIMESTAMP)
                    IsBUSSORTWeekContiguous = False
                    Exit Function
                Else
                End If
            Else
            End If
            l_DBSchedDetail_CLONE.MoveNext
        Loop
    Else
    End If
    
   
'resequence all
    If IsClearedFilterRecordsetEmpty(r_objDBSchedDetail) Then
    End If
    l_strFilter = "ROT_WEEK > '" & l_intContiguousRotWeek & "' AND RECORD_STATUS = 'I'"
    If IsRecordsInFilterRecordSet(l_strFilter, r_objDBSchedDetail) Then 'db matched req keys
        l_intBRWRotCode = r_objDBSchedDetail!ROT_CD
        l_intBRWRotWeek = l_intContiguousRotWeek
        m_lngRecordId = Val(r_objDBSchedDetail!RECORD_ID)
        Do Until r_objDBSchedDetail.EOF
            r_objDBSchedDetail!ROT_WEEK = l_intContiguousRotWeek 'build insert and delete records
            l_intContiguousRotWeek = l_intContiguousRotWeek + 1
            r_objDBSchedDetail.Update
            r_objDBSchedDetail.MoveNext
        Loop
        m_lngErrorNo = trWarRotWeekResequenced
        AppendBusinessError r_rsErrors, LongAppErr(m_lngErrorNo), mstrCBUSINESSWAR, TrsGetErrDescription(LongAppErr(m_lngErrorNo)), mstrCMODULENAME, TrType_RTD, m_lngRecordId, "ROT_WEEK", mstrCVARFIELDNAMES_RTD, l_intBRWRotCode, l_intBRWRotWeek
    Else
    End If
    
    If IsClearedFilterRecordsetEmpty(r_objDBSchedDetail) Then
    End If
    

Exit Function

IsBUSSORTWeekContiguous_EH:

    TrsRaiseError mstrCMODULENAME, "IsBUSSORTWeekContiguous"

End Function


Private Function CreateDetailRecordSet() As ADODB.Recordset
On Error GoTo CreateDetailRecordSet_EH

'Const l_lngCFIELDATTR = adFldUpdatable Or adFldIsNullable
Const l_lngCFIELDATTR = adFldUpdatable

    'Create Recordset
    Dim l_adrsRotDetailSaveEntries As ADODB.Recordset

    Set l_adrsRotDetailSaveEntries = New ADODB.Recordset

    With l_adrsRotDetailSaveEntries
        .Fields.Append "ROT_CD", adInteger, , l_lngCFIELDATTR
        .Fields.Append "ROT_WEEK", adInteger, , l_lngCFIELDATTR
        .Fields.Append "SCHED_CD", adInteger, , l_lngCFIELDATTR
        .Fields.Append "LAST_MOD_EMP_NUM", adInteger, , l_lngCFIELDATTR
        .Fields.Append "CREATED_EMP_NUM", adInteger, , l_lngCFIELDATTR
        .Fields.Append "CREATED_TIMESTAMP", adChar, 26, l_lngCFIELDATTR
        .Fields.Append "LAST_MOD_TIMESTAMP", adChar, 26, l_lngCFIELDATTR
        .Fields.Append "RECORD_STATUS", adChar, 1, l_lngCFIELDATTR
        .Fields.Append "RECORD_ID", adInteger, , l_lngCFIELDATTR
        .Fields.Append "KEY_ROT_CD", adInteger, , l_lngCFIELDATTR
        .Fields.Append "KEY_ROT_WEEK", adInteger, , l_lngCFIELDATTR
        .Open

    End With
    
    Set CreateDetailRecordSet = l_adrsRotDetailSaveEntries


    Exit Function

CreateDetailRecordSet_EH:
    Err.Raise Err.Number, "CreateDetailRecordSet", Err.Description

End Function


Public Function MultiRecordRotation(ByRef r_rsScheduleRotationRequest As ADODB.Recordset, _
                               ByRef r_rsScheduleRotationDetailRequest As ADODB.Recordset, _
                               ByVal v_lngUserId As Long, _
                               ByRef r_rsErrors As ADODB.Recordset _
                                    ) As Boolean

On Error GoTo MultiRecordRotation_EH
    Dim l_rs As ADODB.Recordset
    Dim bln As Boolean
    Dim l_err As String
    Dim l_rsRequestHeader As ADODB.Recordset
    Dim l_rsRequestDetail As ADODB.Recordset
    Dim l_rsErrors As ADODB.Recordset
    
    Dim l_lngSAMEROTCD As Long
    Dim l_strSAMETimeStamp As String
    Dim l_strFilter As String
    
    MultiRecordRotation = True
    
    If Val(v_lngUserId) > 0 Then             ' is the user id valid
        If IsValidEmployee(v_lngUserId) Then
            m_lngUserId = v_lngUserId
            m_lngCreatedUserNo = v_lngUserId
        Else
            Err.Raise trErrboEmpNotFound
        End If
    Else
        If Val(v_lngUserId) = m_lngEMPNOCONVERSIONMINUSONE Then 'processing conversion only
            m_lngUserId = m_lngEMPNOCONVERSIONID 'use default user id for create and last mod emp num
            m_lngCreatedUserNo = r_rsScheduleRotationRequest!CREATED_EMP_NUM
        Else
            Err.Raise trErrdbInvalidZerolngEmpNum
        End If
    End If

   

    Do Until HeaderAndDetailEOF(r_rsScheduleRotationRequest, r_rsScheduleRotationDetailRequest, l_lngSAMEROTCD, l_strSAMETimeStamp) 'r_rsScheduleRotationRequest.EOF And r_rsScheduleRotationDetailRequest.EOF
        If l_lngSAMEROTCD = 0 Then
            l_strFilter = "ROT_CD = " & l_lngSAMEROTCD & _
                          " AND CREATED_TIMESTAMP = " & l_strSAMETimeStamp & _
                          " AND RECORD_STATUS = 'I'"         ' for add processing
        Else
            l_strFilter = "KEY_ROT_CD = " & l_lngSAMEROTCD   ' for Modify processing
        End If
        Set l_rsRequestHeader = Nothing
        Set l_rsRequestDetail = Nothing
        If IsRecordsInFilterRecordSet(l_strFilter, r_rsScheduleRotationRequest) Then '
            Set l_rsRequestHeader = CopyRecordset(r_rsScheduleRotationRequest)
            r_rsScheduleRotationRequest.MoveFirst
            Do Until r_rsScheduleRotationRequest.EOF
                r_rsScheduleRotationRequest!RECORD_STATUS = "T" 'processed
                r_rsScheduleRotationRequest.Update
                r_rsScheduleRotationRequest.MoveNext
            Loop
        End If
        If IsRecordsInFilterRecordSet(l_strFilter, r_rsScheduleRotationDetailRequest) Then '
            Set l_rsRequestDetail = CopyRecordset(r_rsScheduleRotationDetailRequest)
            r_rsScheduleRotationDetailRequest.MoveFirst
            Do Until r_rsScheduleRotationDetailRequest.EOF
                r_rsScheduleRotationDetailRequest!RECORD_STATUS = "T" 'processed
                r_rsScheduleRotationDetailRequest.Update
                r_rsScheduleRotationDetailRequest.MoveNext
            Loop
        End If
        If l_rsRequestHeader Is Nothing And l_rsRequestDetail Is Nothing Then
            Exit Do
        Else
            If RecordRotation(l_rsRequestHeader, l_rsRequestDetail, v_lngUserId, r_rsErrors) Then
                GetObjectContext.SetComplete
                MultiRecordRotation = True
            Else
                GetObjectContext.SetAbort
                MultiRecordRotation = False
                Exit Function
            End If
        End If
    Loop


    GetObjectContext.SetComplete
    
Exit Function

MultiRecordRotation_EH:

    MultiRecordRotation = False
    GetObjectContext.SetAbort
    
    TrsRaiseError mstrCMODULENAME, "MultiRecordRotation", , , , , True

End Function



Private Function HeaderAndDetailEOF(ByRef r_rsScheduleRotationRequest As ADODB.Recordset _
                                  , ByRef r_rsScheduleRotationDetailRequest As ADODB.Recordset _
                                  , ByRef r_lngSAMEROTCD As Long _
                                  , ByRef r_strSAMETimeStamp As String _
                                   ) As Boolean

On Error GoTo HeaderAndDetailEOF_EH

    Dim l_strFilter As String
    
    HeaderAndDetailEOF = True
            
    l_strFilter = "RECORD_STATUS = 'D' OR RECORD_STATUS = 'I' or RECORD_STATUS = 'C'"
    
    If IsRecordsInFilterRecordSet(l_strFilter, r_rsScheduleRotationRequest) Then
        r_lngSAMEROTCD = r_rsScheduleRotationRequest!ROT_CD
        r_strSAMETimeStamp = r_rsScheduleRotationRequest!CREATED_TIMESTAMP
        HeaderAndDetailEOF = False
        Exit Function
    End If
            
    If IsRecordsInFilterRecordSet(l_strFilter, r_rsScheduleRotationDetailRequest) Then
        r_lngSAMEROTCD = r_rsScheduleRotationDetailRequest!ROT_CD
        r_strSAMETimeStamp = r_rsScheduleRotationDetailRequest!CREATED_TIMESTAMP
        HeaderAndDetailEOF = False
    End If
            
Exit Function

HeaderAndDetailEOF_EH:
    TrsRaiseError mstrCMODULENAME, "HeaderAndDetailEOF"

End Function



