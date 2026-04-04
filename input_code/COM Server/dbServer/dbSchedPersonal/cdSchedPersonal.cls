VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "cdSchedPersonal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const m_strCMODULENAME As String = "cdSchedPersonal"
Private Const m_datCOPTIONALSTARTDATE As Date = #1/1/100#
Private Const m_datCOPTIONALENDDATE As Date = #12/31/9999#

'provides the Personal  detail information for History List
Public Function GetPersonalSchedules _
                (ByVal v_lngEmpNum As Long _
                 , Optional ByVal v_datStartDate As Date = m_datCOPTIONALSTARTDATE _
                 , Optional ByVal v_datEndDate As Date = m_datCOPTIONALENDDATE _
                ) As ADODB.Recordset

    
On Error GoTo GetPersonalSchedules_EH
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
   
    If v_lngEmpNum < 1 Then
        Err.Raise trErrdbInvalidZerolngEmpNum
    Else
    End If
   
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
   
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open
    
    l_strSQL = RetrievePersonalSchedules(v_lngEmpNum, v_datStartDate, v_datEndDate)

    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing
    If l_objRS.BOF And l_objRS.EOF Then
        Set GetPersonalSchedules = Nothing
    Else
        Set GetPersonalSchedules = l_objRS
    End If
    GetObjectContext.SetComplete
    Set l_objRS = Nothing
    
Exit Function

GetPersonalSchedules_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    GetObjectContext.SetAbort
    
    Select Case Err.Number
        Case trErrdbInvalidZerolngEmpNum
            TrsRaiseError m_strCMODULENAME, "GetPersonalSchedules", trErrdbInvalidZerolngEmpNum, , , , True
        Case Else
            TrsRaiseError m_strCMODULENAME, "GetPersonalSchedules", , , , DescArgsFrom(l_strSQL), True
    End Select
    
End Function
'start and end dates in between
'provide sql for public method GetPersonalSchedules
'or provide executable transactional method
'allow the user to select all nondemand personal records for history list display
'allow program to detect any overlaps for the given start and end date
'the line number allows to filter the nondemand personal records
'this function provides the demand records to check for overlap
Private Function RetrievePersonalSchedules _
                (ByVal v_lngEmpNum As Long, _
                 Optional ByVal v_datStartDate As Date = m_datCOPTIONALSTARTDATE, _
                 Optional ByVal v_datEndDate As Date = m_datCOPTIONALENDDATE, _
                 Optional ByVal cn As ADODB.Connection, _
                 Optional ByRef r_rsPersonalSchedules As ADODB.Recordset) As String

    
On Error GoTo RetrievePersonalSchedules_EH

    Dim strSQL As String
    Dim errArgs As New Collection

    strSQL = "Select A.EMP_NUM , A.START_DATE, A.END_DATE, A.ROT_CD, A.ROT_WEEK, A.SCHED_CD, " & _
        " A.CREATED_EMP_NUM, A.LAST_MOD_EMP_NUM, CHAR(A.CREATED_TIMESTAMP) AS CREATED_TIMESTAMP, CHAR(A.LAST_MOD_TIMESTAMP) AS LAST_MOD_TIMESTAMP, " & _
        " A.EMP_NUM AS KEY_EMP_NUM, A.START_DATE AS KEY_START_DATE, A.END_DATE AS KEY_END_DATE, ' ' AS RECORD_STATUS, ' ' AS RECORD_ID, B.EMP_NAME " & _
               " FROM METHDBA.TTR_EMP_SCHED A" & _
               " , METHDBA.TTR_EMPLOYEE B"
    If v_datStartDate = m_datCOPTIONALSTARTDATE Then
        strSQL = strSQL & " WHERE B.EMP_NUM = A.EMP_NUM " & _
               " AND A.EMP_NUM = " & v_lngEmpNum & " " & _
               " ORDER BY A.EMP_NUM , A.START_DATE"
    Else
        If v_datEndDate = m_datCOPTIONALENDDATE Then
            strSQL = strSQL & " WHERE B.EMP_NUM = A.EMP_NUM " & _
                   " AND A.EMP_NUM = " & v_lngEmpNum & " " & _
                   " AND ((A.START_DATE >= " & _
                   " '" & Format(v_datStartDate, "mm/dd/yyyy") & "' " & _
                   " AND A.START_DATE <= " & _
                   " '" & Format(m_datCOPTIONALENDDATE, "mm/dd/yyyy") & "') " & _
                   " OR (A.END_DATE >= " & _
                   " '" & Format(v_datStartDate, "mm/dd/yyyy") & "' " & _
                   " AND A.END_DATE <= " & _
                   " '" & Format(m_datCOPTIONALENDDATE, "mm/dd/yyyy") & "'))" & _
                   " ORDER BY A.EMP_NUM , A.START_DATE"
        Else  'is start date or end date of schedule records contain in the new date range given (not out of the range)
            strSQL = strSQL & " WHERE B.EMP_NUM = A.EMP_NUM " & _
                   " AND A.EMP_NUM = " & v_lngEmpNum & " " & _
                   " AND ((A.START_DATE >= " & _
                   " '" & Format(v_datStartDate, "mm/dd/yyyy") & "' " & _
                   " AND A.START_DATE <= " & _
                   " '" & Format(v_datEndDate, "mm/dd/yyyy") & "') " & _
                   " OR (A.END_DATE >= " & _
                   " '" & Format(v_datStartDate, "mm/dd/yyyy") & "' " & _
                   " AND A.END_DATE <= " & _
                   " '" & Format(v_datEndDate, "mm/dd/yyyy") & "'))" & _
                   " ORDER BY A.EMP_NUM , A.START_DATE"
        End If
    End If

    If cn Is Nothing Then
        RetrievePersonalSchedules = strSQL
    Else
        If strSQL <> "" Then
            Set r_rsPersonalSchedules = cn.Execute(strSQL)
            If RecordsExist(r_rsPersonalSchedules) Then
            Else
               Set r_rsPersonalSchedules = Nothing
            End If
        End If
    End If
    
    
Exit Function
RetrievePersonalSchedules_EH:
    
    errArgs.Add strSQL
    TrsRaiseError m_strCMODULENAME, "RetrievePersonalSchedules", , True, , errArgs, True

End Function




'Detail*********
'provides delete, change and update for Personal  detail
Public Function SavePersonalSchedule(ByVal v_rsPersonalSchedule As ADODB.Recordset) As Boolean

On Error GoTo SavePersonalSchedule_EH
    
    Dim cn As ADODB.Connection
    
    If RecordsExist(v_rsPersonalSchedule) Then
        v_rsPersonalSchedule.Filter = ""
        v_rsPersonalSchedule.MoveFirst
    Else
        SavePersonalSchedule = True
        GetObjectContext.SetComplete
        Exit Function
    End If
    
    Set cn = OpenConnection
       
'  Perform Deletes
    ' Filter for Deletes
    v_rsPersonalSchedule.Filter = "RECORD_STATUS = 'D'"
    If Not v_rsPersonalSchedule.EOF Then
        Do Until v_rsPersonalSchedule.EOF
            DeletePersonalSchedule v_rsPersonalSchedule, cn
            v_rsPersonalSchedule.MoveNext
        Loop
    End If
    
'  Perform Changes
    v_rsPersonalSchedule.Filter = ""
    v_rsPersonalSchedule.MoveFirst
    v_rsPersonalSchedule.Filter = "RECORD_STATUS = 'C'"
    Do Until v_rsPersonalSchedule.EOF
        UpdatePersonalSchedule v_rsPersonalSchedule, cn
        v_rsPersonalSchedule.MoveNext
    Loop
    
'  Perform Inserts
    v_rsPersonalSchedule.Filter = ""
    v_rsPersonalSchedule.MoveFirst
    v_rsPersonalSchedule.Filter = "RECORD_STATUS = 'I'"
    Do Until v_rsPersonalSchedule.EOF
       InsertPersonalSchedule v_rsPersonalSchedule, cn
       v_rsPersonalSchedule.MoveNext
    Loop

    cn.Close
    
    Set cn = Nothing
    GetObjectContext.SetComplete
    
    SavePersonalSchedule = True
    
Exit Function
SavePersonalSchedule_EH:

    GetObjectContext.SetAbort
    Set cn = Nothing
    
    TrsRaiseError m_strCMODULENAME, "SavePersonalSchedule", , , , , True
    

End Function



'insert Personal  detail
Private Sub InsertPersonalSchedule(v_rsPersonalSchedule As ADODB.Recordset, cn As ADODB.Connection)

On Error GoTo InsertPersonalSchedule_EH

    Dim ignoreFields As Variant
    Dim strSQL As String
    Dim errArgs As New Collection
    Dim addDefaultFields() As Variant
    
    
    If Trim(v_rsPersonalSchedule!CREATED_TIMESTAMP) <> "" Then
        ignoreFields = Array("RECORD_STATUS", "RECORD_ID", "KEY_EMP_NUM", "KEY_START_DATE", "KEY_END_DATE")
    Else
        ignoreFields = Array("RECORD_STATUS", "RECORD_ID", "KEY_EMP_NUM", "KEY_START_DATE", "KEY_END_DATE", "CREATED_TIMESTAMP", "LAST_MOD_TIMESTAMP")
    End If
    strSQL = InsertStatement(v_rsPersonalSchedule, "METHDBA.TTR_EMP_SCHED", ignoreFields)
    
    If strSQL <> "" Then
       cn.Execute strSQL
    End If
    
    GetObjectContext.SetComplete
    
Exit Sub

InsertPersonalSchedule_EH:

    GetObjectContext.SetAbort

    Dim errLoop As ADODB.Error
    For Each errLoop In cn.Errors
         If errLoop.NativeError = -803 Then 'duplicate BRV to be handle by server bus component
            Err.Raise trErrbusSchedRotDuplicate
            Exit Sub
         Else
            errArgs.Add strSQL
            TrsRaiseError m_strCMODULENAME, "InsertPersonalSchedule", , True, , errArgs, True
            Exit Sub
         End If
    Next

    errArgs.Add strSQL
    TrsRaiseError m_strCMODULENAME, "InsertPersonalSchedule", , True, , errArgs, True


End Sub
'update Personal  detail record
Private Sub UpdatePersonalSchedule(v_rsPersonalSchedule As ADODB.Recordset, cn As ADODB.Connection)

On Error GoTo UpdatePersonalSchedule_EH
    
    Dim primKeys As Variant
    Dim ignoreFields As Variant
    Dim strSQL As String
    Dim errArgs As New Collection
    Dim lngRecsAffected As Long
    
'    Dim l_utilDBDataServices As utilDBDataServices.cuDBDataServices
'
'    Set l_utilDBDataServices = New utilDBDataServices.cuDBDataServices

    strSQL = "Update METHDBA.TTR_EMP_SCHED " _
        & " SET SCHED_CD = " & v_rsPersonalSchedule!sched_cd _
        & " ,ROT_CD = " & v_rsPersonalSchedule!ROT_CD _
        & " ,ROT_WEEK = " & v_rsPersonalSchedule!ROT_WEEK _
        & " ,LAST_MOD_TIMESTAMP = CURRENT TIMESTAMP" _
        & " ,LAST_MOD_EMP_NUM = " & v_rsPersonalSchedule!LAST_MOD_EMP_NUM _
        & " ,START_DATE = " & "'" & Format(v_rsPersonalSchedule!start_date, "mm/dd/yyyy") & "'" _
        & " ,END_DATE = " & "'" & Format(v_rsPersonalSchedule!END_DATE, "mm/dd/yyyy") & "'" _
        & " WHERE EMP_NUM = " & v_rsPersonalSchedule!KEY_EMP_NUM & " " _
        & " AND START_DATE = '" & Format(v_rsPersonalSchedule!KEY_START_DATE, "mm/dd/yyyy") & "'" _
        & " AND END_DATE = '" & Format(v_rsPersonalSchedule!KEY_END_DATE, "mm/dd/yyyy") & "'" _
        & " AND LAST_MOD_TIMESTAMP = '" & v_rsPersonalSchedule!LAST_MOD_TIMESTAMP & "'"

    If strSQL <> "" Then
       cn.Execute strSQL, lngRecsAffected
       If lngRecsAffected = -1 Or lngRecsAffected = 0 Then
            Err.Raise trErrAppCommitConflict
       End If
    End If
    
    GetObjectContext.SetComplete
    
Exit Sub

UpdatePersonalSchedule_EH:
    
    GetObjectContext.SetAbort
    errArgs.Add strSQL
    
    Select Case Err.Number
        Case trErrAppCommitConflict
            Err.Source = Join( _
                            Array( _
                                CStr(TrType_PER), _
                                CStr(v_rsPersonalSchedule!KEY_EMP_NUM), _
                                Format(v_rsPersonalSchedule!KEY_START_DATE, "mm/dd/yyyy"), _
                                Format(v_rsPersonalSchedule!KEY_END_DATE, "mm/dd/yyyy")) _
                            , ";")
            TrsRaiseError m_strCMODULENAME, "UpdatePersonalSchedule", Err.Number
        Case Else
            TrsRaiseError m_strCMODULENAME, "UpdatePersonalSchedule", , True, , errArgs, True
    End Select


End Sub
'Delete Personal  detail
Private Sub DeletePersonalSchedule(v_rsPersonalSchedule As ADODB.Recordset, cn As ADODB.Connection)

On Error GoTo DeletePersonalSchedule_EH
    
    Dim primKeys As Variant
    Dim strSQL As String
    Dim errArgs As New Collection
    Dim lngRecsAffected As Long
    
'    strSQL = "DELETE FROM METHDBA.TTR_EMP_SCHED" & _
'               " WHERE EMP_NUM = " & v_rsPersonalSchedule!KEY_EMP_NUM & " " & _
'               " AND ((START_DATE >= " & _
'               " '" & Format(v_rsPersonalSchedule!KEY_START_DATE, "mm/dd/yyyy") & "' " & _
'               " AND START_DATE <= " & _
'               " '" & Format(v_rsPersonalSchedule!KEY_END_DATE, "mm/dd/yyyy") & "') " & _
'               " OR (END_DATE >= " & _
'               " '" & Format(v_rsPersonalSchedule!KEY_START_DATE, "mm/dd/yyyy") & "' " & _
'               " AND END_DATE <= " & _
'               " '" & Format(v_rsPersonalSchedule!KEY_END_DATE, "mm/dd/yyyy") & "'))"
 
'        6/10/02 RSG - Support optimistic locking
'        Change the data service to only delete specific records, i.e. timestamp
'        Do not attempt delete of set of records.
 
    strSQL = "DELETE FROM METHDBA.TTR_EMP_SCHED" & _
               " WHERE EMP_NUM = " & v_rsPersonalSchedule!KEY_EMP_NUM & _
               " AND START_DATE = " & _
               " '" & Format(v_rsPersonalSchedule!KEY_START_DATE, "mm/dd/yyyy") & "'" & _
               " AND END_DATE = " & _
               " '" & Format(v_rsPersonalSchedule!KEY_END_DATE, "mm/dd/yyyy") & "'" & _
               " AND LAST_MOD_TIMESTAMP = '" & v_rsPersonalSchedule!LAST_MOD_TIMESTAMP & "'"
  
    If strSQL <> "" Then
       cn.Execute strSQL, lngRecsAffected
       If lngRecsAffected = -1 Or lngRecsAffected = 0 Then
            Err.Raise trErrAppCommitConflict
       End If
    End If
    
    GetObjectContext.SetComplete
    
Exit Sub

DeletePersonalSchedule_EH:
    
    GetObjectContext.SetAbort
    errArgs.Add strSQL
    
    Select Case Err.Number
        Case trErrAppCommitConflict
            Err.Source = Join( _
                            Array( _
                                CStr(TrType_PER), _
                                CStr(v_rsPersonalSchedule!KEY_EMP_NUM), _
                                Format(v_rsPersonalSchedule!KEY_START_DATE, "mm/dd/yyyy"), _
                                Format(v_rsPersonalSchedule!KEY_END_DATE, "mm/dd/yyyy"), _
                                v_rsPersonalSchedule!LAST_MOD_TIMESTAMP) _
                            , ";")
            TrsRaiseError m_strCMODULENAME, "DeletePersonalSchedule", Err.Number
        Case Else
            TrsRaiseError m_strCMODULENAME, "DeletePersonalSchedule", , True, , errArgs, True
    End Select
    
End Sub


'Open connection for transactional items

Private Function OpenConnection() As ADODB.Connection
  
On Error GoTo OpenConnection_Eh
    
    Dim cm As New util_DBConnectMgr.clsuDBConnectMgr
    Dim cn As ADODB.Connection
    
    Set cn = New ADODB.Connection
    With cn
        .CursorLocation = adUseClient       'Must use client to support RecordCount
        .ConnectionString = cm.GetConnectionString
        .Open
    End With
    
    Set OpenConnection = cn
    Set cm = Nothing
    Set cn = Nothing
    Exit Function
    
OpenConnection_Eh:

    Set cm = Nothing
    Set cn = Nothing
    
    TrsRaiseError m_strCMODULENAME, "OpenConnection"
 
End Function
'start and end dates cross
Private Function RetrieveScopePersonalSched _
                (ByVal v_lngEmpNum As Long, _
                  ByVal v_datStartDate As Date, _
                  ByVal v_datEndDate As Date, _
                 Optional ByVal cn As ADODB.Connection, _
                 Optional ByRef r_rsPersonalSchedules As ADODB.Recordset) As String
    
On Error GoTo RetrieveScopePersonalSched_EH

    Dim strSQL As String
    Dim errArgs As New Collection
    Dim l_objRS As ADODB.Recordset
    
    strSQL = "Select EMP_NUM , START_DATE, END_DATE, ROT_CD, ROT_WEEK, SCHED_CD, " & _
        " CREATED_EMP_NUM, LAST_MOD_EMP_NUM, CHAR(CREATED_TIMESTAMP) AS CREATED_TIMESTAMP, CHAR(LAST_MOD_TIMESTAMP) AS LAST_MOD_TIMESTAMP, " & _
        " EMP_NUM AS KEY_EMP_NUM, START_DATE AS KEY_START_DATE, END_DATE AS KEY_END_DATE, ' ' AS RECORD_STATUS, ' ' AS RECORD_ID " & _
               " FROM       METHDBA.TTR_EMP_SCHED"
    strSQL = strSQL & " WHERE EMP_NUM = " & v_lngEmpNum & " " & _
           " AND " & _
           " '" & Format(v_datStartDate, "mm/dd/yyyy") & "' " & _
           " <= END_DATE " & _
           " AND " & _
           " '" & Format(v_datEndDate, "mm/dd/yyyy") & "' " & _
           " >= START_DATE " & _
           " ORDER BY EMP_NUM , START_DATE"
    
    If cn Is Nothing Then
        RetrieveScopePersonalSched = strSQL
    Else
        If strSQL <> "" Then
            Set r_rsPersonalSchedules = cn.Execute(strSQL)
            If RecordsExist(r_rsPersonalSchedules) Then
            Else
               Set r_rsPersonalSchedules = Nothing
            End If
        End If
    End If
    
Exit Function

RetrieveScopePersonalSched_EH:
    
    errArgs.Add strSQL
    TrsRaiseError m_strCMODULENAME, "RetrieveScopePersonalSched", , True, , errArgs, True

End Function

'retrieves the scope of records that will touched by the requested start and requested end dates
'it encompasses all records that need to be changed
'includes all TT, FT, TF (excluding all FF) - boolean
'if DBstart >= ReqStart and if DBEnd <= ReqEnd
'if DBstart >= ReqStart
'   if DBEnd <= ReqEnd
'       delete TT
'   else
'       Update TF
'else
'   Update FT
'end if

Public Function GetScopePersonalSched _
                (ByVal v_lngEmpNum As Long _
                 , ByVal v_datStartDate As Date _
                 , ByVal v_datEndDate As Date _
                ) As ADODB.Recordset

On Error GoTo GetScopePersonalSched_EH
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
   
    If v_lngEmpNum < 1 Then
        Err.Raise trErrdbInvalidZerolngEmpNum
    Else
    End If


   
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
   
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open
    
    l_strSQL = RetrieveScopePersonalSched(v_lngEmpNum, v_datStartDate, v_datEndDate)

    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing
    If l_objRS.BOF And l_objRS.EOF Then
        Set GetScopePersonalSched = Nothing
    Else
        Set GetScopePersonalSched = l_objRS
    End If
    GetObjectContext.SetComplete
    Set l_objRS = Nothing
    Exit Function

GetScopePersonalSched_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    GetObjectContext.SetAbort
    
    
    Select Case Err.Number
        Case trErrdbInvalidZerolngEmpNum
            TrsRaiseError m_strCMODULENAME, "GetScopePersonalSched", trErrdbInvalidZerolngEmpNum, , , , True
        Case Else
            TrsRaiseError m_strCMODULENAME, "GetScopePersonalSched", , , , DescArgsFrom(l_strSQL), True
    End Select
    
    
End Function

'This method is used to check for existing Personal Schedules
Public Function GetExistPersonalSched _
                (ByVal v_lngEmpNum As Long _
                 , ByVal v_datStartDate As Date _
                 , ByVal v_datEndDate As Date _
                ) As ADODB.Recordset

    
On Error GoTo GetExistPersonalSched_EH
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    
    If v_lngEmpNum < 1 Then
        Err.Raise trErrdbInvalidZerolngEmpNum
    Else
    End If
    
   
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
   
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open
    
    l_strSQL = RetrieveExistPersSched(v_lngEmpNum, v_datStartDate, v_datEndDate)

    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing
    If l_objRS.BOF And l_objRS.EOF Then
        Set GetExistPersonalSched = Nothing
    Else
        Set GetExistPersonalSched = l_objRS
    End If
    GetObjectContext.SetComplete
    Set l_objRS = Nothing
    Exit Function

GetExistPersonalSched_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    GetObjectContext.SetAbort
    
    Select Case Err.Number
        Case trErrdbInvalidZerolngEmpNum
            TrsRaiseError m_strCMODULENAME, "GetExistPersonalSched", trErrdbInvalidZerolngEmpNum, , , , True
        Case Else
            TrsRaiseError m_strCMODULENAME, "GetExistPersonalSched", , , , DescArgsFrom(l_strSQL), True
    End Select
   
End Function


Private Function RetrieveExistPersSched _
                (ByVal v_lngEmpNum As Long, _
                  ByVal v_datStartDate As Date, _
                  ByVal v_datEndDate As Date, _
                 Optional ByVal cn As ADODB.Connection, _
                 Optional ByRef r_rsPersonalSchedules As ADODB.Recordset) As String
    
On Error GoTo RetrieveExistPersSched_EH

    Dim strSQL As String
    Dim errArgs As New Collection
    
    strSQL = "Select EMP_NUM , START_DATE, END_DATE, ROT_CD, ROT_WEEK, SCHED_CD, " & _
        " CREATED_EMP_NUM, LAST_MOD_EMP_NUM, CHAR(CREATED_TIMESTAMP) AS CREATED_TIMESTAMP, CHAR(LAST_MOD_TIMESTAMP) AS LAST_MOD_TIMESTAMP, " & _
        " EMP_NUM AS KEY_EMP_NUM, START_DATE AS KEY_START_DATE, END_DATE AS KEY_END_DATE, ' ' AS RECORD_STATUS, ' ' AS RECORD_ID " & _
               " FROM       METHDBA.TTR_EMP_SCHED"
    strSQL = strSQL & " WHERE EMP_NUM = " & v_lngEmpNum & " " & _
           " AND " & _
           " START_DATE = " & _
           " '" & Format(v_datStartDate, "mm/dd/yyyy") & "' " & _
           " AND " & _
           " END_DATE = " & _
           " '" & Format(v_datEndDate, "mm/dd/yyyy") & "' "
    
    If cn Is Nothing Then
        RetrieveExistPersSched = strSQL
    Else
        If strSQL <> "" Then
            Set r_rsPersonalSchedules = cn.Execute(strSQL)
            If RecordsExist(r_rsPersonalSchedules) Then
            Else
               Set r_rsPersonalSchedules = Nothing
            End If
        End If
    End If
    
Exit Function
RetrieveExistPersSched_EH:
    
    errArgs.Add strSQL
    TrsRaiseError m_strCMODULENAME, "RetrieveExistPersSched", , True, , errArgs, True

End Function


Public Function GetPersonalSchedForMultiEmp _
                (ByVal v_aryEmployees As Variant _
                 , Optional ByVal v_datStartDate As Date = m_datCOPTIONALSTARTDATE _
                 , Optional ByVal v_datEndDate As Date = m_datCOPTIONALENDDATE _
                ) As ADODB.Recordset

    
On Error GoTo GetPersonalSchedForMultiEmp_EH
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim l_lngStart As Long
   
    l_lngStart = LBound(v_aryEmployees)
    If Trim(v_aryEmployees(l_lngStart)) = "" Then
        Err.Raise trErrdbInvalidZerolngEmpNum
    Else
    End If
   
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
   
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open
    
    l_strSQL = RetrievePersSchedForMultiEmp(v_aryEmployees, v_datStartDate, v_datEndDate)

    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing
    If l_objRS.BOF And l_objRS.EOF Then
        Set GetPersonalSchedForMultiEmp = Nothing
    Else
        Set GetPersonalSchedForMultiEmp = l_objRS
    End If
    GetObjectContext.SetComplete
    Set l_objRS = Nothing
    
Exit Function

GetPersonalSchedForMultiEmp_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    GetObjectContext.SetAbort
    
    Select Case Err.Number
        Case trErrdbInvalidZerolngEmpNum
            TrsRaiseError m_strCMODULENAME, "GetPersonalSchedForMultiEmp", trErrdbInvalidZerolngEmpNum, , , , True
        Case Else
            TrsRaiseError m_strCMODULENAME, "GetPersonalSchedForMultiEmp", , , , DescArgsFrom(l_strSQL), True
    End Select
    
End Function

Private Function RetrievePersSchedForMultiEmp _
                (ByVal v_aryEmployees As Variant, _
                 Optional ByVal v_datStartDate As Date = m_datCOPTIONALSTARTDATE, _
                 Optional ByVal v_datEndDate As Date = m_datCOPTIONALENDDATE, _
                 Optional ByVal cn As ADODB.Connection, _
                 Optional ByRef r_rsPersonalSchedules As ADODB.Recordset) As String

    
On Error GoTo RetrievePersSchedForMultiEmp_EH

    Dim strSQL As String
    Dim errArgs As New Collection

    strSQL = "Select A.EMP_NUM , A.START_DATE, A.END_DATE, A.ROT_CD, A.ROT_WEEK, A.SCHED_CD, " & _
        " A.CREATED_EMP_NUM, A.LAST_MOD_EMP_NUM, CHAR(A.CREATED_TIMESTAMP) AS CREATED_TIMESTAMP, CHAR(A.LAST_MOD_TIMESTAMP) AS LAST_MOD_TIMESTAMP, " & _
        " A.EMP_NUM AS KEY_EMP_NUM, A.START_DATE AS KEY_START_DATE, A.END_DATE AS KEY_END_DATE, ' ' AS RECORD_STATUS, ' ' AS RECORD_ID, B.EMP_NAME " & _
               " FROM  METHDBA.TTR_EMP_SCHED A" & _
               " , METHDBA.TTR_EMPLOYEE B"
    If v_datStartDate = m_datCOPTIONALSTARTDATE Then
        strSQL = strSQL & " WHERE B.EMP_NUM = A.EMP_NUM " & _
               " AND " & BuildInClause("A.EMP_NUM", v_aryEmployees, False) & " " & _
               " ORDER BY A.EMP_NUM , A.START_DATE"
    Else
        If v_datEndDate = m_datCOPTIONALENDDATE Then
            strSQL = strSQL & " WHERE B.EMP_NUM = A.EMP_NUM " & _
                   " AND " & BuildInClause("A.EMP_NUM", v_aryEmployees, False) & " " & _
                   " AND ((A.START_DATE >= " & _
                   " '" & Format(v_datStartDate, "mm/dd/yyyy") & "' " & _
                   " AND A.START_DATE <= " & _
                   " '" & Format(m_datCOPTIONALENDDATE, "mm/dd/yyyy") & "') " & _
                   " OR (A.END_DATE >= " & _
                   " '" & Format(v_datStartDate, "mm/dd/yyyy") & "' " & _
                   " AND A.END_DATE <= " & _
                   " '" & Format(m_datCOPTIONALENDDATE, "mm/dd/yyyy") & "'))" & _
                   " ORDER BY A.EMP_NUM , A.START_DATE"
        Else  'is start date or end date of schedule records contain in the new date range given (not out of the range)
            strSQL = strSQL & " WHERE B.EMP_NUM = A.EMP_NUM " & _
                   " AND " & BuildInClause("A.EMP_NUM", v_aryEmployees, False) & " " & _
                   " AND ((A.START_DATE >= " & _
                   " '" & Format(v_datStartDate, "mm/dd/yyyy") & "' " & _
                   " AND A.START_DATE <= " & _
                   " '" & Format(v_datEndDate, "mm/dd/yyyy") & "') " & _
                   " OR (A.END_DATE >= " & _
                   " '" & Format(v_datStartDate, "mm/dd/yyyy") & "' " & _
                   " AND A.END_DATE <= " & _
                   " '" & Format(v_datEndDate, "mm/dd/yyyy") & "'))" & _
                   " ORDER BY A.EMP_NUM , A.START_DATE"
        End If
    End If

    If cn Is Nothing Then
        RetrievePersSchedForMultiEmp = strSQL
    Else
        If strSQL <> "" Then
            Set r_rsPersonalSchedules = cn.Execute(strSQL)
            If RecordsExist(r_rsPersonalSchedules) Then
            Else
               Set r_rsPersonalSchedules = Nothing
            End If
        End If
    End If
    
    
Exit Function
RetrievePersSchedForMultiEmp_EH:
    
    errArgs.Add strSQL
    TrsRaiseError m_strCMODULENAME, "RetrievePersSchedForMultiEmp", , True, , errArgs, True

End Function


Public Function GetSchedChange() As ADODB.Recordset

On Error GoTo GetSchedChange_EH
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
   
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open
    
    l_strSQL = RetrieveSchedChange

    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing
    If l_objRS.BOF And l_objRS.EOF Then
        Set GetSchedChange = Nothing
    Else
        Set GetSchedChange = l_objRS
    End If
    GetObjectContext.SetComplete
    Set l_objRS = Nothing
    Exit Function

GetSchedChange_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    GetObjectContext.SetAbort
    
    TrsRaiseError m_strCMODULENAME, "GetSchedChange", , , , DescArgsFrom(l_strSQL), True
   
End Function


Private Function RetrieveSchedChange _
                (Optional ByVal cn As ADODB.Connection, _
                 Optional ByRef r_rsSchedChange As ADODB.Recordset) As String
    
On Error GoTo RetrieveSchedChange_EH

    Dim strSQL As String
    Dim errArgs As New Collection

    strSQL = "select emp_num, sched_cd,  0 AS ROT_CD, 0 AS ROT_WEEK, date AS START_DATE, date AS END_DATE, 'I' AS RECORD_STATUS, " & _
             " CHAR(time_stamp) AS CREATED_TIMESTAMP, CHAR(time_stamp) AS LAST_MOD_TIMESTAMP, last_mod_emp_num as CREATED_EMP_NUM, last_mod_emp_num, '1' AS RECORD_ID " & _
               " From methdba.ttr_base_hrs_exc "
    strSQL = strSQL & " Where sched_type = 10 " & _
           " AND SUBSYS_OWNER not IN ('T') " & _
           " and exception_cd not in ('RWT','RLF','TOA','TOR','TSA','TSR','TWS') " & _
           " and sched_cd not in (-9,-2,-1) " & _
           " AND LINE_NUM = 0 ORDER BY emp_num , date"

    If cn Is Nothing Then
        RetrieveSchedChange = strSQL
    Else
        If strSQL <> "" Then
            Set r_rsSchedChange = cn.Execute(strSQL)
            If RecordsExist(r_rsSchedChange) Then
            Else
               Set r_rsSchedChange = Nothing
            End If
        End If
    End If
    
Exit Function
RetrieveSchedChange_EH:
    
    errArgs.Add strSQL
    TrsRaiseError m_strCMODULENAME, "RetrieveSchedChange", , True, , errArgs, True

End Function

'***
Public Function GetDBDefaultEndDate _
                (ByVal v_lngEmpNum As Long _
                 , ByVal v_datStartDate As Date) As ADODB.Recordset

On Error GoTo GetDBDefaultEndDate_EH
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
   
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open
    
    l_strSQL = RetrieveDBDefaultEndDate(v_lngEmpNum, v_datStartDate)

    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing
    
    If l_objRS.BOF And l_objRS.EOF Then
        Set GetDBDefaultEndDate = Nothing
    Else
        Set GetDBDefaultEndDate = l_objRS
    End If

'    If l_objRS.BOF And l_objRS.EOF Then
'        GetDBDefaultEndDate = #1/1/100#
'    Else
'        GetDBDefaultEndDate = CDate(l_objRS.Fields(0))
'    End If
    GetObjectContext.SetComplete
    Set l_objRS = Nothing
    Exit Function

GetDBDefaultEndDate_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    GetObjectContext.SetAbort
    
    TrsRaiseError m_strCMODULENAME, "GetDBDefaultEndDate", , , , DescArgsFrom(l_strSQL), True
   
End Function

'using min and max optional dates if eof is reached, this determines the default of 12/31/9999
Private Function RetrieveDBDefaultEndDate _
                (ByVal v_lngEmpNum As Long _
                 , ByVal v_datStartDate As Date, _
                 Optional ByVal cn As ADODB.Connection, _
                 Optional ByRef r_rsSchedChange As ADODB.Recordset) As String
'
On Error GoTo RetrieveDBDefaultEndDate_EH

    Dim strSQL As String
    Dim errArgs As New Collection

'select coalesce((Min(start_date) - 1 day),'9999-12-31') as END_DATE
'From methdba.ttr_emp_sched
'Where emp_num = 18098
'and start_date > '1997-01-01'
    
    strSQL = "select coalesce((Min(start_date)),'0100-01-01') as START_DATE , coalesce((Min(end_date)),'9999-12-31') as END_DATE  "
    strSQL = strSQL & " From methdba.ttr_emp_sched "
    strSQL = strSQL & " WHERE EMP_NUM = " & v_lngEmpNum & " " & _
                   " AND START_DATE > " & _
                   " '" & Format(v_datStartDate, "mm/dd/yyyy") & "' "


    If cn Is Nothing Then
        RetrieveDBDefaultEndDate = strSQL
    Else
        If strSQL <> "" Then
            Set r_rsSchedChange = cn.Execute(strSQL)
            If RecordsExist(r_rsSchedChange) Then
            Else
               Set r_rsSchedChange = Nothing
            End If
        End If
    End If
    
Exit Function
RetrieveDBDefaultEndDate_EH:
    
    errArgs.Add strSQL
    TrsRaiseError m_strCMODULENAME, "RetrieveDBDefaultEndDate", , True, , errArgs, True

End Function


