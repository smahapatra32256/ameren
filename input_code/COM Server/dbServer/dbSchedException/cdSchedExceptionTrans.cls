VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "cdSchedExceptionTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const m_strCMODULENAME As String = "cdSchedExceptionTrans"

Public Function SaveExceptions( _
                ByVal v_rsExceptions As ADODB.Recordset _
                , ByVal v_lngUserNum As Long _
                ) As Boolean
On Error GoTo eh
Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
Dim l_objCN As ADODB.Connection
Dim l_strSQL As String
Dim l_strAuditSQL As String
Dim l_colErrArgs As Collection
Dim l_colEmployees As Collection
Dim l_strHoldEmployee As String
Dim l_intI As Integer
Dim l_intNumOfEmps As String
Dim l_lngRecsAffected As Long

    v_rsExceptions.Filter = adFilterPendingRecords
    If v_rsExceptions.EOF Then
        SaveExceptions = True
        GetObjectContext.SetComplete
        Exit Function
    End If

    Set l_objCM = New util_DBConnectMgr.clsuDBConnectMgr
    Set l_objCN = New ADODB.Connection
    l_objCN.ConnectionString = l_objCM.GetConnectionString
    Set l_objCM = Nothing
    l_objCN.CursorLocation = adUseClient
    l_objCN.Open

    Set l_colEmployees = New Collection
    l_strHoldEmployee = ""
    Do Until v_rsExceptions.EOF
        l_strSQL = SaveExceptionSQL(v_rsExceptions, l_strAuditSQL, v_lngUserNum)
        If l_strSQL <> "" Then
            l_objCN.Execute l_strSQL
            AddToEmployeeCollection v_rsExceptions, l_colEmployees, l_strHoldEmployee
        End If
        If l_strAuditSQL <> "" Then
            l_objCN.Execute l_strAuditSQL
        End If
        v_rsExceptions.MoveNext
    Loop
    
    l_intNumOfEmps = l_colEmployees.Count
    For l_intI = 1 To l_intNumOfEmps
        l_strSQL = "UPDATE METHDBA.TTR_EMP_MOD" & _
            " SET LAST_MOD_SCHED_TM = CURRENT TIMESTAMP " & _
            " WHERE EMP_NUM = " & l_colEmployees.Item(l_intI)
        l_objCN.Execute l_strSQL, l_lngRecsAffected
        If l_lngRecsAffected = 0 Or l_lngRecsAffected = -1 Then
            l_strSQL = "INSERT INTO METHDBA.TTR_EMP_MOD" & _
                " (EMP_NUM, LAST_MOD_SCHED_TM)" & _
                " VALUES(" & l_colEmployees.Item(l_intI) & ", CURRENT TIMESTAMP )"
            l_objCN.Execute l_strSQL
        End If

    Next l_intI

    l_objCN.Close
    Set l_objCN = Nothing

    GetObjectContext.SetComplete
    SaveExceptions = True

    Exit Function
eh:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    GetObjectContext.SetAbort
    Set l_colErrArgs = New Collection
    l_colErrArgs.Add "SQL: " & l_strSQL
    TrsRaiseError m_strCMODULENAME, "SaveExceptions", , , , l_colErrArgs, True
End Function
Private Function SaveExceptionSQL( _
                ByVal v_rsException As ADODB.Recordset _
                , ByRef r_strAuditSQL As String _
                , ByVal v_lngUserNum As Long _
                ) As String
On Error GoTo eh
'Note only deletes are being audited
'Also there are no records that were modified by the business process.
'A change was processed as a delete and an insert.
Const CSCHEDTABLE As String = "METHDBA.TTR_EMP_SCHED_EXC"
Dim primKeys As Variant
Dim l_strTimeStamp As String
Dim l_strSQL As String
Dim l_varIgnore As Variant
Dim l_varIgnore2 As Variant
Dim l_varAdd(1, 1) As String

    r_strAuditSQL = ""
    l_varIgnore = Array("CREATED_TIMESTAMP", "LAST_MOD_TIMESTAMP")
    primKeys = Array("EMP_NUM", "CHARGE_DATE", "SHIFT_NUM", "HOURS_OFFSET_NUM")
    Select Case v_rsException.Status
       Case adRecModified
            l_strSQL = UpdateStatement(v_rsException, CSCHEDTABLE, primKeys, , True, "LAST_MOD_TIMESTAMP")
            If l_strSQL <> "" Then
                r_strAuditSQL = BuildAuditSQL(v_rsException, v_lngUserNum)
            End If
       Case adRecNew
           l_strSQL = InsertStatement(v_rsException, CSCHEDTABLE, l_varIgnore)
       Case adRecDeleted
            l_strSQL = DeleteStatement(v_rsException, CSCHEDTABLE, primKeys)
            If l_strSQL <> "" Then
                r_strAuditSQL = BuildAuditSQL(v_rsException, v_lngUserNum)
            End If
        Case Else
           l_strSQL = ""
    End Select
    SaveExceptionSQL = l_strSQL

    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "SaveExceptionSQL"
End Function

Private Function BuildAuditSQL( _
                ByVal v_rsException As ADODB.Recordset _
                , ByVal v_lngUserNum As Long _
                ) As String
On Error GoTo eh
Dim l_strSQL As String
Dim l_objField As ADODB.Field
Dim l_intI As Integer
Dim l_strFieldValue As String
Dim l_intScale As Integer
Dim l_intPrecision As Integer
Dim l_intCount As Integer
Dim l_strFormat As String

    If v_rsException.Fields.Count = 0 Then
        BuildAuditSQL = ""
        Exit Function
    End If
    l_strSQL = "INSERT INTO METHDBA.TTR_EMP_SCH_EX_AUD "
    l_strSQL = l_strSQL & "VALUES ( "
    l_intCount = 0
    For Each l_objField In v_rsException.Fields
        Select Case l_objField.Type
            Case adChar, adVarChar
                l_strFieldValue = "'" & ReplaceString(l_objField.OriginalValue, "'", "") & "'"
            Case adDate, adDBDate
                l_strFieldValue = Format(l_objField.OriginalValue, "yyyy-mm-dd")
                If l_strFieldValue = "0100-01-01" Then
                    l_strFieldValue = "'0001-01-01'"
                Else
                    l_strFieldValue = "'" & l_strFieldValue & "'"
                End If
            Case adDecimal, adNumeric
                l_intScale = l_objField.NumericScale
                l_intPrecision = l_objField.Precision
                l_intI = l_intPrecision - l_intScale
                l_strFormat = ""
                If l_intScale > 0 Then
                    l_strFormat = "." & String(l_intScale, "0")
                End If
                If l_intI > 0 Then
                    l_strFormat = String(l_intI - 1, "#") & "0" & l_strFormat
                End If
                l_strFieldValue = Format(l_objField.OriginalValue, l_strFormat)
            Case adInteger, adSmallInt
                l_strFieldValue = CStr(l_objField.OriginalValue)

        End Select
        If l_intCount > 0 Then
            l_strSQL = l_strSQL & " , " & l_strFieldValue
        Else
            l_strSQL = l_strSQL & l_strFieldValue
        End If
        l_intCount = l_intCount + 1

    Next l_objField
    l_strSQL = l_strSQL & " , " & str(v_lngUserNum)
    l_strSQL = l_strSQL & " , CURRENT TIMESTAMP"
    l_strSQL = l_strSQL & " , 'ES' )"

    BuildAuditSQL = l_strSQL

    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "BuildAuditSQL"

End Function
Private Sub AddToEmployeeCollection( _
            ByVal v_rsExceptions As ADODB.Recordset _
            , ByRef r_colEmployees As Collection _
            , ByRef r_strHoldEmployee As String)
On Error GoTo eh
Dim l_strEmployee As String

    Select Case v_rsExceptions.Status
       Case adRecModified
            l_strEmployee = Trim(str(v_rsExceptions!EMP_NUM))
       Case adRecNew
            l_strEmployee = Trim(str(v_rsExceptions!EMP_NUM))
       Case adRecDeleted
            l_strEmployee = Trim(str(v_rsExceptions.Fields("EMP_NUM").OriginalValue))
    End Select

    If l_strEmployee <> r_strHoldEmployee Then
        r_strHoldEmployee = l_strEmployee
        If Not IncludesString(l_strEmployee, r_colEmployees) Then
            r_colEmployees.Add l_strEmployee
        End If
    End If

    Exit Sub
eh:
    TrsRaiseError m_strCMODULENAME, "AddToEmployeeCollection"

End Sub
