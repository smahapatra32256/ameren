VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cdCollectionPeriod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const m_strCMODULENAME As String = "cdCollPeriodPeriod"

' Static data constants
Private Const CRECORDNOTFOUND = "RECORDNOTFOUND"
Private Const CDEFAULTDATE As String = "01/01/0100"

' Department constants
Private Const CCP_GROUPNAME = "COLLPERIODCACHE"
' SharedProperty value(CCP_ROWVALUES) is a 2 dim array with 3 fields,
' PAY_PER_START_DATE, PAY_PER_END_DATE and PAY_PROCESSED_IND.
' index for record dynamic array
'
' Changed  start index position to 1, since 0 is now reserved for key string.
'
Private Const CCP_PAYROLL_GROUP_CD = 0
Private Const CCP_ROWVALUES = 1
Private Const CCP_LAST_SLOT = 1

Private Const CCP_PAY_PER_START_DATE = 0
Private Const CCP_PAY_PER_END_DATE = 1
Private Const CCP_PAY_PROCESSED_IND = 2

'*********************************************************
' Begin Public procedures
'*********************************************************

Public Function GetCollectionPeriodStatusByPayGroup(ByVal v_datDate As Date, ByVal v_strPayrollCode As String) As ADODB.Recordset
    
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    
    On Error GoTo GetCollectionPeriodStatusByPayGroup_EH
    
    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
   
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open

    l_strSQL = "SELECT PAY_PER_START_DATE, PAY_PER_END_DATE, PAY_PROCESSED_IND, PAYROLL_GROUP_CD "
    l_strSQL = l_strSQL & "From METHDBA.TTR_COLL_PERIODS "
    l_strSQL = l_strSQL & "where payroll_group_cd = '" & v_strPayrollCode & "'"
    l_strSQL = l_strSQL & " and pay_per_start_date <= '" & Format(v_datDate, "mm/dd/yyyy") & "'"
    l_strSQL = l_strSQL & " order by pay_per_start_date desc "

    l_objRS.CursorLocation = adUseClient
    'Must be adLockBatchOptimistic because updated in busCollectionPeriod
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetCollectionPeriodStatusByPayGroup = l_objRS
    
    Set l_objRS = Nothing
    

    Exit Function

GetCollectionPeriodStatusByPayGroup_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    
    TrsRaiseError m_strCMODULENAME, "GetCollectionPeriodStatusByPayGroup", , , , DescArgsFrom(l_strSQL), True
   
End Function

Public Function GetCollPerPayProcInd( _
            ByVal v_datDate As Date, _
            ByVal v_strPayrollCode As String) As String
    
    GetCollPerPayProcInd = _
        PayProcessedIndicator(v_strPayrollCode, v_datDate)
    Exit Function
    
GetCollPerPayProcInd_EH:
    TrsRaiseError m_strCMODULENAME, "GetCollPerPayProcInd", , , , , True

End Function

Public Function GetPayProcIndInRange( _
            ByVal v_strPayrollGroupCode As String, _
            ByVal v_datChargeDate As Date) As String
            
    GetPayProcIndInRange = _
        PayProcessedIndicator(v_strPayrollGroupCode, v_datChargeDate)
    Exit Function
    
GetPayProcIndInRange_EH:
    TrsRaiseError m_strCMODULENAME, "GetPayProcIndInRange", , , , , True

End Function

'Public Function GetCollPerPayProcInd(ByVal v_datDate As Date, ByVal v_strPayrollCode As String) As String
'
'
'    Dim l_strSQL As String
'    Dim l_objCN As ADODB.Connection
'    Dim l_objRS As ADODB.Recordset
'    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
'
'    On Error GoTo GetCollPerPayProcInd_EH
'
'
'    Set l_objCM = New util_DBConnectMgr.clsuDBConnectMgr
'    Set l_objCN = New ADODB.Connection
'    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
'
'    Set l_objCM = Nothing
'
'    l_objCN.Open
'
'    l_strSQL = "SELECT PAY_PROCESSED_IND "
'    l_strSQL = l_strSQL & "From METHDBA.TTR_COLL_PERIODS "
'    l_strSQL = l_strSQL & "where payroll_group_cd = '" & v_strPayrollCode & "'"
'    l_strSQL = l_strSQL & " and pay_per_start_date <= '" & Format(v_datDate, "mm/dd/yyyy") & "'"
'
'    Set l_objRS = New ADODB.Recordset
'
'    l_objRS.CursorLocation = adUseClient
'    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic
'
'    l_objRS.ActiveConnection = Nothing
'    l_objCN.Close
'
'    Set l_objCN = Nothing
'
'    If l_objRS.BOF And l_objRS.EOF Then
'        GetCollPerPayProcInd = ""
'    Else
'        GetCollPerPayProcInd = l_objRS.Fields(0)
'    End If
'
'    Set l_objRS = Nothing
'
'    Exit Function
'
'GetCollPerPayProcInd_EH:
'    Set l_objCM = Nothing
'    Set l_objCN = Nothing
'    Set l_objRS = Nothing
'    TrsRaiseError m_strCMODULENAME, "GetCollPerPayProcInd", , , , DescArgsFrom(l_strSQL), True
'
'End Function
'
'Public Function GetPayProcIndInRange( _
'            ByVal v_strPayrollGroupCode As String, _
'            ByVal v_datChargeDate As Date _
'            ) As String
'
'
'    Dim l_strSQL As String
'    Dim l_objCN As ADODB.Connection
'    Dim l_objRS As ADODB.Recordset
'    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
'
'    On Error GoTo GetPayProcIndInRange_EH
'
'
'    Set l_objCM = New util_DBConnectMgr.clsuDBConnectMgr
'    Set l_objCN = New ADODB.Connection
'    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
'
'    Set l_objCM = Nothing
'
'    l_objCN.Open
'
'    l_strSQL = "SELECT PAY_PROCESSED_IND "
'    l_strSQL = l_strSQL & "FROM METHDBA.TTR_COLL_PERIODS "
'    l_strSQL = l_strSQL & "WHERE PAYROLL_GROUP_CD = '" & v_strPayrollGroupCode & "' "
'    l_strSQL = l_strSQL & "AND '" & Format(v_datChargeDate, "MM/DD/YYYY") & "' BETWEEN PAY_PER_START_DATE AND PAY_PER_END_DATE "
'
'    Set l_objRS = New ADODB.Recordset
'
'    l_objRS.CursorLocation = adUseClient
'    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockBatchOptimistic
'
'    l_objRS.ActiveConnection = Nothing
'    l_objCN.Close
'
'    Set l_objCN = Nothing
'
'    If l_objRS.BOF And l_objRS.EOF Then
'        GetPayProcIndInRange = ""
'    Else
'        GetPayProcIndInRange = l_objRS.Fields(0)
'    End If
'
'    Set l_objRS = Nothing
'
'    Exit Function
'
'GetPayProcIndInRange_EH:
'    Set l_objCM = Nothing
'    Set l_objCN = Nothing
'    Set l_objRS = Nothing
'    TrsRaiseError m_strCMODULENAME, "GetPayProcIndInRange", , , , DescArgsFrom(l_strSQL), True
'
'End Function

Public Function GetCollectionPeriodLockStatus( _
            ByVal v_strDeptNum As String, _
            ByVal v_strPrevDeptNum As String, _
            ByVal v_datStartDate As Date, _
            ByVal v_datEndDate As Date _
            ) As ADODB.Recordset

    
    Dim l_strSQL As String
    Dim l_objCN As ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr

On Error GoTo GetCollectionPeriodLockStatus_EH
        
    
    Set l_objCM = New util_DBConnectMgr.clsuDBConnectMgr
    Set l_objCN = New ADODB.Connection
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open
    
    l_strSQL = "SELECT  DEPT_NUM, PAY_PROCESSED_IND, LOCK_STATUS, PAY_PER_START_DATE, PAY_PER_END_DATE, A.PAYROLL_GROUP_CD "
    l_strSQL = l_strSQL & "FROM    METHDBA.TTR_COLL_PERIODS  A,  METHDBA.TTR_DEPT B "
    l_strSQL = l_strSQL & "Where A.PAYROLL_GROUP_CD = B.PAYROLL_GROUP_CD "
    l_strSQL = l_strSQL & "AND     B.DEPT_NUM IN ('" & v_strDeptNum & "','" & v_strPrevDeptNum & "') "
    l_strSQL = l_strSQL & "AND     PAY_PER_END_DATE  >= '" & Format(v_datStartDate, "mm/dd/yyyy") & "' "
    l_strSQL = l_strSQL & "AND     PAY_PER_START_DATE <= '" & Format(v_datEndDate, "mm/dd/yyyy") & "' "
    l_strSQL = l_strSQL & "ORDER BY DEPT_NUM, PAY_PER_START_DATE "

    Set l_objRS = New ADODB.Recordset
    
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockReadOnly
    
    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing
    
    Set GetCollectionPeriodLockStatus = l_objRS
    
    Set l_objRS = Nothing
    

Exit Function
GetCollectionPeriodLockStatus_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    
    TrsRaiseError m_strCMODULENAME, "GetCollectionPeriodLockStatus", , , , DescArgsFrom(l_strSQL), True

End Function

Public Function GetCollectionPeriods( _
        ByVal v_datStartDate As Date, _
        ByVal v_datEndDate As Date, _
        ByVal v_strPayrollCodes As Collection _
        ) As ADODB.Recordset
    
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_strInClause As String
    Dim l_I, l_Count As Long
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    
    On Error GoTo GetCollectionPeriods_EH
    
    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
   
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open

'Buliding In Clause
    If v_strPayrollCodes Is Nothing Then
        l_strInClause = ""
    Else
        l_Count = v_strPayrollCodes.Count
        If l_Count > 0 Then
            l_strInClause = "and PAYROLL_GROUP_CD IN ( "
            For l_I = 1 To l_Count
                If l_I = 1 Then
                    l_strInClause = l_strInClause & "'" & v_strPayrollCodes.Item(l_I) & "' "
                Else
                    l_strInClause = l_strInClause & ",'" & v_strPayrollCodes.Item(l_I) & "' "
                End If
            Next l_I
            l_strInClause = l_strInClause & ")"
        Else
            l_strInClause = " "
        End If
    End If
    
    l_strSQL = "SELECT PAY_PER_START_DATE, PAY_PER_END_DATE, "
    l_strSQL = l_strSQL & " PAY_PROCESSED_IND, PAYROLL_GROUP_CD "
    l_strSQL = l_strSQL & "From METHDBA.TTR_COLL_PERIODS "
    l_strSQL = l_strSQL & "where pay_per_start_date <= '" & Format(v_datEndDate, "mm/dd/yyyy") & "'"
    l_strSQL = l_strSQL & " and pay_per_end_date >= '" & Format(v_datStartDate, "mm/dd/yyyy") & "'"
    l_strSQL = l_strSQL & l_strInClause
    l_strSQL = l_strSQL & " order by payroll_group_cd,  pay_per_start_date desc "

    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockReadOnly

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetCollectionPeriods = l_objRS
    
    Set l_objRS = Nothing

    Exit Function

GetCollectionPeriods_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    
    TrsRaiseError m_strCMODULENAME, "GetCollectionPeriods", , , , DescArgsFrom(l_strSQL), True
   
End Function

Public Function GetCollectionPeriodsByDepartment( _
        ByVal v_strDeptNum As String, _
        ByVal v_datStartDate As Date, _
        ByVal v_datEndDate As Date _
        ) As ADODB.Recordset
    
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_I, l_Count As Long
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    
    On Error GoTo GetCollectionPeriodsByDepartment_EH
    
    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
   
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open
   
    l_strSQL = "SELECT PAY_PER_START_DATE, PAY_PER_END_DATE, PAY_PROCESSED_IND, A.PAYROLL_GROUP_CD "
    l_strSQL = l_strSQL & "From METHDBA.TTR_COLL_PERIODS A "
    l_strSQL = l_strSQL & "   , METHDBA.TTR_DEPT B "
    l_strSQL = l_strSQL & "where dept_num = '" & v_strDeptNum & "' "
    l_strSQL = l_strSQL & "  and A.PAYROLL_GROUP_CD = B.PAYROLL_GROUP_CD "
    l_strSQL = l_strSQL & "  and pay_per_start_date <= '" & Format(v_datEndDate, "mm/dd/yyyy") & "'"
    l_strSQL = l_strSQL & "  and pay_per_end_date >= '" & Format(v_datStartDate, "mm/dd/yyyy") & "'"
    l_strSQL = l_strSQL & " order by pay_per_start_date desc "

    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockReadOnly

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetCollectionPeriodsByDepartment = l_objRS
    
    Set l_objRS = Nothing
    
    Exit Function

GetCollectionPeriodsByDepartment_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    
    TrsRaiseError m_strCMODULENAME, "GetCollectionPeriodsByDepartment", , , , DescArgsFrom(l_strSQL), True
   
End Function

Public Function GetCollectionPeriodsByWorkGroup( _
        ByVal v_lngWorkGroup As Long, _
        ByVal v_datStartDate As Date, _
        ByVal v_datEndDate As Date _
        ) As ADODB.Recordset
    
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_I, l_Count As Long
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    
    On Error GoTo GetCollectionPeriodsByWorkGroup_EH
    
    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
   
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open
   
    l_strSQL = "SELECT PAY_PER_START_DATE, PAY_PER_END_DATE, PAY_PROCESSED_IND, A.PAYROLL_GROUP_CD "
    l_strSQL = l_strSQL & "From METHDBA.TTR_COLL_PERIODS A "
    l_strSQL = l_strSQL & "   , METHDBA.TTR_DEPT B "
    l_strSQL = l_strSQL & "   , METHDBA.TTR_ASSOCIATIONS C "
    l_strSQL = l_strSQL & "where wg_num = " & str(v_lngWorkGroup)
    l_strSQL = l_strSQL & "  and B.dept_num = C.dept_num "
    l_strSQL = l_strSQL & "  and A.PAYROLL_GROUP_CD = B.PAYROLL_GROUP_CD "
    l_strSQL = l_strSQL & "  and pay_per_start_date <= '" & Format(v_datEndDate, "mm/dd/yyyy") & "'"
    l_strSQL = l_strSQL & "  and pay_per_end_date >= '" & Format(v_datStartDate, "mm/dd/yyyy") & "'"
    l_strSQL = l_strSQL & " order by pay_per_start_date desc "

'    If v_lngWorkGroup > 9999990 Then
'        l_strSQL = "SELECT PAY_PER_START_DATE, PAY_PER_END_DATE, PAY_PROCESSED_IND, A.PAYROLL_GROUP_CD "
'        l_strSQL = l_strSQL & "From METHDBA.TTR_COLL_PERIODS A "
'        l_strSQL = l_strSQL & "where PAYROLL_GROUP_CD = 'E'"
'        l_strSQL = l_strSQL & "  and pay_per_start_date <= '" & Format(v_datEndDate, "mm/dd/yyyy") & "'"
'        l_strSQL = l_strSQL & "  and pay_per_end_date >= '" & Format(v_datStartDate, "mm/dd/yyyy") & "'"
'        l_strSQL = l_strSQL & " order by pay_per_start_date desc "
'    End If
    
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockReadOnly

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetCollectionPeriodsByWorkGroup = l_objRS
    
    Set l_objRS = Nothing

    Exit Function

GetCollectionPeriodsByWorkGroup_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    
    TrsRaiseError m_strCMODULENAME, "GetCollectionPeriodsByWorkGroup", , , , DescArgsFrom(l_strSQL), True
   
End Function

Public Function GetLatestProcessedPayPeriodEndDate(ByVal v_strPayrollGroupCode As String) As Date
    
    
    Dim l_strSQL As String
    Dim l_objCN As ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    
    On Error GoTo GetLatestProcessedPayPeriodEndDate_EH
    
    
    Set l_objCM = New util_DBConnectMgr.clsuDBConnectMgr
    Set l_objCN = New ADODB.Connection
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open
    
    l_strSQL = "SELECT MAX(PAY_PER_END_DATE) "
    l_strSQL = l_strSQL & "FROM METHDBA.TTR_COLL_PERIODS "
    l_strSQL = l_strSQL & "WHERE PAYROLL_GROUP_CD = '" & v_strPayrollGroupCode & "'"
    l_strSQL = l_strSQL & " AND PAY_PROCESSED_IND = 'Y' "
    l_strSQL = l_strSQL & " AND PAY_PROCESS_DATE <= '" & Format(Date, "mm/dd/yyyy") & "'"


    Set l_objRS = New ADODB.Recordset
    
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockReadOnly
    
    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing
    
    If l_objRS.BOF And l_objRS.EOF Then
        GetLatestProcessedPayPeriodEndDate = CDate("0100-01-01")
    ElseIf IsNull(l_objRS.Fields(0)) Then
        GetLatestProcessedPayPeriodEndDate = CDate("0100-01-01")
    Else
        GetLatestProcessedPayPeriodEndDate = l_objRS.Fields(0)
    End If
  
    Set l_objRS = Nothing
    
    Exit Function
    
GetLatestProcessedPayPeriodEndDate_EH:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    
    TrsRaiseError m_strCMODULENAME, "GetLatestProcessedPayPeriodEndDate", , , , DescArgsFrom(l_strSQL), True

End Function

'*********************************************************
' Begin Public getters
'*********************************************************
Public Function PayPeriodStartDate( _
            ByVal v_strPayrollGrpCd As String, _
            ByVal v_datChgDate As Date) As Variant
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
    
On Error GoTo eh
    l_varRecord = _
        GetEffectiveRecord(v_strPayrollGrpCd, v_datChgDate)
    If IsEmpty(l_varRecord) Then
        Exit Function
    End If
    PayPeriodStartDate = l_varRecord(CCP_PAY_PER_START_DATE)
    
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "PayPeriodStartDate"
End Function

Public Function PayPeriodEndDate( _
            ByVal v_strPayrollGrpCd As String, _
            ByVal v_datChgDate As Date) As Variant
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
    
On Error GoTo eh
    l_varRecord = _
        GetEffectiveRecord(v_strPayrollGrpCd, v_datChgDate)
    If IsEmpty(l_varRecord) Then
        Exit Function
    End If
    PayPeriodEndDate = l_varRecord(CCP_PAY_PER_END_DATE)

    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "PayPeriodEndDate"
End Function

Public Function PayProcessedIndicator( _
            ByVal v_strPayrollGrpCd As String, _
            ByVal v_datChgDate As Date) As Variant
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
    
On Error GoTo eh
    l_varRecord = _
        GetEffectiveRecord(v_strPayrollGrpCd, v_datChgDate)
    If IsEmpty(l_varRecord) Then
        Exit Function
    End If
    PayProcessedIndicator = l_varRecord(CCP_PAY_PROCESSED_IND)
    
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "PayProcessedIndicator"
End Function

'*********************************************************
' End Public getters
'*********************************************************


'*********************************************************
' Begin Private procedures
'*********************************************************
Private Function GetEffectiveRecord( _
            ByVal v_strPayrollGrpCd As String, _
            ByVal v_datChgDate As Date) As Variant
    Dim l_varRecord As Variant
    Dim l_varRowValues As Variant
    Dim l_varEffRecord(2) As Variant
    Dim l_intRecIndex As Integer
    Dim lblnExists As Boolean
    
On Error GoTo eh
    l_varRecord = _
        GetRecord(v_strPayrollGrpCd, CCP_GROUPNAME, lblnExists)
    If IsEmpty(l_varRecord) Then
        GetEffectiveRecord = Empty
        Exit Function
    End If

    l_varRowValues = l_varRecord(CCP_ROWVALUES)
    l_intRecIndex = GetRecordIndex(v_datChgDate, l_varRowValues)
    If l_intRecIndex = -1 Then
        GetEffectiveRecord = Empty
    Else
        l_varEffRecord(CCP_PAY_PER_START_DATE) = _
            l_varRowValues(CCP_PAY_PER_START_DATE, l_intRecIndex)
        l_varEffRecord(CCP_PAY_PER_END_DATE) = _
            l_varRowValues(CCP_PAY_PER_END_DATE, l_intRecIndex)
        l_varEffRecord(CCP_PAY_PROCESSED_IND) = _
            l_varRowValues(CCP_PAY_PROCESSED_IND, l_intRecIndex)
        GetEffectiveRecord = l_varEffRecord
    End If
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "GetEffectiveRecord"

End Function

Private Function GetRecordIndex( _
            ByVal v_datEffDate As Date, _
            l_varRowValues As Variant) As Integer
' Return the record index v_datEffDate is between; or -1.
' l_varRowValues is a 2 dim array sorted in PAY_PER_START_DATE
' descending order.
    Dim i As Integer

On Error GoTo eh

    If IsEmpty(l_varRowValues) Then
        GetRecordIndex = -1
        Exit Function
    End If
    
    For i = LBound(l_varRowValues, 2) To UBound(l_varRowValues, 2)
        If v_datEffDate >= l_varRowValues(CCP_PAY_PER_START_DATE, i) And _
            v_datEffDate <= l_varRowValues(CCP_PAY_PER_END_DATE, i) Then
            GetRecordIndex = i
            Exit For
        End If
    Next

    If i > UBound(l_varRowValues, 2) Then
        GetRecordIndex = -1
    End If
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "GetRecordIndex"
End Function

'*********************************************************
' Begin Public Cache management methods
'*********************************************************

Public Function GetGroupRecords( _
                    ByVal v_strGroupName As String, _
                    ByVal v_varKeys As Variant, _
                    ByRef r_varKeyArray As Variant, _
                    ByRef r_varColumnArray As Variant) As ADODB.Recordset
    Dim l_rsRecords As ADODB.Recordset
    Dim l_colErrArgs As Collection
On Error GoTo eh

    Select Case v_strGroupName
        Case CCP_GROUPNAME
            Set l_rsRecords = GetCollectionPeriodCacheRecords(v_varKeys)
            r_varKeyArray = Array("PAYROLL_GROUP_CD")
            r_varColumnArray = Array("PAY_PER_START_DATE", "PAY_PER_END_DATE", "PAY_PROCESSED_IND")
    End Select
    Set GetGroupRecords = l_rsRecords

    Set l_rsRecords = Nothing
    Exit Function
eh:
    Set l_colErrArgs = New Collection
    l_colErrArgs.Add "v_strGroupName = " & v_strGroupName
    l_colErrArgs.Add "v_varKeys = " & v_varKeys
    l_colErrArgs.Add "l_rsRecords.Source = " & l_rsRecords.Source
    
    Set l_rsRecords = Nothing
    TrsRaiseError m_strCMODULENAME, "GetGroupRecords", , , , l_colErrArgs, True
End Function

Private Function GetCollectionPeriodCacheRecords(Optional ByVal v_varKeys As Variant) As ADODB.Recordset
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_strInClause As String
    Dim ldatStartDate As Date
    Dim ldatEndDate As Date

On Error GoTo eh
    
    If IsMissing(v_varKeys) Or IsEmpty(v_varKeys) Then
        l_strInClause = ""
    ElseIf IsArray(v_varKeys) Then
        l_strInClause = " AND " & BuildInClause("PAYROLL_GROUP_CD", v_varKeys, True)
    Else
        l_strInClause = " AND PAYROLL_GROUP_CD = '" & Trim(v_varKeys) & "'"
    End If

    ldatStartDate = DateAdd("yyyy", -1, Date)
    ldatEndDate = DateAdd("m", 1, Date)
    
    l_strSQL = "SELECT PAY_PER_START_DATE, PAY_PER_END_DATE, "
    l_strSQL = l_strSQL & " PAY_PROCESSED_IND, PAYROLL_GROUP_CD "
    l_strSQL = l_strSQL & "From METHDBA.TTR_COLL_PERIODS "
    l_strSQL = l_strSQL & "where pay_per_start_date <= '" & Format(ldatEndDate, "mm/dd/yyyy") & "'"
    l_strSQL = l_strSQL & " and pay_per_end_date >= '" & Format(ldatStartDate, "mm/dd/yyyy") & "'"
    l_strSQL = l_strSQL & l_strInClause
    l_strSQL = l_strSQL & " order by payroll_group_cd,  pay_per_start_date desc "
    
    Set l_objRS = GetCacheRecords(l_strSQL)
    
    Set GetCollectionPeriodCacheRecords = l_objRS
    
    Set l_objRS = Nothing
    
    Exit Function
eh:
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetCollectionPeriodCacheRecords", , , , DescArgsFrom(l_strSQL), True
End Function

'Private Function FormatCacheKey(ByVal v_lngWgNum As Long) As String
'    Dim l_lngLength As Long
'    Dim l_strInteger As String
'
'    l_lngLength = 11
'
'    l_strInteger = String(l_lngLength, "0")  '//-----PadLeft with zeros
'    l_strInteger = l_strInteger & Abs(Format(v_lngWgNum))
'    FormatCacheKey = IIf(v_lngWgNum < 0, "-", "0") & Right(l_strInteger, l_lngLength - 1)
'End Function

'*********************************************************
' Begin Private Caching methods
'*********************************************************

Private Function GetRecord( _
            ByVal v_varKey As Variant, _
            ByVal v_strGroupName As String, _
            ByRef r_blnExists As Boolean) As Variant
' Answer an variant array, representing a record, or empty.
' This method should not invoke InitializeGroup unless the
' group is setup for bulk initialization.  If setup for lazy
' initialization it should only get the single record and add it
' to the cache.  If the record is not found, using lazy, it should
' add a value to the cache to denote invalid so that it does not
' attempt to retrieve it from the database again.  Currently, a string
' constant is added for record not found and the spmCache component
' will answer that it exists, but the property value is a
' RECORDNOTFOUND string.
'
    Dim recCache As utilRecordCache.cuRecordCache
    Dim l_colErrArgs As Collection
           
On Error GoTo eh

'//--------------------------------------------------------------------------------------------------------------//
'//-----http://msdn.microsoft.com/library/default.asp?url=/library/en-us/cossdk/htm/imtxas_2bs6.asp
'//-----Do not use the SafeRef method in COM+. In MTS, SafeRef was used by an object to obtain
'//-----a reference to itself that could be passed outside its context wrapper.
'//-----With contexts and interception fully integrated into COM+, this is no longer necessary.
'//-----
'//-----[C/C++]
'//-----Note   The C++ function SafeRef should be used instead of this method.
'//--------------------------------------------------------------------------------------------------------------//

        Set recCache = CreateObject("utilRecordCache.cuRecordCache")
    
        If IsArray(v_varKey) Then
            GetRecord = recCache.GetRecord2(v_varKey, v_strGroupName, _
                                IsLazyInitialized(v_strGroupName), _
                                GroupSupportsLoadRecord(v_strGroupName), _
                                OneRowForEachKey(v_strGroupName), _
                                r_blnExists, Me)
        Else
            GetRecord = recCache.GetRecord(CStr(v_varKey), v_strGroupName, _
                                IsLazyInitialized(v_strGroupName), _
                                GroupSupportsLoadRecord(v_strGroupName), _
                                OneRowForEachKey(v_strGroupName), _
                                r_blnExists, Me)
        End If
        
        Set recCache = Nothing
    Exit Function
eh:
    Set recCache = Nothing

    Set l_colErrArgs = New Collection
    If IsArray(v_varKey) Then
        l_colErrArgs.Add "v_varKey() = (" & Join(v_varKey, ", ") & ")"
    Else
        l_colErrArgs.Add "v_varKey = " & CStr(v_varKey)
    End If
    l_colErrArgs.Add "v_strGroupName = " & v_strGroupName
    
    TrsRaiseError m_strCMODULENAME, "GetRecord", , , , l_colErrArgs, True
End Function

Public Sub InvalidateCollectionPeriodCache(ByVal v_strPayrollGrpCd As String)

On Error GoTo eh
   
    InvalidateCache Array(v_strPayrollGrpCd), CCP_GROUPNAME
    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "InvalidateRulesCache"
End Sub

'//-----code has been factored into utilRecordCache-----//
Private Sub InvalidateCache( _
                ByVal v_varKey As Variant, _
                ByVal v_strGroupName As String)

    Dim recCache As utilRecordCache.cuRecordCache
    Dim l_colErrArgs As Collection
    Dim l_blnExists As Boolean
On Error GoTo eh

    Set recCache = CreateObject("utilRecordCache.cuRecordCache")
    
    If IsArray(v_varKey) Then '//-----uses metadata to format key, may need to retrieve data to get metadata
        recCache.InvalidateCache2 v_varKey, v_strGroupName, _
                                IsLazyInitialized(v_strGroupName), _
                                GroupSupportsLoadRecord(v_strGroupName), _
                                OneRowForEachKey(v_strGroupName), _
                                l_blnExists, Me
    Else
        recCache.InvalidateCache CStr(v_varKey), v_strGroupName
    End If
        
    Set recCache = Nothing

    Exit Sub
eh:
    Set recCache = Nothing
    
    Set l_colErrArgs = New Collection
    If IsArray(v_varKey) Then
        l_colErrArgs.Add "v_varKey() = (" & Join(v_varKey, ", ") & ")"
    Else
        l_colErrArgs.Add "v_varKey = " & CStr(v_varKey)
    End If
    l_colErrArgs.Add "v_strGroupName = " & v_strGroupName
    
    TrsRaiseError m_strCMODULENAME, "InvalidateCache", , , , l_colErrArgs, True
End Sub

'//-----used by RecordCache-----//

Private Function IsLazyInitialized(ByVal v_strGroupName As String) As Boolean
' If group is not lazy initialized, then it is bulk initialized.
On Error GoTo eh
    Select Case Trim(v_strGroupName)
    Case CCP_GROUPNAME
        IsLazyInitialized = False
    Case Else
        IsLazyInitialized = True
    End Select
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "IsLazyInitialized"
End Function


Private Function GroupSupportsLoadRecord(ByVal v_strGroupName As String) As Boolean
'All group caches that support Invalidate must use parameterized SQL in corresponding GetxxxRecords (GetGroupRecords)
On Error GoTo eh
    Select Case Trim(v_strGroupName)
    Case CCP_GROUPNAME
        GroupSupportsLoadRecord = True
    Case Else
        GroupSupportsLoadRecord = False
    End Select
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "GroupSupportsLoadRecord"
End Function


Private Function OneRowForEachKey(ByVal v_strGroupName As String) As Boolean
'//-----normally each key/row is cached into a 1-dimensional array that represents each column-----//
'//-----union components will have multiple rows per key and will be cached into a 2-dimensional array-----//
On Error GoTo eh
    Select Case Trim(v_strGroupName)
    Case CCP_GROUPNAME
        OneRowForEachKey = False
    Case Else
        OneRowForEachKey = True
    End Select
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "OneRowForEachKey"
End Function

Private Function GetCacheRecords(ByVal v_strSQL As String, Optional eLockType As LockTypeEnum = adLockReadOnly) As ADODB.Recordset
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    Dim stTime As Long
    Dim ndTime As Long
    Dim dfTime As Long
    
On Error GoTo eh
    
    #If TRDEBUG Then
        stTime = GetTickCount()
    #End If
    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
    l_objCN.ConnectionString = l_objCM.GetConnectionString
    Set l_objCM = Nothing
    l_objCN.Open
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open v_strSQL, l_objCN, adOpenStatic, eLockType
    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    Set l_objCN = Nothing
    
    If l_objRS.BOF And l_objRS.EOF Then
        Set l_objRS = Nothing
    End If
    
    Set GetCacheRecords = l_objRS
    
    Set l_objRS = Nothing
    Set l_objCN = Nothing
    Set l_objCM = Nothing
    
    #If TRDEBUG Then
        ndTime = GetTickCount()
        dfTime = ndTime - stTime
        Debug.Print "GetCacheRecords: " & CStr(dfTime / 1000)
    #End If
    
    Exit Function

eh:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetCacheRecords", , , , DescArgsFrom(v_strSQL), True
End Function
Public Function GetCurrentCollectionPeriods( _
        ) As ADODB.Recordset
        
'This gets the first open pay period for each payroll group
'Also it returns pay periods in rewiew or in process
' for each payroll group
    
    
    Dim l_objCN As New ADODB.Connection
    Dim l_objRS As ADODB.Recordset
    Dim l_strSQL As String
    Dim l_I, l_Count As Long
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    
    On Error GoTo eh
    
    
    Set l_objCN = New ADODB.Connection
    Set l_objRS = New ADODB.Recordset
    Set l_objCM = New clsuDBConnectMgr
   
    l_objCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    
    Set l_objCM = Nothing
    
    l_objCN.Open
   
    l_strSQL = "select PAYROLL_GROUP_CD, MIN(PAY_PER_START_DATE) as PAY_PER_START_DATE,  "
    l_strSQL = l_strSQL & "MIN(PAY_PER_END_DATE) as PAY_PER_END_DATE,  "
    l_strSQL = l_strSQL & "PAY_PROCESSED_IND "
    l_strSQL = l_strSQL & "From METHDBA.TTR_COLL_PERIODS  "
    l_strSQL = l_strSQL & "where PAY_PROCESSED_IND <> 'Y' "
    l_strSQL = l_strSQL & "group by PAYROLL_GROUP_CD, PAY_PROCESSED_IND "
    
    l_objRS.CursorLocation = adUseClient
    l_objRS.Open l_strSQL, l_objCN, adOpenStatic, adLockReadOnly

    l_objRS.ActiveConnection = Nothing
    l_objCN.Close
    
    Set l_objCN = Nothing

    Set GetCurrentCollectionPeriods = l_objRS
    
    Set l_objRS = Nothing

    Exit Function

eh:
    Set l_objCM = Nothing
    Set l_objCN = Nothing
    Set l_objRS = Nothing
    
    TrsRaiseError m_strCMODULENAME, "GetCurrentCollectionPeriods", , , , DescArgsFrom(l_strSQL), True
   
End Function


