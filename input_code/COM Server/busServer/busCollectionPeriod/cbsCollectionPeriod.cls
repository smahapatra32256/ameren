VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cbsCollectionPeriod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const m_strCMODULENAME As String = "cbsCollectionPeriod"
Private Function GetPayrollGroup(ByVal v_strDepartment As String) As String
Dim l_objDepartment As dbDepartment.cdDepartment
On Error GoTo GetPayrollGroup_EH

    Set l_objDepartment = CreateObject("dbDepartment.cdDepartment")    'MTS standard
    GetPayrollGroup = l_objDepartment.GetPayrollGroup(v_strDepartment)
    Set l_objDepartment = Nothing

    Exit Function
GetPayrollGroup_EH:
    Set l_objDepartment = Nothing
    TrsRaiseError m_strCMODULENAME, "GetPayrollGroup"

End Function

Private Function GetDefaultEffectiveDate() As Date
    Dim l_obj As busEnvironment.cbsEnvironment

On Error GoTo eh

    Set l_obj = CreateObject("busEnvironment.cbsEnvironment")

    GetDefaultEffectiveDate = l_obj.GetDefaultEffectiveDate()
    Set l_obj = Nothing
    Exit Function
eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "GetDefaultEffectiveDate"
End Function

Public Function GetCollectionPeriodStatusByEmployee(ByVal v_lngEmpNum As Long, Optional ByVal v_datEffectiveDate As Date = "12/30/1899", Optional ByVal v_intMaxNoOfPayPeriods As Integer) As ADODB.Recordset
Dim l_objEmployee As dbEmployee.cdEmployee
Dim l_objPrevPayrollCd As dbDepartment.cdDepartment
Dim l_objCollectionPeriod As dbCollectionPeriod.cdCollectionPeriod

Dim l_objRS As ADODB.Recordset
Dim l_objRSforDeptNumPayTable As ADODB.Recordset
Dim l_objRSforPrevDeptNumPayTable As ADODB.Recordset
Dim l_objRSPayTable As New ADODB.Recordset

Dim l_datEffectiveDate As Date, l_intMaxNoOfPayPeriods As Integer
Dim l_strDeptNumPayGrp As String, l_strDeptNum As String
Dim l_datDeptNumEffectiveDate As Date, l_strPrevDepNum As String
Dim l_strPrevDeptNumPayGrp As String, l_DeptNumdattodate As Date
Dim l_intIndex As Integer, l_DeptNumdatfromdate As Date
Dim l_datDeptNumEndPrevDate As Date, l_intIndexNoOfPayPeriods As Integer
    Dim l_objEnvironment As busEnvironment.cbsEnvironment
        
On Error GoTo GetCollectionPeriodStatusByEmployee_EH

    Set l_objEmployee = CreateObject("dbEmployee.cdEmployee")

    If v_datEffectiveDate = CDate("12/30/1899") Then 'if no optional date is supplied
        l_datEffectiveDate = GetDefaultEffectiveDate() '//Date   'default to current date
    Else
        If IsDate(v_datEffectiveDate) Then
            l_datEffectiveDate = v_datEffectiveDate
        Else
            'error in date
            Exit Function
        End If
    End If
    
    If v_intMaxNoOfPayPeriods < 1 Then
        '//effective date = current date
        '//for testing: effective date could be 30 (30/14=2.14), which would then make l_intMaxNoOfPayPeriods 7, allowing for 2 extra pay periods (14 days)
        Set l_objEnvironment = CreateObject("busEnvironment.cbsEnvironment")
        
        l_intMaxNoOfPayPeriods = 5 + (l_objEnvironment.GetEffectiveDateOffset / 14) 'default
        
        Set l_objEnvironment = Nothing
    Else
        If v_intMaxNoOfPayPeriods > 500 Then
            l_intMaxNoOfPayPeriods = 500  'default
        Else
            l_intMaxNoOfPayPeriods = v_intMaxNoOfPayPeriods
        End If
    End If

    Set l_objRS = l_objEmployee.GetEmployee(v_lngEmpNum)

    If (l_objRS.EOF And l_objRS.BOF) Then
        Exit Function
    Else
        l_strDeptNum = l_objRS!DEPT_NUM
        l_strDeptNumPayGrp = l_objRS!PAYROLL_GROUP_CD
        l_strPrevDepNum = l_objRS!PREV_DEPT_NUM
        l_datDeptNumEffectiveDate = l_objRS!DEPT_EFF_DATE
    End If
    l_objRSPayTable.Fields.Append "PAY_PER_START_DATE", adDate
    l_objRSPayTable.Fields.Append "PAY_PER_END_DATE", adDate
    l_objRSPayTable.Fields.Append "PAY_PROCESSED_IND", adChar, 1
    l_objRSPayTable.Fields.Append "PAYROLL_GROUP_CD", adChar, 1
    
    l_objRSPayTable.Open

    Set l_objCollectionPeriod = CreateObject("dbCollectionPeriod.cdCollectionPeriod")
    Set l_objRSforDeptNumPayTable = l_objCollectionPeriod.GetCollectionPeriodStatusByPayGroup(l_datEffectiveDate, l_strDeptNumPayGrp)
    If (l_objRSforDeptNumPayTable.EOF And l_objRSforDeptNumPayTable.BOF) Then
        Set GetCollectionPeriodStatusByEmployee = l_objRSPayTable
        Set l_objRSPayTable = Nothing
        Set l_objEmployee = Nothing
        Set l_objPrevPayrollCd = Nothing
        Set l_objRS = Nothing
        Set l_objRSforDeptNumPayTable = Nothing
        Set l_objRSforPrevDeptNumPayTable = Nothing
        Set l_objCollectionPeriod = Nothing
        Exit Function
    End If

    l_DeptNumdattodate = l_objRSforDeptNumPayTable!PAY_PER_END_DATE      'l_objRSforDeptNumPayTable(1)
    l_DeptNumdatfromdate = l_objRSforDeptNumPayTable!PAY_PER_START_DATE  'l_objRSforDeptNumPayTable(0)

    For l_intIndex = 1 To l_intMaxNoOfPayPeriods - 1
        l_objRSforDeptNumPayTable.MoveNext
        If l_objRSforDeptNumPayTable.EOF Then
           Exit For
        Else
           l_DeptNumdatfromdate = l_objRSforDeptNumPayTable!PAY_PER_START_DATE  'l_objRSforDeptNumPayTable(0)
        End If
    Next
    
    l_objRSforDeptNumPayTable.MoveFirst
    If l_datDeptNumEffectiveDate >= l_DeptNumdatfromdate _
       And l_datDeptNumEffectiveDate <= l_DeptNumdattodate Then
       If Trim(l_strPrevDepNum) = "" Then
          Do Until l_objRSforDeptNumPayTable.EOF
             l_intIndexNoOfPayPeriods = l_intIndexNoOfPayPeriods + 1
             If l_intIndexNoOfPayPeriods > l_intMaxNoOfPayPeriods Then
                Exit Do
             End If
             UpdateCollectionPeriodStatusTable l_objRSforDeptNumPayTable, l_objRSPayTable
             l_objRSforDeptNumPayTable.MoveNext
          Loop 'return the deptnum pay period from table 1
       End If
        Set l_objPrevPayrollCd = CreateObject("dbDepartment.cdDepartment")    'MTS standard
       l_strPrevDeptNumPayGrp = l_objPrevPayrollCd.GetPayrollGroup(l_strPrevDepNum)
        If l_strPrevDeptNumPayGrp = l_strDeptNumPayGrp Then
            Do Until l_objRSforDeptNumPayTable.EOF 'same pay periods  process current dep num pay periods only
                l_intIndexNoOfPayPeriods = l_intIndexNoOfPayPeriods + 1
                If l_intIndexNoOfPayPeriods > l_intMaxNoOfPayPeriods Then
                    Exit Do
                End If
                UpdateCollectionPeriodStatusTable l_objRSforDeptNumPayTable, l_objRSPayTable
                l_objRSforDeptNumPayTable.MoveNext
            Loop 'build and return the deptnum pay period from table 1
        Else    'payroll group for this prev dept no <> dept no/process splits/       'get old prev dept pay periods table for splits
            Set l_objRSforPrevDeptNumPayTable = l_objCollectionPeriod.GetCollectionPeriodStatusByPayGroup(l_datDeptNumEffectiveDate, l_strPrevDeptNumPayGrp)
        End If
    Else  '        dept eff date outside the pay period date range, process current dep num pay periods only
         Do Until l_objRSforDeptNumPayTable.EOF
             l_intIndexNoOfPayPeriods = l_intIndexNoOfPayPeriods + 1
             If l_intIndexNoOfPayPeriods > l_intMaxNoOfPayPeriods Then
                Exit Do
             End If
             UpdateCollectionPeriodStatusTable l_objRSforDeptNumPayTable, l_objRSPayTable
             l_objRSforDeptNumPayTable.MoveNext
         Loop 'build the deptnum pay period from table 1
    End If
   
    If l_intIndexNoOfPayPeriods > l_intMaxNoOfPayPeriods Then
       'limit reached in l_objRSPayTable return table, bypass exit do
    Else
        Do Until l_objRSforDeptNumPayTable.EOF     '       process first pay period table (for splits) - overlay correct dep no effective date l_intSplitPeriod_flag = 1
            If l_datDeptNumEffectiveDate < l_objRSforDeptNumPayTable(0) Then
                l_intIndexNoOfPayPeriods = l_intIndexNoOfPayPeriods + 1
                If l_intIndexNoOfPayPeriods > l_intMaxNoOfPayPeriods Then
                    Exit Do
                End If
                UpdateCollectionPeriodStatusTable l_objRSforDeptNumPayTable, l_objRSPayTable
                l_objRSforDeptNumPayTable.MoveNext
            Else                                                        'normal
                l_datDeptNumEndPrevDate = l_datDeptNumEffectiveDate - 1  ' If l_datDeptNumEffectiveDate => l_objRSforDeptNumPayTable(0) Then
                l_objRSforDeptNumPayTable(0) = l_datDeptNumEffectiveDate 'use dept effective date for start date
                l_intIndexNoOfPayPeriods = l_intIndexNoOfPayPeriods + 1
                If l_intIndexNoOfPayPeriods > l_intMaxNoOfPayPeriods Then
                    Exit Do
                End If
                UpdateCollectionPeriodStatusTable l_objRSforDeptNumPayTable, l_objRSPayTable
                If l_objRSforPrevDeptNumPayTable Is Nothing Then 'if table 2 is nothing or 'limit reached in l_objRSPayTable return table
                Else    'process table 2 l_objRSforPrevDeptNumPayTable
                    If (l_objRSforPrevDeptNumPayTable.EOF And l_objRSforPrevDeptNumPayTable.BOF) Then
                    Else
                       l_objRSforPrevDeptNumPayTable.MoveFirst 'process second prev dept num pay period table 2 (for splits)
                       Do Until l_objRSforPrevDeptNumPayTable.EOF 'skip records on second table until match
                            If l_datDeptNumEndPrevDate < l_objRSforPrevDeptNumPayTable(0) Then
                               l_objRSforPrevDeptNumPayTable.MoveNext
                            Else
                               l_objRSforPrevDeptNumPayTable(1) = l_datDeptNumEndPrevDate 'use endtodate
                               Exit Do
                            End If
                       Loop
                       Do Until l_objRSforPrevDeptNumPayTable.EOF
                            l_intIndexNoOfPayPeriods = l_intIndexNoOfPayPeriods + 1
                            If l_intIndexNoOfPayPeriods > l_intMaxNoOfPayPeriods Then
                                Exit Do
                            End If
                            UpdateCollectionPeriodStatusTable l_objRSforPrevDeptNumPayTable, l_objRSPayTable
                            l_objRSforPrevDeptNumPayTable.MoveNext
                       Loop
                    End If
                End If
                Exit Do  'add "dept_effec_date > table1 start date period" entries to return recordset and overlay start date/l_intSplitPeriod_flag = 1
            End If
        Loop
    End If
            
    
    l_objRSPayTable.MoveFirst
    Set GetCollectionPeriodStatusByEmployee = l_objRSPayTable
    Set l_objRSPayTable = Nothing
    Set l_objEmployee = Nothing
    Set l_objPrevPayrollCd = Nothing
    Set l_objRS = Nothing
    Set l_objRSforDeptNumPayTable = Nothing
    Set l_objRSforPrevDeptNumPayTable = Nothing
    Set l_objCollectionPeriod = Nothing
    Set l_objEnvironment = Nothing
    
    Exit Function
GetCollectionPeriodStatusByEmployee_EH:
    Set l_objRSPayTable = Nothing
    Set l_objEmployee = Nothing
    Set l_objPrevPayrollCd = Nothing
    Set l_objRS = Nothing
    Set l_objRSforDeptNumPayTable = Nothing
    Set l_objRSforPrevDeptNumPayTable = Nothing
    Set l_objCollectionPeriod = Nothing
    Set l_objEnvironment = Nothing
    
    TrsRaiseError m_strCMODULENAME, "GetCollectionPeriodStatusByEmployee", , , , , True

End Function
Private Function UpdateCollectionPeriodStatusTable(r_objFromTable As ADODB.Recordset, r_objResultTable As ADODB.Recordset)
On Error GoTo UpdateCollectionPeriodStatusTable_EH
    
        r_objResultTable.AddNew
        r_objResultTable(0) = r_objFromTable(0)
        r_objResultTable(1) = r_objFromTable(1)
        r_objResultTable(2) = r_objFromTable(2)
        r_objResultTable(3) = r_objFromTable(3)
        r_objResultTable.Update

    Exit Function
UpdateCollectionPeriodStatusTable_EH:
    TrsRaiseError m_strCMODULENAME, "UpdateCollectionPeriodStatusTable", , , , , True

End Function
Public Function GetPayProcIndInRange(ByVal v_strPayrollGroupCode As String, ByVal v_datChargeDate As Date) As String
Dim l_objdbCollectionPeriod As dbCollectionPeriod.cdCollectionPeriod
On Error GoTo GetPayProcIndInRange_EH

    Set l_objdbCollectionPeriod = CreateObject("dbCollectionPeriod.cdCollectionPeriod")
    GetPayProcIndInRange = l_objdbCollectionPeriod.GetPayProcIndInRange(v_strPayrollGroupCode, v_datChargeDate)
    Set l_objdbCollectionPeriod = Nothing

    Exit Function
GetPayProcIndInRange_EH:
    Set l_objdbCollectionPeriod = Nothing
    TrsRaiseError m_strCMODULENAME, "GetPayProcIndInRange", , , , , True

End Function
Public Function GetCollectionPeriodLockStatus( _
            ByVal v_strDeptNum As String, _
            ByVal v_strPrevDeptNum As String, _
            ByVal v_datStartDate As Date, _
            ByVal v_datEndDate As Date _
            ) As ADODB.Recordset
Dim l_objdbCollectionPeriod As dbCollectionPeriod.cdCollectionPeriod
On Error GoTo GetCollectionPeriodLockStatus_EH

    Set l_objdbCollectionPeriod = CreateObject("dbCollectionPeriod.cdCollectionPeriod")
    Set GetCollectionPeriodLockStatus = l_objdbCollectionPeriod.GetCollectionPeriodLockStatus(v_strDeptNum, v_strPrevDeptNum, v_datStartDate, v_datEndDate)
    Set l_objdbCollectionPeriod = Nothing

    Exit Function
GetCollectionPeriodLockStatus_EH:
    Set l_objdbCollectionPeriod = Nothing
    TrsRaiseError m_strCMODULENAME, "GetCollectoinPeriodLockStatus", , , , , True

End Function
Public Function GetCollectionPeriodsByDepartment( _
        ByVal v_strDeptNum As String _
        ) As ADODB.Recordset
Dim l_objCollectionPeriod As dbCollectionPeriod.cdCollectionPeriod
Dim l_datStartDate As Date, l_datTodaysDate As Date
    Dim l_objEnv As busEnvironment.cbsEnvironment
On Error GoTo GetCollectionPeriodsByDepartment_EH

    Set l_objEnv = CreateObject("busEnvironment.cbsEnvironment")
    l_datTodaysDate = l_objEnv.GetDefaultEffectiveDate 'Date
    Set l_objEnv = Nothing
    l_datStartDate = DateAdd("m", -2, l_datTodaysDate)
    '//TODO: could also subtract the GetEffectiveDateOffset from the start date, but the client only shows the first 5 rows anyway
    '//l_datStartDate = DateAdd("d", -l_objEnv.GetEffectiveDateOffset, l_datStartDate)
    
    Set l_objCollectionPeriod = CreateObject("dbCollectionPeriod.cdCollectionPeriod")    'MTS standard
    Set GetCollectionPeriodsByDepartment = _
        l_objCollectionPeriod.GetCollectionPeriodsByDepartment(v_strDeptNum, l_datStartDate, l_datTodaysDate)
    Set l_objCollectionPeriod = Nothing
    Set l_objEnv = Nothing

    Exit Function
GetCollectionPeriodsByDepartment_EH:
    Set l_objCollectionPeriod = Nothing
    Set GetCollectionPeriodsByDepartment = Nothing
    Set l_objEnv = Nothing

    TrsRaiseError m_strCMODULENAME, "GetCollectionPeriodsByDepartment", , , , , True

End Function
Public Function GetCollectionPeriodsByWorkGroup( _
        ByVal v_lngWorkGroup As Long _
        ) As ADODB.Recordset
Dim l_objCollectionPeriod As dbCollectionPeriod.cdCollectionPeriod
Dim l_datStartDate As Date, l_datTodaysDate As Date
    Dim l_objEnv As busEnvironment.cbsEnvironment
On Error GoTo GetCollectionPeriodsByWorkGroup_EH

    Set l_objEnv = CreateObject("busEnvironment.cbsEnvironment")
    l_datTodaysDate = l_objEnv.GetDefaultEffectiveDate 'Date
    Set l_objEnv = Nothing
    l_datStartDate = DateAdd("m", -2, l_datTodaysDate)
    '//TODO: could also subtract the GetEffectiveDateOffset from the start date, but the client only shows the first 5 rows anyway
    '//l_datStartDate = DateAdd("d", -l_objEnv.GetEffectiveDateOffset, l_datStartDate)

    Set l_objCollectionPeriod = CreateObject("dbCollectionPeriod.cdCollectionPeriod")    'MTS standard
    Set GetCollectionPeriodsByWorkGroup = l_objCollectionPeriod.GetCollectionPeriodsByWorkGroup(v_lngWorkGroup, l_datStartDate, l_datTodaysDate)
    Set l_objCollectionPeriod = Nothing
    Set l_objEnv = Nothing
    
    Exit Function
GetCollectionPeriodsByWorkGroup_EH:
    Set l_objCollectionPeriod = Nothing
    Set GetCollectionPeriodsByWorkGroup = Nothing
    Set l_objEnv = Nothing
    TrsRaiseError m_strCMODULENAME, "GetCollectionPeriodsByWorkGroup", , , , , True
        
End Function
Public Function GetLatestProcessedPayPeriodEndDate(ByVal v_strPayrollGroupCode As String) As Date
Dim l_objCollectionPeriod As dbCollectionPeriod.cdCollectionPeriod
On Error GoTo GetLatestProcessedPayPeriodEndDate_EH
    
    Set l_objCollectionPeriod = CreateObject("dbCollectionPeriod.cdCollectionPeriod")    'MTS standard
    GetLatestProcessedPayPeriodEndDate = l_objCollectionPeriod.GetLatestProcessedPayPeriodEndDate(v_strPayrollGroupCode)
    Set l_objCollectionPeriod = Nothing
             
    Exit Function
GetLatestProcessedPayPeriodEndDate_EH:
    Set l_objCollectionPeriod = Nothing
    TrsRaiseError m_strCMODULENAME, "GetLatestProcessedPayPeriodEndDate", , , , , True

End Function
