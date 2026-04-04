VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "cbsSchedExc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const m_strCMODULENAME As String = "cbsSchedExc"
Private m_rsSchedule As ADODB.Recordset

'//for older TRIS clients
Public Function GetViewCalendarDetails( _
    ByVal v_lngEmpNum As Long, _
    ByVal v_datStartDate As Date, _
    ByVal v_datEndDate As Date, _
    Optional v_datLastClosedPPEndDate As Date = CDate("1980-01-01"), _
    Optional v_datLastDayOfOpenPP As Date = CDate("1980-01-01") _
    ) As ADODB.Recordset

    On Error GoTo eh
      
    Set GetViewCalendarDetails = GetViewCalendarDetailsPrivate(v_lngEmpNum, v_datStartDate, v_datEndDate, v_datLastClosedPPEndDate, v_datLastDayOfOpenPP, 39)
    
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "GetViewCalendarDetails", , , , , True
End Function

'//for newer TRIS clients DFE 48 chars, OTTC pick list description increased to 60 chars
Public Function GetViewCalendarDetails2B( _
    ByVal v_lngEmpNum As Long, _
    ByVal v_datStartDate As Date, _
    ByVal v_datEndDate As Date, _
    Optional v_datLastClosedPPEndDate As Date = CDate("1980-01-01"), _
    Optional v_datLastDayOfOpenPP As Date = CDate("1980-01-01"), _
    Optional ByVal v_intDescriptionLength As Integer = 48 _
    ) As ADODB.Recordset

    On Error GoTo eh
      
    Set GetViewCalendarDetails2B = GetViewCalendarDetailsPrivate(v_lngEmpNum, v_datStartDate, v_datEndDate, v_datLastClosedPPEndDate, v_datLastDayOfOpenPP, v_intDescriptionLength)
    
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "GetViewCalendarDetails2B", , , , , True
End Function


Private Function GetViewCalendarDetailsPrivate( _
    ByVal v_lngEmpNum As Long, _
    ByVal v_datStartDate As Date, _
    ByVal v_datEndDate As Date, _
    Optional v_datLastClosedPPEndDate As Date = CDate("1980-01-01"), _
    Optional v_datLastDayOfOpenPP As Date = CDate("1980-01-01"), _
    Optional ByVal v_intDescriptionLength As Integer = 48 _
    ) As ADODB.Recordset

    On Error GoTo eh

    Dim l_lngEmpNum As Long
    Dim l_datStart As Date
    Dim l_datEnd As Date
    Dim l_datLastPayEnd As Date
    Dim l_datLastDayOpenPP As Date
    Dim l_rsLockStatus As ADODB.Recordset
    Dim l_rsSchedule As ADODB.Recordset
    Dim l_rsActualTime As ADODB.Recordset
    Dim l_rsAdjustments As ADODB.Recordset
    Dim l_rsOpenActualTime As ADODB.Recordset
    Dim l_rsMerge As ADODB.Recordset
    Dim l_strFilter As String
    Dim l_colErrArgs As Collection
    
    l_lngEmpNum = v_lngEmpNum
    l_datStart = v_datStartDate
    l_datEnd = v_datEndDate
    
'''''    If v_datLastClosedPPEndDate = CDate("1980-01-01") _
'''''        Or v_datLastDayOfOpenPP = CDate("1980-01-01") Then
'''''        Set l_rsLockStatus = GetPayPeriodLockStatus(l_lngEmpNum)
'''''        l_datLastPayEnd = GetLastClosedPayPeriodEndDate(l_rsLockStatus)
'''''        If l_datLastPayEnd < l_datStart Then
'''''            l_datLastPayEnd = l_datStart
'''''        End If
'''''        l_rsLockStatus.Filter = adFilterNone
'''''        l_rsLockStatus.Sort = ""
'''''        If RecordsExist(l_rsLockStatus) Then
'''''            l_datLastDayOpenPP = GetCurrentPayPeriodEndDate(l_rsLockStatus)
'''''        Else
'''''            l_datLastDayOpenPP = l_datEnd
'''''        End If
'''''    Else
'''''        l_datLastPayEnd = v_datLastClosedPPEndDate
'''''        l_datLastDayOpenPP = v_datLastDayOfOpenPP
'''''        If l_datLastPayEnd < l_datStart Then
'''''            l_datLastPayEnd = l_datStart
'''''        End If
'''''    End If
'''''
'''''    If l_datLastPayEnd > l_datEnd Then
'''''        l_datLastPayEnd = l_datEnd
'''''    End If
'''''
'''''    If l_datLastDayOpenPP > l_datEnd Then
'''''        l_datLastDayOpenPP = l_datEnd
'''''    End If
    
    If v_datLastClosedPPEndDate < l_datStart Then
        l_datLastPayEnd = l_datStart
    Else
        l_datLastPayEnd = v_datLastClosedPPEndDate
    End If
    If v_datLastDayOfOpenPP < l_datStart Then
        l_datLastDayOpenPP = v_datEndDate
    Else
        l_datLastDayOpenPP = v_datLastDayOfOpenPP
    End If
    Set l_rsAdjustments = GetAdjustmentDetails(l_lngEmpNum, l_datStart, l_datEnd) 'l_datLastPayEnd)
    Set l_rsActualTime = GetActualTime(l_lngEmpNum, l_datStart, l_datLastDayOpenPP, v_intDescriptionLength)
    If l_datLastPayEnd < l_datEnd Then
        If l_datLastPayEnd > l_datStart Then
            Set l_rsSchedule = GetSched(l_lngEmpNum, DateAdd("d", 1, l_datLastPayEnd), l_datEnd, trAdjBaseSchedWithExceptionsAndOvertime)
        Else
            Set l_rsSchedule = GetSched(l_lngEmpNum, l_datStart, l_datEnd, trAdjBaseSchedWithExceptionsAndOvertime)
        End If
    Else
'        Set l_rsSchedule = Nothing
         Set l_rsSchedule = GetSched(l_lngEmpNum, l_datStart, l_datEnd, trAdjBaseSchedWithExceptionsAndOvertime)
    End If
    If RecordsExist(l_rsActualTime) Then
        If l_datLastPayEnd = l_datStart Then
            l_strFilter = "CHARGE_DATE >= " & str(l_datLastPayEnd) & _
                " AND CHARGE_DATE <= " & str(l_datLastDayOpenPP)
        Else
            l_strFilter = "CHARGE_DATE > " & str(l_datLastPayEnd) & _
                " AND CHARGE_DATE <= " & str(l_datLastDayOpenPP)
        End If
        l_strFilter = MDAC25FilterFix(l_strFilter)
        l_rsActualTime.Filter = l_strFilter
        Set l_rsOpenActualTime = CopyRecordset(l_rsActualTime)
    Else
        Set l_rsOpenActualTime = l_rsActualTime
    End If
    l_rsActualTime.Filter = ""
    Set l_rsMerge = BuildMergeRS(l_datStart, l_datEnd, _
        l_datLastPayEnd, l_datLastDayOpenPP, l_rsAdjustments, _
        l_rsActualTime, l_rsSchedule, l_rsOpenActualTime)
        
    Set GetViewCalendarDetailsPrivate = l_rsMerge
    Set l_rsMerge = Nothing
    Set l_rsActualTime = Nothing
    Set l_rsOpenActualTime = Nothing
    Set l_rsSchedule = Nothing
    Set l_rsAdjustments = Nothing
    Set l_rsLockStatus = Nothing
    Exit Function
    
eh:
    Set l_rsMerge = Nothing
    Set l_rsActualTime = Nothing
    Set l_rsOpenActualTime = Nothing
    Set l_rsSchedule = Nothing
    Set l_rsAdjustments = Nothing
    Set l_rsLockStatus = Nothing
    Set l_colErrArgs = New Collection
    l_colErrArgs.Add "GetViewCalendarDetailsPrivate argument values"
    l_colErrArgs.Add "v_lngEmpNum=" & CStr(v_lngEmpNum)
    l_colErrArgs.Add "v_datStartDate=" & Format(v_datStartDate, "mm/dd/yyyy")
    l_colErrArgs.Add "v_datEndDate=" & Format(v_datEndDate, "mm/dd/yyyy")
    l_colErrArgs.Add "v_datLastClosedPPEndDate=" & Format(v_datLastClosedPPEndDate, "mm/dd/yyyy")
    l_colErrArgs.Add "v_datLastDayOfOpenPP=" & Format(v_datLastDayOfOpenPP, "mm/dd/yyyy")
    TrsRaiseError m_strCMODULENAME, "GetViewCalendarDetailsPrivate", , , , l_colErrArgs, True
End Function


Private Function CreateDaysOffRS() As ADODB.Recordset

    Dim l_rs As ADODB.Recordset
    
    On Error GoTo eh
    
    Set l_rs = New ADODB.Recordset

    With l_rs
      .Fields.Append "EMP_NUM", adInteger, , adFldUpdatable
      .Fields.Append "CHARGE_DATE", adDate, , adFldUpdatable
      .Fields.Append "NO_SCHED_IND", adChar, 1, adFldUpdatable
      .CursorLocation = adUseClient
      .CursorType = adOpenDynamic
      .LockType = adLockOptimistic
      .Open
    End With
    
    Set CreateDaysOffRS = l_rs
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "CreateDaysOffRS"

End Function


Private Function GetPayPeriodLockStatus( _
            ByVal v_lngEmpNum As Long _
            ) As ADODB.Recordset
            
    Dim l_obj As busCollectionPeriod.cbsCollectionPeriod
    Dim l_objEnv As busEnvironment.cbsEnvironment
    
    On Error GoTo eh
    
    Set l_obj = CreateObject("busCollectionPeriod.cbsCollectionPeriod")
    Set l_objEnv = CreateObject("busEnvironment.cbsEnvironment")
    
    Set GetPayPeriodLockStatus = l_obj.GetCollectionPeriodStatusByEmployee(v_lngEmpNum, l_objEnv.GetDefaultEffectiveDate)

    Set l_obj = Nothing
    Set l_objEnv = Nothing
    
    Exit Function
eh:
    Set l_obj = Nothing
    Set l_objEnv = Nothing
    TrsRaiseError m_strCMODULENAME, "GetPayPeriodLockStatus"
                       
End Function
'''''Private Function GetEmpBanks( _
'''''            ByVal v_aryEmployees As Variant _
'''''            ) As ADODB.Recordset
'''''
'''''    Dim l_obj As dbAccrualBankMgr.cdAccrualBankMgr
'''''
'''''
'''''    Set l_obj = New dbAccrualBankMgr.cdAccrualBankMgr
'''''
'''''    Set GetEmpBanks = _
'''''        l_obj.GetEmployeeBanks2( _
'''''            v_aryEmployees, Date)
'''''
'''''    Set l_obj = Nothing
'''''
'''''End Function

Private Function GetLastClosedPayPeriodEndDate( _
            ByVal v_rsCollectionPeriods As ADODB.Recordset _
            ) As Date
            
    Dim l_strPayGroup As String
    Dim l_datLastPayEnd As Date
            
    On Error GoTo eh
    
    v_rsCollectionPeriods.Filter = "PAY_PROCESSED_IND = 'Y'"
    v_rsCollectionPeriods.Sort = "PAY_PER_END_DATE DESC"
    
    If v_rsCollectionPeriods.EOF And v_rsCollectionPeriods.BOF Then
        l_datLastPayEnd = CDate("1980-01-01")
    Else
        l_datLastPayEnd = v_rsCollectionPeriods!PAY_PER_END_DATE
    End If
    GetLastClosedPayPeriodEndDate = l_datLastPayEnd
    
    Exit Function
        
eh:
    TrsRaiseError m_strCMODULENAME, "GetLastClosedPayPeriodEndDate"
                       
End Function
Private Function GetCurrentPayPeriodEndDate( _
            ByVal v_rsCollectionPeriods As ADODB.Recordset _
            ) As Date
            
    Dim l_strPayGroup As String
    Dim l_datLastPayEnd As Date
            
    On Error GoTo eh
    
    v_rsCollectionPeriods.Filter = ""
    v_rsCollectionPeriods.Sort = "PAY_PER_END_DATE DESC"
    
    l_datLastPayEnd = v_rsCollectionPeriods!PAY_PER_END_DATE
    GetCurrentPayPeriodEndDate = l_datLastPayEnd
        
    Exit Function
        
eh:
    TrsRaiseError m_strCMODULENAME, "GetCurrentPayPeriodEndDate"
                       
End Function

Private Function GetActualTime( _
            ByVal v_aryEmployees As Variant, _
            ByVal v_datStart As Date, _
            ByVal v_datEnd As Date, _
            ByVal v_intDescriptionLength As Integer _
            ) As ADODB.Recordset
            
    Dim l_obj As busTimecard.cbsTimecard
    Dim l_ary As Variant
    
    On Error GoTo eh
    
    If IsArray(v_aryEmployees) Then
        l_ary = v_aryEmployees
    Else
        l_ary = Array(v_aryEmployees)
    End If
    Set l_obj = CreateObject("busTimecard.cbsTimecard")
    
    Set GetActualTime = l_obj.GetActualTimeDetailWithOtherPay2B( _
            l_ary, v_datStart, v_datEnd, v_intDescriptionLength)
            
    Set l_obj = Nothing
    Exit Function
        
eh:
    TrsRaiseError m_strCMODULENAME, "GetActualTime"
            
            
End Function

Private Function GetSched( _
            ByVal v_aryEmployees As Variant, _
            ByVal v_datStart As Date, _
            ByVal v_datEnd As Date, _
            Optional v_intReturnOption As TRSReturnOption = trAdjustedBaseSched _
            ) As ADODB.Recordset
            
    Dim l_obj As busSchedule.cbsSchedule
    
    On Error GoTo eh
    
    Set l_obj = CreateObject("busSchedule.cbsSchedule")
    
    Set GetSched = l_obj.GetAdjustedBaseSchedule( _
            v_aryEmployees, v_datStart, v_datEnd, _
            Nothing, False, v_intReturnOption, _
            trAllScheduleLevels, , trAllExceptions)
            
    Set l_obj = Nothing
            
    Exit Function
        
eh:
    TrsRaiseError m_strCMODULENAME, "GetSched"
            
End Function



Private Function GetDaysWorked( _
            ByVal v_aryEmployees As Variant, _
            ByVal v_datStart As Date, _
            ByVal v_datEnd As Date _
            ) As ADODB.Recordset
            
    Dim l_obj As dbTimeCard.cdTimeCard
    
    On Error GoTo eh
    
    Set l_obj = CreateObject("dbTimeCard.cdTimeCard")
    
    Set GetDaysWorked = l_obj.GetActualTimeDatesWorked( _
            v_aryEmployees, v_datStart, v_datEnd)
            
    Set l_obj = Nothing
    
    Exit Function
        
eh:
    TrsRaiseError m_strCMODULENAME, "GetDaysWorked"
            
            
            
End Function

Private Function BuildDaysOffRecordSet( _
            ByVal v_rsDaysWorked As ADODB.Recordset, _
            ByVal v_rsDaysSched As ADODB.Recordset _
            ) As ADODB.Recordset
            
    Dim l_rsDaysOffRS As ADODB.Recordset
    Dim l_lngEmpNum As Long
    Dim l_aryEmployees As Variant
    Dim l_lngStart As Long
    Dim l_lngEnd As Long
    Dim l_intI As Integer
    
    On Error GoTo eh
    

    Set l_rsDaysOffRS = CreateDaysOffRS()
    DetermineScheduledDaysOff v_rsDaysSched, l_rsDaysOffRS
    DetermineNoScheduleDays v_rsDaysSched, l_rsDaysOffRS
    RemoveDaysOff v_rsDaysWorked, l_rsDaysOffRS
    
    Set BuildDaysOffRecordSet = l_rsDaysOffRS
    
    Exit Function
        
eh:
    TrsRaiseError m_strCMODULENAME, "BuildDaysOffRecordSet"
    
            
End Function

Private Sub DetermineScheduledDaysOff( _
    ByVal v_rsSched As ADODB.Recordset, _
    ByRef r_rsDaysOffRS As ADODB.Recordset)
    
    Dim l_strFilter As String
    Dim l_datChargeDate As Date
    Dim l_lngEmpNum As Long
    Dim l_blnOTOppsInSched As Boolean
    
    On Error GoTo eh
    
    l_strFilter = "SHIFT_HOURS = 0" & _
            " AND SCHED_CD <> -1"
            
    l_strFilter = MDAC25FilterFix(l_strFilter)
    v_rsSched.Filter = l_strFilter
    
    If Not RecordsExist(v_rsSched) Then
        Exit Sub
    End If
    
    l_blnOTOppsInSched = False
    If FieldExists("OT_OPP_NUM", v_rsSched) Then
        l_blnOTOppsInSched = True
    End If
    
    v_rsSched.Sort = "EMP_NUM, CHARGE_DATE"
    v_rsSched.MoveFirst
    l_datChargeDate = CDate("1980-01-01")
    l_lngEmpNum = -1
        
    While Not v_rsSched.EOF
        If l_datChargeDate <> v_rsSched!CHARGE_DATE Or _
            l_lngEmpNum <> v_rsSched!EMP_NUM Then
            l_lngEmpNum = v_rsSched!EMP_NUM
            l_datChargeDate = v_rsSched!CHARGE_DATE
            r_rsDaysOffRS.AddNew
            r_rsDaysOffRS!EMP_NUM = l_lngEmpNum
            r_rsDaysOffRS!CHARGE_DATE = l_datChargeDate
            r_rsDaysOffRS!NO_SCHED_IND = "N"
        End If
        v_rsSched.MoveNext
    Wend
    
    r_rsDaysOffRS.Update
    'need to show days with scheduled OT as days worked

    If Not l_blnOTOppsInSched Then
    'overtime is not on schedule
        Exit Sub
    End If
    
    l_strFilter = "OT_OPP_NUM > 0"
    l_strFilter = MDAC25FilterFix(l_strFilter)
    v_rsSched.Filter = l_strFilter
    
    If RecordsExist(r_rsDaysOffRS) Then
        r_rsDaysOffRS.MoveFirst
    Else
        Exit Sub
    End If
    
    If Not RecordsExist(v_rsSched) Then
        Exit Sub
    End If
    
        
    v_rsSched.Sort = "EMP_NUM, CHARGE_DATE"
    v_rsSched.MoveFirst
    l_datChargeDate = CDate("1980-01-01")
    l_lngEmpNum = -1
                
    While Not v_rsSched.EOF
        If l_datChargeDate <> v_rsSched!CHARGE_DATE Or _
            l_lngEmpNum <> v_rsSched!EMP_NUM Then
            l_lngEmpNum = v_rsSched!EMP_NUM
            l_datChargeDate = v_rsSched!CHARGE_DATE
            If FindDayOffRec(l_lngEmpNum, l_datChargeDate, r_rsDaysOffRS) Then
                r_rsDaysOffRS.Delete
                r_rsDaysOffRS.Update
                If RecordsExist(r_rsDaysOffRS) Then
                    r_rsDaysOffRS.MoveFirst
                End If
            End If
        End If
        v_rsSched.MoveNext
    Wend


    Exit Sub
        
eh:
    TrsRaiseError m_strCMODULENAME, "DetermineScheduledDaysOff"
    
    
    
End Sub

'''''Private Sub RemoveDaysOff( _
'''''    ByVal v_lngEmpNum As Long, _
'''''    ByVal v_datStart As Date, _
'''''    ByVal v_datEnd As Date, _
'''''    ByVal v_rsDaysWorked As ADODB.Recordset, _
'''''    ByRef r_rsDaysOffRS As ADODB.Recordset)
'''''
'''''
'''''    Dim l_datChargeDate As Date
'''''    Dim l_strFilter As String
'''''
'''''    On Error GoTo eh
'''''
'''''    l_strFilter = "EMP_NUM = " & str(v_lngEmpNum) & _
'''''        " AND CHARGE_DATE >= " & str(v_datStart) & _
'''''        " AND CHARGE_DATE <= " & str(v_datEnd)
'''''
'''''    l_strFilter = MDAC25FilterFix(l_strFilter)
'''''    v_rsDaysWorked.Filter = l_strFilter
'''''
'''''    If Not RecordsExist(v_rsDaysWorked) Then
'''''        Exit Sub
'''''    End If
'''''
'''''    v_rsDaysWorked.Sort = "CHARGE_DATE"
'''''    v_rsDaysWorked.MoveFirst
'''''    l_datChargeDate = CDate("1980-01-01")
'''''    While Not v_rsDaysWorked.EOF
'''''        If l_datChargeDate <> v_rsDaysWorked!CHARGE_DATE Then
'''''            l_datChargeDate = v_rsDaysWorked!CHARGE_DATE
'''''            l_strFilter = "EMP_NUM = " & str(v_lngEmpNum) & _
'''''                " AND CHARGE_DATE = " & str(l_datChargeDate)
'''''            l_strFilter = MDAC25FilterFix(l_strFilter)
'''''            r_rsDaysOffRS.Filter = l_strFilter
'''''            If RecordsExist(r_rsDaysOffRS) Then
'''''                r_rsDaysOffRS.Delete
'''''            End If
'''''        End If
'''''        v_rsDaysWorked.MoveNext
'''''    Wend
'''''
'''''    Exit Sub
'''''
'''''eh:
'''''    TrsRaiseError m_strCMODULENAME, "RemoveDaysOff"
'''''
'''''End Sub

Private Sub RemoveDaysOff( _
    ByVal v_rsDaysWorked As ADODB.Recordset, _
    ByRef r_rsDaysOffRS As ADODB.Recordset)
    
    
    Dim l_datChargeDate As Date
    Dim l_strFilter As String
    Dim l_lngEmpNum As Long
    
    On Error GoTo eh
    
    If Not RecordsExist(r_rsDaysOffRS) Then
        Exit Sub
    End If
    
    If Not RecordsExist(v_rsDaysWorked) Then
        Exit Sub
    End If
    
    v_rsDaysWorked.Sort = "EMP_NUM, CHARGE_DATE"
    r_rsDaysOffRS.Sort = "EMP_NUM, CHARGE_DATE"
    v_rsDaysWorked.MoveFirst
    r_rsDaysOffRS.MoveFirst
    l_datChargeDate = CDate("1980-01-01")
    While Not v_rsDaysWorked.EOF And Not r_rsDaysOffRS.EOF
        l_datChargeDate = v_rsDaysWorked!CHARGE_DATE
        l_lngEmpNum = v_rsDaysWorked!EMP_NUM
        If FindDayOffRec(l_lngEmpNum, l_datChargeDate, r_rsDaysOffRS) Then
            r_rsDaysOffRS.Delete
            If Not r_rsDaysOffRS.EOF Then
                r_rsDaysOffRS.MoveNext
            End If
        End If
        v_rsDaysWorked.MoveNext
    Wend
    
    Exit Sub
        
eh:
    TrsRaiseError m_strCMODULENAME, "RemoveDaysOff"
    
End Sub


Private Sub DetermineActualTimeDaysOff( _
    ByVal v_lngEmpNum As Long, _
    ByVal v_datStart As Date, _
    ByVal v_datEnd As Date, _
    ByVal v_rsDaysWorked As ADODB.Recordset, _
    ByRef r_rsDaysOffRS As ADODB.Recordset)
    
    Dim l_datChargeDate As Date
    Dim l_lngNumOfDays As Long
    Dim l_intI As Integer
    Dim l_datLastATChargeDate As Date
    Dim l_datATChargeDate As Date
    
    On Error GoTo eh
    
    l_lngNumOfDays = DateDiff("d", v_datStart, v_datEnd)
    v_rsDaysWorked.Filter = "EMP_NUM = " & str(v_lngEmpNum)
    
    If RecordsExist(v_rsDaysWorked) Then
        v_rsDaysWorked.Sort = "CHARGE_DATE"
        v_rsDaysWorked.MoveLast
        l_datLastATChargeDate = v_rsDaysWorked!CHARGE_DATE
        v_rsDaysWorked.MoveFirst
    Else
        Exit Sub
    End If
    
    For l_intI = 0 To l_lngNumOfDays
        l_datChargeDate = DateAdd("d", l_intI, v_datStart)
        l_datATChargeDate = GetActualTimeChargeDate(v_rsDaysWorked, v_datEnd)
        If v_rsDaysWorked.EOF Then
            l_datATChargeDate = DateAdd("d", 1, v_datEnd)
        Else
            l_datATChargeDate = v_rsDaysWorked!CHARGE_DATE
        End If
        If l_datChargeDate < l_datATChargeDate Then
            r_rsDaysOffRS.AddNew
            r_rsDaysOffRS!EMP_NUM = v_lngEmpNum
            r_rsDaysOffRS!CHARGE_DATE = l_datChargeDate
            r_rsDaysOffRS!NO_SCHED_IND = "N"
        Else
            MoveToNextChargeDate v_rsDaysWorked, l_datChargeDate
        End If
    Next l_intI
    
    Exit Sub
        
eh:
    TrsRaiseError m_strCMODULENAME, "DetermineActualTimeDaysOff"
    

End Sub

Private Function GetActualTimeChargeDate( _
    ByVal v_rsDaysWorked As ADODB.Recordset, _
    ByVal v_datEnd As Date _
    ) As Date
    

    Dim l_datATChargeDate As Date

    On Error GoTo eh

    If v_rsDaysWorked.EOF Then
        l_datATChargeDate = DateAdd("d", 1, v_datEnd)
    Else
        l_datATChargeDate = v_rsDaysWorked!CHARGE_DATE
    End If
    
    GetActualTimeChargeDate = l_datATChargeDate
    
    Exit Function
        
eh:
    TrsRaiseError m_strCMODULENAME, "GetActualTimeChargeDate"
            
            
End Function

Private Sub MoveToNextChargeDate( _
    ByVal v_rsDaysWorked As ADODB.Recordset, _
    ByVal v_datChargeDate As Date)
    
    Dim l_blnFoundNext As Boolean
    
    On Error GoTo eh
    
    If v_rsDaysWorked.EOF Then
        Exit Sub
    End If
    
    l_blnFoundNext = False
    While Not v_rsDaysWorked.EOF _
        And Not l_blnFoundNext
        If v_datChargeDate = v_rsDaysWorked!CHARGE_DATE Then
            v_rsDaysWorked.MoveNext
        Else
            l_blnFoundNext = True
        End If
    Wend
    
    Exit Sub
        
eh:
    TrsRaiseError m_strCMODULENAME, "MoveToNextChargeDate"
    

    
End Sub

Private Function GetMergedExceptions( _
    ByVal v_aryEmployees As Variant, _
    ByVal v_datStart As Date, _
    ByVal l_datEnd As Date, _
    ByVal l_datLastPayEnd As Date, _
    ByVal l_datLastDayOfOpenPP As Date _
    ) As ADODB.Recordset
    
        
   Dim l_obj As dbSchedException.cdSchedException
   
   On Error GoTo eh
   
   Set l_obj = CreateObject("dbSchedException.cdSchedException")
   
   Set GetMergedExceptions = l_obj.GetExceptions( _
    v_aryEmployees, v_datStart, l_datEnd, l_datLastPayEnd, l_datLastDayOfOpenPP)
   
    Exit Function
        
eh:
    TrsRaiseError m_strCMODULENAME, "GetMergedExceptions"
  
End Function

Private Function GetChargeDates( _
    ByVal v_objRS As ADODB.Recordset, _
    ByVal v_datLastPayEnd As Date _
    ) As Variant
  
    Dim l_colChargeDates As Collection
    Dim l_datChargeDate As Date
    
    On Error GoTo eh
 
 
    Set l_colChargeDates = New Collection
    v_objRS.Sort = "CHARGE_DATE"
    v_objRS.MoveFirst
    While Not v_objRS.EOF
        l_datChargeDate = v_objRS!CHARGE_DATE
        If l_datChargeDate <= v_datLastPayEnd Then
            If Not IsDateInList(l_datChargeDate, l_colChargeDates) Then
                l_colChargeDates.Add l_datChargeDate
            End If
        End If
        v_objRS.MoveNext
    Wend
    
    GetChargeDates = ArrayFromCollection(l_colChargeDates)
    
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "GetChargeDates"
    
End Function
  
Private Function BuildGABRequest( _
    ByVal v_objRS As ADODB.Recordset, _
    ByVal v_datLastPayEnd As Date _
    ) As ADODB.Recordset

    Dim l_rsReq As ADODB.Recordset
    Dim l_lngEmpNum As Long
    Dim l_datChargeDate As Date
    
    On Error GoTo eh

    Set l_rsReq = CreateGABRequestRecordset()

    v_objRS.Filter = ""
    If Not RecordsExist(v_objRS) Then
        Set BuildGABRequest = Nothing
        Exit Function
    End If

    v_objRS.Sort = "EMP_NUM, CHARGE_DATE"
    v_objRS.MoveFirst
    l_lngEmpNum = 0
    l_datChargeDate = CDate("1980-01-01")


    While Not v_objRS.EOF
        If v_objRS!EMP_NUM = l_lngEmpNum Then
            If v_objRS!CHARGE_DATE = l_datChargeDate Then
            Else
                If v_objRS!CHARGE_DATE <= v_datLastPayEnd Then
                    l_rsReq.AddNew
                    l_rsReq!EMP_NUM = v_objRS!EMP_NUM
                    l_rsReq!START_DATE = v_objRS!CHARGE_DATE
                    l_rsReq!END_DATE = v_objRS!CHARGE_DATE
                End If
            End If
        Else
            If v_objRS!CHARGE_DATE <= v_datLastPayEnd Then
                l_rsReq.AddNew
                l_rsReq!EMP_NUM = v_objRS!EMP_NUM
                l_rsReq!START_DATE = v_objRS!CHARGE_DATE
                l_rsReq!END_DATE = v_objRS!CHARGE_DATE
            End If
        End If
        l_lngEmpNum = v_objRS!EMP_NUM
        l_datChargeDate = v_objRS!CHARGE_DATE
        v_objRS.MoveNext
    Wend

    Set BuildGABRequest = l_rsReq
    
    Exit Function
        
eh:
    TrsRaiseError m_strCMODULENAME, "BuildGABRequest"


End Function

Private Function CreateGABRequestRecordset()

    Dim l_rs As ADODB.Recordset
    
    On Error GoTo eh
    
    Set l_rs = New ADODB.Recordset

    With l_rs
      .Fields.Append "EMP_NUM", adInteger, , adFldUpdatable
      .Fields.Append "START_DATE", adDate, , adFldUpdatable
      .Fields.Append "END_DATE", adDate, , adFldUpdatable
      .CursorLocation = adUseClient
      .CursorType = adOpenDynamic
      .LockType = adLockOptimistic
      .Open
    End With
    
    Set CreateGABRequestRecordset = l_rs
    
    Exit Function
        
eh:
    TrsRaiseError m_strCMODULENAME, "CreateGABRequestRecordset"


End Function


Private Function GetSchedOnExceptionDays( _
            ByVal v_rsReq As ADODB.Recordset _
            ) As ADODB.Recordset
            
    Dim l_obj As busSchedule.cbsSchedule
    
    On Error GoTo eh
    
    Set l_obj = CreateObject("busSchedule.cbsSchedule")
    
    Set GetSchedOnExceptionDays = l_obj.GABEmployeeSchedOnSpecificDates( _
             v_rsReq, Nothing, False, _
             trBaseSchedNoDetails, trAllScheduleLevels)
            
    Set l_obj = Nothing
    
    Exit Function
        
eh:
    TrsRaiseError m_strCMODULENAME, "GetSchedOnExceptionDays"
            
            
            
End Function

Private Sub DetermineFullShiftInd( _
    ByVal v_rsSchedule As ADODB.Recordset, _
    ByRef r_rsExceptions As ADODB.Recordset)
    
'This also gets the shift code (A,P,O,D,N) if the shift number is > 0
'Shift number will be 0 if it came from actual time or adjustments

    Dim l_lngEmpNum As Long
    Dim l_datChargeDate As Date
    Dim l_strFilter As String
    Dim l_intShiftNum As Integer
    Dim l_blnSchedFound As Boolean
    
    On Error GoTo eh
    
    If r_rsExceptions Is Nothing Then
        Exit Sub
    End If
    
    r_rsExceptions.Filter = ""
    v_rsSchedule.Filter = ""
    
    If Not RecordsExist(r_rsExceptions) Then
        Exit Sub
    End If
    
    r_rsExceptions.Sort = "EMP_NUM, CHARGE_DATE, SHIFT_NUM"
    r_rsExceptions.MoveFirst
    v_rsSchedule.Sort = "EMP_NUM, CHARGE_DATE, SHIFT_NUM"
'    v_rsSchedule.MoveFirst
    
    While Not r_rsExceptions.EOF
        l_lngEmpNum = r_rsExceptions!EMP_NUM
        l_datChargeDate = r_rsExceptions!CHARGE_DATE
        l_intShiftNum = r_rsExceptions!SHIFT_NUM
        l_blnSchedFound = FindScheduleRec(l_lngEmpNum, l_datChargeDate, l_intShiftNum, v_rsSchedule)
        If l_blnSchedFound Then
            If Not IsOvertime(r_rsExceptions!BANK_CD, r_rsExceptions!EXCEPTION_CD) Then
                If v_rsSchedule!SHIFT_HOURS = r_rsExceptions!Hours _
                    Or (v_rsSchedule!SHIFT_HOURS = 0 _
                        And IsAHoliday(r_rsExceptions!EXCEPTION_CD)) Then
                    r_rsExceptions!FULL_SHIFT_EXC_IND = "Y"
                    If l_intShiftNum > 0 Then
                        r_rsExceptions!SHIFT_CODE = v_rsSchedule!SHIFT_CODE
                    End If
                    r_rsExceptions.Update
                End If
            Else
                'PTS 900:  All OT should be displayed as a full shift exception - OT is outside of the schedule, therefore
                'a partial indicator is relevant.
                r_rsExceptions!FULL_SHIFT_EXC_IND = "Y"
            End If
        End If
        r_rsExceptions.MoveNext
    Wend
    
    Exit Sub
        
eh:
    TrsRaiseError m_strCMODULENAME, "DetermineFullShiftInd"

End Sub

Private Function FindScheduleRec( _
    ByVal v_lngEmpNum As Long, _
    ByVal v_datChargeDate As Date, _
    ByVal v_intShiftNum As Integer, _
    ByRef r_rsSchedule As ADODB.Recordset) _
    As Boolean
    
    
    Dim l_lngEmpNum As Long
    Dim l_datChargeDate As Date
    Dim l_blnSchedFound As Boolean
    Dim l_intShiftNum As Integer
    
    On Error GoTo eh
    
    r_rsSchedule.MoveFirst
    l_blnSchedFound = False
    
    While Not r_rsSchedule.EOF And Not l_blnSchedFound
        l_lngEmpNum = r_rsSchedule!EMP_NUM
        l_datChargeDate = r_rsSchedule!CHARGE_DATE
        l_intShiftNum = r_rsSchedule!SHIFT_NUM
        If l_lngEmpNum = v_lngEmpNum _
            And l_datChargeDate = v_datChargeDate Then
            If v_intShiftNum = 0 Or _
                l_intShiftNum = v_intShiftNum Then
                l_blnSchedFound = True
            Else
                r_rsSchedule.MoveNext
            End If
        Else
            r_rsSchedule.MoveNext
        End If
    Wend
    
    FindScheduleRec = l_blnSchedFound
        
    Exit Function
        
eh:
    TrsRaiseError m_strCMODULENAME, "FindScheduleRec"
    
End Function

Private Function IsAHoliday( _
            ByVal v_strException As String _
            ) As Boolean
            
'May want to change implentation to look at busCodeTables
' before we go live with TRIS 4
            
    Dim l_strExcCd As String
    
    On Error GoTo eh
    
    l_strExcCd = Trim(v_strException)
    
    Select Case l_strExcCd
        Case "H", "HR"
            IsAHoliday = True
        Case Else
            IsAHoliday = False
    End Select
    
    Exit Function
        
eh:
    TrsRaiseError m_strCMODULENAME, "IsAHoliday"

End Function

Private Function GetActualTimeSchedChangePay( _
    ByVal v_aryEmployees As Variant, _
    ByVal v_datStart As Date, _
    ByVal v_datEnd As Date _
    ) As ADODB.Recordset
        
   Dim l_obj As dbSchedException.cdSchedException
   
   On Error GoTo eh
   
   Set l_obj = CreateObject("dbSchedException.cdSchedException")
   
   Set GetActualTimeSchedChangePay = l_obj.GetActualTimeSchedChangePay( _
    v_aryEmployees, v_datStart, v_datEnd)
    
    Exit Function
        
eh:
    TrsRaiseError m_strCMODULENAME, "GetActualTimeSchedChangePay"
       
    
End Function

Private Function GetAdjustments( _
    ByVal v_aryEmployees As Variant, _
    ByVal v_datStart As Date, _
    ByVal v_datEnd As Date _
    ) As ADODB.Recordset
        
   Dim l_obj As dbSchedException.cdSchedException
   
   On Error GoTo eh
   
   Set l_obj = CreateObject("dbSchedException.cdSchedException")
   
   Set GetAdjustments = l_obj.GetAdjustments( _
    v_aryEmployees, v_datStart, v_datEnd)
    
    Exit Function
        
eh:
    TrsRaiseError m_strCMODULENAME, "GetAdjustments"
       
    
End Function

Private Sub ApplyAdjustments( _
    ByVal v_rsAdj As ADODB.Recordset, _
    ByRef r_rsExceptions As ADODB.Recordset)
    
    Dim l_strAdjtype As String
    Dim l_datChargeDate As Date
    Dim l_lngEmpNum As Long
    Dim l_strException As String
    Dim l_blnExceptionFound As Boolean
    Dim l_decHours As Variant
    Dim l_decAdjHours As Variant
    Dim l_sColorGroup As String
    
    On Error GoTo eh
        
    If Not RecordsExist(v_rsAdj) Then
        Exit Sub
    End If
    
    If RecordsExist(r_rsExceptions) Then
        r_rsExceptions.Filter = ""
        r_rsExceptions.Sort = "EMP_NUM, CHARGE_DATE,EXCEPTION_CD"
        r_rsExceptions.MoveFirst
    End If
    
    v_rsAdj.MoveFirst
    While Not v_rsAdj.EOF
        l_strAdjtype = v_rsAdj!ADJ_TYPE
        l_lngEmpNum = v_rsAdj!EMP_NUM
        l_datChargeDate = v_rsAdj!CHARGE_DATE
        l_strException = v_rsAdj!EXCEPTION_CD
        l_decAdjHours = CDec(v_rsAdj!Hours)
        l_sColorGroup = v_rsAdj!EG_SUB_CD
        '************************************************************************************************
        '* NOTE:  the following conversions of EA and EX to an E is a temporary fix for PTS 1143
        '* A permanent change will be implemented with the CILCO implementation when
        '* compatibility will be broken and the adjustments returned to the GUI for display
        '* in the hover help of the VIEW CALENDAR Calendar Tab.
        '* DES 1/23/2003
        '************************************************************************************************
        If (Trim(l_strException) = "EA") Or (Trim(l_strException) = "EX") Then
            l_strException = "E"
        End If
        l_blnExceptionFound = FindException( _
            r_rsExceptions, l_lngEmpNum, l_datChargeDate, l_strException)
        'find Exception
        If l_strAdjtype = "A" Then
            If l_blnExceptionFound Then

                l_decHours = CDec(r_rsExceptions!Hours)
'BEGIN TRIS CONVERSION PHASE 3
'Please note that AS PAID time related adjustments are stored as
'signed numbers now (e.g. -4 hours) TTR_ADJ_TIME
                If (l_decAdjHours < 0) Then
                    l_decHours = l_decHours - Abs(l_decAdjHours)
                Else
                    l_decHours = l_decHours - l_decAdjHours
                End If
'END TRIS CONVERSION PHASE 3
                If l_decHours <= 0 Then
                    r_rsExceptions.Delete
                Else
                    r_rsExceptions!Hours = l_decHours
                    r_rsExceptions.Update
                End If
            End If
        Else
            If l_blnExceptionFound Then
                l_decHours = CDec(r_rsExceptions!Hours)
                l_decHours = l_decHours + l_decAdjHours
                r_rsExceptions!Hours = l_decHours
                r_rsExceptions.Update
            Else
                AddAdjustmentException v_rsAdj, r_rsExceptions
            End If
        End If
        v_rsAdj.MoveNext
    Wend
    
    Exit Sub
        
eh:
    TrsRaiseError m_strCMODULENAME, "ApplyAdjustments"

End Sub

Private Sub AddAdjustmentException( _
    ByVal v_rsAdj As ADODB.Recordset, _
    ByRef r_rsExceptions As ADODB.Recordset)
    
    Dim l_strExcType As String
    
    On Error GoTo eh
    
    l_strExcType = Trim(v_rsAdj!EXCEPTION_TYPE)
    
    If l_strExcType = "P" Then
        r_rsExceptions.AddNew
        r_rsExceptions!EMP_NUM = v_rsAdj!EMP_NUM
        r_rsExceptions!CHARGE_DATE = v_rsAdj!CHARGE_DATE
        r_rsExceptions!EXCEPTION_CD = v_rsAdj!EXCEPTION_CD
        r_rsExceptions!OVER_EXC_CD = ""
        r_rsExceptions!CASE_GROUP = v_rsAdj!CASE_GROUP
        r_rsExceptions!BANK_CD = v_rsAdj!BANK_CD
        r_rsExceptions!FULL_SHIFT_EXC_IND = "N"
        r_rsExceptions!Hours = CDec(v_rsAdj!Hours)
        r_rsExceptions!SHIFT_NUM = 0
        r_rsExceptions!SYMBOL_CD = v_rsAdj!SYMBOL_CD
        r_rsExceptions!EG_SUB_CD = v_rsAdj!EG_SUB_CD
        r_rsExceptions!SHIFT_CODE = ""
        r_rsExceptions!SCHED_CHANGE_PMT = "N"
        r_rsExceptions.Update
    End If
    
    Exit Sub
        
eh:
    TrsRaiseError m_strCMODULENAME, "AddAdjustmentException"

End Sub

Private Function FindException( _
    ByRef r_rsExceptions As ADODB.Recordset, _
    ByVal v_lngEmpNum As Long, _
    ByVal v_datChargeDate As Date, _
    ByVal v_strException As String _
    ) As Boolean
    
    
    Dim l_strFilter As String
    
    On Error GoTo eh
    
    If r_rsExceptions Is Nothing Then
        FindException = False
        Exit Function
    End If
    
    l_strFilter = "EMP_NUM = " & str(v_lngEmpNum) _
       & " AND CHARGE_DATE = " & str(v_datChargeDate) _
       & " AND EXCEPTION_CD = '" & Trim(v_strException) & "'"
    
    l_strFilter = MDAC25FilterFix(l_strFilter)
    r_rsExceptions.Filter = l_strFilter
    
    If RecordsExist(r_rsExceptions) Then
        FindException = True
    Else
        FindException = False
    End If
        
    Exit Function
        
eh:
    TrsRaiseError m_strCMODULENAME, "FindException"
            
            
End Function

Private Function IsDateInList( _
        ByVal v_datDate As Date, _
        ByVal v_colCodes As Collection _
        ) As Boolean

    Dim l_intIndex As Integer
    Dim l_intCount As Integer
    Dim l_bolInList As Boolean
    
    On Error GoTo eh
    
    l_bolInList = False
    l_intIndex = 1
    l_intCount = v_colCodes.Count
    
    While Not l_bolInList And l_intIndex <= l_intCount
        If CDate(v_colCodes.Item(l_intIndex)) = v_datDate Then
            l_bolInList = True
        End If
        l_intIndex = l_intIndex + 1
    Wend
    
    IsDateInList = l_bolInList
    
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "IsDateInList"
    

End Function

Private Sub DetermineNoScheduleDays( _
    ByVal v_rsSched As ADODB.Recordset, _
    ByRef r_rsDaysOffRS As ADODB.Recordset)
    
    Dim l_strFilter As String
    Dim l_datChargeDate As Date
    Dim l_lngEmpNum As Long
    
    On Error GoTo eh
    
    l_strFilter = "SCHED_CD = -1"
            
    v_rsSched.Filter = l_strFilter
    
    If Not RecordsExist(v_rsSched) Then
        Exit Sub
    End If
        
    v_rsSched.Sort = "CHARGE_DATE"
    v_rsSched.MoveFirst
    l_datChargeDate = CDate("1980-01-01")
    l_lngEmpNum = -1
        
    While Not v_rsSched.EOF
        If l_datChargeDate <> v_rsSched!CHARGE_DATE Or _
            l_lngEmpNum <> v_rsSched!EMP_NUM Then
            l_datChargeDate = v_rsSched!CHARGE_DATE
            l_lngEmpNum = v_rsSched!EMP_NUM
            r_rsDaysOffRS.AddNew
            r_rsDaysOffRS!EMP_NUM = l_lngEmpNum
            r_rsDaysOffRS!CHARGE_DATE = l_datChargeDate
            r_rsDaysOffRS!NO_SCHED_IND = "Y"
        End If
        v_rsSched.MoveNext
    Wend
    
    Exit Sub
        
eh:
    TrsRaiseError m_strCMODULENAME, "DetermineNoScheduleDays"
    
End Sub

Private Function IsOvertime( _
    ByVal v_strBankCode As String, _
    ByVal v_strExceptionCd As String _
    ) As Boolean
    
    On Error GoTo eh
    
    If v_strBankCode = "O" Or _
        Trim(v_strExceptionCd) = "ONP" Then
        IsOvertime = True
    Else
        IsOvertime = False
        
    End If
            
    Exit Function
        
eh:
    TrsRaiseError m_strCMODULENAME, "IsOvertime"
    
    
End Function


Private Sub DetermineTradeDoubleInd( _
    ByRef r_rsSched As ADODB.Recordset, _
    ByVal v_datLastPayEnd As Date)
    
    Dim l_strFilter As String
    Dim l_rsEmpChargeDates As ADODB.Recordset
    Dim l_lngEmpNum As Long
    Dim l_datChargeDate As Date
    
    On Error GoTo eh
    
    l_strFilter = "SHIFT_NUM = 2" & _
         " AND CHARGE_DATE > " & str(v_datLastPayEnd) & _
         " AND SHIFT_HOURS > 0"
    l_strFilter = MDAC25FilterFix(l_strFilter)
    r_rsSched.Filter = l_strFilter
    
    If Not RecordsExist(r_rsSched) Then
        r_rsSched.Filter = ""
        Exit Sub
    End If
    
    Set l_rsEmpChargeDates = CreateEmpChargeDateRS()
    r_rsSched.MoveFirst
    While Not r_rsSched.EOF
        l_rsEmpChargeDates.AddNew
        l_rsEmpChargeDates!EMP_NUM = r_rsSched!EMP_NUM
        l_rsEmpChargeDates!CHARGE_DATE = r_rsSched!CHARGE_DATE
        r_rsSched.MoveNext
    Wend
    l_rsEmpChargeDates.Sort = "EMP_NUM,CHARGE_DATE"
    l_rsEmpChargeDates.MoveFirst
    While Not l_rsEmpChargeDates.EOF
        l_lngEmpNum = l_rsEmpChargeDates!EMP_NUM
        l_datChargeDate = l_rsEmpChargeDates!CHARGE_DATE
        l_strFilter = "EMP_NUM = " & str(l_lngEmpNum) & _
         " AND CHARGE_DATE = " & str(l_datChargeDate)
        l_strFilter = MDAC25FilterFix(l_strFilter)
        r_rsSched.Filter = l_strFilter
        r_rsSched.Sort = "SHIFT_NUM"
        r_rsSched.MoveFirst
        If r_rsSched!SHIFT_NUM = 1 Then
            If r_rsSched!SHIFT_HOURS > 0 Then
                While Not r_rsSched.EOF
                    r_rsSched!DOUBLE_SHIFT_IND = "Y"
                    r_rsSched.Update
                    r_rsSched.MoveNext
                Wend
            End If
        End If
        l_rsEmpChargeDates.MoveNext
    Wend
    
    Exit Sub
        
eh:
    TrsRaiseError m_strCMODULENAME, "DetermineTradeDoubleInd"
    
End Sub

Private Function CreateEmpChargeDateRS()

    Dim l_rs As ADODB.Recordset
    
    On Error GoTo eh
    
    Set l_rs = New ADODB.Recordset

    With l_rs
      .Fields.Append "EMP_NUM", adInteger, , adFldUpdatable
      .Fields.Append "CHARGE_DATE", adDate, , adFldUpdatable
      .CursorLocation = adUseClient
      .CursorType = adOpenDynamic
      .LockType = adLockOptimistic
      .Open
    End With
    
    Set CreateEmpChargeDateRS = l_rs
    
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "CreateEmpChargeDateRS"


End Function


Private Function BuildMergeRS( _
    ByVal v_datStart As Date, _
    ByVal v_datEnd As Date, _
    ByVal v_datLastPayEnd As Date, _
    ByVal v_datLastDayOpenPP As Date, _
    ByVal v_rsAdjustments As ADODB.Recordset, _
    ByVal v_rsActualTime As ADODB.Recordset, _
    ByVal v_rsSchedule As ADODB.Recordset, _
    ByVal v_rsActualTimeForOpenPP As ADODB.Recordset _
    ) As ADODB.Recordset
    
    Dim l_rsMerge As ADODB.Recordset
    
    On Error GoTo eh
    
    Set l_rsMerge = CreateMergeRS()
    AddAdjustments v_rsAdjustments, l_rsMerge
    AddActualTime v_rsActualTime, l_rsMerge
    AddSchedule v_rsSchedule, v_datLastPayEnd, _
        v_datLastDayOpenPP, v_datEnd, v_rsActualTimeForOpenPP, l_rsMerge
        
    Set BuildMergeRS = l_rsMerge
    
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "BuildMergeRS"
    
    
End Function
Private Function CreateMergeRS() As ADODB.Recordset

    Dim l_rs As ADODB.Recordset
    
    On Error GoTo eh
    
    Set l_rs = New ADODB.Recordset
    With l_rs
      .Fields.Append "TIME_TYPE", adChar, 1, adFldUpdatable
      .Fields.Append "EMP_NUM", adInteger, , adFldUpdatable
      .Fields.Append "CHARGE_DATE", adDate, , adFldUpdatable
      .Fields.Append "DATE_WORKED", adDate, , adFldUpdatable
      .Fields.Append "DISTRIBUTION_ID", adChar, 39, adFldUpdatable
      .Fields.Append "DISTRIBUTION_DESC", adVarChar, 255, adFldUpdatable
      .Fields.Append "HOURS", adDecimal, , adFldUpdatable
      .Fields.Item("HOURS").Precision = 3
      .Fields.Item("HOURS").NumericScale = 1
      .Fields.Append "START_TIME", adChar, 5, adFldUpdatable
      .Fields.Append "END_TIME", adChar, 5, adFldUpdatable
      .Fields.Append "EXCEPTION_CD", adChar, 3, adFldUpdatable
      .Fields.Append "OVER_EXC_CD", adChar, 3, adFldUpdatable
      .Fields.Append "JOB_NUM", adChar, 6, adFldUpdatable
      .Fields.Append "WG_NUM", adInteger, 7, adFldUpdatable
      .Fields.Append "OT_REASON_DESC", adVarChar, 255, adFldUpdatable
      .Fields.Append "OT_OPP_DESC", adVarChar, 255, adFldUpdatable
      .Fields.Append "CASE_NUM", adInteger, , adFldUpdatable
      .Fields.Append "FMLA_IND", adChar, 1, adFldUpdatable
      .Fields.Append "COMP_SHIFT_CD", adChar, 3, adFldUpdatable
      .Fields.Append "COMP_SHIFT_CD_AMT", adDecimal, , adFldUpdatable
      .Fields.Item("COMP_SHIFT_CD_AMT").Precision = 5
      .Fields.Item("COMP_SHIFT_CD_AMT").NumericScale = 4
      .Fields.Append "COMP_CD", adChar, 3, adFldUpdatable
      .Fields.Append "REL_CD", adChar, 3, adFldUpdatable
      .Fields.Append "SCHED_CHG_EXC_CD", adChar, 3, adFldUpdatable
      .Fields.Append "DISTR_DEPT_NUM", adChar, 3, adFldUpdatable
      .Fields.Append "DEPT_NUM", adChar, 3, adFldUpdatable
      .Fields.Append "LOC_CD", adInteger, , adFldUpdatable
      .Fields.Append "OT_OPP_NUM", adInteger, , adFldUpdatable
      .Fields.Append "OT_REASON_CD", adChar, 3, adFldUpdatable
      .Fields.Append "TRAVEL_PAY_IND", adChar, 1, adFldUpdatable
      .Fields.Append "APPR_BY_EMP_NUM", adInteger, , adFldUpdatable
      .Fields.Append "APPR_DATE", adDate, , adFldUpdatable
      .Fields.Append "AMOUNT", adDecimal, , adFldUpdatable
      .Fields.Item("AMOUNT").Precision = 9
      .Fields.Item("AMOUNT").NumericScale = 2
      .Fields.Append "START_LOC_CD", adInteger, , adFldUpdatable
      .Fields.Append "END_LOC_CD", adInteger, , adFldUpdatable
      .Fields.Append "OT_BANK_CHARGE_DATE", adDate, , adFldUpdatable
      .Fields.Append "LINE_NUM", adInteger, , adFldUpdatable
      .Fields.Append "REQ_LINE_IND", adChar, 1, adFldUpdatable
      .Fields.Append "SCHED_CD", adInteger, , adFldUpdatable
      .Fields.Append "SHIFT_NUM", adInteger, , adFldUpdatable
      .Fields.Append "SHIFT_CODE", adChar, 1, adFldUpdatable
      .Fields.Append "SCHED_SRC_CD", adInteger, , adFldUpdatable
      .Fields.Append "SHIFT_HOURS", adDecimal, , adFldUpdatable
      .Fields.Item("SHIFT_HOURS").Precision = 3
      .Fields.Item("SHIFT_HOURS").NumericScale = 1
      .Fields.Append "ADJ_TYPE", adChar, 1, adFldUpdatable
      .Fields.Append "SEQ_NUM", adInteger, , adFldUpdatable
      .Fields.Append "HOURLY_RATE", adDecimal, adFldUpdatable
      .Fields.Item("HOURLY_RATE").Precision = 9
      .Fields.Item("HOURLY_RATE").NumericScale = 2
      .Fields.Append "REQ_BY_EMP_NUM", adInteger, , adFldUpdatable
      .Fields.Append "REQ_DATE", adDate, , adFldUpdatable
      .Fields.Append "COMPLETED_BY_EMP", adInteger, , adFldUpdatable
      .Fields.Append "COMPLETED_DATE", adDate, , adFldUpdatable
      .Fields.Append "FINAL_APPROVER", adInteger, , adFldUpdatable
      .Fields.Append "FINAL_APPR_DATE", adDate, , adFldUpdatable
      .Fields.Append "SPECIAL_CHECK", adChar, 1, adFldUpdatable
      .Fields.Append "ADJ_PREV_PAID_FL", adChar, 1, adFldUpdatable
      .Fields.Append "REQ_REASON", adVarChar, 255, adFldUpdatable
      .Fields.Append "LAST_MOD_EMP_NUM", adInteger, , adFldUpdatable
      .Fields.Append "INDEX_TIME", adChar, 26, adFldUpdatable
      .Fields.Append "ROT_CD", adInteger, , adFldUpdatable
      .Fields.Append "ROT_WEEK", adInteger, , adFldUpdatable
      .Fields.Append "TRADE_EXC_CD", adChar, 7, adFldUpdatable
      
      .CursorLocation = adUseClient
      .CursorType = adOpenDynamic
      .LockType = adLockOptimistic
      .Open
    End With
    
    Set CreateMergeRS = l_rs
    
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "CreateMergeRS"
    

End Function
Private Sub AddAdjustments( _
            ByVal v_rsAdjustments As ADODB.Recordset _
            , ByRef r_rsMerge As ADODB.Recordset)
On Error GoTo eh

    If Not RecordsExist(v_rsAdjustments) Then
        Exit Sub
    End If

    v_rsAdjustments.MoveFirst

    While Not v_rsAdjustments.EOF
'BEGIN PREM TRIS CONVERSION PHASE 3
        With r_rsMerge
            .AddNew
            !TIME_TYPE = "1"
            !EMP_NUM = v_rsAdjustments!EMP_NUM
            If (v_rsAdjustments!CHARGE_DATE = #1/1/100#) Then
                !CHARGE_DATE = v_rsAdjustments!DATE_WORKED
            Else
                !CHARGE_DATE = v_rsAdjustments!CHARGE_DATE
            End If
            !DATE_WORKED = v_rsAdjustments!DATE_WORKED
            !DISTRIBUTION_ID = Trim(v_rsAdjustments!DISTRIBUTION_ID)
            !DISTRIBUTION_DESC = ""
            !Hours = v_rsAdjustments!Hours
            !START_TIME = v_rsAdjustments!START_TIME
            !END_TIME = v_rsAdjustments!END_TIME
            !EXCEPTION_CD = v_rsAdjustments!EXCEPTION_CD
            !OVER_EXC_CD = v_rsAdjustments!OVER_EXC_CD
            !JOB_NUM = v_rsAdjustments!JOB_NUM
            !WG_NUM = v_rsAdjustments!WG_NUM
            !OT_REASON_DESC = ""
            !OT_OPP_DESC = v_rsAdjustments!OT_OPP_DESC
            !CASE_NUM = v_rsAdjustments!CASE_NUM
            !FMLA_IND = v_rsAdjustments!FAMILY_LEAVE_IND
            !COMP_SHIFT_CD = v_rsAdjustments!SHIFT_COMP
            !COMP_SHIFT_CD_AMT = 0
            If (Trim(v_rsAdjustments!INTERPLANT_CD) <> "") Then
                !COMP_CD = v_rsAdjustments!INTERPLANT_CD
            Else
                !COMP_CD = v_rsAdjustments!COMP_CD
            End If
            !REL_CD = v_rsAdjustments!REL_CD
            !SCHED_CHG_EXC_CD = v_rsAdjustments!SCHED_CHG_EXC_CD
            !DISTR_DEPT_NUM = v_rsAdjustments!DISTR_DEPT_NUM
            !DEPT_NUM = v_rsAdjustments!DEPT_NUM
            !LOC_CD = v_rsAdjustments!LOC_CD
            !OT_OPP_NUM = v_rsAdjustments!OT_OPP_NUM
            !OT_REASON_CD = v_rsAdjustments!OT_REASON_CD
            !TRAVEL_PAY_IND = "N"
            !APPR_BY_EMP_NUM = v_rsAdjustments!APPR_BY_MANAGER
            !APPR_DATE = v_rsAdjustments!APPR_DATE
            !AMOUNT = v_rsAdjustments!AMOUNT
            !START_LOC_CD = v_rsAdjustments!START_LOC_CD
            !END_LOC_CD = v_rsAdjustments!END_LOC_CD
            !OT_BANK_CHARGE_DATE = v_rsAdjustments!OT_BANK_CHRG_DT
            !LINE_NUM = 0
            !REQ_LINE_IND = "N"
            !SCHED_CD = v_rsAdjustments!SCHED_CD
            !SHIFT_NUM = 0
            !SHIFT_CODE = ""
            !SCHED_SRC_CD = v_rsAdjustments!SCHED_SRC_CD
            !SHIFT_HOURS = 0
            !ADJ_TYPE = v_rsAdjustments!ADJ_TYPE
            !SEQ_NUM = v_rsAdjustments!SEQ_NUM
            !HOURLY_RATE = v_rsAdjustments!HOURLY_RATE
            !REQ_BY_EMP_NUM = v_rsAdjustments!REQ_BY_EMP_NUM
            !REQ_DATE = v_rsAdjustments!REQ_DATE
            !COMPLETED_BY_EMP = v_rsAdjustments!COMPLETED_BY_EMP
            !COMPLETED_DATE = v_rsAdjustments!COMPLETED_DATE
            !FINAL_APPROVER = v_rsAdjustments!FINAL_APPROVER
            !FINAL_APPR_DATE = v_rsAdjustments!FINAL_APPR_DATE
            !SPECIAL_CHECK = v_rsAdjustments!SPECIAL_CHECK
            !ADJ_PREV_PAID_FL = v_rsAdjustments!ADJ_PREV_PAID_FL
            !REQ_REASON = v_rsAdjustments!REQ_REASON
            !LAST_MOD_EMP_NUM = v_rsAdjustments!LAST_MOD_EMP_NUM
            !INDEX_TIME = v_rsAdjustments!INDEX_TIME
            !ROT_CD = 0
            !ROT_WEEK = 0
            !TRADE_EXC_CD = v_rsAdjustments!TRADE_EXC_CD
            .Update
        End With
'END PREM TRIS CONVERSION PHASE 3
        v_rsAdjustments.MoveNext
    Wend

    Exit Sub
eh:
    TrsRaiseError m_strCMODULENAME, "AddAdjustments"

End Sub

Private Sub AddActualTime( _
    ByVal v_rsActualTime As ADODB.Recordset, _
    ByRef r_rsMerge As ADODB.Recordset)
    
    On Error GoTo eh
    
    If Not RecordsExist(v_rsActualTime) Then
        Exit Sub
    End If
    
    v_rsActualTime.MoveFirst
    
    While Not v_rsActualTime.EOF
        r_rsMerge.AddNew
        r_rsMerge!TIME_TYPE = "2"
        r_rsMerge!EMP_NUM = v_rsActualTime!EMP_NUM
        r_rsMerge!CHARGE_DATE = CDate(v_rsActualTime!CHARGE_DATE)
        r_rsMerge!DATE_WORKED = CDate(v_rsActualTime!DATE_WORKED)
        r_rsMerge!DISTRIBUTION_ID = v_rsActualTime!DISTRIBUTION_ID
        r_rsMerge!DISTRIBUTION_DESC = Trim(v_rsActualTime!DISTRIBUTION_DESC)
        r_rsMerge!Hours = CDec(v_rsActualTime!Hours)
        r_rsMerge!START_TIME = v_rsActualTime!START_TIME
        r_rsMerge!END_TIME = v_rsActualTime!END_TIME
        r_rsMerge!EXCEPTION_CD = v_rsActualTime!EXCEPTION_CD
        r_rsMerge!OVER_EXC_CD = v_rsActualTime!OVER_EXC_CD
        r_rsMerge!JOB_NUM = v_rsActualTime!JOB_NUM
        r_rsMerge!WG_NUM = v_rsActualTime!WG_NUM
        r_rsMerge!OT_REASON_DESC = Trim(v_rsActualTime!OT_REASON_DESC)
        r_rsMerge!OT_OPP_DESC = Trim(v_rsActualTime!OT_OPP_DESC)
        r_rsMerge!CASE_NUM = v_rsActualTime!CASE_NUM
        r_rsMerge!FMLA_IND = v_rsActualTime!FAMILY_LEAVE_IND
        r_rsMerge!COMP_SHIFT_CD = v_rsActualTime!COMP_SHIFT_CD
        r_rsMerge!COMP_SHIFT_CD_AMT = v_rsActualTime!COMP_SHIFT_CD_AMT
        r_rsMerge!COMP_CD = v_rsActualTime!COMP_CD
        r_rsMerge!REL_CD = v_rsActualTime!REL_CD
        r_rsMerge!SCHED_CHG_EXC_CD = v_rsActualTime!SCHED_CHG_EXC_CD
        r_rsMerge!DISTR_DEPT_NUM = v_rsActualTime!DISTR_DEPT_NUM
        r_rsMerge!DEPT_NUM = v_rsActualTime!DEPT_NUM
        r_rsMerge!LOC_CD = v_rsActualTime!LOC_CD
        r_rsMerge!OT_OPP_NUM = v_rsActualTime!OT_OPP_NUM
        r_rsMerge!OT_REASON_CD = v_rsActualTime!OT_REASON_CD
        r_rsMerge!TRAVEL_PAY_IND = v_rsActualTime!TRAVEL_PAY_IND
        r_rsMerge!APPR_BY_EMP_NUM = v_rsActualTime!APPR_BY_EMP_NUM
        r_rsMerge!APPR_DATE = CDate(v_rsActualTime!APPR_DATE)
        r_rsMerge!AMOUNT = CDec(v_rsActualTime!AMOUNT)
        r_rsMerge!START_LOC_CD = v_rsActualTime!START_LOC_CD
        r_rsMerge!END_LOC_CD = v_rsActualTime!END_LOC_CD
        r_rsMerge!OT_BANK_CHARGE_DATE = v_rsActualTime!OT_BANK_CHRG_DT
        r_rsMerge!LINE_NUM = 0
        r_rsMerge!REQ_LINE_IND = "N"
        r_rsMerge!SCHED_CD = v_rsActualTime!SCHED_CD
        r_rsMerge!SHIFT_NUM = 0
        r_rsMerge!SHIFT_CODE = ""
        r_rsMerge!SCHED_SRC_CD = v_rsActualTime!SCHED_SRC_CD
        r_rsMerge!SHIFT_HOURS = 0
        r_rsMerge!ADJ_TYPE = ""
        r_rsMerge!SEQ_NUM = 0
        r_rsMerge!HOURLY_RATE = CDec(v_rsActualTime!HOURLY_RATE)
        r_rsMerge!REQ_BY_EMP_NUM = 0
        r_rsMerge!REQ_DATE = CDate("0100-01-01")
        r_rsMerge!COMPLETED_BY_EMP = 0
        r_rsMerge!COMPLETED_DATE = CDate("0100-01-01")
        r_rsMerge!FINAL_APPROVER = 0
        r_rsMerge!FINAL_APPR_DATE = CDate("0100-01-01")
        r_rsMerge!SPECIAL_CHECK = ""
        r_rsMerge!ADJ_PREV_PAID_FL = ""
        r_rsMerge!REQ_REASON = ""
        r_rsMerge!LAST_MOD_EMP_NUM = v_rsActualTime!LAST_MOD_EMP_NUM
        r_rsMerge!INDEX_TIME = v_rsActualTime!INDEX_TIME
        r_rsMerge!ROT_CD = 0
        r_rsMerge!ROT_WEEK = 0
        r_rsMerge!TRADE_EXC_CD = v_rsActualTime!TRADE_EXC_CD
        r_rsMerge.Update
        
        v_rsActualTime.MoveNext
    Wend
    
    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "AddActualTime"
    
    
    
End Sub

Private Sub AddSchedule( _
    ByVal v_rsSchedule As ADODB.Recordset, _
    ByVal v_datLastPayEnd As Date, _
    ByVal v_datLastDayOpenPP As Date, _
    ByVal v_datEnd As Date, _
    ByVal v_rsActualTime As ADODB.Recordset, _
    ByRef r_rsMerge As ADODB.Recordset)
    
    Dim l_datChargeDate As Date
    Dim l_lngEmpNum As Long
    Dim l_strFilter As String
    Dim l_lngOTOppNum As Long
    
    On Error GoTo eh
    
    If Not RecordsExist(v_rsSchedule) Then
        Exit Sub
    End If
    
    v_rsSchedule.Sort = "EMP_NUM, CHARGE_DATE, SHIFT_NUM"
    v_rsSchedule.MoveFirst
    
    While Not v_rsSchedule.EOF
        l_datChargeDate = v_rsSchedule!CHARGE_DATE
        l_lngOTOppNum = v_rsSchedule!OT_OPP_NUM
        If l_datChargeDate > v_datLastDayOpenPP Then
            AddScheduleFields v_rsSchedule, r_rsMerge
        Else
            l_lngEmpNum = v_rsSchedule!EMP_NUM
            l_strFilter = "EMP_NUM = " & str(l_lngEmpNum) & _
                " AND CHARGE_DATE = " & str(l_datChargeDate) & _
                " AND OT_OPP_NUM =  " & str(l_lngOTOppNum)
            l_strFilter = MDAC25FilterFix(l_strFilter)
            v_rsActualTime.Filter = l_strFilter
            If Not RecordsExist(v_rsActualTime) Then
                AddScheduleFields v_rsSchedule, r_rsMerge
            Else
'                FillInMissingTime v_rsSchedule, v_rsActualTime, r_rsMerge
                'Do I try to fill in missing time with schedule
            End If
            
        End If
        v_rsSchedule.MoveNext
    Wend
    
    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "AddSchedule"
        
End Sub
Private Sub AddScheduleFields( _
    ByVal v_rsSchedule As ADODB.Recordset, _
    ByRef r_rsMerge As ADODB.Recordset, _
    Optional v_strStartTime As String = "", _
    Optional v_strEndTime As String = "")
    
    Dim l_decHours As Variant
    Dim l_strSystemExcCd As String
    Dim l_intSchedSource As Integer
    
    On Error GoTo eh
    
    l_strSystemExcCd = Trim(v_rsSchedule!SYSTEM_EXC_CD)
    If l_strSystemExcCd = "SLT" Then
        Exit Sub
    End If
    l_intSchedSource = v_rsSchedule!SCHED_SRC_CD
    r_rsMerge.AddNew
    r_rsMerge!TIME_TYPE = "3"
    r_rsMerge!EMP_NUM = v_rsSchedule!EMP_NUM
    r_rsMerge!CHARGE_DATE = CDate(v_rsSchedule!CHARGE_DATE)
    r_rsMerge!DATE_WORKED = CDate(v_rsSchedule!DATE_WORKED)
    r_rsMerge!DISTRIBUTION_ID = ""
    r_rsMerge!DISTRIBUTION_DESC = ""
    If v_strStartTime = "" Then
        r_rsMerge!Hours = CDec(v_rsSchedule!Hours)
        r_rsMerge!START_TIME = v_rsSchedule!START_TIME
        r_rsMerge!END_TIME = v_rsSchedule!END_TIME
    Else
        r_rsMerge!START_TIME = v_strStartTime
        r_rsMerge!END_TIME = v_strEndTime
        l_decHours = CDec(Left(v_strEndTime, 2)) + CDec(Right(v_strEndTime, 2)) / 60 _
            - CDec(Left(v_strStartTime, 2)) - CDec(Right(v_strStartTime, 2)) / 60
        r_rsMerge!Hours = l_decHours
    End If
    If r_rsMerge!END_TIME = "24:00" Then
        r_rsMerge!END_TIME = "00:00"
    End If
    r_rsMerge!EXCEPTION_CD = v_rsSchedule!EXCEPTION_CD
    r_rsMerge!OVER_EXC_CD = v_rsSchedule!OVER_EXC_CD
    If Trim(v_rsSchedule!SCHED_JOB_NUM) <> "" Then
        If Trim(v_rsSchedule!SCHED_JOB_NUM) <> Trim(v_rsSchedule!EMP_JOB_NUM) Then
            r_rsMerge!JOB_NUM = v_rsSchedule!SCHED_JOB_NUM
        End If
    End If
    r_rsMerge!WG_NUM = 0
    r_rsMerge!OT_REASON_DESC = ""
    r_rsMerge!OT_OPP_DESC = ""
    r_rsMerge!CASE_NUM = v_rsSchedule!CASE_NUM
    r_rsMerge!FMLA_IND = v_rsSchedule!FMLA_IND
    r_rsMerge!COMP_SHIFT_CD = v_rsSchedule!COMP_SHIFT_CD
    r_rsMerge!COMP_SHIFT_CD_AMT = 0
    r_rsMerge!COMP_CD = v_rsSchedule!COMP_CD
    r_rsMerge!REL_CD = v_rsSchedule!REL_CD
    r_rsMerge!SCHED_CHG_EXC_CD = v_rsSchedule!SCHED_CHG_EXC_CD
    r_rsMerge!DISTR_DEPT_NUM = ""
    r_rsMerge!DEPT_NUM = ""
    r_rsMerge!LOC_CD = 0
    r_rsMerge!OT_OPP_NUM = v_rsSchedule!OT_OPP_NUM
    r_rsMerge!OT_REASON_CD = ""
    If v_rsSchedule!TRAV_PAY_HOURS > 0 Then
        r_rsMerge!TRAVEL_PAY_IND = "Y"
    Else
        r_rsMerge!TRAVEL_PAY_IND = "N"
    End If
    r_rsMerge!APPR_BY_EMP_NUM = 0
    r_rsMerge!APPR_DATE = CDate("0100-01-01")
    r_rsMerge!AMOUNT = CDec(0)
    r_rsMerge!START_LOC_CD = 0
    r_rsMerge!END_LOC_CD = 0
    r_rsMerge!OT_BANK_CHARGE_DATE = CDate("0100-01-01")
    r_rsMerge!LINE_NUM = v_rsSchedule!LINE_NUM
    r_rsMerge!REQ_LINE_IND = v_rsSchedule!REQ_LINE_IND
    r_rsMerge!SCHED_CD = v_rsSchedule!SCHED_CD
    r_rsMerge!SHIFT_NUM = v_rsSchedule!SHIFT_NUM
    r_rsMerge!SHIFT_CODE = v_rsSchedule!SHIFT_CODE
    r_rsMerge!SCHED_SRC_CD = l_intSchedSource
    r_rsMerge!SHIFT_HOURS = v_rsSchedule!SHIFT_HOURS
    r_rsMerge!ADJ_TYPE = ""
    r_rsMerge!SEQ_NUM = 0
    r_rsMerge!HOURLY_RATE = CDec(0)
    r_rsMerge!REQ_BY_EMP_NUM = 0
    r_rsMerge!REQ_DATE = CDate("0100-01-01")
    r_rsMerge!COMPLETED_BY_EMP = 0
    r_rsMerge!COMPLETED_DATE = CDate("0100-01-01")
    r_rsMerge!FINAL_APPROVER = 0
    r_rsMerge!FINAL_APPR_DATE = CDate("0100-01-01")
    r_rsMerge!SPECIAL_CHECK = ""
    r_rsMerge!ADJ_PREV_PAID_FL = ""
    r_rsMerge!REQ_REASON = ""
    r_rsMerge!LAST_MOD_EMP_NUM = 0
    r_rsMerge!INDEX_TIME = ""
    r_rsMerge!ROT_CD = v_rsSchedule!ROT_CD
    r_rsMerge!ROT_WEEK = v_rsSchedule!ROT_WEEK
    
    If l_intSchedSource = 50 Or l_intSchedSource = 30 Then
        r_rsMerge!TRADE_EXC_CD = l_strSystemExcCd
    Else
        r_rsMerge!TRADE_EXC_CD = ""
    End If
    r_rsMerge.Update
    
    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "AddScheduleFields"
    

    
    
End Sub

Private Sub FillInMissingTime( _
    ByVal v_rsSchedule As ADODB.Recordset, _
    ByVal v_rsActualTime As ADODB.Recordset, _
    ByRef r_rsMerge As ADODB.Recordset)
    
    
    Dim l_lngEmpNum As Long
    Dim l_datChargeDate As Date
    Dim l_strStartTime As String
    Dim l_strEndTime As String
    Dim l_datDateWorked As Date
    Dim l_decShiftHours As Variant
'''''    Dim l_strShiftStartTime As String
'''''    Dim l_strShiftEndTime As String
'''''    Dim l_datShiftStartDate As Date
'''''    Dim l_datShiftEndDate As Date
    Dim l_decSchedHours As Variant
    Dim l_decActTimeHours As Variant
    Dim l_strActStartTime As String
    Dim l_strActEndTime As String
    Dim l_strSystemExcCd As String
    Dim l_strFilter As String
    Dim l_intSplitType As Integer
    
    
    On Error GoTo eh
    
    l_strSystemExcCd = v_rsSchedule!SYSTEM_EXC_CD
    If l_strSystemExcCd = "SLT" Then
        Exit Sub
    End If
    
    l_lngEmpNum = v_rsSchedule!EMP_NUM
    l_datChargeDate = v_rsSchedule!CHARGE_DATE
    l_datDateWorked = v_rsSchedule!DATE_WORKED
    l_strStartTime = v_rsSchedule!START_TIME
    l_strEndTime = v_rsSchedule!END_TIME
    l_decSchedHours = CDec(v_rsSchedule!Hours)
    If l_strEndTime = "24:00" Then
        l_strEndTime = "00:00"
    End If
    
'Does schedule match actual time
    l_strFilter = "EMP_NUM = " & str(l_lngEmpNum) & _
        " AND CHARGE_DATE = " & str(l_datChargeDate) & _
        " AND DATE_WORKED = " & str(l_datDateWorked) & _
        " AND START_TIME = '" & l_strStartTime & "'" & _
        " AND END_TIME = '" & l_strEndTime & "'" & _
        " AND OT_OPP_NUM = 0"
    l_strFilter = MDAC25FilterFix(l_strFilter)
    v_rsActualTime.Filter = l_strFilter
    If RecordsExist(v_rsActualTime) Then
        'Schedule time is on actual time
        Exit Sub
    End If
    
    If l_strEndTime = "00:00" Then
        l_strEndTime = "24:00"
    End If
    'Does the Schedule Overlapp any actual time
    l_strFilter = "EMP_NUM = " & str(l_lngEmpNum) & _
        " AND CHARGE_DATE = " & str(l_datChargeDate) & _
        " AND DATE_WORKED = " & str(l_datDateWorked) & _
        " AND START_TIME <= '" & l_strEndTime & "'" & _
        " AND END_TIME >= '" & l_strStartTime & "'" & _
        " AND OT_OPP_NUM = 0"
    l_strFilter = MDAC25FilterFix(l_strFilter)
    v_rsActualTime.Filter = l_strFilter
    If Not RecordsExist(v_rsActualTime) Then
        'No Overlapp
        AddScheduleFields v_rsSchedule, r_rsMerge
        Exit Sub
    End If
    
    l_decActTimeHours = CDec(0)
    v_rsActualTime.MoveFirst
    While Not v_rsActualTime.EOF
        l_decActTimeHours = l_decActTimeHours + CDec(v_rsActualTime!Hours)
        v_rsActualTime.MoveNext
    Wend
    
    If l_decSchedHours = l_decActTimeHours Then
        'Actual Time Splits equal the schedule
        Exit Sub
    End If
    
    'Some part of the schedule is missing
    'Assumes schedule and time can not cross physical days
    v_rsActualTime.Sort = "START_TIME"
    v_rsActualTime.MoveFirst
    While Not v_rsActualTime.EOF
        l_strActStartTime = v_rsActualTime!START_TIME
        l_strActEndTime = v_rsActualTime!END_TIME
        If l_strActEndTime = "00:00" Then
            l_strActEndTime = "24:00"
        End If
        l_intSplitType = DeterimineSplit( _
            l_strStartTime, l_strEndTime, _
            l_strActStartTime, l_strActEndTime)
        Select Case l_intSplitType
            Case 0
            'Do nothing
            Case 1
                AddScheduleFields v_rsSchedule, r_rsMerge, _
                    l_strStartTime, l_strActStartTime
            Case 2
                AddScheduleFields v_rsSchedule, r_rsMerge, _
                    l_strActEndTime, l_strEndTime
            Case 3
                AddScheduleFields v_rsSchedule, r_rsMerge, _
                    l_strStartTime, l_strActStartTime
                AddScheduleFields v_rsSchedule, r_rsMerge, _
                    l_strActEndTime, l_strEndTime
        End Select
        
        v_rsActualTime.MoveNext
    Wend
    
    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "FillInMissingTime"

End Sub

Private Function DeterimineSplit( _
    ByVal v_strStartTime1 As String, _
    ByVal v_strEndTime1 As String, _
    ByVal v_strStartTime2 As String, _
    ByVal v_strEndTime2 As String) As Integer
    
    Dim l_blnFillInFront As Boolean
    Dim l_blnFillInRear As Boolean
    
    On Error GoTo eh
    
    If v_strStartTime1 < v_strStartTime2 Then
        l_blnFillInFront = True
    Else
        l_blnFillInFront = False
    End If
    
    If v_strEndTime1 > v_strEndTime2 Then
        l_blnFillInRear = True
    Else
        l_blnFillInRear = False
    End If
    
    If l_blnFillInFront Then
        If l_blnFillInRear Then
            DeterimineSplit = 3
        Else
            DeterimineSplit = 1
        End If
        Exit Function
    Else
        If l_blnFillInRear Then
            DeterimineSplit = 2
            Exit Function
        End If
    End If
    
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "DeterimineSplit"

    
    
End Function



Private Function GetAdjustmentDetails( _
    ByVal v_aryEmployees As Variant, _
    ByVal v_datStart As Date, _
    ByVal v_datEnd As Date _
    ) As ADODB.Recordset
        
   Dim l_obj As dbAdjustments.cdAdjustments
   
   On Error GoTo eh
   
   Set l_obj = CreateObject("dbAdjustments.cdAdjustments")
    
   
    Set GetAdjustmentDetails = l_obj.GetAdjustmentDetails( _
        v_aryEmployees, v_datStart, v_datEnd)
        
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "GetAdjustmentDetails"
       
    
End Function


Public Function GetScheduledExceptions( _
    ByVal v_aryEmployees As Variant, _
    ByVal v_datStart As Date, _
    ByVal l_datEnd As Date _
    ) As ADODB.Recordset
    
        
   Dim l_rsExc As ADODB.Recordset
   Dim l_rsReq As ADODB.Recordset
   Dim l_rsSchedExc As ADODB.Recordset
   
   On Error GoTo eh
   
    Set l_rsSchedExc = CreateMaintainExceptionsRS()
    Set l_rsExc = GetScheduledExceptionsDB(v_aryEmployees, v_datStart, l_datEnd)
    If Not RecordsExist(l_rsExc) Then
        Set GetScheduledExceptions = l_rsSchedExc
        Set l_rsExc = Nothing
        Set l_rsSchedExc = Nothing
        Exit Function
    End If
    
    Set l_rsReq = BuildGABRequest(l_rsExc, l_datEnd)
    Set m_rsSchedule = GetSchedOnExceptionDays(l_rsReq)
    Set l_rsReq = Nothing
    l_rsExc.Sort = "EMP_NUM, CHARGE_DATE, SHIFT_NUM"
    l_rsExc.MoveFirst
    While Not l_rsExc.EOF
        AddException l_rsExc, l_rsSchedExc
        l_rsExc.MoveNext
    Wend
    
    Set GetScheduledExceptions = l_rsSchedExc
   
    Exit Function
        
eh:
    TrsRaiseError m_strCMODULENAME, "GetScheduledExceptions", , , , , True
  
End Function

Private Function GetScheduledExceptionsDB( _
    ByVal v_aryEmployees As Variant, _
    ByVal v_datStart As Date, _
    ByVal l_datEnd As Date _
    ) As ADODB.Recordset
    
        
   Dim l_obj As dbSchedException.cdSchedException
   
   On Error GoTo eh
   
   Set l_obj = CreateObject("dbSchedException.cdSchedException")
   
   Set GetScheduledExceptionsDB = l_obj.GetScheduledExceptions( _
    v_aryEmployees, v_datStart, l_datEnd)
   
    Exit Function
        
eh:
    TrsRaiseError m_strCMODULENAME, "GetScheduledExceptionsDB"
  
End Function


Private Function CreateMaintainExceptionsRS() As ADODB.Recordset

    Dim l_rs As ADODB.Recordset
    
    On Error GoTo eh
    
    Set l_rs = New ADODB.Recordset
    With l_rs
      .Fields.Append "EMP_NUM", adInteger, , adFldUpdatable
      .Fields.Append "CHARGE_DATE", adDate, , adFldUpdatable
      .Fields.Append "SHIFT_NUM", adInteger, , adFldUpdatable
      .Fields.Append "FULL_SHIFT_EXC_IND", adChar, 1, adFldUpdatable
      .Fields.Append "START_TIME", adChar, 5, adFldUpdatable
      .Fields.Append "END_TIME", adChar, 5, adFldUpdatable
      .Fields.Append "EXCEPTION_CD", adChar, 3, adFldUpdatable
      .Fields.Append "OVER_EXC_CD", adChar, 3, adFldUpdatable
      .Fields.Append "FMLA_IND", adChar, 1, adFldUpdatable
      .Fields.Append "CASE_NUM", adInteger, , adFldUpdatable
      .Fields.Append "WG_NUM", adInteger, 7, adFldUpdatable
      .Fields.Append "JOB_NUM", adChar, 6, adFldUpdatable
      .Fields.Append "RELIEF_COMP", adChar, 3, adFldUpdatable
      .Fields.Append "COMP_CD", adChar, 3, adFldUpdatable
      .Fields.Append "SHIFT_COMP", adChar, 3, adFldUpdatable
      .Fields.Append "REC_STATUS", adChar, 1, adFldUpdatable
      .Fields.Append "DELETE_ACTUAL_TIME", adChar, 1, adFldUpdatable
      .Fields.Append "APPR_AUTH_IND", adChar, 1, adFldUpdatable
      .Fields.Append "CREATED_TIMESTAMP", adChar, 26, adFldUpdatable
      .Fields.Append "CREATED_EMP_NUM", adInteger, , adFldUpdatable
      .Fields.Append "LAST_MOD_EMP_NUM", adInteger, , adFldUpdatable
      .Fields.Append "LAST_MOD_TIMESTAMP", adChar, 26, adFldUpdatable
      .Fields.Append "HOURS", adDecimal, adFldUpdatable
      .Fields.Item("HOURS").Precision = 3
      .Fields.Item("HOURS").NumericScale = 1
      .Fields.Append "EXC_START_DATE", adDate, , adFldUpdatable
      .CursorLocation = adUseClient
      .CursorType = adOpenDynamic
      .LockType = adLockOptimistic
      .Open
    End With
    
    Set CreateMaintainExceptionsRS = l_rs
    
    Exit Function
        
eh:
    TrsRaiseError m_strCMODULENAME, "CreateMaintainExceptionsRS"
  
    

End Function


Private Sub AddException( _
    ByVal v_rsExc As ADODB.Recordset, _
    ByRef r_rsSchedExc As ADODB.Recordset)
    
    Dim i As Integer
    Dim l_strField As String
    Dim l_decOffSetHours As Variant
    Dim l_decHours As Variant
    Dim l_decShiftHours As Variant
    Dim l_strStartTime As String
    Dim l_strEndTime As String
    Dim l_lngEmpNum As Long
    Dim l_datChargeDate As Date
    Dim l_lngShiftNum As Long
    Dim l_strException As String
    Dim l_strFilter As String
    Dim l_datStartDate As Date
    Dim l_colErrArgs As Collection
    
    
    On Error GoTo eh
    
    r_rsSchedExc.AddNew
    For i = 1 To v_rsExc.Fields.Count
        l_strField = v_rsExc.Fields(i - 1).Name
        Select Case l_strField
            Case "HOURS_OFFSET_NUM", _
                 "SCHED_CD", "SCHED_SRC_CD"
            'Fields not in output
            Case Else
                r_rsSchedExc.Fields(l_strField) = v_rsExc.Fields(l_strField)
        End Select
    Next i
    r_rsSchedExc!DELETE_ACTUAL_TIME = "N"
    r_rsSchedExc!APPR_AUTH_IND = "N"
    r_rsSchedExc!REC_STATUS = "L"
    
    l_decHours = CDec(v_rsExc!Hours)
    l_decOffSetHours = CDec(v_rsExc!HOURS_OFFSET_NUM)
    l_strException = Trim(v_rsExc!EXCEPTION_CD)
    
    m_rsSchedule.Filter = ""
    l_lngEmpNum = v_rsExc!EMP_NUM
    l_datChargeDate = v_rsExc!CHARGE_DATE
    l_lngShiftNum = v_rsExc!SHIFT_NUM
    l_strFilter = "EMP_NUM = " & str(l_lngEmpNum)
    l_strFilter = l_strFilter & "  AND  "
    l_strFilter = l_strFilter & "CHARGE_DATE = " & str(l_datChargeDate)
    l_strFilter = l_strFilter & "  AND  "
    l_strFilter = l_strFilter & "SHIFT_NUM = " & str(l_lngShiftNum)
    m_rsSchedule.Filter = MDAC25FilterFix(l_strFilter)
    If Not RecordsExist(m_rsSchedule) Then
            Err.Raise trErrbusShiftInvalidStartTime, m_strCMODULENAME & "." & "AddException", "Schedule does not match exception"
    End If
    l_decShiftHours = m_rsSchedule!SHIFT_HOURS
    If (l_decHours + l_decOffSetHours) _
        > l_decShiftHours Then
        If l_decShiftHours = 0 And _
            CanBeScheduledOnOffDay(l_strException) Then
            r_rsSchedExc!START_TIME = "08:00"
            r_rsSchedExc!END_TIME = "16:00"
            r_rsSchedExc!FULL_SHIFT_EXC_IND = "Y"
            r_rsSchedExc!EXC_START_DATE = l_datChargeDate
        Else
            Err.Raise trErrbusShiftInvalidStartTime, m_strCMODULENAME & "." & "AddException", "Schedule does not match exception"
        End If
    Else
        If l_decHours = l_decShiftHours Then
            r_rsSchedExc!START_TIME = m_rsSchedule!SHIFT_START_TIME
            r_rsSchedExc!END_TIME = m_rsSchedule!SHIFT_END_TIME
            r_rsSchedExc!FULL_SHIFT_EXC_IND = "Y"
            r_rsSchedExc!EXC_START_DATE = m_rsSchedule!SHIFT_START_DATE
        Else
            DetermineExceptionTimes l_strStartTime, l_strEndTime, l_datStartDate, m_rsSchedule, v_rsExc
            r_rsSchedExc!START_TIME = l_strStartTime
            r_rsSchedExc!END_TIME = l_strEndTime
            r_rsSchedExc!FULL_SHIFT_EXC_IND = "N"
            r_rsSchedExc!EXC_START_DATE = l_datStartDate
        End If
    End If
    
    r_rsSchedExc.Update
    
    Exit Sub
    
eh:
    Set l_colErrArgs = New Collection
    l_colErrArgs.Add m_strCMODULENAME & ":AddException arguments"
    l_colErrArgs.Add "l_lngEmpNum=" & CStr(l_lngEmpNum)
    l_colErrArgs.Add "l_datChargeDate=" & Format(l_datChargeDate, "mm/dd/yyyy")
    l_colErrArgs.Add "l_lngShiftNum=" & CStr(l_lngShiftNum)
    l_colErrArgs.Add "l_strException=" & l_strException
    l_colErrArgs.Add "l_decHours=" & CStr(l_decHours)
    l_colErrArgs.Add "l_decOffSetHours=" & CStr(l_decOffSetHours)
    If Err.Number = trErrbusShiftInvalidStartTime Then
        TrsRaiseError m_strCMODULENAME, "AddException", , , , l_colErrArgs, True
    Else
        TrsRaiseError m_strCMODULENAME, "AddException", , , , l_colErrArgs
    End If
    
End Sub

Private Sub DetermineExceptionTimes( _
    ByRef r_strExceptionStartTime As String, _
    ByRef r_strExceptionEndTime As String, _
    ByRef r_strExceptionStartDate As Date, _
    ByVal v_rsAdjustedBaseSched As ADODB.Recordset, _
    ByVal v_rsExceptionSched As ADODB.Recordset)
    
    
    Dim l_decExceptionHours As Variant
    Dim l_decExceptionOffSetHours As Variant
    Dim l_decShiftHours As Variant
    Dim l_strShiftStartTime As String
    Dim l_strShiftEndTime As String
    Dim l_datShiftStartDate As Date
    Dim l_datShiftStartDateTime As Date
    Dim l_datShiftEndDate As Date
    Dim l_intExceptionHours As Integer
    Dim l_intExceptionMinutes As Integer
    Dim l_datExceptionStartDateTime As Date
    Dim l_datExceptionEndDateTime As Date
    Dim l_strLunchStartTime As String
    Dim l_strLunchEndTime As String
    Dim l_decHoursOfLunch As Variant
    Dim l_decHoursBeforeLunch As Variant
    Dim l_datLunchStartDate As Date
    Dim l_datLunchStartDateTime As Date
    Dim l_datLunchEndDate As Date
    Dim l_datLunchEndDateTime As Date
    Dim l_decExceptionPhysicalHours As Variant
    

    On Error GoTo eh

    l_datShiftStartDate = v_rsAdjustedBaseSched!SHIFT_START_DATE
    l_strShiftStartTime = Trim(v_rsAdjustedBaseSched!SHIFT_START_TIME)
    l_strLunchStartTime = Trim(v_rsAdjustedBaseSched!SHIFT_LUNCH_START_TIME)
    l_strLunchEndTime = Trim(v_rsAdjustedBaseSched!SHIFT_LUNCH_END_TIME)
    If l_strLunchStartTime = "" Or l_strLunchEndTime = "" Then
        l_strLunchStartTime = l_strShiftStartTime
        l_strLunchEndTime = l_strShiftStartTime
    End If
    l_decExceptionOffSetHours = CDec(v_rsExceptionSched!HOURS_OFFSET_NUM)
    l_decExceptionHours = v_rsExceptionSched!Hours
    If l_strShiftStartTime > l_strLunchStartTime Then
        l_datLunchStartDate = DateAdd("d", 1, l_datShiftStartDate)
    Else
        l_datLunchStartDate = l_datShiftStartDate
    End If
'There are currently no lunches that cross midnight (2002-04-12)
    If l_strLunchStartTime > l_strLunchEndTime Then
        l_datLunchEndDate = DateAdd("d", 1, l_datLunchStartDate)
    Else
        l_datLunchEndDate = l_datLunchStartDate
    End If
    l_datShiftStartDateTime = DateTimeFrom(l_datShiftStartDate, l_strShiftStartTime)
    l_datLunchStartDateTime = DateTimeFrom(l_datLunchStartDate, l_strLunchStartTime)
    l_decHoursBeforeLunch = CDec(DateDiff("n", l_datShiftStartDateTime, l_datLunchStartDateTime) / 60)
    l_datLunchEndDateTime = DateTimeFrom(l_datLunchEndDate, l_strLunchEndTime)
    l_decHoursOfLunch = CDec(DateDiff("n", l_datLunchStartDateTime, l_datLunchEndDateTime) / 60)
    
    
    If l_decExceptionOffSetHours >= l_decHoursBeforeLunch Then
    'Exception starts after lunch
        l_datExceptionStartDateTime = DateAdd("n", ((l_decExceptionOffSetHours + l_decHoursOfLunch) * 60), l_datShiftStartDateTime)
        l_datExceptionEndDateTime = DateAdd("n", l_decExceptionHours * 60, l_datExceptionStartDateTime)
    Else
    'Exception starts before lunch
        l_datExceptionStartDateTime = DateAdd("n", (l_decExceptionOffSetHours * 60), l_datShiftStartDateTime)
        'Does exception end before lunch
        If (l_decExceptionHours + l_decExceptionOffSetHours) <= l_decHoursBeforeLunch Then
        'Yes have the correct hours to calcualate end time
            l_decExceptionPhysicalHours = l_decExceptionHours
        Else
            'NO, must add lunch time to calculate end time
            l_decExceptionPhysicalHours = l_decHoursOfLunch + l_decExceptionHours
        End If
        l_datExceptionEndDateTime = _
                DateAdd("n", l_decExceptionPhysicalHours * 60, _
                          l_datExceptionStartDateTime)
    End If

    r_strExceptionStartTime = Format(l_datExceptionStartDateTime, "HH:MM")
    r_strExceptionEndTime = Format(l_datExceptionEndDateTime, "HH:MM")
    r_strExceptionStartDate = l_datExceptionStartDateTime
    
    
    Exit Sub
    
eh:

    TrsRaiseError m_strCMODULENAME, "DetermineExceptionTimes"
    
    
End Sub


Private Function CanBeScheduledOnOffDay( _
    ByVal v_strExcCd As String _
    ) As Boolean
    
    On Error GoTo eh
    
    Select Case Trim(v_strExcCd)
        Case "H", "HR", "HB"
            CanBeScheduledOnOffDay = True
        Case Else
            CanBeScheduledOnOffDay = False
    End Select
    Exit Function
eh:

    TrsRaiseError m_strCMODULENAME, "CanBeScheduledOnOffDay"
    
End Function
Private Sub DetermineSchedChgPayments( _
    ByVal v_aryEmployees As Variant, _
    ByVal v_datStart As Date, _
    ByVal v_datEnd As Date, _
    ByRef r_rsExceptions As ADODB.Recordset)
    
    Dim l_rsPayments As ADODB.Recordset
    Dim l_lngEmpNum As Long
    Dim l_datChargeDate As Date
    Dim l_strFilter As String
           
    On Error GoTo eh
    
    Set l_rsPayments = GetPayments(v_aryEmployees, v_datStart, v_datEnd)
    
    If Not RecordsExist(l_rsPayments) Then
        Exit Sub
    End If
    l_rsPayments.Sort = "EMP_NUM, CHARGE_DATE"
    l_rsPayments.MoveFirst
    While Not l_rsPayments.EOF
        l_lngEmpNum = l_rsPayments!EMP_NUM
        l_datChargeDate = l_rsPayments!CHARGE_DATE
        l_strFilter = "EMP_NUM = " & str(l_lngEmpNum) _
            & " AND CHARGE_DATE = " & str(l_datChargeDate)
        l_strFilter = MDAC25FilterFix(l_strFilter)
        r_rsExceptions.Filter = l_strFilter
        If RecordsExist(r_rsExceptions) Then
            While Not r_rsExceptions.EOF
                r_rsExceptions!SCHED_CHANGE_PMT = "Y"
                r_rsExceptions.Update
                r_rsExceptions.MoveNext
            Wend
        Else
            r_rsExceptions.AddNew
            r_rsExceptions!EMP_NUM = l_lngEmpNum
            r_rsExceptions!CHARGE_DATE = l_datChargeDate
            r_rsExceptions!EXCEPTION_CD = l_rsPayments!EXCEPTION_CD
            r_rsExceptions!OVER_EXC_CD = ""
            r_rsExceptions!CASE_GROUP = ""
            r_rsExceptions!BANK_CD = "N"
            r_rsExceptions!FULL_SHIFT_EXC_IND = "N"
            r_rsExceptions!Hours = 0
            r_rsExceptions!SHIFT_NUM = 0
            r_rsExceptions!SYMBOL_CD = ""
            r_rsExceptions!SHIFT_CODE = ""
            r_rsExceptions!SCHED_CHANGE_PMT = "Y"
            r_rsExceptions.Update
        End If
        l_rsPayments.MoveNext
    Wend
    
    Exit Sub
eh:

    TrsRaiseError m_strCMODULENAME, "DetermineSchedChgPayments"
    
End Sub
Private Function GetPayments( _
    ByVal v_aryEmployees As Variant, _
    ByVal v_datStart As Date, _
    ByVal v_datEnd As Date _
    ) As ADODB.Recordset
    

    Dim l_obj As dbSchedule.cdSchedule

    On Error GoTo eh
    
    Set l_obj = CreateObject("dbSchedule.cdSchedule")
    
    Set GetPayments = l_obj.GetScheduleChangePayments(v_aryEmployees, v_datStart, v_datEnd)
    
    Set l_obj = Nothing
    
    Exit Function
eh:

    TrsRaiseError m_strCMODULENAME, "GetPayments"
End Function



Private Function FindDayOffRec( _
    ByVal v_lngEmpNum As Long, _
    ByVal v_datChargeDate As Date, _
    ByRef r_rsDaysOff As ADODB.Recordset) _
    As Boolean
    
    
    Dim l_lngEmpNum As Long
    Dim l_datChargeDate As Date
    Dim l_blnDayOffFound As Boolean
    Dim l_blnExitLoop As Boolean
    
    On Error GoTo eh
    
    FindDayOffRec = False
    l_blnDayOffFound = False
    If Not RecordsExist(r_rsDaysOff) Then
        Exit Function
    End If
    
    If r_rsDaysOff.EOF Then
        Exit Function
    End If
    
    l_lngEmpNum = r_rsDaysOff!EMP_NUM
    l_datChargeDate = r_rsDaysOff!CHARGE_DATE
    If l_lngEmpNum = v_lngEmpNum And _
        l_datChargeDate = v_datChargeDate Then
        FindDayOffRec = True
        Exit Function
    End If

    l_blnExitLoop = False
    While Not r_rsDaysOff.EOF And Not l_blnDayOffFound And Not l_blnExitLoop
        l_lngEmpNum = r_rsDaysOff!EMP_NUM
        l_datChargeDate = r_rsDaysOff!CHARGE_DATE
        If l_lngEmpNum > v_lngEmpNum Then
            l_blnExitLoop = True
        ElseIf l_lngEmpNum < v_lngEmpNum Then
            r_rsDaysOff.MoveNext
        Else
            If l_datChargeDate > v_datChargeDate Then
                l_blnExitLoop = True
            ElseIf l_datChargeDate < v_datChargeDate Then
                r_rsDaysOff.MoveNext
            Else
                l_blnDayOffFound = True
            End If
        End If
    Wend
    
    FindDayOffRec = l_blnDayOffFound
        
    Exit Function
        
eh:
    TrsRaiseError m_strCMODULENAME, "FindDayOffRec"
    
End Function

'//for older TRIS clients
Public Function GetViewCalendarDetailsByEmployee( _
    ByVal v_aryEmployees As Variant, _
    ByVal v_datStartDate As Date, _
    ByVal v_datEndDate As Date, _
    Optional v_datLastClosedPPEndDate As Date = CDate("1980-01-01"), _
    Optional v_datLastDayOfOpenPP As Date = CDate("1980-01-01") _
    ) As ADODB.Recordset

    On Error GoTo eh
    
    Set GetViewCalendarDetailsByEmployee = GetViewCalendarDetailsByEmployeePrivate(v_aryEmployees, v_datStartDate, v_datEndDate, v_datLastClosedPPEndDate, v_datLastDayOfOpenPP, 39)
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "GetViewCalendarDetailsByEmployee", , , , , True
End Function

'//for newer TRIS clients DFE 48 chars, OTTC pick list description increased to 60 chars
Public Function GetViewCalendarDetailsByEmployee2B( _
    ByVal v_aryEmployees As Variant, _
    ByVal v_datStartDate As Date, _
    ByVal v_datEndDate As Date, _
    Optional v_datLastClosedPPEndDate As Date = CDate("1980-01-01"), _
    Optional v_datLastDayOfOpenPP As Date = CDate("1980-01-01"), _
    Optional ByVal v_intDescriptionLength As Integer = 48 _
    ) As ADODB.Recordset

    On Error GoTo eh
    
    Set GetViewCalendarDetailsByEmployee2B = GetViewCalendarDetailsByEmployeePrivate(v_aryEmployees, v_datStartDate, v_datEndDate, v_datLastClosedPPEndDate, v_datLastDayOfOpenPP, v_intDescriptionLength)
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "GetViewCalendarDetailsByEmployee2B", , , , , True
End Function

Private Function GetViewCalendarDetailsByEmployeePrivate( _
    ByVal v_aryEmployees As Variant, _
    ByVal v_datStartDate As Date, _
    ByVal v_datEndDate As Date, _
    Optional v_datLastClosedPPEndDate As Date = CDate("1980-01-01"), _
    Optional v_datLastDayOfOpenPP As Date = CDate("1980-01-01"), _
    Optional ByVal v_intDescriptionLength As Integer = 48 _
    ) As ADODB.Recordset

    On Error GoTo eh

    Dim l_lngEmpNum As Long
    Dim l_datStart As Date
    Dim l_datEnd As Date
    Dim l_datLastPayEnd As Date
    Dim l_datLastDayOpenPP As Date
    Dim l_rsLockStatus As ADODB.Recordset
    Dim l_rsSchedule As ADODB.Recordset
    Dim l_rsActualTime As ADODB.Recordset
    Dim l_rsAdjustments As ADODB.Recordset
    Dim l_rsOpenActualTime As ADODB.Recordset
    Dim l_rsMerge As ADODB.Recordset
    Dim l_strFilter As String
    
    l_datStart = v_datStartDate
    l_datEnd = v_datEndDate
    
    If v_datLastClosedPPEndDate < l_datStart Then
        l_datLastPayEnd = l_datStart
    Else
        l_datLastPayEnd = v_datLastClosedPPEndDate
    End If
    If v_datLastDayOfOpenPP < l_datStart Then
        l_datLastDayOpenPP = v_datEndDate
    Else
        l_datLastDayOpenPP = v_datLastDayOfOpenPP
    End If
    
    
    Set l_rsAdjustments = GetAdjustmentDetails(v_aryEmployees, l_datStart, l_datEnd)
    Set l_rsActualTime = GetActualTime(v_aryEmployees, l_datStart, l_datLastDayOpenPP, v_intDescriptionLength)
    If l_datLastPayEnd < l_datEnd Then
        If l_datLastPayEnd > l_datStart Then
            Set l_rsSchedule = GetSched(v_aryEmployees, DateAdd("d", 1, l_datLastPayEnd), l_datEnd, trAdjBaseSchedWithExceptionsAndOvertime)
        Else
            Set l_rsSchedule = GetSched(v_aryEmployees, l_datStart, l_datEnd, trAdjBaseSchedWithExceptionsAndOvertime)
        End If
    Else
'        Set l_rsSchedule = Nothing
        Set l_rsSchedule = GetSched(v_aryEmployees, l_datStart, l_datEnd, trAdjBaseSchedWithExceptionsAndOvertime)
    End If
    If RecordsExist(l_rsActualTime) Then
        If l_datLastPayEnd = l_datStart Then
            l_strFilter = "CHARGE_DATE >= " & str(l_datLastPayEnd) & _
                " AND CHARGE_DATE <= " & str(l_datLastDayOpenPP)
        Else
            l_strFilter = "CHARGE_DATE > " & str(l_datLastPayEnd) & _
                " AND CHARGE_DATE <= " & str(l_datLastDayOpenPP)
        End If
        l_strFilter = MDAC25FilterFix(l_strFilter)
        l_rsActualTime.Filter = l_strFilter
        Set l_rsOpenActualTime = CopyRecordset(l_rsActualTime)
    Else
        Set l_rsOpenActualTime = l_rsActualTime
    End If
    l_rsActualTime.Filter = ""
    Set l_rsMerge = BuildMergeRS(l_datStart, l_datEnd, _
        l_datLastPayEnd, l_datLastDayOpenPP, l_rsAdjustments, _
        l_rsActualTime, l_rsSchedule, l_rsOpenActualTime)
        
    Set GetViewCalendarDetailsByEmployeePrivate = l_rsMerge
    Set l_rsMerge = Nothing
    Set l_rsActualTime = Nothing
    Set l_rsOpenActualTime = Nothing
    Set l_rsSchedule = Nothing
    Set l_rsAdjustments = Nothing
    Set l_rsLockStatus = Nothing
    Exit Function
    
eh:
    Set l_rsMerge = Nothing
    Set l_rsActualTime = Nothing
    Set l_rsOpenActualTime = Nothing
    Set l_rsSchedule = Nothing
    Set l_rsAdjustments = Nothing
    Set l_rsLockStatus = Nothing
    TrsRaiseError m_strCMODULENAME, "GetViewCalendarDetailsByEmployeePrivate", , , , , True
    
End Function

'//for older TRIS clients
Public Function GetViewCalendarDetailsByWorkGroup( _
    ByVal v_aryWorkGroups As Variant, _
    ByVal v_datStartDate As Date, _
    ByVal v_datEndDate As Date, _
    Optional v_datLastClosedPPEndDate As Date = CDate("1980-01-01"), _
    Optional v_datLastDayOfOpenPP As Date = CDate("1980-01-01") _
    ) As ADODB.Recordset

    On Error GoTo eh

    Set GetViewCalendarDetailsByWorkGroup = GetViewCalendarDetailsByWorkGroupPrivate(v_aryWorkGroups, v_datStartDate, v_datEndDate, v_datLastClosedPPEndDate, v_datLastDayOfOpenPP, 39)
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "GetViewCalendarDetailsByWorkGroup", , , , , True
End Function

'//for newer TRIS clients DFE 48 chars, OTTC pick list description increased to 60 chars
Public Function GetViewCalendarDetailsByWorkGroup2B( _
    ByVal v_aryWorkGroups As Variant, _
    ByVal v_datStartDate As Date, _
    ByVal v_datEndDate As Date, _
    Optional v_datLastClosedPPEndDate As Date = CDate("1980-01-01"), _
    Optional v_datLastDayOfOpenPP As Date = CDate("1980-01-01"), _
    Optional ByVal v_intDescriptionLength As Integer = 48 _
    ) As ADODB.Recordset

    On Error GoTo eh
    
    Set GetViewCalendarDetailsByWorkGroup2B = GetViewCalendarDetailsByWorkGroupPrivate(v_aryWorkGroups, v_datStartDate, v_datEndDate, v_datLastClosedPPEndDate, v_datLastDayOfOpenPP, v_intDescriptionLength)
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "GetViewCalendarDetailsByWorkGroup2B", , , , , True
End Function

Private Function GetViewCalendarDetailsByWorkGroupPrivate( _
    ByVal v_aryWorkGroups As Variant, _
    ByVal v_datStartDate As Date, _
    ByVal v_datEndDate As Date, _
    Optional v_datLastClosedPPEndDate As Date = CDate("1980-01-01"), _
    Optional v_datLastDayOfOpenPP As Date = CDate("1980-01-01"), _
    Optional ByVal v_intDescriptionLength As Integer = 48 _
    ) As ADODB.Recordset

    On Error GoTo eh

    Dim l_lngEmpNum As Long
    Dim l_datStart As Date
    Dim l_datEnd As Date
    Dim l_datLastPayEnd As Date
    Dim l_datLastDayOpenPP As Date
    Dim l_rsLockStatus As ADODB.Recordset
    Dim l_rsSchedule As ADODB.Recordset
    Dim l_rsActualTime As ADODB.Recordset
    Dim l_rsAdjustments As ADODB.Recordset
    Dim l_rsOpenActualTime As ADODB.Recordset
    Dim l_rsMerge As ADODB.Recordset
    Dim l_strFilter As String
    Dim l_aryEmployees As Variant
    
    l_datStart = v_datStartDate
    l_datEnd = v_datEndDate
    l_aryEmployees = EmployeesInWorkGroup(v_aryWorkGroups, l_datStart, l_datEnd)
    If IsEmpty(l_aryEmployees) Then
        Set GetViewCalendarDetailsByWorkGroupPrivate = Nothing
        GoTo Xit
    End If
    
    If v_datLastClosedPPEndDate < l_datStart Then
        l_datLastPayEnd = l_datStart
    Else
        l_datLastPayEnd = v_datLastClosedPPEndDate
    End If
    If v_datLastDayOfOpenPP < l_datStart Then
        l_datLastDayOpenPP = v_datEndDate
    Else
        l_datLastDayOpenPP = v_datLastDayOfOpenPP
    End If
    
    Set l_rsAdjustments = GetAdjustmentDetails(l_aryEmployees, l_datStart, l_datEnd)
    Set l_rsActualTime = GetActualTime(l_aryEmployees, l_datStart, l_datLastDayOpenPP, v_intDescriptionLength)
    If l_datLastPayEnd < l_datEnd Then
        If l_datLastPayEnd > l_datStart Then
            Set l_rsSchedule = GetSched(l_aryEmployees, DateAdd("d", 1, l_datLastPayEnd), l_datEnd, trAdjBaseSchedWithExceptionsAndOvertime)
        Else
            Set l_rsSchedule = GetSched(l_aryEmployees, l_datStart, l_datEnd, trAdjBaseSchedWithExceptionsAndOvertime)
        End If
    Else
'        Set l_rsSchedule = Nothing
        Set l_rsSchedule = GetSched(l_aryEmployees, l_datStart, l_datEnd, trAdjBaseSchedWithExceptionsAndOvertime)
    End If
    If RecordsExist(l_rsActualTime) Then
        If l_datLastPayEnd = l_datStart Then
            l_strFilter = "CHARGE_DATE >= " & str(l_datLastPayEnd) & _
                " AND CHARGE_DATE <= " & str(l_datLastDayOpenPP)
        Else
            l_strFilter = "CHARGE_DATE > " & str(l_datLastPayEnd) & _
                " AND CHARGE_DATE <= " & str(l_datLastDayOpenPP)
        End If
        l_strFilter = MDAC25FilterFix(l_strFilter)
        l_rsActualTime.Filter = l_strFilter
        Set l_rsOpenActualTime = CopyRecordset(l_rsActualTime)
    Else
        Set l_rsOpenActualTime = l_rsActualTime
    End If
    l_rsActualTime.Filter = ""
    Set l_rsMerge = BuildMergeRS(l_datStart, l_datEnd, _
        l_datLastPayEnd, l_datLastDayOpenPP, l_rsAdjustments, _
        l_rsActualTime, l_rsSchedule, l_rsOpenActualTime)
        
    Set GetViewCalendarDetailsByWorkGroupPrivate = l_rsMerge
    Set l_rsMerge = Nothing
    Set l_rsActualTime = Nothing
    Set l_rsOpenActualTime = Nothing
    Set l_rsSchedule = Nothing
    Set l_rsAdjustments = Nothing
    Set l_rsLockStatus = Nothing
Xit:
    Exit Function
    
eh:
    Set l_rsMerge = Nothing
    Set l_rsActualTime = Nothing
    Set l_rsOpenActualTime = Nothing
    Set l_rsSchedule = Nothing
    Set l_rsAdjustments = Nothing
    Set l_rsLockStatus = Nothing
    TrsRaiseError m_strCMODULENAME, "GetViewCalendarDetailsByWorkGroupPrivate", , , , , True
    
End Function



Private Function EmployeesInWorkGroup( _
    ByVal v_aryWorkGroups As Variant, _
    ByVal v_datStart As Date, _
    ByVal v_datEnd As Date _
    ) As Variant
    
    Dim l_rs As ADODB.Recordset
    Dim l_colEmployees As Collection
    
    On Error GoTo eh
    
    Set l_rs = GetEmployeesAssignedToWorkGroup(v_aryWorkGroups, v_datStart, v_datEnd)
    If Not RecordsExist(l_rs) Then
        EmployeesInWorkGroup = Empty
        Exit Function
    End If
    
    Set l_colEmployees = GetEmployeesInRequest(l_rs)
    EmployeesInWorkGroup = ArrayFromCollection(l_colEmployees)

    Exit Function
    
eh:

TrsRaiseError m_strCMODULENAME, "EmployeesInWorkGroup"

    
    
End Function


Private Function GetEmployeesAssignedToWorkGroup( _
    ByVal v_aryWorkGroups As Variant, _
    ByVal v_datStart As Date, _
    ByVal v_datEnd As Date _
    ) As ADODB.Recordset
    
    Dim l_obj As dbWorkGroup.cdWorkGroup
    
    On Error GoTo eh
    
    Set l_obj = CreateObject("dbWorkGroup.cdWorkGroup")
    Set GetEmployeesAssignedToWorkGroup = _
        l_obj.GetEmployeesAssignedToWorkGroup(v_aryWorkGroups, v_datStart, v_datEnd)
    Set l_obj = Nothing

    Exit Function
    
eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "GetEmployeesAssignedToWorkGroup"

    
End Function

Private Function GetEmployeesInRequest( _
    ByVal v_objRS As ADODB.Recordset _
    ) As Collection

    Dim l_colEmployees As Collection
    Dim l_strEmployee As String
    Dim l_aryEmployees As Variant

    On Error GoTo eh

    Set l_colEmployees = New Collection

    If Not RecordsExist(v_objRS) Then
        Set GetEmployeesInRequest = l_colEmployees
        Exit Function
    End If

    v_objRS.Sort = "EMP_NUM"
    v_objRS.MoveFirst
    l_strEmployee = ""
    While Not v_objRS.EOF
        If Trim(str(v_objRS!EMP_NUM)) <> l_strEmployee Then
            l_strEmployee = Trim(str(v_objRS!EMP_NUM))
            If Val(l_strEmployee) > 0 Then
                If Not IncludesString(l_strEmployee, l_colEmployees) Then
                    l_colEmployees.Add l_strEmployee
                End If
            End If
        End If
        v_objRS.MoveNext
    Wend
    v_objRS.Sort = ""
    v_objRS.MoveFirst

    Set GetEmployeesInRequest = l_colEmployees

    Exit Function

eh:

  TrsRaiseError m_strCMODULENAME, "GetEmployeesInRequest"

End Function
            
            
            

Public Function ViewCalendar( _
    ByVal v_lngEmpNum As Long, _
    ByVal v_datStartDate As Date, _
    ByVal v_datEndDate As Date, _
    ByRef r_datLastClosedPPEndDate As Date, _
    ByRef r_rsDaysOff As ADODB.Recordset, _
    ByRef r_rsExceptions As ADODB.Recordset, _
    ByRef r_rsSchedule As ADODB.Recordset, _
    ByRef r_rsActualTimeInfo As ADODB.Recordset) As Boolean

    On Error GoTo eh

    Dim l_lngEmpNum As Long
    Dim l_datStart As Date
    Dim l_datEnd As Date
    Dim l_rsTime As ADODB.Recordset
    Dim l_datLastPayEnd As Date
    Dim l_datLastDayOfOpenPP As Date
    Dim l_rsLockStatus As ADODB.Recordset
    Dim l_rsSched As ADODB.Recordset
    Dim l_rsDatesWorked As ADODB.Recordset
    Dim l_rsDaysOff As ADODB.Recordset
    Dim l_rsExceptions As ADODB.Recordset
    Dim l_rsAdj As ADODB.Recordset
    Dim l_rsActualTimeSCP As ADODB.Recordset
    
    l_lngEmpNum = v_lngEmpNum
    l_datStart = v_datStartDate
    l_datEnd = v_datEndDate
    Set l_rsLockStatus = GetPayPeriodLockStatus(l_lngEmpNum)
    l_datLastPayEnd = GetLastClosedPayPeriodEndDate(l_rsLockStatus)
    If l_datLastPayEnd < l_datStart Then
        l_datLastPayEnd = l_datStart
    End If
    
    l_rsLockStatus.Filter = adFilterNone
    l_rsLockStatus.Sort = ""
    If RecordsExist(l_rsLockStatus) Then
        l_datLastDayOfOpenPP = GetCurrentPayPeriodEndDate(l_rsLockStatus)
    Else
        l_datLastDayOfOpenPP = l_datEnd
    End If
    If l_datStart <= l_datEnd Then
        Set l_rsSched = GetSched( _
                l_lngEmpNum, l_datStart, l_datEnd, trAdjBaseSchedWithExceptionsAndOvertime)
    Else
        Set l_rsSched = Nothing
    End If
    
    If l_datLastPayEnd > l_datEnd Then
        l_datLastPayEnd = l_datEnd
    End If
    
    If l_datLastDayOfOpenPP > l_datEnd Then
        l_datLastDayOfOpenPP = l_datEnd
    End If
    
    Set l_rsDatesWorked = GetDaysWorked( _
            l_lngEmpNum, l_datStart, l_datLastDayOfOpenPP)
    Set l_rsDaysOff = BuildDaysOffRecordSet( _
        l_rsDatesWorked, l_rsSched)
    Set l_rsExceptions = GetMergedExceptions( _
        l_lngEmpNum, l_datStart, l_datEnd, _
        l_datLastPayEnd, l_datLastDayOfOpenPP)
    Set l_rsAdj = GetAdjustments(l_lngEmpNum, l_datStart, l_datEnd)
    
    ApplyAdjustments l_rsAdj, l_rsExceptions
    
    l_rsSched.Filter = ""

    DetermineFullShiftInd l_rsSched, l_rsExceptions
    DetermineTradeDoubleInd l_rsSched, l_datStart
    DetermineSchedChgPayments l_lngEmpNum, l_datLastPayEnd, l_datEnd, l_rsExceptions
    'PTS 1130
    Set l_rsActualTimeSCP = GetActualTimeSchedChangePay(l_lngEmpNum, l_datStart, l_datEnd)
    
    l_rsDaysOff.Filter = ""
    l_rsExceptions.Filter = ""
    l_rsSched.Filter = ""
    
    Set r_rsActualTimeInfo = l_rsActualTimeSCP
    Set r_rsDaysOff = l_rsDaysOff
    Set r_rsExceptions = l_rsExceptions
    Set r_rsSchedule = l_rsSched
    
    ViewCalendar = True
    Set l_rsDaysOff = Nothing
    Set l_rsExceptions = Nothing
    Set l_rsSched = Nothing
    Set l_rsAdj = Nothing
    Set l_rsLockStatus = Nothing
    Set l_rsDatesWorked = Nothing

    Exit Function
    
eh:
    Set r_rsDaysOff = Nothing
    Set r_rsExceptions = Nothing
    Set r_rsSchedule = Nothing
    Set l_rsDaysOff = Nothing
    Set l_rsExceptions = Nothing
    Set l_rsSched = Nothing
    Set l_rsAdj = Nothing
    Set l_rsLockStatus = Nothing
    Set l_rsDatesWorked = Nothing
    TrsRaiseError m_strCMODULENAME, "ViewCalendar", , , , , True
    
End Function

Public Function ViewCalendarByEmployee( _
    ByVal v_aryEmployees As Variant, _
    ByVal v_datStartDate As Date, _
    ByVal v_datEndDate As Date, _
    ByRef r_rsDaysOff As ADODB.Recordset, _
    ByRef r_rsExceptions As ADODB.Recordset, _
    ByRef r_rsSchedule As ADODB.Recordset, _
    ByRef r_rsActualTimeInfo As ADODB.Recordset) As Boolean

    On Error GoTo eh

    Dim l_lngEmpNum As Long
    Dim l_datStart As Date
    Dim l_datEnd As Date
    Dim l_rsTime As ADODB.Recordset
    Dim l_datLastPayEnd As Date
    Dim l_datLastDayOfOpenPP As Date
    Dim l_rsSched As ADODB.Recordset
    Dim l_rsDatesWorked As ADODB.Recordset
    Dim l_rsDaysOff As ADODB.Recordset
    Dim l_rsExceptions As ADODB.Recordset
    Dim l_rsAdj As ADODB.Recordset
    Dim l_datEndAdjScope As Date
    Dim l_rsActualTimeSCP As ADODB.Recordset
    Dim l_objEnv As busEnvironment.cbsEnvironment
    
    l_datStart = v_datStartDate
    l_datEnd = v_datEndDate
    l_datLastPayEnd = DateAdd("d", -16, Date)
    If l_datLastPayEnd < l_datStart Then
        l_datLastPayEnd = l_datStart
        l_datEndAdjScope = l_datEnd
    Else
        l_datEndAdjScope = l_datLastPayEnd
    End If
        
    Set l_objEnv = CreateObject("busEnvironment.cbsEnvironment")
    l_datLastDayOfOpenPP = DateAdd("d", 16, l_objEnv.GetDefaultEffectiveDate)
    Set l_objEnv = Nothing
    
    If l_datStart <= l_datEnd Then
        Set l_rsSched = GetSched( _
                v_aryEmployees, l_datStart, l_datEnd)
    Else
        Set l_rsSched = Nothing
    End If
    
    Set l_rsDatesWorked = GetDaysWorked( _
            v_aryEmployees, l_datStart, l_datLastDayOfOpenPP)
    Set l_rsDaysOff = BuildDaysOffRecordSet( _
         l_rsDatesWorked, l_rsSched)
    Set l_rsExceptions = GetMergedExceptions( _
        v_aryEmployees, l_datStart, l_datEnd, _
        l_datLastPayEnd, l_datLastDayOfOpenPP)
    Set l_rsAdj = GetAdjustments(v_aryEmployees, l_datStart, l_datEndAdjScope)
    ApplyAdjustments l_rsAdj, l_rsExceptions
    l_rsSched.Filter = ""

    DetermineFullShiftInd l_rsSched, l_rsExceptions
    DetermineTradeDoubleInd l_rsSched, l_datStart
    DetermineSchedChgPayments v_aryEmployees, l_datLastPayEnd, l_datEnd, l_rsExceptions
    
    Set l_rsActualTimeSCP = GetActualTimeSchedChangePay(v_aryEmployees, l_datStart, l_datEnd)
    
    l_rsDaysOff.Filter = ""
    l_rsExceptions.Filter = ""
    l_rsSched.Filter = ""
    
    Set r_rsActualTimeInfo = l_rsActualTimeSCP
    Set r_rsDaysOff = l_rsDaysOff
    Set r_rsExceptions = l_rsExceptions
    Set r_rsSchedule = l_rsSched
    
    ViewCalendarByEmployee = True
    Set l_rsDaysOff = Nothing
    Set l_rsExceptions = Nothing
    Set l_rsSched = Nothing
    Set l_rsAdj = Nothing
    Set l_rsDatesWorked = Nothing
    Set l_objEnv = Nothing

    Exit Function
    
eh:
    Set r_rsDaysOff = Nothing
    Set r_rsExceptions = Nothing
    Set r_rsSchedule = Nothing
    Set l_rsDaysOff = Nothing
    Set l_rsExceptions = Nothing
    Set l_rsSched = Nothing
    Set l_rsAdj = Nothing
    Set l_rsDatesWorked = Nothing
    Set l_objEnv = Nothing
    TrsRaiseError m_strCMODULENAME, "ViewCalendarByEmployee", , , , , True
    
End Function

Public Function ViewCalendarByWorkGroup( _
    ByVal v_aryWorkGroups As Variant, _
    ByVal v_datStartDate As Date, _
    ByVal v_datEndDate As Date, _
    ByRef r_rsDaysOff As ADODB.Recordset, _
    ByRef r_rsExceptions As ADODB.Recordset, _
    ByRef r_rsSchedule As ADODB.Recordset, _
    ByRef r_rsActualTimeInfo As ADODB.Recordset) As Boolean

    On Error GoTo eh

    Dim l_lngEmpNum As Long
    Dim l_datStart As Date
    Dim l_datEnd As Date
    Dim l_rsTime As ADODB.Recordset
    Dim l_datLastPayEnd As Date
    Dim l_datLastDayOfOpenPP As Date
    Dim l_rsSched As ADODB.Recordset
    Dim l_rsDatesWorked As ADODB.Recordset
    Dim l_rsDaysOff As ADODB.Recordset
    Dim l_rsExceptions As ADODB.Recordset
    Dim l_rsAdj As ADODB.Recordset
    Dim l_datEndAdjScope As Date
    Dim l_aryEmployees As Variant
    Dim l_rsActualTimeSCP As ADODB.Recordset
    Dim l_objEnv As busEnvironment.cbsEnvironment
    
    l_datStart = v_datStartDate
    l_datEnd = v_datEndDate
    l_aryEmployees = EmployeesInWorkGroup(v_aryWorkGroups, l_datStart, l_datEnd)
    If IsEmpty(l_aryEmployees) Then
        Set r_rsDaysOff = Nothing
        Set r_rsExceptions = Nothing
        Set r_rsSchedule = Nothing
        ViewCalendarByWorkGroup = False
        GoTo Xit
    End If
    
    l_datLastPayEnd = DateAdd("d", -16, Date)
    If l_datLastPayEnd < l_datStart Then
        l_datLastPayEnd = l_datStart
        l_datEndAdjScope = l_datEnd
    Else
        l_datEndAdjScope = l_datLastPayEnd
    End If
    
    Set l_objEnv = CreateObject("busEnvironment.cbsEnvironment")
    l_datLastDayOfOpenPP = DateAdd("d", 16, l_objEnv.GetDefaultEffectiveDate)
    Set l_objEnv = Nothing
    If l_datStart <= l_datEnd Then
        Set l_rsSched = GetSched(l_aryEmployees, l_datStart, l_datEnd)
    Else
        Set l_rsSched = Nothing
    End If
    
    Set l_rsDatesWorked = GetDaysWorked( _
            l_aryEmployees, l_datStart, l_datLastDayOfOpenPP)
    Set l_rsDaysOff = BuildDaysOffRecordSet( _
         l_rsDatesWorked, l_rsSched)
    Set l_rsExceptions = GetMergedExceptions( _
        l_aryEmployees, l_datStart, l_datEnd, _
        l_datLastPayEnd, l_datLastDayOfOpenPP)
    Set l_rsAdj = GetAdjustments(l_aryEmployees, l_datStart, l_datEndAdjScope)
    ApplyAdjustments l_rsAdj, l_rsExceptions
    l_rsSched.Filter = ""

    DetermineFullShiftInd l_rsSched, l_rsExceptions
    DetermineTradeDoubleInd l_rsSched, l_datStart
    DetermineSchedChgPayments l_aryEmployees, l_datLastPayEnd, l_datEnd, l_rsExceptions
    
    Set l_rsActualTimeSCP = GetActualTimeSchedChangePay(l_aryEmployees, l_datStart, l_datEnd)
    
    l_rsDaysOff.Filter = ""
    l_rsExceptions.Filter = ""
    l_rsSched.Filter = ""
    
    Set r_rsActualTimeInfo = l_rsActualTimeSCP
    Set r_rsDaysOff = l_rsDaysOff
    Set r_rsExceptions = l_rsExceptions
    Set r_rsSchedule = l_rsSched
    
    ViewCalendarByWorkGroup = True
    Set l_rsDaysOff = Nothing
    Set l_rsExceptions = Nothing
    Set l_rsSched = Nothing
    Set l_rsAdj = Nothing
    Set l_rsDatesWorked = Nothing
    Set l_objEnv = Nothing
Xit:
    Exit Function
    
eh:
    Set r_rsDaysOff = Nothing
    Set r_rsExceptions = Nothing
    Set r_rsSchedule = Nothing
    Set l_rsDaysOff = Nothing
    Set l_rsExceptions = Nothing
    Set l_rsSched = Nothing
    Set l_rsAdj = Nothing
    Set l_rsDatesWorked = Nothing
    Set l_objEnv = Nothing
    TrsRaiseError m_strCMODULENAME, "ViewCalendarByWorkGroup", , , , , True
    
End Function

Private Sub MergeActualSchedChgPayments( _
    ByVal r_rsActualTimeSCP As ADODB.Recordset, _
    ByRef r_rsTargetRecordset As ADODB.Recordset)
    
    Dim l_rsPayments As ADODB.Recordset
    Dim l_lngEmpNum As Long
    Dim l_datChargeDate As Date
    Dim l_strFilter As String
           
    On Error GoTo eh
    
    With r_rsActualTimeSCP
        .Sort = "EMP_NUM, CHARGE_DATE"
        .MoveFirst
        While Not .EOF
            l_lngEmpNum = !EMP_NUM
            l_datChargeDate = !CHARGE_DATE
            
            l_strFilter = "EMP_NUM = " & str(l_lngEmpNum) _
                & " AND CHARGE_DATE = " & str(l_datChargeDate)
            l_strFilter = MDAC25FilterFix(l_strFilter)
            r_rsTargetRecordset.Filter = l_strFilter
            
            If RecordsExist(r_rsTargetRecordset) Then
                While Not r_rsTargetRecordset.EOF
                    If r_rsTargetRecordset!SCHED_CHANGE_PMT = "N" Then
                        r_rsTargetRecordset!SCHED_CHANGE_PMT = "Y"
                        r_rsTargetRecordset!FULL_SHIFT_EXC_IND = "N"
                        r_rsTargetRecordset.Update
                    End If
                    r_rsTargetRecordset.MoveNext
                Wend
            Else
                r_rsTargetRecordset.AddNew
                r_rsTargetRecordset!EMP_NUM = l_lngEmpNum
                r_rsTargetRecordset!CHARGE_DATE = l_datChargeDate
'                r_rsTargetRecordset!SCHED_SRC_CD = 30
                r_rsTargetRecordset!EXCEPTION_CD = !SCHED_CHG_EXC_CD
                r_rsTargetRecordset!OVER_EXC_CD = ""
                r_rsTargetRecordset!CASE_GROUP = ""
                r_rsTargetRecordset!BANK_CD = "N"
                r_rsTargetRecordset!FULL_SHIFT_EXC_IND = "N"
                r_rsTargetRecordset!Hours = !Hours
                r_rsTargetRecordset!SHIFT_NUM = 0
                r_rsTargetRecordset!SYMBOL_CD = ""
                r_rsTargetRecordset!SHIFT_CODE = ""
                r_rsTargetRecordset!SCHED_CHANGE_PMT = "Y"
                r_rsTargetRecordset.Update
            End If
            .MoveNext
        Wend
        
    End With
    
    Exit Sub
eh:

    TrsRaiseError m_strCMODULENAME, "MergeActualSchedChgPayments"
    
End Sub

