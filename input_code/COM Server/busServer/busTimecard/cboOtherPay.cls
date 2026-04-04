VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cboOtherPay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const mstrCMODULENAME As String = "cboOtherPay"

Private mobjEmployee As boCommonLib.cboEmployee
Private mobjSchedule As cboSchedule
Private mcolOtherPayEntries As Collection
Private mrsOtherPayEntries As ADODB.Recordset
Private mblnOvertimeRequest As Boolean

Private mobjBSTimecard As cbsTimecard

Public Function AddOtherPayEntry(rsOtherPayRequest As ADODB.Recordset) As cboOtherPayEntry
' Add OtherPayEntry to Collection

    Dim NewOtherPayEntry As cboOtherPayEntry

    
On Error GoTo AddOtherPayEntry_EH

    
    Set NewOtherPayEntry = CreateObject("busTimeCard.cboOtherPayEntry")
    NewOtherPayEntry.SetNew
    
    NewOtherPayEntry.EmployeeNumber = rsOtherPayRequest!EMP_NUM
    NewOtherPayEntry.ChargeDate = rsOtherPayRequest!CHARGE_DATE
    NewOtherPayEntry.DateWorked = rsOtherPayRequest!DATE_WORKED
    NewOtherPayEntry.startTime = rsOtherPayRequest!START_TIME
    NewOtherPayEntry.EndTime = rsOtherPayRequest!END_TIME
    NewOtherPayEntry.ExceptionCode = rsOtherPayRequest!EXCEPTION_CD
'    NewOtherPayEntry.RecordStatus = rsOtherPayRequest!RECORD_STATUS
        
    ' Job Number
    If Not IsNull(rsOtherPayRequest!JOB_NUM) Then
        NewOtherPayEntry.JobNumber = rsOtherPayRequest!JOB_NUM
    End If
    
    ' Work Group Number
    If Not IsNull(rsOtherPayRequest!WG_NUM) Then
        NewOtherPayEntry.WorkgroupNumber = rsOtherPayRequest!WG_NUM
    End If
    
    ' Hours
    If Not IsNull(rsOtherPayRequest!Hours) Then
        NewOtherPayEntry.hrs = rsOtherPayRequest!Hours
    End If
    
    ' Amount
    If Not IsNull(rsOtherPayRequest!AMOUNT) Then
        NewOtherPayEntry.Amt = rsOtherPayRequest!AMOUNT
    End If
    
    ' Start Location
    If Not IsNull(rsOtherPayRequest!START_LOC_CD) Then
        NewOtherPayEntry.StartLocation = rsOtherPayRequest!START_LOC_CD
    End If
    
    ' End Location
    If Not IsNull(rsOtherPayRequest!END_LOC_CD) Then
        NewOtherPayEntry.EndLocation = rsOtherPayRequest!END_LOC_CD
    End If
    
    ' Component Shift Code
    If Not IsNull(rsOtherPayRequest!COMP_SHIFT_CD) Then
        NewOtherPayEntry.ShiftComponentCode = rsOtherPayRequest!COMP_SHIFT_CD
    End If
    
    ' Distribution Id
    If Not IsNull(rsOtherPayRequest!DISTRIBUTION_ID) Then
        NewOtherPayEntry.DistributionId = rsOtherPayRequest!DISTRIBUTION_ID
    End If
    
    ' Distribution Department Number
    If Not IsNull(rsOtherPayRequest!DISTR_DEPT_NUM) Then
        NewOtherPayEntry.DistributionDeptNum = rsOtherPayRequest!DISTR_DEPT_NUM
    End If
    
    ' Distribution Type
    If Not IsNull(rsOtherPayRequest!DISTRIBUTION_TYPE) Then
        NewOtherPayEntry.DistributionType = rsOtherPayRequest!DISTRIBUTION_TYPE
    End If

'***OP2***
    ' Chem Metal Tester
    If Not IsNull(rsOtherPayRequest!CHEM_MTL_TSTR_IND) Then
        NewOtherPayEntry.ChemTesterInd = rsOtherPayRequest!CHEM_MTL_TSTR_IND
    Else
        NewOtherPayEntry.ChemTesterInd = "N"
    End If
    
    ' Half Trip Ind
    If Not IsNull(rsOtherPayRequest!HALF_TRIP_IND) Then
        NewOtherPayEntry.HalfTripInd = rsOtherPayRequest!HALF_TRIP_IND
    Else
        NewOtherPayEntry.HalfTripInd = "N"
    End If
'***OP2***
        
    ' Approval Indicator
    If Not IsNull(rsOtherPayRequest!APPR_IND) Then
        NewOtherPayEntry.ApprovalInd = rsOtherPayRequest!APPR_IND
    Else
        NewOtherPayEntry.ApprovalInd = "N"
    End If

    If NewOtherPayEntry.ApprovalInd = "Y" Then
        NewOtherPayEntry.ApprovalDate = Date
        NewOtherPayEntry.ApproveByEmpNum = UserEmployeeNumber
    Else
        NewOtherPayEntry.ApprovalDate = CDate(CDEFAULTDATE)
        NewOtherPayEntry.ApproveByEmpNum = 0
    End If
    
    ' Overtime Opportunity Number
    If IsOvertime Then
        If Not IsNull(rsOtherPayRequest!OT_OPP_NUM) Then
            NewOtherPayEntry.OvertimeOpportunityNumber = rsOtherPayRequest!OT_OPP_NUM
        End If
    End If
    
    '  Add to the OtherPay Collection
    Me.AddOtherPay NewOtherPayEntry
    
    '  Return OtherPay Entry
    Set AddOtherPayEntry = NewOtherPayEntry

    Exit Function
    
AddOtherPayEntry_EH:
    TrsRaiseError mstrCMODULENAME, "AddOtherPayEntry"
    
End Function

Private Sub UpdateOtherPayEntry(r_objOtherPayEntry As cboOtherPayEntry, _
                                rsOtherPayRequest As ADODB.Recordset)
    
On Error GoTo UpdateOtherPayEntry_EH
    '  Move data in recordset into collection
    
    SetUnchangedFieldsToNull r_objOtherPayEntry, rsOtherPayRequest
    
    '  Exception Code/Earning Code
    If Not IsNull(rsOtherPayRequest!EXCEPTION_CD) Then
        r_objOtherPayEntry.ExceptionCode = rsOtherPayRequest!EXCEPTION_CD
        r_objOtherPayEntry.DetermineEarningsCode
    End If
    
    '  Hours
    If Not IsNull(rsOtherPayRequest!Hours) Then
        r_objOtherPayEntry.hrs = rsOtherPayRequest!Hours
    End If
    
    '  Job Number
    If Not IsNull(rsOtherPayRequest!JOB_NUM) Then
        r_objOtherPayEntry.UpdateJobNumber (rsOtherPayRequest!JOB_NUM)
    End If

    '  Workgroup
    If Not IsNull(rsOtherPayRequest!WG_NUM) Then
        r_objOtherPayEntry.UpdateChargedWorkgroupNumber (rsOtherPayRequest!WG_NUM)
    End If
    
    '  Union Code
    r_objOtherPayEntry.DetermineUnionCode
    
    '  Amount
    If Not IsNull(rsOtherPayRequest!AMOUNT) Then
        r_objOtherPayEntry.Amt = rsOtherPayRequest!AMOUNT
    End If
    
    '  Start Location
    If Not IsNull(rsOtherPayRequest!START_LOC_CD) Then
        r_objOtherPayEntry.StartLocation = rsOtherPayRequest!START_LOC_CD
    End If
    
    '  End Location
    If Not IsNull(rsOtherPayRequest!END_LOC_CD) Then
        r_objOtherPayEntry.EndLocation = rsOtherPayRequest!END_LOC_CD
    End If
    
    '  Component Shift Code Component Shift Amount
    If Not IsNull(rsOtherPayRequest!COMP_SHIFT_CD) Then
        r_objOtherPayEntry.ShiftComponentCode = rsOtherPayRequest!COMP_SHIFT_CD
        r_objOtherPayEntry.DetermineShiftComponentAmount
    End If
    
    '  Distribution Id
    If Not IsNull(rsOtherPayRequest!DISTRIBUTION_ID) Then
        r_objOtherPayEntry.DistributionId = rsOtherPayRequest!DISTRIBUTION_ID
    End If
    
    '  Distribution Type
    If Not IsNull(rsOtherPayRequest!DISTRIBUTION_TYPE) Then
        r_objOtherPayEntry.DistributionType = rsOtherPayRequest!DISTRIBUTION_TYPE
    End If
    
    '  Distribution Department Number
    If Not IsNull(rsOtherPayRequest!DISTR_DEPT_NUM) Then
        r_objOtherPayEntry.UpdateDistDeptNumber (rsOtherPayRequest!DISTR_DEPT_NUM)
    End If

    ' Chem Metal Tester
    If Not IsNull(rsOtherPayRequest!CHEM_MTL_TSTR_IND) Then
        r_objOtherPayEntry.ChemTesterInd = rsOtherPayRequest!CHEM_MTL_TSTR_IND
    End If
    
    ' Half Trip Ind
    If Not IsNull(rsOtherPayRequest!HALF_TRIP_IND) Then
        r_objOtherPayEntry.HalfTripInd = rsOtherPayRequest!HALF_TRIP_IND
    End If

    '  Approval Indicator
    If Not IsNull(rsOtherPayRequest!APPR_IND) Then
        r_objOtherPayEntry.SetApprovalIndicator rsOtherPayRequest!APPR_IND, UserEmployeeNumber
    End If
    
    ' Determine TTA Amount, DistId, Dist Type
    If Trim(r_objOtherPayEntry.ExceptionCode) = "TTA" Then
        r_objOtherPayEntry.DetermineTTAAmount
    End If

    Exit Sub

UpdateOtherPayEntry_EH:
    TrsRaiseError mstrCMODULENAME, "UpdateOtherPayEntry"
    
End Sub


Private Sub SetUnchangedFieldsToNull( _
        r_objOtherPayEntry As cboOtherPayEntry, _
        r_rsOtherPayRequest As ADODB.Recordset)

On Error GoTo eh

' ignore unchanged values


    If Not IsNull(r_rsOtherPayRequest.Fields("JOB_NUM")) Then
        If Trim(r_rsOtherPayRequest.Fields("JOB_NUM")) = _
                r_objOtherPayEntry.JobNumber Then
            r_rsOtherPayRequest.Fields("JOB_NUM") = Null
        End If
    End If

    If Not IsNull(r_rsOtherPayRequest.Fields("WG_NUM")) Then
        If r_rsOtherPayRequest.Fields("WG_NUM") = _
                r_objOtherPayEntry.WorkgroupNumber Then
            r_rsOtherPayRequest.Fields("WG_NUM") = Null
        End If
    End If

    If Not IsNull(r_rsOtherPayRequest.Fields("HOURS")) Then
        If r_rsOtherPayRequest.Fields("HOURS") = _
            r_objOtherPayEntry.hrs Then
            r_rsOtherPayRequest.Fields("HOURS") = Null
        End If
    End If

    If Not IsNull(r_rsOtherPayRequest.Fields("AMOUNT")) Then
        If r_rsOtherPayRequest.Fields("AMOUNT") = _
            r_objOtherPayEntry.Amt Then
            r_rsOtherPayRequest.Fields("AMOUNT") = Null
        End If
    End If

    If Not IsNull(r_rsOtherPayRequest.Fields("START_LOC_CD")) Then
        If Trim(r_rsOtherPayRequest.Fields("START_LOC_CD")) = _
                r_objOtherPayEntry.StartLocation Then
            r_rsOtherPayRequest.Fields("START_LOC_CD") = Null
        End If
    End If

    If Not IsNull(r_rsOtherPayRequest.Fields("END_LOC_CD")) Then
        If Trim(r_rsOtherPayRequest.Fields("END_LOC_CD")) = _
                r_objOtherPayEntry.EndLocation Then
            r_rsOtherPayRequest.Fields("END_LOC_CD") = Null
        End If
    End If

    If Not IsNull(r_rsOtherPayRequest.Fields("COMP_SHIFT_CD")) Then
        If Trim(r_rsOtherPayRequest.Fields("COMP_SHIFT_CD")) = _
                r_objOtherPayEntry.ShiftComponentCode Then
            r_rsOtherPayRequest.Fields("COMP_SHIFT_CD") = Null
        End If
    End If
    
    If Not IsNull(r_rsOtherPayRequest.Fields("DISTRIBUTION_TYPE")) Then
        If Trim(r_rsOtherPayRequest.Fields("DISTRIBUTION_TYPE")) = _
                r_objOtherPayEntry.DistributionType Then
            r_rsOtherPayRequest.Fields("DISTRIBUTION_TYPE") = Null
        End If
    End If
    
    If Not IsNull(r_rsOtherPayRequest.Fields("DISTRIBUTION_ID")) Then
        If Trim(r_rsOtherPayRequest.Fields("DISTRIBUTION_ID")) = _
                r_objOtherPayEntry.DistributionId Then
            r_rsOtherPayRequest.Fields("DISTRIBUTION_ID") = Null
        End If
    End If
    
    If Not IsNull(r_rsOtherPayRequest.Fields("DISTR_DEPT_NUM")) Then
        If Trim(r_rsOtherPayRequest.Fields("DISTR_DEPT_NUM")) = _
                r_objOtherPayEntry.DistributionDeptNum Then
            r_rsOtherPayRequest.Fields("DISTR_DEPT_NUM") = Null
        End If
    End If
    
    If Not IsNull(r_rsOtherPayRequest.Fields("EXCEPTION_CD")) Then
        If Trim(r_rsOtherPayRequest.Fields("EXCEPTION_CD")) = _
                r_objOtherPayEntry.ExceptionCode Then
            r_rsOtherPayRequest.Fields("EXCEPTION_CD") = Null
        End If
    End If
    
    If Not IsNull(r_rsOtherPayRequest.Fields("CHEM_MTL_TSTR_IND")) Then
        If Trim(r_rsOtherPayRequest.Fields("CHEM_MTL_TSTR_IND")) = _
                r_objOtherPayEntry.ChemTesterInd Then
            r_rsOtherPayRequest.Fields("CHEM_MTL_TSTR_IND") = Null
        End If
    End If

    If Not IsNull(r_rsOtherPayRequest.Fields("HALF_TRIP_IND")) Then
        If Trim(r_rsOtherPayRequest.Fields("HALF_TRIP_IND")) = _
                r_objOtherPayEntry.HalfTripInd Then
            r_rsOtherPayRequest.Fields("HALF_TRIP_IND") = Null
        End If
    End If

    If Not IsNull(r_rsOtherPayRequest.Fields("APPR_IND")) Then
        If Trim(r_rsOtherPayRequest.Fields("APPR_IND")) = _
                r_objOtherPayEntry.ApprovalInd Then
            r_rsOtherPayRequest.Fields("APPR_IND") = Null
        End If
    End If

    Exit Sub
    
eh:
    TrsRaiseError mstrCMODULENAME, "SetUnchangedFieldsToNull"

End Sub

'Private Function CaptureKeyFieldValues( _
'        r_rsOtherPayRequest As ADODB.Recordset) As Variant
'
'On Error GoTo eh
'
'    Dim l_varKeyFields(7) As Variant
'
'    If Not IsNull(r_rsOtherPayRequest.Fields("EMP_NUM")) Then
'        l_varKeyFields(0) = r_rsOtherPayRequest.Fields("EMP_NUM")
'    Else
'        l_varKeyFields(0) = 0
'    End If
'
'    If Not IsNull(r_rsOtherPayRequest.Fields("CHARGE_DATE")) Then
'        l_varKeyFields(1) = r_rsOtherPayRequest.Fields("CHARGE_DATE")
'    Else
'        l_varKeyFields(1) = Date
'    End If
'
'    If Not IsNull(r_rsOtherPayRequest.Fields("DATE_WORKED")) Then
'        l_varKeyFields(2) = r_rsOtherPayRequest.Fields("DATE_WORKED")
'    Else
'        l_varKeyFields(2) = Date
'    End If
'
'    If Not IsNull(r_rsOtherPayRequest.Fields("START_TIME")) Then
'        l_varKeyFields(3) = r_rsOtherPayRequest.Fields("START_TIME")
'    Else
'        l_varKeyFields(3) = "99:00"
'    End If
'
'    If Not IsNull(r_rsOtherPayRequest.Fields("END_TIME")) Then
'        l_varKeyFields(4) = r_rsOtherPayRequest.Fields("END_TIME")
'    Else
'        l_varKeyFields(4) = "99:00"
'    End If
'
'    If Not IsNull(r_rsOtherPayRequest.Fields("HOURS")) Then
'        l_varKeyFields(5) = r_rsOtherPayRequest.Fields("HOURS")
'    Else
'        l_varKeyFields(5) = CDec("0.0")
'    End If
'
'    If Not IsNull(r_rsOtherPayRequest.Fields("EXCEPTION_CD")) Then
'        l_varKeyFields(6) = r_rsOtherPayRequest.Fields("EXCEPTION_CD")
'    Else
'        l_varKeyFields(6) = ""
'    End If
'
'    l_varKeyFields(7) = ""
'
'    CaptureKeyFieldValues = l_varKeyFields
'
'    Exit Function
'
'eh:
'    TrsRaiseError mstrCMODULENAME, "CaptureKeyFieldValues", Err.Number
'
'End Function


Private Sub DeleteOtherPayEntry(r_objOtherPayEntry As cboOtherPayEntry, _
                            rsOtherPayRequest As ADODB.Recordset)
    ' Mark for deletion
    r_objOtherPayEntry.Delete

End Sub

Public Function AddOtherPay(objOtherPayEntry As cboOtherPayEntry)
    mcolOtherPayEntries.Add objOtherPayEntry
    Set objOtherPayEntry.OtherPay = SafeRef(Me)
End Function

Public Function UserEmployeeNumber() As Long
    UserEmployeeNumber = GetBSTimecard.UserEmployeeNumber
End Function

Public Function SetEmployee(emp As cboEmployee)
    Set mobjEmployee = emp
End Function

Public Function Employee() As cboEmployee
    Set Employee = mobjEmployee
End Function

Public Function Schedule() As cboSchedule
    Set Schedule = mobjSchedule
End Function

Public Function SetSchedule(sched As cboSchedule)
    Set mobjSchedule = sched
End Function

Public Function SetOvertime(ByVal blnOvertime As Boolean)
    mblnOvertimeRequest = blnOvertime
End Function

Public Function IsOvertime() As Boolean
    IsOvertime = mblnOvertimeRequest
End Function


Private Function ApplyRules(OtherPayEntryRequest As cboOtherPayEntry) As Boolean
    
''''    Dim OtherPayEntry As cboOtherPayEntry
    
On Error GoTo ApplyRules_EH
    
    '  Default Missing
    BuildMissingComponents OtherPayEntryRequest
    Exit Function
    
ApplyRules_EH:
    TrsRaiseError mstrCMODULENAME, "ApplyRules"

End Function

Private Sub BuildMissingComponents(OtherPayEntry As cboOtherPayEntry)

Dim mobjEmployee As boCommonLib.cboEmployee
Dim mEmployee As cboEmployee
    
On Error GoTo BuildMissingComponents_EH

    ' End Time  *  Client should be passing
    If OtherPayEntry.EndTime = "" Then
        OtherPayEntry.EndTime = OtherPayEntry.startTime
    End If
    
    ' Job Number (Per Prem and Greg - Don't worry about Job Upgrade, it will be sent) *
    If Trim(OtherPayEntry.JobNumber) = "" Then
        OtherPayEntry.DetermineJobNumber
    End If
    
    ' WorkGroup *  /Location (Depends on WorkGroup)  *
    If OtherPayEntry.WorkgroupNumber = 0 Then
        OtherPayEntry.DetermineChargedWorkgroupNumber
    Else
        OtherPayEntry.DetermineLocationCode
    End If
    
    ' Department Number  *
    If Trim(OtherPayEntry.DepartmentNumber) = "" Then
        OtherPayEntry.DeterminePayrollDepartmentNumber
    End If

    ' Union Code  *
    OtherPayEntry.DetermineUnionCode
    
    ' Distribution Department Number (Depends on WorkGroup Num)  *
'JIMB Begin
'  Always determine Distribution Department number
    OtherPayEntry.DistributionDeptNum = ""
'JimB end
    If Trim(OtherPayEntry.DistributionDeptNum) = "" Then
        OtherPayEntry.DetermineDistributionDepartmentNumber
    End If
    
    If Trim(OtherPayEntry.ShiftComponentCode) = "" Then
        OtherPayEntry.DetermineShiftComponentCode
    End If
    ' Shift Component Amount  *
    OtherPayEntry.DetermineShiftComponentAmount
    
    ' Determine Earnings Code   (Depends on Exception Code)  *
    OtherPayEntry.DetermineEarningsCode

'***OP2***
    ' Determine TTA Amount, DistId, Dist Type
    If Trim(OtherPayEntry.ExceptionCode) = "TTA" Then
        OtherPayEntry.DetermineTTAAmount
    End If
'***OP2***
    
    Exit Sub

BuildMissingComponents_EH:
    TrsRaiseError mstrCMODULENAME, "BuildMissingComponents"

End Sub

Public Function RecordOtherPay(ByRef v_rsOtherPayRequest As ADODB.Recordset, _
                                ByRef r_rsErrors As ADODB.Recordset _
                                ) As Boolean

    Dim l_objOtherPayEntry As cboOtherPayEntry
    Dim lcolErrDescArgs As Collection
    Dim lcolErrArgs As Collection
    Dim l_intPos As Integer
    
On Error GoTo RecordOtherPay_EH

'    Set mobjErrors = r_rsErrors
    '  This is already created in restore
'    Set mcolOtherPayEntries = New Collection

'  Determine RecordStatus
    '  BuildMissingComp
    Select Case v_rsOtherPayRequest!RECORD_STATUS
    Case "I"
        ' Insert Other Pay Entry
        Set l_objOtherPayEntry = AddOtherPayEntry(v_rsOtherPayRequest)
        ApplyRules l_objOtherPayEntry
    Case "C"
        ' Does the entry exist
        Set l_objOtherPayEntry = FindOtherPayEntry(v_rsOtherPayRequest)
        ' Mark for Change
        If Not l_objOtherPayEntry Is Nothing Then
            UpdateOtherPayEntry l_objOtherPayEntry, v_rsOtherPayRequest
        Else
            Set lcolErrArgs = New Collection
            lcolErrArgs.Add "v_rsOtherPayRequest.Filter=" & v_rsOtherPayRequest.Filter
            'Getting all Requests up to the current Request
            l_intPos = v_rsOtherPayRequest.AbsolutePosition
            lcolErrArgs.Add "Current Request Row = " & str(l_intPos)
            v_rsOtherPayRequest.MoveFirst
            lcolErrArgs.Add v_rsOtherPayRequest.GetString()
            v_rsOtherPayRequest.AbsolutePosition = l_intPos
            Err.Raise trErrbusInvalidOPChangeEntryNotFound
        End If
    Case "D"
        ' Does the entry exist
        Set l_objOtherPayEntry = FindOtherPayEntry(v_rsOtherPayRequest)
        ' Mark for Deletion
        If Not l_objOtherPayEntry Is Nothing Then
            DeleteOtherPayEntry l_objOtherPayEntry, v_rsOtherPayRequest
        Else
            Set lcolErrArgs = New Collection
            lcolErrArgs.Add "v_rsOtherPayRequest.Filter=" & v_rsOtherPayRequest.Filter
            'Getting all Requests up to the current Request
            l_intPos = v_rsOtherPayRequest.AbsolutePosition
            lcolErrArgs.Add "Current Request Row = " & str(l_intPos)
            v_rsOtherPayRequest.MoveFirst
            lcolErrArgs.Add v_rsOtherPayRequest.GetString()
            v_rsOtherPayRequest.AbsolutePosition = l_intPos
            Err.Raise trErrbusInvalidOPDeleteEntryNotFound
        End If
        
    Case Else
        Set lcolErrDescArgs = New Collection
        lcolErrDescArgs.Add CStr(v_rsOtherPayRequest!RECORD_STATUS)
        Err.Raise trErrbusInvalidRecordStatus
    End Select
    Set l_objOtherPayEntry = Nothing
    RecordOtherPay = True
    Exit Function

RecordOtherPay_EH:
    Set l_objOtherPayEntry = Nothing
    Select Case Err.Number
        Case trErrbusInvalidRecordStatus
            TrsRaiseError mstrCMODULENAME, "RecordOtherPay", Err.Number, , lcolErrDescArgs
        Case trErrbusInvalidOPChangeEntryNotFound, trErrbusInvalidOPDeleteEntryNotFound
            TrsRaiseError mstrCMODULENAME, "RecordOtherPay", Err.Number, , lcolErrArgs
        Case Else
            TrsRaiseError mstrCMODULENAME, "RecordOtherPay"
    End Select

End Function

Private Function FindOtherPayEntry(OtherPayEntry As ADODB.Recordset) As cboOtherPayEntry
' OtherPay Entry Unique Identifiers
' Employee Number, Charge Date, StartTime, IndexTime
' Answer nothing if a matching timecard entry is not found.
    Dim l_objOtherPayEntry As cboOtherPayEntry
    Dim l_lngEmpNum As Long
    Dim l_datChargeDate As Date
    Dim l_strStartTime As String
    Dim l_strIndexTime As String

On Error GoTo FindOtherPayEntry_EH

    l_lngEmpNum = OtherPayEntry!KEY_EMP_NUM
    l_datChargeDate = OtherPayEntry!KEY_CHARGE_DATE
    l_strStartTime = Trim(OtherPayEntry!KEY_START_TIME)
    l_strIndexTime = Trim(OtherPayEntry!KEY_INDEX_TIME)

    For Each l_objOtherPayEntry In OtherPayEntries
        If l_objOtherPayEntry.EmployeeNumber = l_lngEmpNum And _
            l_objOtherPayEntry.ChargeDate = l_datChargeDate And _
            l_objOtherPayEntry.startTime = l_strStartTime And _
            l_objOtherPayEntry.IndexTime = l_strIndexTime Then
            Set FindOtherPayEntry = l_objOtherPayEntry
            Exit Function
        End If
    Next
    Set l_objOtherPayEntry = Nothing
    Exit Function
    
FindOtherPayEntry_EH:
    Set l_objOtherPayEntry = Nothing
    TrsRaiseError mstrCMODULENAME, "FindOtherPayEntry"

End Function

'Public Function ErrorSet() As ADODB.Recordset
'    Set ErrorSet = mobjErrors
'End Function

'Error set is for validation and warnings only.
Public Function ErrorSet() As ADODB.Recordset
    Set ErrorSet = GetBSTimecard.ErrorSet
End Function

Public Function GetBSTimecard() As cbsTimecard
    Set GetBSTimecard = mobjBSTimecard
End Function

Public Function GetJobManager() As dbCodeTables.cdCodeTables
    Set GetJobManager = Me.GetBSTimecard.GetJobManager
End Function

Public Function GetDepartmentManager() As dbDepartment.cdDepartment
    Set GetDepartmentManager = Me.GetBSTimecard.GetDepartmentManager
End Function

Public Property Get OtherPayEntries() As Collection
' filter deleted entries from collection
    Dim colCopy As Collection
    Dim ope As cboOtherPayEntry
    
On Error GoTo eh
    Set colCopy = New Collection
    For Each ope In OtherPayEntriesPrivate
        If Not ope.IsDeleted Then
            colCopy.Add ope
        End If
    Next
    Set OtherPayEntries = colCopy
    Set colCopy = Nothing
    Set ope = Nothing
    Exit Property
    
eh:
    Set colCopy = Nothing
    Set ope = Nothing
    TrsRaiseError mstrCMODULENAME, "OtherPayEntries"
        
End Property

Public Function OtherPayEntriesPrivate() As Collection
' reference to private collection, only allow friends access.
    Set OtherPayEntriesPrivate = mcolOtherPayEntries
End Function


Public Sub SetBSTimecard(r_objBSTimecard As cbsTimecard)
    Set mobjBSTimecard = r_objBSTimecard
End Sub

Public Sub Restore(ByVal v_lngEmpNum As Long, _
        ByVal v_datStartDate As Variant, v_datEndDate As Variant)
    ' Get existing OtherPay for this employee for this time period
    

    Dim dbTimeCard As dbTimeCard.cdTimeCard
    Dim rsOtherPay As ADODB.Recordset
    
On Error GoTo Restore_eh
    
    
    Set dbTimeCard = CreateObject("dbTimeCard.cdTimeCard")
    Set rsOtherPay = dbTimeCard.Get_OtherPay(v_lngEmpNum, v_datStartDate, v_datEndDate)
    
    ' Set the state of other pay using datamanager
    SetState rsOtherPay
    Exit Sub
    
Restore_eh:
    TrsRaiseError mstrCMODULENAME, "RecordOtherPay"

End Sub


Public Sub RestoreWithState(ByRef r_rsOtherpayEntries As ADODB.Recordset)

    On Error GoTo Restore_eh
    
    SetState r_rsOtherpayEntries
    Exit Sub
    
Restore_eh:
    TrsRaiseError mstrCMODULENAME, "RestoreWithState"
End Sub


Public Function ResetRecordStatus()
'Ensure all record status has been reset to restored.

    If RecordsExist(mrsOtherPayEntries) Then
        mrsOtherPayEntries.MoveFirst
        Do Until mrsOtherPayEntries.EOF
            mrsOtherPayEntries!RECORD_STATUS = " "
            mrsOtherPayEntries.MoveNext
        Loop
        mrsOtherPayEntries.MoveFirst
    End If

End Function

Public Function OtherPayEntryRecordset() As ADODB.Recordset
    Set OtherPayEntryRecordset = mrsOtherPayEntries
End Function

Public Function SetState(rsOtherPay As ADODB.Recordset)
    Dim restoredEntry As cboOtherPayEntry

    
On Error GoTo eh

    If rsOtherPay Is Nothing Then
        Exit Function
    End If
    
    
    Set mrsOtherPayEntries = rsOtherPay
    
' create collection of cboOtherPayEntries
    Set mcolOtherPayEntries = New Collection
    
' REMOVE DATAMANAGER
    If RecordsExist(rsOtherPay) Then
        rsOtherPay.MoveFirst
        Do Until rsOtherPay.EOF
            Set restoredEntry = CreateObject("busTimecard.cboOtherPayEntry")
            restoredEntry.Restore rsOtherPay
            AddOtherPay restoredEntry
            rsOtherPay.MoveNext
        Loop
    End If
    
    
    Set restoredEntry = Nothing
    Exit Function
    
eh:
    
    Set restoredEntry = Nothing
    TrsRaiseError mstrCMODULENAME, "SetState"
End Function

Public Sub Release()
    Dim ope As cboOtherPayEntry

'Donot enable an error handler since this is called from an error handler.

    If Not mobjEmployee Is Nothing Then
        Set mobjEmployee = Nothing
    End If
    
    If Not mobjSchedule Is Nothing Then
        Set mobjSchedule = Nothing
    End If
    
    If Not mobjBSTimecard Is Nothing Then
        Set mobjBSTimecard = Nothing
    End If
    
    For Each ope In mcolOtherPayEntries
        ope.Release
    Next
    Set ope = Nothing
    Set mcolOtherPayEntries = Nothing
    
    Set mrsOtherPayEntries = Nothing
    
    GetObjectContext.SetComplete
    
End Sub
