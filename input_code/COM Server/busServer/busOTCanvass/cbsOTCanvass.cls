VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "cbsOTCanvass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'NOTES: PREM CREATED NEW DB METHODS FOR CERTAIN RETRIEVALS
'IN OCTOBER 2005
'ConstructCanvassSheet2, GetOvertimeCanvasses2
', GetOTOpportunity2, GetOTOppCanvassSheet2
', GetOTOppPhoneCalls2, GetOpportunityResponses2
', GetNonCanavassedOvertimeCanvasses2, GetCanvasses2
', GetOTOppCharges2, GetOTOppIneligibilities2
'THESE METHODS USE PARAMETERIZED QUERIES
Private m_rsTemplateData As ADODB.Recordset
Private m_strSortColumns(1 To 24) As String
Private m_strGroupSortColumns(1 To 4) As String
Private m_rsAllExceptions As ADODB.Recordset
Private m_blnReqEmpStatus As Boolean
Private m_rsEmpLocs As ADODB.Recordset
Private m_rsCanvLocs As ADODB.Recordset
 
Private Const m_strCMODULENAME As String = "cbsOTCanvass"
Public Function GetMaintainIneligibilities( _
    ByVal v_aryEmployees As Variant, _
    Optional v_datStartDate As Date = CDate("1980-01-01"), _
    Optional v_datEndDate As Date = CDate("9999-12-31") _
    ) As ADODB.Recordset
    
    Dim l_obj As dbOTCanvass.cdOTCanvass
    
    
    On Error GoTo eh
    
    Set l_obj = CreateObject("dbOTCanvass.cdOTCanvass")
    
    Set GetMaintainIneligibilities = l_obj.GetMaintainIneligibilities(v_aryEmployees, v_datStartDate, v_datEndDate)
    
    Set l_obj = Nothing
    
    Exit Function

eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "GetMaintainIneligibilities", , , , , True
    
End Function
Public Function GetOvertimeCanvasses( _
    ByVal v_intLocationCode As Integer, _
    ByVal v_datStartDate As Date, _
    ByVal v_datEndDate As Date, _
    Optional v_intOTGroup As Integer = 0, _
    Optional v_blnOpenCanvassesOnly As Boolean = False, _
    Optional v_blnPostedCanvassesOnly As Boolean = False _
    ) As ADODB.Recordset

    Dim l_obj As dbOTCanvass.cdOTCanvass

    On Error GoTo eh

    Set l_obj = CreateObject("dbOTCanvass.cdOTCanvass")
'BEGIN PREM TRIS 5.6
    Set GetOvertimeCanvasses = l_obj.GetOvertimeCanvasses2( _
        v_intLocationCode, v_datStartDate, v_datEndDate, v_intOTGroup, _
        v_blnOpenCanvassesOnly, v_blnPostedCanvassesOnly)
'END PREM TRIS 5.6
    Set l_obj = Nothing

    Exit Function
eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "GetOvertimeCanvasses", , , , , True

End Function
Public Function GetOvertimeCanvassData( _
                    ByVal v_lngOTOppNum As Long, _
                    ByRef r_rsOTOpportunity As ADODB.Recordset, _
                    ByRef r_rsCanvassSheet As ADODB.Recordset, _
                    ByRef r_rsOTOppTimeCharges As ADODB.Recordset, _
                    ByRef r_rsOTOppInelig As ADODB.Recordset, _
                    ByRef r_rsOTOppPhoneCalls As ADODB.Recordset, _
                    ByRef r_rsValidOppResponses As ADODB.Recordset, _
                    ByVal v_blnReqEmpStatus As Boolean, _
                    ByRef r_rsEmpStatus As ADODB.Recordset, _
                    ByRef r_rsOTOppBankCharges As ADODB.Recordset, _
                    ByRef r_rsOtherOTBankCharges As ADODB.Recordset _
                ) As Boolean
'
'
''RSG TEST only until compat changed
'Public Function GetOvertimeCanvassData( _
'    ByVal v_lngOTOppNum As Long, _
'    ByRef r_rsOTOpportunity As ADODB.Recordset, _
'    ByRef r_rsCanvassSheet As ADODB.Recordset, _
'    ByRef r_rsOTOppTimeCharges As ADODB.Recordset, _
'    ByRef r_rsOTOppInelig As ADODB.Recordset, _
'    ByRef r_rsOTOppPhoneCalls As ADODB.Recordset, _
'    ByRef r_rsValidOppResponses As ADODB.Recordset, _
'    ByVal v_blnReqEmpStatus As Boolean, _
'    ByRef r_rsEmpStatus As ADODB.Recordset _
'    ) As Boolean

'Please Note
'r_rsOtherOTBankCharges is future use
'if used, please delete this comment
    
''RSG TEST only until compat changed
'    Dim r_rsOTOppBankCharges As ADODB.Recordset
'    Dim r_rsOtherOTBankCharges As ADODB.Recordset
    
    On Error GoTo eh
    
    m_blnReqEmpStatus = v_blnReqEmpStatus
    GetOvertimeCanvassData = GetOvertimeCanvassDataPrivate( _
        v_lngOTOppNum, r_rsOTOpportunity, r_rsCanvassSheet, _
        r_rsOTOppTimeCharges, r_rsOTOppInelig, r_rsOTOppPhoneCalls, _
        r_rsValidOppResponses, r_rsEmpStatus, r_rsOTOppBankCharges, r_rsOtherOTBankCharges)

    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "GetOvertimeCanvassData", , , , , True

End Function


Private Function GetOvertimeCanvassDataPrivate( _
                    ByVal v_lngOTOppNum As Long, _
                    ByRef r_rsOTOpportunity As ADODB.Recordset, _
                    ByRef r_rsCanvassSheet As ADODB.Recordset, _
                    ByRef r_rsOTOppCharges As ADODB.Recordset, _
                    ByRef r_rsOTOppInelig As ADODB.Recordset, _
                    ByRef r_rsOTOppPhoneCalls As ADODB.Recordset, _
                    ByRef r_rsValidOppResponses As ADODB.Recordset, _
                    ByRef r_rsEmpStatus As ADODB.Recordset, _
                    ByRef r_rsOTOppBankCharges As ADODB.Recordset, _
                    ByRef r_rsOtherOTBankCharges As ADODB.Recordset _
                ) As Boolean
    
    Dim l_obj As dbOTCanvass.cdOTCanvass
    Dim l_rsCanvassSheet As ADODB.Recordset
    Dim l_rsResponses As ADODB.Recordset

On Error GoTo eh

    Set l_obj = CreateObject("dbOTCanvass.cdOTCanvass")

'BEGIN PREM TRIS 5.6
    Set r_rsOTOpportunity = l_obj.GetOTOpportunity2(v_lngOTOppNum)
    Set l_rsCanvassSheet = l_obj.GetOTOppCanvassSheet2(v_lngOTOppNum)
    Set r_rsOTOppCharges = l_obj.GetOTOppCharges2(v_lngOTOppNum)
    Set r_rsOTOppInelig = l_obj.GetOTOppIneligibilities2(v_lngOTOppNum)
    Set r_rsOTOppPhoneCalls = l_obj.GetOTOppPhoneCalls2(v_lngOTOppNum)
    Set l_rsResponses = l_obj.GetOpportunityResponses2(v_lngOTOppNum)
'END PREM TRIS 5.6

    If m_blnReqEmpStatus Then
        Set m_rsAllExceptions = l_obj.GetAllExceptions
    End If
    
'BEGIN PREM TRIS CONVERSION PHASE 3
    Set r_rsOTOppBankCharges = l_obj.GetOpportunityBankCharges(v_lngOTOppNum)
    Set r_rsOtherOTBankCharges = Nothing
'END PREM TRIS CONVERSION PHASE 3

    Set l_obj = Nothing

    If RecordsExist(l_rsResponses) Then
        Set r_rsValidOppResponses = l_rsResponses
    Else
        Set r_rsValidOppResponses = GetOvertimeResponses()
    End If
    SetEmployeeData v_lngOTOppNum, l_rsCanvassSheet, r_rsEmpStatus

    Set r_rsCanvassSheet = l_rsCanvassSheet

    GetOvertimeCanvassDataPrivate = True
    Exit Function
eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "GetOvertimeCanvassDataPrivate"

End Function
Private Function GetOvertimeResponses() As ADODB.Recordset

    Dim l_obj As dbOTCanvass.cdOTCanvass
    
    On Error GoTo eh

    Set l_obj = CreateObject("dbOTCanvass.cdOTCanvass")
    Set GetOvertimeResponses = l_obj.GetOvertimeResponses()
    Exit Function

eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "GetOvertimeResponses"
End Function


Private Sub SetEmployeeData( _
                ByVal v_lngOTOppNum As Long, _
                ByRef r_rsCanvassSheet As ADODB.Recordset, _
                ByRef r_rsEmpStatus As ADODB.Recordset)

    Dim l_aryEmps As Variant
    Dim l_rsQuals As ADODB.Recordset
    Dim l_rsSkills As ADODB.Recordset
    Dim l_obj As dbEmployee.cdEmployee
    Dim l_lngEmployee As Long
    Dim l_strQuals As String
    Dim l_strSkills As String
    Dim l_strPhone As String
    Dim l_rsEmpStatus As ADODB.Recordset

On Error GoTo eh

    l_aryEmps = ExtractEmployees(r_rsCanvassSheet)
    If IsEmpty(l_aryEmps) Then
        Exit Sub
    Else
        Set l_obj = CreateObject("dbEmployee.cdEmployee")
        Set l_rsQuals = l_obj.GetEmployeeQualifications(l_aryEmps)
        Set l_rsSkills = l_obj.GetEmployeeSkills(l_aryEmps)
'        Set l_rsPhoneNumbers = l_obj.GetPeopleSoftPhoneNumber(l_aryEmps)
        If m_blnReqEmpStatus Then
            Set l_rsEmpStatus = GetEmpStatusByOpp(v_lngOTOppNum, Now)
        End If
        Set l_obj = Nothing
    End If
    
    r_rsCanvassSheet.MoveFirst
    While Not r_rsCanvassSheet.EOF
        l_lngEmployee = r_rsCanvassSheet!EMP_NUM
        l_strQuals = ExtractQualifications(l_lngEmployee, l_rsQuals)
        l_strSkills = ExtractSkills(l_lngEmployee, l_rsSkills)
'        l_strPhone = ExtractPhone(l_lngEmployee, l_rsPhoneNumbers)
        r_rsCanvassSheet!QUAL_CDS = l_strQuals
        If l_strQuals <> "" Or l_strSkills <> "" Then
            r_rsCanvassSheet!QUAL_SKILL_IND = "Y"
        End If
        r_rsCanvassSheet!SPEC_SKILL_CDS = l_strSkills
        If l_strPhone <> "" Then
            r_rsCanvassSheet!HOME_PHONE_NUM = l_strPhone
        End If
        If Not l_rsEmpStatus Is Nothing Then
            SetOnDuty l_lngEmployee, l_rsEmpStatus, r_rsCanvassSheet, Now
        End If
        r_rsCanvassSheet.MoveNext
    Wend

    Set r_rsEmpStatus = l_rsEmpStatus
    Exit Sub
    
eh:
    Set l_obj = Nothing
    Set l_rsEmpStatus = Nothing
    TrsRaiseError m_strCMODULENAME, "SetEmployeeData"
    
End Sub


Private Function GetEmployeeSchedules( _
                    ByVal v_aryEmployees As Variant, _
                    ByVal v_datStart As Date, _
                    ByVal v_datEnd As Date) _
                As ADODB.Recordset

    Dim l_obj As busSchedule.cbsSchedule

On Error GoTo eh
    Set l_obj = CreateObject("busSchedule.cbsSchedule")
'    Set GetEmployeeSchedules = l_obj.GetAdjustedBaseSchedule( _
'        v_aryEmployees, v_datStart, v_datEnd, Nothing, , trAdjBaseSchedWithExceptionsAndOvertime)

    Set GetEmployeeSchedules = l_obj.GetScheduleForOTCanvassing( _
        v_aryEmployees, v_datStart, v_datEnd)


    Set l_obj = Nothing
    Exit Function

eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "GetEmployeeSchedules"

End Function


Private Sub SetOnDuty( _
                ByVal v_lngEmpNum As Long, _
                ByVal v_rsEmpStatus As ADODB.Recordset, _
                ByRef r_rsCanvassSheet As ADODB.Recordset, _
                ByVal v_datDateTime As Date)

    Dim l_datStartDateTime As Date
    Dim l_datEndDateTime As Date
                    
On Error GoTo eh
    
    r_rsCanvassSheet!ON_DUTY_IND = "N"
    
    v_rsEmpStatus.Filter = "EMP_NUM = " & v_lngEmpNum
    Do Until v_rsEmpStatus.EOF
        l_datStartDateTime = DateTimeFrom(v_rsEmpStatus!START_DATE, v_rsEmpStatus!START_TIME)
        l_datEndDateTime = DateTimeFrom(v_rsEmpStatus!END_DATE, v_rsEmpStatus!END_TIME)
        If v_datDateTime >= l_datStartDateTime And v_datDateTime <= l_datEndDateTime Then
            If IsActualWork(v_rsEmpStatus!EXCEPTION_CD) Then
                r_rsCanvassSheet!ON_DUTY_IND = "Y"
            Else
                r_rsCanvassSheet!ON_DUTY_IND = "N"
            End If
            r_rsCanvassSheet!ON_DUTY_EXC_CD = v_rsEmpStatus!EXCEPTION_CD
            Exit Do
        End If
        v_rsEmpStatus.MoveNext
    Loop

    v_rsEmpStatus.Filter = adFilterNone
    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "SetOnDuty"
    
End Sub


'Private Function IsOnDuty( _
'                    ByVal v_lngEmpNum As Long, _
'                    ByVal v_rsEmpSchedules As ADODB.Recordset, _
'                    ByVal v_datDateTime As Date) As String
'' Answer Y or N
'
'    Dim l_datStartDateTime As Date
'    Dim l_datEndDateTime As Date
'
'On Error GoTo eh
'
'    IsOnDuty = "N"
'    v_rsEmpSchedules.Filter = "EMP_NUM = " & v_lngEmpNum
'
'    Do Until v_rsEmpSchedules.EOF
'        If IsScheduledWorktime(v_rsEmpSchedules) Then
'            l_datStartDateTime = DateTimeFrom(v_rsEmpSchedules!DATE_WORKED, v_rsEmpSchedules!START_TIME)
'            l_datEndDateTime = DateTimeFrom(v_rsEmpSchedules!DATE_WORKED, v_rsEmpSchedules!END_TIME)
'            If v_datDateTime >= l_datStartDateTime And v_datDateTime <= l_datEndDateTime Then
'                If IsActualWork(v_rsEmpSchedules!EXCEPTION_CD) Then
'                    IsOnDuty = "Y"
'                Else
'                    IsOnDuty = "N"
'                End If
'                Exit Do
'            End If
'        End If
'        v_rsEmpSchedules.MoveNext
'    Loop
'
'    Exit Function
'
'eh:
'    TrsRaiseError m_strCMODULENAME, "IsOnDuty"
'
'End Function

Private Function IsScheduledOvertime(ByVal v_rsEmpSched As ADODB.Recordset) As Boolean
' EXC_CD = HR is also SHIFT_HOURS = 0 and START_TIME <> "" e.g. 08:00 16:00.
' Can this simply look at OT_OPP_NUM?

    IsScheduledOvertime = _
        (v_rsEmpSched!SHIFT_HOURS = 0 And Trim(v_rsEmpSched!START_TIME) <> "") And _
            v_rsEmpSched!OT_OPP_NUM > 0

End Function

Private Function IsRegularScheduledShift(ByVal v_rsEmpSched As ADODB.Recordset) As Boolean

    IsRegularScheduledShift = v_rsEmpSched!SHIFT_HOURS > 0

End Function

Private Function IsScheduledWorktime(ByVal v_rsEmpSched As ADODB.Recordset) As Boolean
' EXC_CD = HR is also SHIFT_HOURS = 0 and START_TIME <> "" e.g. 08:00 16:00.
' RSG - Can this simply look at OT_OPP_NUM?

'    IsScheduledWorktime = (Trim(v_rsEmpSched!START_TIME) <> "" And _
'                                Trim(v_rsEmpSched!SYSTEM_EXC_CD) <> "SLT")

    IsScheduledWorktime = False
    If (Trim(v_rsEmpSched!START_TIME) <> "" And _
            Trim(v_rsEmpSched!SYSTEM_EXC_CD) <> "SLT") Then
        If IsActualWork(v_rsEmpSched!EXCEPTION_CD) Then
            IsScheduledWorktime = True
        End If
    End If

End Function


Private Function IsActualWork(ByVal v_strExc_Cd As String) As Boolean

On Error GoTo eh

    If Trim(v_strExc_Cd) = "" Then
        IsActualWork = True
    Else
        IsActualWork = GetWorkRestInd(v_strExc_Cd) = "W"
    End If
    
    Exit Function

eh:
        TrsRaiseError m_strCMODULENAME, "IsActualWork"

End Function


Private Function GetWorkRestInd(ByVal v_strExc_Cd As String, _
                                Optional v_strGroupInd As String) As String


On Error GoTo eh

    If m_rsAllExceptions Is Nothing Then
        GetAllExceptions
    End If
    
    If v_strGroupInd = "" Then
        m_rsAllExceptions.MoveFirst
    Else
        m_rsAllExceptions.Filter = "GROUP_IND = '" & Trim(v_strGroupInd) & "'"
    End If
    
    m_rsAllExceptions.Find "EXCEPTION_CD = '" & v_strExc_Cd & "'"
    If m_rsAllExceptions.EOF Then
        GetWorkRestInd = ""
    Else
        GetWorkRestInd = Trim(m_rsAllExceptions!WORK_REST_IND)
    End If

    m_rsAllExceptions.Filter = adFilterNone
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "GetWorkRestInd"

End Function


Private Function GetAllExceptions() As ADODB.Recordset

    Dim l_obj As dbOTCanvass.cdOTCanvass

On Error GoTo eh
    
    If m_rsAllExceptions Is Nothing Then
        Set l_obj = CreateObject("dbOTCanvass.cdOTCanvass")
        Set m_rsAllExceptions = l_obj.GetAllExceptions
        Set l_obj = Nothing
    End If

    Set GetAllExceptions = m_rsAllExceptions
     
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "GetAllExceptions"
    
End Function


Private Function ExtractEmployees( _
    ByVal v_rs As ADODB.Recordset _
    ) As Variant

Dim l_colEmps As Collection
Dim l_strEmp As String

On Error GoTo eh

    v_rs.Filter = ""
    If Not RecordsExist(v_rs) Then
        ExtractEmployees = Empty
        Exit Function
    End If
    
    Set l_colEmps = New Collection
    v_rs.MoveFirst
    While Not v_rs.EOF
        l_strEmp = Trim(str(v_rs!EMP_NUM))
        l_colEmps.Add l_strEmp
        v_rs.MoveNext
    Wend
    
    ExtractEmployees = ArrayFromCollection(l_colEmps)
    

    Exit Function

eh:

    TrsRaiseError m_strCMODULENAME, "ExtractEmployees"
    
End Function


Private Function ExtractQualifications( _
    ByVal v_lngEmployee As Long, _
    ByVal v_rsQuals As ADODB.Recordset _
    ) As String
    
Dim l_intCount As Integer
Dim l_strQuals As String

'Returns up to 5 qualifications for the employee
On Error GoTo eh
    
    If Not RecordsExist(v_rsQuals) Then
        ExtractQualifications = ""
        Exit Function
    End If
    
    v_rsQuals.Filter = "EMP_NUM = " & str(v_lngEmployee)
    If Not RecordsExist(v_rsQuals) Then
        v_rsQuals.Filter = ""
        ExtractQualifications = ""
        Exit Function
    End If
    
    l_intCount = 0
    While Not v_rsQuals.EOF And l_intCount < 4
        l_intCount = l_intCount + 1
        If l_intCount = 1 Then
            l_strQuals = Trim(v_rsQuals!QUAL_CD)
        Else
            l_strQuals = l_strQuals & "," & Trim(v_rsQuals!QUAL_CD)
        End If
        v_rsQuals.MoveNext
    Wend
    v_rsQuals.Filter = ""
    
    ExtractQualifications = l_strQuals
    
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "ExtractQualifications"
    
End Function


Private Function ExtractSkills( _
    ByVal v_lngEmployee As Long, _
    ByVal v_rsSkills As ADODB.Recordset _
    ) As String
    
Dim l_intCount As Integer
Dim l_strSkills As String

'Returns up to 5 special skills for the employee
On Error GoTo eh
    
    If Not RecordsExist(v_rsSkills) Then
        ExtractSkills = ""
        Exit Function
    End If
    
    v_rsSkills.Filter = "EMP_NUM = " & str(v_lngEmployee)
    If Not RecordsExist(v_rsSkills) Then
        v_rsSkills.Filter = ""
        ExtractSkills = ""
        Exit Function
    End If
    
    l_intCount = 0
    While Not v_rsSkills.EOF And l_intCount < 4
        l_intCount = l_intCount + 1
        If l_intCount = 1 Then
            l_strSkills = Trim(v_rsSkills!SPEC_SKILL_CD)
        Else
            l_strSkills = l_strSkills & "," & Trim(v_rsSkills!SPEC_SKILL_CD)
        End If
        v_rsSkills.MoveNext
    Wend
    
    v_rsSkills.Filter = ""
    ExtractSkills = l_strSkills

    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "ExtractSkills"
    
End Function



Private Function ExtractPhone( _
    ByVal v_lngEmployee As Long, _
    ByVal v_rsPhone As ADODB.Recordset _
    ) As String
    
Dim l_strPhone As String

On Error GoTo eh

    If Not RecordsExist(v_rsPhone) Then
        ExtractPhone = ""
        Exit Function
    End If
    
    v_rsPhone.Filter = "EMP_NUM = " & str(v_lngEmployee)
    If Not RecordsExist(v_rsPhone) Then
        l_strPhone = ""
    Else
        l_strPhone = Trim(v_rsPhone!HOME_PHONE_NUM)
    End If
    v_rsPhone.Filter = ""
    
    ExtractPhone = l_strPhone

    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "ExtractPhone"
    
End Function

Public Function GetOTTimeChargeExceptions() As ADODB.Recordset


    Dim l_obj As dbOTCanvass.cdOTCanvass
    Dim l_blnError As Boolean
    Dim l_rs As ADODB.Recordset
    
    On Error GoTo eh
    
    Set l_obj = CreateObject("dbOTCanvass.cdOTCanvass")
    Set l_rs = l_obj.GetOvertimeExceptions()
    Set l_obj = Nothing
    
    l_rs.AddNew
    l_rs!EXCEPTION_CD = "ONP"
    l_rs!EXCEPTION_DESC = "Overtime Not Paid"

    Set GetOTTimeChargeExceptions = l_rs
    

    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "GetOTTimeChargeExceptions", , , , , True

End Function

Public Function GetNonCanvassedOvertimeCanvasses( _
    ByVal v_intLocationCode As Integer, _
    ByVal v_datStartDate As Date, _
    ByVal v_datEndDate As Date, _
    Optional v_intOTGroup As Integer = 0, _
    Optional v_blnOpenCanvassesOnly As Boolean = False, _
    Optional v_blnPostedCanvassesOnly As Boolean = False _
    ) As ADODB.Recordset
    
    Dim l_obj As dbOTCanvass.cdOTCanvass
    
    
    On Error GoTo eh
    
    Set l_obj = CreateObject("dbOTCanvass.cdOTCanvass")
    
    'Begin Prem TRIS 5.6
    Set GetNonCanvassedOvertimeCanvasses = l_obj.GetNonCanavassedOvertimeCanvasses2( _
        v_intLocationCode, v_datStartDate, v_datEndDate, v_intOTGroup, _
        v_blnOpenCanvassesOnly, v_blnPostedCanvassesOnly)
    'End Prem TRIS 5.6
    Set l_obj = Nothing
    
    Exit Function

eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "GetNonCanvassedOvertimeCanvasses", , , , , True
    
End Function


Public Function GetCanvasses( _
    ByVal v_intLocationCode As Integer, _
    ByVal v_datStartDate As Date, _
    ByVal v_datEndDate As Date, _
    Optional v_strCanvassType As String = "A", _
    Optional v_intOTGroup As Integer = 0, _
    Optional v_blnOpenCanvassesOnly As Boolean = False, _
    Optional v_blnPostedCanvassesOnly As Boolean = False _
    ) As ADODB.Recordset

    Dim l_obj As dbOTCanvass.cdOTCanvass


    On Error GoTo eh

    Set l_obj = CreateObject("dbOTCanvass.cdOTCanvass")

    'Begin Prem TRIS 5.6
    Set GetCanvasses = l_obj.GetCanvasses2( _
        v_intLocationCode, v_datStartDate, v_datEndDate, v_strCanvassType, v_intOTGroup, _
        v_blnOpenCanvassesOnly, v_blnPostedCanvassesOnly)
    'End Prem TRIS 5.6
    Set l_obj = Nothing

    Exit Function

eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "GetCanvasses", , , , , True

End Function


Public Function GetCanvassOTBankCharges( _
                ByVal v_lngOTOppNum As Long _
                , ByRef r_rsOTOppBankCharges As ADODB.Recordset _
                , ByRef r_rsOtherOTBankCharges As ADODB.Recordset _
                ) As Boolean
''
End Function


Public Function IsCallawayNewRules( _
                ByVal v_lngOTOppNum As Long _
                , ByRef r_blnOTOppNumExists As Boolean _
                ) As Boolean
' This will be called by the client to determine which process to determine ineligibilities.
' Callaway will use the new rule engine for Fatigue Management rules.
                
    Dim l_obj As dbOTCanvass.cdOTCanvass
    Dim l_objCT As dbCodeTables.cdCodeTables
    Dim rs As ADODB.Recordset
    Dim strEffDt As String
    Dim datEffDt As Date

On Error GoTo eh
    Set l_obj = CreateObject("dbOTCanvass.cdOTCanvass")
    Set rs = l_obj.GetOTOpportunity(v_lngOTOppNum)
    If Not RecordsExist(rs) Then
        r_blnOTOppNumExists = False
        IsCallawayNewRules = False
        Exit Function
    End If
    r_blnOTOppNumExists = True
    
    Set l_obj = Nothing
    Set l_objCT = CreateObject("dbCodeTables.cdCodeTables")
    strEffDt = Trim(l_objCT.GetNRCRuleEffectiveDate)
    If strEffDt = "" Then
        'default
        datEffDt = CDate("01/01/2010")
    Else
        datEffDt = CDate(strEffDt)
    End If
    
    If rs!LOC_CD = 59 _
            And rs!CHARGE_DATE >= datEffDt Then
        IsCallawayNewRules = True
    End If

    Set l_obj = Nothing
    Set rs = Nothing

    Exit Function

eh:
    Set l_obj = Nothing
    Set rs = Nothing
    TrsRaiseError m_strCMODULENAME, "IsCallawayNewRules", , , , , True
    
End Function


Public Function GetEmployeeStatusByOpportunity( _
                    ByVal v_lngOTOppNum As Long _
                ) As ADODB.Recordset

On Error GoTo eh

    Set GetEmployeeStatusByOpportunity = _
            GetEmpStatusByOpp(v_lngOTOppNum, Now)
    
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "GetEmployeeStatusByOpportunity", , , , , True

End Function


Public Function GetEmployeeStatusByOpportunityDate( _
                    ByVal v_lngOTOppNum As Long, _
                    ByVal v_datDateTime As Date _
                ) As ADODB.Recordset

On Error GoTo eh

    Set GetEmployeeStatusByOpportunityDate = _
            GetEmpStatusByOpp(v_lngOTOppNum, v_datDateTime)
            
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "GetEmployeeStatusByOpportunityDate", , , , , True
    
End Function


Public Function GetEmployeeStatusByEmployee( _
                    ByVal v_aryEmpNums As Variant _
                ) As ADODB.Recordset

    Dim l_aryEmpNums As Variant

On Error GoTo eh

    Set GetEmployeeStatusByEmployee = _
            GetEmpStatusByEmp(v_aryEmpNums, Now)
    
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "GetEmployeeStatusByEmployee", , , , , True

End Function


Public Function GetEmployeeStatusByEmployeeDate( _
                    ByVal v_aryEmpNums As Variant, _
                    ByVal v_datDateTime As Date _
                ) As ADODB.Recordset


On Error GoTo eh
   
    Set GetEmployeeStatusByEmployeeDate = _
            GetEmpStatusByEmp(v_aryEmpNums, v_datDateTime)

    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "GetEmployeeStatusByEmployeeDate", , , , , True

End Function


Private Function IsInteger(ByVal v_varData As Variant) As Boolean

    Dim l_intType As Integer

On Error GoTo eh
    l_intType = VarType(v_varData)

    Select Case l_intType
        Case 2, 3
            IsInteger = True
        Case Else
            IsInteger = False
    End Select

    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "IsInteger"

End Function


'Private Function GetEmployeeStatusPrivate( _
'                    ByVal v_lngOTOppNum As Long, _
'                    ByVal v_aryEmpNums As Variant, _
'                    ByVal v_datDateTime As Date _
'                ) As ADODB.Recordset
'
'    Dim l_rsEmpSchedules As ADODB.Recordset
'    Dim l_obj As dbOTCanvass.cdOTCanvass
'    Dim l_rsEmpStatus As ADODB.Recordset
'    Dim l_aryOTOppNum As Variant
'
'On Error GoTo eh
'
'    Set l_rsEmpSchedules = GetEmployeeSchedules( _
'                v_aryEmpNums, v_datDateTime - 1, v_datDateTime + 1)
'
'    Set l_obj = CreateObject("dbOTCanvass.cdOTCanvass")
'
'    If v_lngOTOppNum > 0 Then
'        Set m_rsEmpLocs = l_obj.GetEmployeePrimaryLocation( _
'                            v_lngOTOppNum, v_datDateTime)
'    Else
'        Set m_rsEmpLocs = l_obj.GetEmployeePrimaryLocationFromList( _
'                            v_aryEmpNums, v_datDateTime)
'    End If
'
'    l_aryOTOppNum = GetDistinctListOTOppNum(l_rsEmpSchedules)
'    If Not IsEmpty(l_aryOTOppNum) Then
'        Set m_rsCanvLocs = l_obj.GetCanvassLocation(l_aryOTOppNum)
'    End If
'    Set l_obj = Nothing
'
'    Set l_rsEmpStatus = _
'        BuildEmployeeStatusSet(v_aryEmpNums, v_datDateTime, l_rsEmpSchedules)
'
'    Set GetEmployeeStatusPrivate = l_rsEmpStatus
'
'    Exit Function
'
'eh:
'    TrsRaiseError m_strCMODULENAME, "GetEmployeeStatusPrivate"
'
'End Function


Private Function GetEmpStatusByEmp( _
                    ByVal v_aryEmpNums As Variant, _
                    ByVal v_datDateTime As Date _
                ) As ADODB.Recordset

    Dim l_rsEmpSchedules As ADODB.Recordset
    Dim l_obj As dbOTCanvass.cdOTCanvass
    Dim l_rsEmpStatus As ADODB.Recordset
    Dim l_aryOTOppNum As Variant
    Dim l_aryEmpNums As Variant

On Error GoTo eh

    If IsInteger(v_aryEmpNums) Then
        l_aryEmpNums = Array(v_aryEmpNums)
    Else
        l_aryEmpNums = v_aryEmpNums
    End If

    Set l_obj = CreateObject("dbOTCanvass.cdOTCanvass")
    
    Set m_rsEmpLocs = l_obj.GetEmployeePrimaryLocationFromList( _
                            l_aryEmpNums, v_datDateTime)
    
    Set l_rsEmpSchedules = GetEmployeeSchedules(l_aryEmpNums, _
            MidnightDate(v_datDateTime - 1), MidnightDate(v_datDateTime + 1))
    
    l_aryOTOppNum = GetDistinctListOTOppNum(l_rsEmpSchedules)
    If Not IsEmpty(l_aryOTOppNum) Then
        Set m_rsCanvLocs = l_obj.GetCanvassLocation(l_aryOTOppNum)
    End If
    Set l_obj = Nothing

    Set l_rsEmpStatus = _
        BuildEmployeeStatusSet(l_aryEmpNums, v_datDateTime, l_rsEmpSchedules)

    Set GetEmpStatusByEmp = l_rsEmpStatus

    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "GetEmpStatusByEmp"

End Function


Private Function GetEmpStatusByOpp( _
                    ByVal v_lngOTOppNum As Long, _
                    ByVal v_datDateTime As Date _
                ) As ADODB.Recordset

    Dim l_obj As dbOTCanvass.cdOTCanvass
    Dim l_rsEmpSchedules As ADODB.Recordset
    Dim l_rsEmpStatus As ADODB.Recordset
    Dim l_aryEmpNums As Variant
    Dim l_aryOTOppNum As Variant

On Error GoTo eh

    Set l_obj = CreateObject("dbOTCanvass.cdOTCanvass")
    
    Set m_rsEmpLocs = l_obj.GetEmployeePrimaryLocation( _
                            v_lngOTOppNum, v_datDateTime)
    l_aryEmpNums = ArrayOfFieldValues("EMP_NUM", m_rsEmpLocs)

    Set l_rsEmpSchedules = GetEmployeeSchedules(l_aryEmpNums, _
            MidnightDate(v_datDateTime - 1), MidnightDate(v_datDateTime + 1))
'//debug
'//ViewRecordset l_rsEmpSchedules, "l_rsEmpSchedules602627B"
    l_aryOTOppNum = GetDistinctListOTOppNum(l_rsEmpSchedules)
    
    If Not IsEmpty(l_aryOTOppNum) Then
        Set m_rsCanvLocs = l_obj.GetCanvassLocation(l_aryOTOppNum)
    End If
    Set l_obj = Nothing

    Set l_rsEmpStatus = _
        BuildEmployeeStatusSet(l_aryEmpNums, v_datDateTime, l_rsEmpSchedules)

    Set GetEmpStatusByOpp = l_rsEmpStatus

    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "GetEmpStatusByOpp"

End Function


Private Function GetDistinctListOTOppNum( _
                    ByVal v_rsEmpSchedules As ADODB.Recordset) As Variant

    Dim l_cuSet As utilTRISCommonLib.cuSet

On Error GoTo eh

    If Not RecordsExist(v_rsEmpSchedules) Then
        Exit Function
    End If

    Set l_cuSet = New cuSet
    Do Until v_rsEmpSchedules.EOF
        If v_rsEmpSchedules!OT_OPP_NUM <> 0 Then
            l_cuSet.Add v_rsEmpSchedules.Fields("OT_OPP_NUM").Value
        End If
        v_rsEmpSchedules.MoveNext
    Loop

    GetDistinctListOTOppNum = ArrayFromCollection(l_cuSet.Contents)
    Set l_cuSet = Nothing
    v_rsEmpSchedules.MoveFirst
    Exit Function

eh:
    Set l_cuSet = Nothing
    TrsRaiseError m_strCMODULENAME, "GetDistinctListOTOppNum"

End Function


Private Function BuildEmployeeStatusSet( _
                    ByVal v_aryEmpNums As Variant, _
                    ByVal v_datDateTime As Date, _
                    ByVal v_rsEmpSchedules As ADODB.Recordset _
                ) As ADODB.Recordset

    Dim l_rs As ADODB.Recordset
    Dim l_rsEmpSchedC As ADODB.Recordset
    Dim i As Integer

On Error GoTo eh

    Set l_rs = NewEmployeeStatusSet
    If ArraySize(v_aryEmpNums) = 0 Then
        GoTo xit
    End If

    Set l_rsEmpSchedC = v_rsEmpSchedules.Clone
    For i = LBound(v_aryEmpNums) To UBound(v_aryEmpNums)
        l_rsEmpSchedC.Filter = "EMP_NUM = " & v_aryEmpNums(i)
        l_rsEmpSchedC.Sort = "CHARGE_DATE ASC, DATE_WORKED ASC, START_TIME ASC"
Debug.Print ""
Debug.Print "BuildEmployeeStatusSet: EMP_NUM = " & v_aryEmpNums(i)

        Do Until l_rsEmpSchedC.EOF
If Not (l_rsEmpSchedC.BOF Or l_rsEmpSchedC.EOF) Then
    Debug.Print "BuildEmployeeStatusSet: l_rsEmpSchedC.Bookmark = " & l_rsEmpSchedC.Bookmark
    Debug.Print StringFormat("BuildEmployeeStatusSet: CHARGE_DATE = {0}, DATE_WORKED = {1},  START_TIME = {2}", _
                            l_rsEmpSchedC!CHARGE_DATE, l_rsEmpSchedC!DATE_WORKED, l_rsEmpSchedC!START_TIME)
End If
            If IsRegularScheduledShift(l_rsEmpSchedC) Then
                If Not IsPartialShiftException(l_rsEmpSchedC) Then
                    AddRollupRegularShift l_rsEmpSchedC, l_rs
                Else
                    AddAllShiftDetails l_rsEmpSchedC, l_rs
                End If
            ElseIf IsScheduledOvertime(l_rsEmpSchedC) Then
                AddRollupOvertimeShift l_rsEmpSchedC, l_rs
            Else
                l_rsEmpSchedC.MoveNext
            End If
            'already positioned on next shift
        Loop
    Next

xit:
    Set BuildEmployeeStatusSet = l_rs
    Set l_rs = Nothing
    Set l_rsEmpSchedC = Nothing
    Exit Function
eh:
    Set l_rs = Nothing
    Set l_rsEmpSchedC = Nothing
    TrsRaiseError m_strCMODULENAME, "BuildEmployeeStatusSet"

End Function


Private Function IsPartialShiftException( _
                    ByVal v_rsEmpSchedules As ADODB.Recordset) As Boolean
' Assuming positioned on first schedule detail row of regular scheduled shift.
' Do any schedule entries for this regular shift have different off-duty
' exception codes?

    Dim l_varBookmark As Variant
    Dim l_strExcCd As String
    Dim l_lngEmpNum As Long
    Dim l_datChgDate As Date
    Dim l_strShiftStartTime As String
    Dim l_blnReturn As Boolean
    Dim l_blnShiftChange As Boolean

On Error GoTo eh

    l_varBookmark = v_rsEmpSchedules.Bookmark
    l_strExcCd = Trim(v_rsEmpSchedules!EXCEPTION_CD)
    
'    If Not IsScheduledWorktime(v_rsEmpSchedules) Then
'        l_blnReturn = True
'        GoTo xit
'    End If
    
    l_lngEmpNum = v_rsEmpSchedules!EMP_NUM
    l_datChgDate = v_rsEmpSchedules!CHARGE_DATE
    l_strShiftStartTime = v_rsEmpSchedules!SHIFT_START_TIME
'    v_rsEmpSchedules.MoveNext
    
    Do Until v_rsEmpSchedules.EOF
        If Not (v_rsEmpSchedules!EMP_NUM = l_lngEmpNum And _
                v_rsEmpSchedules!CHARGE_DATE = l_datChgDate And _
                v_rsEmpSchedules!SHIFT_START_TIME = l_strShiftStartTime) Then
            Exit Do
        End If
    
        If IsRegularScheduledShift(v_rsEmpSchedules) Then
            If Trim(v_rsEmpSchedules!EXCEPTION_CD) <> l_strExcCd And _
                    (Not IsActualWork(l_strExcCd) Or Not IsActualWork(v_rsEmpSchedules!EXCEPTION_CD)) And _
                    Trim(v_rsEmpSchedules!SYSTEM_EXC_CD) <> "SLT" Then
                l_blnReturn = True
                GoTo xit
            End If
        End If
        v_rsEmpSchedules.MoveNext
    Loop

xit:
    v_rsEmpSchedules.Bookmark = l_varBookmark
    IsPartialShiftException = l_blnReturn
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "IsPartialShiftException"

End Function


Private Sub AddAllShiftDetails( _
                ByVal v_rsEmpSchedules As ADODB.Recordset, _
                ByRef r_rsEmpStatus As ADODB.Recordset)
' Add all shift details to employee status when there is a partial shift exception.
' Otherwise, rollup shift details into single shift record.
    Dim l_datShiftStart As Date
    Dim l_datShiftEnd As Date
    Dim l_datStartDateTime As Date
    Dim l_datEndDateTime As Date

    Dim l_strExcCd As String

On Error GoTo eh

    l_datShiftStart = _
        DateTimeFrom(v_rsEmpSchedules!SHIFT_START_DATE, v_rsEmpSchedules!SHIFT_START_TIME)
    l_datShiftEnd = _
        DateTimeFrom(v_rsEmpSchedules!SHIFT_END_DATE, v_rsEmpSchedules!SHIFT_END_TIME, True)

    Do Until v_rsEmpSchedules.EOF
        If IsRegularScheduledShift(v_rsEmpSchedules) Then
            l_datStartDateTime = DateTimeFrom(v_rsEmpSchedules!DATE_WORKED, v_rsEmpSchedules!START_TIME)
            l_datEndDateTime = DateTimeFrom(v_rsEmpSchedules!DATE_WORKED, v_rsEmpSchedules!END_TIME, True)
            If l_datStartDateTime >= l_datShiftEnd Then
                Exit Do
            End If
            If TimePeriodsIntersect( _
                    l_datShiftStart, l_datShiftEnd, _
                    l_datStartDateTime, l_datEndDateTime) Then
                r_rsEmpStatus.AddNew
                r_rsEmpStatus!EMP_NUM = v_rsEmpSchedules!EMP_NUM
                r_rsEmpStatus!START_DATE = v_rsEmpSchedules!DATE_WORKED
                r_rsEmpStatus!START_TIME = v_rsEmpSchedules!START_TIME
                r_rsEmpStatus!END_DATE = v_rsEmpSchedules!DATE_WORKED
                r_rsEmpStatus!END_TIME = v_rsEmpSchedules!END_TIME
                
                r_rsEmpStatus!EXCEPTION_CD = v_rsEmpSchedules!EXCEPTION_CD
                r_rsEmpStatus!OVER_EXC_CD = v_rsEmpSchedules!OVER_EXC_CD
                r_rsEmpStatus!EXCEPTION_DESC = _
                    LookupExceptionDescription(v_rsEmpSchedules!EXCEPTION_CD)
                
                r_rsEmpStatus!EMP_NAME = _
                    LookupEmployeeLocationField( _
                            v_rsEmpSchedules!EMP_NUM, "EMP_NAME", m_rsEmpLocs)
                            
                l_strExcCd = Trim(v_rsEmpSchedules!EXCEPTION_CD)
                If IsFentonTraining(l_strExcCd) Then
                    r_rsEmpStatus!LOC_CD = "60"
                    r_rsEmpStatus!LOC_DESC = "Fenton Training Center"
                Else
                    r_rsEmpStatus!LOC_CD = _
                        LookupEmployeeLocationField( _
                                v_rsEmpSchedules!EMP_NUM, "LOC_CD", m_rsEmpLocs)
                    r_rsEmpStatus!LOC_DESC = _
                        LookupEmployeeLocationField( _
                                v_rsEmpSchedules!EMP_NUM, "LOC_DESC", m_rsEmpLocs)
                End If
                
                r_rsEmpStatus!OT_OPP_NUM = v_rsEmpSchedules!OT_OPP_NUM
                r_rsEmpStatus!OT_OPP_DESC = ""
            End If
        End If
        v_rsEmpSchedules.MoveNext
    Loop

    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "AddAllShiftDetails"

End Sub


Private Sub AddRollupRegularShift( _
                ByVal v_rsEmpSchedules As ADODB.Recordset, _
                ByRef r_rsEmpStatus As ADODB.Recordset)
' Add all shift details to employee status when there is a partial shift exception.
' Otherwise, rollup shift details into single shift record.

    Dim l_lngEmpNum As Long
    Dim l_datChgDate As Date
    Dim l_strShiftStartTime As String

On Error GoTo eh

    AddRegularScheduledShift v_rsEmpSchedules, r_rsEmpStatus
    
    l_lngEmpNum = v_rsEmpSchedules!EMP_NUM
    l_datChgDate = v_rsEmpSchedules!CHARGE_DATE
    l_strShiftStartTime = v_rsEmpSchedules!SHIFT_START_TIME
' position on next shift
    Do Until v_rsEmpSchedules.EOF
        If Not (v_rsEmpSchedules!EMP_NUM = l_lngEmpNum And _
                v_rsEmpSchedules!CHARGE_DATE = l_datChgDate And _
                v_rsEmpSchedules!SHIFT_START_TIME = l_strShiftStartTime) Then
            Exit Do
        End If
        v_rsEmpSchedules.MoveNext
    Loop
        
    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "AddRollupRegularShift"

End Sub


Private Sub AddRollupOvertimeShift( _
                ByVal v_rsEmpSchedules As ADODB.Recordset, _
                ByRef r_rsEmpStatus As ADODB.Recordset)
    
    Dim l_varBookmark As Variant
    Dim l_datStartDate As Date
    Dim l_strStartTime As String
    Dim l_datEndDate As Date
    Dim l_strEndTime As String

    Dim l_strExcCd As String
    Dim l_strLoc As String
    
On Error GoTo eh

    l_datStartDate = v_rsEmpSchedules!DATE_WORKED
    l_strStartTime = v_rsEmpSchedules!START_TIME

    l_varBookmark = LastBookmark( _
        v_rsEmpSchedules!EMP_NUM, v_rsEmpSchedules!OT_OPP_NUM, v_rsEmpSchedules)
    If Not IsEmpty(l_varBookmark) Then
        v_rsEmpSchedules.Bookmark = l_varBookmark
    End If
    
    l_datEndDate = v_rsEmpSchedules!DATE_WORKED
    l_strEndTime = v_rsEmpSchedules!END_TIME
    
    r_rsEmpStatus.AddNew
    r_rsEmpStatus!EMP_NUM = v_rsEmpSchedules!EMP_NUM
    r_rsEmpStatus!START_DATE = l_datStartDate
    r_rsEmpStatus!START_TIME = l_strStartTime
    r_rsEmpStatus!END_DATE = l_datEndDate
    r_rsEmpStatus!END_TIME = l_strEndTime
    
    r_rsEmpStatus!EXCEPTION_CD = v_rsEmpSchedules!EXCEPTION_CD
    r_rsEmpStatus!OVER_EXC_CD = v_rsEmpSchedules!OVER_EXC_CD
    r_rsEmpStatus!EXCEPTION_DESC = _
        LookupExceptionDescription(v_rsEmpSchedules!EXCEPTION_CD)
    
    r_rsEmpStatus!EMP_NAME = _
        LookupEmployeeLocationField( _
                v_rsEmpSchedules!EMP_NUM, "EMP_NAME", m_rsEmpLocs)
                
    l_strExcCd = Trim(v_rsEmpSchedules!EXCEPTION_CD)
    If IsFentonTraining(l_strExcCd) Then
        r_rsEmpStatus!LOC_CD = "60"
        r_rsEmpStatus!LOC_DESC = "Fenton Training Center"
    Else
        l_strLoc = LookupOvertimeLocationField( _
                    v_rsEmpSchedules!OT_OPP_NUM, "LOC_CD", m_rsCanvLocs)
        r_rsEmpStatus!LOC_CD = l_strLoc
        If Trim(l_strLoc) = "0" Then
            r_rsEmpStatus!LOC_DESC = " "
        Else
            r_rsEmpStatus!LOC_DESC = _
                LookupOvertimeLocationField( _
                    v_rsEmpSchedules!OT_OPP_NUM, "LOC_DESC", m_rsCanvLocs)
        End If
    End If
    r_rsEmpStatus!OT_OPP_NUM = v_rsEmpSchedules!OT_OPP_NUM
    If Trim(l_strLoc) = "0" Then
        r_rsEmpStatus!OT_OPP_DESC = " "
    Else
        r_rsEmpStatus!OT_OPP_DESC = _
            LookupOvertimeLocationField( _
                v_rsEmpSchedules!OT_OPP_NUM, "OT_OPP_DESC", m_rsCanvLocs)
    End If

xit:
    v_rsEmpSchedules.MoveNext   'position on next shift
    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "AddRollupOvertimeShift"

End Sub


Private Function LastBookmark( _
                    ByVal v_lngEmpNum As Long, _
                    ByVal v_lngOTOppNum As Long, _
                    ByVal v_rsEmpSchedules As ADODB.Recordset) As Variant

    Dim l_rsEmpSched As ADODB.Recordset
    Dim l_varBookmark As Variant

On Error GoTo eh

    Set l_rsEmpSched = v_rsEmpSchedules.Clone
    l_rsEmpSched.Filter = "EMP_NUM = " & v_lngEmpNum & _
                        " AND OT_OPP_NUM = " & v_lngOTOppNum
    l_rsEmpSched.Sort = "DATE_WORKED ASC, START_TIME ASC"
    l_rsEmpSched.MoveLast
If Not (l_rsEmpSched.BOF Or l_rsEmpSched.EOF) Then
    Debug.Print "LastBookmark(in): EMP_NUM = " & v_lngEmpNum & ", OT_OPP_NUM = " & v_lngOTOppNum
End If

    If l_rsEmpSched.BOF Or l_rsEmpSched.EOF Then
        LastBookmark = Empty
    Else
        LastBookmark = l_rsEmpSched.Bookmark
    End If

Debug.Print "LastBookmark(out): = " & IIf(IsEmpty(LastBookmark), "Empty", LastBookmark)
If Not (l_rsEmpSched.BOF Or l_rsEmpSched.EOF) Then
    Debug.Print "LastBookmark(out): = DATE_WORKED = " & l_rsEmpSched!DATE_WORKED & ", START_TIME = " & l_rsEmpSched!START_TIME
Else

End If

    Set l_rsEmpSched = Nothing
    Exit Function

eh:
    Set l_rsEmpSched = Nothing
    TrsRaiseError m_strCMODULENAME, "LastBookmark"
    
End Function


Private Sub PositionOnNextShift(ByVal v_rsEmpSchedules As ADODB.Recordset)

' this is breaking based on current positioned on regular shift.
    Dim l_lngEmpNum As Long
    Dim l_datChgDate As Date
    Dim l_strShiftStartTime As String

On Error GoTo eh
    
    l_lngEmpNum = v_rsEmpSchedules!EMP_NUM
    l_datChgDate = v_rsEmpSchedules!CHARGE_DATE
    l_strShiftStartTime = v_rsEmpSchedules!SHIFT_START_TIME
' position on next shift
    Do Until v_rsEmpSchedules.EOF
        If Not (v_rsEmpSchedules!EMP_NUM = l_lngEmpNum And _
                v_rsEmpSchedules!CHARGE_DATE = l_datChgDate And _
                v_rsEmpSchedules!SHIFT_START_TIME = l_strShiftStartTime) Then
            Exit Do
        End If
        v_rsEmpSchedules.MoveNext
    Loop

    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "PositionOnNextShift"

End Sub


Private Sub AddRegularScheduledShift( _
                ByVal v_rsEmpSchedules As ADODB.Recordset, _
                ByRef r_rsEmpStatus As ADODB.Recordset)

' Add all shift details to employee status when there is a partial shift exception.
' Otherwise, rollup shift details into single shift record.

    Dim l_strExcCd As String

On Error GoTo eh

    r_rsEmpStatus.AddNew
    r_rsEmpStatus!EMP_NUM = v_rsEmpSchedules!EMP_NUM
    r_rsEmpStatus!START_DATE = v_rsEmpSchedules!SHIFT_START_DATE
    r_rsEmpStatus!START_TIME = v_rsEmpSchedules!SHIFT_START_TIME
    r_rsEmpStatus!END_DATE = v_rsEmpSchedules!SHIFT_END_DATE
    r_rsEmpStatus!END_TIME = v_rsEmpSchedules!SHIFT_END_TIME
    r_rsEmpStatus!EXCEPTION_CD = v_rsEmpSchedules!EXCEPTION_CD
    r_rsEmpStatus!OVER_EXC_CD = v_rsEmpSchedules!OVER_EXC_CD
    r_rsEmpStatus!EXCEPTION_DESC = _
        LookupExceptionDescription(v_rsEmpSchedules!EXCEPTION_CD)
    r_rsEmpStatus!EMP_NAME = _
        LookupEmployeeLocationField( _
                v_rsEmpSchedules!EMP_NUM, "EMP_NAME", m_rsEmpLocs)
    l_strExcCd = Trim(v_rsEmpSchedules!EXCEPTION_CD)
    If IsFentonTraining(l_strExcCd) Then
        r_rsEmpStatus!LOC_CD = "60"
        r_rsEmpStatus!LOC_DESC = "Fenton Training Center"
    Else
        r_rsEmpStatus!LOC_CD = _
            LookupEmployeeLocationField( _
                    v_rsEmpSchedules!EMP_NUM, "LOC_CD", m_rsEmpLocs)
        r_rsEmpStatus!LOC_DESC = _
            LookupEmployeeLocationField( _
                    v_rsEmpSchedules!EMP_NUM, "LOC_DESC", m_rsEmpLocs)
    End If
    r_rsEmpStatus!OT_OPP_NUM = 0
    r_rsEmpStatus!OT_OPP_DESC = ""

    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "AddRegularScheduledShift"

End Sub

Private Function IsFentonTraining(ByVal v_strExcCd As String) As Boolean

    IsFentonTraining = v_strExcCd = "TF" Or v_strExcCd = "TFI"

End Function



Private Function NewEmployeeStatusSet() As ADODB.Recordset

    Const l_lngCFIELDATTR = adFldUpdatable Or adFldIsNullable
    Dim l_rs As ADODB.Recordset

On Error GoTo eh

    Set l_rs = New ADODB.Recordset
    With l_rs
        .Fields.Append "EMP_NUM", adInteger, , l_lngCFIELDATTR
        .Fields.Append "EMP_NAME", adChar, 30, l_lngCFIELDATTR
        .Fields.Append "START_DATE", adDate, , l_lngCFIELDATTR
        .Fields.Append "START_TIME", adChar, 5, l_lngCFIELDATTR
        .Fields.Append "END_DATE", adDate, , l_lngCFIELDATTR
        .Fields.Append "END_TIME", adChar, 5, l_lngCFIELDATTR
        .Fields.Append "LOC_CD", adChar, 5, l_lngCFIELDATTR
        .Fields.Append "LOC_DESC", adChar, 30, l_lngCFIELDATTR
        .Fields.Append "EXCEPTION_CD", adChar, 3, l_lngCFIELDATTR
        .Fields.Append "OVER_EXC_CD", adChar, 3, l_lngCFIELDATTR
        .Fields.Append "EXCEPTION_DESC", adChar, 30, l_lngCFIELDATTR
        .Fields.Append "OT_OPP_NUM", adInteger, , l_lngCFIELDATTR
        .Fields.Append "OT_OPP_DESC", adChar, 30, l_lngCFIELDATTR

        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open
    End With

    Set NewEmployeeStatusSet = l_rs
    Set l_rs = Nothing

    Exit Function

eh:
    Set l_rs = Nothing
    Set NewEmployeeStatusSet = Nothing
    TrsRaiseError m_strCMODULENAME, "NewEmployeeStatusSet"

End Function


Private Function LookupEmployeeLocationField( _
                ByVal v_lngEmpNum As Long, _
                ByVal v_strFieldName As String, _
                ByVal v_rsEmpLocs As ADODB.Recordset) As String

On Error GoTo eh

    If FindEmployeeLocation(v_lngEmpNum, v_rsEmpLocs) Then
        LookupEmployeeLocationField = v_rsEmpLocs.Fields(v_strFieldName).Value
    Else
        LookupEmployeeLocationField = ""
    End If

    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "LookupEmployeeLocationField"
                
End Function


Private Function LookupExceptionDescription( _
                    ByVal v_strExcCd As String) As String

On Error GoTo eh

    m_rsAllExceptions.MoveFirst
    m_rsAllExceptions.Find "EXCEPTION_CD = '" & v_strExcCd & "'"
    If m_rsAllExceptions.EOF Then
        LookupExceptionDescription = ""
    Else
        LookupExceptionDescription = Trim(m_rsAllExceptions!EXCEPTION_DESC)
    End If

    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "LookupExceptionDescription"
                
End Function


Private Function FindEmployeeLocation( _
                ByVal v_lngEmpNum As Long, _
                ByRef r_rsEmpLocs As ADODB.Recordset) As Boolean

On Error GoTo eh

    If Not RecordsExist(r_rsEmpLocs) Then
        FindEmployeeLocation = False
        Exit Function
    ElseIf Not r_rsEmpLocs.EOF Then
        If r_rsEmpLocs!EMP_NUM = v_lngEmpNum Then
            FindEmployeeLocation = True
            Exit Function
        End If
    End If

    r_rsEmpLocs.MoveFirst
    r_rsEmpLocs.Find "EMP_NUM = " & v_lngEmpNum
    FindEmployeeLocation = Not r_rsEmpLocs.EOF

    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "FindEmployeeLocation"
                
End Function


Private Function LookupOvertimeLocationField( _
                ByVal v_lngOTOppNum As Long, _
                ByVal v_strFieldName As String, _
                ByVal v_rsOTLocs As ADODB.Recordset) As String

On Error GoTo eh

    If FindOvertimeLocation(v_lngOTOppNum, v_rsOTLocs) Then
        LookupOvertimeLocationField = v_rsOTLocs.Fields(v_strFieldName).Value
    Else
        LookupOvertimeLocationField = "0"
    End If

    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "LookupOvertimeLocationField"
                
End Function


Private Function FindOvertimeLocation( _
                ByVal v_lngOTOppNum As Long, _
                ByRef r_rsOTLocs As ADODB.Recordset) As Boolean

On Error GoTo eh

    If Not RecordsExist(r_rsOTLocs) Then
        FindOvertimeLocation = False
        Exit Function
'    ElseIf r_rsOTLocs!OT_OPP_NUM = v_lngOTOppNum Then
'        FindOvertimeLocation = True
'        Exit Function
    End If

    r_rsOTLocs.MoveFirst
    r_rsOTLocs.Find "OT_OPP_NUM = " & v_lngOTOppNum
    FindOvertimeLocation = Not r_rsOTLocs.EOF

    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "FindOvertimeLocation"
                
End Function



