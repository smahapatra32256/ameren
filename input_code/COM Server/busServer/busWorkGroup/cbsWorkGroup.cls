VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cbsWorkGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const m_strCMODULENAME As String = "cbsWorkGroup"

Public Function GetWorkGroup(ByVal v_lngEmpNum As Long, ByVal v_datStartDate As Date, ByVal v_datEndDate As Date) As ADODB.Recordset
    
    Dim l_objWkGp As dbWorkGroup.cdWorkGroup
    
    Dim l_objRS As Recordset
    
    On Error GoTo GetWorkGroup_EH
    
    
    Set l_objWkGp = CreateObject("dbWorkGroup.cdWorkGroup")

    Set l_objRS = l_objWkGp.GetWorkGroup(v_lngEmpNum, v_datStartDate, v_datEndDate)
    Set GetWorkGroup = l_objRS
    Set l_objRS = Nothing
    Set l_objWkGp = Nothing
    
    Exit Function

GetWorkGroup_EH:
    Set l_objRS = Nothing
    Set l_objWkGp = Nothing
    
    
    TrsRaiseError m_strCMODULENAME, "GetWorkGroup", , , , , True

End Function


Public Function GetAccounting(ByVal v_lngWgNum As Long) As ADODB.Recordset
    
'Similar to GetWorkGroupDistribition but requires less database resources
'The intent is to just read the data
    
    Dim l_objWkGp As dbWorkGroup.cdWorkGroup
    
    Dim l_objRS As Recordset
    
    On Error GoTo GetAccounting_EH
    
    
    Set l_objWkGp = CreateObject("dbWorkGroup.cdWorkGroup")

    Set l_objRS = l_objWkGp.GetAccounting(v_lngWgNum)
    Set GetAccounting = l_objRS
    Set l_objRS = Nothing
    Set l_objWkGp = Nothing
    
    Exit Function

GetAccounting_EH:
    Set l_objRS = Nothing
    Set l_objWkGp = Nothing
    
    TrsRaiseError m_strCMODULENAME, "GetAccounting", , , , , True

End Function


Public Function GetDOJMDistrictAndHeadquarters(ByVal v_lngWGNumber As Long, r_strDOJMDistrict, r_strDOJMHeadQtr)
    
    Dim l_strPrefixC As String
    Dim l_strPrefixU As String
    Dim l_objWkGp As dbWorkGroup.cdWorkGroup
    
    Dim l_strCrewNum As String
    Dim l_objRS As Recordset
    
    
    
    On Error GoTo GetDOJMDistrictAndHeadquarters_EH
    
    Select Case v_lngWGNumber
        Case 100 To 999
            l_strPrefixC = "C00"
            l_strPrefixU = "U00"
        Case 1000 To 1899
        'this is CIPS DOJM C00 or U00
            l_strPrefixC = "C00"
            l_strPrefixU = "U00"
        Case 1900 To 1999
            'C0 or U0
            l_strPrefixC = "C0"
            l_strPrefixU = "U0"
        Case 2000 To 2899
        'this is CILCO DOJM C00 or U00
            l_strPrefixC = "C00"
            l_strPrefixU = "U00"
        Case Else
        'Not DOJM
            r_strDOJMDistrict = ""
            r_strDOJMHeadQtr = ""
            Exit Function
    End Select
    
    Set l_objWkGp = CreateObject("dbWorkGroup.cdWorkGroup")

    l_strCrewNum = l_objWkGp.GetCrewNumber(v_lngWGNumber)
    
    If Trim(l_strCrewNum) = "" Then
        l_strCrewNum = CStr(v_lngWGNumber)
' rsg 7/31/00 use the workgroup when crew number is not found.  Is this correct?
'Juan - review WHITE089
'       r_strDOJMDistrict = ""
'       r_strDOJMHeadQtr = ""
'       Exit Function
    End If
        
    Set l_objRS = l_objWkGp.GetDOJMDistrictAndHeadquarters(l_strPrefixC & l_strCrewNum)
    
    If l_objRS.BOF And l_objRS.EOF Then
        Set l_objRS = l_objWkGp.GetDOJMDistrictAndHeadquarters(l_strPrefixU & l_strCrewNum)
        If l_objRS.BOF And l_objRS.EOF Then
           Set l_objRS = Nothing
           Set l_objWkGp = Nothing
           
           Exit Function
        End If
    End If
    
    r_strDOJMDistrict = l_objRS.Fields(0)
    r_strDOJMHeadQtr = Left(l_objRS.Fields(1), 3)
    
    Set l_objWkGp = Nothing
    Set l_objRS = Nothing
    
    Exit Function

GetDOJMDistrictAndHeadquarters_EH:
    Set l_objRS = Nothing
    Set l_objWkGp = Nothing
    
    TrsRaiseError m_strCMODULENAME, "GetDOJMDistrictAndHeadquarters", , , , , True

End Function


Public Function GetEmployeesInWorkGroup( _
        ByVal v_aryWorkGroups As Variant, _
        ByVal v_datStartDate As Date, _
        ByVal v_datEndDate As Date, _
        Optional v_bolActiveEmployeesOnly As Boolean = False _
        ) As ADODB.Recordset
        
'Eventually this method should go away
'The client should use GetEmployeesInWorkGroupViewTime
' or GetEmployeesAssignedToWorkGroup instead
'This method gives the same results as GetEmployeesInWorkGroupViewTime
    
    Dim l_objWkGp As dbWorkGroup.cdWorkGroup
    
    Dim l_objRS As Recordset
    Dim l_aryWorkGroups As Variant
    
    On Error GoTo GetEmployeesInWorkGroup_EH
    
    If IsArray(v_aryWorkGroups) Then
        l_aryWorkGroups = v_aryWorkGroups
    Else
        ReDim l_aryWorkGroups(0)
        l_aryWorkGroups(0) = v_aryWorkGroups
    End If
    
    Set l_objWkGp = CreateObject("dbWorkGroup.cdWorkGroup")
    Set l_objRS = l_objWkGp.GetEmployeesInWorkGroup(l_aryWorkGroups, v_datStartDate, v_datEndDate, v_bolActiveEmployeesOnly)
    Set GetEmployeesInWorkGroup = l_objRS
    Set l_objRS = Nothing
    Set l_objWkGp = Nothing

    Exit Function

GetEmployeesInWorkGroup_EH:
    Set l_objRS = Nothing
    Set l_objWkGp = Nothing
    
    TrsRaiseError m_strCMODULENAME, "GetEmployeesInWorkGroup", , , , , True

End Function

Public Function GetActivities(ByVal v_lngWgNum As Long) As ADODB.Recordset
    
'Similar to GetWorkGroupActivities but requires less database resources
'The intent is to just read the data
    
    Dim l_objWkGp As dbWorkGroup.cdWorkGroup
    
    Dim l_objRS As Recordset
    
    On Error GoTo GetActivities_EH
    
    
    Set l_objWkGp = CreateObject("dbWorkGroup.cdWorkGroup")

    Set l_objRS = l_objWkGp.GetActivities(v_lngWgNum)
    Set GetActivities = l_objRS
    Set l_objRS = Nothing
    Set l_objWkGp = Nothing
    
    Exit Function

GetActivities_EH:
    Set l_objRS = Nothing
    Set l_objWkGp = Nothing
    
    TrsRaiseError m_strCMODULENAME, "GetActivities", , , , , True

End Function

Public Function GetSteps(ByVal v_sJobTask As String, ByVal v_sDBID As String) As ADODB.Recordset
On Error GoTo eh

    Set GetSteps = GetStepsPrivate(v_sJobTask, v_sDBID)
    
    '//HACK: sometimes the EMPRV database id is blank rather than 'SUBS'
    If Not RecordsExist(GetSteps) And Trim(v_sDBID) = "" Then
        Set GetSteps = GetStepsPrivate(v_sJobTask, "SUBS")
    End If
    Exit Function
eh:
    Set GetSteps = Nothing
    
    TrsRaiseError m_strCMODULENAME, "GetSteps", , , , , True
End Function

Private Function GetStepsPrivate(ByVal v_sJobTask As String, ByVal v_sDBID As String) As ADODB.Recordset
    
'Similar to GetWorkGroupSteps but requires less database resources
'The intent is to just read the data
    
    Dim l_objWkGp As dbWorkGroup.cdWorkGroup
    
    On Error GoTo GetSteps_EH
    
    
    Set l_objWkGp = CreateObject("dbWorkGroup.cdWorkGroup")

    Dim l_RSsteps As ADODB.Recordset
    Set l_RSsteps = l_objWkGp.GetSteps(v_sJobTask, v_sDBID)
    
    If Not l_RSsteps Is Nothing Then
        Dim l_rsgetsteps As ADODB.Recordset
        Set l_rsgetsteps = New ADODB.Recordset
        
        With l_rsgetsteps
            .Fields.Append "STEP_NBR", adInteger
            .Fields.Append "STEP_DESC", adVarChar, 60
            
            .CursorLocation = adUseClient
            .LockType = adLockBatchOptimistic
            .CursorType = adOpenStatic
            .Open
            .ActiveConnection = Nothing
        
            If RecordsExist(l_RSsteps) Then
                l_RSsteps.MoveFirst
                While Not l_RSsteps.EOF
                    .AddNew
                    !STEP_NBR = l_RSsteps.Fields("step_nbr").Value
                    !STEP_DESC = l_RSsteps.Fields("step_desc").Value
                    .Update
                    l_RSsteps.MoveNext
                Wend
            End If
        End With
    End If

    Set GetStepsPrivate = l_rsgetsteps
    Set l_objWkGp = Nothing
    
    Exit Function

GetSteps_EH:
    Set l_RSsteps = Nothing
    Set l_rsgetsteps = Nothing
    Set l_objWkGp = Nothing
    
    TrsRaiseError m_strCMODULENAME, "GetStepsPrivate", , , , , True

End Function

Public Function GetTransferWorkGroup() As ADODB.Recordset
    
    Dim l_objWkGp As dbWorkGroup.cdWorkGroup
    
    Dim l_objRS As Recordset
    
    On Error GoTo GetTransferWorkGroup_EH
    
    
    Set l_objWkGp = CreateObject("dbWorkGroup.cdWorkGroup")

    Set l_objRS = l_objWkGp.GetWorkgroupRecord(GetTransferWG())
    Set GetTransferWorkGroup = l_objRS
    Set l_objRS = Nothing
    Set l_objWkGp = Nothing
    
    Exit Function

GetTransferWorkGroup_EH:
    Set l_objRS = Nothing
    Set l_objWkGp = Nothing
    
    TrsRaiseError m_strCMODULENAME, "GetTransferWorkGroup", , , , , True

End Function
Public Function GetEmployeeAssignmentsInWorkGroup( _
        ByVal v_lngWorkGroup As Long, _
        ByVal v_datDate As Date _
        ) As ADODB.Recordset
    
    Dim l_objWkGp As dbWorkGroup.cdWorkGroup
    
    Dim l_objRS As Recordset
    Dim l_objRS2 As Recordset
    
    On Error GoTo GetEmployeeAssignmentsInWorkGroup_EH
    
    
    Set l_objWkGp = CreateObject("dbWorkGroup.cdWorkGroup")

    Set l_objRS = l_objWkGp.GetEmployeeAssignmentsInWorkGroup(v_lngWorkGroup, v_datDate)
    Set GetEmployeeAssignmentsInWorkGroup = l_objRS
    Set l_objRS = Nothing
    Set l_objWkGp = Nothing

    Exit Function

GetEmployeeAssignmentsInWorkGroup_EH:
    Set l_objRS = Nothing
    Set l_objWkGp = Nothing
    
    TrsRaiseError m_strCMODULENAME, "GetEmployeeAssignmentsInWorkGroup", , , , , True

End Function

Public Function GetWorkGroupDistribution(ByVal v_lngWgNum As Long) As ADODB.Recordset

'Similar to GetAccounting but requires more database resources
'The intent is to update the data
    
    Dim l_objWkGp As dbWorkGroup.cdWorkGroup
    
    Dim l_objRS As Recordset
    
    On Error GoTo GetWorkGroupDistribution_EH
    
    
    Set l_objWkGp = CreateObject("dbWorkGroup.cdWorkGroup")

    Set l_objRS = l_objWkGp.GetWorkGroupDistribution(v_lngWgNum)
    Set GetWorkGroupDistribution = l_objRS
    Set l_objRS = Nothing
    Set l_objWkGp = Nothing
    
    Exit Function

GetWorkGroupDistribution_EH:
    Set l_objRS = Nothing
    Set l_objWkGp = Nothing
    
    TrsRaiseError m_strCMODULENAME, "GetWorkGroupDistribution", , , , , True

End Function


Public Function GetWorkGroupActivities(ByVal v_lngWgNum As Long) As ADODB.Recordset
    
    Dim l_objWkGp As dbWorkGroup.cdWorkGroup
    Dim l_objRS As Recordset
    
'Similar to GetActivities but requires more database resources
'The intent is to update the data

    
    On Error GoTo GetWorkGroupActivities_EH
    
    
    Set l_objWkGp = CreateObject("dbWorkGroup.cdWorkGroup")

    Set l_objRS = l_objWkGp.GetWorkGroupActivities(v_lngWgNum)
    Set GetWorkGroupActivities = l_objRS
    Set l_objRS = Nothing
    Set l_objWkGp = Nothing
    
    Exit Function

GetWorkGroupActivities_EH:
    Set l_objRS = Nothing
    Set l_objWkGp = Nothing
    
    TrsRaiseError m_strCMODULENAME, "GetWorkGroupActivities", , , , , True

End Function


Public Function GetWorkGroupAttributes( _
            ByVal v_lngWorkGroups As Variant _
            ) As ADODB.Recordset

    Dim l_rs As ADODB.Recordset
    Dim l_obj As dbWorkGroup.cdWorkGroup

On Error GoTo eh


    'Set l_obj = CreateObject("dbWorkGroup.cdWorkGroup")
    Dim l As Object
    Set l = CreateObject("dbWorkGroup.cdWorkGroup")
    Set l_rs = l.GetWorkGroupAttributes(v_lngWorkGroups)
    Set GetWorkGroupAttributes = l_rs
    Exit Function
    
eh:
    
  TrsRaiseError m_strCMODULENAME, "GetWorkGroupAttributes"

End Function

Public Function GetEmployeesAssignedToWorkGroup( _
        ByVal v_aryWorkGroups As Variant, _
        ByVal v_datStartDate As Date, _
        ByVal v_datEndDate As Date, _
        Optional v_bolActiveEmployeesOnly As Boolean = False _
        ) As ADODB.Recordset
    
    Dim l_objWkGp As dbWorkGroup.cdWorkGroup
    
    Dim l_objRS As Recordset
    Dim l_aryWorkGroups As Variant
    
    On Error GoTo eh
    
    If IsArray(v_aryWorkGroups) Then
        l_aryWorkGroups = v_aryWorkGroups
    Else
        ReDim l_aryWorkGroups(0)
        l_aryWorkGroups(0) = v_aryWorkGroups
    End If
    
    Set l_objWkGp = CreateObject("dbWorkGroup.cdWorkGroup")
    Set l_objRS = l_objWkGp.GetEmployeesAssignedToWorkGroup(l_aryWorkGroups, v_datStartDate, v_datEndDate, v_bolActiveEmployeesOnly)
    Set GetEmployeesAssignedToWorkGroup = l_objRS
    Set l_objRS = Nothing
    Set l_objWkGp = Nothing

    Exit Function

eh:
    Set l_objRS = Nothing
    Set l_objWkGp = Nothing
    
    TrsRaiseError m_strCMODULENAME, "GetEmployeesAssignedToWorkGroup", , , , , True

End Function

Public Function GetEmployeesInWorkGroupViewTime( _
        ByVal v_aryWorkGroups As Variant, _
        ByVal v_datStartDate As Date, _
        ByVal v_datEndDate As Date, _
        Optional v_bolActiveEmployeesOnly As Boolean = False _
        ) As ADODB.Recordset
    
    Dim l_objWkGp As dbWorkGroup.cdWorkGroup
    
    Dim l_objRS As Recordset
    Dim l_aryWorkGroups As Variant
    
    On Error GoTo eh
    
    If IsArray(v_aryWorkGroups) Then
        l_aryWorkGroups = v_aryWorkGroups
    Else
        ReDim l_aryWorkGroups(0)
        l_aryWorkGroups(0) = v_aryWorkGroups
    End If
    
    Set l_objWkGp = CreateObject("dbWorkGroup.cdWorkGroup")
    Set l_objRS = l_objWkGp.GetEmployeesInWorkGroupViewTime(l_aryWorkGroups, v_datStartDate, v_datEndDate, v_bolActiveEmployeesOnly)
    Set GetEmployeesInWorkGroupViewTime = l_objRS
    Set l_objRS = Nothing
    Set l_objWkGp = Nothing

    Exit Function

eh:
    Set l_objRS = Nothing
    Set l_objWkGp = Nothing
    
    TrsRaiseError m_strCMODULENAME, "GetEmployeesInWorkGroupViewTime", , , , , True

End Function
'Function gets WorkGroup Schedule for WG Schedule Report
Public Function GetWGSchedule(ByVal vlngWGNum As Long, ByVal vdtStartDate As Date, _
                                ByVal vdtEndDate As Date, ByVal vblnUnmannedLines As Boolean) As ADODB.Recordset

    On Error GoTo GetWGSchedule_eh

    Dim ladrsWG As ADODB.Recordset
    Dim lobjWG As dbWorkGroup.cdWorkGroup
    Dim ladrsEmpWG As ADODB.Recordset
    Dim ladrsEmpsinWG As ADODB.Recordset
    Dim llngIndex As Long
    Dim llngPrevEmpNum As Long
    Dim llngPrevActimEmp As Long
    Dim ldtPrevActimDate As Date
    Dim llngEmpCount As Long
    Dim lobjSched As busSchedule.cbsSchedule
    Dim lobjdbSched As dbSchedule.cdSchedule
    Dim lobjTimeCard As dbTimeCard.cdTimeCard
    Dim lobjCodeTable As dbCodeTables.cdCodeTables
    Dim lobjEmployee As busEmployee.cbsEmployee
    Dim lobjEmp As dbEmployee.cdEmployee

    Dim ladrsJobPosition As ADODB.Recordset
    Dim ladrsProjectInfo As ADODB.Recordset
    Dim ladrsActualTime As ADODB.Recordset
    Dim ladrsEmpSched As ADODB.Recordset
    Dim ladrsEmpSchedLine As ADODB.Recordset
    Dim ladrsUnmannedSched As ADODB.Recordset
    Dim ladrsDummy As ADODB.Recordset
    Dim lblnDemandInd As Boolean
    Dim lstrWGDesc As String
    Dim llngProjectID As Long
    Dim lstrProjectDesc As String
    Dim ladrsEmp As ADODB.Recordset
    Dim ladrsQuals As ADODB.Recordset
    Dim llngarrEmps() As Long
    Dim lstrarrEmps As String
    Dim llngarrLineNum() As Long
    Dim lstrarrJobNum() As String
    Dim lstrarrReqJobInd() As String
    Dim i As Integer
    Const mconFIELDATTR = adFldUpdatable Or adFldIsNullable

    Set ladrsWG = GetWorkGroupAttributes(vlngWGNum)
    lblnDemandInd = False
    If mADOExtensions.RecordsExist(ladrsWG) Then
        If ladrsWG!DEMAND_IND = "Y" Then
            lblnDemandInd = True
        End If
        lstrWGDesc = ladrsWG!WG_DESC
        llngProjectID = ladrsWG!PROJECT_ID
        If llngProjectID > 0 Then
            Set lobjCodeTable = CreateObject( _
                                "dbCodeTables.cdCodeTables")
            'Set lobjCodeTable = New dbCodeTables.cdCodeTables
            Set ladrsProjectInfo = lobjCodeTable.GetProjectInfo(llngProjectID)
            Set lobjCodeTable = Nothing

            lstrProjectDesc = ladrsProjectInfo!PROJECT_DESC
        Else
            lstrProjectDesc = ""
        End If
    End If
    Set ladrsWG = Nothing
    Set ladrsProjectInfo = Nothing
    
    Set lobjWG = CreateObject("dbWorkGroup.cdWorkGroup")
    'Set lobjWG = New dbWorkGroup.cdWorkGroup
    Set ladrsEmpsinWG = lobjWG.GetEmployeesForWGSchedRpt(vlngWGNum, vblnUnmannedLines, lblnDemandInd, _
                                    vdtStartDate, vdtEndDate)
    Set lobjWG = Nothing

    'Set lobjdbSched = New dbSchedule.cdSchedule
    Set lobjdbSched = CreateObject("dbSchedule.cdSchedule")
    Set ladrsJobPosition = lobjdbSched.GetJobPosition(vlngWGNum, vdtStartDate, vdtEndDate)
    Set lobjdbSched = Nothing

    If mADOExtensions.RecordsExist(ladrsEmpsinWG) Then
        With ladrsEmpsinWG
            .Sort = "EMP_NUM"
            .Filter = "EMP_NUM <> 0"
            Set ladrsEmpSchedLine = New ADODB.Recordset
            ladrsEmpSchedLine.Fields.Append "EMP_NUM", adInteger, , mconFIELDATTR
            ladrsEmpSchedLine.Fields.Append "EMP_NAME", adVarChar, 256, mconFIELDATTR
            ladrsEmpSchedLine.Fields.Append "CHARGE_DATE", adDate, , mconFIELDATTR
            ladrsEmpSchedLine.Fields.Append "DATE_WORKED", adDate, , mconFIELDATTR
            ladrsEmpSchedLine.Fields.Append "SHIFT_START_TIME", adChar, 5, mconFIELDATTR
            ladrsEmpSchedLine.Fields.Append "SHIFT_START_DATE", adDate, , mconFIELDATTR
            ladrsEmpSchedLine.Fields.Append "SHIFT_END_TIME", adChar, 5, mconFIELDATTR
            ladrsEmpSchedLine.Fields.Append "SHIFT_END_DATE", adDate, , mconFIELDATTR
            ladrsEmpSchedLine.Fields.Append "SHIFT_HOURS", adDecimal, , mconFIELDATTR
            ladrsEmpSchedLine.Fields.Append "START_TIME", adChar, 5, mconFIELDATTR
            ladrsEmpSchedLine.Fields.Append "END_TIME", adChar, 5, mconFIELDATTR
            ladrsEmpSchedLine.Fields.Append "HOURS", adDecimal, , mconFIELDATTR
            ladrsEmpSchedLine.Fields.Append "EXCEPTION_CD", adChar, 3, mconFIELDATTR
            ladrsEmpSchedLine.Fields.Append "OVER_EXC_CD", adChar, 3, mconFIELDATTR
            ladrsEmpSchedLine.Fields.Append "JOB_NUM", adChar, 6, mconFIELDATTR
            ladrsEmpSchedLine.Fields.Append "JOB_DESC", adVarChar, 256, mconFIELDATTR
            ladrsEmpSchedLine.Fields.Append "QUALS", adVarChar, 256, mconFIELDATTR
            ladrsEmpSchedLine.Fields.Append "SKILLS", adVarChar, 256, mconFIELDATTR
            ladrsEmpSchedLine.Fields.Append "JOB_POSITION", adChar, 3, mconFIELDATTR
            ladrsEmpSchedLine.Fields.Append "PROJ_ID", adInteger, , mconFIELDATTR
            ladrsEmpSchedLine.Fields.Append "PROJ_DESC", adVarChar, 256, mconFIELDATTR
            ladrsEmpSchedLine.Fields.Append "COMP_SHIFT_CD", adChar, 3, mconFIELDATTR
            ladrsEmpSchedLine.Fields.Append "SYS_COMP_SHIFT_CD", adChar, 3, mconFIELDATTR
            ladrsEmpSchedLine.Fields.Append "WG_NUM", adInteger, , mconFIELDATTR
            ladrsEmpSchedLine.Fields.Append "WG_DESC", adVarChar, 256, mconFIELDATTR
            ladrsEmpSchedLine.Fields.Append "PROJECT_ID", adInteger, , mconFIELDATTR
            ladrsEmpSchedLine.Fields.Append "CASE_NUM", adInteger, , mconFIELDATTR
            ladrsEmpSchedLine.Fields.Append "FMLA_IND", adChar, 1, mconFIELDATTR
            ladrsEmpSchedLine.Fields.Append "COMP_CD", adChar, 3, mconFIELDATTR
            ladrsEmpSchedLine.Fields.Append "REL_CD", adChar, 3, mconFIELDATTR
            ladrsEmpSchedLine.Fields.Append "SYSTEM_EXC_CD", adChar, 3, mconFIELDATTR
            ladrsEmpSchedLine.Fields.Append "SHIFT_CODE", adChar, 1, mconFIELDATTR
            ladrsEmpSchedLine.Fields.Append "SCHED_SRC_CD", adInteger, , mconFIELDATTR
            ladrsEmpSchedLine.Fields.Append "SCHED_CD", adInteger, , mconFIELDATTR
            ladrsEmpSchedLine.Fields.Append "SCHED_CHANGE_PMT", adChar, 1, mconFIELDATTR
            ladrsEmpSchedLine.Fields.Append "OT_OPP_NUM", adInteger, , mconFIELDATTR
            ladrsEmpSchedLine.Fields.Append "TRAV_PAY_HOURS", adDecimal, , mconFIELDATTR
            ladrsEmpSchedLine.Fields.Append "LINE_NUM", adInteger, , mconFIELDATTR
            ladrsEmpSchedLine.Fields.Append "REQ_LINE_IND", adChar, 1, mconFIELDATTR
            ladrsEmpSchedLine.Fields.Append "SHIFT_START_END", adChar, 11, mconFIELDATTR
            ladrsEmpSchedLine.Fields.Append "SCHED_CD_FOR_SHIFT_COMP", adInteger, , mconFIELDATTR
            ladrsEmpSchedLine.Fields.Append "BASE_SCHED_SRC_CD", adInteger, , mconFIELDATTR
            ladrsEmpSchedLine.Fields.Append "LINE_SRC_CD", adInteger, , mconFIELDATTR
            ladrsEmpSchedLine.Fields.Append "SHIFT_NUM", adInteger, , mconFIELDATTR
            ladrsEmpSchedLine.Fields.Append "SHIFT_LUNCH_START_TIME", adChar, 5, mconFIELDATTR
            ladrsEmpSchedLine.Fields.Append "SHIFT_LUNCH_END_TIME", adChar, 5, mconFIELDATTR
            ladrsEmpSchedLine.Fields.Append "TRADE_NUM", adInteger, , mconFIELDATTR
            ladrsEmpSchedLine.Fields.Append "WG_LINE_IND", adChar, 1, mconFIELDATTR
            ladrsEmpSchedLine.Fields.Append "FULL_SHIFT_EXC_IND", adChar, 1, mconFIELDATTR
            ladrsEmpSchedLine.Fields.Append "DOUBLE_SHIFT_IND", adChar, 1, mconFIELDATTR
            ladrsEmpSchedLine.Fields.Append "LAST_MOD_SCHED_TM", adChar, 26, mconFIELDATTR
            ladrsEmpSchedLine.Fields.Append "LINE_WG", adInteger, , mconFIELDATTR
            ladrsEmpSchedLine.Fields.Append "LAST_MOD_EMP_NUM", adInteger, , mconFIELDATTR
            ladrsEmpSchedLine.Fields.Append "LAST_MOD_TIMESTAMP", adChar, 26, mconFIELDATTR
            ladrsEmpSchedLine.Fields.Append "LINE_JOB_NUM", adChar, 6, mconFIELDATTR
            ladrsEmpSchedLine.Fields.Append "SCHED_CHG_EXC_CD", adChar, 3, mconFIELDATTR
            ladrsEmpSchedLine.Fields.Append "ROT_CD", adInteger, , mconFIELDATTR
            ladrsEmpSchedLine.Fields.Append "ROT_WEEK", adInteger, , mconFIELDATTR
            ladrsEmpSchedLine.CursorLocation = adUseClient
            ladrsEmpSchedLine.CursorType = adOpenStatic
            ladrsEmpSchedLine.LockType = adLockBatchOptimistic
            ladrsEmpSchedLine.Open
        End With
    End If
    
    If mADOExtensions.RecordsExist(ladrsEmpsinWG) Then
        With ladrsEmpsinWG
            ReDim llngarrEmps(1 To .RecordCount)
            ReDim llngarrLineNum(1 To .RecordCount)
            ReDim lstrarrJobNum(1 To .RecordCount)
            ReDim lstrarrReqJobInd(1 To .RecordCount)
            '.Sort = "EMP_NUM"
            '.Filter = "EMP_NUM <> 0"
            .MoveFirst
            llngEmpCount = 1
            For llngIndex = 1 To .RecordCount
                llngarrEmps(llngEmpCount) = !EMP_NUM
                If llngIndex = .RecordCount Then
                    lstrarrEmps = lstrarrEmps & !EMP_NUM
                Else
                    lstrarrEmps = lstrarrEmps & !EMP_NUM & ","
                End If
                If lblnDemandInd Then
                    llngarrLineNum(llngEmpCount) = !LINE_NUM
                    lstrarrReqJobInd(llngEmpCount) = !REQ_JOB_IND
                    If Trim(!JOB_NUM) <> "" Then
                        lstrarrJobNum(llngEmpCount) = !JOB_NUM
                    Else
                        lstrarrJobNum(llngEmpCount) = ""
                    End If
                Else
                    llngarrLineNum(llngEmpCount) = 0
                    lstrarrJobNum(llngEmpCount) = ""
                    lstrarrReqJobInd(llngEmpCount) = ""
                End If
                llngEmpCount = llngEmpCount + 1
                .MoveNext
            Next llngIndex
        End With
        'Set ladrsEmpSchedLine = New ADODB.Recordset
        'With ladrsEmpSchedLine
            '.Fields.Append "EMP_NUM", adInteger, , mconFIELDATTR
            '.Fields.Append "EMP_NAME", adVarChar, 256, mconFIELDATTR
            '.Fields.Append "CHARGE_DATE", adDate, , mconFIELDATTR
            '.Fields.Append "DATE_WORKED", adDate, , mconFIELDATTR
            '.Fields.Append "SHIFT_START_TIME", adChar, 5, mconFIELDATTR
            '.Fields.Append "SHIFT_START_DATE", adDate, , mconFIELDATTR
            '.Fields.Append "SHIFT_END_TIME", adChar, 5, mconFIELDATTR
            '.Fields.Append "SHIFT_END_DATE", adDate, , mconFIELDATTR
            '.Fields.Append "SHIFT_HOURS", adDecimal, , mconFIELDATTR
            '.Fields.Append "START_TIME", adChar, 5, mconFIELDATTR
            '.Fields.Append "END_TIME", adChar, 5, mconFIELDATTR
            '.Fields.Append "HOURS", adDecimal, , mconFIELDATTR
            '.Fields.Append "EXCEPTION_CD", adChar, 3, mconFIELDATTR
            '.Fields.Append "OVER_EXC_CD", adChar, 3, mconFIELDATTR
            '.Fields.Append "JOB_NUM", adChar, 6, mconFIELDATTR
            '.Fields.Append "JOB_DESC", adVarChar, 256, mconFIELDATTR
            '.Fields.Append "QUALS", adVarChar, 256, mconFIELDATTR
            '.Fields.Append "SKILLS", adVarChar, 256, mconFIELDATTR
            '.Fields.Append "JOB_POSITION", adChar, 3, mconFIELDATTR
            '.Fields.Append "PROJ_ID", adInteger, , mconFIELDATTR
            '.Fields.Append "PROJ_DESC", adVarChar, 256, mconFIELDATTR
            '.Fields.Append "COMP_SHIFT_CD", adChar, 3, mconFIELDATTR
            '.Fields.Append "SYS_COMP_SHIFT_CD", adChar, 3, mconFIELDATTR
            '.Fields.Append "WG_NUM", adInteger, , mconFIELDATTR
            '.Fields.Append "WG_DESC", adVarChar, 256, mconFIELDATTR
            '.Fields.Append "PROJECT_ID", adInteger, , mconFIELDATTR
            '.Fields.Append "CASE_NUM", adInteger, , mconFIELDATTR
            '.Fields.Append "FMLA_IND", adChar, 1, mconFIELDATTR
            '.Fields.Append "COMP_CD", adChar, 3, mconFIELDATTR
            '.Fields.Append "REL_CD", adChar, 3, mconFIELDATTR
            '.Fields.Append "SYSTEM_EXC_CD", adChar, 3, mconFIELDATTR
            '.Fields.Append "SHIFT_CODE", adChar, 1, mconFIELDATTR
            '.Fields.Append "SCHED_SRC_CD", adInteger, , mconFIELDATTR
            '.Fields.Append "SCHED_CD", adInteger, , mconFIELDATTR
            '.Fields.Append "SCHED_CHANGE_PMT", adChar, 1, mconFIELDATTR
            '.Fields.Append "OT_OPP_NUM", adInteger, , mconFIELDATTR
            '.Fields.Append "TRAV_PAY_HOURS", adDecimal, , mconFIELDATTR
            '.Fields.Append "LINE_NUM", adInteger, , mconFIELDATTR
            '.Fields.Append "REQ_LINE_IND", adChar, 1, mconFIELDATTR
            '.Fields.Append "SHIFT_START_END", adChar, 11, mconFIELDATTR
            '.Fields.Append "SCHED_CD_FOR_SHIFT_COMP", adInteger, , mconFIELDATTR
            '.Fields.Append "BASE_SCHED_SRC_CD", adInteger, , mconFIELDATTR
            '.Fields.Append "LINE_SRC_CD", adInteger, , mconFIELDATTR
            '.Fields.Append "SHIFT_NUM", adInteger, , mconFIELDATTR
            '.Fields.Append "SHIFT_LUNCH_START_TIME", adChar, 5, mconFIELDATTR
            '.Fields.Append "SHIFT_LUNCH_END_TIME", adChar, 5, mconFIELDATTR
            '.Fields.Append "TRADE_NUM", adInteger, , mconFIELDATTR
            '.Fields.Append "WG_LINE_IND", adChar, 1, mconFIELDATTR
            '.Fields.Append "FULL_SHIFT_EXC_IND", adChar, 1, mconFIELDATTR
            '.Fields.Append "DOUBLE_SHIFT_IND", adChar, 1, mconFIELDATTR
            '.Fields.Append "LAST_MOD_SCHED_TM", adChar, 26, mconFIELDATTR
            '.Fields.Append "LINE_WG", adInteger, , mconFIELDATTR
            '.Fields.Append "LAST_MOD_EMP_NUM", adInteger, , mconFIELDATTR
            '.Fields.Append "LAST_MOD_TIMESTAMP", adChar, 26, mconFIELDATTR
            '.Fields.Append "LINE_JOB_NUM", adChar, 6, mconFIELDATTR
            '.Fields.Append "SCHED_CHG_EXC_CD", adChar, 3, mconFIELDATTR
            '.Fields.Append "ROT_CD", adInteger, , mconFIELDATTR
            '.Fields.Append "ROT_WEEK", adInteger, , mconFIELDATTR
            '.CursorLocation = adUseClient
            '.CursorType = adOpenStatic
            '.LockType = adLockBatchOptimistic
            '.Open
        'End With
        
        If llngProjectID > 0 Then
            Set lobjEmp = CreateObject("dbEmployee.cdEmployee")
            'Set lobjEmp = New dbEmployee.cdEmployee
            'change to getqualsandskillsforemployees2
            Set ladrsQuals = lobjEmp.GetQualsAndSkillsForEmployees2(lstrarrEmps _
                                , vdtStartDate, vdtEndDate)
            Set lobjEmp = Nothing
        Else
            Set ladrsQuals = Nothing
        End If
        
        Set lobjEmployee = CreateObject("busEmployee.cbsEmployee")
        'Set lobjEmployee = New busEmployee.cbsEmployee
        Set ladrsEmp = lobjEmployee.GetEmployee(llngarrEmps)
        Set lobjEmployee = Nothing

        Set lobjWG = CreateObject("dbWorkGroup.cdWorkGroup")
        'Set lobjWG = New dbWorkGroup.cdWorkGroup
        Set ladrsEmpWG = lobjWG.GetEmployeeWorkGroups(llngarrEmps, vdtStartDate, vdtEndDate)
        Set lobjWG = Nothing

        llngPrevActimEmp = 0
        ldtPrevActimDate = "9999-12-31"
        
        Set lobjTimeCard = CreateObject("dbTimeCard.cdTimeCard")
        'Set lobjTimeCard = New dbTimeCard.cdTimeCard
        Set ladrsActualTime = lobjTimeCard.GetActualTimeEntriesWithoutOtherPay( _
                              llngarrEmps, vdtStartDate, vdtEndDate)
        Set lobjTimeCard = Nothing

        Set lobjSched = CreateObject("busSchedule.cbsSchedule")
        'Set lobjSched = New busSchedule.cbsSchedule

        For llngIndex = 1 To llngEmpCount - 1
            Set ladrsEmpSched = lobjSched.GetAdjustedBaseSchedule(llngarrEmps(llngIndex), vdtStartDate _
                              , vdtEndDate, ladrsDummy, False, trAdjBaseSchedWithExceptionsAndOvertime _
                              , trAllScheduleLevels, trRelief, trAllExceptions)
            Set ladrsDummy = Nothing

            With ladrsEmpSched
                .MoveFirst
                While Not ladrsEmpSched.EOF
                    If lblnDemandInd Then
                        If (llngarrLineNum(llngIndex)) = 0 Or (!LINE_NUM = llngarrLineNum(llngIndex)) Then
                            ladrsEmpSchedLine.AddNew
                            ladrsEmpSchedLine.Fields("EMP_NUM") = !EMP_NUM
                            If mADOExtensions.RecordsExist(ladrsEmp) Then
                                ladrsEmp.MoveFirst
                                ladrsEmp.Find "EMP_NUM = " & !EMP_NUM
                                If Not (ladrsEmp.EOF) Then
                                    ladrsEmpSchedLine.Fields("EMP_NAME") = ladrsEmp!EMP_NAME
                                End If
                            End If
                            ladrsEmpSchedLine.Fields("CHARGE_DATE") = !CHARGE_DATE
                            ladrsEmpSchedLine.Fields("DATE_WORKED") = !DATE_WORKED
                            ladrsEmpSchedLine.Fields("SHIFT_START_TIME") = !SHIFT_START_TIME
                            ladrsEmpSchedLine.Fields("SHIFT_START_DATE") = !SHIFT_START_DATE
                            ladrsEmpSchedLine.Fields("SHIFT_END_TIME") = !SHIFT_END_TIME
                            ladrsEmpSchedLine.Fields("SHIFT_END_DATE") = !SHIFT_END_DATE
                            ladrsEmpSchedLine.Fields("SHIFT_HOURS") = !SHIFT_HOURS
                            ladrsEmpSchedLine.Fields("START_TIME") = !START_TIME
                            ladrsEmpSchedLine.Fields("END_TIME") = !END_TIME
                            ladrsEmpSchedLine.Fields("HOURS") = !HOURS
                            ladrsEmpSchedLine.Fields("EXCEPTION_CD") = !EXCEPTION_CD
                            ladrsEmpSchedLine.Fields("OVER_EXC_CD") = !OVER_EXC_CD
                            If Trim(lstrarrJobNum(llngIndex)) <> "" Then
                                ladrsEmpSchedLine.Fields("JOB_NUM") = lstrarrJobNum(llngIndex)
                            Else
                                If Trim(!SCHED_JOB_NUM) = "" Then
                                    ladrsEmpSchedLine.Fields("JOB_NUM") = !EMP_JOB_NUM
                                Else
                                    ladrsEmpSchedLine.Fields("JOB_NUM") = !SCHED_JOB_NUM
                                End If
                            End If
                            ladrsEmpSchedLine.Fields("JOB_DESC") = GetJobDesc(ladrsEmpSchedLine.Fields("JOB_NUM"))
                            If llngProjectID > 0 Then
                                If mADOExtensions.RecordsExist(ladrsQuals) Then
                                    ladrsQuals.MoveFirst
                                    ladrsQuals.Filter = "EMP_NUM = " & !EMP_NUM & " AND EFF_START_DATE <= '" & vdtStartDate & _
                                                        "' AND EFF_END_DATE >= '" & vdtEndDate & "'"
                                    If Not ladrsQuals.EOF Then
                                        ladrsQuals.Sort = "QUAL_CD, SPEC_SKILL_CD"
                                    End If
                                    For i = 1 To ladrsQuals.RecordCount
                                        If i > 5 Then Exit For
                                        If i = 1 Then
                                            ladrsEmpSchedLine.Fields("QUALS") = Trim(ladrsQuals!QUAL_CD)
                                            ladrsEmpSchedLine.Fields("SKILLS") = Trim(ladrsQuals!SPEC_SKILL_CD)
                                        Else
                                            ladrsEmpSchedLine.Fields("QUALS") = _
                                            Trim(ladrsEmpSchedLine.Fields("QUALS")) & " " & Trim(ladrsQuals!QUAL_CD)
                                            ladrsEmpSchedLine.Fields("SKILLS") = _
                                            Trim(ladrsEmpSchedLine.Fields("SKILLS")) & " " & Trim(ladrsQuals!SPEC_SKILL_CD)
                                        End If
                                        ladrsQuals.MoveNext
                                    Next i
                                    ladrsQuals.Filter = adFilterNone
                                Else
                                    ladrsEmpSchedLine.Fields("QUALS") = ""
                                    ladrsEmpSchedLine.Fields("SKILLS") = ""
                                End If
                            End If
                            If mADOExtensions.RecordsExist(ladrsJobPosition) Then
                                ladrsJobPosition.MoveFirst
                                ladrsJobPosition.Filter = "CHARGE_DATE = '" & !CHARGE_DATE & "' AND LINE_NUM = " & !LINE_NUM
                                If Not ladrsJobPosition.EOF Then
                                    ladrsEmpSchedLine.Fields("JOB_POSITION") = ladrsJobPosition.Fields("JOB_POSITION")
                                Else
                                    ladrsEmpSchedLine.Fields("JOB_POSITION") = ""
                                End If
                            Else
                                ladrsEmpSchedLine.Fields("JOB_POSITION") = ""
                            End If
                            ladrsJobPosition.Filter = adFilterNone
                            ladrsEmpSchedLine.Fields("PROJ_ID") = llngProjectID
                            ladrsEmpSchedLine.Fields("PROJ_DESC") = lstrProjectDesc
                            ladrsEmpSchedLine.Fields("COMP_SHIFT_CD") = !COMP_SHIFT_CD
                            ladrsEmpSchedLine.Fields("SYS_COMP_SHIFT_CD") = !SYS_COMP_SHIFT_CD
                            If !LINE_NUM = 0 Then
                                If mADOExtensions.RecordsExist(ladrsEmpWG) Then
                                    ladrsEmpWG.MoveFirst
                                    ladrsEmpWG.Filter = "EMP_NUM = " & !EMP_NUM & " AND EFF_START_DATE <= #" & !CHARGE_DATE & "#" & _
                                    " AND EFF_END_DATE >= #" & !CHARGE_DATE & "#"
                                    If Not (ladrsEmpWG.EOF) Then
                                        ladrsEmpSchedLine.Fields("WG_NUM") = ladrsEmpWG!WG_NUM
                                    Else
                                        ladrsEmpSchedLine.Fields("WG_NUM") = !WG_NUM
                                    End If
                                    ladrsEmpWG.Filter = adFilterNone
                                Else
                                    ladrsEmpSchedLine.Fields("WG_NUM") = !WG_NUM
                                End If
                            Else
                                ladrsEmpSchedLine.Fields("WG_NUM") = !WG_NUM
                            End If
                            ladrsEmpSchedLine.Fields("WG_DESC") = lstrWGDesc
                            ladrsEmpSchedLine.Fields("CASE_NUM") = !CASE_NUM
                            ladrsEmpSchedLine.Fields("FMLA_IND") = !FMLA_IND
                            ladrsEmpSchedLine.Fields("COMP_CD") = !COMP_CD
                            ladrsEmpSchedLine.Fields("REL_CD") = !REL_CD
                            ladrsEmpSchedLine.Fields("SYSTEM_EXC_CD") = !SYSTEM_EXC_CD
                            ladrsEmpSchedLine.Fields("SHIFT_CODE") = !SHIFT_CODE
                            ladrsEmpSchedLine.Fields("SCHED_SRC_CD") = !SCHED_SRC_CD
                            ladrsEmpSchedLine.Fields("SCHED_CD") = !SCHED_CD
                            ladrsEmpSchedLine.Fields("SCHED_CHANGE_PMT") = !SCHED_CHANGE_PMT
                            ladrsEmpSchedLine.Fields("OT_OPP_NUM") = !OT_OPP_NUM
                            ladrsEmpSchedLine.Fields("TRAV_PAY_HOURS") = !TRAV_PAY_HOURS
                            ladrsEmpSchedLine.Fields("LINE_NUM") = !LINE_NUM
                            ladrsEmpSchedLine.Fields("REQ_LINE_IND") = lstrarrReqJobInd(llngIndex)
                            ladrsEmpSchedLine.Fields("SHIFT_START_END") = !SHIFT_START_END
                            ladrsEmpSchedLine.Fields("SCHED_CD_FOR_SHIFT_COMP") = !SCHED_CD_FOR_SHIFT_COMP
                            ladrsEmpSchedLine.Fields("BASE_SCHED_SRC_CD") = !BASE_SCHED_SRC_CD
                            ladrsEmpSchedLine.Fields("LINE_SRC_CD") = !LINE_SRC_CD
                            ladrsEmpSchedLine.Fields("SHIFT_NUM") = !SHIFT_NUM
                            ladrsEmpSchedLine.Fields("SHIFT_LUNCH_START_TIME") = !SHIFT_LUNCH_START_TIME
                            ladrsEmpSchedLine.Fields("SHIFT_LUNCH_END_TIME") = !SHIFT_LUNCH_END_TIME
                            ladrsEmpSchedLine.Fields("TRADE_NUM") = !TRADE_NUM
                            ladrsEmpSchedLine.Fields("WG_LINE_IND") = "Y"
                            ladrsEmpSchedLine.Fields("FULL_SHIFT_EXC_IND") = !FULL_SHIFT_EXC_IND
                            ladrsEmpSchedLine.Fields("DOUBLE_SHIFT_IND") = !DOUBLE_SHIFT_IND
                            ladrsEmpSchedLine.Fields("LAST_MOD_SCHED_TM") = !LAST_MOD_SCHED_TM
                            ladrsEmpSchedLine.Fields("LINE_WG") = !LINE_WG
                            ladrsEmpSchedLine.Fields("LAST_MOD_EMP_NUM") = !LAST_MOD_EMP_NUM
                            ladrsEmpSchedLine.Fields("LAST_MOD_TIMESTAMP") = !LAST_MOD_TIMESTAMP
                            ladrsEmpSchedLine.Fields("LINE_JOB_NUM") = !LINE_JOB_NUM
                            ladrsEmpSchedLine.Fields("SCHED_CHG_EXC_CD") = !SCHED_CHG_EXC_CD
                            ladrsEmpSchedLine.Fields("ROT_CD") = !ROT_CD
                            ladrsEmpSchedLine.Fields("ROT_WEEK") = !ROT_WEEK
                            ladrsEmpSchedLine.Update
                        End If
                    Else
                        ladrsEmpSchedLine.AddNew
                        ladrsEmpSchedLine.Fields("EMP_NUM") = !EMP_NUM
                        If mADOExtensions.RecordsExist(ladrsEmp) Then
                            ladrsEmp.MoveFirst
                            ladrsEmp.Find "EMP_NUM = " & !EMP_NUM
                            If Not (ladrsEmp.EOF) Then
                                ladrsEmpSchedLine.Fields("EMP_NAME") = ladrsEmp!EMP_NAME
                            End If
                        End If
                        ladrsEmpSchedLine.Fields("CHARGE_DATE") = !CHARGE_DATE
                        ladrsEmpSchedLine.Fields("DATE_WORKED") = !DATE_WORKED
                        ladrsEmpSchedLine.Fields("SHIFT_START_TIME") = !SHIFT_START_TIME
                        ladrsEmpSchedLine.Fields("SHIFT_START_DATE") = !SHIFT_START_DATE
                        ladrsEmpSchedLine.Fields("SHIFT_END_TIME") = !SHIFT_END_TIME
                        ladrsEmpSchedLine.Fields("SHIFT_END_DATE") = !SHIFT_END_DATE
                        ladrsEmpSchedLine.Fields("SHIFT_HOURS") = !SHIFT_HOURS
                        ladrsEmpSchedLine.Fields("START_TIME") = !START_TIME
                        ladrsEmpSchedLine.Fields("END_TIME") = !END_TIME
                        ladrsEmpSchedLine.Fields("HOURS") = !HOURS
                        ladrsEmpSchedLine.Fields("EXCEPTION_CD") = !EXCEPTION_CD
                        ladrsEmpSchedLine.Fields("OVER_EXC_CD") = !OVER_EXC_CD
                        If Trim(!SCHED_JOB_NUM) = "" Then
                            ladrsEmpSchedLine.Fields("JOB_NUM") = !EMP_JOB_NUM
                        Else
                            ladrsEmpSchedLine.Fields("JOB_NUM") = !SCHED_JOB_NUM
                        End If
                        ladrsEmpSchedLine.Fields("JOB_DESC") = GetJobDesc(ladrsEmpSchedLine.Fields("JOB_NUM"))
                        If llngProjectID > 0 Then
                            If mADOExtensions.RecordsExist(ladrsQuals) Then
                                ladrsQuals.MoveFirst
                                ladrsQuals.Filter = "EMP_NUM = " & !EMP_NUM & " AND EFF_START_DATE <= '" & vdtStartDate & _
                                                    "' AND EFF_END_DATE >= '" & vdtEndDate & "'"
                                If Not ladrsQuals.EOF Then
                                        ladrsQuals.Sort = "QUAL_CD, SPEC_SKILL_CD"
                                End If
                                For i = 1 To ladrsQuals.RecordCount
                                    If i > 5 Then Exit For
                                    If i = 1 Then
                                        ladrsEmpSchedLine.Fields("QUALS") = Trim(ladrsQuals!QUAL_CD)
                                        ladrsEmpSchedLine.Fields("SKILLS") = Trim(ladrsQuals!SPEC_SKILL_CD)
                                    Else
                                        ladrsEmpSchedLine.Fields("QUALS") = _
                                        ladrsEmpSchedLine.Fields("QUALS") & " " & Trim(ladrsQuals!QUAL_CD)
                                        ladrsEmpSchedLine.Fields("SKILLS") = _
                                        ladrsEmpSchedLine.Fields("SKILLS") & " " & Trim(ladrsQuals!SPEC_SKILL_CD)
                                    End If
                                    ladrsQuals.MoveNext
                                Next i
                                ladrsQuals.Filter = adFilterNone
                            Else
                                ladrsEmpSchedLine.Fields("QUALS") = ""
                                ladrsEmpSchedLine.Fields("SKILLS") = ""
                            End If
                        End If
                        ladrsEmpSchedLine.Fields("JOB_POSITION") = ""
                        ladrsEmpSchedLine.Fields("PROJ_ID") = llngProjectID
                        ladrsEmpSchedLine.Fields("PROJ_DESC") = lstrProjectDesc
                        ladrsEmpSchedLine.Fields("COMP_SHIFT_CD") = !COMP_SHIFT_CD
                        ladrsEmpSchedLine.Fields("SYS_COMP_SHIFT_CD") = !SYS_COMP_SHIFT_CD
                        If mADOExtensions.RecordsExist(ladrsEmpWG) Then
                            ladrsEmpWG.MoveFirst
                            ladrsEmpWG.Filter = "EMP_NUM = " & !EMP_NUM & " AND EFF_START_DATE <= #" & !CHARGE_DATE & "#" & _
                            " AND EFF_END_DATE >= #" & !CHARGE_DATE & "#"
                            If Not (ladrsEmpWG.EOF) Then
                                ladrsEmpSchedLine.Fields("WG_NUM") = ladrsEmpWG!WG_NUM
                            Else
                                ladrsEmpSchedLine.Fields("WG_NUM") = !WG_NUM
                            End If
                            ladrsEmpWG.Filter = adFilterNone
                        Else
                            ladrsEmpSchedLine.Fields("WG_NUM") = !WG_NUM
                        End If
                        ladrsEmpSchedLine.Fields("WG_DESC") = lstrWGDesc
                        ladrsEmpSchedLine.Fields("CASE_NUM") = !CASE_NUM
                        ladrsEmpSchedLine.Fields("FMLA_IND") = !FMLA_IND
                        ladrsEmpSchedLine.Fields("COMP_CD") = !COMP_CD
                        ladrsEmpSchedLine.Fields("REL_CD") = !REL_CD
                        ladrsEmpSchedLine.Fields("SYSTEM_EXC_CD") = !SYSTEM_EXC_CD
                        ladrsEmpSchedLine.Fields("SHIFT_CODE") = !SHIFT_CODE
                        ladrsEmpSchedLine.Fields("SCHED_SRC_CD") = !SCHED_SRC_CD
                        ladrsEmpSchedLine.Fields("SCHED_CD") = !SCHED_CD
                        ladrsEmpSchedLine.Fields("SCHED_CHANGE_PMT") = !SCHED_CHANGE_PMT
                        ladrsEmpSchedLine.Fields("OT_OPP_NUM") = !OT_OPP_NUM
                        ladrsEmpSchedLine.Fields("TRAV_PAY_HOURS") = !TRAV_PAY_HOURS
                        ladrsEmpSchedLine.Fields("LINE_NUM") = !LINE_NUM
                        ladrsEmpSchedLine.Fields("REQ_LINE_IND") = !REQ_LINE_IND
                        ladrsEmpSchedLine.Fields("SHIFT_START_END") = !SHIFT_START_END
                        ladrsEmpSchedLine.Fields("SCHED_CD_FOR_SHIFT_COMP") = !SCHED_CD_FOR_SHIFT_COMP
                        ladrsEmpSchedLine.Fields("BASE_SCHED_SRC_CD") = !BASE_SCHED_SRC_CD
                        ladrsEmpSchedLine.Fields("LINE_SRC_CD") = !LINE_SRC_CD
                        ladrsEmpSchedLine.Fields("SHIFT_NUM") = !SHIFT_NUM
                        ladrsEmpSchedLine.Fields("SHIFT_LUNCH_START_TIME") = !SHIFT_LUNCH_START_TIME
                        ladrsEmpSchedLine.Fields("SHIFT_LUNCH_END_TIME") = !SHIFT_LUNCH_END_TIME
                        ladrsEmpSchedLine.Fields("TRADE_NUM") = !TRADE_NUM
                        ladrsEmpSchedLine.Fields("WG_LINE_IND") = "N"
                        ladrsEmpSchedLine.Fields("FULL_SHIFT_EXC_IND") = !FULL_SHIFT_EXC_IND
                        ladrsEmpSchedLine.Fields("DOUBLE_SHIFT_IND") = !DOUBLE_SHIFT_IND
                        ladrsEmpSchedLine.Fields("LAST_MOD_SCHED_TM") = !LAST_MOD_SCHED_TM
                        ladrsEmpSchedLine.Fields("LINE_WG") = !LINE_WG
                        ladrsEmpSchedLine.Fields("LAST_MOD_EMP_NUM") = !LAST_MOD_EMP_NUM
                        ladrsEmpSchedLine.Fields("LAST_MOD_TIMESTAMP") = !LAST_MOD_TIMESTAMP
                        ladrsEmpSchedLine.Fields("LINE_JOB_NUM") = !LINE_JOB_NUM
                        ladrsEmpSchedLine.Fields("SCHED_CHG_EXC_CD") = !SCHED_CHG_EXC_CD
                        ladrsEmpSchedLine.Fields("ROT_CD") = !ROT_CD
                        ladrsEmpSchedLine.Fields("ROT_WEEK") = !ROT_WEEK
                        ladrsEmpSchedLine.Update
                    End If
                    'get actual time
                    If ladrsEmpSched!EMP_NUM <> llngPrevActimEmp Or _
                       ladrsEmpSched!CHARGE_DATE <> ldtPrevActimDate Then
                        ladrsActualTime.Filter = adFilterNone
                        ladrsActualTime.Filter = "EMP_NUM = " & ladrsEmpSched!EMP_NUM & _
                        " AND CHARGE_DATE = #" & ladrsEmpSched!CHARGE_DATE & "#"
                        While Not ladrsActualTime.EOF
                            If Not ladrsActualTime.EOF Then
                                If lblnDemandInd Then
                                    If (llngarrLineNum(llngIndex)) = 0 Or (!LINE_NUM = llngarrLineNum(llngIndex)) Then
                                        ladrsEmpSchedLine.AddNew
                                        ladrsEmpSchedLine.Fields("EMP_NUM") = !EMP_NUM
                                        If mADOExtensions.RecordsExist(ladrsEmp) Then
                                            ladrsEmp.MoveFirst
                                            ladrsEmp.Find "EMP_NUM = " & !EMP_NUM
                                            If Not (ladrsEmp.EOF) Then
                                                ladrsEmpSchedLine.Fields("EMP_NAME") = ladrsEmp!EMP_NAME
                                            End If
                                        End If
                                        ladrsEmpSchedLine.Fields("CHARGE_DATE") = ladrsActualTime!CHARGE_DATE
                                        ladrsEmpSchedLine.Fields("DATE_WORKED") = ladrsActualTime!DATE_WORKED
                                        ladrsEmpSchedLine.Fields("SHIFT_START_TIME") = !SHIFT_START_TIME
                                        ladrsEmpSchedLine.Fields("SHIFT_START_DATE") = !SHIFT_START_DATE
                                        ladrsEmpSchedLine.Fields("SHIFT_END_TIME") = !SHIFT_END_TIME
                                        ladrsEmpSchedLine.Fields("SHIFT_END_DATE") = !SHIFT_END_DATE
                                        ladrsEmpSchedLine.Fields("SHIFT_HOURS") = !SHIFT_HOURS
                                        ladrsEmpSchedLine.Fields("START_TIME") = ladrsActualTime!START_TIME
                                        ladrsEmpSchedLine.Fields("END_TIME") = ladrsActualTime!END_TIME
                                        ladrsEmpSchedLine.Fields("HOURS") = ladrsActualTime!HOURS
                                        ladrsEmpSchedLine.Fields("EXCEPTION_CD") = ladrsActualTime!EXCEPTION_CD
                                        ladrsEmpSchedLine.Fields("OVER_EXC_CD") = ladrsActualTime!OVER_EXC_CD
                                        If Trim(lstrarrJobNum(llngIndex)) <> "" Then
                                            ladrsEmpSchedLine.Fields("JOB_NUM") = lstrarrJobNum(llngIndex)
                                        Else
                                            If Trim(!SCHED_JOB_NUM) = "" Then
                                                ladrsEmpSchedLine.Fields("JOB_NUM") = !EMP_JOB_NUM
                                            Else
                                                ladrsEmpSchedLine.Fields("JOB_NUM") = !SCHED_JOB_NUM
                                            End If
                                        End If
                                        ladrsEmpSchedLine.Fields("JOB_DESC") = GetJobDesc(ladrsEmpSchedLine.Fields("JOB_NUM"))
                                        If llngProjectID > 0 Then
                                            If mADOExtensions.RecordsExist(ladrsQuals) Then
                                                ladrsQuals.MoveFirst
                                                ladrsQuals.Filter = "EMP_NUM = " & !EMP_NUM & " AND EFF_START_DATE <= '" & vdtStartDate & _
                                                                    "' AND EFF_END_DATE >= '" & vdtEndDate & "'"
                                                If Not ladrsQuals.EOF Then
                                                    ladrsQuals.Sort = "QUAL_CD, SPEC_SKILL_CD"
                                                End If
                                                For i = 1 To ladrsQuals.RecordCount
                                                    If i > 5 Then Exit For
                                                    If i = 1 Then
                                                        ladrsEmpSchedLine.Fields("QUALS") = Trim(ladrsQuals!QUAL_CD)
                                                        ladrsEmpSchedLine.Fields("SKILLS") = Trim(ladrsQuals!SPEC_SKILL_CD)
                                                    Else
                                                        ladrsEmpSchedLine.Fields("QUALS") = _
                                                        ladrsEmpSchedLine.Fields("QUALS") & " " & Trim(ladrsQuals!QUAL_CD)
                                                        ladrsEmpSchedLine.Fields("SKILLS") = _
                                                        ladrsEmpSchedLine.Fields("SKILLS") & " " & Trim(ladrsQuals!SPEC_SKILL_CD)
                                                    End If
                                                    ladrsQuals.MoveNext
                                                Next i
                                                ladrsQuals.Filter = adFilterNone
                                            Else
                                                ladrsEmpSchedLine.Fields("QUALS") = ""
                                                ladrsEmpSchedLine.Fields("SKILLS") = ""
                                            End If
                                        End If
                                        If mADOExtensions.RecordsExist(ladrsJobPosition) Then
                                            ladrsJobPosition.MoveFirst
                                            ladrsJobPosition.Filter = "CHARGE_DATE = '" & !CHARGE_DATE & "' AND LINE_NUM = " & !LINE_NUM
                                            If Not ladrsJobPosition.EOF Then
                                                ladrsEmpSchedLine.Fields("JOB_POSITION") = ladrsJobPosition.Fields("JOB_POSITION")
                                            Else
                                                ladrsEmpSchedLine.Fields("JOB_POSITION") = ""
                                            End If
                                        Else
                                            ladrsEmpSchedLine.Fields("JOB_POSITION") = ""
                                        End If
                                        ladrsJobPosition.Filter = adFilterNone
                                        ladrsEmpSchedLine.Fields("PROJ_ID") = llngProjectID
                                        ladrsEmpSchedLine.Fields("PROJ_DESC") = lstrProjectDesc
                                        ladrsEmpSchedLine.Fields("COMP_SHIFT_CD") = ladrsActualTime!COMP_SHIFT_CD
                                        ladrsEmpSchedLine.Fields("SYS_COMP_SHIFT_CD") = !SYS_COMP_SHIFT_CD
                                        ladrsEmpSchedLine.Fields("WG_NUM") = !WG_NUM
                                        ladrsEmpSchedLine.Fields("WG_DESC") = lstrWGDesc
                                        ladrsEmpSchedLine.Fields("CASE_NUM") = !CASE_NUM
                                        ladrsEmpSchedLine.Fields("FMLA_IND") = !FMLA_IND
                                        ladrsEmpSchedLine.Fields("COMP_CD") = ladrsActualTime!COMP_CD
                                        ladrsEmpSchedLine.Fields("REL_CD") = ladrsActualTime!REL_CD
                                        ladrsEmpSchedLine.Fields("SYSTEM_EXC_CD") = !SYSTEM_EXC_CD
                                        ladrsEmpSchedLine.Fields("SHIFT_CODE") = !SHIFT_CODE
                                        ladrsEmpSchedLine.Fields("SCHED_SRC_CD") = !SCHED_SRC_CD
                                        ladrsEmpSchedLine.Fields("SCHED_CD") = !SCHED_CD
                                        ladrsEmpSchedLine.Fields("SCHED_CHANGE_PMT") = !SCHED_CHANGE_PMT
                                        ladrsEmpSchedLine.Fields("OT_OPP_NUM") = !OT_OPP_NUM
                                        ladrsEmpSchedLine.Fields("TRAV_PAY_HOURS") = !TRAV_PAY_HOURS
                                        ladrsEmpSchedLine.Fields("LINE_NUM") = !LINE_NUM
                                        ladrsEmpSchedLine.Fields("REQ_LINE_IND") = lstrarrReqJobInd(llngIndex)
                                        ladrsEmpSchedLine.Fields("SHIFT_START_END") = !SHIFT_START_END
                                        ladrsEmpSchedLine.Fields("SCHED_CD_FOR_SHIFT_COMP") = !SCHED_CD_FOR_SHIFT_COMP
                                        ladrsEmpSchedLine.Fields("BASE_SCHED_SRC_CD") = !BASE_SCHED_SRC_CD
                                        ladrsEmpSchedLine.Fields("LINE_SRC_CD") = !LINE_SRC_CD
                                        ladrsEmpSchedLine.Fields("SHIFT_NUM") = !SHIFT_NUM
                                        ladrsEmpSchedLine.Fields("SHIFT_LUNCH_START_TIME") = !SHIFT_LUNCH_START_TIME
                                        ladrsEmpSchedLine.Fields("SHIFT_LUNCH_END_TIME") = !SHIFT_LUNCH_END_TIME
                                        ladrsEmpSchedLine.Fields("TRADE_NUM") = !TRADE_NUM
                                        ladrsEmpSchedLine.Fields("WG_LINE_IND") = "Y"
                                        ladrsEmpSchedLine.Fields("FULL_SHIFT_EXC_IND") = !FULL_SHIFT_EXC_IND
                                        ladrsEmpSchedLine.Fields("DOUBLE_SHIFT_IND") = !DOUBLE_SHIFT_IND
                                        ladrsEmpSchedLine.Fields("LAST_MOD_SCHED_TM") = !LAST_MOD_SCHED_TM
                                        ladrsEmpSchedLine.Fields("LINE_WG") = !LINE_WG
                                        ladrsEmpSchedLine.Fields("LAST_MOD_EMP_NUM") = !LAST_MOD_EMP_NUM
                                        ladrsEmpSchedLine.Fields("LAST_MOD_TIMESTAMP") = !LAST_MOD_TIMESTAMP
                                        ladrsEmpSchedLine.Fields("LINE_JOB_NUM") = !LINE_JOB_NUM
                                        ladrsEmpSchedLine.Fields("SCHED_CHG_EXC_CD") = !SCHED_CHG_EXC_CD
                                        ladrsEmpSchedLine.Fields("ROT_CD") = !ROT_CD
                                        ladrsEmpSchedLine.Fields("ROT_WEEK") = !ROT_WEEK
                                        ladrsEmpSchedLine.Update
                                    End If
                                Else
                                    ladrsEmpSchedLine.AddNew
                                    ladrsEmpSchedLine.Fields("EMP_NUM") = !EMP_NUM
                                    If mADOExtensions.RecordsExist(ladrsEmp) Then
                                        ladrsEmp.MoveFirst
                                        ladrsEmp.Find "EMP_NUM = " & !EMP_NUM
                                        If Not (ladrsEmp.EOF) Then
                                            ladrsEmpSchedLine.Fields("EMP_NAME") = ladrsEmp!EMP_NAME
                                        End If
                                    End If
                                    ladrsEmpSchedLine.Fields("CHARGE_DATE") = ladrsActualTime!CHARGE_DATE
                                    ladrsEmpSchedLine.Fields("DATE_WORKED") = ladrsActualTime!DATE_WORKED
                                    ladrsEmpSchedLine.Fields("SHIFT_START_TIME") = !SHIFT_START_TIME
                                    ladrsEmpSchedLine.Fields("SHIFT_START_DATE") = !SHIFT_START_DATE
                                    ladrsEmpSchedLine.Fields("SHIFT_END_TIME") = !SHIFT_END_TIME
                                    ladrsEmpSchedLine.Fields("SHIFT_END_DATE") = !SHIFT_END_DATE
                                    ladrsEmpSchedLine.Fields("SHIFT_HOURS") = !SHIFT_HOURS
                                    ladrsEmpSchedLine.Fields("START_TIME") = ladrsActualTime!START_TIME
                                    ladrsEmpSchedLine.Fields("END_TIME") = ladrsActualTime!END_TIME
                                    ladrsEmpSchedLine.Fields("HOURS") = ladrsActualTime!HOURS
                                    ladrsEmpSchedLine.Fields("EXCEPTION_CD") = ladrsActualTime!EXCEPTION_CD
                                    ladrsEmpSchedLine.Fields("OVER_EXC_CD") = ladrsActualTime!OVER_EXC_CD
                                    If Trim(!SCHED_JOB_NUM) = "" Then
                                        ladrsEmpSchedLine.Fields("JOB_NUM") = !EMP_JOB_NUM
                                    Else
                                        ladrsEmpSchedLine.Fields("JOB_NUM") = !SCHED_JOB_NUM
                                    End If
                                    ladrsEmpSchedLine.Fields("JOB_DESC") = GetJobDesc(ladrsEmpSchedLine.Fields("JOB_NUM"))
                                    If llngProjectID > 0 Then
                                        If mADOExtensions.RecordsExist(ladrsQuals) Then
                                            ladrsQuals.MoveFirst
                                            ladrsQuals.Filter = "EMP_NUM = " & !EMP_NUM & " AND EFF_START_DATE <= '" & vdtStartDate & _
                                                                "' AND EFF_END_DATE >= '" & vdtEndDate & "'"
                                            If Not ladrsQuals.EOF Then
                                                ladrsQuals.Sort = "QUAL_CD, SPEC_SKILL_CD"
                                            End If
                                            For i = 1 To ladrsQuals.RecordCount
                                                If i > 5 Then Exit For
                                                If i = 1 Then
                                                    ladrsEmpSchedLine.Fields("QUALS") = Trim(ladrsQuals!QUAL_CD)
                                                    ladrsEmpSchedLine.Fields("SKILLS") = Trim(ladrsQuals!SPEC_SKILL_CD)
                                                Else
                                                    ladrsEmpSchedLine.Fields("QUALS") = _
                                                    ladrsEmpSchedLine.Fields("QUALS") & " " & Trim(ladrsQuals!QUAL_CD)
                                                    ladrsEmpSchedLine.Fields("SKILLS") = _
                                                    ladrsEmpSchedLine.Fields("SKILLS") & " " & Trim(ladrsQuals!SPEC_SKILL_CD)
                                                End If
                                                ladrsQuals.MoveNext
                                            Next i
                                            ladrsQuals.Filter = adFilterNone
                                        Else
                                            ladrsEmpSchedLine.Fields("QUALS") = ""
                                            ladrsEmpSchedLine.Fields("SKILLS") = ""
                                        End If
                                    End If
                                    ladrsEmpSchedLine.Fields("JOB_POSITION") = ""
                                    ladrsEmpSchedLine.Fields("PROJ_ID") = llngProjectID
                                    ladrsEmpSchedLine.Fields("PROJ_DESC") = lstrProjectDesc
                                    ladrsEmpSchedLine.Fields("COMP_SHIFT_CD") = ladrsActualTime!COMP_SHIFT_CD
                                    ladrsEmpSchedLine.Fields("SYS_COMP_SHIFT_CD") = !SYS_COMP_SHIFT_CD
                                    If mADOExtensions.RecordsExist(ladrsEmpWG) Then
                                        ladrsEmpWG.MoveFirst
                                        ladrsEmpWG.Filter = "EMP_NUM = " & !EMP_NUM & _
                                        " AND EFF_START_DATE <= #" & !CHARGE_DATE & "#" & " AND EFF_END_DATE >= #" & !CHARGE_DATE & "#"
                                        If Not (ladrsEmpWG.EOF) Then
                                            ladrsEmpSchedLine.Fields("WG_NUM") = ladrsEmpWG!WG_NUM
                                        Else
                                            ladrsEmpSchedLine.Fields("WG_NUM") = !WG_NUM
                                        End If
                                        ladrsEmpWG.Filter = adFilterNone
                                    Else
                                        ladrsEmpSchedLine.Fields("WG_NUM") = !WG_NUM
                                    End If
                                    ladrsEmpSchedLine.Fields("WG_DESC") = lstrWGDesc
                                    ladrsEmpSchedLine.Fields("CASE_NUM") = !CASE_NUM
                                    ladrsEmpSchedLine.Fields("FMLA_IND") = !FMLA_IND
                                    ladrsEmpSchedLine.Fields("COMP_CD") = ladrsActualTime!COMP_CD
                                    ladrsEmpSchedLine.Fields("REL_CD") = ladrsActualTime!REL_CD
                                    ladrsEmpSchedLine.Fields("SYSTEM_EXC_CD") = !SYSTEM_EXC_CD
                                    ladrsEmpSchedLine.Fields("SHIFT_CODE") = !SHIFT_CODE
                                    ladrsEmpSchedLine.Fields("SCHED_SRC_CD") = !SCHED_SRC_CD
                                    ladrsEmpSchedLine.Fields("SCHED_CD") = !SCHED_CD
                                    ladrsEmpSchedLine.Fields("SCHED_CHANGE_PMT") = !SCHED_CHANGE_PMT
                                    ladrsEmpSchedLine.Fields("OT_OPP_NUM") = !OT_OPP_NUM
                                    ladrsEmpSchedLine.Fields("TRAV_PAY_HOURS") = !TRAV_PAY_HOURS
                                    ladrsEmpSchedLine.Fields("LINE_NUM") = !LINE_NUM
                                    ladrsEmpSchedLine.Fields("REQ_LINE_IND") = !REQ_LINE_IND
                                    ladrsEmpSchedLine.Fields("SHIFT_START_END") = !SHIFT_START_END
                                    ladrsEmpSchedLine.Fields("SCHED_CD_FOR_SHIFT_COMP") = !SCHED_CD_FOR_SHIFT_COMP
                                    ladrsEmpSchedLine.Fields("BASE_SCHED_SRC_CD") = !BASE_SCHED_SRC_CD
                                    ladrsEmpSchedLine.Fields("LINE_SRC_CD") = !LINE_SRC_CD
                                    ladrsEmpSchedLine.Fields("SHIFT_NUM") = !SHIFT_NUM
                                    ladrsEmpSchedLine.Fields("SHIFT_LUNCH_START_TIME") = !SHIFT_LUNCH_START_TIME
                                    ladrsEmpSchedLine.Fields("SHIFT_LUNCH_END_TIME") = !SHIFT_LUNCH_END_TIME
                                    ladrsEmpSchedLine.Fields("TRADE_NUM") = !TRADE_NUM
                                    ladrsEmpSchedLine.Fields("WG_LINE_IND") = "N"
                                    ladrsEmpSchedLine.Fields("FULL_SHIFT_EXC_IND") = !FULL_SHIFT_EXC_IND
                                    ladrsEmpSchedLine.Fields("DOUBLE_SHIFT_IND") = !DOUBLE_SHIFT_IND
                                    ladrsEmpSchedLine.Fields("LAST_MOD_SCHED_TM") = !LAST_MOD_SCHED_TM
                                    ladrsEmpSchedLine.Fields("LINE_WG") = !LINE_WG
                                    ladrsEmpSchedLine.Fields("LAST_MOD_EMP_NUM") = !LAST_MOD_EMP_NUM
                                    ladrsEmpSchedLine.Fields("LAST_MOD_TIMESTAMP") = !LAST_MOD_TIMESTAMP
                                    ladrsEmpSchedLine.Fields("LINE_JOB_NUM") = !LINE_JOB_NUM
                                    ladrsEmpSchedLine.Fields("SCHED_CHG_EXC_CD") = !SCHED_CHG_EXC_CD
                                    ladrsEmpSchedLine.Fields("ROT_CD") = !ROT_CD
                                    ladrsEmpSchedLine.Fields("ROT_WEEK") = !ROT_WEEK
                                    ladrsEmpSchedLine.Update
                                End If
                                llngPrevActimEmp = !EMP_NUM
                                ldtPrevActimDate = !CHARGE_DATE
                                ladrsActualTime.MoveNext
                            End If
                        Wend
                    End If
                    .MoveNext
                Wend
            End With
        Next llngIndex
        Set lobjSched = Nothing
    End If
    Set ladrsEmpWG = Nothing

    'Get Schedule for Unmanned Lines, if needed
    If lblnDemandInd Then
        If vblnUnmannedLines Then
            Set lobjSched = CreateObject("busSchedule.cbsSchedule")
            'Set lobjSched = New busSchedule.cbsSchedule
            Set ladrsUnmannedSched = lobjSched.GetAdjustedBaseWorkGroupSchedules(vlngWGNum, vdtStartDate, _
                                vdtEndDate, trAdjBaseSchedWithExceptionsAndOvertime, trAllScheduleLevels)
            Set lobjSched = Nothing

            With ladrsUnmannedSched
                .Filter = "EMP_NUM = 0 AND LINE_NUM <> 0"
                If Not ladrsUnmannedSched.EOF Then
                    .MoveFirst
                End If
                While Not ladrsUnmannedSched.EOF
                    ladrsEmpSchedLine.AddNew
                    ladrsEmpSchedLine.Fields("EMP_NUM") = !EMP_NUM
                    ladrsEmpSchedLine.Fields("EMP_NAME") = "UNMANNED"
                    ladrsEmpSchedLine.Fields("CHARGE_DATE") = !CHARGE_DATE
                    ladrsEmpSchedLine.Fields("DATE_WORKED") = !DATE_WORKED
                    ladrsEmpSchedLine.Fields("SHIFT_START_TIME") = !SHIFT_START_TIME
                    ladrsEmpSchedLine.Fields("SHIFT_START_DATE") = !SHIFT_START_DATE
                    ladrsEmpSchedLine.Fields("SHIFT_END_TIME") = !SHIFT_END_TIME
                    ladrsEmpSchedLine.Fields("SHIFT_END_DATE") = !SHIFT_END_DATE
                    ladrsEmpSchedLine.Fields("SHIFT_HOURS") = !SHIFT_HOURS
                    ladrsEmpSchedLine.Fields("START_TIME") = !START_TIME
                    ladrsEmpSchedLine.Fields("END_TIME") = !END_TIME
                    ladrsEmpSchedLine.Fields("HOURS") = !HOURS
                    ladrsEmpSchedLine.Fields("EXCEPTION_CD") = !EXCEPTION_CD
                    ladrsEmpSchedLine.Fields("OVER_EXC_CD") = !OVER_EXC_CD
                    If Trim(!SCHED_JOB_NUM) = "" Then
                        ladrsEmpSchedLine.Fields("JOB_NUM") = !EMP_JOB_NUM
                    Else
                        ladrsEmpSchedLine.Fields("JOB_NUM") = !SCHED_JOB_NUM
                    End If
                    ladrsEmpSchedLine.Fields("JOB_DESC") = GetJobDesc(ladrsEmpSchedLine.Fields("JOB_NUM"))
                    ladrsEmpSchedLine.Fields("QUALS") = ""
                    ladrsEmpSchedLine.Fields("SKILLS") = ""
                    If mADOExtensions.RecordsExist(ladrsJobPosition) Then
                        ladrsJobPosition.MoveFirst
                        ladrsJobPosition.Filter = "CHARGE_DATE = '" & !CHARGE_DATE & "' AND LINE_NUM = " & !LINE_NUM
                        If Not ladrsJobPosition.EOF Then
                            ladrsEmpSchedLine.Fields("JOB_POSITION") = ladrsJobPosition.Fields("JOB_POSITION")
                        Else
                            ladrsEmpSchedLine.Fields("JOB_POSITION") = ""
                        End If
                    Else
                        ladrsEmpSchedLine.Fields("JOB_POSITION") = ""
                    End If
                    ladrsJobPosition.Filter = adFilterNone
                    ladrsEmpSchedLine.Fields("PROJ_ID") = 0
                    ladrsEmpSchedLine.Fields("PROJ_DESC") = ""
                    ladrsEmpSchedLine.Fields("COMP_SHIFT_CD") = !COMP_SHIFT_CD
                    ladrsEmpSchedLine.Fields("SYS_COMP_SHIFT_CD") = !SYS_COMP_SHIFT_CD
                    ladrsEmpSchedLine.Fields("WG_NUM") = !WG_NUM
                    ladrsEmpSchedLine.Fields("WG_DESC") = lstrWGDesc
                    ladrsEmpSchedLine.Fields("CASE_NUM") = !CASE_NUM
                    ladrsEmpSchedLine.Fields("FMLA_IND") = !FMLA_IND
                    ladrsEmpSchedLine.Fields("COMP_CD") = !COMP_CD
                    ladrsEmpSchedLine.Fields("REL_CD") = !REL_CD
                    ladrsEmpSchedLine.Fields("SYSTEM_EXC_CD") = !SYSTEM_EXC_CD
                    ladrsEmpSchedLine.Fields("SHIFT_CODE") = !SHIFT_CODE
                    ladrsEmpSchedLine.Fields("SCHED_SRC_CD") = !SCHED_SRC_CD
                    ladrsEmpSchedLine.Fields("SCHED_CD") = !SCHED_CD
                    ladrsEmpSchedLine.Fields("SCHED_CHANGE_PMT") = !SCHED_CHANGE_PMT
                    ladrsEmpSchedLine.Fields("OT_OPP_NUM") = !OT_OPP_NUM
                    ladrsEmpSchedLine.Fields("TRAV_PAY_HOURS") = !TRAV_PAY_HOURS
                    ladrsEmpSchedLine.Fields("LINE_NUM") = !LINE_NUM
                    ladrsEmpSchedLine.Fields("REQ_LINE_IND") = !REQ_LINE_IND
                    ladrsEmpSchedLine.Fields("SHIFT_START_END") = !SHIFT_START_END
                    ladrsEmpSchedLine.Fields("SCHED_CD_FOR_SHIFT_COMP") = !SCHED_CD_FOR_SHIFT_COMP
                    ladrsEmpSchedLine.Fields("BASE_SCHED_SRC_CD") = !BASE_SCHED_SRC_CD
                    ladrsEmpSchedLine.Fields("LINE_SRC_CD") = !LINE_SRC_CD
                    ladrsEmpSchedLine.Fields("SHIFT_NUM") = !SHIFT_NUM
                    ladrsEmpSchedLine.Fields("SHIFT_LUNCH_START_TIME") = !SHIFT_LUNCH_START_TIME
                    ladrsEmpSchedLine.Fields("SHIFT_LUNCH_END_TIME") = !SHIFT_LUNCH_END_TIME
                    ladrsEmpSchedLine.Fields("TRADE_NUM") = !TRADE_NUM
                    ladrsEmpSchedLine.Fields("WG_LINE_IND") = "Y"
                    ladrsEmpSchedLine.Fields("FULL_SHIFT_EXC_IND") = !FULL_SHIFT_EXC_IND
                    ladrsEmpSchedLine.Fields("DOUBLE_SHIFT_IND") = !DOUBLE_SHIFT_IND
                    ladrsEmpSchedLine.Fields("LAST_MOD_SCHED_TM") = !LAST_MOD_SCHED_TM
                    ladrsEmpSchedLine.Fields("LINE_WG") = !LINE_WG
                    ladrsEmpSchedLine.Fields("LAST_MOD_EMP_NUM") = !LAST_MOD_EMP_NUM
                    ladrsEmpSchedLine.Fields("LAST_MOD_TIMESTAMP") = !LAST_MOD_TIMESTAMP
                    ladrsEmpSchedLine.Fields("LINE_JOB_NUM") = !LINE_JOB_NUM
                    ladrsEmpSchedLine.Fields("SCHED_CHG_EXC_CD") = !SCHED_CHG_EXC_CD
                    ladrsEmpSchedLine.Fields("ROT_CD") = !ROT_CD
                    ladrsEmpSchedLine.Fields("ROT_WEEK") = !ROT_WEEK
                    ladrsEmpSchedLine.Update
                    .MoveNext
                Wend
            End With
        End If
    End If
        
    If mADOExtensions.RecordsExist(ladrsEmpSchedLine) Then
        Set GetWGSchedule = ladrsEmpSchedLine
    Else
        Set GetWGSchedule = Nothing
    End If

    Set lobjSched = Nothing
    Set lobjdbSched = Nothing
    Set lobjTimeCard = Nothing
    Set lobjCodeTable = Nothing
    Set lobjEmployee = Nothing
    Set lobjEmp = Nothing

    Set ladrsWG = Nothing
    Set ladrsEmpWG = Nothing
    Set ladrsJobPosition = Nothing
    Set ladrsProjectInfo = Nothing
    Set ladrsActualTime = Nothing
    Set ladrsEmpSched = Nothing
    Set ladrsEmpSchedLine = Nothing
    Set ladrsUnmannedSched = Nothing
    Set ladrsDummy = Nothing
    Set ladrsEmp = Nothing
    Set ladrsQuals = Nothing
    Set ladrsEmpsinWG = Nothing
    Set ladrsEmpSched = Nothing

    Exit Function
GetWGSchedule_eh:
    Set lobjSched = Nothing
    Set lobjdbSched = Nothing
    Set lobjTimeCard = Nothing
    Set lobjCodeTable = Nothing
    Set lobjEmployee = Nothing
    Set lobjEmp = Nothing

    Set ladrsWG = Nothing
    Set ladrsEmpWG = Nothing
    Set ladrsJobPosition = Nothing
    Set ladrsProjectInfo = Nothing
    Set ladrsActualTime = Nothing
    Set ladrsEmpSched = Nothing
    Set ladrsEmpSchedLine = Nothing
    Set ladrsUnmannedSched = Nothing
    Set ladrsDummy = Nothing
    Set ladrsEmp = Nothing
    Set ladrsQuals = Nothing
    Set ladrsEmpsinWG = Nothing
    Set ladrsEmpSched = Nothing

    TrsRaiseError m_strCMODULENAME, "GetWGSchedule"

End Function
'Function gets WorkGroup Schedule for Daily Roster Report
Public Function GetDailyRosterSchedule(ByVal vlngWGNum As Long, _
                                                    ByVal vdtRequestDate As Date) _
                                                    As ADODB.Recordset

    'On Error GoTo GetDailyRosterSchedule_eh

    'Dim ladrsWG As ADODB.Recordset
    'Dim lobjWG As dbWorkGroup.cdWorkGroup
    'Dim ladrsEmpsinWG As ADODB.Recordset
    'Dim llngIndex As Long
    'Dim llngPrevActimEmp As Long
    'Dim ldtPrevActimDate As Date
    'Dim llngEmpCount As Long
    'Dim lobjSched As busSchedule.cbsSchedule
    'Dim lobjTimeCard As dbTimeCard.cdTimeCard
    'Dim lobjCodeTable As dbCodeTables.cdCodeTables
    'Dim lobjEmployee As busEmployee.cbsEmployee

    'Dim ladrsProjectInfo As ADODB.Recordset
    'Dim ladrsActualTime As ADODB.Recordset
    'Dim ladrsEmpSched As ADODB.Recordset
    'Dim ladrsEmpSchedLine As ADODB.Recordset
    'Dim ladrsUnmannedSched As ADODB.Recordset
    'Dim ladrsDummy As ADODB.Recordset
    'Dim lblnDemandInd As Boolean
    'Dim lstrWGDesc As String
    'Dim ladrsEmp As ADODB.Recordset
    'Dim llngarrEmps() As Long
    'Dim lstrarrEmps As String
    'Dim i As Integer
    'Const mconFIELDATTR = adFldUpdatable Or adFldIsNullable

    'Set ladrsWG = GetWorkGroupAttributes(vlngWGNum)
    'lblnDemandInd = False
    'If mADOExtensions.RecordsExist(ladrsWG) Then
    '    If ladrsWG!DEMAND_IND = "Y" Then
    '        lblnDemandInd = True
    '    End If
    '    lstrWGDesc = ladrsWG!WG_DESC
    'End If
    'Set ladrsWG = Nothing
    
    'Set lobjWG = CreateObject("dbWorkGroup.cdWorkGroup")
    'Set ladrsEmpsinWG = lobjWG.GetEmployeesForDailyRosterRpt(vlngWGNum, _
                                                             vdtRequestDate)
    'Set lobjWG = Nothing

    'If mADOExtensions.RecordsExist(ladrsEmpsinWG) Then
    '    With ladrsEmpsinWG
    '        ReDim llngarrEmps(1 To .RecordCount)
    '        .Sort = "EMP_NUM"
    '        .Filter = "EMP_NUM <> 0"
    '        .MoveFirst
    '        llngEmpCount = 1
    '        For llngIndex = 1 To .RecordCount
    '            llngarrEmps(llngEmpCount) = !EMP_NUM
    '            If llngIndex = .RecordCount Then
    '                lstrarrEmps = lstrarrEmps & !EMP_NUM
    '            Else
    '                lstrarrEmps = lstrarrEmps & !EMP_NUM & ","
    '            End If
    '            llngEmpCount = llngEmpCount + 1
    '            .MoveNext
    '        Next llngIndex
    '    End With
    '    Set ladrsEmpSchedLine = New ADODB.Recordset
    '    With ladrsEmpSchedLine
    '        .Fields.Append "EMP_NUM", adInteger, , mconFIELDATTR
    '        .Fields.Append "EMP_NAME", adVarChar, 256, mconFIELDATTR
    '        .Fields.Append "CHARGE_DATE", adDate, , mconFIELDATTR
    '        .Fields.Append "DATE_WORKED", adDate, , mconFIELDATTR
    '        .Fields.Append "SHIFT_START_TIME", adChar, 5, mconFIELDATTR
    '        .Fields.Append "SHIFT_START_DATE", adDate, , mconFIELDATTR
    '        .Fields.Append "SHIFT_END_TIME", adChar, 5, mconFIELDATTR
    '        .Fields.Append "SHIFT_END_DATE", adDate, , mconFIELDATTR
    '        .Fields.Append "SHIFT_HOURS", adDecimal, , mconFIELDATTR
    '        .Fields.Append "START_TIME", adChar, 5, mconFIELDATTR
    '        .Fields.Append "END_TIME", adChar, 5, mconFIELDATTR
    '        .Fields.Append "HOURS", adDecimal, , mconFIELDATTR
    '        .Fields.Append "EXCEPTION_CD", adChar, 3, mconFIELDATTR
    '        .Fields.Append "OVER_EXC_CD", adChar, 3, mconFIELDATTR
    '        .Fields.Append "WG_NUM", adInteger, , mconFIELDATTR
    '        .Fields.Append "WG_DESC", adVarChar, 256, mconFIELDATTR
    '        .Fields.Append "SYSTEM_EXC_CD", adChar, 3, mconFIELDATTR
    '        .Fields.Append "SHIFT_CODE", adChar, 1, mconFIELDATTR
    '        .Fields.Append "SCHED_SRC_CD", adInteger, , mconFIELDATTR
    '        .Fields.Append "SCHED_CD", adInteger, , mconFIELDATTR
    '        .Fields.Append "OT_OPP_NUM", adInteger, , mconFIELDATTR
    '        .Fields.Append "LINE_NUM", adInteger, , mconFIELDATTR
    '        .Fields.Append "SHIFT_START_END", adChar, 11, mconFIELDATTR
    '        .Fields.Append "SHIFT_NUM", adInteger, , mconFIELDATTR
    '        .Fields.Append "SHIFT_LUNCH_START_TIME", adChar, 5, mconFIELDATTR
    '        .Fields.Append "SHIFT_LUNCH_END_TIME", adChar, 5, mconFIELDATTR
    '        .CursorLocation = adUseClient
    '        .CursorType = adOpenStatic
    '        .LockType = adLockBatchOptimistic
    '        .Open
    '    End With
        
    '    Set lobjEmployee = CreateObject("busEmployee.cbsEmployee")
    '    Set lobjEmployee = New busEmployee.cbsEmployee
    '    Set ladrsEmp = lobjEmployee.GetEmployee(llngarrEmps)
    '    Set lobjEmployee = Nothing

    '    llngPrevActimEmp = 0
    '    ldtPrevActimDate = "9999-12-31"
        
    '    Set lobjTimeCard = CreateObject("dbTimeCard.cdTimeCard")
    '    Set ladrsActualTime = lobjTimeCard.GetActualTimeEntriesWithoutOtherPay( _
                              llngarrEmps, vdtRequestDate, vdtRequestDate)
    '    Set lobjTimeCard = Nothing

    '    Set lobjSched = CreateObject("busSchedule.cbsSchedule")
        'Set lobjSched = New busSchedule.cbsSchedule

    '    For llngIndex = 1 To llngEmpCount - 1
    '        Set ladrsEmpSched = lobjSched.GetAdjustedBaseSchedule(llngarrEmps(llngIndex), vdtRequestDate _
    '                          , vdtRequestDate, ladrsDummy, False, trAdjBaseSchedWithExceptionsAndOvertime _
    '                          , trAllScheduleLevels, trRelief, trAllExceptions)
    '        Set ladrsDummy = Nothing

    '        With ladrsEmpSched
    '            .MoveFirst
    '            While Not ladrsEmpSched.EOF
    '                If lblnDemandInd Then
    '                    ladrsEmpSchedLine.AddNew
    '                    ladrsEmpSchedLine.Fields("EMP_NUM") = !EMP_NUM
    '                    If mADOExtensions.RecordsExist(ladrsEmp) Then
    '                        ladrsEmp.MoveFirst
    '                        ladrsEmp.Find "EMP_NUM = " & !EMP_NUM
    '                        If Not (ladrsEmp.EOF) Then
    '                            ladrsEmpSchedLine.Fields("EMP_NAME") = ladrsEmp!EMP_NAME
    '                        End If
    '                    End If
    '                    ladrsEmpSchedLine.Fields("CHARGE_DATE") = !CHARGE_DATE
    '                    ladrsEmpSchedLine.Fields("DATE_WORKED") = !DATE_WORKED
    '                    ladrsEmpSchedLine.Fields("SHIFT_START_TIME") = !SHIFT_START_TIME
    '                    ladrsEmpSchedLine.Fields("SHIFT_START_DATE") = !SHIFT_START_DATE
    '                    ladrsEmpSchedLine.Fields("SHIFT_END_TIME") = !SHIFT_END_TIME
    '                    ladrsEmpSchedLine.Fields("SHIFT_END_DATE") = !SHIFT_END_DATE
    '                    ladrsEmpSchedLine.Fields("SHIFT_HOURS") = !SHIFT_HOURS
    '                    ladrsEmpSchedLine.Fields("START_TIME") = !START_TIME
    '                    ladrsEmpSchedLine.Fields("END_TIME") = !END_TIME
    '                    ladrsEmpSchedLine.Fields("HOURS") = !HOURS
    '                    ladrsEmpSchedLine.Fields("EXCEPTION_CD") = !EXCEPTION_CD
    '                    ladrsEmpSchedLine.Fields("OVER_EXC_CD") = !OVER_EXC_CD
    '                    ladrsEmpSchedLine.Fields("WG_NUM") = !WG_NUM
    '                    ladrsEmpSchedLine.Fields("WG_DESC") = lstrWGDesc
    '                    ladrsEmpSchedLine.Fields("SYSTEM_EXC_CD") = !SYSTEM_EXC_CD
    '                    ladrsEmpSchedLine.Fields("SHIFT_CODE") = !SHIFT_CODE
    '                    ladrsEmpSchedLine.Fields("SCHED_SRC_CD") = !SCHED_SRC_CD
    '                    ladrsEmpSchedLine.Fields("SCHED_CD") = !SCHED_CD
    '                    ladrsEmpSchedLine.Fields("OT_OPP_NUM") = !OT_OPP_NUM
    '                    ladrsEmpSchedLine.Fields("LINE_NUM") = !LINE_NUM
    '                    ladrsEmpSchedLine.Fields("SHIFT_START_END") = !SHIFT_START_END
    '                    ladrsEmpSchedLine.Fields("SHIFT_NUM") = !SHIFT_NUM
    '                    ladrsEmpSchedLine.Fields("SHIFT_LUNCH_START_TIME") = !SHIFT_LUNCH_START_TIME
    '                    ladrsEmpSchedLine.Fields("SHIFT_LUNCH_END_TIME") = !SHIFT_LUNCH_END_TIME
    '                    ladrsEmpSchedLine.Update
    '                Else
    '                    ladrsEmpSchedLine.AddNew
    '                    ladrsEmpSchedLine.Fields("EMP_NUM") = !EMP_NUM
    '                    If mADOExtensions.RecordsExist(ladrsEmp) Then
    '                        ladrsEmp.MoveFirst
    '                        ladrsEmp.Find "EMP_NUM = " & !EMP_NUM
    '                        If Not (ladrsEmp.EOF) Then
    '                            ladrsEmpSchedLine.Fields("EMP_NAME") = ladrsEmp!EMP_NAME
    '                        End If
    '                    End If
    '                    ladrsEmpSchedLine.Fields("CHARGE_DATE") = !CHARGE_DATE
    '                    ladrsEmpSchedLine.Fields("DATE_WORKED") = !DATE_WORKED
    '                    ladrsEmpSchedLine.Fields("SHIFT_START_TIME") = !SHIFT_START_TIME
    '                    ladrsEmpSchedLine.Fields("SHIFT_START_DATE") = !SHIFT_START_DATE
    '                    ladrsEmpSchedLine.Fields("SHIFT_END_TIME") = !SHIFT_END_TIME
    '                    ladrsEmpSchedLine.Fields("SHIFT_END_DATE") = !SHIFT_END_DATE
    '                    ladrsEmpSchedLine.Fields("SHIFT_HOURS") = !SHIFT_HOURS
    '                    ladrsEmpSchedLine.Fields("START_TIME") = !START_TIME
    '                    ladrsEmpSchedLine.Fields("END_TIME") = !END_TIME
    '                    ladrsEmpSchedLine.Fields("HOURS") = !HOURS
    '                    ladrsEmpSchedLine.Fields("EXCEPTION_CD") = !EXCEPTION_CD
    '                    ladrsEmpSchedLine.Fields("OVER_EXC_CD") = !OVER_EXC_CD
    '                    ladrsEmpSchedLine.Fields("WG_NUM") = !WG_NUM
    '                    ladrsEmpSchedLine.Fields("WG_DESC") = lstrWGDesc
    '                    ladrsEmpSchedLine.Fields("SYSTEM_EXC_CD") = !SYSTEM_EXC_CD
    '                    ladrsEmpSchedLine.Fields("SHIFT_CODE") = !SHIFT_CODE
    '                    ladrsEmpSchedLine.Fields("OT_OPP_NUM") = !OT_OPP_NUM
    '                    ladrsEmpSchedLine.Fields("LINE_NUM") = !LINE_NUM
    '                    ladrsEmpSchedLine.Fields("SHIFT_START_END") = !SHIFT_START_END
    '                    ladrsEmpSchedLine.Fields("SHIFT_NUM") = !SHIFT_NUM
    '                    ladrsEmpSchedLine.Fields("SHIFT_LUNCH_START_TIME") = !SHIFT_LUNCH_START_TIME
    '                    ladrsEmpSchedLine.Fields("SHIFT_LUNCH_END_TIME") = !SHIFT_LUNCH_END_TIME
    '                    ladrsEmpSchedLine.Update
    '                End If
    '                'get actual time
    '                If ladrsEmpSched!EMP_NUM <> llngPrevActimEmp Or _
    '                   ladrsEmpSched!CHARGE_DATE <> ldtPrevActimDate Then
    '                    ladrsActualTime.Filter = adFilterNone
    '                    ladrsActualTime.Filter = "EMP_NUM = " & ladrsEmpSched!EMP_NUM & _
    '                    " AND CHARGE_DATE = #" & ladrsEmpSched!CHARGE_DATE & "#"
    '                    While Not ladrsActualTime.EOF
    '                        If Not ladrsActualTime.EOF Then
    '                            If lblnDemandInd Then
    '                                ladrsEmpSchedLine.AddNew
    '                                ladrsEmpSchedLine.Fields("EMP_NUM") = !EMP_NUM
    '                                If mADOExtensions.RecordsExist(ladrsEmp) Then
    '                                    ladrsEmp.MoveFirst
    '                                    ladrsEmp.Find "EMP_NUM = " & !EMP_NUM
    '                                    If Not (ladrsEmp.EOF) Then
    '                                        ladrsEmpSchedLine.Fields("EMP_NAME") = ladrsEmp!EMP_NAME
    '                                    End If
    '                                End If
    '                                ladrsEmpSchedLine.Fields("CHARGE_DATE") = ladrsActualTime!CHARGE_DATE
    '                                ladrsEmpSchedLine.Fields("DATE_WORKED") = ladrsActualTime!DATE_WORKED
    '                                ladrsEmpSchedLine.Fields("SHIFT_START_TIME") = !SHIFT_START_TIME
    '                                ladrsEmpSchedLine.Fields("SHIFT_START_DATE") = !SHIFT_START_DATE
    '                                ladrsEmpSchedLine.Fields("SHIFT_END_TIME") = !SHIFT_END_TIME
    '                                ladrsEmpSchedLine.Fields("SHIFT_END_DATE") = !SHIFT_END_DATE
    '                                ladrsEmpSchedLine.Fields("SHIFT_HOURS") = !SHIFT_HOURS
    '                                ladrsEmpSchedLine.Fields("START_TIME") = ladrsActualTime!START_TIME
    '                                ladrsEmpSchedLine.Fields("END_TIME") = ladrsActualTime!END_TIME
    '                                ladrsEmpSchedLine.Fields("HOURS") = ladrsActualTime!HOURS
    '                                ladrsEmpSchedLine.Fields("EXCEPTION_CD") = ladrsActualTime!EXCEPTION_CD
    '                                ladrsEmpSchedLine.Fields("OVER_EXC_CD") = ladrsActualTime!OVER_EXC_CD
    '                                ladrsEmpSchedLine.Fields("WG_NUM") = !WG_NUM
    '                                ladrsEmpSchedLine.Fields("WG_DESC") = lstrWGDesc
    '                                ladrsEmpSchedLine.Fields("SYSTEM_EXC_CD") = !SYSTEM_EXC_CD
    '                                ladrsEmpSchedLine.Fields("SHIFT_CODE") = !SHIFT_CODE
    '                                ladrsEmpSchedLine.Fields("OT_OPP_NUM") = !OT_OPP_NUM
    '                                ladrsEmpSchedLine.Fields("LINE_NUM") = !LINE_NUM
    '                                ladrsEmpSchedLine.Fields("SHIFT_START_END") = !SHIFT_START_END
    '                                ladrsEmpSchedLine.Fields("SHIFT_NUM") = !SHIFT_NUM
    '                                ladrsEmpSchedLine.Fields("SHIFT_LUNCH_START_TIME") = !SHIFT_LUNCH_START_TIME
    '                                ladrsEmpSchedLine.Fields("SHIFT_LUNCH_END_TIME") = !SHIFT_LUNCH_END_TIME
    '                                ladrsEmpSchedLine.Update
    '                            Else
    '                                ladrsEmpSchedLine.AddNew
    '                                ladrsEmpSchedLine.Fields("EMP_NUM") = !EMP_NUM
    '                                If mADOExtensions.RecordsExist(ladrsEmp) Then
    '                                    ladrsEmp.MoveFirst
    '                                    ladrsEmp.Find "EMP_NUM = " & !EMP_NUM
    '                                    If Not (ladrsEmp.EOF) Then
    '                                        ladrsEmpSchedLine.Fields("EMP_NAME") = ladrsEmp!EMP_NAME
    '                                    End If
    '                                End If
    '                                ladrsEmpSchedLine.Fields("CHARGE_DATE") = ladrsActualTime!CHARGE_DATE
    '                                ladrsEmpSchedLine.Fields("DATE_WORKED") = ladrsActualTime!DATE_WORKED
    '                                ladrsEmpSchedLine.Fields("SHIFT_START_TIME") = !SHIFT_START_TIME
    '                                ladrsEmpSchedLine.Fields("SHIFT_START_DATE") = !SHIFT_START_DATE
    '                                ladrsEmpSchedLine.Fields("SHIFT_END_TIME") = !SHIFT_END_TIME
    '                                ladrsEmpSchedLine.Fields("SHIFT_END_DATE") = !SHIFT_END_DATE
    '                                ladrsEmpSchedLine.Fields("SHIFT_HOURS") = !SHIFT_HOURS
    '                                ladrsEmpSchedLine.Fields("START_TIME") = ladrsActualTime!START_TIME
    '                                ladrsEmpSchedLine.Fields("END_TIME") = ladrsActualTime!END_TIME
    '                                ladrsEmpSchedLine.Fields("HOURS") = ladrsActualTime!HOURS
    '                                ladrsEmpSchedLine.Fields("EXCEPTION_CD") = ladrsActualTime!EXCEPTION_CD
    '                                ladrsEmpSchedLine.Fields("OVER_EXC_CD") = ladrsActualTime!OVER_EXC_CD
    '                                ladrsEmpSchedLine.Fields("WG_NUM") = !WG_NUM
    '                                ladrsEmpSchedLine.Fields("WG_DESC") = lstrWGDesc
    '                                ladrsEmpSchedLine.Fields("SYSTEM_EXC_CD") = !SYSTEM_EXC_CD
    '                                ladrsEmpSchedLine.Fields("SHIFT_CODE") = !SHIFT_CODE
    '                                ladrsEmpSchedLine.Fields("SCHED_SRC_CD") = !SCHED_SRC_CD
    '                                ladrsEmpSchedLine.Fields("SCHED_CD") = !SCHED_CD
    '                                ladrsEmpSchedLine.Fields("OT_OPP_NUM") = !OT_OPP_NUM
    '                                ladrsEmpSchedLine.Fields("LINE_NUM") = !LINE_NUM
    '                                ladrsEmpSchedLine.Fields("SHIFT_START_END") = !SHIFT_START_END
    '                                ladrsEmpSchedLine.Fields("SHIFT_NUM") = !SHIFT_NUM
    '                                ladrsEmpSchedLine.Fields("SHIFT_LUNCH_START_TIME") = !SHIFT_LUNCH_START_TIME
    '                                ladrsEmpSchedLine.Fields("SHIFT_LUNCH_END_TIME") = !SHIFT_LUNCH_END_TIME
    '                                ladrsEmpSchedLine.Update
    '                            End If
    '                            llngPrevActimEmp = !EMP_NUM
    '                            ldtPrevActimDate = !CHARGE_DATE
    '                            ladrsActualTime.MoveNext
    '                        End If
    '                    Wend
    '                End If
    '                .MoveNext
    '            Wend
    '        End With
    '    Next llngIndex
    '    Set lobjSched = Nothing
    'End If
        
    'If mADOExtensions.RecordsExist(ladrsEmpSchedLine) Then
    '    Set GetDailyRosterSchedule = ladrsEmpSchedLine
    'Else
    '    Set GetDailyRosterSchedule = Nothing
    'End If

    'Set lobjSched = Nothing
    'Set lobjTimeCard = Nothing
    'Set lobjCodeTable = Nothing
    'Set lobjEmployee = Nothing
    
    'Set ladrsWG = Nothing
    'Set ladrsActualTime = Nothing
    'Set ladrsEmpSched = Nothing
    'Set ladrsEmpSchedLine = Nothing
    'Set ladrsUnmannedSched = Nothing
    'Set ladrsDummy = Nothing
    'Set ladrsEmp = Nothing
    'Set ladrsEmpsinWG = Nothing
    'Set ladrsEmpSched = Nothing

    'Exit Function
'GetDailyRosterSchedule_eh:
    'Set lobjSched = Nothing
    'Set lobjTimeCard = Nothing
    'Set lobjCodeTable = Nothing
    'Set lobjEmployee = Nothing
    
    'Set ladrsWG = Nothing
    'Set ladrsActualTime = Nothing
    'Set ladrsEmpSched = Nothing
    'Set ladrsEmpSchedLine = Nothing
    'Set ladrsUnmannedSched = Nothing
    'Set ladrsDummy = Nothing
    'Set ladrsEmp = Nothing
    'Set ladrsEmpsinWG = Nothing
    'Set ladrsEmpSched = Nothing

    'TrsRaiseError m_strCMODULENAME, "GetDailyRosterSchedule"

    
End Function
'Function gets WorkGroup Schedule for Daily Roster Report
Public Function GetDailyRosterSchedule2(ByVal vlngWGNum As Long, _
                                                    ByVal vdtStartDate As Date, _
                                                    ByVal vdtEndDate As Date) _
                                                    As ADODB.Recordset

    'On Error GoTo GetDailyRosterSchedule2_eh

    'Dim ladrsWG As ADODB.Recordset
    'Dim lobjWG As dbWorkGroup.cdWorkGroup
    'Dim ladrsEmpsinWG As ADODB.Recordset
    'Dim llngIndex As Long
    'Dim llngPrevActimEmp As Long
    'Dim ldtPrevActimDate As Date
    'Dim llngEmpCount As Long
    'Dim lobjSched As busSchedule.cbsSchedule
    'Dim lobjTimeCard As dbTimeCard.cdTimeCard
    'Dim lobjCodeTable As dbCodeTables.cdCodeTables
    'Dim lobjEmployee As busEmployee.cbsEmployee

    'Dim ladrsProjectInfo As ADODB.Recordset
    'Dim ladrsActualTime As ADODB.Recordset
    'Dim ladrsEmpSched As ADODB.Recordset
    'Dim ladrsEmpSchedLine As ADODB.Recordset
    'Dim ladrsUnmannedSched As ADODB.Recordset
    'Dim ladrsDummy As ADODB.Recordset
    'Dim lblnDemandInd As Boolean
    'Dim lstrWGDesc As String
    'Dim ladrsEmp As ADODB.Recordset
    'Dim llngarrEmps() As Long
    'Dim lstrarrEmps As String
    'Dim i As Integer
    'Const mconFIELDATTR = adFldUpdatable Or adFldIsNullable

    'Set ladrsWG = GetWorkGroupAttributes(vlngWGNum)
    'lblnDemandInd = False
    'If mADOExtensions.RecordsExist(ladrsWG) Then
    '    If ladrsWG!DEMAND_IND = "Y" Then
    '        lblnDemandInd = True
    '    End If
    '    lstrWGDesc = ladrsWG!WG_DESC
    'End If
    'Set ladrsWG = Nothing
    
    'Set lobjWG = CreateObject("dbWorkGroup.cdWorkGroup")
    'Set ladrsEmpsinWG = lobjWG.GetEmployeesForDailyRosterRpt2(vlngWGNum, _
                                                             vdtStartDate, vdtEndDate)
    'Set lobjWG = Nothing

    'If mADOExtensions.RecordsExist(ladrsEmpsinWG) Then
    '    With ladrsEmpsinWG
    '        ReDim llngarrEmps(1 To .RecordCount)
    '        .Sort = "EMP_NUM"
    '        .Filter = "EMP_NUM <> 0"
    '        .MoveFirst
    '        llngEmpCount = 1
    '        For llngIndex = 1 To .RecordCount
    '            llngarrEmps(llngEmpCount) = !EMP_NUM
    '            If llngIndex = .RecordCount Then
    '                lstrarrEmps = lstrarrEmps & !EMP_NUM
    '            Else
    '                lstrarrEmps = lstrarrEmps & !EMP_NUM & ","
    '            End If
    '            llngEmpCount = llngEmpCount + 1
    '            .MoveNext
    '        Next llngIndex
    '    End With
    '    Set ladrsEmpSchedLine = New ADODB.Recordset
    '    With ladrsEmpSchedLine
    '        .Fields.Append "EMP_NUM", adInteger, , mconFIELDATTR
    '        .Fields.Append "EMP_NAME", adVarChar, 256, mconFIELDATTR
    '        .Fields.Append "CHARGE_DATE", adDate, , mconFIELDATTR
    '        .Fields.Append "DATE_WORKED", adDate, , mconFIELDATTR
    '        .Fields.Append "SHIFT_START_TIME", adChar, 5, mconFIELDATTR
    '        .Fields.Append "SHIFT_START_DATE", adDate, , mconFIELDATTR
    '        .Fields.Append "SHIFT_END_TIME", adChar, 5, mconFIELDATTR
    '        .Fields.Append "SHIFT_END_DATE", adDate, , mconFIELDATTR
    '        .Fields.Append "SHIFT_HOURS", adDecimal, , mconFIELDATTR
    '        .Fields.Append "START_TIME", adChar, 5, mconFIELDATTR
    '        .Fields.Append "END_TIME", adChar, 5, mconFIELDATTR
    '        .Fields.Append "HOURS", adDecimal, , mconFIELDATTR
    '        .Fields.Append "EXCEPTION_CD", adChar, 3, mconFIELDATTR
    '        .Fields.Append "OVER_EXC_CD", adChar, 3, mconFIELDATTR
    '        .Fields.Append "WG_NUM", adInteger, , mconFIELDATTR
    '        .Fields.Append "WG_DESC", adVarChar, 256, mconFIELDATTR
    '        .Fields.Append "SYSTEM_EXC_CD", adChar, 3, mconFIELDATTR
    '        .Fields.Append "SHIFT_CODE", adChar, 1, mconFIELDATTR
    '        .Fields.Append "SCHED_SRC_CD", adInteger, , mconFIELDATTR
    '        .Fields.Append "SCHED_CD", adInteger, , mconFIELDATTR
    '        .Fields.Append "OT_OPP_NUM", adInteger, , mconFIELDATTR
    '        .Fields.Append "LINE_NUM", adInteger, , mconFIELDATTR
    '        .Fields.Append "SHIFT_START_END", adChar, 11, mconFIELDATTR
    '        .Fields.Append "SHIFT_NUM", adInteger, , mconFIELDATTR
    '        .Fields.Append "SHIFT_LUNCH_START_TIME", adChar, 5, mconFIELDATTR
    '        .Fields.Append "SHIFT_LUNCH_END_TIME", adChar, 5, mconFIELDATTR
    '        .CursorLocation = adUseClient
    '        .CursorType = adOpenStatic
    '        .LockType = adLockBatchOptimistic
    '        .Open
    '    End With
        
    '    Set lobjEmployee = CreateObject("busEmployee.cbsEmployee")
    '    Set lobjEmployee = New busEmployee.cbsEmployee
    '    Set ladrsEmp = lobjEmployee.GetEmployee(llngarrEmps)
    '    Set lobjEmployee = Nothing

    '    llngPrevActimEmp = 0
    '    ldtPrevActimDate = "9999-12-31"
        
    '    Set lobjTimeCard = CreateObject("dbTimeCard.cdTimeCard")
    '    Set ladrsActualTime = lobjTimeCard.GetActualTimeEntriesWithoutOtherPay( _
    '                          llngarrEmps, vdtStartDate, vdtEndDate)
    '    Set lobjTimeCard = Nothing

    '    Set lobjSched = CreateObject("busSchedule.cbsSchedule")
    '    'Set lobjSched = New busSchedule.cbsSchedule

    '    For llngIndex = 1 To llngEmpCount - 1
    '        Set ladrsEmpSched = lobjSched.GetAdjustedBaseSchedule(llngarrEmps(llngIndex), vdtStartDate _
    '                          , vdtEndDate, ladrsDummy, False, trAdjBaseSchedWithExceptionsAndOvertime _
    '                          , trAllScheduleLevels, trRelief, trAllExceptions)
    '        Set ladrsDummy = Nothing

    '        With ladrsEmpSched
    '            .MoveFirst
    '            While Not ladrsEmpSched.EOF
    '                If lblnDemandInd Then
    '                    ladrsEmpSchedLine.AddNew
    '                    ladrsEmpSchedLine.Fields("EMP_NUM") = !EMP_NUM
    '                    If mADOExtensions.RecordsExist(ladrsEmp) Then
    '                        ladrsEmp.MoveFirst
    '                        ladrsEmp.Find "EMP_NUM = " & !EMP_NUM
    '                        If Not (ladrsEmp.EOF) Then
    '                            ladrsEmpSchedLine.Fields("EMP_NAME") = ladrsEmp!EMP_NAME
    '                        End If
    '                    End If
    '                    ladrsEmpSchedLine.Fields("CHARGE_DATE") = !CHARGE_DATE
    '                    ladrsEmpSchedLine.Fields("DATE_WORKED") = !DATE_WORKED
    '                    ladrsEmpSchedLine.Fields("SHIFT_START_TIME") = !SHIFT_START_TIME
    '                    ladrsEmpSchedLine.Fields("SHIFT_START_DATE") = !SHIFT_START_DATE
    '                    ladrsEmpSchedLine.Fields("SHIFT_END_TIME") = !SHIFT_END_TIME
    '                    ladrsEmpSchedLine.Fields("SHIFT_END_DATE") = !SHIFT_END_DATE
    '                    ladrsEmpSchedLine.Fields("SHIFT_HOURS") = !SHIFT_HOURS
    '                    ladrsEmpSchedLine.Fields("START_TIME") = !START_TIME
    '                    ladrsEmpSchedLine.Fields("END_TIME") = !END_TIME
    '                    ladrsEmpSchedLine.Fields("HOURS") = !HOURS
    '                    ladrsEmpSchedLine.Fields("EXCEPTION_CD") = !EXCEPTION_CD
    '                    ladrsEmpSchedLine.Fields("OVER_EXC_CD") = !OVER_EXC_CD
    '                    ladrsEmpSchedLine.Fields("WG_NUM") = !WG_NUM
    '                    ladrsEmpSchedLine.Fields("WG_DESC") = lstrWGDesc
    '                    ladrsEmpSchedLine.Fields("SYSTEM_EXC_CD") = !SYSTEM_EXC_CD
    '                    ladrsEmpSchedLine.Fields("SHIFT_CODE") = !SHIFT_CODE
    '                    ladrsEmpSchedLine.Fields("SCHED_SRC_CD") = !SCHED_SRC_CD
    '                    ladrsEmpSchedLine.Fields("SCHED_CD") = !SCHED_CD
    '                    ladrsEmpSchedLine.Fields("OT_OPP_NUM") = !OT_OPP_NUM
    '                    ladrsEmpSchedLine.Fields("LINE_NUM") = !LINE_NUM
    '                    ladrsEmpSchedLine.Fields("SHIFT_START_END") = !SHIFT_START_END
    '                    ladrsEmpSchedLine.Fields("SHIFT_NUM") = !SHIFT_NUM
    '                    ladrsEmpSchedLine.Fields("SHIFT_LUNCH_START_TIME") = !SHIFT_LUNCH_START_TIME
    '                    ladrsEmpSchedLine.Fields("SHIFT_LUNCH_END_TIME") = !SHIFT_LUNCH_END_TIME
    '                    ladrsEmpSchedLine.Update
    '                Else
    '                    ladrsEmpSchedLine.AddNew
    '                    ladrsEmpSchedLine.Fields("EMP_NUM") = !EMP_NUM
    '                    If mADOExtensions.RecordsExist(ladrsEmp) Then
    '                        ladrsEmp.MoveFirst
    '                        ladrsEmp.Find "EMP_NUM = " & !EMP_NUM
    '                        If Not (ladrsEmp.EOF) Then
    '                            ladrsEmpSchedLine.Fields("EMP_NAME") = ladrsEmp!EMP_NAME
    '                        End If
    '                    End If
    '                    ladrsEmpSchedLine.Fields("CHARGE_DATE") = !CHARGE_DATE
    '                    ladrsEmpSchedLine.Fields("DATE_WORKED") = !DATE_WORKED
    '                    ladrsEmpSchedLine.Fields("SHIFT_START_TIME") = !SHIFT_START_TIME
    '                    ladrsEmpSchedLine.Fields("SHIFT_START_DATE") = !SHIFT_START_DATE
    '                    ladrsEmpSchedLine.Fields("SHIFT_END_TIME") = !SHIFT_END_TIME
    '                    ladrsEmpSchedLine.Fields("SHIFT_END_DATE") = !SHIFT_END_DATE
    '                    ladrsEmpSchedLine.Fields("SHIFT_HOURS") = !SHIFT_HOURS
    '                    ladrsEmpSchedLine.Fields("START_TIME") = !START_TIME
    '                    ladrsEmpSchedLine.Fields("END_TIME") = !END_TIME
    '                    ladrsEmpSchedLine.Fields("HOURS") = !HOURS
    '                    ladrsEmpSchedLine.Fields("EXCEPTION_CD") = !EXCEPTION_CD
    '                    ladrsEmpSchedLine.Fields("OVER_EXC_CD") = !OVER_EXC_CD
    '                    ladrsEmpSchedLine.Fields("WG_NUM") = !WG_NUM
    '                    ladrsEmpSchedLine.Fields("WG_DESC") = lstrWGDesc
    '                    ladrsEmpSchedLine.Fields("SYSTEM_EXC_CD") = !SYSTEM_EXC_CD
    '                    ladrsEmpSchedLine.Fields("SHIFT_CODE") = !SHIFT_CODE
    '                    ladrsEmpSchedLine.Fields("OT_OPP_NUM") = !OT_OPP_NUM
    '                    ladrsEmpSchedLine.Fields("LINE_NUM") = !LINE_NUM
    '                    ladrsEmpSchedLine.Fields("SHIFT_START_END") = !SHIFT_START_END
    '                    ladrsEmpSchedLine.Fields("SHIFT_NUM") = !SHIFT_NUM
    '                    ladrsEmpSchedLine.Fields("SHIFT_LUNCH_START_TIME") = !SHIFT_LUNCH_START_TIME
    '                    ladrsEmpSchedLine.Fields("SHIFT_LUNCH_END_TIME") = !SHIFT_LUNCH_END_TIME
    '                    ladrsEmpSchedLine.Update
    '                End If
                    'get actual time
    '                If ladrsEmpSched!EMP_NUM <> llngPrevActimEmp Or _
    '                   ladrsEmpSched!CHARGE_DATE <> ldtPrevActimDate Then
    '                    ladrsActualTime.Filter = adFilterNone
    '                    ladrsActualTime.Filter = "EMP_NUM = " & ladrsEmpSched!EMP_NUM & _
    '                    " AND CHARGE_DATE = #" & ladrsEmpSched!CHARGE_DATE & "#"
    '                    While Not ladrsActualTime.EOF
    '                        If Not ladrsActualTime.EOF Then
    '                            If lblnDemandInd Then
    '                                ladrsEmpSchedLine.AddNew
    '                                ladrsEmpSchedLine.Fields("EMP_NUM") = !EMP_NUM
    '                                If mADOExtensions.RecordsExist(ladrsEmp) Then
    '                                    ladrsEmp.MoveFirst
    '                                    ladrsEmp.Find "EMP_NUM = " & !EMP_NUM
    '                                    If Not (ladrsEmp.EOF) Then
    '                                        ladrsEmpSchedLine.Fields("EMP_NAME") = ladrsEmp!EMP_NAME
    '                                    End If
    '                                End If
    '                                ladrsEmpSchedLine.Fields("CHARGE_DATE") = ladrsActualTime!CHARGE_DATE
    '                                ladrsEmpSchedLine.Fields("DATE_WORKED") = ladrsActualTime!DATE_WORKED
    '                                ladrsEmpSchedLine.Fields("SHIFT_START_TIME") = !SHIFT_START_TIME
    '                                ladrsEmpSchedLine.Fields("SHIFT_START_DATE") = !SHIFT_START_DATE
    '                                ladrsEmpSchedLine.Fields("SHIFT_END_TIME") = !SHIFT_END_TIME
    '                                ladrsEmpSchedLine.Fields("SHIFT_END_DATE") = !SHIFT_END_DATE
    '                                ladrsEmpSchedLine.Fields("SHIFT_HOURS") = !SHIFT_HOURS
    '                                ladrsEmpSchedLine.Fields("START_TIME") = ladrsActualTime!START_TIME
    '                                ladrsEmpSchedLine.Fields("END_TIME") = ladrsActualTime!END_TIME
    '                                ladrsEmpSchedLine.Fields("HOURS") = ladrsActualTime!HOURS
    '                                ladrsEmpSchedLine.Fields("EXCEPTION_CD") = ladrsActualTime!EXCEPTION_CD
    '                                ladrsEmpSchedLine.Fields("OVER_EXC_CD") = ladrsActualTime!OVER_EXC_CD
    '                                ladrsEmpSchedLine.Fields("WG_NUM") = !WG_NUM
    '                                ladrsEmpSchedLine.Fields("WG_DESC") = lstrWGDesc
    '                                ladrsEmpSchedLine.Fields("SYSTEM_EXC_CD") = !SYSTEM_EXC_CD
    '                                ladrsEmpSchedLine.Fields("SHIFT_CODE") = !SHIFT_CODE
    '                                ladrsEmpSchedLine.Fields("OT_OPP_NUM") = !OT_OPP_NUM
    '                                ladrsEmpSchedLine.Fields("LINE_NUM") = !LINE_NUM
    '                                ladrsEmpSchedLine.Fields("SHIFT_START_END") = !SHIFT_START_END
    '                                ladrsEmpSchedLine.Fields("SHIFT_NUM") = !SHIFT_NUM
    '                                ladrsEmpSchedLine.Fields("SHIFT_LUNCH_START_TIME") = !SHIFT_LUNCH_START_TIME
    '                                ladrsEmpSchedLine.Fields("SHIFT_LUNCH_END_TIME") = !SHIFT_LUNCH_END_TIME
    '                                ladrsEmpSchedLine.Update
    '                            Else
    '                                ladrsEmpSchedLine.AddNew
    '                                ladrsEmpSchedLine.Fields("EMP_NUM") = !EMP_NUM
    '                                If mADOExtensions.RecordsExist(ladrsEmp) Then
    '                                    ladrsEmp.MoveFirst
    '                                    ladrsEmp.Find "EMP_NUM = " & !EMP_NUM
    '                                    If Not (ladrsEmp.EOF) Then
    '                                        ladrsEmpSchedLine.Fields("EMP_NAME") = ladrsEmp!EMP_NAME
    '                                    End If
    '                                End If
    '                                ladrsEmpSchedLine.Fields("CHARGE_DATE") = ladrsActualTime!CHARGE_DATE
    '                                ladrsEmpSchedLine.Fields("DATE_WORKED") = ladrsActualTime!DATE_WORKED
    '                                ladrsEmpSchedLine.Fields("SHIFT_START_TIME") = !SHIFT_START_TIME
    '                                ladrsEmpSchedLine.Fields("SHIFT_START_DATE") = !SHIFT_START_DATE
    '                                ladrsEmpSchedLine.Fields("SHIFT_END_TIME") = !SHIFT_END_TIME
    '                                ladrsEmpSchedLine.Fields("SHIFT_END_DATE") = !SHIFT_END_DATE
    '                                ladrsEmpSchedLine.Fields("SHIFT_HOURS") = !SHIFT_HOURS
    '                                ladrsEmpSchedLine.Fields("START_TIME") = ladrsActualTime!START_TIME
    '                                ladrsEmpSchedLine.Fields("END_TIME") = ladrsActualTime!END_TIME
    '                                ladrsEmpSchedLine.Fields("HOURS") = ladrsActualTime!HOURS
    '                                ladrsEmpSchedLine.Fields("EXCEPTION_CD") = ladrsActualTime!EXCEPTION_CD
    '                                ladrsEmpSchedLine.Fields("OVER_EXC_CD") = ladrsActualTime!OVER_EXC_CD
    '                                ladrsEmpSchedLine.Fields("WG_NUM") = !WG_NUM
    '                                ladrsEmpSchedLine.Fields("WG_DESC") = lstrWGDesc
    '                                ladrsEmpSchedLine.Fields("SYSTEM_EXC_CD") = !SYSTEM_EXC_CD
    '                                ladrsEmpSchedLine.Fields("SHIFT_CODE") = !SHIFT_CODE
    '                                ladrsEmpSchedLine.Fields("SCHED_SRC_CD") = !SCHED_SRC_CD
    '                                ladrsEmpSchedLine.Fields("SCHED_CD") = !SCHED_CD
    '                                ladrsEmpSchedLine.Fields("OT_OPP_NUM") = !OT_OPP_NUM
    '                                ladrsEmpSchedLine.Fields("LINE_NUM") = !LINE_NUM
    '                                ladrsEmpSchedLine.Fields("SHIFT_START_END") = !SHIFT_START_END
    '                                ladrsEmpSchedLine.Fields("SHIFT_NUM") = !SHIFT_NUM
    '                                ladrsEmpSchedLine.Fields("SHIFT_LUNCH_START_TIME") = !SHIFT_LUNCH_START_TIME
    '                                ladrsEmpSchedLine.Fields("SHIFT_LUNCH_END_TIME") = !SHIFT_LUNCH_END_TIME
    '                                ladrsEmpSchedLine.Update
    '                            End If
    '                            llngPrevActimEmp = !EMP_NUM
    '                            ldtPrevActimDate = !CHARGE_DATE
    '                            ladrsActualTime.MoveNext
    '                        End If
    '                    Wend
    '                End If
    '                .MoveNext
    '            Wend
    '        End With
    '    Next llngIndex
    '    Set lobjSched = Nothing
    'End If
        
    'If mADOExtensions.RecordsExist(ladrsEmpSchedLine) Then
    '    Set GetDailyRosterSchedule2 = ladrsEmpSchedLine
    'Else
    '    Set GetDailyRosterSchedule2 = Nothing
    'End If

    'Set lobjSched = Nothing
    'Set lobjTimeCard = Nothing
    'Set lobjCodeTable = Nothing
    'Set lobjEmployee = Nothing
    
    'Set ladrsWG = Nothing
    'Set ladrsActualTime = Nothing
    'Set ladrsEmpSched = Nothing
    'Set ladrsEmpSchedLine = Nothing
    'Set ladrsUnmannedSched = Nothing
    'Set ladrsDummy = Nothing
    'Set ladrsEmp = Nothing
    'Set ladrsEmpsinWG = Nothing
    'Set ladrsEmpSched = Nothing

    'Exit Function
'GetDailyRosterSchedule2_eh:
    'Set lobjSched = Nothing
    'Set lobjTimeCard = Nothing
    'Set lobjCodeTable = Nothing
    'Set lobjEmployee = Nothing
    
    'Set ladrsWG = Nothing
    'Set ladrsActualTime = Nothing
    'Set ladrsEmpSched = Nothing
    'Set ladrsEmpSchedLine = Nothing
    'Set ladrsUnmannedSched = Nothing
    'Set ladrsDummy = Nothing
    'Set ladrsEmp = Nothing
    'Set ladrsEmpsinWG = Nothing
    'Set ladrsEmpSched = Nothing

    'TrsRaiseError m_strCMODULENAME, "GetDailyRosterSchedule2"

End Function
'Function gets WorkGroup Schedule for Daily Roster Report
Public Function GetDailyRosterSchedule3(ByVal vlngWGNum As Long, _
                                            ByVal varrEmployees As Variant, _
                                            ByVal vdtStartDate As Date, _
                                            ByVal vdtEndDate As Date) _
                                            As ADODB.Recordset

    On Error GoTo GetDailyRosterSchedule3_eh

    Dim ladrsWG As ADODB.Recordset
    Dim lobjWG As dbWorkGroup.cdWorkGroup
    Dim ladrsEmpsinWG As ADODB.Recordset
    Dim llngIndex As Long
    Dim llngPrevActimEmp As Long
    Dim ldtPrevActimDate As Date
    Dim llngEmpCount As Long
    Dim lobjSched As busSchedule.cbsSchedule
    Dim lobjTimeCard As dbTimeCard.cdTimeCard
    Dim lobjCodeTable As dbCodeTables.cdCodeTables
    Dim lobjEmployee As busEmployee.cbsEmployee
    Dim lobjdbEmployee As dbEmployee.cdEmployee

    Dim ladrsProjectInfo As ADODB.Recordset
    Dim ladrsActualTime As ADODB.Recordset
    Dim ladrsEmpSched As ADODB.Recordset
    Dim ladrsEmpSchedLine As ADODB.Recordset
    Dim ladrsUnmannedSched As ADODB.Recordset
    Dim ladrsDummy As ADODB.Recordset
    Dim lblnDemandInd As Boolean
    Dim lblnIsEmployeeInWG As Boolean
    Dim lstrWGDesc As String
    Dim ladrsEmp As ADODB.Recordset
    Dim llngarrEmps() As Long
    Dim lstrarrEmps As String
    Dim llngStart As Long
    Dim llngEnd As Long
    Dim i As Integer
    Const mconFIELDATTR = adFldUpdatable Or adFldIsNullable

    Set ladrsWG = GetWorkGroupAttributes(vlngWGNum)
    lblnDemandInd = False
    If mADOExtensions.RecordsExist(ladrsWG) Then
        If ladrsWG!DEMAND_IND = "Y" Then
            lblnDemandInd = True
        End If
        lstrWGDesc = ladrsWG!WG_DESC
    End If
    Set ladrsWG = Nothing
    
    'Set lobjWG = CreateObject("dbWorkGroup.cdWorkGroup")
    'Set ladrsEmpsinWG = lobjWG.GetEmployeesForDailyRosterRpt2(vlngWGNum, _
                                                             vdtStartDate, vdtEndDate)
    'Set lobjWG = Nothing

    'If mADOExtensions.RecordsExist(ladrsEmpsinWG) Then
        'With ladrsEmpsinWG
        llngStart = LBound(varrEmployees)
        llngEnd = UBound(varrEmployees)
            'ReDim llngarrEmps(1 To .RecordCount)
            '.Sort = "EMP_NUM"
            '.Filter = "EMP_NUM <> 0"
            '.MoveFirst
            'llngEmpCount = 1
            'For llngIndex = 1 To .RecordCount
            '    llngarrEmps(llngEmpCount) = !EMP_NUM
            '    If llngIndex = .RecordCount Then
            '        lstrarrEmps = lstrarrEmps & !EMP_NUM
            '    Else
            '        lstrarrEmps = lstrarrEmps & !EMP_NUM & ","
            '    End If
            '    llngEmpCount = llngEmpCount + 1
            '    .MoveNext
            'Next llngIndex
        'End With
        Set ladrsEmpSchedLine = New ADODB.Recordset
        With ladrsEmpSchedLine
            .Fields.Append "EMP_NUM", adInteger, , mconFIELDATTR
            .Fields.Append "EMP_NAME", adVarChar, 256, mconFIELDATTR
            .Fields.Append "CHARGE_DATE", adDate, , mconFIELDATTR
            .Fields.Append "DATE_WORKED", adDate, , mconFIELDATTR
            .Fields.Append "SHIFT_START_TIME", adChar, 5, mconFIELDATTR
            .Fields.Append "SHIFT_START_DATE", adDate, , mconFIELDATTR
            .Fields.Append "SHIFT_END_TIME", adChar, 5, mconFIELDATTR
            .Fields.Append "SHIFT_END_DATE", adDate, , mconFIELDATTR
            .Fields.Append "SHIFT_HOURS", adDecimal, , mconFIELDATTR
            .Fields.Append "START_TIME", adChar, 5, mconFIELDATTR
            .Fields.Append "END_TIME", adChar, 5, mconFIELDATTR
            .Fields.Append "HOURS", adDecimal, , mconFIELDATTR
            .Fields.Append "EXCEPTION_CD", adChar, 3, mconFIELDATTR
            .Fields.Append "OVER_EXC_CD", adChar, 3, mconFIELDATTR
            .Fields.Append "WG_NUM", adInteger, , mconFIELDATTR
            .Fields.Append "WG_DESC", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SYSTEM_EXC_CD", adChar, 3, mconFIELDATTR
            .Fields.Append "SHIFT_CODE", adChar, 1, mconFIELDATTR
            .Fields.Append "SCHED_SRC_CD", adInteger, , mconFIELDATTR
            .Fields.Append "SCHED_CD", adInteger, , mconFIELDATTR
            .Fields.Append "OT_OPP_NUM", adInteger, , mconFIELDATTR
            .Fields.Append "LINE_NUM", adInteger, , mconFIELDATTR
            .Fields.Append "SHIFT_START_END", adChar, 11, mconFIELDATTR
            .Fields.Append "SHIFT_NUM", adInteger, , mconFIELDATTR
            .Fields.Append "SHIFT_LUNCH_START_TIME", adChar, 5, mconFIELDATTR
            .Fields.Append "SHIFT_LUNCH_END_TIME", adChar, 5, mconFIELDATTR
            .CursorLocation = adUseClient
            .CursorType = adOpenStatic
            .LockType = adLockBatchOptimistic
            .Open
        End With
        
        Set lobjEmployee = CreateObject("busEmployee.cbsEmployee")
        Set lobjEmployee = New busEmployee.cbsEmployee
        Set ladrsEmp = lobjEmployee.GetEmployee(varrEmployees)
        Set lobjEmployee = Nothing

        llngPrevActimEmp = 0
        ldtPrevActimDate = "9999-12-31"
        
        Set lobjTimeCard = CreateObject("dbTimeCard.cdTimeCard")
        Set ladrsActualTime = lobjTimeCard.GetActualTimeEntriesWithoutOtherPay( _
                              varrEmployees, vdtStartDate, vdtEndDate)
        Set lobjTimeCard = Nothing

        Set lobjdbEmployee = CreateObject("dbEmployee.cdEmployee")
        
        Set lobjSched = CreateObject("busSchedule.cbsSchedule")
        'Set lobjSched = New busSchedule.cbsSchedule

        For llngIndex = llngStart To llngEnd
            Set ladrsEmpSched = lobjSched.GetAdjustedBaseSchedule(varrEmployees(llngIndex), vdtStartDate _
                              , vdtEndDate, ladrsDummy, False, trAdjBaseSchedWithExceptionsAndOvertime _
                              , trAllScheduleLevels, trRelief, trAllExceptions)
            Set ladrsDummy = Nothing

            With ladrsEmpSched
                .MoveFirst
                While Not ladrsEmpSched.EOF
                    lblnIsEmployeeInWG = lobjdbEmployee.IsEmployeeInWorkGroup(!EMP_NUM, _
                                                                                  vlngWGNum, _
                                                                                  !CHARGE_DATE)
                    If lblnIsEmployeeInWG Then
                        If lblnDemandInd Then
                            ladrsEmpSchedLine.AddNew
                            ladrsEmpSchedLine.Fields("EMP_NUM") = !EMP_NUM
                            If mADOExtensions.RecordsExist(ladrsEmp) Then
                                ladrsEmp.MoveFirst
                                ladrsEmp.Find "EMP_NUM = " & !EMP_NUM
                                If Not (ladrsEmp.EOF) Then
                                    ladrsEmpSchedLine.Fields("EMP_NAME") = ladrsEmp!EMP_NAME
                                End If
                            End If
                            ladrsEmpSchedLine.Fields("CHARGE_DATE") = !CHARGE_DATE
                            ladrsEmpSchedLine.Fields("DATE_WORKED") = !DATE_WORKED
                            ladrsEmpSchedLine.Fields("SHIFT_START_TIME") = !SHIFT_START_TIME
                            ladrsEmpSchedLine.Fields("SHIFT_START_DATE") = !SHIFT_START_DATE
                            ladrsEmpSchedLine.Fields("SHIFT_END_TIME") = !SHIFT_END_TIME
                            ladrsEmpSchedLine.Fields("SHIFT_END_DATE") = !SHIFT_END_DATE
                            ladrsEmpSchedLine.Fields("SHIFT_HOURS") = !SHIFT_HOURS
                            ladrsEmpSchedLine.Fields("START_TIME") = !START_TIME
                            ladrsEmpSchedLine.Fields("END_TIME") = !END_TIME
                            ladrsEmpSchedLine.Fields("HOURS") = !HOURS
                            ladrsEmpSchedLine.Fields("EXCEPTION_CD") = !EXCEPTION_CD
                            ladrsEmpSchedLine.Fields("OVER_EXC_CD") = !OVER_EXC_CD
                            ladrsEmpSchedLine.Fields("WG_NUM") = !WG_NUM
                            ladrsEmpSchedLine.Fields("WG_DESC") = lstrWGDesc
                            ladrsEmpSchedLine.Fields("SYSTEM_EXC_CD") = !SYSTEM_EXC_CD
                            ladrsEmpSchedLine.Fields("SHIFT_CODE") = !SHIFT_CODE
                            ladrsEmpSchedLine.Fields("SCHED_SRC_CD") = !SCHED_SRC_CD
                            ladrsEmpSchedLine.Fields("SCHED_CD") = !SCHED_CD
                            ladrsEmpSchedLine.Fields("OT_OPP_NUM") = !OT_OPP_NUM
                            ladrsEmpSchedLine.Fields("LINE_NUM") = !LINE_NUM
                            ladrsEmpSchedLine.Fields("SHIFT_START_END") = !SHIFT_START_END
                            ladrsEmpSchedLine.Fields("SHIFT_NUM") = !SHIFT_NUM
                            ladrsEmpSchedLine.Fields("SHIFT_LUNCH_START_TIME") = !SHIFT_LUNCH_START_TIME
                            ladrsEmpSchedLine.Fields("SHIFT_LUNCH_END_TIME") = !SHIFT_LUNCH_END_TIME
                            ladrsEmpSchedLine.Update
                        Else
                            ladrsEmpSchedLine.AddNew
                            ladrsEmpSchedLine.Fields("EMP_NUM") = !EMP_NUM
                            If mADOExtensions.RecordsExist(ladrsEmp) Then
                                ladrsEmp.MoveFirst
                                ladrsEmp.Find "EMP_NUM = " & !EMP_NUM
                                If Not (ladrsEmp.EOF) Then
                                    ladrsEmpSchedLine.Fields("EMP_NAME") = ladrsEmp!EMP_NAME
                                End If
                            End If
                            ladrsEmpSchedLine.Fields("CHARGE_DATE") = !CHARGE_DATE
                            ladrsEmpSchedLine.Fields("DATE_WORKED") = !DATE_WORKED
                            ladrsEmpSchedLine.Fields("SHIFT_START_TIME") = !SHIFT_START_TIME
                            ladrsEmpSchedLine.Fields("SHIFT_START_DATE") = !SHIFT_START_DATE
                            ladrsEmpSchedLine.Fields("SHIFT_END_TIME") = !SHIFT_END_TIME
                            ladrsEmpSchedLine.Fields("SHIFT_END_DATE") = !SHIFT_END_DATE
                            ladrsEmpSchedLine.Fields("SHIFT_HOURS") = !SHIFT_HOURS
                            ladrsEmpSchedLine.Fields("START_TIME") = !START_TIME
                            ladrsEmpSchedLine.Fields("END_TIME") = !END_TIME
                            ladrsEmpSchedLine.Fields("HOURS") = !HOURS
                            ladrsEmpSchedLine.Fields("EXCEPTION_CD") = !EXCEPTION_CD
                            ladrsEmpSchedLine.Fields("OVER_EXC_CD") = !OVER_EXC_CD
                            ladrsEmpSchedLine.Fields("WG_NUM") = !WG_NUM
                            ladrsEmpSchedLine.Fields("WG_DESC") = lstrWGDesc
                            ladrsEmpSchedLine.Fields("SYSTEM_EXC_CD") = !SYSTEM_EXC_CD
                            ladrsEmpSchedLine.Fields("SHIFT_CODE") = !SHIFT_CODE
                            ladrsEmpSchedLine.Fields("OT_OPP_NUM") = !OT_OPP_NUM
                            ladrsEmpSchedLine.Fields("LINE_NUM") = !LINE_NUM
                            ladrsEmpSchedLine.Fields("SHIFT_START_END") = !SHIFT_START_END
                            ladrsEmpSchedLine.Fields("SHIFT_NUM") = !SHIFT_NUM
                            ladrsEmpSchedLine.Fields("SHIFT_LUNCH_START_TIME") = !SHIFT_LUNCH_START_TIME
                            ladrsEmpSchedLine.Fields("SHIFT_LUNCH_END_TIME") = !SHIFT_LUNCH_END_TIME
                            ladrsEmpSchedLine.Update
                        End If
                        'get actual time
                        If ladrsEmpSched!EMP_NUM <> llngPrevActimEmp Or _
                           ladrsEmpSched!CHARGE_DATE <> ldtPrevActimDate Then
                            ladrsActualTime.Filter = adFilterNone
                            ladrsActualTime.Filter = "EMP_NUM = " & ladrsEmpSched!EMP_NUM & _
                            " AND CHARGE_DATE = #" & ladrsEmpSched!CHARGE_DATE & "#"
                            While Not ladrsActualTime.EOF
                                If Not ladrsActualTime.EOF Then
                                    lblnIsEmployeeInWG = lobjdbEmployee.IsEmployeeInWorkGroup(!EMP_NUM, _
                                                              vlngWGNum, _
                                                              !CHARGE_DATE)
                                    If lblnIsEmployeeInWG Then
                                        If lblnDemandInd Then
                                            ladrsEmpSchedLine.AddNew
                                            ladrsEmpSchedLine.Fields("EMP_NUM") = !EMP_NUM
                                            If mADOExtensions.RecordsExist(ladrsEmp) Then
                                                ladrsEmp.MoveFirst
                                                ladrsEmp.Find "EMP_NUM = " & !EMP_NUM
                                                If Not (ladrsEmp.EOF) Then
                                                    ladrsEmpSchedLine.Fields("EMP_NAME") = ladrsEmp!EMP_NAME
                                                End If
                                            End If
                                            ladrsEmpSchedLine.Fields("CHARGE_DATE") = ladrsActualTime!CHARGE_DATE
                                            ladrsEmpSchedLine.Fields("DATE_WORKED") = ladrsActualTime!DATE_WORKED
                                            ladrsEmpSchedLine.Fields("SHIFT_START_TIME") = !SHIFT_START_TIME
                                            ladrsEmpSchedLine.Fields("SHIFT_START_DATE") = !SHIFT_START_DATE
                                            ladrsEmpSchedLine.Fields("SHIFT_END_TIME") = !SHIFT_END_TIME
                                            ladrsEmpSchedLine.Fields("SHIFT_END_DATE") = !SHIFT_END_DATE
                                            ladrsEmpSchedLine.Fields("SHIFT_HOURS") = !SHIFT_HOURS
                                            ladrsEmpSchedLine.Fields("START_TIME") = ladrsActualTime!START_TIME
                                            ladrsEmpSchedLine.Fields("END_TIME") = ladrsActualTime!END_TIME
                                            ladrsEmpSchedLine.Fields("HOURS") = ladrsActualTime!HOURS
                                            ladrsEmpSchedLine.Fields("EXCEPTION_CD") = ladrsActualTime!EXCEPTION_CD
                                            ladrsEmpSchedLine.Fields("OVER_EXC_CD") = ladrsActualTime!OVER_EXC_CD
                                            ladrsEmpSchedLine.Fields("WG_NUM") = !WG_NUM
                                            ladrsEmpSchedLine.Fields("WG_DESC") = lstrWGDesc
                                            ladrsEmpSchedLine.Fields("SYSTEM_EXC_CD") = !SYSTEM_EXC_CD
                                            ladrsEmpSchedLine.Fields("SHIFT_CODE") = !SHIFT_CODE
                                            ladrsEmpSchedLine.Fields("OT_OPP_NUM") = !OT_OPP_NUM
                                            ladrsEmpSchedLine.Fields("LINE_NUM") = !LINE_NUM
                                            ladrsEmpSchedLine.Fields("SHIFT_START_END") = !SHIFT_START_END
                                            ladrsEmpSchedLine.Fields("SHIFT_NUM") = !SHIFT_NUM
                                            ladrsEmpSchedLine.Fields("SHIFT_LUNCH_START_TIME") = !SHIFT_LUNCH_START_TIME
                                            ladrsEmpSchedLine.Fields("SHIFT_LUNCH_END_TIME") = !SHIFT_LUNCH_END_TIME
                                            ladrsEmpSchedLine.Update
                                        Else
                                            ladrsEmpSchedLine.AddNew
                                            ladrsEmpSchedLine.Fields("EMP_NUM") = !EMP_NUM
                                            If mADOExtensions.RecordsExist(ladrsEmp) Then
                                                ladrsEmp.MoveFirst
                                                ladrsEmp.Find "EMP_NUM = " & !EMP_NUM
                                                If Not (ladrsEmp.EOF) Then
                                                    ladrsEmpSchedLine.Fields("EMP_NAME") = ladrsEmp!EMP_NAME
                                                End If
                                            End If
                                            ladrsEmpSchedLine.Fields("CHARGE_DATE") = ladrsActualTime!CHARGE_DATE
                                            ladrsEmpSchedLine.Fields("DATE_WORKED") = ladrsActualTime!DATE_WORKED
                                            ladrsEmpSchedLine.Fields("SHIFT_START_TIME") = !SHIFT_START_TIME
                                            ladrsEmpSchedLine.Fields("SHIFT_START_DATE") = !SHIFT_START_DATE
                                            ladrsEmpSchedLine.Fields("SHIFT_END_TIME") = !SHIFT_END_TIME
                                            ladrsEmpSchedLine.Fields("SHIFT_END_DATE") = !SHIFT_END_DATE
                                            ladrsEmpSchedLine.Fields("SHIFT_HOURS") = !SHIFT_HOURS
                                            ladrsEmpSchedLine.Fields("START_TIME") = ladrsActualTime!START_TIME
                                            ladrsEmpSchedLine.Fields("END_TIME") = ladrsActualTime!END_TIME
                                            ladrsEmpSchedLine.Fields("HOURS") = ladrsActualTime!HOURS
                                            ladrsEmpSchedLine.Fields("EXCEPTION_CD") = ladrsActualTime!EXCEPTION_CD
                                            ladrsEmpSchedLine.Fields("OVER_EXC_CD") = ladrsActualTime!OVER_EXC_CD
                                            ladrsEmpSchedLine.Fields("WG_NUM") = !WG_NUM
                                            ladrsEmpSchedLine.Fields("WG_DESC") = lstrWGDesc
                                            ladrsEmpSchedLine.Fields("SYSTEM_EXC_CD") = !SYSTEM_EXC_CD
                                            ladrsEmpSchedLine.Fields("SHIFT_CODE") = !SHIFT_CODE
                                            ladrsEmpSchedLine.Fields("SCHED_SRC_CD") = !SCHED_SRC_CD
                                            ladrsEmpSchedLine.Fields("SCHED_CD") = !SCHED_CD
                                            ladrsEmpSchedLine.Fields("OT_OPP_NUM") = !OT_OPP_NUM
                                            ladrsEmpSchedLine.Fields("LINE_NUM") = !LINE_NUM
                                            ladrsEmpSchedLine.Fields("SHIFT_START_END") = !SHIFT_START_END
                                            ladrsEmpSchedLine.Fields("SHIFT_NUM") = !SHIFT_NUM
                                            ladrsEmpSchedLine.Fields("SHIFT_LUNCH_START_TIME") = !SHIFT_LUNCH_START_TIME
                                            ladrsEmpSchedLine.Fields("SHIFT_LUNCH_END_TIME") = !SHIFT_LUNCH_END_TIME
                                            ladrsEmpSchedLine.Update
                                        End If
                                    End If
                                    llngPrevActimEmp = !EMP_NUM
                                    ldtPrevActimDate = !CHARGE_DATE
                                    ladrsActualTime.MoveNext
                                End If
                            Wend
                        End If
                    End If
                    .MoveNext
                Wend
            End With
        Next llngIndex
        Set lobjSched = Nothing
        Set lobjdbEmployee = Nothing
    'End If
        
    If mADOExtensions.RecordsExist(ladrsEmpSchedLine) Then
        Set GetDailyRosterSchedule3 = ladrsEmpSchedLine
    Else
        Set GetDailyRosterSchedule3 = Nothing
    End If

    Set lobjSched = Nothing
    Set lobjTimeCard = Nothing
    Set lobjCodeTable = Nothing
    Set lobjEmployee = Nothing
    Set lobjdbEmployee = Nothing
    
    Set ladrsWG = Nothing
    Set ladrsActualTime = Nothing
    Set ladrsEmpSched = Nothing
    Set ladrsEmpSchedLine = Nothing
    Set ladrsUnmannedSched = Nothing
    Set ladrsDummy = Nothing
    Set ladrsEmp = Nothing
    Set ladrsEmpsinWG = Nothing
    Set ladrsEmpSched = Nothing

    Exit Function
GetDailyRosterSchedule3_eh:
    Set lobjSched = Nothing
    Set lobjTimeCard = Nothing
    Set lobjCodeTable = Nothing
    Set lobjEmployee = Nothing
    Set lobjdbEmployee = Nothing
    
    Set ladrsWG = Nothing
    Set ladrsActualTime = Nothing
    Set ladrsEmpSched = Nothing
    Set ladrsEmpSchedLine = Nothing
    Set ladrsUnmannedSched = Nothing
    Set ladrsDummy = Nothing
    Set ladrsEmp = Nothing
    Set ladrsEmpsinWG = Nothing
    Set ladrsEmpSched = Nothing

    TrsRaiseError m_strCMODULENAME, "GetDailyRosterSchedule3"

End Function


Public Function GetJobDesc(ByVal vlngJobNum) As String
On Error GoTo GetJobDesc_EH

    Dim lobjJobInfo As dbCodeTables.cdCodeTables
    Dim ladrsJobInfo As ADODB.Recordset

    Set lobjJobInfo = CreateObject("dbCodeTables.cdCodeTables")
    Set ladrsJobInfo = lobjJobInfo.GetJob(vlngJobNum)
    Set lobjJobInfo = Nothing

    If mADOExtensions.RecordsExist(ladrsJobInfo) Then
        GetJobDesc = ladrsJobInfo!JOB_DESC
    Else
        GetJobDesc = ""
    End If

    Set ladrsJobInfo = Nothing

    Exit Function
GetJobDesc_EH:
    Set lobjJobInfo = Nothing
    Set ladrsJobInfo = Nothing

    TrsRaiseError m_strCMODULENAME, "GetJobDesc"

End Function



Public Function IsJobTaskWorkgroup(ByVal v_lngWgNum As Long) As Boolean
    
'Does the workgroup use EMPRV JobTasks?
    
    Dim l_objWkGp As dbWorkGroup.cdWorkGroup
    
    Dim l_objBln As Boolean
    
    On Error GoTo IsJobTaskWorkgroup_EH
        
    Set l_objWkGp = CreateObject("dbWorkGroup.cdWorkGroup")

    l_objBln = l_objWkGp.IsJobTaskWorkgroup(v_lngWgNum)
    IsJobTaskWorkgroup = l_objBln
   
    Set l_objWkGp = Nothing
    
    Exit Function

IsJobTaskWorkgroup_EH:
    Set l_objWkGp = Nothing
    
    TrsRaiseError m_strCMODULENAME, "IsJobTaskWorkgroup", , , , , True

End Function

