VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "rptException"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim m_dtStartDate As Date, m_dtEndDate As Date

Dim m_adrsSelectedEmployees As ADODB.Recordset
Dim m_adrsSelectedDepartments As ADODB.Recordset
Dim m_adrsSelectedWorkgroups As ADODB.Recordset
Dim m_adrsSelectedJobs As ADODB.Recordset

'Supporting Information
Dim m_adrsExceptionCodes As ADODB.Recordset

'Currently being printed
Dim m_lProcessLevel As Long
Dim m_lWorkgroup As Long, m_lCurrentEmpNum As Long
Dim m_sDeptNum As String, m_sJobNum As String

'Recordset for report totals
Dim m_adrsWGTotals As ADODB.Recordset
Dim m_adrsEmpTotals As ADODB.Recordset
Dim m_adrsDeptTotals As ADODB.Recordset
Dim m_adrsJobTotals As ADODB.Recordset
Dim m_adrsEmpVacBanks As ADODB.Recordset

'Flags for department level report
Dim m_bActualTimeDetails As Boolean, m_bSchedEmpDetails As Boolean

'Define the main report
Dim m_adrsReportData As ADODB.Recordset
Dim WithEvents rptMain As ActiveReport
Dim m_lReportLevel As Long

'Sub Reports
Dim WithEvents rptSubReport1 As ActiveReport
Dim WithEvents rptSubReport2 As ActiveReport
Dim WithEvents rptSubReport3 As ActiveReport

Dim m_ActiveReports_ErrNumber As Integer
Dim m_ActiveReports_ErrDescription As String
Dim m_ActiveReports_ErrSource As String
Dim m_ActiveReports_OrigErrDescription As String
Dim m_ActiveReports_OrigErrSource As String

'WG Schedule report variables
Dim m_adrsWorkgroups As ADODB.Recordset

'Public l_adrsTemp As ADODB.Recordset
'Dim m_arrDays() As Variant
'Dim m_adrsEmployeeSchedule As ADODB.Recordset
'Dim m_iNumberOfDays As Integer
'Dim m_adrsEmployees As ADODB.Recordset
'Dim m_adrsAllDepartments As ADODB.Recordset
'Dim m_adrsAllJobs As ADODB.Recordset
'Dim m_adrsAllEmployees As ADODB.Recordset

Public l_adrsTemp As ADODB.Recordset

Private Const m_conFIELDATTR = adFldUpdatable Or adFldIsNullable
Private Const MODULENAME As String = "rptEngine.rptException"
Public Property Get ReportLevel() As Long
    ReportLevel = m_lReportLevel
End Property
Public Property Get JobNumber() As String
    JobNumber = m_sJobNum
End Property
Public Property Let JobNumber(ByVal v_sValue As String)
    m_sJobNum = v_sValue
End Property
Public Property Get Department() As String
    Department = m_sDeptNum
End Property
Public Property Let Department(ByVal v_sValue As String)
    m_sDeptNum = v_sValue
End Property
Public Property Get Workgroup() As Long
    Workgroup = m_lWorkgroup
End Property
Public Property Let Workgroup(ByVal v_lValue As Long)
    m_lWorkgroup = v_lValue
End Property
Public Property Get Employee() As Long
    Employee = m_lCurrentEmpNum
End Property
Public Property Let Employee(ByVal v_lValue As Long)
    m_lCurrentEmpNum = v_lValue
End Property
Public Property Get ProcessLevel() As Long
    ProcessLevel = m_lProcessLevel
End Property
Public Property Let ProcessLevel(ByVal v_lValue As Long)
    m_lProcessLevel = v_lValue
End Property
'Input Parameters:
' - v_sReportPath:  File for the report's xml definition
' - v_dtStartDate:  User entered Start Date
' - v_dtEndDate:  User entered End Date
' - v_vReportLevel:
'       1:  Employee
'       2:  Workgroup
'       3:  Department
'       4:  Job Code/Craft
' - v_adrsParametersSelected: User selection of employees, WGs, Depts, or Jobs
' - v_adrsSupportingInfo1:  Employees' information from the client
' - v_adrsSupportingInfo2:  Recordset with WGs, dept or Jobs information

Public Function GetExceptionEmpRpt( _
        ByVal v_dtStartDate As Date _
        , ByVal v_dtEndDate As Date _
        , ByVal v_adrsSelectedEmployees As ADODB.Recordset _
        , ByVal v_adrsSelectedExceptionCodes As ADODB.Recordset _
        , ByVal v_adrsAllExceptionCodes As ADODB.Recordset _
        , ByVal v_bPrintExceptionCodes As Boolean _
        , ByRef r_bReportGenerated As Boolean _
        , ByVal v_lpapersize) As ActiveReport
On Error GoTo GetExceptionEmpRpt_EH
Dim lsFooter As String

    'Initialize return
    r_bReportGenerated = False
    ResetActiveReportsError
    Set GetExceptionEmpRpt = Nothing

    'Check for existence
    If (RecordsExist(v_adrsSelectedEmployees)) Then
        Set m_adrsSelectedEmployees = v_adrsSelectedEmployees
    Else
        Exit Function
    End If

    If (RecordsExist(v_adrsAllExceptionCodes)) Then
        Set m_adrsExceptionCodes = v_adrsAllExceptionCodes
    Else
        Exit Function
    End If

    Set rptMain = CreateObject("DDActiveReports2.ActiveReport")
    
    'sometimes receive error - Method '~' of object '~' failed
    'rptMain.Printer.PaperSize = v_lpapersize

    'Set module level variables
    m_lReportLevel = 1
    m_dtStartDate = v_dtStartDate
    m_dtEndDate = v_dtEndDate

    'Load XML file into the report object
    rptMain.LoadLayout "C:\TRIS Solution\TRISReports\RPX Files\rptExceptionEmp.rpx"
    'Add a reference in the rpt object to this (rptException) class
    Call rptMain.AddNamedItem("MyClass", Me)

    Set rptSubReport1 = CreateObject("DDActiveReports2.ActiveReport")
    rptSubReport1.LoadLayout "C:\TRIS Solution\TRISReports\RPX Files\rptExceptionTotalsSub.rpx"
    'Add a reference in the rpt object to this (rptException) class
    Call rptSubReport1.AddNamedItem("MyClass", Me)

    Set rptSubReport2 = CreateObject("DDActiveReports2.ActiveReport")
    rptSubReport2.LoadLayout "C:\TRIS Solution\TRISReports\RPX Files\rptExceptionEmpBanks.rpx"
    'Add a reference in the rpt object to this (rptException) class
    Call rptSubReport2.AddNamedItem("MyClass", Me)

    rptMain.Sections("GroupFooter1").Controls("SubReport1").object = _
        rptSubReport1
    rptMain.Sections("GroupFooter1").Controls("SubReport2").object = _
        rptSubReport2
    rptMain.Sections("PageHeader").Controls("txtHdrRunDate").Text = _
        Now()
    rptMain.Sections("PageHeader").Controls("txtHdrStartDate").Text = _
        CStr(Format(v_dtStartDate, "mm/dd/yyyy"))
    rptMain.Sections("PageHeader").Controls("txtHdrEndDate").Text = _
        CStr(Format(v_dtEndDate, "mm/dd/yyyy"))

    lsFooter = ""
    If (v_bPrintExceptionCodes) Then
        lsFooter = GetFooter(v_adrsSelectedExceptionCodes)
    End If
    rptMain.Sections("PageFooter").Controls("txtExceptionCodeFooter").Text = _
        lsFooter

    Set m_adrsReportData = GetEmpExceptionReportData( _
                            m_adrsSelectedEmployees _
                            , v_dtStartDate _
                            , v_dtEndDate _
                            , v_adrsSelectedExceptionCodes)
    Set m_adrsEmpVacBanks = GetEmpVacationBanks(m_adrsSelectedEmployees)

    If RecordsExist(m_adrsReportData) Then
        DefineEmpTotals

        With m_adrsReportData
            .Sort = "EMP_NUM, EXCEPTION_CD, OVER_EXC_CD, SCHED_CHG_EXC_CD"
            .MoveFirst

            Do While Not .EOF
                ComputeEmpTotals "" _
                    , 0 _
                    , !EMP_NUM _
                    , Trim(!EXCEPTION_CD) _
                    , Trim(!OVER_EXC_CD) _
                    , Trim(!SCHED_CHG_EXC_CD) _
                    , Trim(!ADJ_TYPE) _
                    , !Hours _
                    , !AMOUNT _
                    , !TRAVEL_PAY_HRS _
                    , Trim(!PROCESSED_IND)

                .MoveNext
            Loop

            .Sort = "EMP_NUM, EXCEPTION_CD, OVER_EXC_CD, SCHED_CHG_EXC_CD, CHARGE_DATE, ADJ_TYPE"
            .MoveFirst
        End With

        r_bReportGenerated = True
        rptMain.Run

        '//check for any errors from the active reports script or the vb6 functions being called back
        If m_ActiveReports_ErrNumber <> 0 Then
            Err.Raise m_ActiveReports_ErrNumber, m_ActiveReports_ErrSource, m_ActiveReports_ErrDescription
        End If
        
        'Return the report object to the calling application
        Set GetExceptionEmpRpt = rptMain
    End If

    '//-----clean up public properties used by ActiveReports-----//
    DestroyState
    
    Exit Function
GetExceptionEmpRpt_EH:
    DestroyState
    
    Select Case Err.Number
    Case m_ActiveReports_ErrNumber
        TrsRaiseError MODULENAME, "GetExceptionEmpRpt", Err.Number, , , _
            DescArgsFrom("[OrigErrDescription]===>" & m_ActiveReports_OrigErrDescription & "<===", "[OrigErrSource]===>" & m_ActiveReports_OrigErrSource & "<==="), True
    Case Else
        TrsRaiseError MODULENAME, "GetExceptionEmpRpt", , , , , True
    End Select
End Function

Public Function GetExceptionWGRpt( _
        ByVal v_dtStartDate As Date _
        , ByVal v_dtEndDate As Date _
        , ByVal v_adrsWorkgroupsSelected As ADODB.Recordset _
        , ByVal v_adrsSelectedExceptionCodes As ADODB.Recordset _
        , ByVal v_adrsAllExceptionCodes As ADODB.Recordset _
        , ByVal v_bPrintExceptionCodes As Boolean _
        , ByRef r_bReportGenerated As Boolean) As ActiveReport
On Error GoTo GetExceptionWGRpt_EH
Dim lsFooter As String

    'Initialize return
    r_bReportGenerated = False
    ResetActiveReportsError
    Set GetExceptionWGRpt = Nothing

    If (RecordsExist(v_adrsWorkgroupsSelected)) Then
        Set m_adrsSelectedWorkgroups = v_adrsWorkgroupsSelected
    Else
        Exit Function
    End If

    If (RecordsExist(v_adrsAllExceptionCodes)) Then
        Set m_adrsExceptionCodes = v_adrsAllExceptionCodes
    Else
        Exit Function
    End If

    m_lReportLevel = 2

    m_dtStartDate = v_dtStartDate
    m_dtEndDate = v_dtEndDate

    Set m_adrsWorkgroups = v_adrsWorkgroupsSelected.Clone

    'Load XML file into the report object
    Set rptMain = CreateObject("DDActiveReports2.ActiveReport")
    rptMain.LoadLayout "C:\TRIS Solution\TRISReports\RPX Files\rptExceptionWG.rpx"
    'Add a reference in the rpt object to this (rptException) class
    Call rptMain.AddNamedItem("MyClass", Me)

    Set rptSubReport1 = CreateObject("DDActiveReports2.ActiveReport")
    rptSubReport1.LoadLayout "C:\TRIS Solution\TRISReports\RPX Files\rptExceptionTotalsSub.rpx"
    'Add a reference in the rpt object to this (rptException) class
    Call rptSubReport1.AddNamedItem("MyClass", Me)

    Set rptSubReport2 = CreateObject("DDActiveReports2.ActiveReport")
    rptSubReport2.LoadLayout "C:\TRIS Solution\TRISReports\RPX Files\rptExceptionTotalsSub.rpx"
    'Add a reference in the rpt object to this (rptException) class
    Call rptSubReport2.AddNamedItem("MyClass", Me)

    rptMain.Sections("GroupFooter1").Controls("subReport1").object = _
        rptSubReport1
    rptMain.Sections("GroupFooter2").Controls("subReport2").object = _
        rptSubReport2
    rptMain.Sections("PageHeader").Controls("txtHdrRunDate").Text = _
        Now()
    rptMain.Sections("PageHeader").Controls("txtHdrStartDate").Text = _
        CStr(Format(v_dtStartDate, "mm/dd/yyyy"))
    rptMain.Sections("PageHeader").Controls("txtHdrEndDate").Text = _
        CStr(Format(v_dtEndDate, "mm/dd/yyyy"))

    lsFooter = ""
    If (v_bPrintExceptionCodes) Then
        lsFooter = GetFooter(v_adrsSelectedExceptionCodes)
    End If
    rptMain.Sections("PageFooter").Controls("txtExceptionCodeFooter").Text = _
        lsFooter

    Set m_adrsReportData = GetWGExceptionReportData( _
            m_adrsSelectedWorkgroups _
            , v_dtStartDate _
            , v_dtEndDate _
            , v_adrsSelectedExceptionCodes)

    If (RecordsExist(m_adrsReportData)) Then
        DefineWGTotals
        DefineEmpTotals

        With m_adrsReportData
            .Sort = "WG_NUM, EXCEPTION_CD, OVER_EXC_CD, SCHED_CHG_EXC_CD"
            .MoveFirst

            Do While Not .EOF
                ComputeWGTotals "" _
                    , !WG_NUM _
                    , Trim(!EXCEPTION_CD) _
                    , Trim(!OVER_EXC_CD) _
                    , Trim(!SCHED_CHG_EXC_CD) _
                    , Trim(!ADJ_TYPE) _
                    , !Hours _
                    , !AMOUNT _
                    , !TRAVEL_PAY_HRS _
                    , Trim(!PROCESSED_IND)

                ComputeEmpTotals "" _
                    , !WG_NUM _
                    , !EMP_NUM _
                    , Trim(!EXCEPTION_CD) _
                    , Trim(!OVER_EXC_CD) _
                    , Trim(!SCHED_CHG_EXC_CD) _
                    , Trim(!ADJ_TYPE) _
                    , !Hours _
                    , !AMOUNT _
                    , !TRAVEL_PAY_HRS _
                    , Trim(!PROCESSED_IND)

                .MoveNext
            Loop

            .Sort = "WG_NUM, EMP_NUM, EXCEPTION_CD, OVER_EXC_CD, SCHED_CHG_EXC_CD, CHARGE_DATE, ADJ_TYPE"
            .MoveFirst
        End With

        r_bReportGenerated = True
        rptMain.Run

        '//check for any errors from the active reports script or the vb6 functions being called back
        If m_ActiveReports_ErrNumber <> 0 Then
            Err.Raise m_ActiveReports_ErrNumber, m_ActiveReports_ErrSource, m_ActiveReports_ErrDescription
        End If
        
        'Return the report object to the calling application
        Set GetExceptionWGRpt = rptMain
    End If

    '//-----clean up public properties used by ActiveReports-----//
    DestroyState
    
    Exit Function
GetExceptionWGRpt_EH:
    DestroyState
    
    Select Case Err.Number
    Case m_ActiveReports_ErrNumber
        TrsRaiseError MODULENAME, "GetExceptionWGRpt", Err.Number, , , _
            DescArgsFrom("[OrigErrDescription]===>" & m_ActiveReports_OrigErrDescription & "<===", "[OrigErrSource]===>" & m_ActiveReports_OrigErrSource & "<==="), True
    Case Else
        TrsRaiseError MODULENAME, "GetExceptionWGRpt", , , , , True
    End Select
End Function

Public Function GetExceptionDeptRpt2( _
        ByVal v_dtStartDate As Date _
        , ByVal v_dtEndDate As Date _
        , ByVal v_adrsSelectedDepartments As ADODB.Recordset _
        , ByVal v_adrsSelectedExceptionCodes As ADODB.Recordset _
        , ByVal v_adrsAllExceptionCodes As ADODB.Recordset _
        , ByVal v_bPrintExceptionCodes As Boolean _
        , ByVal v_adrsWorkgroups As ADODB.Recordset _
        , ByVal v_bActualTimeDetails As Boolean _
        , ByVal v_bSchedEmpDetails As Boolean _
        , ByVal v_bEmpDetails As Boolean _
        , ByRef r_bReportGenerated As Boolean _
        ) As ActiveReport
On Error GoTo GetExceptionDeptRpt2_EH
Dim lsFooter As String

    'Initialize return
    r_bReportGenerated = False
    ResetActiveReportsError
    Set GetExceptionDeptRpt2 = Nothing

    If (RecordsExist(v_adrsSelectedDepartments)) Then
        Set m_adrsSelectedDepartments = v_adrsSelectedDepartments
    Else
        Exit Function
    End If

    If (RecordsExist(v_adrsAllExceptionCodes)) Then
        Set m_adrsExceptionCodes = v_adrsAllExceptionCodes
    Else
        Exit Function
    End If

    m_lReportLevel = 3
    m_dtStartDate = v_dtStartDate
    m_dtEndDate = v_dtEndDate

    Set m_adrsWorkgroups = v_adrsWorkgroups


    'Load XML file into the report object
    Set rptMain = CreateObject("DDActiveReports2.ActiveReport")
    rptMain.LoadLayout "C:\TRIS Solution\TRISReports\RPX Files\rptExceptionDept.rpx"
    'Add a reference in the rpt object to this (rptException) class
    Call rptMain.AddNamedItem("MyClass", Me)

    Set rptSubReport1 = CreateObject("DDActiveReports2.ActiveReport")
    rptSubReport1.LoadLayout "C:\TRIS Solution\TRISReports\RPX Files\rptExceptionTotalsSub.rpx"
    'Add a reference in the rpt object to this (rptException) class
    Call rptSubReport1.AddNamedItem("MyClass", Me)

    Set rptSubReport2 = CreateObject("DDActiveReports2.ActiveReport")
    rptSubReport2.LoadLayout "C:\TRIS Solution\TRISReports\RPX Files\rptExceptionTotalsSub.rpx"
    'Add a reference in the rpt object to this (rptException) class
    Call rptSubReport2.AddNamedItem("MyClass", Me)

    Set rptSubReport3 = CreateObject("DDActiveReports2.ActiveReport")
    rptSubReport3.LoadLayout "C:\TRIS Solution\TRISReports\RPX Files\rptExceptionTotalsSub.rpx"
    'Add a reference in the rpt object to this (rptException) class
    Call rptSubReport3.AddNamedItem("MyClass", Me)

    rptMain.Sections("GroupFooter1").Controls("SubReportDeptTotals").object = _
        rptSubReport1
    'rptMain.Sections("GroupFooter2").Controls("SubReportWGTotals").object = _
    '    rptSubReport2
    rptMain.Sections("GroupFooter3").Controls("SubReportEmpTotals").object = _
        rptSubReport3
    rptMain.Sections("PageHeader").Controls("txtHdrRunDate").Text = _
        Now()
    rptMain.Sections("PageHeader").Controls("txtHdrStartDate").Text = _
        CStr(Format(v_dtStartDate, "mm/dd/yyyy"))
    rptMain.Sections("PageHeader").Controls("txtHdrEndDate").Text = _
        CStr(Format(v_dtEndDate, "mm/dd/yyyy"))
    lsFooter = ""
    If (v_bPrintExceptionCodes) Then
        lsFooter = GetFooter(v_adrsSelectedExceptionCodes)
    End If
    rptMain.Sections("PageFooter").Controls("txtExceptionCodeFooter").Text = _
        lsFooter

    m_bActualTimeDetails = v_bActualTimeDetails
    m_bSchedEmpDetails = v_bSchedEmpDetails

    If (v_bEmpDetails) Then
        rptMain.Sections("Detail").Visible = True
        rptMain.Sections("GroupHeader3").Visible = True
        rptMain.Sections("GroupFooter3").Visible = True
    Else
         rptMain.Sections("Detail").Visible = False
         rptMain.Sections("GroupHeader3").Visible = False
         rptMain.Sections("GroupHeader4").Visible = False
         rptMain.Sections("GroupFooter3").Visible = False
    End If

    rptMain.Sections("PageHeader").Controls("lblReportHeader3").Caption = ""
    If (m_bActualTimeDetails) Then
        If Not (m_bSchedEmpDetails) Then
            rptMain.Sections("PageHeader").Controls("lblReportHeader3").Caption = _
                "Actual Time Exceptions"
        End If
    Else
        If (m_bSchedEmpDetails) Then
            rptMain.Sections("PageHeader").Controls("lblReportHeader3").Caption = _
                "Scheduled Exceptions"
        End If
    End If

    Set m_adrsReportData = GetDeptExceptions( _
             m_adrsSelectedDepartments _
             , v_dtStartDate _
             , v_dtEndDate _
             , v_adrsSelectedExceptionCodes)

    If (RecordsExist(m_adrsReportData)) Then
        DefineEmpTotals
        DefineDeptTotals
        DefineWGTotals

        With m_adrsReportData
            .Sort = "DEPT_NUM, EXCEPTION_CD, OVER_EXC_CD, SCHED_CHG_EXC_CD"
            .MoveFirst

            If (m_bSchedEmpDetails And Not m_bActualTimeDetails) Then
                m_adrsReportData.Filter = MDAC25FilterFix("PROCESSED_IND = 1")
            ElseIf (m_bActualTimeDetails And Not m_bSchedEmpDetails) Then
                m_adrsReportData.Filter = MDAC25FilterFix("PROCESSED_IND > 1")
            Else
                m_adrsReportData.Filter = adFilterNone
            End If

            Do While Not .EOF
                ComputeDeptTotals _
                    Trim(!DEPT_NUM) _
                    , Trim(!EXCEPTION_CD) _
                    , Trim(!OVER_EXC_CD) _
                    , Trim(!SCHED_CHG_EXC_CD) _
                    , Trim(!ADJ_TYPE) _
                    , !Hours _
                    , !AMOUNT _
                    , !TRAVEL_PAY_HRS _
                    , Trim(!PROCESSED_IND)

                'ComputeWGTotals _
                '    Trim(!DEPT_NUM) _
                '    , !WG_NUM _
                '    , Trim(!EXCEPTION_CD) _
                '    , Trim(!OVER_EXC_CD) _
                '    , Trim(!SCHED_CHG_EXC_CD) _
                '    , Trim(!ADJ_TYPE) _
                '    , !Hours _
                '    , !AMOUNT _
                '    , Trim(!PROCESSED_IND)

                ComputeEmpTotals _
                    Trim(!DEPT_NUM) _
                    , 0 _
                    , !EMP_NUM _
                    , Trim(!EXCEPTION_CD) _
                    , Trim(!OVER_EXC_CD) _
                    , Trim(!SCHED_CHG_EXC_CD) _
                    , Trim(!ADJ_TYPE) _
                    , !Hours _
                    , !AMOUNT _
                    , !TRAVEL_PAY_HRS _
                    , Trim(!PROCESSED_IND)

                .MoveNext
            Loop

            If (RecordsExist(m_adrsReportData)) Then
                .Sort = "DEPT_NUM, EMP_NUM, EXCEPTION_CD, OVER_EXC_CD, SCHED_CHG_EXC_CD, CHARGE_DATE, ADJ_TYPE"
                .MoveFirst
            End If
        End With

        If (RecordsExist(m_adrsReportData)) Then
            r_bReportGenerated = True
            rptMain.Run
            
            '//check for any errors from the active reports script or the vb6 functions being called back
            If m_ActiveReports_ErrNumber <> 0 Then
                Err.Raise m_ActiveReports_ErrNumber, m_ActiveReports_ErrSource, m_ActiveReports_ErrDescription
            End If
        
            'Return the report object to the calling application
            Set GetExceptionDeptRpt2 = rptMain
        End If

    End If

    '//-----clean up public properties used by ActiveReports-----//
    DestroyState
    
    Exit Function
GetExceptionDeptRpt2_EH:
    DestroyState
    
    Select Case Err.Number
    Case m_ActiveReports_ErrNumber
        TrsRaiseError MODULENAME, "GetExceptionDeptRpt2", Err.Number, , , _
            DescArgsFrom("[OrigErrDescription]===>" & m_ActiveReports_OrigErrDescription & "<===", "[OrigErrSource]===>" & m_ActiveReports_OrigErrSource & "<==="), True
    Case Else
        TrsRaiseError MODULENAME, "GetExceptionDeptRpt2", , , , , True
    End Select
End Function

Public Function GetExceptionJobRpt2( _
        ByVal v_dtStartDate As Date _
        , ByVal v_dtEndDate As Date _
        , ByVal v_adrsSelectedJobs As ADODB.Recordset _
        , ByVal v_adrsSelectedExceptionCodes As ADODB.Recordset _
        , ByVal v_lLocation As Long _
        , ByVal v_sLocationDesc As String _
        , ByVal v_adrsAllExceptionCodes As ADODB.Recordset _
        , ByVal v_bPrintExceptionCodes As Boolean _
        , ByRef r_bReportGenerated As Boolean _
        ) As ActiveReport
On Error GoTo GetExceptionJobRpt2_EH
Dim lsFooter As String

    'Initialize return
    r_bReportGenerated = False
    ResetActiveReportsError
    Set GetExceptionJobRpt2 = Nothing

    If (RecordsExist(v_adrsSelectedJobs)) Then
        Set m_adrsSelectedJobs = v_adrsSelectedJobs
    Else
        Exit Function
    End If

    If (RecordsExist(v_adrsAllExceptionCodes)) Then
        Set m_adrsExceptionCodes = v_adrsAllExceptionCodes
    Else
        Exit Function
    End If

    m_lReportLevel = 4
    m_dtStartDate = v_dtStartDate
    m_dtEndDate = v_dtEndDate

    'Load XML file into the report object
    Set rptMain = CreateObject("DDActiveReports2.ActiveReport")
    rptMain.LoadLayout "C:\TRIS Solution\TRISReports\RPX Files\rptExceptionJob.rpx"
    'Add a reference in the rpt object to this (rptException) class
    Call rptMain.AddNamedItem("MyClass", Me)

    Set rptSubReport1 = CreateObject("DDActiveReports2.ActiveReport")
    rptSubReport1.LoadLayout "C:\TRIS Solution\TRISReports\RPX Files\rptExceptionTotalsSub.rpx"
    'Add a reference in the rpt object to this (rptException) class
    Call rptSubReport1.AddNamedItem("MyClass", Me)

    Set rptSubReport2 = CreateObject("DDActiveReports2.ActiveReport")
    rptSubReport2.LoadLayout "C:\TRIS Solution\TRISReports\RPX Files\rptExceptionTotalsSub.rpx"
    'Add a reference in the rpt object to this (rptException) class
    Call rptSubReport2.AddNamedItem("MyClass", Me)

    rptMain.Sections("GroupFooter1").Controls("SubReportJobTotals").object = _
        rptSubReport1
    rptMain.Sections("GroupFooter2").Controls("SubReportEmpTotals").object = _
        rptSubReport2
    rptMain.Sections("PageHeader").Controls("txtHdrRunDate").Text = _
        Now()
    rptMain.Sections("PageHeader").Controls("txtHdrStartDate").Text = _
        CStr(Format(v_dtStartDate, "mm/dd/yyyy"))
    rptMain.Sections("PageHeader").Controls("txtHdrEndDate").Text = _
        CStr(Format(v_dtEndDate, "mm/dd/yyyy"))
    lsFooter = ""
    If (v_bPrintExceptionCodes) Then
        lsFooter = GetFooter(v_adrsSelectedExceptionCodes)
    End If
    rptMain.Sections("PageFooter").Controls("txtExceptionCodeFooter").Text = _
        lsFooter

    Set m_adrsReportData = GetJobExceptionReportData( _
         m_adrsSelectedJobs _
         , v_dtStartDate _
         , v_dtEndDate _
         , v_adrsSelectedExceptionCodes _
         , v_lLocation)
    
    If (RecordsExist(m_adrsReportData)) Then
        DefineEmpJobTotals
        DefineJobTotals

        With m_adrsReportData
            .Sort = "JOB_NUM, EXCEPTION_CD, OVER_EXC_CD, SCHED_CHG_EXC_CD"
            .MoveFirst

            Do While Not .EOF
                ComputeJobTotals _
                    Trim(!JOB_NUM) _
                    , Trim(!EXCEPTION_CD) _
                    , Trim(!OVER_EXC_CD) _
                    , Trim(!SCHED_CHG_EXC_CD) _
                    , Trim(!ADJ_TYPE) _
                    , !Hours _
                    , !AMOUNT _
                    , !TRAVEL_PAY_HRS _
                    , Trim(!PROCESSED_IND)

                ComputeEmpJobTotals _
                    Trim(!JOB_NUM) _
                    , !EMP_NUM _
                    , Trim(!EXCEPTION_CD) _
                    , Trim(!OVER_EXC_CD) _
                    , Trim(!SCHED_CHG_EXC_CD) _
                    , Trim(!ADJ_TYPE) _
                    , !Hours _
                    , !AMOUNT _
                    , !TRAVEL_PAY_HRS _
                    , Trim(!PROCESSED_IND)

                .MoveNext
            Loop

            .Sort = "JOB_NUM, EMP_NUM, EXCEPTION_CD, OVER_EXC_CD, SCHED_CHG_EXC_CD, CHARGE_DATE, ADJ_TYPE"
            .MoveFirst
        End With

        r_bReportGenerated = True
        rptMain.Run

        '//check for any errors from the active reports script or the vb6 functions being called back
        If m_ActiveReports_ErrNumber <> 0 Then
            Err.Raise m_ActiveReports_ErrNumber, m_ActiveReports_ErrSource, m_ActiveReports_ErrDescription
        End If
        
        'Return the report object to the calling application
        Set GetExceptionJobRpt2 = rptMain
    End If

    '//-----clean up public properties used by ActiveReports-----//
    DestroyState
    
    Exit Function
GetExceptionJobRpt2_EH:
    DestroyState
    
    Select Case Err.Number
    Case m_ActiveReports_ErrNumber
        TrsRaiseError MODULENAME, "GetExceptionJobRpt2", Err.Number, , , _
            DescArgsFrom("[OrigErrDescription]===>" & m_ActiveReports_OrigErrDescription & "<===", "[OrigErrSource]===>" & m_ActiveReports_OrigErrSource & "<==="), True
    Case Else
        TrsRaiseError MODULENAME, "GetExceptionJobRpt2", , , , , True
    End Select
End Function

Public Function GetExceptionDeptRpt( _
        ByVal v_dtStartDate As Date, _
        ByVal v_dtEndDate As Date, _
        ByVal v_adrsSelectedDepartments As ADODB.Recordset, _
        ByVal v_adrsSelectedExceptionCodes As ADODB.Recordset, _
        ByVal v_adrsAllExceptionCodes As ADODB.Recordset, _
        ByVal v_bPrintExceptionCodes As Boolean, _
        ByVal v_adrsWorkgroups As ADODB.Recordset, _
        ByVal v_bActualTimeDetails As Boolean, _
        ByVal v_bSchedEmpDetails As Boolean, _
        ByRef r_bReportGenerated As Boolean) As ActiveReport
           
'Logic moved to GetExceptionDeptRpt2 - kept for compatibility (5/10/2004 - des)

End Function
Public Function GetExceptionJobRpt( _
        ByVal v_dtStartDate As Date, _
        ByVal v_dtEndDate As Date, _
        ByVal v_adrsSelectedJobs As ADODB.Recordset, _
        ByVal v_adrsSelectedExceptionCodes As ADODB.Recordset, _
        ByVal v_lLocation As Long, _
        ByVal v_sLocationDesc As String, _
        ByVal v_adrsAllExceptionCodes As ADODB.Recordset, _
        ByVal v_bPrintExceptionCodes As Boolean, _
        ByVal v_bActualTimeDetails As Boolean, _
        ByVal v_bSchedEmpDetails As Boolean, _
        ByRef r_bReportGenerated As Boolean) As ActiveReport
           
'Logic moved to GetExceptionJobRpt2 - kept for compatibility (5/10/2004 - des)

End Function
Private Function GetFooter(ByVal vadrsExceptions As ADODB.Recordset _
                    ) As String
On Error GoTo GetFooter_EH
Dim llIndex As Long, lsFooter As String

    lsFooter = "Report for all Exception Codes."

    If (RecordsExist(vadrsExceptions)) Then
        With vadrsExceptions
            .MoveFirst

            lsFooter = "Report for Exception Codes: "
            For llIndex = 1 To .RecordCount
                lsFooter = lsFooter & Trim(!EXCEPTION_CD)
                If (llIndex < .RecordCount) Then
                    lsFooter = lsFooter & ", "
                End If

                .MoveNext
            Next llIndex

            .MoveFirst
        End With
    End If

    GetFooter = lsFooter

    Exit Function
GetFooter_EH:
    mstdErrorHandling.TrsRaiseError MODULENAME _
        , "GetFooter"

End Function
Public Function GetExceptionReportData() As ADODB.Recordset
On Error GoTo GetExceptionReportData_EH

    'Call to pass the report data into the
    Set GetExceptionReportData = m_adrsReportData

    Exit Function
GetExceptionReportData_EH:
    TrsRaiseError MODULENAME _
        , "GetExceptionReportData" _
        , , , , , True

End Function
Public Function GetEmployeeBanks() As ADODB.Recordset
On Error GoTo GetEmployeeBanks_EH
Dim l_sFilter As String

    If RecordsExist(m_adrsSelectedEmployees) Then
        Set m_adrsEmpVacBanks = GetEmpVacationBanks(m_adrsSelectedEmployees)
    End If
        
    If Not (RecordsExist(m_adrsEmpVacBanks)) Then
        Set GetEmployeeBanks = Nothing
        Exit Function
    End If

    With m_adrsEmpVacBanks
        .Filter = adFilterNone
        .MoveFirst

        l_sFilter = "EMP_NUM = " & m_lCurrentEmpNum & _
            " AND EFF_YEAR = " & DatePart("YYYY", m_dtStartDate)
        .Filter = MDAC25FilterFix(l_sFilter)

        Set GetEmployeeBanks = m_adrsEmpVacBanks
    End With

    rptMain.Sections("GroupFooter1").Controls("lblEmployeeBanks").Caption = _
        "Vacation Bank " & DatePart("yyyy", m_dtStartDate)

    Exit Function
GetEmployeeBanks_EH:
    TrsRaiseError MODULENAME _
        , "GetEmployeeBanks"

End Function
Public Function GetExceptionTotals() As ADODB.Recordset
On Error GoTo GetExceptionTotals_EH
Dim l_sFilter As String

    Select Case (m_lReportLevel)
        Case 1  'Employee Level Report
            With m_adrsEmpTotals
                .Filter = adFilterNone

                l_sFilter = "TOT_EMP_NUM = " & m_lCurrentEmpNum
                .Filter = MDAC25FilterFix(l_sFilter)

                If Not (.EOF) Then
                    .Sort = "TOT_EXCEPTION_CD, TOT_OVER_EXC_CD, TOT_SCHED_CHG_EXC_CD"
                    .MoveFirst
                End If

                Set GetExceptionTotals = m_adrsEmpTotals
            End With

        Case 2  'Workgroup Level Report
            Select Case (m_lProcessLevel)
                Case 1
                    With m_adrsEmpTotals
                        .Filter = adFilterNone
                        l_sFilter = "TOT_WG_NUM = " & m_lWorkgroup & _
                            " AND TOT_EMP_NUM = " & m_lCurrentEmpNum
                        .Filter = MDAC25FilterFix(l_sFilter)

                        If Not (.EOF) Then
                            .Sort = "TOT_EXCEPTION_CD, TOT_OVER_EXC_CD, TOT_SCHED_CHG_EXC_CD"
                            .MoveFirst
                        End If

                        Set GetExceptionTotals = m_adrsEmpTotals
                    End With

                Case 2
                    With m_adrsWGTotals
                        .Filter = adFilterNone

                        l_sFilter = "TOT_WG_NUM = " & m_lWorkgroup
                        .Filter = MDAC25FilterFix(l_sFilter)
                        If Not (.EOF) Then
                            .MoveFirst
                        End If

                        Set GetExceptionTotals = m_adrsWGTotals
                    End With
            End Select

        Case 3  'Department Level Report
            Select Case (m_lProcessLevel)
                Case 1
                    With m_adrsEmpTotals
                        .Filter = adFilterNone
                        l_sFilter = "TOT_DEPT_NUM = '" & m_sDeptNum & "' " & _
                            " and TOT_WG_NUM = " & m_lWorkgroup & _
                            " AND TOT_EMP_NUM = " & m_lCurrentEmpNum
                        .Filter = MDAC25FilterFix(l_sFilter)

                        If Not (.EOF) Then
                            .Sort = "TOT_EXCEPTION_CD, TOT_OVER_EXC_CD, TOT_SCHED_CHG_EXC_CD"
                            .MoveFirst
                        End If

                        Set GetExceptionTotals = m_adrsEmpTotals
                    End With

                Case 2
                    With m_adrsWGTotals
                        .Filter = adFilterNone
                        l_sFilter = "TOT_WG_NUM = " & m_lWorkgroup & _
                            " AND TOT_DEPT_NUM = '" & m_sDeptNum & " ' "
                        .Filter = MDAC25FilterFix(l_sFilter)

                        If Not (.EOF) Then
                            .Sort = "TOT_EXCEPTION_CD, TOT_OVER_EXC_CD, TOT_SCHED_CHG_EXC_CD"
                            .MoveFirst
                        End If

                        Set GetExceptionTotals = m_adrsWGTotals
                    End With

                Case 3
                    With m_adrsDeptTotals
                        .Filter = adFilterNone
                        l_sFilter = "TOT_DEPT_NUM = '" & m_sDeptNum & "' "
                        .Filter = MDAC25FilterFix(l_sFilter)

                        If Not (.EOF) Then
                            .Sort = "TOT_EXCEPTION_CD, TOT_OVER_EXC_CD, TOT_SCHED_CHG_EXC_CD"
                            .MoveFirst
                        End If

                        Set GetExceptionTotals = m_adrsDeptTotals
                    End With
            End Select

        Case 4  'Job Level Report
            Select Case m_lProcessLevel
                Case 1
                    With m_adrsEmpTotals
                        .Filter = adFilterNone
                        l_sFilter = "TOT_JOB_NUM = '" & m_sJobNum & "' " & _
                            " AND TOT_EMP_NUM = " & m_lCurrentEmpNum
                        .Filter = MDAC25FilterFix(l_sFilter)

                        If Not (.EOF) Then
                            .Sort = "TOT_EXCEPTION_CD, TOT_OVER_EXC_CD, TOT_SCHED_CHG_EXC_CD"
                            .MoveFirst
                        End If

                        Set GetExceptionTotals = m_adrsEmpTotals
                    End With

                Case 4
                    With m_adrsJobTotals
                        .Filter = adFilterNone
                        l_sFilter = "TOT_JOB_NUM = '" & m_sJobNum & "'"
                        .Filter = MDAC25FilterFix(l_sFilter)

                        If Not (.EOF) Then
                            .Sort = "TOT_EXCEPTION_CD, TOT_OVER_EXC_CD, TOT_SCHED_CHG_EXC_CD"
                            .MoveFirst
                        End If

                        Set GetExceptionTotals = m_adrsJobTotals
                    End With
            End Select
    End Select

    Exit Function
GetExceptionTotals_EH:
    TrsRaiseError MODULENAME _
        , "GetExceptionTotals"

End Function
Public Function GetExceptionDesc(ByVal v_sExceptionCode As String _
                , ByVal v_sOverrideCode As String _
                ) As String
On Error GoTo GetExceptionDesc_EH

    GetExceptionDesc = ""

    With m_adrsExceptionCodes
        If (Trim(v_sOverrideCode) <> "") Then
            .MoveFirst

            .Find "EXCEPTION_CD = '" & Trim(v_sOverrideCode) & "'"
            If Not .EOF Then
                GetExceptionDesc = Trim(!EXCEPTION_DESC)
            End If
        ElseIf (Trim(v_sExceptionCode) <> "") Then
            .MoveFirst

            .Find "EXCEPTION_CD = '" & Trim(v_sExceptionCode) & "'"
            If Not .EOF Then
                GetExceptionDesc = Trim(!EXCEPTION_DESC)
            End If
        End If
    End With

    Exit Function
GetExceptionDesc_EH:
    TrsRaiseError MODULENAME _
        , "GetExceptionDesc"

End Function
Public Function GetJobDesc(ByVal v_sJobNum As String) As String
On Error GoTo GetJobDesc_EH

    GetJobDesc = ""

    If (Trim(v_sJobNum) <> "") Then
        With m_adrsSelectedJobs
            .MoveFirst

            .Find "JOB_NUM = '" & Trim(v_sJobNum) & "'"
            If Not .EOF Then
                GetJobDesc = Trim(!JOB_DESC)
            End If
        End With
    End If

    Exit Function
GetJobDesc_EH:
    TrsRaiseError MODULENAME _
        , "GetJobDesc"

End Function
Public Function GetWGDesc(ByVal v_sWGNum As String) As String
On Error GoTo GetWGDesc_EH

    GetWGDesc = ""

    If (Trim(v_sWGNum) <> "") Then
        With m_adrsWorkgroups
            .MoveFirst

            .Find "WG_NUM = " & CLng(Trim(v_sWGNum))
            If Not .EOF Then
                GetWGDesc = Trim(!WG_DESC)
            End If
        End With
    End If

    Exit Function
GetWGDesc_EH:
    TrsRaiseError MODULENAME _
        , "GetWGDesc"

End Function
Public Function GetDepartmentName(ByVal v_sDeptNum As String) As String
On Error GoTo GetDepartmentName_EH

    GetDepartmentName = ""

    With m_adrsSelectedDepartments
        .MoveFirst

        .Find "DEPT_NUM = '" & Trim(v_sDeptNum) & "'"
        If Not .EOF Then
            GetDepartmentName = Trim(!DEPT_NAME)
        End If
    End With

    Exit Function
GetDepartmentName_EH:
    TrsRaiseError MODULENAME _
        , "GetDepartmentName"

End Function
Private Function GetEmpVacationBanks( _
                    v_adrsSelectedEmployees As ADODB.Recordset _
                    ) As ADODB.Recordset
On Error GoTo GetEmpVacationBanks_EH
Dim l_objAccrualBankMgr As dbAccrualBankMgr.cdAccrualBankMgr
Dim l_lARREmployees() As Long, l_lEmpIndex As Long

    Set GetEmpVacationBanks = Nothing

    If Not (RecordsExist(v_adrsSelectedEmployees)) Then
        Exit Function
    End If

    With v_adrsSelectedEmployees
        .MoveFirst

         ReDim l_lARREmployees(.RecordCount - 1)

         For l_lEmpIndex = 1 To .RecordCount
             l_lARREmployees(l_lEmpIndex - 1) = !EMP_NUM

             .MoveNext
         Next l_lEmpIndex
    End With

    Set l_objAccrualBankMgr = CreateObject("dbAccrualBankMgr.cdAccrualBankMgr")
    Set GetEmpVacationBanks = l_objAccrualBankMgr.GetEmployeeBanks( _
                                l_lARREmployees _
                                , Now())
    Set l_objAccrualBankMgr = Nothing

    Exit Function
GetEmpVacationBanks_EH:
    Set l_objAccrualBankMgr = Nothing
    TrsRaiseError MODULENAME _
        , "GetEmpVacationBanks"

End Function
Private Function GetEmpExceptionReportData( _
        v_adrsSelectedEmployees As ADODB.Recordset _
        , v_dtStartDate As Date _
        , v_dtEndDate As Date _
        , v_adrsSelectedExceptionCodes As ADODB.Recordset _
        ) As ADODB.Recordset
On Error GoTo GetEmpExceptionReportData_EH
Dim l_objSchedException As dbSchedException.cdSchedException
Dim l_lARREmployees() As Long, l_lEmpIndex As Long
Dim ladrs As ADODB.Recordset

    Set GetEmpExceptionReportData = Nothing

    If Not (RecordsExist(v_adrsSelectedEmployees)) Then
        Exit Function
    End If

    With v_adrsSelectedEmployees
        .MoveFirst

         ReDim l_lARREmployees(.RecordCount - 1)

         For l_lEmpIndex = 1 To .RecordCount
             l_lARREmployees(l_lEmpIndex - 1) = !EMP_NUM

             .MoveNext
         Next l_lEmpIndex
    End With

    Set l_objSchedException = CreateObject("dbSchedException.cdSchedException")
    Set ladrs = l_objSchedException.GetEmpExceptions( _
                    l_lARREmployees _
                    , v_dtStartDate _
                    , v_dtEndDate _
                    , v_adrsSelectedExceptionCodes)
    Set l_objSchedException = Nothing

    If (RecordsExist(ladrs)) Then
        ladrs.Sort = "EMP_NUM, CHARGE_DATE"
        ladrs.MoveFirst
    End If

    Set GetEmpExceptionReportData = ladrs
    Set ladrs = Nothing

    Exit Function
GetEmpExceptionReportData_EH:
    Set l_objSchedException = Nothing
    Set ladrs = Nothing

    TrsRaiseError MODULENAME _
        , "GetEmpExceptionReportData"

End Function
Private Function GetDeptExceptions( _
        v_adrsSelectedDepartments As ADODB.Recordset _
        , v_dtStartDate As Date _
        , v_dtEndDate As Date _
        , v_adrsSelectedExceptionCodes As ADODB.Recordset _
        ) As ADODB.Recordset
On Error GoTo GetDeptExceptions_EH
Dim l_objSchedException As dbSchedException.cdSchedException
Dim l_lDeptIndex As Long, l_lARRDepartments() As String
Dim ladrs As ADODB.Recordset

    Set GetDeptExceptions = Nothing

    If Not (RecordsExist(v_adrsSelectedDepartments)) Then
        Exit Function
    Else
        'Copy employees into an array
        With v_adrsSelectedDepartments
            .MoveFirst

            ReDim l_lARRDepartments(.RecordCount - 1)

            For l_lDeptIndex = 1 To .RecordCount
                l_lARRDepartments(l_lDeptIndex - 1) = !DEPT_NUM

                .MoveNext
            Next l_lDeptIndex
        End With
    End If

    Set l_objSchedException = CreateObject("dbSchedException.cdSchedException")
    Set ladrs = l_objSchedException.GetDeptExceptions( _
                    l_lARRDepartments _
                    , v_dtStartDate _
                    , v_dtEndDate _
                    , v_adrsSelectedExceptionCodes)
    Set l_objSchedException = Nothing

    If (RecordsExist(ladrs)) Then
        ladrs.Sort = "DEPT_NUM, EMP_NUM, CHARGE_DATE"
        ladrs.MoveFirst
    End If

    Set GetDeptExceptions = ladrs
    Set ladrs = Nothing

    Exit Function
GetDeptExceptions_EH:
    Set l_objSchedException = Nothing
    Set ladrs = Nothing

    TrsRaiseError MODULENAME _
        , "GetDeptExceptions"

End Function
Private Function GetWGExceptionReportData( _
        v_adrsSelectedWorkgroups As ADODB.Recordset _
        , v_dtStartDate As Date _
        , v_dtEndDate As Date _
        , v_adrsSelectedExceptionCodes As ADODB.Recordset _
        ) As ADODB.Recordset
On Error GoTo GetWGExceptionReportData_EH
Dim l_objSchedException As dbSchedException.cdSchedException
Dim l_lWGIndex As Long, l_lARRWorkgroups() As Long
Dim ladrs As ADODB.Recordset

    Set GetWGExceptionReportData = Nothing

    If Not (RecordsExist(v_adrsSelectedWorkgroups)) Then
        Exit Function
    Else
        'Copy employees into an array
        With v_adrsSelectedWorkgroups
            .MoveFirst

            ReDim l_lARRWorkgroups(.RecordCount - 1)

            For l_lWGIndex = 1 To .RecordCount
                l_lARRWorkgroups(l_lWGIndex - 1) = !WG_NUM

                .MoveNext
            Next l_lWGIndex
        End With
    End If

    Set l_objSchedException = CreateObject("dbSchedException.cdSchedException")
    Set ladrs = l_objSchedException.GetWGExceptions( _
                    l_lARRWorkgroups _
                    , v_dtStartDate _
                    , v_dtEndDate _
                    , v_adrsSelectedExceptionCodes)
    Set l_objSchedException = Nothing
    
    If (RecordsExist(ladrs)) Then
        ladrs.Sort = "WG_NUM,  EMP_NUM, CHARGE_DATE"
        ladrs.MoveFirst
    End If

    Set GetWGExceptionReportData = ladrs
    Set ladrs = Nothing

    Exit Function
GetWGExceptionReportData_EH:
    Set l_objSchedException = Nothing
    Set ladrs = Nothing

    TrsRaiseError MODULENAME _
        , "GetWGExceptionReportData"

End Function
Private Function GetJobExceptionReportData( _
                v_adrsSelectedJobs As ADODB.Recordset _
                , v_dtStartDate As Date _
                , v_dtEndDate As Date _
                , v_adrsSelectedExceptionCodes As ADODB.Recordset _
                , v_lLocation As Long _
                ) As ADODB.Recordset
On Error GoTo GetJobExceptionReportData_EH
Dim l_objSchedException As dbSchedException.cdSchedException
Dim l_lJobIndex As Long, l_sARRJobs() As String
Dim ladrs As ADODB.Recordset

    Set GetJobExceptionReportData = Nothing

    If Not (RecordsExist(v_adrsSelectedJobs)) Then
        Exit Function
    Else
        'Copy employees into an array
        With v_adrsSelectedJobs
            .MoveFirst

            ReDim l_sARRJobs(.RecordCount - 1)

            For l_lJobIndex = 1 To .RecordCount
                l_sARRJobs(l_lJobIndex - 1) = !JOB_NUM

                .MoveNext
            Next l_lJobIndex
        End With
    End If

    Set l_objSchedException = CreateObject("dbSchedException.cdSchedException")
    Set ladrs = l_objSchedException.GetJobExceptions2( _
                    l_sARRJobs _
                    , v_dtStartDate _
                    , v_dtEndDate _
                    , v_adrsSelectedExceptionCodes _
                    , v_lLocation)
    Set l_objSchedException = Nothing
    If (RecordsExist(ladrs)) Then
        ladrs.Sort = "JOB_NUM,  EMP_NUM, CHARGE_DATE"
        ladrs.MoveFirst
    End If

    Set GetJobExceptionReportData = ladrs
    Set ladrs = Nothing

    Exit Function
GetJobExceptionReportData_EH:
    Set l_objSchedException = Nothing
    Set ladrs = Nothing

    TrsRaiseError MODULENAME _
        , "GetJobExceptionReportData"

End Function
Private Sub DefineDeptTotals()
On Error GoTo DefineDeptTotals_EH

    If Not (m_adrsDeptTotals Is Nothing) Then
        Set m_adrsDeptTotals = Nothing
    End If

    Set m_adrsDeptTotals = New ADODB.Recordset
    With m_adrsDeptTotals
        .Fields.Append "TOT_DEPT_NUM", adChar, 5, m_conFIELDATTR
        .Fields.Append "TOT_EXCEPTION_CD", adChar, 3, m_conFIELDATTR
        .Fields.Append "TOT_OVER_EXC_CD", adChar, 3, m_conFIELDATTR
        .Fields.Append "TOT_SCHED_CHG_EXC_CD", adChar, 3, m_conFIELDATTR
        .Fields.Append "TOT_ADJ", adChar, 3, m_conFIELDATTR
        .Fields.Append "TOT_AMOUNT", adDouble, , m_conFIELDATTR
        .Fields.Append "TOT_EXCEPTION_DESC", adChar, 30, m_conFIELDATTR
        .Fields.Append "TOT_HOURS", adDouble, , m_conFIELDATTR
        .Fields.Append "TOT_TRAVEL_PAY_HRS", adDouble, , m_conFIELDATTR

        .CursorLocation = adUseClient
        .LockType = adLockBatchOptimistic
        .CursorType = adOpenStatic
        .Open
    End With

    Exit Sub
DefineDeptTotals_EH:
    TrsRaiseError MODULENAME _
        , "DefineDeptTotals"

End Sub
Private Sub DefineJobTotals()
On Error GoTo DefineJobTotals_EH

    If Not (m_adrsJobTotals Is Nothing) Then
        Set m_adrsJobTotals = Nothing
    End If

    Set m_adrsJobTotals = New ADODB.Recordset
    With m_adrsJobTotals
        .Fields.Append "TOT_JOB_NUM", adChar, 6, m_conFIELDATTR
        .Fields.Append "TOT_EXCEPTION_CD", adChar, 3, m_conFIELDATTR
        .Fields.Append "TOT_OVER_EXC_CD", adChar, 3, m_conFIELDATTR
        .Fields.Append "TOT_SCHED_CHG_EXC_CD", adChar, 3, m_conFIELDATTR
        .Fields.Append "TOT_ADJ", adChar, 3, m_conFIELDATTR
        .Fields.Append "TOT_AMOUNT", adDouble, , m_conFIELDATTR
        .Fields.Append "TOT_EXCEPTION_DESC", adChar, 30, m_conFIELDATTR
        .Fields.Append "TOT_HOURS", adDouble, , m_conFIELDATTR
        .Fields.Append "TOT_TRAVEL_PAY_HRS", adDouble, , m_conFIELDATTR

        .CursorLocation = adUseClient
        .LockType = adLockBatchOptimistic
        .CursorType = adOpenStatic
        .Open
    End With

    Exit Sub
DefineJobTotals_EH:
    TrsRaiseError MODULENAME _
        , "DefineJobTotals"

End Sub
Private Sub DefineWGTotals()
On Error GoTo DefineWGTotals_EH

    If Not (m_adrsWGTotals Is Nothing) Then
        Set m_adrsWGTotals = Nothing
    End If

    Set m_adrsWGTotals = New ADODB.Recordset
    With m_adrsWGTotals
        .Fields.Append "TOT_WG_NUM", adInteger, , m_conFIELDATTR
        .Fields.Append "TOT_DEPT_NUM", adChar, 3, m_conFIELDATTR
        .Fields.Append "TOT_EXCEPTION_CD", adChar, 3, m_conFIELDATTR
        .Fields.Append "TOT_OVER_EXC_CD", adChar, 3, m_conFIELDATTR
        .Fields.Append "TOT_SCHED_CHG_EXC_CD", adChar, 3, m_conFIELDATTR
        .Fields.Append "TOT_ADJ", adChar, 3, m_conFIELDATTR
        .Fields.Append "TOT_AMOUNT", adDouble, , m_conFIELDATTR
        .Fields.Append "TOT_EXCEPTION_DESC", adChar, 30, m_conFIELDATTR
        .Fields.Append "TOT_HOURS", adDouble, , m_conFIELDATTR
        .Fields.Append "TOT_TRAVEL_PAY_HRS", adDouble, , m_conFIELDATTR

        .CursorLocation = adUseClient
        .LockType = adLockBatchOptimistic
        .CursorType = adOpenStatic
        .Open
    End With

    Exit Sub
DefineWGTotals_EH:
    TrsRaiseError MODULENAME _
        , "DefineWGTotals"

End Sub
Private Sub DefineEmpTotals()
On Error GoTo DefineEmpTotals_EH

    If Not (m_adrsEmpTotals Is Nothing) Then
        Set m_adrsEmpTotals = Nothing
    End If

    Set m_adrsEmpTotals = New ADODB.Recordset
    With m_adrsEmpTotals
        .Fields.Append "TOT_WG_NUM", adInteger, , m_conFIELDATTR
        .Fields.Append "TOT_DEPT_NUM", adChar, 3, m_conFIELDATTR
        .Fields.Append "TOT_EMP_NUM", adInteger, , m_conFIELDATTR
        .Fields.Append "TOT_AMOUNT", adDouble, , m_conFIELDATTR
        .Fields.Append "TOT_EXCEPTION_CD", adChar, 3, m_conFIELDATTR
        .Fields.Append "TOT_OVER_EXC_CD", adChar, 3, m_conFIELDATTR
        .Fields.Append "TOT_SCHED_CHG_EXC_CD", adChar, 3, m_conFIELDATTR
        .Fields.Append "TOT_ADJ", adChar, 3, m_conFIELDATTR
        .Fields.Append "TOT_EXCEPTION_DESC", adChar, 30, m_conFIELDATTR
        .Fields.Append "TOT_HOURS", adDouble, , m_conFIELDATTR
        .Fields.Append "TOT_TRAVEL_PAY_HRS", adDouble, , m_conFIELDATTR

        .CursorLocation = adUseClient
        .LockType = adLockBatchOptimistic
        .CursorType = adOpenStatic
        .Open
    End With

    Exit Sub
DefineEmpTotals_EH:
    TrsRaiseError MODULENAME _
        , "DefineEmpTotals"

End Sub
Private Sub DefineEmpJobTotals()
On Error GoTo DefineEmpJobTotals_EH

    If Not (m_adrsEmpTotals Is Nothing) Then
        Set m_adrsEmpTotals = Nothing
    End If

    Set m_adrsEmpTotals = New ADODB.Recordset
    With m_adrsEmpTotals
        .Fields.Append "TOT_JOB_NUM", adChar, 6, m_conFIELDATTR
        .Fields.Append "TOT_EMP_NUM", adInteger, , m_conFIELDATTR
        .Fields.Append "TOT_EXCEPTION_CD", adChar, 3, m_conFIELDATTR
        .Fields.Append "TOT_OVER_EXC_CD", adChar, 3, m_conFIELDATTR
        .Fields.Append "TOT_SCHED_CHG_EXC_CD", adChar, 3, m_conFIELDATTR
        .Fields.Append "TOT_AMOUNT", adDouble, , m_conFIELDATTR
        .Fields.Append "TOT_ADJ", adChar, 3, m_conFIELDATTR
        .Fields.Append "TOT_EXCEPTION_DESC", adChar, 30, m_conFIELDATTR
        .Fields.Append "TOT_HOURS", adDouble, , m_conFIELDATTR
        .Fields.Append "TOT_TRAVEL_PAY_HRS", adDouble, , m_conFIELDATTR

        .CursorLocation = adUseClient
        .LockType = adLockBatchOptimistic
        .CursorType = adOpenStatic
        .Open
    End With

    Exit Sub
DefineEmpJobTotals_EH:
    TrsRaiseError MODULENAME _
        , "DefineEmpJobTotals"

End Sub
Private Sub ComputeEmpTotals( _
            ByVal v_sDeptNum As String _
            , ByVal v_lWGNum As Long _
            , ByVal v_lEmpNum As Long _
            , ByVal v_sExceptionCode As String _
            , ByVal v_sOverrideCode As String _
            , ByVal v_sSchedChgCode As String _
            , ByVal v_sAdjustment As String _
            , ByVal v_dHours As Double _
            , ByVal v_dAmount As Double _
            , ByVal v_dTravelPayHrs As Double _
            , ByVal v_sProcessedInd As String)
On Error GoTo ComputeEmpTotals_EH
Dim l_sFilter As String

    With m_adrsEmpTotals
        .Filter = adFilterNone

        l_sFilter = "TOT_WG_NUM = " & v_lWGNum & _
            " AND TOT_EMP_NUM = " & v_lEmpNum & _
            " AND TOT_DEPT_NUM = '" & v_sDeptNum & "'" & _
            " AND TOT_EXCEPTION_CD = '" & v_sExceptionCode & "'" & _
            " AND TOT_OVER_EXC_CD = '" & Trim(v_sOverrideCode) & "'" & _
            " AND TOT_SCHED_CHG_EXC_CD = '" & Trim(v_sSchedChgCode) & "'" & _
            " AND TOT_ADJ = '" & Trim(v_sAdjustment) & "'"
        .Filter = MDAC25FilterFix(l_sFilter)

        If Not (.EOF) Then
            !TOT_HOURS = !TOT_HOURS + v_dHours
            !TOT_AMOUNT = !TOT_AMOUNT + v_dAmount
            !TOT_TRAVEL_PAY_HRS = !TOT_TRAVEL_PAY_HRS + v_dTravelPayHrs

            .Update
        Else
            .AddNew

            !TOT_WG_NUM = v_lWGNum
            !TOT_EMP_NUM = v_lEmpNum
            !TOT_DEPT_NUM = v_sDeptNum
            !TOT_EXCEPTION_CD = v_sExceptionCode
            !TOT_OVER_EXC_CD = v_sOverrideCode
            !TOT_SCHED_CHG_EXC_CD = v_sSchedChgCode
            !TOT_ADJ = v_sAdjustment
            !TOT_HOURS = v_dHours
            !TOT_AMOUNT = v_dAmount
            !TOT_TRAVEL_PAY_HRS = v_dTravelPayHrs

            .Update
        End If

        .Filter = adFilterNone
    End With

    Exit Sub
ComputeEmpTotals_EH:
    TrsRaiseError MODULENAME _
        , "ComputeEmpTotals"

End Sub
Private Sub ComputeDeptTotals( _
            ByVal v_sDeptNum As String _
            , ByVal v_sExceptionCode As String _
            , ByVal v_sOverrideCode As String _
            , ByVal v_sSchedChgCode As String _
            , ByVal v_sAdjustment As String _
            , ByVal v_dHours As Double _
            , ByVal v_dAmount As Double _
            , ByVal v_dTravelPayHrs As Double _
            , ByVal v_sProcessedInd As String)
On Error GoTo ComputeDeptTotals_EH
Dim l_sFilter As String

    With m_adrsDeptTotals
        .Filter = adFilterNone

        l_sFilter = "TOT_DEPT_NUM = '" & v_sDeptNum & "' " & _
            " AND TOT_EXCEPTION_CD = '" & Trim(v_sExceptionCode) & "'" & _
            " AND TOT_OVER_EXC_CD = '" & Trim(v_sOverrideCode) & "' " & _
            " AND TOT_SCHED_CHG_EXC_CD = '" & Trim(v_sSchedChgCode) & "' " & _
            " AND TOT_ADJ = '" & Trim(v_sAdjustment) & "'"
        .Filter = MDAC25FilterFix(l_sFilter)

        If Not (.EOF) Then
            !TOT_HOURS = !TOT_HOURS + v_dHours
            !TOT_AMOUNT = !TOT_AMOUNT + v_dAmount
            !TOT_TRAVEL_PAY_HRS = !TOT_TRAVEL_PAY_HRS + v_dTravelPayHrs
            .Update
        Else
            .AddNew

            !TOT_DEPT_NUM = v_sDeptNum
            !TOT_EXCEPTION_CD = v_sExceptionCode
            !TOT_OVER_EXC_CD = v_sOverrideCode
            !TOT_SCHED_CHG_EXC_CD = v_sSchedChgCode
            !TOT_ADJ = v_sAdjustment
            !TOT_HOURS = v_dHours
            !TOT_AMOUNT = v_dAmount
            !TOT_TRAVEL_PAY_HRS = v_dTravelPayHrs

            .Update
        End If

        .Filter = adFilterNone
    End With

    Exit Sub
ComputeDeptTotals_EH:
    TrsRaiseError MODULENAME _
        , "ComputeDeptTotals"

End Sub
Private Sub ComputeWGTotals( _
            ByVal v_sDeptNum As String _
            , ByVal v_lWGNum As Long _
            , ByVal v_sExceptionCode As String _
            , ByVal v_sOverrideCode As String _
            , ByVal v_sSchedChgCode As String _
            , ByVal v_sAdjustment As String _
            , ByVal v_dHours As Double _
            , ByVal v_dAmount As Double _
            , ByVal v_dTravelPayHrs As Double _
            , ByVal v_sProcessedInd As String)
On Error GoTo ComputeWGTotals_EH
Dim l_sFilter As String

    With m_adrsWGTotals
        .Filter = adFilterNone

        l_sFilter = "TOT_WG_NUM = " & v_lWGNum & _
            " AND TOT_DEPT_NUM = '" & v_sDeptNum & "'" & _
            " AND TOT_EXCEPTION_CD = '" & v_sExceptionCode & "'" & _
            " AND TOT_OVER_EXC_CD = '" & Trim(v_sOverrideCode) & "' " & _
            " AND TOT_SCHED_CHG_EXC_CD = '" & Trim(v_sSchedChgCode) & "' " & _
            " AND TOT_ADJ = '" & Trim(v_sAdjustment) & "' "
        .Filter = MDAC25FilterFix(l_sFilter)

        If Not (.EOF) Then
            !TOT_HOURS = !TOT_HOURS + v_dHours
            !TOT_AMOUNT = !TOT_AMOUNT + v_dAmount
            !TOT_TRAVEL_PAY_HRS = !TOT_TRAVEL_PAY_HRS + v_dTravelPayHrs
            .Update
        Else
            .AddNew

            !TOT_WG_NUM = v_lWGNum
            !TOT_DEPT_NUM = Trim(v_sDeptNum)
            !TOT_EXCEPTION_CD = Trim(v_sExceptionCode)
            !TOT_OVER_EXC_CD = Trim(v_sOverrideCode)
            !TOT_SCHED_CHG_EXC_CD = Trim(v_sSchedChgCode)
            !TOT_ADJ = Trim(v_sAdjustment)
            !TOT_HOURS = v_dHours
            !TOT_AMOUNT = v_dAmount
            !TOT_TRAVEL_PAY_HRS = v_dTravelPayHrs

            .Update
        End If

        .Filter = adFilterNone
    End With
    
    Exit Sub
ComputeWGTotals_EH:
    TrsRaiseError MODULENAME _
        , "ComputeWGTotals"

End Sub
Private Sub ComputeJobTotals( _
            ByVal v_sJobNum As String _
            , ByVal v_sExceptionCode As String _
            , ByVal v_sOverrideCode As String _
            , ByVal v_sSchedChgCode As String _
            , ByVal v_sAdjustment As String _
            , ByVal v_dHours As Double _
            , ByVal v_dAmount As Double _
            , ByVal v_dTravelPayHrs As Double _
            , ByVal v_sProcessedInd As String)
On Error GoTo ComputeJobTotals_EH
Dim l_sFilter As String

    With m_adrsJobTotals
        .Filter = adFilterNone

        l_sFilter = "TOT_JOB_NUM = '" & v_sJobNum & "' " & _
            " AND TOT_EXCEPTION_CD = '" & v_sExceptionCode & "'" & _
            " AND TOT_OVER_EXC_CD = '" & Trim(v_sOverrideCode) & "' " & _
            " AND TOT_SCHED_CHG_EXC_CD = '" & Trim(v_sSchedChgCode) & "' " & _
            " AND TOT_ADJ = '" & Trim(v_sAdjustment) & "'"
        .Filter = MDAC25FilterFix(l_sFilter)

        If Not (.EOF) Then
            !TOT_HOURS = !TOT_HOURS + v_dHours
            !TOT_AMOUNT = !TOT_AMOUNT + v_dAmount
            !TOT_TRAVEL_PAY_HRS = !TOT_TRAVEL_PAY_HRS + v_dTravelPayHrs

            .Update
        Else
            .AddNew

            !TOT_JOB_NUM = v_sJobNum
            !TOT_EXCEPTION_CD = v_sExceptionCode
            !TOT_OVER_EXC_CD = v_sOverrideCode
            !TOT_SCHED_CHG_EXC_CD = v_sSchedChgCode
            !TOT_ADJ = v_sAdjustment
            !TOT_HOURS = v_dHours
            !TOT_AMOUNT = v_dAmount
            !TOT_TRAVEL_PAY_HRS = v_dTravelPayHrs

            .Update
        End If

        .Filter = adFilterNone
    End With

    Exit Sub
ComputeJobTotals_EH:
    TrsRaiseError MODULENAME _
        , "ComputeJobTotals"
    
End Sub
Private Sub ComputeEmpJobTotals( _
            ByVal v_sJobNum As String _
            , ByVal v_lEmpNum As Long _
            , ByVal v_sExceptionCode As String _
            , ByVal v_sOverrideCode As String _
            , ByVal v_sSchedChgCode As String _
            , ByVal v_sAdjustment As String _
            , ByVal v_dHours As Double _
            , ByVal v_dAmount As Double _
            , ByVal v_dTravelPayHrs As Double _
            , ByVal v_sProcessedInd As String)
On Error GoTo ComputeEmpJobTotals_EH
Dim l_sFilter As String

    With m_adrsEmpTotals
        .Filter = adFilterNone

        l_sFilter = "TOT_JOB_NUM = '" & v_sJobNum & "' " & _
            " AND TOT_EMP_NUM = " & v_lEmpNum & _
            " AND TOT_EXCEPTION_CD = '" & Trim(v_sExceptionCode) & "'" & _
            " AND TOT_OVER_EXC_CD = '" & Trim(v_sOverrideCode) & "' " & _
            " AND TOT_SCHED_CHG_EXC_CD = '" & Trim(v_sSchedChgCode) & "' " & _
            " AND TOT_ADJ = '" & Trim(v_sAdjustment) & "'"

        .Filter = MDAC25FilterFix(l_sFilter)
        If Not (.EOF) Then
            !TOT_HOURS = !TOT_HOURS + v_dHours
            !TOT_AMOUNT = !TOT_AMOUNT + v_dAmount
            !TOT_TRAVEL_PAY_HRS = !TOT_TRAVEL_PAY_HRS + v_dTravelPayHrs

            .Update
        Else
            .AddNew

            !TOT_JOB_NUM = v_sJobNum
            !TOT_EMP_NUM = v_lEmpNum
            !TOT_EXCEPTION_CD = v_sExceptionCode
            !TOT_OVER_EXC_CD = v_sOverrideCode
            !TOT_SCHED_CHG_EXC_CD = v_sSchedChgCode
            !TOT_ADJ = v_sAdjustment
            !TOT_HOURS = v_dHours
            !TOT_AMOUNT = v_dAmount
            !TOT_TRAVEL_PAY_HRS = v_dTravelPayHrs

            .Update
        End If

        .Filter = adFilterNone
    End With
    
    Exit Sub
ComputeEmpJobTotals_EH:
    TrsRaiseError MODULENAME _
        , "ComputeEmpJobTotals"
    
End Sub
Private Function GetEmployeesInWorkGroup( _
                ByVal v_adrsSelectedWorkgroups As ADODB.Recordset _
                , v_dtStartDate As Date _
                , v_dtEndDate As Date _
                ) As ADODB.Recordset
On Error GoTo GetEmployeesInWorkgroup_EH
Dim l_lRecordIndex As Long, l_lARRWorkgroups() As Long
Dim l_objWorkgroup As busWorkGroup.cbsWorkGroup

    Set GetEmployeesInWorkGroup = Nothing

    If Not (RecordsExist(v_adrsSelectedWorkgroups)) Then
        Exit Function
    End If

    With v_adrsSelectedWorkgroups
        .MoveFirst

        ReDim l_lARRWorkgroups(.RecordCount - 1)
        For l_lRecordIndex = 1 To .RecordCount
            l_lARRWorkgroups(l_lRecordIndex - 1) = !WG_NUM

            .MoveNext
        Next l_lRecordIndex
    End With

    Set l_objWorkgroup = CreateObject("busWorkGroup.cbsWorkGroup")
    Set GetEmployeesInWorkGroup = l_objWorkgroup.GetEmployeesInWorkGroup( _
                            l_lARRWorkgroups _
                            , v_dtStartDate _
                            , v_dtEndDate)
    Set l_objWorkgroup = Nothing

    Exit Function
GetEmployeesInWorkgroup_EH:
    Set l_objWorkgroup = Nothing

    TrsRaiseError MODULENAME _
        , "GetEmployeesInWorkgroup"

End Function

Private Sub DestroyState()
    Set m_adrsSelectedEmployees = Nothing
    Set m_adrsSelectedDepartments = Nothing
    Set m_adrsSelectedWorkgroups = Nothing
    Set m_adrsSelectedJobs = Nothing
    
    'Supporting Information
    Set m_adrsExceptionCodes = Nothing
    
    'Recordset for report totals
    Set m_adrsWGTotals = Nothing
    Set m_adrsEmpTotals = Nothing
    Set m_adrsDeptTotals = Nothing
    Set m_adrsJobTotals = Nothing
    Set m_adrsEmpVacBanks = Nothing
    
    'Define the main report
    Set m_adrsReportData = Nothing
    Set rptMain = Nothing
    
    'Sub Reports
    Set rptSubReport1 = Nothing
    Set rptSubReport2 = Nothing
    Set rptSubReport3 = Nothing
    
    'WG Schedule report variables
    Set m_adrsWorkgroups = Nothing
    Set l_adrsTemp = Nothing

End Sub

Private Function ResetActiveReportsError()
    m_ActiveReports_ErrNumber = 0
    m_ActiveReports_ErrDescription = ""
    m_ActiveReports_ErrSource = ""
    m_ActiveReports_OrigErrDescription = ""
    m_ActiveReports_OrigErrSource = ""
End Function

Private Sub rptMain_Error(ByVal Number As Integer, ByVal Description As DDActiveReports2.IReturnString, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As DDActiveReports2.IReturnBool)
    CancelDisplay = True

    If Number <> 5007 Then '//No Printer Installed. Using default papersize.
        TrsTraceError MODULENAME, "rptMain_Error", , , , DescArgsFrom(Source, CStr(Number), CStr(Description)), True
        
        '//keep for raising back to client making initial call
        m_ActiveReports_OrigErrDescription = mstdErrorHandling.mstrErrDescRet '//may be what caused the scripting error
        m_ActiveReports_OrigErrSource = mstdErrorHandling.mstrErrSourceRet

        m_ActiveReports_ErrNumber = Number
        m_ActiveReports_ErrDescription = Description '//active reports scripting error
        m_ActiveReports_ErrSource = Source
    End If
End Sub

Private Sub rptSubReport1_Error(ByVal Number As Integer, ByVal Description As DDActiveReports2.IReturnString, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As DDActiveReports2.IReturnBool)
    CancelDisplay = True

    If Number <> 5007 Then '//No Printer Installed. Using default papersize.
        TrsTraceError MODULENAME, "rptSubReport1_Error", , , , DescArgsFrom(Source, CStr(Number), CStr(Description)), True
        
        '//keep for raising back to client making initial call
        m_ActiveReports_OrigErrDescription = mstdErrorHandling.mstrErrDescRet '//may be what caused the scripting error
        m_ActiveReports_OrigErrSource = mstdErrorHandling.mstrErrSourceRet

        m_ActiveReports_ErrNumber = Number
        m_ActiveReports_ErrDescription = Description '//active reports scripting error
        m_ActiveReports_ErrSource = Source
    End If
End Sub

Private Sub rptSubReport2_Error(ByVal Number As Integer, ByVal Description As DDActiveReports2.IReturnString, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As DDActiveReports2.IReturnBool)
    CancelDisplay = True

    If Number <> 5007 Then '//No Printer Installed. Using default papersize.
        TrsTraceError MODULENAME, "rptSubReport2_Error", , , , DescArgsFrom(Source, CStr(Number), CStr(Description)), True
        
        '//keep for raising back to client making initial call
        m_ActiveReports_OrigErrDescription = mstdErrorHandling.mstrErrDescRet '//may be what caused the scripting error
        m_ActiveReports_OrigErrSource = mstdErrorHandling.mstrErrSourceRet

        m_ActiveReports_ErrNumber = Number
        m_ActiveReports_ErrDescription = Description '//active reports scripting error
        m_ActiveReports_ErrSource = Source
    End If
End Sub

Private Sub rptSubReport3_Error(ByVal Number As Integer, ByVal Description As DDActiveReports2.IReturnString, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As DDActiveReports2.IReturnBool)
    CancelDisplay = True

    If Number <> 5007 Then '//No Printer Installed. Using default papersize.
        TrsTraceError MODULENAME, "rptSubReport3_Error", , , , DescArgsFrom(Source, CStr(Number), CStr(Description)), True
        
        '//keep for raising back to client making initial call
        m_ActiveReports_OrigErrDescription = mstdErrorHandling.mstrErrDescRet '//may be what caused the scripting error
        m_ActiveReports_OrigErrSource = mstdErrorHandling.mstrErrSourceRet

        m_ActiveReports_ErrNumber = Number
        m_ActiveReports_ErrDescription = Description '//active reports scripting error
        m_ActiveReports_ErrSource = Source
    End If
End Sub

