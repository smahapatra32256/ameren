VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "rptActualTimeHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Support information
Dim m_adrsAllDepartments As ADODB.Recordset
Dim m_adrsAllWorkgroups As ADODB.Recordset
Dim m_adrsAllEmployees As ADODB.Recordset
Dim m_adrsExceptionCodes As ADODB.Recordset

'Totals
Dim m_adrsEmpActualTimeTotals As ADODB.Recordset
Dim m_adrsDeptActualTimeTotals As ADODB.Recordset
Dim m_adrsWGActualTimeTotals As ADODB.Recordset
'Dim m_adrsEmpDistActualTimeTotals As ADODB.Recordset
Dim m_adrsDistActualTimeTotals As ADODB.Recordset

'Selected
Dim m_adrsSelectedEmployees As ADODB.Recordset
Dim m_adrsSelectedDistributions As ADODB.Recordset
Dim m_adrsSelectedWorkgroups As ADODB.Recordset
Dim m_adrsSelectedDepartments As ADODB.Recordset

Dim m_dtStartDate As Date, m_dtEndDate As Date

Dim WithEvents rptMain As ActiveReport
Attribute rptMain.VB_VarHelpID = -1
Dim WithEvents rptSubReport1 As ActiveReport
Attribute rptSubReport1.VB_VarHelpID = -1
Dim WithEvents rptSubReport2 As ActiveReport
Attribute rptSubReport2.VB_VarHelpID = -1
Dim m_adrsReportData As ADODB.Recordset

Dim m_dTotalHours As Double, m_dOTHours As Double
Dim m_dSymHours As Double, m_dSTHours As Double
Dim m_lCurrentEmpNum As Long, m_sDeptNum As String
Dim m_lWorkgroup As Long
Dim m_sDistributionId As String, m_sDistributionType As String

Dim m_sHoldDeptNum As String, m_sHoldDeptName As String

'Report Level
    '1 - Employee
    '2 - Workgroup
    '3 - Department
    '4 - Distribution
Dim m_lProcessLevel As Long, m_lReportLevel As Long
Dim m_ActiveReports_ErrNumber As Integer
Dim m_ActiveReports_ErrDescription As String
Dim m_ActiveReports_ErrSource As String
Dim m_ActiveReports_OrigErrDescription As String
Dim m_ActiveReports_OrigErrSource As String

Private Const m_conFIELDATTR = adFldUpdatable Or adFldIsNullable
Private Const MODULENAME As String = "rptEngine.rptActualTimeHistory"
Private Sub DestroyState()

    Set m_adrsAllEmployees = Nothing
    Set m_adrsAllDepartments = Nothing
    Set m_adrsAllWorkgroups = Nothing

    Set m_adrsEmpActualTimeTotals = Nothing
    Set m_adrsDeptActualTimeTotals = Nothing
    Set m_adrsWGActualTimeTotals = Nothing
    Set m_adrsDistActualTimeTotals = Nothing

    Set m_adrsExceptionCodes = Nothing

    Set m_adrsSelectedEmployees = Nothing
    Set m_adrsSelectedDistributions = Nothing
    Set m_adrsSelectedWorkgroups = Nothing
    Set m_adrsSelectedDepartments = Nothing

    Set m_adrsReportData = Nothing

    Set rptMain = Nothing
    Set rptSubReport1 = Nothing
    Set rptSubReport2 = Nothing

End Sub
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
Public Property Get DistributionId() As String
    DistributionId = m_sDistributionId
End Property
Public Property Let DistributionId(ByVal v_sValue As String)
    m_sDistributionId = v_sValue
End Property
Public Property Get DistributionType() As String
    DistributionType = m_sDistributionType
End Property
Public Property Let DistributionType(ByVal v_sValue As String)
    m_sDistributionType = v_sValue
End Property
Public Property Get TotalHours() As String
    'Report must invoke GetEmployeeTotals to populated the module level variables
    TotalHours = FormatNumber(m_dTotalHours, 1)
    m_dTotalHours = 0
End Property
Public Property Get OTHours() As String
    'Report must invoke GetEmployeeTotals to populated the module level variables
    OTHours = FormatNumber(m_dOTHours, 1)
    m_dOTHours = 0
End Property
Public Property Get SymHours() As String
    'Report must invoke GetEmployeeTotals to
    'populated the module level variables
    SymHours = FormatNumber(m_dSymHours, 1)
    m_dSymHours = 0
End Property
Public Property Get STHours() As String
    'Report must invoke GetEmployeeTotals to
    'populated the module level variables
    STHours = FormatNumber(m_dSTHours, 1)
    m_dSTHours = 0
End Property
Public Function GetActualTimeHistoryEmpRpt( _
                ByVal v_dtStartDate As Date _
                , ByVal v_dtEndDate As Date _
                , ByVal v_adrsEmployeesSelected As ADODB.Recordset _
                , ByVal v_adrsExceptionCodes As ADODB.Recordset _
                , ByRef r_bReportGenerated As Boolean _
                , ByVal v_bActiveEmployees As Boolean _
                , ByVal v_adrsAllDepartments As ADODB.Recordset _
                ) As ActiveReport
On Error GoTo GetActualTimeHistoryEmpRpt_EH

'//TrsTraceError MODULENAME, "GetActualTimeHistoryEmpRpt", , , , DescArgsFrom("starting"), True
    'Initialize return
    r_bReportGenerated = False
    ResetActiveReportsError
    
    Set GetActualTimeHistoryEmpRpt = Nothing

    If (RecordsExist(v_adrsEmployeesSelected)) Then
        Set m_adrsSelectedEmployees = v_adrsEmployeesSelected
        Set m_adrsAllEmployees = m_adrsSelectedEmployees.Clone
    Else
        Exit Function
    End If

    If (RecordsExist(v_adrsAllDepartments)) Then
        Set m_adrsAllDepartments = v_adrsAllDepartments
    Else
        Exit Function
    End If

    m_lReportLevel = 1 'Employee Level Report

    m_dtStartDate = v_dtStartDate
    m_dtEndDate = v_dtEndDate

    Set m_adrsExceptionCodes = v_adrsExceptionCodes

    'Load XML file into the report object
    Set rptMain = CreateObject("DDActiveReports2.ActiveReport")
    rptMain.LoadLayout "C:\TRIS Solution\TRISReports\RPX Files\rptActualTimeHistoryEmp.rpx"
    'Add a reference in the rpt object to this (rptException) class
    Call rptMain.AddNamedItem("MyClass", Me)

    Set rptSubReport1 = CreateObject("DDActiveReports2.ActiveReport")
    rptSubReport1.LoadLayout "C:\TRIS Solution\TRISReports\RPX Files\rptActualTimeTotalsSub.rpx"
    'Add a reference in the rpt object to this (rptException) class
    Call rptSubReport1.AddNamedItem("MyClass", Me)

    rptMain.Sections("GroupFooter1").Controls("subRptTotals").object = _
        rptSubReport1
    rptMain.Sections("PageHeader").Controls("txtHdrRunDate").Text = _
        Now()
    rptMain.Sections("PageHeader").Controls("txtHdrStartDate").Text = _
        CStr(Format(v_dtStartDate, "mm/dd/yyyy"))
    rptMain.Sections("PageHeader").Controls("txtHdrEndDate").Text = _
        CStr(Format(v_dtEndDate, "mm/dd/yyyy"))

    Set m_adrsReportData = GetActualTimeHistoryEmp( _
                            m_adrsSelectedEmployees _
                            , v_dtStartDate _
                            , v_dtEndDate _
                            , v_bActiveEmployees)
    If (RecordsExist(m_adrsReportData)) Then
    'place filter here if distributions are supplied
        r_bReportGenerated = True

        m_adrsReportData.Sort = "EMP_NUM, CHARGE_DATE"
        m_adrsReportData.MoveFirst

        CalculateEmpTotals

        m_adrsReportData.Sort = "EMP_NUM, CHARGE_DATE, DATE_WORKED, START_TIME, END_TIME, SEQ_NUM, HOURS"
        m_adrsReportData.MoveFirst

        'Return the report object to the calling application
'//TrsTraceError MODULENAME, "GetActualTimeHistoryEmpRpt", , , , DescArgsFrom("before running active reports"), True
        rptMain.Run

        '//check for any errors from the active reports script or the vb6 functions being called back
        If m_ActiveReports_ErrNumber <> 0 Then
            Err.Raise m_ActiveReports_ErrNumber, m_ActiveReports_ErrSource, m_ActiveReports_ErrDescription
        End If
        
'//TrsTraceError MODULENAME, "GetActualTimeHistoryEmpRpt", , , , DescArgsFrom("after running active reports"), True
'//Err.Source = ""
        Set GetActualTimeHistoryEmpRpt = rptMain
    End If

    '//-----clean up public properties used by ActiveReports-----//
    DestroyState

'//TrsTraceError MODULENAME, "GetActualTimeHistoryEmpRpt", , , , DescArgsFrom("ending"), True
'//Err.Source = ""
    Exit Function
GetActualTimeHistoryEmpRpt_EH:

    DestroyState
    
    Select Case Err.Number
    Case m_ActiveReports_ErrNumber
        TrsRaiseError MODULENAME, "GetActualTimeHistoryEmpRpt", Err.Number, , , _
            DescArgsFrom("[OrigErrDescription]===>" & m_ActiveReports_OrigErrDescription & "<===", "[OrigErrSource]===>" & m_ActiveReports_OrigErrSource & "<==="), True
    Case Else
        TrsRaiseError MODULENAME, "GetActualTimeHistoryEmpRpt", , , , , True
    End Select
End Function

Public Function GetActualTimeHistoryWGRpt( _
                ByVal v_dtStartDate As Date _
                , ByVal v_dtEndDate As Date _
                , ByVal v_adrsSelectedWorkgroups As ADODB.Recordset _
                , ByVal v_adrsExceptionCodes As ADODB.Recordset _
                , ByRef r_bReportGenerated As Boolean _
                , ByVal v_bActiveEmployees As Boolean _
                , ByVal v_adrsAllEmployees As ADODB.Recordset _
                ) As ActiveReport
On Error GoTo GetActualTimeHistoryWGRpt_EH

    'Initialize return
    r_bReportGenerated = False
    ResetActiveReportsError
    Set GetActualTimeHistoryWGRpt = Nothing

    If (RecordsExist(v_adrsSelectedWorkgroups)) Then
        Set m_adrsSelectedWorkgroups = v_adrsSelectedWorkgroups
        Set m_adrsAllWorkgroups = m_adrsSelectedWorkgroups.Clone
    Else
        Exit Function
    End If
    If (RecordsExist(v_adrsAllEmployees)) Then
        Set m_adrsAllEmployees = v_adrsAllEmployees
    Else
        Exit Function
    End If

    m_lReportLevel = 2

    m_dtStartDate = v_dtStartDate
    m_dtEndDate = v_dtEndDate

    Set m_adrsExceptionCodes = v_adrsExceptionCodes

    'Load XML file into the report object
    Set rptMain = CreateObject("DDActiveReports2.ActiveReport")
    rptMain.LoadLayout "C:\TRIS Solution\TRISReports\RPX Files\rptActualTimeHistoryWG.rpx"
    'Add a reference in the rpt object to this (rptException) class
    Call rptMain.AddNamedItem("MyClass", Me)

    Set rptSubReport1 = CreateObject("DDActiveReports2.ActiveReport")
    rptSubReport1.LoadLayout "C:\TRIS Solution\TRISReports\RPX Files\rptActualTimeTotalsSub.rpx"
    'Add a reference in the rpt object to this (rptException) class
    Call rptSubReport1.AddNamedItem("MyClass", Me)

    Set rptSubReport2 = CreateObject("DDActiveReports2.ActiveReport")
    rptSubReport2.LoadLayout "C:\TRIS Solution\TRISReports\RPX Files\rptActualTimeTotalsSub.rpx"
    'Add a reference in the rpt object to this (rptException) class
    Call rptSubReport2.AddNamedItem("MyClass", Me)

    rptMain.Sections("GroupFooter2").Controls("subRptEmpTotals").object = _
        rptSubReport1
    rptMain.Sections("GroupFooter1").Controls("subRptWGTotals").object = _
        rptSubReport2
    rptMain.Sections("PageHeader").Controls("txtHdrRunDate").Text = _
        Now()
    rptMain.Sections("PageHeader").Controls("txtHdrStartDate").Text = _
        CStr(Format(v_dtStartDate, "mm/dd/yyyy"))
    rptMain.Sections("PageHeader").Controls("txtHdrEndDate").Text = _
        CStr(Format(v_dtEndDate, "mm/dd/yyyy"))

    Set m_adrsReportData = GetActualTimeHistoryWG( _
                            m_adrsSelectedWorkgroups _
                            , v_dtStartDate _
                            , v_dtEndDate _
                            , v_bActiveEmployees)
    If (RecordsExist(m_adrsReportData)) Then
    'place filter here if distributions are supplied
        r_bReportGenerated = True

        m_adrsReportData.Sort = "WG_NUM, EXCEPTION_CD, CHARGE_DATE"
        m_adrsReportData.MoveFirst

        CalculateWGTotals

        m_adrsReportData.Sort = "WG_NUM, EMP_NUM, EXCEPTION_CD, CHARGE_DATE"
        m_adrsReportData.MoveFirst

        CalculateEmpWGTotals

        m_adrsReportData.Sort = "WG_NUM, EMP_NUM, CHARGE_DATE, DATE_WORKED, START_TIME, END_TIME, SEQ_NUM, HOURS"
        m_adrsReportData.MoveFirst

        'Return the report object to the calling application
        rptMain.Run
        
        '//check for any errors from the active reports script or the vb6 functions being called back
        If m_ActiveReports_ErrNumber <> 0 Then
            Err.Raise m_ActiveReports_ErrNumber, m_ActiveReports_ErrSource, m_ActiveReports_ErrDescription
        End If
        
        Set GetActualTimeHistoryWGRpt = rptMain
    End If
    
    '//-----clean up public properties used by ActiveReports-----//
    DestroyState

    Exit Function
GetActualTimeHistoryWGRpt_EH:
    DestroyState
    
    Select Case Err.Number
    Case m_ActiveReports_ErrNumber
        TrsRaiseError MODULENAME, "GetActualTimeHistoryWGRpt", Err.Number, , , _
            DescArgsFrom("[OrigErrDescription]===>" & m_ActiveReports_OrigErrDescription & "<===", "[OrigErrSource]===>" & m_ActiveReports_OrigErrSource & "<==="), True
    Case Else
        TrsRaiseError MODULENAME, "GetActualTimeHistoryWGRpt", , , , , True
    End Select
End Function

Public Function GetActualTimeHistoryDeptRpt( _
                ByVal v_dtStartDate As Date _
                , ByVal v_dtEndDate As Date _
                , ByVal v_adrsSelectedDepartments As ADODB.Recordset _
                , ByVal v_adrsExceptionCodes As ADODB.Recordset _
                , ByRef r_bReportGenerated As Boolean _
                , ByVal v_bActiveEmployees As Boolean _
                , ByVal v_adrsAllEmployees As ADODB.Recordset _
                ) As ActiveReport
On Error GoTo GetActualTimeHistoryDeptRpt_EH

    'Initialize return
    r_bReportGenerated = False
    ResetActiveReportsError
    Set GetActualTimeHistoryDeptRpt = Nothing

    If (RecordsExist(v_adrsSelectedDepartments)) Then
        Set m_adrsSelectedDepartments = v_adrsSelectedDepartments
        Set m_adrsAllDepartments = m_adrsSelectedDepartments.Clone
        Set m_adrsAllEmployees = v_adrsAllEmployees
    Else
        Exit Function
    End If

    m_lReportLevel = 3

    m_dtStartDate = v_dtStartDate
    m_dtEndDate = v_dtEndDate

    Set m_adrsExceptionCodes = v_adrsExceptionCodes

    'Load XML file into the report object
    Set rptMain = CreateObject("DDActiveReports2.ActiveReport")
    rptMain.LoadLayout "C:\TRIS Solution\TRISReports\RPX Files\rptActualTimeHistoryDept.rpx"
    'Add a reference in the rpt object to this (rptException) class
    Call rptMain.AddNamedItem("MyClass", Me)

    Set rptSubReport1 = CreateObject("DDActiveReports2.ActiveReport")
    rptSubReport1.LoadLayout "C:\TRIS Solution\TRISReports\RPX Files\rptActualTimeTotalsSub.rpx"
    'Add a reference in the rpt object to this (rptException) class
    Call rptSubReport1.AddNamedItem("MyClass", Me)

    Set rptSubReport2 = CreateObject("DDActiveReports2.ActiveReport")
    rptSubReport2.LoadLayout "C:\TRIS Solution\TRISReports\RPX Files\rptActualTimeTotalsSub.rpx"
    'Add a reference in the rpt object to this (rptException) class
    Call rptSubReport2.AddNamedItem("MyClass", Me)

    rptMain.Sections("GroupFooter2").Controls("subRptEmpTotals").object = _
        rptSubReport1
    rptMain.Sections("GroupFooter1").Controls("subRptDeptTotals").object = _
        rptSubReport2
    rptMain.Sections("PageHeader").Controls("txtHdrRunDate").Text = _
        Now()
    rptMain.Sections("PageHeader").Controls("txtHdrStartDate").Text = _
        CStr(Format(v_dtStartDate, "mm/dd/yyyy"))
    rptMain.Sections("PageHeader").Controls("txtHdrEndDate").Text = _
        CStr(Format(v_dtEndDate, "mm/dd/yyyy"))

    Set m_adrsReportData = GetActualTimeHistoryDept( _
                            m_adrsSelectedDepartments _
                            , v_dtStartDate _
                            , v_dtEndDate _
                            , v_bActiveEmployees)
    If (RecordsExist(m_adrsReportData)) Then
    'place filter here if distributions are supplied
        r_bReportGenerated = True

        m_adrsReportData.Sort = "DEPT_NUM, EXCEPTION_CD"
        m_adrsReportData.MoveFirst

        CalculateDeptTotals

        m_adrsReportData.Sort = "DEPT_NUM, EMP_NUM, EXCEPTION_CD"
        m_adrsReportData.MoveFirst

        CalculateEmpDeptTotals

        m_adrsReportData.Sort = "DEPT_NUM, EMP_NUM, CHARGE_DATE, DATE_WORKED, START_TIME, END_TIME, SEQ_NUM, HOURS"
        m_adrsReportData.MoveFirst

        rptMain.Run
        
        '//check for any errors from the active reports script or the vb6 functions being called back
        If m_ActiveReports_ErrNumber <> 0 Then
            Err.Raise m_ActiveReports_ErrNumber, m_ActiveReports_ErrSource, m_ActiveReports_ErrDescription
        End If

        'Return the report object to the calling application
        Set GetActualTimeHistoryDeptRpt = rptMain
    End If

    '//-----clean up public properties used by ActiveReports-----//
    DestroyState
        
    Exit Function
GetActualTimeHistoryDeptRpt_EH:
    DestroyState
    
    Select Case Err.Number
    Case m_ActiveReports_ErrNumber
        TrsRaiseError MODULENAME, "GetActualTimeHistoryDeptRpt", Err.Number, , , _
            DescArgsFrom("[OrigErrDescription]===>" & m_ActiveReports_OrigErrDescription & "<===", "[OrigErrSource]===>" & m_ActiveReports_OrigErrSource & "<==="), True
    Case Else
        TrsRaiseError MODULENAME, "GetActualTimeHistoryDeptRpt", , , , , True
    End Select
End Function

Public Function GetActualTimeHistoryDistRpt( _
        ByVal v_dtStartDate As Date _
        , ByVal v_dtEndDate As Date _
        , ByVal v_adrsSelectedDistributions As ADODB.Recordset _
        , ByVal v_adrsExceptionCodes As ADODB.Recordset _
        , ByRef r_bReportGenerated As Boolean _
        , ByVal v_bActiveEmployees As Boolean _
        , ByVal v_adrsAllEmployees As ADODB.Recordset _
        ) As ActiveReport
On Error GoTo GetActualTimeHistoryDistRpt_EH

    'Initialize return
    r_bReportGenerated = False
    ResetActiveReportsError
    Set GetActualTimeHistoryDistRpt = Nothing

    If (RecordsExist(v_adrsSelectedDistributions)) Then
        Set m_adrsSelectedDistributions = v_adrsSelectedDistributions
        Set m_adrsAllEmployees = v_adrsAllEmployees
    Else
        Exit Function
    End If

    m_lReportLevel = 4

    m_dtStartDate = v_dtStartDate
    m_dtEndDate = v_dtEndDate

    Set m_adrsExceptionCodes = v_adrsExceptionCodes

    'Load XML file into the report object
    Set rptMain = CreateObject("DDActiveReports2.ActiveReport")
    rptMain.LoadLayout "C:\TRIS Solution\TRISReports\RPX Files\rptActualTimeHistoryDist.rpx"
    'Add a reference in the rpt object to this (rptException) class
    Call rptMain.AddNamedItem("MyClass", Me)

    Set rptSubReport1 = CreateObject("DDActiveReports2.ActiveReport")
    rptSubReport1.LoadLayout "C:\TRIS Solution\TRISReports\RPX Files\rptActualTimeTotalsSub.rpx"
    'Add a reference in the rpt object to this (rptException) class
    Call rptSubReport1.AddNamedItem("MyClass", Me)

    Set rptSubReport2 = CreateObject("DDActiveReports2.ActiveReport")
    rptSubReport2.LoadLayout "C:\TRIS Solution\TRISReports\RPX Files\rptActualTimeTotalsSub.rpx"
    'Add a reference in the rpt object to this (rptException) class
    Call rptSubReport2.AddNamedItem("MyClass", Me)

    rptMain.Sections("GroupFooter2").Controls("subRptEmpTotals").object = _
        rptSubReport1
    rptMain.Sections("GroupFooter1").Controls("subRptDistTotals").object = _
        rptSubReport2
    rptMain.Sections("PageHeader").Controls("txtHdrRunDate").Text = _
        Now()
    rptMain.Sections("PageHeader").Controls("txtHdrStartDate").Text = _
        CStr(Format(v_dtStartDate, "mm/dd/yyyy"))
    rptMain.Sections("PageHeader").Controls("txtHdrEndDate").Text = _
        CStr(Format(v_dtEndDate, "mm/dd/yyyy"))

    Set m_adrsReportData = GetActualTimeHistoryDist( _
                            m_adrsSelectedDistributions _
                            , v_dtStartDate _
                            , v_dtEndDate _
                            , v_bActiveEmployees)
    If (RecordsExist(m_adrsReportData)) Then
        r_bReportGenerated = True

        m_adrsReportData.Sort = "DISTRIBUTION_ID, DISTRIBUTION_TYPE, EXCEPTION_CD"
        m_adrsReportData.MoveFirst

        CalculateDistTotals

        m_adrsReportData.Sort = "DISTRIBUTION_ID, DISTRIBUTION_TYPE, EMP_NUM, EXCEPTION_CD"
        m_adrsReportData.MoveFirst

        CalculateEmpDistTotals

        m_adrsReportData.Sort = "DISTRIBUTION_ID, DISTRIBUTION_TYPE, EMP_NUM, EXCEPTION_CD"
        m_adrsReportData.MoveFirst

        'Return the report object to the calling application
        rptMain.Run
        
        '//check for any errors from the active reports script or the vb6 functions being called back
        If m_ActiveReports_ErrNumber <> 0 Then
            Err.Raise m_ActiveReports_ErrNumber, m_ActiveReports_ErrSource, m_ActiveReports_ErrDescription
        End If
        
        Set GetActualTimeHistoryDistRpt = rptMain
    End If

    '//-----clean up public properties used by ActiveReports-----//
    DestroyState
        
    Exit Function
GetActualTimeHistoryDistRpt_EH:
    DestroyState
    
    Select Case Err.Number
    Case m_ActiveReports_ErrNumber
        TrsRaiseError MODULENAME, "GetActualTimeHistoryDistRpt", Err.Number, , , _
            DescArgsFrom("[OrigErrDescription]===>" & m_ActiveReports_OrigErrDescription & "<===", "[OrigErrSource]===>" & m_ActiveReports_OrigErrSource & "<==="), True
    Case Else
        TrsRaiseError MODULENAME, "GetActualTimeHistoryDistRpt", , , , , True
    End Select
End Function

Public Function GetWGDesc(ByVal v_sWGNum As String) As String
On Error GoTo GetWGDesc_EH
Dim l_objWorkgroup As dbWorkGroup.cdWorkGroup
Dim l_adrsWorkgroup As ADODB.Recordset

    If (Trim(v_sWGNum) <> "") Then
        With m_adrsAllWorkgroups
            .MoveFirst

            .Find "WG_NUM = " & CLng(Trim(v_sWGNum))
            If Not (.EOF) Then
                GetWGDesc = Trim(!WG_DESC)
            Else
                Set l_objWorkgroup = CreateObject("dbWorkGroup.cdWorkGroup")
                Set l_adrsWorkgroup = l_objWorkgroup.GetWorkGroupAttributes(CVar(v_sWGNum))
                If (RecordsExist(l_adrsWorkgroup)) Then
                    l_adrsWorkgroup.MoveFirst
                    GetWGDesc = l_adrsWorkgroup!WG_DESC
                End If
            End If
        End With
    End If

    Set l_objWorkgroup = Nothing
    Set l_adrsWorkgroup = Nothing

    Exit Function
GetWGDesc_EH:
    Set l_objWorkgroup = Nothing
    Set l_adrsWorkgroup = Nothing

    TrsRaiseError MODULENAME _
        , "GetWGDesc" _
        , , , , , True
End Function
Public Function GetDepartmentName(ByVal v_sDeptNum As String) As String
On Error GoTo GetDepartmentName_EH

    If (Trim(v_sDeptNum) = Trim(m_sHoldDeptNum)) Then
        GetDepartmentName = Trim(m_sHoldDeptName)
    Else
        With m_adrsAllDepartments
            .MoveFirst

            .Find "DEPT_NUM = '" & Trim(v_sDeptNum) & "'"
            If Not (.EOF) Then
                GetDepartmentName = Trim(!DEPT_NAME)

                m_sHoldDeptNum = Trim(v_sDeptNum)
                m_sHoldDeptName = Trim(!DEPT_NAME)
            Else
                m_sHoldDeptNum = ""
                m_sHoldDeptName = ""

                GetDepartmentName = ""
            End If
        End With
    End If

'//test active reports error handling
'//Dim o As ADODB.Recordset
'//o.MoveFirst

    Exit Function
GetDepartmentName_EH:
    TrsRaiseError MODULENAME _
        , "GetDepartmentName" _
        , , , , , True
End Function
Public Function GetActualTimeReportData() As ADODB.Recordset
On Error GoTo GetActualTimeReportData_EH

'//    TrsTraceError MODULENAME, "GetActualTimeReportData", , , , DescArgsFrom("called back from active reports script"), True
    
    m_adrsReportData.MoveFirst
    'Call to pass the report data into the
    Set GetActualTimeReportData = m_adrsReportData

    Exit Function
GetActualTimeReportData_EH:
    TrsRaiseError MODULENAME _
        , "GetActualTimeReportData" _
        , , , , , True

End Function
Public Function GetEmployeeName(ByVal v_lEmpNum As Long) As String
On Error GoTo GetEmployeeName_EH

'//TrsTraceError MODULENAME, "GetEmployeeName", , , , DescArgsFrom("called back from active reports script", "v_lEmpNum=" & v_lEmpNum), True
    GetEmployeeName = ""

    If (v_lEmpNum > 0) Then
        With m_adrsAllEmployees
            .MoveFirst

            .Find "EMP_NUM = " & v_lEmpNum
            If Not .EOF Then
                GetEmployeeName = Trim(!EMP_NAME)
            End If

            .MoveFirst
        End With
    End If

    Exit Function
GetEmployeeName_EH:
    TrsRaiseError MODULENAME _
        , "GetEmployeeName" _
        , , , , , True

End Function

'//new length is 60, but don't know who might be calling this, so setting it back to only 30 chars
Public Function GetDistributionDesc(ByVal v_sDistributionId As String _
                , ByVal v_sDistributionType As String _
                ) As String
On Error GoTo GetDistributionDesc_EH
Dim l_sFilter As String

    GetDistributionDesc = ""

    If (Trim(v_sDistributionId) <> "") Then
        With m_adrsSelectedDistributions
            .Filter = adFilterNone

            l_sFilter = "DISTRIBUTION_ID = '" & v_sDistributionId & "'" & _
                " AND DISTRIBUTION_TYPE = '" & v_sDistributionType & "'"
            .Filter = MDAC25FilterFix(l_sFilter)

            If (RecordsExist(m_adrsSelectedDistributions)) Then
                GetDistributionDesc = Left(Trim(!DISTRIBUTION_DESC), 30) '//new length is 60, but don't know who might be calling this, so setting it back to only 30 chars
            End If

            .Filter = adFilterNone
        End With
    End If

    Exit Function
GetDistributionDesc_EH:
    TrsRaiseError MODULENAME _
        , "GetDistributionDesc" _
        , , , , , True

End Function
Public Function GetActualTimeTotals() As ADODB.Recordset
On Error GoTo GetActualTimeTotals_EH
Dim l_sFilter As String

'//TrsTraceError MODULENAME, "GetActualTimeTotals", , , , DescArgsFrom("called back from active reports script", "m_lProcessLevel=" & m_lProcessLevel, "m_lReportLevel=" & m_lReportLevel), True
'//Err.Source = ""
    'Report Level
    Select Case (m_lProcessLevel)
        Case 1  'Employee Totals
            Select Case (m_lReportLevel)
                Case 1  'Employee Report
                    With m_adrsEmpActualTimeTotals
                        .Filter = adFilterNone

                        l_sFilter = "EMP_NUM = " & m_lCurrentEmpNum
                        .Filter = MDAC25FilterFix(l_sFilter)
                        If Not (.EOF) Then
                            .MoveFirst
                        End If

                        Set GetActualTimeTotals = m_adrsEmpActualTimeTotals
                    End With

                Case 2 'Workgroup report
                    With m_adrsEmpActualTimeTotals
                        .Filter = adFilterNone

                        l_sFilter = "EMP_NUM = " & m_lCurrentEmpNum & _
                                " AND WG_NUM = " & m_lWorkgroup
                        .Filter = MDAC25FilterFix(l_sFilter)
                        If Not (.EOF) Then
                            .MoveFirst
                        End If

                        Set GetActualTimeTotals = m_adrsEmpActualTimeTotals
                    End With

                Case 3 'Department report
                    With m_adrsEmpActualTimeTotals
                        .Filter = adFilterNone

                        l_sFilter = "EMP_NUM = " & m_lCurrentEmpNum & _
                                " AND DEPT_NUM = '" & m_sDeptNum & "'"
                        .Filter = MDAC25FilterFix(l_sFilter)
                        If Not (.EOF) Then
                            .MoveFirst
                        End If

                        Set GetActualTimeTotals = m_adrsEmpActualTimeTotals
                    End With

                Case 4 'Distribution report
                    With m_adrsEmpActualTimeTotals
                        .Filter = adFilterNone
                        l_sFilter = "EMP_NUM = " & m_lCurrentEmpNum & _
                                " AND DISTRIBUTION_ID = '" & m_sDistributionId & "'" & _
                                " AND DISTRIBUTION_TYPE = '" & m_sDistributionType & "'"
                        .Filter = MDAC25FilterFix(l_sFilter)
                        If Not (.EOF) Then
                            .MoveFirst
                        End If

                        Set GetActualTimeTotals = m_adrsEmpActualTimeTotals
                    End With
            End Select

        Case 2 'Workgroup Totals
            With m_adrsWGActualTimeTotals
                .Filter = adFilterNone

                l_sFilter = "WG_NUM = " & m_lWorkgroup
                .Filter = MDAC25FilterFix(l_sFilter)
                If Not (.EOF) Then
                    .MoveFirst
                End If

                Set GetActualTimeTotals = m_adrsWGActualTimeTotals
            End With

        Case 3 'Department Totals
            With m_adrsDeptActualTimeTotals
                .Filter = adFilterNone

                l_sFilter = "DEPT_NUM = '" & m_sDeptNum & "'"
                .Filter = MDAC25FilterFix(l_sFilter)
                If Not (.EOF) Then
                    .MoveFirst
                End If

                Set GetActualTimeTotals = m_adrsDeptActualTimeTotals
            End With

        Case 4 'Distribution Totals
            With m_adrsDistActualTimeTotals
                .Filter = adFilterNone

                l_sFilter = "DISTRIBUTION_ID = '" & m_sDistributionId & "'" & _
                    " AND DISTRIBUTION_TYPE = '" & m_sDistributionType & "'"
                .Filter = MDAC25FilterFix(l_sFilter)
                If Not (.EOF) Then
                    .MoveFirst
                End If

                Set GetActualTimeTotals = m_adrsDistActualTimeTotals
            End With
    End Select

    Exit Function
GetActualTimeTotals_EH:
    TrsRaiseError MODULENAME _
        , "GetActualTimeTotals" _
        , , , , , True

End Function
Private Function GetActualTimeHistoryEmp( _
                ByVal v_adrsSelectedEmployees As ADODB.Recordset _
                , ByVal v_dtStartDate As Date _
                , ByVal v_dtEndDate As Date _
                , ByVal v_bActiveEmployees As Boolean _
                ) As ADODB.Recordset
On Error GoTo GetActualTimeHistoryEmp_EH
Dim l_obj As busTimecard.cbsTimecard
Dim l_objAdj As dbAdjustments.cdAdjustments
Dim l_lARREmps() As Long, l_lEmpIndex As Long
Dim ladrsActualTime As ADODB.Recordset
Dim ladrsAdj As ADODB.Recordset
Dim ladrsActualTimeHistoryEmp As ADODB.Recordset

    Set GetActualTimeHistoryEmp = Nothing

    If Not (RecordsExist(v_adrsSelectedEmployees)) Then
        Exit Function
    Else
        With v_adrsSelectedEmployees
            .MoveFirst

            ReDim l_lARREmps(.RecordCount - 1)

            For l_lEmpIndex = 1 To .RecordCount
                l_lARREmps(l_lEmpIndex - 1) = !EMP_NUM

                .MoveNext
            Next l_lEmpIndex
        End With
    End If
    
    Set ladrsActualTimeHistoryEmp = New ADODB.Recordset
    With ladrsActualTimeHistoryEmp
        .Fields.Append "EMP_NUM", adInteger, , m_conFIELDATTR
        .Fields.Append "DEPT_NUM", adChar, 3, m_conFIELDATTR
        .Fields.Append "OT_OPP_NUM", adInteger, , m_conFIELDATTR
        .Fields.Append "CHARGE_DATE", adDate, , m_conFIELDATTR
        .Fields.Append "DATE_WORKED", adDate, , m_conFIELDATTR
        .Fields.Append "START_TIME", adChar, 5, m_conFIELDATTR
        .Fields.Append "END_TIME", adChar, 5, m_conFIELDATTR
        .Fields.Append "LOC_CD", adInteger, , m_conFIELDATTR
        .Fields.Append "JOB_NUM", adChar, 6, m_conFIELDATTR
        .Fields.Append "HOURS", adDecimal, , m_conFIELDATTR
        .Fields.Item("HOURS").Precision = 3
        .Fields.Item("HOURS").NumericScale = 1
        .Fields.Append "WG_NUM", adInteger, 5, m_conFIELDATTR
        .Fields.Append "REL_CD", adChar, 3, m_conFIELDATTR
        .Fields.Append "COMP_CD", adChar, 3, m_conFIELDATTR
        .Fields.Append "COMP_SHIFT_CD", adChar, 3, m_conFIELDATTR
        .Fields.Append "EXCEPTION_CD", adChar, 3, m_conFIELDATTR
        .Fields.Append "EARNINGS_CD", adChar, 3, m_conFIELDATTR
        .Fields.Append "OVER_EXC_CD", adChar, 3, m_conFIELDATTR
        .Fields.Append "CASE_NUM", adInteger, , m_conFIELDATTR
        .Fields.Append "START_LOC_CD", adInteger, , m_conFIELDATTR
        .Fields.Append "END_LOC_CD", adInteger, , m_conFIELDATTR
        .Fields.Append "AMOUNT", adDecimal, , m_conFIELDATTR
        .Fields.Item("AMOUNT").Precision = 7
        .Fields.Item("AMOUNT").NumericScale = 2
        .Fields.Append "APPR_BY_EMP_NUM", adInteger, , m_conFIELDATTR
        .Fields.Append "DISTRIBUTION_TYPE", adChar, 1, m_conFIELDATTR
        .Fields.Append "DISTRIBUTION_ID", adChar, 39, m_conFIELDATTR
        .Fields.Append "OT_REASON_CD", adChar, 3, m_conFIELDATTR
        .Fields.Append "ADJ_FLAG", adChar, 3, m_conFIELDATTR
        .Fields.Append "SEQ_NUM", adInteger, , m_conFIELDATTR
        .Fields.Append "ADJ_TYPE", adChar, 1, m_conFIELDATTR
        .Fields.Append "TRAVEL_PAY_IND", adChar, 1, m_conFIELDATTR
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open
    End With

    '//internally within DistMgr.GetDistributionList (due to OTTC) the description length must be greater than or equal to 39
    '//since distribution_id, description, and job_task will all be 60 chars
    '//but timecard will return distribution_id as 39 chars
    Set l_obj = CreateObject("busTimecard.cbsTimecard")
    Set ladrsActualTime = l_obj.GetActualTimeDetailWithOtherPay2B( _
                                    l_lARREmps _
                                    , v_dtStartDate _
                                    , v_dtEndDate, 60)
    Set l_obj = Nothing
    
    If RecordsExist(ladrsActualTime) Then
        With ladrsActualTime
            .MoveFirst
            While Not ladrsActualTime.EOF
                ladrsActualTimeHistoryEmp.AddNew
                ladrsActualTimeHistoryEmp.Fields("EMP_NUM") = !EMP_NUM
                ladrsActualTimeHistoryEmp.Fields("DEPT_NUM") = !DEPT_NUM
                ladrsActualTimeHistoryEmp.Fields("OT_OPP_NUM") = !OT_OPP_NUM
                ladrsActualTimeHistoryEmp.Fields("CHARGE_DATE") = !CHARGE_DATE
                ladrsActualTimeHistoryEmp.Fields("DATE_WORKED") = !DATE_WORKED
                ladrsActualTimeHistoryEmp.Fields("START_TIME") = !START_TIME
                ladrsActualTimeHistoryEmp.Fields("END_TIME") = !END_TIME
                ladrsActualTimeHistoryEmp.Fields("LOC_CD") = !LOC_CD
                ladrsActualTimeHistoryEmp.Fields("JOB_NUM") = !JOB_NUM
                ladrsActualTimeHistoryEmp.Fields("HOURS") = !Hours
                ladrsActualTimeHistoryEmp.Fields("WG_NUM") = !WG_NUM
                ladrsActualTimeHistoryEmp.Fields("REL_CD") = !REL_CD
                ladrsActualTimeHistoryEmp.Fields("COMP_CD") = !COMP_CD
                ladrsActualTimeHistoryEmp.Fields("COMP_SHIFT_CD") = !COMP_SHIFT_CD
                ladrsActualTimeHistoryEmp.Fields("EXCEPTION_CD") = !EXCEPTION_CD
                ladrsActualTimeHistoryEmp.Fields("EARNINGS_CD") = !EARNINGS_CD
                ladrsActualTimeHistoryEmp.Fields("OVER_EXC_CD") = !OVER_EXC_CD
                ladrsActualTimeHistoryEmp.Fields("CASE_NUM") = !CASE_NUM
                ladrsActualTimeHistoryEmp.Fields("START_LOC_CD") = !START_LOC_CD
                ladrsActualTimeHistoryEmp.Fields("END_LOC_CD") = !END_LOC_CD
                ladrsActualTimeHistoryEmp.Fields("AMOUNT") = !AMOUNT
                ladrsActualTimeHistoryEmp.Fields("APPR_BY_EMP_NUM") = !APPR_BY_EMP_NUM
                ladrsActualTimeHistoryEmp.Fields("DISTRIBUTION_TYPE") = !DISTRIBUTION_TYPE
                ladrsActualTimeHistoryEmp.Fields("DISTRIBUTION_ID") = !DISTRIBUTION_ID
                ladrsActualTimeHistoryEmp.Fields("OT_REASON_CD") = !OT_REASON_CD
                ladrsActualTimeHistoryEmp.Fields("ADJ_FLAG") = "ACT"
                ladrsActualTimeHistoryEmp.Fields("SEQ_NUM") = 0
                ladrsActualTimeHistoryEmp.Fields("ADJ_TYPE") = ""
                ladrsActualTimeHistoryEmp.Fields("TRAVEL_PAY_IND") = !TRAVEL_PAY_IND
                ladrsActualTimeHistoryEmp.Update
                .MoveNext
            Wend
        End With
    End If
            
    Set l_objAdj = CreateObject("dbAdjustments.cdAdjustments")
    Set ladrsAdj = l_objAdj.GetAdjustmentDetails(l_lARREmps _
                                                 , v_dtStartDate _
                                                 , v_dtEndDate)
    Set l_objAdj = Nothing
    
    If RecordsExist(ladrsAdj) Then
        With ladrsAdj
            .MoveFirst
            While Not ladrsAdj.EOF
                ladrsActualTimeHistoryEmp.AddNew
                ladrsActualTimeHistoryEmp.Fields("EMP_NUM") = !EMP_NUM
                ladrsActualTimeHistoryEmp.Fields("DEPT_NUM") = !DEPT_NUM
                ladrsActualTimeHistoryEmp.Fields("OT_OPP_NUM") = !OT_OPP_NUM
                ladrsActualTimeHistoryEmp.Fields("CHARGE_DATE") = !DATE_WORKED
                ladrsActualTimeHistoryEmp.Fields("DATE_WORKED") = !DATE_WORKED
                ladrsActualTimeHistoryEmp.Fields("START_TIME") = !START_TIME
                ladrsActualTimeHistoryEmp.Fields("END_TIME") = !END_TIME
                ladrsActualTimeHistoryEmp.Fields("LOC_CD") = !LOC_CD
                ladrsActualTimeHistoryEmp.Fields("JOB_NUM") = !JOB_NUM
                ladrsActualTimeHistoryEmp.Fields("HOURS") = !Hours
                ladrsActualTimeHistoryEmp.Fields("WG_NUM") = !WG_NUM
                ladrsActualTimeHistoryEmp.Fields("REL_CD") = !REL_CD
                ladrsActualTimeHistoryEmp.Fields("COMP_CD") = !COMP_CD
                ladrsActualTimeHistoryEmp.Fields("COMP_SHIFT_CD") = !SHIFT_COMP
                ladrsActualTimeHistoryEmp.Fields("EXCEPTION_CD") = !EXCEPTION_CD
                ladrsActualTimeHistoryEmp.Fields("EARNINGS_CD") = !EARNINGS_CD
                ladrsActualTimeHistoryEmp.Fields("OVER_EXC_CD") = !OVER_EXC_CD
                ladrsActualTimeHistoryEmp.Fields("CASE_NUM") = !CASE_NUM
                ladrsActualTimeHistoryEmp.Fields("START_LOC_CD") = !START_LOC_CD
                ladrsActualTimeHistoryEmp.Fields("END_LOC_CD") = !END_LOC_CD
                ladrsActualTimeHistoryEmp.Fields("AMOUNT") = !AMOUNT
                ladrsActualTimeHistoryEmp.Fields("APPR_BY_EMP_NUM") = !APPR_BY_MANAGER
                ladrsActualTimeHistoryEmp.Fields("DISTRIBUTION_TYPE") = !DISTRIBUTION_TYPE
                ladrsActualTimeHistoryEmp.Fields("DISTRIBUTION_ID") = !DISTRIBUTION_ID
                ladrsActualTimeHistoryEmp.Fields("OT_REASON_CD") = !OT_REASON_CD
                ladrsActualTimeHistoryEmp.Fields("ADJ_FLAG") = "ADJ"
                ladrsActualTimeHistoryEmp.Fields("SEQ_NUM") = !SEQ_NUM
                ladrsActualTimeHistoryEmp.Fields("ADJ_TYPE") = !ADJ_TYPE
                If !TRAVEL_PAY_HRS <> 0# Then
                    ladrsActualTimeHistoryEmp.Fields("TRAVEL_PAY_IND") = "Y"
                Else
                    ladrsActualTimeHistoryEmp.Fields("TRAVEL_PAY_IND") = "N"
                End If
                ladrsActualTimeHistoryEmp.Update
                .MoveNext
            Wend
        End With
    End If
    
    Set GetActualTimeHistoryEmp = ladrsActualTimeHistoryEmp

    Set ladrsActualTime = Nothing
    Set ladrsAdj = Nothing
    Set ladrsActualTimeHistoryEmp = Nothing

    Exit Function
GetActualTimeHistoryEmp_EH:
    Set l_obj = Nothing
    Set l_objAdj = Nothing
    Set ladrsActualTime = Nothing
    Set ladrsAdj = Nothing
    Set ladrsActualTimeHistoryEmp = Nothing

    TrsRaiseError MODULENAME _
        , "GetActualTimeHistoryEmp"

End Function
Private Function GetActualTimeHistoryEmpbyDist( _
                ByVal v_adrsSelectedEmployees As ADODB.Recordset _
                , ByVal v_dtStartDate As Date _
                , ByVal v_dtEndDate As Date _
                , ByVal v_adrsDistributions As ADODB.Recordset _
                ) As ADODB.Recordset
On Error GoTo GetActualTimeHistoryEmpbyDist_EH
Dim l_obj As dbTimeCard.cdTimeCard
Dim l_objAdj As dbAdjustments.cdAdjustments
Dim l_lARREmps() As Long, l_lEmpIndex As Long
Dim ladrsActualTime As ADODB.Recordset
Dim ladrsAdj As ADODB.Recordset
Dim ladrsActualTimeHistoryEmpbyDist As ADODB.Recordset

    Set GetActualTimeHistoryEmpbyDist = Nothing

    If Not (RecordsExist(v_adrsSelectedEmployees)) Then
        Exit Function
    Else
        With v_adrsSelectedEmployees
            .MoveFirst

            ReDim l_lARREmps(.RecordCount - 1)

            For l_lEmpIndex = 1 To .RecordCount
                l_lARREmps(l_lEmpIndex - 1) = !EMP_NUM

                .MoveNext
            Next l_lEmpIndex
        End With
    End If
    
    Set ladrsActualTimeHistoryEmpbyDist = New ADODB.Recordset
    With ladrsActualTimeHistoryEmpbyDist
        .Fields.Append "EMP_NUM", adInteger, , m_conFIELDATTR
        .Fields.Append "DEPT_NUM", adChar, 3, m_conFIELDATTR
        .Fields.Append "OT_OPP_NUM", adInteger, , m_conFIELDATTR
        .Fields.Append "CHARGE_DATE", adDate, , m_conFIELDATTR
        .Fields.Append "DATE_WORKED", adDate, , m_conFIELDATTR
        .Fields.Append "START_TIME", adChar, 5, m_conFIELDATTR
        .Fields.Append "END_TIME", adChar, 5, m_conFIELDATTR
        .Fields.Append "LOC_CD", adInteger, , m_conFIELDATTR
        .Fields.Append "JOB_NUM", adChar, 6, m_conFIELDATTR
        .Fields.Append "HOURS", adDecimal, , m_conFIELDATTR
        .Fields.Item("HOURS").Precision = 3
        .Fields.Item("HOURS").NumericScale = 1
        .Fields.Append "WG_NUM", adInteger, 5, m_conFIELDATTR
        .Fields.Append "REL_CD", adChar, 3, m_conFIELDATTR
        .Fields.Append "COMP_CD", adChar, 3, m_conFIELDATTR
        .Fields.Append "COMP_SHIFT_CD", adChar, 3, m_conFIELDATTR
        .Fields.Append "EXCEPTION_CD", adChar, 3, m_conFIELDATTR
        .Fields.Append "EARNINGS_CD", adChar, 3, m_conFIELDATTR
        .Fields.Append "OVER_EXC_CD", adChar, 3, m_conFIELDATTR
        .Fields.Append "CASE_NUM", adInteger, , m_conFIELDATTR
        .Fields.Append "START_LOC_CD", adInteger, , m_conFIELDATTR
        .Fields.Append "END_LOC_CD", adInteger, , m_conFIELDATTR
        .Fields.Append "AMOUNT", adDecimal, , m_conFIELDATTR
        .Fields.Item("AMOUNT").Precision = 7
        .Fields.Item("AMOUNT").NumericScale = 2
        .Fields.Append "APPR_BY_EMP_NUM", adInteger, , m_conFIELDATTR
        .Fields.Append "DISTRIBUTION_TYPE", adChar, 1, m_conFIELDATTR
        .Fields.Append "DISTRIBUTION_ID", adChar, 39, m_conFIELDATTR
        .Fields.Append "OT_REASON_CD", adChar, 3, m_conFIELDATTR
        .Fields.Append "ADJ_FLAG", adChar, 3, m_conFIELDATTR
        .Fields.Append "SEQ_NUM", adInteger, , m_conFIELDATTR
        .Fields.Append "ADJ_TYPE", adChar, 1, m_conFIELDATTR
        .Fields.Append "TRAVEL_PAY_IND", adChar, 1, m_conFIELDATTR
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open
    End With

    Set l_obj = CreateObject("dbTimecard.cdTimecard")
    Set ladrsActualTime = l_obj.GetDistributionActualTimeEntries2( _
                                      v_adrsDistributions _
                                    , v_dtStartDate _
                                    , v_dtEndDate _
                                    , l_lARREmps)
    Set l_obj = Nothing
    
    If RecordsExist(ladrsActualTime) Then
        With ladrsActualTime
            .MoveFirst
            While Not ladrsActualTime.EOF
                ladrsActualTimeHistoryEmpbyDist.AddNew
                ladrsActualTimeHistoryEmpbyDist.Fields("EMP_NUM") = !EMP_NUM
                ladrsActualTimeHistoryEmpbyDist.Fields("DEPT_NUM") = !DEPT_NUM
                ladrsActualTimeHistoryEmpbyDist.Fields("OT_OPP_NUM") = !OT_OPP_NUM
                ladrsActualTimeHistoryEmpbyDist.Fields("CHARGE_DATE") = !CHARGE_DATE
                ladrsActualTimeHistoryEmpbyDist.Fields("DATE_WORKED") = !DATE_WORKED
                ladrsActualTimeHistoryEmpbyDist.Fields("START_TIME") = !START_TIME
                ladrsActualTimeHistoryEmpbyDist.Fields("END_TIME") = !END_TIME
                ladrsActualTimeHistoryEmpbyDist.Fields("LOC_CD") = !LOC_CD
                ladrsActualTimeHistoryEmpbyDist.Fields("JOB_NUM") = !JOB_NUM
                ladrsActualTimeHistoryEmpbyDist.Fields("HOURS") = !Hours
                ladrsActualTimeHistoryEmpbyDist.Fields("WG_NUM") = !WG_NUM
                ladrsActualTimeHistoryEmpbyDist.Fields("REL_CD") = !REL_CD
                ladrsActualTimeHistoryEmpbyDist.Fields("COMP_CD") = !COMP_CD
                ladrsActualTimeHistoryEmpbyDist.Fields("COMP_SHIFT_CD") = !COMP_SHIFT_CD
                ladrsActualTimeHistoryEmpbyDist.Fields("EXCEPTION_CD") = !EXCEPTION_CD
                ladrsActualTimeHistoryEmpbyDist.Fields("EARNINGS_CD") = !EARNINGS_CD
                ladrsActualTimeHistoryEmpbyDist.Fields("OVER_EXC_CD") = !OVER_EXC_CD
                ladrsActualTimeHistoryEmpbyDist.Fields("CASE_NUM") = !CASE_NUM
                ladrsActualTimeHistoryEmpbyDist.Fields("START_LOC_CD") = !START_LOC_CD
                ladrsActualTimeHistoryEmpbyDist.Fields("END_LOC_CD") = !END_LOC_CD
                ladrsActualTimeHistoryEmpbyDist.Fields("AMOUNT") = !AMOUNT
                ladrsActualTimeHistoryEmpbyDist.Fields("APPR_BY_EMP_NUM") = !APPR_BY_EMP_NUM
                ladrsActualTimeHistoryEmpbyDist.Fields("DISTRIBUTION_TYPE") = !DISTRIBUTION_TYPE
                ladrsActualTimeHistoryEmpbyDist.Fields("DISTRIBUTION_ID") = !DISTRIBUTION_ID
                ladrsActualTimeHistoryEmpbyDist.Fields("OT_REASON_CD") = !OT_REASON_CD
                ladrsActualTimeHistoryEmpbyDist.Fields("ADJ_FLAG") = "ACT"
                ladrsActualTimeHistoryEmpbyDist.Fields("SEQ_NUM") = 0
                ladrsActualTimeHistoryEmpbyDist.Fields("ADJ_TYPE") = ""
                ladrsActualTimeHistoryEmpbyDist.Fields("TRAVEL_PAY_IND") = !TRAVEL_PAY_IND
                ladrsActualTimeHistoryEmpbyDist.Update
                .MoveNext
            Wend
        End With
    End If
            
    Set l_objAdj = CreateObject("dbAdjustments.cdAdjustments")
    Set ladrsAdj = l_objAdj.GetDistributionAdjustmentDetails(l_lARREmps _
                                                 , v_dtStartDate _
                                                 , v_dtEndDate _
                                                 , v_adrsDistributions)
    Set l_objAdj = Nothing
    
    If RecordsExist(ladrsAdj) Then
        With ladrsAdj
            .MoveFirst
            While Not ladrsAdj.EOF
                ladrsActualTimeHistoryEmpbyDist.AddNew
                ladrsActualTimeHistoryEmpbyDist.Fields("EMP_NUM") = !EMP_NUM
                ladrsActualTimeHistoryEmpbyDist.Fields("DEPT_NUM") = !DEPT_NUM
                ladrsActualTimeHistoryEmpbyDist.Fields("OT_OPP_NUM") = !OT_OPP_NUM
                ladrsActualTimeHistoryEmpbyDist.Fields("CHARGE_DATE") = !DATE_WORKED
                ladrsActualTimeHistoryEmpbyDist.Fields("DATE_WORKED") = !DATE_WORKED
                ladrsActualTimeHistoryEmpbyDist.Fields("START_TIME") = !START_TIME
                ladrsActualTimeHistoryEmpbyDist.Fields("END_TIME") = !END_TIME
                ladrsActualTimeHistoryEmpbyDist.Fields("LOC_CD") = !LOC_CD
                ladrsActualTimeHistoryEmpbyDist.Fields("JOB_NUM") = !JOB_NUM
                ladrsActualTimeHistoryEmpbyDist.Fields("HOURS") = !Hours
                ladrsActualTimeHistoryEmpbyDist.Fields("WG_NUM") = !WG_NUM
                ladrsActualTimeHistoryEmpbyDist.Fields("REL_CD") = !REL_CD
                ladrsActualTimeHistoryEmpbyDist.Fields("COMP_CD") = !COMP_CD
                ladrsActualTimeHistoryEmpbyDist.Fields("COMP_SHIFT_CD") = !SHIFT_COMP
                ladrsActualTimeHistoryEmpbyDist.Fields("EXCEPTION_CD") = !EXCEPTION_CD
                ladrsActualTimeHistoryEmpbyDist.Fields("EARNINGS_CD") = !EARNINGS_CD
                ladrsActualTimeHistoryEmpbyDist.Fields("OVER_EXC_CD") = !OVER_EXC_CD
                ladrsActualTimeHistoryEmpbyDist.Fields("CASE_NUM") = !CASE_NUM
                ladrsActualTimeHistoryEmpbyDist.Fields("START_LOC_CD") = !START_LOC_CD
                ladrsActualTimeHistoryEmpbyDist.Fields("END_LOC_CD") = !END_LOC_CD
                ladrsActualTimeHistoryEmpbyDist.Fields("AMOUNT") = !AMOUNT
                ladrsActualTimeHistoryEmpbyDist.Fields("APPR_BY_EMP_NUM") = !APPR_BY_MANAGER
                ladrsActualTimeHistoryEmpbyDist.Fields("DISTRIBUTION_TYPE") = !DISTRIBUTION_TYPE
                ladrsActualTimeHistoryEmpbyDist.Fields("DISTRIBUTION_ID") = !DISTRIBUTION_ID
                ladrsActualTimeHistoryEmpbyDist.Fields("OT_REASON_CD") = !OT_REASON_CD
                ladrsActualTimeHistoryEmpbyDist.Fields("ADJ_FLAG") = "ADJ"
                ladrsActualTimeHistoryEmpbyDist.Fields("SEQ_NUM") = !SEQ_NUM
                ladrsActualTimeHistoryEmpbyDist.Fields("ADJ_TYPE") = !ADJ_TYPE
                If !TRAVEL_PAY_HRS <> 0# Then
                    ladrsActualTimeHistoryEmpbyDist.Fields("TRAVEL_PAY_IND") = "Y"
                Else
                    ladrsActualTimeHistoryEmpbyDist.Fields("TRAVEL_PAY_IND") = "N"
                End If
                ladrsActualTimeHistoryEmpbyDist.Update
                .MoveNext
            Wend
        End With
    End If
    
    Set GetActualTimeHistoryEmpbyDist = ladrsActualTimeHistoryEmpbyDist

    Exit Function
GetActualTimeHistoryEmpbyDist_EH:
    Set l_obj = Nothing
    Set l_objAdj = Nothing
    Set ladrsActualTime = Nothing
    Set ladrsAdj = Nothing
    Set ladrsActualTimeHistoryEmpbyDist = Nothing

    TrsRaiseError MODULENAME _
        , "GetActualTimeHistoryEmpbyDist"

End Function
Private Function GetActualTimeHistoryDist( _
                ByVal v_adrsSelectedDistributions As ADODB.Recordset _
                , ByVal v_dtStartDate As Date _
                , ByVal v_dtEndDate As Date _
                , ByVal v_bActiveEmployees As Boolean _
                ) As ADODB.Recordset
On Error GoTo GetActualTimeHistoryDist_EH
Dim l_obj As dbTimeCard.cdTimeCard

    Set l_obj = CreateObject("dbTimecard.cdTimecard")
    Set GetActualTimeHistoryDist = l_obj.GetDistributionActualTimeEntries( _
                                    v_adrsSelectedDistributions _
                                    , v_dtStartDate _
                                    , v_dtEndDate _
                                    , v_bActiveEmployees)
    Set l_obj = Nothing

    Exit Function
GetActualTimeHistoryDist_EH:
    Set l_obj = Nothing

    TrsRaiseError MODULENAME _
        , "GetActualTimeHistoryDist"

End Function
Private Function GetActualTimeHistoryDept( _
                ByVal v_adrsSelectedDepartments As ADODB.Recordset _
                , ByVal v_dtStartDate As Date _
                , ByVal v_dtEndDate As Date _
                , ByVal v_bActiveEmployees As Boolean _
                ) As ADODB.Recordset
On Error GoTo GetActualTimeHistoryDept_EH
Dim l_obj As busTimecard.cbsTimecard
Dim l_objWorkgroup As dbWorkGroup.cdWorkGroup
Dim l_objAdj As dbAdjustments.cdAdjustments
Dim ladrsActualTimeHistoryDept As ADODB.Recordset
Dim ladrsActualTime As ADODB.Recordset
Dim ladrsAdj As ADODB.Recordset

Dim l_adrsEmployees As New ADODB.Recordset
Dim l_lARREmployees() As Long, l_sArrDepts() As String
Dim l_lArrayIndex As Long

    Set GetActualTimeHistoryDept = Nothing

    If Not (RecordsExist(v_adrsSelectedDepartments)) Then
        Exit Function
    Else
        With v_adrsSelectedDepartments
            .MoveFirst

            ReDim l_sArrDepts(.RecordCount - 1)
            For l_lArrayIndex = 1 To .RecordCount
                l_sArrDepts(l_lArrayIndex - 1) = !DEPT_NUM

                .MoveNext
            Next l_lArrayIndex
        End With
    End If

    Set l_objWorkgroup = CreateObject("dbWorkGroup.cdWorkGroup")
    Set l_adrsEmployees = l_objWorkgroup.GetEmployeesInDepartmentTimeCharged( _
                            l_sArrDepts _
                            , v_dtStartDate _
                            , v_dtEndDate _
                            , v_bActiveEmployees)
    Set l_objWorkgroup = Nothing

    If (RecordsExist(l_adrsEmployees)) Then
        With l_adrsEmployees
            .MoveFirst

            ReDim l_lARREmployees(.RecordCount - 1)

            For l_lArrayIndex = 1 To .RecordCount
                l_lARREmployees(l_lArrayIndex - 1) = !EMP_NUM

                .MoveNext
            Next l_lArrayIndex
        End With
    Else
        Set l_obj = Nothing
        Set l_objWorkgroup = Nothing
        Set l_adrsEmployees = Nothing
        Exit Function
    End If
    
    Set ladrsActualTimeHistoryDept = New ADODB.Recordset
    With ladrsActualTimeHistoryDept
        .Fields.Append "EMP_NUM", adInteger, , m_conFIELDATTR
        .Fields.Append "DEPT_NUM", adChar, 3, m_conFIELDATTR
        .Fields.Append "OT_OPP_NUM", adInteger, , m_conFIELDATTR
        .Fields.Append "CHARGE_DATE", adDate, , m_conFIELDATTR
        .Fields.Append "DATE_WORKED", adDate, , m_conFIELDATTR
        .Fields.Append "START_TIME", adChar, 5, m_conFIELDATTR
        .Fields.Append "END_TIME", adChar, 5, m_conFIELDATTR
        .Fields.Append "LOC_CD", adInteger, , m_conFIELDATTR
        .Fields.Append "JOB_NUM", adChar, 6, m_conFIELDATTR
        .Fields.Append "HOURS", adDecimal, , m_conFIELDATTR
        .Fields.Item("HOURS").Precision = 3
        .Fields.Item("HOURS").NumericScale = 1
        .Fields.Append "WG_NUM", adInteger, 5, m_conFIELDATTR
        .Fields.Append "REL_CD", adChar, 3, m_conFIELDATTR
        .Fields.Append "COMP_CD", adChar, 3, m_conFIELDATTR
        .Fields.Append "COMP_SHIFT_CD", adChar, 3, m_conFIELDATTR
        .Fields.Append "EXCEPTION_CD", adChar, 3, m_conFIELDATTR
        .Fields.Append "EARNINGS_CD", adChar, 3, m_conFIELDATTR
        .Fields.Append "OVER_EXC_CD", adChar, 3, m_conFIELDATTR
        .Fields.Append "CASE_NUM", adInteger, , m_conFIELDATTR
        .Fields.Append "START_LOC_CD", adInteger, , m_conFIELDATTR
        .Fields.Append "END_LOC_CD", adInteger, , m_conFIELDATTR
        .Fields.Append "AMOUNT", adDecimal, , m_conFIELDATTR
        .Fields.Item("AMOUNT").Precision = 7
        .Fields.Item("AMOUNT").NumericScale = 2
        .Fields.Append "APPR_BY_EMP_NUM", adInteger, , m_conFIELDATTR
        .Fields.Append "DISTRIBUTION_TYPE", adChar, 1, m_conFIELDATTR
        .Fields.Append "DISTRIBUTION_ID", adChar, 39, m_conFIELDATTR
        .Fields.Append "OT_REASON_CD", adChar, 3, m_conFIELDATTR
        .Fields.Append "ADJ_FLAG", adChar, 3, m_conFIELDATTR
        .Fields.Append "SEQ_NUM", adInteger, , m_conFIELDATTR
        .Fields.Append "ADJ_TYPE", adChar, 1, m_conFIELDATTR
        .Fields.Append "TRAVEL_PAY_IND", adChar, 1, m_conFIELDATTR
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open
    End With

    Set l_adrsEmployees = Nothing

    '//internally within DistMgr.GetDistributionList (due to OTTC) the description length must be greater than or equal to 39
    '//since distribution_id, description, and job_task will all be 60 chars
    '//but timecard will return distribution_id as 39 chars
    Set l_obj = CreateObject("busTimecard.cbsTimecard")
    Set ladrsActualTime = l_obj.GetActualTimeDetailWithOtherPay2B( _
                                    l_lARREmployees _
                                    , v_dtStartDate _
                                    , v_dtEndDate, 60)
    Set l_obj = Nothing
    
    If RecordsExist(ladrsActualTime) Then
        With ladrsActualTime
            .MoveFirst
            While Not ladrsActualTime.EOF
                ladrsActualTimeHistoryDept.AddNew
                ladrsActualTimeHistoryDept.Fields("EMP_NUM") = !EMP_NUM
                ladrsActualTimeHistoryDept.Fields("DEPT_NUM") = !DEPT_NUM
                ladrsActualTimeHistoryDept.Fields("OT_OPP_NUM") = !OT_OPP_NUM
                ladrsActualTimeHistoryDept.Fields("CHARGE_DATE") = !CHARGE_DATE
                ladrsActualTimeHistoryDept.Fields("DATE_WORKED") = !DATE_WORKED
                ladrsActualTimeHistoryDept.Fields("START_TIME") = !START_TIME
                ladrsActualTimeHistoryDept.Fields("END_TIME") = !END_TIME
                ladrsActualTimeHistoryDept.Fields("LOC_CD") = !LOC_CD
                ladrsActualTimeHistoryDept.Fields("JOB_NUM") = !JOB_NUM
                ladrsActualTimeHistoryDept.Fields("HOURS") = !Hours
                ladrsActualTimeHistoryDept.Fields("WG_NUM") = !WG_NUM
                ladrsActualTimeHistoryDept.Fields("REL_CD") = !REL_CD
                ladrsActualTimeHistoryDept.Fields("COMP_CD") = !COMP_CD
                ladrsActualTimeHistoryDept.Fields("COMP_SHIFT_CD") = !COMP_SHIFT_CD
                ladrsActualTimeHistoryDept.Fields("EXCEPTION_CD") = !EXCEPTION_CD
                ladrsActualTimeHistoryDept.Fields("EARNINGS_CD") = !EARNINGS_CD
                ladrsActualTimeHistoryDept.Fields("OVER_EXC_CD") = !OVER_EXC_CD
                ladrsActualTimeHistoryDept.Fields("CASE_NUM") = !CASE_NUM
                ladrsActualTimeHistoryDept.Fields("START_LOC_CD") = !START_LOC_CD
                ladrsActualTimeHistoryDept.Fields("END_LOC_CD") = !END_LOC_CD
                ladrsActualTimeHistoryDept.Fields("AMOUNT") = !AMOUNT
                ladrsActualTimeHistoryDept.Fields("APPR_BY_EMP_NUM") = !APPR_BY_EMP_NUM
                ladrsActualTimeHistoryDept.Fields("DISTRIBUTION_TYPE") = !DISTRIBUTION_TYPE
                ladrsActualTimeHistoryDept.Fields("DISTRIBUTION_ID") = !DISTRIBUTION_ID
                ladrsActualTimeHistoryDept.Fields("OT_REASON_CD") = !OT_REASON_CD
                ladrsActualTimeHistoryDept.Fields("ADJ_FLAG") = "ACT"
                ladrsActualTimeHistoryDept.Fields("SEQ_NUM") = 0
                ladrsActualTimeHistoryDept.Fields("ADJ_TYPE") = ""
                ladrsActualTimeHistoryDept.Fields("TRAVEL_PAY_IND") = !TRAVEL_PAY_IND
                ladrsActualTimeHistoryDept.Update
                .MoveNext
            Wend
        End With
    End If
            
    Set l_objAdj = CreateObject("dbAdjustments.cdAdjustments")
    Set ladrsAdj = l_objAdj.GetAdjustmentDetails(l_lARREmployees _
                                                 , v_dtStartDate _
                                                 , v_dtEndDate)
    Set l_objAdj = Nothing
    
    If RecordsExist(ladrsAdj) Then
        With ladrsAdj
            .MoveFirst
            While Not ladrsAdj.EOF
                ladrsActualTimeHistoryDept.AddNew
                ladrsActualTimeHistoryDept.Fields("EMP_NUM") = !EMP_NUM
                ladrsActualTimeHistoryDept.Fields("DEPT_NUM") = !DEPT_NUM
                ladrsActualTimeHistoryDept.Fields("OT_OPP_NUM") = !OT_OPP_NUM
                ladrsActualTimeHistoryDept.Fields("CHARGE_DATE") = !DATE_WORKED
                ladrsActualTimeHistoryDept.Fields("DATE_WORKED") = !DATE_WORKED
                ladrsActualTimeHistoryDept.Fields("START_TIME") = !START_TIME
                ladrsActualTimeHistoryDept.Fields("END_TIME") = !END_TIME
                ladrsActualTimeHistoryDept.Fields("LOC_CD") = !LOC_CD
                ladrsActualTimeHistoryDept.Fields("JOB_NUM") = !JOB_NUM
                ladrsActualTimeHistoryDept.Fields("HOURS") = !Hours
                ladrsActualTimeHistoryDept.Fields("WG_NUM") = !WG_NUM
                ladrsActualTimeHistoryDept.Fields("REL_CD") = !REL_CD
                ladrsActualTimeHistoryDept.Fields("COMP_CD") = !COMP_CD
                ladrsActualTimeHistoryDept.Fields("COMP_SHIFT_CD") = !SHIFT_COMP
                ladrsActualTimeHistoryDept.Fields("EXCEPTION_CD") = !EXCEPTION_CD
                ladrsActualTimeHistoryDept.Fields("EARNINGS_CD") = !EARNINGS_CD
                ladrsActualTimeHistoryDept.Fields("OVER_EXC_CD") = !OVER_EXC_CD
                ladrsActualTimeHistoryDept.Fields("CASE_NUM") = !CASE_NUM
                ladrsActualTimeHistoryDept.Fields("START_LOC_CD") = !START_LOC_CD
                ladrsActualTimeHistoryDept.Fields("END_LOC_CD") = !END_LOC_CD
                ladrsActualTimeHistoryDept.Fields("AMOUNT") = !AMOUNT
                ladrsActualTimeHistoryDept.Fields("APPR_BY_EMP_NUM") = !APPR_BY_MANAGER
                ladrsActualTimeHistoryDept.Fields("DISTRIBUTION_TYPE") = !DISTRIBUTION_TYPE
                ladrsActualTimeHistoryDept.Fields("DISTRIBUTION_ID") = !DISTRIBUTION_ID
                ladrsActualTimeHistoryDept.Fields("OT_REASON_CD") = !OT_REASON_CD
                ladrsActualTimeHistoryDept.Fields("ADJ_FLAG") = "ADJ"
                ladrsActualTimeHistoryDept.Fields("SEQ_NUM") = !SEQ_NUM
                ladrsActualTimeHistoryDept.Fields("ADJ_TYPE") = !ADJ_TYPE
                If !TRAVEL_PAY_HRS <> 0# Then
                    ladrsActualTimeHistoryDept.Fields("TRAVEL_PAY_IND") = "Y"
                Else
                    ladrsActualTimeHistoryDept.Fields("TRAVEL_PAY_IND") = "N"
                End If
                ladrsActualTimeHistoryDept.Update
                .MoveNext
            Wend
        End With
    End If
    
    Set GetActualTimeHistoryDept = ladrsActualTimeHistoryDept

    Exit Function
GetActualTimeHistoryDept_EH:
    Set l_objWorkgroup = Nothing
    Set l_obj = Nothing
    Set l_adrsEmployees = Nothing
    Set ladrsAdj = Nothing
    Set ladrsActualTime = Nothing
    Set ladrsActualTimeHistoryDept = Nothing

    TrsRaiseError MODULENAME _
        , "GetActualTimeHistoryDept"

End Function
Private Function GetActualTimeHistoryDeptbyDist( _
                ByVal v_adrsSelectedDepartments As ADODB.Recordset _
                , ByVal v_dtStartDate As Date _
                , ByVal v_dtEndDate As Date _
                , ByVal v_adrsDistributions As ADODB.Recordset _
                ) As ADODB.Recordset
On Error GoTo GetActualTimeHistoryDeptbyDist_EH
Dim l_obj As dbTimeCard.cdTimeCard
Dim l_objWorkgroup As dbWorkGroup.cdWorkGroup
Dim l_objAdj As dbAdjustments.cdAdjustments
Dim ladrsActualTimeHistoryDeptbyDist As ADODB.Recordset
Dim ladrsActualTime As ADODB.Recordset
Dim ladrsAdj As ADODB.Recordset

Dim l_adrsEmployees As New ADODB.Recordset
Dim l_lARREmployees() As Long, l_sArrDepts() As String
Dim l_lArrayIndex As Long

    Set GetActualTimeHistoryDeptbyDist = Nothing

    If Not (RecordsExist(v_adrsSelectedDepartments)) Then
        Exit Function
    Else
        With v_adrsSelectedDepartments
            .MoveFirst

            ReDim l_sArrDepts(.RecordCount - 1)
            For l_lArrayIndex = 1 To .RecordCount
                l_sArrDepts(l_lArrayIndex - 1) = !DEPT_NUM

                .MoveNext
            Next l_lArrayIndex
        End With
    End If

    Set l_objWorkgroup = CreateObject("dbWorkGroup.cdWorkGroup")
    Set l_adrsEmployees = l_objWorkgroup.GetEmployeesInDepartmentTimeCharged( _
                            l_sArrDepts _
                            , v_dtStartDate _
                            , v_dtEndDate _
                            , False)
                            
    Set l_objWorkgroup = Nothing

    If (RecordsExist(l_adrsEmployees)) Then
        With l_adrsEmployees
            .MoveFirst

            ReDim l_lARREmployees(.RecordCount - 1)

            For l_lArrayIndex = 1 To .RecordCount
                l_lARREmployees(l_lArrayIndex - 1) = !EMP_NUM

                .MoveNext
            Next l_lArrayIndex
        End With
    Else
        Set l_obj = Nothing
        Set l_objWorkgroup = Nothing
        Set l_adrsEmployees = Nothing
        Exit Function
    End If
    
    Set ladrsActualTimeHistoryDeptbyDist = New ADODB.Recordset
    With ladrsActualTimeHistoryDeptbyDist
        .Fields.Append "EMP_NUM", adInteger, , m_conFIELDATTR
        .Fields.Append "DEPT_NUM", adChar, 3, m_conFIELDATTR
        .Fields.Append "OT_OPP_NUM", adInteger, , m_conFIELDATTR
        .Fields.Append "CHARGE_DATE", adDate, , m_conFIELDATTR
        .Fields.Append "DATE_WORKED", adDate, , m_conFIELDATTR
        .Fields.Append "START_TIME", adChar, 5, m_conFIELDATTR
        .Fields.Append "END_TIME", adChar, 5, m_conFIELDATTR
        .Fields.Append "LOC_CD", adInteger, , m_conFIELDATTR
        .Fields.Append "JOB_NUM", adChar, 6, m_conFIELDATTR
        .Fields.Append "HOURS", adDecimal, , m_conFIELDATTR
        .Fields.Item("HOURS").Precision = 3
        .Fields.Item("HOURS").NumericScale = 1
        .Fields.Append "WG_NUM", adInteger, 5, m_conFIELDATTR
        .Fields.Append "REL_CD", adChar, 3, m_conFIELDATTR
        .Fields.Append "COMP_CD", adChar, 3, m_conFIELDATTR
        .Fields.Append "COMP_SHIFT_CD", adChar, 3, m_conFIELDATTR
        .Fields.Append "EXCEPTION_CD", adChar, 3, m_conFIELDATTR
        .Fields.Append "EARNINGS_CD", adChar, 3, m_conFIELDATTR
        .Fields.Append "OVER_EXC_CD", adChar, 3, m_conFIELDATTR
        .Fields.Append "CASE_NUM", adInteger, , m_conFIELDATTR
        .Fields.Append "START_LOC_CD", adInteger, , m_conFIELDATTR
        .Fields.Append "END_LOC_CD", adInteger, , m_conFIELDATTR
        .Fields.Append "AMOUNT", adDecimal, , m_conFIELDATTR
        .Fields.Item("AMOUNT").Precision = 7
        .Fields.Item("AMOUNT").NumericScale = 2
        .Fields.Append "APPR_BY_EMP_NUM", adInteger, , m_conFIELDATTR
        .Fields.Append "DISTRIBUTION_TYPE", adChar, 1, m_conFIELDATTR
        .Fields.Append "DISTRIBUTION_ID", adChar, 39, m_conFIELDATTR
        .Fields.Append "OT_REASON_CD", adChar, 3, m_conFIELDATTR
        .Fields.Append "ADJ_FLAG", adChar, 3, m_conFIELDATTR
        .Fields.Append "SEQ_NUM", adInteger, , m_conFIELDATTR
        .Fields.Append "ADJ_TYPE", adChar, 1, m_conFIELDATTR
        .Fields.Append "TRAVEL_PAY_IND", adChar, 1, m_conFIELDATTR
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open
    End With

    Set l_adrsEmployees = Nothing

    Set l_obj = CreateObject("dbTimecard.cdTimecard")
    Set ladrsActualTime = l_obj.GetDistributionActualTimeEntries2( _
                                    v_adrsDistributions _
                                    , v_dtStartDate _
                                    , v_dtEndDate _
                                    , l_lARREmployees)
    Set l_obj = Nothing
    
    If RecordsExist(ladrsActualTime) Then
        With ladrsActualTime
            .MoveFirst
            While Not ladrsActualTime.EOF
                ladrsActualTimeHistoryDeptbyDist.AddNew
                ladrsActualTimeHistoryDeptbyDist.Fields("EMP_NUM") = !EMP_NUM
                ladrsActualTimeHistoryDeptbyDist.Fields("DEPT_NUM") = !DEPT_NUM
                ladrsActualTimeHistoryDeptbyDist.Fields("OT_OPP_NUM") = !OT_OPP_NUM
                ladrsActualTimeHistoryDeptbyDist.Fields("CHARGE_DATE") = !CHARGE_DATE
                ladrsActualTimeHistoryDeptbyDist.Fields("DATE_WORKED") = !DATE_WORKED
                ladrsActualTimeHistoryDeptbyDist.Fields("START_TIME") = !START_TIME
                ladrsActualTimeHistoryDeptbyDist.Fields("END_TIME") = !END_TIME
                ladrsActualTimeHistoryDeptbyDist.Fields("LOC_CD") = !LOC_CD
                ladrsActualTimeHistoryDeptbyDist.Fields("JOB_NUM") = !JOB_NUM
                ladrsActualTimeHistoryDeptbyDist.Fields("HOURS") = !Hours
                ladrsActualTimeHistoryDeptbyDist.Fields("WG_NUM") = !WG_NUM
                ladrsActualTimeHistoryDeptbyDist.Fields("REL_CD") = !REL_CD
                ladrsActualTimeHistoryDeptbyDist.Fields("COMP_CD") = !COMP_CD
                ladrsActualTimeHistoryDeptbyDist.Fields("COMP_SHIFT_CD") = !COMP_SHIFT_CD
                ladrsActualTimeHistoryDeptbyDist.Fields("EXCEPTION_CD") = !EXCEPTION_CD
                ladrsActualTimeHistoryDeptbyDist.Fields("EARNINGS_CD") = !EARNINGS_CD
                ladrsActualTimeHistoryDeptbyDist.Fields("OVER_EXC_CD") = !OVER_EXC_CD
                ladrsActualTimeHistoryDeptbyDist.Fields("CASE_NUM") = !CASE_NUM
                ladrsActualTimeHistoryDeptbyDist.Fields("START_LOC_CD") = !START_LOC_CD
                ladrsActualTimeHistoryDeptbyDist.Fields("END_LOC_CD") = !END_LOC_CD
                ladrsActualTimeHistoryDeptbyDist.Fields("AMOUNT") = !AMOUNT
                ladrsActualTimeHistoryDeptbyDist.Fields("APPR_BY_EMP_NUM") = !APPR_BY_EMP_NUM
                ladrsActualTimeHistoryDeptbyDist.Fields("DISTRIBUTION_TYPE") = !DISTRIBUTION_TYPE
                ladrsActualTimeHistoryDeptbyDist.Fields("DISTRIBUTION_ID") = !DISTRIBUTION_ID
                ladrsActualTimeHistoryDeptbyDist.Fields("OT_REASON_CD") = !OT_REASON_CD
                ladrsActualTimeHistoryDeptbyDist.Fields("ADJ_FLAG") = "ACT"
                ladrsActualTimeHistoryDeptbyDist.Fields("SEQ_NUM") = 0
                ladrsActualTimeHistoryDeptbyDist.Fields("ADJ_TYPE") = ""
                ladrsActualTimeHistoryDeptbyDist.Fields("TRAVEL_PAY_IND") = !TRAVEL_PAY_IND
                ladrsActualTimeHistoryDeptbyDist.Update
                .MoveNext
            Wend
        End With
    End If
            
    Set l_objAdj = CreateObject("dbAdjustments.cdAdjustments")
    Set ladrsAdj = l_objAdj.GetDistributionAdjustmentDetails(l_lARREmployees _
                                                 , v_dtStartDate _
                                                 , v_dtEndDate _
                                                 , v_adrsDistributions)
    Set l_objAdj = Nothing
    
    If RecordsExist(ladrsAdj) Then
        With ladrsAdj
            .MoveFirst
            While Not ladrsAdj.EOF
                ladrsActualTimeHistoryDeptbyDist.AddNew
                ladrsActualTimeHistoryDeptbyDist.Fields("EMP_NUM") = !EMP_NUM
                ladrsActualTimeHistoryDeptbyDist.Fields("DEPT_NUM") = !DEPT_NUM
                ladrsActualTimeHistoryDeptbyDist.Fields("OT_OPP_NUM") = !OT_OPP_NUM
                ladrsActualTimeHistoryDeptbyDist.Fields("CHARGE_DATE") = !DATE_WORKED
                ladrsActualTimeHistoryDeptbyDist.Fields("DATE_WORKED") = !DATE_WORKED
                ladrsActualTimeHistoryDeptbyDist.Fields("START_TIME") = !START_TIME
                ladrsActualTimeHistoryDeptbyDist.Fields("END_TIME") = !END_TIME
                ladrsActualTimeHistoryDeptbyDist.Fields("LOC_CD") = !LOC_CD
                ladrsActualTimeHistoryDeptbyDist.Fields("JOB_NUM") = !JOB_NUM
                ladrsActualTimeHistoryDeptbyDist.Fields("HOURS") = !Hours
                ladrsActualTimeHistoryDeptbyDist.Fields("WG_NUM") = !WG_NUM
                ladrsActualTimeHistoryDeptbyDist.Fields("REL_CD") = !REL_CD
                ladrsActualTimeHistoryDeptbyDist.Fields("COMP_CD") = !COMP_CD
                ladrsActualTimeHistoryDeptbyDist.Fields("COMP_SHIFT_CD") = !SHIFT_COMP
                ladrsActualTimeHistoryDeptbyDist.Fields("EXCEPTION_CD") = !EXCEPTION_CD
                ladrsActualTimeHistoryDeptbyDist.Fields("EARNINGS_CD") = !EARNINGS_CD
                ladrsActualTimeHistoryDeptbyDist.Fields("OVER_EXC_CD") = !OVER_EXC_CD
                ladrsActualTimeHistoryDeptbyDist.Fields("CASE_NUM") = !CASE_NUM
                ladrsActualTimeHistoryDeptbyDist.Fields("START_LOC_CD") = !START_LOC_CD
                ladrsActualTimeHistoryDeptbyDist.Fields("END_LOC_CD") = !END_LOC_CD
                ladrsActualTimeHistoryDeptbyDist.Fields("AMOUNT") = !AMOUNT
                ladrsActualTimeHistoryDeptbyDist.Fields("APPR_BY_EMP_NUM") = !APPR_BY_MANAGER
                ladrsActualTimeHistoryDeptbyDist.Fields("DISTRIBUTION_TYPE") = !DISTRIBUTION_TYPE
                ladrsActualTimeHistoryDeptbyDist.Fields("DISTRIBUTION_ID") = !DISTRIBUTION_ID
                ladrsActualTimeHistoryDeptbyDist.Fields("OT_REASON_CD") = !OT_REASON_CD
                ladrsActualTimeHistoryDeptbyDist.Fields("ADJ_FLAG") = "ADJ"
                ladrsActualTimeHistoryDeptbyDist.Fields("SEQ_NUM") = !SEQ_NUM
                ladrsActualTimeHistoryDeptbyDist.Fields("ADJ_TYPE") = !ADJ_TYPE
                If !TRAVEL_PAY_HRS <> 0# Then
                    ladrsActualTimeHistoryDeptbyDist.Fields("TRAVEL_PAY_IND") = "Y"
                Else
                    ladrsActualTimeHistoryDeptbyDist.Fields("TRAVEL_PAY_IND") = "N"
                End If
                ladrsActualTimeHistoryDeptbyDist.Update
                .MoveNext
            Wend
        End With
    End If
    
    Set GetActualTimeHistoryDeptbyDist = ladrsActualTimeHistoryDeptbyDist

    Exit Function
GetActualTimeHistoryDeptbyDist_EH:
    Set l_objWorkgroup = Nothing
    Set l_obj = Nothing
    Set l_adrsEmployees = Nothing
    Set ladrsAdj = Nothing
    Set ladrsActualTime = Nothing
    Set ladrsActualTimeHistoryDeptbyDist = Nothing

    TrsRaiseError MODULENAME _
        , "GetActualTimeHistoryDeptbyDist"

End Function
Private Function GetActualTimeHistoryWG( _
                ByVal v_adrsSelectedWorkgroups As ADODB.Recordset _
                , ByVal v_dtStartDate As Date _
                , ByVal v_dtEndDate As Date _
                , ByVal v_bActiveEmployees As Boolean _
                ) As ADODB.Recordset
On Error GoTo GetActualTimeHistoryWG_EH

Dim l_obj As busTimecard.cbsTimecard
Dim l_objWorkgroup As dbWorkGroup.cdWorkGroup
Dim l_objAdj As dbAdjustments.cdAdjustments
Dim l_objEmployee As dbEmployee.cdEmployee
Dim ladrsActualTime As ADODB.Recordset
Dim ladrsAdj As ADODB.Recordset
Dim ladrsActualTimeHistoryWG As ADODB.Recordset
Dim ladrsPrimaryWG As ADODB.Recordset

Dim l_lARREmployees() As Long, l_lARRWorkgroups() As Long
Dim llarrEmp(0) As Long
Dim l_adrsEmployees As ADODB.Recordset
Dim l_lArrayIndex As Long

    Set GetActualTimeHistoryWG = Nothing

    If Not (RecordsExist(v_adrsSelectedWorkgroups)) Then
            Exit Function
    Else
        With v_adrsSelectedWorkgroups
            .MoveFirst

            ReDim l_lARRWorkgroups(.RecordCount - 1)

            For l_lArrayIndex = 1 To .RecordCount
                l_lARRWorkgroups(l_lArrayIndex - 1) = !WG_NUM

                .MoveNext
            Next l_lArrayIndex
        End With
    End If

    Set l_objWorkgroup = CreateObject("dbWorkGroup.cdWorkGroup")
    Set l_adrsEmployees = l_objWorkgroup.GetEmployeesInWorkGroupViewTime( _
                            l_lARRWorkgroups _
                            , v_dtStartDate _
                            , v_dtEndDate _
                            , v_bActiveEmployees)
    Set l_objWorkgroup = Nothing

    If (RecordsExist(l_adrsEmployees)) Then
        With l_adrsEmployees
            .MoveFirst

            ReDim l_lARREmployees(.RecordCount - 1)

            For l_lArrayIndex = 1 To .RecordCount
                l_lARREmployees(l_lArrayIndex - 1) = !EMP_NUM

                .MoveNext
            Next l_lArrayIndex
        End With
    Else
        Set l_objWorkgroup = Nothing
        Set l_obj = Nothing
        Set l_adrsEmployees = Nothing

        Exit Function
    End If
    
    Set ladrsActualTimeHistoryWG = New ADODB.Recordset
    With ladrsActualTimeHistoryWG
        .Fields.Append "EMP_NUM", adInteger, , m_conFIELDATTR
        .Fields.Append "DEPT_NUM", adChar, 3, m_conFIELDATTR
        .Fields.Append "OT_OPP_NUM", adInteger, , m_conFIELDATTR
        .Fields.Append "CHARGE_DATE", adDate, , m_conFIELDATTR
        .Fields.Append "DATE_WORKED", adDate, , m_conFIELDATTR
        .Fields.Append "START_TIME", adChar, 5, m_conFIELDATTR
        .Fields.Append "END_TIME", adChar, 5, m_conFIELDATTR
        .Fields.Append "LOC_CD", adInteger, , m_conFIELDATTR
        .Fields.Append "JOB_NUM", adChar, 6, m_conFIELDATTR
        .Fields.Append "HOURS", adDecimal, , m_conFIELDATTR
        .Fields.Item("HOURS").Precision = 3
        .Fields.Item("HOURS").NumericScale = 1
        .Fields.Append "WG_NUM", adInteger, 5, m_conFIELDATTR
        .Fields.Append "REL_CD", adChar, 3, m_conFIELDATTR
        .Fields.Append "COMP_CD", adChar, 3, m_conFIELDATTR
        .Fields.Append "COMP_SHIFT_CD", adChar, 3, m_conFIELDATTR
        .Fields.Append "EXCEPTION_CD", adChar, 3, m_conFIELDATTR
        .Fields.Append "EARNINGS_CD", adChar, 3, m_conFIELDATTR
        .Fields.Append "OVER_EXC_CD", adChar, 3, m_conFIELDATTR
        .Fields.Append "CASE_NUM", adInteger, , m_conFIELDATTR
        .Fields.Append "START_LOC_CD", adInteger, , m_conFIELDATTR
        .Fields.Append "END_LOC_CD", adInteger, , m_conFIELDATTR
        .Fields.Append "AMOUNT", adDecimal, , m_conFIELDATTR
        .Fields.Item("AMOUNT").Precision = 7
        .Fields.Item("AMOUNT").NumericScale = 2
        .Fields.Append "APPR_BY_EMP_NUM", adInteger, , m_conFIELDATTR
        .Fields.Append "DISTRIBUTION_TYPE", adChar, 1, m_conFIELDATTR
        .Fields.Append "DISTRIBUTION_ID", adChar, 39, m_conFIELDATTR
        .Fields.Append "OT_REASON_CD", adChar, 3, m_conFIELDATTR
        .Fields.Append "ADJ_FLAG", adChar, 3, m_conFIELDATTR
        .Fields.Append "SEQ_NUM", adInteger, , m_conFIELDATTR
        .Fields.Append "ADJ_TYPE", adChar, 1, m_conFIELDATTR
        .Fields.Append "TRAVEL_PAY_IND", adChar, 1, m_conFIELDATTR
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open
    End With

    Set l_adrsEmployees = Nothing

    '//internally within DistMgr.GetDistributionList (due to OTTC) the description length must be greater than or equal to 39
    '//since distribution_id, description, and job_task will all be 60 chars
    '//but timecard will return distribution_id as 39 chars
    Set l_obj = CreateObject("busTimecard.cbsTimecard")
    Set ladrsActualTime = l_obj.GetActualTimeDetailWithOtherPay2B( _
                                        l_lARREmployees _
                                        , v_dtStartDate _
                                        , v_dtEndDate, 60)
    Set l_obj = Nothing
    
    If RecordsExist(ladrsActualTime) Then
        With ladrsActualTime
            .MoveFirst
            While Not ladrsActualTime.EOF
                ladrsActualTimeHistoryWG.AddNew
                ladrsActualTimeHistoryWG.Fields("EMP_NUM") = !EMP_NUM
                ladrsActualTimeHistoryWG.Fields("DEPT_NUM") = !DEPT_NUM
                ladrsActualTimeHistoryWG.Fields("OT_OPP_NUM") = !OT_OPP_NUM
                ladrsActualTimeHistoryWG.Fields("CHARGE_DATE") = !CHARGE_DATE
                ladrsActualTimeHistoryWG.Fields("DATE_WORKED") = !DATE_WORKED
                ladrsActualTimeHistoryWG.Fields("START_TIME") = !START_TIME
                ladrsActualTimeHistoryWG.Fields("END_TIME") = !END_TIME
                ladrsActualTimeHistoryWG.Fields("LOC_CD") = !LOC_CD
                ladrsActualTimeHistoryWG.Fields("JOB_NUM") = !JOB_NUM
                ladrsActualTimeHistoryWG.Fields("HOURS") = !Hours
                ladrsActualTimeHistoryWG.Fields("WG_NUM") = !WG_NUM
                ladrsActualTimeHistoryWG.Fields("REL_CD") = !REL_CD
                ladrsActualTimeHistoryWG.Fields("COMP_CD") = !COMP_CD
                ladrsActualTimeHistoryWG.Fields("COMP_SHIFT_CD") = !COMP_SHIFT_CD
                ladrsActualTimeHistoryWG.Fields("EXCEPTION_CD") = !EXCEPTION_CD
                ladrsActualTimeHistoryWG.Fields("EARNINGS_CD") = !EARNINGS_CD
                ladrsActualTimeHistoryWG.Fields("OVER_EXC_CD") = !OVER_EXC_CD
                ladrsActualTimeHistoryWG.Fields("CASE_NUM") = !CASE_NUM
                ladrsActualTimeHistoryWG.Fields("START_LOC_CD") = !START_LOC_CD
                ladrsActualTimeHistoryWG.Fields("END_LOC_CD") = !END_LOC_CD
                ladrsActualTimeHistoryWG.Fields("AMOUNT") = !AMOUNT
                ladrsActualTimeHistoryWG.Fields("APPR_BY_EMP_NUM") = !APPR_BY_EMP_NUM
                ladrsActualTimeHistoryWG.Fields("DISTRIBUTION_TYPE") = !DISTRIBUTION_TYPE
                ladrsActualTimeHistoryWG.Fields("DISTRIBUTION_ID") = !DISTRIBUTION_ID
                ladrsActualTimeHistoryWG.Fields("OT_REASON_CD") = !OT_REASON_CD
                ladrsActualTimeHistoryWG.Fields("ADJ_FLAG") = "ACT"
                ladrsActualTimeHistoryWG.Fields("SEQ_NUM") = 0
                ladrsActualTimeHistoryWG.Fields("ADJ_TYPE") = ""
                ladrsActualTimeHistoryWG.Fields("TRAVEL_PAY_IND") = !TRAVEL_PAY_IND
                ladrsActualTimeHistoryWG.Update
                .MoveNext
            Wend
        End With
    End If
            
    Set l_objAdj = CreateObject("dbAdjustments.cdAdjustments")
    Set ladrsAdj = l_objAdj.GetAdjustmentDetails(l_lARREmployees _
                                                 , v_dtStartDate _
                                                 , v_dtEndDate)
    Set l_objAdj = Nothing
    
    If RecordsExist(ladrsAdj) Then
        With ladrsAdj
            .MoveFirst
            While Not ladrsAdj.EOF
                ladrsActualTimeHistoryWG.AddNew
                ladrsActualTimeHistoryWG.Fields("EMP_NUM") = !EMP_NUM
                ladrsActualTimeHistoryWG.Fields("DEPT_NUM") = !DEPT_NUM
                ladrsActualTimeHistoryWG.Fields("OT_OPP_NUM") = !OT_OPP_NUM
                ladrsActualTimeHistoryWG.Fields("CHARGE_DATE") = !DATE_WORKED
                ladrsActualTimeHistoryWG.Fields("DATE_WORKED") = !DATE_WORKED
                ladrsActualTimeHistoryWG.Fields("START_TIME") = !START_TIME
                ladrsActualTimeHistoryWG.Fields("END_TIME") = !END_TIME
                ladrsActualTimeHistoryWG.Fields("LOC_CD") = !LOC_CD
                ladrsActualTimeHistoryWG.Fields("JOB_NUM") = !JOB_NUM
                ladrsActualTimeHistoryWG.Fields("HOURS") = !Hours
                Set l_objEmployee = CreateObject("dbEmployee.cdEmployee")
                llarrEmp(0) = !EMP_NUM
                Set ladrsPrimaryWG = l_objEmployee.GetEmployeePrimaryWorkGroup(llarrEmp(0))
                Set l_objEmployee = Nothing
                If RecordsExist(ladrsPrimaryWG) Then
                    ladrsActualTimeHistoryWG.Fields("WG_NUM") = ladrsPrimaryWG.Fields("WG_NUM")
                Else
                    ladrsActualTimeHistoryWG.Fields("WG_NUM") = 0
                End If
                ladrsActualTimeHistoryWG.Fields("REL_CD") = !REL_CD
                ladrsActualTimeHistoryWG.Fields("COMP_CD") = !COMP_CD
                ladrsActualTimeHistoryWG.Fields("COMP_SHIFT_CD") = !SHIFT_COMP
                ladrsActualTimeHistoryWG.Fields("EXCEPTION_CD") = !EXCEPTION_CD
                ladrsActualTimeHistoryWG.Fields("EARNINGS_CD") = !EARNINGS_CD
                ladrsActualTimeHistoryWG.Fields("OVER_EXC_CD") = !OVER_EXC_CD
                ladrsActualTimeHistoryWG.Fields("CASE_NUM") = !CASE_NUM
                ladrsActualTimeHistoryWG.Fields("START_LOC_CD") = !START_LOC_CD
                ladrsActualTimeHistoryWG.Fields("END_LOC_CD") = !END_LOC_CD
                ladrsActualTimeHistoryWG.Fields("AMOUNT") = !AMOUNT
                ladrsActualTimeHistoryWG.Fields("APPR_BY_EMP_NUM") = !APPR_BY_MANAGER
                ladrsActualTimeHistoryWG.Fields("DISTRIBUTION_TYPE") = !DISTRIBUTION_TYPE
                ladrsActualTimeHistoryWG.Fields("DISTRIBUTION_ID") = !DISTRIBUTION_ID
                ladrsActualTimeHistoryWG.Fields("OT_REASON_CD") = !OT_REASON_CD
                ladrsActualTimeHistoryWG.Fields("ADJ_FLAG") = "ADJ"
                ladrsActualTimeHistoryWG.Fields("SEQ_NUM") = !SEQ_NUM
                ladrsActualTimeHistoryWG.Fields("ADJ_TYPE") = !ADJ_TYPE
                If !TRAVEL_PAY_HRS <> 0# Then
                    ladrsActualTimeHistoryWG.Fields("TRAVEL_PAY_IND") = "Y"
                Else
                    ladrsActualTimeHistoryWG.Fields("TRAVEL_PAY_IND") = "N"
                End If
                ladrsActualTimeHistoryWG.Update
                .MoveNext
            Wend
        End With
    End If
    
    Set GetActualTimeHistoryWG = ladrsActualTimeHistoryWG


    Exit Function
GetActualTimeHistoryWG_EH:
    Set l_objWorkgroup = Nothing
    Set l_obj = Nothing
    Set l_adrsEmployees = Nothing
    Set l_objAdj = Nothing
    Set ladrsActualTime = Nothing
    Set ladrsAdj = Nothing
    Set ladrsActualTimeHistoryWG = Nothing
    Set l_objEmployee = Nothing
    Set ladrsPrimaryWG = Nothing

    TrsRaiseError MODULENAME _
        , "GetActualTimeHistoryWG"

End Function
Private Function GetActualTimeHistoryWGbyDist( _
                ByVal v_adrsSelectedWorkgroups As ADODB.Recordset _
                , ByVal v_dtStartDate As Date _
                , ByVal v_dtEndDate As Date _
                , ByVal v_adrsDistributions As ADODB.Recordset _
                ) As ADODB.Recordset
On Error GoTo GetActualTimeHistoryWGbyDist_EH

Dim l_obj As dbTimeCard.cdTimeCard
Dim l_objWorkgroup As dbWorkGroup.cdWorkGroup
Dim l_objAdj As dbAdjustments.cdAdjustments
Dim l_objEmployee As dbEmployee.cdEmployee
Dim ladrsActualTime As ADODB.Recordset
Dim ladrsAdj As ADODB.Recordset
Dim ladrsActualTimeHistoryWGbyDist As ADODB.Recordset
Dim ladrsPrimaryWG As ADODB.Recordset

Dim l_lARREmployees() As Long, l_lARRWorkgroups() As Long
Dim llarrEmp(0) As Long
Dim l_adrsEmployees As ADODB.Recordset
Dim l_lArrayIndex As Long

    Set GetActualTimeHistoryWGbyDist = Nothing

    If Not (RecordsExist(v_adrsSelectedWorkgroups)) Then
            Exit Function
    Else
        With v_adrsSelectedWorkgroups
            .MoveFirst

            ReDim l_lARRWorkgroups(.RecordCount - 1)

            For l_lArrayIndex = 1 To .RecordCount
                l_lARRWorkgroups(l_lArrayIndex - 1) = !WG_NUM

                .MoveNext
            Next l_lArrayIndex
        End With
    End If

    Set l_objWorkgroup = CreateObject("dbWorkGroup.cdWorkGroup")
    Set l_adrsEmployees = l_objWorkgroup.GetEmployeesInWorkGroupViewTime( _
                            l_lARRWorkgroups _
                            , v_dtStartDate _
                            , v_dtEndDate _
                            , False)
    Set l_objWorkgroup = Nothing

    If (RecordsExist(l_adrsEmployees)) Then
        With l_adrsEmployees
            .MoveFirst

            ReDim l_lARREmployees(.RecordCount - 1)

            For l_lArrayIndex = 1 To .RecordCount
                l_lARREmployees(l_lArrayIndex - 1) = !EMP_NUM

                .MoveNext
            Next l_lArrayIndex
        End With
    Else
        Set l_objWorkgroup = Nothing
        Set l_obj = Nothing
        Set l_adrsEmployees = Nothing

        Exit Function
    End If
    
    Set ladrsActualTimeHistoryWGbyDist = New ADODB.Recordset
    With ladrsActualTimeHistoryWGbyDist
        .Fields.Append "EMP_NUM", adInteger, , m_conFIELDATTR
        .Fields.Append "DEPT_NUM", adChar, 3, m_conFIELDATTR
        .Fields.Append "OT_OPP_NUM", adInteger, , m_conFIELDATTR
        .Fields.Append "CHARGE_DATE", adDate, , m_conFIELDATTR
        .Fields.Append "DATE_WORKED", adDate, , m_conFIELDATTR
        .Fields.Append "START_TIME", adChar, 5, m_conFIELDATTR
        .Fields.Append "END_TIME", adChar, 5, m_conFIELDATTR
        .Fields.Append "LOC_CD", adInteger, , m_conFIELDATTR
        .Fields.Append "JOB_NUM", adChar, 6, m_conFIELDATTR
        .Fields.Append "HOURS", adDecimal, , m_conFIELDATTR
        .Fields.Item("HOURS").Precision = 3
        .Fields.Item("HOURS").NumericScale = 1
        .Fields.Append "WG_NUM", adInteger, 5, m_conFIELDATTR
        .Fields.Append "REL_CD", adChar, 3, m_conFIELDATTR
        .Fields.Append "COMP_CD", adChar, 3, m_conFIELDATTR
        .Fields.Append "COMP_SHIFT_CD", adChar, 3, m_conFIELDATTR
        .Fields.Append "EXCEPTION_CD", adChar, 3, m_conFIELDATTR
        .Fields.Append "EARNINGS_CD", adChar, 3, m_conFIELDATTR
        .Fields.Append "OVER_EXC_CD", adChar, 3, m_conFIELDATTR
        .Fields.Append "CASE_NUM", adInteger, , m_conFIELDATTR
        .Fields.Append "START_LOC_CD", adInteger, , m_conFIELDATTR
        .Fields.Append "END_LOC_CD", adInteger, , m_conFIELDATTR
        .Fields.Append "AMOUNT", adDecimal, , m_conFIELDATTR
        .Fields.Item("AMOUNT").Precision = 7
        .Fields.Item("AMOUNT").NumericScale = 2
        .Fields.Append "APPR_BY_EMP_NUM", adInteger, , m_conFIELDATTR
        .Fields.Append "DISTRIBUTION_TYPE", adChar, 1, m_conFIELDATTR
        .Fields.Append "DISTRIBUTION_ID", adChar, 39, m_conFIELDATTR
        .Fields.Append "OT_REASON_CD", adChar, 3, m_conFIELDATTR
        .Fields.Append "ADJ_FLAG", adChar, 3, m_conFIELDATTR
        .Fields.Append "SEQ_NUM", adInteger, , m_conFIELDATTR
        .Fields.Append "ADJ_TYPE", adChar, 1, m_conFIELDATTR
        .Fields.Append "TRAVEL_PAY_IND", adChar, 1, m_conFIELDATTR
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open
    End With

    Set l_adrsEmployees = Nothing

    Set l_obj = CreateObject("dbTimecard.cdTimecard")
    Set ladrsActualTime = l_obj.GetDistributionActualTimeEntries2( _
                                        v_adrsDistributions _
                                        , v_dtStartDate _
                                        , v_dtEndDate _
                                        , l_lARREmployees)
    Set l_obj = Nothing
    
    If RecordsExist(ladrsActualTime) Then
        With ladrsActualTime
            .MoveFirst
            While Not ladrsActualTime.EOF
                ladrsActualTimeHistoryWGbyDist.AddNew
                ladrsActualTimeHistoryWGbyDist.Fields("EMP_NUM") = !EMP_NUM
                ladrsActualTimeHistoryWGbyDist.Fields("DEPT_NUM") = !DEPT_NUM
                ladrsActualTimeHistoryWGbyDist.Fields("OT_OPP_NUM") = !OT_OPP_NUM
                ladrsActualTimeHistoryWGbyDist.Fields("CHARGE_DATE") = !CHARGE_DATE
                ladrsActualTimeHistoryWGbyDist.Fields("DATE_WORKED") = !DATE_WORKED
                ladrsActualTimeHistoryWGbyDist.Fields("START_TIME") = !START_TIME
                ladrsActualTimeHistoryWGbyDist.Fields("END_TIME") = !END_TIME
                ladrsActualTimeHistoryWGbyDist.Fields("LOC_CD") = !LOC_CD
                ladrsActualTimeHistoryWGbyDist.Fields("JOB_NUM") = !JOB_NUM
                ladrsActualTimeHistoryWGbyDist.Fields("HOURS") = !Hours
                ladrsActualTimeHistoryWGbyDist.Fields("WG_NUM") = !WG_NUM
                ladrsActualTimeHistoryWGbyDist.Fields("REL_CD") = !REL_CD
                ladrsActualTimeHistoryWGbyDist.Fields("COMP_CD") = !COMP_CD
                ladrsActualTimeHistoryWGbyDist.Fields("COMP_SHIFT_CD") = !COMP_SHIFT_CD
                ladrsActualTimeHistoryWGbyDist.Fields("EXCEPTION_CD") = !EXCEPTION_CD
                ladrsActualTimeHistoryWGbyDist.Fields("EARNINGS_CD") = !EARNINGS_CD
                ladrsActualTimeHistoryWGbyDist.Fields("OVER_EXC_CD") = !OVER_EXC_CD
                ladrsActualTimeHistoryWGbyDist.Fields("CASE_NUM") = !CASE_NUM
                ladrsActualTimeHistoryWGbyDist.Fields("START_LOC_CD") = !START_LOC_CD
                ladrsActualTimeHistoryWGbyDist.Fields("END_LOC_CD") = !END_LOC_CD
                ladrsActualTimeHistoryWGbyDist.Fields("AMOUNT") = !AMOUNT
                ladrsActualTimeHistoryWGbyDist.Fields("APPR_BY_EMP_NUM") = !APPR_BY_EMP_NUM
                ladrsActualTimeHistoryWGbyDist.Fields("DISTRIBUTION_TYPE") = !DISTRIBUTION_TYPE
                ladrsActualTimeHistoryWGbyDist.Fields("DISTRIBUTION_ID") = !DISTRIBUTION_ID
                ladrsActualTimeHistoryWGbyDist.Fields("OT_REASON_CD") = !OT_REASON_CD
                ladrsActualTimeHistoryWGbyDist.Fields("ADJ_FLAG") = "ACT"
                ladrsActualTimeHistoryWGbyDist.Fields("SEQ_NUM") = 0
                ladrsActualTimeHistoryWGbyDist.Fields("ADJ_TYPE") = ""
                ladrsActualTimeHistoryWGbyDist.Fields("TRAVEL_PAY_IND") = !TRAVEL_PAY_IND
                ladrsActualTimeHistoryWGbyDist.Update
                .MoveNext
            Wend
        End With
    End If
            
    Set l_objAdj = CreateObject("dbAdjustments.cdAdjustments")
    Set ladrsAdj = l_objAdj.GetDistributionAdjustmentDetails(l_lARREmployees _
                                                           , v_dtStartDate _
                                                           , v_dtEndDate _
                                                           , v_adrsDistributions)
    Set l_objAdj = Nothing
    
    If RecordsExist(ladrsAdj) Then
        With ladrsAdj
            .MoveFirst
            While Not ladrsAdj.EOF
                ladrsActualTimeHistoryWGbyDist.AddNew
                ladrsActualTimeHistoryWGbyDist.Fields("EMP_NUM") = !EMP_NUM
                ladrsActualTimeHistoryWGbyDist.Fields("DEPT_NUM") = !DEPT_NUM
                ladrsActualTimeHistoryWGbyDist.Fields("OT_OPP_NUM") = !OT_OPP_NUM
                ladrsActualTimeHistoryWGbyDist.Fields("CHARGE_DATE") = !DATE_WORKED
                ladrsActualTimeHistoryWGbyDist.Fields("DATE_WORKED") = !DATE_WORKED
                ladrsActualTimeHistoryWGbyDist.Fields("START_TIME") = !START_TIME
                ladrsActualTimeHistoryWGbyDist.Fields("END_TIME") = !END_TIME
                ladrsActualTimeHistoryWGbyDist.Fields("LOC_CD") = !LOC_CD
                ladrsActualTimeHistoryWGbyDist.Fields("JOB_NUM") = !JOB_NUM
                ladrsActualTimeHistoryWGbyDist.Fields("HOURS") = !Hours
                Set l_objEmployee = CreateObject("dbEmployee.cdEmployee")
                llarrEmp(0) = !EMP_NUM
                Set ladrsPrimaryWG = l_objEmployee.GetEmployeePrimaryWorkGroup(llarrEmp(0))
                Set l_objEmployee = Nothing
                If RecordsExist(ladrsPrimaryWG) Then
                    ladrsActualTimeHistoryWGbyDist.Fields("WG_NUM") = ladrsPrimaryWG.Fields("WG_NUM")
                Else
                    ladrsActualTimeHistoryWGbyDist.Fields("WG_NUM") = 0
                End If
                ladrsActualTimeHistoryWGbyDist.Fields("REL_CD") = !REL_CD
                ladrsActualTimeHistoryWGbyDist.Fields("COMP_CD") = !COMP_CD
                ladrsActualTimeHistoryWGbyDist.Fields("COMP_SHIFT_CD") = !SHIFT_COMP
                ladrsActualTimeHistoryWGbyDist.Fields("EXCEPTION_CD") = !EXCEPTION_CD
                ladrsActualTimeHistoryWGbyDist.Fields("EARNINGS_CD") = !EARNINGS_CD
                ladrsActualTimeHistoryWGbyDist.Fields("OVER_EXC_CD") = !OVER_EXC_CD
                ladrsActualTimeHistoryWGbyDist.Fields("CASE_NUM") = !CASE_NUM
                ladrsActualTimeHistoryWGbyDist.Fields("START_LOC_CD") = !START_LOC_CD
                ladrsActualTimeHistoryWGbyDist.Fields("END_LOC_CD") = !END_LOC_CD
                ladrsActualTimeHistoryWGbyDist.Fields("AMOUNT") = !AMOUNT
                ladrsActualTimeHistoryWGbyDist.Fields("APPR_BY_EMP_NUM") = !APPR_BY_MANAGER
                ladrsActualTimeHistoryWGbyDist.Fields("DISTRIBUTION_TYPE") = !DISTRIBUTION_TYPE
                ladrsActualTimeHistoryWGbyDist.Fields("DISTRIBUTION_ID") = !DISTRIBUTION_ID
                ladrsActualTimeHistoryWGbyDist.Fields("OT_REASON_CD") = !OT_REASON_CD
                ladrsActualTimeHistoryWGbyDist.Fields("ADJ_FLAG") = "ADJ"
                ladrsActualTimeHistoryWGbyDist.Fields("SEQ_NUM") = !SEQ_NUM
                ladrsActualTimeHistoryWGbyDist.Fields("ADJ_TYPE") = !ADJ_TYPE
                If !TRAVEL_PAY_HRS <> 0# Then
                    ladrsActualTimeHistoryWGbyDist.Fields("TRAVEL_PAY_IND") = "Y"
                Else
                    ladrsActualTimeHistoryWGbyDist.Fields("TRAVEL_PAY_IND") = "N"
                End If
                ladrsActualTimeHistoryWGbyDist.Update
                .MoveNext
            Wend
        End With
    End If
    
    Set GetActualTimeHistoryWGbyDist = ladrsActualTimeHistoryWGbyDist


    Exit Function
GetActualTimeHistoryWGbyDist_EH:
    Set l_objWorkgroup = Nothing
    Set l_obj = Nothing
    Set l_adrsEmployees = Nothing
    Set l_objAdj = Nothing
    Set ladrsActualTime = Nothing
    Set ladrsAdj = Nothing
    Set ladrsActualTimeHistoryWGbyDist = Nothing
    Set l_objEmployee = Nothing
    Set ladrsPrimaryWG = Nothing

    TrsRaiseError MODULENAME _
        , "GetActualTimeHistoryWGbyDist"

End Function
Private Sub CalculateEmpDistTotals()
On Error GoTo CalculateEmpDistTotals_EH
Dim l_dDistSymHours As Double, l_dDistOTHours As Double
Dim l_dDistSTHours As Double, l_dDistTotHours As Double
Dim l_sHoldDistId As String, l_sHoldDistType As String
Dim l_sHoldException As String, l_lHoldEmpNum As Long
Dim l_bExceptionsPopulated As Boolean, l_sSymbolTimeInd As String
Dim l_sBankCode As String

    DefineEmpDistTotal

    If (RecordsExist(m_adrsExceptionCodes)) Then
        l_bExceptionsPopulated = True
    Else
        l_bExceptionsPopulated = False
    End If

    With m_adrsReportData
        .MoveFirst

        Do While Not (.EOF)
            l_sHoldDistId = !DISTRIBUTION_ID
            l_sHoldDistType = !DISTRIBUTION_TYPE
            l_lHoldEmpNum = !EMP_NUM

            l_dDistSymHours = 0
            l_dDistOTHours = 0
            l_dDistSTHours = 0
            l_dDistTotHours = 0

            Do While (l_sHoldDistId = !DISTRIBUTION_ID) And _
                (l_lHoldEmpNum = !EMP_NUM) And _
                (l_sHoldDistType = !DISTRIBUTION_TYPE)

                If (!EXCEPTION_CD <> l_sHoldException) Then
                    l_sSymbolTimeInd = "N"
                    l_sHoldException = !EXCEPTION_CD

                    If (l_bExceptionsPopulated) Then
                        m_adrsExceptionCodes.MoveFirst
                        m_adrsExceptionCodes.Find "EXCEPTION_CD = '" & Trim(!EXCEPTION_CD) & "'"

                        If Not m_adrsExceptionCodes.EOF Then
                            l_sSymbolTimeInd = Trim(m_adrsExceptionCodes!SYMBOL_TIME_IND)
                            l_sBankCode = Trim(m_adrsExceptionCodes!EXCEPTION_BANK_CD)
                        End If
                    End If
                End If

                'If (l_sSymbolTimeInd = "N") And _
                '    (!OT_OPP_NUM <= 0) And (!EXCEPTION_CD <> "SSO") And _
                '    (Left(!START_TIME, 2) < "99") And (Left(!END_TIME, 2) < "99") Then
                '    l_dDistSTHours = l_dDistSTHours + !Hours
                'ElseIf (l_sSymbolTimeInd = "N") And _
                '        ((!OT_OPP_NUM > 0) Or (!EXCEPTION_CD = "SSO")) And _
                '        (Left(!START_TIME, 2) < "99") And _
                '        (Left(!END_TIME, 2) < "99") Then
                '        l_dDistOTHours = l_dDistOTHours + !Hours
                'ElseIf (l_sSymbolTimeInd = "Y") And _
                '        ((!OT_OPP_NUM <= 0) Or (!EXCEPTION_CD <> "SSO")) And _
                '        (Left(!START_TIME, 2) < "99") And _
                '        (Left(!END_TIME, 2) < "99") Then
                '        l_dDistSymHours = l_dDistSymHours + !Hours
                'End If
                
                If (l_sSymbolTimeInd = "Y") And (!OT_OPP_NUM = 0) And (Trim(!EXCEPTION_CD) <> "SSO") Then
                    l_dDistSymHours = l_dDistSymHours + !Hours
                    l_dDistTotHours = l_dDistTotHours + !Hours
                ElseIf (!OT_OPP_NUM <> 0) Or (Trim(!EXCEPTION_CD) = "SSO") Or (l_sBankCode = "O") Then
                    l_dDistOTHours = l_dDistOTHours + !Hours
                    l_dDistTotHours = l_dDistTotHours + !Hours
                ElseIf (Trim(!EXCEPTION_CD) <> "LU") Then
                    l_dDistSTHours = l_dDistSTHours + !Hours
                    l_dDistTotHours = l_dDistTotHours + !Hours
                End If

                .MoveNext

                If (.EOF) Then
                    Exit Do
                End If
            Loop

            With m_adrsEmpActualTimeTotals
                .AddNew

                !EMP_NUM = l_lHoldEmpNum
                !DISTRIBUTION_ID = l_sHoldDistId
                !DISTRIBUTION_TYPE = l_sHoldDistType
                !TOT_OT_HOURS = l_dDistOTHours
                !TOT_ST_HOURS = l_dDistSTHours
                !TOT_SYM_HOURS = l_dDistSymHours
                !TOT_HOURS = l_dDistTotHours

                .Update
            End With

            If (.EOF) Then
                Exit Do
            End If
        Loop
    End With

    Exit Sub
CalculateEmpDistTotals_EH:
    TrsRaiseError MODULENAME _
        , "CalculateEmpDistTotals"

End Sub
Private Sub CalculateEmpWGTotals()
On Error GoTo CalculateEmpWGTotals_EH
Dim l_dEmpSymHours As Double, l_dEmpOTHours As Double
Dim l_dEmpSTHours As Double, l_dEmpTotHours As Double
Dim l_lHoldWorkgroup As Long, l_sHoldException As String
Dim l_lHoldEmpNum As Long
Dim l_bExceptionsPopulated As Boolean, l_sSymbolTimeInd As String
Dim l_sBankCode As String

    DefineEmpWGTotal

    If (RecordsExist(m_adrsExceptionCodes)) Then
        l_bExceptionsPopulated = True
    Else
        l_bExceptionsPopulated = False
    End If

    With m_adrsReportData
        .MoveFirst

        Do While Not (.EOF)
            l_lHoldWorkgroup = !WG_NUM
            l_lHoldEmpNum = !EMP_NUM

            l_dEmpSymHours = 0
            l_dEmpOTHours = 0
            l_dEmpSTHours = 0
            l_dEmpTotHours = 0

            Do While (l_lHoldWorkgroup = !WG_NUM) And _
                    (l_lHoldEmpNum = !EMP_NUM)

                If (!EXCEPTION_CD <> l_sHoldException) Then
                    l_sSymbolTimeInd = "N"
                    l_sHoldException = !EXCEPTION_CD

                    If (l_bExceptionsPopulated) Then
                        m_adrsExceptionCodes.MoveFirst
                        m_adrsExceptionCodes.Find "EXCEPTION_CD = '" & Trim(!EXCEPTION_CD) & "'"

                        If Not m_adrsExceptionCodes.EOF Then
                            l_sSymbolTimeInd = Trim(m_adrsExceptionCodes!SYMBOL_TIME_IND)
                            l_sBankCode = Trim(m_adrsExceptionCodes!EXCEPTION_BANK_CD)
                        End If
                    End If
                End If

                'If (l_sSymbolTimeInd = "N") And _
                '    (!OT_OPP_NUM <= 0) And (!EXCEPTION_CD <> "SSO") And _
                '    (Left(!START_TIME, 2) < "99") And (Left(!END_TIME, 2) < "99") Then
                '    l_dEmpSTHours = l_dEmpSTHours + !Hours
                'ElseIf (l_sSymbolTimeInd = "N") And _
                '        ((!OT_OPP_NUM > 0) Or (!EXCEPTION_CD = "SSO")) And _
                '        (Left(!START_TIME, 2) < "99") And _
                '        (Left(!END_TIME, 2) < "99") Then
                '        l_dEmpOTHours = l_dEmpOTHours + !Hours
                'ElseIf (l_sSymbolTimeInd = "Y") And _
                '        ((!OT_OPP_NUM <= 0) Or (!EXCEPTION_CD <> "SSO")) And _
                '        (Left(!START_TIME, 2) < "99") And _
                '        (Left(!END_TIME, 2) < "99") Then
                '        l_dEmpSymHours = l_dEmpSymHours + !Hours
                'End If
                
                If (l_sSymbolTimeInd = "Y") And (!OT_OPP_NUM = 0) And (Trim(!EXCEPTION_CD) <> "SSO") Then
                    l_dEmpSymHours = l_dEmpSymHours + !Hours
                    l_dEmpTotHours = l_dEmpTotHours + !Hours
                ElseIf (!OT_OPP_NUM <> 0) Or (Trim(!EXCEPTION_CD) = "SSO") Or (l_sBankCode = "O") Then
                    l_dEmpOTHours = l_dEmpOTHours + !Hours
                    l_dEmpTotHours = l_dEmpTotHours + !Hours
                ElseIf (Trim(!EXCEPTION_CD) <> "LU") Then
                    l_dEmpSTHours = l_dEmpSTHours + !Hours
                    l_dEmpTotHours = l_dEmpTotHours + !Hours
                End If

                .MoveNext

                If .EOF Then
                    Exit Do
                End If
            Loop

            With m_adrsEmpActualTimeTotals
                .AddNew

                !EMP_NUM = l_lHoldEmpNum
                !WG_NUM = l_lHoldWorkgroup
                !TOT_OT_HOURS = l_dEmpOTHours
                !TOT_ST_HOURS = l_dEmpSTHours
                !TOT_SYM_HOURS = l_dEmpSymHours
                !TOT_HOURS = l_dEmpTotHours

                .Update
            End With

            If .EOF Then
                Exit Do
            End If
        Loop
    End With

    Exit Sub
CalculateEmpWGTotals_EH:
    TrsRaiseError MODULENAME _
        , "CalculateEmpWGTotals"
End Sub
Private Sub CalculateDeptTotals()
On Error GoTo CalculateDeptTotals_EH
Dim l_dDeptSymHours As Double, l_dDeptOTHours As Double
Dim l_dDeptSTHours As Double, l_dDeptTotHours As Double
Dim l_sHoldDeptNum As String, l_sHoldException As String
Dim l_bExceptionsPopulated As Boolean, l_sSymbolTimeInd As String
Dim l_sBankCode As String

    DefineDepartmentTotal

    If (RecordsExist(m_adrsExceptionCodes)) Then
        l_bExceptionsPopulated = True
    Else
        l_bExceptionsPopulated = False
    End If

    With m_adrsReportData
        .MoveFirst

        Do While Not .EOF
            l_sHoldDeptNum = !DEPT_NUM

            l_dDeptSymHours = 0
            l_dDeptOTHours = 0
            l_dDeptSTHours = 0
            l_dDeptTotHours = 0

            Do While (l_sHoldDeptNum = !DEPT_NUM)

                If (!EXCEPTION_CD <> l_sHoldException) Then
                    l_sSymbolTimeInd = "N"
                    l_sHoldException = !EXCEPTION_CD

                    If (l_bExceptionsPopulated) Then
                        m_adrsExceptionCodes.MoveFirst
                        m_adrsExceptionCodes.Find "EXCEPTION_CD = '" & Trim(!EXCEPTION_CD) & "'"
                        If Not (m_adrsExceptionCodes.EOF) Then
                            l_sSymbolTimeInd = Trim(m_adrsExceptionCodes!SYMBOL_TIME_IND)
                            l_sBankCode = Trim(m_adrsExceptionCodes!EXCEPTION_BANK_CD)
                        End If
                    End If
                End If

                'If (l_sSymbolTimeInd = "N") And _
                '    (!OT_OPP_NUM <= 0) And (!EXCEPTION_CD <> "SSO") And _
                '    (Left(!START_TIME, 2) < "99") And (Left(!END_TIME, 2) < "99") Then
                '    l_dDeptSTHours = l_dDeptSTHours + !Hours
                'ElseIf (l_sSymbolTimeInd = "N") And _
                '        ((!OT_OPP_NUM > 0) Or (!EXCEPTION_CD = "SSO")) And _
                '        (Left(!START_TIME, 2) < "99") And _
                '        (Left(!END_TIME, 2) < "99") Then
                '        l_dDeptOTHours = l_dDeptOTHours + !Hours
                'ElseIf (l_sSymbolTimeInd = "Y") And _
                '        ((!OT_OPP_NUM <= 0) Or (!EXCEPTION_CD <> "SSO")) And _
                '        (Left(!START_TIME, 2) < "99") And _
                '        (Left(!END_TIME, 2) < "99") Then
                '        l_dDeptSymHours = l_dDeptSymHours + !Hours
                'End If
                
                If (l_sSymbolTimeInd = "Y") And (!OT_OPP_NUM = 0) And (Trim(!EXCEPTION_CD) <> "SSO") Then
                    l_dDeptSymHours = l_dDeptSymHours + !Hours
                    l_dDeptTotHours = l_dDeptTotHours + !Hours
                ElseIf (!OT_OPP_NUM <> 0) Or (Trim(!EXCEPTION_CD) = "SSO") Or (l_sBankCode = "O") Then
                    l_dDeptOTHours = l_dDeptOTHours + !Hours
                    l_dDeptTotHours = l_dDeptTotHours + !Hours
                ElseIf (Trim(!EXCEPTION_CD) <> "LU") Then
                    l_dDeptSTHours = l_dDeptSTHours + !Hours
                    l_dDeptTotHours = l_dDeptTotHours + !Hours
                End If

                .MoveNext

                If .EOF Then
                    Exit Do
                End If
            Loop

            With m_adrsDeptActualTimeTotals
                .AddNew

                !DEPT_NUM = l_sHoldDeptNum
                !TOT_OT_HOURS = l_dDeptOTHours
                !TOT_ST_HOURS = l_dDeptSTHours
                !TOT_SYM_HOURS = l_dDeptSymHours
                !TOT_HOURS = l_dDeptTotHours

                .Update
            End With

            If .EOF Then
                Exit Do
            End If
        Loop
    End With

    Exit Sub
CalculateDeptTotals_EH:
    TrsRaiseError MODULENAME _
        , "CalculateDeptTotals"
End Sub
Private Sub CalculateDistTotals()
On Error GoTo CalculateDistTotals_EH
Dim l_dDistSymHours As Double, l_dDistOTHours As Double
Dim l_dDistSTHours As Double, l_dDistTotHours As Double
Dim l_sHoldDistId As String, l_sHoldDistType As String
Dim l_sHoldException As String
Dim l_bExceptionsPopulated As Boolean, l_sSymbolTimeInd As String
Dim l_sBankCode As String

    DefineDistTotal

    If (RecordsExist(m_adrsExceptionCodes)) Then
        l_bExceptionsPopulated = True
    Else
        l_bExceptionsPopulated = False
    End If

    With m_adrsReportData
        .MoveFirst

        Do While Not .EOF
            l_sHoldDistId = !DISTRIBUTION_ID
            l_sHoldDistType = !DISTRIBUTION_TYPE

            l_dDistSymHours = 0
            l_dDistOTHours = 0
            l_dDistSTHours = 0
            l_dDistTotHours = 0
            Do While (l_sHoldDistId = !DISTRIBUTION_ID) And _
                    (l_sHoldDistType = !DISTRIBUTION_TYPE)

                If (!EXCEPTION_CD <> l_sHoldException) Then
                    l_sSymbolTimeInd = "N"
                    l_sHoldException = !EXCEPTION_CD

                    If (l_bExceptionsPopulated) Then
                        m_adrsExceptionCodes.MoveFirst
                        m_adrsExceptionCodes.Find "EXCEPTION_CD = '" & Trim(!EXCEPTION_CD) & "'"
                        If Not (m_adrsExceptionCodes.EOF) Then
                            l_sSymbolTimeInd = Trim(m_adrsExceptionCodes!SYMBOL_TIME_IND)
                            l_sBankCode = Trim(m_adrsExceptionCodes!EXCEPTION_BANK_CD)
                        End If
                    End If
                End If

                'If (l_sSymbolTimeInd = "N") And _
                '    (!OT_OPP_NUM <= 0) And (!EXCEPTION_CD <> "SSO") And _
                '    (Left(!START_TIME, 2) < "99") And (Left(!END_TIME, 2) < "99") Then
                '    l_dDistSTHours = l_dDistSTHours + !Hours
                'ElseIf (l_sSymbolTimeInd = "N") And _
                '        ((!OT_OPP_NUM > 0) Or (!EXCEPTION_CD = "SSO")) And _
                '        (Left(!START_TIME, 2) < "99") And _
                '        (Left(!END_TIME, 2) < "99") Then
                '        l_dDistOTHours = l_dDistOTHours + !Hours
                'ElseIf (l_sSymbolTimeInd = "Y") And _
                '        ((!OT_OPP_NUM <= 0) Or (!EXCEPTION_CD <> "SSO")) And _
                '        (Left(!START_TIME, 2) < "99") And _
                '        (Left(!END_TIME, 2) < "99") Then
                '        l_dDistSymHours = l_dDistSymHours + !Hours
                'End If
                
                If (l_sSymbolTimeInd = "Y") And (!OT_OPP_NUM = 0) And (Trim(!EXCEPTION_CD) <> "SSO") Then
                    l_dDistSymHours = l_dDistSymHours + !Hours
                    l_dDistTotHours = l_dDistTotHours + !Hours
                ElseIf (!OT_OPP_NUM <> 0) Or (Trim(!EXCEPTION_CD) = "SSO") Or (l_sBankCode = "O") Then
                    l_dDistOTHours = l_dDistOTHours + !Hours
                    l_dDistTotHours = l_dDistTotHours + !Hours
                ElseIf (Trim(!EXCEPTION_CD) <> "LU") Then
                    l_dDistSTHours = l_dDistSTHours + !Hours
                    l_dDistTotHours = l_dDistTotHours + !Hours
                End If

                .MoveNext

                If .EOF Then
                    Exit Do
                End If
            Loop

            With m_adrsDistActualTimeTotals
                .AddNew

                !DISTRIBUTION_ID = l_sHoldDistId
                !DISTRIBUTION_TYPE = l_sHoldDistType
                !TOT_OT_HOURS = l_dDistOTHours
                !TOT_ST_HOURS = l_dDistSTHours
                !TOT_SYM_HOURS = l_dDistSymHours
                !TOT_HOURS = l_dDistTotHours

                .Update
            End With

            If .EOF Then
                Exit Do
            End If
        Loop
    End With

    Exit Sub
CalculateDistTotals_EH:
    TrsRaiseError MODULENAME _
        , "CalculateDistTotals"

End Sub
Private Sub CalculateWGTotals()
On Error GoTo CalculateWGTotals_EH
Dim l_dWGSymHours As Double, l_dWGOTHours As Double
Dim l_dWGSTHours As Double, l_dWGTotHours As Double
Dim l_lHoldWGNum As Long, l_sHoldException As String
Dim l_bExceptionsPopulated As Boolean, l_sSymbolTimeInd As String
Dim l_sBankCode As String

    DefineWGTotal

    If (RecordsExist(m_adrsExceptionCodes)) Then
        l_bExceptionsPopulated = True
    Else
        l_bExceptionsPopulated = False
    End If

    With m_adrsReportData
        .MoveFirst

        Do While Not .EOF
            l_lHoldWGNum = !WG_NUM

            l_dWGSymHours = 0
            l_dWGOTHours = 0
            l_dWGSTHours = 0
            l_dWGTotHours = 0

            Do While (l_lHoldWGNum = !WG_NUM)

                If (!EXCEPTION_CD <> l_sHoldException) Then
                    l_sSymbolTimeInd = "N"
                    l_sHoldException = !EXCEPTION_CD

                    If (l_bExceptionsPopulated) Then
                        m_adrsExceptionCodes.MoveFirst
                        m_adrsExceptionCodes.Find "EXCEPTION_CD = '" & Trim(!EXCEPTION_CD) & "'"

                        If Not (m_adrsExceptionCodes.EOF) Then
                            l_sSymbolTimeInd = Trim(m_adrsExceptionCodes!SYMBOL_TIME_IND)
                            l_sBankCode = Trim(m_adrsExceptionCodes!EXCEPTION_BANK_CD)
                        End If
                    End If
                End If

                'If (l_sSymbolTimeInd = "N") And _
                '    (!OT_OPP_NUM <= 0) And (!EXCEPTION_CD <> "SSO") And _
                '    (Left(!START_TIME, 2) < "99") And (Left(!END_TIME, 2) < "99") Then
                '    l_dWGSTHours = l_dWGSTHours + !Hours
                'ElseIf (l_sSymbolTimeInd = "N") And _
                '        ((!OT_OPP_NUM > 0) Or (!EXCEPTION_CD = "SSO")) And _
                '        (Left(!START_TIME, 2) < "99") And _
                '        (Left(!END_TIME, 2) < "99") Then
                '        l_dWGOTHours = l_dWGOTHours + !Hours
                'ElseIf (l_sSymbolTimeInd = "Y") And _
                '        ((!OT_OPP_NUM <= 0) Or (!EXCEPTION_CD <> "SSO")) And _
                '        (Left(!START_TIME, 2) < "99") And _
                '        (Left(!END_TIME, 2) < "99") Then
                '        l_dWGSymHours = l_dWGSymHours + !Hours
                'End If
                
                If (l_sSymbolTimeInd = "Y") And (!OT_OPP_NUM = 0) And (Trim(!EXCEPTION_CD) <> "SSO") Then
                    l_dWGSymHours = l_dWGSymHours + !Hours
                    l_dWGTotHours = l_dWGTotHours + !Hours
                ElseIf (!OT_OPP_NUM <> 0) Or (Trim(!EXCEPTION_CD) = "SSO") Or (l_sBankCode = "O") Then
                    l_dWGOTHours = l_dWGOTHours + !Hours
                    l_dWGTotHours = l_dWGTotHours + !Hours
                ElseIf (Trim(!EXCEPTION_CD) <> "LU") Then
                    l_dWGSTHours = l_dWGSTHours + !Hours
                    l_dWGTotHours = l_dWGTotHours + !Hours
                End If

                .MoveNext

                If .EOF Then
                    Exit Do
                End If
            Loop

            With m_adrsWGActualTimeTotals
                .AddNew

                !WG_NUM = l_lHoldWGNum
                !TOT_OT_HOURS = l_dWGOTHours
                !TOT_ST_HOURS = l_dWGSTHours
                !TOT_SYM_HOURS = l_dWGSymHours
                !TOT_HOURS = l_dWGTotHours

                .Update
            End With

            If .EOF Then
                Exit Do
            End If
        Loop
    End With

    Exit Sub
CalculateWGTotals_EH:
    TrsRaiseError MODULENAME _
        , "CalculateWGTotals"
End Sub
Private Sub CalculateEmpTotals()
On Error GoTo CalculateEmpTotals_EH
Dim l_dEmpSymHours As Double, l_dEmpOTHours As Double
Dim l_dEmpSTHours As Double, l_dEmpTotHours As Double
Dim l_lHoldEmpNum As Long, l_sHoldException As String
Dim l_bExceptionsPopulated As Boolean, l_sSymbolTimeInd As String
Dim l_sBankCode As String

    DefineEmployeeTotal

    If (RecordsExist(m_adrsExceptionCodes)) Then
        l_bExceptionsPopulated = True
    Else
        l_bExceptionsPopulated = False
    End If

    With m_adrsReportData
        .MoveFirst

        Do While Not .EOF
            l_lHoldEmpNum = !EMP_NUM

            l_dEmpSymHours = 0
            l_dEmpOTHours = 0
            l_dEmpSTHours = 0
            l_dEmpTotHours = 0

            Do While (l_lHoldEmpNum = !EMP_NUM)

                If (!EXCEPTION_CD <> l_sHoldException) Then
                    l_sSymbolTimeInd = "N"
                    l_sHoldException = !EXCEPTION_CD

                    If (l_bExceptionsPopulated) Then
                        m_adrsExceptionCodes.MoveFirst
                        m_adrsExceptionCodes.Find "EXCEPTION_CD = '" & Trim(!EXCEPTION_CD) & "'"

                        If Not (m_adrsExceptionCodes.EOF) Then
                            l_sSymbolTimeInd = Trim(m_adrsExceptionCodes!SYMBOL_TIME_IND)
                            l_sBankCode = Trim(m_adrsExceptionCodes!EXCEPTION_BANK_CD)
                        End If
                    End If
                End If

                'If (l_sSymbolTimeInd = "N") And _
                '    (!OT_OPP_NUM <= 0) And (!EXCEPTION_CD <> "SSO") And _
                '    (Left(!START_TIME, 2) < "99") And (Left(!END_TIME, 2) < "99") Then
                '    l_dEmpSTHours = l_dEmpSTHours + !Hours
                'ElseIf (l_sSymbolTimeInd = "N") And _
                '        ((!OT_OPP_NUM > 0) Or (!EXCEPTION_CD = "SSO")) And _
                '        (Left(!START_TIME, 2) < "99") And _
                '        (Left(!END_TIME, 2) < "99") Then
                '        l_dEmpOTHours = l_dEmpOTHours + !Hours
                'ElseIf (l_sSymbolTimeInd = "Y") And _
                '        ((!OT_OPP_NUM <= 0) Or (!EXCEPTION_CD <> "SSO")) And _
                '        (Left(!START_TIME, 2) < "99") And _
                '        (Left(!END_TIME, 2) < "99") Then
                '        l_dEmpSymHours = l_dEmpSymHours + !Hours
                'End If
                
                If (l_sSymbolTimeInd = "Y") And (!OT_OPP_NUM = 0) And (Trim(!EXCEPTION_CD) <> "SSO") Then
                    l_dEmpSymHours = l_dEmpSymHours + !Hours
                    l_dEmpTotHours = l_dEmpTotHours + !Hours
                ElseIf (!OT_OPP_NUM <> 0) Or (Trim(!EXCEPTION_CD) = "SSO") Or (l_sBankCode = "O") Then
                    l_dEmpOTHours = l_dEmpOTHours + !Hours
                    l_dEmpTotHours = l_dEmpTotHours + !Hours
                ElseIf (Trim(!EXCEPTION_CD) <> "LU") Then
                    l_dEmpSTHours = l_dEmpSTHours + !Hours
                    l_dEmpTotHours = l_dEmpTotHours + !Hours
                End If

                .MoveNext

                If .EOF Then
                    Exit Do
                End If
            Loop

            With m_adrsEmpActualTimeTotals
                .AddNew

                !EMP_NUM = l_lHoldEmpNum
                !TOT_OT_HOURS = l_dEmpOTHours
                !TOT_ST_HOURS = l_dEmpSTHours
                !TOT_SYM_HOURS = l_dEmpSymHours
                !TOT_HOURS = l_dEmpTotHours

                .Update
            End With

            If .EOF Then
                Exit Do
            End If
        Loop
    End With

    Exit Sub
CalculateEmpTotals_EH:
    TrsRaiseError MODULENAME _
        , "CalculateEmpTotals"
End Sub
Private Sub CalculateEmpDeptTotals()
On Error GoTo CalculateEmpDeptTotals_EH
Dim l_dEmpSymHours As Double, l_dEmpOTHours As Double
Dim l_dEmpSTHours As Double, l_dEmpTotHours As Double
Dim l_lHoldEmpNum As Long, l_sHoldException As String
Dim l_sHoldDeptNum  As String, l_sSymbolTimeInd As String
Dim l_sBankCode As String
Dim l_bExceptionsPopulated As Boolean

    DefineEmpDeptTotals

    If (RecordsExist(m_adrsExceptionCodes)) Then
        l_bExceptionsPopulated = True
    Else
        l_bExceptionsPopulated = False
    End If

    With m_adrsReportData
        .MoveFirst

        Do While Not .EOF
            l_lHoldEmpNum = !EMP_NUM
            l_sHoldDeptNum = !DEPT_NUM

            l_dEmpSymHours = 0
            l_dEmpOTHours = 0
            l_dEmpSTHours = 0
            l_dEmpTotHours = 0

            Do While (l_lHoldEmpNum = !EMP_NUM) And _
                    (l_sHoldDeptNum = !DEPT_NUM)

                If (!EXCEPTION_CD <> l_sHoldException) Then
                    l_sSymbolTimeInd = "N"
                    l_sHoldException = !EXCEPTION_CD

                    If (l_bExceptionsPopulated) Then
                        m_adrsExceptionCodes.MoveFirst
                        m_adrsExceptionCodes.Find "EXCEPTION_CD = '" & Trim(!EXCEPTION_CD) & "'"

                        If Not (m_adrsExceptionCodes.EOF) Then
                            l_sSymbolTimeInd = Trim(m_adrsExceptionCodes!SYMBOL_TIME_IND)
                            l_sBankCode = Trim(m_adrsExceptionCodes!EXCEPTION_BANK_CD)
                        End If
                    End If
                End If

                'If (l_sSymbolTimeInd = "N") And _
                '    (!OT_OPP_NUM <= 0) And (!EXCEPTION_CD <> "SSO") And _
                '    (Left(!START_TIME, 2) < "99") And (Left(!END_TIME, 2) < "99") Then
                '    l_dEmpSTHours = l_dEmpSTHours + !Hours
                'ElseIf (l_sSymbolTimeInd = "N") And _
                '        ((!OT_OPP_NUM > 0) Or (!EXCEPTION_CD = "SSO")) And _
                '        (Left(!START_TIME, 2) < "99") And _
                '        (Left(!END_TIME, 2) < "99") Then
                '        l_dEmpOTHours = l_dEmpOTHours + !Hours
                'ElseIf (l_sSymbolTimeInd = "Y") And _
                '        ((!OT_OPP_NUM <= 0) Or (!EXCEPTION_CD <> "SSO")) And _
                '        (Left(!START_TIME, 2) < "99") And _
                '        (Left(!END_TIME, 2) < "99") Then
                '        l_dEmpSymHours = l_dEmpSymHours + !Hours
                'End If
                
                If (l_sSymbolTimeInd = "Y") And (!OT_OPP_NUM = 0) And (Trim(!EXCEPTION_CD) <> "SSO") Then
                    l_dEmpSymHours = l_dEmpSymHours + !Hours
                    l_dEmpTotHours = l_dEmpTotHours + !Hours
                ElseIf (!OT_OPP_NUM <> 0) Or (Trim(!EXCEPTION_CD) = "SSO") Or (l_sBankCode = "O") Then
                    l_dEmpOTHours = l_dEmpOTHours + !Hours
                    l_dEmpTotHours = l_dEmpTotHours + !Hours
                ElseIf (Trim(!EXCEPTION_CD) <> "LU") Then
                    l_dEmpSTHours = l_dEmpSTHours + !Hours
                    l_dEmpTotHours = l_dEmpTotHours + !Hours
                End If

                .MoveNext

                If .EOF Then
                    Exit Do
                End If
            Loop

            With m_adrsEmpActualTimeTotals
                .AddNew

                !EMP_NUM = l_lHoldEmpNum
                !DEPT_NUM = l_sHoldDeptNum
                !TOT_OT_HOURS = l_dEmpOTHours
                !TOT_ST_HOURS = l_dEmpSTHours
                !TOT_SYM_HOURS = l_dEmpSymHours
                !TOT_HOURS = l_dEmpTotHours

                .Update
            End With

            If .EOF Then
                Exit Do
            End If
        Loop
    End With

    Exit Sub
CalculateEmpDeptTotals_EH:
    TrsRaiseError MODULENAME _
        , "CalculateEmpDeptTotals"
End Sub
Private Sub DefineEmployeeTotal()
On Error GoTo DefineEmployeeTotal_EH

    Set m_adrsEmpActualTimeTotals = New ADODB.Recordset
    With m_adrsEmpActualTimeTotals
        .Fields.Append "EMP_NUM", adInteger, , m_conFIELDATTR
        .Fields.Append "TOT_OT_HOURS", adDouble, , m_conFIELDATTR
        .Fields.Append "TOT_ST_HOURS", adDouble, , m_conFIELDATTR
        .Fields.Append "TOT_SYM_HOURS", adDouble, , m_conFIELDATTR
        .Fields.Append "TOT_HOURS", adDouble, , m_conFIELDATTR

        .CursorLocation = adUseClient
        .LockType = adLockBatchOptimistic
        .CursorType = adOpenStatic
        .Open
    End With
    
    Exit Sub
DefineEmployeeTotal_EH:
    TrsRaiseError MODULENAME _
        , "DefineEmployeeTotal"

End Sub
Private Sub DefineEmpDeptTotals()
On Error GoTo DefineEmpDeptTotals_EH

    Set m_adrsEmpActualTimeTotals = New ADODB.Recordset
    With m_adrsEmpActualTimeTotals
        .Fields.Append "DEPT_NUM", adChar, 3, m_conFIELDATTR
        .Fields.Append "EMP_NUM", adInteger, , m_conFIELDATTR
        .Fields.Append "TOT_OT_HOURS", adDouble, , m_conFIELDATTR
        .Fields.Append "TOT_ST_HOURS", adDouble, , m_conFIELDATTR
        .Fields.Append "TOT_SYM_HOURS", adDouble, , m_conFIELDATTR
        .Fields.Append "TOT_HOURS", adDouble, , m_conFIELDATTR

        .CursorLocation = adUseClient
        .LockType = adLockBatchOptimistic
        .CursorType = adOpenStatic
        .Open
    End With

    Exit Sub
DefineEmpDeptTotals_EH:
    TrsRaiseError MODULENAME _
        , "DefineEmpDeptTotals"

End Sub
Private Sub DefineWGTotal()

    Set m_adrsWGActualTimeTotals = New ADODB.Recordset
    With m_adrsWGActualTimeTotals
        .Fields.Append "WG_NUM", adInteger, , m_conFIELDATTR
        .Fields.Append "TOT_OT_HOURS", adDouble, , m_conFIELDATTR
        .Fields.Append "TOT_ST_HOURS", adDouble, , m_conFIELDATTR
        .Fields.Append "TOT_SYM_HOURS", adDouble, , m_conFIELDATTR
        .Fields.Append "TOT_HOURS", adDouble, , m_conFIELDATTR
        .CursorLocation = adUseClient
        .LockType = adLockBatchOptimistic
        .CursorType = adOpenStatic
        .Open
    End With

    Exit Sub
GetActualTimeHistoryEmpRpt_EH:
    TrsRaiseError MODULENAME _
        , "GetActualTimeHistoryEmpRpt"

End Sub
Private Sub DefineDepartmentTotal()
On Error GoTo DefineDepartmentTotal_EH

    Set m_adrsDeptActualTimeTotals = New ADODB.Recordset
    With m_adrsDeptActualTimeTotals
        .Fields.Append "DEPT_NUM", adChar, 3, m_conFIELDATTR
        .Fields.Append "TOT_OT_HOURS", adDouble, , m_conFIELDATTR
        .Fields.Append "TOT_ST_HOURS", adDouble, , m_conFIELDATTR
        .Fields.Append "TOT_SYM_HOURS", adDouble, , m_conFIELDATTR
        .Fields.Append "TOT_HOURS", adDouble, , m_conFIELDATTR

        .CursorLocation = adUseClient
        .LockType = adLockBatchOptimistic
        .CursorType = adOpenStatic
        .Open
    End With

    Exit Sub
DefineDepartmentTotal_EH:
    TrsRaiseError MODULENAME _
        , "DefineDepartmentTotal"

End Sub
Private Sub DefineEmpDistTotal()
On Error GoTo DefineEmpDistTotal_EH

    Set m_adrsEmpActualTimeTotals = New ADODB.Recordset
    With m_adrsEmpActualTimeTotals
        .Fields.Append "DISTRIBUTION_ID", adChar, 40, m_conFIELDATTR
        .Fields.Append "DISTRIBUTION_TYPE", adChar, 1, m_conFIELDATTR
        .Fields.Append "EMP_NUM", adInteger, , m_conFIELDATTR
        .Fields.Append "TOT_OT_HOURS", adDouble, , m_conFIELDATTR
        .Fields.Append "TOT_ST_HOURS", adDouble, , m_conFIELDATTR
        .Fields.Append "TOT_SYM_HOURS", adDouble, , m_conFIELDATTR
        .Fields.Append "TOT_HOURS", adDouble, , m_conFIELDATTR

        .CursorLocation = adUseClient
        .LockType = adLockBatchOptimistic
        .CursorType = adOpenStatic
        .Open
    End With

    Exit Sub
DefineEmpDistTotal_EH:
    TrsRaiseError MODULENAME _
        , "DefineEmpDistTotal"

End Sub
Private Sub DefineEmpWGTotal()
On Error GoTo DefineEmpWGTotal_EH

    Set m_adrsEmpActualTimeTotals = New ADODB.Recordset
    With m_adrsEmpActualTimeTotals
        .Fields.Append "WG_NUM", adChar, 40, m_conFIELDATTR
        .Fields.Append "EMP_NUM", adInteger, , m_conFIELDATTR
        .Fields.Append "TOT_OT_HOURS", adDouble, , m_conFIELDATTR
        .Fields.Append "TOT_ST_HOURS", adDouble, , m_conFIELDATTR
        .Fields.Append "TOT_SYM_HOURS", adDouble, , m_conFIELDATTR
        .Fields.Append "TOT_HOURS", adDouble, , m_conFIELDATTR

        .CursorLocation = adUseClient
        .LockType = adLockBatchOptimistic
        .CursorType = adOpenStatic
        .Open
    End With

    Exit Sub
DefineEmpWGTotal_EH:
    TrsRaiseError MODULENAME _
        , "DefineEmpWGTotal"

End Sub
Private Sub DefineDistTotal()
On Error GoTo DefineDistTotal_EH

    If Not m_adrsDistActualTimeTotals Is Nothing Then
        Set m_adrsDistActualTimeTotals = Nothing
    End If

    Set m_adrsDistActualTimeTotals = New ADODB.Recordset
    With m_adrsDistActualTimeTotals
        .Fields.Append "DISTRIBUTION_ID", adChar, 40, m_conFIELDATTR
        .Fields.Append "DISTRIBUTION_TYPE", adChar, 1, m_conFIELDATTR
        .Fields.Append "TOT_OT_HOURS", adDouble, , m_conFIELDATTR
        .Fields.Append "TOT_ST_HOURS", adDouble, , m_conFIELDATTR
        .Fields.Append "TOT_SYM_HOURS", adDouble, , m_conFIELDATTR
        .Fields.Append "TOT_HOURS", adDouble, , m_conFIELDATTR

        .CursorLocation = adUseClient
        .LockType = adLockBatchOptimistic
        .CursorType = adOpenStatic
        .Open
    End With

    Exit Sub
DefineDistTotal_EH:
    TrsRaiseError MODULENAME _
        , "DefineDistTotal"

End Sub
Public Function GetActualTimeEmpbyDistRpt( _
                ByVal v_dtStartDate As Date _
                , ByVal v_dtEndDate As Date _
                , ByVal v_adrsEmployeesSelected As ADODB.Recordset _
                , ByVal v_adrsExceptionCodes As ADODB.Recordset _
                , ByRef r_bReportGenerated As Boolean _
                , ByVal v_bActiveEmployees As Boolean _
                , ByVal v_adrsAllDepartments As ADODB.Recordset _
                , ByVal v_adrsSelDistributions As ADODB.Recordset _
                ) As ActiveReport
On Error GoTo GetActualTimeEmpbyDistRpt_EH

Dim lstrFilter As String
Dim lstrDist As String
Dim llngIndex As Long
Dim i As Integer

    'Initialize return
    r_bReportGenerated = False
    ResetActiveReportsError
    Set GetActualTimeEmpbyDistRpt = Nothing

    If (RecordsExist(v_adrsEmployeesSelected)) Then
        Set m_adrsSelectedEmployees = v_adrsEmployeesSelected
        Set m_adrsAllEmployees = m_adrsSelectedEmployees.Clone
    Else
        Exit Function
    End If

    If (RecordsExist(v_adrsAllDepartments)) Then
        Set m_adrsAllDepartments = v_adrsAllDepartments
    Else
        Exit Function
    End If

    m_lReportLevel = 1 'Employee Level Report

    m_dtStartDate = v_dtStartDate
    m_dtEndDate = v_dtEndDate

    Set m_adrsExceptionCodes = v_adrsExceptionCodes

    'Load XML file into the report object
    Set rptMain = CreateObject("DDActiveReports2.ActiveReport")
    rptMain.LoadLayout "C:\TRIS Solution\TRISReports\RPX Files\rptActualTimeHistoryEmp.rpx"
    'Add a reference in the rpt object to this (rptException) class
    Call rptMain.AddNamedItem("MyClass", Me)

    Set rptSubReport1 = CreateObject("DDActiveReports2.ActiveReport")
    rptSubReport1.LoadLayout "C:\TRIS Solution\TRISReports\RPX Files\rptActualTimeTotalsSub.rpx"
    'Add a reference in the rpt object to this (rptException) class
    Call rptSubReport1.AddNamedItem("MyClass", Me)

    rptMain.Sections("GroupFooter1").Controls("subRptTotals").object = _
        rptSubReport1
    rptMain.Sections("PageHeader").Controls("txtHdrRunDate").Text = _
        Now()
    rptMain.Sections("PageHeader").Controls("txtHdrStartDate").Text = _
        CStr(Format(v_dtStartDate, "mm/dd/yyyy"))
    rptMain.Sections("PageHeader").Controls("txtHdrEndDate").Text = _
        CStr(Format(v_dtEndDate, "mm/dd/yyyy"))

    Set m_adrsReportData = GetActualTimeHistoryEmpbyDist( _
                            m_adrsSelectedEmployees _
                            , v_dtStartDate _
                            , v_dtEndDate _
                            , v_adrsSelDistributions)
    'If (RecordsExist(m_adrsReportData)) Then
    'place filter here if distributions are supplied
    '    lstrFilter = ""
    '    lstrDist = ""
    '    If RecordsExist(v_adrsSelDistributions) Then
    '        With v_adrsSelDistributions
    '            .MoveFirst
    '            lstrDist = Trim(.Fields("DISTRIBUTION_ID"))
    '            For i = 1 To Len(lstrDist)
    '                If Mid(lstrDist, i, 1) = " " Then
    '                    Mid(lstrDist, i, 1) = "%"
    '                End If
    '            Next i
    '            lstrFilter = _
    '            "DISTRIBUTION_ID LIKE '%" & _
    '            lstrDist & "%'"
    '            .MoveNext
    '            If Not .EOF Then
    '                For llngIndex = 2 To .RecordCount
    '                    lstrDist = Trim(.Fields("DISTRIBUTION_ID"))
    '                    For i = 0 To Len(lstrDist)
    '                        If Mid(lstrDist, i, 1) = " " Then
    '                            Mid(lstrDist, i, 1) = "%"
    '                        End If
    '                    Next i
    '                    lstrFilter = lstrFilter & " OR" & _
    '                    " DISTRIBUTION_ID LIKE '%" & _
    '                    lstrDist & "%'"
    '                    .MoveNext
    '                Next llngIndex
    '            End If
    '        End With
        
    '        If lstrFilter <> "" Then
                'm_adrsReportData.Filter = stdmTRISCommon.MDAC25FilterFix(lstrFilter)
    '            m_adrsReportData.Filter = lstrFilter
    '        End If
            
            If m_adrsReportData.EOF Then
               Exit Function
            End If
    '    End If
        
        r_bReportGenerated = True

        m_adrsReportData.Sort = "EMP_NUM, CHARGE_DATE"
        m_adrsReportData.MoveFirst

        CalculateEmpTotals

        m_adrsReportData.Sort = "EMP_NUM, CHARGE_DATE, DATE_WORKED, START_TIME, END_TIME, SEQ_NUM, HOURS"
        m_adrsReportData.MoveFirst

        'Return the report object to the calling application
        rptMain.Run
        
        '//check for any errors from the active reports script or the vb6 functions being called back
        If m_ActiveReports_ErrNumber <> 0 Then
            Err.Raise m_ActiveReports_ErrNumber, m_ActiveReports_ErrSource, m_ActiveReports_ErrDescription
        End If
        
        Set GetActualTimeEmpbyDistRpt = rptMain
    'End If

    '//-----clean up public properties used by ActiveReports-----//
    DestroyState
        
    Exit Function
GetActualTimeEmpbyDistRpt_EH:
    DestroyState
    
    Select Case Err.Number
    Case m_ActiveReports_ErrNumber
        TrsRaiseError MODULENAME, "GetActualTimeEmpbyDistRpt", Err.Number, , , _
            DescArgsFrom("[OrigErrDescription]===>" & m_ActiveReports_OrigErrDescription & "<===", "[OrigErrSource]===>" & m_ActiveReports_OrigErrSource & "<==="), True
    Case Else
        TrsRaiseError MODULENAME, "GetActualTimeEmpbyDistRpt", , , , , True
    End Select
End Function

Public Function GetActualTimeDeptbyDistRpt( _
                ByVal v_dtStartDate As Date _
                , ByVal v_dtEndDate As Date _
                , ByVal v_adrsSelectedDepartments As ADODB.Recordset _
                , ByVal v_adrsExceptionCodes As ADODB.Recordset _
                , ByRef r_bReportGenerated As Boolean _
                , ByVal v_bActiveEmployees As Boolean _
                , ByVal v_adrsAllEmployees As ADODB.Recordset _
                , ByVal v_adrsSelDistributions As ADODB.Recordset _
                ) As ActiveReport
On Error GoTo GetActualTimeDeptbyDistRpt_EH

Dim lstrFilter As String
Dim lstrDist As String
Dim llngIndex As Long
Dim i As Integer

    'Initialize return
    r_bReportGenerated = False
    ResetActiveReportsError
    Set GetActualTimeDeptbyDistRpt = Nothing

    If (RecordsExist(v_adrsSelectedDepartments)) Then
        Set m_adrsSelectedDepartments = v_adrsSelectedDepartments
        Set m_adrsAllDepartments = m_adrsSelectedDepartments.Clone
        Set m_adrsAllEmployees = v_adrsAllEmployees
    Else
        Exit Function
    End If

    m_lReportLevel = 3

    m_dtStartDate = v_dtStartDate
    m_dtEndDate = v_dtEndDate

    Set m_adrsExceptionCodes = v_adrsExceptionCodes

    'Load XML file into the report object
    Set rptMain = CreateObject("DDActiveReports2.ActiveReport")
    rptMain.LoadLayout "C:\TRIS Solution\TRISReports\RPX Files\rptActualTimeHistoryDept.rpx"
    'Add a reference in the rpt object to this (rptException) class
    Call rptMain.AddNamedItem("MyClass", Me)

    Set rptSubReport1 = CreateObject("DDActiveReports2.ActiveReport")
    rptSubReport1.LoadLayout "C:\TRIS Solution\TRISReports\RPX Files\rptActualTimeTotalsSub.rpx"
    'Add a reference in the rpt object to this (rptException) class
    Call rptSubReport1.AddNamedItem("MyClass", Me)

    Set rptSubReport2 = CreateObject("DDActiveReports2.ActiveReport")
    rptSubReport2.LoadLayout "C:\TRIS Solution\TRISReports\RPX Files\rptActualTimeTotalsSub.rpx"
    'Add a reference in the rpt object to this (rptException) class
    Call rptSubReport2.AddNamedItem("MyClass", Me)

    rptMain.Sections("GroupFooter2").Controls("subRptEmpTotals").object = _
        rptSubReport1
    rptMain.Sections("GroupFooter1").Controls("subRptDeptTotals").object = _
        rptSubReport2
    rptMain.Sections("PageHeader").Controls("txtHdrRunDate").Text = _
        Now()
    rptMain.Sections("PageHeader").Controls("txtHdrStartDate").Text = _
        CStr(Format(v_dtStartDate, "mm/dd/yyyy"))
    rptMain.Sections("PageHeader").Controls("txtHdrEndDate").Text = _
        CStr(Format(v_dtEndDate, "mm/dd/yyyy"))

    Set m_adrsReportData = GetActualTimeHistoryDeptbyDist( _
                            m_adrsSelectedDepartments _
                            , v_dtStartDate _
                            , v_dtEndDate _
                            , v_adrsSelDistributions)
    'If (RecordsExist(m_adrsReportData)) Then
    'place filter here if distributions are supplied
    '    lstrFilter = ""
    '    lstrDist = ""
    '    If RecordsExist(v_adrsSelDistributions) Then
    '        With v_adrsSelDistributions
    '            .MoveFirst
    '            lstrDist = Trim(.Fields("DISTRIBUTION_ID"))
    '            For i = 0 To Len(lstrDist)
    '                If Mid(lstrDist, i, 1) = " " Then
    '                    Mid(lstrDist, i, 1) = "%"
    '                End If
    '            Next i
    '            lstrFilter = _
    '            "DISTRIBUTION_ID LIKE '%" & _
    '            lstrDist & "%'"
    '            If Not .EOF Then
    '                .MoveNext
    '                For llngIndex = 2 To .RecordCount
    '                        lstrDist = Trim(.Fields("DISTRIBUTION_ID"))
    '                        For i = 1 To Len(lstrDist)
    '                            If Mid(lstrDist, i, 1) = " " Then
    '                                Mid(lstrDist, i, 1) = "%"
    '                            End If
    '                        Next i
    '                        lstrFilter = lstrFilter & " OR" & _
    '                        " DISTRIBUTION_ID LIKE '%" & _
    '                        lstrDist & "%'"
    '                    .MoveNext
    '                Next llngIndex
    '            End If
    '        End With
        
    '        If lstrFilter <> "" Then
    '            m_adrsReportData.Filter = stdmTRISCommon.MDAC25FilterFix(lstrFilter)
    '        End If
            
            If m_adrsReportData.EOF Then
               Exit Function
            End If
    '    End If
        
        r_bReportGenerated = True

        m_adrsReportData.Sort = "DEPT_NUM, EXCEPTION_CD"
        m_adrsReportData.MoveFirst

        CalculateDeptTotals

        m_adrsReportData.Sort = "DEPT_NUM, EMP_NUM, EXCEPTION_CD"
        m_adrsReportData.MoveFirst

        CalculateEmpDeptTotals

        m_adrsReportData.Sort = "DEPT_NUM, EMP_NUM, CHARGE_DATE, DATE_WORKED, START_TIME, END_TIME, SEQ_NUM, HOURS"
        m_adrsReportData.MoveFirst

        rptMain.Run
        
        '//check for any errors from the active reports script or the vb6 functions being called back
        If m_ActiveReports_ErrNumber <> 0 Then
            Err.Raise m_ActiveReports_ErrNumber, m_ActiveReports_ErrSource, m_ActiveReports_ErrDescription
        End If

        'Return the report object to the calling application
        Set GetActualTimeDeptbyDistRpt = rptMain
    'End If

    '//-----clean up public properties used by ActiveReports-----//
    DestroyState
        
    Exit Function
GetActualTimeDeptbyDistRpt_EH:
    DestroyState
    
    Select Case Err.Number
    Case m_ActiveReports_ErrNumber
        TrsRaiseError MODULENAME, "GetActualTimeDeptbyDistRpt", Err.Number, , , _
            DescArgsFrom("[OrigErrDescription]===>" & m_ActiveReports_OrigErrDescription & "<===", "[OrigErrSource]===>" & m_ActiveReports_OrigErrSource & "<==="), True
    Case Else
        TrsRaiseError MODULENAME, "GetActualTimeDeptbyDistRpt", , , , , True
    End Select
End Function

Public Function GetActualTimeWGbyDistRpt( _
                ByVal v_dtStartDate As Date _
                , ByVal v_dtEndDate As Date _
                , ByVal v_adrsSelectedWorkgroups As ADODB.Recordset _
                , ByVal v_adrsExceptionCodes As ADODB.Recordset _
                , ByRef r_bReportGenerated As Boolean _
                , ByVal v_bActiveEmployees As Boolean _
                , ByVal v_adrsAllEmployees As ADODB.Recordset _
                , ByVal v_adrsSelDistributions As ADODB.Recordset _
                ) As ActiveReport
On Error GoTo GetActualTimeWGbyDistRpt_EH

Dim lstrFilter As String
Dim llngIndex As Long

    'Initialize return
    r_bReportGenerated = False
    ResetActiveReportsError
    Set GetActualTimeWGbyDistRpt = Nothing

    If (RecordsExist(v_adrsSelectedWorkgroups)) Then
        Set m_adrsSelectedWorkgroups = v_adrsSelectedWorkgroups
        Set m_adrsAllWorkgroups = m_adrsSelectedWorkgroups.Clone
    Else
        Exit Function
    End If
    If (RecordsExist(v_adrsAllEmployees)) Then
        Set m_adrsAllEmployees = v_adrsAllEmployees
    Else
        Exit Function
    End If

    m_lReportLevel = 2

    m_dtStartDate = v_dtStartDate
    m_dtEndDate = v_dtEndDate

    Set m_adrsExceptionCodes = v_adrsExceptionCodes

    'Load XML file into the report object
    Set rptMain = CreateObject("DDActiveReports2.ActiveReport")
    rptMain.LoadLayout "C:\TRIS Solution\TRISReports\RPX Files\rptActualTimeHistoryWG.rpx"
    'Add a reference in the rpt object to this (rptException) class
    Call rptMain.AddNamedItem("MyClass", Me)

    Set rptSubReport1 = CreateObject("DDActiveReports2.ActiveReport")
    rptSubReport1.LoadLayout "C:\TRIS Solution\TRISReports\RPX Files\rptActualTimeTotalsSub.rpx"
    'Add a reference in the rpt object to this (rptException) class
    Call rptSubReport1.AddNamedItem("MyClass", Me)

    Set rptSubReport2 = CreateObject("DDActiveReports2.ActiveReport")
    rptSubReport2.LoadLayout "C:\TRIS Solution\TRISReports\RPX Files\rptActualTimeTotalsSub.rpx"
    'Add a reference in the rpt object to this (rptException) class
    Call rptSubReport2.AddNamedItem("MyClass", Me)

    rptMain.Sections("GroupFooter2").Controls("subRptEmpTotals").object = _
        rptSubReport1
    rptMain.Sections("GroupFooter1").Controls("subRptWGTotals").object = _
        rptSubReport2
    rptMain.Sections("PageHeader").Controls("txtHdrRunDate").Text = _
        Now()
    rptMain.Sections("PageHeader").Controls("txtHdrStartDate").Text = _
        CStr(Format(v_dtStartDate, "mm/dd/yyyy"))
    rptMain.Sections("PageHeader").Controls("txtHdrEndDate").Text = _
        CStr(Format(v_dtEndDate, "mm/dd/yyyy"))

    Set m_adrsReportData = GetActualTimeHistoryWGbyDist( _
                            m_adrsSelectedWorkgroups _
                            , v_dtStartDate _
                            , v_dtEndDate _
                            , v_adrsSelDistributions)
    'If (RecordsExist(m_adrsReportData)) Then
    'place filter here if distributions are supplied
    '    lstrFilter = ""
    '    If RecordsExist(v_adrsSelDistributions) Then
    '        With v_adrsSelDistributions
    '            .MoveFirst
    '            lstrFilter = _
    '            "DISTRIBUTION_ID LIKE '%" & _
    '            Trim(.Fields("DISTRIBUTION_ID")) & "%'"
    '            If Not .EOF Then
    '                .MoveNext
    '                For llngIndex = 2 To .RecordCount
    '                        lstrFilter = lstrFilter & " OR" & _
    '                        " DISTRIBUTION_ID LIKE '%" & _
    '                        Trim(.Fields("DISTRIBUTION_ID")) & "%'"
    '                    .MoveNext
    '                Next llngIndex
    '            End If
    '        End With
        
    '        If lstrFilter <> "" Then
    '            m_adrsReportData.Filter = stdmTRISCommon.MDAC25FilterFix(lstrFilter)
    '        End If
    '
    '        If m_adrsReportData.EOF Then
    '           Exit Function
    '        End If
    '    End If
        
    If RecordsExist(m_adrsReportData) Then
        r_bReportGenerated = True

        m_adrsReportData.Sort = "WG_NUM, EXCEPTION_CD, CHARGE_DATE"
        m_adrsReportData.MoveFirst

        CalculateWGTotals

        m_adrsReportData.Sort = "WG_NUM, EMP_NUM, EXCEPTION_CD, CHARGE_DATE"
        m_adrsReportData.MoveFirst

        CalculateEmpWGTotals

        m_adrsReportData.Sort = "WG_NUM, EMP_NUM, CHARGE_DATE, DATE_WORKED, START_TIME, END_TIME, SEQ_NUM, HOURS"
        m_adrsReportData.MoveFirst

        'Return the report object to the calling application
        rptMain.Run
        
        '//check for any errors from the active reports script or the vb6 functions being called back
        If m_ActiveReports_ErrNumber <> 0 Then
            Err.Raise m_ActiveReports_ErrNumber, m_ActiveReports_ErrSource, m_ActiveReports_ErrDescription
        End If
        
        Set GetActualTimeWGbyDistRpt = rptMain
    End If

    '//-----clean up public properties used by ActiveReports-----//
    DestroyState

    Exit Function
GetActualTimeWGbyDistRpt_EH:
    DestroyState
    
    Select Case Err.Number
    Case m_ActiveReports_ErrNumber
        TrsRaiseError MODULENAME, "GetActualTimeWGbyDistRpt", Err.Number, , , _
            DescArgsFrom("[OrigErrDescription]===>" & m_ActiveReports_OrigErrDescription & "<===", "[OrigErrSource]===>" & m_ActiveReports_OrigErrSource & "<==="), True
    Case Else
        TrsRaiseError MODULENAME, "GetActualTimeWGbyDistRpt", , , , , True
    End Select
End Function

Private Function ResetActiveReportsError()
    m_ActiveReports_ErrNumber = 0
    m_ActiveReports_ErrDescription = ""
    m_ActiveReports_ErrSource = ""
    m_ActiveReports_OrigErrDescription = ""
    m_ActiveReports_OrigErrSource = ""
End Function

'//called from Active Reports scripts
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
        m_ActiveReports_ErrDescription = Description
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
        m_ActiveReports_ErrDescription = Description
        m_ActiveReports_ErrSource = Source
    End If
End Sub

