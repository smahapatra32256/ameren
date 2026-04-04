VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "rptDailyRoster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Project:       TRIS Conversion
'Date:          Oct 2004
'Author:        Lisa Bertke
'Description:

Option Explicit

Public mdtRequestDate As Date
Public mdtStartDate As Date
Public mdtEndDate As Date
Public mblnPrintQualsSkills As Boolean
Public mblnSortbyLine As Boolean
Public mblnIneligOnPage As Boolean
Public mblnNewPage As Boolean
Public mlngPageNum As Long
Public mblnPrintNewPage As Boolean
Dim marrWorkGroups() As Long
Dim marrWorkGroupDesc() As String
Dim marrDepartments() As String
Dim marrDepartmentDesc() As String
Dim marrEmployees() As Long
Dim mlngWGCount As Long
Dim mlngWGIndex As Long
Dim mlngEmpCount As Long
Dim mlngEmpIndex As Long
Dim mlngDeptCount As Long
Dim mlngDeptIndex As Long
Dim mblnDemandInd As Boolean
Dim mblnFirstRecord As Boolean
Dim mlngPrevEmp As Long
Dim mdtPrevDate As Date
Dim mstrPrevExcep As String
Dim mstrPrevDescription As String
Dim mlngExceptionCount As Long
Dim marrRequestDate() As Date
Dim mlngDateIndex As Long
Dim mlngDateCount As Long
    

Dim WithEvents rptMain As ActiveReport
Dim m_ActiveReports_ErrNumber As Integer
Dim m_ActiveReports_ErrDescription As String
Dim m_ActiveReports_ErrSource As String
Dim m_ActiveReports_OrigErrDescription As String
Dim m_ActiveReports_OrigErrSource As String

Dim mobjCodeTables As dbCodeTables.cdCodeTables
Dim mobjEmployee As dbEmployee.cdEmployee
Dim mobjTimeCard As dbTimeCard.cdTimeCard
Dim mobjRules As dbRules.cdRules

Dim madrsDummy As ADODB.Recordset
Dim madrsRosterData As ADODB.Recordset
Dim madrsFormatRptData As ADODB.Recordset
Dim madrsEmpsinWG As ADODB.Recordset
Dim madrsEmpsinDept As ADODB.Recordset
Dim madrsQuals As ADODB.Recordset
Dim madrsSkills As ADODB.Recordset
Dim madrsEmployeeInelig As ADODB.Recordset

Const mconFIELDATTR = adFldUpdatable Or adFldIsNullable
Private Const MODULENAME As String = "rptDailyRoster"

Public Function GetDailyRosterRptWG( _
        ByVal vdtRequestDate As Date, _
        ByVal vblnPrintQualsSkills As Boolean, _
        ByVal vblnSortbyLine As Boolean, _
        ByVal vadrsSelectedWorkgroups As ADODB.Recordset, _
        ByRef rbReportGenerated As Boolean) As ActiveReport
                                                                                                                                                                        
End Function

Public Function GetDailyRosterRptWG2( _
        ByVal vdtStartDate As Date, _
        ByVal vdtEndDate As Date, _
        ByVal vblnPrintQualsSkills As Boolean, _
        ByVal vblnSortbyLine As Boolean, _
        ByVal vadrsSelectedWorkgroups As ADODB.Recordset, _
        ByRef rbReportGenerated As Boolean) As ActiveReport
                                                                                                      
On Error GoTo GetDailyRosterRptWG2_EH

    Dim ladrsWG As ADODB.Recordset
    Dim lobjWG As dbWorkGroup.cdWorkGroup
    Dim lobjEmployee As dbEmployee.cdEmployee
    Dim lobjbusEmployee As busEmployee.cbsEmployee
    Dim llngWG(0) As Long
    Dim llngPrevEmp As Long
    Dim llngNumofEmps As Long
    Dim llngdatecount As Long
    
    mdtStartDate = vdtStartDate
    mdtEndDate = vdtEndDate
    mblnPrintQualsSkills = vblnPrintQualsSkills
    mblnSortbyLine = vblnSortbyLine
    
    rbReportGenerated = False
    ResetActiveReportsError
    Set GetDailyRosterRptWG2 = Nothing
    
    'Load XML File into the report object
    Set rptMain = CreateObject("DDActiveReports2.ActiveReport")
    rptMain.LoadLayout "C:\TRIS Solution\TRISReports\RPX Files\rptDailyRosterWG.rpx"
    
    'Add a reference in the rpt object to this (rptDailyRoster) class
    Call rptMain.AddNamedItem("MyClass", Me)
    
    rptMain.Sections("PageHeader").Controls("txtPgHdrRunDate").Text = Date
    rptMain.Sections("PageHeader").Controls("txtPgHdrRunTime").Text = Time
        
    If mADOExtensions.RecordsExist(vadrsSelectedWorkgroups) Then
        With vadrsSelectedWorkgroups
            mlngWGCount = .RecordCount
            .MoveFirst
            ReDim marrWorkGroups(0 To mlngWGCount - 1)
            ReDim marrWorkGroupDesc(0 To mlngWGCount - 1)
            For mlngWGIndex = 0 To mlngWGCount - 1
                marrWorkGroups(mlngWGIndex) = !WG_NUM
                marrWorkGroupDesc(mlngWGIndex) = !WG_DESC
                .MoveNext
            Next mlngWGIndex
        End With
        mlngDateCount = DateDiff("d", Format(mdtStartDate, "mm/dd/yyyy"), Format(mdtEndDate, "mm/dd/yyyy"))
        'ReDim marrRequestDate(0 To mlngDateCount)
        'For mlngDateIndex = 0 To mlngDateCount
        '    marrRequestDate(mlngDateIndex) = DateAdd("d", mlngDateIndex, mdtStartDate)
        'Next mlngDateIndex

        Set madrsFormatRptData = New ADODB.Recordset
        With madrsFormatRptData
            .Fields.Append "WG_NUM", adInteger, , mconFIELDATTR
            .Fields.Append "WG_DESC", adChar, 30, mconFIELDATTR
            .Fields.Append "EMP_NUM", adInteger, , mconFIELDATTR
            .Fields.Append "EMP_NAME", adChar, 30, mconFIELDATTR
            .Fields.Append "SECONDARY_IND", adVarChar, 1, mconFIELDATTR
            .Fields.Append "QUALS", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SKILLS", adVarChar, 256, mconFIELDATTR
            .Fields.Append "LINE_NUM", adInteger, , mconFIELDATTR
            .Fields.Append "CHARGE_DATE", adDate, , mconFIELDATTR
            .Fields.Append "REG_SHIFT", adChar, 1, mconFIELDATTR
            .Fields.Append "REG_SHIFT_START", adChar, 5, mconFIELDATTR
            .Fields.Append "REG_SHIFT_END", adChar, 5, mconFIELDATTR
            .Fields.Append "SORT_CODE", adInteger, , mconFIELDATTR
            .Fields.Append "TRADE_RAS", adChar, 3, mconFIELDATTR
            .Fields.Append "TRADE_DOUBLE", adChar, 1, mconFIELDATTR
            .Fields.Append "CANVASSED_OT_SHIFT1", adChar, 1, mconFIELDATTR
            .Fields.Append "CANVASSED_OT_SHIFT2", adChar, 1, mconFIELDATTR
            .Fields.Append "CANVASSED_OT_SHIFT3", adChar, 1, mconFIELDATTR
            .Fields.Append "CANVASSED_OT_SHIFT4", adChar, 1, mconFIELDATTR
            .Fields.Append "CANVASSED_OT_SHIFT5", adChar, 1, mconFIELDATTR
            .Fields.Append "CANVASSED_OT_SHIFT6", adChar, 1, mconFIELDATTR
            .Fields.Append "CANVASSED_OT_SHIFT7", adChar, 1, mconFIELDATTR
            .Fields.Append "CANVASSED_OT_SHIFT8", adChar, 1, mconFIELDATTR
            .Fields.Append "CANVASSED_OT_SHIFT9", adChar, 1, mconFIELDATTR
            .Fields.Append "CANVASSED_OT_SHIFT10", adChar, 1, mconFIELDATTR
            .Fields.Append "CANVASSED_OT_SHIFT11", adChar, 1, mconFIELDATTR
            .Fields.Append "CANVASSED_OT_SHIFT12", adChar, 1, mconFIELDATTR
            .Fields.Append "CANVASSED_OT_SHIFT13", adChar, 1, mconFIELDATTR
            .Fields.Append "CANVASSED_OT_SHIFT14", adChar, 1, mconFIELDATTR
            .Fields.Append "CANVASSED_OT_SHIFT15", adChar, 1, mconFIELDATTR
            .Fields.Append "CANVASSED_OT_HOURS1", adChar, 5, mconFIELDATTR
            .Fields.Append "CANVASSED_OT_HOURS2", adChar, 5, mconFIELDATTR
            .Fields.Append "CANVASSED_OT_HOURS3", adChar, 5, mconFIELDATTR
            .Fields.Append "CANVASSED_OT_HOURS4", adChar, 5, mconFIELDATTR
            .Fields.Append "CANVASSED_OT_HOURS5", adChar, 5, mconFIELDATTR
            .Fields.Append "CANVASSED_OT_HOURS6", adChar, 5, mconFIELDATTR
            .Fields.Append "CANVASSED_OT_HOURS7", adChar, 5, mconFIELDATTR
            .Fields.Append "CANVASSED_OT_HOURS8", adChar, 5, mconFIELDATTR
            .Fields.Append "CANVASSED_OT_HOURS9", adChar, 5, mconFIELDATTR
            .Fields.Append "CANVASSED_OT_HOURS10", adChar, 5, mconFIELDATTR
            .Fields.Append "CANVASSED_OT_HOURS11", adChar, 5, mconFIELDATTR
            .Fields.Append "CANVASSED_OT_HOURS12", adChar, 5, mconFIELDATTR
            .Fields.Append "CANVASSED_OT_HOURS13", adChar, 5, mconFIELDATTR
            .Fields.Append "CANVASSED_OT_HOURS14", adChar, 5, mconFIELDATTR
            .Fields.Append "CANVASSED_OT_HOURS15", adChar, 5, mconFIELDATTR
            .Fields.Append "EXCEPTION1", adChar, 3, mconFIELDATTR
            .Fields.Append "EXCEPTION1_DESC", adChar, 30, mconFIELDATTR
            .Fields.Append "EXCEPTION1_START", adChar, 5, mconFIELDATTR
            .Fields.Append "EXCEPTION1_END", adChar, 5, mconFIELDATTR
            .Fields.Append "EXCEPTION2", adChar, 3, mconFIELDATTR
            .Fields.Append "EXCEPTION2_DESC", adChar, 30, mconFIELDATTR
            .Fields.Append "EXCEPTION2_START", adChar, 5, mconFIELDATTR
            .Fields.Append "EXCEPTION2_END", adChar, 5, mconFIELDATTR
            .Fields.Append "EXCEPTION3", adChar, 3, mconFIELDATTR
            .Fields.Append "EXCEPTION3_DESC", adChar, 30, mconFIELDATTR
            .Fields.Append "EXCEPTION3_START", adChar, 5, mconFIELDATTR
            .Fields.Append "EXCEPTION3_END", adChar, 5, mconFIELDATTR
            .Fields.Append "EXCEPTION4", adChar, 3, mconFIELDATTR
            .Fields.Append "EXCEPTION4_DESC", adChar, 30, mconFIELDATTR
            .Fields.Append "EXCEPTION4_START", adChar, 5, mconFIELDATTR
            .Fields.Append "EXCEPTION4_END", adChar, 5, mconFIELDATTR
            .Fields.Append "EXCEPTION5", adChar, 3, mconFIELDATTR
            .Fields.Append "EXCEPTION5_DESC", adChar, 30, mconFIELDATTR
            .Fields.Append "EXCEPTION5_START", adChar, 5, mconFIELDATTR
            .Fields.Append "EXCEPTION5_END", adChar, 5, mconFIELDATTR
            .Fields.Append "EXCEPTION6", adChar, 3, mconFIELDATTR
            .Fields.Append "EXCEPTION6_DESC", adChar, 30, mconFIELDATTR
            .Fields.Append "EXCEPTION6_START", adChar, 5, mconFIELDATTR
            .Fields.Append "EXCEPTION6_END", adChar, 5, mconFIELDATTR
            .Fields.Append "EXCEPTION7", adChar, 3, mconFIELDATTR
            .Fields.Append "EXCEPTION7_DESC", adChar, 30, mconFIELDATTR
            .Fields.Append "EXCEPTION7_START", adChar, 5, mconFIELDATTR
            .Fields.Append "EXCEPTION7_END", adChar, 5, mconFIELDATTR
            .Fields.Append "EXCEPTION8", adChar, 3, mconFIELDATTR
            .Fields.Append "EXCEPTION8_DESC", adChar, 30, mconFIELDATTR
            .Fields.Append "EXCEPTION8_START", adChar, 5, mconFIELDATTR
            .Fields.Append "EXCEPTION8_END", adChar, 5, mconFIELDATTR
            .Fields.Append "EXCEPTION9", adChar, 3, mconFIELDATTR
            .Fields.Append "EXCEPTION9_DESC", adChar, 30, mconFIELDATTR
            .Fields.Append "EXCEPTION9_START", adChar, 5, mconFIELDATTR
            .Fields.Append "EXCEPTION9_END", adChar, 5, mconFIELDATTR
            .Fields.Append "EXCEPTION10", adChar, 3, mconFIELDATTR
            .Fields.Append "EXCEPTION10_DESC", adChar, 30, mconFIELDATTR
            .Fields.Append "EXCEPTION10_START", adChar, 5, mconFIELDATTR
            .Fields.Append "EXCEPTION10_END", adChar, 5, mconFIELDATTR
            .Fields.Append "EXCEPTION11", adChar, 3, mconFIELDATTR
            .Fields.Append "EXCEPTION11_DESC", adChar, 30, mconFIELDATTR
            .Fields.Append "EXCEPTION11_START", adChar, 5, mconFIELDATTR
            .Fields.Append "EXCEPTION11_END", adChar, 5, mconFIELDATTR
            .Fields.Append "EXCEPTION12", adChar, 3, mconFIELDATTR
            .Fields.Append "EXCEPTION12_DESC", adChar, 30, mconFIELDATTR
            .Fields.Append "EXCEPTION12_START", adChar, 5, mconFIELDATTR
            .Fields.Append "EXCEPTION12_END", adChar, 5, mconFIELDATTR
            .Fields.Append "EXCEPTION13", adChar, 3, mconFIELDATTR
            .Fields.Append "EXCEPTION13_DESC", adChar, 30, mconFIELDATTR
            .Fields.Append "EXCEPTION13_START", adChar, 5, mconFIELDATTR
            .Fields.Append "EXCEPTION13_END", adChar, 5, mconFIELDATTR
            .Fields.Append "EXCEPTION14", adChar, 3, mconFIELDATTR
            .Fields.Append "EXCEPTION14_DESC", adChar, 30, mconFIELDATTR
            .Fields.Append "EXCEPTION14_START", adChar, 5, mconFIELDATTR
            .Fields.Append "EXCEPTION14_END", adChar, 5, mconFIELDATTR
            .Fields.Append "EXCEPTION15", adChar, 3, mconFIELDATTR
            .Fields.Append "EXCEPTION15_DESC", adChar, 30, mconFIELDATTR
            .Fields.Append "EXCEPTION15_START", adChar, 5, mconFIELDATTR
            .Fields.Append "EXCEPTION15_END", adChar, 5, mconFIELDATTR
            .Fields.Append "INELIGIBILITY_START1", adDate, , mconFIELDATTR
            .Fields.Append "INELIGIBILITY_START2", adDate, , mconFIELDATTR
            .Fields.Append "INELIGIBILITY_START3", adDate, , mconFIELDATTR
            .Fields.Append "INELIGIBILITY_START4", adDate, , mconFIELDATTR
            .Fields.Append "INELIGIBILITY_START5", adDate, , mconFIELDATTR
            .Fields.Append "INELIGIBILITY_START6", adDate, , mconFIELDATTR
            .Fields.Append "INELIGIBILITY_START7", adDate, , mconFIELDATTR
            .Fields.Append "INELIGIBILITY_START8", adDate, , mconFIELDATTR
            .Fields.Append "INELIGIBILITY_START9", adDate, , mconFIELDATTR
            .Fields.Append "INELIGIBILITY_START10", adDate, , mconFIELDATTR
            .Fields.Append "INELIGIBILITY_START11", adDate, , mconFIELDATTR
            .Fields.Append "INELIGIBILITY_START12", adDate, , mconFIELDATTR
            .Fields.Append "INELIGIBILITY_START13", adDate, , mconFIELDATTR
            .Fields.Append "INELIGIBILITY_START14", adDate, , mconFIELDATTR
            .Fields.Append "INELIGIBILITY_START15", adDate, , mconFIELDATTR
            .Fields.Append "INELIGIBILITY_END1", adDate, , mconFIELDATTR
            .Fields.Append "INELIGIBILITY_END2", adDate, , mconFIELDATTR
            .Fields.Append "INELIGIBILITY_END3", adDate, , mconFIELDATTR
            .Fields.Append "INELIGIBILITY_END4", adDate, , mconFIELDATTR
            .Fields.Append "INELIGIBILITY_END5", adDate, , mconFIELDATTR
            .Fields.Append "INELIGIBILITY_END6", adDate, , mconFIELDATTR
            .Fields.Append "INELIGIBILITY_END7", adDate, , mconFIELDATTR
            .Fields.Append "INELIGIBILITY_END8", adDate, , mconFIELDATTR
            .Fields.Append "INELIGIBILITY_END9", adDate, , mconFIELDATTR
            .Fields.Append "INELIGIBILITY_END10", adDate, , mconFIELDATTR
            .Fields.Append "INELIGIBILITY_END11", adDate, , mconFIELDATTR
            .Fields.Append "INELIGIBILITY_END12", adDate, , mconFIELDATTR
            .Fields.Append "INELIGIBILITY_END13", adDate, , mconFIELDATTR
            .Fields.Append "INELIGIBILITY_END14", adDate, , mconFIELDATTR
            .Fields.Append "INELIGIBILITY_END15", adDate, , mconFIELDATTR
            .CursorLocation = adUseClient
            .CursorType = adOpenStatic
            .LockType = adLockBatchOptimistic
            .Open
        End With
    Else
        Exit Function
    End If
    
    For mlngWGIndex = 0 To mlngWGCount - 1
        Set lobjWG = CreateObject("dbWorkGroup.cdWorkGroup")
        Set ladrsWG = lobjWG.GetWorkGroupAttributes(marrWorkGroups(mlngWGIndex))
        mblnDemandInd = False
        If mADOExtensions.RecordsExist(ladrsWG) Then
            If ladrsWG!DEMAND_IND = "Y" Then
                mblnDemandInd = True
            End If
        End If
        
        'For mlngDateIndex = 0 To mlngDateCount
            llngWG(0) = marrWorkGroups(mlngWGIndex)
            Set lobjWG = CreateObject("dbWorkGroup.cdWorkGroup")
            Set madrsEmpsinWG = Nothing
            'Set madrsEmpsinWG = lobjWG.GetEmployeesForDailyRosterRpt2(llngWG(0), _
                                                        marrRequestDate(mlngDateIndex))
            Set madrsEmpsinWG = lobjWG.GetEmployeesForDailyRosterRpt2(llngWG(0), _
                                                        vdtStartDate, vdtEndDate)
            Set lobjWG = Nothing
            
            If mADOExtensions.RecordsExist(madrsEmpsinWG) Then
                With madrsEmpsinWG
                    mlngEmpCount = .RecordCount
                    .MoveFirst
                    llngPrevEmp = 0
                    llngNumofEmps = 0
                    mlngEmpIndex = 0
                    ReDim marrEmployees(0 To mlngEmpCount - 1)
                    While Not .EOF
                        If !EMP_NUM <> llngPrevEmp Then
                            marrEmployees(mlngEmpIndex) = !EMP_NUM
                            mlngEmpIndex = mlngEmpIndex + 1
                            llngNumofEmps = llngNumofEmps + 1
                        End If
                        llngPrevEmp = !EMP_NUM
                        .MoveNext
                    Wend
                    ReDim Preserve marrEmployees(0 To llngNumofEmps - 1)
                End With
            
                Set lobjbusEmployee = CreateObject("busEmployee.cbsEmployee")
                Set madrsEmployeeInelig = lobjbusEmployee.GetEmployeeIneligibilities(madrsEmpsinWG)
                Set lobjbusEmployee = Nothing
                
                Set lobjEmployee = CreateObject("dbEmployee.cdEmployee")
                
                If mblnPrintQualsSkills Then
                    Set madrsQuals = lobjEmployee.GetEmployeeQualifications(marrEmployees)
                    Set madrsSkills = lobjEmployee.GetEmployeeSkills(marrEmployees)
                End If
                
                Set lobjEmployee = Nothing
                
                'Set madrsRosterData = GetDailyRoster(llngWG(0), _
                                                    marrRequestDate(mlngDateIndex))
                Set madrsRosterData = GetDailyRoster2(llngWG(0), vdtStartDate, vdtEndDate)
                                                                                                        
                mstrPrevExcep = ""
                mstrPrevDescription = ""
                                                                
                If mADOExtensions.RecordsExist(madrsRosterData) Then
                    madrsRosterData.Filter = "SYSTEM_EXC_CD <> 'SLT'"
                    madrsRosterData.MoveFirst
                    If mADOExtensions.RecordsExist(madrsRosterData) Then
                        Set mobjCodeTables = CreateObject("dbCodeTables.cdCodeTables")
                        Set mobjEmployee = CreateObject("dbEmployee.cdEmployee")
                        Set mobjTimeCard = CreateObject("dbTimecard.cdTimecard")
                        Set mobjRules = CreateObject("dbRules.cdRules")
                        rbReportGenerated = True
                        madrsRosterData.Sort = "EMP_NUM, CHARGE_DATE, DATE_WORKED, START_TIME, END_TIME, SCHED_SRC_CD"
                        madrsRosterData.MoveFirst
                        mblnFirstRecord = True
                        mlngPrevEmp = 0
                        mdtPrevDate = "12/31/9999"
                        With madrsRosterData
                            Do While .EOF = False
                                PopulateWG_Record
                                If Not mblnFirstRecord Then
                                    madrsFormatRptData.Update
                                End If
                                mlngPrevEmp = !EMP_NUM
                                mdtPrevDate = !CHARGE_DATE
                                .MoveNext
                            Loop
                        End With
                        Set mobjCodeTables = Nothing
                        Set mobjEmployee = Nothing
                        Set mobjTimeCard = Nothing
                        Set mobjRules = Nothing
                    End If
                End If
            End If
        'Next mlngDateIndex
    Next mlngWGIndex
    
    'Return the report object to the calling application
    If rbReportGenerated = True Then
        rptMain.Run
        
        '//check for any errors from the active reports script or the vb6 functions being called back
        If m_ActiveReports_ErrNumber <> 0 Then
            Err.Raise m_ActiveReports_ErrNumber, m_ActiveReports_ErrSource, m_ActiveReports_ErrDescription
        End If
        Set GetDailyRosterRptWG2 = rptMain
    Else
        Set GetDailyRosterRptWG2 = Nothing
    End If
    
    Set lobjEmployee = Nothing
    Set lobjWG = Nothing
    Set ladrsWG = Nothing
    
    '//-----clean up public properties used by ActiveReports-----//
    DestroyState
    
    Exit Function
GetDailyRosterRptWG2_EH:
    Set lobjEmployee = Nothing
    Set lobjWG = Nothing
    Set ladrsWG = Nothing
    
    DestroyState
                                                                                                                                                                    Select Case Err.Number
    Case m_ActiveReports_ErrNumber
        TrsRaiseError MODULENAME, "GetDailyRosterRptWG2", Err.Number, , , _
            DescArgsFrom("[OrigErrDescription]===>" & m_ActiveReports_OrigErrDescription & "<===", "[OrigErrSource]===>" & m_ActiveReports_OrigErrSource & "<==="), True
    Case Else
        TrsRaiseError MODULENAME, "GetDailyRosterRptWG2", , , , , True
    End Select
End Function

Public Function GetDailyRosterRptDept( _
        ByVal vdtRequestDate As Date, _
        ByVal vblnPrintQualsSkills As Boolean, _
        ByVal vblnSortbyLine As Boolean, _
        ByVal vadrsSelectedDepartments As ADODB.Recordset, _
        ByRef rbReportGenerated As Boolean) As ActiveReport

End Function
Public Function GetDailyRosterRptDept2( _
        ByVal vdtStartDate As Date, _
        ByVal vdtEndDate As Date, _
        ByVal vblnPrintQualsSkills As Boolean, _
        ByVal vblnSortbyLine As Boolean, _
        ByVal vadrsSelectedDepartments As ADODB.Recordset, _
        ByRef rbReportGenerated As Boolean) As ActiveReport
        
On Error GoTo GetDailyRosterRptDept2_EH

    Dim ladrsWG As ADODB.Recordset
    Dim lobjWG As dbWorkGroup.cdWorkGroup
    Dim lobjEmployee As dbEmployee.cdEmployee
    Dim lobjbusEmployee As busEmployee.cbsEmployee
    Dim lobjDepartment As dbDepartment.cdDepartment
    Dim llngWG(0) As Long
    Dim llngPrevEmp As Long
    Dim llngPrevWG As Long
    Dim ldtPrevDate As Date
    Dim llngNumofEmps As Long
    
    mdtStartDate = vdtStartDate
    mdtEndDate = vdtEndDate
    mblnPrintQualsSkills = vblnPrintQualsSkills
    mblnSortbyLine = vblnSortbyLine
    
    rbReportGenerated = False
    ResetActiveReportsError
    Set GetDailyRosterRptDept2 = Nothing
    
    'Load XML File into the report object
    Set rptMain = CreateObject("DDActiveReports2.ActiveReport")
    rptMain.LoadLayout "C:\TRIS Solution\TRISReports\RPX Files\rptDailyRosterDept.rpx"
    
    'Add a reference in the rpt object to this (rptDailyRoster) class
    Call rptMain.AddNamedItem("MyClass", Me)
    
    rptMain.Sections("PageHeader").Controls("txtPgHdrRunDate").Text = Date
    rptMain.Sections("PageHeader").Controls("txtPgHdrRunTime").Text = Time
    
    If mADOExtensions.RecordsExist(vadrsSelectedDepartments) Then
        With vadrsSelectedDepartments
            mlngDeptCount = .RecordCount
            .MoveFirst
            ReDim marrDepartments(0 To mlngDeptCount - 1)
            ReDim marrDepartmentDesc(0 To mlngDeptCount - 1)
            For mlngDeptIndex = 0 To mlngDeptCount - 1
                marrDepartments(mlngDeptIndex) = !DEPT_NUM
                marrDepartmentDesc(mlngDeptIndex) = !DEPT_NAME
                .MoveNext
            Next mlngDeptIndex
        End With
        mlngDateCount = DateDiff("d", Format(mdtStartDate, "mm/dd/yyyy"), Format(mdtEndDate, "mm/dd/yyyy"))
        'ReDim marrRequestDate(0 To mlngDateCount)
        'For mlngDateIndex = 0 To mlngDateCount
        '    marrRequestDate(mlngDateIndex) = DateAdd("d", mlngDateIndex, mdtStartDate)
        'Next mlngDateIndex

        Set madrsFormatRptData = New ADODB.Recordset
        With madrsFormatRptData
            .Fields.Append "DEPT_NUM", adChar, 3, mconFIELDATTR
            .Fields.Append "DEPT_DESC", adChar, 30, mconFIELDATTR
            .Fields.Append "WG_NUM", adInteger, , mconFIELDATTR
            .Fields.Append "WG_DESC", adChar, 30, mconFIELDATTR
            .Fields.Append "EMP_NUM", adInteger, , mconFIELDATTR
            .Fields.Append "EMP_NAME", adChar, 30, mconFIELDATTR
            .Fields.Append "SECONDARY_IND", adVarChar, 1, mconFIELDATTR
            .Fields.Append "QUALS", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SKILLS", adVarChar, 256, mconFIELDATTR
            .Fields.Append "LINE_NUM", adInteger, , mconFIELDATTR
            .Fields.Append "CHARGE_DATE", adDate, , mconFIELDATTR
            .Fields.Append "REG_SHIFT", adChar, 1, mconFIELDATTR
            .Fields.Append "REG_SHIFT_START", adChar, 5, mconFIELDATTR
            .Fields.Append "REG_SHIFT_END", adChar, 5, mconFIELDATTR
            .Fields.Append "SORT_CODE", adInteger, , mconFIELDATTR
            .Fields.Append "TRADE_RAS", adChar, 3, mconFIELDATTR
            .Fields.Append "TRADE_DOUBLE", adChar, 1, mconFIELDATTR
            .Fields.Append "CANVASSED_OT_SHIFT1", adChar, 1, mconFIELDATTR
            .Fields.Append "CANVASSED_OT_SHIFT2", adChar, 1, mconFIELDATTR
            .Fields.Append "CANVASSED_OT_SHIFT3", adChar, 1, mconFIELDATTR
            .Fields.Append "CANVASSED_OT_SHIFT4", adChar, 1, mconFIELDATTR
            .Fields.Append "CANVASSED_OT_SHIFT5", adChar, 1, mconFIELDATTR
            .Fields.Append "CANVASSED_OT_SHIFT6", adChar, 1, mconFIELDATTR
            .Fields.Append "CANVASSED_OT_SHIFT7", adChar, 1, mconFIELDATTR
            .Fields.Append "CANVASSED_OT_SHIFT8", adChar, 1, mconFIELDATTR
            .Fields.Append "CANVASSED_OT_SHIFT9", adChar, 1, mconFIELDATTR
            .Fields.Append "CANVASSED_OT_SHIFT10", adChar, 1, mconFIELDATTR
            .Fields.Append "CANVASSED_OT_SHIFT11", adChar, 1, mconFIELDATTR
            .Fields.Append "CANVASSED_OT_SHIFT12", adChar, 1, mconFIELDATTR
            .Fields.Append "CANVASSED_OT_SHIFT13", adChar, 1, mconFIELDATTR
            .Fields.Append "CANVASSED_OT_SHIFT14", adChar, 1, mconFIELDATTR
            .Fields.Append "CANVASSED_OT_SHIFT15", adChar, 1, mconFIELDATTR
            .Fields.Append "CANVASSED_OT_HOURS1", adChar, 5, mconFIELDATTR
            .Fields.Append "CANVASSED_OT_HOURS2", adChar, 5, mconFIELDATTR
            .Fields.Append "CANVASSED_OT_HOURS3", adChar, 5, mconFIELDATTR
            .Fields.Append "CANVASSED_OT_HOURS4", adChar, 5, mconFIELDATTR
            .Fields.Append "CANVASSED_OT_HOURS5", adChar, 5, mconFIELDATTR
            .Fields.Append "CANVASSED_OT_HOURS6", adChar, 5, mconFIELDATTR
            .Fields.Append "CANVASSED_OT_HOURS7", adChar, 5, mconFIELDATTR
            .Fields.Append "CANVASSED_OT_HOURS8", adChar, 5, mconFIELDATTR
            .Fields.Append "CANVASSED_OT_HOURS9", adChar, 5, mconFIELDATTR
            .Fields.Append "CANVASSED_OT_HOURS10", adChar, 5, mconFIELDATTR
            .Fields.Append "CANVASSED_OT_HOURS11", adChar, 5, mconFIELDATTR
            .Fields.Append "CANVASSED_OT_HOURS12", adChar, 5, mconFIELDATTR
            .Fields.Append "CANVASSED_OT_HOURS13", adChar, 5, mconFIELDATTR
            .Fields.Append "CANVASSED_OT_HOURS14", adChar, 5, mconFIELDATTR
            .Fields.Append "CANVASSED_OT_HOURS15", adChar, 5, mconFIELDATTR
            .Fields.Append "EXCEPTION1", adChar, 3, mconFIELDATTR
            .Fields.Append "EXCEPTION1_DESC", adChar, 30, mconFIELDATTR
            .Fields.Append "EXCEPTION1_START", adChar, 5, mconFIELDATTR
            .Fields.Append "EXCEPTION1_END", adChar, 5, mconFIELDATTR
            .Fields.Append "EXCEPTION2", adChar, 3, mconFIELDATTR
            .Fields.Append "EXCEPTION2_DESC", adChar, 30, mconFIELDATTR
            .Fields.Append "EXCEPTION2_START", adChar, 5, mconFIELDATTR
            .Fields.Append "EXCEPTION2_END", adChar, 5, mconFIELDATTR
            .Fields.Append "EXCEPTION3", adChar, 3, mconFIELDATTR
            .Fields.Append "EXCEPTION3_DESC", adChar, 30, mconFIELDATTR
            .Fields.Append "EXCEPTION3_START", adChar, 5, mconFIELDATTR
            .Fields.Append "EXCEPTION3_END", adChar, 5, mconFIELDATTR
            .Fields.Append "EXCEPTION4", adChar, 3, mconFIELDATTR
            .Fields.Append "EXCEPTION4_DESC", adChar, 30, mconFIELDATTR
            .Fields.Append "EXCEPTION4_START", adChar, 5, mconFIELDATTR
            .Fields.Append "EXCEPTION4_END", adChar, 5, mconFIELDATTR
            .Fields.Append "EXCEPTION5", adChar, 3, mconFIELDATTR
            .Fields.Append "EXCEPTION5_DESC", adChar, 30, mconFIELDATTR
            .Fields.Append "EXCEPTION5_START", adChar, 5, mconFIELDATTR
            .Fields.Append "EXCEPTION5_END", adChar, 5, mconFIELDATTR
            .Fields.Append "EXCEPTION6", adChar, 3, mconFIELDATTR
            .Fields.Append "EXCEPTION6_DESC", adChar, 30, mconFIELDATTR
            .Fields.Append "EXCEPTION6_START", adChar, 5, mconFIELDATTR
            .Fields.Append "EXCEPTION6_END", adChar, 5, mconFIELDATTR
            .Fields.Append "EXCEPTION7", adChar, 3, mconFIELDATTR
            .Fields.Append "EXCEPTION7_DESC", adChar, 30, mconFIELDATTR
            .Fields.Append "EXCEPTION7_START", adChar, 5, mconFIELDATTR
            .Fields.Append "EXCEPTION7_END", adChar, 5, mconFIELDATTR
            .Fields.Append "EXCEPTION8", adChar, 3, mconFIELDATTR
            .Fields.Append "EXCEPTION8_DESC", adChar, 30, mconFIELDATTR
            .Fields.Append "EXCEPTION8_START", adChar, 5, mconFIELDATTR
            .Fields.Append "EXCEPTION8_END", adChar, 5, mconFIELDATTR
            .Fields.Append "EXCEPTION9", adChar, 3, mconFIELDATTR
            .Fields.Append "EXCEPTION9_DESC", adChar, 30, mconFIELDATTR
            .Fields.Append "EXCEPTION9_START", adChar, 5, mconFIELDATTR
            .Fields.Append "EXCEPTION9_END", adChar, 5, mconFIELDATTR
            .Fields.Append "EXCEPTION10", adChar, 3, mconFIELDATTR
            .Fields.Append "EXCEPTION10_DESC", adChar, 30, mconFIELDATTR
            .Fields.Append "EXCEPTION10_START", adChar, 5, mconFIELDATTR
            .Fields.Append "EXCEPTION10_END", adChar, 5, mconFIELDATTR
            .Fields.Append "EXCEPTION11", adChar, 3, mconFIELDATTR
            .Fields.Append "EXCEPTION11_DESC", adChar, 30, mconFIELDATTR
            .Fields.Append "EXCEPTION11_START", adChar, 5, mconFIELDATTR
            .Fields.Append "EXCEPTION11_END", adChar, 5, mconFIELDATTR
            .Fields.Append "EXCEPTION12", adChar, 3, mconFIELDATTR
            .Fields.Append "EXCEPTION12_DESC", adChar, 30, mconFIELDATTR
            .Fields.Append "EXCEPTION12_START", adChar, 5, mconFIELDATTR
            .Fields.Append "EXCEPTION12_END", adChar, 5, mconFIELDATTR
            .Fields.Append "EXCEPTION13", adChar, 3, mconFIELDATTR
            .Fields.Append "EXCEPTION13_DESC", adChar, 30, mconFIELDATTR
            .Fields.Append "EXCEPTION13_START", adChar, 5, mconFIELDATTR
            .Fields.Append "EXCEPTION13_END", adChar, 5, mconFIELDATTR
            .Fields.Append "EXCEPTION14", adChar, 3, mconFIELDATTR
            .Fields.Append "EXCEPTION14_DESC", adChar, 30, mconFIELDATTR
            .Fields.Append "EXCEPTION14_START", adChar, 5, mconFIELDATTR
            .Fields.Append "EXCEPTION14_END", adChar, 5, mconFIELDATTR
            .Fields.Append "EXCEPTION15", adChar, 3, mconFIELDATTR
            .Fields.Append "EXCEPTION15_DESC", adChar, 30, mconFIELDATTR
            .Fields.Append "EXCEPTION15_START", adChar, 5, mconFIELDATTR
            .Fields.Append "EXCEPTION15_END", adChar, 5, mconFIELDATTR
            .Fields.Append "INELIGIBILITY_START1", adDate, , mconFIELDATTR
            .Fields.Append "INELIGIBILITY_START2", adDate, , mconFIELDATTR
            .Fields.Append "INELIGIBILITY_START3", adDate, , mconFIELDATTR
            .Fields.Append "INELIGIBILITY_START4", adDate, , mconFIELDATTR
            .Fields.Append "INELIGIBILITY_START5", adDate, , mconFIELDATTR
            .Fields.Append "INELIGIBILITY_START6", adDate, , mconFIELDATTR
            .Fields.Append "INELIGIBILITY_START7", adDate, , mconFIELDATTR
            .Fields.Append "INELIGIBILITY_START8", adDate, , mconFIELDATTR
            .Fields.Append "INELIGIBILITY_START9", adDate, , mconFIELDATTR
            .Fields.Append "INELIGIBILITY_START10", adDate, , mconFIELDATTR
            .Fields.Append "INELIGIBILITY_START11", adDate, , mconFIELDATTR
            .Fields.Append "INELIGIBILITY_START12", adDate, , mconFIELDATTR
            .Fields.Append "INELIGIBILITY_START13", adDate, , mconFIELDATTR
            .Fields.Append "INELIGIBILITY_START14", adDate, , mconFIELDATTR
            .Fields.Append "INELIGIBILITY_START15", adDate, , mconFIELDATTR
            .Fields.Append "INELIGIBILITY_END1", adDate, , mconFIELDATTR
            .Fields.Append "INELIGIBILITY_END2", adDate, , mconFIELDATTR
            .Fields.Append "INELIGIBILITY_END3", adDate, , mconFIELDATTR
            .Fields.Append "INELIGIBILITY_END4", adDate, , mconFIELDATTR
            .Fields.Append "INELIGIBILITY_END5", adDate, , mconFIELDATTR
            .Fields.Append "INELIGIBILITY_END6", adDate, , mconFIELDATTR
            .Fields.Append "INELIGIBILITY_END7", adDate, , mconFIELDATTR
            .Fields.Append "INELIGIBILITY_END8", adDate, , mconFIELDATTR
            .Fields.Append "INELIGIBILITY_END9", adDate, , mconFIELDATTR
            .Fields.Append "INELIGIBILITY_END10", adDate, , mconFIELDATTR
            .Fields.Append "INELIGIBILITY_END11", adDate, , mconFIELDATTR
            .Fields.Append "INELIGIBILITY_END12", adDate, , mconFIELDATTR
            .Fields.Append "INELIGIBILITY_END13", adDate, , mconFIELDATTR
            .Fields.Append "INELIGIBILITY_END14", adDate, , mconFIELDATTR
            .Fields.Append "INELIGIBILITY_END15", adDate, , mconFIELDATTR
            .CursorLocation = adUseClient
            .CursorType = adOpenStatic
            .LockType = adLockBatchOptimistic
            .Open
        End With
    Else
        Exit Function
    End If
    
    For mlngDeptIndex = 0 To mlngDeptCount - 1
        
        'For mlngDateIndex = 0 To mlngDateCount
        
            Set lobjDepartment = CreateObject("dbDepartment.cdDepartment")
            'Set madrsEmpsinDept = lobjDepartment.GetEmpsinDept( _
                                        marrDepartments(mlngDeptIndex), _
                                        marrRequestDate(mlngDateIndex))
            Set madrsEmpsinDept = lobjDepartment.GetEmpsinDept2( _
                                        marrDepartments(mlngDeptIndex), _
                                        vdtStartDate, _
                                        vdtEndDate)
            Set lobjDepartment = Nothing
            
            If mADOExtensions.RecordsExist(madrsEmpsinDept) Then
                With madrsEmpsinDept
                    mlngEmpCount = .RecordCount
                    .MoveFirst
                    llngPrevEmp = 0
                    llngPrevWG = 0
                    ldtPrevDate = "12/31/9999"
                    llngNumofEmps = 0
                    mlngEmpIndex = 0
                    ReDim marrEmployees(0 To mlngEmpCount - 1)
                    ReDim marrWorkGroups(0 To mlngEmpCount - 1)
                    ReDim marrWorkGroupDesc(0 To mlngEmpCount - 1)
                    While Not .EOF
                        If !EMP_NUM <> llngPrevEmp Or !WG_NUM <> llngPrevWG Then
                            marrEmployees(mlngEmpIndex) = !EMP_NUM
                            marrWorkGroups(mlngEmpIndex) = !WG_NUM
                            marrWorkGroupDesc(mlngEmpIndex) = !WG_DESC
                            mlngEmpIndex = mlngEmpIndex + 1
                            llngNumofEmps = llngNumofEmps + 1
                        End If
                        llngPrevEmp = !EMP_NUM
                        llngPrevWG = !WG_NUM
                        .MoveNext
                    Wend
                    ReDim Preserve marrEmployees(0 To llngNumofEmps - 1)
                    ReDim Preserve marrWorkGroups(0 To llngNumofEmps - 1)
                    ReDim Preserve marrWorkGroupDesc(0 To llngNumofEmps - 1)
                End With
            
                Set lobjbusEmployee = CreateObject("busEmployee.cbsEmployee")
                Set madrsEmployeeInelig = lobjbusEmployee.GetEmployeeIneligibilities(madrsEmpsinDept)
                Set lobjbusEmployee = Nothing
                
                Set lobjEmployee = CreateObject("dbEmployee.cdEmployee")
                
                If mblnPrintQualsSkills Then
                    Set madrsQuals = lobjEmployee.GetEmployeeQualifications(marrEmployees)
                    Set madrsSkills = lobjEmployee.GetEmployeeSkills(marrEmployees)
                End If
                
                Set lobjEmployee = Nothing
                
                'Set madrsRosterData = GetDailyRoster(marrEmployees, _
                '                                    vdtRequestDate)
                                                    
                'madrsRosterData.Filter = "SYSTEM_EXC_CD <> 'SLT'"
                llngPrevWG = 0
                mstrPrevExcep = ""
                mstrPrevDescription = ""
                
                For mlngEmpIndex = 0 To llngNumofEmps - 1
                    With madrsRosterData
                        If marrWorkGroups(mlngEmpIndex) <> llngPrevWG Then
                            llngWG(0) = marrWorkGroups(mlngEmpIndex)
                            'Set madrsRosterData = GetDailyRoster2(llngWG(0), _
                                                        marrRequestDate(mlngDateIndex))
                            Set madrsRosterData = GetDailyRoster2(llngWG(0), _
                                                        vdtStartDate, _
                                                        vdtEndDate)
                        End If
                        If mADOExtensions.RecordsExist(madrsRosterData) Then
                            madrsRosterData.Filter = "SYSTEM_EXC_CD <> 'SLT'"
                            madrsRosterData.MoveFirst
                            If mADOExtensions.RecordsExist(madrsRosterData) Then
                                Set mobjCodeTables = CreateObject("dbCodeTables.cdCodeTables")
                                Set mobjEmployee = CreateObject("dbEmployee.cdEmployee")
                                Set mobjTimeCard = CreateObject("dbTimecard.cdTimecard")
                                Set mobjRules = CreateObject("dbRules.cdRules")
                                rbReportGenerated = True
                                madrsRosterData.Filter = adFilterNone
                                madrsRosterData.Filter = "EMP_NUM = " & _
                                            marrEmployees(mlngEmpIndex) & " AND " & _
                                            "SYSTEM_EXC_CD <> 'SLT'"
                                madrsRosterData.Sort = "EMP_NUM, CHARGE_DATE, DATE_WORKED, START_TIME, END_TIME, SCHED_SRC_CD"
                                madrsRosterData.MoveFirst
                                mblnFirstRecord = True
                                mlngPrevEmp = 0
                                mdtPrevDate = "12/31/9999"
                                Do While madrsRosterData.EOF = False
                                    PopulateDept_Record
                                    If Not mblnFirstRecord Then
                                        madrsFormatRptData.Update
                                    End If
                                    mlngPrevEmp = madrsRosterData!EMP_NUM
                                    llngPrevWG = marrWorkGroups(mlngEmpIndex)
                                    mdtPrevDate = madrsRosterData!CHARGE_DATE
                                madrsRosterData.MoveNext
                                Loop
                                Set mobjCodeTables = Nothing
                                Set mobjEmployee = Nothing
                                Set mobjTimeCard = Nothing
                                Set mobjRules = Nothing
                            End If
                        End If
                    End With
                Next mlngEmpIndex
            End If
        'Next mlngDateIndex
    Next mlngDeptIndex
    
    'Return the report object to the calling application
    If rbReportGenerated = True Then
        rptMain.Run
        
        '//check for any errors from the active reports script or the vb6 functions being called back
        If m_ActiveReports_ErrNumber <> 0 Then
            Err.Raise m_ActiveReports_ErrNumber, m_ActiveReports_ErrSource, m_ActiveReports_ErrDescription
        End If
        Set GetDailyRosterRptDept2 = rptMain
    Else
        Set GetDailyRosterRptDept2 = Nothing
    End If
    
    Set lobjEmployee = Nothing
    Set lobjDepartment = Nothing
    Set lobjWG = Nothing
    Set ladrsWG = Nothing
    
    '//-----clean up public properties used by ActiveReports-----//
    DestroyState
    
    Exit Function
GetDailyRosterRptDept2_EH:
    Set lobjEmployee = Nothing
    Set lobjDepartment = Nothing
    Set lobjWG = Nothing
    Set ladrsWG = Nothing
    
    DestroyState
    
    Select Case Err.Number
    Case m_ActiveReports_ErrNumber
        TrsRaiseError MODULENAME, "GetDailyRosterRptDept2", Err.Number, , , _
            DescArgsFrom("[OrigErrDescription]===>" & m_ActiveReports_OrigErrDescription & "<===", "[OrigErrSource]===>" & m_ActiveReports_OrigErrSource & "<==="), True
    Case Else
        TrsRaiseError MODULENAME, "GetDailyRosterRptDept2", , , , , True
    End Select
End Function

Private Sub PopulateWG_Record()
On Error GoTo PopulateWG_Record_EH

    Dim llngQualsIndex As Long
    Dim llngSkillsIndex As Long
    Dim ladrsOTShift As ADODB.Recordset
    Dim lblnSecondaryWG As Boolean
    Dim lintTotalStartMin As Integer
    Dim lintTotalEndMin As Integer
    Dim lblnProcessExcep As Boolean
    
    If mADOExtensions.RecordsExist(madrsEmpsinWG) Then
        With madrsRosterData
            If mlngPrevEmp <> !EMP_NUM Or (mlngPrevEmp = !EMP_NUM And _
                                           mdtPrevDate <> !CHARGE_DATE) Then
                If Not mblnFirstRecord Then
                    If mADOExtensions.RecordsExist(madrsEmployeeInelig) Then
                        madrsEmployeeInelig.MoveFirst
                        'madrsEmployeeInelig.Filter = "EMP_NUM = " & mlngPrevEmp & _
                            " AND EFF_START_DATE < #" & marrRequestDate(mlngDateIndex) & _
                            "# AND " & " EFF_END_DATE > #" & marrRequestDate(mlngDateIndex) & "#"
                        madrsEmployeeInelig.Filter = "EMP_NUM = " & mlngPrevEmp & _
                            " AND EFF_START_DATE <= #" & mdtPrevDate & _
                            "# AND " & " EFF_END_DATE >= #" & mdtPrevDate & "#"
                        If Not madrsEmployeeInelig.EOF Then
                            madrsEmployeeInelig.MoveFirst
                            Select Case mlngExceptionCount
                                Case 0
                                    madrsFormatRptData.Fields("EXCEPTION1") = _
                                        madrsEmployeeInelig.Fields("EXCEPTION_CD")
                                    madrsFormatRptData.Fields("EXCEPTION1_DESC") = _
                                        mobjCodeTables.ExceptionDescription( _
                                        madrsFormatRptData.Fields("EXCEPTION1"))
                                    mlngExceptionCount = mlngExceptionCount + 1
                                    madrsFormatRptData.Fields("INELIGIBILITY_START1") = _
                                        madrsEmployeeInelig.Fields("EFF_START_DATE")
                                    madrsFormatRptData.Fields("INELIGIBILITY_END1") = _
                                        madrsEmployeeInelig.Fields("EFF_END_DATE")
                                    madrsFormatRptData.Fields("EXCEPTION1_START") = _
                                        madrsEmployeeInelig.Fields("EFF_START_TIME")
                                    madrsFormatRptData.Fields("EXCEPTION1_END") = _
                                        madrsEmployeeInelig.Fields("EFF_END_TIME")
                                    madrsFormatRptData.Update
                                Case 1
                                    madrsFormatRptData.Fields("EXCEPTION2") = _
                                        madrsEmployeeInelig.Fields("EXCEPTION_CD")
                                    madrsFormatRptData.Fields("EXCEPTION2_DESC") = _
                                        mobjCodeTables.ExceptionDescription( _
                                        madrsFormatRptData.Fields("EXCEPTION2"))
                                    mlngExceptionCount = mlngExceptionCount + 1
                                    madrsFormatRptData.Fields("INELIGIBILITY_START2") = _
                                        madrsEmployeeInelig.Fields("EFF_START_DATE")
                                    madrsFormatRptData.Fields("INELIGIBILITY_END2") = _
                                        madrsEmployeeInelig.Fields("EFF_END_DATE")
                                    madrsFormatRptData.Fields("EXCEPTION2_START") = _
                                        madrsEmployeeInelig.Fields("EFF_START_TIME")
                                    madrsFormatRptData.Fields("EXCEPTION2_END") = _
                                        madrsEmployeeInelig.Fields("EFF_END_TIME")
                                    madrsFormatRptData.Update
                                Case 2
                                    madrsFormatRptData.Fields("EXCEPTION3") = _
                                        madrsEmployeeInelig.Fields("EXCEPTION_CD")
                                    madrsFormatRptData.Fields("EXCEPTION3_DESC") = _
                                        mobjCodeTables.ExceptionDescription( _
                                        madrsFormatRptData.Fields("EXCEPTION3"))
                                    mlngExceptionCount = mlngExceptionCount + 1
                                    madrsFormatRptData.Fields("INELIGIBILITY_START3") = _
                                        madrsEmployeeInelig.Fields("EFF_START_DATE")
                                    madrsFormatRptData.Fields("INELIGIBILITY_END3") = _
                                        madrsEmployeeInelig.Fields("EFF_END_DATE")
                                    madrsFormatRptData.Fields("EXCEPTION3_START") = _
                                        madrsEmployeeInelig.Fields("EFF_START_TIME")
                                    madrsFormatRptData.Fields("EXCEPTION3_END") = _
                                        madrsEmployeeInelig.Fields("EFF_END_TIME")
                                    madrsFormatRptData.Update
                                Case 3
                                    madrsFormatRptData.Fields("EXCEPTION4") = _
                                        madrsEmployeeInelig.Fields("EXCEPTION_CD")
                                    madrsFormatRptData.Fields("EXCEPTION4_DESC") = _
                                        mobjCodeTables.ExceptionDescription( _
                                        madrsFormatRptData.Fields("EXCEPTION4"))
                                    mlngExceptionCount = mlngExceptionCount + 1
                                    madrsFormatRptData.Fields("INELIGIBILITY_START4") = _
                                        madrsEmployeeInelig.Fields("EFF_START_DATE")
                                    madrsFormatRptData.Fields("INELIGIBILITY_END4") = _
                                        madrsEmployeeInelig.Fields("EFF_END_DATE")
                                    madrsFormatRptData.Fields("EXCEPTION4_START") = _
                                        madrsEmployeeInelig.Fields("EFF_START_TIME")
                                    madrsFormatRptData.Fields("EXCEPTION4_END") = _
                                        madrsEmployeeInelig.Fields("EFF_END_TIME")
                                    madrsFormatRptData.Update
                                Case 4
                                    madrsFormatRptData.Fields("EXCEPTION5") = _
                                        madrsEmployeeInelig.Fields("EXCEPTION_CD")
                                    madrsFormatRptData.Fields("EXCEPTION5_DESC") = _
                                        mobjCodeTables.ExceptionDescription( _
                                        madrsFormatRptData.Fields("EXCEPTION5"))
                                    mlngExceptionCount = mlngExceptionCount + 1
                                    madrsFormatRptData.Fields("INELIGIBILITY_START5") = _
                                        madrsEmployeeInelig.Fields("EFF_START_DATE")
                                    madrsFormatRptData.Fields("INELIGIBILITY_END5") = _
                                        madrsEmployeeInelig.Fields("EFF_END_DATE")
                                    madrsFormatRptData.Fields("EXCEPTION5_START") = _
                                        madrsEmployeeInelig.Fields("EFF_START_TIME")
                                    madrsFormatRptData.Fields("EXCEPTION5_END") = _
                                        madrsEmployeeInelig.Fields("EFF_END_TIME")
                                    madrsFormatRptData.Update
                                Case 5
                                    madrsFormatRptData.Fields("EXCEPTION6") = _
                                        madrsEmployeeInelig.Fields("EXCEPTION_CD")
                                    madrsFormatRptData.Fields("EXCEPTION6_DESC") = _
                                        mobjCodeTables.ExceptionDescription( _
                                        madrsFormatRptData.Fields("EXCEPTION6"))
                                    mlngExceptionCount = mlngExceptionCount + 1
                                    madrsFormatRptData.Fields("INELIGIBILITY_START6") = _
                                        madrsEmployeeInelig.Fields("EFF_START_DATE")
                                    madrsFormatRptData.Fields("INELIGIBILITY_END6") = _
                                        madrsEmployeeInelig.Fields("EFF_END_DATE")
                                    madrsFormatRptData.Fields("EXCEPTION6_START") = _
                                        madrsEmployeeInelig.Fields("EFF_START_TIME")
                                    madrsFormatRptData.Fields("EXCEPTION6_END") = _
                                        madrsEmployeeInelig.Fields("EFF_END_TIME")
                                    madrsFormatRptData.Update
                                Case 6
                                    madrsFormatRptData.Fields("EXCEPTION7") = _
                                        madrsEmployeeInelig.Fields("EXCEPTION_CD")
                                    madrsFormatRptData.Fields("EXCEPTION7_DESC") = _
                                        mobjCodeTables.ExceptionDescription( _
                                        madrsFormatRptData.Fields("EXCEPTION7"))
                                    mlngExceptionCount = mlngExceptionCount + 1
                                    madrsFormatRptData.Fields("INELIGIBILITY_START7") = _
                                        madrsEmployeeInelig.Fields("EFF_START_DATE")
                                    madrsFormatRptData.Fields("INELIGIBILITY_END7") = _
                                        madrsEmployeeInelig.Fields("EFF_END_DATE")
                                    madrsFormatRptData.Fields("EXCEPTION7_START") = _
                                        madrsEmployeeInelig.Fields("EFF_START_TIME")
                                    madrsFormatRptData.Fields("EXCEPTION7_END") = _
                                        madrsEmployeeInelig.Fields("EFF_END_TIME")
                                    madrsFormatRptData.Update
                                Case 7
                                    madrsFormatRptData.Fields("EXCEPTION8") = _
                                        madrsEmployeeInelig.Fields("EXCEPTION_CD")
                                    madrsFormatRptData.Fields("EXCEPTION8_DESC") = _
                                        mobjCodeTables.ExceptionDescription( _
                                        madrsFormatRptData.Fields("EXCEPTION8"))
                                    mlngExceptionCount = mlngExceptionCount + 1
                                    madrsFormatRptData.Fields("INELIGIBILITY_START8") = _
                                        madrsEmployeeInelig.Fields("EFF_START_DATE")
                                    madrsFormatRptData.Fields("INELIGIBILITY_END8") = _
                                        madrsEmployeeInelig.Fields("EFF_END_DATE")
                                    madrsFormatRptData.Fields("EXCEPTION8_START") = _
                                        madrsEmployeeInelig.Fields("EFF_START_TIME")
                                    madrsFormatRptData.Fields("EXCEPTION8_END") = _
                                        madrsEmployeeInelig.Fields("EFF_END_TIME")
                                    madrsFormatRptData.Update
                                Case 8
                                    madrsFormatRptData.Fields("EXCEPTION9") = _
                                        madrsEmployeeInelig.Fields("EXCEPTION_CD")
                                    madrsFormatRptData.Fields("EXCEPTION9_DESC") = _
                                        mobjCodeTables.ExceptionDescription( _
                                        madrsFormatRptData.Fields("EXCEPTION9"))
                                    mlngExceptionCount = mlngExceptionCount + 1
                                    madrsFormatRptData.Fields("INELIGIBILITY_START9") = _
                                        madrsEmployeeInelig.Fields("EFF_START_DATE")
                                    madrsFormatRptData.Fields("INELIGIBILITY_END9") = _
                                        madrsEmployeeInelig.Fields("EFF_END_DATE")
                                    madrsFormatRptData.Fields("EXCEPTION9_START") = _
                                        madrsEmployeeInelig.Fields("EFF_START_TIME")
                                    madrsFormatRptData.Fields("EXCEPTION9_END") = _
                                        madrsEmployeeInelig.Fields("EFF_END_TIME")
                                    madrsFormatRptData.Update
                                Case 9
                                    madrsFormatRptData.Fields("EXCEPTION10") = _
                                        madrsEmployeeInelig.Fields("EXCEPTION_CD")
                                    madrsFormatRptData.Fields("EXCEPTION10_DESC") = _
                                        mobjCodeTables.ExceptionDescription( _
                                        madrsFormatRptData.Fields("EXCEPTION10"))
                                    mlngExceptionCount = mlngExceptionCount + 1
                                    madrsFormatRptData.Fields("INELIGIBILITY_START10") = _
                                        madrsEmployeeInelig.Fields("EFF_START_DATE")
                                    madrsFormatRptData.Fields("INELIGIBILITY_END10") = _
                                        madrsEmployeeInelig.Fields("EFF_END_DATE")
                                    madrsFormatRptData.Fields("EXCEPTION10_START") = _
                                        madrsEmployeeInelig.Fields("EFF_START_TIME")
                                    madrsFormatRptData.Fields("EXCEPTION10_END") = _
                                        madrsEmployeeInelig.Fields("EFF_END_TIME")
                                    madrsFormatRptData.Update
                                Case 10
                                    madrsFormatRptData.Fields("EXCEPTION11") = _
                                        madrsEmployeeInelig.Fields("EXCEPTION_CD")
                                    madrsFormatRptData.Fields("EXCEPTION11_DESC") = _
                                        mobjCodeTables.ExceptionDescription( _
                                        madrsFormatRptData.Fields("EXCEPTION11"))
                                    mlngExceptionCount = mlngExceptionCount + 1
                                    madrsFormatRptData.Fields("INELIGIBILITY_START11") = _
                                        madrsEmployeeInelig.Fields("EFF_START_DATE")
                                    madrsFormatRptData.Fields("INELIGIBILITY_END11") = _
                                        madrsEmployeeInelig.Fields("EFF_END_DATE")
                                    madrsFormatRptData.Fields("EXCEPTION11_START") = _
                                        madrsEmployeeInelig.Fields("EFF_START_TIME")
                                    madrsFormatRptData.Fields("EXCEPTION11_END") = _
                                        madrsEmployeeInelig.Fields("EFF_END_TIME")
                                    madrsFormatRptData.Update
                                Case 11
                                    madrsFormatRptData.Fields("EXCEPTION12") = _
                                        madrsEmployeeInelig.Fields("EXCEPTION_CD")
                                    madrsFormatRptData.Fields("EXCEPTION12_DESC") = _
                                        mobjCodeTables.ExceptionDescription( _
                                        madrsFormatRptData.Fields("EXCEPTION12"))
                                    mlngExceptionCount = mlngExceptionCount + 1
                                    madrsFormatRptData.Fields("INELIGIBILITY_START12") = _
                                        madrsEmployeeInelig.Fields("EFF_START_DATE")
                                    madrsFormatRptData.Fields("INELIGIBILITY_END12") = _
                                        madrsEmployeeInelig.Fields("EFF_END_DATE")
                                    madrsFormatRptData.Fields("EXCEPTION12_START") = _
                                        madrsEmployeeInelig.Fields("EFF_START_TIME")
                                    madrsFormatRptData.Fields("EXCEPTION12_END") = _
                                        madrsEmployeeInelig.Fields("EFF_END_TIME")
                                    madrsFormatRptData.Update
                                Case 12
                                    madrsFormatRptData.Fields("EXCEPTION13") = _
                                        madrsEmployeeInelig.Fields("EXCEPTION_CD")
                                    madrsFormatRptData.Fields("EXCEPTION13_DESC") = _
                                        mobjCodeTables.ExceptionDescription( _
                                        madrsFormatRptData.Fields("EXCEPTION13"))
                                    mlngExceptionCount = mlngExceptionCount + 1
                                    madrsFormatRptData.Fields("INELIGIBILITY_START13") = _
                                        madrsEmployeeInelig.Fields("EFF_START_DATE")
                                    madrsFormatRptData.Fields("INELIGIBILITY_END13") = _
                                        madrsEmployeeInelig.Fields("EFF_END_DATE")
                                    madrsFormatRptData.Fields("EXCEPTION13_START") = _
                                        madrsEmployeeInelig.Fields("EFF_START_TIME")
                                    madrsFormatRptData.Fields("EXCEPTION13_END") = _
                                        madrsEmployeeInelig.Fields("EFF_END_TIME")
                                    madrsFormatRptData.Update
                                Case 13
                                    madrsFormatRptData.Fields("EXCEPTION14") = _
                                        madrsEmployeeInelig.Fields("EXCEPTION_CD")
                                    madrsFormatRptData.Fields("EXCEPTION14_DESC") = _
                                        mobjCodeTables.ExceptionDescription( _
                                        madrsFormatRptData.Fields("EXCEPTION14"))
                                    mlngExceptionCount = mlngExceptionCount + 1
                                    madrsFormatRptData.Fields("INELIGIBILITY_START14") = _
                                        madrsEmployeeInelig.Fields("EFF_START_DATE")
                                    madrsFormatRptData.Fields("INELIGIBILITY_END14") = _
                                        madrsEmployeeInelig.Fields("EFF_END_DATE")
                                    madrsFormatRptData.Fields("EXCEPTION14_START") = _
                                        madrsEmployeeInelig.Fields("EFF_START_TIME")
                                    madrsFormatRptData.Fields("EXCEPTION14_END") = _
                                        madrsEmployeeInelig.Fields("EFF_END_TIME")
                                    madrsFormatRptData.Update
                                Case 14
                                    madrsFormatRptData.Fields("EXCEPTION15") = _
                                        madrsEmployeeInelig.Fields("EXCEPTION_CD")
                                    madrsFormatRptData.Fields("EXCEPTION15_DESC") = _
                                        mobjCodeTables.ExceptionDescription( _
                                        madrsFormatRptData.Fields("EXCEPTION15"))
                                    mlngExceptionCount = mlngExceptionCount + 1
                                    madrsFormatRptData.Fields("INELIGIBILITY_START15") = _
                                        madrsEmployeeInelig.Fields("EFF_START_DATE")
                                    madrsFormatRptData.Fields("INELIGIBILITY_END15") = _
                                        madrsEmployeeInelig.Fields("EFF_END_DATE")
                                    madrsFormatRptData.Fields("EXCEPTION15_START") = _
                                        madrsEmployeeInelig.Fields("EFF_START_TIME")
                                    madrsFormatRptData.Fields("EXCEPTION15_END") = _
                                        madrsEmployeeInelig.Fields("EFF_END_TIME")
                                    madrsFormatRptData.Update
                            End Select
                        End If
                        madrsEmployeeInelig.Filter = adFilterNone
                    End If
                End If
                madrsFormatRptData.AddNew
                mblnFirstRecord = False
                madrsFormatRptData.Fields("WG_NUM") = marrWorkGroups(mlngWGIndex)
                madrsFormatRptData.Fields("WG_DESC") = marrWorkGroupDesc(mlngWGIndex)
                madrsFormatRptData.Fields("EMP_NUM") = !EMP_NUM
                madrsEmpsinWG.MoveFirst
                madrsEmpsinWG.Find "EMP_NUM = " & !EMP_NUM
                madrsFormatRptData.Fields("EMP_NAME") = madrsEmpsinWG.Fields("EMP_NAME")
                'lblnSecondaryWG = mobjEmployee.IsSecondaryWorkGroup(!EMP_NUM, _
                    marrWorkGroups(mlngWGIndex), marrRequestDate(mlngDateIndex))
                lblnSecondaryWG = mobjEmployee.IsSecondaryWorkGroup(!EMP_NUM, _
                    marrWorkGroups(mlngWGIndex), !CHARGE_DATE)
                If lblnSecondaryWG = True Then
                    madrsFormatRptData.Fields("SECONDARY_IND") = "*"
                Else
                    madrsFormatRptData.Fields("SECONDARY_IND") = Null
                End If
                If mblnPrintQualsSkills Then
                    If mADOExtensions.RecordsExist(madrsQuals) Then
                        madrsQuals.MoveFirst
                        madrsQuals.Filter = "EMP_NUM = " & !EMP_NUM
                        If Not madrsQuals.EOF Then
                            madrsQuals.MoveFirst
                            llngQualsIndex = 1
                            While llngQualsIndex < 19 And Not madrsQuals.EOF
                                If llngQualsIndex > 1 Then
                                    madrsFormatRptData.Fields("QUALS") = _
                                        madrsFormatRptData.Fields("QUALS") & ", " & _
                                        madrsQuals.Fields("QUAL_CD")
                                Else
                                    madrsFormatRptData.Fields("QUALS") = _
                                        "QUALIFICATIONS: " & madrsQuals.Fields("QUAL_CD")
                                End If
                                llngQualsIndex = llngQualsIndex + 1
                                madrsQuals.MoveNext
                            Wend
                        Else
                            madrsFormatRptData.Fields("QUALS") = Null
                        End If
                    Else
                        madrsFormatRptData.Fields("QUALS") = Null
                    End If
                    madrsQuals.Filter = adFilterNone
                    If mADOExtensions.RecordsExist(madrsSkills) Then
                        madrsSkills.MoveFirst
                        madrsSkills.Filter = "EMP_NUM = " & !EMP_NUM
                        If Not madrsSkills.EOF Then
                            madrsSkills.MoveFirst
                            llngSkillsIndex = 1
                            While llngSkillsIndex < 19 And Not madrsSkills.EOF
                                If llngSkillsIndex > 1 Then
                                    madrsFormatRptData.Fields("SKILLS") = _
                                        madrsFormatRptData.Fields("SKILLS") & ", " & _
                                        madrsSkills.Fields("SPEC_SKILL_CD")
                                Else
                                    madrsFormatRptData.Fields("SKILLS") = _
                                        "SKILLS: " & madrsSkills.Fields("SPEC_SKILL_CD")
                                End If
                                llngSkillsIndex = llngSkillsIndex + 1
                                madrsSkills.MoveNext
                            Wend
                        Else
                            madrsFormatRptData.Fields("SKILLS") = Null
                        End If
                    Else
                        madrsFormatRptData.Fields("SKILLS") = Null
                    End If
                    madrsSkills.Filter = adFilterNone
                Else
                    madrsFormatRptData.Fields("QUALS") = Null
                    madrsFormatRptData.Fields("SKILLS") = Null
                End If
                If mblnDemandInd Then
                    madrsFormatRptData.Fields("LINE_NUM") = !LINE_NUM
                Else
                    madrsFormatRptData.Fields("LINE_NUM") = 0
                End If
                mlngExceptionCount = 0
            End If
            madrsFormatRptData.Fields("CHARGE_DATE") = !CHARGE_DATE
            If !SHIFT_NUM = 1 Then
                madrsFormatRptData.Fields("REG_SHIFT") = !SHIFT_CODE
                Select Case Trim(!SHIFT_CODE)
                    Case "O"
                        madrsFormatRptData.Fields("SORT_CODE") = 1
                    Case "A"
                        madrsFormatRptData.Fields("SORT_CODE") = 2
                    Case "P"
                        madrsFormatRptData.Fields("SORT_CODE") = 3
                    Case Else
                        madrsFormatRptData.Fields("SORT_CODE") = 0
                End Select
                If Trim(!SHIFT_START_TIME) = "24:00" Then
                    madrsFormatRptData.Fields("REG_SHIFT_START") = "00:00"
                Else
                    madrsFormatRptData.Fields("REG_SHIFT_START") = !SHIFT_START_TIME
                End If
                If Trim(!SHIFT_END_TIME) = "24:00" Then
                    madrsFormatRptData.Fields("REG_SHIFT_END") = "00:00"
                Else
                    madrsFormatRptData.Fields("REG_SHIFT_END") = !SHIFT_END_TIME
                End If
            End If
            Select Case !SCHED_SRC_CD
                Case "50"
                    Select Case !SYSTEM_EXC_CD
                        Case "TWS"
                            madrsFormatRptData.Fields("TRADE_RAS") = "*S*"
                        Case "TOR"
                            madrsFormatRptData.Fields("TRADE_RAS") = "*R*"
                        Case "TSR"
                            madrsFormatRptData.Fields("TRADE_RAS") = "*R*"
                        Case "RWT"
                            madrsFormatRptData.Fields("TRADE_RAS") = "*R*"
                        Case "TOA"
                            madrsFormatRptData.Fields("TRADE_RAS") = "*A*"
                        Case "TSA"
                            madrsFormatRptData.Fields("TRADE_RAS") = "*A*"
                    End Select
                Case "30"
                    Select Case !SYSTEM_EXC_CD
                        Case "TWS"
                            madrsFormatRptData.Fields("TRADE_RAS") = "*S*"
                        Case "TOR"
                            madrsFormatRptData.Fields("TRADE_RAS") = "*R*"
                        Case "TSR"
                            madrsFormatRptData.Fields("TRADE_RAS") = "*R*"
                        Case "RWT"
                            madrsFormatRptData.Fields("TRADE_RAS") = "*R*"
                        Case "TOA"
                            madrsFormatRptData.Fields("TRADE_RAS") = "*A*"
                        Case "TSA"
                            madrsFormatRptData.Fields("TRADE_RAS") = "*A*"
                    End Select
            End Select
            If !SHIFT_NUM = 2 Then
                madrsFormatRptData.Fields("TRADE_DOUBLE") = !SHIFT_CODE
            End If
            If Trim(!EXCEPTION_CD) <> "" And mlngExceptionCount < 15 Then
                lblnProcessExcep = False
                If Trim(!OVER_EXC_CD) <> "" Then
                    If ValidExcep(Trim(!OVER_EXC_CD)) Then
                        Select Case mlngExceptionCount
                            Case 0
                                madrsFormatRptData.Fields("EXCEPTION1") = !OVER_EXC_CD
                                lblnProcessExcep = True
                            Case 1
                                If ExcepExists(!OVER_EXC_CD) Then
                                    lblnProcessExcep = False
                                Else
                                    madrsFormatRptData.Fields("EXCEPTION2") = !OVER_EXC_CD
                                    lblnProcessExcep = True
                                End If
                            Case 2
                                If ExcepExists(!OVER_EXC_CD) Then
                                    lblnProcessExcep = False
                                Else
                                    madrsFormatRptData.Fields("EXCEPTION3") = !OVER_EXC_CD
                                    lblnProcessExcep = True
                                End If
                            Case 3
                                If ExcepExists(!OVER_EXC_CD) Then
                                    lblnProcessExcep = False
                                Else
                                    madrsFormatRptData.Fields("EXCEPTION4") = !OVER_EXC_CD
                                    lblnProcessExcep = True
                                End If
                            Case 4
                                If ExcepExists(!OVER_EXC_CD) Then
                                    lblnProcessExcep = False
                                Else
                                    madrsFormatRptData.Fields("EXCEPTION5") = !OVER_EXC_CD
                                    lblnProcessExcep = True
                                End If
                            Case 5
                                If ExcepExists(!OVER_EXC_CD) Then
                                    lblnProcessExcep = False
                                Else
                                    madrsFormatRptData.Fields("EXCEPTION6") = !OVER_EXC_CD
                                    lblnProcessExcep = True
                                End If
                            Case 6
                                If ExcepExists(!OVER_EXC_CD) Then
                                    lblnProcessExcep = False
                                Else
                                    madrsFormatRptData.Fields("EXCEPTION7") = !OVER_EXC_CD
                                    lblnProcessExcep = True
                                End If
                            Case 7
                                If ExcepExists(!OVER_EXC_CD) Then
                                    lblnProcessExcep = False
                                Else
                                    madrsFormatRptData.Fields("EXCEPTION8") = !OVER_EXC_CD
                                    lblnProcessExcep = True
                                End If
                            Case 8
                                If ExcepExists(!OVER_EXC_CD) Then
                                    lblnProcessExcep = False
                                Else
                                    madrsFormatRptData.Fields("EXCEPTION9") = !OVER_EXC_CD
                                    lblnProcessExcep = True
                                End If
                            Case 9
                                If ExcepExists(!OVER_EXC_CD) Then
                                    lblnProcessExcep = False
                                Else
                                    madrsFormatRptData.Fields("EXCEPTION10") = !OVER_EXC_CD
                                    lblnProcessExcep = True
                                End If
                            Case 10
                                If ExcepExists(!OVER_EXC_CD) Then
                                    lblnProcessExcep = False
                                Else
                                    madrsFormatRptData.Fields("EXCEPTION11") = !OVER_EXC_CD
                                    lblnProcessExcep = True
                                End If
                            Case 11
                                If ExcepExists(!OVER_EXC_CD) Then
                                    lblnProcessExcep = False
                                Else
                                    madrsFormatRptData.Fields("EXCEPTION12") = !OVER_EXC_CD
                                    lblnProcessExcep = True
                                End If
                            Case 12
                                If ExcepExists(!OVER_EXC_CD) Then
                                    lblnProcessExcep = False
                                Else
                                    madrsFormatRptData.Fields("EXCEPTION13") = !OVER_EXC_CD
                                    lblnProcessExcep = True
                                End If
                            Case 13
                                If ExcepExists(!OVER_EXC_CD) Then
                                    lblnProcessExcep = False
                                Else
                                    madrsFormatRptData.Fields("EXCEPTION14") = !OVER_EXC_CD
                                    lblnProcessExcep = True
                                End If
                            Case 14
                                If ExcepExists(!OVER_EXC_CD) Then
                                    lblnProcessExcep = False
                                Else
                                    madrsFormatRptData.Fields("EXCEPTION15") = !OVER_EXC_CD
                                    lblnProcessExcep = True
                                End If
                        End Select
                    End If
                Else
                    If ValidExcep(Trim(!EXCEPTION_CD)) Then
                        Select Case mlngExceptionCount
                            Case 0
                                madrsFormatRptData.Fields("EXCEPTION1") = !EXCEPTION_CD
                                lblnProcessExcep = True
                            Case 1
                                If ExcepExists(!EXCEPTION_CD) Then
                                    lblnProcessExcep = False
                                Else
                                    madrsFormatRptData.Fields("EXCEPTION2") = !EXCEPTION_CD
                                    lblnProcessExcep = True
                                End If
                            Case 2
                                If ExcepExists(!EXCEPTION_CD) Then
                                    lblnProcessExcep = False
                                Else
                                    madrsFormatRptData.Fields("EXCEPTION3") = !EXCEPTION_CD
                                    lblnProcessExcep = True
                                End If
                            Case 3
                                If ExcepExists(!EXCEPTION_CD) Then
                                    lblnProcessExcep = False
                                Else
                                    madrsFormatRptData.Fields("EXCEPTION4") = !EXCEPTION_CD
                                    lblnProcessExcep = True
                                End If
                            Case 4
                                If ExcepExists(!EXCEPTION_CD) Then
                                    lblnProcessExcep = False
                                Else
                                    madrsFormatRptData.Fields("EXCEPTION5") = !EXCEPTION_CD
                                    lblnProcessExcep = True
                                End If
                            Case 5
                                If ExcepExists(!EXCEPTION_CD) Then
                                    lblnProcessExcep = False
                                Else
                                    madrsFormatRptData.Fields("EXCEPTION6") = !EXCEPTION_CD
                                    lblnProcessExcep = True
                                End If
                            Case 6
                                If ExcepExists(!EXCEPTION_CD) Then
                                    lblnProcessExcep = False
                                Else
                                    madrsFormatRptData.Fields("EXCEPTION7") = !EXCEPTION_CD
                                    lblnProcessExcep = True
                                End If
                            Case 7
                                If ExcepExists(!EXCEPTION_CD) Then
                                    lblnProcessExcep = False
                                Else
                                    madrsFormatRptData.Fields("EXCEPTION8") = !EXCEPTION_CD
                                    lblnProcessExcep = True
                                End If
                            Case 8
                                If ExcepExists(!EXCEPTION_CD) Then
                                    lblnProcessExcep = False
                                Else
                                    madrsFormatRptData.Fields("EXCEPTION9") = !EXCEPTION_CD
                                    lblnProcessExcep = True
                                End If
                            Case 9
                                If ExcepExists(!EXCEPTION_CD) Then
                                    lblnProcessExcep = False
                                Else
                                    madrsFormatRptData.Fields("EXCEPTION10") = !EXCEPTION_CD
                                    lblnProcessExcep = True
                                End If
                            Case 10
                                If ExcepExists(!EXCEPTION_CD) Then
                                    lblnProcessExcep = False
                                Else
                                    madrsFormatRptData.Fields("EXCEPTION11") = !EXCEPTION_CD
                                    lblnProcessExcep = True
                                End If
                            Case 11
                                If ExcepExists(!EXCEPTION_CD) Then
                                    lblnProcessExcep = False
                                Else
                                    madrsFormatRptData.Fields("EXCEPTION12") = !EXCEPTION_CD
                                    lblnProcessExcep = True
                                End If
                            Case 12
                                If ExcepExists(!EXCEPTION_CD) Then
                                    lblnProcessExcep = False
                                Else
                                    madrsFormatRptData.Fields("EXCEPTION13") = !EXCEPTION_CD
                                    lblnProcessExcep = True
                                End If
                            Case 13
                                If ExcepExists(!EXCEPTION_CD) Then
                                    lblnProcessExcep = False
                                Else
                                    madrsFormatRptData.Fields("EXCEPTION14") = !EXCEPTION_CD
                                    lblnProcessExcep = True
                                End If
                            Case 14
                                If ExcepExists(!EXCEPTION_CD) Then
                                    lblnProcessExcep = False
                                Else
                                    madrsFormatRptData.Fields("EXCEPTION15") = !EXCEPTION_CD
                                    lblnProcessExcep = True
                                End If
                        End Select
                    End If
                End If
                If lblnProcessExcep Then
                    Select Case mlngExceptionCount
                        Case 0
                            madrsFormatRptData.Fields("EXCEPTION1_DESC") = _
                            mobjCodeTables.ExceptionDescription(madrsFormatRptData.Fields("EXCEPTION1"))
                            madrsFormatRptData.Fields("EXCEPTION1_START") = !START_TIME
                            madrsFormatRptData.Fields("EXCEPTION1_END") = !END_TIME
                            If Trim(madrsFormatRptData.Fields("EXCEPTION1_START")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION1_START") = "00:00"
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION1_END")) = "00:00" Then
                                madrsFormatRptData.Fields("EXCEPTION1_END") = "24:00"
                            End If
                            mlngExceptionCount = mlngExceptionCount + 1
                            If OvertimeExcep(Trim(madrsFormatRptData.Fields("EXCEPTION1"))) _
                                And !OT_OPP_NUM > 0 Then
                                lintTotalStartMin = _
                                (Format(Mid(madrsFormatRptData.Fields("EXCEPTION1_START"), 1, 2), "00") * 60) + _
                                 Format(Mid(madrsFormatRptData.Fields("EXCEPTION1_START"), 4, 2), "00")
                                lintTotalEndMin = _
                                (Format(Mid(madrsFormatRptData.Fields("EXCEPTION1_END"), 1, 2), "00") * 60) + _
                                 Format(Mid(madrsFormatRptData.Fields("EXCEPTION1_END"), 4, 2), "00")
                                madrsFormatRptData.Fields("CANVASSED_OT_HOURS1") = _
                                CStr(Abs(((lintTotalEndMin - lintTotalStartMin) / 60)))
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION1_START")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION1_START") = "00:00"
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION1_END")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION1_END") = "00:00"
                            End If
                            If !OT_OPP_NUM > 0 Then
                                Set ladrsOTShift = mobjTimeCard.GetOTShiftCode(!OT_OPP_NUM)
                                If mADOExtensions.RecordsExist(ladrsOTShift) Then
                                    madrsFormatRptData.Fields("CANVASSED_OT_SHIFT1") = _
                                    ladrsOTShift.Fields("SHIFT_CD")
                                End If
                            End If
                        Case 1
                            madrsFormatRptData.Fields("EXCEPTION2_DESC") = _
                            mobjCodeTables.ExceptionDescription(madrsFormatRptData.Fields("EXCEPTION2"))
                            madrsFormatRptData.Fields("EXCEPTION2_START") = !START_TIME
                            madrsFormatRptData.Fields("EXCEPTION2_END") = !END_TIME
                            If Trim(madrsFormatRptData.Fields("EXCEPTION2_START")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION2_START") = "00:00"
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION2_END")) = "00:00" Then
                                madrsFormatRptData.Fields("EXCEPTION2_END") = "24:00"
                            End If
                            mlngExceptionCount = mlngExceptionCount + 1
                            If OvertimeExcep(Trim(madrsFormatRptData.Fields("EXCEPTION2"))) _
                                And !OT_OPP_NUM > 0 Then
                                lintTotalStartMin = _
                                (Format(Mid(madrsFormatRptData.Fields("EXCEPTION2_START"), 1, 2), "00") * 60) + _
                                 Format(Mid(madrsFormatRptData.Fields("EXCEPTION2_START"), 4, 2), "00")
                                lintTotalEndMin = _
                                (Format(Mid(madrsFormatRptData.Fields("EXCEPTION2_END"), 1, 2), "00") * 60) + _
                                 Format(Mid(madrsFormatRptData.Fields("EXCEPTION2_END"), 4, 2), "00")
                                madrsFormatRptData.Fields("CANVASSED_OT_HOURS2") = _
                                CStr(Abs(((lintTotalEndMin - lintTotalStartMin) / 60)))
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION2_START")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION2_START") = "00:00"
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION2_END")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION2_END") = "00:00"
                            End If
                            If !OT_OPP_NUM > 0 Then
                                Set ladrsOTShift = mobjTimeCard.GetOTShiftCode(!OT_OPP_NUM)
                                If mADOExtensions.RecordsExist(ladrsOTShift) Then
                                    madrsFormatRptData.Fields("CANVASSED_OT_SHIFT2") = _
                                    ladrsOTShift.Fields("SHIFT_CD")
                                End If
                            End If
                        Case 2
                            madrsFormatRptData.Fields("EXCEPTION3_DESC") = _
                            mobjCodeTables.ExceptionDescription(madrsFormatRptData.Fields("EXCEPTION3"))
                            madrsFormatRptData.Fields("EXCEPTION3_START") = !START_TIME
                            madrsFormatRptData.Fields("EXCEPTION3_END") = !END_TIME
                            If Trim(madrsFormatRptData.Fields("EXCEPTION3_START")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION3_START") = "00:00"
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION3_END")) = "00:00" Then
                                madrsFormatRptData.Fields("EXCEPTION3_END") = "24:00"
                            End If
                            mlngExceptionCount = mlngExceptionCount + 1
                            If OvertimeExcep(Trim(madrsFormatRptData.Fields("EXCEPTION3"))) _
                                And !OT_OPP_NUM > 0 Then
                                lintTotalStartMin = _
                                (Format(Mid(madrsFormatRptData.Fields("EXCEPTION3_START"), 1, 2), "00") * 60) + _
                                 Format(Mid(madrsFormatRptData.Fields("EXCEPTION3_START"), 4, 2), "00")
                                lintTotalEndMin = _
                                (Format(Mid(madrsFormatRptData.Fields("EXCEPTION3_END"), 1, 2), "00") * 60) + _
                                 Format(Mid(madrsFormatRptData.Fields("EXCEPTION3_END"), 4, 2), "00")
                                madrsFormatRptData.Fields("CANVASSED_OT_HOURS3") = _
                                CStr(Abs(((lintTotalEndMin - lintTotalStartMin) / 60)))
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION3_START")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION3_START") = "00:00"
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION3_END")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION3_END") = "00:00"
                            End If
                            If !OT_OPP_NUM > 0 Then
                                Set ladrsOTShift = mobjTimeCard.GetOTShiftCode(!OT_OPP_NUM)
                                If mADOExtensions.RecordsExist(ladrsOTShift) Then
                                    madrsFormatRptData.Fields("CANVASSED_OT_SHIFT3") = _
                                    ladrsOTShift.Fields("SHIFT_CD")
                                End If
                            End If
                        Case 3
                            madrsFormatRptData.Fields("EXCEPTION4_DESC") = _
                            mobjCodeTables.ExceptionDescription(madrsFormatRptData.Fields("EXCEPTION4"))
                            madrsFormatRptData.Fields("EXCEPTION4_START") = !START_TIME
                            madrsFormatRptData.Fields("EXCEPTION4_END") = !END_TIME
                            If Trim(madrsFormatRptData.Fields("EXCEPTION4_START")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION4_START") = "00:00"
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION4_END")) = "00:00" Then
                                madrsFormatRptData.Fields("EXCEPTION4_END") = "24:00"
                            End If
                            mlngExceptionCount = mlngExceptionCount + 1
                            If OvertimeExcep(Trim(madrsFormatRptData.Fields("EXCEPTION4"))) _
                                And !OT_OPP_NUM > 0 Then
                                lintTotalStartMin = _
                                (Format(Mid(madrsFormatRptData.Fields("EXCEPTION4_START"), 1, 2), "00") * 60) + _
                                 Format(Mid(madrsFormatRptData.Fields("EXCEPTION4_START"), 4, 2), "00")
                                lintTotalEndMin = _
                                (Format(Mid(madrsFormatRptData.Fields("EXCEPTION4_END"), 1, 2), "00") * 60) + _
                                 Format(Mid(madrsFormatRptData.Fields("EXCEPTION4_END"), 4, 2), "00")
                                madrsFormatRptData.Fields("CANVASSED_OT_HOURS4") = _
                                CStr(Abs(((lintTotalEndMin - lintTotalStartMin) / 60)))
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION4_START")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION4_START") = "00:00"
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION4_END")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION4_END") = "00:00"
                            End If
                            If !OT_OPP_NUM > 0 Then
                                Set ladrsOTShift = mobjTimeCard.GetOTShiftCode(!OT_OPP_NUM)
                                If mADOExtensions.RecordsExist(ladrsOTShift) Then
                                    madrsFormatRptData.Fields("CANVASSED_OT_SHIFT4") = _
                                    ladrsOTShift.Fields("SHIFT_CD")
                                End If
                            End If
                        Case 4
                            madrsFormatRptData.Fields("EXCEPTION5_DESC") = _
                            mobjCodeTables.ExceptionDescription(madrsFormatRptData.Fields("EXCEPTION5"))
                            madrsFormatRptData.Fields("EXCEPTION5_START") = !START_TIME
                            madrsFormatRptData.Fields("EXCEPTION5_END") = !END_TIME
                            If Trim(madrsFormatRptData.Fields("EXCEPTION5_START")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION5_START") = "00:00"
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION5_END")) = "00:00" Then
                                madrsFormatRptData.Fields("EXCEPTION5_END") = "24:00"
                            End If
                            mlngExceptionCount = mlngExceptionCount + 1
                            If OvertimeExcep(Trim(madrsFormatRptData.Fields("EXCEPTION5"))) _
                                And !OT_OPP_NUM > 0 Then
                                lintTotalStartMin = _
                                (Format(Mid(madrsFormatRptData.Fields("EXCEPTION5_START"), 1, 2), "00") * 60) + _
                                 Format(Mid(madrsFormatRptData.Fields("EXCEPTION5_START"), 4, 2), "00")
                                lintTotalEndMin = _
                                (Format(Mid(madrsFormatRptData.Fields("EXCEPTION5_END"), 1, 2), "00") * 60) + _
                                 Format(Mid(madrsFormatRptData.Fields("EXCEPTION5_END"), 4, 2), "00")
                                madrsFormatRptData.Fields("CANVASSED_OT_HOURS5") = _
                                CStr(Abs(((lintTotalEndMin - lintTotalStartMin) / 60)))
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION5_START")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION5_START") = "00:00"
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION5_END")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION5_END") = "00:00"
                            End If
                            If !OT_OPP_NUM > 0 Then
                                Set ladrsOTShift = mobjTimeCard.GetOTShiftCode(!OT_OPP_NUM)
                                If mADOExtensions.RecordsExist(ladrsOTShift) Then
                                    madrsFormatRptData.Fields("CANVASSED_OT_SHIFT5") = _
                                    ladrsOTShift.Fields("SHIFT_CD")
                                End If
                            End If
                        Case 5
                            madrsFormatRptData.Fields("EXCEPTION6_DESC") = _
                            mobjCodeTables.ExceptionDescription(madrsFormatRptData.Fields("EXCEPTION6"))
                            madrsFormatRptData.Fields("EXCEPTION6_START") = !START_TIME
                            madrsFormatRptData.Fields("EXCEPTION6_END") = !END_TIME
                            If Trim(madrsFormatRptData.Fields("EXCEPTION6_START")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION6_START") = "00:00"
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION6_END")) = "00:00" Then
                                madrsFormatRptData.Fields("EXCEPTION6_END") = "24:00"
                            End If
                            mlngExceptionCount = mlngExceptionCount + 1
                            If OvertimeExcep(Trim(madrsFormatRptData.Fields("EXCEPTION6"))) _
                                And !OT_OPP_NUM > 0 Then
                                lintTotalStartMin = _
                                (Format(Mid(madrsFormatRptData.Fields("EXCEPTION6_START"), 1, 2), "00") * 60) + _
                                 Format(Mid(madrsFormatRptData.Fields("EXCEPTION6_START"), 4, 2), "00")
                                lintTotalEndMin = _
                                (Format(Mid(madrsFormatRptData.Fields("EXCEPTION6_END"), 1, 2), "00") * 60) + _
                                 Format(Mid(madrsFormatRptData.Fields("EXCEPTION6_END"), 4, 2), "00")
                                madrsFormatRptData.Fields("CANVASSED_OT_HOURS6") = _
                                CStr(Abs(((lintTotalEndMin - lintTotalStartMin) / 60)))
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION6_START")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION6_START") = "00:00"
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION6_END")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION6_END") = "00:00"
                            End If
                            If !OT_OPP_NUM > 0 Then
                                Set ladrsOTShift = mobjTimeCard.GetOTShiftCode(!OT_OPP_NUM)
                                If mADOExtensions.RecordsExist(ladrsOTShift) Then
                                    madrsFormatRptData.Fields("CANVASSED_OT_SHIFT6") = _
                                    ladrsOTShift.Fields("SHIFT_CD")
                                End If
                            End If
                        Case 6
                            madrsFormatRptData.Fields("EXCEPTION7_DESC") = _
                            mobjCodeTables.ExceptionDescription(madrsFormatRptData.Fields("EXCEPTION7"))
                            madrsFormatRptData.Fields("EXCEPTION7_START") = !START_TIME
                            madrsFormatRptData.Fields("EXCEPTION7_END") = !END_TIME
                            If Trim(madrsFormatRptData.Fields("EXCEPTION7_START")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION7_START") = "00:00"
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION7_END")) = "00:00" Then
                                madrsFormatRptData.Fields("EXCEPTION7_END") = "24:00"
                            End If
                            mlngExceptionCount = mlngExceptionCount + 1
                            If OvertimeExcep(Trim(madrsFormatRptData.Fields("EXCEPTION7"))) _
                                And !OT_OPP_NUM > 0 Then
                                lintTotalStartMin = _
                                (Format(Mid(madrsFormatRptData.Fields("EXCEPTION7_START"), 1, 2), "00") * 60) + _
                                 Format(Mid(madrsFormatRptData.Fields("EXCEPTION7_START"), 4, 2), "00")
                                lintTotalEndMin = _
                                (Format(Mid(madrsFormatRptData.Fields("EXCEPTION7_END"), 1, 2), "00") * 60) + _
                                 Format(Mid(madrsFormatRptData.Fields("EXCEPTION7_END"), 4, 2), "00")
                                madrsFormatRptData.Fields("CANVASSED_OT_HOURS7") = _
                                CStr(Abs(((lintTotalEndMin - lintTotalStartMin) / 60)))
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION7_START")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION7_START") = "00:00"
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION7_END")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION7_END") = "00:00"
                            End If
                            If !OT_OPP_NUM > 0 Then
                                Set ladrsOTShift = mobjTimeCard.GetOTShiftCode(!OT_OPP_NUM)
                                If mADOExtensions.RecordsExist(ladrsOTShift) Then
                                    madrsFormatRptData.Fields("CANVASSED_OT_SHIFT7") = _
                                    ladrsOTShift.Fields("SHIFT_CD")
                                End If
                            End If
                        Case 7
                            madrsFormatRptData.Fields("EXCEPTION8_DESC") = _
                            mobjCodeTables.ExceptionDescription(madrsFormatRptData.Fields("EXCEPTION8"))
                            madrsFormatRptData.Fields("EXCEPTION8_START") = !START_TIME
                            madrsFormatRptData.Fields("EXCEPTION8_END") = !END_TIME
                            If Trim(madrsFormatRptData.Fields("EXCEPTION8_START")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION8_START") = "00:00"
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION8_END")) = "00:00" Then
                                madrsFormatRptData.Fields("EXCEPTION8_END") = "24:00"
                            End If
                            mlngExceptionCount = mlngExceptionCount + 1
                            If OvertimeExcep(Trim(madrsFormatRptData.Fields("EXCEPTION8"))) _
                                And !OT_OPP_NUM > 0 Then
                                lintTotalStartMin = _
                                (Format(Mid(madrsFormatRptData.Fields("EXCEPTION8_START"), 1, 2), "00") * 60) + _
                                 Format(Mid(madrsFormatRptData.Fields("EXCEPTION8_START"), 4, 2), "00")
                                lintTotalEndMin = _
                                (Format(Mid(madrsFormatRptData.Fields("EXCEPTION8_END"), 1, 2), "00") * 60) + _
                                 Format(Mid(madrsFormatRptData.Fields("EXCEPTION8_END"), 4, 2), "00")
                                madrsFormatRptData.Fields("CANVASSED_OT_HOURS8") = _
                                CStr(Abs(((lintTotalEndMin - lintTotalStartMin) / 60)))
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION8_START")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION8_START") = "00:00"
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION8_END")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION8_END") = "00:00"
                            End If
                            If !OT_OPP_NUM > 0 Then
                                Set ladrsOTShift = mobjTimeCard.GetOTShiftCode(!OT_OPP_NUM)
                                If mADOExtensions.RecordsExist(ladrsOTShift) Then
                                    madrsFormatRptData.Fields("CANVASSED_OT_SHIFT8") = _
                                    ladrsOTShift.Fields("SHIFT_CD")
                                End If
                            End If
                        Case 8
                            madrsFormatRptData.Fields("EXCEPTION9_DESC") = _
                            mobjCodeTables.ExceptionDescription(madrsFormatRptData.Fields("EXCEPTION9"))
                            madrsFormatRptData.Fields("EXCEPTION9_START") = !START_TIME
                            madrsFormatRptData.Fields("EXCEPTION9_END") = !END_TIME
                            If Trim(madrsFormatRptData.Fields("EXCEPTION9_START")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION9_START") = "00:00"
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION9_END")) = "00:00" Then
                                madrsFormatRptData.Fields("EXCEPTION9_END") = "24:00"
                            End If
                            mlngExceptionCount = mlngExceptionCount + 1
                            If OvertimeExcep(Trim(madrsFormatRptData.Fields("EXCEPTION9"))) _
                                And !OT_OPP_NUM > 0 Then
                                lintTotalStartMin = _
                                (Format(Mid(madrsFormatRptData.Fields("EXCEPTION9_START"), 1, 2), "00") * 60) + _
                                 Format(Mid(madrsFormatRptData.Fields("EXCEPTION9_START"), 4, 2), "00")
                                lintTotalEndMin = _
                                (Format(Mid(madrsFormatRptData.Fields("EXCEPTION9_END"), 1, 2), "00") * 60) + _
                                 Format(Mid(madrsFormatRptData.Fields("EXCEPTION9_END"), 4, 2), "00")
                                madrsFormatRptData.Fields("CANVASSED_OT_HOURS9") = _
                                CStr(Abs(((lintTotalEndMin - lintTotalStartMin) / 60)))
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION9_START")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION9_START") = "00:00"
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION9_END")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION9_END") = "00:00"
                            End If
                            If !OT_OPP_NUM > 0 Then
                                Set ladrsOTShift = mobjTimeCard.GetOTShiftCode(!OT_OPP_NUM)
                                If mADOExtensions.RecordsExist(ladrsOTShift) Then
                                    madrsFormatRptData.Fields("CANVASSED_OT_SHIFT9") = _
                                    ladrsOTShift.Fields("SHIFT_CD")
                                End If
                            End If
                        Case 9
                            madrsFormatRptData.Fields("EXCEPTION10_DESC") = _
                            mobjCodeTables.ExceptionDescription(madrsFormatRptData.Fields("EXCEPTION10"))
                            madrsFormatRptData.Fields("EXCEPTION10_START") = !START_TIME
                            madrsFormatRptData.Fields("EXCEPTION10_END") = !END_TIME
                            If Trim(madrsFormatRptData.Fields("EXCEPTION10_START")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION10_START") = "00:00"
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION10_END")) = "00:00" Then
                                madrsFormatRptData.Fields("EXCEPTION10_END") = "24:00"
                            End If
                            mlngExceptionCount = mlngExceptionCount + 1
                            If OvertimeExcep(Trim(madrsFormatRptData.Fields("EXCEPTION10"))) _
                                And !OT_OPP_NUM > 0 Then
                                lintTotalStartMin = _
                                (Format(Mid(madrsFormatRptData.Fields("EXCEPTION10_START"), 1, 2), "00") * 60) + _
                                 Format(Mid(madrsFormatRptData.Fields("EXCEPTION10_START"), 4, 2), "00")
                                lintTotalEndMin = _
                                (Format(Mid(madrsFormatRptData.Fields("EXCEPTION10_END"), 1, 2), "00") * 60) + _
                                 Format(Mid(madrsFormatRptData.Fields("EXCEPTION10_END"), 4, 2), "00")
                                madrsFormatRptData.Fields("CANVASSED_OT_HOURS10") = _
                                CStr(Abs(((lintTotalEndMin - lintTotalStartMin) / 60)))
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION10_START")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION10_START") = "00:00"
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION10_END")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION10_END") = "00:00"
                            End If
                            If !OT_OPP_NUM > 0 Then
                                Set ladrsOTShift = mobjTimeCard.GetOTShiftCode(!OT_OPP_NUM)
                                If mADOExtensions.RecordsExist(ladrsOTShift) Then
                                    madrsFormatRptData.Fields("CANVASSED_OT_SHIFT10") = _
                                    ladrsOTShift.Fields("SHIFT_CD")
                                End If
                            End If
                        Case 10
                            madrsFormatRptData.Fields("EXCEPTION11_DESC") = _
                            mobjCodeTables.ExceptionDescription(madrsFormatRptData.Fields("EXCEPTION11"))
                            madrsFormatRptData.Fields("EXCEPTION11_START") = !START_TIME
                            madrsFormatRptData.Fields("EXCEPTION11_END") = !END_TIME
                            If Trim(madrsFormatRptData.Fields("EXCEPTION11_START")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION11_START") = "00:00"
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION11_END")) = "00:00" Then
                                madrsFormatRptData.Fields("EXCEPTION11_END") = "24:00"
                            End If
                            mlngExceptionCount = mlngExceptionCount + 1
                            If OvertimeExcep(Trim(madrsFormatRptData.Fields("EXCEPTION11"))) _
                                And !OT_OPP_NUM > 0 Then
                                lintTotalStartMin = _
                                (Format(Mid(madrsFormatRptData.Fields("EXCEPTION11_START"), 1, 2), "00") * 60) + _
                                 Format(Mid(madrsFormatRptData.Fields("EXCEPTION11_START"), 4, 2), "00")
                                lintTotalEndMin = _
                                (Format(Mid(madrsFormatRptData.Fields("EXCEPTION11_END"), 1, 2), "00") * 60) + _
                                 Format(Mid(madrsFormatRptData.Fields("EXCEPTION11_END"), 4, 2), "00")
                                madrsFormatRptData.Fields("CANVASSED_OT_HOURS11") = _
                                CStr(Abs(((lintTotalEndMin - lintTotalStartMin) / 60)))
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION11_START")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION11_START") = "00:00"
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION11_END")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION11_END") = "00:00"
                            End If
                            If !OT_OPP_NUM > 0 Then
                                Set ladrsOTShift = mobjTimeCard.GetOTShiftCode(!OT_OPP_NUM)
                                If mADOExtensions.RecordsExist(ladrsOTShift) Then
                                    madrsFormatRptData.Fields("CANVASSED_OT_SHIFT11") = _
                                    ladrsOTShift.Fields("SHIFT_CD")
                                End If
                            End If
                        Case 11
                            madrsFormatRptData.Fields("EXCEPTION12_DESC") = _
                            mobjCodeTables.ExceptionDescription(madrsFormatRptData.Fields("EXCEPTION12"))
                            madrsFormatRptData.Fields("EXCEPTION12_START") = !START_TIME
                            madrsFormatRptData.Fields("EXCEPTION12_END") = !END_TIME
                            If Trim(madrsFormatRptData.Fields("EXCEPTION12_START")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION12_START") = "00:00"
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION12_END")) = "00:00" Then
                                madrsFormatRptData.Fields("EXCEPTION12_END") = "24:00"
                            End If
                            mlngExceptionCount = mlngExceptionCount + 1
                            If OvertimeExcep(Trim(madrsFormatRptData.Fields("EXCEPTION12"))) _
                                And !OT_OPP_NUM > 0 Then
                                lintTotalStartMin = _
                                (Format(Mid(madrsFormatRptData.Fields("EXCEPTION12_START"), 1, 2), "00") * 60) + _
                                 Format(Mid(madrsFormatRptData.Fields("EXCEPTION12_START"), 4, 2), "00")
                                lintTotalEndMin = _
                                (Format(Mid(madrsFormatRptData.Fields("EXCEPTION12_END"), 1, 2), "00") * 60) + _
                                 Format(Mid(madrsFormatRptData.Fields("EXCEPTION12_END"), 4, 2), "00")
                                madrsFormatRptData.Fields("CANVASSED_OT_HOURS12") = _
                                CStr(Abs(((lintTotalEndMin - lintTotalStartMin) / 60)))
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION12_START")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION12_START") = "00:00"
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION12_END")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION12_END") = "00:00"
                            End If
                            If !OT_OPP_NUM > 0 Then
                                Set ladrsOTShift = mobjTimeCard.GetOTShiftCode(!OT_OPP_NUM)
                                If mADOExtensions.RecordsExist(ladrsOTShift) Then
                                    madrsFormatRptData.Fields("CANVASSED_OT_SHIFT12") = _
                                    ladrsOTShift.Fields("SHIFT_CD")
                                End If
                            End If
                        Case 12
                            madrsFormatRptData.Fields("EXCEPTION13_DESC") = _
                            mobjCodeTables.ExceptionDescription(madrsFormatRptData.Fields("EXCEPTION13"))
                            madrsFormatRptData.Fields("EXCEPTION13_START") = !START_TIME
                            madrsFormatRptData.Fields("EXCEPTION13_END") = !END_TIME
                            If Trim(madrsFormatRptData.Fields("EXCEPTION13_START")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION13_START") = "00:00"
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION13_END")) = "00:00" Then
                                madrsFormatRptData.Fields("EXCEPTION13_END") = "24:00"
                            End If
                            mlngExceptionCount = mlngExceptionCount + 1
                            If OvertimeExcep(Trim(madrsFormatRptData.Fields("EXCEPTION13"))) _
                                And !OT_OPP_NUM > 0 Then
                                lintTotalStartMin = _
                                (Format(Mid(madrsFormatRptData.Fields("EXCEPTION13_START"), 1, 2), "00") * 60) + _
                                 Format(Mid(madrsFormatRptData.Fields("EXCEPTION13_START"), 4, 2), "00")
                                lintTotalEndMin = _
                                (Format(Mid(madrsFormatRptData.Fields("EXCEPTION13_END"), 1, 2), "00") * 60) + _
                                 Format(Mid(madrsFormatRptData.Fields("EXCEPTION13_END"), 4, 2), "00")
                                madrsFormatRptData.Fields("CANVASSED_OT_HOURS13") = _
                                CStr(Abs(((lintTotalEndMin - lintTotalStartMin) / 60)))
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION13_START")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION13_START") = "00:00"
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION13_END")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION13_END") = "00:00"
                            End If
                            If !OT_OPP_NUM > 0 Then
                                Set ladrsOTShift = mobjTimeCard.GetOTShiftCode(!OT_OPP_NUM)
                                If mADOExtensions.RecordsExist(ladrsOTShift) Then
                                    madrsFormatRptData.Fields("CANVASSED_OT_SHIFT13") = _
                                    ladrsOTShift.Fields("SHIFT_CD")
                                End If
                            End If
                        Case 13
                            madrsFormatRptData.Fields("EXCEPTION14_DESC") = _
                            mobjCodeTables.ExceptionDescription(madrsFormatRptData.Fields("EXCEPTION14"))
                            madrsFormatRptData.Fields("EXCEPTION14_START") = !START_TIME
                            madrsFormatRptData.Fields("EXCEPTION14_END") = !END_TIME
                            If Trim(madrsFormatRptData.Fields("EXCEPTION14_START")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION14_START") = "00:00"
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION14_END")) = "00:00" Then
                                madrsFormatRptData.Fields("EXCEPTION14_END") = "24:00"
                            End If
                            mlngExceptionCount = mlngExceptionCount + 1
                            If OvertimeExcep(Trim(madrsFormatRptData.Fields("EXCEPTION14"))) _
                                And !OT_OPP_NUM > 0 Then
                                lintTotalStartMin = _
                                (Format(Mid(madrsFormatRptData.Fields("EXCEPTION14_START"), 1, 2), "00") * 60) + _
                                 Format(Mid(madrsFormatRptData.Fields("EXCEPTION14_START"), 4, 2), "00")
                                lintTotalEndMin = _
                                (Format(Mid(madrsFormatRptData.Fields("EXCEPTION14_END"), 1, 2), "00") * 60) + _
                                 Format(Mid(madrsFormatRptData.Fields("EXCEPTION14_END"), 4, 2), "00")
                                madrsFormatRptData.Fields("CANVASSED_OT_HOURS14") = _
                                CStr(Abs(((lintTotalEndMin - lintTotalStartMin) / 60)))
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION14_START")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION14_START") = "00:00"
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION14_END")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION14_END") = "00:00"
                            End If
                            If !OT_OPP_NUM > 0 Then
                                Set ladrsOTShift = mobjTimeCard.GetOTShiftCode(!OT_OPP_NUM)
                                If mADOExtensions.RecordsExist(ladrsOTShift) Then
                                    madrsFormatRptData.Fields("CANVASSED_OT_SHIFT14") = _
                                    ladrsOTShift.Fields("SHIFT_CD")
                                End If
                            End If
                        Case 14
                            madrsFormatRptData.Fields("EXCEPTION15_DESC") = _
                            mobjCodeTables.ExceptionDescription(madrsFormatRptData.Fields("EXCEPTION15"))
                            madrsFormatRptData.Fields("EXCEPTION15_START") = !START_TIME
                            madrsFormatRptData.Fields("EXCEPTION15_END") = !END_TIME
                            If Trim(madrsFormatRptData.Fields("EXCEPTION15_START")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION15_START") = "00:00"
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION15_END")) = "00:00" Then
                                madrsFormatRptData.Fields("EXCEPTION15_END") = "24:00"
                            End If
                            mlngExceptionCount = mlngExceptionCount + 1
                            If OvertimeExcep(Trim(madrsFormatRptData.Fields("EXCEPTION15"))) _
                                And !OT_OPP_NUM > 0 Then
                                lintTotalStartMin = _
                                (Format(Mid(madrsFormatRptData.Fields("EXCEPTION15_START"), 1, 2), "00") * 60) + _
                                 Format(Mid(madrsFormatRptData.Fields("EXCEPTION15_START"), 4, 2), "00")
                                lintTotalEndMin = _
                                (Format(Mid(madrsFormatRptData.Fields("EXCEPTION15_END"), 1, 2), "00") * 60) + _
                                 Format(Mid(madrsFormatRptData.Fields("EXCEPTION15_END"), 4, 2), "00")
                                madrsFormatRptData.Fields("CANVASSED_OT_HOURS15") = _
                                CStr(Abs(((lintTotalEndMin - lintTotalStartMin) / 60)))
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION15_START")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION15_START") = "00:00"
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION15_END")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION15_END") = "00:00"
                            End If
                            If !OT_OPP_NUM > 0 Then
                                Set ladrsOTShift = mobjTimeCard.GetOTShiftCode(!OT_OPP_NUM)
                                If mADOExtensions.RecordsExist(ladrsOTShift) Then
                                    madrsFormatRptData.Fields("CANVASSED_OT_SHIFT15") = _
                                    ladrsOTShift.Fields("SHIFT_CD")
                                End If
                            End If
                    End Select
                End If
            End If
        End With
    End If
    
    Exit Sub
PopulateWG_Record_EH:
    
    'Raise Error depending on the Error Number
    TrsRaiseError MODULENAME _
            , "PopulateWG_Record"
End Sub
Private Sub PopulateDept_Record()
On Error GoTo PopulateDept_Record_EH

    Dim llngQualsIndex As Long
    Dim llngSkillsIndex As Long
    Dim ladrsOTShift As ADODB.Recordset
    Dim lblnSecondaryWG As Boolean
    Dim lintTotalStartMin As Integer
    Dim lintTotalEndMin As Integer
    Dim lblnProcessExcep As Boolean
    
    If mADOExtensions.RecordsExist(madrsEmpsinDept) Then
        With madrsRosterData
            If mlngPrevEmp <> !EMP_NUM Or (mlngPrevEmp = !EMP_NUM And _
                                           mdtPrevDate <> !CHARGE_DATE) Then
                If Not mblnFirstRecord Then
                    If mADOExtensions.RecordsExist(madrsEmployeeInelig) Then
                        madrsEmployeeInelig.MoveFirst
                        'madrsEmployeeInelig.Filter = "EMP_NUM = " & mlngPrevEmp & _
                            " AND EFF_START_DATE <= #" & marrRequestDate(mlngDateIndex) & _
                            "# AND " & " EFF_END_DATE >= #" &  (mlngDateIndex) & "#"
                        madrsEmployeeInelig.Filter = "EMP_NUM = " & mlngPrevEmp & _
                            " AND EFF_START_DATE <= #" & mdtPrevDate & _
                            "# AND " & " EFF_END_DATE >= #" & mdtPrevDate & "#"
                        If Not madrsEmployeeInelig.EOF Then
                            madrsEmployeeInelig.MoveFirst
                            Select Case mlngExceptionCount
                                Case 0
                                    madrsFormatRptData.Fields("EXCEPTION1") = _
                                        madrsEmployeeInelig.Fields("EXCEPTION_CD")
                                    madrsFormatRptData.Fields("EXCEPTION1_DESC") = _
                                        mobjCodeTables.ExceptionDescription( _
                                        madrsFormatRptData.Fields("EXCEPTION1"))
                                    mlngExceptionCount = mlngExceptionCount + 1
                                    madrsFormatRptData.Fields("INELIGIBILITY_START1") = _
                                        madrsEmployeeInelig.Fields("EFF_START_DATE")
                                    madrsFormatRptData.Fields("INELIGIBILITY_END1") = _
                                        madrsEmployeeInelig.Fields("EFF_END_DATE")
                                    madrsFormatRptData.Fields("EXCEPTION1_START") = _
                                        madrsEmployeeInelig.Fields("EFF_START_TIME")
                                    madrsFormatRptData.Fields("EXCEPTION1_END") = _
                                        madrsEmployeeInelig.Fields("EFF_END_TIME")
                                    madrsFormatRptData.Update
                                Case 1
                                    madrsFormatRptData.Fields("EXCEPTION2") = _
                                        madrsEmployeeInelig.Fields("EXCEPTION_CD")
                                    madrsFormatRptData.Fields("EXCEPTION2_DESC") = _
                                        mobjCodeTables.ExceptionDescription( _
                                        madrsFormatRptData.Fields("EXCEPTION2"))
                                    mlngExceptionCount = mlngExceptionCount + 1
                                    madrsFormatRptData.Fields("INELIGIBILITY_START2") = _
                                        madrsEmployeeInelig.Fields("EFF_START_DATE")
                                    madrsFormatRptData.Fields("INELIGIBILITY_END2") = _
                                        madrsEmployeeInelig.Fields("EFF_END_DATE")
                                    madrsFormatRptData.Fields("EXCEPTION2_START") = _
                                        madrsEmployeeInelig.Fields("EFF_START_TIME")
                                    madrsFormatRptData.Fields("EXCEPTION2_END") = _
                                        madrsEmployeeInelig.Fields("EFF_END_TIME")
                                    madrsFormatRptData.Update
                                Case 2
                                    madrsFormatRptData.Fields("EXCEPTION3") = _
                                        madrsEmployeeInelig.Fields("EXCEPTION_CD")
                                    madrsFormatRptData.Fields("EXCEPTION3_DESC") = _
                                        mobjCodeTables.ExceptionDescription( _
                                        madrsFormatRptData.Fields("EXCEPTION3"))
                                    mlngExceptionCount = mlngExceptionCount + 1
                                    madrsFormatRptData.Fields("INELIGIBILITY_START3") = _
                                        madrsEmployeeInelig.Fields("EFF_START_DATE")
                                    madrsFormatRptData.Fields("INELIGIBILITY_END3") = _
                                        madrsEmployeeInelig.Fields("EFF_END_DATE")
                                    madrsFormatRptData.Fields("EXCEPTION3_START") = _
                                        madrsEmployeeInelig.Fields("EFF_START_TIME")
                                    madrsFormatRptData.Fields("EXCEPTION3_END") = _
                                        madrsEmployeeInelig.Fields("EFF_END_TIME")
                                    madrsFormatRptData.Update
                                Case 3
                                    madrsFormatRptData.Fields("EXCEPTION4") = _
                                        madrsEmployeeInelig.Fields("EXCEPTION_CD")
                                    madrsFormatRptData.Fields("EXCEPTION4_DESC") = _
                                        mobjCodeTables.ExceptionDescription( _
                                        madrsFormatRptData.Fields("EXCEPTION4"))
                                    mlngExceptionCount = mlngExceptionCount + 1
                                    madrsFormatRptData.Fields("INELIGIBILITY_START4") = _
                                        madrsEmployeeInelig.Fields("EFF_START_DATE")
                                    madrsFormatRptData.Fields("INELIGIBILITY_END4") = _
                                        madrsEmployeeInelig.Fields("EFF_END_DATE")
                                    madrsFormatRptData.Fields("EXCEPTION4_START") = _
                                        madrsEmployeeInelig.Fields("EFF_START_TIME")
                                    madrsFormatRptData.Fields("EXCEPTION4_END") = _
                                        madrsEmployeeInelig.Fields("EFF_END_TIME")
                                    madrsFormatRptData.Update
                                Case 4
                                    madrsFormatRptData.Fields("EXCEPTION5") = _
                                        madrsEmployeeInelig.Fields("EXCEPTION_CD")
                                    madrsFormatRptData.Fields("EXCEPTION5_DESC") = _
                                        mobjCodeTables.ExceptionDescription( _
                                        madrsFormatRptData.Fields("EXCEPTION5"))
                                    mlngExceptionCount = mlngExceptionCount + 1
                                    madrsFormatRptData.Fields("INELIGIBILITY_START5") = _
                                        madrsEmployeeInelig.Fields("EFF_START_DATE")
                                    madrsFormatRptData.Fields("INELIGIBILITY_END5") = _
                                        madrsEmployeeInelig.Fields("EFF_END_DATE")
                                    madrsFormatRptData.Fields("EXCEPTION5_START") = _
                                        madrsEmployeeInelig.Fields("EFF_START_TIME")
                                    madrsFormatRptData.Fields("EXCEPTION5_END") = _
                                        madrsEmployeeInelig.Fields("EFF_END_TIME")
                                    madrsFormatRptData.Update
                                Case 5
                                    madrsFormatRptData.Fields("EXCEPTION6") = _
                                        madrsEmployeeInelig.Fields("EXCEPTION_CD")
                                    madrsFormatRptData.Fields("EXCEPTION6_DESC") = _
                                        mobjCodeTables.ExceptionDescription( _
                                        madrsFormatRptData.Fields("EXCEPTION6"))
                                    mlngExceptionCount = mlngExceptionCount + 1
                                    madrsFormatRptData.Fields("INELIGIBILITY_START6") = _
                                        madrsEmployeeInelig.Fields("EFF_START_DATE")
                                    madrsFormatRptData.Fields("INELIGIBILITY_END6") = _
                                        madrsEmployeeInelig.Fields("EFF_END_DATE")
                                    madrsFormatRptData.Fields("EXCEPTION6_START") = _
                                        madrsEmployeeInelig.Fields("EFF_START_TIME")
                                    madrsFormatRptData.Fields("EXCEPTION6_END") = _
                                        madrsEmployeeInelig.Fields("EFF_END_TIME")
                                    madrsFormatRptData.Update
                                Case 6
                                    madrsFormatRptData.Fields("EXCEPTION7") = _
                                        madrsEmployeeInelig.Fields("EXCEPTION_CD")
                                    madrsFormatRptData.Fields("EXCEPTION7_DESC") = _
                                        mobjCodeTables.ExceptionDescription( _
                                        madrsFormatRptData.Fields("EXCEPTION7"))
                                    mlngExceptionCount = mlngExceptionCount + 1
                                    madrsFormatRptData.Fields("INELIGIBILITY_START7") = _
                                        madrsEmployeeInelig.Fields("EFF_START_DATE")
                                    madrsFormatRptData.Fields("INELIGIBILITY_END7") = _
                                        madrsEmployeeInelig.Fields("EFF_END_DATE")
                                    madrsFormatRptData.Fields("EXCEPTION7_START") = _
                                        madrsEmployeeInelig.Fields("EFF_START_TIME")
                                    madrsFormatRptData.Fields("EXCEPTION7_END") = _
                                        madrsEmployeeInelig.Fields("EFF_END_TIME")
                                    madrsFormatRptData.Update
                                Case 7
                                    madrsFormatRptData.Fields("EXCEPTION8") = _
                                        madrsEmployeeInelig.Fields("EXCEPTION_CD")
                                    madrsFormatRptData.Fields("EXCEPTION8_DESC") = _
                                        mobjCodeTables.ExceptionDescription( _
                                        madrsFormatRptData.Fields("EXCEPTION8"))
                                    mlngExceptionCount = mlngExceptionCount + 1
                                    madrsFormatRptData.Fields("INELIGIBILITY_START8") = _
                                        madrsEmployeeInelig.Fields("EFF_START_DATE")
                                    madrsFormatRptData.Fields("INELIGIBILITY_END8") = _
                                        madrsEmployeeInelig.Fields("EFF_END_DATE")
                                    madrsFormatRptData.Fields("EXCEPTION8_START") = _
                                        madrsEmployeeInelig.Fields("EFF_START_TIME")
                                    madrsFormatRptData.Fields("EXCEPTION8_END") = _
                                        madrsEmployeeInelig.Fields("EFF_END_TIME")
                                    madrsFormatRptData.Update
                                Case 8
                                    madrsFormatRptData.Fields("EXCEPTION9") = _
                                        madrsEmployeeInelig.Fields("EXCEPTION_CD")
                                    madrsFormatRptData.Fields("EXCEPTION9_DESC") = _
                                        mobjCodeTables.ExceptionDescription( _
                                        madrsFormatRptData.Fields("EXCEPTION9"))
                                    mlngExceptionCount = mlngExceptionCount + 1
                                    madrsFormatRptData.Fields("INELIGIBILITY_START9") = _
                                        madrsEmployeeInelig.Fields("EFF_START_DATE")
                                    madrsFormatRptData.Fields("INELIGIBILITY_END9") = _
                                        madrsEmployeeInelig.Fields("EFF_END_DATE")
                                    madrsFormatRptData.Fields("EXCEPTION9_START") = _
                                        madrsEmployeeInelig.Fields("EFF_START_TIME")
                                    madrsFormatRptData.Fields("EXCEPTION9_END") = _
                                        madrsEmployeeInelig.Fields("EFF_END_TIME")
                                    madrsFormatRptData.Update
                                Case 9
                                    madrsFormatRptData.Fields("EXCEPTION10") = _
                                        madrsEmployeeInelig.Fields("EXCEPTION_CD")
                                    madrsFormatRptData.Fields("EXCEPTION10_DESC") = _
                                        mobjCodeTables.ExceptionDescription( _
                                        madrsFormatRptData.Fields("EXCEPTION10"))
                                    mlngExceptionCount = mlngExceptionCount + 1
                                    madrsFormatRptData.Fields("INELIGIBILITY_START10") = _
                                        madrsEmployeeInelig.Fields("EFF_START_DATE")
                                    madrsFormatRptData.Fields("INELIGIBILITY_END10") = _
                                        madrsEmployeeInelig.Fields("EFF_END_DATE")
                                    madrsFormatRptData.Fields("EXCEPTION10_START") = _
                                        madrsEmployeeInelig.Fields("EFF_START_TIME")
                                    madrsFormatRptData.Fields("EXCEPTION10_END") = _
                                        madrsEmployeeInelig.Fields("EFF_END_TIME")
                                    madrsFormatRptData.Update
                                Case 10
                                    madrsFormatRptData.Fields("EXCEPTION11") = _
                                        madrsEmployeeInelig.Fields("EXCEPTION_CD")
                                    madrsFormatRptData.Fields("EXCEPTION11_DESC") = _
                                        mobjCodeTables.ExceptionDescription( _
                                        madrsFormatRptData.Fields("EXCEPTION11"))
                                    mlngExceptionCount = mlngExceptionCount + 1
                                    madrsFormatRptData.Fields("INELIGIBILITY_START11") = _
                                        madrsEmployeeInelig.Fields("EFF_START_DATE")
                                    madrsFormatRptData.Fields("INELIGIBILITY_END11") = _
                                        madrsEmployeeInelig.Fields("EFF_END_DATE")
                                    madrsFormatRptData.Fields("EXCEPTION11_START") = _
                                        madrsEmployeeInelig.Fields("EFF_START_TIME")
                                    madrsFormatRptData.Fields("EXCEPTION11_END") = _
                                        madrsEmployeeInelig.Fields("EFF_END_TIME")
                                    madrsFormatRptData.Update
                                Case 11
                                    madrsFormatRptData.Fields("EXCEPTION12") = _
                                        madrsEmployeeInelig.Fields("EXCEPTION_CD")
                                    madrsFormatRptData.Fields("EXCEPTION12_DESC") = _
                                        mobjCodeTables.ExceptionDescription( _
                                        madrsFormatRptData.Fields("EXCEPTION12"))
                                    mlngExceptionCount = mlngExceptionCount + 1
                                    madrsFormatRptData.Fields("INELIGIBILITY_START12") = _
                                        madrsEmployeeInelig.Fields("EFF_START_DATE")
                                    madrsFormatRptData.Fields("INELIGIBILITY_END12") = _
                                        madrsEmployeeInelig.Fields("EFF_END_DATE")
                                    madrsFormatRptData.Fields("EXCEPTION12_START") = _
                                        madrsEmployeeInelig.Fields("EFF_START_TIME")
                                    madrsFormatRptData.Fields("EXCEPTION12_END") = _
                                        madrsEmployeeInelig.Fields("EFF_END_TIME")
                                    madrsFormatRptData.Update
                                Case 12
                                    madrsFormatRptData.Fields("EXCEPTION13") = _
                                        madrsEmployeeInelig.Fields("EXCEPTION_CD")
                                    madrsFormatRptData.Fields("EXCEPTION13_DESC") = _
                                        mobjCodeTables.ExceptionDescription( _
                                        madrsFormatRptData.Fields("EXCEPTION13"))
                                    mlngExceptionCount = mlngExceptionCount + 1
                                    madrsFormatRptData.Fields("INELIGIBILITY_START13") = _
                                        madrsEmployeeInelig.Fields("EFF_START_DATE")
                                    madrsFormatRptData.Fields("INELIGIBILITY_END13") = _
                                        madrsEmployeeInelig.Fields("EFF_END_DATE")
                                    madrsFormatRptData.Fields("EXCEPTION13_START") = _
                                        madrsEmployeeInelig.Fields("EFF_START_TIME")
                                    madrsFormatRptData.Fields("EXCEPTION13_END") = _
                                        madrsEmployeeInelig.Fields("EFF_END_TIME")
                                    madrsFormatRptData.Update
                                Case 13
                                    madrsFormatRptData.Fields("EXCEPTION14") = _
                                        madrsEmployeeInelig.Fields("EXCEPTION_CD")
                                    madrsFormatRptData.Fields("EXCEPTION14_DESC") = _
                                        mobjCodeTables.ExceptionDescription( _
                                        madrsFormatRptData.Fields("EXCEPTION14"))
                                    mlngExceptionCount = mlngExceptionCount + 1
                                    madrsFormatRptData.Fields("INELIGIBILITY_START14") = _
                                        madrsEmployeeInelig.Fields("EFF_START_DATE")
                                    madrsFormatRptData.Fields("INELIGIBILITY_END14") = _
                                        madrsEmployeeInelig.Fields("EFF_END_DATE")
                                    madrsFormatRptData.Fields("EXCEPTION14_START") = _
                                        madrsEmployeeInelig.Fields("EFF_START_TIME")
                                    madrsFormatRptData.Fields("EXCEPTION14_END") = _
                                        madrsEmployeeInelig.Fields("EFF_END_TIME")
                                    madrsFormatRptData.Update
                                Case 14
                                    madrsFormatRptData.Fields("EXCEPTION15") = _
                                        madrsEmployeeInelig.Fields("EXCEPTION_CD")
                                    madrsFormatRptData.Fields("EXCEPTION15_DESC") = _
                                        mobjCodeTables.ExceptionDescription( _
                                        madrsFormatRptData.Fields("EXCEPTION15"))
                                    mlngExceptionCount = mlngExceptionCount + 1
                                    madrsFormatRptData.Fields("INELIGIBILITY_START15") = _
                                        madrsEmployeeInelig.Fields("EFF_START_DATE")
                                    madrsFormatRptData.Fields("INELIGIBILITY_END15") = _
                                        madrsEmployeeInelig.Fields("EFF_END_DATE")
                                    madrsFormatRptData.Fields("EXCEPTION15_START") = _
                                        madrsEmployeeInelig.Fields("EFF_START_TIME")
                                    madrsFormatRptData.Fields("EXCEPTION15_END") = _
                                        madrsEmployeeInelig.Fields("EFF_END_TIME")
                                    madrsFormatRptData.Update
                            End Select
                        End If
                        madrsEmployeeInelig.Filter = adFilterNone
                    End If
                End If
                madrsFormatRptData.AddNew
                mblnFirstRecord = False
                madrsEmpsinDept.MoveFirst
                madrsEmpsinDept.Find "EMP_NUM = " & !EMP_NUM
                madrsFormatRptData.Fields("DEPT_NUM") = marrDepartments(mlngDeptIndex)
                madrsFormatRptData.Fields("DEPT_DESC") = marrDepartmentDesc(mlngDeptIndex)
                madrsFormatRptData.Fields("WG_NUM") = marrWorkGroups(mlngEmpIndex)
                madrsFormatRptData.Fields("WG_DESC") = marrWorkGroupDesc(mlngEmpIndex)
                madrsFormatRptData.Fields("EMP_NUM") = !EMP_NUM
                madrsFormatRptData.Fields("EMP_NAME") = madrsEmpsinDept.Fields("EMP_NAME")
                'lblnSecondaryWG = mobjEmployee.IsSecondaryWorkGroup(!EMP_NUM, _
                            madrsEmpsinDept.Fields("WG_NUM"), marrRequestDate(mlngDateIndex))
                lblnSecondaryWG = mobjEmployee.IsSecondaryWorkGroup(!EMP_NUM, _
                            madrsEmpsinDept.Fields("WG_NUM"), !CHARGE_DATE)
                If lblnSecondaryWG = True Then
                    madrsFormatRptData.Fields("SECONDARY_IND") = "*"
                Else
                    madrsFormatRptData.Fields("SECONDARY_IND") = Null
                End If
                If mblnPrintQualsSkills Then
                    If mADOExtensions.RecordsExist(madrsQuals) Then
                        madrsQuals.MoveFirst
                        madrsQuals.Filter = "EMP_NUM = " & !EMP_NUM
                        If Not madrsQuals.EOF Then
                            madrsQuals.MoveFirst
                            llngQualsIndex = 1
                            While llngQualsIndex < 19 And Not madrsQuals.EOF
                                If llngQualsIndex > 1 Then
                                    madrsFormatRptData.Fields("QUALS") = _
                                        madrsFormatRptData.Fields("QUALS") & ", " & _
                                        madrsQuals.Fields("QUAL_CD")
                                Else
                                    madrsFormatRptData.Fields("QUALS") = _
                                        "QUALIFICATIONS: " & madrsQuals.Fields("QUAL_CD")
                                End If
                                llngQualsIndex = llngQualsIndex + 1
                                madrsQuals.MoveNext
                            Wend
                        Else
                            madrsFormatRptData.Fields("QUALS") = Null
                        End If
                    Else
                        madrsFormatRptData.Fields("QUALS") = Null
                    End If
                    madrsQuals.Filter = adFilterNone
                    If mADOExtensions.RecordsExist(madrsSkills) Then
                        madrsSkills.MoveFirst
                        madrsSkills.Filter = "EMP_NUM = " & !EMP_NUM
                        If Not madrsSkills.EOF Then
                            madrsSkills.MoveFirst
                            llngSkillsIndex = 1
                            While llngSkillsIndex < 19 And Not madrsSkills.EOF
                                If llngSkillsIndex > 1 Then
                                    madrsFormatRptData.Fields("SKILLS") = _
                                        madrsFormatRptData.Fields("SKILLS") & ", " & _
                                        madrsSkills.Fields("SPEC_SKILL_CD")
                                Else
                                    madrsFormatRptData.Fields("SKILLS") = _
                                        "SKILLS: " & madrsSkills.Fields("SPEC_SKILL_CD")
                                End If
                                llngSkillsIndex = llngSkillsIndex + 1
                                madrsSkills.MoveNext
                            Wend
                        Else
                            madrsFormatRptData.Fields("SKILLS") = Null
                        End If
                    Else
                        madrsFormatRptData.Fields("SKILLS") = Null
                    End If
                    madrsSkills.Filter = adFilterNone
                Else
                    madrsFormatRptData.Fields("QUALS") = Null
                    madrsFormatRptData.Fields("SKILLS") = Null
                End If
                If madrsEmpsinDept.Fields("DEMAND_IND") = "Y" Then
                    madrsFormatRptData.Fields("LINE_NUM") = !LINE_NUM
                Else
                    madrsFormatRptData.Fields("LINE_NUM") = 0
                End If
                mlngExceptionCount = 0
            End If
            madrsFormatRptData.Fields("CHARGE_DATE") = !CHARGE_DATE
            If !SHIFT_NUM = 1 Then
                madrsFormatRptData.Fields("REG_SHIFT") = !SHIFT_CODE
                Select Case Trim(!SHIFT_CODE)
                    Case "O"
                        madrsFormatRptData.Fields("SORT_CODE") = 1
                    Case "A"
                        madrsFormatRptData.Fields("SORT_CODE") = 2
                    Case "P"
                        madrsFormatRptData.Fields("SORT_CODE") = 3
                    Case Else
                        madrsFormatRptData.Fields("SORT_CODE") = 0
                End Select
                If Trim(!SHIFT_START_TIME) = "24:00" Then
                    madrsFormatRptData.Fields("REG_SHIFT_START") = "00:00"
                Else
                    madrsFormatRptData.Fields("REG_SHIFT_START") = !SHIFT_START_TIME
                End If
                If Trim(!SHIFT_END_TIME) = "24:00" Then
                    madrsFormatRptData.Fields("REG_SHIFT_END") = "00:00"
                Else
                    madrsFormatRptData.Fields("REG_SHIFT_END") = !SHIFT_END_TIME
                End If
            End If
            Select Case !SCHED_SRC_CD
                Case "50"
                    Select Case !SYSTEM_EXC_CD
                        Case "TWS"
                            madrsFormatRptData.Fields("TRADE_RAS") = "*S*"
                        Case "TOR"
                            madrsFormatRptData.Fields("TRADE_RAS") = "*R*"
                        Case "TSR"
                            madrsFormatRptData.Fields("TRADE_RAS") = "*R*"
                        Case "RWT"
                            madrsFormatRptData.Fields("TRADE_RAS") = "*R*"
                        Case "TOA"
                            madrsFormatRptData.Fields("TRADE_RAS") = "*A*"
                        Case "TSA"
                            madrsFormatRptData.Fields("TRADE_RAS") = "*A*"
                    End Select
                Case "30"
                    Select Case !SYSTEM_EXC_CD
                        Case "TWS"
                            madrsFormatRptData.Fields("TRADE_RAS") = "*S*"
                        Case "TOR"
                            madrsFormatRptData.Fields("TRADE_RAS") = "*R*"
                        Case "TSR"
                            madrsFormatRptData.Fields("TRADE_RAS") = "*R*"
                        Case "RWT"
                            madrsFormatRptData.Fields("TRADE_RAS") = "*R*"
                        Case "TOA"
                            madrsFormatRptData.Fields("TRADE_RAS") = "*A*"
                        Case "TSA"
                            madrsFormatRptData.Fields("TRADE_RAS") = "*A*"
                    End Select
            End Select
            If !SHIFT_NUM = 2 Then
                madrsFormatRptData.Fields("TRADE_DOUBLE") = !SHIFT_CODE
            End If
            If Trim(!EXCEPTION_CD) <> "" And mlngExceptionCount < 15 Then
                lblnProcessExcep = False
                If ValidExcep(Trim(!OVER_EXC_CD)) Then
                    Select Case mlngExceptionCount
                        Case 0
                            madrsFormatRptData.Fields("EXCEPTION1") = !OVER_EXC_CD
                            lblnProcessExcep = True
                        Case 1
                            If ExcepExists(!OVER_EXC_CD) Then
                                lblnProcessExcep = False
                            Else
                                madrsFormatRptData.Fields("EXCEPTION2") = !OVER_EXC_CD
                                lblnProcessExcep = True
                            End If
                        Case 2
                            If ExcepExists(!OVER_EXC_CD) Then
                                lblnProcessExcep = False
                            Else
                                madrsFormatRptData.Fields("EXCEPTION3") = !OVER_EXC_CD
                                lblnProcessExcep = True
                            End If
                        Case 3
                            If ExcepExists(!OVER_EXC_CD) Then
                                lblnProcessExcep = False
                            Else
                                madrsFormatRptData.Fields("EXCEPTION4") = !OVER_EXC_CD
                                lblnProcessExcep = True
                            End If
                        Case 4
                            If ExcepExists(!OVER_EXC_CD) Then
                                lblnProcessExcep = False
                            Else
                                madrsFormatRptData.Fields("EXCEPTION5") = !OVER_EXC_CD
                                lblnProcessExcep = True
                            End If
                        Case 5
                            If ExcepExists(!OVER_EXC_CD) Then
                                lblnProcessExcep = False
                            Else
                                madrsFormatRptData.Fields("EXCEPTION6") = !OVER_EXC_CD
                                lblnProcessExcep = True
                            End If
                        Case 6
                            If ExcepExists(!OVER_EXC_CD) Then
                                lblnProcessExcep = False
                            Else
                                madrsFormatRptData.Fields("EXCEPTION7") = !OVER_EXC_CD
                                lblnProcessExcep = True
                            End If
                        Case 7
                            If ExcepExists(!OVER_EXC_CD) Then
                                lblnProcessExcep = False
                            Else
                                madrsFormatRptData.Fields("EXCEPTION8") = !OVER_EXC_CD
                                lblnProcessExcep = True
                            End If
                        Case 8
                            If ExcepExists(!OVER_EXC_CD) Then
                                lblnProcessExcep = False
                            Else
                                madrsFormatRptData.Fields("EXCEPTION9") = !OVER_EXC_CD
                                lblnProcessExcep = True
                            End If
                        Case 9
                            If ExcepExists(!OVER_EXC_CD) Then
                                lblnProcessExcep = False
                            Else
                                madrsFormatRptData.Fields("EXCEPTION10") = !OVER_EXC_CD
                                lblnProcessExcep = True
                            End If
                        Case 10
                            If ExcepExists(!OVER_EXC_CD) Then
                                lblnProcessExcep = False
                            Else
                                madrsFormatRptData.Fields("EXCEPTION11") = !OVER_EXC_CD
                                lblnProcessExcep = True
                            End If
                        Case 11
                            If ExcepExists(!OVER_EXC_CD) Then
                                lblnProcessExcep = False
                            Else
                                madrsFormatRptData.Fields("EXCEPTION12") = !OVER_EXC_CD
                                lblnProcessExcep = True
                            End If
                        Case 12
                            If ExcepExists(!OVER_EXC_CD) Then
                                lblnProcessExcep = False
                            Else
                                madrsFormatRptData.Fields("EXCEPTION13") = !OVER_EXC_CD
                                lblnProcessExcep = True
                            End If
                        Case 13
                            If ExcepExists(!OVER_EXC_CD) Then
                                lblnProcessExcep = False
                            Else
                                madrsFormatRptData.Fields("EXCEPTION14") = !OVER_EXC_CD
                                lblnProcessExcep = True
                            End If
                        Case 14
                            If ExcepExists(!OVER_EXC_CD) Then
                                lblnProcessExcep = False
                            Else
                                madrsFormatRptData.Fields("EXCEPTION15") = !OVER_EXC_CD
                                lblnProcessExcep = True
                            End If
                    End Select
                Else
                    If ValidExcep(Trim(!EXCEPTION_CD)) Then
                        Select Case mlngExceptionCount
                            Case 0
                                madrsFormatRptData.Fields("EXCEPTION1") = !EXCEPTION_CD
                                lblnProcessExcep = True
                            Case 1
                                If ExcepExists(!EXCEPTION_CD) Then
                                    lblnProcessExcep = False
                                Else
                                    madrsFormatRptData.Fields("EXCEPTION2") = !EXCEPTION_CD
                                    lblnProcessExcep = True
                                End If
                            Case 2
                                If ExcepExists(!EXCEPTION_CD) Then
                                    lblnProcessExcep = False
                                Else
                                    madrsFormatRptData.Fields("EXCEPTION3") = !EXCEPTION_CD
                                    lblnProcessExcep = True
                                End If
                            Case 3
                                If ExcepExists(!EXCEPTION_CD) Then
                                    lblnProcessExcep = False
                                Else
                                    madrsFormatRptData.Fields("EXCEPTION4") = !EXCEPTION_CD
                                    lblnProcessExcep = True
                                End If
                            Case 4
                                If ExcepExists(!EXCEPTION_CD) Then
                                    lblnProcessExcep = False
                                Else
                                    madrsFormatRptData.Fields("EXCEPTION5") = !EXCEPTION_CD
                                    lblnProcessExcep = True
                                End If
                            Case 5
                                If ExcepExists(!EXCEPTION_CD) Then
                                    lblnProcessExcep = False
                                Else
                                    madrsFormatRptData.Fields("EXCEPTION6") = !EXCEPTION_CD
                                    lblnProcessExcep = True
                                End If
                            Case 6
                                If ExcepExists(!EXCEPTION_CD) Then
                                    lblnProcessExcep = False
                                Else
                                    madrsFormatRptData.Fields("EXCEPTION7") = !EXCEPTION_CD
                                    lblnProcessExcep = True
                                End If
                            Case 7
                                If ExcepExists(!EXCEPTION_CD) Then
                                    lblnProcessExcep = False
                                Else
                                    madrsFormatRptData.Fields("EXCEPTION8") = !EXCEPTION_CD
                                    lblnProcessExcep = True
                                End If
                            Case 8
                                If ExcepExists(!EXCEPTION_CD) Then
                                    lblnProcessExcep = False
                                Else
                                    madrsFormatRptData.Fields("EXCEPTION9") = !EXCEPTION_CD
                                    lblnProcessExcep = True
                                End If
                            Case 9
                                If ExcepExists(!EXCEPTION_CD) Then
                                    lblnProcessExcep = False
                                Else
                                    madrsFormatRptData.Fields("EXCEPTION10") = !EXCEPTION_CD
                                    lblnProcessExcep = True
                                End If
                            Case 10
                                If ExcepExists(!EXCEPTION_CD) Then
                                    lblnProcessExcep = False
                                Else
                                    madrsFormatRptData.Fields("EXCEPTION11") = !EXCEPTION_CD
                                    lblnProcessExcep = True
                                End If
                            Case 11
                                If ExcepExists(!EXCEPTION_CD) Then
                                    lblnProcessExcep = False
                                Else
                                    madrsFormatRptData.Fields("EXCEPTION12") = !EXCEPTION_CD
                                    lblnProcessExcep = True
                                End If
                            Case 12
                                If ExcepExists(!EXCEPTION_CD) Then
                                    lblnProcessExcep = False
                                Else
                                    madrsFormatRptData.Fields("EXCEPTION13") = !EXCEPTION_CD
                                    lblnProcessExcep = True
                                End If
                            Case 13
                                If ExcepExists(!EXCEPTION_CD) Then
                                    lblnProcessExcep = False
                                Else
                                    madrsFormatRptData.Fields("EXCEPTION14") = !EXCEPTION_CD
                                    lblnProcessExcep = True
                                End If
                            Case 14
                                If ExcepExists(!EXCEPTION_CD) Then
                                    lblnProcessExcep = False
                                Else
                                    madrsFormatRptData.Fields("EXCEPTION15") = !EXCEPTION_CD
                                    lblnProcessExcep = True
                                End If
                        End Select
                    End If
                End If
                If lblnProcessExcep Then
                    Select Case mlngExceptionCount
                        Case 0
                            madrsFormatRptData.Fields("EXCEPTION1_DESC") = _
                            mobjCodeTables.ExceptionDescription(madrsFormatRptData.Fields("EXCEPTION1"))
                            madrsFormatRptData.Fields("EXCEPTION1_START") = !START_TIME
                            madrsFormatRptData.Fields("EXCEPTION1_END") = !END_TIME
                            If Trim(madrsFormatRptData.Fields("EXCEPTION1_START")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION1_START") = "00:00"
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION1_END")) = "00:00" Then
                                madrsFormatRptData.Fields("EXCEPTION1_END") = "24:00"
                            End If
                            mlngExceptionCount = mlngExceptionCount + 1
                            If OvertimeExcep(Trim(madrsFormatRptData.Fields("EXCEPTION1"))) _
                                And !OT_OPP_NUM > 0 Then
                                lintTotalStartMin = _
                                (Format(Mid(madrsFormatRptData.Fields("EXCEPTION1_START"), 1, 2), "00") * 60) + _
                                 Format(Mid(madrsFormatRptData.Fields("EXCEPTION1_START"), 4, 2), "00")
                                lintTotalEndMin = _
                                (Format(Mid(madrsFormatRptData.Fields("EXCEPTION1_END"), 1, 2), "00") * 60) + _
                                 Format(Mid(madrsFormatRptData.Fields("EXCEPTION1_END"), 4, 2), "00")
                                madrsFormatRptData.Fields("CANVASSED_OT_HOURS1") = _
                                CStr(Abs(((lintTotalEndMin - lintTotalStartMin) / 60)))
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION1_START")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION1_START") = "00:00"
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION1_END")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION1_END") = "00:00"
                            End If
                            If !OT_OPP_NUM > 0 Then
                                Set ladrsOTShift = mobjTimeCard.GetOTShiftCode(!OT_OPP_NUM)
                                If mADOExtensions.RecordsExist(ladrsOTShift) Then
                                    madrsFormatRptData.Fields("CANVASSED_OT_SHIFT1") = _
                                    ladrsOTShift.Fields("SHIFT_CD")
                                End If
                            End If
                        Case 1
                            madrsFormatRptData.Fields("EXCEPTION2_DESC") = _
                            mobjCodeTables.ExceptionDescription(madrsFormatRptData.Fields("EXCEPTION2"))
                            madrsFormatRptData.Fields("EXCEPTION2_START") = !START_TIME
                            madrsFormatRptData.Fields("EXCEPTION2_END") = !END_TIME
                            If Trim(madrsFormatRptData.Fields("EXCEPTION2_START")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION2_START") = "00:00"
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION2_END")) = "00:00" Then
                                madrsFormatRptData.Fields("EXCEPTION2_END") = "24:00"
                            End If
                            mlngExceptionCount = mlngExceptionCount + 1
                            If OvertimeExcep(Trim(madrsFormatRptData.Fields("EXCEPTION2"))) _
                                And !OT_OPP_NUM > 0 Then
                                lintTotalStartMin = _
                                (Format(Mid(madrsFormatRptData.Fields("EXCEPTION2_START"), 1, 2), "00") * 60) + _
                                 Format(Mid(madrsFormatRptData.Fields("EXCEPTION2_START"), 4, 2), "00")
                                lintTotalEndMin = _
                                (Format(Mid(madrsFormatRptData.Fields("EXCEPTION2_END"), 1, 2), "00") * 60) + _
                                 Format(Mid(madrsFormatRptData.Fields("EXCEPTION2_END"), 4, 2), "00")
                                madrsFormatRptData.Fields("CANVASSED_OT_HOURS2") = _
                                CStr(Abs(((lintTotalEndMin - lintTotalStartMin) / 60)))
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION2_START")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION2_START") = "00:00"
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION2_END")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION2_END") = "00:00"
                            End If
                            If !OT_OPP_NUM > 0 Then
                                Set ladrsOTShift = mobjTimeCard.GetOTShiftCode(!OT_OPP_NUM)
                                If mADOExtensions.RecordsExist(ladrsOTShift) Then
                                    madrsFormatRptData.Fields("CANVASSED_OT_SHIFT2") = _
                                    ladrsOTShift.Fields("SHIFT_CD")
                                End If
                            End If
                        Case 2
                            madrsFormatRptData.Fields("EXCEPTION3_DESC") = _
                            mobjCodeTables.ExceptionDescription(madrsFormatRptData.Fields("EXCEPTION3"))
                            madrsFormatRptData.Fields("EXCEPTION3_START") = !START_TIME
                            madrsFormatRptData.Fields("EXCEPTION3_END") = !END_TIME
                            If Trim(madrsFormatRptData.Fields("EXCEPTION3_START")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION3_START") = "00:00"
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION3_END")) = "00:00" Then
                                madrsFormatRptData.Fields("EXCEPTION3_END") = "24:00"
                            End If
                            mlngExceptionCount = mlngExceptionCount + 1
                            If OvertimeExcep(Trim(madrsFormatRptData.Fields("EXCEPTION3"))) _
                                And !OT_OPP_NUM > 0 Then
                                lintTotalStartMin = _
                                (Format(Mid(madrsFormatRptData.Fields("EXCEPTION3_START"), 1, 2), "00") * 60) + _
                                 Format(Mid(madrsFormatRptData.Fields("EXCEPTION3_START"), 4, 2), "00")
                                lintTotalEndMin = _
                                (Format(Mid(madrsFormatRptData.Fields("EXCEPTION3_END"), 1, 2), "00") * 60) + _
                                 Format(Mid(madrsFormatRptData.Fields("EXCEPTION3_END"), 4, 2), "00")
                                madrsFormatRptData.Fields("CANVASSED_OT_HOURS3") = _
                                CStr(Abs(((lintTotalEndMin - lintTotalStartMin) / 60)))
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION3_START")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION3_START") = "00:00"
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION3_END")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION3_END") = "00:00"
                            End If
                            If !OT_OPP_NUM > 0 Then
                                Set ladrsOTShift = mobjTimeCard.GetOTShiftCode(!OT_OPP_NUM)
                                If mADOExtensions.RecordsExist(ladrsOTShift) Then
                                    madrsFormatRptData.Fields("CANVASSED_OT_SHIFT3") = _
                                    ladrsOTShift.Fields("SHIFT_CD")
                                End If
                            End If
                        Case 3
                            madrsFormatRptData.Fields("EXCEPTION4_DESC") = _
                            mobjCodeTables.ExceptionDescription(madrsFormatRptData.Fields("EXCEPTION4"))
                            madrsFormatRptData.Fields("EXCEPTION4_START") = !START_TIME
                            madrsFormatRptData.Fields("EXCEPTION4_END") = !END_TIME
                            If Trim(madrsFormatRptData.Fields("EXCEPTION4_START")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION4_START") = "00:00"
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION4_END")) = "00:00" Then
                                madrsFormatRptData.Fields("EXCEPTION4_END") = "24:00"
                            End If
                            mlngExceptionCount = mlngExceptionCount + 1
                            If OvertimeExcep(Trim(madrsFormatRptData.Fields("EXCEPTION4"))) _
                                And !OT_OPP_NUM > 0 Then
                                lintTotalStartMin = _
                                (Format(Mid(madrsFormatRptData.Fields("EXCEPTION4_START"), 1, 2), "00") * 60) + _
                                 Format(Mid(madrsFormatRptData.Fields("EXCEPTION4_START"), 4, 2), "00")
                                lintTotalEndMin = _
                                (Format(Mid(madrsFormatRptData.Fields("EXCEPTION4_END"), 1, 2), "00") * 60) + _
                                 Format(Mid(madrsFormatRptData.Fields("EXCEPTION4_END"), 4, 2), "00")
                                madrsFormatRptData.Fields("CANVASSED_OT_HOURS4") = _
                                CStr(Abs(((lintTotalEndMin - lintTotalStartMin) / 60)))
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION4_START")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION4_START") = "00:00"
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION4_END")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION4_END") = "00:00"
                            End If
                            If !OT_OPP_NUM > 0 Then
                                Set ladrsOTShift = mobjTimeCard.GetOTShiftCode(!OT_OPP_NUM)
                                If mADOExtensions.RecordsExist(ladrsOTShift) Then
                                    madrsFormatRptData.Fields("CANVASSED_OT_SHIFT4") = _
                                    ladrsOTShift.Fields("SHIFT_CD")
                                End If
                            End If
                        Case 4
                            madrsFormatRptData.Fields("EXCEPTION5_DESC") = _
                            mobjCodeTables.ExceptionDescription(madrsFormatRptData.Fields("EXCEPTION5"))
                            madrsFormatRptData.Fields("EXCEPTION5_START") = !START_TIME
                            madrsFormatRptData.Fields("EXCEPTION5_END") = !END_TIME
                            If Trim(madrsFormatRptData.Fields("EXCEPTION5_START")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION5_START") = "00:00"
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION5_END")) = "00:00" Then
                                madrsFormatRptData.Fields("EXCEPTION5_END") = "24:00"
                            End If
                            mlngExceptionCount = mlngExceptionCount + 1
                            If OvertimeExcep(Trim(madrsFormatRptData.Fields("EXCEPTION5"))) _
                                And !OT_OPP_NUM > 0 Then
                                lintTotalStartMin = _
                                (Format(Mid(madrsFormatRptData.Fields("EXCEPTION5_START"), 1, 2), "00") * 60) + _
                                 Format(Mid(madrsFormatRptData.Fields("EXCEPTION5_START"), 4, 2), "00")
                                lintTotalEndMin = _
                                (Format(Mid(madrsFormatRptData.Fields("EXCEPTION5_END"), 1, 2), "00") * 60) + _
                                 Format(Mid(madrsFormatRptData.Fields("EXCEPTION5_END"), 4, 2), "00")
                                madrsFormatRptData.Fields("CANVASSED_OT_HOURS5") = _
                                CStr(Abs(((lintTotalEndMin - lintTotalStartMin) / 60)))
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION5_START")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION5_START") = "00:00"
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION5_END")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION5_END") = "00:00"
                            End If
                            If !OT_OPP_NUM > 0 Then
                                Set ladrsOTShift = mobjTimeCard.GetOTShiftCode(!OT_OPP_NUM)
                                If mADOExtensions.RecordsExist(ladrsOTShift) Then
                                    madrsFormatRptData.Fields("CANVASSED_OT_SHIFT5") = _
                                    ladrsOTShift.Fields("SHIFT_CD")
                                End If
                            End If
                        Case 5
                            madrsFormatRptData.Fields("EXCEPTION6_DESC") = _
                            mobjCodeTables.ExceptionDescription(madrsFormatRptData.Fields("EXCEPTION6"))
                            madrsFormatRptData.Fields("EXCEPTION6_START") = !START_TIME
                            madrsFormatRptData.Fields("EXCEPTION6_END") = !END_TIME
                            mlngExceptionCount = mlngExceptionCount + 1
                            If Trim(madrsFormatRptData.Fields("EXCEPTION6_START")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION6_START") = "00:00"
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION6_END")) = "00:00" Then
                                madrsFormatRptData.Fields("EXCEPTION6_END") = "24:00"
                            End If
                            If OvertimeExcep(Trim(madrsFormatRptData.Fields("EXCEPTION6"))) _
                                And !OT_OPP_NUM > 0 Then
                                lintTotalStartMin = _
                                (Format(Mid(madrsFormatRptData.Fields("EXCEPTION6_START"), 1, 2), "00") * 60) + _
                                 Format(Mid(madrsFormatRptData.Fields("EXCEPTION6_START"), 4, 2), "00")
                                lintTotalEndMin = _
                                (Format(Mid(madrsFormatRptData.Fields("EXCEPTION6_END"), 1, 2), "00") * 60) + _
                                 Format(Mid(madrsFormatRptData.Fields("EXCEPTION6_END"), 4, 2), "00")
                                madrsFormatRptData.Fields("CANVASSED_OT_HOURS6") = _
                                CStr(Abs(((lintTotalEndMin - lintTotalStartMin) / 60)))
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION6_START")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION6_START") = "00:00"
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION6_END")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION6_END") = "00:00"
                            End If
                            If !OT_OPP_NUM > 0 Then
                                Set ladrsOTShift = mobjTimeCard.GetOTShiftCode(!OT_OPP_NUM)
                                If mADOExtensions.RecordsExist(ladrsOTShift) Then
                                    madrsFormatRptData.Fields("CANVASSED_OT_SHIFT6") = _
                                    ladrsOTShift.Fields("SHIFT_CD")
                                End If
                            End If
                        Case 6
                            madrsFormatRptData.Fields("EXCEPTION7_DESC") = _
                            mobjCodeTables.ExceptionDescription(madrsFormatRptData.Fields("EXCEPTION7"))
                            madrsFormatRptData.Fields("EXCEPTION7_START") = !START_TIME
                            madrsFormatRptData.Fields("EXCEPTION7_END") = !END_TIME
                            If Trim(madrsFormatRptData.Fields("EXCEPTION7_START")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION7_START") = "00:00"
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION7_END")) = "00:00" Then
                                madrsFormatRptData.Fields("EXCEPTION7_END") = "24:00"
                            End If
                            mlngExceptionCount = mlngExceptionCount + 1
                            If OvertimeExcep(Trim(madrsFormatRptData.Fields("EXCEPTION7"))) _
                                And !OT_OPP_NUM > 0 Then
                                lintTotalStartMin = _
                                (Format(Mid(madrsFormatRptData.Fields("EXCEPTION7_START"), 1, 2), "00") * 60) + _
                                 Format(Mid(madrsFormatRptData.Fields("EXCEPTION7_START"), 4, 2), "00")
                                lintTotalEndMin = _
                                (Format(Mid(madrsFormatRptData.Fields("EXCEPTION7_END"), 1, 2), "00") * 60) + _
                                 Format(Mid(madrsFormatRptData.Fields("EXCEPTION7_END"), 4, 2), "00")
                                madrsFormatRptData.Fields("CANVASSED_OT_HOURS7") = _
                                CStr(Abs(((lintTotalEndMin - lintTotalStartMin) / 60)))
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION7_START")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION7_START") = "00:00"
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION7_END")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION7_END") = "00:00"
                            End If
                            If !OT_OPP_NUM > 0 Then
                                Set ladrsOTShift = mobjTimeCard.GetOTShiftCode(!OT_OPP_NUM)
                                If mADOExtensions.RecordsExist(ladrsOTShift) Then
                                    madrsFormatRptData.Fields("CANVASSED_OT_SHIFT7") = _
                                    ladrsOTShift.Fields("SHIFT_CD")
                                End If
                            End If
                        Case 7
                            madrsFormatRptData.Fields("EXCEPTION8_DESC") = _
                            mobjCodeTables.ExceptionDescription(madrsFormatRptData.Fields("EXCEPTION8"))
                            madrsFormatRptData.Fields("EXCEPTION8_START") = !START_TIME
                            madrsFormatRptData.Fields("EXCEPTION8_END") = !END_TIME
                            mlngExceptionCount = mlngExceptionCount + 1
                            If Trim(madrsFormatRptData.Fields("EXCEPTION8_START")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION8_START") = "00:00"
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION8_END")) = "00:00" Then
                                madrsFormatRptData.Fields("EXCEPTION8_END") = "24:00"
                            End If
                            If OvertimeExcep(Trim(madrsFormatRptData.Fields("EXCEPTION8"))) _
                                And !OT_OPP_NUM > 0 Then
                                lintTotalStartMin = _
                                (Format(Mid(madrsFormatRptData.Fields("EXCEPTION8_START"), 1, 2), "00") * 60) + _
                                 Format(Mid(madrsFormatRptData.Fields("EXCEPTION8_START"), 4, 2), "00")
                                lintTotalEndMin = _
                                (Format(Mid(madrsFormatRptData.Fields("EXCEPTION8_END"), 1, 2), "00") * 60) + _
                                 Format(Mid(madrsFormatRptData.Fields("EXCEPTION8_END"), 4, 2), "00")
                                madrsFormatRptData.Fields("CANVASSED_OT_HOURS8") = _
                                CStr(Abs(((lintTotalEndMin - lintTotalStartMin) / 60)))
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION8_START")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION8_START") = "00:00"
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION8_END")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION8_END") = "00:00"
                            End If
                            If !OT_OPP_NUM > 0 Then
                                Set ladrsOTShift = mobjTimeCard.GetOTShiftCode(!OT_OPP_NUM)
                                If mADOExtensions.RecordsExist(ladrsOTShift) Then
                                    madrsFormatRptData.Fields("CANVASSED_OT_SHIFT8") = _
                                    ladrsOTShift.Fields("SHIFT_CD")
                                End If
                            End If
                        Case 8
                            madrsFormatRptData.Fields("EXCEPTION9_DESC") = _
                            mobjCodeTables.ExceptionDescription(madrsFormatRptData.Fields("EXCEPTION9"))
                            madrsFormatRptData.Fields("EXCEPTION9_START") = !START_TIME
                            madrsFormatRptData.Fields("EXCEPTION9_END") = !END_TIME
                            mlngExceptionCount = mlngExceptionCount + 1
                            If Trim(madrsFormatRptData.Fields("EXCEPTION9_START")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION9_START") = "00:00"
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION9_END")) = "00:00" Then
                                madrsFormatRptData.Fields("EXCEPTION9_END") = "24:00"
                            End If
                            If OvertimeExcep(Trim(madrsFormatRptData.Fields("EXCEPTION9"))) _
                                And !OT_OPP_NUM > 0 Then
                                lintTotalStartMin = _
                                (Format(Mid(madrsFormatRptData.Fields("EXCEPTION9_START"), 1, 2), "00") * 60) + _
                                 Format(Mid(madrsFormatRptData.Fields("EXCEPTION9_START"), 4, 2), "00")
                                lintTotalEndMin = _
                                (Format(Mid(madrsFormatRptData.Fields("EXCEPTION9_END"), 1, 2), "00") * 60) + _
                                 Format(Mid(madrsFormatRptData.Fields("EXCEPTION9_END"), 4, 2), "00")
                                madrsFormatRptData.Fields("CANVASSED_OT_HOURS9") = _
                                CStr(Abs(((lintTotalEndMin - lintTotalStartMin) / 60)))
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION9_START")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION9_START") = "00:00"
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION9_END")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION9_END") = "00:00"
                            End If
                            If !OT_OPP_NUM > 0 Then
                                Set ladrsOTShift = mobjTimeCard.GetOTShiftCode(!OT_OPP_NUM)
                                If mADOExtensions.RecordsExist(ladrsOTShift) Then
                                    madrsFormatRptData.Fields("CANVASSED_OT_SHIFT9") = _
                                    ladrsOTShift.Fields("SHIFT_CD")
                                End If
                            End If
                        Case 9
                            madrsFormatRptData.Fields("EXCEPTION10_DESC") = _
                            mobjCodeTables.ExceptionDescription(madrsFormatRptData.Fields("EXCEPTION10"))
                            madrsFormatRptData.Fields("EXCEPTION10_START") = !START_TIME
                            madrsFormatRptData.Fields("EXCEPTION10_END") = !END_TIME
                            mlngExceptionCount = mlngExceptionCount + 1
                            If Trim(madrsFormatRptData.Fields("EXCEPTION10_START")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION10_START") = "00:00"
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION10_END")) = "00:00" Then
                                madrsFormatRptData.Fields("EXCEPTION10_END") = "24:00"
                            End If
                            If OvertimeExcep(Trim(madrsFormatRptData.Fields("EXCEPTION10"))) _
                                And !OT_OPP_NUM > 0 Then
                                lintTotalStartMin = _
                                (Format(Mid(madrsFormatRptData.Fields("EXCEPTION10_START"), 1, 2), "00") * 60) + _
                                 Format(Mid(madrsFormatRptData.Fields("EXCEPTION10_START"), 4, 2), "00")
                                lintTotalEndMin = _
                                (Format(Mid(madrsFormatRptData.Fields("EXCEPTION10_END"), 1, 2), "00") * 60) + _
                                 Format(Mid(madrsFormatRptData.Fields("EXCEPTION10_END"), 4, 2), "00")
                                madrsFormatRptData.Fields("CANVASSED_OT_HOURS10") = _
                                CStr(Abs(((lintTotalEndMin - lintTotalStartMin) / 60)))
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION10_START")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION10_START") = "00:00"
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION10_END")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION10_END") = "00:00"
                            End If
                            If !OT_OPP_NUM > 0 Then
                                Set ladrsOTShift = mobjTimeCard.GetOTShiftCode(!OT_OPP_NUM)
                                If mADOExtensions.RecordsExist(ladrsOTShift) Then
                                    madrsFormatRptData.Fields("CANVASSED_OT_SHIFT10") = _
                                    ladrsOTShift.Fields("SHIFT_CD")
                                End If
                            End If
                        Case 10
                            madrsFormatRptData.Fields("EXCEPTION11_DESC") = _
                            mobjCodeTables.ExceptionDescription(madrsFormatRptData.Fields("EXCEPTION11"))
                            madrsFormatRptData.Fields("EXCEPTION11_START") = !START_TIME
                            madrsFormatRptData.Fields("EXCEPTION11_END") = !END_TIME
                            If Trim(madrsFormatRptData.Fields("EXCEPTION11_START")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION11_START") = "00:00"
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION11_END")) = "00:00" Then
                                madrsFormatRptData.Fields("EXCEPTION11_END") = "24:00"
                            End If
                            mlngExceptionCount = mlngExceptionCount + 1
                            If OvertimeExcep(Trim(madrsFormatRptData.Fields("EXCEPTION11"))) _
                                And !OT_OPP_NUM > 0 Then
                                lintTotalStartMin = _
                                (Format(Mid(madrsFormatRptData.Fields("EXCEPTION11_START"), 1, 2), "00") * 60) + _
                                 Format(Mid(madrsFormatRptData.Fields("EXCEPTION11_START"), 4, 2), "00")
                                lintTotalEndMin = _
                                (Format(Mid(madrsFormatRptData.Fields("EXCEPTION11_END"), 1, 2), "00") * 60) + _
                                 Format(Mid(madrsFormatRptData.Fields("EXCEPTION11_END"), 4, 2), "00")
                                madrsFormatRptData.Fields("CANVASSED_OT_HOURS11") = _
                                CStr(Abs(((lintTotalEndMin - lintTotalStartMin) / 60)))
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION11_START")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION11_START") = "00:00"
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION11_END")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION11_END") = "00:00"
                            End If
                            If !OT_OPP_NUM > 0 Then
                                Set ladrsOTShift = mobjTimeCard.GetOTShiftCode(!OT_OPP_NUM)
                                If mADOExtensions.RecordsExist(ladrsOTShift) Then
                                    madrsFormatRptData.Fields("CANVASSED_OT_SHIFT11") = _
                                    ladrsOTShift.Fields("SHIFT_CD")
                                End If
                            End If
                        Case 11
                            madrsFormatRptData.Fields("EXCEPTION12_DESC") = _
                            mobjCodeTables.ExceptionDescription(madrsFormatRptData.Fields("EXCEPTION12"))
                            madrsFormatRptData.Fields("EXCEPTION12_START") = !START_TIME
                            madrsFormatRptData.Fields("EXCEPTION12_END") = !END_TIME
                            If Trim(madrsFormatRptData.Fields("EXCEPTION12_START")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION12_START") = "00:00"
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION12_END")) = "00:00" Then
                                madrsFormatRptData.Fields("EXCEPTION12_END") = "24:00"
                            End If
                            mlngExceptionCount = mlngExceptionCount + 1
                            If OvertimeExcep(Trim(madrsFormatRptData.Fields("EXCEPTION12"))) _
                                And !OT_OPP_NUM > 0 Then
                                lintTotalStartMin = _
                                (Format(Mid(madrsFormatRptData.Fields("EXCEPTION12_START"), 1, 2), "00") * 60) + _
                                 Format(Mid(madrsFormatRptData.Fields("EXCEPTION12_START"), 4, 2), "00")
                                lintTotalEndMin = _
                                (Format(Mid(madrsFormatRptData.Fields("EXCEPTION12_END"), 1, 2), "00") * 60) + _
                                 Format(Mid(madrsFormatRptData.Fields("EXCEPTION12_END"), 4, 2), "00")
                                madrsFormatRptData.Fields("CANVASSED_OT_HOURS12") = _
                                CStr(Abs(((lintTotalEndMin - lintTotalStartMin) / 60)))
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION12_START")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION12_START") = "00:00"
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION12_END")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION12_END") = "00:00"
                            End If
                            If !OT_OPP_NUM > 0 Then
                                Set ladrsOTShift = mobjTimeCard.GetOTShiftCode(!OT_OPP_NUM)
                                If mADOExtensions.RecordsExist(ladrsOTShift) Then
                                    madrsFormatRptData.Fields("CANVASSED_OT_SHIFT12") = _
                                    ladrsOTShift.Fields("SHIFT_CD")
                                End If
                            End If
                        Case 12
                            madrsFormatRptData.Fields("EXCEPTION13_DESC") = _
                            mobjCodeTables.ExceptionDescription(madrsFormatRptData.Fields("EXCEPTION13"))
                            madrsFormatRptData.Fields("EXCEPTION13_START") = !START_TIME
                            madrsFormatRptData.Fields("EXCEPTION13_END") = !END_TIME
                            If Trim(madrsFormatRptData.Fields("EXCEPTION13_START")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION13_START") = "00:00"
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION13_END")) = "00:00" Then
                                madrsFormatRptData.Fields("EXCEPTION13_END") = "24:00"
                            End If
                            mlngExceptionCount = mlngExceptionCount + 1
                            If OvertimeExcep(Trim(madrsFormatRptData.Fields("EXCEPTION13"))) _
                                And !OT_OPP_NUM > 0 Then
                                lintTotalStartMin = _
                                (Format(Mid(madrsFormatRptData.Fields("EXCEPTION13_START"), 1, 2), "00") * 60) + _
                                 Format(Mid(madrsFormatRptData.Fields("EXCEPTION13_START"), 4, 2), "00")
                                lintTotalEndMin = _
                                (Format(Mid(madrsFormatRptData.Fields("EXCEPTION13_END"), 1, 2), "00") * 60) + _
                                 Format(Mid(madrsFormatRptData.Fields("EXCEPTION13_END"), 4, 2), "00")
                                madrsFormatRptData.Fields("CANVASSED_OT_HOURS13") = _
                                CStr(Abs(((lintTotalEndMin - lintTotalStartMin) / 60)))
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION13_START")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION13_START") = "00:00"
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION13_END")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION13_END") = "00:00"
                            End If
                            If !OT_OPP_NUM > 0 Then
                                Set ladrsOTShift = mobjTimeCard.GetOTShiftCode(!OT_OPP_NUM)
                                If mADOExtensions.RecordsExist(ladrsOTShift) Then
                                    madrsFormatRptData.Fields("CANVASSED_OT_SHIFT13") = _
                                    ladrsOTShift.Fields("SHIFT_CD")
                                End If
                            End If
                        Case 13
                            madrsFormatRptData.Fields("EXCEPTION14_DESC") = _
                            mobjCodeTables.ExceptionDescription(madrsFormatRptData.Fields("EXCEPTION14"))
                            madrsFormatRptData.Fields("EXCEPTION14_START") = !START_TIME
                            madrsFormatRptData.Fields("EXCEPTION14_END") = !END_TIME
                            If Trim(madrsFormatRptData.Fields("EXCEPTION14_START")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION14_START") = "00:00"
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION14_END")) = "00:00" Then
                                madrsFormatRptData.Fields("EXCEPTION14_END") = "24:00"
                            End If
                            mlngExceptionCount = mlngExceptionCount + 1
                            If OvertimeExcep(Trim(madrsFormatRptData.Fields("EXCEPTION14"))) _
                                And !OT_OPP_NUM > 0 Then
                                lintTotalStartMin = _
                                (Format(Mid(madrsFormatRptData.Fields("EXCEPTION14_START"), 1, 2), "00") * 60) + _
                                 Format(Mid(madrsFormatRptData.Fields("EXCEPTION14_START"), 4, 2), "00")
                                lintTotalEndMin = _
                                (Format(Mid(madrsFormatRptData.Fields("EXCEPTION14_END"), 1, 2), "00") * 60) + _
                                 Format(Mid(madrsFormatRptData.Fields("EXCEPTION14_END"), 4, 2), "00")
                                madrsFormatRptData.Fields("CANVASSED_OT_HOURS14") = _
                                CStr(Abs(((lintTotalEndMin - lintTotalStartMin) / 60)))
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION14_START")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION14_START") = "00:00"
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION14_END")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION14_END") = "00:00"
                            End If
                            If !OT_OPP_NUM > 0 Then
                                Set ladrsOTShift = mobjTimeCard.GetOTShiftCode(!OT_OPP_NUM)
                                If mADOExtensions.RecordsExist(ladrsOTShift) Then
                                    madrsFormatRptData.Fields("CANVASSED_OT_SHIFT14") = _
                                    ladrsOTShift.Fields("SHIFT_CD")
                                End If
                            End If
                        Case 14
                            madrsFormatRptData.Fields("EXCEPTION15_DESC") = _
                            mobjCodeTables.ExceptionDescription(madrsFormatRptData.Fields("EXCEPTION15"))
                            madrsFormatRptData.Fields("EXCEPTION15_START") = !START_TIME
                            madrsFormatRptData.Fields("EXCEPTION15_END") = !END_TIME
                            If Trim(madrsFormatRptData.Fields("EXCEPTION15_START")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION15_START") = "00:00"
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION15_END")) = "00:00" Then
                                madrsFormatRptData.Fields("EXCEPTION15_END") = "24:00"
                            End If
                            mlngExceptionCount = mlngExceptionCount + 1
                            If OvertimeExcep(Trim(madrsFormatRptData.Fields("EXCEPTION15"))) _
                                And !OT_OPP_NUM > 0 Then
                                lintTotalStartMin = _
                                (Format(Mid(madrsFormatRptData.Fields("EXCEPTION15_START"), 1, 2), "00") * 60) + _
                                 Format(Mid(madrsFormatRptData.Fields("EXCEPTION15_START"), 4, 2), "00")
                                lintTotalEndMin = _
                                (Format(Mid(madrsFormatRptData.Fields("EXCEPTION15_END"), 1, 2), "00") * 60) + _
                                 Format(Mid(madrsFormatRptData.Fields("EXCEPTION15_END"), 4, 2), "00")
                                madrsFormatRptData.Fields("CANVASSED_OT_HOURS15") = _
                                CStr(Abs(((lintTotalEndMin - lintTotalStartMin) / 60)))
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION15_START")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION15_START") = "00:00"
                            End If
                            If Trim(madrsFormatRptData.Fields("EXCEPTION15_END")) = "24:00" Then
                                madrsFormatRptData.Fields("EXCEPTION15_END") = "00:00"
                            End If
                            If !OT_OPP_NUM > 0 Then
                                Set ladrsOTShift = mobjTimeCard.GetOTShiftCode(!OT_OPP_NUM)
                                If mADOExtensions.RecordsExist(ladrsOTShift) Then
                                    madrsFormatRptData.Fields("CANVASSED_OT_SHIFT15") = _
                                    ladrsOTShift.Fields("SHIFT_CD")
                                End If
                            End If
                    End Select
                End If
            End If
        End With
    End If
    
    Exit Sub
PopulateDept_Record_EH:
    
    'Raise Error depending on the Error Number
    TrsRaiseError MODULENAME _
            , "PopulateDept_Record"
End Sub
Private Function GetDailyRoster _
    (ByVal vlngWGNum As Long, _
    ByVal vdtRequestDate As Date) As ADODB.Recordset
    
On Error GoTo GetDailyRoster_EH

    Dim lobjWorkGroup As busWorkGroup.cbsWorkGroup
    
    Set lobjWorkGroup = CreateObject("busWorkGroup.cbsWorkGroup")
    Set GetDailyRoster = lobjWorkGroup.GetDailyRosterSchedule(vlngWGNum, _
                                            marrRequestDate(mlngDateIndex))
    Set lobjWorkGroup = Nothing
    
    Exit Function
GetDailyRoster_EH:
    Set lobjWorkGroup = Nothing
    
    'Raise Error depending on the Error Number
    TrsRaiseError MODULENAME _
            , "GetDailyRoster"

            
End Function
Private Function GetDailyRoster2 _
    (ByVal vlngWGNum As Long, _
    ByVal vdtStartDate As Date, _
    ByVal vdtEndDate As Date) As ADODB.Recordset
    
On Error GoTo GetDailyRoster2_EH

    Dim lobjWorkGroup As busWorkGroup.cbsWorkGroup
    
    Set lobjWorkGroup = CreateObject("busWorkGroup.cbsWorkGroup")
    'Set GetDailyRoster2 = lobjWorkGroup.GetDailyRosterSchedule2(vlngWGNum, _
                                        vdtStartDate, _
                                        vdtEndDate)
    Set GetDailyRoster2 = lobjWorkGroup.GetDailyRosterSchedule3(vlngWGNum, _
                                        marrEmployees, _
                                        vdtStartDate, _
                                        vdtEndDate)
    
    Set lobjWorkGroup = Nothing
    
    Exit Function
GetDailyRoster2_EH:
    Set lobjWorkGroup = Nothing
    
    'Raise error depending on the error number
    TrsRaiseError MODULENAME _
            , "GetDailyRoster2"
            
End Function

Public Function GetDailyRosterWGReportData() As ADODB.Recordset
On Error GoTo GetDailyRosterWGReportData_EH

    madrsFormatRptData.MoveFirst
        
    'Call to pass the report data
    If mADOExtensions.RecordsExist(madrsFormatRptData) Then
        If mblnSortbyLine Then
            madrsFormatRptData.Sort = "WG_NUM, CHARGE_DATE, LINE_NUM, EMP_NUM"
        Else
            madrsFormatRptData.Sort = "WG_NUM, CHARGE_DATE, SORT_CODE, REG_SHIFT"
        End If
        madrsFormatRptData.MoveFirst
        Set GetDailyRosterWGReportData = madrsFormatRptData
    Else
        Set GetDailyRosterWGReportData = Nothing
    End If
    
    Exit Function
GetDailyRosterWGReportData_EH:
    'Raise Error depending on the Error Number
    TrsRaiseError MODULENAME _
            , "GetDailyRosterWGReportData" _
            , , , , , True
    
End Function
Public Function GetDailyRosterDeptReportData() As ADODB.Recordset
On Error GoTo GetDailyRosterDeptReportData_EH

    madrsFormatRptData.MoveFirst
        
    'Call to pass the report data
    If mADOExtensions.RecordsExist(madrsFormatRptData) Then
        If mblnSortbyLine Then
            madrsFormatRptData.Sort = "DEPT_NUM, WG_NUM, CHARGE_DATE, LINE_NUM"
        Else
            madrsFormatRptData.Sort = "DEPT_NUM, WG_NUM, CHARGE_DATE, SORT_CODE, REG_SHIFT"
        End If
        madrsFormatRptData.MoveFirst
        Set GetDailyRosterDeptReportData = madrsFormatRptData
    Else
        Set GetDailyRosterDeptReportData = Nothing
    End If
    
    Exit Function
GetDailyRosterDeptReportData_EH:
    'Raise Error depending on the Error Number
    TrsRaiseError MODULENAME _
            , "GetDailyRosterDeptReportData" _
            , , , , , True
    
End Function
Private Function ValidExcep(ByVal vstrExcep As String) As Boolean
On Error GoTo ValidExcep_EH

    ValidExcep = False
    
    Select Case Trim(vstrExcep)
        Case "SLT"
            ValidExcep = False
        Case "SWT"
            ValidExcep = False
        Case "RLF"
            ValidExcep = False
        Case ""
            ValidExcep = False
        Case Else
            ValidExcep = True
    End Select
    
    If Mid(Trim(vstrExcep), 1, 1) = "S" Or Mid(Trim(vstrExcep), 1, 1) = "Z" Then
        ValidExcep = False
    End If
    
    Exit Function
ValidExcep_EH:
    'Raise Error depending on the Error Number
    TrsRaiseError MODULENAME _
            , "ValidExcep"
            
End Function
Private Function OvertimeExcep(ByVal vstrExcep As String) As Boolean
On Error GoTo OvertimeExcep_EH

    OvertimeExcep = False
    
    OvertimeExcep = mobjRules.DoesRuleApply(133, trExceptionCode, vstrExcep)
    
    Exit Function
OvertimeExcep_EH:
    'Raise error depending on the Error Number
    TrsRaiseError MODULENAME _
            , "OvertimeExcep"
            
End Function
Private Function ExcepExists(ByVal vstrExcep As String) As Boolean
On Error GoTo ExcepExists_EH

Dim lstrStartTime As String
Dim lstrEndTime As String

    ExcepExists = False
    
    With madrsRosterData
    
        lstrStartTime = Trim(!START_TIME)
        lstrEndTime = Trim(!END_TIME)
        If lstrStartTime = "24:00" Then
            lstrStartTime = "00:00"
        End If
        If lstrEndTime = "24:00" Then
            lstrEndTime = "00:00"
        End If
        
        Select Case mlngExceptionCount
            Case 0
                ExcepExists = False
            Case 1
                If (madrsFormatRptData.Fields("EXCEPTION1") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION1_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION1_END") = lstrEndTime) Then
                    ExcepExists = True
                Else
                    ExcepExists = False
                End If
            Case 2
                If (madrsFormatRptData.Fields("EXCEPTION1") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION1_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION1_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION2") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION2_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION2_END") = lstrEndTime) Then
                    ExcepExists = True
                Else
                    ExcepExists = False
                End If
            Case 3
                If (madrsFormatRptData.Fields("EXCEPTION1") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION1_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION1_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION2") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION2_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION2_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION3") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION3_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION3_END") = lstrEndTime) Then
                    ExcepExists = True
                Else
                    ExcepExists = False
                End If
            Case 4
                If (madrsFormatRptData.Fields("EXCEPTION1") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION1_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION1_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION2") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION2_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION2_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION3") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION3_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION3_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION4") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION4_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION4_END") = lstrEndTime) Then
                    ExcepExists = True
                Else
                    ExcepExists = False
                End If
            Case 5
                If (madrsFormatRptData.Fields("EXCEPTION1") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION1_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION1_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION2") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION2_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION2_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION3") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION3_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION3_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION4") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION4_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION4_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION5") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION5_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION5_END") = lstrEndTime) Then
                    ExcepExists = True
                Else
                    ExcepExists = False
                End If
            Case 6
                If (madrsFormatRptData.Fields("EXCEPTION1") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION1_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION1_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION2") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION2_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION2_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION3") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION3_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION3_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION4") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION4_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION4_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION5") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION5_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION5_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION6") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION6_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION6_END") = lstrEndTime) Then
                    ExcepExists = True
                Else
                    ExcepExists = False
                End If
            Case 7
                If (madrsFormatRptData.Fields("EXCEPTION1") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION1_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION1_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION2") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION2_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION2_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION3") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION3_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION3_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION4") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION4_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION4_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION5") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION5_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION5_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION6") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION6_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION6_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION7") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION7_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION7_END") = lstrEndTime) Then
                    ExcepExists = True
                Else
                    ExcepExists = False
                End If
            Case 8
                If (madrsFormatRptData.Fields("EXCEPTION1") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION1_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION1_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION2") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION2_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION2_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION3") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION3_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION3_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION4") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION4_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION4_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION5") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION5_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION5_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION6") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION6_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION6_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION7") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION7_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION7_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION8") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION8_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION8_END") = lstrEndTime) Then
                    ExcepExists = True
                Else
                    ExcepExists = False
                End If
            Case 9
                If (madrsFormatRptData.Fields("EXCEPTION1") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION1_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION1_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION2") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION2_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION2_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION3") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION3_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION3_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION4") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION4_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION4_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION5") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION5_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION5_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION6") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION6_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION6_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION7") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION7_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION7_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION8") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION8_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION8_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION9") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION9_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION9_END") = lstrEndTime) Then
                    ExcepExists = True
                Else
                    ExcepExists = False
                End If
            Case 10
                If (madrsFormatRptData.Fields("EXCEPTION1") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION1_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION1_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION2") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION2_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION2_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION3") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION3_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION3_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION4") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION4_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION4_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION5") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION5_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION5_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION6") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION6_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION6_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION7") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION7_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION7_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION8") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION8_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION8_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION9") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION9_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION9_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION10") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION10_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION10_END") = lstrEndTime) Then
                    ExcepExists = True
                Else
                    ExcepExists = False
                End If
            Case 11
                If (madrsFormatRptData.Fields("EXCEPTION1") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION1_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION1_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION2") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION2_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION2_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION3") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION3_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION3_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION4") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION4_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION4_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION5") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION5_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION5_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION6") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION6_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION6_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION7") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION7_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION7_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION8") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION8_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION8_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION9") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION9_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION9_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION10") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION10_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION10_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION11") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION11_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION11_END") = lstrEndTime) Then
                    ExcepExists = True
                Else
                    ExcepExists = False
                End If
            Case 12
                If (madrsFormatRptData.Fields("EXCEPTION1") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION1_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION1_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION2") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION2_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION2_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION3") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION3_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION3_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION4") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION4_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION4_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION5") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION5_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION5_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION6") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION6_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION6_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION7") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION7_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION7_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION8") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION8_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION8_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION9") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION9_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION9_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION10") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION10_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION10_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION11") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION11_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION11_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION12") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION12_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION12_END") = lstrEndTime) Then
                    ExcepExists = True
                Else
                    ExcepExists = False
                End If
            Case 13
                If (madrsFormatRptData.Fields("EXCEPTION1") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION1_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION1_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION2") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION2_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION2_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION3") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION3_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION3_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION4") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION4_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION4_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION5") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION5_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION5_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION6") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION6_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION6_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION7") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION7_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION7_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION8") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION8_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION8_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION9") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION9_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION9_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION10") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION10_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION10_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION11") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION11_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION11_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION12") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION12_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION12_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION13") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION13_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION13_END") = lstrEndTime) Then
                    ExcepExists = True
                Else
                    ExcepExists = False
                End If
            Case 14
                If (madrsFormatRptData.Fields("EXCEPTION1") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION1_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION1_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION2") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION2_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION2_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION3") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION3_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION3_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION4") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION4_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION4_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION5") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION5_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION5_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION6") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION6_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION6_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION7") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION7_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION7_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION8") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION8_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION8_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION9") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION9_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION9_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION10") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION10_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION10_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION11") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION11_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION11_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION12") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION12_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION12_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION13") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION13_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION13_END") = lstrEndTime) Or _
                   (madrsFormatRptData.Fields("EXCEPTION14") = vstrExcep And madrsFormatRptData.Fields("EXCEPTION14_START") = lstrStartTime And madrsFormatRptData.Fields("EXCEPTION14_END") = lstrEndTime) Then
                    ExcepExists = True
                Else
                    ExcepExists = False
                End If
        End Select
    End With
    
    Exit Function
ExcepExists_EH:
    'Raise Error depending on the Error Number
    TrsRaiseError MODULENAME _
            , "ExcepExists"
            
End Function

Private Sub DestroyState()

    Set mobjCodeTables = Nothing
    Set mobjEmployee = Nothing
    Set mobjTimeCard = Nothing
    Set mobjRules = Nothing
    
    Set madrsRosterData = Nothing
    Set madrsFormatRptData = Nothing
    Set madrsEmpsinWG = Nothing
    Set madrsQuals = Nothing
    Set madrsSkills = Nothing
    
    Set rptMain = Nothing
    
    'Goodbye
    Set madrsDummy = Nothing

    Set madrsEmployeeInelig = Nothing
    
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
