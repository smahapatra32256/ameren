VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "rptWorkGroupSchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Project:       TRIS Conversion
'Date:          May 2004
'Author:        Lisa Bertke
'Description:

Option Explicit

Public mdtStartDate As Date, mdtEndDate As Date
Public mstrNewPage As String
Public mstrNewPagePrinted As String
Public mlngPageNum As Long
Public RlfRptPrinted As String
Public PrintRlfRpt As String
Public ReliefHdrPrint As String
Public HdrPrinted As String
Public PrintHdr As String
Public mlngRlfPageNum As Long
Dim marrWorkGroups() As Long
Dim marrExcepCount(1 To 42) As Integer
Dim marrRlfExcepCount(1 To 42) As Integer
Dim mblnStartTime As Boolean
Dim mblnUnmannedLines As Boolean
Dim mblnSortbyLine As Boolean
Dim mblnSortbyEmployee As Boolean
Dim mlngWGCount As Long
Dim mlngWGIndex As Long
Dim mlngChargeDtWG As Long
Dim mblnDemandInd As Boolean
Dim mintJobPosition As Integer
Dim mblnSecondShift As Boolean
Dim mblnJobPosition As Boolean
Dim mblnRlfSecondShift As Boolean
Dim mblnRlfJobPosition As Boolean
Dim mblnReliefReport As Boolean
Dim mblnFirstRecord As Boolean
Dim mdtPrevChargeDt As Date
Dim mlngPrevEmp As Long
Dim mlngPrevRlfEmp As Long
Dim mlngPrevLine As Long
Dim mblnOtherWGJob As Boolean
Dim mstrOtherWGJobNum As String
Dim mstrOtherWGJobDesc As String
Dim mvarBookMark As Variant
Dim madrsReportData As ADODB.Recordset
Dim mlngCurrLineCount As Long
Dim mlngPrevLineCount As Long

Dim WithEvents rptMain As ActiveReport
Dim m_ActiveReports_ErrNumber As Integer
Dim m_ActiveReports_ErrDescription As String
Dim m_ActiveReports_ErrSource As String
Dim m_ActiveReports_OrigErrDescription As String
Dim m_ActiveReports_OrigErrSource As String

Dim madrsScheduleData As ADODB.Recordset
Dim madrsFormatRptData As ADODB.Recordset
Dim madrsFormatRlfRptData As ADODB.Recordset
Dim madrsEmployeeInfo As ADODB.Recordset

Dim mobjEmployee As dbEmployee.cdEmployee
Dim mobjRules As dbRules.cdRules

Const mconFIELDATTR = adFldUpdatable Or adFldIsNullable
Private Const MODULENAME As String = "rptWorkgroupSchedule"

Public Function GetWorkGroupScheduleRpt( _
        ByVal vdtStartDate As Date, _
        ByVal vdtEndDate As Date, _
        ByVal vblnStartTime As Boolean, _
        ByVal vblnUnmannedLines As Boolean, _
        ByVal vblnSortbyLine As Boolean, _
        ByVal vadrsSelectedWorkgroups As ADODB.Recordset, _
        ByRef rbReportGenerated As Boolean, _
        ByVal vbactiveworkgroups As Boolean, _
        ByVal vadrsAllWorkGroups As ADODB.Recordset) As ActiveReport
On Error GoTo GetWorkGroupScheduleRpt_EH

    Dim ldtDay As Date
    Dim lintDayCount As Integer
    Dim lintLeft As Integer
    Dim lstrPrevMonth As String
    Dim ladrsWG As ADODB.Recordset
    Dim lobjWG As dbWorkGroup.cdWorkGroup
 
    mdtStartDate = vdtStartDate
    mdtEndDate = vdtEndDate
    mblnStartTime = vblnStartTime
    mblnUnmannedLines = vblnUnmannedLines
    mblnSortbyLine = vblnSortbyLine
    If mblnSortbyLine = True Then
        mblnSortbyEmployee = False
    Else
        mblnSortbyEmployee = True
    End If
    ResetActiveReportsError

    'Load XML file into the report object
    Set rptMain = CreateObject("DDActiveReports2.ActiveReport")
    rptMain.LoadLayout "C:\TRIS Solution\TRISReports\RPX Files\rptWGSchedule.rpx"
    
    'Add a reference in the rpt object to this (rptException) class
    Call rptMain.AddNamedItem("MyClass", Me)
                
    rptMain.Sections("PageHeader").Controls("txtHdrRunDate").Text = Date
    rptMain.Sections("PageHeader").Controls("txtHdrRunTime").Text = Time
    
    lintDayCount = 1
    lintLeft = 450
    lstrPrevMonth = ""
    For ldtDay = mdtStartDate To mdtEndDate
        If lintDayCount < 43 Then
            If lstrPrevMonth <> CStr(Format(ldtDay, "mmm")) Then
                rptMain.Sections("GroupHeader1").Controls("Day" & lintDayCount).Caption = _
                                                            CStr(Format(ldtDay, "mmm"))
            End If
            rptMain.Sections("GroupHeader1").Controls("Day" & lintDayCount).Caption = _
            rptMain.Sections("GroupHeader1").Controls("Day" & lintDayCount).Caption _
                                               & vbCrLf & CStr(Format(ldtDay, "d")) _
                                               & vbCrLf & CStr(Format(ldtDay, "ddd"))
            If CStr(Format(ldtDay, "ddd")) = "Sun" Then
                lintLeft = lintLeft + 90
            End If
            rptMain.Sections("GroupHeader1").Controls("Day" & lintDayCount).Left = _
                                                lintLeft
            rptMain.Sections("Detail").Controls("txtDtlDay" & lintDayCount).Left = _
                                                lintLeft
            lstrPrevMonth = CStr(Format(ldtDay, "mmm"))
            lintDayCount = lintDayCount + 1
            lintLeft = lintLeft + 320
        End If
    Next ldtDay
    
    If lintDayCount < 43 Then
        For lintDayCount = lintDayCount To 42
            rptMain.Sections("GroupHeader1").Controls("Day" & lintDayCount).Left = _
                                                lintLeft
            rptMain.Sections("Detail").Controls("txtDtlDay" & lintDayCount).Left = _
                                                lintLeft
            lintLeft = lintLeft + 320
        Next lintDayCount
    End If
    
    If RecordsExist(vadrsSelectedWorkgroups) Then
        With vadrsSelectedWorkgroups
            mlngWGCount = .RecordCount
            .MoveFirst
            ReDim marrWorkGroups(1 To mlngWGCount)
            For mlngWGIndex = 1 To mlngWGCount
                marrWorkGroups(mlngWGIndex) = !WG_NUM
                .MoveNext
            Next mlngWGIndex
        End With
        If mlngWGCount > 1 Then
            rptMain.Sections("GroupHeader3").GrpKeepTogether = ddGrpNone
        Else
            rptMain.Sections("GroupHeader3").GrpKeepTogether = ddGrpFirstDetail
        End If
        Set madrsFormatRptData = New ADODB.Recordset
        With madrsFormatRptData
            .Fields.Append "WG_NUM", adInteger, , mconFIELDATTR
            .Fields.Append "WG_DESC", adVarChar, 256, mconFIELDATTR
            .Fields.Append "EMP_NUM", adInteger, , mconFIELDATTR
            .Fields.Append "EMP_NAME", adVarChar, 256, mconFIELDATTR
            .Fields.Append "JOB_NUM1", adChar, 6, mconFIELDATTR
            .Fields.Append "JOB_DESC1", adVarChar, 256, mconFIELDATTR
            .Fields.Append "JOB_NUM2", adChar, 6, mconFIELDATTR
            .Fields.Append "JOB_DESC2", adVarChar, 256, mconFIELDATTR
            .Fields.Append "LINE_NUM", adInteger, , mconFIELDATTR
            .Fields.Append "CHARGE_WG_NUM", adInteger, , mconFIELDATTR
            .Fields.Append "REQ_LINE_IND", adChar, 1, mconFIELDATTR
            .Fields.Append "DEMAND_LINE", adChar, 1, mconFIELDATTR
            .Fields.Append "OTHER_WGS", adVarChar, 256, mconFIELDATTR
            .Fields.Append "PROJ_ID", adInteger, , mconFIELDATTR
            .Fields.Append "PROJ_DESC", adVarChar, 256, mconFIELDATTR
            .Fields.Append "QUALS", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SKILLS", adVarChar, 256, mconFIELDATTR
            .Fields.Append "REL_REPORT", adChar, 1, mconFIELDATTR
            .Fields.Append "SCHED_DAY1_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY2_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY3_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY4_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY5_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY6_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY7_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY8_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY9_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY10_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY11_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY12_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY13_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY14_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY15_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY16_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY17_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY18_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY19_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY20_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY21_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY22_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY23_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY24_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY25_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY26_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY27_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY28_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY29_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY30_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY31_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY32_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY33_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY34_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY35_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY36_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY37_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY38_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY39_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY40_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY41_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY42_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY1_EXCEP", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY2_EXCEP", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY3_EXCEP", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY4_EXCEP", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY5_EXCEP", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY6_EXCEP", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY7_EXCEP", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY8_EXCEP", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY9_EXCEP", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY10_EXCEP", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY11_EXCEP", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY12_EXCEP", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY13_EXCEP", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY14_EXCEP", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY15_EXCEP", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY16_EXCEP", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY17_EXCEP", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY18_EXCEP", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY19_EXCEP", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY20_EXCEP", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY21_EXCEP", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY22_EXCEP", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY23_EXCEP", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY24_EXCEP", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY25_EXCEP", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY26_EXCEP", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY27_EXCEP", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY28_EXCEP", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY29_EXCEP", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY30_EXCEP", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY31_EXCEP", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY32_EXCEP", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY33_EXCEP", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY34_EXCEP", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY35_EXCEP", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY36_EXCEP", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY37_EXCEP", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY38_EXCEP", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY39_EXCEP", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY40_EXCEP", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY41_EXCEP", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY42_EXCEP", adVarChar, 256, mconFIELDATTR
            .CursorLocation = adUseClient
            .CursorType = adOpenStatic
            .LockType = adLockBatchOptimistic
            .Open
        End With
        Set madrsFormatRlfRptData = New ADODB.Recordset
        With madrsFormatRlfRptData
            .Fields.Append "WG_NUM", adInteger, , mconFIELDATTR
            .Fields.Append "WG_DESC", adVarChar, 256, mconFIELDATTR
            .Fields.Append "EMP_NUM", adInteger, , mconFIELDATTR
            .Fields.Append "EMP_NAME", adVarChar, 256, mconFIELDATTR
            .Fields.Append "JOB_NUM1", adChar, 6, mconFIELDATTR
            .Fields.Append "JOB_DESC1", adVarChar, 256, mconFIELDATTR
            .Fields.Append "JOB_NUM2", adChar, 6, mconFIELDATTR
            .Fields.Append "JOB_DESC2", adVarChar, 256, mconFIELDATTR
            .Fields.Append "LINE_NUM", adInteger, , mconFIELDATTR
            .Fields.Append "CHARGE_WG_NUM", adInteger, , mconFIELDATTR
            .Fields.Append "REQ_LINE_IND", adChar, 1, mconFIELDATTR
            .Fields.Append "DEMAND_LINE", adChar, 1, mconFIELDATTR
            .Fields.Append "OTHER_WGS", adVarChar, 256, mconFIELDATTR
            .Fields.Append "PROJ_ID", adInteger, , mconFIELDATTR
            .Fields.Append "PROJ_DESC", adVarChar, 256, mconFIELDATTR
            .Fields.Append "QUALS", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SKILLS", adVarChar, 256, mconFIELDATTR
            .Fields.Append "REL_REPORT", adChar, 1, mconFIELDATTR
            .Fields.Append "SCHED_DAY1_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY2_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY3_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY4_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY5_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY6_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY7_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY8_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY9_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY10_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY11_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY12_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY13_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY14_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY15_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY16_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY17_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY18_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY19_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY20_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY21_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY22_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY23_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY24_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY25_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY26_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY27_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY28_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY29_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY30_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY31_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY32_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY33_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY34_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY35_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY36_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY37_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY38_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY39_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY40_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY41_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY42_SHIFT", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY1_EXCEP", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY2_EXCEP", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY3_EXCEP", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY4_EXCEP", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY5_EXCEP", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY6_EXCEP", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY7_EXCEP", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY8_EXCEP", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY9_EXCEP", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY10_EXCEP", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY11_EXCEP", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY12_EXCEP", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY13_EXCEP", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY14_EXCEP", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY15_EXCEP", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY16_EXCEP", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY17_EXCEP", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY18_EXCEP", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY19_EXCEP", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY20_EXCEP", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY21_EXCEP", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY22_EXCEP", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY23_EXCEP", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY24_EXCEP", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY25_EXCEP", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY26_EXCEP", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY27_EXCEP", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY28_EXCEP", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY29_EXCEP", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY30_EXCEP", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY31_EXCEP", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY32_EXCEP", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY33_EXCEP", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY34_EXCEP", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY35_EXCEP", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY36_EXCEP", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY37_EXCEP", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY38_EXCEP", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY39_EXCEP", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY40_EXCEP", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY41_EXCEP", adVarChar, 256, mconFIELDATTR
            .Fields.Append "SCHED_DAY42_EXCEP", adVarChar, 256, mconFIELDATTR
            .CursorLocation = adUseClient
            .CursorType = adOpenStatic
            .LockType = adLockBatchOptimistic
            .Open
        End With
    Else
        Exit Function
    End If
    
    For mlngWGIndex = 1 To mlngWGCount
        Set lobjWG = CreateObject("dbWorkGroup.cdWorkGroup")
        Set ladrsWG = lobjWG.GetWorkGroupAttributes(marrWorkGroups(mlngWGIndex))
        Set lobjWG = Nothing
        mblnDemandInd = False
        mintJobPosition = 0
        If mADOExtensions.RecordsExist(ladrsWG) Then
            If ladrsWG!DEMAND_IND = "Y" Then
                mblnDemandInd = True
            End If
        End If
        Set madrsScheduleData = GetWorkGroupSchedule _
            (marrWorkGroups(mlngWGIndex), vdtStartDate, vdtEndDate, vbactiveworkgroups)
            
        If RecordsExist(madrsScheduleData) Then
            rbReportGenerated = True
            madrsScheduleData.Sort = "EMP_NUM, LINE_NUM, CHARGE_DATE, SHIFT_NUM, SHIFT_CODE"
            madrsScheduleData.MoveFirst
            mblnFirstRecord = True
            mblnSecondShift = False
            mblnJobPosition = False
            mblnRlfSecondShift = False
            mblnRlfJobPosition = False
            mblnReliefReport = False
            mblnOtherWGJob = False
            Set mobjEmployee = CreateObject("dbEmployee.cdEmployee")
            Set mobjRules = CreateObject("dbRules.cdRules")
            With madrsScheduleData
                mlngPrevEmp = 0
                mlngPrevLine = 0
                mlngPrevRlfEmp = 0
                mdtPrevChargeDt = "9999-12-31"
                mlngChargeDtWG = 0
                Do While .EOF = False
                    If (Not mblnUnmannedLines And !EMP_NUM = 0) Then
                        .MoveNext
                    Else
                        Populate_Record
                        If Not mblnFirstRecord Then
                            madrsFormatRptData.Update
                        End If
                        If mblnReliefReport Then
                            Populate_Relief_Record
                            madrsFormatRlfRptData.Update
                            If mvarBookMark <> "" Then
                                madrsFormatRlfRptData.Bookmark = mvarBookMark
                            End If
                            mlngPrevRlfEmp = !EMP_NUM
                        End If
                        mlngPrevEmp = !EMP_NUM
                        mlngPrevLine = !LINE_NUM
                        mdtPrevChargeDt = !CHARGE_DATE
                        .MoveNext
                    End If
                Loop
            End With
        End If
    Next mlngWGIndex

    'Return the report object to the calling application
    If rbReportGenerated = True Then
        rptMain.Run
        
        '//check for any errors from the active reports script or the vb6 functions being called back
        If m_ActiveReports_ErrNumber <> 0 Then
            Err.Raise m_ActiveReports_ErrNumber, m_ActiveReports_ErrSource, m_ActiveReports_ErrDescription
        End If
        Set GetWorkGroupScheduleRpt = rptMain
    Else
        Set GetWorkGroupScheduleRpt = Nothing
    End If

    Set lobjWG = Nothing
    Set ladrsWG = Nothing

    '//-----clean up public properties used by ActiveReports-----//
    DestroyState
    
    Exit Function
GetWorkGroupScheduleRpt_EH:
    Set lobjWG = Nothing
    Set ladrsWG = Nothing

    DestroyState
    
    Select Case Err.Number
    Case m_ActiveReports_ErrNumber
        TrsRaiseError MODULENAME, "GetWorkGroupScheduleRpt", Err.Number, , , _
            DescArgsFrom("[OrigErrDescription]===>" & m_ActiveReports_OrigErrDescription & "<===", "[OrigErrSource]===>" & m_ActiveReports_OrigErrSource & "<==="), True
    Case Else
        TrsRaiseError MODULENAME, "GetWorkGroupScheduleRpt", , , , , True
    End Select
End Function

Private Sub ProcessShift(ByVal lstrshift As String)
On Error GoTo ProcessShift_EH

    With madrsScheduleData
        If mblnStartTime Then
            If !SHIFT_NUM = 1 Then
                If mlngChargeDtWG <> 0 And marrWorkGroups(mlngWGIndex) <> mlngChargeDtWG Then
                    madrsFormatRptData.Fields(lstrshift) = !SHIFT_CODE & _
                    Mid(!SHIFT_START_TIME, 2, 1) & _
                    Mid(!SHIFT_START_TIME, 4, 1) & vbCrLf & "*****"
                    mblnOtherWGJob = True
                    mstrOtherWGJobNum = !JOB_NUM
                    mstrOtherWGJobDesc = !JOB_DESC
                Else
                    If mblnDemandInd Then
                        If Not mblnJobPosition Then
                            If Trim(!JOB_POSITION) = "" Then
                                madrsFormatRptData.Fields(lstrshift) = _
                                !SHIFT_CODE & _
                                Mid(!SHIFT_START_TIME, 2, 1) & _
                                Mid(!SHIFT_START_TIME, 4, 1)
                            Else
                                madrsFormatRptData.Fields(lstrshift) = _
                                !JOB_POSITION & vbCrLf & !SHIFT_CODE & _
                                Mid(!SHIFT_START_TIME, 2, 1) & _
                                Mid(!SHIFT_START_TIME, 4, 1)
                            End If
                            mblnJobPosition = True
                        End If
                    Else
                        madrsFormatRptData.Fields(lstrshift) = !SHIFT_CODE & _
                        Mid(!SHIFT_START_TIME, 2, 1) & _
                        Mid(!SHIFT_START_TIME, 4, 1)
                    End If
                End If
            Else
                If Not mblnSecondShift Then
                    mblnSecondShift = True
                    If Len(madrsFormatRptData.Fields(lstrshift)) >= 3 Then
                        madrsFormatRptData.Fields(lstrshift) = _
                        madrsFormatRptData.Fields(lstrshift) & vbCrLf & _
                        !SHIFT_CODE & _
                        Mid(!SHIFT_START_TIME, 2, 1) & _
                        Mid(!SHIFT_START_TIME, 4, 1)
                    Else
                        If IsNull(Len(madrsFormatRptData.Fields(lstrshift))) Then
                            If mblnDemandInd Then
                                If Not mblnJobPosition Then
                                    If Trim(!JOB_POSITION) = "" Then
                                        madrsFormatRptData.Fields(lstrshift) = _
                                        madrsFormatRptData.Fields(lstrshift) & _
                                        vbCrLf & " " & vbCrLf & _
                                        !SHIFT_CODE & _
                                        Mid(!SHIFT_START_TIME, 2, 1) & _
                                        Mid(!SHIFT_START_TIME, 4, 1)
                                    Else
                                        madrsFormatRptData.Fields(lstrshift) = _
                                        !JOB_POSITION & vbCrLf & " " & vbCrLf & _
                                        !SHIFT_CODE & _
                                        Mid(!SHIFT_START_TIME, 2, 1) & _
                                        Mid(!SHIFT_START_TIME, 4, 1)
                                    End If
                                End If
                            Else
                                madrsFormatRptData.Fields(lstrshift) = _
                                madrsFormatRptData.Fields(lstrshift) & _
                                vbCrLf & " " & vbCrLf & !SHIFT_CODE & _
                                Mid(!SHIFT_START_TIME, 2, 1) & _
                                Mid(!SHIFT_START_TIME, 4, 1)
                            End If
                        End If
                    End If
                End If
            End If
        Else
            If !SHIFT_NUM = 1 Then
                If mlngChargeDtWG <> 0 And marrWorkGroups(mlngWGIndex) <> mlngChargeDtWG Then
                    madrsFormatRptData.Fields(lstrshift) = !SHIFT_CODE & vbCrLf & _
                    "*****"
                    mblnOtherWGJob = True
                    mstrOtherWGJobNum = !JOB_NUM
                    mstrOtherWGJobDesc = !JOB_DESC
                Else
                    If mblnDemandInd Then
                        If Not mblnJobPosition Then
                            If Trim(!JOB_POSITION) = "" Then
                                madrsFormatRptData.Fields(lstrshift) = _
                                !SHIFT_CODE
                            Else
                                madrsFormatRptData.Fields(lstrshift) = _
                                !JOB_POSITION & vbCrLf & !SHIFT_CODE
                            End If
                            mblnJobPosition = True
                        End If
                    Else
                        madrsFormatRptData.Fields(lstrshift) = !SHIFT_CODE
                    End If
                End If
            Else
                If Not mblnSecondShift Then
                    mblnSecondShift = True
                    madrsFormatRptData.Fields(lstrshift) = _
                    madrsFormatRptData.Fields(lstrshift) & vbCrLf & _
                    !SHIFT_CODE
                End If
            End If
        End If
    End With

    Exit Sub
ProcessShift_EH:
    'Raise Error depending on the Error Number
    TrsRaiseError MODULENAME _
            , "ProcessShift"

End Sub
Private Sub ProcessRlfShift(ByVal lstrshift As String)
On Error GoTo ProcessRlfShift_EH

    With madrsScheduleData
        If mblnStartTime Then
            If !SHIFT_NUM = 1 Then
                If mblnDemandInd Then
                    If Not mblnRlfJobPosition Then
                        If Trim(!JOB_POSITION) = "" Then
                            madrsFormatRlfRptData.Fields(lstrshift) = _
                            !SHIFT_CODE & _
                            Mid(!SHIFT_START_TIME, 2, 1) & _
                            Mid(!SHIFT_START_TIME, 4, 1)
                        Else
                            madrsFormatRlfRptData.Fields(lstrshift) = _
                            !JOB_POSITION & vbCrLf & !SHIFT_CODE & _
                            Mid(!SHIFT_START_TIME, 2, 1) & _
                            Mid(!SHIFT_START_TIME, 4, 1)
                        End If
                        mblnRlfJobPosition = True
                    End If
                Else
                    madrsFormatRlfRptData.Fields(lstrshift) = !SHIFT_CODE & vbCrLf & _
                    Mid(!SHIFT_START_TIME, 2, 1) & _
                    Mid(!SHIFT_START_TIME, 4, 1)
                End If
            Else
                If Not mblnRlfSecondShift Then
                    mblnRlfSecondShift = True
                    If Len(madrsFormatRlfRptData.Fields(lstrshift)) >= 3 Then
                        madrsFormatRlfRptData.Fields(lstrshift) = _
                        madrsFormatRlfRptData.Fields(lstrshift) & _
                        vbCrLf & !SHIFT_CODE & _
                        Mid(!SHIFT_START_TIME, 2, 1) & _
                        Mid(!SHIFT_START_TIME, 4, 1)
                    Else
                        If IsNull(Len(madrsFormatRlfRptData.Fields(lstrshift))) Then
                            If mblnDemandInd Then
                                If Not mblnRlfJobPosition Then
                                    If Trim(!JOB_POSITION) = "" Then
                                        madrsFormatRlfRptData.Fields(lstrshift) = _
                                        madrsFormatRlfRptData.Fields(lstrshift) & _
                                        vbCrLf & " " & vbCrLf & !SHIFT_CODE & _
                                        Mid(!SHIFT_START_TIME, 2, 1) & _
                                        Mid(!SHIFT_START_TIME, 4, 1)
                                    Else
                                        madrsFormatRlfRptData.Fields(lstrshift) = _
                                        !JOB_POSITION & vbCrLf & " " & vbCrLf & _
                                        !SHIFT_CODE & _
                                        Mid(!SHIFT_START_TIME, 2, 1) & _
                                        Mid(!SHIFT_START_TIME, 4, 1)
                                    End If
                                End If
                            Else
                                madrsFormatRlfRptData.Fields(lstrshift) = _
                                madrsFormatRlfRptData.Fields(lstrshift) & vbCrLf & _
                                " " & vbCrLf & !SHIFT_CODE & _
                                Mid(!SHIFT_START_TIME, 2, 1) & _
                                Mid(!SHIFT_START_TIME, 4, 1)
                            End If
                        Else
                            madrsFormatRlfRptData.Fields(lstrshift) = _
                            madrsFormatRlfRptData.Fields(lstrshift) & vbCrLf & _
                            !SHIFT_CODE
                        End If
                    End If
                End If
            End If
        Else
            If !SHIFT_NUM = 1 Then
                If mblnDemandInd Then
                    If Not mblnRlfJobPosition Then
                        If Trim(!JOB_POSITION) = "" Then
                            madrsFormatRlfRptData.Fields(lstrshift) = _
                            !SHIFT_CODE
                        Else
                            madrsFormatRlfRptData.Fields(lstrshift) = _
                            !JOB_POSITION & vbCrLf & !SHIFT_CODE
                        End If
                        mblnRlfJobPosition = True
                    End If
                Else
                    madrsFormatRlfRptData.Fields(lstrshift) = !SHIFT_CODE
                End If
            Else
                If Not mblnRlfSecondShift Then
                    mblnRlfSecondShift = True
                    madrsFormatRlfRptData.Fields(lstrshift) = _
                    madrsFormatRlfRptData.Fields(lstrshift) & _
                    vbCrLf & !SHIFT_CODE
                End If
            End If
        End If
    End With

    Exit Sub
ProcessRlfShift_EH:
    'Raise Error depending on the Error Number
    TrsRaiseError MODULENAME _
            , "ProcessRlfShift"

End Sub
Private Sub ProcessExcep(ByVal lstrExcep As String, ByVal lintDay As Integer)
On Error GoTo ProcessExcep_EH

    Dim lobjOT As dbTimeCard.cdTimeCard
    Dim ladrsOT As ADODB.Recordset
    
    With madrsScheduleData
        If marrExcepCount(lintDay) = 0 Then
            If ValidExcep(!OVER_EXC_CD) Then
                If OvertimeExcep(!OVER_EXC_CD) Then
                    Set lobjOT = CreateObject( _
                                "dbTimeCard.cdTimeCard")
                    Set ladrsOT = lobjOT.GetOTShiftCode(!OT_OPP_NUM)
                    Set lobjOT = Nothing
                    If mADOExtensions.RecordsExist(ladrsOT) Then
                        madrsFormatRptData.Fields(lstrExcep) = "OT" & ladrsOT!SHIFT_CD
                    Else
                        madrsFormatRptData.Fields(lstrExcep) = "OT"
                    End If
                    marrExcepCount(lintDay) = marrExcepCount(lintDay) + 1
                Else
                    If Trim(!SHIFT_START_TIME) = "" Then
                        If Trim(Mid(!OVER_EXC_CD, 1, 1)) = "H" Then
                            madrsFormatRptData.Fields(lstrExcep) = Trim(!OVER_EXC_CD)
                            marrExcepCount(lintDay) = marrExcepCount(lintDay) + 1
                        End If
                    Else
                        madrsFormatRptData.Fields(lstrExcep) = Trim(!OVER_EXC_CD)
                        marrExcepCount(lintDay) = marrExcepCount(lintDay) + 1
                    End If
                End If
            Else
                If ValidExcep(!EXCEPTION_CD) Then
                    If OvertimeExcep(!EXCEPTION_CD) Then
                        Set lobjOT = CreateObject( _
                                "dbTimeCard.cdTimeCard")
                        Set ladrsOT = lobjOT.GetOTShiftCode(!OT_OPP_NUM)
                        Set lobjOT = Nothing
                        If mADOExtensions.RecordsExist(ladrsOT) Then
                            madrsFormatRptData.Fields(lstrExcep) = "OT" & _
                            ladrsOT!SHIFT_CD
                        Else
                            madrsFormatRptData.Fields(lstrExcep) = "OT"
                        End If
                        marrExcepCount(lintDay) = marrExcepCount(lintDay) + 1
                    Else
                        If Trim(!SHIFT_START_TIME) = "" Then
                            If Trim(Mid(!EXCEPTION_CD, 1, 1)) = "H" Then
                                madrsFormatRptData.Fields(lstrExcep) = _
                                Trim(!EXCEPTION_CD)
                                marrExcepCount(lintDay) = marrExcepCount(lintDay) + 1
                            End If
                        Else
                            If Trim(Mid(!EXCEPTION_CD, 1, 1)) = "P" Then
                                madrsFormatRptData.Fields(lstrExcep) = "PBU"
                            Else
                                madrsFormatRptData.Fields(lstrExcep) = Trim(!EXCEPTION_CD)
                            End If
                            marrExcepCount(lintDay) = marrExcepCount(lintDay) + 1
                        End If
                    End If
                Else
                    If ValidExcep(!SYSTEM_EXC_CD) Then
                        If OvertimeExcep(!SYSTEM_EXC_CD) Then
                            Set lobjOT = CreateObject( _
                                        "dbTimeCard.cdTimeCard")
                            Set ladrsOT = lobjOT.GetOTShiftCode(!OT_OPP_NUM)
                            Set lobjOT = Nothing
                            If mADOExtensions.RecordsExist(ladrsOT) Then
                                madrsFormatRptData.Fields(lstrExcep) = "OT" & _
                                ladrsOT!SHIFT_CD
                            Else
                                madrsFormatRptData.Fields(lstrExcep) = "OT"
                            End If
                            marrExcepCount(lintDay) = marrExcepCount(lintDay) + 1
                        Else
                            If Trim(!SHIFT_START_TIME) = "" Then
                                If Trim(Mid(!SYSTEM_EXC_CD, 1, 1)) = "H" Then
                                    madrsFormatRptData.Fields(lstrExcep) = _
                                    Trim(!SYSTEM_EXC_CD)
                                    marrExcepCount(lintDay) = _
                                    marrExcepCount(lintDay) + 1
                                End If
                            Else
                                madrsFormatRptData.Fields(lstrExcep) = _
                                Trim(!SYSTEM_EXC_CD)
                                marrExcepCount(lintDay) = marrExcepCount(lintDay) + 1
                            End If
                        End If
                    End If
                End If
            End If
        Else
            If marrExcepCount(lintDay) < 5 Then
                If ValidExcep(!OVER_EXC_CD) Then
                    If OvertimeExcep(!OVER_EXC_CD) Then
                        If InStr(madrsFormatRptData.Fields(lstrExcep), "OT") = 0 Then
                            Set lobjOT = CreateObject( _
                                        "dbTimeCard.cdTimeCard")
                            Set ladrsOT = lobjOT.GetOTShiftCode(!OT_OPP_NUM)
                            Set lobjOT = Nothing
                            If mADOExtensions.RecordsExist(ladrsOT) Then
                                madrsFormatRptData.Fields(lstrExcep) = _
                                madrsFormatRptData.Fields(lstrExcep) & vbCrLf & _
                                "OT" & ladrsOT!SHIFT_CD
                            Else
                                madrsFormatRptData.Fields(lstrExcep) = _
                                madrsFormatRptData.Fields(lstrExcep) & vbCrLf & "OT"
                            End If
                            marrExcepCount(lintDay) = marrExcepCount(lintDay) + 1
                        End If
                    Else
                        If Trim(!SHIFT_START_TIME) = "" Then
                            If Trim(Mid(!OVER_EXC_CD, 1, 1)) = "H" Then
                                If InStr(madrsFormatRptData.Fields(lstrExcep), _
                                        Trim(!OVER_EXC_CD)) = 0 Then
                                    madrsFormatRptData.Fields(lstrExcep) = _
                                    madrsFormatRptData.Fields(lstrExcep) & vbCrLf & _
                                    Trim(!OVER_EXC_CD)
                                    marrExcepCount(lintDay) = _
                                    marrExcepCount(lintDay) + 1
                                End If
                            End If
                        Else
                            If InStr(madrsFormatRptData.Fields(lstrExcep), _
                                    Trim(!OVER_EXC_CD)) = 0 Then
                                madrsFormatRptData.Fields(lstrExcep) = _
                                madrsFormatRptData.Fields(lstrExcep) & vbCrLf & _
                                Trim(!OVER_EXC_CD)
                                marrExcepCount(lintDay) = marrExcepCount(lintDay) + 1
                            End If
                        End If
                    End If
                Else
                    If ValidExcep(!EXCEPTION_CD) Then
                        If OvertimeExcep(!EXCEPTION_CD) Then
                            If InStr(madrsFormatRptData.Fields(lstrExcep), "OT") = 0 Then
                                Set lobjOT = CreateObject( _
                                            "dbTimeCard.cdTimeCard")
                                Set ladrsOT = lobjOT.GetOTShiftCode(!OT_OPP_NUM)
                                Set lobjOT = Nothing
                                If mADOExtensions.RecordsExist(ladrsOT) Then
                                    madrsFormatRptData.Fields(lstrExcep) = _
                                    madrsFormatRptData.Fields(lstrExcep) & vbCrLf & _
                                    "OT" & ladrsOT!SHIFT_CD
                                Else
                                    madrsFormatRptData.Fields(lstrExcep) = _
                                    madrsFormatRptData.Fields(lstrExcep) & vbCrLf & _
                                    "OT"
                                End If
                                marrExcepCount(lintDay) = marrExcepCount(lintDay) + 1
                            End If
                        Else
                            If Trim(!SHIFT_START_TIME) = "" Then
                                If Trim(Mid(!EXCEPTION_CD, 1, 1)) = "H" Then
                                    If InStr(madrsFormatRptData.Fields(lstrExcep), _
                                            Trim(!EXCEPTION_CD)) = 0 Then
                                        madrsFormatRptData.Fields(lstrExcep) = _
                                        madrsFormatRptData.Fields(lstrExcep) & _
                                        vbCrLf & Trim(!EXCEPTION_CD)
                                        marrExcepCount(lintDay) = _
                                        marrExcepCount(lintDay) + 1
                                    End If
                                End If
                            Else
                                If InStr(madrsFormatRptData.Fields(lstrExcep), _
                                        Trim(!EXCEPTION_CD)) = 0 Then
                                    If Trim(Mid(!EXCEPTION_CD, 1, 1)) = "P" Then
                                        madrsFormatRptData.Fields(lstrExcep) = _
                                        madrsFormatRptData.Fields(lstrExcep) & vbCrLf & _
                                        "PBU"
                                    Else
                                        madrsFormatRptData.Fields(lstrExcep) = _
                                        madrsFormatRptData.Fields(lstrExcep) & vbCrLf & _
                                        Trim(!EXCEPTION_CD)
                                    End If
                                    marrExcepCount(lintDay) = _
                                    marrExcepCount(lintDay) + 1
                                End If
                            End If
                        End If
                    Else
                        If ValidExcep(!SYSTEM_EXC_CD) Then
                            If OvertimeExcep(!SYSTEM_EXC_CD) Then
                                If InStr(madrsFormatRptData.Fields(lstrExcep), "OT") = 0 Then
                                    Set lobjOT = CreateObject( _
                                                "dbTimeCard.cdTimeCard")
                                    Set ladrsOT = lobjOT.GetOTShiftCode(!OT_OPP_NUM)
                                    Set lobjOT = Nothing
                                    If mADOExtensions.RecordsExist(ladrsOT) Then
                                        madrsFormatRptData.Fields(lstrExcep) = _
                                        madrsFormatRptData.Fields(lstrExcep) & _
                                        vbCrLf & "OT" & ladrsOT!SHIFT_CD
                                    Else
                                        madrsFormatRptData.Fields(lstrExcep) = _
                                        madrsFormatRptData.Fields(lstrExcep) & _
                                        vbCrLf & "OT"
                                    End If
                                    marrExcepCount(lintDay) = _
                                    marrExcepCount(lintDay) + 1
                                End If
                            Else
                                If Trim(!SHIFT_START_TIME) = "" Then
                                    If Trim(Mid(!SYSTEM_EXC_CD, 1, 1)) = "H" Then
                                        If InStr(madrsFormatRptData.Fields(lstrExcep), _
                                                Trim(!SYSTEM_EXC_CD)) = 0 Then
                                            madrsFormatRptData.Fields(lstrExcep) = _
                                            madrsFormatRptData.Fields(lstrExcep) & _
                                            vbCrLf & Trim(!SYSTEM_EXC_CD)
                                            marrExcepCount(lintDay) = _
                                            marrExcepCount(lintDay) + 1
                                        End If
                                    End If
                                Else
                                    If InStr(madrsFormatRptData.Fields(lstrExcep), _
                                            Trim(!SYSTEM_EXC_CD)) = 0 Then
                                        madrsFormatRptData.Fields(lstrExcep) = _
                                        madrsFormatRptData.Fields(lstrExcep) & _
                                        vbCrLf & Trim(!SYSTEM_EXC_CD)
                                        marrExcepCount(lintDay) = _
                                        marrExcepCount(lintDay) + 1
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
        'process relief component
        If marrExcepCount(lintDay) = 0 Then
            If Trim(!REL_CD) <> "" Then
                mblnReliefReport = True
                Select Case Trim(!REL_CD)
                    Case "RH"
                        madrsFormatRptData.Fields(lstrExcep) = "RLH"
                        marrExcepCount(lintDay) = marrExcepCount(lintDay) + 1
                    Case "RH1"
                        madrsFormatRptData.Fields(lstrExcep) = "RLH"
                        marrExcepCount(lintDay) = marrExcepCount(lintDay) + 1
                    Case "RH2"
                        madrsFormatRptData.Fields(lstrExcep) = "RLH"
                        marrExcepCount(lintDay) = marrExcepCount(lintDay) + 1
                    Case "RL"
                        madrsFormatRptData.Fields(lstrExcep) = "RLL"
                        marrExcepCount(lintDay) = marrExcepCount(lintDay) + 1
                    Case "RL1"
                        madrsFormatRptData.Fields(lstrExcep) = "RLL"
                        marrExcepCount(lintDay) = marrExcepCount(lintDay) + 1
                    Case "RL2"
                        madrsFormatRptData.Fields(lstrExcep) = "RLL"
                        marrExcepCount(lintDay) = marrExcepCount(lintDay) + 1
                    Case "RNC"
                        madrsFormatRptData.Fields(lstrExcep) = "RNC"
                        marrExcepCount(lintDay) = marrExcepCount(lintDay) + 1
                    Case Else
                        madrsFormatRptData.Fields(lstrExcep) = "RLF"
                        marrExcepCount(lintDay) = marrExcepCount(lintDay) + 1
                End Select
            End If
        Else
            If marrExcepCount(lintDay) < 5 Then
                If Trim(!REL_CD) <> "" Then
                    mblnReliefReport = True
                    Select Case Trim(!REL_CD)
                        Case "RH"
                            If InStr(madrsFormatRptData.Fields(lstrExcep), "RLH") = 0 Then
                                madrsFormatRptData.Fields(lstrExcep) = _
                                madrsFormatRptData.Fields(lstrExcep) & vbCrLf & "RLH"
                                marrExcepCount(lintDay) = marrExcepCount(lintDay) + 1
                            End If
                        Case "RH1"
                            If InStr(madrsFormatRptData.Fields(lstrExcep), "RLH") = 0 Then
                                madrsFormatRptData.Fields(lstrExcep) = _
                                madrsFormatRptData.Fields(lstrExcep) & vbCrLf & "RLH"
                                marrExcepCount(lintDay) = marrExcepCount(lintDay) + 1
                            End If
                        Case "RH2"
                            If InStr(madrsFormatRptData.Fields(lstrExcep), "RLH") = 0 Then
                                madrsFormatRptData.Fields(lstrExcep) = _
                                madrsFormatRptData.Fields(lstrExcep) & vbCrLf & "RLH"
                                marrExcepCount(lintDay) = marrExcepCount(lintDay) + 1
                            End If
                        Case "RL"
                            If InStr(madrsFormatRptData.Fields(lstrExcep), "RLL") = 0 Then
                                madrsFormatRptData.Fields(lstrExcep) = _
                                madrsFormatRptData.Fields(lstrExcep) & vbCrLf & "RLL"
                                marrExcepCount(lintDay) = marrExcepCount(lintDay) + 1
                            End If
                        Case "RL1"
                            If InStr(madrsFormatRptData.Fields(lstrExcep), "RLL") = 0 Then
                                madrsFormatRptData.Fields(lstrExcep) = _
                                madrsFormatRptData.Fields(lstrExcep) & vbCrLf & "RLL"
                                marrExcepCount(lintDay) = marrExcepCount(lintDay) + 1
                            End If
                        Case "RL2"
                            If InStr(madrsFormatRptData.Fields(lstrExcep), "RLL") = 0 Then
                                madrsFormatRptData.Fields(lstrExcep) = _
                                madrsFormatRptData.Fields(lstrExcep) & vbCrLf & "RLL"
                                marrExcepCount(lintDay) = marrExcepCount(lintDay) + 1
                            End If
                        Case "RNC"
                            If InStr(madrsFormatRptData.Fields(lstrExcep), "RNC") = 0 Then
                                madrsFormatRptData.Fields(lstrExcep) = _
                                madrsFormatRptData.Fields(lstrExcep) & vbCrLf & "RNC"
                                marrExcepCount(lintDay) = marrExcepCount(lintDay) + 1
                            End If
                        Case Else
                            If InStr(madrsFormatRptData.Fields(lstrExcep), "RLF") = 0 Then
                                madrsFormatRptData.Fields(lstrExcep) = _
                                madrsFormatRptData.Fields(lstrExcep) & vbCrLf & "RLF"
                                marrExcepCount(lintDay) = marrExcepCount(lintDay) + 1
                            End If
                    End Select
                End If
            End If
        End If
    End With

    Exit Sub
ProcessExcep_EH:
    'Raise Error depending on the Error Number
    TrsRaiseError MODULENAME _
            , "ProcessExcep"
End Sub
Private Sub ProcessRlfExcep(ByVal lstrExcep As String, ByVal lintDay As Integer)
On Error GoTo ProcessRlfExcep_EH

    Dim lobjOT As dbTimeCard.cdTimeCard
    Dim ladrsOT As ADODB.Recordset
    
    With madrsScheduleData
        'process relief component
        If Trim(!REL_CD) <> "" Then
            If marrRlfExcepCount(lintDay) = 0 Then
                Select Case Trim(!REL_CD)
                    Case "RH"
                        madrsFormatRlfRptData.Fields(lstrExcep) = "RLH"
                        marrRlfExcepCount(lintDay) = marrRlfExcepCount(lintDay) + 1
                    Case "RH1"
                        madrsFormatRlfRptData.Fields(lstrExcep) = "RLH"
                        marrRlfExcepCount(lintDay) = marrRlfExcepCount(lintDay) + 1
                    Case "RH2"
                        madrsFormatRlfRptData.Fields(lstrExcep) = "RLH"
                        marrRlfExcepCount(lintDay) = marrRlfExcepCount(lintDay) + 1
                    Case "RL"
                        madrsFormatRlfRptData.Fields(lstrExcep) = "RLL"
                        marrRlfExcepCount(lintDay) = marrRlfExcepCount(lintDay) + 1
                    Case "RL1"
                        madrsFormatRlfRptData.Fields(lstrExcep) = "RLL"
                        marrRlfExcepCount(lintDay) = marrRlfExcepCount(lintDay) + 1
                    Case "RL2"
                        madrsFormatRlfRptData.Fields(lstrExcep) = "RLL"
                        marrRlfExcepCount(lintDay) = marrRlfExcepCount(lintDay) + 1
                    Case "RNC"
                        madrsFormatRlfRptData.Fields(lstrExcep) = "RNC"
                        marrRlfExcepCount(lintDay) = marrRlfExcepCount(lintDay) + 1
                    Case Else
                        madrsFormatRlfRptData.Fields(lstrExcep) = "RLF"
                        marrRlfExcepCount(lintDay) = marrRlfExcepCount(lintDay) + 1
                End Select
            Else
                If marrRlfExcepCount(lintDay) < 5 Then
                    Select Case Trim(!REL_CD)
                        Case "RH"
                            If InStr(madrsFormatRlfRptData.Fields(lstrExcep), "RLH") = 0 Then
                                madrsFormatRlfRptData.Fields(lstrExcep) = "RLH" & _
                                vbCrLf & madrsFormatRlfRptData.Fields(lstrExcep)
                                marrRlfExcepCount(lintDay) = _
                                marrRlfExcepCount(lintDay) + 1
                            End If
                        Case "RH1"
                            If InStr(madrsFormatRlfRptData.Fields(lstrExcep), "RLH") = 0 Then
                                madrsFormatRlfRptData.Fields(lstrExcep) = "RLH" & _
                                vbCrLf & madrsFormatRlfRptData.Fields(lstrExcep)
                                marrRlfExcepCount(lintDay) = _
                                marrRlfExcepCount(lintDay) + 1
                            End If
                        Case "RH2"
                            If InStr(madrsFormatRlfRptData.Fields(lstrExcep), "RLH") = 0 Then
                                madrsFormatRlfRptData.Fields(lstrExcep) = "RLH" & _
                                vbCrLf & madrsFormatRlfRptData.Fields(lstrExcep)
                                marrRlfExcepCount(lintDay) = _
                                marrRlfExcepCount(lintDay) + 1
                            End If
                        Case "RL"
                            If InStr(madrsFormatRlfRptData.Fields(lstrExcep), "RLL") = 0 Then
                                madrsFormatRlfRptData.Fields(lstrExcep) = "RLL" & _
                                vbCrLf & madrsFormatRlfRptData.Fields(lstrExcep)
                                marrRlfExcepCount(lintDay) = _
                                marrRlfExcepCount(lintDay) + 1
                            End If
                        Case "RL1"
                            If InStr(madrsFormatRlfRptData.Fields(lstrExcep), "RLL") = 0 Then
                                madrsFormatRlfRptData.Fields(lstrExcep) = "RLL" & _
                                vbCrLf & madrsFormatRlfRptData.Fields(lstrExcep)
                                marrRlfExcepCount(lintDay) = _
                                marrRlfExcepCount(lintDay) + 1
                            End If
                        Case "RL2"
                            If InStr(madrsFormatRlfRptData.Fields(lstrExcep), "RLL") = 0 Then
                                madrsFormatRlfRptData.Fields(lstrExcep) = "RLL" & _
                                vbCrLf & madrsFormatRlfRptData.Fields(lstrExcep)
                                marrRlfExcepCount(lintDay) = _
                                marrRlfExcepCount(lintDay) + 1
                            End If
                        Case "RNC"
                            If InStr(madrsFormatRlfRptData.Fields(lstrExcep), "RNC") = 0 Then
                                madrsFormatRlfRptData.Fields(lstrExcep) = "RNC" & _
                                vbCrLf & madrsFormatRlfRptData.Fields(lstrExcep)
                                marrRlfExcepCount(lintDay) = _
                                marrRlfExcepCount(lintDay) + 1
                            End If
                        Case Else
                            If InStr(madrsFormatRlfRptData.Fields(lstrExcep), "RLF") = 0 Then
                                madrsFormatRlfRptData.Fields(lstrExcep) = "RLF" & _
                                vbCrLf & madrsFormatRlfRptData.Fields(lstrExcep)
                                marrRlfExcepCount(lintDay) = _
                                marrRlfExcepCount(lintDay) + 1
                            End If
                    End Select
                End If
            End If
        End If
        If marrExcepCount(lintDay) = 0 Then
            If ValidExcep(!OVER_EXC_CD) Then
                If OvertimeExcep(!OVER_EXC_CD) Then
                    Set lobjOT = CreateObject( _
                                "dbTimeCard.cdTimeCard")
                    Set ladrsOT = lobjOT.GetOTShiftCode(!OT_OPP_NUM)
                    Set lobjOT = Nothing
                    If mADOExtensions.RecordsExist(ladrsOT) Then
                        madrsFormatRlfRptData.Fields(lstrExcep) = "OT" & _
                        ladrsOT!SHIFT_CD
                    Else
                        madrsFormatRlfRptData.Fields(lstrExcep) = "OT"
                    End If
                    marrRlfExcepCount(lintDay) = marrRlfExcepCount(lintDay) + 1
                Else
                    If Trim(!SHIFT_START_TIME) = "" Then
                        If Trim(Mid(!OVER_EXC_CD, 1, 1)) = "H" Then
                            madrsFormatRlfRptData.Fields(lstrExcep) = _
                            Trim(!OVER_EXC_CD)
                            marrRlfExcepCount(lintDay) = marrRlfExcepCount(lintDay) + 1
                        End If
                    Else
                        madrsFormatRlfRptData.Fields(lstrExcep) = Trim(!OVER_EXC_CD)
                        marrRlfExcepCount(lintDay) = marrRlfExcepCount(lintDay) + 1
                    End If
                End If
            Else
                If ValidExcep(!EXCEPTION_CD) Then
                    If OvertimeExcep(!EXCEPTION_CD) Then
                        Set lobjOT = CreateObject( _
                                "dbTimeCard.cdTimeCard")
                        Set ladrsOT = lobjOT.GetOTShiftCode(!OT_OPP_NUM)
                        Set lobjOT = Nothing
                        If mADOExtensions.RecordsExist(ladrsOT) Then
                            madrsFormatRlfRptData.Fields(lstrExcep) = "OT" & _
                            ladrsOT!SHIFT_CD
                        Else
                            madrsFormatRlfRptData.Fields(lstrExcep) = "OT"
                        End If
                        marrRlfExcepCount(lintDay) = marrRlfExcepCount(lintDay) + 1
                    Else
                        If Trim(!SHIFT_START_TIME) = "" Then
                            If Trim(Mid(!EXCEPTION_CD, 1, 1)) = "H" Then
                                madrsFormatRlfRptData.Fields(lstrExcep) = _
                                Trim(!EXCEPTION_CD)
                                marrRlfExcepCount(lintDay) = _
                                marrRlfExcepCount(lintDay) + 1
                            End If
                        Else
                            If Trim(Mid(!EXCEPTION_CD, 1, 1)) = "P" Then
                                madrsFormatRlfRptData.Fields(lstrExcep) = _
                                "PBU"
                            Else
                                madrsFormatRlfRptData.Fields(lstrExcep) = _
                                Trim(!EXCEPTION_CD)
                            End If
                            marrRlfExcepCount(lintDay) = marrRlfExcepCount(lintDay) + 1
                        End If
                    End If
                Else
                    If ValidExcep(!SYSTEM_EXC_CD) Then
                        If OvertimeExcep(!SYSTEM_EXC_CD) Then
                            Set lobjOT = CreateObject( _
                                    "dbTimeCard.cdTimeCard")
                            Set ladrsOT = lobjOT.GetOTShiftCode(!OT_OPP_NUM)
                            Set lobjOT = Nothing
                            If mADOExtensions.RecordsExist(ladrsOT) Then
                                madrsFormatRlfRptData.Fields(lstrExcep) = "OT" & _
                                ladrsOT!SHIFT_CD
                            Else
                                madrsFormatRlfRptData.Fields(lstrExcep) = "OT"
                            End If
                            marrRlfExcepCount(lintDay) = marrRlfExcepCount(lintDay) + 1
                        Else
                            If Trim(!SHIFT_START_TIME) = "" Then
                                If Trim(Mid(!SYSTEM_EXC_CD, 1, 1)) = "H" Then
                                    madrsFormatRlfRptData.Fields(lstrExcep) = _
                                    Trim(!SYSTEM_EXC_CD)
                                    marrRlfExcepCount(lintDay) = _
                                    marrRlfExcepCount(lintDay) + 1
                                End If
                            Else
                                madrsFormatRlfRptData.Fields(lstrExcep) = _
                                Trim(!SYSTEM_EXC_CD)
                                marrRlfExcepCount(lintDay) = _
                                marrRlfExcepCount(lintDay) + 1
                            End If
                        End If
                    End If
                End If
            End If
        Else
            If marrRlfExcepCount(lintDay) < 5 Then
                If ValidExcep(!OVER_EXC_CD) Then
                    If OvertimeExcep(!OVER_EXC_CD) Then
                        If InStr(madrsFormatRlfRptData.Fields(lstrExcep), "OT") = 0 Then
                            Set lobjOT = CreateObject( _
                                    "dbTimeCard.cdTimeCard")
                            Set ladrsOT = lobjOT.GetOTShiftCode(!OT_OPP_NUM)
                            Set lobjOT = Nothing
                            If mADOExtensions.RecordsExist(ladrsOT) Then
                                madrsFormatRlfRptData.Fields(lstrExcep) = _
                                madrsFormatRlfRptData.Fields(lstrExcep) & vbCrLf & "OT" & ladrsOT!SHIFT_CD
                            Else
                                madrsFormatRlfRptData.Fields(lstrExcep) = _
                                madrsFormatRlfRptData.Fields(lstrExcep) & vbCrLf & "OT"
                            End If
                            marrRlfExcepCount(lintDay) = marrRlfExcepCount(lintDay) + 1
                        End If
                    Else
                        If Trim(!SHIFT_START_TIME) = "" Then
                            If Trim(Mid(!OVER_EXC_CD, 1, 1)) = "H" Then
                                If InStr(madrsFormatRlfRptData.Fields(lstrExcep), _
                                        Trim(!OVER_EXC_CD)) = 0 Then
                                    madrsFormatRlfRptData.Fields(lstrExcep) = _
                                    madrsFormatRlfRptData.Fields(lstrExcep) & _
                                    vbCrLf & Trim(!OVER_EXC_CD)
                                    marrRlfExcepCount(lintDay) = _
                                    marrRlfExcepCount(lintDay) + 1
                                End If
                            End If
                        Else
                            If InStr(madrsFormatRlfRptData.Fields(lstrExcep), _
                                    Trim(!OVER_EXC_CD)) = 0 Then
                                madrsFormatRlfRptData.Fields(lstrExcep) = _
                                madrsFormatRlfRptData.Fields(lstrExcep) & _
                                vbCrLf & Trim(!OVER_EXC_CD)
                                marrRlfExcepCount(lintDay) = _
                                marrRlfExcepCount(lintDay) + 1
                            End If
                        End If
                    End If
                Else
                    If ValidExcep(!EXCEPTION_CD) Then
                        If OvertimeExcep(!EXCEPTION_CD) Then
                            If InStr(madrsFormatRlfRptData.Fields(lstrExcep), "OT") = 0 Then
                                Set lobjOT = CreateObject( _
                                        "dbTimeCard.cdTimeCard")
                                Set ladrsOT = lobjOT.GetOTShiftCode(!OT_OPP_NUM)
                                Set lobjOT = Nothing
                                If mADOExtensions.RecordsExist(ladrsOT) Then
                                    madrsFormatRlfRptData.Fields(lstrExcep) = _
                                    madrsFormatRlfRptData.Fields(lstrExcep) & _
                                    vbCrLf & "OT" & ladrsOT!SHIFT_CD
                                Else
                                    madrsFormatRlfRptData.Fields(lstrExcep) = _
                                    madrsFormatRlfRptData.Fields(lstrExcep) & _
                                    vbCrLf & "OT"
                                End If
                                marrRlfExcepCount(lintDay) = _
                                marrRlfExcepCount(lintDay) + 1
                            End If
                        Else
                            If Trim(!SHIFT_START_TIME) = "" Then
                                If Trim(Mid(!EXCEPTION_CD, 1, 1)) = "H" Then
                                    If InStr(madrsFormatRlfRptData.Fields(lstrExcep), _
                                            Trim(!EXCEPTION_CD)) = 0 Then
                                        madrsFormatRlfRptData.Fields(lstrExcep) = _
                                        madrsFormatRlfRptData.Fields(lstrExcep) & _
                                        vbCrLf & Trim(!EXCEPTION_CD)
                                        marrRlfExcepCount(lintDay) = _
                                        marrRlfExcepCount(lintDay) + 1
                                    End If
                                End If
                            Else
                                If InStr(madrsFormatRlfRptData.Fields(lstrExcep), _
                                        Trim(!EXCEPTION_CD)) = 0 Then
                                    If Trim(Mid(!EXCEPTION_CD, 1, 1)) = "P" Then
                                        madrsFormatRlfRptData.Fields(lstrExcep) = _
                                        madrsFormatRlfRptData.Fields(lstrExcep) & _
                                        vbCrLf & "PBU"
                                    Else
                                        madrsFormatRlfRptData.Fields(lstrExcep) = _
                                        madrsFormatRlfRptData.Fields(lstrExcep) & _
                                        vbCrLf & Trim(!EXCEPTION_CD)
                                    End If
                                    marrRlfExcepCount(lintDay) = _
                                    marrRlfExcepCount(lintDay) + 1
                                End If
                            End If
                        End If
                    Else
                        If ValidExcep(!SYSTEM_EXC_CD) Then
                            If OvertimeExcep(!SYSTEM_EXC_CD) Then
                                If InStr(madrsFormatRlfRptData.Fields(lstrExcep), "OT") = 0 Then
                                    Set lobjOT = CreateObject( _
                                            "dbTimeCard.cdTimeCard")
                                    Set ladrsOT = lobjOT.GetOTShiftCode(!OT_OPP_NUM)
                                    Set lobjOT = Nothing
                                    If mADOExtensions.RecordsExist(ladrsOT) Then
                                        madrsFormatRlfRptData.Fields(lstrExcep) = _
                                        madrsFormatRlfRptData.Fields(lstrExcep) & _
                                        vbCrLf & "OT" & ladrsOT!SHIFT_CD
                                    Else
                                        madrsFormatRlfRptData.Fields(lstrExcep) = _
                                        madrsFormatRlfRptData.Fields(lstrExcep) & _
                                        vbCrLf & "OT"
                                    End If
                                    marrRlfExcepCount(lintDay) = _
                                    marrRlfExcepCount(lintDay) + 1
                                End If
                            Else
                                If Trim(!SHIFT_START_TIME) = "" Then
                                    If Trim(Mid(!SYSTEM_EXC_CD, 1, 1)) = "H" Then
                                        If InStr(madrsFormatRlfRptData.Fields(lstrExcep), _
                                        Trim(!SYSTEM_EXC_CD)) = 0 Then
                                            madrsFormatRlfRptData.Fields(lstrExcep) = _
                                            madrsFormatRlfRptData.Fields(lstrExcep) & _
                                            vbCrLf & Trim(!SYSTEM_EXC_CD)
                                            marrRlfExcepCount(lintDay) = _
                                            marrRlfExcepCount(lintDay) + 1
                                        End If
                                    End If
                                Else
                                    If InStr(madrsFormatRlfRptData.Fields(lstrExcep), _
                                            Trim(!SYSTEM_EXC_CD)) = 0 Then
                                        madrsFormatRlfRptData.Fields(lstrExcep) = _
                                        madrsFormatRlfRptData.Fields(lstrExcep) & _
                                        vbCrLf & Trim(!SYSTEM_EXC_CD)
                                        marrRlfExcepCount(lintDay) = _
                                        marrRlfExcepCount(lintDay) + 1
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End With

    Set ladrsOT = Nothing
    Set lobjOT = Nothing

    Exit Sub
ProcessRlfExcep_EH:
    Set ladrsOT = Nothing
    Set lobjOT = Nothing

    'Raise Error depending on the Error Number
    TrsRaiseError MODULENAME _
            , "ProcessRlfExcep"
End Sub
Sub InitializeExcep()
On Error GoTo InitializeExcep_EH

    Dim i As Integer

    For i = 1 To 42
        marrExcepCount(i) = 0
    Next i
    madrsFormatRptData.Fields("SCHED_DAY1_EXCEP") = ""
    madrsFormatRptData.Fields("SCHED_DAY2_EXCEP") = ""
    madrsFormatRptData.Fields("SCHED_DAY3_EXCEP") = ""
    madrsFormatRptData.Fields("SCHED_DAY4_EXCEP") = ""
    madrsFormatRptData.Fields("SCHED_DAY5_EXCEP") = ""
    madrsFormatRptData.Fields("SCHED_DAY6_EXCEP") = ""
    madrsFormatRptData.Fields("SCHED_DAY7_EXCEP") = ""
    madrsFormatRptData.Fields("SCHED_DAY8_EXCEP") = ""
    madrsFormatRptData.Fields("SCHED_DAY9_EXCEP") = ""
    madrsFormatRptData.Fields("SCHED_DAY10_EXCEP") = ""
    madrsFormatRptData.Fields("SCHED_DAY11_EXCEP") = ""
    madrsFormatRptData.Fields("SCHED_DAY12_EXCEP") = ""
    madrsFormatRptData.Fields("SCHED_DAY13_EXCEP") = ""
    madrsFormatRptData.Fields("SCHED_DAY14_EXCEP") = ""
    madrsFormatRptData.Fields("SCHED_DAY15_EXCEP") = ""
    madrsFormatRptData.Fields("SCHED_DAY16_EXCEP") = ""
    madrsFormatRptData.Fields("SCHED_DAY17_EXCEP") = ""
    madrsFormatRptData.Fields("SCHED_DAY18_EXCEP") = ""
    madrsFormatRptData.Fields("SCHED_DAY19_EXCEP") = ""
    madrsFormatRptData.Fields("SCHED_DAY20_EXCEP") = ""
    madrsFormatRptData.Fields("SCHED_DAY21_EXCEP") = ""
    madrsFormatRptData.Fields("SCHED_DAY22_EXCEP") = ""
    madrsFormatRptData.Fields("SCHED_DAY23_EXCEP") = ""
    madrsFormatRptData.Fields("SCHED_DAY24_EXCEP") = ""
    madrsFormatRptData.Fields("SCHED_DAY25_EXCEP") = ""
    madrsFormatRptData.Fields("SCHED_DAY26_EXCEP") = ""
    madrsFormatRptData.Fields("SCHED_DAY27_EXCEP") = ""
    madrsFormatRptData.Fields("SCHED_DAY28_EXCEP") = ""
    madrsFormatRptData.Fields("SCHED_DAY29_EXCEP") = ""
    madrsFormatRptData.Fields("SCHED_DAY30_EXCEP") = ""
    madrsFormatRptData.Fields("SCHED_DAY31_EXCEP") = ""
    madrsFormatRptData.Fields("SCHED_DAY32_EXCEP") = ""
    madrsFormatRptData.Fields("SCHED_DAY33_EXCEP") = ""
    madrsFormatRptData.Fields("SCHED_DAY34_EXCEP") = ""
    madrsFormatRptData.Fields("SCHED_DAY35_EXCEP") = ""
    madrsFormatRptData.Fields("SCHED_DAY36_EXCEP") = ""
    madrsFormatRptData.Fields("SCHED_DAY37_EXCEP") = ""
    madrsFormatRptData.Fields("SCHED_DAY38_EXCEP") = ""
    madrsFormatRptData.Fields("SCHED_DAY39_EXCEP") = ""
    madrsFormatRptData.Fields("SCHED_DAY40_EXCEP") = ""
    madrsFormatRptData.Fields("SCHED_DAY41_EXCEP") = ""
    madrsFormatRptData.Fields("SCHED_DAY42_EXCEP") = ""

    Exit Sub
InitializeExcep_EH:
    'Raise Error depending on the Error Number
    TrsRaiseError MODULENAME _
            , "InitializeExcep"

End Sub
Sub InitializeRlfExcep()
On Error GoTo InitializeRlfExcep_EH

    Dim i As Integer

    For i = 1 To 42
        marrRlfExcepCount(i) = 0
    Next i
    madrsFormatRlfRptData.Fields("SCHED_DAY1_EXCEP") = ""
    madrsFormatRlfRptData.Fields("SCHED_DAY2_EXCEP") = ""
    madrsFormatRlfRptData.Fields("SCHED_DAY3_EXCEP") = ""
    madrsFormatRlfRptData.Fields("SCHED_DAY4_EXCEP") = ""
    madrsFormatRlfRptData.Fields("SCHED_DAY5_EXCEP") = ""
    madrsFormatRlfRptData.Fields("SCHED_DAY6_EXCEP") = ""
    madrsFormatRlfRptData.Fields("SCHED_DAY7_EXCEP") = ""
    madrsFormatRlfRptData.Fields("SCHED_DAY8_EXCEP") = ""
    madrsFormatRlfRptData.Fields("SCHED_DAY9_EXCEP") = ""
    madrsFormatRlfRptData.Fields("SCHED_DAY10_EXCEP") = ""
    madrsFormatRlfRptData.Fields("SCHED_DAY11_EXCEP") = ""
    madrsFormatRlfRptData.Fields("SCHED_DAY12_EXCEP") = ""
    madrsFormatRlfRptData.Fields("SCHED_DAY13_EXCEP") = ""
    madrsFormatRlfRptData.Fields("SCHED_DAY14_EXCEP") = ""
    madrsFormatRlfRptData.Fields("SCHED_DAY15_EXCEP") = ""
    madrsFormatRlfRptData.Fields("SCHED_DAY16_EXCEP") = ""
    madrsFormatRlfRptData.Fields("SCHED_DAY17_EXCEP") = ""
    madrsFormatRlfRptData.Fields("SCHED_DAY18_EXCEP") = ""
    madrsFormatRlfRptData.Fields("SCHED_DAY19_EXCEP") = ""
    madrsFormatRlfRptData.Fields("SCHED_DAY20_EXCEP") = ""
    madrsFormatRlfRptData.Fields("SCHED_DAY21_EXCEP") = ""
    madrsFormatRlfRptData.Fields("SCHED_DAY22_EXCEP") = ""
    madrsFormatRlfRptData.Fields("SCHED_DAY23_EXCEP") = ""
    madrsFormatRlfRptData.Fields("SCHED_DAY24_EXCEP") = ""
    madrsFormatRlfRptData.Fields("SCHED_DAY25_EXCEP") = ""
    madrsFormatRlfRptData.Fields("SCHED_DAY26_EXCEP") = ""
    madrsFormatRlfRptData.Fields("SCHED_DAY27_EXCEP") = ""
    madrsFormatRlfRptData.Fields("SCHED_DAY28_EXCEP") = ""
    madrsFormatRlfRptData.Fields("SCHED_DAY29_EXCEP") = ""
    madrsFormatRlfRptData.Fields("SCHED_DAY30_EXCEP") = ""
    madrsFormatRlfRptData.Fields("SCHED_DAY31_EXCEP") = ""
    madrsFormatRlfRptData.Fields("SCHED_DAY32_EXCEP") = ""
    madrsFormatRlfRptData.Fields("SCHED_DAY33_EXCEP") = ""
    madrsFormatRlfRptData.Fields("SCHED_DAY34_EXCEP") = ""
    madrsFormatRlfRptData.Fields("SCHED_DAY35_EXCEP") = ""
    madrsFormatRlfRptData.Fields("SCHED_DAY36_EXCEP") = ""
    madrsFormatRlfRptData.Fields("SCHED_DAY37_EXCEP") = ""
    madrsFormatRlfRptData.Fields("SCHED_DAY38_EXCEP") = ""
    madrsFormatRlfRptData.Fields("SCHED_DAY39_EXCEP") = ""
    madrsFormatRlfRptData.Fields("SCHED_DAY40_EXCEP") = ""
    madrsFormatRlfRptData.Fields("SCHED_DAY41_EXCEP") = ""
    madrsFormatRlfRptData.Fields("SCHED_DAY42_EXCEP") = ""

    Exit Sub
InitializeRlfExcep_EH:
    'Raise Error depending on the Error Number
    TrsRaiseError MODULENAME _
            , "InitializeRlfExcep"

End Sub

Sub Populate_Record()
On Error GoTo Populate_Record_EH

    Dim lintDayCount As Integer
    Dim lblnDataFound As Boolean
    Dim ldtDay As Date
    Dim lvarBookMark As Variant
    Dim lblnOtherWGs As Boolean
    Dim llngEmpNum(0) As String

    With madrsScheduleData
        If mlngPrevEmp <> !EMP_NUM Then
            mblnFirstRecord = False
            mblnOtherWGJob = False
            madrsFormatRptData.AddNew
            madrsFormatRptData.Fields("WG_NUM") = marrWorkGroups(mlngWGIndex)
            madrsFormatRptData.Fields("WG_DESC") = !WG_DESC
            madrsFormatRptData.Fields("EMP_NUM") = !EMP_NUM
            madrsFormatRptData.Fields("EMP_NAME") = !EMP_NAME
            madrsFormatRptData.Fields("JOB_NUM1") = !JOB_NUM
            madrsFormatRptData.Fields("JOB_DESC1") = !JOB_DESC
            madrsFormatRptData.Fields("JOB_NUM2") = ""
            madrsFormatRptData.Fields("JOB_DESC2") = ""
            madrsFormatRptData.Fields("QUALS") = !QUALS
            madrsFormatRptData.Fields("SKILLS") = !SKILLS
            madrsFormatRptData.Fields("PROJ_ID") = !PROJ_ID
            madrsFormatRptData.Fields("PROJ_DESC") = !PROJ_DESC
            If mblnDemandInd Then
                If marrWorkGroups(mlngWGIndex) <> !WG_NUM Then
                    madrsFormatRptData.Fields("LINE_NUM") = 0
                Else
                    madrsFormatRptData.Fields("LINE_NUM") = !LINE_NUM
                End If
            Else
                madrsFormatRptData.Fields("LINE_NUM") = 0
            End If
            madrsFormatRptData.Fields("REL_REPORT") = "N"
            madrsFormatRptData.Fields("CHARGE_WG_NUM") = !WG_NUM
            madrsFormatRptData.Fields("REQ_LINE_IND") = !REQ_LINE_IND
            madrsFormatRptData.Fields("DEMAND_LINE") = !WG_LINE_IND
            InitializeExcep
        Else
            If mblnDemandInd And mlngPrevLine <> !LINE_NUM Then
                If madrsFormatRptData.EOF Then
                    lblnOtherWGs = False
                Else
                    lvarBookMark = madrsFormatRptData.Bookmark
                    madrsFormatRptData.Filter = "EMP_NUM = " & !EMP_NUM
                    If Not madrsFormatRptData.EOF Then
                        madrsFormatRptData.Find "LINE_NUM = 0"
                        If Not madrsFormatRptData.EOF Then
                            lblnOtherWGs = True
                        Else
                            lblnOtherWGs = False
                        End If
                        madrsFormatRptData.Filter = adFilterNone
                        madrsFormatRptData.Bookmark = lvarBookMark
                    End If
                End If
                If Not lblnOtherWGs Then
                    mblnFirstRecord = False
                    mblnOtherWGJob = False
                    madrsFormatRptData.AddNew
                    madrsFormatRptData.Fields("WG_NUM") = marrWorkGroups(mlngWGIndex)
                    madrsFormatRptData.Fields("WG_DESC") = !WG_DESC
                    madrsFormatRptData.Fields("EMP_NUM") = !EMP_NUM
                    madrsFormatRptData.Fields("EMP_NAME") = !EMP_NAME
                    madrsFormatRptData.Fields("JOB_NUM1") = !JOB_NUM
                    madrsFormatRptData.Fields("JOB_DESC1") = !JOB_DESC
                    madrsFormatRptData.Fields("JOB_NUM2") = ""
                    madrsFormatRptData.Fields("JOB_DESC2") = ""
                    madrsFormatRptData.Fields("QUALS") = !QUALS
                    madrsFormatRptData.Fields("SKILLS") = !SKILLS
                    madrsFormatRptData.Fields("PROJ_ID") = !PROJ_ID
                    madrsFormatRptData.Fields("PROJ_DESC") = !PROJ_DESC
                    If mblnDemandInd Then
                        If marrWorkGroups(mlngWGIndex) <> !WG_NUM Then
                            madrsFormatRptData.Fields("LINE_NUM") = 0
                        Else
                            madrsFormatRptData.Fields("LINE_NUM") = !LINE_NUM
                        End If
                    Else
                        madrsFormatRptData.Fields("LINE_NUM") = 0
                    End If
                    madrsFormatRptData.Fields("REL_REPORT") = "N"
                    madrsFormatRptData.Fields("CHARGE_WG_NUM") = !WG_NUM
                    madrsFormatRptData.Fields("REQ_LINE_IND") = !REQ_LINE_IND
                    madrsFormatRptData.Fields("DEMAND_LINE") = !WG_LINE_IND
                    InitializeExcep
                End If
            End If
        End If
        If mblnOtherWGJob Then
            llngEmpNum(0) = !EMP_NUM
            Set madrsEmployeeInfo = mobjEmployee.GetEmployee(llngEmpNum(0))
            If madrsEmployeeInfo.Fields("JOB_NUM") <> !JOB_NUM Then
                madrsFormatRptData.Fields("JOB_NUM2") = !JOB_NUM
                madrsFormatRptData.Fields("JOB_DESC2") = !JOB_DESC
                madrsFormatRptData.Fields("JOB_NUM1") = madrsEmployeeInfo.Fields("JOB_NUM")
                madrsFormatRptData.Fields("JOB_DESC1") = _
                    GetJobDesc(madrsEmployeeInfo.Fields("JOB_NUM"))
            End If
        End If
        If mdtPrevChargeDt <> !CHARGE_DATE Or mlngPrevEmp <> !EMP_NUM Then
            mblnSecondShift = False
            mblnJobPosition = False
            lvarBookMark = .Bookmark
            .Filter = "CHARGE_DATE = " & !CHARGE_DATE & " AND EMP_NUM = " & !EMP_NUM
            If Not .EOF Then
                .Find "WG_NUM <> " & marrWorkGroups(mlngWGIndex)
                If Not .EOF Then
                    mlngChargeDtWG = !WG_NUM
                    .Filter = adFilterNone
                    .Bookmark = lvarBookMark
                Else
                    mlngChargeDtWG = marrWorkGroups(mlngWGIndex)
                End If
            Else
                mlngChargeDtWG = marrWorkGroups(mlngWGIndex)
            End If
            .Filter = adFilterNone
            .Bookmark = lvarBookMark
        End If
        If marrWorkGroups(mlngWGIndex) <> mlngChargeDtWG Then
            If mlngChargeDtWG <> 0 Then
                If IsNull(madrsFormatRptData.Fields("OTHER_WGS")) Then
                    madrsFormatRptData.Fields("OTHER_WGS") = mlngChargeDtWG
                Else
                    If InStr(1, madrsFormatRptData.Fields("OTHER_WGS"), mlngChargeDtWG) = 0 Then
                        madrsFormatRptData.Fields("OTHER_WGS") = _
                        madrsFormatRptData.Fields("OTHER_WGS") & " " & _
                        mlngChargeDtWG
                    End If
                End If
            End If
        End If
        lintDayCount = 0
        lblnDataFound = False
        mblnReliefReport = False
        For ldtDay = mdtStartDate To mdtEndDate
            lintDayCount = lintDayCount + 1
            If !CHARGE_DATE = ldtDay Then
                lblnDataFound = True
                Select Case lintDayCount
                    Case 1
                        ProcessShift "SCHED_DAY1_SHIFT"
                        ProcessExcep "SCHED_DAY1_EXCEP", 1
                    Case 2
                        ProcessShift "SCHED_DAY2_SHIFT"
                        ProcessExcep "SCHED_DAY2_EXCEP", 2
                    Case 3
                        ProcessShift "SCHED_DAY3_SHIFT"
                        ProcessExcep "SCHED_DAY3_EXCEP", 3
                    Case 4
                        ProcessShift "SCHED_DAY4_SHIFT"
                        ProcessExcep "SCHED_DAY4_EXCEP", 4
                    Case 5
                        ProcessShift "SCHED_DAY5_SHIFT"
                        ProcessExcep "SCHED_DAY5_EXCEP", 5
                    Case 6
                        ProcessShift "SCHED_DAY6_SHIFT"
                        ProcessExcep "SCHED_DAY6_EXCEP", 6
                    Case 7
                        ProcessShift "SCHED_DAY7_SHIFT"
                        ProcessExcep "SCHED_DAY7_EXCEP", 7
                    Case 8
                        ProcessShift "SCHED_DAY8_SHIFT"
                        ProcessExcep "SCHED_DAY8_EXCEP", 8
                    Case 9
                        ProcessShift "SCHED_DAY9_SHIFT"
                        ProcessExcep "SCHED_DAY9_EXCEP", 9
                    Case 10
                        ProcessShift "SCHED_DAY10_SHIFT"
                        ProcessExcep "SCHED_DAY10_EXCEP", 10
                    Case 11
                        ProcessShift "SCHED_DAY11_SHIFT"
                        ProcessExcep "SCHED_DAY11_EXCEP", 11
                    Case 12
                        ProcessShift "SCHED_DAY12_SHIFT"
                        ProcessExcep "SCHED_DAY12_EXCEP", 12
                    Case 13
                        ProcessShift "SCHED_DAY13_SHIFT"
                        ProcessExcep "SCHED_DAY13_EXCEP", 13
                    Case 14
                        ProcessShift "SCHED_DAY14_SHIFT"
                        ProcessExcep "SCHED_DAY14_EXCEP", 14
                    Case 15
                        ProcessShift "SCHED_DAY15_SHIFT"
                        ProcessExcep "SCHED_DAY15_EXCEP", 15
                    Case 16
                        ProcessShift "SCHED_DAY16_SHIFT"
                        ProcessExcep "SCHED_DAY16_EXCEP", 16
                    Case 17
                        ProcessShift "SCHED_DAY17_SHIFT"
                        ProcessExcep "SCHED_DAY17_EXCEP", 17
                    Case 18
                        ProcessShift "SCHED_DAY18_SHIFT"
                        ProcessExcep "SCHED_DAY18_EXCEP", 18
                    Case 19
                        ProcessShift "SCHED_DAY19_SHIFT"
                        ProcessExcep "SCHED_DAY19_EXCEP", 19
                    Case 20
                        ProcessShift "SCHED_DAY20_SHIFT"
                        ProcessExcep "SCHED_DAY20_EXCEP", 20
                    Case 21
                        ProcessShift "SCHED_DAY21_SHIFT"
                        ProcessExcep "SCHED_DAY21_EXCEP", 21
                    Case 22
                        ProcessShift "SCHED_DAY22_SHIFT"
                        ProcessExcep "SCHED_DAY22_EXCEP", 22
                    Case 23
                        ProcessShift "SCHED_DAY23_SHIFT"
                        ProcessExcep "SCHED_DAY23_EXCEP", 23
                    Case 24
                        ProcessShift "SCHED_DAY24_SHIFT"
                        ProcessExcep "SCHED_DAY24_EXCEP", 24
                    Case 25
                        ProcessShift "SCHED_DAY25_SHIFT"
                        ProcessExcep "SCHED_DAY25_EXCEP", 25
                    Case 26
                        ProcessShift "SCHED_DAY26_SHIFT"
                        ProcessExcep "SCHED_DAY26_EXCEP", 26
                    Case 27
                        ProcessShift "SCHED_DAY27_SHIFT"
                        ProcessExcep "SCHED_DAY27_EXCEP", 27
                    Case 28
                        ProcessShift "SCHED_DAY28_SHIFT"
                        ProcessExcep "SCHED_DAY28_EXCEP", 28
                    Case 29
                        ProcessShift "SCHED_DAY29_SHIFT"
                        ProcessExcep "SCHED_DAY29_EXCEP", 29
                    Case 30
                        ProcessShift "SCHED_DAY30_SHIFT"
                        ProcessExcep "SCHED_DAY30_EXCEP", 30
                    Case 31
                        ProcessShift "SCHED_DAY31_SHIFT"
                        ProcessExcep "SCHED_DAY31_EXCEP", 31
                    Case 32
                        ProcessShift "SCHED_DAY32_SHIFT"
                        ProcessExcep "SCHED_DAY32_EXCEP", 32
                    Case 33
                        ProcessShift "SCHED_DAY33_SHIFT"
                        ProcessExcep "SCHED_DAY33_EXCEP", 33
                    Case 34
                        ProcessShift "SCHED_DAY34_SHIFT"
                        ProcessExcep "SCHED_DAY34_EXCEP", 34
                    Case 35
                        ProcessShift "SCHED_DAY35_SHIFT"
                        ProcessExcep "SCHED_DAY35_EXCEP", 35
                    Case 36
                        ProcessShift "SCHED_DAY36_SHIFT"
                        ProcessExcep "SCHED_DAY36_EXCEP", 36
                    Case 37
                        ProcessShift "SCHED_DAY37_SHIFT"
                        ProcessExcep "SCHED_DAY37_EXCEP", 37
                    Case 38
                        ProcessShift "SCHED_DAY38_SHIFT"
                        ProcessExcep "SCHED_DAY38_EXCEP", 38
                    Case 39
                        ProcessShift "SCHED_DAY39_SHIFT"
                        ProcessExcep "SCHED_DAY39_EXCEP", 39
                    Case 40
                        ProcessShift "SCHED_DAY40_SHIFT"
                        ProcessExcep "SCHED_DAY40_EXCEP", 40
                    Case 41
                        ProcessShift "SCHED_DAY41_SHIFT"
                        ProcessExcep "SCHED_DAY41_EXCEP", 41
                    Case 42
                        ProcessShift "SCHED_DAY42_SHIFT"
                        ProcessExcep "SCHED_DAY42_EXCEP", 42
                End Select
                If lblnDataFound Then Exit For
            End If
        Next ldtDay
    End With

    Exit Sub
Populate_Record_EH:
    'Raise Error depending on the Error Number
    TrsRaiseError MODULENAME _
            , "Populate_Record"
End Sub
Sub Populate_Relief_Record()
On Error GoTo Populate_Relief_Record_EH

    Dim lintDayCount As Integer
    Dim lblnDataFound As Boolean
    Dim ldtDay As Date
    Dim lvarBookMark As Variant

    With madrsScheduleData
        If mlngPrevRlfEmp <> !EMP_NUM Then
            If Not madrsFormatRlfRptData.EOF Then
                mvarBookMark = madrsFormatRlfRptData.Bookmark
                madrsFormatRlfRptData.Find "EMP_NUM = " & !EMP_NUM
                If madrsFormatRlfRptData.EOF Then
                    mvarBookMark = ""
                    madrsFormatRlfRptData.AddNew
                    madrsFormatRlfRptData.Fields("WG_NUM") = marrWorkGroups(mlngWGIndex)
                    madrsFormatRlfRptData.Fields("WG_DESC") = !WG_DESC
                    madrsFormatRlfRptData.Fields("EMP_NUM") = !EMP_NUM
                    madrsFormatRlfRptData.Fields("EMP_NAME") = !EMP_NAME
                    madrsFormatRlfRptData.Fields("JOB_NUM1") = !JOB_NUM
                    madrsFormatRlfRptData.Fields("JOB_DESC1") = !JOB_DESC
                    madrsFormatRlfRptData.Fields("JOB_NUM2") = ""
                    madrsFormatRlfRptData.Fields("JOB_DESC2") = ""
                    madrsFormatRlfRptData.Fields("QUALS") = !QUALS
                    madrsFormatRlfRptData.Fields("SKILLS") = !SKILLS
                    madrsFormatRlfRptData.Fields("PROJ_ID") = !PROJ_ID
                    madrsFormatRlfRptData.Fields("PROJ_DESC") = !PROJ_DESC
                    madrsFormatRlfRptData.Fields("LINE_NUM") = 0
                    madrsFormatRlfRptData.Fields("REL_REPORT") = "Y"
                    madrsFormatRlfRptData.Fields("CHARGE_WG_NUM") = !WG_NUM
                    madrsFormatRlfRptData.Fields("REQ_LINE_IND") = !REQ_LINE_IND
                    madrsFormatRlfRptData.Fields("DEMAND_LINE") = !WG_LINE_IND
                    InitializeRlfExcep
                End If
            Else
                madrsFormatRlfRptData.AddNew
                madrsFormatRlfRptData.Fields("WG_NUM") = marrWorkGroups(mlngWGIndex)
                madrsFormatRlfRptData.Fields("WG_DESC") = !WG_DESC
                madrsFormatRlfRptData.Fields("EMP_NUM") = !EMP_NUM
                madrsFormatRlfRptData.Fields("EMP_NAME") = !EMP_NAME
                madrsFormatRlfRptData.Fields("JOB_NUM1") = !JOB_NUM
                madrsFormatRlfRptData.Fields("JOB_DESC1") = !JOB_DESC
                madrsFormatRlfRptData.Fields("JOB_NUM2") = ""
                madrsFormatRlfRptData.Fields("JOB_DESC2") = ""
                madrsFormatRlfRptData.Fields("QUALS") = !QUALS
                madrsFormatRlfRptData.Fields("SKILLS") = !SKILLS
                madrsFormatRlfRptData.Fields("PROJ_ID") = !PROJ_ID
                madrsFormatRlfRptData.Fields("PROJ_DESC") = !PROJ_DESC
                madrsFormatRlfRptData.Fields("LINE_NUM") = 0
                madrsFormatRlfRptData.Fields("REL_REPORT") = "Y"
                madrsFormatRlfRptData.Fields("CHARGE_WG_NUM") = !WG_NUM
                madrsFormatRlfRptData.Fields("REQ_LINE_IND") = !REQ_LINE_IND
                madrsFormatRlfRptData.Fields("DEMAND_LINE") = !WG_LINE_IND
                InitializeRlfExcep
            End If
        End If
        lintDayCount = 0
        lblnDataFound = False
        If mdtPrevChargeDt <> !CHARGE_DATE Or mlngPrevEmp <> !EMP_NUM Then
            mblnRlfSecondShift = False
            mblnRlfJobPosition = False
        End If
        For ldtDay = mdtStartDate To mdtEndDate
            lintDayCount = lintDayCount + 1
            If !CHARGE_DATE = ldtDay Then
                lblnDataFound = True
                Select Case lintDayCount
                    Case 1
                        ProcessRlfShift "SCHED_DAY1_SHIFT"
                        ProcessRlfExcep "SCHED_DAY1_EXCEP", 1
                    Case 2
                        ProcessRlfShift "SCHED_DAY2_SHIFT"
                        ProcessRlfExcep "SCHED_DAY2_EXCEP", 2
                    Case 3
                        ProcessRlfShift "SCHED_DAY3_SHIFT"
                        ProcessRlfExcep "SCHED_DAY3_EXCEP", 3
                    Case 4
                        ProcessRlfShift "SCHED_DAY4_SHIFT"
                        ProcessRlfExcep "SCHED_DAY4_EXCEP", 4
                    Case 5
                        ProcessRlfShift "SCHED_DAY5_SHIFT"
                        ProcessRlfExcep "SCHED_DAY5_EXCEP", 5
                    Case 6
                        ProcessRlfShift "SCHED_DAY6_SHIFT"
                        ProcessRlfExcep "SCHED_DAY6_EXCEP", 6
                    Case 7
                        ProcessRlfShift "SCHED_DAY7_SHIFT"
                        ProcessRlfExcep "SCHED_DAY7_EXCEP", 7
                    Case 8
                        ProcessRlfShift "SCHED_DAY8_SHIFT"
                        ProcessRlfExcep "SCHED_DAY8_EXCEP", 8
                    Case 9
                        ProcessRlfShift "SCHED_DAY9_SHIFT"
                        ProcessRlfExcep "SCHED_DAY9_EXCEP", 9
                    Case 10
                        ProcessRlfShift "SCHED_DAY10_SHIFT"
                        ProcessRlfExcep "SCHED_DAY10_EXCEP", 10
                    Case 11
                        ProcessRlfShift "SCHED_DAY11_SHIFT"
                        ProcessRlfExcep "SCHED_DAY11_EXCEP", 11
                    Case 12
                        ProcessRlfShift "SCHED_DAY12_SHIFT"
                        ProcessRlfExcep "SCHED_DAY12_EXCEP", 12
                    Case 13
                        ProcessRlfShift "SCHED_DAY13_SHIFT"
                        ProcessRlfExcep "SCHED_DAY13_EXCEP", 13
                    Case 14
                        ProcessRlfShift "SCHED_DAY14_SHIFT"
                        ProcessRlfExcep "SCHED_DAY14_EXCEP", 14
                    Case 15
                        ProcessRlfShift "SCHED_DAY15_SHIFT"
                        ProcessRlfExcep "SCHED_DAY15_EXCEP", 15
                    Case 16
                        ProcessRlfShift "SCHED_DAY16_SHIFT"
                        ProcessRlfExcep "SCHED_DAY16_EXCEP", 16
                    Case 17
                        ProcessRlfShift "SCHED_DAY17_SHIFT"
                        ProcessRlfExcep "SCHED_DAY17_EXCEP", 17
                    Case 18
                        ProcessRlfShift "SCHED_DAY18_SHIFT"
                        ProcessRlfExcep "SCHED_DAY18_EXCEP", 18
                    Case 19
                        ProcessRlfShift "SCHED_DAY19_SHIFT"
                        ProcessRlfExcep "SCHED_DAY19_EXCEP", 19
                    Case 20
                        ProcessRlfShift "SCHED_DAY20_SHIFT"
                        ProcessRlfExcep "SCHED_DAY20_EXCEP", 20
                    Case 21
                        ProcessRlfShift "SCHED_DAY21_SHIFT"
                        ProcessRlfExcep "SCHED_DAY21_EXCEP", 21
                    Case 22
                        ProcessRlfShift "SCHED_DAY22_SHIFT"
                        ProcessRlfExcep "SCHED_DAY22_EXCEP", 22
                    Case 23
                        ProcessRlfShift "SCHED_DAY23_SHIFT"
                        ProcessRlfExcep "SCHED_DAY23_EXCEP", 23
                    Case 24
                        ProcessRlfShift "SCHED_DAY24_SHIFT"
                        ProcessRlfExcep "SCHED_DAY24_EXCEP", 24
                    Case 25
                        ProcessRlfShift "SCHED_DAY25_SHIFT"
                        ProcessRlfExcep "SCHED_DAY25_EXCEP", 25
                    Case 26
                        ProcessRlfShift "SCHED_DAY26_SHIFT"
                        ProcessRlfExcep "SCHED_DAY26_EXCEP", 26
                    Case 27
                        ProcessRlfShift "SCHED_DAY27_SHIFT"
                        ProcessRlfExcep "SCHED_DAY27_EXCEP", 27
                    Case 28
                        ProcessRlfShift "SCHED_DAY28_SHIFT"
                        ProcessRlfExcep "SCHED_DAY28_EXCEP", 28
                    Case 29
                        ProcessRlfShift "SCHED_DAY29_SHIFT"
                        ProcessRlfExcep "SCHED_DAY29_EXCEP", 29
                    Case 30
                        ProcessRlfShift "SCHED_DAY30_SHIFT"
                        ProcessRlfExcep "SCHED_DAY30_EXCEP", 30
                    Case 31
                        ProcessRlfShift "SCHED_DAY31_SHIFT"
                        ProcessRlfExcep "SCHED_DAY31_EXCEP", 31
                    Case 32
                        ProcessRlfShift "SCHED_DAY32_SHIFT"
                        ProcessRlfExcep "SCHED_DAY32_EXCEP", 32
                    Case 33
                        ProcessRlfShift "SCHED_DAY33_SHIFT"
                        ProcessRlfExcep "SCHED_DAY33_EXCEP", 33
                    Case 34
                        ProcessRlfShift "SCHED_DAY34_SHIFT"
                        ProcessRlfExcep "SCHED_DAY34_EXCEP", 34
                    Case 35
                        ProcessRlfShift "SCHED_DAY35_SHIFT"
                        ProcessRlfExcep "SCHED_DAY35_EXCEP", 35
                    Case 36
                        ProcessRlfShift "SCHED_DAY36_SHIFT"
                        ProcessRlfExcep "SCHED_DAY36_EXCEP", 36
                    Case 37
                        ProcessRlfShift "SCHED_DAY37_SHIFT"
                        ProcessRlfExcep "SCHED_DAY37_EXCEP", 37
                    Case 38
                        ProcessRlfShift "SCHED_DAY38_SHIFT"
                        ProcessRlfExcep "SCHED_DAY38_EXCEP", 38
                    Case 39
                        ProcessRlfShift "SCHED_DAY39_SHIFT"
                        ProcessRlfExcep "SCHED_DAY39_EXCEP", 39
                    Case 40
                        ProcessRlfShift "SCHED_DAY40_SHIFT"
                        ProcessRlfExcep "SCHED_DAY40_EXCEP", 40
                    Case 41
                        ProcessRlfShift "SCHED_DAY41_SHIFT"
                        ProcessRlfExcep "SCHED_DAY41_EXCEP", 41
                    Case 42
                        ProcessRlfShift "SCHED_DAY42_SHIFT"
                        ProcessRlfExcep "SCHED_DAY42_EXCEP", 42
                End Select
                If lblnDataFound Then Exit For
            End If
        Next ldtDay
    End With

    Exit Sub
Populate_Relief_Record_EH:
    'Raise Error depending on the Error Number
    TrsRaiseError MODULENAME _
            , "Populate_Relief_Record"

End Sub
Sub ProcessDayFirst()
On Error GoTo ProcessDayFirst_EH

    With madrsReportData
        !SCHED_DAY_1 = madrsFormatRptData.Fields("SCHED_DAY1_SHIFT") & vbCrLf & _
                       madrsFormatRptData.Fields("SCHED_DAY1_EXCEP")
        If Trim(!SCHED_DAY_1) <> "" Then
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_1, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        !SCHED_DAY_2 = madrsFormatRptData.Fields("SCHED_DAY2_SHIFT") & vbCrLf & _
                       madrsFormatRptData.Fields("SCHED_DAY2_EXCEP")
        If Trim(!SCHED_DAY_2) <> "" Then
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_2, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        !SCHED_DAY_3 = madrsFormatRptData.Fields("SCHED_DAY3_SHIFT") & vbCrLf & _
                       madrsFormatRptData.Fields("SCHED_DAY3_EXCEP")
        If Trim(!SCHED_DAY_3) <> "" Then
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_3, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        !sched_day_4 = madrsFormatRptData.Fields("SCHED_DAY4_SHIFT") & vbCrLf & _
                       madrsFormatRptData.Fields("SCHED_DAY4_EXCEP")
        If Trim(!sched_day_4) <> "" Then
            mlngPrevLineCount = CountOccurrences(!sched_day_4, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        !SCHED_DAY_5 = madrsFormatRptData.Fields("SCHED_DAY5_SHIFT") & vbCrLf & _
                       madrsFormatRptData.Fields("SCHED_DAY5_EXCEP")
        If Trim(!SCHED_DAY_5) <> "" Then
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_5, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        !SCHED_DAY_6 = madrsFormatRptData.Fields("SCHED_DAY6_SHIFT") & vbCrLf & _
                       madrsFormatRptData.Fields("SCHED_DAY6_EXCEP")
        If Trim(!SCHED_DAY_6) <> "" Then
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_6, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        !SCHED_DAY_7 = madrsFormatRptData.Fields("SCHED_DAY7_SHIFT") & vbCrLf & _
                       madrsFormatRptData.Fields("SCHED_DAY7_EXCEP")
        If Trim(!SCHED_DAY_7) <> "" Then
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_7, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        !SCHED_DAY_8 = madrsFormatRptData.Fields("SCHED_DAY8_SHIFT") & vbCrLf & _
                       madrsFormatRptData.Fields("SCHED_DAY8_EXCEP")
        If Trim(!SCHED_DAY_8) <> "" Then
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_8, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        !SCHED_DAY_9 = madrsFormatRptData.Fields("SCHED_DAY9_SHIFT") & vbCrLf & _
                       madrsFormatRptData.Fields("SCHED_DAY9_EXCEP")
        If Trim(!SCHED_DAY_9) <> "" Then
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_9, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        !SCHED_DAY_10 = madrsFormatRptData.Fields("SCHED_DAY10_SHIFT") & vbCrLf & _
                        madrsFormatRptData.Fields("SCHED_DAY10_EXCEP")
        If Trim(!SCHED_DAY_10) <> "" Then
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_10, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        !SCHED_DAY_11 = madrsFormatRptData.Fields("SCHED_DAY11_SHIFT") & vbCrLf & _
                        madrsFormatRptData.Fields("SCHED_DAY11_EXCEP")
        If Trim(!SCHED_DAY_11) <> "" Then
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_11, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        !SCHED_DAY_12 = madrsFormatRptData.Fields("SCHED_DAY12_SHIFT") & vbCrLf & _
                        madrsFormatRptData.Fields("SCHED_DAY12_EXCEP")
        If Trim(!SCHED_DAY_12) <> "" Then
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_12, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        !SCHED_DAY_13 = madrsFormatRptData.Fields("SCHED_DAY13_SHIFT") & vbCrLf & _
                        madrsFormatRptData.Fields("SCHED_DAY13_EXCEP")
        If Trim(!SCHED_DAY_13) <> "" Then
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_13, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        !SCHED_DAY_14 = madrsFormatRptData.Fields("SCHED_DAY14_SHIFT") & vbCrLf & _
                        madrsFormatRptData.Fields("SCHED_DAY14_EXCEP")
        If Trim(!SCHED_DAY_14) <> "" Then
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_14, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        !SCHED_DAY_15 = madrsFormatRptData.Fields("SCHED_DAY15_SHIFT") & vbCrLf & _
                        madrsFormatRptData.Fields("SCHED_DAY15_EXCEP")
        If Trim(!SCHED_DAY_15) <> "" Then
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_15, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        !SCHED_DAY_16 = madrsFormatRptData.Fields("SCHED_DAY16_SHIFT") & vbCrLf & _
                        madrsFormatRptData.Fields("SCHED_DAY16_EXCEP")
        If Trim(!SCHED_DAY_16) <> "" Then
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_16, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        !SCHED_DAY_17 = madrsFormatRptData.Fields("SCHED_DAY17_SHIFT") & vbCrLf & _
                        madrsFormatRptData.Fields("SCHED_DAY17_EXCEP")
        If Trim(!SCHED_DAY_17) <> "" Then
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_17, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        !SCHED_DAY_18 = madrsFormatRptData.Fields("SCHED_DAY18_SHIFT") & vbCrLf & _
                        madrsFormatRptData.Fields("SCHED_DAY18_EXCEP")
        If Trim(!SCHED_DAY_18) <> "" Then
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_18, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        !SCHED_DAY_19 = madrsFormatRptData.Fields("SCHED_DAY19_SHIFT") & vbCrLf & _
                        madrsFormatRptData.Fields("SCHED_DAY19_EXCEP")
        If Trim(!SCHED_DAY_19) <> "" Then
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_19, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        !SCHED_DAY_20 = madrsFormatRptData.Fields("SCHED_DAY20_SHIFT") & vbCrLf & _
                        madrsFormatRptData.Fields("SCHED_DAY20_EXCEP")
        If Trim(!SCHED_DAY_20) <> "" Then
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_20, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        !SCHED_DAY_21 = madrsFormatRptData.Fields("SCHED_DAY21_SHIFT") & vbCrLf & _
                        madrsFormatRptData.Fields("SCHED_DAY21_EXCEP")
        If Trim(!SCHED_DAY_21) <> "" Then
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_21, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        !SCHED_DAY_22 = madrsFormatRptData.Fields("SCHED_DAY22_SHIFT") & vbCrLf & _
                        madrsFormatRptData.Fields("SCHED_DAY22_EXCEP")
        If Trim(!SCHED_DAY_22) <> "" Then
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_22, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        !SCHED_DAY_23 = madrsFormatRptData.Fields("SCHED_DAY23_SHIFT") & vbCrLf & _
                        madrsFormatRptData.Fields("SCHED_DAY23_EXCEP")
        If Trim(!SCHED_DAY_23) <> "" Then
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_23, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        !SCHED_DAY_24 = madrsFormatRptData.Fields("SCHED_DAY24_SHIFT") & vbCrLf & _
                        madrsFormatRptData.Fields("SCHED_DAY24_EXCEP")
        If Trim(!SCHED_DAY_24) <> "" Then
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_24, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        !SCHED_DAY_25 = madrsFormatRptData.Fields("SCHED_DAY25_SHIFT") & vbCrLf & _
                        madrsFormatRptData.Fields("SCHED_DAY25_EXCEP")
        If Trim(!SCHED_DAY_25) <> "" Then
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_25, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        !SCHED_DAY_26 = madrsFormatRptData.Fields("SCHED_DAY26_SHIFT") & vbCrLf & _
                        madrsFormatRptData.Fields("SCHED_DAY26_EXCEP")
        If Trim(!SCHED_DAY_26) <> "" Then
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_26, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        !SCHED_DAY_27 = madrsFormatRptData.Fields("SCHED_DAY27_SHIFT") & vbCrLf & _
                        madrsFormatRptData.Fields("SCHED_DAY27_EXCEP")
        If Trim(!SCHED_DAY_27) <> "" Then
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_27, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        !SCHED_DAY_28 = madrsFormatRptData.Fields("SCHED_DAY28_SHIFT") & vbCrLf & _
                        madrsFormatRptData.Fields("SCHED_DAY28_EXCEP")
        If Trim(!SCHED_DAY_28) <> "" Then
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_28, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        !SCHED_DAY_29 = madrsFormatRptData.Fields("SCHED_DAY29_SHIFT") & vbCrLf & _
                        madrsFormatRptData.Fields("SCHED_DAY29_EXCEP")
        If Trim(!SCHED_DAY_29) <> "" Then
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_29, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        !SCHED_DAY_30 = madrsFormatRptData.Fields("SCHED_DAY30_SHIFT") & vbCrLf & _
                        madrsFormatRptData.Fields("SCHED_DAY30_EXCEP")
        If Trim(!SCHED_DAY_30) <> "" Then
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_30, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        !SCHED_DAY_31 = madrsFormatRptData.Fields("SCHED_DAY31_SHIFT") & vbCrLf & _
                        madrsFormatRptData.Fields("SCHED_DAY31_EXCEP")
        If Trim(!SCHED_DAY_31) <> "" Then
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_31, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        !SCHED_DAY_32 = madrsFormatRptData.Fields("SCHED_DAY32_SHIFT") & vbCrLf & _
                        madrsFormatRptData.Fields("SCHED_DAY32_EXCEP")
        If Trim(!SCHED_DAY_32) <> "" Then
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_32, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        !SCHED_DAY_33 = madrsFormatRptData.Fields("SCHED_DAY33_SHIFT") & vbCrLf & _
                        madrsFormatRptData.Fields("SCHED_DAY33_EXCEP")
        If Trim(!SCHED_DAY_33) <> "" Then
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_33, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        !SCHED_DAY_34 = madrsFormatRptData.Fields("SCHED_DAY34_SHIFT") & vbCrLf & _
                        madrsFormatRptData.Fields("SCHED_DAY34_EXCEP")
        If Trim(!SCHED_DAY_34) <> "" Then
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_34, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        !SCHED_DAY_35 = madrsFormatRptData.Fields("SCHED_DAY35_SHIFT") & vbCrLf & _
                        madrsFormatRptData.Fields("SCHED_DAY35_EXCEP")
        If Trim(!SCHED_DAY_35) <> "" Then
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_35, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        !SCHED_DAY_36 = madrsFormatRptData.Fields("SCHED_DAY36_SHIFT") & vbCrLf & _
                        madrsFormatRptData.Fields("SCHED_DAY36_EXCEP")
        If Trim(!SCHED_DAY_36) <> "" Then
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_36, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        !SCHED_DAY_37 = madrsFormatRptData.Fields("SCHED_DAY37_SHIFT") & vbCrLf & _
                        madrsFormatRptData.Fields("SCHED_DAY37_EXCEP")
        If Trim(!SCHED_DAY_37) <> "" Then
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_37, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        !SCHED_DAY_38 = madrsFormatRptData.Fields("SCHED_DAY38_SHIFT") & vbCrLf & _
                        madrsFormatRptData.Fields("SCHED_DAY38_EXCEP")
        If Trim(!SCHED_DAY_38) <> "" Then
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_38, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        !SCHED_DAY_39 = madrsFormatRptData.Fields("SCHED_DAY39_SHIFT") & vbCrLf & _
                        madrsFormatRptData.Fields("SCHED_DAY39_EXCEP")
        If Trim(!SCHED_DAY_39) <> "" Then
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_39, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        !SCHED_DAY_40 = madrsFormatRptData.Fields("SCHED_DAY40_SHIFT") & vbCrLf & _
                        madrsFormatRptData.Fields("SCHED_DAY40_EXCEP")
        If Trim(!SCHED_DAY_40) <> "" Then
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_40, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        !SCHED_DAY_41 = madrsFormatRptData.Fields("SCHED_DAY41_SHIFT") & vbCrLf & _
                        madrsFormatRptData.Fields("SCHED_DAY41_EXCEP")
        If Trim(!SCHED_DAY_41) <> "" Then
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_41, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        !SCHED_DAY_42 = madrsFormatRptData.Fields("SCHED_DAY42_SHIFT") & vbCrLf & _
                        madrsFormatRptData.Fields("SCHED_DAY42_EXCEP")
        If Trim(!SCHED_DAY_42) <> "" Then
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_42, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
    End With
                
    Exit Sub

ProcessDayFirst_EH:
'Raise Error depending on the Error Number
    TrsRaiseError MODULENAME _
            , "ProcessDayFirst"

End Sub
Private Sub ProcessDay()
On Error GoTo ProcessDay_EH

Dim lstrshift As String

    With madrsReportData
        lstrshift = Replace(!SCHED_DAY_1, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !SCHED_DAY_1 = madrsFormatRptData.Fields("SCHED_DAY1_SHIFT") & vbCrLf & _
                        madrsFormatRptData.Fields("SCHED_DAY1_EXCEP")
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_1, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        lstrshift = Replace(!SCHED_DAY_2, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !SCHED_DAY_2 = madrsFormatRptData.Fields("SCHED_DAY2_SHIFT") & vbCrLf & _
                        madrsFormatRptData.Fields("SCHED_DAY2_EXCEP")
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_2, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        lstrshift = Replace(!SCHED_DAY_3, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !SCHED_DAY_3 = madrsFormatRptData.Fields("SCHED_DAY3_SHIFT") & vbCrLf & _
                        madrsFormatRptData.Fields("SCHED_DAY3_EXCEP")
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_3, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        lstrshift = Replace(!sched_day_4, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !sched_day_4 = madrsFormatRptData.Fields("SCHED_DAY4_SHIFT") & vbCrLf & _
                        madrsFormatRptData.Fields("SCHED_DAY4_EXCEP")
            mlngPrevLineCount = CountOccurrences(!sched_day_4, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        lstrshift = Replace(!SCHED_DAY_5, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !SCHED_DAY_5 = madrsFormatRptData.Fields("SCHED_DAY5_SHIFT") & vbCrLf & _
                       madrsFormatRptData.Fields("SCHED_DAY5_EXCEP")
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_5, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        lstrshift = Replace(!SCHED_DAY_6, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !SCHED_DAY_6 = madrsFormatRptData.Fields("SCHED_DAY6_SHIFT") & vbCrLf & _
                        madrsFormatRptData.Fields("SCHED_DAY6_EXCEP")
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_6, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        lstrshift = Replace(!SCHED_DAY_7, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !SCHED_DAY_7 = madrsFormatRptData.Fields("SCHED_DAY7_SHIFT") & vbCrLf & _
                        madrsFormatRptData.Fields("SCHED_DAY7_EXCEP")
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_7, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        lstrshift = Replace(!SCHED_DAY_8, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !SCHED_DAY_8 = madrsFormatRptData.Fields("SCHED_DAY8_SHIFT") & vbCrLf & _
                        madrsFormatRptData.Fields("SCHED_DAY8_EXCEP")
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_8, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        lstrshift = Replace(!SCHED_DAY_9, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !SCHED_DAY_9 = madrsFormatRptData.Fields("SCHED_DAY9_SHIFT") & vbCrLf & _
                        madrsFormatRptData.Fields("SCHED_DAY9_EXCEP")
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_9, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        lstrshift = Replace(!SCHED_DAY_10, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !SCHED_DAY_10 = madrsFormatRptData.Fields("SCHED_DAY10_SHIFT") & vbCrLf & _
                        madrsFormatRptData.Fields("SCHED_DAY10_EXCEP")
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_10, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        lstrshift = Replace(!SCHED_DAY_11, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !SCHED_DAY_11 = madrsFormatRptData.Fields("SCHED_DAY11_SHIFT") & vbCrLf & _
                        madrsFormatRptData.Fields("SCHED_DAY11_EXCEP")
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_11, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        lstrshift = Replace(!SCHED_DAY_12, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !SCHED_DAY_12 = madrsFormatRptData.Fields("SCHED_DAY12_SHIFT") & vbCrLf & _
                        madrsFormatRptData.Fields("SCHED_DAY12_EXCEP")
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_12, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        lstrshift = Replace(!SCHED_DAY_13, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !SCHED_DAY_13 = madrsFormatRptData.Fields("SCHED_DAY13_SHIFT") & vbCrLf & _
                        madrsFormatRptData.Fields("SCHED_DAY13_EXCEP")
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_13, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        lstrshift = Replace(!SCHED_DAY_14, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !SCHED_DAY_14 = madrsFormatRptData.Fields("SCHED_DAY14_SHIFT") & vbCrLf & _
                        madrsFormatRptData.Fields("SCHED_DAY14_EXCEP")
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_14, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        lstrshift = Replace(!SCHED_DAY_15, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !SCHED_DAY_15 = madrsFormatRptData.Fields("SCHED_DAY15_SHIFT") & vbCrLf & _
                        madrsFormatRptData.Fields("SCHED_DAY15_EXCEP")
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_15, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        lstrshift = Replace(!SCHED_DAY_16, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !SCHED_DAY_16 = madrsFormatRptData.Fields("SCHED_DAY16_SHIFT") & vbCrLf & _
                        madrsFormatRptData.Fields("SCHED_DAY16_EXCEP")
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_16, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        lstrshift = Replace(!SCHED_DAY_17, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !SCHED_DAY_17 = madrsFormatRptData.Fields("SCHED_DAY17_SHIFT") & vbCrLf & _
                        madrsFormatRptData.Fields("SCHED_DAY17_EXCEP")
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_17, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        lstrshift = Replace(!SCHED_DAY_18, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !SCHED_DAY_18 = madrsFormatRptData.Fields("SCHED_DAY18_SHIFT") & vbCrLf & _
                        madrsFormatRptData.Fields("SCHED_DAY18_EXCEP")
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_18, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        lstrshift = Replace(!SCHED_DAY_19, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !SCHED_DAY_19 = madrsFormatRptData.Fields("SCHED_DAY19_SHIFT") & vbCrLf & _
                        madrsFormatRptData.Fields("SCHED_DAY19_EXCEP")
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_19, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        lstrshift = Replace(!SCHED_DAY_20, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !SCHED_DAY_20 = madrsFormatRptData.Fields("SCHED_DAY20_SHIFT") & vbCrLf & _
                        madrsFormatRptData.Fields("SCHED_DAY20_EXCEP")
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_20, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        lstrshift = Replace(!SCHED_DAY_21, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !SCHED_DAY_21 = madrsFormatRptData.Fields("SCHED_DAY21_SHIFT") & vbCrLf & _
                        madrsFormatRptData.Fields("SCHED_DAY21_EXCEP")
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_21, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        lstrshift = Replace(!SCHED_DAY_22, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !SCHED_DAY_22 = madrsFormatRptData.Fields("SCHED_DAY22_SHIFT") & vbCrLf & _
                        madrsFormatRptData.Fields("SCHED_DAY22_EXCEP")
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_22, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        lstrshift = Replace(!SCHED_DAY_23, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !SCHED_DAY_23 = madrsFormatRptData.Fields("SCHED_DAY23_SHIFT") & vbCrLf & _
                        madrsFormatRptData.Fields("SCHED_DAY23_EXCEP")
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_23, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        lstrshift = Replace(!SCHED_DAY_24, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !SCHED_DAY_24 = madrsFormatRptData.Fields("SCHED_DAY24_SHIFT") & vbCrLf & _
                        madrsFormatRptData.Fields("SCHED_DAY24_EXCEP")
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_24, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        lstrshift = Replace(!SCHED_DAY_25, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !SCHED_DAY_25 = madrsFormatRptData.Fields("SCHED_DAY25_SHIFT") & vbCrLf & _
                        madrsFormatRptData.Fields("SCHED_DAY25_EXCEP")
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_25, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        lstrshift = Replace(!SCHED_DAY_26, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !SCHED_DAY_26 = madrsFormatRptData.Fields("SCHED_DAY26_SHIFT") & vbCrLf & _
                        madrsFormatRptData.Fields("SCHED_DAY26_EXCEP")
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_26, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        lstrshift = Replace(!SCHED_DAY_27, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !SCHED_DAY_27 = madrsFormatRptData.Fields("SCHED_DAY27_SHIFT") & vbCrLf & _
                        madrsFormatRptData.Fields("SCHED_DAY27_EXCEP")
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_27, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        lstrshift = Replace(!SCHED_DAY_28, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !SCHED_DAY_28 = madrsFormatRptData.Fields("SCHED_DAY28_SHIFT") & vbCrLf & _
                        madrsFormatRptData.Fields("SCHED_DAY28_EXCEP")
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_28, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        lstrshift = Replace(!SCHED_DAY_29, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !SCHED_DAY_29 = madrsFormatRptData.Fields("SCHED_DAY29_SHIFT") & vbCrLf & _
                        madrsFormatRptData.Fields("SCHED_DAY29_EXCEP")
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_29, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        lstrshift = Replace(!SCHED_DAY_30, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !SCHED_DAY_30 = madrsFormatRptData.Fields("SCHED_DAY30_SHIFT") & vbCrLf & _
                        madrsFormatRptData.Fields("SCHED_DAY30_EXCEP")
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_30, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        lstrshift = Replace(!SCHED_DAY_31, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !SCHED_DAY_31 = madrsFormatRptData.Fields("SCHED_DAY31_SHIFT") & vbCrLf & _
                        madrsFormatRptData.Fields("SCHED_DAY31_EXCEP")
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_31, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        lstrshift = Replace(!SCHED_DAY_32, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !SCHED_DAY_32 = madrsFormatRptData.Fields("SCHED_DAY32_SHIFT") & vbCrLf & _
                        madrsFormatRptData.Fields("SCHED_DAY32_EXCEP")
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_32, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        lstrshift = Replace(!SCHED_DAY_33, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !SCHED_DAY_33 = madrsFormatRptData.Fields("SCHED_DAY33_SHIFT") & vbCrLf & _
                        madrsFormatRptData.Fields("SCHED_DAY33_EXCEP")
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_33, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        lstrshift = Replace(!SCHED_DAY_34, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !SCHED_DAY_34 = madrsFormatRptData.Fields("SCHED_DAY34_SHIFT") & vbCrLf & _
                        madrsFormatRptData.Fields("SCHED_DAY34_EXCEP")
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_34, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        lstrshift = Replace(!SCHED_DAY_35, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !SCHED_DAY_35 = madrsFormatRptData.Fields("SCHED_DAY35_SHIFT") & vbCrLf & _
                        madrsFormatRptData.Fields("SCHED_DAY35_EXCEP")
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_35, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        lstrshift = Replace(!SCHED_DAY_36, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !SCHED_DAY_36 = madrsFormatRptData.Fields("SCHED_DAY36_SHIFT") & vbCrLf & _
                        madrsFormatRptData.Fields("SCHED_DAY36_EXCEP")
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_36, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        lstrshift = Replace(!SCHED_DAY_37, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !SCHED_DAY_37 = madrsFormatRptData.Fields("SCHED_DAY37_SHIFT") & vbCrLf & _
                        madrsFormatRptData.Fields("SCHED_DAY37_EXCEP")
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_37, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        lstrshift = Replace(!SCHED_DAY_38, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !SCHED_DAY_38 = madrsFormatRptData.Fields("SCHED_DAY38_SHIFT") & vbCrLf & _
                        madrsFormatRptData.Fields("SCHED_DAY38_EXCEP")
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_38, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        lstrshift = Replace(!SCHED_DAY_39, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !SCHED_DAY_39 = madrsFormatRptData.Fields("SCHED_DAY39_SHIFT") & vbCrLf & _
                        madrsFormatRptData.Fields("SCHED_DAY39_EXCEP")
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_39, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        lstrshift = Replace(!SCHED_DAY_40, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !SCHED_DAY_40 = madrsFormatRptData.Fields("SCHED_DAY40_SHIFT") & vbCrLf & _
                        madrsFormatRptData.Fields("SCHED_DAY40_EXCEP")
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_40, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        lstrshift = Replace(!SCHED_DAY_41, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !SCHED_DAY_41 = madrsFormatRptData.Fields("SCHED_DAY41_SHIFT") & vbCrLf & _
                        madrsFormatRptData.Fields("SCHED_DAY41_EXCEP")
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_41, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        lstrshift = Replace(!SCHED_DAY_42, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !SCHED_DAY_42 = madrsFormatRptData.Fields("SCHED_DAY42_SHIFT") & vbCrLf & _
                        madrsFormatRptData.Fields("SCHED_DAY42_EXCEP")
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_42, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
    End With
                
    Exit Sub

ProcessDay_EH:
'Raise Error depending on the Error Number
    TrsRaiseError MODULENAME _
            , "ProcessDay"

End Sub
Sub ProcessRlfDayFirst()
On Error GoTo ProcessRlfDayFirst_EH

    With madrsReportData
        !SCHED_DAY_1 = madrsFormatRlfRptData.Fields("SCHED_DAY1_SHIFT") & vbCrLf & _
                       madrsFormatRlfRptData.Fields("SCHED_DAY1_EXCEP")
        If Trim(!SCHED_DAY_1) <> "" Then
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_1, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        !SCHED_DAY_2 = madrsFormatRlfRptData.Fields("SCHED_DAY2_SHIFT") & vbCrLf & _
                       madrsFormatRlfRptData.Fields("SCHED_DAY2_EXCEP")
        If Trim(!SCHED_DAY_2) <> "" Then
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_2, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        !SCHED_DAY_3 = madrsFormatRlfRptData.Fields("SCHED_DAY3_SHIFT") & vbCrLf & _
                       madrsFormatRlfRptData.Fields("SCHED_DAY3_EXCEP")
        If Trim(!SCHED_DAY_3) <> "" Then
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_3, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        !sched_day_4 = madrsFormatRlfRptData.Fields("SCHED_DAY4_SHIFT") & vbCrLf & _
                       madrsFormatRlfRptData.Fields("SCHED_DAY4_EXCEP")
        If Trim(!sched_day_4) <> "" Then
            mlngPrevLineCount = CountOccurrences(!sched_day_4, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        !SCHED_DAY_5 = madrsFormatRlfRptData.Fields("SCHED_DAY5_SHIFT") & vbCrLf & _
                       madrsFormatRlfRptData.Fields("SCHED_DAY5_EXCEP")
        If Trim(!SCHED_DAY_5) <> "" Then
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_5, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        !SCHED_DAY_6 = madrsFormatRlfRptData.Fields("SCHED_DAY6_SHIFT") & vbCrLf & _
                       madrsFormatRlfRptData.Fields("SCHED_DAY6_EXCEP")
        If Trim(!SCHED_DAY_6) <> "" Then
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_6, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        !SCHED_DAY_7 = madrsFormatRlfRptData.Fields("SCHED_DAY7_SHIFT") & vbCrLf & _
                       madrsFormatRlfRptData.Fields("SCHED_DAY7_EXCEP")
        If Trim(!SCHED_DAY_7) <> "" Then
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_7, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        !SCHED_DAY_8 = madrsFormatRlfRptData.Fields("SCHED_DAY8_SHIFT") & vbCrLf & _
                       madrsFormatRlfRptData.Fields("SCHED_DAY8_EXCEP")
        If Trim(!SCHED_DAY_8) <> "" Then
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_8, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        !SCHED_DAY_9 = madrsFormatRlfRptData.Fields("SCHED_DAY9_SHIFT") & vbCrLf & _
                       madrsFormatRlfRptData.Fields("SCHED_DAY9_EXCEP")
        If Trim(!SCHED_DAY_9) <> "" Then
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_9, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        !SCHED_DAY_10 = madrsFormatRlfRptData.Fields("SCHED_DAY10_SHIFT") & vbCrLf & _
                        madrsFormatRlfRptData.Fields("SCHED_DAY10_EXCEP")
        If Trim(!SCHED_DAY_10) <> "" Then
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_10, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        !SCHED_DAY_11 = madrsFormatRlfRptData.Fields("SCHED_DAY11_SHIFT") & vbCrLf & _
                        madrsFormatRlfRptData.Fields("SCHED_DAY11_EXCEP")
        If Trim(!SCHED_DAY_11) <> "" Then
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_11, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        !SCHED_DAY_12 = madrsFormatRlfRptData.Fields("SCHED_DAY12_SHIFT") & vbCrLf & _
                        madrsFormatRlfRptData.Fields("SCHED_DAY12_EXCEP")
        If Trim(!SCHED_DAY_12) <> "" Then
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_12, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        !SCHED_DAY_13 = madrsFormatRlfRptData.Fields("SCHED_DAY13_SHIFT") & vbCrLf & _
                        madrsFormatRlfRptData.Fields("SCHED_DAY13_EXCEP")
        If Trim(!SCHED_DAY_13) <> "" Then
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_13, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        !SCHED_DAY_14 = madrsFormatRlfRptData.Fields("SCHED_DAY14_SHIFT") & vbCrLf & _
                        madrsFormatRlfRptData.Fields("SCHED_DAY14_EXCEP")
        If Trim(!SCHED_DAY_14) <> "" Then
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_14, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        !SCHED_DAY_15 = madrsFormatRlfRptData.Fields("SCHED_DAY15_SHIFT") & vbCrLf & _
                        madrsFormatRlfRptData.Fields("SCHED_DAY15_EXCEP")
        If Trim(!SCHED_DAY_15) <> "" Then
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_15, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        !SCHED_DAY_16 = madrsFormatRlfRptData.Fields("SCHED_DAY16_SHIFT") & vbCrLf & _
                        madrsFormatRlfRptData.Fields("SCHED_DAY16_EXCEP")
        If Trim(!SCHED_DAY_16) <> "" Then
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_16, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        !SCHED_DAY_17 = madrsFormatRlfRptData.Fields("SCHED_DAY17_SHIFT") & vbCrLf & _
                        madrsFormatRlfRptData.Fields("SCHED_DAY17_EXCEP")
        If Trim(!SCHED_DAY_17) <> "" Then
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_17, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        !SCHED_DAY_18 = madrsFormatRlfRptData.Fields("SCHED_DAY18_SHIFT") & vbCrLf & _
                        madrsFormatRlfRptData.Fields("SCHED_DAY18_EXCEP")
        If Trim(!SCHED_DAY_18) <> "" Then
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_18, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        !SCHED_DAY_19 = madrsFormatRlfRptData.Fields("SCHED_DAY19_SHIFT") & vbCrLf & _
                        madrsFormatRlfRptData.Fields("SCHED_DAY19_EXCEP")
        If Trim(!SCHED_DAY_19) <> "" Then
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_19, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        !SCHED_DAY_20 = madrsFormatRlfRptData.Fields("SCHED_DAY20_SHIFT") & vbCrLf & _
                        madrsFormatRlfRptData.Fields("SCHED_DAY20_EXCEP")
        If Trim(!SCHED_DAY_20) <> "" Then
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_20, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        !SCHED_DAY_21 = madrsFormatRlfRptData.Fields("SCHED_DAY21_SHIFT") & vbCrLf & _
                        madrsFormatRlfRptData.Fields("SCHED_DAY21_EXCEP")
        If Trim(!SCHED_DAY_21) <> "" Then
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_21, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        !SCHED_DAY_22 = madrsFormatRlfRptData.Fields("SCHED_DAY22_SHIFT") & vbCrLf & _
                        madrsFormatRlfRptData.Fields("SCHED_DAY22_EXCEP")
        If Trim(!SCHED_DAY_22) <> "" Then
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_22, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        !SCHED_DAY_23 = madrsFormatRlfRptData.Fields("SCHED_DAY23_SHIFT") & vbCrLf & _
                        madrsFormatRlfRptData.Fields("SCHED_DAY23_EXCEP")
        If Trim(!SCHED_DAY_23) <> "" Then
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_23, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        !SCHED_DAY_24 = madrsFormatRlfRptData.Fields("SCHED_DAY24_SHIFT") & vbCrLf & _
                        madrsFormatRlfRptData.Fields("SCHED_DAY24_EXCEP")
        If Trim(!SCHED_DAY_24) <> "" Then
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_24, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        !SCHED_DAY_25 = madrsFormatRlfRptData.Fields("SCHED_DAY25_SHIFT") & vbCrLf & _
                        madrsFormatRlfRptData.Fields("SCHED_DAY25_EXCEP")
        If Trim(!SCHED_DAY_25) <> "" Then
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_25, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        !SCHED_DAY_26 = madrsFormatRlfRptData.Fields("SCHED_DAY26_SHIFT") & vbCrLf & _
                        madrsFormatRlfRptData.Fields("SCHED_DAY26_EXCEP")
        If Trim(!SCHED_DAY_26) <> "" Then
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_26, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        !SCHED_DAY_27 = madrsFormatRlfRptData.Fields("SCHED_DAY27_SHIFT") & vbCrLf & _
                        madrsFormatRlfRptData.Fields("SCHED_DAY27_EXCEP")
        If Trim(!SCHED_DAY_27) <> "" Then
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_27, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        !SCHED_DAY_28 = madrsFormatRlfRptData.Fields("SCHED_DAY28_SHIFT") & vbCrLf & _
                        madrsFormatRlfRptData.Fields("SCHED_DAY28_EXCEP")
        If Trim(!SCHED_DAY_28) <> "" Then
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_28, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        !SCHED_DAY_29 = madrsFormatRlfRptData.Fields("SCHED_DAY29_SHIFT") & vbCrLf & _
                        madrsFormatRlfRptData.Fields("SCHED_DAY29_EXCEP")
        If Trim(!SCHED_DAY_29) <> "" Then
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_29, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        !SCHED_DAY_30 = madrsFormatRlfRptData.Fields("SCHED_DAY30_SHIFT") & vbCrLf & _
                        madrsFormatRlfRptData.Fields("SCHED_DAY30_EXCEP")
        If Trim(!SCHED_DAY_30) <> "" Then
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_30, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        !SCHED_DAY_31 = madrsFormatRlfRptData.Fields("SCHED_DAY31_SHIFT") & vbCrLf & _
                        madrsFormatRlfRptData.Fields("SCHED_DAY31_EXCEP")
        If Trim(!SCHED_DAY_31) <> "" Then
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_31, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        !SCHED_DAY_32 = madrsFormatRlfRptData.Fields("SCHED_DAY32_SHIFT") & vbCrLf & _
                        madrsFormatRlfRptData.Fields("SCHED_DAY32_EXCEP")
        If Trim(!SCHED_DAY_32) <> "" Then
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_32, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        !SCHED_DAY_33 = madrsFormatRlfRptData.Fields("SCHED_DAY33_SHIFT") & vbCrLf & _
                        madrsFormatRlfRptData.Fields("SCHED_DAY33_EXCEP")
        If Trim(!SCHED_DAY_33) <> "" Then
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_33, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        !SCHED_DAY_34 = madrsFormatRlfRptData.Fields("SCHED_DAY34_SHIFT") & vbCrLf & _
                        madrsFormatRlfRptData.Fields("SCHED_DAY34_EXCEP")
        If Trim(!SCHED_DAY_34) <> "" Then
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_34, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        !SCHED_DAY_35 = madrsFormatRlfRptData.Fields("SCHED_DAY35_SHIFT") & vbCrLf & _
                        madrsFormatRlfRptData.Fields("SCHED_DAY35_EXCEP")
        If Trim(!SCHED_DAY_35) <> "" Then
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_35, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        !SCHED_DAY_36 = madrsFormatRlfRptData.Fields("SCHED_DAY36_SHIFT") & vbCrLf & _
                        madrsFormatRlfRptData.Fields("SCHED_DAY36_EXCEP")
        If Trim(!SCHED_DAY_36) <> "" Then
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_36, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        !SCHED_DAY_37 = madrsFormatRlfRptData.Fields("SCHED_DAY37_SHIFT") & vbCrLf & _
                        madrsFormatRlfRptData.Fields("SCHED_DAY37_EXCEP")
        If Trim(!SCHED_DAY_37) <> "" Then
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_37, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        !SCHED_DAY_38 = madrsFormatRlfRptData.Fields("SCHED_DAY38_SHIFT") & vbCrLf & _
                        madrsFormatRlfRptData.Fields("SCHED_DAY38_EXCEP")
        If Trim(!SCHED_DAY_38) <> "" Then
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_38, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        !SCHED_DAY_39 = madrsFormatRlfRptData.Fields("SCHED_DAY39_SHIFT") & vbCrLf & _
                        madrsFormatRlfRptData.Fields("SCHED_DAY39_EXCEP")
        If Trim(!SCHED_DAY_39) <> "" Then
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_39, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        !SCHED_DAY_40 = madrsFormatRlfRptData.Fields("SCHED_DAY40_SHIFT") & vbCrLf & _
                        madrsFormatRlfRptData.Fields("SCHED_DAY40_EXCEP")
        If Trim(!SCHED_DAY_40) <> "" Then
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_40, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        !SCHED_DAY_41 = madrsFormatRlfRptData.Fields("SCHED_DAY41_SHIFT") & vbCrLf & _
                        madrsFormatRlfRptData.Fields("SCHED_DAY41_EXCEP")
        If Trim(!SCHED_DAY_41) <> "" Then
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_41, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        !SCHED_DAY_42 = madrsFormatRlfRptData.Fields("SCHED_DAY42_SHIFT") & vbCrLf & _
                        madrsFormatRlfRptData.Fields("SCHED_DAY42_EXCEP")
        If Trim(!SCHED_DAY_42) <> "" Then
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_42, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
    End With
    
    Exit Sub

ProcessRlfDayFirst_EH:
'Raise Error depending on the Error Number
    TrsRaiseError MODULENAME _
            , "ProcessRlfDayFirst"
            
End Sub
Sub ProcessRlfDay()
On Error GoTo ProcessRlfDay_EH

Dim lstrshift As String

    With madrsReportData
        lstrshift = Replace(!SCHED_DAY_1, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !SCHED_DAY_1 = madrsFormatRlfRptData.Fields("SCHED_DAY1_SHIFT") & vbCrLf & _
                       madrsFormatRlfRptData.Fields("SCHED_DAY1_EXCEP")
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_1, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        lstrshift = Replace(!SCHED_DAY_2, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !SCHED_DAY_2 = madrsFormatRlfRptData.Fields("SCHED_DAY2_SHIFT") & vbCrLf & _
                       madrsFormatRlfRptData.Fields("SCHED_DAY2_EXCEP")
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_2, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        lstrshift = Replace(!SCHED_DAY_3, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !SCHED_DAY_3 = madrsFormatRlfRptData.Fields("SCHED_DAY3_SHIFT") & vbCrLf & _
                       madrsFormatRlfRptData.Fields("SCHED_DAY3_EXCEP")
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_3, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        lstrshift = Replace(!sched_day_4, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !sched_day_4 = madrsFormatRlfRptData.Fields("SCHED_DAY4_SHIFT") & vbCrLf & _
                       madrsFormatRlfRptData.Fields("SCHED_DAY4_EXCEP")
            mlngPrevLineCount = CountOccurrences(!sched_day_4, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        lstrshift = Replace(!SCHED_DAY_5, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !SCHED_DAY_5 = madrsFormatRlfRptData.Fields("SCHED_DAY5_SHIFT") & vbCrLf & _
                       madrsFormatRlfRptData.Fields("SCHED_DAY5_EXCEP")
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_5, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        lstrshift = Replace(!SCHED_DAY_6, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !SCHED_DAY_6 = madrsFormatRlfRptData.Fields("SCHED_DAY6_SHIFT") & vbCrLf & _
                       madrsFormatRlfRptData.Fields("SCHED_DAY6_EXCEP")
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_6, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        lstrshift = Replace(!SCHED_DAY_7, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !SCHED_DAY_7 = madrsFormatRlfRptData.Fields("SCHED_DAY7_SHIFT") & vbCrLf & _
                       madrsFormatRlfRptData.Fields("SCHED_DAY7_EXCEP")
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_7, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        lstrshift = Replace(!SCHED_DAY_8, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !SCHED_DAY_8 = madrsFormatRlfRptData.Fields("SCHED_DAY8_SHIFT") & vbCrLf & _
                       madrsFormatRlfRptData.Fields("SCHED_DAY8_EXCEP")
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_8, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        lstrshift = Replace(!SCHED_DAY_9, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !SCHED_DAY_9 = madrsFormatRlfRptData.Fields("SCHED_DAY9_SHIFT") & vbCrLf & _
                       madrsFormatRlfRptData.Fields("SCHED_DAY9_EXCEP")
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_9, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        lstrshift = Replace(!SCHED_DAY_10, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !SCHED_DAY_10 = madrsFormatRlfRptData.Fields("SCHED_DAY10_SHIFT") & vbCrLf & _
                        madrsFormatRlfRptData.Fields("SCHED_DAY10_EXCEP")
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_10, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        lstrshift = Replace(!SCHED_DAY_11, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !SCHED_DAY_11 = madrsFormatRlfRptData.Fields("SCHED_DAY11_SHIFT") & vbCrLf & _
                        madrsFormatRlfRptData.Fields("SCHED_DAY11_EXCEP")
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_11, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        lstrshift = Replace(!SCHED_DAY_12, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !SCHED_DAY_12 = madrsFormatRlfRptData.Fields("SCHED_DAY12_SHIFT") & vbCrLf & _
                        madrsFormatRlfRptData.Fields("SCHED_DAY12_EXCEP")
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_12, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        lstrshift = Replace(!SCHED_DAY_13, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !SCHED_DAY_13 = madrsFormatRlfRptData.Fields("SCHED_DAY13_SHIFT") & vbCrLf & _
                        madrsFormatRlfRptData.Fields("SCHED_DAY13_EXCEP")
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_13, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        lstrshift = Replace(!SCHED_DAY_14, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !SCHED_DAY_14 = madrsFormatRlfRptData.Fields("SCHED_DAY14_SHIFT") & vbCrLf & _
                        madrsFormatRlfRptData.Fields("SCHED_DAY14_EXCEP")
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_14, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        lstrshift = Replace(!SCHED_DAY_15, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !SCHED_DAY_15 = madrsFormatRlfRptData.Fields("SCHED_DAY15_SHIFT") & vbCrLf & _
                        madrsFormatRlfRptData.Fields("SCHED_DAY15_EXCEP")
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_15, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        lstrshift = Replace(!SCHED_DAY_16, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !SCHED_DAY_16 = madrsFormatRlfRptData.Fields("SCHED_DAY16_SHIFT") & vbCrLf & _
                        madrsFormatRlfRptData.Fields("SCHED_DAY16_EXCEP")
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_16, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        lstrshift = Replace(!SCHED_DAY_17, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !SCHED_DAY_17 = madrsFormatRlfRptData.Fields("SCHED_DAY17_SHIFT") & vbCrLf & _
                        madrsFormatRlfRptData.Fields("SCHED_DAY17_EXCEP")
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_17, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        lstrshift = Replace(!SCHED_DAY_18, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !SCHED_DAY_18 = madrsFormatRlfRptData.Fields("SCHED_DAY18_SHIFT") & vbCrLf & _
                        madrsFormatRlfRptData.Fields("SCHED_DAY18_EXCEP")
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_18, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        lstrshift = Replace(!SCHED_DAY_19, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !SCHED_DAY_19 = madrsFormatRlfRptData.Fields("SCHED_DAY19_SHIFT") & vbCrLf & _
                        madrsFormatRlfRptData.Fields("SCHED_DAY19_EXCEP")
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_19, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        lstrshift = Replace(!SCHED_DAY_20, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !SCHED_DAY_20 = madrsFormatRlfRptData.Fields("SCHED_DAY20_SHIFT") & vbCrLf & _
                        madrsFormatRlfRptData.Fields("SCHED_DAY20_EXCEP")
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_20, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        lstrshift = Replace(!SCHED_DAY_21, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !SCHED_DAY_21 = madrsFormatRlfRptData.Fields("SCHED_DAY21_SHIFT") & vbCrLf & _
                        madrsFormatRlfRptData.Fields("SCHED_DAY21_EXCEP")
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_21, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        lstrshift = Replace(!SCHED_DAY_22, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !SCHED_DAY_22 = madrsFormatRlfRptData.Fields("SCHED_DAY22_SHIFT") & vbCrLf & _
                        madrsFormatRlfRptData.Fields("SCHED_DAY22_EXCEP")
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_22, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        lstrshift = Replace(!SCHED_DAY_23, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !SCHED_DAY_23 = madrsFormatRlfRptData.Fields("SCHED_DAY23_SHIFT") & vbCrLf & _
                        madrsFormatRlfRptData.Fields("SCHED_DAY23_EXCEP")
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_23, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        lstrshift = Replace(!SCHED_DAY_24, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !SCHED_DAY_24 = madrsFormatRlfRptData.Fields("SCHED_DAY24_SHIFT") & vbCrLf & _
                        madrsFormatRlfRptData.Fields("SCHED_DAY24_EXCEP")
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_24, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        lstrshift = Replace(!SCHED_DAY_25, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !SCHED_DAY_25 = madrsFormatRlfRptData.Fields("SCHED_DAY25_SHIFT") & vbCrLf & _
                        madrsFormatRlfRptData.Fields("SCHED_DAY25_EXCEP")
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_25, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        lstrshift = Replace(!SCHED_DAY_26, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !SCHED_DAY_26 = madrsFormatRlfRptData.Fields("SCHED_DAY26_SHIFT") & vbCrLf & _
                        madrsFormatRlfRptData.Fields("SCHED_DAY26_EXCEP")
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_26, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        lstrshift = Replace(!SCHED_DAY_27, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !SCHED_DAY_27 = madrsFormatRlfRptData.Fields("SCHED_DAY27_SHIFT") & vbCrLf & _
                        madrsFormatRlfRptData.Fields("SCHED_DAY27_EXCEP")
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_27, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        lstrshift = Replace(!SCHED_DAY_28, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !SCHED_DAY_28 = madrsFormatRlfRptData.Fields("SCHED_DAY28_SHIFT") & vbCrLf & _
                        madrsFormatRlfRptData.Fields("SCHED_DAY28_EXCEP")
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_28, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        lstrshift = Replace(!SCHED_DAY_29, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !SCHED_DAY_29 = madrsFormatRlfRptData.Fields("SCHED_DAY29_SHIFT") & vbCrLf & _
                        madrsFormatRlfRptData.Fields("SCHED_DAY29_EXCEP")
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_29, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        lstrshift = Replace(!SCHED_DAY_30, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !SCHED_DAY_30 = madrsFormatRlfRptData.Fields("SCHED_DAY30_SHIFT") & vbCrLf & _
                        madrsFormatRlfRptData.Fields("SCHED_DAY30_EXCEP")
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_30, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        lstrshift = Replace(!SCHED_DAY_31, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !SCHED_DAY_31 = madrsFormatRlfRptData.Fields("SCHED_DAY31_SHIFT") & vbCrLf & _
                        madrsFormatRlfRptData.Fields("SCHED_DAY31_EXCEP")
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_31, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        lstrshift = Replace(!SCHED_DAY_32, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !SCHED_DAY_32 = madrsFormatRlfRptData.Fields("SCHED_DAY32_SHIFT") & vbCrLf & _
                        madrsFormatRlfRptData.Fields("SCHED_DAY32_EXCEP")
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_32, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        lstrshift = Replace(!SCHED_DAY_33, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !SCHED_DAY_33 = madrsFormatRlfRptData.Fields("SCHED_DAY33_SHIFT") & vbCrLf & _
                        madrsFormatRlfRptData.Fields("SCHED_DAY33_EXCEP")
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_33, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        lstrshift = Replace(!SCHED_DAY_34, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !SCHED_DAY_34 = madrsFormatRlfRptData.Fields("SCHED_DAY34_SHIFT") & vbCrLf & _
                        madrsFormatRlfRptData.Fields("SCHED_DAY34_EXCEP")
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_34, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        lstrshift = Replace(!SCHED_DAY_35, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !SCHED_DAY_35 = madrsFormatRlfRptData.Fields("SCHED_DAY35_SHIFT") & vbCrLf & _
                        madrsFormatRlfRptData.Fields("SCHED_DAY35_EXCEP")
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_35, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        lstrshift = Replace(!SCHED_DAY_36, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !SCHED_DAY_36 = madrsFormatRlfRptData.Fields("SCHED_DAY36_SHIFT") & vbCrLf & _
                        madrsFormatRlfRptData.Fields("SCHED_DAY36_EXCEP")
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_36, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        lstrshift = Replace(!SCHED_DAY_37, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !SCHED_DAY_37 = madrsFormatRlfRptData.Fields("SCHED_DAY37_SHIFT") & vbCrLf & _
                        madrsFormatRlfRptData.Fields("SCHED_DAY37_EXCEP")
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_37, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        lstrshift = Replace(!SCHED_DAY_38, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !SCHED_DAY_38 = madrsFormatRlfRptData.Fields("SCHED_DAY38_SHIFT") & vbCrLf & _
                        madrsFormatRlfRptData.Fields("SCHED_DAY38_EXCEP")
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_38, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        lstrshift = Replace(!SCHED_DAY_39, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !SCHED_DAY_39 = madrsFormatRlfRptData.Fields("SCHED_DAY39_SHIFT") & vbCrLf & _
                        madrsFormatRlfRptData.Fields("SCHED_DAY39_EXCEP")
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_39, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        lstrshift = Replace(!SCHED_DAY_40, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !SCHED_DAY_40 = madrsFormatRlfRptData.Fields("SCHED_DAY40_SHIFT") & vbCrLf & _
                        madrsFormatRlfRptData.Fields("SCHED_DAY40_EXCEP")
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_40, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        lstrshift = Replace(!SCHED_DAY_41, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !SCHED_DAY_41 = madrsFormatRlfRptData.Fields("SCHED_DAY41_SHIFT") & vbCrLf & _
                        madrsFormatRlfRptData.Fields("SCHED_DAY41_EXCEP")
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_41, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
        lstrshift = Replace(!SCHED_DAY_42, vbCrLf, "")
        If Trim(lstrshift) = "" Then
            !SCHED_DAY_42 = madrsFormatRlfRptData.Fields("SCHED_DAY42_SHIFT") & vbCrLf & _
                        madrsFormatRlfRptData.Fields("SCHED_DAY42_EXCEP")
            mlngPrevLineCount = CountOccurrences(!SCHED_DAY_42, vbCrLf)
            If mlngPrevLineCount > mlngCurrLineCount Then
                mlngCurrLineCount = mlngPrevLineCount
            End If
        End If
    End With
    
    Exit Sub

ProcessRlfDay_EH:
'Raise Error depending on the Error Number
    TrsRaiseError MODULENAME _
            , "ProcessRlfDay"
            
End Sub

Private Function GetWorkGroupSchedule _
    (ByVal vlngSelectedWorkgroup As Long, _
    ByVal vdtStartDate As Date, _
    ByVal vdtEndDate As Date, _
    ByVal vbactiveworkgroups As Boolean) As ADODB.Recordset
On Error GoTo GetWorkGroupSchedule_EH

    Dim lobj As busWorkGroup.cbsWorkGroup
    'Dim lobj As Object

    Set lobj = CreateObject( _
            "busWorkGroup.cbsWorkGroup")
    'Set lobj = New busWorkGroup.cbsWorkGroup

    Set GetWorkGroupSchedule = lobj.GetWGSchedule(vlngSelectedWorkgroup _
                                , vdtStartDate _
                                , vdtEndDate _
                                , mblnUnmannedLines)

    Set lobj = Nothing
        
    Exit Function
GetWorkGroupSchedule_EH:
    Set lobj = Nothing

    'Raise Error depending on the Error Number
    TrsRaiseError MODULENAME _
            , "GetWorkGroupSchedule"
        
End Function
Public Function GetWorkGroupScheduleReportData() As ADODB.Recordset
On Error GoTo GetWorkGroupScheduleReportData_EH

    Dim llngPrevEmpNum As Long
    Dim llngPrevLineNum As Long

    Set madrsReportData = New ADODB.Recordset
    With madrsReportData
        .Fields.Append "WG_NUM", adInteger, 256, mconFIELDATTR
        .Fields.Append "WG_DESC", adVarChar, 256, mconFIELDATTR
        .Fields.Append "EMP_NUM", adInteger, , mconFIELDATTR
        .Fields.Append "EMP_NAME", adVarChar, 256, mconFIELDATTR
        .Fields.Append "JOB_NUM1", adChar, 6, mconFIELDATTR
        .Fields.Append "JOB_DESC1", adVarChar, 256, mconFIELDATTR
        .Fields.Append "JOB_NUM2", adChar, 6, mconFIELDATTR
        .Fields.Append "JOB_DESC2", adVarChar, 256, mconFIELDATTR
        .Fields.Append "QUALS", adVarChar, 256, mconFIELDATTR
        .Fields.Append "SKILLS", adVarChar, 256, mconFIELDATTR
        .Fields.Append "PROJ_ID", adInteger, 256, mconFIELDATTR
        .Fields.Append "PROJ_DESC", adVarChar, 256, mconFIELDATTR
        .Fields.Append "OTHER_WGS", adVarChar, 256, mconFIELDATTR
        .Fields.Append "LINE_NUM", adInteger, , mconFIELDATTR
        .Fields.Append "REQ_LINE_IND", adChar, 1, mconFIELDATTR
        .Fields.Append "DEMAND_LINE", adChar, 1, mconFIELDATTR
        .Fields.Append "REL_REPORT", adChar, 1, mconFIELDATTR
        .Fields.Append "LINE_COUNT", adInteger, , mconFIELDATTR
        .Fields.Append "SCHED_DAY_1", adVarChar, 256, mconFIELDATTR
        .Fields.Append "SCHED_DAY_2", adVarChar, 256, mconFIELDATTR
        .Fields.Append "SCHED_DAY_3", adVarChar, 256, mconFIELDATTR
        .Fields.Append "SCHED_DAY_4", adVarChar, 256, mconFIELDATTR
        .Fields.Append "SCHED_DAY_5", adVarChar, 256, mconFIELDATTR
        .Fields.Append "SCHED_DAY_6", adVarChar, 256, mconFIELDATTR
        .Fields.Append "SCHED_DAY_7", adVarChar, 256, mconFIELDATTR
        .Fields.Append "SCHED_DAY_8", adVarChar, 256, mconFIELDATTR
        .Fields.Append "SCHED_DAY_9", adVarChar, 256, mconFIELDATTR
        .Fields.Append "SCHED_DAY_10", adVarChar, 256, mconFIELDATTR
        .Fields.Append "SCHED_DAY_11", adVarChar, 256, mconFIELDATTR
        .Fields.Append "SCHED_DAY_12", adVarChar, 256, mconFIELDATTR
        .Fields.Append "SCHED_DAY_13", adVarChar, 256, mconFIELDATTR
        .Fields.Append "SCHED_DAY_14", adVarChar, 256, mconFIELDATTR
        .Fields.Append "SCHED_DAY_15", adVarChar, 256, mconFIELDATTR
        .Fields.Append "SCHED_DAY_16", adVarChar, 256, mconFIELDATTR
        .Fields.Append "SCHED_DAY_17", adVarChar, 256, mconFIELDATTR
        .Fields.Append "SCHED_DAY_18", adVarChar, 256, mconFIELDATTR
        .Fields.Append "SCHED_DAY_19", adVarChar, 256, mconFIELDATTR
        .Fields.Append "SCHED_DAY_20", adVarChar, 256, mconFIELDATTR
        .Fields.Append "SCHED_DAY_21", adVarChar, 256, mconFIELDATTR
        .Fields.Append "SCHED_DAY_22", adVarChar, 256, mconFIELDATTR
        .Fields.Append "SCHED_DAY_23", adVarChar, 256, mconFIELDATTR
        .Fields.Append "SCHED_DAY_24", adVarChar, 256, mconFIELDATTR
        .Fields.Append "SCHED_DAY_25", adVarChar, 256, mconFIELDATTR
        .Fields.Append "SCHED_DAY_26", adVarChar, 256, mconFIELDATTR
        .Fields.Append "SCHED_DAY_27", adVarChar, 256, mconFIELDATTR
        .Fields.Append "SCHED_DAY_28", adVarChar, 256, mconFIELDATTR
        .Fields.Append "SCHED_DAY_29", adVarChar, 256, mconFIELDATTR
        .Fields.Append "SCHED_DAY_30", adVarChar, 256, mconFIELDATTR
        .Fields.Append "SCHED_DAY_31", adVarChar, 256, mconFIELDATTR
        .Fields.Append "SCHED_DAY_32", adVarChar, 256, mconFIELDATTR
        .Fields.Append "SCHED_DAY_33", adVarChar, 256, mconFIELDATTR
        .Fields.Append "SCHED_DAY_34", adVarChar, 256, mconFIELDATTR
        .Fields.Append "SCHED_DAY_35", adVarChar, 256, mconFIELDATTR
        .Fields.Append "SCHED_DAY_36", adVarChar, 256, mconFIELDATTR
        .Fields.Append "SCHED_DAY_37", adVarChar, 256, mconFIELDATTR
        .Fields.Append "SCHED_DAY_38", adVarChar, 256, mconFIELDATTR
        .Fields.Append "SCHED_DAY_39", adVarChar, 256, mconFIELDATTR
        .Fields.Append "SCHED_DAY_40", adVarChar, 256, mconFIELDATTR
        .Fields.Append "SCHED_DAY_41", adVarChar, 256, mconFIELDATTR
        .Fields.Append "SCHED_DAY_42", adVarChar, 256, mconFIELDATTR

        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open
    End With
    
    mlngCurrLineCount = 0
    mlngPrevLineCount = 0
    llngPrevEmpNum = 0
    llngPrevLineNum = 0
    
    With madrsReportData
        If Not madrsFormatRptData.EOF Then
            If mblnSortbyEmployee Then
                madrsFormatRptData.Sort = "REL_REPORT, WG_NUM, EMP_NUM, LINE_NUM"
            Else
                madrsFormatRptData.Sort = "REL_REPORT, WG_NUM, LINE_NUM, EMP_NUM"
            End If
            madrsFormatRptData.MoveFirst
        End If
        While Not madrsFormatRptData.EOF
            If madrsFormatRptData.Fields("EMP_NUM") <> llngPrevEmpNum _
                Or madrsFormatRptData.Fields("EMP_NUM") = 0 _
                Or (madrsFormatRptData.Fields("EMP_NUM") = llngPrevEmpNum And _
                    madrsFormatRptData.Fields("LINE_NUM") <> llngPrevLineNum) Then
                .AddNew
                !WG_NUM = madrsFormatRptData.Fields("WG_NUM")
                !WG_DESC = madrsFormatRptData.Fields("WG_DESC")
                !EMP_NUM = madrsFormatRptData.Fields("EMP_NUM")
                !EMP_NAME = madrsFormatRptData.Fields("EMP_NAME")
                !JOB_NUM1 = madrsFormatRptData.Fields("JOB_NUM1")
                !JOB_DESC1 = madrsFormatRptData.Fields("JOB_DESC1")
                !JOB_NUM2 = madrsFormatRptData.Fields("JOB_NUM2")
                !JOB_DESC2 = madrsFormatRptData.Fields("JOB_DESC2")
                !QUALS = madrsFormatRptData.Fields("QUALS")
                !SKILLS = madrsFormatRptData.Fields("SKILLS")
                !PROJ_ID = madrsFormatRptData.Fields("PROJ_ID")
                !PROJ_DESC = madrsFormatRptData.Fields("PROJ_DESC")
                !OTHER_WGS = madrsFormatRptData.Fields("OTHER_WGS")
                If mblnSortbyEmployee Then
                    !LINE_NUM = 0
                Else
                    !LINE_NUM = madrsFormatRptData.Fields("LINE_NUM")
                End If
                !REQ_LINE_IND = madrsFormatRptData.Fields("REQ_LINE_IND")
                !DEMAND_LINE = madrsFormatRptData.Fields("DEMAND_LINE")
                !REL_REPORT = madrsFormatRptData.Fields("REL_REPORT")
                mlngPrevLineCount = 0
                mlngCurrLineCount = 0
                ProcessDayFirst
            Else
                ProcessDay
            End If
            !LINE_COUNT = mlngCurrLineCount
            llngPrevEmpNum = !EMP_NUM
            llngPrevLineNum = !LINE_NUM
            .Update
            madrsFormatRptData.MoveNext
        Wend
        If Not madrsFormatRlfRptData.EOF Then
            madrsFormatRlfRptData.Sort = "WG_NUM, EMP_NUM"
            madrsFormatRlfRptData.MoveFirst
            mlngCurrLineCount = 0
            mlngPrevLineCount = 0
            llngPrevEmpNum = 0
            llngPrevLineNum = 0
        End If
        While Not madrsFormatRlfRptData.EOF
            If madrsFormatRlfRptData.Fields("EMP_NUM") <> llngPrevEmpNum _
                Or madrsFormatRlfRptData.Fields("EMP_NUM") = 0 _
                Or (madrsFormatRlfRptData.Fields("EMP_NUM") = llngPrevEmpNum _
                And madrsFormatRlfRptData.Fields("LINE_NUM") <> llngPrevLineNum) Then
                .AddNew
                !WG_NUM = madrsFormatRlfRptData.Fields("WG_NUM")
                !WG_DESC = madrsFormatRlfRptData.Fields("WG_DESC")
                !EMP_NUM = madrsFormatRlfRptData.Fields("EMP_NUM")
                !EMP_NAME = madrsFormatRlfRptData.Fields("EMP_NAME")
                !JOB_NUM1 = madrsFormatRlfRptData.Fields("JOB_NUM1")
                !JOB_DESC1 = madrsFormatRlfRptData.Fields("JOB_DESC1")
                !JOB_NUM2 = madrsFormatRlfRptData.Fields("JOB_NUM2")
                !JOB_DESC2 = madrsFormatRlfRptData.Fields("JOB_DESC2")
                !QUALS = madrsFormatRlfRptData.Fields("QUALS")
                !SKILLS = madrsFormatRlfRptData.Fields("SKILLS")
                !PROJ_ID = madrsFormatRlfRptData.Fields("PROJ_ID")
                !PROJ_DESC = madrsFormatRlfRptData.Fields("PROJ_DESC")
                !OTHER_WGS = madrsFormatRlfRptData.Fields("OTHER_WGS")
                If mblnSortbyEmployee Then
                    !LINE_NUM = 0
                Else
                    !LINE_NUM = madrsFormatRlfRptData.Fields("LINE_NUM")
                End If
                !REQ_LINE_IND = madrsFormatRlfRptData.Fields("REQ_LINE_IND")
                !DEMAND_LINE = madrsFormatRlfRptData.Fields("DEMAND_LINE")
                !REL_REPORT = madrsFormatRlfRptData.Fields("REL_REPORT")
                mlngPrevLineCount = 0
                mlngCurrLineCount = 0
                ProcessRlfDayFirst
            Else
                ProcessRlfDay
            End If
            !LINE_COUNT = mlngCurrLineCount
            llngPrevEmpNum = !EMP_NUM
            llngPrevLineNum = !LINE_NUM
            .Update
            madrsFormatRlfRptData.MoveNext
        Wend
        .Sort = "WG_NUM, REL_REPORT, LINE_NUM, EMP_NUM"
        .MoveFirst
    End With
    
    'Call to pass the report data into the
    If madrsReportData.EOF And madrsReportData.BOF Then
        Set GetWorkGroupScheduleReportData = Nothing
    Else
        Set GetWorkGroupScheduleReportData = madrsReportData
    End If
    
    Set madrsReportData = Nothing

    Exit Function
GetWorkGroupScheduleReportData_EH:
    Set madrsReportData = Nothing

    'Raise Error depending on the Error Number
    TrsRaiseError MODULENAME _
            , "GetWorkGroupScheduleReportData" _
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
        Case "TSA"
            ValidExcep = False
        Case "TSR"
            ValidExcep = False
        Case "TOA"
            ValidExcep = False
        Case "TOR"
            ValidExcep = False
        Case "TWS"
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
    'Raise Error depending on the Error Number
    TrsRaiseError MODULENAME _
            , "OvertimeExcep"

End Function
Private Function GetJobPosition() As ADODB.Recordset
On Error GoTo GetJobPosition_EH

    Dim lobjSched As dbSchedule.cdSchedule

    Set lobjSched = CreateObject( _
                    "dbSchedule.cdSchedule")
    Set GetJobPosition = lobjSched.GetJobPosition(madrsScheduleData!WG_NUM, madrsScheduleData!CHARGE_DATE, _
    madrsScheduleData!LINE_NUM)

    Set lobjSched = Nothing

    Exit Function
GetJobPosition_EH:
    'Raise Error depending on the Error Number
    TrsRaiseError MODULENAME _
            , "GetJobPosition"

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

    TrsRaiseError MODULENAME, "GetJobDesc"

End Function


Private Sub DestroyState()

    Set madrsScheduleData = Nothing
    Set madrsFormatRptData = Nothing
    Set madrsFormatRlfRptData = Nothing
    Set madrsEmployeeInfo = Nothing
    
    Set mobjEmployee = Nothing
    Set mobjRules = Nothing

    Set rptMain = Nothing

    'Goodbye
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

