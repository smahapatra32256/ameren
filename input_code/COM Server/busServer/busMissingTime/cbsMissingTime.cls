VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "cbsMISSINGTIME"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Project:   TRIS Time to Pay
'Author:    Prem Chander
'Date:      January 2006
'Description: This module will retrieve get missing time employees, compare their schedules
'(if management) and email them.

Option Explicit
'General Declarations Section
Implements IObjectConstruct
Dim m_ConstructString As String

'#Const DEBUGGING = 0 '1 - ON;0 - OFF

Private Const MODULENAME As String = "busMissingTime.cbsMissingTime"
Private Const MTEXECUTIVE As Integer = 1
Private Const MTCONTRACT As Integer = 2

Dim mlOffset As Long, mbDoNotProcess As Boolean
Dim mbDayAfterDeadline As Boolean, mbCopySupervisor As Boolean
Dim msPayrollGroupCode As String
Dim mdtPayPeriodStart As Date, mdtPayPeriodEnd As Date
Dim m_colConstructStringProperty As Collection
Dim DEBUGGING As Integer '1 - ON;0 - OFF
Dim TRACING As Integer '1 - ON;0 - OFF
Dim TestRunNumber As Integer
Dim TestRunSuffix As String
Dim m_MailTO As String
Dim m_MailCC As String
Dim m_MailBCC As String

Private Sub Class_Initialize()
    '//C:\TRIS SOLUTION\Tasks\Logs\Run8A
    TRACING = 0
    TestRunNumber = 8
    '//TestRunSuffix = "D" 'A payroll "A"=Old, "B"=2/23 13:30, "C"=2/26 13:00, "D"=2/27 5:00
    '//TestRunSuffix = "A" 'B payroll "A"=3/1 1:30
    '//TestRunSuffix = "B" 'B payroll "B"=3/4 13:00
    TestRunSuffix = "C" 'B payroll "C"=3/5 5:00
End Sub

'in Component Services, component property, activation tab, check enable object construction
'set Constructor String to ... (Note: TO and CC, not yet implemented)
'DEBUGGING=1;TO=26649@ameren.com, 26895@ameren.com, 41802@ameren.com;CC=TRIS Development
'DEBUGGING=1;TO=26895@ameren.com;CC=TRIS Development
    
Private Sub IObjectConstruct_Construct(ByVal pCtorObj As Object)
    Dim mConstructString As IObjectConstructString
    
    Set mConstructString = pCtorObj
    
    m_ConstructString = mConstructString.ConstructString
    
    Set m_colConstructStringProperty = ParseContructString(m_ConstructString)
    
    PostConstruct
    App.LogEvent MODULENAME & " ObjectConstructString = " & m_ConstructString, vbLogEventTypeInformation
End Sub

Private Sub PostConstruct()
    If DoesPropertyExists(m_colConstructStringProperty, "DEBUGGING") Then
        If IsNumeric(m_colConstructStringProperty.Item("DEBUGGING")) Then
            DEBUGGING = CInt(m_colConstructStringProperty.Item("DEBUGGING"))
        End If
    End If
    
    If DoesPropertyExists(m_colConstructStringProperty, "TRACING") Then
        If IsNumeric(m_colConstructStringProperty.Item("TRACING")) Then
            TRACING = CInt(m_colConstructStringProperty.Item("TRACING"))
        End If
    End If
    
    If DoesPropertyExists(m_colConstructStringProperty, "TO") Then
        m_MailTO = CStr(m_colConstructStringProperty.Item("TO"))
    End If

    If DoesPropertyExists(m_colConstructStringProperty, "CC") Then
        m_MailCC = CStr(m_colConstructStringProperty.Item("CC"))
    End If
    
    If DoesPropertyExists(m_colConstructStringProperty, "BCC") Then
        m_MailBCC = CStr(m_colConstructStringProperty.Item("BCC"))
    End If
End Sub
    
Private Function DoesPropertyExists(ByVal colPropertyValue As Collection, ByVal Property As String) As Boolean
    Dim oValue As Variant

    On Error GoTo eh
    
    oValue = colPropertyValue.Item(Property)
    
    DoesPropertyExists = True
    Exit Function
eh:
    DoesPropertyExists = False
End Function

Private Function ParseContructString(ByVal ConstructString As String) As Collection
    Dim colConstructStringProperty As Collection
    Dim PropertyValueArray() As String
    Dim PropertyValue As Variant
    Dim PropertyValueArray2() As String

    Set colConstructStringProperty = New Collection
    PropertyValueArray = Split(ConstructString, ";")
    
    For Each PropertyValue In PropertyValueArray
        PropertyValueArray2 = Split(PropertyValue, "=")
        If UBound(PropertyValueArray2) >= 1 Then
            colConstructStringProperty.Add PropertyValueArray2(1), PropertyValueArray2(0)
        End If
    Next
       
    Set ParseContructString = colConstructStringProperty
End Function

'Common
Public Sub NotifyMissingTimeEmployees(ByVal vsPayrollGroup As String, Optional ByVal vbNotifySupervisor As Boolean = False)
    '//use the scheduled/run date and payroll group code to look up the pay process date
    Dim ldtPayProcessDate As Date
    Dim rs As ADODB.Recordset
    On Error GoTo eh
    
    ldtPayProcessDate = GetPayProcessedDate(vsPayrollGroup, Now)
    Set rs = NotifyMissingTimeEmployees2(vsPayrollGroup, vbNotifySupervisor, ldtPayProcessDate, Now, True)
    
    Set rs = Nothing
    Exit Sub
eh:
    Set rs = Nothing

    TrsRaiseError MODULENAME, "NotifyMissingTimeEmployees", , , , , True
End Sub


'//Test
Public Function NotifyMissingTimeEmployees3(ByVal vsPayrollGroup As String, ByVal vbNotifySupervisor As Boolean, _
                                            ByVal vdtScheduleDateTime As Date, ByVal vbSendMail As Boolean) As ADODB.Recordset
    '//use the scheduled/run date and payroll group code to look up the pay process date
    Dim ldtPayProcessDate As Date
    Dim rs As ADODB.Recordset
    On Error GoTo eh
    
    ldtPayProcessDate = GetPayProcessedDate(vsPayrollGroup, vdtScheduleDateTime)
    Set rs = NotifyMissingTimeEmployees2(vsPayrollGroup, vbNotifySupervisor, ldtPayProcessDate, vdtScheduleDateTime, vbSendMail)
    
    Set NotifyMissingTimeEmployees3 = rs
    Set rs = Nothing
    Exit Function
eh:
    Set rs = Nothing

    TrsRaiseError MODULENAME, "NotifyMissingTimeEmployees3", , , , , True
End Function

'NOTE: vdtPayProcessDate is not used by Management, it is not the same day as the scheduled datetime
Public Function NotifyMissingTimeEmployees2(ByVal vsPayrollGroup As String, ByVal vbNotifySupervisor As Boolean, _
                                            ByVal vdtPayProcessDate As Date, ByVal vdtScheduleDateTime As Date, ByVal vbSendMail As Boolean) As ADODB.Recordset
On Error GoTo eh
    Dim liExecContractIndicator As Integer, ladrs As ADODB.Recordset
    Dim lxml As MSXML2.DOMDocument
    Dim rsEMail As ADODB.Recordset
    Dim ldtDeadlineDateTime As Date
    
    mbCopySupervisor = vbNotifySupervisor
    msPayrollGroupCode = UCase(Trim(vsPayrollGroup))

    Select Case (msPayrollGroupCode)
        Case "A", "B", "C", "N"
            liExecContractIndicator = MTCONTRACT
        Case "E"
            liExecContractIndicator = MTEXECUTIVE
    End Select

    Set ladrs = GetMissingTimeInformation2(msPayrollGroupCode, vdtPayProcessDate, vdtScheduleDateTime)

    If (IsTheRSPopulated(ladrs)) Then
        With ladrs
            If (.RecordCount > 0) Then
                .MoveFirst

                Set lxml = New MSXML2.DOMDocument

                .Save lxml, adPersistXML

                lxml.Save "C:\TRIS Solution\Tasks\Logs\" & msPayrollGroupCode & "MT" & Format(Now, "MMddyyyyhhmmss") & ".xml"
                Set lxml = Nothing

                .MoveFirst
            End If
        End With
    End If

    Set rsEMail = CreateEMailRecordset()
    
    'Call Method to notify
    'NOTE: If TRACING is enabled, BCC will be set to "TRIS Development"
    ldtDeadlineDateTime = DateAdd("h", 15, DateAdd("d", -1, MidnightDate(vdtPayProcessDate)))
    SendMail ladrs, liExecContractIndicator, mbCopySupervisor, vbSendMail, rsEMail, vdtPayProcessDate, vdtScheduleDateTime, ldtDeadlineDateTime
    
    If TRACING = 1 Then
        SaveEMailRecordset rsEMail, GetTraceFolder(vsPayrollGroup, vdtPayProcessDate, vdtScheduleDateTime)
    End If

    Set ladrs = Nothing
    Set NotifyMissingTimeEmployees2 = rsEMail
    Exit Function
eh:
    Set lxml = Nothing
    Set ladrs = Nothing

    TrsRaiseError MODULENAME, "NotifyMissingTimeEmployees2", , , , , True
End Function

Private Function GetTraceFolder(ByVal vsPayrollGroup As String, ByVal vdtPayProcessDate As Date, ByVal vdtScheduleDateTime As Date) As String
    GetTraceFolder = "C:\TRIS SOLUTION\Tasks\Logs\Run" & TestRunNumber & TestRunSuffix & "\SendMail_" & _
                    Format(vdtPayProcessDate, "YYYYMMDD") & "_" & Format(vdtScheduleDateTime, "YYYYMMDD.hh.mm") & vsPayrollGroup
End Function

'TODO: Make this obsolete
Public Function GetMissingTimeInformation(ByVal vsPayrollGroupCode As String) As ADODB.Recordset
    Set GetMissingTimeInformation = GetMissingTimeInformation2(vsPayrollGroupCode, Date, Now)
End Function

Public Function GetMissingTimeInformation2(ByVal vsPayrollGroupCode As String, ByVal vdtPayProcessDate As Date, ByVal vdtScheduleDateTime As Date) As ADODB.Recordset
On Error GoTo eh

    Set GetMissingTimeInformation2 = Nothing

    Select Case (UCase(Trim(vsPayrollGroupCode)))
        Case "E"
            Set GetMissingTimeInformation2 = GetMgmtMissingTimeEmployees2(vdtScheduleDateTime)

        Case "A", "B", "C", "N"
            Set GetMissingTimeInformation2 = GetContractEmployeesMissingTime2(vsPayrollGroupCode, vdtPayProcessDate, vdtScheduleDateTime)
    End Select

    Exit Function
eh:
    TrsRaiseError MODULENAME, "GetMissingTimeInformation2", , , , , True
End Function

'Process flow for sending missing time notifications for contract employees:
'
'1. Send to Timekeepers (those that have entered or modified actual time (time collection))
'
'If the CopySupervisor = True
'   2. Send to Supervisor
'   3. Send to Timekeeper’s Supervisor
'
'4. Send Summary to Payroll

'Process flow for sending missing time notifications for executive employees:
'1. e-mail to individuals (see SEND_IND and COPY_SUPERVISOR_IND columns in dbMissingTime)
'
'4. Mail summary to Payroll

'TODO: remove </CENTER></BODY>, cannot find <CENTER> nor <BODY> ???
'cannot see (This is an automated message. Please do not reply to this message.) when html is viewed in Word
'TODO: Split into 2 functions, 1 for contract, 1 for management
'NOTE: vdtPayProcessDate is not used
Private Sub SendMail(ByVal vadrsMissingTime As ADODB.Recordset _
                    , ByVal viExecContract As Integer _
                    , ByVal vbCopySupervisor As Boolean _
                    , ByVal vbSendMail As Boolean _
                    , ByVal vadrsEMail As ADODB.Recordset _
                    , ByVal vdtPayProcessDate As Date, ByVal vdtScheduleDateTime As Date, ByVal vdtDeadlineDateTime As Date)
On Error GoTo SendMail_EH
'Components
Dim lobjCnMgr As util_DBConnectMgr.clsuDBConnectMgr
'Region
Dim lsDB2 As String, lbNothing As Boolean
Dim lbNone As Boolean
'Other
Dim llIndex As Long, llPreviousTK As Long
Dim llPreviousSupv As Long, lsHTMLHeader As String
Dim lsHTMLTblHeader As String, lsPreviousCompany As String
Dim lsHTMLBody As String
'SMTP
Dim lobjCDOMsg As CDO.Message
Dim lobjCDOConfig As CDO.Configuration
Dim ladFlds As ADODB.Fields

    If (mbDoNotProcess) Then
        Exit Sub
    End If

    Select Case (viExecContract)
        Case MTCONTRACT
            App.LogEvent "Contract Missing Time notification e-mail process started at " & Now, vbLogEventTypeInformation
        Case MTEXECUTIVE
            App.LogEvent "Executive Missing Time notification e-mail process started at " & Now, vbLogEventTypeInformation
    End Select

    'Get Region
    Set lobjCnMgr = CreateObject("util_DBConnectMgr.clsuDBConnectMgr")
    lsDB2 = lobjCnMgr.GetDataSourceName
    Set lobjCnMgr = Nothing

    If (InStr(1, lsDB2, "LUW") > 0) Then
        lsDB2 = Trim(Left(lsDB2, 7))
    Else
        lsDB2 = Trim(Left(lsDB2, 4))
    End If

    'Configure
    Set lobjCDOConfig = CreateObject("CDO.Configuration")
    Set ladFlds = lobjCDOConfig.Fields
    
    Dim sSendUsingMethod As String
    Dim sSMPTAutenticate As String
    
    sSendUsingMethod = GetEmailConfig("cdoSendUsingMethod")
    
    If UCase(sSendUsingMethod) = UCase("cdoSendUsingPort") Then
        ladFlds(cdoSendUsingMethod) = cdoSendUsingPort

        ladFlds(cdoSMTPServer) = GetEmailConfig("cdoSMTPServer") '//"mail.ameren.com"
        ladFlds(cdoSMTPServerPort) = GetEmailConfig("cdoSMTPServerPort") '//25
        sSMPTAutenticate = GetEmailConfig("cdoSMTPAuthenticate") '//cdoAnonymous
        If UCase(sSMPTAutenticate) = UCase("cdoAnonymous") Then
            ladFlds(cdoSMTPAuthenticate) = cdoAnonymous
        End If
    Else
        If UCase(sSendUsingMethod) = UCase("cdoSendUsingPickup") Then
            ladFlds(cdoSendUsingMethod) = cdoSendUsingPickup
            
            '//ladFlds(cdoSMTPServerPickupDirectory) = "\\plzntpvtrscomp1\C\TRIS SOLUTION\EMailOutbox\MissingTime"
            ladFlds(cdoSMTPServerPickupDirectory) = GetEmailConfig("cdoSMTPServerPickupDirectory") '//"C:\TRIS SOLUTION\EMailOutbox\MissingTime"
        Else
    
        End If
    End If
    
'    ladFlds(cdoSendUsingMethod) = cdoSendUsingPort
'
'    'Per information from Mr. Travis Foley
''    Select Case (Environ("COMPUTERNAME"))
''        Case "TRISQA1", "TRISNLB1", "TRISNLB2" _
''            , "TRISPROD1", "TRISPROD3", "TRISPROD4" _
''            , "TRISPROD5", "TRISPROD6", "TRISPROD7", "PLZNTPVTRIS1", "PLZNTPVTRIS2", "PLZNTPVTRIS3"
'            ladFlds(cdoSMTPServer) = "mail.ameren.com"
'            ladFlds(cdoSMTPServerPort) = 25
''        Case Else
''            ladFlds(cdoSMTPServer) = "appmta.corp.dir.ameren.com"
''            ladFlds(cdoSMTPServerPort) = 7257
''    End Select
'
'    ladFlds(cdoSMTPAuthenticate) = cdoAnonymous
    ladFlds.Update

'//TEST
'    Set lobjCDOMsg = CreateObject("CDO.Message")
'    With lobjCDOMsg
'        Set .Configuration = lobjCDOConfig
'        .Subject = "test missing time send email"
'        '//.AutoGenerateTextBody = True
'        .Fields(cdoPriority) = 1 '//Values: Urgent 1; Normal 0; Non-Urgent -1.
'        .Fields(cdoImportance) = 2 '//Values: Low 0, Normal 1, High 2
'        .To = "tvincent@ameren.com"
'        .TextBody = "this is a test"
'        '//.HTMLBody
'        .Send
'    End With

    'Notify
    If (IsTheRSPopulated(vadrsMissingTime)) Then
        lbNothing = False

        vadrsMissingTime.Filter = adFilterNone

        Select Case (viExecContract)
            Case MTCONTRACT
                vadrsMissingTime.Sort = "TK_EMP_NUM, CMPY_CD, DEPT_NUM, SUPERVISOR_NUM, EMP_NUM, CHARGE_DATE, OT_IND"
                vadrsMissingTime.MoveFirst

                Set lobjCDOMsg = CreateObject("CDO.Message")
                With lobjCDOMsg
                    Set .Configuration = lobjCDOConfig

                    .From = "HRCentralPayroll@ameren.com"
                    .Subject = "ACTION REQUIRED: Missing Time for pay period " & _
                        FormatDate(vadrsMissingTime!PAY_PER_START_DATE) & " - " & _
                        FormatDate(vadrsMissingTime!PAY_PER_END_DATE) & "..."
                    .AutoGenerateTextBody = True
                    .Fields(cdoPriority) = 1 '//Values: Urgent 1; Normal 0; Non-Urgent -1.
                    .Fields(cdoImportance) = 2 '//Values: Low 0, Normal 1, High 2
                    .Fields.Update
                    
                    .To = ""
                    .CC = ""
                    .BCC = ""
                    llPreviousTK = 0
                    llPreviousSupv = 0

                    lsHTMLTblHeader = "<TABLE ALIGN=CENTER CELLSPACING=1 CELLPADDING=0 BORDER=1>" & _
                            "<TR><TD WIDTH=235 ALIGN=CENTER BGCOLOR=#800000>" & _
                            "<FONT STYLE=TIMES NEW ROMAN COLOR=#ffffff SIZE=3>" & _
                                "Employee</FONT></TD>" & _
                            "<TD WIDTH=235 ALIGN=CENTER BGCOLOR=#800000>" & _
                            "<FONT STYLE=TIMES NEW ROMAN COLOR=#ffffff SIZE=3>" & _
                                "Department</FONT></TD>" & _
                            "<TD WIDTH=235 ALIGN=CENTER BGCOLOR=#800000>" & _
                            "<FONT STYLE=TIMES NEW ROMAN COLOR=#ffffff SIZE=3>" & _
                                "Supervisor</FONT></TD>" & _
                            "<TD WIDTH=80 ALIGN=CENTER BGCOLOR=#800000>" & _
                            "<FONT STYLE=TIMES NEW ROMAN COLOR=#ffffff SIZE=3>" & _
                                "Charge Date</FONT></TD>" & _
                            "<TD WIDTH=50 ALIGN=CENTER BGCOLOR=#800000>" & _
                            "<FONT STYLE=TIMES NEW ROMAN COLOR=#ffffff SIZE=3>" & _
                                "Scheduled OT</FONT></TD>" & _
                            "<TD WIDTH=65 ALIGN=CENTER BGCOLOR=#800000>" & _
                            "<FONT STYLE=TIMES NEW ROMAN COLOR=#ffffff SIZE=3>" & _
                                "Hours Entered</FONT></TD>" & _
                            "<TD WIDTH=65 ALIGN=CENTER BGCOLOR=#800000>" & _
                            "<FONT STYLE=TIMES NEW ROMAN COLOR=#ffffff SIZE=3>" & _
                                "Hours Scheduled</FONT></TD></TR>"

                    If (mbDayAfterDeadline) Then
                        lsHTMLHeader = "<HTML>This message is to inform you that" & _
                            " the following employees are missing time" & _
                            " for the associated charge dates." & _
                            " Time must be entered immediately." & _
                            "<P>Please contact Central Payroll to unlock locked departments to report time.</P>"
                    Else
                        If MidnightDate(vdtScheduleDateTime) = MidnightDate(vdtDeadlineDateTime) Then
                            lsHTMLHeader = "<HTML>This message is to inform you that" & _
                                " the following employees are missing time" & _
                                " for the associated charge dates." & _
                                " Time must be entered by 3:00PM today (" & FormatDate(vdtScheduleDateTime) & ")."
                        Else
                            lsHTMLHeader = "<HTML>This message is to inform you that" & _
                                " the following employees are missing time" & _
                                " for the associated charge dates." & _
                                " Time must be entered by " & FormatDateTime2(vdtDeadlineDateTime) & "."
                        End If
                    End If

                    For llIndex = 1 To vadrsMissingTime.RecordCount
                        If (llPreviousTK <> vadrsMissingTime!TK_EMP_NUM) Then
                            If (llIndex > 1) Then
                                If (.HTMLBody <> "") Then
                                    .HTMLBody = .HTMLBody & "</TABLE></P>"
                                End If
                            End If

                            If (Trim(.To) = "") Then
                                .To = CStr(vadrsMissingTime!TK_EMP_NUM) & "@ameren.com"
                            Else
                                If (InStr(1, CStr(vadrsMissingTime!TK_EMP_NUM) & "@ameren.com", .To) = 0) And _
                                    (InStr(1, CStr(vadrsMissingTime!TK_EMP_NUM) & "@ameren.com", .CC) = 0) Then
                                    .To = .To & ", " & CStr(vadrsMissingTime!TK_EMP_NUM) & "@ameren.com"
                                End If
                            End If

                            .HTMLBody = lsHTMLHeader & _
                                "<P><FONT STYLE=TIMES NEW ROMAN SIZE=3>" & _
                                "Company: " & Trim(vadrsMissingTime!CMPY_DESC_TXT) & _
                                " [" & Trim(vadrsMissingTime!CMPY_CD) & "]</FONT></TD>" & _
                                lsHTMLTblHeader
                        Else
                            If (lsPreviousCompany <> vadrsMissingTime!CMPY_CD) Then
                                If (llIndex > 1) Then
                                    If (.HTMLBody <> "") Then
                                        .HTMLBody = .HTMLBody & "</TABLE></P>"
                                    End If
                                End If

                                .HTMLBody = .HTMLBody & _
                                    "<P><FONT STYLE=TIMES NEW ROMAN SIZE=3>" & _
                                    "Company: " & Trim(vadrsMissingTime!CMPY_DESC_TXT) & _
                                    " [" & Trim(vadrsMissingTime!CMPY_CD) & "]</FONT></TD>" & _
                                    lsHTMLTblHeader
                            End If
                        End If

                        .HTMLBody = .HTMLBody & _
                            "<TR><TD ALIGN=LEFT><FONT STYLE =TIMES NEW ROMAN SIZE=2>" & _
                                Trim(vadrsMissingTime!EMP_NAME) & _
                                " [" & CStr(vadrsMissingTime!EMP_NUM) & "]" & _
                            "</FONT></TD>" & _
                            "<TD ALIGN=LEFT><FONT STYLE =TIMES NEW ROMAN SIZE=2>" & _
                                Trim(vadrsMissingTime!DEPT_NAME) & _
                                " [" & Trim(vadrsMissingTime!DEPT_NUM) & "]" & _
                            "</FONT></TD>" & _
                            "<TD ALIGN=LEFT><FONT STYLE =TIMES NEW ROMAN SIZE=2>" & _
                                Trim(vadrsMissingTime!SUPERVISOR_NAME) & _
                                " [" & CStr(vadrsMissingTime!SUPERVISOR_NUM) & "]" & _
                            "</FONT></TD>" & _
                            "<TD ALIGN=LEFT><FONT STYLE =TIMES NEW ROMAN SIZE=2>" & _
                                FormatDate(vadrsMissingTime!CHARGE_DATE) & _
                            "</FONT></TD>" & _
                            "<TD ALIGN=CENTER><FONT STYLE =TIMES NEW ROMAN SIZE=2>" & _
                                Trim(vadrsMissingTime!OT_IND) & _
                            "</FONT></TD>" & _
                            "<TD ALIGN=RIGHT><FONT STYLE =TIMES NEW ROMAN SIZE=2>" & _
                                Format(vadrsMissingTime!ENTERED_HOURS, "#0.0") & _
                            "</FONT></TD>" & _
                            "<TD ALIGN=RIGHT><FONT STYLE =TIMES NEW ROMAN SIZE=2>" & _
                                Format(vadrsMissingTime!SCHEDULED_HOURS, "#0.0") & _
                            "</FONT></TD>"

                        llPreviousTK = vadrsMissingTime!TK_EMP_NUM
                        llPreviousSupv = vadrsMissingTime!SUPERVISOR_NUM
                        lsPreviousCompany = vadrsMissingTime!CMPY_CD

                        vadrsMissingTime.MoveNext
                        
'TODO: use indicator to flush out last row on level breaks instead of "If (llIndex = vadrsMissingTime.RecordCount) Then"
'Dim blnFlushEMail As Boolean
'blnFlushEMail = False
'                        If Not (vadrsMissingTime.EOF) Then
'                            If (vadrsMissingTime!TK_EMP_NUM <> llPreviousTK) Then
'                                blnFlushEMail = True
'                            End If
'                        Else
'                            If (llIndex = vadrsMissingTime.RecordCount) Then
'                                blnFlushEMail = True
'                            End If
'                        End If
'
'                        If blnFlushEMail Then
'                        End If
                        
                        If Not (vadrsMissingTime.EOF) Then
                            If (vadrsMissingTime!TK_EMP_NUM <> llPreviousTK) Then
                                If (vbCopySupervisor) Then
                                    .HTMLBody = .HTMLBody & "</TABLE></P></CENTER></BODY>" & _
                                        "<P>Respective supervisors will be notified regarding missing time.</P>" & _
                                        "<P>If you are no longer responsible for the time entry of any of the listed employees" & _
                                        ", please inform the right person regarding missing time.</P>" & _
                                        "<P>(This is an automated message. Please do not reply to this message.)</P>"
                                Else
                                    .HTMLBody = .HTMLBody & "</TABLE></P></CENTER></BODY>" & _
                                        "<P>If you are no longer responsible for the time entry of any of the listed employees" & _
                                        ", please inform the right person regarding missing time.</P>" & _
                                        "<P>(This is an automated message. Please do not reply to this message.)</P>"
                                End If

                                Select Case (lsDB2)
                                    Case "DB2P", "DB2PLUW"
                                        .HTMLBody = .HTMLBody & "</HTML>"
                                    Case Else
                                        .HTMLBody = .HTMLBody & _
                                            "<P>THIS IS A TEST MESSAGE. PLEASE DISREGARD THIS MESSAGE.</P></HTML>"
                                End Select

                                TraceEMail lobjCDOMsg
                                AddEMailRecord vadrsEMail, lobjCDOMsg, "Timekeeper"
                                If vbSendMail Then
                                    On Error Resume Next
                                    .Send
                                    If (Err.Number <> 0) Then
                                        Err.Clear
                                    End If
                                    On Error GoTo SendMail_EH
                                End If

                                .To = ""
                                .CC = ""
                                .BCC = ""
                                .HTMLBody = ""
                            End If
                        Else
                            If (llIndex = vadrsMissingTime.RecordCount) Then
'                                On Error Resume Next
                                If (vbCopySupervisor) Then
                                    .HTMLBody = .HTMLBody & "</TABLE></P></CENTER></BODY>" & _
                                        "<P>Respective supervisors will be notified regarding missing time.</P>" & _
                                        "<P>If you are no longer responsible for the time entry of any of the listed employees" & _
                                        ", please inform the right person regarding missing time.</P>" & _
                                        "<P>(This is an automated message. Please do not reply to this message.)</P>"
                                Else
                                    .HTMLBody = .HTMLBody & "</TABLE></P></CENTER></BODY>" & _
                                        "<P>If you are no longer responsible for the time entry of any of the listed employees" & _
                                        ", please inform the right person regarding missing time.</P>" & _
                                        "<P>(This is an automated message. Please do not reply to this message.)</P>"
                                End If

                                Select Case (lsDB2)
                                    Case "DB2P", "DB2PLUW"
                                        .HTMLBody = .HTMLBody & "</HTML>"
                                    Case Else
                                        .HTMLBody = .HTMLBody & _
                                            "<P>THIS IS A TEST MESSAGE. PLEASE DISREGARD THIS MESSAGE.</P></HTML>"
                                End Select

                                TraceEMail lobjCDOMsg
                                AddEMailRecord vadrsEMail, lobjCDOMsg, "Timekeeper"
                                If vbSendMail Then
                                    On Error Resume Next
                                    .Send
                                    If (Err.Number <> 0) Then
                                        Err.Clear
                                    End If
                                    On Error GoTo SendMail_EH
                                End If

                                .To = ""
                                .CC = ""
                                .BCC = ""
                                .HTMLBody = ""
                            End If
                        End If
                        DoEvents
                    Next llIndex

                    'Send e-mails to Supervisors.
                    If (vbCopySupervisor) Then
                        'Send e-mails to employee supervisors
                        vadrsMissingTime.Sort = "SUPERVISOR_NUM, CMPY_CD, DEPT_NUM, EMP_NUM, CHARGE_DATE, OT_IND, TK_EMP_NUM"
                        vadrsMissingTime.MoveFirst

                        .To = ""
                        .CC = ""
                        .BCC = ""
                        .HTMLBody = ""

                        llPreviousSupv = 0
                        lsPreviousCompany = ""

                        If (mbDayAfterDeadline) Then
                            lsHTMLHeader = "<HTML>This message is to inform you that" & _
                                " the following employees are missing time" & _
                                " for the associated charge dates." & _
                                " Time must be entered immediately." & _
                                " Respective Timekeepers have been notified to enter time for these employees immediately."
                        Else
                            If MidnightDate(vdtScheduleDateTime) = MidnightDate(vdtDeadlineDateTime) Then
                                lsHTMLHeader = "<HTML>This message is to inform you that" & _
                                    " the following employees are missing time" & _
                                    " for the associated charge dates." & _
                                    " Time must be entered by 3:00PM today (" & FormatDate(vdtScheduleDateTime) & ")." & _
                                    " Respective Timekeepers have been notified to enter time for these employees immediately."
                            Else
                                lsHTMLHeader = "<HTML>This message is to inform you that" & _
                                    " the following employees are missing time" & _
                                    " for the associated charge dates." & _
                                    " Time must be entered by " & FormatDateTime2(vdtDeadlineDateTime) & "." & _
                                    " Respective Timekeepers have been notified to enter time for these employees immediately."
                            End If
                        End If

                        lsHTMLTblHeader = "<TABLE ALIGN=CENTER CELLSPACING=1 CELLPADDING=0 BORDER=1>" & _
                                "<TR><TD WIDTH=235 ALIGN=CENTER BGCOLOR=#800000>" & _
                                "<FONT STYLE=TIMES NEW ROMAN COLOR=#ffffff SIZE=3>" & _
                                    "Employee</FONT></TD>" & _
                                "<TD WIDTH=235 ALIGN=CENTER BGCOLOR=#800000>" & _
                                "<FONT STYLE=TIMES NEW ROMAN COLOR=#ffffff SIZE=3>" & _
                                    "Department</FONT></TD>" & _
                                "<TD WIDTH=235 ALIGN=CENTER BGCOLOR=#800000>" & _
                                "<FONT STYLE=TIMES NEW ROMAN COLOR=#ffffff SIZE=3>" & _
                                    "Timekeeper</FONT></TD>" & _
                                "<TD WIDTH=80 ALIGN=CENTER BGCOLOR=#800000>" & _
                                "<FONT STYLE=TIMES NEW ROMAN COLOR=#ffffff SIZE=3>" & _
                                    "Charge Date</FONT></TD>" & _
                                "<TD WIDTH=50 ALIGN=CENTER BGCOLOR=#800000>" & _
                                "<FONT STYLE=TIMES NEW ROMAN COLOR=#ffffff SIZE=3>" & _
                                    "Scheduled OT</FONT></TD>" & _
                                "<TD WIDTH=65 ALIGN=CENTER BGCOLOR=#800000>" & _
                                "<FONT STYLE=TIMES NEW ROMAN COLOR=#ffffff SIZE=3>" & _
                                    "Hours Entered</FONT></TD>" & _
                                "<TD WIDTH=65 ALIGN=CENTER BGCOLOR=#800000>" & _
                                "<FONT STYLE=TIMES NEW ROMAN COLOR=#ffffff SIZE=3>" & _
                                    "Hours Scheduled</FONT></TD></TR>"

                        For llIndex = 1 To vadrsMissingTime.RecordCount
                            If (llPreviousSupv <> vadrsMissingTime!SUPERVISOR_NUM) Then
                                If (llIndex > 1) Then
                                    If (.HTMLBody <> "") Then
                                        .HTMLBody = .HTMLBody & "</TABLE></P>"
                                    End If
                                End If

                                .To = CStr(vadrsMissingTime!SUPERVISOR_NUM) & "@ameren.com"
                                .CC = ""
                                .BCC = ""
                            
                                .HTMLBody = lsHTMLHeader & _
                                    "<P><FONT STYLE=TIMES NEW ROMAN SIZE=3>" & _
                                    "Company: " & Trim(vadrsMissingTime!CMPY_DESC_TXT) & _
                                    " [" & Trim(vadrsMissingTime!CMPY_CD) & "]</FONT></TD>" & _
                                    lsHTMLTblHeader
                            Else
                                If (lsPreviousCompany <> vadrsMissingTime!CMPY_CD) Then
                                    If (llIndex > 1) Then
                                        If (.HTMLBody <> "") Then
                                            .HTMLBody = .HTMLBody & "</TABLE></P>"
                                        End If
                                    End If

                                    .HTMLBody = .HTMLBody & _
                                        "<P><FONT STYLE=TIMES NEW ROMAN SIZE=3>" & _
                                        "Company: " & Trim(vadrsMissingTime!CMPY_DESC_TXT) & _
                                        " [" & Trim(vadrsMissingTime!CMPY_CD) & "]</FONT></TD>" & _
                                        lsHTMLTblHeader
                                End If
                            End If

                            .HTMLBody = .HTMLBody & _
                                "<TR><TD ALIGN=LEFT><FONT STYLE =TIMES NEW ROMAN SIZE=2>" & _
                                    Trim(vadrsMissingTime!EMP_NAME) & _
                                    " [" & CStr(vadrsMissingTime!EMP_NUM) & "]" & _
                                "</FONT></TD>" & _
                                "<TD ALIGN=LEFT><FONT STYLE =TIMES NEW ROMAN SIZE=2>" & _
                                    Trim(vadrsMissingTime!DEPT_NAME) & _
                                    " [" & Trim(vadrsMissingTime!DEPT_NUM) & "]" & _
                                "</FONT></TD>" & _
                                "<TD ALIGN=LEFT><FONT STYLE =TIMES NEW ROMAN SIZE=2>" & _
                                    Trim(vadrsMissingTime!TK_EMP_NAME) & _
                                    " [" & CStr(vadrsMissingTime!TK_EMP_NUM) & "]" & _
                                "</FONT></TD>" & _
                                "<TD ALIGN=LEFT><FONT STYLE =TIMES NEW ROMAN SIZE=2>" & _
                                    FormatDate(vadrsMissingTime!CHARGE_DATE) & _
                                "</FONT></TD>" & _
                                "<TD ALIGN=CENTER><FONT STYLE =TIMES NEW ROMAN SIZE=2>" & _
                                    Trim(vadrsMissingTime!OT_IND) & _
                                "</FONT></TD>" & _
                                "<TD ALIGN=RIGHT><FONT STYLE =TIMES NEW ROMAN SIZE=2>" & _
                                    Format(vadrsMissingTime!ENTERED_HOURS, "#0.0") & _
                                "</FONT></TD>" & _
                                "<TD ALIGN=RIGHT><FONT STYLE =TIMES NEW ROMAN SIZE=2>" & _
                                    Format(vadrsMissingTime!SCHEDULED_HOURS, "#0.0") & _
                                "</FONT></TD>"

                            llPreviousSupv = vadrsMissingTime!SUPERVISOR_NUM
                            lsPreviousCompany = vadrsMissingTime!CMPY_CD

                            vadrsMissingTime.MoveNext
                            If Not (vadrsMissingTime.EOF) Then
                                If (vadrsMissingTime!SUPERVISOR_NUM <> llPreviousSupv) Then
                                    .HTMLBody = .HTMLBody & "</TABLE></P></CENTER></BODY>" & _
                                        "<P>(This is an automated message. Please do not reply to this message.)</P>"

                                    Select Case (lsDB2)
                                        Case "DB2P", "DB2PLUW"
                                            .HTMLBody = .HTMLBody & "</HTML>"
                                        Case Else
                                            .HTMLBody = .HTMLBody & _
                                                "<P>THIS IS A TEST MESSAGE. PLEASE DISREGARD THIS MESSAGE.</P></HTML>"
                                    End Select

                                    TraceEMail lobjCDOMsg
                                    AddEMailRecord vadrsEMail, lobjCDOMsg, "Supervisor"
                                    If vbSendMail Then
                                        On Error Resume Next
                                        .Send
                                        If (Err.Number <> 0) Then
                                            Err.Clear
                                        End If
                                        On Error GoTo SendMail_EH
                                    End If

                                    .To = ""
                                    .CC = ""
                                    .BCC = ""
                                    .HTMLBody = ""
                                End If
                            Else
                                If (llIndex = vadrsMissingTime.RecordCount) Then
                                    .HTMLBody = .HTMLBody & "</TABLE></P></CENTER></BODY>" & _
                                        "<P>(This is an automated message. Please do not reply to this message.)</P>"

                                    Select Case (lsDB2)
                                        Case "DB2P", "DB2PLUW"
                                            .HTMLBody = .HTMLBody & "</HTML>"
                                        Case Else
                                            .HTMLBody = .HTMLBody & _
                                                "<P>THIS IS A TEST MESSAGE. PLEASE DISREGARD THIS MESSAGE.</P></HTML>"
                                    End Select

                                    TraceEMail lobjCDOMsg
                                    AddEMailRecord vadrsEMail, lobjCDOMsg, "Supervisor"
                                    If vbSendMail Then
                                        On Error Resume Next
                                        .Send
                                        If (Err.Number <> 0) Then
                                            Err.Clear
                                        End If
                                        On Error GoTo SendMail_EH
                                    End If

                                    .To = ""
                                    .CC = ""
                                    .BCC = ""
                                    .HTMLBody = ""
                                End If
                            End If
                            DoEvents
                        Next llIndex

                        'Send e-mails to Timekeeper supervisors
                        vadrsMissingTime.Sort = "TK_SUPV_NUM, CMPY_CD, DEPT_NUM, TK_EMP_NUM, SUPERVISOR_NUM, EMP_NUM, CHARGE_DATE, OT_IND"
                        vadrsMissingTime.MoveFirst

                        .To = ""
                        .CC = ""
                        .BCC = ""
                        .HTMLBody = ""

                        llPreviousSupv = 0
                        lsPreviousCompany = ""

                        If (mbDayAfterDeadline) Then
                            lsHTMLHeader = "<HTML>This message is to inform you that" & _
                                " the following employees are missing time" & _
                                " for the associated charge dates." & _
                                " Time must be entered immediately." & _
                                " Respective Timekeepers have been notified to enter time for these employees immediately."
                        Else
                            If MidnightDate(vdtScheduleDateTime) = MidnightDate(vdtDeadlineDateTime) Then
                                lsHTMLHeader = "<HTML>This message is to inform you that" & _
                                    " the following employees are missing time" & _
                                    " for the associated charge dates." & _
                                    " Time must be entered by 3:00PM today (" & FormatDate(vdtScheduleDateTime) & ")." & _
                                    " Respective Timekeepers have been notified to enter time for these employees immediately."
                            Else
                                lsHTMLHeader = "<HTML>This message is to inform you that" & _
                                    " the following employees are missing time" & _
                                    " for the associated charge dates." & _
                                    " Time must be entered by " & FormatDateTime2(vdtDeadlineDateTime) & "." & _
                                    " Respective Timekeepers have been notified to enter time for these employees immediately."
                            End If
                        End If

                        lsHTMLTblHeader = "<TABLE ALIGN=CENTER CELLSPACING=1 CELLPADDING=0 BORDER=1>" & _
                                "<TR><TD WIDTH=235 ALIGN=CENTER BGCOLOR=#800000>" & _
                                "<FONT STYLE=TIMES NEW ROMAN COLOR=#ffffff SIZE=3>" & _
                                    "Employee</FONT></TD>" & _
                                "<TD WIDTH=235 ALIGN=CENTER BGCOLOR=#800000>" & _
                                "<FONT STYLE=TIMES NEW ROMAN COLOR=#ffffff SIZE=3>" & _
                                    "Department</FONT></TD>" & _
                                "<TD WIDTH=235 ALIGN=CENTER BGCOLOR=#800000>" & _
                                "<FONT STYLE=TIMES NEW ROMAN COLOR=#ffffff SIZE=3>" & _
                                    "Supervisor</FONT></TD>" & _
                                "<TD WIDTH=235 ALIGN=CENTER BGCOLOR=#800000>" & _
                                "<FONT STYLE=TIMES NEW ROMAN COLOR=#ffffff SIZE=3>" & _
                                    "Timekeeper</FONT></TD>" & _
                                "<TD WIDTH=80 ALIGN=CENTER BGCOLOR=#800000>" & _
                                "<FONT STYLE=TIMES NEW ROMAN COLOR=#ffffff SIZE=3>" & _
                                    "Charge Date</FONT></TD>" & _
                                "<TD WIDTH=50 ALIGN=CENTER BGCOLOR=#800000>" & _
                                "<FONT STYLE=TIMES NEW ROMAN COLOR=#ffffff SIZE=3>" & _
                                    "Scheduled OT</FONT></TD>" & _
                                "<TD WIDTH=65 ALIGN=CENTER BGCOLOR=#800000>" & _
                                "<FONT STYLE=TIMES NEW ROMAN COLOR=#ffffff SIZE=3>" & _
                                    "Hours Entered</FONT></TD>" & _
                                "<TD WIDTH=65 ALIGN=CENTER BGCOLOR=#800000>" & _
                                "<FONT STYLE=TIMES NEW ROMAN COLOR=#ffffff SIZE=3>" & _
                                    "Hours Scheduled</FONT></TD></TR>"

                        For llIndex = 1 To vadrsMissingTime.RecordCount
                            If (llPreviousSupv <> vadrsMissingTime!TK_SUPV_NUM) Then
                                If (llIndex > 1) Then
                                    If (.HTMLBody <> "") Then
                                        .HTMLBody = .HTMLBody & "</TABLE></P>"
                                    End If
                                End If

                                .To = CStr(vadrsMissingTime!TK_SUPV_NUM) & "@ameren.com"
                                .CC = ""
                                .BCC = ""

                                .HTMLBody = lsHTMLHeader & _
                                    "<P><FONT STYLE=TIMES NEW ROMAN SIZE=3>" & _
                                    "Company: " & Trim(vadrsMissingTime!CMPY_DESC_TXT) & _
                                    " [" & Trim(vadrsMissingTime!CMPY_CD) & "]</FONT></TD>" & _
                                    lsHTMLTblHeader
                            Else
                                If (lsPreviousCompany <> vadrsMissingTime!CMPY_CD) Then
                                    If (llIndex > 1) Then
                                        If (.HTMLBody <> "") Then
                                            .HTMLBody = .HTMLBody & "</TABLE></P>"
                                        End If
                                    End If

                                    .HTMLBody = .HTMLBody & _
                                        "<P><FONT STYLE=TIMES NEW ROMAN SIZE=3>" & _
                                        "Company: " & Trim(vadrsMissingTime!CMPY_DESC_TXT) & _
                                        " [" & Trim(vadrsMissingTime!CMPY_CD) & "]</FONT></TD>" & _
                                        lsHTMLTblHeader
                                End If
                            End If

                            .HTMLBody = .HTMLBody & _
                                "<TR><TD ALIGN=LEFT><FONT STYLE =TIMES NEW ROMAN SIZE=2>" & _
                                    Trim(vadrsMissingTime!EMP_NAME) & _
                                    " [" & CStr(vadrsMissingTime!EMP_NUM) & "]" & _
                                "</FONT></TD>" & _
                                "<TD ALIGN=LEFT><FONT STYLE =TIMES NEW ROMAN SIZE=2>" & _
                                    Trim(vadrsMissingTime!DEPT_NAME) & _
                                    " [" & Trim(vadrsMissingTime!DEPT_NUM) & "]" & _
                                "</FONT></TD>" & _
                                "<TD ALIGN=LEFT><FONT STYLE =TIMES NEW ROMAN SIZE=2>" & _
                                    Trim(vadrsMissingTime!SUPERVISOR_NUM) & _
                                    " [" & CStr(vadrsMissingTime!SUPERVISOR_NAME) & "]" & _
                                "</FONT></TD>" & _
                                "<TD ALIGN=LEFT><FONT STYLE =TIMES NEW ROMAN SIZE=2>" & _
                                    Trim(vadrsMissingTime!TK_EMP_NAME) & _
                                    " [" & CStr(vadrsMissingTime!TK_EMP_NUM) & "]" & _
                                "</FONT></TD>"

                            .HTMLBody = .HTMLBody & _
                                "<TD ALIGN=LEFT><FONT STYLE =TIMES NEW ROMAN SIZE=2>" & _
                                    FormatDate(vadrsMissingTime!CHARGE_DATE) & _
                                "</FONT></TD>" & _
                                "<TD ALIGN=CENTER><FONT STYLE =TIMES NEW ROMAN SIZE=2>" & _
                                    Trim(vadrsMissingTime!OT_IND) & _
                                "</FONT></TD>" & _
                                "<TD ALIGN=RIGHT><FONT STYLE =TIMES NEW ROMAN SIZE=2>" & _
                                    Format(vadrsMissingTime!ENTERED_HOURS, "#0.0") & _
                                "</FONT></TD>" & _
                                "<TD ALIGN=RIGHT><FONT STYLE =TIMES NEW ROMAN SIZE=2>" & _
                                    Format(vadrsMissingTime!SCHEDULED_HOURS, "#0.0") & _
                                "</FONT></TD>"

                            llPreviousSupv = vadrsMissingTime!TK_SUPV_NUM
                            lsPreviousCompany = vadrsMissingTime!CMPY_CD

                            vadrsMissingTime.MoveNext
                            If Not (vadrsMissingTime.EOF) Then
                                If (vadrsMissingTime!TK_SUPV_NUM <> llPreviousSupv) Then
                                    .HTMLBody = .HTMLBody & "</TABLE></P></CENTER></BODY>" & _
                                        "<P>(This is an automated message. Please do not reply to this message.)</P>"

                                    Select Case (lsDB2)
                                        Case "DB2P", "DB2PLUW"
                                            .HTMLBody = .HTMLBody & "</HTML>"
                                        Case Else
                                            .HTMLBody = .HTMLBody & _
                                                "<P>THIS IS A TEST MESSAGE. PLEASE DISREGARD THIS MESSAGE.</P></HTML>"
                                    End Select

                                    TraceEMail lobjCDOMsg
                                    AddEMailRecord vadrsEMail, lobjCDOMsg, "TimekeeperSupervisor"
                                    If vbSendMail Then
                                        On Error Resume Next
                                        .Send
                                        If (Err.Number <> 0) Then
                                            Err.Clear
                                        End If
                                        On Error GoTo SendMail_EH
                                    End If

                                    .To = ""
                                    .CC = ""
                                    .BCC = ""
                                    .HTMLBody = ""
                                End If
                            Else
                                If (llIndex = vadrsMissingTime.RecordCount) Then
                                    .HTMLBody = .HTMLBody & "</TABLE></P></CENTER></BODY>" & _
                                        "<P>(This is an automated message. Please do not reply to this message.)</P>"

                                    Select Case (lsDB2)
                                        Case "DB2P", "DB2PLUW"
                                            .HTMLBody = .HTMLBody & "</HTML>"
                                        Case Else
                                            .HTMLBody = .HTMLBody & _
                                                "<P>THIS IS A TEST MESSAGE. PLEASE DISREGARD THIS MESSAGE.</P></HTML>"
                                    End Select

                                    TraceEMail lobjCDOMsg
                                    AddEMailRecord vadrsEMail, lobjCDOMsg, "TimekeeperSupervisor"
                                    If vbSendMail Then
                                        On Error Resume Next
                                        .Send
                                        If (Err.Number <> 0) Then
                                            Err.Clear
                                        End If
                                        On Error GoTo SendMail_EH
                                    End If

                                    .To = ""
                                    .CC = ""
                                    .BCC = ""
                                    .HTMLBody = ""
                                End If
                            End If
                            DoEvents
                        Next llIndex
                    End If
                End With

                Set lobjCDOMsg = Nothing

            Case MTEXECUTIVE
                lbNone = False

                vadrsMissingTime.Filter = "SEND_IND = 'Y'"
                If Not (vadrsMissingTime.EOF) Then
                    vadrsMissingTime.Sort = "CMPY_CD, DEPT_NUM, EMP_NUM"
                    vadrsMissingTime.MoveFirst

                    'e-mail to individuals
                    Set lobjCDOMsg = CreateObject("CDO.Message")
                    With lobjCDOMsg
                        Set .Configuration = lobjCDOConfig
                        .From = "HRExecPayroll@ameren.com"

                        For llIndex = 1 To vadrsMissingTime.RecordCount
                            .BCC = ""
                            If (vbCopySupervisor) Then
                                .To = vadrsMissingTime!NOTIFY_EMP_NUM ' & "@ameren.com"
                                If vadrsMissingTime!COPY_SUPERVISOR_IND = "Y" Then
                                    .CC = vadrsMissingTime!SUPERVISOR_NUM & "@ameren.com"
                                Else
                                    .CC = ""
                                End If
                            Else
                                .To = vadrsMissingTime!NOTIFY_EMP_NUM ' & "@ameren.com"
                                .CC = ""
                            End If

                            .Subject = "ACTION REQUIRED: Missing Time for pay period " & _
                                FormatDate(vadrsMissingTime!PAY_PER_START_DATE) & " - " & _
                                FormatDate(vadrsMissingTime!PAY_PER_END_DATE) & "..."
                                
                            .AutoGenerateTextBody = True
                            '//TODO: will this work for non-html body email ???
                            .Fields(cdoPriority) = 1 '//Values: Urgent 1; Normal 0; Non-Urgent -1.
                            .Fields(cdoImportance) = 2 '//Values: Low 0, Normal 1, High 2
                            .Fields.Update
                            Dim htmlMessage As String
                            
                            htmlMessage = "<HTML><P>This message is to inform you (" & _
                                    Trim(vadrsMissingTime!EMP_NAME) & " [" & CStr(vadrsMissingTime!EMP_NUM) & _
                                    "]) of missing time for the pay period beginning " & _
                                    FormatDate(vadrsMissingTime!PAY_PER_START_DATE) & _
                                    " and ending " & FormatDate(vadrsMissingTime!PAY_PER_END_DATE) & "." & _
                                    " You have entered a total of " & Format(vadrsMissingTime!TOT_HRS, "#0.0") & _
                                    " hours (excluding overtime and ONP time) against the pay period total of " & _
                                    Format(vadrsMissingTime!PAY_PER_TOT_HRS, "#0.0") & " hours.  " & _
                                    "To ensure that your time is appropriately charged to the correct accounting distribution, it is very important that your time is entered into TRIS.<\P>" & _
                                    vbCrLf & vbCrLf & _
                                    "<P>Please enter the remainder of your time in TRIS immediately.</P>"

                            If (vbCopySupervisor) Then
                                If vadrsMissingTime!COPY_SUPERVISOR_IND = "Y" Then
                                htmlMessage = htmlMessage & _
                                    vbCrLf & vbCrLf & _
                                    "<P>Your supervisor " & Trim(vadrsMissingTime!SUPERVISOR_NAME) & _
                                        " [" & CStr(vadrsMissingTime!SUPERVISOR_NUM) & _
                                    "] has been copied on this e-mail.</P>"
                                End If
                            End If

                            htmlMessage = htmlMessage & _
                                        vbCrLf & vbCrLf & _
                                        "<P>(This is an automated message. Please do not reply to this message.)</P>"

                            Select Case (lsDB2)
                                Case "DB2P", "DB2PLUW"
                                    '
                                Case Else
                                    htmlMessage = htmlMessage & _
                                        vbNewLine & vbNewLine & _
                                        "<P>THIS IS A TEST MESSAGE. PLEASE DISREGARD THIS MESSAGE.</P>"
                            End Select
                            htmlMessage = htmlMessage & "</HTML>"
                            .HTMLBody = htmlMessage

                            TraceEMail lobjCDOMsg
                            AddEMailRecord vadrsEMail, lobjCDOMsg, "Employee"
                            If vbSendMail Then
                                On Error Resume Next
                                .Send
                                If (Err.Number <> 0) Then
                                    Err.Clear
                                End If
                                On Error GoTo SendMail_EH
                            End If

                            vadrsMissingTime.MoveNext

                            .To = ""
                            .CC = ""
                            .BCC = ""
                            .Subject = ""
                            .TextBody = ""
                            .HTMLBody = ""
                            DoEvents
                        Next llIndex
                    End With

                    Set lobjCDOMsg = Nothing

                    vadrsMissingTime.MoveFirst
                Else
                    lbNone = True
                End If
        End Select
    Else
        lbNothing = True
    End If

    'Mail summary to Payroll
    Set lobjCDOMsg = CreateObject("CDO.Message")
    With lobjCDOMsg
        Set .Configuration = lobjCDOConfig

        .To = ""
        .CC = ""
        .BCC = ""
        .Subject = ""
        .TextBody = ""
        .HTMLBody = ""
        .Subject = "TRIS Time reminder on " & FormatDateTime(vdtScheduleDateTime)

        Select Case (viExecContract)
            Case MTCONTRACT
                .From = "HRCentralPayroll@ameren.com"
                If msPayrollGroupCode = "N" Then
                    '.To = "e26895@ameren.com"
                    'Cindy Sebacher, Lori Twillman, Stacey Joost
                    .To = "e24000@ameren.com, e21375@ameren.com, e92253@ameren.com"
                Else
                    .To = "HRCentralPayroll@ameren.com, #HRCentralPayroll@ameren.com"
                End If
                .CC = "TRISDevelopment@ameren.com"

                If (lbNothing) Then '// this means Not RecordsExists !!!
                    .Subject = .Subject & " for " & msPayrollGroupCode & _
                        " payroll, for pay period " & _
                        FormatDate(mdtPayPeriodStart) & " - " & FormatDate(mdtPayPeriodEnd) & "..."

                    .HTMLBody = "<HTML><P>List of contract employees (" & _
                        msPayrollGroupCode & " payroll) missing time for prior pay period " & _
                        FormatDate(mdtPayPeriodStart) & " - " & FormatDate(mdtPayPeriodEnd) & _
                        ", notified through e-mail.<P>" & _
                        "<P>NONE</P>" & _
                        "<P>(This is an automated message. Please do not reply to this message.)<P>"
                Else
                    vadrsMissingTime.Sort = "CMPY_CD, DEPT_NUM, SUPERVISOR_NUM, EMP_NUM, CHARGE_DATE, TK_EMP_NUM"
                    vadrsMissingTime.MoveFirst

                    .Subject = .Subject & " for " & msPayrollGroupCode & _
                        " payroll, for pay period " & _
                        FormatDate(vadrsMissingTime!PAY_PER_START_DATE) & " - " & _
                        FormatDate(vadrsMissingTime!PAY_PER_END_DATE) & "..."

                    lsHTMLTblHeader = "<TABLE ALIGN=CENTER CELLSPACING=1 CELLPADDING=0 BORDER=1>" & _
                                "<TR><TD WIDTH=235 ALIGN=CENTER BGCOLOR=#800000>" & _
                                "<FONT STYLE=TIMES NEW ROMAN COLOR=#ffffff SIZE=3>" & _
                                    "Employee</FONT></TD>" & _
                                "<TD WIDTH=235 ALIGN=CENTER BGCOLOR=#800000>" & _
                                "<FONT STYLE=TIMES NEW ROMAN COLOR=#ffffff SIZE=3>" & _
                                    "Department</FONT></TD>" & _
                                "<TD WIDTH=235 ALIGN=CENTER BGCOLOR=#800000>" & _
                                "<FONT STYLE=TIMES NEW ROMAN COLOR=#ffffff SIZE=3>" & _
                                    "Supervisor</FONT></TD>" & _
                                "<TD WIDTH=235 ALIGN=CENTER BGCOLOR=#800000>" & _
                                "<FONT STYLE=TIMES NEW ROMAN COLOR=#ffffff SIZE=3>" & _
                                    "Timekeeper</FONT></TD>" & _
                                "<TD WIDTH=80 ALIGN=CENTER BGCOLOR=#800000>" & _
                                "<FONT STYLE=TIMES NEW ROMAN COLOR=#ffffff SIZE=3>" & _
                                    "Charge Date</FONT></TD>" & _
                                "<TD WIDTH=30 ALIGN=CENTER BGCOLOR=#800000>" & _
                                "<FONT STYLE=TIMES NEW ROMAN COLOR=#ffffff SIZE=3>" & _
                                    "Sch. OT</FONT></TD>" & _
                                "<TD WIDTH=65 ALIGN=CENTER BGCOLOR=#800000>" & _
                                "<FONT STYLE=TIMES NEW ROMAN COLOR=#ffffff SIZE=3>" & _
                                    "Hours Entered</FONT></TD>" & _
                                "<TD WIDTH=65 ALIGN=CENTER BGCOLOR=#800000>" & _
                                "<FONT STYLE=TIMES NEW ROMAN COLOR=#ffffff SIZE=3>" & _
                                    "Hours Scheduled</FONT></TD></TR>"

                    If (vbCopySupervisor) Then
                        .HTMLBody = "<HTML>This message is to inform you that the following employees (" & _
                            msPayrollGroupCode & " payroll) were missing time for" & _
                            " the associated charge dates." & _
                            " Associated timekeepers and supervisors were notified of the missing time."
                    Else
                        .HTMLBody = "<HTML>This message is to inform you that the following employees (" & _
                            msPayrollGroupCode & " payroll) were missing time for" & _
                            " the associated charge dates." & _
                            " Associated timekeepers were notified of the missing time."
                    End If

                    lsPreviousCompany = ""
                    mlOffset = 0

                    For llIndex = 1 To vadrsMissingTime.RecordCount
                        If (lsPreviousCompany <> vadrsMissingTime!CMPY_CD) Then
                            If (llIndex > 1) Then
                                If (lsHTMLBody <> "") Then
                                    lsHTMLBody = lsHTMLBody & "</TABLE></P>"
                                End If
                            End If

                            mlOffset = 0

                            lsHTMLBody = lsHTMLBody & _
                                "<P><FONT STYLE=TIMES NEW ROMAN SIZE=3>" & _
                                "Company: " & Trim(vadrsMissingTime!CMPY_DESC_TXT) & _
                                " [" & Trim(vadrsMissingTime!CMPY_CD) & "]</FONT></TD>" & _
                                lsHTMLTblHeader
                        End If

                        If (mlOffset = 0) Then
                            mlOffset = Len(lsHTMLBody)
                        End If

                        FastConcatString lsHTMLBody _
                            , "<TR><TD ALIGN=LEFT><FONT STYLE =TIMES NEW ROMAN SIZE=2>" & _
                                Trim(vadrsMissingTime!EMP_NAME) & _
                                " [" & CStr(vadrsMissingTime!EMP_NUM) & "]" & _
                            "</FONT></TD>" & _
                            "<TD ALIGN=LEFT><FONT STYLE =TIMES NEW ROMAN SIZE=2>" & _
                                Trim(vadrsMissingTime!DEPT_NAME) & _
                                " [" & Trim(vadrsMissingTime!DEPT_NUM) & "]" & _
                            "</FONT></TD>" & _
                            "<TD ALIGN=LEFT><FONT STYLE =TIMES NEW ROMAN SIZE=2>" & _
                                Trim(vadrsMissingTime!SUPERVISOR_NAME) & _
                                " [" & CStr(vadrsMissingTime!SUPERVISOR_NUM) & "]" & _
                            "</FONT></TD>" & _
                            "<TD ALIGN=LEFT><FONT STYLE =TIMES NEW ROMAN SIZE=2>" & _
                                Trim(vadrsMissingTime!TK_EMP_NAME) & _
                                " [" & CStr(vadrsMissingTime!TK_EMP_NUM) & "]" & _
                            "</FONT></TD>" _
                            , mlOffset _
                            , 30000

                        If (Len(lsHTMLBody) * 4 = 64860) Then
                            .HTMLBody = .HTMLBody & lsHTMLBody
                            lsHTMLBody = ""
                            mlOffset = 0
                        End If

                        FastConcatString lsHTMLBody _
                            , "<TD ALIGN=LEFT><FONT STYLE =TIMES NEW ROMAN SIZE=2>" & _
                                FormatDate(vadrsMissingTime!CHARGE_DATE) & _
                            "</FONT></TD>" & _
                            "<TD ALIGN=CENTER><FONT STYLE =TIMES NEW ROMAN SIZE=2>" & _
                                Trim(vadrsMissingTime!OT_IND) & _
                            "</FONT></TD>" & _
                            "<TD ALIGN=RIGHT><FONT STYLE =TIMES NEW ROMAN SIZE=2>" & _
                                Format(vadrsMissingTime!ENTERED_HOURS, "#0.0") & _
                            "</FONT></TD>" & _
                            "<TD ALIGN=RIGHT><FONT STYLE =TIMES NEW ROMAN SIZE=2>" & _
                                Format(vadrsMissingTime!SCHEDULED_HOURS, "#0.0") & _
                            "</FONT></TD>" _
                            , mlOffset _
                            , 25000

                        lsPreviousCompany = vadrsMissingTime!CMPY_CD

                        vadrsMissingTime.MoveNext
                        DoEvents
                    Next llIndex

                    If (lsHTMLBody <> "") Then
                        .HTMLBody = .HTMLBody & lsHTMLBody
                    End If

                    vadrsMissingTime.MoveFirst

                    .HTMLBody = .HTMLBody & "</TABLE></P></CENTER></BODY>" & _
                        "<P>(This is an automated message. Please do not reply to this message.)</P>"
                End If

            Case MTEXECUTIVE
                .From = "HRExecPayroll@ameren.com"
                .To = "#HRExecPayroll@ameren.com, HRExecPayroll@ameren.com"
                .CC = "TRISDevelopment@ameren.com"
                .BCC = ""
                .AutoGenerateTextBody = True

                If (lbNothing) Then
                    .Subject = .Subject & " for prior pay period..."

                    .HTMLBody = "<HTML><P>List of employees missing time for prior pay period " & _
                        ", notified through e-mail.</P>" & _
                        "<P>NONE</P>" & _
                        "<P>(This is an automated message. Please do not reply to this message.)</P>"
                Else
                    vadrsMissingTime.Filter = adFilterNone
                    vadrsMissingTime.MoveFirst
                    If (lbNone) Then
                        .Subject = .Subject & " for pay period " & _
                                    FormatDate(vadrsMissingTime!PAY_PER_START_DATE) & " - " & _
                                    FormatDate(vadrsMissingTime!PAY_PER_END_DATE) & "..."

                        .HTMLBody = "<HTML><P>List of employees missing time for pay period " & _
                            FormatDate(vadrsMissingTime!PAY_PER_START_DATE) & " - " & _
                            FormatDate(vadrsMissingTime!PAY_PER_END_DATE) & _
                            ", notified through e-mail.</P>" & _
                            "<P>NONE</P>" & _
                            "<P>(This is an automated message. Please do not reply to this message.)</P>"
                    Else
                        vadrsMissingTime.Filter = "SEND_IND = 'Y'"
                        If Not (vadrsMissingTime.EOF) Then
                            vadrsMissingTime.Sort = "CMPY_CD, DEPT_NUM, EMP_NUM"
                            vadrsMissingTime.MoveFirst

                            .Subject = .Subject & " for pay period " & _
                                    FormatDate(vadrsMissingTime!PAY_PER_START_DATE) & " - " & _
                                    FormatDate(vadrsMissingTime!PAY_PER_END_DATE) & "..."

                            lsHTMLHeader = "<HTML>List of employees (by Company) missing time for pay period " & _
                                        FormatDate(vadrsMissingTime!PAY_PER_START_DATE) & " - " & _
                                        FormatDate(vadrsMissingTime!PAY_PER_END_DATE) & _
                                        ", notified through e-mail." & _
                                        "<BR>Number of Employees notified = " & _
                                        CStr(vadrsMissingTime.RecordCount) & "</BR>"
                            .HTMLBody = lsHTMLHeader

                            lsHTMLTblHeader = "<TABLE ALIGN=CENTER CELLSPACING=1 CELLPADDING=0 BORDER=1>" & _
                                        "<TR><TD WIDTH=235 ALIGN=CENTER BGCOLOR=#800000>" & _
                                        "<FONT STYLE=TIMES NEW ROMAN COLOR=#ffffff SIZE=3>" & _
                                            "Employee</FONT></TD>" & _
                                        "<TD WIDTH=235 ALIGN=CENTER BGCOLOR=#800000>" & _
                                        "<FONT STYLE=TIMES NEW ROMAN COLOR=#ffffff SIZE=3>" & _
                                            "Department</FONT></TD>" & _
                                        "<TD WIDTH=235 ALIGN=CENTER BGCOLOR=#800000>" & _
                                        "<FONT STYLE=TIMES NEW ROMAN COLOR=#ffffff SIZE=3>" & _
                                            "Supervisor</FONT></TD>" & _
                                        "<TD WIDTH=65 ALIGN=CENTER BGCOLOR=#800000>" & _
                                        "<FONT STYLE=TIMES NEW ROMAN COLOR=#ffffff SIZE=3>" & _
                                            "Hours Entered</FONT></TD>" & _
                                        "<TD WIDTH=65 ALIGN=CENTER BGCOLOR=#800000>" & _
                                        "<FONT STYLE=TIMES NEW ROMAN COLOR=#ffffff SIZE=3>" & _
                                            "Pay Per. Hours</FONT></TD></TR>"

                            lsPreviousCompany = ""
                            vadrsMissingTime.MoveFirst

                            For llIndex = 1 To vadrsMissingTime.RecordCount
                                If (lsPreviousCompany <> vadrsMissingTime!CMPY_CD) Then
                                    If (llIndex > 1) Then
                                        .HTMLBody = .HTMLBody & "</TABLE></P>"
                                    End If

                                    .HTMLBody = .HTMLBody & _
                                                "<P><FONT STYLE=TIMES NEW ROMAN SIZE=3>" & _
                                                "Company: " & Trim(vadrsMissingTime!CMPY_DESC_TXT) & _
                                                " [" & Trim(vadrsMissingTime!CMPY_CD) & "]</FONT></TD>" & _
                                                lsHTMLTblHeader
                                End If

                                .HTMLBody = .HTMLBody & _
                                    "<TR><TD ALIGN=LEFT><FONT STYLE =TIMES NEW ROMAN SIZE=2>" & _
                                        Trim(vadrsMissingTime!EMP_NAME) & _
                                        " [" & CStr(vadrsMissingTime!EMP_NUM) & "]" & _
                                    "</FONT></TD>" & _
                                    "<TD ALIGN=LEFT><FONT STYLE =TIMES NEW ROMAN SIZE=2>" & _
                                        Trim(vadrsMissingTime!DEPT_NAME) & _
                                        " [" & Trim(vadrsMissingTime!DEPT_NUM) & "]" & _
                                    "</FONT></TD>" & _
                                    "<TD ALIGN=LEFT><FONT STYLE =TIMES NEW ROMAN SIZE=2>" & _
                                        Trim(vadrsMissingTime!SUPERVISOR_NAME) & _
                                        " [" & CStr(vadrsMissingTime!SUPERVISOR_NUM) & "]" & _
                                    "</FONT></TD>" & _
                                    "<TD ALIGN=RIGHT><FONT STYLE =TIMES NEW ROMAN SIZE=2>" & _
                                        Format(vadrsMissingTime!TOT_HRS, "#0.0") & _
                                    "</FONT></TD>" & _
                                    "<TD ALIGN=RIGHT><FONT STYLE =TIMES NEW ROMAN SIZE=2>" & _
                                        Format(vadrsMissingTime!PAY_PER_TOT_HRS, "#0.0") & _
                                    "</FONT></TD>"

                                lsPreviousCompany = vadrsMissingTime!CMPY_CD

                                vadrsMissingTime.MoveNext
                                DoEvents
                            Next llIndex

                            vadrsMissingTime.MoveFirst

                            If (vbCopySupervisor) Then
                                .HTMLBody = .HTMLBody & "</TABLE></P></CENTER></BODY>" & _
                                    "<P>Respective Supervisors have been notified regarding Missing Time Employees.</P>" & _
                                    "<P>(This is an automated message. Please do not reply to this message.)</P>"
                            Else
                                .HTMLBody = .HTMLBody & "</TABLE></P></CENTER></BODY>" & _
                                    "<P>(This is an automated message. Please do not reply to this message.)</P>"
                            End If
                        End If
                    End If
                End If
        End Select

        If (.HTMLBody <> "") Then
            Select Case (lsDB2)
                Case "DB2P", "DB2PLUW"
                    .HTMLBody = .HTMLBody & "</HTML>"
                Case Else
                    .HTMLBody = .HTMLBody & "<P>THIS IS A TEST MESSAGE. PLEASE DISREGARD THIS MESSAGE.</P></HTML>"
            End Select
        Else
            Select Case (lsDB2)
                Case "DB2P", "DB2PLUW"
                    '
                Case Else
                    .TextBody = .TextBody & vbNewLine & vbNewLine & "THIS IS A TEST MESSAGE. PLEASE DISREGARD THIS MESSAGE."
            End Select
        End If

        TraceEMail lobjCDOMsg
        AddEMailRecord vadrsEMail, lobjCDOMsg, "SummaryToPayroll"
        If vbSendMail Then
            On Error Resume Next
            .Send
            If (Err.Number <> 0) Then
                Err.Clear
            End If
            On Error GoTo SendMail_EH
        End If

    End With

    Set lobjCDOMsg = Nothing

    Set ladFlds = Nothing
    Set lobjCDOConfig = Nothing

    Select Case (viExecContract)
        Case MTCONTRACT
            App.LogEvent "Contract Missing Time notification e-mail process ended at " & Now, vbLogEventTypeInformation
        Case MTEXECUTIVE
            App.LogEvent "Executive Missing Time notification e-mail process ended at " & Now, vbLogEventTypeInformation
    End Select

    Exit Sub
SendMail_EH:
    Set lobjCnMgr = Nothing
    Set lobjCDOMsg = Nothing
    Set ladFlds = Nothing
    Set lobjCDOConfig = Nothing

    TrsRaiseError MODULENAME, "SendMail", , , , , True
End Sub

Private Function GetEmailConfig(ByVal strValueName As String) As String
    Dim strValue As String
    Dim blnReturn As Boolean
    Dim strFullSubKey As String
    Dim ral As REGTool5.Registry
    
On Error GoTo eh
    Set ral = New REGTool5.Registry
       
    strFullSubKey = "Software\Ameren\TRIS\Email\"

    blnReturn = ral.GetKeyValue(HKEY_LOCAL_MACHINE, strFullSubKey, strValueName, strValue)
    GetEmailConfig = strValue
    Set ral = Nothing
    Exit Function
eh:
    Set ral = Nothing
    TrsRaiseError MODULENAME, "GetEmailConfig", , True, DescArgsFrom(strFullSubKey, strValueName)
End Function

Private Function FormatDate(ByVal vdtDate As Date) As String
    FormatDate = Format(vdtDate, "M/D/YYYY")
End Function

Private Function FormatDateTime(ByVal vdtDateTime As Date) As String
    FormatDateTime = Format(vdtDateTime, "M/D/YYYY H:mm AM/PM")
End Function

'//3:00 PM on Monday, February 26
Private Function FormatDateTime2(ByVal vdtDateTime As Date) As String
    FormatDateTime2 = StringFormat("{0} on {1}, {2} {3}", FormatTime2(vdtDateTime), WeekdayName(Weekday(vdtDateTime)), MonthName(Month(vdtDateTime)), Day(vdtDateTime))
End Function

Private Function FormatTime2(ByVal vdtDateTime As Date) As String
    FormatTime2 = Format(vdtDateTime, "H:mm AM/PM")
End Function

'Private Function DayNameOfDate2(aDate As Date) As String
'' RED116
'    Dim l_intWeekday As Integer
'
'    l_intWeekday = Weekday(aDate)
'    Select Case l_intWeekday
'        Case 1
'            DayNameOfDate = "Sunday"
'        Case 2
'            DayNameOfDate = "Monday"
'        Case 3
'            DayNameOfDate = "Tuesday"
'        Case 4
'            DayNameOfDate = "Wednesday"
'        Case 5
'            DayNameOfDate = "Thursday"
'        Case 6
'            DayNameOfDate = "Friday"
'        Case 7
'            DayNameOfDate = "Saturday"
'    End Select
'End Function

Private Function IsTheRSPopulated(ByVal vadrs As ADODB.Recordset) As Boolean
On Error GoTo IsTheRSPopulated_EH

    IsTheRSPopulated = False

    If Not (vadrs Is Nothing) Then
        If Not (vadrs.BOF And vadrs.EOF) Then
            IsTheRSPopulated = True
        End If
    End If

    Exit Function
IsTheRSPopulated_EH:
    TrsRaiseError MODULENAME, "IsTheRSPopulated" _
        , , , , , True

End Function

'run times PayGroups (A,B,C,N) N does not Copy Supervisor
'weeks with holidays will shift the schedule
'                              | PayProcessDate
'                              |
'                             \ /
' Sched                        V
' Task   Sunday   Monday   Tuesday   Wednesday
' 5:00            Run      Run & Copy
'10:00            Run
'13:00            Run & Copy

'Contract Missing Time

'TODO: Make this obsolete
Public Function GetContractEmployeesMissingTime(ByVal vsPayrollGroupCode As String) As ADODB.Recordset
    Set GetContractEmployeesMissingTime = GetContractEmployeesMissingTime2(vsPayrollGroupCode, Date, Now)
End Function

Public Function GetContractEmployeesMissingTime2(ByVal vsPayrollGroupCode As String, ByVal vdtPayProcessDate As Date, ByVal vdtScheduleDateTime As Date) As ADODB.Recordset
On Error GoTo eh
    'Components
    Dim lobjdb As dbMissingTime.cdMissingTime
    Dim lobjdbTrans As dbMissingTime.cdMissingTimeTrans
    
    'Recordsets
    Dim ladrsDeptPP As ADODB.Recordset, ladrsBriefSchedule As ADODB.Recordset

    App.LogEvent vsPayrollGroupCode & " GetContractEmployeesMissingTime2 started at " & Now, vbLogEventTypeInformation
    
    'Initialize
    Set GetContractEmployeesMissingTime2 = Nothing

    mbDayAfterDeadline = False
    mbDoNotProcess = False

    App.LogEvent vsPayrollGroupCode & " GetDepartmentsAndPayPeriod2 started at " & Now, vbLogEventTypeInformation

    'Get Contract Pay Period and Departments for Yesterday
    Set ladrsDeptPP = GetDepartmentsAndPayPeriod2(vsPayrollGroupCode, vdtScheduleDateTime) '//actually the scheduled or run datetime

    If TRACING = 1 Then
        SaveRecordset ladrsDeptPP, GetTraceFolder(vsPayrollGroupCode, vdtPayProcessDate, vdtScheduleDateTime), "GetDepartmentsAndPayPeriod2.xml"
    End If

    App.LogEvent vsPayrollGroupCode & " GetDepartmentsAndPayPeriod2 ended at " & Now, vbLogEventTypeInformation

    'Get Schedules for Employees in Contract Departments for Pay Period
    If (IsTheRSPopulated(ladrsDeptPP)) Then
        ladrsDeptPP.MoveFirst

        '//used for the summary sent to payroll, only when there is no one to notify
        'Set module level vars.
        mdtPayPeriodStart = ladrsDeptPP!PAY_PER_START_DATE
        mdtPayPeriodEnd = ladrsDeptPP!PAY_PER_END_DATE

        '//DetermineContractSchedule vsPayrollGroupCode, ladrsDeptPP!PAY_PROCESS_DATE, vdtScheduleDateTime
        DetermineContractSchedule vsPayrollGroupCode, vdtPayProcessDate, vdtScheduleDateTime
    Else
        mbDoNotProcess = True
    End If

    'Check if further processing is allowed
    If (mbDoNotProcess) Then
        Set ladrsDeptPP = Nothing
        Set ladrsBriefSchedule = Nothing

        App.LogEvent vsPayrollGroupCode & " GetContractEmployeesMissingTime2 (no department pay period information matched run criteria) at " & Now, vbLogEventTypeInformation

        'Exit "Gracefully"
        Exit Function
    End If
'''''    Else
'''''        App.LogEvent vsPayrollGroupCode & " Get and prepare schedule started at " & Now, vbLogEventTypeInformation
'''''
'''''        'Prepare Brief Schedule (for ST and OT)
'''''        Set ladrsBriefSchedule = PrepareBriefScheduleRS( _
'''''                                GetDepartmentSchedule(ladrsDeptPP))
'''''        Set ladrsDeptPP = Nothing
'''''
'''''        App.LogEvent vsPayrollGroupCode & " Get and prepare schedule ended at " & Now, vbLogEventTypeInformation
'''''    End If
'''''
'''''    'Destroy just in case
'''''    Set ladrsDeptPP = Nothing
'''''
'''''    App.LogEvent vsPayrollGroupCode & " DeleteEmployeeBriefSchedule & InsertEmployeeBriefSchedule started at " & Now, vbLogEventTypeInformation
'''''
'''''    'Call db to Delete and Insert Brief Schedule in TTR_EMPL_SCH
'''''    Set lobjdbTrans = CreateObject("dbMissingTime.cdMissingTimeTrans")
'''''    lobjdbTrans.DeleteEmployeeBriefSchedule
'''''    If (IsTheRSPopulated(ladrsBriefSchedule)) Then
'''''        lobjdbTrans.InsertEmployeeBriefSchedule ladrsBriefSchedule
'''''    End If
'''''    Set lobjdbTrans = Nothing
'''''    Set ladrsBriefSchedule = Nothing
'''''
'''''    App.LogEvent vsPayrollGroupCode & " DeleteEmployeeBriefSchedule & InsertEmployeeBriefSchedule ended at " & Now, vbLogEventTypeInformation

    App.LogEvent vsPayrollGroupCode & " dbGetContractMissingTimeEmployees2 started at " & Now, vbLogEventTypeInformation

    'Call db to Get Missing Time
    Set lobjdb = CreateObject("dbMissingTime.cdMissingTime")
    
    Set GetContractEmployeesMissingTime2 = lobjdb.GetContractMissingTimeEmployees2(vsPayrollGroupCode, vdtPayProcessDate)

    Set lobjdb = Nothing

    If TRACING = 1 Then
        SaveRecordset GetContractEmployeesMissingTime2, GetTraceFolder(vsPayrollGroupCode, vdtPayProcessDate, vdtScheduleDateTime), "GetContractEmployeesMissingTime2.xml"
    End If
    
    App.LogEvent vsPayrollGroupCode & " dbGetContractMissingTimeEmployees2 ended at " & Now, vbLogEventTypeInformation
    
    '//HACK: Would sometimes get a deadlock here after calling GetContractMissingTimeEmployees2 (now contains common table expressions holding locks)
    '//Now, call this db method from the vbs, trying to get the locks from the get to be released
'    App.LogEvent vsPayrollGroupCode & " DeleteEmployeeBriefSchedule started at " & Now, vbLogEventTypeInformation
'
'    'Call db to Delete records from TTR_EMPL_SCH
'    Set lobjdbTrans = CreateObject("dbMissingTime.cdMissingTimeTrans")
'    lobjdbTrans.DeleteEmployeeBriefSchedule
'    Set lobjdbTrans = Nothing
'
'    App.LogEvent vsPayrollGroupCode & " DeleteEmployeeBriefSchedule ended at " & Now, vbLogEventTypeInformation

    App.LogEvent vsPayrollGroupCode & " GetContractEmployeesMissingTime2 ended at " & Now, vbLogEventTypeInformation

    Exit Function
eh:
    Set lobjdb = Nothing
    Set lobjdbTrans = Nothing

    Set ladrsDeptPP = Nothing
    Set ladrsBriefSchedule = Nothing

    TrsRaiseError MODULENAME, "GetContractEmployeesMissingTime2", , , , , True
End Function

'//TODO: Comment out before checking in
'Public Function Test()
'    Dim TestPassed As Boolean
'
'    TestPassed = False
'    DetermineContractSchedule "A", "1/20/2010", "1/18/2010 5:00AM"
'    If mbDoNotProcess And Not mbCopySupervisor And Not mbDayAfterDeadline Then
'        TestPassed = True
'    End If
'
'    TestPassed = False
'    DetermineContractSchedule "A", "1/20/2010", "1/19/2010 5:00AM"
'    If Not mbDoNotProcess And Not mbCopySupervisor And Not mbDayAfterDeadline Then
'        TestPassed = True
'    End If
'
'    TestPassed = False
'    DetermineContractSchedule "A", "1/20/2010", "1/19/2010 1:00PM"
'    If Not mbDoNotProcess And mbCopySupervisor And Not mbDayAfterDeadline Then
'        TestPassed = True
'    End If
'
'    TestPassed = False
'    DetermineContractSchedule "A", "1/20/2010", "1/20/2010 5:00AM"
'    If Not mbDoNotProcess And mbCopySupervisor And mbDayAfterDeadline Then
'        TestPassed = True
'    End If
'    Dim v As Variant
'    Dim i As Integer, j As Integer
'    Dim s As String
'
'    v = GetRegExSubMatches("""HRCentralPayroll@ameren.com"" <HRCentralPayroll@ameren.com>, ""#HRCentralPayroll@ameren.com"" <#HRCentralPayroll@ameren.com>", _
'                            """([a-zA-Z@#.]*)""\s*\<[a-zA-Z@#.]*\>,?\s*")
'    For i = LBound(v) To UBound(v)
'        For j = LBound(v(i)) To UBound(v(i))
'            s = s & v(i)(j)
'        Next j
'        If i < UBound(v) Then
'            s = s & ", "
'        End If
'    Next i
'End Function

'//for testing new schedule of a Friday 1:30pm run (2/5/2024)
Public Function DetermineContractSchedule2(ByVal vsPayrollGroupCode As String, ByVal vdtDeptPayProcessDate As Date, ByVal vdtScheduleDateTime As Date, _
                                            ByRef rbProcess As Boolean, ByRef rbCopySupervisor As Boolean, ByRef rbDayAfterDeadline As Boolean)
On Error GoTo eh

    DetermineContractSchedule vsPayrollGroupCode, vdtDeptPayProcessDate, vdtScheduleDateTime
    
    rbProcess = Not mbDoNotProcess
    rbCopySupervisor = mbCopySupervisor
    rbDayAfterDeadline = mbDayAfterDeadline
    
    Exit Function
eh:
    TrsRaiseError MODULENAME, "DetermineContractSchedule2", , , , , True
End Function

Private Function DetermineContractSchedule(ByVal vsPayrollGroupCode As String, ByVal vdtDeptPayProcessDate As Date, ByVal vdtScheduleDateTime As Date)
On Error GoTo eh
    mbDoNotProcess = True '//gets set to false (default) in GetContractEmployeesMissingTime2, which doesn't make sense
    mbCopySupervisor = False '//this gets initialized in NotifyMissingTimeEmployees2, but only gets set (non default) for Mgmt
    mbDayAfterDeadline = False
    
    'Check for run criteria
    If vdtDeptPayProcessDate = MidnightDate(vdtScheduleDateTime) Then
        mbDayAfterDeadline = True

        Select Case (Hour(vdtScheduleDateTime))
            Case 5, 6, 7
                mbDoNotProcess = False
                If vsPayrollGroupCode <> "N" Then
                    mbCopySupervisor = True
                End If
            Case Else

        End Select
    'the day before
    ElseIf DateAdd("d", -1, vdtDeptPayProcessDate) = MidnightDate(vdtScheduleDateTime) Then
        Select Case (Hour(vdtScheduleDateTime))
            Case 5, 6, 7, 10, 11, 12
                mbDoNotProcess = False
                mbCopySupervisor = False

            Case 13, 14, 15
                mbDoNotProcess = False
                If vsPayrollGroupCode <> "N" Then
                    mbCopySupervisor = True
                End If
            Case Else
                mbDoNotProcess = True
        End Select
    'Friday before 1:30pm-ish
    Else
        Dim daysBeforePayroll As Integer
        daysBeforePayroll = DateDiff("d", MidnightDate(vdtScheduleDateTime), vdtDeptPayProcessDate)
        If daysBeforePayroll >= 0 And daysBeforePayroll < 7 Then '//4 or 5 days before payroll runs
            If DayNameOfDate(vdtScheduleDateTime) = "FRI" Then
                Select Case (Hour(vdtScheduleDateTime))
                    Case 13, 14, 15
                        mbDoNotProcess = False
                        mbCopySupervisor = False
                End Select
            End If
        End If
    End If
    Exit Function
eh:
    TrsRaiseError MODULENAME, "DetermineContractSchedule", , , , , True
End Function

'//created for testing db.GetContractMissingTimeEmployees2
Public Sub BuildEmployeeBriefSchedule(ByVal vsPayrollGroupCode As String, ByVal vdtPayProcessDate As Date)
On Error GoTo eh
    'Components
    Dim lobjdb As dbMissingTime.cdMissingTime
    Dim lobjdbTrans As dbMissingTime.cdMissingTimeTrans
    
    'Recordsets
    Dim ladrsDeptPP As ADODB.Recordset, ladrsBriefSchedule As ADODB.Recordset
   
    'Get Contract Pay Period and Departments for Yesterday
    Set ladrsDeptPP = GetDepartmentsAndPayPeriod2(vsPayrollGroupCode, vdtPayProcessDate)

    'Prepare Brief Schedule (for ST and OT)
    Set ladrsBriefSchedule = PrepareBriefScheduleRS( _
                            GetDepartmentSchedule(ladrsDeptPP))
    'Destroy just in case
    Set ladrsDeptPP = Nothing

    'Call db to Delete and Insert Brief Schedule in TTR_EMPL_SCH
    Set lobjdbTrans = CreateObject("dbMissingTime.cdMissingTimeTrans")
    lobjdbTrans.DeleteEmployeeBriefSchedule
    If (IsTheRSPopulated(ladrsBriefSchedule)) Then
        lobjdbTrans.InsertEmployeeBriefSchedule ladrsBriefSchedule
    End If
    Set lobjdbTrans = Nothing
    Set ladrsBriefSchedule = Nothing

'    'Call db to Get Missing Time
'    Set lobjdb = CreateObject("dbMissingTime.cdMissingTime")
'    Set GetContractEmployeesMissingTime2 = lobjdb.GetContractMissingTimeEmployees2(vsPayrollGroupCode, vdtPayProcessDate)
'    Set lobjdb = Nothing

    'Call db to Delete records from TTR_EMPL_SCH
'    Set lobjdbTrans = CreateObject("dbMissingTime.cdMissingTimeTrans")
'    lobjdbTrans.DeleteEmployeeBriefSchedule
'    Set lobjdbTrans = Nothing

    Exit Sub
eh:
    Set lobjdb = Nothing
    Set lobjdbTrans = Nothing

    Set ladrsDeptPP = Nothing
    Set ladrsBriefSchedule = Nothing

    TrsRaiseError MODULENAME, "BuildEmployeeBriefSchedule", , , , , True
End Sub

'TODO: Make this obsolete
Public Function GetDepartmentsAndPayPeriod(ByVal vsPayrollGroupCode As String) As ADODB.Recordset
    Set GetDepartmentsAndPayPeriod = GetDepartmentsAndPayPeriod2(vsPayrollGroupCode, Date)
    'TrsRaiseError MODULENAME, "GetDepartmentsAndPayPeriod", trErrObsoleteMethod, , , , True
End Function

Public Function GetDepartmentsAndPayPeriod2(ByVal vsPayrollGroupCode As String, ByVal vdtPayProcessDate As Date) As ADODB.Recordset
On Error GoTo eh
'Components
Dim lobjdb As dbMissingTime.cdMissingTime

    Set GetDepartmentsAndPayPeriod2 = Nothing

    Set lobjdb = CreateObject("dbMissingTime.cdMissingTime")
    Set GetDepartmentsAndPayPeriod2 = lobjdb.GetDepartmentsAndPayPeriod2(vsPayrollGroupCode, vdtPayProcessDate)
    Set lobjdb = Nothing

    Exit Function
eh:
    Set lobjdb = Nothing

    TrsRaiseError MODULENAME, "GetDepartmentsAndPayPeriod2", , , , , True
End Function

Public Function GetPayProcessedDate(ByVal vsPayrollGroupCode As String, ByVal vdtScheduleDateTime As Date) As Date
On Error GoTo eh
'Components
    Dim lobjdb As dbMissingTime.cdMissingTime

    GetPayProcessedDate = Date

    Set lobjdb = CreateObject("dbMissingTime.cdMissingTime")
    GetPayProcessedDate = lobjdb.GetPayProcessedDate(vsPayrollGroupCode, vdtScheduleDateTime)
    Set lobjdb = Nothing

    Exit Function
eh:
    Set lobjdb = Nothing

    TrsRaiseError MODULENAME, "GetPayProcessedDate", , , , , True
End Function

Public Function GetDepartmentSchedule(ByVal vadrsDepartmentPayPeriod As ADODB.Recordset _
                ) As ADODB.Recordset
On Error GoTo GetDepartmentSchedule_EH
Dim lobjSchedule As busSchedule.cbsSchedule
'Other
Dim llRecordIndex As Long, lsArrDept

    Set GetDepartmentSchedule = Nothing

    If (IsTheRSPopulated(vadrsDepartmentPayPeriod)) Then
        With vadrsDepartmentPayPeriod
            .Sort = "DEPT_NUM"
            .MoveFirst

            ReDim lsArrDept(.RecordCount - 1)

            For llRecordIndex = 1 To .RecordCount
                lsArrDept(llRecordIndex - 1) = !DEPT_NUM

                .MoveNext
            Next llRecordIndex

            .MoveFirst

            App.LogEvent "GetAdjustedBaseDepartmentSchedules (Number of Depts.: " & CStr(.RecordCount) & ")" & _
                    " started at " & Now _
                    , vbLogEventTypeInformation

            Set lobjSchedule = CreateObject("busSchedule.cbsSchedule")
            Set GetDepartmentSchedule = lobjSchedule.GetAdjustedBaseDepartmentSchedules( _
                lsArrDept _
                , !PAY_PER_START_DATE _
                , !PAY_PER_END_DATE _
                , trBaseSchedNoDetails _
                , trAllScheduleLevels)
            Set lobjSchedule = Nothing

            App.LogEvent "GetAdjustedBaseDepartmentSchedules (Number of Depts.: " & CStr(.RecordCount) & ")" & _
                    " ended at " & Now _
                    , vbLogEventTypeInformation
        End With
    End If

    Erase lsArrDept

    Exit Function
GetDepartmentSchedule_EH:
    Set lobjSchedule = Nothing
    Erase lsArrDept

    TrsRaiseError MODULENAME, "GetDepartmentSchedule" _
        , , , , , True
End Function
Private Function PrepareBriefScheduleRS(ByVal vadrsSchedule As ADODB.Recordset _
                    ) As ADODB.Recordset
On Error GoTo PrepareBriefScheduleRS_EH
'Recordsets
Dim ladrsBriefSchedule As ADODB.Recordset
'Index
Dim llRecordIndex As Long
Dim llPrevEmployee As Long, ldtPrevDate As Date
Dim lbAddRecord As Boolean, lbUpdateRecord As Boolean

    Set PrepareBriefScheduleRS = Nothing

    If (IsTheRSPopulated(vadrsSchedule)) Then
        Set ladrsBriefSchedule = New ADODB.Recordset

        With ladrsBriefSchedule
            .Fields.Append "EMP_NUM", adInteger, , adFldUpdatable + adFldIsNullable
            .Fields.Append "CHARGE_DATE", adDate, , adFldUpdatable + adFldIsNullable
            .Fields.Append "HOURS", adDecimal, , adFldUpdatable + adFldIsNullable
                .Fields.Item("HOURS").Precision = 3
                .Fields.Item("HOURS").NumericScale = 1
            .Fields.Append "OT_IND", adChar, 1, adFldUpdatable + adFldIsNullable

            .CursorLocation = adUseClient
            .CursorType = adOpenStatic
            .LockType = adLockBatchOptimistic

            .Open
        End With

        With vadrsSchedule
            .Filter = adFilterNone

            .Filter = "SHIFT_HOURS > 0.0"
            If Not (.EOF) Then
                .Sort = "EMP_NUM, CHARGE_DATE, SHIFT_NUM"
                .MoveFirst

                llPrevEmployee = 0
                ldtPrevDate = vbNull

                For llRecordIndex = 1 To .RecordCount
                    lbAddRecord = False
                    lbUpdateRecord = False

                    If (llPrevEmployee <> !EMP_NUM) Then
                        lbAddRecord = True

                        llPrevEmployee = 0
                        ldtPrevDate = vbNull
                    Else
                        If (ldtPrevDate <> !CHARGE_DATE) Then
                            lbAddRecord = True

                            ldtPrevDate = vbNull
                        Else
                            lbUpdateRecord = True
                        End If
                    End If

                    If (lbUpdateRecord) Then
                        If (IsTheRSPopulated(ladrsBriefSchedule)) Then
                            ladrsBriefSchedule.Filter = adFilterNone
                            ladrsBriefSchedule.Filter = "EMP_NUM = " & CStr(!EMP_NUM) & _
                                " AND CHARGE_DATE = #" & !CHARGE_DATE & "#" & _
                                " AND OT_IND = 'N'"

                            If Not (ladrsBriefSchedule.EOF) Then
                                ladrsBriefSchedule.MoveFirst
                                ladrsBriefSchedule!Hours = ladrsBriefSchedule!Hours + !SHIFT_HOURS
                                ladrsBriefSchedule.Update

                                ladrsBriefSchedule.Filter = adFilterNone
                            Else
                                ladrsBriefSchedule.Filter = adFilterNone

                                lbAddRecord = True
                            End If

                            ladrsBriefSchedule.Filter = adFilterNone
                        Else
                            lbAddRecord = True
                        End If
                    End If

                    If (lbAddRecord) Then
                        If (!SHIFT_HOURS > 0) Then
                            ladrsBriefSchedule.AddNew
                            ladrsBriefSchedule!EMP_NUM = !EMP_NUM
                            ladrsBriefSchedule!CHARGE_DATE = !CHARGE_DATE
                            ladrsBriefSchedule!Hours = !SHIFT_HOURS
                            ladrsBriefSchedule!OT_IND = "N"
                            ladrsBriefSchedule.Update
                        End If
                    End If

                    llPrevEmployee = !EMP_NUM
                    ldtPrevDate = !CHARGE_DATE

                    .MoveNext
                Next llRecordIndex
            End If

            .Filter = adFilterNone

            .Filter = "OT_HOURS > 0.0"
            If Not (.EOF) Then
                .Sort = "EMP_NUM, CHARGE_DATE, SHIFT_NUM"
                .MoveFirst

                llPrevEmployee = 0
                ldtPrevDate = vbNull

                For llRecordIndex = 1 To .RecordCount
                    lbAddRecord = False
                    lbUpdateRecord = False

                    If (llPrevEmployee <> !EMP_NUM) Then
                        lbAddRecord = True

                        llPrevEmployee = 0
                        ldtPrevDate = vbNull
                    Else
                        If (ldtPrevDate <> !CHARGE_DATE) Then
                            lbAddRecord = True

                            ldtPrevDate = vbNull
                        Else
                            lbUpdateRecord = True
                        End If
                    End If

                    If (lbUpdateRecord) Then
                        If (IsTheRSPopulated(ladrsBriefSchedule)) Then
                            ladrsBriefSchedule.Filter = adFilterNone
                            ladrsBriefSchedule.Filter = "EMP_NUM = " & CStr(!EMP_NUM) & _
                                " AND CHARGE_DATE = #" & !CHARGE_DATE & "#" & _
                                " AND OT_IND = 'Y'"

                            If Not (ladrsBriefSchedule.EOF) Then
                                ladrsBriefSchedule.MoveFirst
                                ladrsBriefSchedule!Hours = ladrsBriefSchedule!Hours + !OT_HOURS
                                ladrsBriefSchedule.Update

                                ladrsBriefSchedule.Filter = adFilterNone
                            Else
                                ladrsBriefSchedule.Filter = adFilterNone

                                lbAddRecord = True
                            End If

                            ladrsBriefSchedule.Filter = adFilterNone
                        Else
                            lbAddRecord = True
                        End If
                    End If

                    If (lbAddRecord) Then
                        If (!OT_HOURS > 0) Then
                            ladrsBriefSchedule.AddNew
                            ladrsBriefSchedule!EMP_NUM = !EMP_NUM
                            ladrsBriefSchedule!CHARGE_DATE = !CHARGE_DATE
                            ladrsBriefSchedule!Hours = !OT_HOURS
                            ladrsBriefSchedule!OT_IND = "Y"
                            ladrsBriefSchedule.Update
                        End If
                    End If

                    llPrevEmployee = !EMP_NUM
                    ldtPrevDate = !CHARGE_DATE

                    .MoveNext
                Next llRecordIndex
            End If

            .Filter = adFilterNone
        End With
    End If

    If (IsTheRSPopulated(ladrsBriefSchedule)) Then
        Set PrepareBriefScheduleRS = ladrsBriefSchedule
    Else
        Set PrepareBriefScheduleRS = Nothing
    End If

    Set ladrsBriefSchedule = Nothing

    Exit Function
PrepareBriefScheduleRS_EH:
    Set ladrsBriefSchedule = Nothing

    TrsRaiseError MODULENAME, "PrepareBriefScheduleRS" _
        , , , , , True
    
End Function
'run times PayGroups (E)
'weeks with holidays will shift the schedule (day 1  - day 15)
'                             PayProcessDate (day 21 - day 25)
' Sched  Day
' Task   16     17     18     19            20
'10:00   Run    Run    Run    Run & Copy    Run & Copy

'weeks with holidays will shift the schedule (day 16 - day 31)
'                             PayProcessDate (day 7  - day 9)
' Sched  Day
' Task   1      2      3      4             5
'10:00   Run    Run    Run    Run & Copy    Run & Copy

'Management Missing Time

'TODO: Make this obsolete
Public Function GetMgmtMissingTimeEmployees() As ADODB.Recordset
    Set GetMgmtMissingTimeEmployees = GetMgmtMissingTimeEmployees2(Now)
End Function

Public Function GetMgmtMissingTimeEmployees2(ByVal vdtScheduleDateTime As Date) As ADODB.Recordset
On Error GoTo eh
'Components
Dim lobjdb As dbMissingTime.cdMissingTime
Dim lobjSchedule As busSchedule.cbsSchedule
'RS
Dim ladrsNotify As ADODB.Recordset
Dim ladrsSchedule As ADODB.Recordset
'Other
Dim llIndex As Long, llArrOfEmps() As Long
Dim ldtStart As Date, ldtEnd As Date
Dim llScheduleIndex As Long, ldblHours As Double

    App.LogEvent "E GetMgmtMissingTimeEmployees2 started at " & Now, vbLogEventTypeInformation

    Set GetMgmtMissingTimeEmployees2 = Nothing

    'Get Missing Time Employees
    App.LogEvent "E db GetMgmtMissingTimeEmployees2 started at " & Now, vbLogEventTypeInformation

    Set lobjdb = CreateObject("dbMissingTime.cdMissingTime")
    Set ladrsNotify = lobjdb.GetMgmtMissingTimeEmployees2(vdtScheduleDateTime)
    Set lobjdb = Nothing

    App.LogEvent "E db GetMgmtMissingTimeEmployees2 ended at " & Now, vbLogEventTypeInformation
'''''
'''''    App.LogEvent "E GetAdjustedBaseSchedule & Set Send Indicator started at " & Now, vbLogEventTypeInformation
'''''
'''''    'Get and Compare against schedule. Determine and set if the employee
'''''    'should be notified
'''''    If (IsTheRSPopulated(ladrsNotify)) Then
'''''        With ladrsNotify
'''''            .Sort = "EMP_NUM"
'''''            .MoveFirst
'''''
'''''            ldtStart = !PAY_PER_START_DATE
'''''            ldtEnd = !PAY_PER_END_DATE
'''''
'''''            ReDim llArrOfEmps(.RecordCount - 1)
'''''
'''''            For llIndex = 1 To .RecordCount
'''''                llArrOfEmps(llIndex - 1) = !EMP_NUM
'''''
'''''                .MoveNext
'''''            Next llIndex
'''''
'''''            .MoveFirst
'''''
'''''            App.LogEvent "E GetAdjustedBaseSchedule started at " & Now, vbLogEventTypeInformation
'''''
'''''            Set lobjSchedule = CreateObject("busSchedule.cbsSchedule")
'''''            Set ladrsSchedule = lobjSchedule.GetAdjustedBaseSchedule(llArrOfEmps _
'''''                                    , ldtStart _
'''''                                    , ldtEnd _
'''''                                    , Nothing _
'''''                                    , False _
'''''                                    , trAdjBaseSchedWithExceptions _
'''''                                    , trAllScheduleLevels _
'''''                                    , trRelief _
'''''                                    , trAllExceptions)
'''''            Set lobjSchedule = Nothing
'''''
'''''            App.LogEvent "E GetAdjustedBaseSchedule ended at " & Now, vbLogEventTypeInformation
'''''
'''''            App.LogEvent "E Set Send Indicator started at " & Now, vbLogEventTypeInformation
'''''
'''''            If Not (ladrsSchedule Is Nothing) Then
'''''                If Not (ladrsSchedule.EOF And ladrsSchedule.BOF) Then
'''''                    ladrsSchedule.Sort = "EMP_NUM, CHARGE_DATE, DATE_WORKED, START_TIME"
'''''                    ladrsSchedule.MoveFirst
'''''
'''''                    For llIndex = .RecordCount To 1 Step -1
'''''                        ladrsSchedule.Filter = adFilterNone
'''''                        ladrsSchedule.Filter = "(EMP_NUM = " & !EMP_NUM & _
'''''                                        " AND SHIFT_HOURS > 0.0" & _
'''''                                        " AND SYSTEM_EXC_CD = 'SWT') OR " & _
'''''                                        " (EMP_NUM = " & !EMP_NUM & _
'''''                                        " AND SHIFT_HOURS > 0.0" & _
'''''                                        " AND SYSTEM_EXC_CD = 'RWT') OR " & _
'''''                                        " (EMP_NUM = " & !EMP_NUM & _
'''''                                        " AND SHIFT_HOURS > 0.0" & _
'''''                                        " AND SYSTEM_EXC_CD = 'TSA') OR " & _
'''''                                        " (EMP_NUM = " & !EMP_NUM & _
'''''                                        " AND SHIFT_HOURS > 0.0" & _
'''''                                        " AND SYSTEM_EXC_CD = 'TSR') OR " & _
'''''                                        " (EMP_NUM = " & !EMP_NUM & _
'''''                                        " AND SHIFT_HOURS > 0.0" & _
'''''                                        " AND SYSTEM_EXC_CD = 'RLF') OR " & _
'''''                                        " (EMP_NUM = " & !EMP_NUM & _
'''''                                        " AND SHIFT_HOURS > 0.0" & _
'''''                                        " AND SYSTEM_EXC_CD = 'TWS')"
'''''                        ldblHours = 0#
'''''
'''''                        If Not (ladrsSchedule.EOF) Then
'''''                            ladrsSchedule.MoveFirst
'''''
'''''                            For llScheduleIndex = 1 To ladrsSchedule.RecordCount
'''''                                ldblHours = ldblHours + ladrsSchedule!Hours
'''''
'''''                                ladrsSchedule.MoveNext
'''''                            Next llScheduleIndex
'''''                        End If
'''''
'''''                        ladrsSchedule.Filter = adFilterNone
'''''
'''''                        If (ldblHours > 0) Then
'''''                            If (!TOT_HRS >= ldblHours) Then
'''''                                !SEND_IND = "N"
'''''                                .Update
'''''                            End If
'''''                        End If
'''''
'''''                        .MoveNext
'''''                        DoEvents
'''''                    Next llIndex
'''''                End If
'''''            End If
'''''
'''''            App.LogEvent "E Set Send Indicator ended at " & Now, vbLogEventTypeInformation
'''''
'''''            Set ladrsSchedule = Nothing
'''''        End With
'''''    End If

    Set lobjSchedule = Nothing
    Set lobjdb = Nothing

    App.LogEvent "E GetAdjustedBaseSchedule & Set Send Indicator ended at " & Now, vbLogEventTypeInformation

    Set ladrsSchedule = Nothing

    Set GetMgmtMissingTimeEmployees2 = ladrsNotify
    Set ladrsNotify = Nothing

    App.LogEvent "E GetMgmtMissingTimeEmployees2 ended at " & Now, vbLogEventTypeInformation

    Exit Function
eh:
    Set lobjSchedule = Nothing
    Set lobjdb = Nothing

    Set ladrsSchedule = Nothing
    Set ladrsNotify = Nothing

    TrsRaiseError MODULENAME, "GetMgmtMissingTimeEmployees2", , , , , True
End Function

'//for testing without actually sending an email!!!
Private Function CreateEMailRecordset() As ADODB.Recordset
    Dim rs As ADODB.Recordset
    
On Error GoTo eh:
    Set rs = New ADODB.Recordset
    With rs
        .Fields.Append "From", adVarChar, 100, adFldUpdatable
        .Fields.Append "To", adVarChar, 100, adFldUpdatable ' + adFldIsNullable
        .Fields.Append "CC", adVarChar, 100, adFldUpdatable
        .Fields.Append "BCC", adVarChar, 100, adFldUpdatable
        .Fields.Append "Subject", adVarChar, 1000, adFldUpdatable
        .Fields.Append "TextBody", adVarChar, 2550000, adFldUpdatable
        .Fields.Append "HTMLBody", adVarChar, 2550000, adFldUpdatable
        .Fields.Append "Tag", adVarChar, 100, adFldUpdatable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open
    End With
    Set CreateEMailRecordset = rs
Exit Function
eh:
    Set rs = Nothing
    TrsRaiseError MODULENAME, "CreateEMailRecordset"
End Function

Private Function SaveRecordset(ByVal rs As ADODB.Recordset, ByVal sFolderPath As String, ByVal sFileName As String)
    Dim sFilePath As String
    Dim oFSO As Scripting.FileSystemObject
    
    Set oFSO = New Scripting.FileSystemObject

    CreateFolder sFolderPath, True
    sFilePath = oFSO.BuildPath(sFolderPath, sFileName)
    sFilePath = GetUniqueFileName(sFilePath)
    If Not rs Is Nothing Then
        rs.Save sFilePath, adPersistXML
    End If
    
    Set oFSO = Nothing
End Function

Private Function SaveEMailRecordset(ByVal rs As ADODB.Recordset, ByVal sFolderPath As String)
    Dim sHTMLBody As String, sTextBody As String
    Dim sFileName As String, sFilePath As String
    Dim oFSO As Scripting.FileSystemObject
    Dim ts As TextStream
    
    Set oFSO = New Scripting.FileSystemObject
    
    If IsTheRSPopulated(rs) Then
        rs.MoveFirst

        CreateFolder sFolderPath, True
        sFilePath = oFSO.BuildPath(sFolderPath, "EMailRecordset.xml")
        sFilePath = GetUniqueFileName(sFilePath)
        rs.Save sFilePath, adPersistXML
        
        sFilePath = oFSO.BuildPath(sFolderPath, "EMailRecordset.adtg")
        sFilePath = GetUniqueFileName(sFilePath)
        rs.Save sFilePath, adPersistADTG
        
        Do While Not rs.EOF
            sHTMLBody = rs.Fields("HTMLBody").Value
            sTextBody = rs.Fields("TextBody").Value
            
            sFileName = rs.Fields("Subject").Value & "_" & rs.Fields("To").Value & "_" & rs.Fields("Tag").Value
            sFileName = Replace(sFileName, " ", "")
            sFileName = Replace(sFileName, """", "'")
            sFileName = Replace(sFileName, ":", ".")
            sFileName = Replace(sFileName, "<", "[")
            sFileName = Replace(sFileName, ">", "]")
            sFileName = Replace(sFileName, "/", ".")
            sFilePath = oFSO.BuildPath(sFolderPath, sFileName)
            
            If Trim(sHTMLBody) <> "" Then
                sFilePath = sFilePath & ".html"
                sFilePath = GetUniqueFileName(sFilePath)
                Set ts = oFSO.CreateTextFile(sFilePath, False)
                ts.Write sHTMLBody
                ts.Close
            ElseIf Trim(sTextBody) <> "" Then
                sFilePath = sFilePath & ".txt"
                sFilePath = GetUniqueFileName(sFilePath)
                Set ts = oFSO.CreateTextFile(sFilePath, False)
                ts.Write sTextBody
                ts.Close
            End If
            
            rs.MoveNext
        Loop
    End If
    Set oFSO = Nothing
End Function

Private Function GetUniqueFileName(ByVal sFilePath As String) As String
    Dim oFSO As Scripting.FileSystemObject
    Dim nSuffix As Integer, sSuffix As String
    Dim sFolderPath As String, sFileName As String, sExt As String, sNewFilePath As String
    Dim i As Integer
    
    Set oFSO = New Scripting.FileSystemObject
    
    nSuffix = 1
    sNewFilePath = sFilePath
    Do While oFSO.FileExists(sNewFilePath) And nSuffix < 26
        sSuffix = Chr(64 + nSuffix)
        
        sFolderPath = oFSO.GetParentFolderName(sFilePath)
        sFileName = oFSO.GetFileName(sFilePath)
        '//remove extension
        i = InStrRev(sFileName, ".")
        If i >= 1 And i <= Len(sFileName) Then
            sFileName = Left(sFileName, i - 1)
        End If
        
        sExt = oFSO.GetExtensionName(sFilePath)
        sNewFilePath = oFSO.BuildPath(sFolderPath, sFileName & sSuffix & "." & sExt)
        nSuffix = nSuffix + 1
    Loop

    GetUniqueFileName = sNewFilePath
    Set oFSO = Nothing
End Function

'//-----Create any missing parent folders as well-----//
Private Function CreateFolder(ByVal sFolder As String, ByVal blnAddMissingParentFolders As Boolean)
    Dim aryFolder() As String
    Dim oFSO As FileSystemObject
    Dim nFolderCount As Integer, nFolderExists As Integer, nFolderCreate As Integer
    Dim sParent As String
    
    Set oFSO = New FileSystemObject
    
    If Not oFSO.FolderExists(sFolder) Then
        If blnAddMissingParentFolders Then
            '//-----build array of folders in reverse order
            nFolderCount = 0
            sParent = sFolder
            Do While sParent <> ""
                ReDim Preserve aryFolder(nFolderCount)
                aryFolder(nFolderCount) = sParent
                nFolderCount = nFolderCount + 1
              
                sParent = oFSO.GetParentFolderName(sParent)
            Loop
            
            '//-----determine which folder does exist
            For nFolderExists = 0 To nFolderCount - 1
                If oFSO.FolderExists(aryFolder(nFolderExists)) Then
                   Exit For
                End If
            Next
            
            For nFolderCreate = nFolderExists - 1 To 0 Step -1
                oFSO.CreateFolder aryFolder(nFolderCreate)
            Next
        Else
            oFSO.CreateFolder sFolder
        End If
    End If
    Set oFSO = Nothing
End Function

Private Sub AddEMailRecord(ByVal rs As ADODB.Recordset, ByVal vMessage As CDO.Message, ByVal vsTag As String)
On Error GoTo eh:
   
    'Dim EMailObject(0 To 6) As Variant
    Dim EMailObject As cboEMail

    Set EMailObject = CreateEMailObject(vMessage)

    ValidateEMailRecord rs, EMailObject
    
    rs.AddNew
    
    rs.Fields("From").Value = Left(EMailObject.From, rs.Fields("From").DefinedSize)
    rs.Fields("To").Value = Left(EMailObject.To_, rs.Fields("To").DefinedSize)
    rs.Fields("CC").Value = Left(EMailObject.CC, rs.Fields("CC").DefinedSize)
    rs.Fields("BCC").Value = Left(EMailObject.BCC, rs.Fields("BCC").DefinedSize)
    rs.Fields("Subject").Value = Left(EMailObject.Subject, rs.Fields("Subject").DefinedSize)
    rs.Fields("TextBody").Value = Left(EMailObject.TextBody, rs.Fields("TextBody").DefinedSize)
    rs.Fields("HTMLBody").Value = Left(EMailObject.HTMLBody, rs.Fields("HTMLBody").DefinedSize)
    rs.Fields("Tag").Value = Left(vsTag, rs.Fields("Tag").DefinedSize)
    rs.Update
Exit Sub
eh:
    TrsRaiseError MODULENAME, "AddEMailRecord"
End Sub

Private Sub TraceEMail(ByVal vMessage As CDO.Message)
On Error GoTo eh:
  
    If TRACING = 1 Then
        If Trim(vMessage.BCC) = "" Then
            vMessage.BCC = "TRISDevelopment@ameren.com"
        Else
            If ParseEMailAddress(Trim(vMessage.BCC)) <> "TRISDevelopment@ameren.com" Then
                vMessage.BCC = vMessage.BCC & ", TRISDevelopment@ameren.com"
            End If
        End If
    End If

Exit Sub
eh:
    TrsRaiseError MODULENAME, "TraceEMail"
End Sub

Private Function CreateEMailObject(ByVal vMessage As CDO.Message) As cboEMail
On Error GoTo eh:
   
   Dim EMailObject As cboEMail
'ParseEMailAddress

    Set EMailObject = New cboEMail
    EMailObject.From = ParseEMailAddress(vMessage.From)
    EMailObject.To_ = ParseEMailAddress(vMessage.To)
    EMailObject.CC = ParseEMailAddress(vMessage.CC)
    EMailObject.BCC = ParseEMailAddress(vMessage.BCC)
    EMailObject.Subject = vMessage.Subject
    EMailObject.TextBody = vMessage.TextBody
    EMailObject.HTMLBody = vMessage.HTMLBody

    Set CreateEMailObject = EMailObject
Exit Function
eh:
    TrsRaiseError MODULENAME, "CreateEMailObject"
End Function

Private Sub ValidateEMailRecord(ByVal rs As ADODB.Recordset, ByVal EMailObject As cboEMail)
On Error GoTo eh:
   
    If Len(EMailObject.From) > rs.Fields("From").DefinedSize Then
        App.LogEvent "ValidateEMailRecord Len(EMailObject.From) = " & Len(EMailObject.From), vbLogEventTypeError
    End If
    
    If Len(EMailObject.To_) > rs.Fields("To").DefinedSize Then
        App.LogEvent "ValidateEMailRecord Len(EMailObject.To) = " & Len(EMailObject.To_), vbLogEventTypeError
    End If
    
    If Len(EMailObject.CC) > rs.Fields("CC").DefinedSize Then
        App.LogEvent "ValidateEMailRecord Len(EMailObject.CC) = " & Len(EMailObject.CC), vbLogEventTypeError
    End If
    
    If Len(EMailObject.BCC) > rs.Fields("BCC").DefinedSize Then
        App.LogEvent "ValidateEMailRecord Len(EMailObject.BCC) = " & Len(EMailObject.BCC), vbLogEventTypeError
    End If
    
    If Len(EMailObject.Subject) > rs.Fields("Subject").DefinedSize Then
        App.LogEvent "ValidateEMailRecord Len(EMailObject.Subject) = " & Len(EMailObject.Subject), vbLogEventTypeError
    End If
    
    If Len(EMailObject.TextBody) > rs.Fields("TextBody").DefinedSize Then
        App.LogEvent "ValidateEMailRecord Len(EMailObject.TextBody) = " & Len(EMailObject.TextBody), vbLogEventTypeError
    End If
    
    If Len(EMailObject.HTMLBody) > rs.Fields("HTMLBody").DefinedSize Then
        App.LogEvent "ValidateEMailRecord Len(EMailObject.HTMLBody) = " & Len(EMailObject.HTMLBody), vbLogEventTypeError
    End If
Exit Sub
eh:
    TrsRaiseError MODULENAME, "ValidateEMailRecord"
End Sub

Private Function AddEMailRecord2(ByVal rs As ADODB.Recordset, ByVal vsFrom As String, ByVal vsTo As String, _
                                ByVal vsCC As String, ByVal vsBCC As String, _
                                ByVal vsSubject As String, ByVal vsTextBody As String, ByVal vsHTMLBody As String _
                                ) As ADODB.Recordset
On Error GoTo eh:
    Set rs = New ADODB.Recordset
    
    rs.AddNew
    rs.Fields("From").Value = vsFrom
    rs.Fields("To").Value = vsTo
    rs.Fields("CC").Value = vsCC
    rs.Fields("BCC").Value = vsBCC
    rs.Fields("Subject").Value = vsSubject
    rs.Fields("TextBody").Value = vsTextBody
    rs.Fields("HTMLBody").Value = vsHTMLBody
    rs.Update
Exit Function
eh:
    TrsRaiseError MODULENAME, "AddEMailRecord2"
End Function

Private Function ParseEMailAddress(ByVal sEMailAddress) As String
    Dim v As Variant
    Dim i As Integer, j As Integer
    Dim s As String
    
'    v = GetRegExSubMatches("""HRCentralPayroll@ameren.com"" <HRCentralPayroll@ameren.com>, ""#HRCentralPayroll@ameren.com"" <#HRCentralPayroll@ameren.com>", _
'                            """([a-zA-Z@#.]*)""\s*\<[a-zA-Z@#.]*\>,?\s*")
    
'    If Trim(sEMailAddress) <> "" Then
        v = GetRegExSubMatches(sEMailAddress, """([a-zA-Z0-9@#.]*)""\s*\<[a-zA-Z0-9@#.]*\>,?\s*")
        For i = LBound(v) To UBound(v)
            If Not IsEmpty(v(i)) Then
                For j = LBound(v(i)) To UBound(v(i))
                    s = s & v(i)(j)
                Next j
                If i < UBound(v) Then
                    s = s & ", "
                End If
            End If
            
        Next i
        ParseEMailAddress = s
'    End If
    
    If ParseEMailAddress = "" Then
        ParseEMailAddress = sEMailAddress
    End If
End Function

'//TODO:Refactor
Private Function IsRegExMatch(ByVal Source As String, ByVal Pattern As String) As Boolean
    Dim re As RegExp
    Dim matches As MatchCollection
    Set re = New RegExp
    
    re.IgnoreCase = True
    re.Global = True
    re.Pattern = Pattern
    
    Set matches = re.Execute(Source)
    
    IsRegExMatch = (matches.Count > 0)
    Set re = Nothing
End Function

Private Function GetRegExGroups(ByVal Source As String, ByVal Pattern As String) As Variant
    Dim re As RegExp
    Dim matches As MatchCollection
    Dim groups() As Variant
    Dim i As Integer
    
    Set re = New RegExp
  
    re.IgnoreCase = True
    re.Global = True
    re.Pattern = Pattern
    
    ReDim groups(0)
    
    Set matches = re.Execute(Source)
    
    If matches.Count > 0 Then
        'Print matches.Item(0).SubMatches.Item(1)
        If matches.Item(0).SubMatches.Count > 0 Then
            ReDim groups(matches.Item(0).SubMatches.Count - 1)
            For i = 0 To matches.Item(0).SubMatches.Count - 1
                groups(i) = matches.Item(0).SubMatches.Item(i)
            Next i
        End If
    End If
    
    GetRegExGroups = groups
    Set re = Nothing
End Function

Private Function GetRegExSubMatches(ByVal Source As String, ByVal Pattern As String) As Variant
    Dim re As RegExp
    Dim matches As MatchCollection
    Dim varMatches() As Variant
    Dim varSubMatches() As Variant
    Dim i As Integer, j As Integer
    
    Set re = New RegExp
  
    re.IgnoreCase = True
    re.Global = True
    re.Pattern = Pattern
    
    ReDim varMatches(0)
    
    Set matches = re.Execute(Source)
    
    If matches.Count > 0 Then
        ReDim varMatches(matches.Count - 1)
        For j = 0 To matches.Count - 1
            If matches.Item(j).SubMatches.Count > 0 Then
                ReDim varSubMatches(matches.Item(j).SubMatches.Count - 1)
                For i = 0 To matches.Item(j).SubMatches.Count - 1
                    varSubMatches(i) = matches.Item(j).SubMatches.Item(i)
                Next i
                varMatches(j) = varSubMatches
            End If
        Next j
    End If
    
    GetRegExSubMatches = varMatches
    Set re = Nothing
End Function

