VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "cbsEAMNotify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Project:   Post TRIS 5.5 Employee Absence Management
'Author:    Prem Chander
'Date:      October 2005
'Description:

'//--------------------------------------------------------------------------------------------------------------------------------------//
'//all but GetEAMNotifyPreference is obsolete and has been moved to Notification or Notification2
'//--------------------------------------------------------------------------------------------------------------------------------------//

Option Explicit

Private Const MODULENAME As String = "busEAM.cbsEAMNotify"
Public Function GetEAMNotifyPreference(ByVal vlTRISUserNumber As Long _
                ) As ADODB.Recordset
On Error GoTo GetEAMNotifyPreference_EH
Dim lobjdb As dbEAM.cdEAMNotify, lsUser As String

    Set GetEAMNotifyPreference = Nothing

    Select Case (Len(CStr(vlTRISUserNumber)))
        Case 0
            Exit Function
        Case 1
            lsUser = "00000000" & CStr(vlTRISUserNumber)
        Case 2
            lsUser = "0000000" & CStr(vlTRISUserNumber)
        Case 3
            lsUser = "000000" & CStr(vlTRISUserNumber)
        Case 4
            lsUser = "00000" & CStr(vlTRISUserNumber)
        Case 5
            lsUser = "0000" & CStr(vlTRISUserNumber)
        Case 6
            lsUser = "000" & CStr(vlTRISUserNumber)
        Case 7
            lsUser = "00" & CStr(vlTRISUserNumber)
        Case 8
            lsUser = "0" & CStr(vlTRISUserNumber)
        Case 9
            lsUser = CStr(vlTRISUserNumber)
    End Select

    Set lobjdb = CreateObject("dbEAM.cdEAMNotify")
    Set GetEAMNotifyPreference = lobjdb.GetEAMNotifyPreference(lsUser)
    Set lobjdb = Nothing

    Exit Function
GetEAMNotifyPreference_EH:
    Set lobjdb = Nothing

    TrsRaiseError MODULENAME _
        , "GetEAMNotifyPreference" _
        , , , , _
        , True

End Function
Public Sub SendEAMFollowUpEmails()
On Error GoTo SendEAMFollowUpEmails_EH

    ''NO LONGER USED

    Exit Sub
SendEAMFollowUpEmails_EH:
    TrsRaiseError MODULENAME _
        , "SendEAMFollowUpEmails" _
        , , , , _
        , True
End Sub
Public Sub SendEAMFollowUpEmails2(ByVal viByPreferenceOrDate As Integer)
On Error GoTo SendEAMFollowUpEmails2_EH
'Notify
Dim ladrsNotify As ADODB.Recordset, ladrsEMPWG As ADODB.Recordset

    Select Case (viByPreferenceOrDate)
        Case 1
            Set ladrsNotify = GetContactsDueByPreference
        Case 2
            Set ladrsNotify = GetContactsDueByFUDate
    End Select

    Set ladrsEMPWG = GetEmployeesInEAMNotifyWG(ladrsNotify)

    SendMail ladrsNotify _
        , ladrsEMPWG _
        , viByPreferenceOrDate

    Set ladrsEMPWG = Nothing
    Set ladrsNotify = Nothing

    Exit Sub
SendEAMFollowUpEmails2_EH:
    Set ladrsEMPWG = Nothing
    Set ladrsNotify = Nothing

    TrsRaiseError MODULENAME _
        , "SendEAMFollowUpEmails2" _
        , , , , _
        , True
End Sub
Public Sub SendEAMRtnToWrkEmails()
On Error GoTo SendEAMRtnToWrkEmails_EH
'Notify
Dim ladrsNotify As ADODB.Recordset, ladrsEMPWG As ADODB.Recordset

    Set ladrsNotify = GetRtnToWrksDue
    Set ladrsEMPWG = GetEmployeesInEAMNotifyWG(ladrsNotify)

    SendMail ladrsNotify _
        , ladrsEMPWG _
        , 3

    Set ladrsEMPWG = Nothing
    Set ladrsNotify = Nothing

    Exit Sub
SendEAMRtnToWrkEmails_EH:
    Set ladrsEMPWG = Nothing
    Set ladrsNotify = Nothing

    TrsRaiseError MODULENAME _
        , "SendEAMRtnToWrkEmails" _
        , , , , _
        , True
End Sub
Public Function GetContactsDueByPreference() As ADODB.Recordset
On Error GoTo GetContactsDueByPreference_EH
'Get Notify Info
Dim lobjdb As dbEAM.cdEAMNotify

    Set GetContactsDueByPreference = Nothing

    Set lobjdb = CreateObject("dbEAM.cdEAMNotify")
    Set GetContactsDueByPreference = lobjdb.GetContactsDueByPreference
    Set lobjdb = Nothing

    Exit Function
GetContactsDueByPreference_EH:
    Set lobjdb = Nothing

    TrsRaiseError MODULENAME _
        , "GetContactsDueByPreference" _
        , , , , _
        , True

End Function
Public Function GetContactsDueByFUDate() As ADODB.Recordset
On Error GoTo GetContactsDueByFUDate_EH
'Get Notify Info
Dim lobjdb As dbEAM.cdEAMNotify

    Set GetContactsDueByFUDate = Nothing

    Set lobjdb = CreateObject("dbEAM.cdEAMNotify")
    Set GetContactsDueByFUDate = lobjdb.GetContactsDueByFUDate
    Set lobjdb = Nothing

    Exit Function
GetContactsDueByFUDate_EH:
    Set lobjdb = Nothing

    TrsRaiseError MODULENAME _
        , "GetContactsDueByFUDate" _
        , , , , _
        , True

End Function
Public Function GetRtnToWrksDue() As ADODB.Recordset
On Error GoTo GetRtnToWrksDue_EH
'Get Notify Info
Dim lobjdb As dbEAM.cdEAMNotify

    Set GetRtnToWrksDue = Nothing

    Set lobjdb = CreateObject("dbEAM.cdEAMNotify")
    Set GetRtnToWrksDue = lobjdb.GetCurrentRtnToWrksDue
    Set lobjdb = Nothing

    Exit Function
GetRtnToWrksDue_EH:
    Set lobjdb = Nothing

    TrsRaiseError MODULENAME _
        , "GetRtnToWrksDue" _
        , , , , _
        , True

End Function
Private Function GetEmployeesInEAMNotifyWG(ByVal vadrsNotify As ADODB.Recordset _
                ) As ADODB.Recordset
On Error GoTo GetEmployeesInEAMNotifyWG_EH
'Component
Dim lobjdbWG As dbWorkGroup.cdWorkGroup
'Other
Dim llArrOfWG() As Long, llPrevWGNum As Long
Dim llIndex As Long

    Set GetEmployeesInEAMNotifyWG = Nothing

    If (RecordsExist(vadrsNotify)) Then
        With vadrsNotify
            .Filter = adFilterNone

            .Filter = "ABS_NOTIFY_WG_NUM > 0"
            If Not (.EOF) Then
                .Sort = "ABS_NOTIFY_WG_NUM"
                .MoveFirst

                ReDim llArrOfWG(.RecordCount - 1)

                llPrevWGNum = 0

                For llIndex = 1 To .RecordCount
                    If (llPrevWGNum <> !ABS_NOTIFY_WG_NUM) Then
                        llArrOfWG(llIndex - 1) = !ABS_NOTIFY_WG_NUM
                    End If

                    llPrevWGNum = !ABS_NOTIFY_WG_NUM

                    .MoveNext
                Next llIndex

                Set lobjdbWG = CreateObject("dbWorkGroup.cdWorkGroup")
                Set GetEmployeesInEAMNotifyWG = _
                            lobjdbWG.GetEmployeesAssignedToWorkGroup(llArrOfWG _
                                    , Date, Date _
                                    , True)
                Set lobjdbWG = Nothing
            End If

            .Filter = adFilterNone
            .MoveFirst
        End With
    End If

    Exit Function
GetEmployeesInEAMNotifyWG_EH:
    Set lobjdbWG = Nothing

    TrsRaiseError MODULENAME _
        , "GetEmployeesInEAMNotifyWG" _
        , , , , _
        , True

End Function
Private Sub SendMail(ByVal vadrsNotify As ADODB.Recordset _
            , ByVal vadrsEmpInWG As ADODB.Recordset _
            , ByVal viContactByPrefOrDateOrRtnToWrk As Integer)
On Error GoTo SendMail_EH
'Get Region
Dim lobjCnMgr As util_DBConnectMgr.clsuDBConnectMgr
Dim lsDB2 As String
'SMTP
Dim lobjCDOMsg As CDO.Message
Dim lobjCDOConfig As CDO.Configuration
Dim ladFlds As ADODB.Fields, lsUser As String
'Other
Dim llIndex As Long, llAnotherIndex As Long

    Set lobjCnMgr = CreateObject("util_DBConnectMgr.clsuDBConnectMgr")
    lsDB2 = lobjCnMgr.GetDataSourceName
    Set lobjCnMgr = Nothing

    If (InStr(1, lsDB2, "LUW") > 0) Then
        lsDB2 = Trim(Left(lsDB2, 7))
    Else
        lsDB2 = Trim(Left(lsDB2, 4))
    End If

    Set lobjCDOConfig = CreateObject("CDO.Configuration")
    Set ladFlds = lobjCDOConfig.Fields
    ladFlds(cdoSendUsingMethod) = cdoSendUsingPort

    'Per information from Mr. Travis Foley
'    Select Case (Environ("COMPUTERNAME"))
'        Case "TRISQA1", "TRISNLB1", "TRISNLB2" _
'            , "TRISPROD1", "TRISPROD3", "TRISPROD4" _
'            , "TRISPROD5", "TRISPROD6", "TRISPROD7"
            ladFlds(cdoSMTPServer) = "MAIL.AMEREN.COM"
            ladFlds(cdoSMTPServerPort) = 25
'        Case Else
'            ladFlds(cdoSMTPServer) = "appmta.corp.dir.ameren.com"
'            ladFlds(cdoSMTPServerPort) = 7257
'    End Select

    ladFlds(cdoSMTPAuthenticate) = cdoAnonymous
    ladFlds.Update

    If (RecordsExist(vadrsNotify)) Then
        With vadrsNotify
            .MoveFirst

            For llIndex = 1 To .RecordCount
                Set lobjCDOMsg = CreateObject("CDO.Message")

                With lobjCDOMsg
                    Set .Configuration = lobjCDOConfig

                    .From = "ITAppMsg@ameren.com"
                    If (vadrsNotify!FU_NOTIFY_EMP_NUM <> vadrsNotify!ABS_PRIM_CONT_EMP) Then
                        If (vadrsNotify!FU_NOTIFY_EMP_NUM > 0) Then
                            .To = CStr(vadrsNotify!FU_NOTIFY_EMP_NUM) & "@ameren.com"
                        End If
                        If (vadrsNotify!ABS_PRIM_CONT_EMP > 0) Then
                            .CC = CStr(vadrsNotify!ABS_PRIM_CONT_EMP) & "@ameren.com"
                        End If
                    Else
                        If (vadrsNotify!FU_NOTIFY_EMP_NUM > 0) Then
                            .To = CStr(vadrsNotify!FU_NOTIFY_EMP_NUM) & "@ameren.com"
                        End If
                    End If

                    .BCC = "TRISDevelopment@ameren.com"

                    If (vadrsNotify!ABS_NOTIFY_WG_NUM > 0) Then
                        If (RecordsExist(vadrsEmpInWG)) Then
                            vadrsEmpInWG.Filter = adFilterNone

                            vadrsEmpInWG.Filter = "WG_NUM = " & _
                                                CStr(vadrsNotify!ABS_NOTIFY_WG_NUM)
                            If Not (vadrsEmpInWG.EOF) Then
                                vadrsEmpInWG.Sort = "EMP_NUM"
                                vadrsEmpInWG.MoveFirst

                                For llAnotherIndex = 1 To vadrsEmpInWG.RecordCount
                                    If (Trim(.CC) <> "") Then
                                        .CC = .CC & ", " & _
                                                CStr(vadrsEmpInWG!EMP_NUM) & "@ameren.com"
                                    Else
                                        .CC = .CC & _
                                                CStr(vadrsEmpInWG!EMP_NUM) & "@ameren.com"
                                    End If

                                    vadrsEmpInWG.MoveNext
                                Next llAnotherIndex

                                vadrsEmpInWG.MoveFirst
                            End If

                            vadrsEmpInWG.Filter = adFilterNone
                            vadrsEmpInWG.MoveFirst
                        End If
                    End If

                    .Subject = "TRIS Employee Absence Management"

                    Select Case (viContactByPrefOrDateOrRtnToWrk)
                        Case 1  'Contact by Preference
                            .Subject = .Subject & " (Follow-up Notification)..."
                            .TextBody = "This is to remind you that a Follow-up is due for" & _
                                " employee " & Trim(vadrsNotify!EMP_NAME) & " [" & CStr(vadrsNotify!EMP_NUM) & "]" & _
                                " regarding absence " & Trim(vadrsNotify!ABS_DESC) & " [" & CStr(vadrsNotify!ABS_NUM) & "]" & _
                                " which started on " & Trim(vadrsNotify!ABS_ST_TS) & "." & vbCrLf & vbCrLf & _
                                "Employee was last contacted " & CStr(vadrsNotify!ELAPSED_DAYS) & _
                                " days ago, on " & CStr(vadrsNotify!LAST_CONT_DT) & "." & vbCrLf & vbCrLf & _
                                "(This is an automated message. Please do not reply to this message.)"

                        Case 2  'Contact by Follow-up Date
                            .Subject = .Subject & " (Follow-up Notification)..."
                            .TextBody = "This is to remind you that a Follow-up is due for" & _
                                " employee " & Trim(vadrsNotify!EMP_NAME) & " [" & CStr(vadrsNotify!EMP_NUM) & "]" & _
                                " regarding absence " & Trim(vadrsNotify!ABS_DESC) & " [" & CStr(vadrsNotify!ABS_NUM) & "]" & _
                                " which started on " & Trim(vadrsNotify!ABS_ST_TS) & "." & vbCrLf & vbCrLf & _
                                "Follow-up is scheduled for " & CStr(vadrsNotify!FOLW_UP_DT) & "." & vbCrLf & vbCrLf & _
                                "(This is an automated message. Please do not reply to this message.)"

                        Case 3  'Return to Work
                            .Subject = .Subject & " (Return to Work Notification)..."
                            .TextBody = "This is to remind you that " & _
                                " employee " & Trim(vadrsNotify!EMP_NAME) & " [" & CStr(vadrsNotify!EMP_NUM) & "]" & _
                                " is scheduled to return to work on " & Trim(vadrsNotify!RTN_TO_WRK_TS) & _
                                " from absence " & Trim(vadrsNotify!ABS_DESC) & " [" & CStr(vadrsNotify!ABS_NUM) & "]" & _
                                " which started on " & Trim(vadrsNotify!ABS_ST_TS) & "." & vbCrLf & vbCrLf & _
                                "(This is an automated message. Please do not reply to this message.)"
                    End Select

                    Select Case (lsDB2)
                        Case "DB2P", "DB2PLUW"
                            '
                        Case Else
                            .TextBody = .TextBody & vbCrLf & vbCrLf & _
                                "THIS IS A TEST MESSAGE, ABSENCE AND NOTIFICATION PREFERENCE WERE CREATED/MODIFIED DURING TEST." & vbCrLf & vbCrLf & _
                                "PLEASE DISREGARD THIS MESSAGE."
                    End Select

                    On Error Resume Next
                    .Send
                    If (Err.Number <> 0) Then
                        Err.Clear
                    End If

                    On Error GoTo SendMail_EH
                End With

                Set lobjCDOMsg = Nothing

                .MoveNext
            Next llIndex
        End With
    End If

    Set lobjCnMgr = Nothing

    Set lobjCDOConfig = Nothing
    Set ladFlds = Nothing
    Set lobjCDOMsg = Nothing

    Exit Sub
SendMail_EH:
    Set lobjCnMgr = Nothing

    Set lobjCDOConfig = Nothing
    Set ladFlds = Nothing
    Set lobjCDOMsg = Nothing

    TrsRaiseError MODULENAME _
        , "SendMail" _
        , , , , _
        , True

End Sub
'    If Not (GetEAMNotifyPreference Is Nothing) Then
'        With GetEAMNotifyPreference
'            If Not (.BOF And .EOF) Then
'                .MoveFirst
'                .Find "RECORD_TYPE = 'CNTCT'"
'                If (.EOF) Then
'                    .AddNew
'                    !RG_VALUE_CD = ""
'                    !USER_NUM_AS_STR = lsUser
'                    !ELAPSED_DAYS = 0
'                    !NOTIFY_IND = "N"
'                    !RECORD_TYPE = "CNTCT"
'                    .Update
'                End If
'
'                .MoveFirst
'                .Find "RECORD_TYPE = 'RTW'"
'                If (.EOF) Then
'                    .AddNew
'                    !RG_VALUE_CD = ""
'                    !USER_NUM_AS_STR = lsUser
'                    !ELAPSED_DAYS = 0
'                    !NOTIFY_IND = "N"
'                    !RECORD_TYPE = "RTW"
'                    .Update
'                End If
'            Else
'                .AddNew
'                !RG_VALUE_CD = ""
'                !USER_NUM_AS_STR = lsUser
'                !ELAPSED_DAYS = 0
'                !NOTIFY_IND = "N"
'                !RECORD_TYPE = "CNTCT"
'                .Update
'
'                .AddNew
'                !RG_VALUE_CD = ""
'                !USER_NUM_AS_STR = lsUser
'                !ELAPSED_DAYS = 0
'                !NOTIFY_IND = "N"
'                !RECORD_TYPE = "RTW"
'                .Update
'            End If
'        End With
'    Else
'        Set GetEAMNotifyPreference = New ADODB.Recordset
'        With GetEAMNotifyPreference
'            .AddNew
'            !RG_VALUE_CD = ""
'            !USER_NUM_AS_STR = lsUser
'            !ELAPSED_DAYS = 0
'            !NOTIFY_IND = "N"
'            !RECORD_TYPE = "CNTCT"
'            .Update
'
'            .AddNew
'            !RG_VALUE_CD = ""
'            !USER_NUM_AS_STR = lsUser
'            !ELAPSED_DAYS = 0
'            !NOTIFY_IND = "N"
'            !RECORD_TYPE = "RTW"
'            .Update
'        End With
'    End If

