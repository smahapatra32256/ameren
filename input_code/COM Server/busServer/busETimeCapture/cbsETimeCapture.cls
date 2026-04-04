VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "cbsETimeCapture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const m_strCMODULENAME As String = "busETimeCapture.cbsETimeCapture"

'//phase 3, errors from server import, for the client to see in Time Collection By Timesheet
'Public Function GetBRVLogAsNewErrorset(ByVal vlBRVHdrId As Long) As ADODB.Recordset '//in time collection format
'On Error GoTo eh
'    Dim db As dbETimeCapture.cdETimeCapture
'    Dim rs As ADODB.Recordset
'
'    Set db = CreateObject("dbETimeCapture.cdETimeCapture")
'    Set rs = db.GetBRVLogAsNewErrorset(vlBRVHdrId)
'    Set GetBRVLogAsNewErrorset = rs
'    Exit Function
'eh:
'    Set rs = Nothing
'
'    TrsRaiseError m_strCMODULENAME, "GetBRVLogAsNewErrorset", , , , DescArgsFrom(CStr(vlBRVHdrId)), True
'End Function

'//phase 2, errors from mobile device submit, for the client to see in Time Collection By Timesheet
Public Function GetValidationResultsAsNewErrorset(ByVal vlTimesheetId As Double) As ADODB.Recordset '//in time collection format
On Error GoTo eh
    Dim db As dbETimeCapture.cdETimeCapture
    Dim rs As ADODB.Recordset
    
    Set db = CreateObject("dbETimeCapture.cdETimeCapture")
    Set rs = db.GetValidationResultsAsNewErrorset(vlTimesheetId)
    Set GetValidationResultsAsNewErrorset = rs
    Exit Function
eh:
    Set rs = Nothing
    
    TrsRaiseError m_strCMODULENAME, "GetValidationResultsAsNewErrorset", , , , DescArgsFrom(CStr(vlTimesheetId)), True
End Function

'//obsolete
Public Function GetTimeCaptureAndConvertToActualTimeForEdit(ByVal vdblTimesheetId As Double, ByRef rsDetail As ADODB.Recordset, ByRef rsSummary As ADODB.Recordset) As Boolean '//in time collection format
On Error GoTo eh
    Dim tc As Ameren_TRIS_Server_COMIntegration2.TimeCapture
    Dim rc As Boolean
    
    Set tc = New Ameren_TRIS_Server_COMIntegration2.TimeCapture
    rc = tc.GetTimeCaptureAndConvertToActualTimeForEdit(vdblTimesheetId, rsDetail, rsSummary)
    
    Set tc = Nothing
    Exit Function
eh:
    Set tc = Nothing
    TrsRaiseError m_strCMODULENAME, "GetTimeCaptureAndConvertToActualTimeForEdit", , , , DescArgsFrom(CStr(vdblTimesheetId)), True
End Function

Public Function GetTimeCaptureAndConvertToActualTimeForEdit2(ByVal vdblTimesheetId As Double, ByRef rsDetail As ADODB.Recordset, ByRef rsSummary As ADODB.Recordset, ByVal vARREmployees As Variant) As Boolean  '//in time collection format
On Error GoTo eh
    Dim tc As Ameren_TRIS_Server_COMIntegration2.TimeCapture
    Dim rc As Boolean
    
    Set tc = New Ameren_TRIS_Server_COMIntegration2.TimeCapture
    rc = tc.GetTimeCaptureAndConvertToActualTimeForEdit2(vdblTimesheetId, rsDetail, rsSummary, vARREmployees)
    
    Set tc = Nothing
    Exit Function
eh:
    Set tc = Nothing
    TrsRaiseError m_strCMODULENAME, "GetTimeCaptureAndConvertToActualTimeForEdit2", , , , DescArgsFrom(CStr(vdblTimesheetId)), True
End Function

'//obsolete
Public Function GetTimeCaptureAndConvertToOtherPayForEdit(ByVal vdblTimesheetId As Double) As ADODB.Recordset '//in time collection format
On Error GoTo eh
    Dim tc As Ameren_TRIS_Server_COMIntegration2.TimeCapture
    Dim rs As ADODB.Recordset
    
    Set tc = New Ameren_TRIS_Server_COMIntegration2.TimeCapture
    Set rs = tc.GetTimeCaptureAndConvertToOtherPayForEdit(vdblTimesheetId)
    
    Set tc = Nothing
    Set GetTimeCaptureAndConvertToOtherPayForEdit = rs
    Exit Function
eh:
    Set tc = Nothing
    TrsRaiseError m_strCMODULENAME, "GetTimeCaptureAndConvertToOtherPayForEdit", , , , DescArgsFrom(CStr(vdblTimesheetId)), True
End Function

Public Function GetTimeCaptureAndConvertToOtherPayForEdit2(ByVal vdblTimesheetId As Double, ByVal vARREmployees As Variant) As ADODB.Recordset '//in time collection format
On Error GoTo eh
    Dim tc As Ameren_TRIS_Server_COMIntegration2.TimeCapture
    Dim rs As ADODB.Recordset
    
    Set tc = New Ameren_TRIS_Server_COMIntegration2.TimeCapture
    Set rs = tc.GetTimeCaptureAndConvertToOtherPayForEdit2(vdblTimesheetId, vARREmployees)
    
    Set tc = Nothing
    Set GetTimeCaptureAndConvertToOtherPayForEdit2 = rs
    Exit Function
eh:
    Set tc = Nothing
    TrsRaiseError m_strCMODULENAME, "GetTimeCaptureAndConvertToOtherPayForEdit2", , , , DescArgsFrom(CStr(vdblTimesheetId)), True
End Function

'Public Function GetTimeCaptureSelectionForTimekeeper(ByVal vlngTimekeeper As Long, ByVal vdt_StartDate As Date, ByVal vdt_EndDate As Date) As ADODB.Recordset
'On Error GoTo eh
'    Dim db As dbETimeCapture.cdETimeCapture
'    Dim rs As ADODB.Recordset
'
'    Set db = CreateObject("dbETimeCapture.cdETimeCapture")
'    Set rs = db.GetTimeCaptureSelectionForTimekeeper(vlngTimekeeper, vdt_StartDate, vdt_EndDate)
'    Set GetTimeCaptureSelectionForTimekeeper = rs
'    Set db = Nothing
'
'    Exit Function
'eh:
'    Set db = Nothing
'    Set rs = Nothing
'
'    TrsRaiseError m_strCMODULENAME, "GetTimeCaptureSelectionForTimekeeper", , , , DescArgsFrom(CStr(vlngTimekeeper), CStr(vdt_StartDate), CStr(vdt_EndDate)), True
'End Function

Public Function GetTimeCaptureSelectionAsHTML(ByVal vlngEmpNum As Long, ByVal vdt_StartDate As Date, ByVal vdt_EndDate As Date) As String
On Error GoTo eh
    Dim db As dbETimeCapture.cdETimeCapture
    Dim html As String
    
    Dim o As Ameren_TRIS_Server_COMIntegration2.TimeCapture
    
    Set o = New Ameren_TRIS_Server_COMIntegration2.TimeCapture
    
    html = o.GetTimeCaptureSelectionAsHTML(vlngEmpNum, vdt_StartDate, vdt_EndDate)
    
    GetTimeCaptureSelectionAsHTML = html
    
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "GetTimeCaptureSelectionAsHTML", , , , DescArgsFrom(CStr(vlngEmpNum), CStr(vdt_StartDate), CStr(vdt_EndDate)), True
End Function

Public Function GetTimeCaptureSelectionForTimekeeperAsHTML(ByVal vlngTimekeeper As Long, ByVal vdt_StartDate As Date, ByVal vdt_EndDate As Date) As String
On Error GoTo eh
    Dim db As dbETimeCapture.cdETimeCapture
    Dim html As String
    
    Dim o As Ameren_TRIS_Server_COMIntegration2.TimeCapture
    
    Set o = New Ameren_TRIS_Server_COMIntegration2.TimeCapture
    
    html = o.GetTimeCaptureSelectionForTimekeeperAsHTML(vlngTimekeeper, vdt_StartDate, vdt_EndDate)
    
    GetTimeCaptureSelectionForTimekeeperAsHTML = html
    
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "GetTimeCaptureSelectionForTimekeeperAsHTML", , , , DescArgsFrom(CStr(vlngTimekeeper), CStr(vdt_StartDate), CStr(vdt_EndDate)), True
End Function

'Public Function GetTimeCaptureSelectionPreview(ByVal v_dblRowId As Double, ByRef r_strXML As String, ByRef r_strXSLT As String)
'On Error GoTo eh
'
'    Exit Function
'eh:
'    'Set db = Nothing
'
'    TrsRaiseError m_strCMODULENAME, "GetTimeCaptureSelectionPreview", , , , DescArgsFrom(CStr(v_dblRowId)), True
'End Function

'obsolete
'Public Function GetTimeCaptureSelectionPreview2(ByVal v_dblRowId As Double, ByRef r_strXML As String, ByRef r_strXSLT As String, ByRef r_strHTML As String)
'On Error GoTo eh
'
'    Dim o As Ameren_TRIS_Server_COMIntegration2.TimeCapture
'
'    Set o = New Ameren_TRIS_Server_COMIntegration2.TimeCapture
'
'    o.GetTimeCaptureSelectionPreview v_dblRowId, r_strXML, r_strXSLT, r_strHTML
'    Exit Function
'eh:
'    TrsRaiseError m_strCMODULENAME, "GetTimeCaptureSelectionPreview2", , , , DescArgsFrom(CStr(v_dblRowId)), True
'End Function

'obsolete
'Public Function GetTimeCaptureSelectionPreviewBrvs(ByVal v_dblTimesheetId As Double, ByRef r_strXML As String, ByRef r_strXSLT As String, ByRef r_strHTML As String)
'On Error GoTo eh
'    Dim o As Ameren_TRIS_Server_COMIntegration2.TimeCapture
'
'    Set o = New Ameren_TRIS_Server_COMIntegration2.TimeCapture
'
'    o.GetTimeCaptureSelectionPreviewBrvs v_dblTimesheetId, r_strXML, r_strXSLT, r_strHTML
'    Exit Function
'eh:
'    TrsRaiseError m_strCMODULENAME, "GetTimeCaptureSelectionPreviewBrvs", , , , DescArgsFrom(CStr(v_dblTimesheetId)), True
'End Function

Public Function GetTimesheetViewAsHTML(ByVal v_dblTimesheetId As Double) As String
On Error GoTo eh
    Dim o As Ameren_TRIS_Server_COMIntegration2.TimeCapture
    
    Set o = New Ameren_TRIS_Server_COMIntegration2.TimeCapture
    
    GetTimesheetViewAsHTML = o.GetTimesheetViewAsHTML(v_dblTimesheetId)
    Set o = Nothing
    Exit Function
eh:
    Set o = Nothing
    TrsRaiseError m_strCMODULENAME, "GetTimesheetViewAsHTML", , , , DescArgsFrom(CStr(v_dblTimesheetId)), True
End Function

'obsolete
Public Function GetTimesheetViewExtractedFromActualTimeAsHTML(ByVal v_datChargeDate As Date, ByVal v_arrEmployees As Variant) As String
On Error GoTo eh
    Dim o As Ameren_TRIS_Server_COMIntegration2.TimeCapture
    
    Set o = New Ameren_TRIS_Server_COMIntegration2.TimeCapture
    
    GetTimesheetViewExtractedFromActualTimeAsHTML = o.GetTimesheetViewExtractedFromActualTimeAsHTML(v_datChargeDate, v_arrEmployees)
    Set o = Nothing
    Exit Function
eh:
    Set o = Nothing
    Dim EmpList As String
    Dim i As Integer
    If Not IsEmpty(v_arrEmployees) Then
        If IsArray(v_arrEmployees) Then
            For i = LBound(v_arrEmployees) To UBound(v_arrEmployees)
                EmpList = EmpList + CStr(v_arrEmployees(i))
                If i < UBound(v_arrEmployees) Then
                    EmpList = EmpList + ", "
                End If
            Next i
            '//EmpList = Join(v_arrEmployees, ",")
        End If
    End If
    
    TrsRaiseError m_strCMODULENAME, "GetTimesheetViewExtractedFromActualTimeAsHTML", , , , DescArgsFrom(CStr(v_datChargeDate), EmpList), True
End Function

Public Function GetTimesheetViewsExtractedFromActualTimeAsHTML(ByVal v_datChargeDate As Date, ByVal v_arrEmployees As Variant, _
                                        ByRef TimesheetView As String, ByRef TimesheetEditView As String) As Boolean
On Error GoTo eh
    Dim o As Ameren_TRIS_Server_COMIntegration2.TimeCapture
    
    Set o = New Ameren_TRIS_Server_COMIntegration2.TimeCapture

    GetTimesheetViewsExtractedFromActualTimeAsHTML = o.GetTimesheetViewsExtractedFromActualTimeAsHTML(v_datChargeDate, v_arrEmployees, TimesheetView, TimesheetEditView)

    Set o = Nothing
    Exit Function
eh:
    Set o = Nothing
    Dim EmpList As String
    Dim i As Integer
    If Not IsEmpty(v_arrEmployees) Then
        If IsArray(v_arrEmployees) Then
            For i = LBound(v_arrEmployees) To UBound(v_arrEmployees)
                EmpList = EmpList + CStr(v_arrEmployees(i))
                If i < UBound(v_arrEmployees) Then
                    EmpList = EmpList + ", "
                End If
            Next i
            '//EmpList = Join(v_arrEmployees, ",")
        End If
    End If
    
    TrsRaiseError m_strCMODULENAME, "GetTimesheetViewsExtractedFromActualTimeAsHTML", , , , DescArgsFrom(CStr(v_datChargeDate), EmpList), True
End Function

Public Function GetTimesheetEntryViewAsHTML(ByVal v_dblTimesheetId As Double) As String
On Error GoTo eh
    Dim o As Ameren_TRIS_Server_COMIntegration2.TimeCapture
    
    Set o = New Ameren_TRIS_Server_COMIntegration2.TimeCapture
    
    GetTimesheetEntryViewAsHTML = o.GetTimesheetEntryViewAsHTML(v_dblTimesheetId)
    Set o = Nothing
    Exit Function
eh:
    Set o = Nothing
    TrsRaiseError m_strCMODULENAME, "GetTimesheetEntryViewAsHTML", , , , DescArgsFrom(CStr(v_dblTimesheetId)), True
End Function

Public Function GetTimesheetApprovalViewAsHTML(ByVal v_dblTimesheetId As Double) As String
On Error GoTo eh
    Dim o As Ameren_TRIS_Server_COMIntegration2.TimeCapture
    
    Set o = New Ameren_TRIS_Server_COMIntegration2.TimeCapture
    
    GetTimesheetApprovalViewAsHTML = o.GetTimesheetApprovalViewAsHTML(v_dblTimesheetId)
    Set o = Nothing
    Exit Function
eh:
    Set o = Nothing
    TrsRaiseError m_strCMODULENAME, "GetTimesheetApprovalViewAsHTML", , , , DescArgsFrom(CStr(v_dblTimesheetId)), True
End Function

Public Function GetTimesheetViewsAsHTML(ByVal v_dblTimesheetId As Double, ByRef TimesheetView As String, ByRef TimesheetEntryView As String, _
                                        ByRef TimesheetApprovalView As String, ByRef HistoryView As String) As Boolean
On Error GoTo eh
    Dim o As Ameren_TRIS_Server_COMIntegration2.TimeCapture
    
    Set o = New Ameren_TRIS_Server_COMIntegration2.TimeCapture
    
    GetTimesheetViewsAsHTML = o.GetTimesheetViewsAsHTML(v_dblTimesheetId, TimesheetView, TimesheetEntryView, TimesheetApprovalView, HistoryView)
    Set o = Nothing
    Exit Function
eh:
    Set o = Nothing
    TrsRaiseError m_strCMODULENAME, "GetTimesheetViewsAsHTML", , , , DescArgsFrom(CStr(v_dblTimesheetId)), True
End Function

Public Function GetTimesheetEntryBrvsViewAsHTML(ByVal v_dblTimesheetId As Double) As String
On Error GoTo eh
    Dim o As Ameren_TRIS_Server_COMIntegration2.TimeCapture
    
    Set o = New Ameren_TRIS_Server_COMIntegration2.TimeCapture
    
    GetTimesheetEntryBrvsViewAsHTML = o.GetTimesheetEntryBrvsViewAsHTML(v_dblTimesheetId)
    Set o = Nothing
    Exit Function
eh:
    Set o = Nothing
    TrsRaiseError m_strCMODULENAME, "GetTimesheetEntryBrvsViewAsHTML", , , , DescArgsFrom(CStr(v_dblTimesheetId)), True
End Function

Public Function GetTimesheetHistoryViewAsHTML(ByVal v_dblTimesheetId As Double) As String
On Error GoTo eh
    Dim o As Ameren_TRIS_Server_COMIntegration2.TimeCapture
    
    Set o = New Ameren_TRIS_Server_COMIntegration2.TimeCapture
    
    GetTimesheetHistoryViewAsHTML = o.GetTimesheetHistoryViewAsHTML(v_dblTimesheetId)
    Set o = Nothing
    Exit Function
eh:
    Set o = Nothing
    TrsRaiseError m_strCMODULENAME, "GetTimesheetHistoryViewAsHTML", , , , DescArgsFrom(CStr(v_dblTimesheetId)), True
End Function


'obsolete
'Public Function GetTimeCaptureSelectionDetails(ByVal v_dblRowId As Double, ByRef r_rsEmployees As ADODB.Recordset, ByRef r_rsDates As ADODB.Recordset)
'On Error GoTo eh
'    Dim db As dbETimeCapture.cdETimeCapture
'    Dim rs As ADODB.Recordset
'
'    Set db = CreateObject("dbETimeCapture.cdETimeCapture")
'    db.GetTimeCaptureSelectionDetails v_dblRowId, r_rsEmployees, r_rsDates
'
'    Set db = Nothing
'
'    Exit Function
'eh:
'    Set db = Nothing
'    Set rs = Nothing
'
'    TrsRaiseError m_strCMODULENAME, "GetTimeCaptureSelectionDetails", , , , DescArgsFrom(CStr(v_dblRowId)), True
'End Function

Public Function GetTimeCaptureSelectionDetails(ByVal vdblTimesheetId As Double, ByRef r_rsEmployees As ADODB.Recordset, ByRef r_rsDates As ADODB.Recordset, _
                                                ByRef r_lngLastModEmpNum As Long, ByRef r_strLastModTimestamp As String)
On Error GoTo eh
    Dim db As dbETimeCapture.cdETimeCapture
    Dim rs As ADODB.Recordset
    
    Set db = CreateObject("dbETimeCapture.cdETimeCapture")
    db.GetTimeCaptureSelectionDetails vdblTimesheetId, r_rsEmployees, r_rsDates, r_lngLastModEmpNum, r_strLastModTimestamp

    Set db = Nothing
    
    Exit Function
eh:
    Set db = Nothing
    Set rs = Nothing
    
    TrsRaiseError m_strCMODULENAME, "GetTimeCaptureSelectionDetails", , , , DescArgsFrom(CStr(vdblTimesheetId)), True
End Function

'//let COM Integration manage the transaction
Public Function UpdateTimeCaptureTimekeeper(ByVal vdblTimesheetId As Double, ByVal vlngLastModEmpNum As Long, ByVal vstrLastModTimestamp As String, _
                                            ByVal vintNewTimekeeper As Long) As Integer  '//in time collection format
On Error GoTo eh
    Dim tc As Ameren_TRIS_Server_COMIntegration2.TimeCapture
    
    Set tc = New Ameren_TRIS_Server_COMIntegration2.TimeCapture
    UpdateTimeCaptureTimekeeper = tc.UpdateTimeCaptureTimekeeper(vdblTimesheetId, vlngLastModEmpNum, vstrLastModTimestamp, vintNewTimekeeper)
    
    Set tc = Nothing
    Exit Function
eh:
    UpdateTimeCaptureTimekeeper = -1
    
    TrsRaiseError m_strCMODULENAME, "UpdateTimeCaptureTimekeeper", , , , DescArgsFrom(CStr(vdblTimesheetId), CStr(vlngLastModEmpNum), vstrLastModTimestamp, CStr(vintNewTimekeeper)), True
End Function

'//return data from TTR_E_TIME_CAPTURE_XML formatted to look like it came from GetActualTimeDetailAndSummary
'// call cominterop to call ClientProcessor
'Public Function GetETimeSheetEntriesForEdit(ByVal vlRowId As Long) As ADODB.Recordset '//in time collection format
'On Error GoTo eh
'    Dim db As dbETimeCapture.cdETimeCapture
'    Dim rs As ADODB.Recordset
'
'    Set db = CreateObject("dbETimeCapture.cdETimeCapture")
'    Set rs = db.GetETimeSheetEntriesForEdit(vlRowId)
'    Set GetETimeSheetEntriesForEdit = rs
'    Exit Function
'eh:
'    Set rs = Nothing
'
'    TrsRaiseError m_strCMODULENAME, "CreateeTimeSheetEntriesForEdit", , , , DescArgsFrom(CStr(vlRowId)), True
'End Function
