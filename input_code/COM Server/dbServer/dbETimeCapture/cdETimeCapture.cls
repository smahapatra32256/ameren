VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cdETimeCapture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const m_strCMODULENAME As String = "dbETimeCapture.cdETimeCapture"
Private Const m_DBSCRIPTSROOT As String = "C:\TRIS SOLUTION\dbScripts\TRIS_db2"
Private Const m_DBSCRIPTSPATH As String = "\dbETimeCapture\cdETimeCapture"
'//Private Const m_lngXMLChunkSize As Long = 102400 '//100K

Private m_ResourceMgr As utilADOHelper.cuResourceMgr

'Public Function GetStoredBRVs(ByVal vlEmployeeNumber As Long _
'                    , ByVal vdtStartDate As Date _
'                    , ByVal vdtEndDate As Date _
'                    ) As ADODB.Recordset
'//phase 3
'Public Function GetBRVLogAsNewErrorset(ByVal vdblBRVHdrId As Double) As ADODB.Recordset '//in time collection format
'On Error GoTo eh
'    Dim SQL As String
'    Dim oADOHelper As utilADOHelper.cuADOHelper
'    Dim rs As ADODB.Recordset
'
'    'get the SQL
'    SQL = m_ResourceMgr.GetStringResource(m_DBSCRIPTSROOT, m_DBSCRIPTSPATH, "GetBRVLogAsNewErrorset.sql")
'
'    'create parameters
'    Set oADOHelper = New utilADOHelper.cuADOHelper
'    oADOHelper.ADO.CreateNumericParameter "vdblBRVHdrId", adDouble, 20, 0, vdblBRVHdrId
'
'    'execute sql
'    Set rs = oADOHelper.ExecuteSelect(SQL, adOpenStatic, adLockBatchOptimistic)
'
'    Set oADOHelper = Nothing
'    Set GetBRVLogAsNewErrorset = rs
'    Exit Function
'eh:
'    Set oADOHelper = Nothing
'    Set rs = Nothing
'
'    TrsRaiseError m_strCMODULENAME, "GetBRVLogAsNewErrorset", , , , DescArgsFrom(SQL, CStr(vdblBRVHdrId)), True
'End Function

'//phase 2
Public Function GetValidationResultsAsNewErrorset(ByVal vdblTimesheetId As Double) As ADODB.Recordset '//in time collection format
On Error GoTo eh
    Dim SQL As String
    Dim oADOHelper As utilADOHelper.cuADOHelper
    Dim rs As ADODB.Recordset
    
    'get the SQL
    SQL = m_ResourceMgr.GetStringResource(m_DBSCRIPTSROOT, m_DBSCRIPTSPATH, "GetValidationResultsAsNewErrorset.sql")
    
    'create parameters
    Set oADOHelper = New utilADOHelper.cuADOHelper
    oADOHelper.ADO.CreateNumericParameter "vdblTimesheetId", adDouble, 20, 0, vdblTimesheetId
    oADOHelper.ADO.CreateNumericParameter "vdblTimesheetId2", adDouble, 20, 0, vdblTimesheetId
    oADOHelper.ADO.CreateNumericParameter "vdblTimesheetId3", adDouble, 20, 0, vdblTimesheetId
    oADOHelper.ADO.CreateNumericParameter "vdblTimesheetId4", adDouble, 20, 0, vdblTimesheetId
    
    'execute sql
    Set rs = oADOHelper.ExecuteSelect(SQL, adOpenStatic, adLockBatchOptimistic)

    Set oADOHelper = Nothing
    Set GetValidationResultsAsNewErrorset = rs
    Exit Function
eh:
    Set oADOHelper = Nothing
    Set rs = Nothing
    
    TrsRaiseError m_strCMODULENAME, "GetValidationResultsAsNewErrorset", , , , DescArgsFrom(SQL, CStr(vdblTimesheetId)), True
End Function

'Public Function GetTimeCaptureSelectionForTimekeeper(ByVal vlngTimekeeper As Long, ByVal vdt_StartDate As Date, ByVal vdt_EndDate As Date) As ADODB.Recordset
'On Error GoTo eh
'    Dim SQL As String
'    Dim oADOHelper As utilADOHelper.cuADOHelper
'    Dim rs As ADODB.Recordset
'
'    'get the SQL
'    SQL = m_ResourceMgr.GetStringResource(m_DBSCRIPTSROOT, m_DBSCRIPTSPATH, "GetTimeCaptureSelectionForTimekeeper.sql")
'
'    'create parameters
'    Set oADOHelper = New utilADOHelper.cuADOHelper
'
'    oADOHelper.ADO.CreateNumericParameter "StartDate", adDate, 0, 0, vdt_StartDate
'    oADOHelper.ADO.CreateNumericParameter "EndDate", adDate, 0, 0, vdt_EndDate
'    oADOHelper.ADO.CreateNumericParameter "StartDate2", adDate, 0, 0, vdt_StartDate
'    oADOHelper.ADO.CreateNumericParameter "EndDate2", adDate, 0, 0, vdt_EndDate
'    oADOHelper.ADO.CreateNumericParameter "Timekeeper", adInteger, 0, 0, vlngTimekeeper
'
'    'execute sql
'    Set rs = oADOHelper.ExecuteSelect(SQL, adOpenStatic, adLockBatchOptimistic)
'
'    Set oADOHelper = Nothing
'    Set GetTimeCaptureSelectionForTimekeeper = rs
'    Exit Function
'eh:
'    Set oADOHelper = Nothing
'    Set rs = Nothing
'
'    TrsRaiseError m_strCMODULENAME, "GetTimeCaptureSelectionForTimekeeper", , , , DescArgsFrom(SQL, CStr(vlngTimekeeper), CStr(vdt_StartDate), CStr(vdt_EndDate)), True
'End Function

'Public Function GetTimeCaptureSelectionPreview(ByVal v_dblRowId As Double, ByRef r_strXML As String, ByRef r_strXSLT As String)
'On Error GoTo eh
'    Dim SQL As String
'    Dim oADOHelper As utilADOHelper.cuADOHelper
'    Dim rs As ADODB.Recordset
'
'    'get the SQL
'    SQL = m_ResourceMgr.GetStringResource(m_DBSCRIPTSROOT, m_DBSCRIPTSPATH, "GetTimeCaptureSelectionPreview.sql")
'
'    'create parameters
'    Set oADOHelper = New utilADOHelper.cuADOHelper
'
'    oADOHelper.ADO.CreateNumericParameter "v_dblRowId", adDecimal, 20, 0, v_dblRowId
'
'    'execute sql
'    Set rs = oADOHelper.ExecuteSelect(SQL, adOpenStatic, adLockBatchOptimistic)
'
'    If RecordsExist(rs) Then
'        rs.MoveFirst
'        '//r_strXML = rs.Fields("XML_Data").GetChunk(m_lngXMLChunkSize)
'        r_strXML = rs.Fields("XML_Data").Value
'    End If
'
'    '//r_strXSLT = m_ResourceMgr.GetStringResource(m_DBSCRIPTSROOT, m_DBSCRIPTSPATH, "GetTimeCaptureSelectionPreviewXSLT.txt")
'    Dim fs As FileSystemObject
'    Set fs = New FileSystemObject
'
'    Dim ts As TextStream
'    Set ts = fs.OpenTextFile("C:\TRISSandboxes\TRISSandbox\TRISMaster\Tester\Ameren.TRIS.Server.Business.eTimeCaptureTester\eTimeCapture.xslt")
'    r_strXSLT = ts.ReadAll
'    Set ts = Nothing
'    Set fs = Nothing
'
'    Set oADOHelper = Nothing
'    Set rs = Nothing
'    Exit Function
'eh:
'    Set oADOHelper = Nothing
'    Set rs = Nothing
'
'    TrsRaiseError m_strCMODULENAME, "GetTimeCaptureSelectionPreview", , , , DescArgsFrom(SQL, CStr(v_dblRowId)), True
'End Function

'obsolete
'Public Function GetTimeCaptureSelectionDetails(ByVal v_dblRowId As Double, ByRef r_rsEmployees As ADODB.Recordset, ByRef r_rsDates As ADODB.Recordset)
'On Error GoTo eh
'
'    Dim r_lngLastModEmpNum As Long
'    Dim r_strLastModTimestamp As String
'
'    GetTimeCaptureSelectionDetails = GetTimeCaptureSelectionDetails2(v_dblRowId, r_rsEmployees, r_rsDates, r_lngLastModEmpNum, r_strLastModTimestamp)
'eh:
'
'    TrsRaiseError m_strCMODULENAME, "GetTimeCaptureSelectionPreview", , , , , True
'End Function

Public Function GetTimeCaptureSelectionDetails(ByVal v_dblTimesheetId As Double, ByRef r_rsEmployees As ADODB.Recordset, ByRef r_rsDates As ADODB.Recordset, _
                                                ByRef r_lngLastModEmpNum As Long, ByRef r_strLastModTimestamp As String)
On Error GoTo eh
    Dim SQL1 As String, SQL2 As String, SQL3 As String
    Dim oADOHelper As utilADOHelper.cuADOHelper
    Dim rs As ADODB.Recordset
    Dim l_rsLastMod As ADODB.Recordset
    
    'get the SQL for SelectedEmployees
    SQL1 = m_ResourceMgr.GetStringResource(m_DBSCRIPTSROOT, m_DBSCRIPTSPATH, "GetTimeCaptureSelectionDetailsEmployees.sql")
    
    'create parameters
    Set oADOHelper = New utilADOHelper.cuADOHelper

    oADOHelper.ADO.CreateNumericParameter "v_dblTimesheetId", adDecimal, 20, 0, v_dblTimesheetId
    
    'execute sql
    Set r_rsEmployees = oADOHelper.ExecuteSelect(SQL1, adOpenStatic, adLockBatchOptimistic)


    'get the SQL for SelectedDates
    SQL2 = m_ResourceMgr.GetStringResource(m_DBSCRIPTSROOT, m_DBSCRIPTSPATH, "GetTimeCaptureSelectionDetailsDates.sql")
    
    'create parameters
    Set oADOHelper = New utilADOHelper.cuADOHelper

    oADOHelper.ADO.CreateNumericParameter "v_dblTimesheetId", adDecimal, 20, 0, v_dblTimesheetId
    
    'execute sql
    Set r_rsDates = oADOHelper.ExecuteSelect(SQL2, adOpenStatic, adLockBatchOptimistic)
    
    'for optimistic locking
    'get the SQL for last mod empnum and last mod timestamp
    SQL3 = m_ResourceMgr.GetStringResource(m_DBSCRIPTSROOT, m_DBSCRIPTSPATH, "GetTimeCaptureSelectionDetailsLastMod.sql")
    
    'create parameters
    Set oADOHelper = New utilADOHelper.cuADOHelper

    oADOHelper.ADO.CreateNumericParameter "v_dblTimesheetId", adDecimal, 20, 0, v_dblTimesheetId
    
    'execute sql
    Set l_rsLastMod = oADOHelper.ExecuteSelect(SQL3, adOpenStatic, adLockBatchOptimistic)
    If RecordsExist(l_rsLastMod) Then
        l_rsLastMod.MoveFirst
        r_lngLastModEmpNum = l_rsLastMod.Fields("LAST_MOD_EMP_NUM").Value
        r_strLastModTimestamp = l_rsLastMod.Fields("LAST_MOD_TIMESTAMP").Value
    Else
        r_lngLastModEmpNum = -1
        r_strLastModTimestamp = "1/1/100"
    End If
    
    Set oADOHelper = Nothing
    Exit Function
eh:
    Set oADOHelper = Nothing
    Set r_rsEmployees = Nothing
    Set r_rsDates = Nothing
    
    TrsRaiseError m_strCMODULENAME, "GetTimeCaptureSelectionDetails", , , , DescArgsFrom(SQL1, SQL2, CStr(v_dblTimesheetId)), True
End Function

'Public Function GetETimeSheetEntriesForEdit(ByVal vlRowId As Long) As ADODB.Recordset '//in time collection format
'On Error GoTo eh
'    Dim SQL As String
'    Dim oADOHelper As utilADOHelper.cuADOHelper
'    Dim rs As ADODB.Recordset
'
'    'get the SQL
'    SQL = m_ResourceMgr.GetStringResource(m_DBSCRIPTSROOT, m_DBSCRIPTSPATH, "GetBRVLogAsNewErrorset.sql")
'
'    'create parameters
'    Set oADOHelper = New utilADOHelper.cuADOHelper
'    oADOHelper.ADO.CreateNumericParameter "vlBRVHdrId", adInteger, 4, 0, vlBRVHdrId
'
'    'execute sql
'    Set rs = oADOHelper.ExecuteSelect(SQL, adOpenStatic, adLockBatchOptimistic)
'
'    Set oADOHelper = Nothing
'    Set GetBRVLogAsNewErrorset = rs
'    Exit Function
'eh:
'    Set oADOHelper = Nothing
'    Set rs = Nothing
'
'    TrsRaiseError m_strCMODULENAME, "GetBRVLogAsNewErrorset", , , , DescArgsFrom(SQL, CStr(vlBRVHdrId)), True
'End Function

Private Sub Class_Initialize()
    Set m_ResourceMgr = New utilADOHelper.cuResourceMgr
End Sub

Private Sub Class_Terminate()
    Set m_ResourceMgr = Nothing
End Sub
