VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "cdETimeCaptureTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const m_strCMODULENAME As String = "dbETimeCapture.cdETimeCaptureTrans"
Private Const m_DBSCRIPTSROOT As String = "C:\TRIS SOLUTION\dbScripts\TRIS_db2"
Private Const m_DBSCRIPTSPATH As String = "\dbETimeCapture\cdETimeCaptureTrans"
Private m_ResourceMgr As utilADOHelper.cuResourceMgr


Private Sub Class_Initialize()
    Set m_ResourceMgr = New utilADOHelper.cuResourceMgr
End Sub

Private Sub Class_Terminate()
    Set m_ResourceMgr = Nothing
End Sub

Public Function UpdateTimeCaptureStatus(ByVal vdblTimesheetId As Double, ByVal vlngLastModEmpNum As Long, ByVal vstrLastModTimestamp As String, _
                                        ByVal vintNewStatus As Integer, ByVal vlngNewLastModEmpNum As Long) As Integer '//in time collection format
On Error GoTo eh
    Dim SQL As String
    Dim oADOHelper As utilADOHelper.cuADOHelperTrans
    Dim rc As Integer
    
    'get the SQL
    SQL = m_ResourceMgr.GetStringResource(m_DBSCRIPTSROOT, m_DBSCRIPTSPATH, "UpdateTimeCaptureStatus.sql")
    
    'create parameters
    Set oADOHelper = New utilADOHelper.cuADOHelperTrans
    oADOHelper.ADO.CreateNumericParameter "NewStatus", adInteger, 0, 0, vintNewStatus
    oADOHelper.ADO.CreateNumericParameter "NewLastModEmpNum", adInteger, 0, 0, vlngNewLastModEmpNum
    
    oADOHelper.ADO.CreateNumericParameter "vdblTimeSheetId", adDouble, 20, 0, vdblTimesheetId
    
    oADOHelper.ADO.CreateNumericParameter "LastModEmpNum", adInteger, 0, 0, vlngLastModEmpNum
    oADOHelper.ADO.CreateCharacterParameter "LastModTimestamp", adChar, 26, vstrLastModTimestamp
    
    'execute sql
    rc = oADOHelper.ExecuteSQL(SQL, adCmdText)

    UpdateTimeCaptureStatus = rc
    Set oADOHelper = Nothing
    Exit Function
eh:
    Set oADOHelper = Nothing
    
    TrsRaiseError m_strCMODULENAME, "UpdateTimeCaptureStatus", , , , DescArgsFrom(SQL, CStr(vdblTimesheetId), CStr(vlngLastModEmpNum), vstrLastModTimestamp), True
End Function

Public Function GetUpdatedTimeCaptureStatus(ByVal vdblTimesheetId As Double, ByRef rlngUpdatedLastModEmpNum As Long, ByRef rstrUpdatedLastModTimestamp As String) As Integer '//in time collection format
On Error GoTo eh
    Dim SQL As String
    Dim oADOHelper As utilADOHelper.cuADOHelperTrans
    Dim rs As ADODB.Recordset
    
    GetUpdatedTimeCaptureStatus = 0
    'get the SQL
    SQL = m_ResourceMgr.GetStringResource(m_DBSCRIPTSROOT, m_DBSCRIPTSPATH, "GetUpdatedTimeCaptureStatus.sql")
    
    'create parameters
    Set oADOHelper = New utilADOHelper.cuADOHelperTrans
    oADOHelper.ADO.CreateNumericParameter "vdblTimeSheetId", adDouble, 0, 0, vdblTimesheetId
    
    'execute sql
    Set rs = oADOHelper.ExecuteSelect(SQL)
    
    If RecordsExist(rs) Then
        rs.MoveFirst
        rlngUpdatedLastModEmpNum = rs.Fields("LAST_MOD_EMP_NUM").Value
        rstrUpdatedLastModTimestamp = rs.Fields("LAST_MOD_TIMESTAMP").Value
        GetUpdatedTimeCaptureStatus = 1
    End If
    
    Set oADOHelper = Nothing
    Exit Function
eh:
    Set oADOHelper = Nothing
    
    TrsRaiseError m_strCMODULENAME, "GetUpdatedTimeCaptureStatus", , , , DescArgsFrom(SQL, CStr(vdblTimesheetId)), True
End Function
