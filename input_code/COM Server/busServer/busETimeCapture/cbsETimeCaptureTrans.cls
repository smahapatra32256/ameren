VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "cbsETimeCaptureTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const m_strCMODULENAME As String = "busETimeCapture.cbsETimeCaptureTrans"

'Public Function UpdateTimeCaptureStatus(ByVal vdblTimeSheetId As Double, ByVal vlngLastModEmpNum As Long, ByVal vstrLastModTimestamp As String, _
'                                        ByVal vintNewStatus As Integer, ByVal vlngNewLastModEmpNum As Long) As Integer '//in time collection format
''//todo: throw obsolete
'End Function

Public Function UpdateTimeCaptureStatus(ByVal vdblTimesheetId As Double, ByVal vlngLastModEmpNum As Long, ByVal vstrLastModTimestamp As String, _
                                        ByVal vintNewStatus As Integer, ByVal vlngNewLastModEmpNum As Long, _
                                        ByRef rlngUpdatedLastModEmpNum As Long, ByRef rstrUpdatedLastModTimestamp As String) As Integer '//in time collection format
On Error GoTo eh

    Dim db As dbETimeCapture.cdETimeCaptureTrans
    Dim rc As Integer
    
    Set db = CreateObject("dbETimeCapture.cdETimeCaptureTrans")
    
    
    UpdateTimeCaptureStatus = db.UpdateTimeCaptureStatus(vdblTimesheetId, vlngLastModEmpNum, vstrLastModTimestamp, _
                                        vintNewStatus, vlngNewLastModEmpNum)
                                        
    rc = db.GetUpdatedTimeCaptureStatus(vdblTimesheetId, rlngUpdatedLastModEmpNum, rstrUpdatedLastModTimestamp)
    
    Set db = Nothing
    Exit Function
eh:
    Set db = Nothing
    
    TrsRaiseError m_strCMODULENAME, "UpdateTimeCaptureStatus", , , , DescArgsFrom(CStr(vdblTimesheetId), CStr(vlngLastModEmpNum), vstrLastModTimestamp, CStr(vintNewStatus)), True
End Function

