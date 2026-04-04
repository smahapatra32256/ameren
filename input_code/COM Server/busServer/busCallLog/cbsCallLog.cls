VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 4  'RequiresNewTransaction
END
Attribute VB_Name = "cbsCallLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const m_strCMODULENAME As String = "cbsCallLog"

Public Function OTStartCall( _
                    ByVal v_lngOTOppNum As Long, _
                    ByVal v_lngEmpNum As Long, _
                    ByVal v_strNumDialed As String, _
                    ByVal v_strUserID As String _
                ) As String
' Return a 26 char timestamp representing the server time the call was initiated.

    Dim l_cdCallLog As cdCallLog
    
On Error GoTo eh

    Set l_cdCallLog = CreateObject("dbCallLog.cdCallLog")
    OTStartCall = l_cdCallLog.OTStartCall(v_lngOTOppNum, v_lngEmpNum, v_strNumDialed, v_strUserID)
    
    Set l_cdCallLog = Nothing
    GetObjectContext.SetComplete
    Exit Function

eh:
    GetObjectContext.SetAbort
    Set l_cdCallLog = Nothing
    TrsRaiseError m_strCMODULENAME, "OTStartCall", , , , , True

End Function


Public Sub OTEndCall( _
                    ByVal v_lngOTOppNum As Long, _
                    ByVal v_lngEmpNum As Long, _
                    ByVal v_strCallStartID As String, _
                    ByVal v_strUserID As String)

    Dim l_cdCallLog As cdCallLog
    
On Error GoTo eh

    Set l_cdCallLog = CreateObject("dbCallLog.cdCallLog")
    l_cdCallLog.OTEndCall v_lngOTOppNum, v_lngEmpNum, v_strCallStartID, v_strUserID
    
    Set l_cdCallLog = Nothing
    GetObjectContext.SetComplete
    Exit Sub

eh:
    GetObjectContext.SetAbort
    Set l_cdCallLog = Nothing
    TrsRaiseError m_strCMODULENAME, "OTEndCall", , , , , True

End Sub
