VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cbsSchedRotation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const mstrCMODULENAME As String = "busSchedRotation"
'Private Const mstrCZERO As String = "0"
'Private Const mstrCVARFIELDNAMES_RTH As String = "ROT_CD"
'Private Const mstrCVARFIELDNAMES_RTD As String = "ROT_CD;ROT_WEEK"
'Private m_strUcaseYorN As String
'Private m_lngUserId As Long
'Private m_lngErrorNo As Long
'Private m_lngRecordId As Long
'Private m_intNextRotationCode As Integer

'public call to get rotation header
Public Function GetRotationSchedHeader _
                (Optional ByVal v_intRotationCode As Integer _
                ) As ADODB.Recordset

On Error GoTo GetRotationSchedHeader_EH
    
    Dim l_dbSchedRotation As dbSchedRotation.cdSchedRotation

    Set l_dbSchedRotation = CreateObject("dbSchedRotation.cdSchedRotation")

    Set GetRotationSchedHeader = l_dbSchedRotation.GetRotationSchedHeader(v_intRotationCode)

    Set l_dbSchedRotation = Nothing
    
    Exit Function

GetRotationSchedHeader_EH:
    Set l_dbSchedRotation = Nothing
    TrsRaiseError mstrCMODULENAME, "GetRotationSchedHeader"

End Function
'public call to get rotation detail
Public Function GetRotationSchedDetail _
                (Optional ByVal v_intRotationCode As Integer, _
                 Optional ByVal v_intRotationWeek As Integer _
                ) As ADODB.Recordset

On Error GoTo GetRotationSchedDetail_EH

    Dim l_dbSchedRotation As dbSchedRotation.cdSchedRotation

    Set l_dbSchedRotation = CreateObject("dbSchedRotation.cdSchedRotation")

    Set GetRotationSchedDetail = l_dbSchedRotation.GetRotationSchedDetail(v_intRotationCode, v_intRotationWeek)

    Set l_dbSchedRotation = Nothing
Exit Function

GetRotationSchedDetail_EH:
    Set l_dbSchedRotation = Nothing
    TrsRaiseError mstrCMODULENAME, "GetRotationSchedDetail"

End Function

