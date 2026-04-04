VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "cbsSchedCodeMaint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const mstrCMODULENAME As String = "busSchedCodeMaint"

Public Function GetRotationSchedDetail _
                (ByVal v_iSchedCode As Integer) As ADODB.Recordset

On Error GoTo GetRotationSchedDetail_EH

    Dim l_dbSchedRotation As dbSchedCodeMaint.cdSchedCodeMaint

    Set l_dbSchedRotation = CreateObject("dbSchedCodeMaint.cdSchedCodeMaint")

    Set GetRotationSchedDetail = l_dbSchedRotation.GetScheduleDetail(v_iSchedCode)

    Set l_dbSchedRotation = Nothing
Exit Function

GetRotationSchedDetail_EH:
    Set l_dbSchedRotation = Nothing
    TrsRaiseError mstrCMODULENAME, "GetRotationSchedDetail"

End Function

