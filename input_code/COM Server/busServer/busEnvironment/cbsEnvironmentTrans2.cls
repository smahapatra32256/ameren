VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "cbsEnvironmentTrans2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const MODULENAME As String = "cbsFunSecurityTrans2"
Private Const m_strCMODULENAME As String = MODULENAME

Public Function SetUserPreference(ByVal v_lngEmpNum As Long, _
                                    ByVal v_eTypeCd As PreferenceTypeCode, _
                                    ByVal v_strOtherTypeCd As String, _
                                    ByVal v_rsPreference As ADODB.Recordset) As Boolean
                                    
    On Error GoTo eh

    Dim dbEnvTrans As dbEnvironment.cdEnvironmentTrans2
    Dim strTypeCd As String
    
'    Set dbEnvTrans = CreateObject("dbEnvironment.cdEnvironmentTrans2")
    
    If v_eTypeCd <> PreferenceTypeCode.Other Then
        strTypeCd = PreferenceTypeCodeToString(v_eTypeCd)
    Else
        strTypeCd = v_strOtherTypeCd
    End If
        
    Set dbEnvTrans = CreateObject("dbEnvironment.cdEnvironmentTrans2")

    SetUserPreference = dbEnvTrans.SetUserPreference(v_lngEmpNum, strTypeCd, v_rsPreference)
    
    Set dbEnvTrans = Nothing
    
'    GetObjectContext.SetComplete
    Exit Function
eh:
    'Abort Transaction
'    GetObjectContext.SetAbort
    Set dbEnvTrans = Nothing
    
    TrsRaiseError MODULENAME _
        , "SetUserPreference" _
        , _
        , , DescArgsFrom(CStr(v_lngEmpNum), PreferenceTypeCodeToString(v_eTypeCd)), _
        , True
End Function

