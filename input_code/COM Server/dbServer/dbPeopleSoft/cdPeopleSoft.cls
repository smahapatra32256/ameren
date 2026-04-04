VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "cdPeopleSoft"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const m_strCMODULENAME As String = "cdEmployee"
Private Const m_strC1SPACESTRING As String = " "


Public Function GetPeopleSoftPhoneNumber( _
            ByVal v_aryEmployees As Variant _
            ) As ADODB.Recordset
            
    Dim l_strEmployeeSelection As String
    Dim datahub As Ameren_TRIS_Server_COMIntegration4.HRDataHubEmployee
    Dim l_objRS As ADODB.Recordset

    On Error GoTo eh
    
    If IsArray(v_aryEmployees) Then
        l_strEmployeeSelection = BuildInClause("EMPLID", v_aryEmployees, False) & " "
    Else
        l_strEmployeeSelection = "EMPLID = '" & CStr(CLng(v_aryEmployees)) & "' "
    End If

'    l_strSQL = "SELECT CAST(SUBSTR(EMPLID,1,9) AS INTEGER) as EMP_NUM, " & _
'                " SUBSTR(PHONE,1,3) || SUBSTR(PHONE,5,3) || " & _
'                " SUBSTR(PHONE,9,4) as HOME_PHONE_NUM " & _
'            " FROM PS_PERSONAL_PHONE " & _
'            " WHERE PHONE_TYPE = 'HOME' " & l_strEmployeeSelection & _
'            " ORDER BY 1"
    Set datahub = New Ameren_TRIS_Server_COMIntegration4.HRDataHubEmployee
    Set l_objRS = datahub.GetPeopleSoftPhoneNumber_(v_aryEmployees, "HOME")

    Set GetPeopleSoftPhoneNumber = l_objRS
    
    Set l_objRS = Nothing
    Set datahub = Nothing

    Exit Function
eh:
    Set datahub = Nothing
    Set l_objRS = Nothing
    TrsRaiseError m_strCMODULENAME, "GetPeopleSoftPhoneNumber", , , , DescArgsFrom(l_strEmployeeSelection)
End Function

