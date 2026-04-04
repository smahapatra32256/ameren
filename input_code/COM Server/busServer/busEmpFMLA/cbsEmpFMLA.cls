VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "cbsEmpFMLA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub Process_TTR_EMP_FMLA(Optional ByVal siOption_Num As Integer = 0, _
                                Optional ByVal strStart_Dt As String = "01-01-1994", _
                                Optional ByVal strEnd_Dt As String = "12-31-9999", _
                                Optional blnAudit As Boolean = False)
  
    Dim l_objdbEmpFMLA As dbEmpFMLA.cdEmpFMLA
    
On Error GoTo eh

    Set l_objdbEmpFMLA = CreateObject("dbEmpFMLA.cdEmpFMLA")
    
    l_objdbEmpFMLA.Process_TTR_EMP_FMLA siOption_Num, strStart_Dt, strEnd_Dt, blnAudit
    
    Set l_objdbEmpFMLA = Nothing
    
    Exit Sub
    
eh:
    Set l_objdbEmpFMLA = Nothing

    TrsRaiseError "busEmpFMLA.cbsEmpFMLA", "Process_TTR_EMP_FMLA", , , , _
            DescArgsFrom(CStr(siOption_Num), strStart_Dt, strEnd_Dt, CStr(blnAudit)), True
End Sub

