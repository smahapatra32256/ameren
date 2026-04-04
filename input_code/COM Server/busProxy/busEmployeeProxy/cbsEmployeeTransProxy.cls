VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "cbsEmployeeTransProxy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'//-----The following code was generated to be used by the Client only-----//




Public Function ProcessEmpQualsSkillsEntries ( _
	ByVal vadrsQAndSInformation As ADODB.Recordset,  _
	ByRef radrsErrors As ADODB.Recordset,  _
	ByVal vlUserEmployeeNumber As Long _
	)  As  Boolean
  
    Dim l_objcbsEmployeeTrans As busEmployee.cbsEmployeeTrans
    
On Error GoTo eh

    Set l_objcbsEmployeeTrans = CreateObject("busEmployee.cbsEmployeeTrans")
    
    ProcessEmpQualsSkillsEntries = l_objcbsEmployeeTrans.ProcessEmpQualsSkillsEntries(vadrsQAndSInformation, radrsErrors, vlUserEmployeeNumber)
    
    Set l_objcbsEmployeeTrans = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsEmployeeTrans = Nothing

    TrsRaiseError "busEmployee.cbsEmployeeTrans", "ProcessEmpQualsSkillsEntries", , , , , True

End Function


Public Function ProcessSeniorityEntries ( _
	ByVal vadrsSeniorityInformation As ADODB.Recordset,  _
	ByVal vadrsFrozenSeniorityInformation As ADODB.Recordset,  _
	ByRef radrsErrors As ADODB.Recordset,  _
	ByVal vlSeniorityEmployeeNumber As Long,  _
	ByVal vlUserEmployeeNumber As Long _
	)  As  Boolean
  
    Dim l_objcbsEmployeeTrans As busEmployee.cbsEmployeeTrans
    
On Error GoTo eh

    Set l_objcbsEmployeeTrans = CreateObject("busEmployee.cbsEmployeeTrans")
    
    ProcessSeniorityEntries = l_objcbsEmployeeTrans.ProcessSeniorityEntries(vadrsSeniorityInformation, vadrsFrozenSeniorityInformation, radrsErrors, vlSeniorityEmployeeNumber, vlUserEmployeeNumber)
    
    Set l_objcbsEmployeeTrans = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsEmployeeTrans = Nothing

    TrsRaiseError "busEmployee.cbsEmployeeTrans", "ProcessSeniorityEntries", , , , , True

End Function


Public Function EmployeeCompany ( _
	ByVal v_datDeptEffDate As Date,  _
	ByVal v_strDeptNum As String,  _
	ByVal v_strPriorDeptNum As String,  _
	Optional ByVal v_strCompanyCode As String = "NA",  _
	Optional ByVal v_strPriorCompanyCode As String = "NA",  _
	Optional ByVal v_strChgDate As Date = "01/01/100" _
	)  As  String
  
    Dim l_objcbsEmployeeTrans As busEmployee.cbsEmployeeTrans
    
On Error GoTo eh

    Set l_objcbsEmployeeTrans = CreateObject("busEmployee.cbsEmployeeTrans")
    
    EmployeeCompany = l_objcbsEmployeeTrans.EmployeeCompany(v_datDeptEffDate, v_strDeptNum, v_strPriorDeptNum, v_strCompanyCode, v_strPriorCompanyCode, v_strChgDate)
    
    Set l_objcbsEmployeeTrans = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsEmployeeTrans = Nothing

    TrsRaiseError "busEmployee.cbsEmployeeTrans", "EmployeeCompany", , , , , True

End Function


Public Function ProcessEmpCanvassJobEntries ( _
	ByVal vadrsCanvassJobInformation As ADODB.Recordset,  _
	ByRef radrsErrors As ADODB.Recordset,  _
	ByVal vlUserEmployeeNumber As Long _
	)  As  Boolean
  
    Dim l_objcbsEmployeeTrans As busEmployee.cbsEmployeeTrans
    
On Error GoTo eh

    Set l_objcbsEmployeeTrans = CreateObject("busEmployee.cbsEmployeeTrans")
    
    ProcessEmpCanvassJobEntries = l_objcbsEmployeeTrans.ProcessEmpCanvassJobEntries(vadrsCanvassJobInformation, radrsErrors, vlUserEmployeeNumber)
    
    Set l_objcbsEmployeeTrans = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsEmployeeTrans = Nothing

    TrsRaiseError "busEmployee.cbsEmployeeTrans", "ProcessEmpCanvassJobEntries", , , , , True

End Function


Public Function ProcessEmpIneligibilityEntries ( _
	ByVal vadrsIneligibilityInformation As ADODB.Recordset,  _
	ByRef radrsErrors As ADODB.Recordset,  _
	ByVal vlUserEmployeeNumber As Long _
	)  As  Boolean
  
    Dim l_objcbsEmployeeTrans As busEmployee.cbsEmployeeTrans
    
On Error GoTo eh

    Set l_objcbsEmployeeTrans = CreateObject("busEmployee.cbsEmployeeTrans")
    
    ProcessEmpIneligibilityEntries = l_objcbsEmployeeTrans.ProcessEmpIneligibilityEntries(vadrsIneligibilityInformation, radrsErrors, vlUserEmployeeNumber)
    
    Set l_objcbsEmployeeTrans = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsEmployeeTrans = Nothing

    TrsRaiseError "busEmployee.cbsEmployeeTrans", "ProcessEmpIneligibilityEntries", , , , , True

End Function


Public Function ProcessVolunteerEntries ( _
	ByVal vadrsVolunteerInformation As ADODB.Recordset,  _
	ByRef radrsErrors As ADODB.Recordset,  _
	ByVal vlUserEmployeeNumber As Long _
	)  As  Boolean
  
    Dim l_objcbsEmployeeTrans As busEmployee.cbsEmployeeTrans
    
On Error GoTo eh

    Set l_objcbsEmployeeTrans = CreateObject("busEmployee.cbsEmployeeTrans")
    
    ProcessVolunteerEntries = l_objcbsEmployeeTrans.ProcessVolunteerEntries(vadrsVolunteerInformation, radrsErrors, vlUserEmployeeNumber)
    
    Set l_objcbsEmployeeTrans = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsEmployeeTrans = Nothing

    TrsRaiseError "busEmployee.cbsEmployeeTrans", "ProcessVolunteerEntries", , , , , True

End Function


Public Function ProcessEmployeeECWorkReasons ( _
	ByVal vadrsEmployeeECWorkReasons As ADODB.Recordset,  _
	ByRef radrsErrors As ADODB.Recordset,  _
	ByVal vlUserEmployeeNumber As Long _
	)  As  Boolean
  
    Dim l_objcbsEmployeeTrans As busEmployee.cbsEmployeeTrans
    
On Error GoTo eh

    Set l_objcbsEmployeeTrans = CreateObject("busEmployee.cbsEmployeeTrans")
    
    ProcessEmployeeECWorkReasons = l_objcbsEmployeeTrans.ProcessEmployeeECWorkReasons(vadrsEmployeeECWorkReasons, radrsErrors, vlUserEmployeeNumber)
    
    Set l_objcbsEmployeeTrans = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsEmployeeTrans = Nothing

    TrsRaiseError "busEmployee.cbsEmployeeTrans", "ProcessEmployeeECWorkReasons", , , , , True

End Function