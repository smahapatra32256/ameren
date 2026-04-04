VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "cbsCodeTablesTransProxy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'//-----The following code was generated to be used by the Client only-----//




Public Function ProcessOTCodeTableEntries ( _
	ByVal vadrsOTCodeTableInformation As ADODB.Recordset,  _
	ByRef radrsErrors As ADODB.Recordset,  _
	ByVal vlUserEmployeeNumber As Long _
	)  As  Boolean
  
    Dim l_objcbsCodeTablesTrans As busCodeTables.cbsCodeTablesTrans
    
On Error GoTo eh

    Set l_objcbsCodeTablesTrans = CreateObject("busCodeTables.cbsCodeTablesTrans")
    
    ProcessOTCodeTableEntries = l_objcbsCodeTablesTrans.ProcessOTCodeTableEntries(vadrsOTCodeTableInformation, radrsErrors, vlUserEmployeeNumber)
    
    Set l_objcbsCodeTablesTrans = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsCodeTablesTrans = Nothing

    TrsRaiseError "busCodeTables.cbsCodeTablesTrans", "ProcessOTCodeTableEntries", , , , , True

End Function


Public Function ProcessPromotionalLinesEntries ( _
	ByVal vadrsPromotionalLines As ADODB.Recordset,  _
	ByRef radrsErrors As ADODB.Recordset,  _
	ByVal vlUserEmployeeNumber As Long _
	)  As  Boolean
  
    Dim l_objcbsCodeTablesTrans As busCodeTables.cbsCodeTablesTrans
    
On Error GoTo eh

    Set l_objcbsCodeTablesTrans = CreateObject("busCodeTables.cbsCodeTablesTrans")
    
    ProcessPromotionalLinesEntries = l_objcbsCodeTablesTrans.ProcessPromotionalLinesEntries(vadrsPromotionalLines, radrsErrors, vlUserEmployeeNumber)
    
    Set l_objcbsCodeTablesTrans = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsCodeTablesTrans = Nothing

    TrsRaiseError "busCodeTables.cbsCodeTablesTrans", "ProcessPromotionalLinesEntries", , , , , True

End Function