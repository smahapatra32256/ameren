VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "cdCodeTablesTransProxy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'//-----The following code was generated to be used by the Client only-----//
Public Enum TrsComponentType
    trReliefComponent = 1
    trShiftComponent = 2
    trOtherComponent = 3
End Enum





Public Function dbProcessCraftSave ( _
	ByVal vadrsOTCodeTableEntries As ADODB.Recordset,  _
	ByVal vlUserEmployeeNumber As Long _
	)  As  Boolean
  
    Dim l_objcdCodeTablesTrans As dbCodeTables.cdCodeTablesTrans
    
On Error GoTo eh

    Set l_objcdCodeTablesTrans = CreateObject("dbCodeTables.cdCodeTablesTrans")
    
    dbProcessCraftSave = l_objcdCodeTablesTrans.dbProcessCraftSave(vadrsOTCodeTableEntries, vlUserEmployeeNumber)
    
    Set l_objcdCodeTablesTrans = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTablesTrans = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTablesTrans", "dbProcessCraftSave", , , , , True

End Function


Public Function dbProcessQualificationsSave ( _
	ByVal vadrsOTCodeTableEntries As ADODB.Recordset,  _
	ByVal vlUserEmployeeNumber As Long _
	)  As  Boolean
  
    Dim l_objcdCodeTablesTrans As dbCodeTables.cdCodeTablesTrans
    
On Error GoTo eh

    Set l_objcdCodeTablesTrans = CreateObject("dbCodeTables.cdCodeTablesTrans")
    
    dbProcessQualificationsSave = l_objcdCodeTablesTrans.dbProcessQualificationsSave(vadrsOTCodeTableEntries, vlUserEmployeeNumber)
    
    Set l_objcdCodeTablesTrans = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTablesTrans = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTablesTrans", "dbProcessQualificationsSave", , , , , True

End Function


Public Function dbProcessSpecialSkillsSave ( _
	ByVal vadrsOTCodeTableEntries As ADODB.Recordset,  _
	ByVal vlUserEmployeeNumber As Long _
	)  As  Boolean
  
    Dim l_objcdCodeTablesTrans As dbCodeTables.cdCodeTablesTrans
    
On Error GoTo eh

    Set l_objcdCodeTablesTrans = CreateObject("dbCodeTables.cdCodeTablesTrans")
    
    dbProcessSpecialSkillsSave = l_objcdCodeTablesTrans.dbProcessSpecialSkillsSave(vadrsOTCodeTableEntries, vlUserEmployeeNumber)
    
    Set l_objcdCodeTablesTrans = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTablesTrans = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTablesTrans", "dbProcessSpecialSkillsSave", , , , , True

End Function


Public Function dbProcessOTReasonsSave ( _
	ByVal vadrsOTCodeTableEntries As ADODB.Recordset,  _
	ByVal vlUserEmployeeNumber As Long _
	)  As  Boolean
  
    Dim l_objcdCodeTablesTrans As dbCodeTables.cdCodeTablesTrans
    
On Error GoTo eh

    Set l_objcdCodeTablesTrans = CreateObject("dbCodeTables.cdCodeTablesTrans")
    
    dbProcessOTReasonsSave = l_objcdCodeTablesTrans.dbProcessOTReasonsSave(vadrsOTCodeTableEntries, vlUserEmployeeNumber)
    
    Set l_objcdCodeTablesTrans = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTablesTrans = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTablesTrans", "dbProcessOTReasonsSave", , , , , True

End Function


Public Function dbProcessOTResponsesSave ( _
	ByVal vadrsOTCodeTableEntries As ADODB.Recordset,  _
	ByVal vlUserEmployeeNumber As Long _
	)  As  Boolean
  
    Dim l_objcdCodeTablesTrans As dbCodeTables.cdCodeTablesTrans
    
On Error GoTo eh

    Set l_objcdCodeTablesTrans = CreateObject("dbCodeTables.cdCodeTablesTrans")
    
    dbProcessOTResponsesSave = l_objcdCodeTablesTrans.dbProcessOTResponsesSave(vadrsOTCodeTableEntries, vlUserEmployeeNumber)
    
    Set l_objcdCodeTablesTrans = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTablesTrans = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTablesTrans", "dbProcessOTResponsesSave", , , , , True

End Function


Public Function dbProcessAssociateCraftToJob ( _
	ByVal vadrsOTCodeTableEntries As ADODB.Recordset,  _
	ByVal vlUserEmployeeNumber As Long _
	)  As  Boolean
  
    Dim l_objcdCodeTablesTrans As dbCodeTables.cdCodeTablesTrans
    
On Error GoTo eh

    Set l_objcdCodeTablesTrans = CreateObject("dbCodeTables.cdCodeTablesTrans")
    
    dbProcessAssociateCraftToJob = l_objcdCodeTablesTrans.dbProcessAssociateCraftToJob(vadrsOTCodeTableEntries, vlUserEmployeeNumber)
    
    Set l_objcdCodeTablesTrans = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTablesTrans = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTablesTrans", "dbProcessAssociateCraftToJob", , , , , True

End Function


Public Function dbProcessPromotionalLines ( _
	ByVal vadrsPromotionalLinesEntries As ADODB.Recordset,  _
	ByVal vlUserEmployeeNumber As Long _
	)  As  Boolean
  
    Dim l_objcdCodeTablesTrans As dbCodeTables.cdCodeTablesTrans
    
On Error GoTo eh

    Set l_objcdCodeTablesTrans = CreateObject("dbCodeTables.cdCodeTablesTrans")
    
    dbProcessPromotionalLines = l_objcdCodeTablesTrans.dbProcessPromotionalLines(vadrsPromotionalLinesEntries, vlUserEmployeeNumber)
    
    Set l_objcdCodeTablesTrans = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTablesTrans = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTablesTrans", "dbProcessPromotionalLines", , , , , True

End Function


Public Function dbProcessOTBankHeaderSave ( _
	ByVal vadrsOTCodeTableEntries As ADODB.Recordset,  _
	ByVal vlUserEmployeeNumber As Long _
	)  As  Boolean
  
    Dim l_objcdCodeTablesTrans As dbCodeTables.cdCodeTablesTrans
    
On Error GoTo eh

    Set l_objcdCodeTablesTrans = CreateObject("dbCodeTables.cdCodeTablesTrans")
    
    dbProcessOTBankHeaderSave = l_objcdCodeTablesTrans.dbProcessOTBankHeaderSave(vadrsOTCodeTableEntries, vlUserEmployeeNumber)
    
    Set l_objcdCodeTablesTrans = Nothing
    
    Exit Function
    
eh:
    Set l_objcdCodeTablesTrans = Nothing

    TrsRaiseError "dbCodeTables.cdCodeTablesTrans", "dbProcessOTBankHeaderSave", , , , , True

End Function