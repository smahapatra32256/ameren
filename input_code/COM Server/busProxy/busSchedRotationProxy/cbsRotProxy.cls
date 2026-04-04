VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "cbsRotProxy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'//-----The following code was generated to be used by the Client only-----//




Public Function GetRotationSchedHeader ( _
	Optional ByVal v_intRotationCode As Integer _
	)  As  ADODB.Recordset
  
    Dim l_objcbsSchedRotation As busSchedRotation.cbsSchedRotation
    
On Error GoTo eh

    Set l_objcbsSchedRotation = CreateObject("busSchedRotation.cbsSchedRotation")
    
    Set GetRotationSchedHeader = l_objcbsSchedRotation.GetRotationSchedHeader(v_intRotationCode)
    
    Set l_objcbsSchedRotation = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsSchedRotation = Nothing

    TrsRaiseError "busSchedRotation.cbsSchedRotation", "GetRotationSchedHeader", , , , , True

End Function


Public Function GetRotationSchedDetail ( _
	Optional ByVal v_intRotationCode As Integer,  _
	Optional ByVal v_intRotationWeek As Integer _
	)  As  ADODB.Recordset
  
    Dim l_objcbsSchedRotation As busSchedRotation.cbsSchedRotation
    
On Error GoTo eh

    Set l_objcbsSchedRotation = CreateObject("busSchedRotation.cbsSchedRotation")
    
    Set GetRotationSchedDetail = l_objcbsSchedRotation.GetRotationSchedDetail(v_intRotationCode, v_intRotationWeek)
    
    Set l_objcbsSchedRotation = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsSchedRotation = Nothing

    TrsRaiseError "busSchedRotation.cbsSchedRotation", "GetRotationSchedDetail", , , , , True

End Function