VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "cbsRotTransProxy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'//-----The following code was generated to be used by the Client only-----//




Public Function RecordRotation ( _
	ByRef r_rsScheduleRotationRequest As ADODB.Recordset,  _
	ByRef r_rsScheduleRotationDetailRequest As ADODB.Recordset,  _
	ByVal v_lngUserId As Long,  _
	ByRef r_rsErrors As ADODB.Recordset _
	)  As  Boolean
  
    Dim l_objcbsSchedRotationTrans As busSchedRotation.cbsSchedRotationTrans
    
On Error GoTo eh

    Set l_objcbsSchedRotationTrans = CreateObject("busSchedRotation.cbsSchedRotationTrans")
    
    RecordRotation = l_objcbsSchedRotationTrans.RecordRotation(r_rsScheduleRotationRequest, r_rsScheduleRotationDetailRequest, v_lngUserId, r_rsErrors)
    
    Set l_objcbsSchedRotationTrans = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsSchedRotationTrans = Nothing

    TrsRaiseError "busSchedRotation.cbsSchedRotationTrans", "RecordRotation", , , , , True

End Function


Public Function IsRotDetPatternExists ( _
	ByRef r_rsScheduleRotationRequest As ADODB.Recordset,  _
	ByRef r_rsScheduleRotationDetailRequest As ADODB.Recordset _
	)  As  Boolean
  
    Dim l_objcbsSchedRotationTrans As busSchedRotation.cbsSchedRotationTrans
    
On Error GoTo eh

    Set l_objcbsSchedRotationTrans = CreateObject("busSchedRotation.cbsSchedRotationTrans")
    
    IsRotDetPatternExists = l_objcbsSchedRotationTrans.IsRotDetPatternExists(r_rsScheduleRotationRequest, r_rsScheduleRotationDetailRequest)
    
    Set l_objcbsSchedRotationTrans = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsSchedRotationTrans = Nothing

    TrsRaiseError "busSchedRotation.cbsSchedRotationTrans", "IsRotDetPatternExists", , , , , True

End Function


Public Function MultiRecordRotation ( _
	ByRef r_rsScheduleRotationRequest As ADODB.Recordset,  _
	ByRef r_rsScheduleRotationDetailRequest As ADODB.Recordset,  _
	ByVal v_lngUserId As Long,  _
	ByRef r_rsErrors As ADODB.Recordset _
	)  As  Boolean
  
    Dim l_objcbsSchedRotationTrans As busSchedRotation.cbsSchedRotationTrans
    
On Error GoTo eh

    Set l_objcbsSchedRotationTrans = CreateObject("busSchedRotation.cbsSchedRotationTrans")
    
    MultiRecordRotation = l_objcbsSchedRotationTrans.MultiRecordRotation(r_rsScheduleRotationRequest, r_rsScheduleRotationDetailRequest, v_lngUserId, r_rsErrors)
    
    Set l_objcbsSchedRotationTrans = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsSchedRotationTrans = Nothing

    TrsRaiseError "busSchedRotation.cbsSchedRotationTrans", "MultiRecordRotation", , , , , True

End Function