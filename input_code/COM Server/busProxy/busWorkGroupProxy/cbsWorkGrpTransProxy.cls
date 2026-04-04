VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "cbsWorkGrpTransProxy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'//-----The following code was generated to be used by the Client only-----//




Public Function ProcessSaveWorkGroup ( _
	ByVal v_rsWorkGroupAttributes As ADODB.Recordset,  _
	ByVal v_rsWorkGroupSchedule As ADODB.Recordset,  _
	ByVal v_rsWorkGroupAssignments As ADODB.Recordset,  _
	ByVal v_rsWorkGroupDistributionList As ADODB.Recordset,  _
	ByVal v_rsWorkGroupDistributionActivities As ADODB.Recordset,  _
	ByVal v_UserId As String,  _
	ByVal v_blnConfirmConflicts As Boolean,  _
	ByRef r_rsErrors As ADODB.Recordset,  _
	ByRef r_rsWGAttributeUpdates As ADODB.Recordset,  _
	ByRef r_rsWGBaseScheduleUpdates As ADODB.Recordset,  _
	ByRef r_rsWGLineScheduleUpdates As ADODB.Recordset,  _
	ByRef r_rsWGAssignUpdates As ADODB.Recordset,  _
	ByRef r_rsWGAssignOtherUpdates As ADODB.Recordset,  _
	ByRef r_rsWGDistributionUpdates As ADODB.Recordset,  _
	ByRef r_rsWGActivityUpdates As ADODB.Recordset,  _
	ByRef r_rsWGGLActivityUpdates As ADODB.Recordset,  _
	ByRef r_rsCascadeConflicts As ADODB.Recordset,  _
	ByRef r_rsExceptionSPA As ADODB.Recordset,  _
	ByRef r_rsAvailScheduleSPA As ADODB.Recordset,  _
	ByRef r_rsEmpSchedDatesSPA As ADODB.Recordset,  _
	ByRef r_lngWorkGroup As Long,  _
	ByRef r_blnDemandWorkGroup As Boolean _
	)  As  Boolean
  
    Dim l_objcbsWorkGroupTrans As busWorkGroup.cbsWorkGroupTrans
    
On Error GoTo eh

    Set l_objcbsWorkGroupTrans = CreateObject("busWorkGroup.cbsWorkGroupTrans")
    
    ProcessSaveWorkGroup = l_objcbsWorkGroupTrans.ProcessSaveWorkGroup(v_rsWorkGroupAttributes, v_rsWorkGroupSchedule, v_rsWorkGroupAssignments, v_rsWorkGroupDistributionList, v_rsWorkGroupDistributionActivities, v_UserId, v_blnConfirmConflicts, r_rsErrors, r_rsWGAttributeUpdates, r_rsWGBaseScheduleUpdates, r_rsWGLineScheduleUpdates, r_rsWGAssignUpdates, r_rsWGAssignOtherUpdates, r_rsWGDistributionUpdates, r_rsWGActivityUpdates, r_rsWGGLActivityUpdates, r_rsCascadeConflicts, r_rsExceptionSPA, r_rsAvailScheduleSPA, r_rsEmpSchedDatesSPA, r_lngWorkGroup, r_blnDemandWorkGroup)
    
    Set l_objcbsWorkGroupTrans = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsWorkGroupTrans = Nothing

    TrsRaiseError "busWorkGroup.cbsWorkGroupTrans", "ProcessSaveWorkGroup", , , , , True

End Function


Public Function ResumeProcessSaveWorkGroup ( _
	ByRef r_lngWorkGroup As Long,  _
	ByVal v_blnDemandWorkGroup As Boolean,  _
	ByVal v_rsWGAttributeUpdates As ADODB.Recordset,  _
	ByVal v_rsWGBaseScheduleUpdates As ADODB.Recordset,  _
	ByVal v_rsWGLineScheduleUpdates As ADODB.Recordset,  _
	ByVal v_rsWGAssignUpdates As ADODB.Recordset,  _
	ByVal v_rsWGAssignOtherUpdates As ADODB.Recordset,  _
	ByVal v_rsWGDistributionUpdates As ADODB.Recordset,  _
	ByVal v_rsWGActivityUpdates As ADODB.Recordset,  _
	ByVal v_rsWGGLActivityUpdates As ADODB.Recordset,  _
	ByVal v_rsCascadeConflicts As ADODB.Recordset,  _
	ByVal v_rsExceptionSPA As ADODB.Recordset,  _
	ByVal v_rsAvailScheduleSPA As ADODB.Recordset,  _
	ByVal v_rsEmpSchedDatesSPA As ADODB.Recordset,  _
	ByRef r_rsErrors As ADODB.Recordset,  _
	ByVal v_UserId As String _
	)  As  Boolean
  
    Dim l_objcbsWorkGroupTrans As busWorkGroup.cbsWorkGroupTrans
    
On Error GoTo eh

    Set l_objcbsWorkGroupTrans = CreateObject("busWorkGroup.cbsWorkGroupTrans")
    
    ResumeProcessSaveWorkGroup = l_objcbsWorkGroupTrans.ResumeProcessSaveWorkGroup(r_lngWorkGroup, v_blnDemandWorkGroup, v_rsWGAttributeUpdates, v_rsWGBaseScheduleUpdates, v_rsWGLineScheduleUpdates, v_rsWGAssignUpdates, v_rsWGAssignOtherUpdates, v_rsWGDistributionUpdates, v_rsWGActivityUpdates, v_rsWGGLActivityUpdates, v_rsCascadeConflicts, v_rsExceptionSPA, v_rsAvailScheduleSPA, v_rsEmpSchedDatesSPA, r_rsErrors, v_UserId)
    
    Set l_objcbsWorkGroupTrans = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsWorkGroupTrans = Nothing

    TrsRaiseError "busWorkGroup.cbsWorkGroupTrans", "ResumeProcessSaveWorkGroup", , , , , True

End Function