VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "cdWorkGroupTransProxy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'//-----The following code was generated to be used by the Client only-----//




Public Function SaveWGAssignments ( _
	ByVal v_rsWGAssignments As ADODB.Recordset,  _
	ByVal v_rsAdjPrimaries As ADODB.Recordset _
	)  As  Boolean
  
    Dim l_objcdWorkGroupTrans As dbWorkGroup.cdWorkGroupTrans
    
On Error GoTo eh

    Set l_objcdWorkGroupTrans = CreateObject("dbWorkGroup.cdWorkGroupTrans")
    
    SaveWGAssignments = l_objcdWorkGroupTrans.SaveWGAssignments(v_rsWGAssignments, v_rsAdjPrimaries)
    
    Set l_objcdWorkGroupTrans = Nothing
    
    Exit Function
    
eh:
    Set l_objcdWorkGroupTrans = Nothing

    TrsRaiseError "dbWorkGroup.cdWorkGroupTrans", "SaveWGAssignments", , , , , True

End Function


Public Function SaveWGLineSchedules ( _
	ByRef r_rsWGSchedules As ADODB.Recordset _
	)  As  Boolean
  
    Dim l_objcdWorkGroupTrans As dbWorkGroup.cdWorkGroupTrans
    
On Error GoTo eh

    Set l_objcdWorkGroupTrans = CreateObject("dbWorkGroup.cdWorkGroupTrans")
    
    SaveWGLineSchedules = l_objcdWorkGroupTrans.SaveWGLineSchedules(r_rsWGSchedules)
    
    Set l_objcdWorkGroupTrans = Nothing
    
    Exit Function
    
eh:
    Set l_objcdWorkGroupTrans = Nothing

    TrsRaiseError "dbWorkGroup.cdWorkGroupTrans", "SaveWGLineSchedules", , , , , True

End Function


Public Function SaveWGBaseSchedules ( _
	ByRef r_rsWGSchedules As ADODB.Recordset _
	)  As  Boolean
  
    Dim l_objcdWorkGroupTrans As dbWorkGroup.cdWorkGroupTrans
    
On Error GoTo eh

    Set l_objcdWorkGroupTrans = CreateObject("dbWorkGroup.cdWorkGroupTrans")
    
    SaveWGBaseSchedules = l_objcdWorkGroupTrans.SaveWGBaseSchedules(r_rsWGSchedules)
    
    Set l_objcdWorkGroupTrans = Nothing
    
    Exit Function
    
eh:
    Set l_objcdWorkGroupTrans = Nothing

    TrsRaiseError "dbWorkGroup.cdWorkGroupTrans", "SaveWGBaseSchedules", , , , , True

End Function


Public Function SaveWGAttributes ( _
	ByVal v_rsWGAttributes As ADODB.Recordset,  _
	ByVal v_strUserId As String _
	)  As  Boolean
  
    Dim l_objcdWorkGroupTrans As dbWorkGroup.cdWorkGroupTrans
    
On Error GoTo eh

    Set l_objcdWorkGroupTrans = CreateObject("dbWorkGroup.cdWorkGroupTrans")
    
    SaveWGAttributes = l_objcdWorkGroupTrans.SaveWGAttributes(v_rsWGAttributes, v_strUserId)
    
    Set l_objcdWorkGroupTrans = Nothing
    
    Exit Function
    
eh:
    Set l_objcdWorkGroupTrans = Nothing

    TrsRaiseError "dbWorkGroup.cdWorkGroupTrans", "SaveWGAttributes", , , , , True

End Function


Public Function SaveWGDistribution ( _
	ByVal v_rsWGDistribution As ADODB.Recordset _
	)  As  Boolean
  
    Dim l_objcdWorkGroupTrans As dbWorkGroup.cdWorkGroupTrans
    
On Error GoTo eh

    Set l_objcdWorkGroupTrans = CreateObject("dbWorkGroup.cdWorkGroupTrans")
    
    SaveWGDistribution = l_objcdWorkGroupTrans.SaveWGDistribution(v_rsWGDistribution)
    
    Set l_objcdWorkGroupTrans = Nothing
    
    Exit Function
    
eh:
    Set l_objcdWorkGroupTrans = Nothing

    TrsRaiseError "dbWorkGroup.cdWorkGroupTrans", "SaveWGDistribution", , , , , True

End Function


Public Function SaveWGActivities ( _
	ByVal v_rsWGActivities As ADODB.Recordset _
	)  As  Boolean
  
    Dim l_objcdWorkGroupTrans As dbWorkGroup.cdWorkGroupTrans
    
On Error GoTo eh

    Set l_objcdWorkGroupTrans = CreateObject("dbWorkGroup.cdWorkGroupTrans")
    
    SaveWGActivities = l_objcdWorkGroupTrans.SaveWGActivities(v_rsWGActivities)
    
    Set l_objcdWorkGroupTrans = Nothing
    
    Exit Function
    
eh:
    Set l_objcdWorkGroupTrans = Nothing

    TrsRaiseError "dbWorkGroup.cdWorkGroupTrans", "SaveWGActivities", , , , , True

End Function


Public Function SaveGLActivities ( _
	ByVal v_rsGLActivities As ADODB.Recordset _
	)  As  Boolean
  
    Dim l_objcdWorkGroupTrans As dbWorkGroup.cdWorkGroupTrans
    
On Error GoTo eh

    Set l_objcdWorkGroupTrans = CreateObject("dbWorkGroup.cdWorkGroupTrans")
    
    SaveGLActivities = l_objcdWorkGroupTrans.SaveGLActivities(v_rsGLActivities)
    
    Set l_objcdWorkGroupTrans = Nothing
    
    Exit Function
    
eh:
    Set l_objcdWorkGroupTrans = Nothing

    TrsRaiseError "dbWorkGroup.cdWorkGroupTrans", "SaveGLActivities", , , , , True

End Function