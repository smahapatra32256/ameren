VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "cbsDistMgrProxy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'//-----The following code was generated to be used by the Client only-----//
Public Enum TrsDistributionMgrErrorEnum
    trErrdb_E_Invalid_Zero_v_lngWgNum_F = -2147191423
End Enum





Public Function GetDefaultAcct ( _
	ByVal v_strExceptionCd As String,  _
	ByVal v_strDistDept As String,  _
	Optional ByVal v_strBank As String = "",  _
	Optional ByVal v_strOverrideExcCd As String = "",  _
	Optional ByVal v_strAcctDiv As String = "" _
	)  As  String
  
    Dim l_objcbsDistributionMgr As busDistributionMgr.cbsDistributionMgr
    
On Error GoTo eh

    Set l_objcbsDistributionMgr = CreateObject("busDistributionMgr.cbsDistributionMgr")
    
    GetDefaultAcct = l_objcbsDistributionMgr.GetDefaultAcct(v_strExceptionCd, v_strDistDept, v_strBank, v_strOverrideExcCd, v_strAcctDiv)
    
    Set l_objcbsDistributionMgr = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsDistributionMgr = Nothing

    TrsRaiseError "busDistributionMgr.cbsDistributionMgr", "GetDefaultAcct", , , , , True

End Function


Public Function IsDOJMWorkRequest ( _
	ByVal v_varDistributionId As Variant _
	)  As  Boolean
  
    Dim l_objcbsDistributionMgr As busDistributionMgr.cbsDistributionMgr
    
On Error GoTo eh

    Set l_objcbsDistributionMgr = CreateObject("busDistributionMgr.cbsDistributionMgr")
    
    IsDOJMWorkRequest = l_objcbsDistributionMgr.IsDOJMWorkRequest(v_varDistributionId)
    
    Set l_objcbsDistributionMgr = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsDistributionMgr = Nothing

    TrsRaiseError "busDistributionMgr.cbsDistributionMgr", "IsDOJMWorkRequest", , , , , True

End Function


Public Function IsEMPRVJobTaskWG ( _
	ByVal v_varWorkGroup As Variant _
	)  As  Boolean
  
    Dim l_objcbsDistributionMgr As busDistributionMgr.cbsDistributionMgr
    
On Error GoTo eh

    Set l_objcbsDistributionMgr = CreateObject("busDistributionMgr.cbsDistributionMgr")
    
    IsEMPRVJobTaskWG = l_objcbsDistributionMgr.IsEMPRVJobTaskWG(v_varWorkGroup)
    
    Set l_objcbsDistributionMgr = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsDistributionMgr = Nothing

    TrsRaiseError "busDistributionMgr.cbsDistributionMgr", "IsEMPRVJobTaskWG", , , , , True

End Function


Public Function IsServiceRequest ( _
	ByVal v_varDistributionId As Variant _
	)  As  Boolean
  
    Dim l_objcbsDistributionMgr As busDistributionMgr.cbsDistributionMgr
    
On Error GoTo eh

    Set l_objcbsDistributionMgr = CreateObject("busDistributionMgr.cbsDistributionMgr")
    
    IsServiceRequest = l_objcbsDistributionMgr.IsServiceRequest(v_varDistributionId)
    
    Set l_objcbsDistributionMgr = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsDistributionMgr = Nothing

    TrsRaiseError "busDistributionMgr.cbsDistributionMgr", "IsServiceRequest", , , , , True

End Function


Public Function GetAccountDescription ( _
	ByVal v_objRsPickListDistribution As ADODB.Recordset,  _
	ByVal v_lngWgNum As Long,  _
	ByVal v_strDistributionId As String,  _
	ByVal v_strDistributionType As String,  _
	ByVal v_strDistributionDeptNum As String _
	)  As  String
  
    Dim l_objcbsDistributionMgr As busDistributionMgr.cbsDistributionMgr
    
On Error GoTo eh

    Set l_objcbsDistributionMgr = CreateObject("busDistributionMgr.cbsDistributionMgr")
    
    GetAccountDescription = l_objcbsDistributionMgr.GetAccountDescription(v_objRsPickListDistribution, v_lngWgNum, v_strDistributionId, v_strDistributionType, v_strDistributionDeptNum)
    
    Set l_objcbsDistributionMgr = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsDistributionMgr = Nothing

    TrsRaiseError "busDistributionMgr.cbsDistributionMgr", "GetAccountDescription", , , , , True

End Function


Public Function GetAccountDescription2 ( _
	ByVal v_objRsPickListDistribution As ADODB.Recordset,  _
	ByVal v_lngWgNum As Long,  _
	ByVal v_strDistributionId As String,  _
	ByVal v_strDistributionType As String,  _
	ByVal v_strDistributionDeptNum As String,  _
	ByVal v_intDescriptionLength As Variant _
	)  As  String
  
    Dim l_objcbsDistributionMgr As busDistributionMgr.cbsDistributionMgr
    
On Error GoTo eh

    Set l_objcbsDistributionMgr = CreateObject("busDistributionMgr.cbsDistributionMgr")
    
    GetAccountDescription2 = l_objcbsDistributionMgr.GetAccountDescription2(v_objRsPickListDistribution, v_lngWgNum, v_strDistributionId, v_strDistributionType, v_strDistributionDeptNum, v_intDescriptionLength)
    
    Set l_objcbsDistributionMgr = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsDistributionMgr = Nothing

    TrsRaiseError "busDistributionMgr.cbsDistributionMgr", "GetAccountDescription2", , , , , True

End Function


Public Function GetAccountDescriptionJobTask ( _
	ByVal v_objRsPickListDistribution As ADODB.Recordset,  _
	ByVal v_lngWgNum As Long,  _
	ByVal v_strDistributionId As String,  _
	ByVal v_strDistributionType As String,  _
	ByVal v_strJobTask As String,  _
	ByVal v_strEMPRVDBid As String,  _
	ByVal v_strDistributionDeptNum As String _
	)  As  String
  
    Dim l_objcbsDistributionMgr As busDistributionMgr.cbsDistributionMgr
    
On Error GoTo eh

    Set l_objcbsDistributionMgr = CreateObject("busDistributionMgr.cbsDistributionMgr")
    
    GetAccountDescriptionJobTask = l_objcbsDistributionMgr.GetAccountDescriptionJobTask(v_objRsPickListDistribution, v_lngWgNum, v_strDistributionId, v_strDistributionType, v_strJobTask, v_strEMPRVDBid, v_strDistributionDeptNum)
    
    Set l_objcbsDistributionMgr = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsDistributionMgr = Nothing

    TrsRaiseError "busDistributionMgr.cbsDistributionMgr", "GetAccountDescriptionJobTask", , , , , True

End Function


Public Function GetAccountDescriptionJobTask2 ( _
	ByVal v_objRsPickListDistribution As ADODB.Recordset,  _
	ByVal v_lngWgNum As Long,  _
	ByVal v_strDistributionId As String,  _
	ByVal v_strDistributionType As String,  _
	ByVal v_strJobTask As String,  _
	ByVal v_strEMPRVDBid As String,  _
	ByVal v_strDistributionDeptNum As String,  _
	ByVal v_intDescriptionLength As Integer _
	)  As  String
  
    Dim l_objcbsDistributionMgr As busDistributionMgr.cbsDistributionMgr
    
On Error GoTo eh

    Set l_objcbsDistributionMgr = CreateObject("busDistributionMgr.cbsDistributionMgr")
    
    GetAccountDescriptionJobTask2 = l_objcbsDistributionMgr.GetAccountDescriptionJobTask2(v_objRsPickListDistribution, v_lngWgNum, v_strDistributionId, v_strDistributionType, v_strJobTask, v_strEMPRVDBid, v_strDistributionDeptNum, v_intDescriptionLength)
    
    Set l_objcbsDistributionMgr = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsDistributionMgr = Nothing

    TrsRaiseError "busDistributionMgr.cbsDistributionMgr", "GetAccountDescriptionJobTask2", , , , , True

End Function


Public Function GetWGDistributionIDDescriptions ( _
	ByVal v_colWgNum As Collection _
	)  As  ADODB.Recordset
  
    Dim l_objcbsDistributionMgr As busDistributionMgr.cbsDistributionMgr
    
On Error GoTo eh

    Set l_objcbsDistributionMgr = CreateObject("busDistributionMgr.cbsDistributionMgr")
    
    Set GetWGDistributionIDDescriptions = l_objcbsDistributionMgr.GetWGDistributionIDDescriptions(v_colWgNum)
    
    Set l_objcbsDistributionMgr = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsDistributionMgr = Nothing

    TrsRaiseError "busDistributionMgr.cbsDistributionMgr", "GetWGDistributionIDDescriptions", , , , , True

End Function


Public Function GetWGDistributionIDDescriptions2B ( _
	ByVal v_colWgNum As Collection,  _
	ByVal v_intDescriptionLength As Integer _
	)  As  ADODB.Recordset
  
    Dim l_objcbsDistributionMgr As busDistributionMgr.cbsDistributionMgr
    
On Error GoTo eh

    Set l_objcbsDistributionMgr = CreateObject("busDistributionMgr.cbsDistributionMgr")
    
    Set GetWGDistributionIDDescriptions2B = l_objcbsDistributionMgr.GetWGDistributionIDDescriptions2B(v_colWgNum, v_intDescriptionLength)
    
    Set l_objcbsDistributionMgr = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsDistributionMgr = Nothing

    TrsRaiseError "busDistributionMgr.cbsDistributionMgr", "GetWGDistributionIDDescriptions2B", , , , , True

End Function


Public Function GetWGDistributionIDDescriptions3 ( _
	ByVal v_colWgNum As Collection,  _
	ByVal v_intDescriptionLength As Integer,  _
	ByVal v_sCallerMode As String _
	)  As  ADODB.Recordset
  
    Dim l_objcbsDistributionMgr As busDistributionMgr.cbsDistributionMgr
    
On Error GoTo eh

    Set l_objcbsDistributionMgr = CreateObject("busDistributionMgr.cbsDistributionMgr")
    
    Set GetWGDistributionIDDescriptions3 = l_objcbsDistributionMgr.GetWGDistributionIDDescriptions3(v_colWgNum, v_intDescriptionLength, v_sCallerMode)
    
    Set l_objcbsDistributionMgr = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsDistributionMgr = Nothing

    TrsRaiseError "busDistributionMgr.cbsDistributionMgr", "GetWGDistributionIDDescriptions3", , , , , True

End Function


Public Function GetDOJMWorkRequestsByRegion ( _
	ByVal v_strDistrict As String,  _
	ByVal v_strHeadquarters As String _
	)  As  ADODB.Recordset
  
    Dim l_objcbsDistributionMgr As busDistributionMgr.cbsDistributionMgr
    
On Error GoTo eh

    Set l_objcbsDistributionMgr = CreateObject("busDistributionMgr.cbsDistributionMgr")
    
    Set GetDOJMWorkRequestsByRegion = l_objcbsDistributionMgr.GetDOJMWorkRequestsByRegion(v_strDistrict, v_strHeadquarters)
    
    Set l_objcbsDistributionMgr = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsDistributionMgr = Nothing

    TrsRaiseError "busDistributionMgr.cbsDistributionMgr", "GetDOJMWorkRequestsByRegion", , , , , True

End Function


Public Function GetDOJMWorkRequestsByWorkGroup ( _
	ByVal v_lngWgNum As Long,  _
	ByRef r_strDistrict As String,  _
	ByRef r_strHeadquarters As String _
	)  As  ADODB.Recordset
  
    Dim l_objcbsDistributionMgr As busDistributionMgr.cbsDistributionMgr
    
On Error GoTo eh

    Set l_objcbsDistributionMgr = CreateObject("busDistributionMgr.cbsDistributionMgr")
    
    Set GetDOJMWorkRequestsByWorkGroup = l_objcbsDistributionMgr.GetDOJMWorkRequestsByWorkGroup(v_lngWgNum, r_strDistrict, r_strHeadquarters)
    
    Set l_objcbsDistributionMgr = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsDistributionMgr = Nothing

    TrsRaiseError "busDistributionMgr.cbsDistributionMgr", "GetDOJMWorkRequestsByWorkGroup", , , , , True

End Function


Public Function IsDOJMWorkGroup ( _
	ByVal v_lngWgNum As Long _
	)  As  Boolean
  
    Dim l_objcbsDistributionMgr As busDistributionMgr.cbsDistributionMgr
    
On Error GoTo eh

    Set l_objcbsDistributionMgr = CreateObject("busDistributionMgr.cbsDistributionMgr")
    
    IsDOJMWorkGroup = l_objcbsDistributionMgr.IsDOJMWorkGroup(v_lngWgNum)
    
    Set l_objcbsDistributionMgr = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsDistributionMgr = Nothing

    TrsRaiseError "busDistributionMgr.cbsDistributionMgr", "IsDOJMWorkGroup", , , , , True

End Function


Public Function GetDistributionList ( _
	ByVal v_lngWgNum As Long,  _
	Optional ByRef r_strDistrict As String,  _
	Optional ByRef r_strHeadquarters As String _
	)  As  ADODB.Recordset
  
    Dim l_objcbsDistributionMgr As busDistributionMgr.cbsDistributionMgr
    
On Error GoTo eh

    Set l_objcbsDistributionMgr = CreateObject("busDistributionMgr.cbsDistributionMgr")
    
    Set GetDistributionList = l_objcbsDistributionMgr.GetDistributionList(v_lngWgNum, r_strDistrict, r_strHeadquarters)
    
    Set l_objcbsDistributionMgr = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsDistributionMgr = Nothing

    TrsRaiseError "busDistributionMgr.cbsDistributionMgr", "GetDistributionList", , , , , True

End Function


Public Function GetDistributionList2 ( _
	ByVal v_lngWgNum As Long,  _
	Optional ByRef r_strDistrict As String,  _
	Optional ByRef r_strHeadquarters As String _
	)  As  ADODB.Recordset
  
    Dim l_objcbsDistributionMgr As busDistributionMgr.cbsDistributionMgr
    
On Error GoTo eh

    Set l_objcbsDistributionMgr = CreateObject("busDistributionMgr.cbsDistributionMgr")
    
    Set GetDistributionList2 = l_objcbsDistributionMgr.GetDistributionList2(v_lngWgNum, r_strDistrict, r_strHeadquarters)
    
    Set l_objcbsDistributionMgr = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsDistributionMgr = Nothing

    TrsRaiseError "busDistributionMgr.cbsDistributionMgr", "GetDistributionList2", , , , , True

End Function


Public Function GetDistributionList2B ( _
	ByVal v_lngWgNum As Long,  _
	ByVal v_intDescriptionLength As Integer,  _
	Optional ByRef r_strDistrict As String,  _
	Optional ByRef r_strHeadquarters As String _
	)  As  ADODB.Recordset
  
    Dim l_objcbsDistributionMgr As busDistributionMgr.cbsDistributionMgr
    
On Error GoTo eh

    Set l_objcbsDistributionMgr = CreateObject("busDistributionMgr.cbsDistributionMgr")
    
    Set GetDistributionList2B = l_objcbsDistributionMgr.GetDistributionList2B(v_lngWgNum, v_intDescriptionLength, r_strDistrict, r_strHeadquarters)
    
    Set l_objcbsDistributionMgr = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsDistributionMgr = Nothing

    TrsRaiseError "busDistributionMgr.cbsDistributionMgr", "GetDistributionList2B", , , , , True

End Function


Public Function GetDistributionListReverse ( _
	ByVal v_lngWgNum As Long,  _
	Optional ByRef r_strDistrict As String,  _
	Optional ByRef r_strHeadquarters As String _
	)  As  ADODB.Recordset
  
    Dim l_objcbsDistributionMgr As busDistributionMgr.cbsDistributionMgr
    
On Error GoTo eh

    Set l_objcbsDistributionMgr = CreateObject("busDistributionMgr.cbsDistributionMgr")
    
    Set GetDistributionListReverse = l_objcbsDistributionMgr.GetDistributionListReverse(v_lngWgNum, r_strDistrict, r_strHeadquarters)
    
    Set l_objcbsDistributionMgr = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsDistributionMgr = Nothing

    TrsRaiseError "busDistributionMgr.cbsDistributionMgr", "GetDistributionListReverse", , , , , True

End Function


Public Function GetDistributionListReverse2 ( _
	ByVal v_lngWgNum As Long,  _
	Optional ByRef r_strDistrict As String,  _
	Optional ByRef r_strHeadquarters As String _
	)  As  ADODB.Recordset
  
    Dim l_objcbsDistributionMgr As busDistributionMgr.cbsDistributionMgr
    
On Error GoTo eh

    Set l_objcbsDistributionMgr = CreateObject("busDistributionMgr.cbsDistributionMgr")
    
    Set GetDistributionListReverse2 = l_objcbsDistributionMgr.GetDistributionListReverse2(v_lngWgNum, r_strDistrict, r_strHeadquarters)
    
    Set l_objcbsDistributionMgr = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsDistributionMgr = Nothing

    TrsRaiseError "busDistributionMgr.cbsDistributionMgr", "GetDistributionListReverse2", , , , , True

End Function


Public Function GetDistributionListReverse2B ( _
	ByVal v_lngWgNum As Long,  _
	ByVal v_intDescriptionLength As Integer,  _
	Optional ByRef r_strDistrict As String,  _
	Optional ByRef r_strHeadquarters As String _
	)  As  ADODB.Recordset
  
    Dim l_objcbsDistributionMgr As busDistributionMgr.cbsDistributionMgr
    
On Error GoTo eh

    Set l_objcbsDistributionMgr = CreateObject("busDistributionMgr.cbsDistributionMgr")
    
    Set GetDistributionListReverse2B = l_objcbsDistributionMgr.GetDistributionListReverse2B(v_lngWgNum, v_intDescriptionLength, r_strDistrict, r_strHeadquarters)
    
    Set l_objcbsDistributionMgr = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsDistributionMgr = Nothing

    TrsRaiseError "busDistributionMgr.cbsDistributionMgr", "GetDistributionListReverse2B", , , , , True

End Function


Public Function GetDistributionListDOJM ( _
	ByVal v_lngWgNum As Long,  _
	Optional ByRef r_strDistrict As String,  _
	Optional ByRef r_strHeadquarters As String _
	)  As  ADODB.Recordset
  
    Dim l_objcbsDistributionMgr As busDistributionMgr.cbsDistributionMgr
    
On Error GoTo eh

    Set l_objcbsDistributionMgr = CreateObject("busDistributionMgr.cbsDistributionMgr")
    
    Set GetDistributionListDOJM = l_objcbsDistributionMgr.GetDistributionListDOJM(v_lngWgNum, r_strDistrict, r_strHeadquarters)
    
    Set l_objcbsDistributionMgr = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsDistributionMgr = Nothing

    TrsRaiseError "busDistributionMgr.cbsDistributionMgr", "GetDistributionListDOJM", , , , , True

End Function


Public Function GetDistributionListEMPRV ( _
	ByVal v_lngWgNum As Long,  _
	ByVal v_blnReverseJobTasks As Boolean _
	)  As  ADODB.Recordset
  
    Dim l_objcbsDistributionMgr As busDistributionMgr.cbsDistributionMgr
    
On Error GoTo eh

    Set l_objcbsDistributionMgr = CreateObject("busDistributionMgr.cbsDistributionMgr")
    
    Set GetDistributionListEMPRV = l_objcbsDistributionMgr.GetDistributionListEMPRV(v_lngWgNum, v_blnReverseJobTasks)
    
    Set l_objcbsDistributionMgr = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsDistributionMgr = Nothing

    TrsRaiseError "busDistributionMgr.cbsDistributionMgr", "GetDistributionListEMPRV", , , , , True

End Function


Public Function GetDistributionListEMPRV2 ( _
	ByVal v_lngWgNum As Long,  _
	ByVal v_blnReverseJobTasks As Boolean _
	)  As  ADODB.Recordset
  
    Dim l_objcbsDistributionMgr As busDistributionMgr.cbsDistributionMgr
    
On Error GoTo eh

    Set l_objcbsDistributionMgr = CreateObject("busDistributionMgr.cbsDistributionMgr")
    
    Set GetDistributionListEMPRV2 = l_objcbsDistributionMgr.GetDistributionListEMPRV2(v_lngWgNum, v_blnReverseJobTasks)
    
    Set l_objcbsDistributionMgr = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsDistributionMgr = Nothing

    TrsRaiseError "busDistributionMgr.cbsDistributionMgr", "GetDistributionListEMPRV2", , , , , True

End Function


Public Function GetDistributionListEMPRV2B ( _
	ByVal v_lngWgNum As Long,  _
	ByVal v_blnReverseJobTasks As Boolean,  _
	ByVal v_intDescriptionLength As Integer _
	)  As  ADODB.Recordset
  
    Dim l_objcbsDistributionMgr As busDistributionMgr.cbsDistributionMgr
    
On Error GoTo eh

    Set l_objcbsDistributionMgr = CreateObject("busDistributionMgr.cbsDistributionMgr")
    
    Set GetDistributionListEMPRV2B = l_objcbsDistributionMgr.GetDistributionListEMPRV2B(v_lngWgNum, v_blnReverseJobTasks, v_intDescriptionLength)
    
    Set l_objcbsDistributionMgr = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsDistributionMgr = Nothing

    TrsRaiseError "busDistributionMgr.cbsDistributionMgr", "GetDistributionListEMPRV2B", , , , , True

End Function


Public Sub GetECSupportInformation ( _
	ByRef r_rsAllWorkNumbers As ADODB.Recordset,  _
	ByRef r_rsAllClassNumbers As ADODB.Recordset,  _
	ByRef r_rsAllDrawingSizes As ADODB.Recordset,  _
	ByRef r_rsAllWorkReasons As ADODB.Recordset,  _
	ByRef r_strECAccountNumNonProdTime As String,  _
	ByRef r_strECDescriptionNonProdTime As String _
	)  
  
    Dim l_objcbsDistributionMgr As busDistributionMgr.cbsDistributionMgr
    
On Error GoTo eh

    Set l_objcbsDistributionMgr = CreateObject("busDistributionMgr.cbsDistributionMgr")
    
    l_objcbsDistributionMgr.GetECSupportInformation r_rsAllWorkNumbers, r_rsAllClassNumbers, r_rsAllDrawingSizes, r_rsAllWorkReasons, r_strECAccountNumNonProdTime, r_strECDescriptionNonProdTime
    
    Set l_objcbsDistributionMgr = Nothing
    
    Exit Sub
    
eh:
    Set l_objcbsDistributionMgr = Nothing

    TrsRaiseError "busDistributionMgr.cbsDistributionMgr", "GetECSupportInformation", , , , , True

End Sub


Public Sub GetAllECInformation ( _
	ByRef r_rsECClassNumbers As ADODB.Recordset,  _
	ByRef r_rsECWorkNumbers As ADODB.Recordset,  _
	ByRef r_rsECDrawingSizes As ADODB.Recordset,  _
	ByRef r_rsECWorkReasons As ADODB.Recordset _
	)  
  
    Dim l_objcbsDistributionMgr As busDistributionMgr.cbsDistributionMgr
    
On Error GoTo eh

    Set l_objcbsDistributionMgr = CreateObject("busDistributionMgr.cbsDistributionMgr")
    
    l_objcbsDistributionMgr.GetAllECInformation r_rsECClassNumbers, r_rsECWorkNumbers, r_rsECDrawingSizes, r_rsECWorkReasons
    
    Set l_objcbsDistributionMgr = Nothing
    
    Exit Sub
    
eh:
    Set l_objcbsDistributionMgr = Nothing

    TrsRaiseError "busDistributionMgr.cbsDistributionMgr", "GetAllECInformation", , , , , True

End Sub


Public Function ValidateDistributionIdsFromTimeEntries ( _
	ByVal v_rsTimeEntries As ADODB.Recordset,  _
	ByVal v_rsOtherPayEntries As ADODB.Recordset,  _
	ByVal v_rsEmployeeData As ADODB.Recordset,  _
	ByVal v_rsEmployeeWGData As ADODB.Recordset,  _
	ByVal v_datStartDate As Date,  _
	ByVal v_datEndDate As Date,  _
	ByRef r_rsBRVs As ADODB.Recordset _
	)  As  Boolean
  
    Dim l_objcbsDistributionMgr As busDistributionMgr.cbsDistributionMgr
    
On Error GoTo eh

    Set l_objcbsDistributionMgr = CreateObject("busDistributionMgr.cbsDistributionMgr")
    
    ValidateDistributionIdsFromTimeEntries = l_objcbsDistributionMgr.ValidateDistributionIdsFromTimeEntries(v_rsTimeEntries, v_rsOtherPayEntries, v_rsEmployeeData, v_rsEmployeeWGData, v_datStartDate, v_datEndDate, r_rsBRVs)
    
    Set l_objcbsDistributionMgr = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsDistributionMgr = Nothing

    TrsRaiseError "busDistributionMgr.cbsDistributionMgr", "ValidateDistributionIdsFromTimeEntries", , , , , True

End Function


Public Function ValidateDistributionIdsFromTimeEntries2 ( _
	ByVal v_rsTimeEntries As ADODB.Recordset,  _
	ByVal v_rsOtherPayEntries As ADODB.Recordset,  _
	ByVal v_rsEmployeeData As ADODB.Recordset,  _
	ByVal v_rsEmployeeWGData As ADODB.Recordset,  _
	ByVal v_datStartDate As Date,  _
	ByVal v_datEndDate As Date,  _
	ByRef r_rsBRVs As ADODB.Recordset,  _
	ByVal v_rsJobTaskSteps As ADODB.Recordset _
	)  As  Boolean
  
    Dim l_objcbsDistributionMgr As busDistributionMgr.cbsDistributionMgr
    
On Error GoTo eh

    Set l_objcbsDistributionMgr = CreateObject("busDistributionMgr.cbsDistributionMgr")
    
    ValidateDistributionIdsFromTimeEntries2 = l_objcbsDistributionMgr.ValidateDistributionIdsFromTimeEntries2(v_rsTimeEntries, v_rsOtherPayEntries, v_rsEmployeeData, v_rsEmployeeWGData, v_datStartDate, v_datEndDate, r_rsBRVs, v_rsJobTaskSteps)
    
    Set l_objcbsDistributionMgr = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsDistributionMgr = Nothing

    TrsRaiseError "busDistributionMgr.cbsDistributionMgr", "ValidateDistributionIdsFromTimeEntries2", , , , , True

End Function


Public Function ValidateDistributionIDs ( _
	ByVal v_strReasonForEdit As String,  _
	ByVal v_strCompanyCode As String,  _
	ByVal v_strDistributionIDs As Variant,  _
	ByRef r_strDistributionIDsReturned As Variant,  _
	ByRef r_strReturnCodes As Variant,  _
	ByRef r_strReturnMsgs As Variant,  _
	ByRef r_strHighlightFields As Variant,  _
	ByRef r_rsBRVs As ADODB.Recordset _
	)  As  Boolean
  
    Dim l_objcbsDistributionMgr As busDistributionMgr.cbsDistributionMgr
    
On Error GoTo eh

    Set l_objcbsDistributionMgr = CreateObject("busDistributionMgr.cbsDistributionMgr")
    
    ValidateDistributionIDs = l_objcbsDistributionMgr.ValidateDistributionIDs(v_strReasonForEdit, v_strCompanyCode, v_strDistributionIDs, r_strDistributionIDsReturned, r_strReturnCodes, r_strReturnMsgs, r_strHighlightFields, r_rsBRVs)
    
    Set l_objcbsDistributionMgr = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsDistributionMgr = Nothing

    TrsRaiseError "busDistributionMgr.cbsDistributionMgr", "ValidateDistributionIDs", , , , , True

End Function


Public Function ValidateDistributionId ( _
	ByVal v_strDistributionId As Variant,  _
	ByRef r_strReturnMsg As String,  _
	ByVal v_strCorporateCode As String,  _
	ByVal v_strDistributionType As String,  _
	ByVal v_strPayrollDept As String,  _
	ByVal v_lngChargeWGNum As Long,  _
	ByRef r_strDistributionIDReturned As Variant,  _
	ByRef r_strHighlightFields As String,  _
	Optional ByRef v_strReasonForEdit As String = "FE" _
	)  As  Boolean
  
    Dim l_objcbsDistributionMgr As busDistributionMgr.cbsDistributionMgr
    
On Error GoTo eh

    Set l_objcbsDistributionMgr = CreateObject("busDistributionMgr.cbsDistributionMgr")
    
    ValidateDistributionId = l_objcbsDistributionMgr.ValidateDistributionId(v_strDistributionId, r_strReturnMsg, v_strCorporateCode, v_strDistributionType, v_strPayrollDept, v_lngChargeWGNum, r_strDistributionIDReturned, r_strHighlightFields, v_strReasonForEdit)
    
    Set l_objcbsDistributionMgr = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsDistributionMgr = Nothing

    TrsRaiseError "busDistributionMgr.cbsDistributionMgr", "ValidateDistributionId", , , , , True

End Function


Public Function ValidateECDraftingAcctNum ( _
	ByVal v_strDistributionId As String,  _
	ByRef r_strReturnMsg As String,  _
	ByVal v_strCorporateCode As String,  _
	ByVal v_strPayrollDept As String,  _
	ByVal v_lngChargeWGNum As Long,  _
	ByRef r_strDistributionIDReturned As String,  _
	ByRef r_strHighlightFields As String,  _
	Optional ByRef v_strReasonForEdit As String = "FE" _
	)  As  Boolean
  
    Dim l_objcbsDistributionMgr As busDistributionMgr.cbsDistributionMgr
    
On Error GoTo eh

    Set l_objcbsDistributionMgr = CreateObject("busDistributionMgr.cbsDistributionMgr")
    
    ValidateECDraftingAcctNum = l_objcbsDistributionMgr.ValidateECDraftingAcctNum(v_strDistributionId, r_strReturnMsg, v_strCorporateCode, v_strPayrollDept, v_lngChargeWGNum, r_strDistributionIDReturned, r_strHighlightFields, v_strReasonForEdit)
    
    Set l_objcbsDistributionMgr = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsDistributionMgr = Nothing

    TrsRaiseError "busDistributionMgr.cbsDistributionMgr", "ValidateECDraftingAcctNum", , , , , True

End Function


Public Function ValidateECDraftingAcctNum2 ( _
	ByVal v_strDistributionId As String,  _
	ByVal v_strDistributionType As String,  _
	ByRef r_strReturnMsg As String,  _
	ByVal v_strCorporateCode As String,  _
	ByVal v_strPayrollDept As String,  _
	ByVal v_lngChargeWGNum As Long,  _
	ByRef r_strDistributionIDReturned As String,  _
	ByRef r_strHighlightFields As String,  _
	Optional ByRef v_strReasonForEdit As String = "FE" _
	)  As  Boolean
  
    Dim l_objcbsDistributionMgr As busDistributionMgr.cbsDistributionMgr
    
On Error GoTo eh

    Set l_objcbsDistributionMgr = CreateObject("busDistributionMgr.cbsDistributionMgr")
    
    ValidateECDraftingAcctNum2 = l_objcbsDistributionMgr.ValidateECDraftingAcctNum2(v_strDistributionId, v_strDistributionType, r_strReturnMsg, v_strCorporateCode, v_strPayrollDept, v_lngChargeWGNum, r_strDistributionIDReturned, r_strHighlightFields, v_strReasonForEdit)
    
    Set l_objcbsDistributionMgr = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsDistributionMgr = Nothing

    TrsRaiseError "busDistributionMgr.cbsDistributionMgr", "ValidateECDraftingAcctNum2", , , , , True

End Function