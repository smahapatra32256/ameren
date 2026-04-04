VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "cbsSysAdminProxy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'//-----The following code was generated to be used by the Client only-----//




Public Function ProcessSysAdminRequest ( _
	ByVal vsRequest As String,  _
	ByVal vlUserNumber As Long,  _
	ByRef radrsMulti As Variant,  _
	ByRef rlRecordsAffected As Long,  _
	ByRef rsErrorString As String,  _
	Optional ByVal vsUserID As String = "" _
	)  As  Long
  
    Dim l_objcbsSysAdmin As busSysAdmin.cbsSysAdmin
    Dim l_radrsMulti() As ADODB.Recordset
    
On Error GoTo eh
    If Not IsEmpty(radrsMulti) Then
        l_radrsMulti = radrsMulti
    End If
    
    Set l_objcbsSysAdmin = CreateObject("busSysAdmin.cbsSysAdmin")
    
    ProcessSysAdminRequest = l_objcbsSysAdmin.ProcessSysAdminRequest(vsRequest, vlUserNumber, l_radrsMulti, rlRecordsAffected, rsErrorString, vsUserID)
    
    Set l_objcbsSysAdmin = Nothing
    radrsMulti = l_radrsMulti
    
    Exit Function
    
eh:
    Set l_objcbsSysAdmin = Nothing

    TrsRaiseError "busSysAdmin.cbsSysAdmin", "ProcessSysAdminRequest", , , , , True

End Function


Public Function ProcessSysAdminSELECT ( _
	ByVal vsRequest As String,  _
	ByVal vlUserNumber As Long,  _
	ByRef radrsMulti As Variant,  _
	ByRef rsErrorString As String,  _
	Optional ByVal vsUserID As String = "" _
	)  As  Long
  
    Dim l_objcbsSysAdmin As busSysAdmin.cbsSysAdmin
    Dim l_radrsMulti() As ADODB.Recordset
    
On Error GoTo eh
    If Not IsEmpty(radrsMulti) Then
        l_radrsMulti = radrsMulti
    End If
    
    Set l_objcbsSysAdmin = CreateObject("busSysAdmin.cbsSysAdmin")
    
    ProcessSysAdminSELECT = l_objcbsSysAdmin.ProcessSysAdminSELECT(vsRequest, vlUserNumber, l_radrsMulti, rsErrorString, vsUserID)
    
    Set l_objcbsSysAdmin = Nothing
    radrsMulti = l_radrsMulti
    
    Exit Function
    
eh:
    Set l_objcbsSysAdmin = Nothing

    TrsRaiseError "busSysAdmin.cbsSysAdmin", "ProcessSysAdminSELECT", , , , , True

End Function


Public Function ProcessSysAdminINSUPDDEL ( _
	ByVal vsRequest As String,  _
	ByVal vlUserNumber As Long,  _
	ByRef rsErrorString As String,  _
	Optional ByVal vsUserID As String = "" _
	)  As  Long
  
    Dim l_objcbsSysAdmin As busSysAdmin.cbsSysAdmin
    
On Error GoTo eh

    Set l_objcbsSysAdmin = CreateObject("busSysAdmin.cbsSysAdmin")
    
    ProcessSysAdminINSUPDDEL = l_objcbsSysAdmin.ProcessSysAdminINSUPDDEL(vsRequest, vlUserNumber, rsErrorString, vsUserID)
    
    Set l_objcbsSysAdmin = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsSysAdmin = Nothing

    TrsRaiseError "busSysAdmin.cbsSysAdmin", "ProcessSysAdminINSUPDDEL", , , , , True

End Function


Public Sub GetAllCatalogInformation ( _
	ByRef radrsTables As ADODB.Recordset,  _
	ByRef radrsColumns As ADODB.Recordset,  _
	ByRef radrsIndices As ADODB.Recordset,  _
	ByRef radrsKeys As ADODB.Recordset,  _
	ByVal vlUserNumber As Long,  _
	Optional ByVal vsUserID As String _
	)  
  
    Dim l_objcbsSysAdmin As busSysAdmin.cbsSysAdmin
    
On Error GoTo eh

    Set l_objcbsSysAdmin = CreateObject("busSysAdmin.cbsSysAdmin")
    
    l_objcbsSysAdmin.GetAllCatalogInformation radrsTables, radrsColumns, radrsIndices, radrsKeys, vlUserNumber, vsUserID
    
    Set l_objcbsSysAdmin = Nothing
    
    Exit Sub
    
eh:
    Set l_objcbsSysAdmin = Nothing

    TrsRaiseError "busSysAdmin.cbsSysAdmin", "GetAllCatalogInformation", , , , , True

End Sub


Public Function GetTableCatalogInformation ( _
	)  As  ADODB.Recordset
  
    Dim l_objcbsSysAdmin As busSysAdmin.cbsSysAdmin
    
On Error GoTo eh

    Set l_objcbsSysAdmin = CreateObject("busSysAdmin.cbsSysAdmin")
    
    Set GetTableCatalogInformation = l_objcbsSysAdmin.GetTableCatalogInformation()
    
    Set l_objcbsSysAdmin = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsSysAdmin = Nothing

    TrsRaiseError "busSysAdmin.cbsSysAdmin", "GetTableCatalogInformation", , , , , True

End Function


Public Function GetColumnCatalogInformation ( _
	)  As  ADODB.Recordset
  
    Dim l_objcbsSysAdmin As busSysAdmin.cbsSysAdmin
    
On Error GoTo eh

    Set l_objcbsSysAdmin = CreateObject("busSysAdmin.cbsSysAdmin")
    
    Set GetColumnCatalogInformation = l_objcbsSysAdmin.GetColumnCatalogInformation()
    
    Set l_objcbsSysAdmin = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsSysAdmin = Nothing

    TrsRaiseError "busSysAdmin.cbsSysAdmin", "GetColumnCatalogInformation", , , , , True

End Function


Public Function GetIndexCatalogInformation ( _
	)  As  ADODB.Recordset
  
    Dim l_objcbsSysAdmin As busSysAdmin.cbsSysAdmin
    
On Error GoTo eh

    Set l_objcbsSysAdmin = CreateObject("busSysAdmin.cbsSysAdmin")
    
    Set GetIndexCatalogInformation = l_objcbsSysAdmin.GetIndexCatalogInformation()
    
    Set l_objcbsSysAdmin = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsSysAdmin = Nothing

    TrsRaiseError "busSysAdmin.cbsSysAdmin", "GetIndexCatalogInformation", , , , , True

End Function


Public Function GetIndexKeysCatalogInformation ( _
	)  As  ADODB.Recordset
  
    Dim l_objcbsSysAdmin As busSysAdmin.cbsSysAdmin
    
On Error GoTo eh

    Set l_objcbsSysAdmin = CreateObject("busSysAdmin.cbsSysAdmin")
    
    Set GetIndexKeysCatalogInformation = l_objcbsSysAdmin.GetIndexKeysCatalogInformation()
    
    Set l_objcbsSysAdmin = Nothing
    
    Exit Function
    
eh:
    Set l_objcbsSysAdmin = Nothing

    TrsRaiseError "busSysAdmin.cbsSysAdmin", "GetIndexKeysCatalogInformation", , , , , True

End Function