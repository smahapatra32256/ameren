VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "cbsClientInformation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Constant
Private Const MODULENAME As String = "busSecurity.cbsClientInformation"
Public Function MaintainTRISCustomerInfo( _
                ByVal v_sTRISCustomerID As String _
                , ByVal v_lCustomerNumber As Long _
                , ByVal v_sClientMachineID As String _
                , ByVal v_sRemoteServerName As String _
                , ByVal v_sClientVersionNumber As String _
                , ByVal v_sClientEdition As String) As Boolean
On Error GoTo MaintainTRISCustomerInfo_EH
Dim lobjDataAccess As dbSecurity.cdFunSecurity
Dim ladrsTRISCustInfo As ADODB.Recordset
Dim lsPrevSessTimeStamp As String, lsPrevSessCltMCID As String
Dim lsPrevSessRSN As String

    'Construct the Data Access Object
    Set lobjDataAccess = CreateObject( _
                            "dbSecurity.cdFunSecurity")
    'Call Data Access Object Method and assign return
    Set ladrsTRISCustInfo = lobjDataAccess.db_GetCustomerInfo( _
                            v_sTRISCustomerID)
    'Release Original
    Set lobjDataAccess = Nothing

    'Prepare Values
    'User ID is 18 characters long
    v_sTRISCustomerID = Left(v_sTRISCustomerID, 18)
    'Machine ID is 18 characters long
    v_sClientMachineID = Left(v_sClientMachineID, 18)
    'Remote Server Name is 18 characters long
    v_sRemoteServerName = Left(v_sRemoteServerName, 18)

    lsPrevSessTimeStamp = ""
    lsPrevSessRSN = ""
    lsPrevSessCltMCID = ""

    If Not (ladrsTRISCustInfo Is Nothing) Then
        If Not (ladrsTRISCustInfo.EOF And _
            ladrsTRISCustInfo.BOF) Then
            ladrsTRISCustInfo.MoveFirst

            lsPrevSessTimeStamp = ladrsTRISCustInfo!CURR_TIMESTAMP
            lsPrevSessCltMCID = ladrsTRISCustInfo!CURR_MACH_ID
            lsPrevSessRSN = ladrsTRISCustInfo!CURR_RMT_SRVR_NM

            'Prepare these values properly
            'Time Stamp 26 characters long
            lsPrevSessTimeStamp = Left(lsPrevSessTimeStamp, 26)
            'Machine ID is 18 characters long
            lsPrevSessCltMCID = Left(lsPrevSessCltMCID, 18)
            'RSN is 18 characters long
            lsPrevSessRSN = Left(lsPrevSessRSN, 18)

            'Construct the Data Access Object
            Set lobjDataAccess = CreateObject( _
                                        "dbSecurity.cdFunSecurity")
            'Update Customer Record
            MaintainTRISCustomerInfo = _
                        lobjDataAccess.db_UpdCustomerInfo( _
                                v_sTRISCustomerID _
                                , v_lCustomerNumber _
                                , v_sClientMachineID _
                                , v_sRemoteServerName _
                                , lsPrevSessTimeStamp _
                                , lsPrevSessCltMCID _
                                , lsPrevSessRSN _
                                , v_sClientVersionNumber _
                                , v_sClientEdition)
            'Release
            Set lobjDataAccess = Nothing
        Else
            'Construct the Data Access Object
            Set lobjDataAccess = CreateObject( _
                                        "dbSecurity.cdFunSecurity")
            'Create a new record for customer
            MaintainTRISCustomerInfo = _
                        lobjDataAccess.db_InsCustomerInfo( _
                                v_sTRISCustomerID _
                                , v_lCustomerNumber _
                                , v_sClientMachineID _
                                , v_sRemoteServerName _
                                , v_sClientVersionNumber _
                                , v_sClientEdition)
            'Release
            Set lobjDataAccess = Nothing
        End If
    Else
        'Construct the Data Access Object
        Set lobjDataAccess = CreateObject( _
                                    "dbSecurity.cdFunSecurity")
        'Create a new record for customer
        MaintainTRISCustomerInfo = _
                        lobjDataAccess.db_InsCustomerInfo( _
                            v_sTRISCustomerID _
                            , v_lCustomerNumber _
                            , v_sClientMachineID _
                            , v_sRemoteServerName _
                            , v_sClientVersionNumber _
                            , v_sClientEdition)
        'Release
        Set lobjDataAccess = Nothing
    End If

    Set ladrsTRISCustInfo = Nothing

    GetObjectContext.SetComplete

    Exit Function
MaintainTRISCustomerInfo_EH:
    GetObjectContext.SetAbort

    Set lobjDataAccess = Nothing
    Set ladrsTRISCustInfo = Nothing

    'Raise Error depending on the Error Number
    TrsRaiseError MODULENAME _
            , "MaintainTRISCustomerInfo"

End Function
