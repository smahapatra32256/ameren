VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "cbsFunSecurity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Project:   TRIS Conversion Phase 1
'Author:    Prem M. Chander
'Date:      December 1999 (Firstcut)
'Description: This class module will reside in a Server Package
'under Microsoft Transaction Server (MTS). This module will be called by
'the client. This module provides the client with availability information, user existence,
'authorized business functions and other system support information. This module will call its associated
'data access component to retrieve pertinant information from the database.

Option Explicit
'Constant
Private Const MODULENAME As String = "busSecurity.cbsFunSecurity"
'This function will call functions to see if TRIS is Available for the User and
'check for the existence of the User
Public Function ApplyFunctionalSecurity( _
                ByVal v_bDeveloper As Boolean _
                , ByVal v_lEmpNum As Long _
                , ByRef r_adrsEmployee As ADODB.Recordset _
                , ByRef r_adrsABF As ADODB.Recordset _
                , ByRef r_bIsAvailable As Boolean _
                , Optional ByVal v_sTRISCustomerID As String _
                , Optional ByVal v_sClientMachineID As String _
                , Optional ByVal v_sClientVersionNum As String _
                , Optional ByVal v_sClientEdition As String _
                ) As Boolean
On Error GoTo ApplyFunctionalSecurity_EH
Dim llSecLevel As Long, lsServerName As String
Dim lsTmpStr As String

    'The Employee is not a developer then check for TRIS Availability
    If (v_bDeveloper = False) Then
        'Call to see if TRIS is Available
        r_bIsAvailable = TRISAvailable()
        
        If (r_bIsAvailable) Then
            'Check Employee Existence
            Set r_adrsEmployee = VerifyUserExistence(v_lEmpNum)

            If Not (r_adrsEmployee Is Nothing) Then
                'If Employee Exists
                If (r_adrsEmployee.RecordCount = 1) Then
                    If Not (CheckTRISRelease(r_adrsEmployee!Release _
                            , CDate(Format(r_adrsEmployee!Release_EFF_DATE, "mm/dd/yyyy")))) Then
                        Err.Raise trErrInValidTRISRelease
                    End If

                    'Get Security Level
                    llSecLevel = r_adrsEmployee!SEC_LEVEL
                    'Get Authorized Business Functions
                    Set r_adrsABF = GetAuthorizedBusFuns(llSecLevel)
                Else
                    Set r_adrsEmployee = Nothing
                    Set r_adrsABF = Nothing
                End If
            Else
                Set r_adrsEmployee = Nothing
                Set r_adrsABF = Nothing
            End If
        End If
    Else
        'Check Employee Existence
        Set r_adrsEmployee = VerifyUserExistence(v_lEmpNum)
        If Not (r_adrsEmployee Is Nothing) Then

            'If Employee Exists
            If Not (r_adrsEmployee.EOF And _
                r_adrsEmployee.BOF) Then
                r_adrsEmployee.MoveFirst

                lsTmpStr = ParseCustomerID(v_sTRISCustomerID)
                If (lsTmpStr <> "" And _
                    lsTmpStr <> CStr(v_lEmpNum)) Then
                    If Not (CheckTRISRelease(r_adrsEmployee!Release _
                            , CDate(Format(r_adrsEmployee!Release_EFF_DATE, "mm/dd/yyyy")))) Then
                        Err.Raise trErrInValidTRISRelease
                    End If
                End If

                'Get Security Level
                llSecLevel = r_adrsEmployee!SEC_LEVEL
                'Get Authorized Business Functions
                Set r_adrsABF = GetAuthorizedBusFuns(llSecLevel)
            Else
                Set r_adrsEmployee = Nothing
                Set r_adrsABF = Nothing
            End If
        Else
            Set r_adrsEmployee = Nothing
            Set r_adrsABF = Nothing
        End If
    End If

    'Record TRIS Customer Information
    lsServerName = Environ("COMPUTERNAME")

    '//-----include the COM+ Partition name in the remote server name column-----//
    Dim o As COMAdmin.COMAdminCatalog

    Set o = CreateObject("COMAdmin.COMAdminCatalog")
    If o.CurrentPartitionName = "Base Application Partition" Then
        lsServerName = lsServerName & ":BASE"
    Else
        lsServerName = lsServerName & ":" & Replace(o.CurrentPartitionName, "TRISPART", "")
    End If
    Set o = Nothing

    If (TRISCustomerInfo( _
        v_sTRISCustomerID _
        , v_lEmpNum _
        , v_sClientMachineID _
        , lsServerName _
        , v_sClientVersionNum _
        , v_sClientEdition)) Then
        ApplyFunctionalSecurity = True
    Else
        Err.Raise trErrApplyFunSecurity_F
    End If

    'Return the Current Date to the Client for
    'Department Effective Date Check
    'r_dtServerDate = CDate(Format(Now, "mm/dd/yyyy"))

    Exit Function
ApplyFunctionalSecurity_EH:
    'Destroy objects
    Set r_adrsEmployee = Nothing
    Set r_adrsABF = Nothing
    r_bIsAvailable = False

    'Raise Error depending on the Error Number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "ApplyFunctionalSecurity"

            Case trErrdb_CheckEmployee_F, trErrdb_TRISAvailable_F _
                , trErrdb_GetMenuItems_F, trErrdb_GetApplData_F _
                , trErrdb_ExecuteSQL_F _
                , trErrTRISAvailable_F, trErrVerifyUserExistence_F _
                , trErrGetAuthorizedBusFuns_F, trErrGetBusinessData_F _
                , trErrSpecialSecurity_F, trErrGetClientInfo _
                , trErrdb_GetCltInfo, trErrdb_InsCltInfo _
                , trErrdb_UpdCltInfo, trErrInValidTRISRelease
                TrsRaiseError MODULENAME _
                    , "ApplyFunctionalSecurity" _
                    , Err.Number _
                    , , , DescArgsFrom("v_lEmpNum=" & CStr(v_lEmpNum), "v_sTRISCustomerID=" & v_sTRISCustomerID), _
                    True

            Case Else
                TrsRaiseError MODULENAME _
                    , "ApplyFunctionalSecurity" _
                    , trErrApplyFunSecurity_F _
                    , , , , _
                    True
    End Select

End Function
'The function will return a boolean to the client depending on
'the availability of TRIS Database
Private Function TRISAvailable() As Boolean
On Error GoTo TRISAvailable_EH
'Data Access Object
Dim lobjDataAccess As dbSecurity.cdFunSecurity

    'Construct the Data Access Object
    Set lobjDataAccess = CreateObject( _
                            "dbSecurity.cdFunSecurity")
    'Call data access method
    TRISAvailable = lobjDataAccess.db_TRISAvailable
    'destroy
    Set lobjDataAccess = Nothing

    Exit Function
TRISAvailable_EH:
    Set lobjDataAccess = Nothing

    'Raise Error depending on the Error Number
    TrsRaiseError MODULENAME _
            , "TRISAvailable"

End Function
'The function returns an ADODB Recordset to the client if the employee exists
'in the TRIS employee database table.
Private Function VerifyUserExistence(ByVal v_lEmpNum As Long _
                    ) As ADODB.Recordset
On Error GoTo VerifyUserExistence_EH
'Data Access Object
Dim lobjDataAccess As dbSecurity.cdFunSecurity

    'Construct the Data Access Object
    Set lobjDataAccess = CreateObject( _
                            "dbSecurity.cdFunSecurity")
    'NOTE: It is assumed that a valid employee number will be passed
    'Call Data Access Object method and assign return
    Set VerifyUserExistence = lobjDataAccess.db_CheckEmployee(v_lEmpNum)
    'destroy
    Set lobjDataAccess = Nothing

    Exit Function
VerifyUserExistence_EH:
    'destroy
    Set lobjDataAccess = Nothing
    'just in case
    Set VerifyUserExistence = Nothing

    'Raise Error depending on the Error Number
    TrsRaiseError MODULENAME _
            , "VerifyUserExistence"

End Function
'This function returns all the authorized menu item (business functions)
'the user has access to.
Private Function GetAuthorizedBusFuns(ByVal v_lSecLevel As Long _
                    ) As ADODB.Recordset
On Error GoTo GetAuthorizedBusFuns_EH
'Data Access Object
Dim lobjDataAccess As dbSecurity.cdFunSecurity

    'Construct the Data Access Object
    Set lobjDataAccess = CreateObject( _
                            "dbSecurity.cdFunSecurity")
    'Call Data Access Object Method and assign return
    Set GetAuthorizedBusFuns = lobjDataAccess.db_GetMenuItems(v_lSecLevel)
    Set lobjDataAccess = Nothing

    Exit Function

GetAuthorizedBusFuns_EH:
    'Destroy Object
    Set lobjDataAccess = Nothing

    'just in case
    Set GetAuthorizedBusFuns = Nothing

    'Raise Error depending on the Error Number
    TrsRaiseError MODULENAME _
        , "GetAuthorizedBusFuns"
                
End Function
Private Function CheckTRISRelease(ByVal vsRelease As String _
                    , ByVal vdtReleaseEffDate As Date) As Boolean
On Error GoTo CheckTRISRelease_EH

    Select Case (Trim(vsRelease))
        Case "5.0.0"
            If (vdtReleaseEffDate > CDate(Format(Now, "mm/dd/yyyy"))) Then
                Err.Raise trErrInValidTRISRelease
            Else
                CheckTRISRelease = True
            End If
''''            Case "4.5.0" '//-----since this will be deployed July 29, this is to keep out CILCO users until Aug. 4, 2003
''''                If (vdtReleaseEffDate > CDate(Format(Now, "mm/dd/yyyy"))) Then
''''                    Err.Raise trErrInValidTRISRelease
''''                Else
''''                    CheckTRISRelease = True
''''                End If

        Case Else
            Err.Raise trErrInValidTRISRelease
    End Select

    Exit Function

CheckTRISRelease_EH:
    TrsRaiseError MODULENAME _
        , "CheckTRISRelease"

End Function
Private Function ParseCustomerID(ByVal v_sTRISCustomerID As String _
                    ) As String
On Error GoTo ParseCustomerID_EH
Dim llIndex As Long, lsNumber As String
    
    If (Trim(v_sTRISCustomerID) <> "") Then
        For llIndex = 1 To Len(v_sTRISCustomerID)
            If (IsNumeric(Mid(v_sTRISCustomerID, llIndex, 1))) Then
                'If numeric then build string
                lsNumber = lsNumber & Mid(v_sTRISCustomerID _
                                        , llIndex _
                                        , 1)
            End If
        Next llIndex
        ParseCustomerID = Trim(lsNumber)
    Else
        ParseCustomerID = ""
    End If

    Exit Function
ParseCustomerID_EH:
    TrsRaiseError MODULENAME _
        , "ParseCustomerID"

End Function
Private Function TRISCustomerInfo(ByVal v_sTRISCustomerID As String _
                    , ByVal v_lCustomerNumber As Long _
                    , ByVal v_sClientMachineID As String _
                    , ByVal v_sRemoteServerName As String _
                    , ByVal v_sClientVersionNumber As String _
                    , ByVal v_sClientEdition As String _
                    ) As Boolean
On Error GoTo TRISCustomerInfo_EH
Dim lobjClientInfo As busSecurity.cbsClientInformation

    'Construct the Data Access Object
    Set lobjClientInfo = CreateObject( _
                            "busSecurity.cbsClientInformation")
    'Call Data Access Object Method and assign return
    TRISCustomerInfo = lobjClientInfo.MaintainTRISCustomerInfo( _
                                v_sTRISCustomerID _
                                , v_lCustomerNumber _
                                , v_sClientMachineID _
                                , v_sRemoteServerName _
                                , v_sClientVersionNumber _
                                , v_sClientEdition)

    Set lobjClientInfo = Nothing

    Exit Function
TRISCustomerInfo_EH:
    'destroy
    Set lobjClientInfo = Nothing

    'Raise Error depending on the Error Number
    TrsRaiseError MODULENAME _
            , "TRISCustomerInfo"

End Function
Public Function RefreshSecurityInformation( _
                    ByVal v_lEmployee As Long _
                    , ByRef r_adrsEmployeeInfo As ADODB.Recordset _
                    , ByRef r_adrsAuthrzBusFunc As ADODB.Recordset _
                    ) As Boolean
On Error GoTo RefreshSecurityInformation_EH
Dim llSecurityLevel As Long

    RefreshSecurityInformation = False

    'Check Employee Existence
    Set r_adrsEmployeeInfo = VerifyUserExistence(v_lEmployee)
    If Not (r_adrsEmployeeInfo Is Nothing) Then
        'If Employee Exists
        If Not (r_adrsEmployeeInfo.EOF And _
            r_adrsEmployeeInfo.BOF) Then
            r_adrsEmployeeInfo.MoveFirst

            'Get Security Level
            llSecurityLevel = r_adrsEmployeeInfo!SEC_LEVEL

            'Get Authorized Business Functions
            Set r_adrsAuthrzBusFunc = GetAuthorizedBusFuns(llSecurityLevel)
        End If
    End If

    RefreshSecurityInformation = True

    Exit Function
RefreshSecurityInformation_EH:
    'Destroy objects
    Set r_adrsEmployeeInfo = Nothing
    Set r_adrsAuthrzBusFunc = Nothing

    'Raise Error depending on the Error Number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "RefreshSecurityInformation"

            Case trErrdb_CheckEmployee_F, trErrdb_GetMenuItems_F _
                , trErrdb_GetApplData_F, trErrdb_ExecuteSQL_F _
                , trErrVerifyUserExistence_F, trErrGetAuthorizedBusFuns_F _
                , trErrGetBusinessData_F, trErrSpecialSecurity_F
                TrsRaiseError MODULENAME _
                    , "RefreshSecurityInformation" _
                    , Err.Number _
                    , , , , _
                    True

            Case Else
                TrsRaiseError MODULENAME _
                    , "RefreshSecurityInformation" _
                    , trErrApplyFunSecurity_F _
                    , , , , _
                    True
    End Select

End Function
