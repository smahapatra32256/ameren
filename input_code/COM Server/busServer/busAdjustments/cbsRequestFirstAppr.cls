VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "cbsRequestFirstAppr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Project:   TRIS Conversion Phase 1
'Author:    Prem M. Chander
'Date:      December 2000 (Firstcut)
'Description: This class module will reside in a Server Package
'under Microsoft Transaction Server (MTS). This module will be called by
'the client. This module provides the client with Adustment Information. It also
' processes any requests from the Client. This module will call its associated
'data access component to retrieve/modify pertinant information from the database.

Option Explicit

'Constant
Private Const MODULENAME As String = "busAdjustments.cbsRequestFirstAppr"
Public Function GetRequestForAdjustments( _
                ByVal v_lRequestedByEmployee As Long _
                ) As ADODB.Recordset
On Error GoTo GetRequestForAdjustments_EH
Dim lobjDataAccess As dbAdjustments.cdRequestFirstAppr

    Set GetRequestForAdjustments = Nothing

    'Construct the Data Access Object
    Set lobjDataAccess = CreateObject( _
                    "dbAdjustments.cdRequestFirstAppr")
    Set GetRequestForAdjustments = lobjDataAccess.db_GetAdjustments( _
                        ADJUSTMENTREQUESTFOR _
                        , v_lRequestedByEmployee)
    Set lobjDataAccess = Nothing

    'destroy objects
    Set lobjDataAccess = Nothing

    Exit Function
GetRequestForAdjustments_EH:
    Set lobjDataAccess = Nothing

    'Raise Error depending on the Error Number
    Select Case (Err.Number)
        Case trErrutOpenConnection
            TrsRaiseError MODULENAME _
                , "GetRequestForAdjustments"

        Case trErrdbGetReqAdj, trErrdbGetAdjEmpInfo _
            , trErrdb_ExecuteSQL_F
            TrsRaiseError MODULENAME _
                , "GetRequestForAdjustments" _
                , Err.Number _
                , , , , True

        Case Else
            TrsRaiseError MODULENAME _
                , "GetRequestForAdjustments" _
                , trErrGetRequestAdjust_F _
                , , , , True
    End Select

End Function
Public Function GetFirstApprovalAdjustments( _
                ByVal v_vArrayOfDepartments As Variant _
                ) As ADODB.Recordset
On Error GoTo GetFirstApprovalAdjustments_EH
Dim lsDepartments As String
Dim lobjDataAccess As dbAdjustments.cdRequestFirstAppr

    Set GetFirstApprovalAdjustments = Nothing

    'Prepare the data
    lsDepartments = PrepareDepartments(v_vArrayOfDepartments)

    'Construct the Data Access Object
    Set lobjDataAccess = CreateObject( _
                            "dbAdjustments.cdRequestFirstAppr")
    Set GetFirstApprovalAdjustments = lobjDataAccess.db_GetAdjustments( _
                            ADJUSTMENTFIRSTAPPROVAL _
                            , , lsDepartments)
    Set lobjDataAccess = Nothing

    'destroy objects
    Set lobjDataAccess = Nothing

    Exit Function
GetFirstApprovalAdjustments_EH:
    Set lobjDataAccess = Nothing

    'Raise Error depending on the Error Number
    Select Case Err.Number
        Case trErrutOpenConnection
            TrsRaiseError MODULENAME _
                , "GetFirstApprovalAdjustments"

        Case trErrdbGetFirstAppAdj, trErrdbGetAdjEmpInfo _
            , trErrdb_ExecuteSQL_F
            TrsRaiseError MODULENAME _
                , "GetFirstApprovalAdjustments" _
                , Err.Number _
                , , , , True

        Case Else
            TrsRaiseError MODULENAME _
                , "GetFirstApprovalAdjustments" _
                , trErrGetFirstAppAdjust_F _
                , , , , True
    End Select

End Function
Private Function PrepareDepartments(ByVal vvArrayOfDept As Variant _
                    ) As String
On Error GoTo PrepareDepartments_EH
Dim llIndex As Long, lsDepartments As String

    PrepareDepartments = ""

    If Not (IsEmpty(vvArrayOfDept)) Then
        For llIndex = LBound(vvArrayOfDept) To UBound(vvArrayOfDept)
            If (Trim(vvArrayOfDept(llIndex)) <> "") Then
                lsDepartments = lsDepartments & _
                        ", '" & Trim(vvArrayOfDept(llIndex)) & "'"
            End If
        Next llIndex
    End If

    If (Left(lsDepartments, 1) = ",") Then
        lsDepartments = Right(lsDepartments, Len(lsDepartments) - 1)
    End If

    PrepareDepartments = lsDepartments

    Exit Function
PrepareDepartments_EH:
    TrsRaiseError MODULENAME _
        , "PrepareDepartments"

End Function
'Begin Prem TRIS 5.6
Public Function GetFirstApprovalAdjustments2( _
                ByVal vlTRISUserNumber As Long _
                ) As ADODB.Recordset
On Error GoTo GetFirstApprovalAdjustments2_EH
Dim lsDepartments As String
Dim lobjDataAccess As dbAdjustments.cdRequestFirstAppr

    Set GetFirstApprovalAdjustments2 = Nothing

    'Construct the Data Access Object
    Set lobjDataAccess = CreateObject( _
                            "dbAdjustments.cdRequestFirstAppr")
    Set GetFirstApprovalAdjustments2 = lobjDataAccess.db_GetAdjustments( _
                            ADJUSTMENTFIRSTAPPROVAL _
                            , vlTRISUserNumber)
    Set lobjDataAccess = Nothing

    'destroy objects
    Set lobjDataAccess = Nothing

    Exit Function
GetFirstApprovalAdjustments2_EH:
    Set lobjDataAccess = Nothing

    'Raise Error depending on the Error Number
    Select Case Err.Number
        Case trErrutOpenConnection
            TrsRaiseError MODULENAME _
                , "GetFirstApprovalAdjustments2"

        Case trErrdbGetFirstAppAdj, trErrdbGetAdjEmpInfo _
            , trErrdb_ExecuteSQL_F
            TrsRaiseError MODULENAME _
                , "GetFirstApprovalAdjustments2" _
                , Err.Number _
                , , , , True

        Case Else
            TrsRaiseError MODULENAME _
                , "GetFirstApprovalAdjustments2" _
                , trErrGetFirstAppAdjust_F _
                , , , , True
    End Select

End Function
'End Prem TRIS 5.6

'Begin Prem TRIS 5.8
Public Sub GetAdjustmentsStatusAndDetails(ByVal vlTRISUserNumber As Long _
                    , ByVal vdtStartDate As Date _
                    , ByVal vdtEndDate As Date _
                    , ByRef radrsAdjustmentStatus As ADODB.Recordset _
                    , ByRef radrsAdjustmentDetails As ADODB.Recordset)
On Error GoTo GetAdjustmentsStatusAndDetails_EH
Dim lobjdb As dbAdjustments.cdAdjustments

    Set lobjdb = CreateObject("dbAdjustments.cdAdjustments")
    Set radrsAdjustmentStatus = lobjdb.GetAdjustmentsStatusByScope( _
                                vlTRISUserNumber _
                                , vdtStartDate _
                                , vdtEndDate)
    Set radrsAdjustmentDetails = lobjdb.GetAdjustmentsDetailsByScope( _
                                vlTRISUserNumber _
                                , vdtStartDate _
                                , vdtEndDate)
    Set lobjdb = Nothing

    Exit Sub
GetAdjustmentsStatusAndDetails_EH:
    Set lobjdb = Nothing

    'Raise Error depending on the Error Number
    Select Case Err.Number
        Case trErrutOpenConnection
            TrsRaiseError MODULENAME _
                , "GetAdjustmentsStatusAndDetails"

        Case trErrdb_ExecuteSQL_F
            TrsRaiseError MODULENAME _
                , "GetAdjustmentsStatusAndDetails" _
                , trErrdb_ExecuteSQL_F _
                , , , _
                , True

        Case Else
            TrsRaiseError MODULENAME _
                , "GetAdjustmentsStatusAndDetails" _
                , , , , _
                , True
    End Select

End Sub
Public Function GetAdjustmentsStatusByScope(ByVal vlTRISUserNumber As Long _
                    , ByVal vdtStartDate As Date _
                    , ByVal vdtEndDate As Date _
                    ) As ADODB.Recordset
On Error GoTo GetAdjustmentsStatusByScope_EH
Dim lobjdb As dbAdjustments.cdAdjustments

    Set lobjdb = CreateObject("dbAdjustments.cdAdjustments")
    Set GetAdjustmentsStatusByScope = lobjdb.GetAdjustmentsStatusByScope( _
                                vlTRISUserNumber _
                                , vdtStartDate _
                                , vdtEndDate)
    Set lobjdb = Nothing

    Exit Function
GetAdjustmentsStatusByScope_EH:
    Set lobjdb = Nothing

    'Raise Error depending on the Error Number
    Select Case Err.Number
        Case trErrutOpenConnection
            TrsRaiseError MODULENAME _
                , "GetAdjustmentsStatusByScope"

        Case trErrdb_ExecuteSQL_F
            TrsRaiseError MODULENAME _
                , "GetAdjustmentsStatusByScope" _
                , trErrdb_ExecuteSQL_F _
                , , , _
                , True

        Case Else
            TrsRaiseError MODULENAME _
                , "GetAdjustmentsStatusByScope" _
                , , , , _
                , True
    End Select

End Function
Public Function GetAdjustmentsDetailsByScope(ByVal vlTRISUserNumber As Long _
                    , ByVal vdtStartDate As Date _
                    , ByVal vdtEndDate As Date _
                    ) As ADODB.Recordset
On Error GoTo GetAdjustmentsDetailsByScope_EH
Dim lobjdb As dbAdjustments.cdAdjustments

    Set lobjdb = CreateObject("dbAdjustments.cdAdjustments")
    Set GetAdjustmentsDetailsByScope = lobjdb.GetAdjustmentsDetailsByScope( _
                                vlTRISUserNumber _
                                , vdtStartDate _
                                , vdtEndDate)
    Set lobjdb = Nothing

    Exit Function
GetAdjustmentsDetailsByScope_EH:
    Set lobjdb = Nothing

    'Raise Error depending on the Error Number
    Select Case Err.Number
        Case trErrutOpenConnection
            TrsRaiseError MODULENAME _
                , "GetAdjustmentsDetailsByScope"

        Case trErrdb_ExecuteSQL_F
            TrsRaiseError MODULENAME _
                , "GetAdjustmentsDetailsByScope" _
                , trErrdb_ExecuteSQL_F _
                , , , _
                , True

        Case Else
            TrsRaiseError MODULENAME _
                , "GetAdjustmentsDetailsByScope" _
                , , , , _
                , True
    End Select

End Function
'End Prem TRIS 5.8

'************* Automated Adjustment related additions **************

Public Function SaveAutomatedAdjustments( _
        ByVal vrsTimeEntriesRequest As ADODB.Recordset, _
        ByVal vrsOtherPay As ADODB.Recordset, _
        ByVal vStartDate As Date, _
        ByVal vEndDate As Date, _
        ByVal lrsEmployees As ADODB.Recordset, _
        ByVal lrsEmpWorkgroups As ADODB.Recordset, _
        ByVal vrsSchedules As ADODB.Recordset, _
        ByRef rrsErrors As ADODB.Recordset, _
        ByVal vUserId As String) As Boolean
        
                    
    On Error GoTo SaveAutomatedAdjustments_EH
    Dim retVal As Boolean
    retVal = False
    Set rrsErrors = NewErrorset

    Dim objAutomatedAdjustments As Ameren_TRIS_Server_COMIntegration.AutomatedAdjustments
    Set objAutomatedAdjustments = New Ameren_TRIS_Server_COMIntegration.AutomatedAdjustments
    retVal = objAutomatedAdjustments.Save(vrsTimeEntriesRequest, _
                vrsOtherPay, _
                vStartDate, _
                vEndDate, _
                lrsEmployees, _
                lrsEmpWorkgroups, _
                vrsSchedules, _
                rrsErrors, _
                vUserId)
        
    Set objAutomatedAdjustments = Nothing

    SaveAutomatedAdjustments = retVal
    
    Exit Function

SaveAutomatedAdjustments_EH:
    
    Set objAutomatedAdjustments = Nothing
    SaveAutomatedAdjustments = False
    
    TrsRaiseError MODULENAME _
                , "SaveAutomatedAdjustments" _
                , , , , _
                , True
                
End Function

Public Function GetTimeEntries(ByVal vlEmployeeNumber As Long _
                    , ByVal vdtStartDate As Date _
                    , ByVal vdtEndDate As Date _
                    ) As ADODB.Recordset
                    
  On Error GoTo GetTimeEntries_EH

    Dim objTime As Ameren_TRIS_Server_COMIntegration.Time
    Set objTime = New Ameren_TRIS_Server_COMIntegration.Time
    Set GetTimeEntries = objTime.GetTimeEntries(vlEmployeeNumber, vdtStartDate, vdtEndDate)
    Set objTime = Nothing

    Exit Function

GetTimeEntries_EH:
    
    Set objTime = Nothing
    TrsRaiseError MODULENAME _
                , "GetTimeEntries" _
                , , , , _
                , True
                    
End Function

Public Function GetOtherPayEntries(ByVal vlEmployeeNumber As Long _
                    , ByVal vdtStartDate As Date _
                    , ByVal vdtEndDate As Date _
                    ) As ADODB.Recordset

On Error GoTo GetOtherPayEntries_EH

    Dim objOtherPay As Ameren_TRIS_Server_COMIntegration.OtherPay
    Set objOtherPay = New Ameren_TRIS_Server_COMIntegration.OtherPay
    Set GetOtherPayEntries = objOtherPay.GetOtherPayEntries(vlEmployeeNumber, vdtStartDate, vdtEndDate)
    Set objOtherPay = Nothing

    Exit Function

GetOtherPayEntries_EH:
    
    Set objOtherPay = Nothing
    TrsRaiseError MODULENAME _
                , "GetOtherPayEntries" _
                , , , , _
                , True

End Function

Public Function GetDatesWithStagedAdjustments(vlEmployeeNumber As Long) As ADODB.Recordset

    On Error GoTo GetDatesWithStagedAdjustments_EH

    Dim objAdj As Ameren_TRIS_Server_COMIntegration.AutomatedAdjustments
    Set objAdj = New Ameren_TRIS_Server_COMIntegration.AutomatedAdjustments
    Set GetDatesWithStagedAdjustments = objAdj.GetDatesWithStagedAdjustments(vlEmployeeNumber)
    Set objAdj = Nothing

    Exit Function

GetDatesWithStagedAdjustments_EH:
    
    Set objAdj = Nothing
    TrsRaiseError MODULENAME _
                , "GetDatesWithStagedAdjustments" _
                , , , , _
                , True

End Function

Public Sub RemoveStagedAdjustments(vlEmployeeNumber As Long, vadtChargeDates As Variant, _
    ByRef rrsBrvs As ADODB.Recordset)
    
    On Error GoTo RemoveStagedAdjustments_EH

    Dim Dates() As Date
    Dates = vadtChargeDates

    Dim objAdj As Ameren_TRIS_Server_COMIntegration.AutomatedAdjustments
    Set objAdj = New Ameren_TRIS_Server_COMIntegration.AutomatedAdjustments
    objAdj.RemoveStagedAdjustments vlEmployeeNumber, Dates, rrsBrvs
    Set objAdj = Nothing

    Exit Sub

RemoveStagedAdjustments_EH:
    
    Set objAdj = Nothing
    TrsRaiseError MODULENAME _
                , "RemoveStagedAdjustments" _
                , , , , _
                , True
    
End Sub

Public Function GetApprovalStagedEmployeeList(vlUserID As Long) As ADODB.Recordset

    On Error GoTo GetApprovalStagedEmployeeList_EH

    Dim objTime As Ameren_TRIS_Server_COMIntegration.Time
    Set objTime = New Ameren_TRIS_Server_COMIntegration.Time
    Set GetApprovalStagedEmployeeList = objTime.GetApprovalEmployeeList(vlUserID)
    Set objTime = Nothing

    Exit Function

GetApprovalStagedEmployeeList_EH:
    
    Set objTime = Nothing
    TrsRaiseError MODULENAME _
                , "GetApprovalStagedEmployeeList" _
                , , , , _
                , True

End Function

Public Function GetApprovalStagedEntries(vlEmployeeNumber As Long) As ADODB.Recordset

    On Error GoTo GetApprovalStagedEntries_EH

    Dim objTime As Ameren_TRIS_Server_COMIntegration.Time
    Set objTime = New Ameren_TRIS_Server_COMIntegration.Time
    Set GetApprovalStagedEntries = objTime.GetApprovalTimeEntries(vlEmployeeNumber)
    Set objTime = Nothing

    Exit Function

GetApprovalStagedEntries_EH:
    
    Set objTime = Nothing
    TrsRaiseError MODULENAME _
                , "GetApprovalStagedEntries" _
                , , , , _
                , True

End Function

Public Sub SetApproval(ByVal SetList As ADODB.Recordset, ByVal SetAsApproved As Boolean, ByVal UserID As Long, _
    ByRef Brvs As ADODB.Recordset)
                           
    On Error GoTo SetApproval_EH

    Dim objAdj As Ameren_TRIS_Server_COMIntegration.AutomatedAdjustments
    Set objAdj = New Ameren_TRIS_Server_COMIntegration.AutomatedAdjustments
    objAdj.SetApproval SetList, SetAsApproved, UserID, Brvs
    Set objAdj = Nothing

    Exit Sub

SetApproval_EH:
    
    Set objAdj = Nothing
    TrsRaiseError MODULENAME _
                , "SetApproval" _
                , , , , _
                , True
                           
End Sub

Public Function GetPerDateDefaultInformation(ByVal EmployeeNumber As Long, ByVal StartDate As Date, ByVal EndDate As Date) _
    As ADODB.Recordset

    On Error GoTo GetPerDateDefaultInformation_EH

    Dim objAdj As Ameren_TRIS_Server_COMIntegration.AutomatedAdjustments
    Set objAdj = New Ameren_TRIS_Server_COMIntegration.AutomatedAdjustments
    Set GetPerDateDefaultInformation = objAdj.GetPerDateDefaultInformation(EmployeeNumber, StartDate, EndDate)
    Set objAdj = Nothing

    Exit Function

GetPerDateDefaultInformation_EH:
    
    Set objAdj = Nothing
    TrsRaiseError MODULENAME _
                , "GetPerDateDefaultInformation" _
                , , , , _
                , True

End Function

Public Function GetNotificationListForApproval(ByVal vlTRISUserNumber As Long) As ADODB.Recordset
    On Error GoTo eh

    Dim l_objAdjustments As dbAdjustments.cdAdjustments
    Set l_objAdjustments = CreateObject("dbAdjustments.cdAdjustments")
    Set GetNotificationListForApproval = l_objAdjustments.GetNotificationListForApproval(vlTRISUserNumber)
    Set l_objAdjustments = Nothing

    Exit Function
eh:
    Set l_objAdjustments = Nothing
    TrsRaiseError MODULENAME, "GetNotificationListForApproval", , , , , True
End Function
