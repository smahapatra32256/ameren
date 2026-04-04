VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "cbsBuiltDataScopeTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'this class was created (4/14/2009) to be called by the training reset process as a single transaction (required)
'there were locking issues when attempting to call the component with a trans attrib of not supported
'the batch process for rebuilding datascope should not call this, but rather use cbsBuiltDataScope (not supported)
'so that multiple transactions are created rather than one big transaction

'Constant
Private Const MODULENAME As String = "busSecurity.cbsBuiltDataScopeTrans"

Public Sub GetDataScope(ByVal vlEmployee As Long _
            , ByVal vbActiveInformationOnly As Boolean _
            , ByRef radrsEMP As ADODB.Recordset _
            , ByRef radrsWG As ADODB.Recordset _
            , ByRef radrsDWG As ADODB.Recordset _
            , ByRef radrsNWG As ADODB.Recordset _
            , ByRef radrsDEPT As ADODB.Recordset _
            , ByRef radrsFUN As ADODB.Recordset _
            , ByRef radrsDIV As ADODB.Recordset _
            , ByRef radrsAEEMP As ADODB.Recordset _
            , ByRef radrsAEWG As ADODB.Recordset _
            , ByRef radrsAEDEPT As ADODB.Recordset _
            , ByRef radrsAEFUN As ADODB.Recordset _
            , ByRef radrsAEDIV As ADODB.Recordset _
            , Optional ByVal vbFilterWorkgroups As Boolean = True _
            , Optional ByVal vbFilterApprovalEntry As Boolean = True)
            
On Error GoTo eh

    GetDataScope3 vlEmployee, vbActiveInformationOnly, _
            radrsEMP, radrsWG, radrsDWG, radrsNWG, _
            radrsDEPT, radrsFUN, _
            radrsDIV, radrsAEEMP, _
            radrsAEWG, radrsAEDEPT, _
            radrsAEFUN, radrsAEDIV, _
            vbFilterWorkgroups, vbFilterApprovalEntry, False
    Exit Sub
eh:
    'Abort Transaction
    GetObjectContext().SetAbort

    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
        Case trErrutOpenConnection
            TrsRaiseError MODULENAME _
                , "GetDataScope"

        Case trErrdb_Employee_F, trErrdb_WorkGroup_F _
            , trErrdb_Department_F, trErrdb_Function_F _
            , trErrdb_Division_F, trErrdb_ExecuteSQL_F
            TrsRaiseError MODULENAME _
                , "GetDataScope" _
                , Err.Number _
                , , , , _
                True

        Case Else
            TrsRaiseError MODULENAME _
                , "GetDataScope" _
                , trErrDeriveDataScope_F _
                , , _
                , DescArgsFrom("Employee Number = " & CStr(vlEmployee)) _
                , True
    End Select
End Sub

Public Sub GetDataScope3(ByVal vlEmployee As Long _
            , ByVal vbActiveInformationOnly As Boolean _
            , ByRef radrsEMP As ADODB.Recordset _
            , ByRef radrsWG As ADODB.Recordset _
            , ByRef radrsDWG As ADODB.Recordset _
            , ByRef radrsNWG As ADODB.Recordset _
            , ByRef radrsDEPT As ADODB.Recordset _
            , ByRef radrsFUN As ADODB.Recordset _
            , ByRef radrsDIV As ADODB.Recordset _
            , ByRef radrsAEEMP As ADODB.Recordset _
            , ByRef radrsAEWG As ADODB.Recordset _
            , ByRef radrsAEDEPT As ADODB.Recordset _
            , ByRef radrsAEFUN As ADODB.Recordset _
            , ByRef radrsAEDIV As ADODB.Recordset _
            , Optional ByVal vbFilterWorkgroups As Boolean = True _
            , Optional ByVal vbFilterApprovalEntry As Boolean = True _
            , Optional ByVal vbForceRebuild As Boolean = False)

On Error GoTo eh
    Dim lobjdbDataScope As dbSecurity.cdBuiltDataScope

    Set lobjdbDataScope = CreateObject("dbSecurity.cdBuiltDataScope")
    
    lobjdbDataScope.GetAllDataScopes3 vlEmployee, vbActiveInformationOnly, _
            radrsEMP, radrsWG, radrsDWG, radrsNWG, _
            radrsDEPT, radrsFUN, _
            radrsDIV, radrsAEEMP, _
            radrsAEWG, radrsAEDEPT, _
            radrsAEFUN, radrsAEDIV, _
            vbFilterWorkgroups, vbFilterApprovalEntry, vbForceRebuild

'    Set radrsEMP = lobjdbDataScope.GetDataScopeEmp(vlEmployee, vbActiveInformationOnly)
'    Set radrsWG = lobjdbDataScope.GetDataScopeWG(vlEmployee, vbActiveInformationOnly)
'    Set radrsDEPT = lobjdbDataScope.GetDataScopeDept(vlEmployee)
'    Set radrsFUN = lobjdbDataScope.GetDataScopeFun(vlEmployee)
'    Set radrsDIV = lobjdbDataScope.GetDataScopeDiv(vlEmployee)
    
    Set lobjdbDataScope = Nothing

'    'Create and set Demand and Non-demand WG RS
'    If (vbFilterWorkgroups) Then
'        If (IsRecordsetPopulated(radrsWG)) Then
'            FilterWorkgroups radrsWG _
'                , radrsDWG _
'                , radrsNWG
'        End If
'    End If
'
'    'Create and set the view RS
'    If (vbFilterApprovalEntry) Then
'        SetApprovalEntryRS radrsEMP, radrsWG _
'            , radrsDEPT, radrsFUN _
'            , radrsDIV, radrsAEEMP _
'            , radrsAEWG, radrsAEDEPT _
'            , radrsAEFUN, radrsAEDIV
'    End If

    GetObjectContext().SetComplete
    Exit Sub
eh:
    Set lobjdbDataScope = Nothing
    'Abort Transaction
    GetObjectContext().SetAbort

    Set lobjdbDataScope = Nothing

    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
        Case trErrutOpenConnection
            TrsRaiseError MODULENAME _
                , "GetDataScope3"

        Case trErrdb_Employee_F, trErrdb_WorkGroup_F _
            , trErrdb_Department_F, trErrdb_Function_F _
            , trErrdb_Division_F, trErrdb_ExecuteSQL_F
            TrsRaiseError MODULENAME _
                , "GetDataScope3" _
                , Err.Number _
                , , , , _
                True

        Case Else
            TrsRaiseError MODULENAME _
                , "GetDataScope3" _
                , trErrDeriveDataScope_F _
                , , _
                , DescArgsFrom("Employee Number = " & CStr(vlEmployee)) _
                , True
    End Select

End Sub

Public Function InvalidateDataScopeByEffEmp(ByVal intEmpNum As Long, ByVal blnCascade As Boolean) As Boolean
    Dim lobjdbDataScope As dbSecurity.cdBuiltDataScope

    Set lobjdbDataScope = CreateObject("dbSecurity.cdBuiltDataScope")
    
    InvalidateDataScopeByEffEmp = lobjdbDataScope.InvalidateDataScopeDivByEffEmp(intEmpNum, blnCascade)
    
    Set lobjdbDataScope = Nothing
    Exit Function
End Function

