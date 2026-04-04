VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "cbsBuiltDataScope"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Constant
Private Const MODULENAME As String = "busSecurity.cbsBuiltDataScope"

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

    Exit Sub
eh:
    Set lobjdbDataScope = Nothing
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

'//--------------------------------------------------------------------------------//
'//-----expose the following individually for calling from several threads-----//
'//--------------------------------------------------------------------------------//
Public Sub CheckAndRebuildDataScopeEmp2(ByVal intEmpNum As Long, _
                                        ByVal boolActive As Boolean)
'On Error GoTo eh
    Dim lobjdbDataScope As dbSecurity.cdBuiltDataScope

    Set lobjdbDataScope = CreateObject("dbSecurity.cdBuiltDataScope")
    
    lobjdbDataScope.CheckAndRebuildDataScopeEmp2 intEmpNum, boolActive
    
    Set lobjdbDataScope = Nothing
    Exit Sub
                                                
End Sub

Public Sub CheckAndRebuildDataScopeWG2(ByVal intEmpNum As Long, _
                                        ByVal boolActive As Boolean)
'On Error GoTo eh
    Dim lobjdbDataScope As dbSecurity.cdBuiltDataScope

    Set lobjdbDataScope = CreateObject("dbSecurity.cdBuiltDataScope")
    
    lobjdbDataScope.CheckAndRebuildDataScopeWG2 intEmpNum, boolActive
    
    Set lobjdbDataScope = Nothing
    Exit Sub
                                                
End Sub

Public Sub CheckAndRebuildDataScopeDept2(ByVal intEmpNum As Long)

'On Error GoTo eh
    Dim lobjdbDataScope As dbSecurity.cdBuiltDataScope

    Set lobjdbDataScope = CreateObject("dbSecurity.cdBuiltDataScope")
    
    lobjdbDataScope.CheckAndRebuildDataScopeDept2 intEmpNum
    
    Set lobjdbDataScope = Nothing
    Exit Sub
                                                
End Sub

Public Sub CheckAndRebuildDataScopeFun2(ByVal intEmpNum As Long)

    'On Error GoTo eh
    Dim lobjdbDataScope As dbSecurity.cdBuiltDataScope

    Set lobjdbDataScope = CreateObject("dbSecurity.cdBuiltDataScope")
    
    lobjdbDataScope.CheckAndRebuildDataScopeFun2 intEmpNum
    
    Set lobjdbDataScope = Nothing
    Exit Sub
                                                
End Sub

Public Sub CheckAndRebuildDataScopeDiv2(ByVal intEmpNum As Long)

'On Error GoTo eh
    Dim lobjdbDataScope As dbSecurity.cdBuiltDataScope

    Set lobjdbDataScope = CreateObject("dbSecurity.cdBuiltDataScope")
    
    lobjdbDataScope.CheckAndRebuildDataScopeDiv2 intEmpNum
    
    Set lobjdbDataScope = Nothing
    Exit Sub
                                                
End Sub

Public Function GetDataScopeEmp2(ByVal intEmpNum As Long, _
                                ByVal boolActive As Boolean) As ADODB.Recordset
'On Error GoTo eh
    Dim lobjdbDataScope As dbSecurity.cdBuiltDataScope

    Set lobjdbDataScope = CreateObject("dbSecurity.cdBuiltDataScope")
    
    Set GetDataScopeEmp2 = lobjdbDataScope.GetDataScopeEmp2(intEmpNum, boolActive)
    
    Set lobjdbDataScope = Nothing
    Exit Function
    
End Function

Public Function GetDataScopeWG2(ByVal intEmpNum As Long, _
                                ByVal boolActive As Boolean) As ADODB.Recordset
'On Error GoTo eh
    Dim lobjdbDataScope As dbSecurity.cdBuiltDataScope

    Set lobjdbDataScope = CreateObject("dbSecurity.cdBuiltDataScope")
    
    Set GetDataScopeWG2 = lobjdbDataScope.GetDataScopeWG2(intEmpNum, boolActive)
    
    Set lobjdbDataScope = Nothing
    Exit Function
    
End Function

Public Function GetDataScopeDept2(ByVal intEmpNum As Long) As ADODB.Recordset
'On Error GoTo eh
    Dim lobjdbDataScope As dbSecurity.cdBuiltDataScope

    Set lobjdbDataScope = CreateObject("dbSecurity.cdBuiltDataScope")
    
    Set GetDataScopeDept2 = lobjdbDataScope.GetDataScopeDept2(intEmpNum)
    
    Set lobjdbDataScope = Nothing
    Exit Function
    
End Function

Public Function GetDataScopeFun2(ByVal intEmpNum As Long) As ADODB.Recordset
'On Error GoTo eh
    Dim lobjdbDataScope As dbSecurity.cdBuiltDataScope

    Set lobjdbDataScope = CreateObject("dbSecurity.cdBuiltDataScope")
    
    Set GetDataScopeFun2 = lobjdbDataScope.GetDataScopeFun2(intEmpNum)
    
    Set lobjdbDataScope = Nothing
    Exit Function
    
End Function

Public Function GetDataScopeDiv2(ByVal intEmpNum As Long) As ADODB.Recordset
'On Error GoTo eh
    Dim lobjdbDataScope As dbSecurity.cdBuiltDataScope

    Set lobjdbDataScope = CreateObject("dbSecurity.cdBuiltDataScope")
    
    Set GetDataScopeDiv2 = lobjdbDataScope.GetDataScopeDiv2(intEmpNum)
    
    Set lobjdbDataScope = Nothing
    Exit Function
    
End Function

'//-----for client filtering-----//
Public Function GetDataScopeAEEmp2(ByVal intEmpNum As Long, _
                                ByVal boolActive As Boolean) As ADODB.Recordset
'On Error GoTo eh
    Dim lobjdbDataScope As dbSecurity.cdBuiltDataScope

    Set lobjdbDataScope = CreateObject("dbSecurity.cdBuiltDataScope")
    
    Set GetDataScopeAEEmp2 = lobjdbDataScope.GetDataScopeAEEmp2(intEmpNum, boolActive)
    
    Set lobjdbDataScope = Nothing
    Exit Function
End Function

Public Function GetDataScopeAEWG2(ByVal intEmpNum As Long, _
                                ByVal boolActive As Boolean) As ADODB.Recordset
'On Error GoTo eh
    Dim lobjdbDataScope As dbSecurity.cdBuiltDataScope

    Set lobjdbDataScope = CreateObject("dbSecurity.cdBuiltDataScope")
    
    Set GetDataScopeAEWG2 = lobjdbDataScope.GetDataScopeAEWG2(intEmpNum, boolActive)
    
    Set lobjdbDataScope = Nothing
    Exit Function
End Function

Public Function GetDataScopeAEDept2(ByVal intEmpNum As Long) As ADODB.Recordset
'On Error GoTo eh
    Dim lobjdbDataScope As dbSecurity.cdBuiltDataScope

    Set lobjdbDataScope = CreateObject("dbSecurity.cdBuiltDataScope")
    
    Set GetDataScopeAEDept2 = lobjdbDataScope.GetDataScopeAEDept2(intEmpNum)
    
    Set lobjdbDataScope = Nothing
    Exit Function
End Function

Public Function GetDataScopeAEFun2(ByVal intEmpNum As Long) As ADODB.Recordset
'On Error GoTo eh
    Dim lobjdbDataScope As dbSecurity.cdBuiltDataScope

    Set lobjdbDataScope = CreateObject("dbSecurity.cdBuiltDataScope")
    
    Set GetDataScopeAEFun2 = lobjdbDataScope.GetDataScopeAEFun2(intEmpNum)
    
    Set lobjdbDataScope = Nothing
    Exit Function
End Function

Public Function GetDataScopeAEDiv2(ByVal intEmpNum As Long) As ADODB.Recordset
'On Error GoTo eh
    Dim lobjdbDataScope As dbSecurity.cdBuiltDataScope

    Set lobjdbDataScope = CreateObject("dbSecurity.cdBuiltDataScope")
    
    Set GetDataScopeAEDiv2 = lobjdbDataScope.GetDataScopeAEDiv2(intEmpNum)
    
    Set lobjdbDataScope = Nothing
    Exit Function
End Function

'//-----for client filtering-----//
Public Function GetDataScopeDWG2(ByVal intEmpNum As Long, _
                                ByVal boolActive As Boolean) As ADODB.Recordset
'On Error GoTo eh
    Dim lobjdbDataScope As dbSecurity.cdBuiltDataScope

    Set lobjdbDataScope = CreateObject("dbSecurity.cdBuiltDataScope")
    
    Set GetDataScopeDWG2 = lobjdbDataScope.GetDataScopeDWG2(intEmpNum, boolActive)
    
    Set lobjdbDataScope = Nothing
    Exit Function
End Function

Public Function GetDataScopeNWG2(ByVal intEmpNum As Long, _
                                ByVal boolActive As Boolean) As ADODB.Recordset
'On Error GoTo eh
    Dim lobjdbDataScope As dbSecurity.cdBuiltDataScope

    Set lobjdbDataScope = CreateObject("dbSecurity.cdBuiltDataScope")
    
    Set GetDataScopeNWG2 = lobjdbDataScope.GetDataScopeNWG2(intEmpNum, boolActive)
    
    Set lobjdbDataScope = Nothing
    Exit Function
End Function

'//-----expose for any testing through the provider
Public Function InvalidateDataScopeBatch(ByVal blnCascade As Boolean) As Boolean

'On Error GoTo eh
    Dim lobjdbDataScope As dbSecurity.cdBuiltDataScope

    Set lobjdbDataScope = CreateObject("dbSecurity.cdBuiltDataScope")
    
    InvalidateDataScopeBatch = lobjdbDataScope.InvalidateDataScopeBatch(blnCascade)
    
    Set lobjdbDataScope = Nothing
    Exit Function
                                                
End Function

'//-----expose for testing/fixes
Public Function InvalidateDataScopeByRefEmp(ByVal intEmpNum As Long, ByVal blnCascade As Boolean) As Boolean
    Dim lobjdbDataScope As dbSecurity.cdBuiltDataScope

    Set lobjdbDataScope = CreateObject("dbSecurity.cdBuiltDataScope")
    
    InvalidateDataScopeByRefEmp = lobjdbDataScope.InvalidateDataScopeByRefEmp(intEmpNum, blnCascade)
    
    Set lobjdbDataScope = Nothing
    Exit Function
End Function

Public Function InvalidateDataScopeByRefWG(ByVal intWGNum As Long, ByVal blnCascade As Boolean) As Boolean
    Dim lobjdbDataScope As dbSecurity.cdBuiltDataScope

    Set lobjdbDataScope = CreateObject("dbSecurity.cdBuiltDataScope")
    
    InvalidateDataScopeByRefWG = lobjdbDataScope.InvalidateDataScopeByRefWG(intWGNum, blnCascade)
    
    Set lobjdbDataScope = Nothing
    Exit Function
End Function

Public Function InvalidateDataScopeByRefDept(ByVal strDeptNum As String, ByVal blnCascade As Boolean) As Boolean
    Dim lobjdbDataScope As dbSecurity.cdBuiltDataScope

    Set lobjdbDataScope = CreateObject("dbSecurity.cdBuiltDataScope")
    
    InvalidateDataScopeByRefDept = lobjdbDataScope.InvalidateDataScopeByRefDept(strDeptNum, blnCascade)
    
    Set lobjdbDataScope = Nothing
    Exit Function
End Function

Public Function InvalidateDataScopeByRefFun(ByVal strFunNum As String, ByVal blnCascade As Boolean) As Boolean
    Dim lobjdbDataScope As dbSecurity.cdBuiltDataScope

    Set lobjdbDataScope = CreateObject("dbSecurity.cdBuiltDataScope")
    
    InvalidateDataScopeByRefFun = lobjdbDataScope.InvalidateDataScopeByRefFun(strFunNum, blnCascade)
    
    Set lobjdbDataScope = Nothing
    Exit Function
End Function

Public Function InvalidateDataScopeByRefDiv(ByVal strDivNum As String, ByVal blnCascade As Boolean) As Boolean
    Dim lobjdbDataScope As dbSecurity.cdBuiltDataScope

    Set lobjdbDataScope = CreateObject("dbSecurity.cdBuiltDataScope")
    
    InvalidateDataScopeByRefDiv = lobjdbDataScope.InvalidateDataScopeByRefDiv(strDivNum, blnCascade)
    
    Set lobjdbDataScope = Nothing
    Exit Function
End Function

Public Function InvalidateDataScopeByEffEmp(ByVal intEmpNum As Long, ByVal blnCascade As Boolean) As Boolean
    Dim lobjdbDataScope As dbSecurity.cdBuiltDataScope

    Set lobjdbDataScope = CreateObject("dbSecurity.cdBuiltDataScope")
    
    InvalidateDataScopeByEffEmp = lobjdbDataScope.InvalidateDataScopeDivByEffEmp(intEmpNum, blnCascade)
    
    Set lobjdbDataScope = Nothing
    Exit Function
End Function

Public Function RebuildDataScopeBatch(ByVal blnCascade As Boolean) As Boolean

'On Error GoTo eh
    Dim lobjdbDataScope As dbSecurity.cdBuiltDataScope

    Set lobjdbDataScope = CreateObject("dbSecurity.cdBuiltDataScope")
    
    RebuildDataScopeBatch = lobjdbDataScope.RebuildDataScopeBatch(blnCascade)
    
    Set lobjdbDataScope = Nothing
    Exit Function
                                                
End Function

Public Function RemoveDataScopeBatch(ByVal blnCascade As Boolean) As Boolean

'On Error GoTo eh
    Dim lobjdbDataScope As dbSecurity.cdBuiltDataScope

    Set lobjdbDataScope = CreateObject("dbSecurity.cdBuiltDataScope")
    
    RemoveDataScopeBatch = lobjdbDataScope.RemoveDataScopeBatch(blnCascade)
    
    Set lobjdbDataScope = Nothing
    Exit Function
                                                
End Function

Public Function GetBatchNeedingRebuiltTrace() As ADODB.Recordset
    Dim lobjdbDataScope As dbSecurity.cdBuiltDataScope

    Set lobjdbDataScope = CreateObject("dbSecurity.cdBuiltDataScope")
    
    Set GetBatchNeedingRebuiltTrace = lobjdbDataScope.GetBatchNeedingRebuiltTrace
    
    Set lobjdbDataScope = Nothing
    Exit Function
End Function

Public Function GetBatchSyncAssocationsToAuditTrace() As ADODB.Recordset
    Dim lobjdbDataScope As dbSecurity.cdBuiltDataScope

    Set lobjdbDataScope = CreateObject("dbSecurity.cdBuiltDataScope")
    
    Set GetBatchSyncAssocationsToAuditTrace = lobjdbDataScope.GetBatchSyncAssocationsToAuditTrace
    
    Set lobjdbDataScope = Nothing
    Exit Function
End Function

Public Function GetBatchRemoveTrace() As ADODB.Recordset
    Dim lobjdbDataScope As dbSecurity.cdBuiltDataScope

    Set lobjdbDataScope = CreateObject("dbSecurity.cdBuiltDataScope")
    
    Set GetBatchRemoveTrace = lobjdbDataScope.GetBatchRemoveTrace
    
    Set lobjdbDataScope = Nothing
    Exit Function
End Function

Public Function GetBatchInvalidateTrace() As ADODB.Recordset
    Dim lobjdbDataScope As dbSecurity.cdBuiltDataScope

    Set lobjdbDataScope = CreateObject("dbSecurity.cdBuiltDataScope")
    
    Set GetBatchInvalidateTrace = lobjdbDataScope.GetBatchInvalidateTrace
    
    Set lobjdbDataScope = Nothing
    Exit Function
End Function

Private Sub FilterWorkgroups(ByVal vadrsAllWorkgroups As ADODB.Recordset _
                    , ByRef radrsDemandWorkgroups As ADODB.Recordset _
                    , ByRef radrsNonDemandWorkgroups As ADODB.Recordset)
On Error GoTo FilterWorkgroups_EH
Dim llIndex As Long
'RS Fields
Dim lsFldName As String, llFldIndex As Long
Dim llFldSize As Long, ladFldType As ADODB.DataTypeEnum
Dim ladFldAttrib As ADODB.FieldAttributeEnum

    With vadrsAllWorkgroups
        .Filter = adFilterNone

        .Filter = MDAC25FilterFix("DEMAND_IND = 'Y'")
        If Not (.EOF) Then
            Set radrsDemandWorkgroups = New ADODB.Recordset

            For llIndex = 0 To .Fields.Count - 1
                'Get the fld name, size
                lsFldName = .Fields.Item(llIndex).Name
                ladFldType = .Fields.Item(llIndex).Type
                llFldSize = .Fields.Item(llIndex).DefinedSize
                ladFldAttrib = .Fields.Item(llIndex).Attributes

                If (llFldSize > 0) Then
                    radrsDemandWorkgroups.Fields.Append lsFldName _
                        , ladFldType _
                        , llFldSize _
                        , ladFldAttrib
                Else
                    radrsDemandWorkgroups.Fields.Append lsFldName _
                        , ladFldType _
                        , _
                        , ladFldAttrib
                End If
            Next llIndex

            radrsDemandWorkgroups.CursorLocation = adUseClient
            radrsDemandWorkgroups.LockType = adLockBatchOptimistic
            radrsDemandWorkgroups.CursorType = adOpenStatic
            radrsDemandWorkgroups.Open

            .MoveFirst
            For llIndex = 1 To .RecordCount
                radrsDemandWorkgroups.AddNew

                For llFldIndex = 0 To .Fields.Count - 1
                    radrsDemandWorkgroups.Fields.Item(llFldIndex).Value _
                        = .Fields.Item(llFldIndex).Value
                Next llFldIndex

                radrsDemandWorkgroups.Update

                .MoveNext
            Next llIndex
        End If

        .Filter = adFilterNone

        .Filter = MDAC25FilterFix("DEMAND_IND = 'N'")
        If Not (.EOF) Then
            Set radrsNonDemandWorkgroups = New ADODB.Recordset

            For llIndex = 0 To .Fields.Count - 1
                'Get the fld name, size
                lsFldName = .Fields.Item(llIndex).Name
                ladFldType = .Fields.Item(llIndex).Type
                llFldSize = .Fields.Item(llIndex).DefinedSize
                ladFldAttrib = .Fields.Item(llIndex).Attributes

                If (llFldSize > 0) Then
                    radrsNonDemandWorkgroups.Fields.Append lsFldName _
                        , ladFldType _
                        , llFldSize _
                        , ladFldAttrib
                Else
                    radrsNonDemandWorkgroups.Fields.Append lsFldName _
                        , ladFldType _
                        , _
                        , ladFldAttrib
                End If
            Next llIndex

            radrsNonDemandWorkgroups.CursorLocation = adUseClient
            radrsNonDemandWorkgroups.LockType = adLockBatchOptimistic
            radrsNonDemandWorkgroups.CursorType = adOpenStatic
            radrsNonDemandWorkgroups.Open

            .MoveFirst
            For llIndex = 1 To .RecordCount
                radrsNonDemandWorkgroups.AddNew

                For llFldIndex = 0 To .Fields.Count - 1
                    radrsNonDemandWorkgroups.Fields.Item(llFldIndex).Value _
                        = .Fields.Item(llFldIndex).Value
                Next llFldIndex

                radrsNonDemandWorkgroups.Update

                .MoveNext
            Next llIndex
        End If

        .Filter = adFilterNone
        .MoveFirst
    End With

    Exit Sub
FilterWorkgroups_EH:
    TrsRaiseError MODULENAME _
        , "FilterWorkgroups"
End Sub
Private Sub SetApprovalEntryRS(ByVal vadrsEMP As ADODB.Recordset _
                    , ByVal vadrsWG As ADODB.Recordset _
                    , ByVal vadrsDEPT As ADODB.Recordset _
                    , ByVal vadrsFUN As ADODB.Recordset _
                    , ByVal vadrsDIV As ADODB.Recordset _
                    , ByRef radrsAEEMP As ADODB.Recordset _
                    , ByRef radrsAEWG As ADODB.Recordset _
                    , ByRef radrsAEDEPT As ADODB.Recordset _
                    , ByRef radrsAEFUN As ADODB.Recordset _
                    , ByRef radrsAEDIV As ADODB.Recordset)
On Error GoTo SetApprovalEntryRS_EH
Dim llIndex As Long
'RS Fields
Dim lsFldName As String, llFldIndex As Long
Dim llFldSize As Long, ladFldType As ADODB.DataTypeEnum
Dim ladFldAttrib As ADODB.FieldAttributeEnum

    'Employee
    If (IsRecordsetPopulated(vadrsEMP)) Then
        With vadrsEMP
            .Filter = adFilterNone

            .Filter = MDAC25FilterFix("APP_ENT_AUTHORITY <> 'V'")
            If Not (.EOF) Then
                Set radrsAEEMP = New ADODB.Recordset

                For llIndex = 0 To .Fields.Count - 1
                    'Get the fld name, size
                    lsFldName = .Fields.Item(llIndex).Name
                    ladFldType = .Fields.Item(llIndex).Type
                    llFldSize = .Fields.Item(llIndex).DefinedSize
                    ladFldAttrib = .Fields.Item(llIndex).Attributes

                    If (llFldSize > 0) Then
                        radrsAEEMP.Fields.Append lsFldName _
                            , ladFldType _
                            , llFldSize _
                            , ladFldAttrib
                    Else
                        radrsAEEMP.Fields.Append lsFldName _
                            , ladFldType _
                            , _
                            , ladFldAttrib
                    End If
                Next llIndex

                radrsAEEMP.CursorLocation = adUseClient
                radrsAEEMP.LockType = adLockBatchOptimistic
                radrsAEEMP.CursorType = adOpenStatic
                radrsAEEMP.Open

                .MoveFirst
                For llIndex = 1 To .RecordCount
                    radrsAEEMP.AddNew

                    For llFldIndex = 0 To .Fields.Count - 1
                        radrsAEEMP.Fields.Item(llFldIndex).Value _
                            = .Fields.Item(llFldIndex).Value
                    Next llFldIndex

                    radrsAEEMP.Update

                    .MoveNext
                Next llIndex
            End If

            .Filter = adFilterNone
            .MoveFirst
        End With
    End If

    'Workgroup
    If (IsRecordsetPopulated(vadrsWG)) Then
        With vadrsWG
            .Filter = adFilterNone

            .Filter = MDAC25FilterFix("APP_ENT_AUTHORITY <> 'V'")
            If Not (.EOF) Then
                Set radrsAEWG = New ADODB.Recordset

                For llIndex = 0 To .Fields.Count - 1
                    'Get the fld name, size
                    lsFldName = .Fields.Item(llIndex).Name
                    ladFldType = .Fields.Item(llIndex).Type
                    llFldSize = .Fields.Item(llIndex).DefinedSize
                    ladFldAttrib = .Fields.Item(llIndex).Attributes

                    If (llFldSize > 0) Then
                        radrsAEWG.Fields.Append lsFldName _
                            , ladFldType _
                            , llFldSize _
                            , ladFldAttrib
                    Else
                        radrsAEWG.Fields.Append lsFldName _
                            , ladFldType _
                            , _
                            , ladFldAttrib
                    End If
                Next llIndex

                radrsAEWG.CursorLocation = adUseClient
                radrsAEWG.LockType = adLockBatchOptimistic
                radrsAEWG.CursorType = adOpenStatic
                radrsAEWG.Open

                .MoveFirst
                For llIndex = 1 To .RecordCount
                    radrsAEWG.AddNew

                    For llFldIndex = 0 To .Fields.Count - 1
                        radrsAEWG.Fields.Item(llFldIndex).Value _
                            = .Fields.Item(llFldIndex).Value
                    Next llFldIndex

                    radrsAEWG.Update

                    .MoveNext
                Next llIndex
            End If

            .Filter = adFilterNone
            .MoveFirst
        End With
    End If

    'Department
    If (IsRecordsetPopulated(vadrsDEPT)) Then
        With vadrsDEPT
            .Filter = adFilterNone

            .Filter = MDAC25FilterFix("APP_ENT_AUTHORITY <> 'V'")
            If Not (.EOF) Then
                Set radrsAEDEPT = New ADODB.Recordset

                For llIndex = 0 To .Fields.Count - 1
                    'Get the fld name, size
                    lsFldName = .Fields.Item(llIndex).Name
                    ladFldType = .Fields.Item(llIndex).Type
                    llFldSize = .Fields.Item(llIndex).DefinedSize
                    ladFldAttrib = .Fields.Item(llIndex).Attributes

                    If (llFldSize > 0) Then
                        radrsAEDEPT.Fields.Append lsFldName _
                            , ladFldType _
                            , llFldSize _
                            , ladFldAttrib
                    Else
                        radrsAEDEPT.Fields.Append lsFldName _
                            , ladFldType _
                            , _
                            , ladFldAttrib
                    End If
                Next llIndex

                radrsAEDEPT.CursorLocation = adUseClient
                radrsAEDEPT.LockType = adLockBatchOptimistic
                radrsAEDEPT.CursorType = adOpenStatic
                radrsAEDEPT.Open

                .MoveFirst
                For llIndex = 1 To .RecordCount
                    radrsAEDEPT.AddNew

                    For llFldIndex = 0 To .Fields.Count - 1
                        radrsAEDEPT.Fields.Item(llFldIndex).Value _
                            = .Fields.Item(llFldIndex).Value
                    Next llFldIndex

                    radrsAEDEPT.Update

                    .MoveNext
                Next llIndex
            End If

            .Filter = adFilterNone
            .MoveFirst
        End With
    End If

    'Function
    If (IsRecordsetPopulated(vadrsFUN)) Then
        With vadrsFUN
            .Filter = adFilterNone

            .Filter = MDAC25FilterFix("APP_ENT_AUTHORITY <> 'V'")
            If Not (.EOF) Then
                Set radrsAEFUN = New ADODB.Recordset

                For llIndex = 0 To .Fields.Count - 1
                    'Get the fld name, size
                    lsFldName = .Fields.Item(llIndex).Name
                    ladFldType = .Fields.Item(llIndex).Type
                    llFldSize = .Fields.Item(llIndex).DefinedSize
                    ladFldAttrib = .Fields.Item(llIndex).Attributes

                    If (llFldSize > 0) Then
                        radrsAEFUN.Fields.Append lsFldName _
                            , ladFldType _
                            , llFldSize _
                            , ladFldAttrib
                    Else
                        radrsAEFUN.Fields.Append lsFldName _
                            , ladFldType _
                            , _
                            , ladFldAttrib
                    End If
                Next llIndex

                radrsAEFUN.CursorLocation = adUseClient
                radrsAEFUN.LockType = adLockBatchOptimistic
                radrsAEFUN.CursorType = adOpenStatic
                radrsAEFUN.Open

                .MoveFirst
                For llIndex = 1 To .RecordCount
                    radrsAEFUN.AddNew

                    For llFldIndex = 0 To .Fields.Count - 1
                        radrsAEFUN.Fields.Item(llFldIndex).Value _
                            = .Fields.Item(llFldIndex).Value
                    Next llFldIndex

                    radrsAEFUN.Update

                    .MoveNext
                Next llIndex
            End If

            .Filter = adFilterNone
            .MoveFirst
        End With
    End If

    'Division
    If (IsRecordsetPopulated(vadrsDIV)) Then
        With vadrsDIV
            .Filter = adFilterNone

            .Filter = MDAC25FilterFix("APP_ENT_AUTHORITY <> 'V'")
            If Not (.EOF) Then
                Set radrsAEDIV = New ADODB.Recordset

                For llIndex = 0 To .Fields.Count - 1
                    'Get the fld name, size
                    lsFldName = .Fields.Item(llIndex).Name
                    ladFldType = .Fields.Item(llIndex).Type
                    llFldSize = .Fields.Item(llIndex).DefinedSize
                    ladFldAttrib = .Fields.Item(llIndex).Attributes

                    If (llFldSize > 0) Then
                        radrsAEDIV.Fields.Append lsFldName _
                            , ladFldType _
                            , llFldSize _
                            , ladFldAttrib
                    Else
                        radrsAEDIV.Fields.Append lsFldName _
                            , ladFldType _
                            , _
                            , ladFldAttrib
                    End If
                Next llIndex

                radrsAEDIV.CursorLocation = adUseClient
                radrsAEDIV.LockType = adLockBatchOptimistic
                radrsAEDIV.CursorType = adOpenStatic
                radrsAEDIV.Open

                .MoveFirst
                For llIndex = 1 To .RecordCount
                    radrsAEDIV.AddNew

                    For llFldIndex = 0 To .Fields.Count - 1
                        radrsAEDIV.Fields.Item(llFldIndex).Value _
                            = .Fields.Item(llFldIndex).Value
                    Next llFldIndex

                    radrsAEDIV.Update

                    .MoveNext
                Next llIndex
            End If

            .Filter = adFilterNone
            .MoveFirst
        End With
    End If

    Exit Sub
SetApprovalEntryRS_EH:
    TrsRaiseError MODULENAME _
        , "SetApprovalEntryRS"
End Sub
Private Function IsRecordsetPopulated(ByVal vadrs As ADODB.Recordset) As Boolean
On Error GoTo IsRecordsetPopulated_EH

    IsRecordsetPopulated = False

    If Not (vadrs Is Nothing) Then
        If Not (vadrs.EOF And vadrs.BOF) Then
            IsRecordsetPopulated = True
        End If
    End If

    Exit Function
IsRecordsetPopulated_EH:
    TrsRaiseError MODULENAME _
        , "IsRecordsetPopulated"
End Function
'END PREM POST TRIS 5.6


