VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "cbsDataScope"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Project:   TRIS Conversion Phase 1
'Author:    Prem M. Chander
'CoAuthor:  Lance Chang
'Date:      September 1999 (Firstcut)
'Description: This class module will reside in a Server Package
'under Microsoft Transaction Server (MTS). This module will be called by
'the client and other components in MTS. This module returns datascope
'information for a given employee which include, employees, workgroups,
'departments, functions and dvisions. The determination is driven based on an
'employees approval information. This module will call its associated
'data access component to retrieve pertinant information from the database.

Option Explicit

'Constant
Private Const MODULENAME As String = "busSecurity.cbsDataScope"
Public Function DeriveDataScope(ByVal v_sKeyToRun As String _
                    , ByVal v_lEmpNum As Long _
                    , ByVal v_bActiveInfo As Boolean _
                    , ByRef r_adrsEMP As ADODB.Recordset _
                    , ByRef r_adrsWG As ADODB.Recordset _
                    , ByRef r_adrsDWG As ADODB.Recordset _
                    , ByRef r_adrsNWG As ADODB.Recordset _
                    , ByRef r_adrsDEPT As ADODB.Recordset _
                    , ByRef r_adrsFUN As ADODB.Recordset _
                    , ByRef r_adrsDIV As ADODB.Recordset _
                    ) As Boolean

    DeriveDataScope = DeriveDataScope2(v_sKeyToRun, v_lEmpNum _
                    , v_bActiveInfo _
                    , r_adrsEMP _
                    , r_adrsWG _
                    , r_adrsDWG _
                    , r_adrsNWG _
                    , r_adrsDEPT _
                    , r_adrsFUN _
                    , r_adrsDIV, True)

End Function
'Derives datascope for an employee. This function returns a boolean.
Public Function DeriveDataScope2(ByVal v_sKeyToRun As String _
                    , ByVal v_lEmpNum As Long _
                    , ByVal v_bActiveInfo As Boolean _
                    , ByRef r_adrsEMP As ADODB.Recordset _
                    , ByRef r_adrsWG As ADODB.Recordset _
                    , ByRef r_adrsDWG As ADODB.Recordset _
                    , ByRef r_adrsNWG As ADODB.Recordset _
                    , ByRef r_adrsDEPT As ADODB.Recordset _
                    , ByRef r_adrsFUN As ADODB.Recordset _
                    , ByRef r_adrsDIV As ADODB.Recordset _
                    , Optional ByVal v_bFilterWorkgroups As Boolean = True _
                    ) As Boolean

    DeriveDataScope2 = False

    DeriveDataScope3 v_lEmpNum, v_bActiveInfo _
                    , r_adrsEMP, r_adrsWG _
                    , r_adrsDWG, r_adrsNWG _
                    , r_adrsDEPT, r_adrsFUN _
                    , r_adrsDIV, Nothing _
                    , Nothing, Nothing _
                    , Nothing, Nothing _
                    , True, False

    DeriveDataScope2 = True

End Function
'BEGIN PREM TRIS 5.6
Public Sub DeriveDataScope3(ByVal vlEmployee As Long _
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
On Error GoTo DeriveDataScope3_EH
Dim ladrsApproval As ADODB.Recordset
Dim lobjdbDataScope As dbSecurity.cdDataScope

    Set lobjdbDataScope = CreateObject("dbSecurity.cdDataScope")

    'Get Approval Information
    Set ladrsApproval = lobjdbDataScope.GetApprovalInformation(vlEmployee)
    If (IsRecordsetPopulated(ladrsApproval)) Then
        With ladrsApproval
            .MoveFirst

            'If Division Level Access then get
            'divison, function, dept, wg, and employee
            .Find "APPR_LEVEL_IND = 10"
            If Not (.EOF) Then
                'Get Employees
                Set radrsEMP = lobjdbDataScope.GetEmployees(vlEmployee _
                                        , vbActiveInformationOnly)
                'Get Workgroups Primary and Secondary too.
                Set radrsWG = lobjdbDataScope.GetAllWorkgroups(vlEmployee)
                'Get Departments
                Set radrsDEPT = lobjdbDataScope.GetDepartments(vlEmployee)
                'Get Functions
                Set radrsFUN = lobjdbDataScope.GetFunctions(vlEmployee)
                'Get Divisions
                Set radrsDIV = lobjdbDataScope.GetDivisions(vlEmployee)
            Else
                .MoveFirst

                'If Function Level Access then get
                'function, dept, wg, and employee
                .Find "APPR_LEVEL_IND = 20"
                If Not (.EOF) Then
                    'Get Employees
                    Set radrsEMP = lobjdbDataScope.GetEmployees(vlEmployee _
                                            , vbActiveInformationOnly)
                    'Get Workgroups Primary and Secondary too.
                    Set radrsWG = lobjdbDataScope.GetAllWorkgroups(vlEmployee)
                    'Get Departments
                    Set radrsDEPT = lobjdbDataScope.GetDepartments(vlEmployee)
                    'Get Functions
                    Set radrsFUN = lobjdbDataScope.GetFunctions(vlEmployee)
                Else
                    .MoveFirst

                    'If Department Level Access then get
                    'dept, wg, and employee
                    .Find "APPR_LEVEL_IND = 30"
                    If Not (.EOF) Then
                        'Get Employees
                        Set radrsEMP = lobjdbDataScope.GetEmployees(vlEmployee _
                                                , vbActiveInformationOnly)
                        'Get Workgroups Primary and Secondary too.
                        Set radrsWG = lobjdbDataScope.GetAllWorkgroups(vlEmployee)
                        'Get Departments
                        Set radrsDEPT = lobjdbDataScope.GetDepartments(vlEmployee)
                    Else
                        .MoveFirst

                        'If Workgroup/Employee Level Access then get
                        'wg, and employee

                        'Get Employees
                        Set radrsEMP = lobjdbDataScope.GetEmployees(vlEmployee _
                                                , vbActiveInformationOnly)
                        'Get Workgroups Primary and Secondary too.
                        Set radrsWG = lobjdbDataScope.GetAllWorkgroups(vlEmployee)
                    End If
                End If
            End If
        End With
    End If

    Set lobjdbDataScope = Nothing

    'Eliminate Duplicates
    Set radrsEMP = EliminateEmployeeDuplicates(radrsEMP)
    Set radrsWG = EliminateWorkgroupDuplicates(radrsWG)
    Set radrsDEPT = EliminateDepartmentDuplicates(radrsDEPT)
    Set radrsFUN = EliminateFunctionDuplicates(radrsFUN)

    'Create and set Demand and Non-demand WG RS
    If (vbFilterWorkgroups) Then
        If (IsRecordsetPopulated(radrsWG)) Then
            FilterWorkgroups radrsWG _
                , radrsDWG _
                , radrsNWG
        End If
    End If

    'Create and set the view RS
    If (vbFilterApprovalEntry) Then
        SetApprovalEntryRS radrsEMP, radrsWG _
            , radrsDEPT, radrsFUN _
            , radrsDIV, radrsAEEMP _
            , radrsAEWG, radrsAEDEPT _
            , radrsAEFUN, radrsAEDIV
    End If

    Exit Sub
DeriveDataScope3_EH:
    Set lobjdbDataScope = Nothing

    'Call Standard Module Function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
        Case trErrutOpenConnection
            TrsRaiseError MODULENAME _
                , "DeriveDataScope3"

        Case trErrdb_Employee_F, trErrdb_WorkGroup_F _
            , trErrdb_Department_F, trErrdb_Function_F _
            , trErrdb_Division_F, trErrdb_ExecuteSQL_F
            TrsRaiseError MODULENAME _
                , "DeriveDataScope3" _
                , Err.Number _
                , , , , _
                True

        Case Else
            TrsRaiseError MODULENAME _
                , "DeriveDataScope3" _
                , trErrDeriveDataScope_F _
                , , _
                , DescArgsFrom("Employee Number = " & CStr(vlEmployee)) _
                , True
    End Select
End Sub
Private Function EliminateEmployeeDuplicates(ByVal vadrsSource As ADODB.Recordset _
                    ) As ADODB.Recordset
On Error GoTo EliminateEmployeeDuplicates_EH
Dim ladrs As ADODB.Recordset
Dim llIndex As Long, lbAddRecord As Boolean
Dim lbPopulated As Boolean
'RS Fields
Dim lsFldName As String, llFldIndex As Long
Dim llFldSize As Long, ladFldType As ADODB.DataTypeEnum
Dim ladFldAttrib As ADODB.FieldAttributeEnum

    Set EliminateEmployeeDuplicates = vadrsSource

    If (IsRecordsetPopulated(vadrsSource)) Then
        Set ladrs = New ADODB.Recordset

        With vadrsSource
            For llIndex = 0 To .Fields.Count - 1
                'Get the fld name, size
                lsFldName = .Fields.Item(llIndex).Name
                ladFldType = .Fields.Item(llIndex).Type
                llFldSize = .Fields.Item(llIndex).DefinedSize
                ladFldAttrib = .Fields.Item(llIndex).Attributes

                If (llFldSize > 0) Then
                    ladrs.Fields.Append lsFldName _
                        , ladFldType _
                        , llFldSize _
                        , ladFldAttrib
                Else
                    ladrs.Fields.Append lsFldName _
                        , ladFldType _
                        , _
                        , ladFldAttrib
                End If
            Next llIndex

            ladrs.CursorLocation = adUseClient
            ladrs.LockType = adLockBatchOptimistic
            ladrs.CursorType = adOpenStatic
            ladrs.Open

            .Filter = adFilterNone

            .Filter = MDAC25FilterFix("APPR_LEVEL_IND = 60")
            If Not (.EOF) Then
                .MoveFirst

                For llIndex = 1 To .RecordCount
                    ladrs.AddNew

                    For llFldIndex = 0 To .Fields.Count - 1
                        ladrs.Fields.Item(llFldIndex).Value _
                            = .Fields.Item(llFldIndex).Value
                    Next llFldIndex

                    ladrs.Update

                    .MoveNext
                Next llIndex
            End If

            lbPopulated = False
            If (IsRecordsetPopulated(ladrs)) Then
                ladrs.MoveFirst

                lbPopulated = True
            End If

            .Filter = adFilterNone

            .Filter = MDAC25FilterFix("APPR_LEVEL_IND = 50")
            If Not (.EOF) Then
                .MoveFirst

                For llIndex = 1 To .RecordCount
                    lbAddRecord = True

                    If (lbPopulated) Then
                        ladrs.MoveFirst

                        ladrs.Find "EMP_NUM = " & CStr(!EMP_NUM)
                        If Not (ladrs.EOF) Then
                            lbAddRecord = False
                        End If

                        ladrs.MoveFirst
                    End If

                    If (lbAddRecord) Then
                        ladrs.AddNew

                        For llFldIndex = 0 To .Fields.Count - 1
                            ladrs.Fields.Item(llFldIndex).Value _
                                = .Fields.Item(llFldIndex).Value
                        Next llFldIndex

                        ladrs.Update
                    End If

                    .MoveNext
                Next llIndex
            End If

            If Not (lbPopulated) Then
                If (IsRecordsetPopulated(ladrs)) Then
                    ladrs.MoveFirst

                    lbPopulated = True
                End If
            End If

            .Filter = adFilterNone

            .Filter = MDAC25FilterFix("APPR_LEVEL_IND = 30")
            If Not (.EOF) Then
                .MoveFirst

                For llIndex = 1 To .RecordCount
                    lbAddRecord = True

                    If (lbPopulated) Then
                        ladrs.MoveFirst

                        ladrs.Find "EMP_NUM = " & CStr(!EMP_NUM)
                        If Not (ladrs.EOF) Then
                            lbAddRecord = False
                        End If

                        ladrs.MoveFirst
                    End If

                    If (lbAddRecord) Then
                        ladrs.AddNew

                        For llFldIndex = 0 To .Fields.Count - 1
                            ladrs.Fields.Item(llFldIndex).Value _
                                = .Fields.Item(llFldIndex).Value
                        Next llFldIndex

                        ladrs.Update
                    End If

                    .MoveNext
                Next llIndex
            End If

            If Not (lbPopulated) Then
                If (IsRecordsetPopulated(ladrs)) Then
                    ladrs.MoveFirst

                    lbPopulated = True
                End If
            End If

            .Filter = adFilterNone

            .Filter = MDAC25FilterFix("APPR_LEVEL_IND = 20")
            If Not (.EOF) Then
                .MoveFirst

                For llIndex = 1 To .RecordCount
                    lbAddRecord = True

                    If (lbPopulated) Then
                        ladrs.MoveFirst

                        ladrs.Find "EMP_NUM = " & CStr(!EMP_NUM)
                        If Not (ladrs.EOF) Then
                            lbAddRecord = False
                        End If

                        ladrs.MoveFirst
                    End If

                    If (lbAddRecord) Then
                        ladrs.AddNew

                        For llFldIndex = 0 To .Fields.Count - 1
                            ladrs.Fields.Item(llFldIndex).Value _
                                = .Fields.Item(llFldIndex).Value
                        Next llFldIndex

                        ladrs.Update
                    End If

                    .MoveNext
                Next llIndex
            End If

            If Not (lbPopulated) Then
                If (IsRecordsetPopulated(ladrs)) Then
                    ladrs.MoveFirst

                    lbPopulated = True
                End If
            End If

            .Filter = adFilterNone

            .Filter = MDAC25FilterFix("APPR_LEVEL_IND = 10")
            If Not (.EOF) Then
                .MoveFirst

                For llIndex = 1 To .RecordCount
                    lbAddRecord = True

                    If (lbPopulated) Then
                        ladrs.MoveFirst

                        ladrs.Find "EMP_NUM = " & CStr(!EMP_NUM)
                        If Not (ladrs.EOF) Then
                            lbAddRecord = False
                        End If

                        ladrs.MoveFirst
                    End If

                    If (lbAddRecord) Then
                        ladrs.AddNew

                        For llFldIndex = 0 To .Fields.Count - 1
                            ladrs.Fields.Item(llFldIndex).Value _
                                = .Fields.Item(llFldIndex).Value
                        Next llFldIndex

                        ladrs.Update
                    End If

                    .MoveNext
                Next llIndex
            End If

            .Filter = adFilterNone
        End With
    End If

    If (IsRecordsetPopulated(ladrs)) Then
        ladrs.MoveFirst

        Set EliminateEmployeeDuplicates = ladrs
    End If

    Set ladrs = Nothing

    Exit Function
EliminateEmployeeDuplicates_EH:
    Set ladrs = Nothing

    TrsRaiseError MODULENAME _
        , "EliminateEmployeeDuplicates"

End Function
Private Function EliminateWorkgroupDuplicates(ByVal vadrsSource As ADODB.Recordset _
                    ) As ADODB.Recordset
On Error GoTo EliminateWorkgroupDuplicates_EH
Dim ladrs As ADODB.Recordset
Dim llIndex As Long, lbAddRecord As Boolean
Dim lbPopulated As Boolean
'RS Fields
Dim lsFldName As String, llFldIndex As Long
Dim llFldSize As Long, ladFldType As ADODB.DataTypeEnum
Dim ladFldAttrib As ADODB.FieldAttributeEnum

    Set EliminateWorkgroupDuplicates = vadrsSource

    If (IsRecordsetPopulated(vadrsSource)) Then
        Set ladrs = New ADODB.Recordset
        With vadrsSource
            For llIndex = 0 To .Fields.Count - 1
                'Get the fld name, size
                lsFldName = .Fields.Item(llIndex).Name
                ladFldType = .Fields.Item(llIndex).Type
                llFldSize = .Fields.Item(llIndex).DefinedSize
                ladFldAttrib = .Fields.Item(llIndex).Attributes

                If (llFldSize > 0) Then
                    ladrs.Fields.Append lsFldName _
                        , ladFldType _
                        , llFldSize _
                        , ladFldAttrib
                Else
                    ladrs.Fields.Append lsFldName _
                        , ladFldType _
                        , _
                        , ladFldAttrib
                End If
            Next llIndex

            ladrs.CursorLocation = adUseClient
            ladrs.LockType = adLockBatchOptimistic
            ladrs.CursorType = adOpenStatic
            ladrs.Open

            .Filter = adFilterNone

            .Filter = MDAC25FilterFix("APPR_LEVEL_IND = 50")
            If Not (.EOF) Then
                .MoveFirst

                For llIndex = 1 To .RecordCount
                    ladrs.AddNew

                    For llFldIndex = 0 To .Fields.Count - 1
                        ladrs.Fields.Item(llFldIndex).Value _
                            = .Fields.Item(llFldIndex).Value
                    Next llFldIndex

                    ladrs.Update

                    .MoveNext
                Next llIndex
            End If

            lbPopulated = False
            If (IsRecordsetPopulated(ladrs)) Then
                ladrs.MoveFirst

                lbPopulated = True
            End If

            .Filter = adFilterNone

            .Filter = MDAC25FilterFix("APPR_LEVEL_IND = 30")
            If Not (.EOF) Then
                .MoveFirst

                For llIndex = 1 To .RecordCount
                    lbAddRecord = True

                    If (lbPopulated) Then
                        ladrs.Filter = adFilterNone

                        ladrs.Filter = MDAC25FilterFix("WG_NUM = " & CStr(!WG_NUM) & _
                            " AND APP_ENT_AUTHORITY <> 'P' AND APP_ENT_AUTHORITY <> 'S'")
                        If Not (ladrs.EOF) Then
                            lbAddRecord = False
                        End If

                        ladrs.Filter = adFilterNone
                    End If

                    If (lbAddRecord) Then
                        ladrs.AddNew

                        For llFldIndex = 0 To .Fields.Count - 1
                            ladrs.Fields.Item(llFldIndex).Value _
                                = .Fields.Item(llFldIndex).Value
                        Next llFldIndex

                        ladrs.Update
                    End If

                    .MoveNext
                Next llIndex
            End If

            If Not (lbPopulated) Then
                If (IsRecordsetPopulated(ladrs)) Then
                    ladrs.MoveFirst

                    lbPopulated = True
                End If
            End If

            .Filter = adFilterNone

            .Filter = MDAC25FilterFix("APPR_LEVEL_IND = 20")
            If Not (.EOF) Then
                .MoveFirst

                For llIndex = 1 To .RecordCount
                    lbAddRecord = True

                    If (lbPopulated) Then
                        ladrs.Filter = adFilterNone

                        ladrs.Filter = MDAC25FilterFix("WG_NUM = " & CStr(!WG_NUM) & _
                            " AND APP_ENT_AUTHORITY <> 'P' AND APP_ENT_AUTHORITY <> 'S'")
                        If Not (ladrs.EOF) Then
                            lbAddRecord = False
                        End If

                        ladrs.Filter = adFilterNone
                    End If

                    If (lbAddRecord) Then
                        ladrs.AddNew

                        For llFldIndex = 0 To .Fields.Count - 1
                            ladrs.Fields.Item(llFldIndex).Value _
                                = .Fields.Item(llFldIndex).Value
                        Next llFldIndex

                        ladrs.Update
                    End If

                    .MoveNext
                Next llIndex
            End If

            If Not (lbPopulated) Then
                If (IsRecordsetPopulated(ladrs)) Then
                    ladrs.MoveFirst

                    lbPopulated = True
                End If
            End If

            .Filter = adFilterNone

            .Filter = MDAC25FilterFix("APPR_LEVEL_IND = 10")
            If Not (.EOF) Then
                .MoveFirst

                For llIndex = 1 To .RecordCount
                    lbAddRecord = True

                    If (lbPopulated) Then
                        ladrs.Filter = adFilterNone

                        ladrs.Filter = MDAC25FilterFix("WG_NUM = " & CStr(!WG_NUM) & _
                            " AND APP_ENT_AUTHORITY <> 'P' AND APP_ENT_AUTHORITY <> 'S'")
                        If Not (ladrs.EOF) Then
                            lbAddRecord = False
                        End If

                        ladrs.Filter = adFilterNone
                    End If

                    If (lbAddRecord) Then
                        ladrs.AddNew

                        For llFldIndex = 0 To .Fields.Count - 1
                            ladrs.Fields.Item(llFldIndex).Value _
                                = .Fields.Item(llFldIndex).Value
                        Next llFldIndex

                        ladrs.Update
                    End If

                    .MoveNext
                Next llIndex
            End If

            .Filter = adFilterNone
        End With
    End If

    If (IsRecordsetPopulated(ladrs)) Then
        ladrs.MoveFirst

        Set EliminateWorkgroupDuplicates = ladrs
    End If

    Set ladrs = Nothing

    Exit Function
EliminateWorkgroupDuplicates_EH:
    Set ladrs = Nothing

    TrsRaiseError MODULENAME _
        , "EliminateWorkgroupDuplicates"

End Function
Private Function EliminateDepartmentDuplicates(ByVal vadrsSource As ADODB.Recordset _
                    ) As ADODB.Recordset
On Error GoTo EliminateDepartmentDuplicates_EH
Dim ladrs As ADODB.Recordset
Dim llIndex As Long, lbAddRecord As Boolean
Dim lbPopulated As Boolean
'RS Fields
Dim lsFldName As String, llFldIndex As Long
Dim llFldSize As Long, ladFldType As ADODB.DataTypeEnum
Dim ladFldAttrib As ADODB.FieldAttributeEnum

    Set EliminateDepartmentDuplicates = vadrsSource

    If (IsRecordsetPopulated(vadrsSource)) Then
        Set ladrs = New ADODB.Recordset

        With vadrsSource
            For llIndex = 0 To .Fields.Count - 1
                'Get the fld name, size
                lsFldName = .Fields.Item(llIndex).Name
                ladFldType = .Fields.Item(llIndex).Type
                llFldSize = .Fields.Item(llIndex).DefinedSize
                ladFldAttrib = .Fields.Item(llIndex).Attributes

                If (llFldSize > 0) Then
                    ladrs.Fields.Append lsFldName _
                        , ladFldType _
                        , llFldSize _
                        , ladFldAttrib
                Else
                    ladrs.Fields.Append lsFldName _
                        , ladFldType _
                        , _
                        , ladFldAttrib
                End If
            Next llIndex

            ladrs.CursorLocation = adUseClient
            ladrs.LockType = adLockBatchOptimistic
            ladrs.CursorType = adOpenStatic
            ladrs.Open

            .Filter = adFilterNone

            .Filter = MDAC25FilterFix("APPR_LEVEL_IND = 30")
            If Not (.EOF) Then
                .MoveFirst

                For llIndex = 1 To .RecordCount
                    ladrs.AddNew

                    For llFldIndex = 0 To .Fields.Count - 1
                        ladrs.Fields.Item(llFldIndex).Value _
                            = .Fields.Item(llFldIndex).Value
                    Next llFldIndex

                    ladrs.Update

                    .MoveNext
                Next llIndex
            End If

            lbPopulated = False
            If (IsRecordsetPopulated(ladrs)) Then
                ladrs.MoveFirst

                lbPopulated = True
            End If

            .Filter = adFilterNone

            .Filter = MDAC25FilterFix("APPR_LEVEL_IND = 20")
            If Not (.EOF) Then
                .MoveFirst

                For llIndex = 1 To .RecordCount
                    lbAddRecord = True

                    If (lbPopulated) Then
                        ladrs.MoveFirst

                        ladrs.Find "DEPT_NUM = '" & Trim(!DEPT_NUM) & "'"
                        If Not (ladrs.EOF) Then
                            lbAddRecord = False
                        End If

                        ladrs.MoveFirst
                    End If

                    If (lbAddRecord) Then
                        ladrs.AddNew

                        For llFldIndex = 0 To .Fields.Count - 1
                            ladrs.Fields.Item(llFldIndex).Value _
                                = .Fields.Item(llFldIndex).Value
                        Next llFldIndex

                        ladrs.Update
                    End If

                    .MoveNext
                Next llIndex
            End If

            If Not (lbPopulated) Then
                If (IsRecordsetPopulated(ladrs)) Then
                    ladrs.MoveFirst

                    lbPopulated = True
                End If
            End If

            .Filter = adFilterNone

            .Filter = MDAC25FilterFix("APPR_LEVEL_IND = 10")
            If Not (.EOF) Then
                .MoveFirst

                For llIndex = 1 To .RecordCount
                    lbAddRecord = True

                    If (lbPopulated) Then
                        ladrs.MoveFirst

                        ladrs.Find "DEPT_NUM = '" & Trim(!DEPT_NUM) & "'"
                        If Not (ladrs.EOF) Then
                            lbAddRecord = False
                        End If

                        ladrs.MoveFirst
                    End If

                    If (lbAddRecord) Then
                        ladrs.AddNew

                        For llFldIndex = 0 To .Fields.Count - 1
                            ladrs.Fields.Item(llFldIndex).Value _
                                = .Fields.Item(llFldIndex).Value
                        Next llFldIndex

                        ladrs.Update
                    End If

                    .MoveNext
                Next llIndex
            End If

            .Filter = adFilterNone
        End With
    End If

    If (IsRecordsetPopulated(ladrs)) Then
        ladrs.MoveFirst

        Set EliminateDepartmentDuplicates = ladrs
    End If

    Set ladrs = Nothing

    Exit Function
EliminateDepartmentDuplicates_EH:
    Set ladrs = Nothing

    TrsRaiseError MODULENAME _
        , "EliminateDepartmentDuplicates"

End Function
Private Function EliminateFunctionDuplicates(ByVal vadrsSource As ADODB.Recordset _
                    ) As ADODB.Recordset
On Error GoTo EliminateFunctionDuplicates_EH
Dim ladrs As ADODB.Recordset
Dim llIndex As Long, lbAddRecord As Boolean
Dim lbPopulated As Boolean
'RS Fields
Dim lsFldName As String, llFldIndex As Long
Dim llFldSize As Long, ladFldType As ADODB.DataTypeEnum
Dim ladFldAttrib As ADODB.FieldAttributeEnum

    Set EliminateFunctionDuplicates = vadrsSource

    If (IsRecordsetPopulated(vadrsSource)) Then
        Set ladrs = New ADODB.Recordset

        With vadrsSource
            For llIndex = 0 To .Fields.Count - 1
                'Get the fld name, size
                lsFldName = .Fields.Item(llIndex).Name
                ladFldType = .Fields.Item(llIndex).Type
                llFldSize = .Fields.Item(llIndex).DefinedSize
                ladFldAttrib = .Fields.Item(llIndex).Attributes

                If (llFldSize > 0) Then
                    ladrs.Fields.Append lsFldName _
                        , ladFldType _
                        , llFldSize _
                        , ladFldAttrib
                Else
                    ladrs.Fields.Append lsFldName _
                        , ladFldType _
                        , _
                        , ladFldAttrib
                End If
            Next llIndex

            ladrs.CursorLocation = adUseClient
            ladrs.LockType = adLockBatchOptimistic
            ladrs.CursorType = adOpenStatic
            ladrs.Open

            .Filter = adFilterNone

            .Filter = MDAC25FilterFix("APPR_LEVEL_IND = 20")
            If Not (.EOF) Then
                .MoveFirst

                For llIndex = 1 To .RecordCount
                    ladrs.AddNew

                    For llFldIndex = 0 To .Fields.Count - 1
                        ladrs.Fields.Item(llFldIndex).Value _
                            = .Fields.Item(llFldIndex).Value
                    Next llFldIndex

                    ladrs.Update

                    .MoveNext
                Next llIndex
            End If

            lbPopulated = False
            If (IsRecordsetPopulated(ladrs)) Then
                ladrs.MoveFirst

                lbPopulated = True
            End If

            .Filter = adFilterNone

            .Filter = MDAC25FilterFix("APPR_LEVEL_IND = 10")
            If Not (.EOF) Then
                .MoveFirst

                For llIndex = 1 To .RecordCount
                    lbAddRecord = True

                    If (lbPopulated) Then
                        ladrs.MoveFirst

                        ladrs.Find "FUNCTION_NUM = '" & Trim(!FUNCTION_NUM) & "'"
                        If Not (ladrs.EOF) Then
                            lbAddRecord = False
                        End If

                        ladrs.MoveFirst
                    End If

                    If (lbAddRecord) Then
                        ladrs.AddNew

                        For llFldIndex = 0 To .Fields.Count - 1
                            ladrs.Fields.Item(llFldIndex).Value _
                                = .Fields.Item(llFldIndex).Value
                        Next llFldIndex

                        ladrs.Update
                    End If

                    .MoveNext
                Next llIndex
            End If

            .Filter = adFilterNone
        End With
    End If

    If (IsRecordsetPopulated(ladrs)) Then
        ladrs.MoveFirst

        Set EliminateFunctionDuplicates = ladrs
    End If

    Set ladrs = Nothing

    Exit Function
EliminateFunctionDuplicates_EH:
    Set ladrs = Nothing

    TrsRaiseError MODULENAME _
        , "EliminateFunctionDuplicates"

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
