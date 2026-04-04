VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "cbsOTTemplateTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const m_strCMODULENAME As String = "cbsOTTEmplateTrans"

Private m_strUserID As String
Private m_lngUserNum As Long
Private m_lngTemplateNum As Long
Private m_rsErrors As ADODB.Recordset

Private m_rsTemplateBase As ADODB.Recordset
Private m_rsSortOrder As ADODB.Recordset
Private m_rsForceOrder As ADODB.Recordset
Private m_rsGroupOrder As ADODB.Recordset
Private m_rsOTResponses As ADODB.Recordset
Private m_rsIneligibleRules As ADODB.Recordset
Private m_rsShifts As ADODB.Recordset

Private Const CANVASS_ORDER As String = "C"
Private Const FORCE_ORDER As String = "F"
Private Const GROUP_ORDER As String = "G"


Public Function SaveOvertimeTemplate( _
                    ByRef r_lngTemplateNum As Long, _
                    ByVal v_rsTemplateBase As ADODB.Recordset, _
                    ByVal v_rsSortOrder As ADODB.Recordset, _
                    ByVal v_rsForceOrder As ADODB.Recordset, _
                    ByVal v_rsGroupOrder As ADODB.Recordset, _
                    ByVal v_rsOTResponses As ADODB.Recordset, _
                    ByVal v_rsIneligibleRules As ADODB.Recordset, _
                    ByVal v_rsShifts As ADODB.Recordset, _
                    ByVal v_strUserID As String, _
                    ByRef r_rsBRVs As ADODB.Recordset _
                ) As Boolean

    Dim l_objCtx As ObjectContext
    Dim l_objErrors As ADODB.Recordset
    Dim l_blnOK As Boolean
    
    On Error GoTo eh
    
    m_strUserID = v_strUserID
    m_lngTemplateNum = r_lngTemplateNum
    m_lngUserNum = ParseUserID(m_strUserID)
    
    Set l_objCtx = GetObjectContext
    Set l_objErrors = NewTemplateErrorset
    
    If ARequestExist( _
            v_rsTemplateBase, _
            v_rsSortOrder, v_rsForceOrder, v_rsGroupOrder, _
            v_rsOTResponses, v_rsIneligibleRules, v_rsShifts) Then
        'continue
    Else
        SaveOvertimeTemplate = True
        l_objCtx.SetComplete
        GoTo xit
    End If

'    If Not PreValidateTemplateBase(v_rsTemplateBase, l_objErrors) Or _
'        Not PreValidateSortOrders(v_rsSortOrder, v_rsForceOrder, v_rsGroupOrder, l_objErrors) Or _
'        Not PreValidateResponses(v_rsOTResponses, l_objErrors) Or _
'        Not PreValidateRules(v_rsIneligibleRules, l_objErrors) Or _
'        Not PreValidateShifts(v_rsShifts, l_objErrors) Then
'            l_objCtx.SetAbort
'            SaveOvertimeTemplate = False
'            GoTo xit
'    End If

    If Not PreValidateShifts(v_rsShifts, l_objErrors) Then
        l_objCtx.SetAbort
        SaveOvertimeTemplate = False
        GoTo xit
    End If

'    GetOvertimeTemplateData m_lngTemplateNum, m_rsTemplateBase, _
'            m_rsSortOrder, m_rsForceOrder, m_rsGroupOrder, _
'            m_rsOTResponses, m_rsIneligibleRules, m_rsShifts
    GetOvertimeTemplateData m_lngTemplateNum, v_rsTemplateBase, _
            v_rsSortOrder, v_rsForceOrder, v_rsGroupOrder, _
            v_rsOTResponses, v_rsIneligibleRules, v_rsShifts
    
    If ProcessTemplateAttributes(v_rsTemplateBase) Then
        'continue
    Else
        SaveOvertimeTemplate = False
        l_objCtx.SetAbort
        GoTo xit
    End If
    
    If r_lngTemplateNum = -1 Then
        SetNewTemplateNumber v_rsTemplateBase, _
            v_rsSortOrder, v_rsForceOrder, v_rsGroupOrder, _
            v_rsOTResponses, v_rsIneligibleRules, v_rsShifts
    End If


    If ProcessOrderAttributes(v_rsSortOrder, m_rsSortOrder, CANVASS_ORDER) Then
        'continue
    Else
        SaveOvertimeTemplate = False
        l_objCtx.SetAbort
        GoTo xit
    End If
    

    If ProcessOrderAttributes(v_rsForceOrder, m_rsForceOrder, FORCE_ORDER) Then
        'continue
    Else
        SaveOvertimeTemplate = False
        l_objCtx.SetAbort
        GoTo xit
    End If
    
    
    If ProcessOrderAttributes(v_rsGroupOrder, m_rsGroupOrder, GROUP_ORDER) Then
        'continue
    Else
        SaveOvertimeTemplate = False
        l_objCtx.SetAbort
        GoTo xit
    End If
    
    
    If ProcessOTResponses(v_rsOTResponses, m_rsOTResponses) Then
        'continue
    Else
        SaveOvertimeTemplate = False
        l_objCtx.SetAbort
        GoTo xit
    End If
    
    If ProcessIneligibleRules(v_rsIneligibleRules, m_rsIneligibleRules) Then
        'continue
    Else
        SaveOvertimeTemplate = False
        l_objCtx.SetAbort
        GoTo xit
    End If
    
    If ProcessShifts(v_rsShifts, m_rsShifts) Then
        'continue
    Else
        SaveOvertimeTemplate = False
        l_objCtx.SetAbort
        GoTo xit
    End If

'
'   Pass all template recordsets to data component to ensure same
'   connection is used.
'
    If ProcessOvertimeTemplate(m_lngTemplateNum, v_rsTemplateBase, _
        v_rsSortOrder, v_rsForceOrder, v_rsGroupOrder, _
        v_rsOTResponses, v_rsIneligibleRules, v_rsShifts) Then
        'continue
    Else
        SaveOvertimeTemplate = False
        l_objCtx.SetAbort
        GoTo xit
    End If
    
    
    r_lngTemplateNum = m_lngTemplateNum
    SaveOvertimeTemplate = True
    l_objCtx.SetComplete

    
xit:
    If Not m_rsErrors Is Nothing Then
        If m_rsErrors.RecordCount = 0 Then
            Set r_rsBRVs = Nothing
        Else
            Set r_rsBRVs = m_rsErrors
        End If
    Else
        Set r_rsBRVs = Nothing
    End If
    FreeGlobalRS
    Set l_objCtx = Nothing
    Exit Function
    
eh:
    l_objCtx.SetAbort
    Set l_objCtx = Nothing
    TrsRaiseError m_strCMODULENAME, "SaveOvertimeTemplate", , , , , True
           
End Function


Public Function DeleteTemplate( _
                    ByVal v_lngTemplateNum As Long, _
                    ByVal v_strUserID As String) As Boolean
    Dim l_obj As dbOTTemplate.cdOTTemplate
On Error GoTo eh

    m_strUserID = v_strUserID
    m_lngTemplateNum = v_lngTemplateNum
    m_lngUserNum = ParseUserID(m_strUserID)

    Set l_obj = CreateObject("dbOTTemplate.cdOTTemplate")
    Set m_rsTemplateBase = l_obj.GetOvertimeTemplateBase(v_lngTemplateNum)
    Set l_obj = Nothing

    If Not m_rsTemplateBase.EOF Then
        m_rsTemplateBase!TEMPL_STATUS = "I"
        m_rsTemplateBase!LAST_MOD_EMP_NUM = CLng(m_strUserID)
    Else
        Err.Raise trErrbusOTTemplNotFound
    End If

    If Not ProcessOvertimeTemplate(v_lngTemplateNum, m_rsTemplateBase, _
                Nothing, Nothing, Nothing, Nothing, Nothing, Nothing) Then
        Err.Raise trErrbusSaveOTTemplError
    End If
    
    DeleteTemplate = True
    GetObjectContext.SetComplete
    Set m_rsTemplateBase = Nothing

    Exit Function

eh:
    Dim l_colArgs As Collection
    Dim l_lngErrNum As Long
    
    GetObjectContext.SetAbort
    Set l_obj = Nothing
    Set m_rsTemplateBase = Nothing

    l_lngErrNum = Err.Number
    Select Case l_lngErrNum
        Case trErrbusOTTemplNotFound, trErrbusSaveOTTemplError
            Set l_colArgs = New Collection
            l_colArgs.Add CStr(v_lngTemplateNum)
            Err.Source = FormatOTTemplateBRV(TrType_OT_TEMPL, v_lngTemplateNum)
            TrsRaiseError m_strCMODULENAME, "DeleteTemplate", l_lngErrNum, , l_colArgs, , True
        Case Else
            Set l_colArgs = New Collection
            l_colArgs.Add CStr(v_lngTemplateNum)
            TrsRaiseError m_strCMODULENAME, "DeleteTemplate", , , , l_colArgs, True
    End Select
    
End Function


Private Function ProcessTemplateAttributes( _
                    ByVal v_rsTemplateBase As ADODB.Recordset _
                ) As Boolean
    
    Dim l_blnError As Boolean
    
    On Error GoTo eh
     
    If RecordsExist(v_rsTemplateBase) Then
        v_rsTemplateBase.MoveFirst
    Else
        ProcessTemplateAttributes = True
        Exit Function
    End If
    
    If m_lngTemplateNum = -1 Then
        m_lngTemplateNum = GetNextTemplateNum()
        If m_lngTemplateNum = -1 Then
            l_blnError = AppendTemplateBaseBRV(v_rsTemplateBase, 30000, mstrCBUSINESSERR, "System error, unable to create overtime template.  Please contact TRIS Administrator")
            ProcessTemplateAttributes = False
            Exit Function
        End If
    End If
    
    If Not PreValidateTemplateBase(v_rsTemplateBase) Then
        ProcessTemplateAttributes = False
        Exit Function
    End If

    If v_rsTemplateBase!RECORD_STATUS <> "I" Then
        m_rsTemplateBase.Filter = "TIME_STAMP = '" & _
            v_rsTemplateBase!TIME_STAMP & "'"
        If m_rsTemplateBase.EOF Then
            l_blnError = AppendTemplateBaseBRV(v_rsTemplateBase, 30000, mstrCBUSINESSERR, "Save conflict with another user.  Please refresh and try again.")
            ProcessTemplateAttributes = False
            Exit Function
'            MsgBox "commit conflict ProcessTemplateAttributes"
            'error
        End If
    End If
    
    SetTemplateAttributes v_rsTemplateBase
    
    m_rsTemplateBase.Filter = ""
    ProcessTemplateAttributes = True
    
    Exit Function
    
eh:
    
    TrsRaiseError m_strCMODULENAME, "ProcessTemplateAttributes"
    
    
End Function


'Private Function ProcessSortAttributes( _
'                    ByVal v_rsSortOrder As ADODB.Recordset, _
'                    ByVal v_rsForceOrder As ADODB.Recordset _
'                ) As Boolean
'
'    Dim l_blnError As Boolean
'
'    On Error GoTo eh
'
'    If RecordsExist(v_rsSortOrder) Then
'        v_rsSortOrder.MoveFirst
'    Else
'        ProcessSortAttributes = True
'        Exit Function
'    End If
'
'    If Not PreValidateSort(v_rsSortOrder) Then
'        ProcessSortAttributes = False
'        Exit Function
'    End If
'
'    v_rsSortOrder.Filter = MDAC25FilterFix("RECORD_STATUS = 'D'")
'    If Not v_rsSortOrder.EOF Then
'        Do Until v_rsSortOrder.EOF
'            FindSortRecord m_rsSortOrder, v_rsSortOrder!KEY_SORT_CD, _
'                v_rsSortOrder!KEY_SORT_TYPE_CD, v_rsSortOrder!LAST_MOD_TIMESTAMP
'            If Not m_rsSortOrder.EOF Then
'                m_rsSortOrder.Delete
'            Else
'                'error
'            End If
'            v_rsSortOrder.MoveNext
'        Loop
'    End If
'
'    v_rsSortOrder.Filter = MDAC25FilterFix("RECORD_STATUS = 'I'")
'    If Not v_rsSortOrder.EOF Then
'        Do Until v_rsSortOrder.EOF
'            m_rsSortOrder.AddNew
'            m_rsSortOrder!TEMPLATE_NUM = m_lngTemplateNum
'            m_rsSortOrder!CREATED_EMP_NUM = CLng(m_strUserID)
'            SetSortAttributes v_rsSortOrder, m_rsSortOrder
'            v_rsSortOrder.MoveNext
'        Loop
'    End If
'
'    v_rsSortOrder.Filter = MDAC25FilterFix("RECORD_STATUS = 'C'")
'    If Not v_rsSortOrder.EOF Then
'        Do Until v_rsSortOrder.EOF
'            FindSortRecord m_rsSortOrder, v_rsSortOrder!KEY_SORT_CD, _
'                v_rsSortOrder!KEY_SORT_TYPE_CD, v_rsSortOrder!LAST_MOD_TIMESTAMP
'            If Not m_rsSortOrder.EOF Then
'                SetSortAttributes v_rsSortOrder, m_rsSortOrder
'            Else
'                'error
'            End If
'            v_rsSortOrder.MoveNext
'        Loop
'    End If
'
'    m_rsSortOrder.Filter = ""
'    v_rsSortOrder.Filter = ""
'    ProcessSortAttributes = True
'
'    Exit Function
'
'eh:
'
'    TrsRaiseError m_strCMODULENAME, "ProcessSortAttributes"
'
'End Function
'
'
'Private Function ProcessForceAttributes( _
'                    ByVal v_rsSortOrder As ADODB.Recordset, _
'                    ByVal v_rsForceOrder As ADODB.Recordset _
'                ) As Boolean
'
'    Dim l_blnError As Boolean
'
'    On Error GoTo eh
'
'    If RecordsExist(v_rsForceOrder) Then
'        v_rsForceOrder.MoveFirst
'    Else
'        ProcessForceAttributes = True
'        Exit Function
'    End If
'
'    If Not PreValidateSort(v_rsForceOrder) Then
'        ProcessForceAttributes = False
'        Exit Function
'    End If
'
'    v_rsForceOrder.Filter = MDAC25FilterFix("RECORD_STATUS = 'D'")
'    If Not v_rsForceOrder.EOF Then
'        Do Until v_rsForceOrder.EOF
'            FindSortRecord m_rsForceOrder, v_rsForceOrder!KEY_SORT_CD, v_rsForceOrder!KEY_SORT_TYPE_CD
'            If Not m_rsForceOrder.EOF Then
'                m_rsForceOrder.Delete
'            Else
'                'error
'            End If
'            v_rsForceOrder.MoveNext
'        Loop
'    End If
'
'    v_rsForceOrder.Filter = MDAC25FilterFix("RECORD_STATUS = 'I'")
'    If Not v_rsForceOrder.EOF Then
'        Do Until v_rsForceOrder.EOF
'            m_rsForceOrder.AddNew
'            m_rsForceOrder!TEMPLATE_NUM = m_lngTemplateNum
'            m_rsForceOrder!CREATED_EMP_NUM = CLng(m_strUserID)
'            SetSortAttributes v_rsForceOrder
'            v_rsForceOrder.MoveNext
'        Loop
'    End If
'
'    v_rsForceOrder.Filter = MDAC25FilterFix("RECORD_STATUS = 'C'")
'    If Not v_rsForceOrder.EOF Then
'        Do Until v_rsForceOrder.EOF
'            FindSortRecord m_rsForceOrder, v_rsForceOrder!KEY_SORT_CD, v_rsForceOrder!KEY_SORT_TYPE_CD
'            If Not m_rsForceOrder.EOF Then
'                SetSortAttributes v_rsForceOrder
'            Else
'                'error
'            End If
'            v_rsForceOrder.MoveNext
'        Loop
'    End If
'
'    m_rsForceOrder.Filter = ""
'    v_rsForceOrder.Filter = ""
'    ProcessForceAttributes = True
'
'    Exit Function
'
'eh:
'
'    TrsRaiseError m_strCMODULENAME, "ProcessForceAttributes"
'
'End Function


Private Function ProcessOrderAttributes( _
                    ByVal v_rsOrderRequest As ADODB.Recordset, _
                    ByVal v_rsOrderPersistent As ADODB.Recordset, _
                    ByVal v_strSortTypeCd As String _
                ) As Boolean
    
' v_strSortTypeCd values C, F, G
    
    Dim l_blnError As Boolean
    
    On Error GoTo eh
     
    If RecordsExist(v_rsOrderRequest) Then
        v_rsOrderRequest.MoveFirst
    Else
        ProcessOrderAttributes = True
        Exit Function
    End If
    
    If Not PreValidateSort(v_rsOrderRequest) Then
        ProcessOrderAttributes = False
        Exit Function
    End If

    v_rsOrderRequest.Filter = MDAC25FilterFix("RECORD_STATUS = 'D'")
    If Not v_rsOrderRequest.EOF Then
        Do Until v_rsOrderRequest.EOF
            FindSortRecord v_rsOrderPersistent, v_rsOrderRequest!KEY_SORT_CD, _
                v_strSortTypeCd, v_rsOrderRequest!LAST_MOD_TIMESTAMP
            If Not v_rsOrderPersistent.EOF Then
                v_rsOrderPersistent.Delete
            Else
                MsgBox "Concurrency conflict"
                'error
            End If
            v_rsOrderRequest.MoveNext
        Loop
    End If
    
    v_rsOrderRequest.Filter = MDAC25FilterFix("RECORD_STATUS = 'I'")
    If Not v_rsOrderRequest.EOF Then
        Do Until v_rsOrderRequest.EOF
            v_rsOrderPersistent.AddNew
            v_rsOrderPersistent!TEMPLATE_NUM = m_lngTemplateNum
            v_rsOrderPersistent!SORT_TYPE_CD = v_strSortTypeCd
            v_rsOrderPersistent!CREATED_EMP_NUM = CLng(m_strUserID)
            SetSortAttributes v_rsOrderRequest, v_rsOrderPersistent
            v_rsOrderRequest.MoveNext
        Loop
    End If
    
    v_rsOrderRequest.Filter = MDAC25FilterFix("RECORD_STATUS = 'C'")
    If Not v_rsOrderRequest.EOF Then
        Do Until v_rsOrderRequest.EOF
            FindSortRecord v_rsOrderPersistent, v_rsOrderRequest!KEY_SORT_CD, _
                v_strSortTypeCd, v_rsOrderRequest!LAST_MOD_TIMESTAMP
            If Not v_rsOrderPersistent.EOF Then
                SetSortAttributes v_rsOrderRequest, v_rsOrderPersistent
            Else
                MsgBox "Concurrency conflict"
                'error
            End If
            v_rsOrderRequest.MoveNext
        Loop
    End If
    
    v_rsOrderPersistent.Filter = ""
    v_rsOrderRequest.Filter = ""
    ProcessOrderAttributes = True
    
    Exit Function
    
eh:
    
    TrsRaiseError m_strCMODULENAME, "ProcessOrderAttributes"
    
End Function



Private Sub FindSortRecord( _
                ByRef r_rsOrder As ADODB.Recordset, _
                ByVal v_lngSortCode As Long, _
                ByVal v_strSortTypeCode As String, _
                ByVal v_strLastModTimestamp As String)
    Dim l_strFind As String
    
On Error GoTo eh

    l_strFind = "SORT_CD = " & v_lngSortCode & " AND SORT_TYPE_CD = '" _
        & v_strSortTypeCode & "' AND LAST_MOD_TIMESTAMP = '" & v_strLastModTimestamp & "'"
    r_rsOrder.Filter = l_strFind
   
    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "FindSortRecord"
    
End Sub


Private Sub FindResponseRecord( _
                ByRef r_rsOTResponses As ADODB.Recordset, _
                ByVal v_strResponseCode As String)
    Dim l_strFind As String
    
On Error GoTo eh

    l_strFind = "OT_RESP_CD = '" & v_strResponseCode & "'"
    r_rsOTResponses.Filter = l_strFind
   
    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "FindResponseRecord"
    
End Sub


Private Sub FindIneligibleRulesRecord( _
                ByRef r_rsIneligibleRules As ADODB.Recordset, _
                ByVal v_strIneligibleCode As String)
    Dim l_strFind As String
    
On Error GoTo eh

    l_strFind = "INELG_CD = '" & v_strIneligibleCode & "'"
    r_rsIneligibleRules.Filter = l_strFind
   
    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "FindIneligibleRulesRecord"
    
End Sub


Private Sub FindShiftsRecord( _
                ByRef r_rsShifts As ADODB.Recordset, _
                ByVal v_strShiftCode As String, _
                ByVal v_strLastModTimestamp As String)
    Dim l_strFind As String
    
On Error GoTo eh

    l_strFind = "SHIFT_CD = '" & v_strShiftCode & "'" & _
        " and LAST_MOD_TIMESTAMP = '" & v_strLastModTimestamp & "'"
    r_rsShifts.Filter = l_strFind
   
    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "FindShiftsRecord"
    
End Sub


Private Function ProcessOTResponses( _
                    ByVal v_rsResponseRequest As ADODB.Recordset, _
                    ByVal v_rsResponsePersistent As ADODB.Recordset _
                ) As Boolean
    
    Dim l_blnError As Boolean
    
On Error GoTo eh
     
    If RecordsExist(v_rsResponseRequest) Then
        v_rsResponseRequest.MoveFirst
    Else
        ProcessOTResponses = True
        Exit Function
    End If
    
    If Not PreValidateResponses(v_rsResponseRequest) Then
        ProcessOTResponses = False
        Exit Function
    End If

    v_rsResponseRequest.Filter = MDAC25FilterFix("RECORD_STATUS = 'D'")
    If Not v_rsResponseRequest.EOF Then
        Do Until v_rsResponseRequest.EOF
            FindResponseRecord v_rsResponsePersistent, v_rsResponseRequest!KEY_OT_RESP_CD
            If Not v_rsResponsePersistent.EOF Then
                v_rsResponsePersistent.Delete
            Else
                'error
            End If
            v_rsResponseRequest.MoveNext
        Loop
    End If
    
    v_rsResponseRequest.Filter = MDAC25FilterFix("RECORD_STATUS = 'I'")
    If Not v_rsResponseRequest.EOF Then
        Do Until v_rsResponseRequest.EOF
            v_rsResponsePersistent.AddNew
            v_rsResponsePersistent!TEMPLATE_NUM = m_lngTemplateNum
            v_rsResponsePersistent!OT_RESP_CD = v_rsResponseRequest!OT_RESP_CD
            v_rsResponsePersistent!CREATED_EMP_NUM = CLng(m_strUserID)
            v_rsResponsePersistent!LAST_MOD_EMP_NUM = CLng(m_strUserID)
            v_rsResponsePersistent.Update
            v_rsResponseRequest.MoveNext
        Loop
    End If
    
'    v_rsResponseRequest.Filter = MDAC25FilterFix("RECORD_STATUS = 'C'")
'    If Not v_rsResponseRequest.EOF Then
'        Do Until v_rsResponseRequest.EOF
'            FindResponseRecord v_rsResponsePersistent, v_rsResponseRequest!KEY_OT_RESP_CD
'            If Not v_rsResponsePersistent.EOF Then
'                SetResponses v_rsResponseRequest, v_rsResponsePersistent
'            Else
'                'error
'            End If
'            v_rsResponseRequest.MoveNext
'        Loop
'    End If
    
    v_rsResponsePersistent.Filter = ""
    v_rsResponseRequest.Filter = ""
    ProcessOTResponses = True
    
    Exit Function
    
eh:
    
    TrsRaiseError m_strCMODULENAME, "ProcessOTResponses"
    
End Function


Private Function ProcessIneligibleRules( _
                    ByVal v_rsIneligibleRulesRequest As ADODB.Recordset, _
                    ByVal v_rsIneligibleRulesPersistent As ADODB.Recordset _
                ) As Boolean
    
    Dim l_blnError As Boolean
    
    On Error GoTo eh
     
    If RecordsExist(v_rsIneligibleRulesRequest) Then
        v_rsIneligibleRulesRequest.MoveFirst
    Else
        ProcessIneligibleRules = True
        Exit Function
    End If
    
'    If Not PreValidateIneligibleRules(v_rsIneligibleRulesRequest) Then
'        ProcessIneligibleRules = False
'        Exit Function
'    End If

    v_rsIneligibleRulesRequest.Filter = MDAC25FilterFix("RECORD_STATUS = 'D'")
    If Not v_rsIneligibleRulesRequest.EOF Then
        Do Until v_rsIneligibleRulesRequest.EOF
            FindIneligibleRulesRecord v_rsIneligibleRulesPersistent, v_rsIneligibleRulesRequest!KEY_INELG_CD
            If Not v_rsIneligibleRulesPersistent.EOF Then
                v_rsIneligibleRulesPersistent.Delete
            Else
                'error
            End If
            v_rsIneligibleRulesRequest.MoveNext
        Loop
    End If
    
    v_rsIneligibleRulesRequest.Filter = MDAC25FilterFix("RECORD_STATUS = 'I'")
    If Not v_rsIneligibleRulesRequest.EOF Then
        Do Until v_rsIneligibleRulesRequest.EOF
            v_rsIneligibleRulesPersistent.AddNew
            v_rsIneligibleRulesPersistent!TEMPLATE_NUM = m_lngTemplateNum
            v_rsIneligibleRulesPersistent!INELG_CD = v_rsIneligibleRulesRequest!INELG_CD
            v_rsIneligibleRulesPersistent!CREATED_EMP_NUM = CLng(m_strUserID)
            v_rsIneligibleRulesPersistent!LAST_MOD_EMP_NUM = CLng(m_strUserID)
            v_rsIneligibleRulesPersistent.Update
            v_rsIneligibleRulesRequest.MoveNext
        Loop
    End If
       
    v_rsIneligibleRulesPersistent.Filter = ""
    v_rsIneligibleRulesRequest.Filter = ""
    ProcessIneligibleRules = True
    
    Exit Function
    
eh:
    
    TrsRaiseError m_strCMODULENAME, "ProcessIneligibleRules"
    
End Function


Private Function ProcessShifts( _
                    ByVal v_rsShiftsRequest As ADODB.Recordset, _
                    ByVal v_rsShiftsPersistent As ADODB.Recordset _
                ) As Boolean
    
    Dim l_blnError As Boolean
    
    On Error GoTo eh
     
    If RecordsExist(v_rsShiftsRequest) Then
        v_rsShiftsRequest.MoveFirst
    Else
        ProcessShifts = True
        Exit Function
    End If
    
'    If Not PreValidateShifts(v_rsShiftsRequest) Then
'        ProcessShifts = False
'        Exit Function
'    End If

    v_rsShiftsRequest.Filter = MDAC25FilterFix("RECORD_STATUS = 'D'")
    If Not v_rsShiftsRequest.EOF Then
        Do Until v_rsShiftsRequest.EOF
            FindShiftsRecord v_rsShiftsPersistent, v_rsShiftsRequest!KEY_SHIFT_CD, v_rsShiftsRequest!LAST_MOD_TIMESTAMP
            If Not v_rsShiftsPersistent.EOF Then
                v_rsShiftsPersistent.Delete
            Else
                'error
            End If
            v_rsShiftsRequest.MoveNext
        Loop
    End If
    
    v_rsShiftsRequest.Filter = MDAC25FilterFix("RECORD_STATUS = 'I'")
    If Not v_rsShiftsRequest.EOF Then
        Do Until v_rsShiftsRequest.EOF
            v_rsShiftsPersistent.AddNew
            v_rsShiftsPersistent!TEMPLATE_NUM = m_lngTemplateNum
            v_rsShiftsPersistent!SHIFT_CD = v_rsShiftsRequest!SHIFT_CD
            v_rsShiftsPersistent!START_TIME = v_rsShiftsRequest!START_TIME
            v_rsShiftsPersistent!END_TIME = v_rsShiftsRequest!END_TIME
            v_rsShiftsPersistent!LUNCH_START_TIME = v_rsShiftsRequest!LUNCH_START_TIME
            v_rsShiftsPersistent!LUNCH_END_TIME = v_rsShiftsRequest!LUNCH_END_TIME
            v_rsShiftsPersistent!TURNOVER_START_TM = v_rsShiftsRequest!TURNOVER_START_TM
            v_rsShiftsPersistent!TURNOVER_END_TM = v_rsShiftsRequest!TURNOVER_END_TM
            v_rsShiftsPersistent!COMP_SHIFT_CD = v_rsShiftsRequest!COMP_SHIFT_CD
            v_rsShiftsPersistent!PRIOR_DAY_IND = v_rsShiftsRequest!PRIOR_DAY_IND
            v_rsShiftsPersistent!CREATED_EMP_NUM = CLng(m_strUserID)
            v_rsShiftsPersistent!LAST_MOD_EMP_NUM = CLng(m_strUserID)
            v_rsShiftsPersistent.Update
            v_rsShiftsRequest.MoveNext
        Loop
    End If
       
    v_rsShiftsRequest.Filter = MDAC25FilterFix("RECORD_STATUS = 'C'")
    If Not v_rsShiftsRequest.EOF Then
        Do Until v_rsShiftsRequest.EOF
            FindShiftsRecord v_rsShiftsPersistent, v_rsShiftsRequest!KEY_SHIFT_CD, v_rsShiftsRequest!LAST_MOD_TIMESTAMP
            If Not v_rsShiftsPersistent.EOF Then
                SetShiftAttributes v_rsShiftsRequest, v_rsShiftsPersistent
            Else
                'error
            End If
            v_rsShiftsRequest.MoveNext
        Loop
    End If
       
    v_rsShiftsPersistent.Filter = ""
    v_rsShiftsRequest.Filter = ""
    ProcessShifts = True
    
    Exit Function
    
eh:
    
    TrsRaiseError m_strCMODULENAME, "ProcessShifts"
    
End Function


Private Function PreValidateTemplateBase( _
                    ByVal v_rsTemplateBase As ADODB.Recordset _
                ) As Boolean

On Error GoTo eh

' validate location code
'   and refuse rule num


    PreValidateTemplateBase = True
    Exit Function
    
eh:
    
    TrsRaiseError m_strCMODULENAME, "PreValidateTemplateBase"

End Function


Private Function PreValidateSort( _
                    ByVal v_rsSortOrder As ADODB.Recordset _
                ) As Boolean

On Error GoTo eh

' validate location code
'   and refuse rule num


    PreValidateSort = True
    Exit Function
    
eh:
    
    TrsRaiseError m_strCMODULENAME, "PreValidateSort"

End Function

Private Sub SetTemplateAttributes( _
                ByVal v_rsTemplateBase As ADODB.Recordset)
    
    On Error GoTo eh
    
    If v_rsTemplateBase!RECORD_STATUS = "I" Then
        m_rsTemplateBase.AddNew
        m_rsTemplateBase!TEMPLATE_NUM = m_lngTemplateNum
        m_rsTemplateBase!HRS_OVERRIDE_IND = "Y"
    End If
    
    If v_rsTemplateBase!RECORD_STATUS = "D" Then
        m_rsTemplateBase!TEMPL_STATUS = "I"
    End If
    
    m_rsTemplateBase!CANVASS_DESC = v_rsTemplateBase!CANVASS_DESC
    m_rsTemplateBase!TEMPL_LOCATION = v_rsTemplateBase!TEMPL_LOCATION
    m_rsTemplateBase!REFUSE_RULE_NUM = v_rsTemplateBase!REFUSE_RULE_NUM
    
    m_rsTemplateBase!LAST_MOD_EMP_NUM = CLng(m_strUserID)
    m_rsTemplateBase.Update
    
    Exit Sub
    
eh:
    
  TrsRaiseError m_strCMODULENAME, "SetTemplateAttributes"
    
End Sub

'
'Private Sub SetSortAttributes( _
'                ByVal v_rsSortOrder As ADODB.Recordset)
'
'    On Error GoTo eh
'
''    If v_rsSortOrder!RECORD_STATUS = "I" Then
''        m_rsSortOrder.AddNew
''        m_rsSortOrder!TEMPLATE_NUM = m_lngTemplateNum
''        m_rsSortOrder!CREATED_EMP_NUM = CLng(m_strUserID)
''    End If
'    m_rsSortOrder!SORT_CD = v_rsSortOrder!SORT_CD
'    m_rsSortOrder!SORT_TYPE_CD = v_rsSortOrder!SORT_TYPE_CD
'    m_rsSortOrder!SORT_ORDER = v_rsSortOrder!SORT_ORDER
'    m_rsSortOrder!LAST_MOD_EMP_NUM = CLng(m_strUserID)
'    m_rsSortOrder.Update
'
'    Exit Sub
'
'eh:
'
'  TrsRaiseError m_strCMODULENAME, "SetSortAttributes"
'
'End Sub


Private Sub SetSortAttributes( _
                    ByVal v_rsOrderRequest As ADODB.Recordset, _
                    ByVal v_rsOrderPersistent As ADODB.Recordset)
    
On Error GoTo eh
    If Not IsEmpty(v_rsOrderRequest!SORT_CD) Then
        v_rsOrderPersistent!SORT_CD = v_rsOrderRequest!SORT_CD
    End If
'    If Not IsEmpty(v_rsOrderRequest!SORT_TYPE_CD) Then
'        v_rsOrderPersistent!SORT_TYPE_CD = v_rsOrderRequest!SORT_TYPE_CD
'    End If
    If Not IsEmpty(v_rsOrderRequest!SORT_ORDER) Then
        v_rsOrderPersistent!SORT_ORDER = v_rsOrderRequest!SORT_ORDER
    End If
    v_rsOrderPersistent!LAST_MOD_EMP_NUM = CLng(m_strUserID)
    v_rsOrderPersistent.Update

    Exit Sub
    
eh:
    
  TrsRaiseError m_strCMODULENAME, "SetSortAttributes"
    
End Sub


Private Sub SetShiftAttributes( _
                    ByVal v_rsShiftsRequest As ADODB.Recordset, _
                    ByVal v_rsShiftsPersistent As ADODB.Recordset)
    
    
On Error GoTo eh
    If Not IsEmpty(v_rsShiftsRequest!SHIFT_CD) Then
        v_rsShiftsPersistent!SHIFT_CD = v_rsShiftsRequest!SHIFT_CD
    End If
    If Not IsEmpty(v_rsShiftsRequest!START_TIME) Then
        v_rsShiftsPersistent!START_TIME = v_rsShiftsRequest!START_TIME
    End If
    If Not IsEmpty(v_rsShiftsRequest!END_TIME) Then
        v_rsShiftsPersistent!END_TIME = v_rsShiftsRequest!END_TIME
    End If
    If Not IsEmpty(v_rsShiftsRequest!LUNCH_START_TIME) Then
        v_rsShiftsPersistent!LUNCH_START_TIME = v_rsShiftsRequest!LUNCH_START_TIME
    End If
    If Not IsEmpty(v_rsShiftsRequest!LUNCH_END_TIME) Then
        v_rsShiftsPersistent!LUNCH_END_TIME = v_rsShiftsRequest!LUNCH_END_TIME
    End If
    If Not IsEmpty(v_rsShiftsRequest!TURNOVER_START_TM) Then
        v_rsShiftsPersistent!TURNOVER_START_TM = v_rsShiftsRequest!TURNOVER_START_TM
    End If
    If Not IsEmpty(v_rsShiftsRequest!TURNOVER_END_TM) Then
        v_rsShiftsPersistent!TURNOVER_END_TM = v_rsShiftsRequest!TURNOVER_END_TM
    End If
    If Not IsEmpty(v_rsShiftsRequest!COMP_SHIFT_CD) Then
        v_rsShiftsPersistent!COMP_SHIFT_CD = v_rsShiftsRequest!COMP_SHIFT_CD
    End If
    If Not IsEmpty(v_rsShiftsRequest!PRIOR_DAY_IND) Then
        v_rsShiftsPersistent!PRIOR_DAY_IND = v_rsShiftsRequest!PRIOR_DAY_IND
    End If
    
    v_rsShiftsPersistent!LAST_MOD_EMP_NUM = CLng(m_strUserID)
    v_rsShiftsPersistent.Update

    Exit Sub
    
eh:
    
  TrsRaiseError m_strCMODULENAME, "SetShiftAttributes"
    
End Sub




Private Function AppendTemplateBaseBRV( _
            ByVal v_rsTemplateBase As ADODB.Recordset, _
            ByVal v_lngErrorNumber As Long, _
            ByVal v_strErrorType As String, _
            ByVal v_strErrorDesc As String _
            ) As Boolean
            
    Dim l_bolError As Boolean
    Dim lngKeyTemplateNum As Long
    Dim l_strKeyTime As String
    
    On Error GoTo eh
    
    If Not IsNull(v_rsTemplateBase!KEY_TEMPLATE_NUM) And _
        Not IsEmpty(v_rsTemplateBase!KEY_TEMPLATE_NUM) Then
        lngKeyTemplateNum = v_rsTemplateBase!KEY_TEMPLATE_NUM
    Else
        If Not IsNull(v_rsTemplateBase!TEMPLATE_NUM) And _
            Not IsEmpty(v_rsTemplateBase!TEMPLATE_NUM) Then
            lngKeyTemplateNum = v_rsTemplateBase!TEMPLATE_NUM
        Else
            lngKeyTemplateNum = -1
        End If
    End If

    If IsNull(v_rsTemplateBase!TIME_STAMP) Then
        l_strKeyTime = ""
    Else
        l_strKeyTime = v_rsTemplateBase!TIME_STAMP
    End If
    
    l_bolError = AppendBusinessError(m_rsErrors, _
            v_lngErrorNumber, _
            v_strErrorType, _
            v_strErrorDesc, _
            m_strCMODULENAME, _
            111, , , "TEMPLATE_NUM", _
            CStr(lngKeyTemplateNum), _
            "", _
            "0", _
            "0", _
            "", _
            "")

    AppendTemplateBaseBRV = l_bolError
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "AppendTemplateBaseBRV"

End Function


Private Sub SetNewTemplateNumber( _
                ByVal v_rsTemplateBase As ADODB.Recordset, _
                ByVal v_rsSortOrder As ADODB.Recordset, _
                ByVal v_rsForceOrder As ADODB.Recordset, _
                ByVal v_rsGroupOrder As ADODB.Recordset, _
                ByVal v_rsOTResponses As ADODB.Recordset, _
                ByVal v_rsIneligibleRules As ADODB.Recordset, _
                ByVal v_rsShifts As ADODB.Recordset)
                

    On Error GoTo eh
                
    With v_rsTemplateBase
        If RecordsExist(v_rsTemplateBase) Then
            .MoveFirst
            Do Until .EOF
                !TEMPLATE_NUM = m_lngTemplateNum
                !KEY_TEMPLATE_NUM = m_lngTemplateNum
                .MoveNext
            Loop
        End If
    End With
   
    With v_rsSortOrder
        If RecordsExist(v_rsSortOrder) Then
            .MoveFirst
            Do Until .EOF
                !TEMPLATE_NUM = m_lngTemplateNum
                !KEY_TEMPLATE_NUM = m_lngTemplateNum
                .MoveNext
            Loop
        End If
    End With
    
    With v_rsForceOrder
        If RecordsExist(v_rsForceOrder) Then
            .MoveFirst
            Do Until .EOF
                !TEMPLATE_NUM = m_lngTemplateNum
                !KEY_TEMPLATE_NUM = m_lngTemplateNum
                .MoveNext
            Loop
        End If
    End With
    
    With v_rsGroupOrder
        If RecordsExist(v_rsGroupOrder) Then
            .MoveFirst
            Do Until .EOF
                !TEMPLATE_NUM = m_lngTemplateNum
                !KEY_TEMPLATE_NUM = m_lngTemplateNum
                .MoveNext
            Loop
        End If
    End With
    
    With v_rsOTResponses
        If RecordsExist(v_rsOTResponses) Then
            .MoveFirst
            Do Until .EOF
                !TEMPLATE_NUM = m_lngTemplateNum
                !KEY_TEMPLATE_NUM = m_lngTemplateNum
                .MoveNext
            Loop
        End If
    End With
    
    With v_rsIneligibleRules
        If RecordsExist(v_rsIneligibleRules) Then
            .MoveFirst
            Do Until .EOF
                !TEMPLATE_NUM = m_lngTemplateNum
                !KEY_TEMPLATE_NUM = m_lngTemplateNum
                .MoveNext
            Loop
        End If
    End With
    
    With v_rsShifts
        If RecordsExist(v_rsShifts) Then
            .MoveFirst
            Do Until .EOF
                !TEMPLATE_NUM = m_lngTemplateNum
                !KEY_TEMPLATE_NUM = m_lngTemplateNum
                .MoveNext
            Loop
        End If
    End With
    
    Exit Sub
    
eh:
    
  TrsRaiseError m_strCMODULENAME, "SetNewTemplateNumber"
  
End Sub


'Private Function PreValidateTemplateBase( _
'                    ByRef r_rsTemplateBase As ADODB.Recordset, _
'                    ByRef r_rsErrors As ADODB.Recordset _
'                ) As Boolean
'
'
'On Error GoTo eh
'
'    If Not RecordsExist(r_rsTemplateBase) Then
'        PreValidateTemplateBase = True
'        Exit Function
'    End If
'
'    r_rsTemplateBase.MoveFirst
'    Do Until r_rsTemplateBase.EOF
'        Select Case r_rsTemplateBase!RECORD_STATUS
'            Case "I"
'                If Not ValidateNewTemplate(r_rsTemplateBase, r_rsErrors) Then
'                    l_blnValidationError = True
'                End If
'            Case "C"
'                If Not ValidateChangeTemplate(r_rsTemplateBase, r_rsErrors) Then
'                    l_blnValidationError = True
'                End If
'            Case "D"
'                If Not ValidateDeleteTemplate(r_rsTemplateBase, r_rsErrors) Then
'                    l_blnValidationError = True
'                End If
'            Case Else
'                AppendEntryError r_rsErrors, r_rsTemplateBase, trErrbusValidationError, mstrCBUSINESSERR, "Record status must be I, C, or D"
'                l_blnValidationError = True
'        End Select
'        r_rsTemplateBase.MoveNext
'    Loop
'
'
'    PreValidateTemplateBase = Not l_blnValidationError
'    r_rsTemplateBase.MoveFirst
'    Exit Function
'
'eh:
'    TrsRaiseError m_strCMODULENAME, "PreValidateTemplateBase"
'
'End Function
'

Private Function ValidateNewTemplate( _
                    ByRef r_rsTemplateBase As ADODB.Recordset, _
                    ByRef r_rsErrors As ADODB.Recordset _
                ) As Boolean
                
                
    ValidateNewTemplate = True
'    ValidateNewTemplate = Not l_blnValidationError
    r_rsTemplateBase.MoveFirst
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "ValidateNewTemplate"
                
End Function


Private Function ValidateChangeTemplate( _
                    ByRef r_rsTemplateBase As ADODB.Recordset, _
                    ByRef r_rsErrors As ADODB.Recordset _
                ) As Boolean
                
                
    ValidateChangeTemplate = True
'    ValidateChangeTemplate = Not l_blnValidationError
    r_rsTemplateBase.MoveFirst
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "ValidateChangeTemplate"
                
End Function


Private Function ValidateDeleteTemplate( _
                    ByRef r_rsTemplateBase As ADODB.Recordset, _
                    ByRef r_rsErrors As ADODB.Recordset _
                ) As Boolean
                
                
    ValidateDeleteTemplate = True
'    ValidateDeleteTemplate = Not l_blnValidationError
    r_rsTemplateBase.MoveFirst
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "ValidateDeleteTemplate"
                
End Function


Private Function PreValidateSortOrders( _
                    ByRef r_rsSortOrder As ADODB.Recordset, _
                    ByRef r_rsForceOrder As ADODB.Recordset, _
                    ByRef r_rsGroupOrder As ADODB.Recordset, _
                    ByRef r_rsErrors As ADODB.Recordset _
                ) As Boolean


On Error GoTo eh

    If Not RecordsExist(r_rsSortOrder) Then
        PreValidateSortOrders = True
        Exit Function
    End If

    
    PreValidateSortOrders = True
'    PreValidateSortOrders = Not l_blnValidationError
    r_rsSortOrder.MoveFirst
    r_rsForceOrder.MoveFirst
    r_rsGroupOrder.MoveFirst
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "PreValidateSortOrders"

End Function


Private Function PreValidateResponses( _
                    ByRef r_rsOTResponses As ADODB.Recordset) As Boolean

    Dim l_strRecStatus As String
    Dim l_strErrorMessage As String

On Error GoTo eh

    PreValidateResponses = True
    If Not RecordsExist(r_rsOTResponses) Then
        Exit Function
    End If
    
    If IsNull(r_rsOTResponses!RECORD_STATUS) Or IsEmpty(r_rsOTResponses!RECORD_STATUS) Then
        l_strErrorMessage = "Record Status is missing"
        AppendBusinessError m_rsErrors, _
            trErrbusOTCanvass, _
            mstrCAPPLICATIONERR, _
            l_strErrorMessage, _
            m_strCMODULENAME, _
            TrType_OT_TEMPL, , , "TEMPLATE_NUM", _
            "0"
        PreValidateResponses = False
        Exit Function
    End If
    
    l_strRecStatus = UCase(Trim(r_rsOTResponses!RECORD_STATUS))
    If l_strRecStatus = "C" Or l_strRecStatus = "D" Then
        If IsNull(r_rsOTResponses!KEY_OT_RESP_CD) Or IsEmpty(r_rsOTResponses!KEY_OT_RESP_CD) Then
            l_strErrorMessage = "KEY_OT_RESP_CD is missing"
            AppendBusinessError m_rsErrors, _
                trErrbusOTCanvass, _
                mstrCAPPLICATIONERR, _
                l_strErrorMessage, _
                m_strCMODULENAME, _
                TrType_OT_TEMPL, , , "TEMPLATE_NUM", _
                "0"
            PreValidateResponses = False
            Exit Function
        End If
    ElseIf l_strRecStatus = "I" Then
        If IsNull(r_rsOTResponses!OT_RESP_CD) Or IsEmpty(r_rsOTResponses!OT_RESP_CD) Then
            l_strErrorMessage = "OT_RESP_CD is missing"
            AppendBusinessError m_rsErrors, _
                trErrbusOTCanvass, _
                mstrCAPPLICATIONERR, _
                l_strErrorMessage, _
                m_strCMODULENAME, _
                TrType_OT_TEMPL, , , "TEMPLATE_NUM", _
                "0"
            PreValidateResponses = False
            Exit Function
        End If
    Else
        l_strErrorMessage = "RECORD_STATUS is invalid. Input=" & l_strRecStatus
        AppendBusinessError m_rsErrors, _
                trErrbusOTCanvass, _
                mstrCAPPLICATIONERR, _
                l_strErrorMessage, _
                m_strCMODULENAME, _
                TrType_OT_TEMPL, , , "TEMPLATE_NUM", _
                "0"
            PreValidateResponses = False
            Exit Function
    End If
    
'    PreValidateSortOrders = Not l_blnValidationError
'    r_rsOTResponses.MoveFirst
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "PreValidateResponses"

End Function


Private Function PreValidateRules( _
                    ByRef r_rsIneligibleRules As ADODB.Recordset, _
                    ByRef r_rsErrors As ADODB.Recordset _
                ) As Boolean

On Error GoTo eh

    If Not RecordsExist(r_rsIneligibleRules) Then
        PreValidateRules = True
        Exit Function
    End If

    
    PreValidateRules = True
'    PreValidateRules = Not l_blnValidationError
    r_rsIneligibleRules.MoveFirst
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "PreValidateRules"

End Function


Private Function PreValidateShifts( _
                    ByRef r_rsShifts As ADODB.Recordset, _
                    ByRef r_rsErrors As ADODB.Recordset _
                ) As Boolean

On Error GoTo eh

    If Not RecordsExist(r_rsShifts) Then
        PreValidateShifts = True
        Exit Function
    End If

    r_rsShifts.MoveFirst
    Do Until r_rsShifts.EOF
        If r_rsShifts!RECORD_STATUS <> "D" Then
            If Trim(r_rsShifts!END_TIME) = "24:00" Then
                r_rsShifts!END_TIME = "00:00"
            End If
        End If
        r_rsShifts.MoveNext
    Loop
    
    PreValidateShifts = True
'    PreValidateShifts = Not l_blnValidationError
    r_rsShifts.MoveFirst
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "PreValidateShifts"

End Function


Private Sub GetOvertimeTemplateData( _
                ByVal v_lngTemplateNum As Long, _
                ByVal v_rsTemplateBase As ADODB.Recordset, _
                ByVal v_rsSortOrder As ADODB.Recordset, _
                ByVal v_rsForceOrder As ADODB.Recordset, _
                ByVal v_rsGroupOrder As ADODB.Recordset, _
                ByVal v_rsOTResponses As ADODB.Recordset, _
                ByVal v_rsIneligibleRules As ADODB.Recordset, _
                ByVal v_rsShifts As ADODB.Recordset)
    
    Dim l_obj As dbOTTemplate.cdOTTemplate
   
    
    On Error GoTo eh
    
    Set l_obj = CreateObject("dbOTTemplate.cdOTTemplate")
    If RecordsExist(v_rsTemplateBase) Then
        Set m_rsTemplateBase = l_obj.GetOvertimeTemplateBase(v_lngTemplateNum)
    End If
    If RecordsExist(v_rsSortOrder) Then
        Set m_rsSortOrder = l_obj.GetTemplateOrder(v_lngTemplateNum, CANVASS_ORDER)
    End If
    If RecordsExist(v_rsForceOrder) Then
        Set m_rsForceOrder = l_obj.GetTemplateOrder(v_lngTemplateNum, FORCE_ORDER)
    End If
    If RecordsExist(v_rsGroupOrder) Then
        Set m_rsGroupOrder = l_obj.GetTemplateOrder(v_lngTemplateNum, GROUP_ORDER)
    End If
    If RecordsExist(v_rsOTResponses) Then
        Set m_rsOTResponses = l_obj.GetTemplateOTResponses(v_lngTemplateNum)
    End If
    If RecordsExist(v_rsIneligibleRules) Then
        Set m_rsIneligibleRules = l_obj.GetTemplateIneligibleRules(v_lngTemplateNum)
    End If
    If RecordsExist(v_rsShifts) Then
        Set m_rsShifts = l_obj.GetTemplateShifts(v_lngTemplateNum)
    End If
    
    Set l_obj = Nothing
    Exit Sub

eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "GetOvertimeTemplateData"
    
End Sub



'Private Sub GetOvertimeTemplateData( _
'                ByVal v_lngTemplateNum As Long, _
'                ByRef r_rsTemplateBase As ADODB.Recordset, _
'                ByRef r_rsSortOrder As ADODB.Recordset, _
'                ByRef r_rsForceSortOrder As ADODB.Recordset, _
'                ByRef r_rsGroupSortOrder As ADODB.Recordset, _
'                ByRef r_rsOTResponses As ADODB.Recordset, _
'                ByRef r_rsIneligibleRules As ADODB.Recordset, _
'                ByRef r_rsShifts As ADODB.Recordset)
'
'    Dim l_obj As dbOTTemplate.cdOTTemplate
'
'
'    On Error GoTo eh
'
'    Set l_obj = CreateObject("dbOTTemplate.cdOTTemplate")
'    Set r_rsTemplateBase = l_obj.GetOvertimeTemplateBase(v_lngTemplateNum)
'    Set r_rsSortOrder = l_obj.GetTemplateSort(v_lngTemplateNum, "C")
'    Set r_rsForceSortOrder = l_obj.GetTemplateSort(v_lngTemplateNum, "F")
'    Set r_rsGroupSortOrder = l_obj.GetTemplateSort(v_lngTemplateNum, "G")
'    Set r_rsOTResponses = l_obj.GetTemplateOTResponses(v_lngTemplateNum)
'    Set r_rsIneligibleRules = l_obj.GetTemplateIneligibleRules(v_lngTemplateNum)
'    Set r_rsShifts = l_obj.GetTemplateShifts(v_lngTemplateNum)
'
'    Set l_obj = Nothing
'    Exit Sub
'
'eh:
'    Set l_obj = Nothing
'    TrsRaiseError m_strCMODULENAME, "GetOvertimeTemplateData"
'
'End Sub

'
'Private Function DetermineCriteria( _
'    ByVal v_rsTemplateBase As ADODB.Recordset, _
'    ByVal v_rs As ADODB.Recordset, _
'    ByVal v_strCriteriaType As String _
'    ) As Variant
'
'    Dim l_col As Collection
'
'    On Error GoTo eh
'
'    Set l_col = New Collection
'    If v_strCriteriaType = "L" Then
'        If v_rsTemplateBase!LOC_CD_SELECT = "R" Then
'            l_col.Add v_rsTemplateBase!TEMPL_LOCATION
'        End If
'    End If
'
'    If RecordsExist(v_rs) Then
'        v_rs.MoveFirst
'        While Not v_rs.EOF
'            Select Case v_strCriteriaType
'                Case "L"
'                    l_col.Add v_rs("LOC_CD").Value
'                Case "J"
'                    l_col.Add v_rs("JOB_NUM").Value
'                Case "C"
'                    l_col.Add v_rs("CRAFT_NUM").Value
'                Case "W"
'                    l_col.Add v_rs("WG_NUM").Value
'            End Select
'            v_rs.MoveNext
'        Wend
'    End If
'
'    DetermineCriteria = ArrayFromCollection(l_col)
'
'   Exit Function
'
'eh:
'    Set l_obj = Nothing
'    TrsRaiseError m_strCMODULENAME, "DetermineCriteria"
'
'End Function



Private Function GetNextTemplateNum() As Long

    Dim l_obj As utilDBDataServices.cuDBDataServices

On Error GoTo eh
    Set l_obj = CreateObject("utilDBDataServices.cuDBDataServices")
    GetNextTemplateNum = l_obj.GetNextSequenceNumber("TP")
    Set l_obj = Nothing
    Exit Function
    
eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "GetNextTemplateNum"

End Function


Private Function ProcessOvertimeTemplate( _
                    ByVal v_lngTemplateNum As Long, _
                    ByVal v_rsTemplateBase As ADODB.Recordset, _
                    ByVal v_rsSortOrder As ADODB.Recordset, _
                    ByVal v_rsForceOrder As ADODB.Recordset, _
                    ByVal v_rsGroupOrder As ADODB.Recordset, _
                    ByVal v_rsOTResponses As ADODB.Recordset, _
                    ByVal v_rsIneligibleRules As ADODB.Recordset, _
                    ByVal v_rsShifts As ADODB.Recordset _
                ) As Boolean

    Dim l_obj As dbOTTemplate.cdOTTemplate
    Dim blnReturn As Boolean
    
On Error GoTo eh
    
    
    Set l_obj = CreateObject("dbOTTemplate.cdOTTemplate")
    
    blnReturn = l_obj.SaveOvertimeTemplate( _
                    v_lngTemplateNum, CStr(m_lngUserNum), _
                    m_rsTemplateBase, m_rsSortOrder, m_rsForceOrder, m_rsGroupOrder, _
                    m_rsOTResponses, m_rsIneligibleRules, m_rsShifts)

    ProcessOvertimeTemplate = blnReturn

'    If ProcessTemplateBaseData( _
'        v_rsTemplateBase, v_rsWorkLimitRules, v_rsShifts) Then
'        'continue
'    Else
'        SaveOvertimeTemplate = False
'        l_objCtx.SetAbort
'        GoTo xit
'    End If
'
'    If ProcessTemplateSorts(v_rsSortOrder, v_rsForceOrder, v_rsGroupOrder) Then
'        'continue
'    Else
'        SaveOvertimeTemplate = False
'        l_objCtx.SetAbort
'        GoTo xit
'    End If
'
'    If ProcessTemplateResponses(v_rsOTResponses) Then
'        'continue
'    Else
'        SaveOvertimeTemplate = False
'        l_objCtx.SetAbort
'        GoTo xit
'    End If
'
'    If ProcessIneligibleRules(v_rsIneligibleRules) Then
'        'continue
'    Else
'        SaveOvertimeTemplate = False
'        l_objCtx.SetAbort
'        GoTo xit
'    End If
'
'    If ProcessShifts(v_rsShifts) Then
'        'continue
'    Else
'        SaveOvertimeTemplate = False
'        l_objCtx.SetAbort
'        GoTo xit
'    End If



    Set l_obj = Nothing
    Exit Function
    
eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "ProcessOvertimeTemplate"
    
End Function

'
'Private Function ProcessTemplateSorts( _
'    ByVal v_rsSortOrder As ADODB.Recordset, _
'    ByVal v_rsForceSortOrder As ADODB.Recordset, _
'    ByVal v_rsGroupSortOrder As ADODB.Recordset _
'    ) As Boolean
'
'
'
'On Error GoTo eh
'
'
'
'
'    Exit Function
'
'eh:
'
'    TrsRaiseError m_strCMODULENAME, "ProcessTemplateSorts"
'
'End Function

Private Function FormatOTTemplateBRV( _
                        ByVal v_lngBRVType As TrTypeBRV, _
                        ByVal v_lngTemplateNum As Long) As String
    
    FormatOTTemplateBRV = CStr(v_lngBRVType) & ";" & CStr(v_lngTemplateNum)
    
    Exit Function
    
End Function


Private Function ARequestExist( _
                    ByVal v_rs1 As ADODB.Recordset, _
                    Optional ByVal v_rs2 As ADODB.Recordset, _
                    Optional ByVal v_rs3 As ADODB.Recordset, _
                    Optional ByVal v_rs4 As ADODB.Recordset, _
                    Optional ByVal v_rs5 As ADODB.Recordset, _
                    Optional ByVal v_rs6 As ADODB.Recordset, _
                    Optional ByVal v_rs7 As ADODB.Recordset, _
                    Optional ByVal v_rs8 As ADODB.Recordset, _
                    Optional ByVal v_rs9 As ADODB.Recordset, _
                    Optional ByVal v_rs10 As ADODB.Recordset, _
                    Optional ByVal v_rs11 As ADODB.Recordset _
            ) As Boolean

    On Error GoTo eh

    ARequestExist = RecordsExist(v_rs1) _
        Or RecordsExist(v_rs2) _
        Or RecordsExist(v_rs3) _
        Or RecordsExist(v_rs4) _
        Or RecordsExist(v_rs5) _
        Or RecordsExist(v_rs6) _
        Or RecordsExist(v_rs7) _
        Or RecordsExist(v_rs8) _
        Or RecordsExist(v_rs9) _
        Or RecordsExist(v_rs10) _
        Or RecordsExist(v_rs11)

    Exit Function
    
eh:
    
  TrsRaiseError m_strCMODULENAME, "ARequestExist"
  
End Function


Private Sub FreeGlobalRS()

   Set m_rsErrors = Nothing

End Sub

