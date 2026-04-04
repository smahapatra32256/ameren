VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "cbsOTTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const m_strCMODULENAME As String = "cbsOTTemplate"

'Private m_rsTemplateData As ADODB.Recordset
'Private m_strSortColumns(1 To 24) As String
'Private m_strGroupSortColumns(1 To 4) As String



Public Function GetOvertimeTemplates( _
                    v_intLocationCode As Integer _
                ) As ADODB.Recordset
    
    Dim l_obj As dbOTTemplate.cdOTTemplate
    
    
    On Error GoTo eh
    
'    Dim x As String
'    x = GetObjectContext.Security.GetOriginalCreatorName
'    App.LogEvent ("cbsOTTemplate user: " & x)
    
    Set l_obj = CreateObject("dbOTTemplate.cdOTTemplate")
    
    Set GetOvertimeTemplates = l_obj.GetOvertimeTemplates( _
            v_intLocationCode)
    
    Set l_obj = Nothing
    
    Exit Function

eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "GetOvertimeTemplates"
    
End Function


Public Function GetOvertimeTemplateData( _
                    ByVal v_lngTemplateNum As Long, _
                    ByRef r_rsTemplateBase As ADODB.Recordset, _
                    ByRef r_rsSortOrder As ADODB.Recordset, _
                    ByRef r_rsForceSortOrder As ADODB.Recordset, _
                    ByRef r_rsGroupSortOrder As ADODB.Recordset, _
                    ByRef r_rsOTResponses As ADODB.Recordset, _
                    ByRef r_rsIneligibleRules As ADODB.Recordset, _
                    ByRef r_rsShifts As ADODB.Recordset _
                ) As Boolean
    
    
' RSG - Do we still need ExtractTemplateBaseData?
    
    Dim l_obj As dbOTTemplate.cdOTTemplate
   
    
    On Error GoTo eh
    
    Set l_obj = CreateObject("dbOTTemplate.cdOTTemplate")
    
    Set r_rsTemplateBase = l_obj.GetOvertimeTemplateBase( _
        v_lngTemplateNum)
    If RecordsExist(r_rsTemplateBase) Then
'        Set r_rsTemplateBase = ExtractTemplateBaseData()
        
        Set r_rsSortOrder = GetTemplateSortOrder(v_lngTemplateNum)
        Set r_rsForceSortOrder = GetTemplateForceSortOrder(v_lngTemplateNum)
        Set r_rsGroupSortOrder = GetTemplateGroupSortOrder(v_lngTemplateNum)
        Set r_rsOTResponses = GetTemplateOTResponses(v_lngTemplateNum)
        Set r_rsIneligibleRules = GetTemplateIneligibleRules(v_lngTemplateNum)
        Set r_rsShifts = GetTemplateShifts(v_lngTemplateNum)
    End If
    
    Set l_obj = Nothing
    GetOvertimeTemplateData = True
    Exit Function

eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "GetOvertimeTemplateData"
    
End Function


Public Function GetTemplateSupportInfo( _
                    ByRef r_rsOTResponses As ADODB.Recordset, _
                    ByRef r_rsRefusalChargeRules As ADODB.Recordset, _
                    ByRef r_rsIneligibleRules As ADODB.Recordset, _
                    ByRef r_rsAvailSortColumns As ADODB.Recordset, _
                    ByRef r_rsAvailGroupSortColumns As ADODB.Recordset _
                ) As Boolean
    
    Dim l_obj As dbOTTemplate.cdOTTemplate
    
    On Error GoTo eh
    
    Set r_rsOTResponses = GetOvertimeResponses()
    Set r_rsRefusalChargeRules = GetRefusalChargeRules()
    Set r_rsIneligibleRules = GetIneligibleRules()
    Set r_rsAvailSortColumns = GetAvailableSortColumns()
    Set r_rsAvailGroupSortColumns = GetAvailableGroupColumns()
    
    Set l_obj = Nothing
    GetTemplateSupportInfo = True
    Exit Function

eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "GetTemplateSupportInfo"
        
End Function


Public Function GetShiftsForTemplate( _
                    ByVal v_lngTemplateNum As Long _
                ) As ADODB.Recordset
       
On Error GoTo eh
    
    Set GetShiftsForTemplate = GetTemplateShifts(v_lngTemplateNum)
    
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "GetShiftsForTemplate"
    
End Function


Private Function GetOvertimeResponses() As ADODB.Recordset

    Dim l_obj As dbOTTemplate.cdOTTemplate

On Error GoTo eh

    Set l_obj = CreateObject("dbOTTemplate.cdOTTemplate")
    Set GetOvertimeResponses = l_obj.GetOTResponses()
    Set l_obj = Nothing

    Exit Function
    
eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "GetOvertimeResponses"

End Function


Private Function GetRefusalChargeRules() As ADODB.Recordset

    Dim l_obj As dbOTTemplate.cdOTTemplate

On Error GoTo eh

    Set l_obj = CreateObject("dbOTTemplate.cdOTTemplate")
    Set GetRefusalChargeRules = l_obj.GetRefusalChargeRules()
    Set l_obj = Nothing

    Exit Function
    
eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "GetRefusalChargeRules"

End Function


Private Function GetIneligibleRules() As ADODB.Recordset

    Dim l_obj As dbOTTemplate.cdOTTemplate

On Error GoTo eh

    Set l_obj = CreateObject("dbOTTemplate.cdOTTemplate")
    Set GetIneligibleRules = l_obj.GetIneligibleRules()
    Set l_obj = Nothing

    Exit Function
    
eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "GetIneligibleRules"

End Function


Private Function CreateTemplateBaseRecordset() As ADODB.Recordset

    Dim l_rs As ADODB.Recordset
    
    On Error GoTo eh
    
    Set l_rs = New ADODB.Recordset
    With l_rs
      .Fields.Append "TEMPLATE_NUM", adInteger, , adFldUpdatable
      .Fields.Append "CANVASS_DESC", adChar, 30, adFldUpdatable
      .Fields.Append "TEMPL_LOCATION", adInteger, , adFldUpdatable
      '.Fields.Append "REFUSE_RULE_NUM", adChar, 1, adFldUpdatable
      .Fields.Append "HRS_OVERRIDE_IND", adChar, 1, adFldUpdatable
      .Fields.Append "TEMPL_STATUS", adChar, 1, adFldUpdatable
      .Fields.Append "LAST_MOD_EMP_NUM", adInteger, , adFldUpdatable
      .Fields.Append "LAST_MOD_TIMESTAMP", adChar, 26, adFldUpdatable

      .CursorLocation = adUseClient
      .CursorType = adOpenDynamic
      .LockType = adLockOptimistic
      .Open
    End With

    
    Set CreateTemplateBaseRecordset = l_rs
    Exit Function
    
eh:
    Set l_rs = Nothing
    TrsRaiseError m_strCMODULENAME, "CreateTemplateBaseRecordset"

End Function


'Private Function ExtractTemplateBaseData() _
'    As ADODB.Recordset
'
'    Dim l_rs As ADODB.Recordset
'    Dim l_strField As String
'    Dim l_intI As Integer
'
'    On Error GoTo eh
'
'    Set l_rs = CreateTemplateBaseRecordset
'
'
'    l_rs.AddNew
'    For l_intI = 1 To l_rs.Fields.Count
'        l_strField = l_rs.Fields(l_intI - 1).Name
'        l_rs(l_strField) = m_rsTemplateData(l_strField)
'    Next
'
'    Set ExtractTemplateBaseData = l_rs
'
'    Exit Function
'
'eh:
'    Set l_rs = Nothing
'    TrsRaiseError m_strCMODULENAME, "ExtractTemplateBaseData"
'
'End Function
'
'
'Private Function BuildTemplateSortOrder( _
'    ByVal v_lngTemplateNum As Long, _
'    ByVal v_strSort As String, _
'    Optional v_blnGroupSort As Boolean = False _
'    ) As ADODB.Recordset
'
'    Dim l_rs As ADODB.Recordset
'    Dim l_strSortOrder As String
'    Dim l_intI As Integer
'    Dim l_intNumCols As Integer
'    Dim l_intSortCode As Integer
'    Dim l_lngLength As Long
'    Dim l_intStartPos As Integer
'
'On Error GoTo eh
'
'    l_strSortOrder = Trim(v_strSort)
'    Set l_rs = CreateTemplateSortRecordset()
'    If l_strSortOrder = "" Then
'        l_intNumCols = 0
'    Else
'        l_lngLength = Len(l_strSortOrder)
'        l_intNumCols = l_lngLength / 2
'    End If
'
'    For l_intI = 1 To l_intNumCols
'        l_intStartPos = (l_intI - 1) * 2 + 1
'        l_intSortCode = Mid(l_strSortOrder, l_intStartPos, 2)
'        l_rs.AddNew
'        l_rs!TEMPLATE_NUM = v_lngTemplateNum
'        l_rs!SORT_CD = l_intSortCode
'        l_rs!SORT_ORDER = l_intI
'        If v_blnGroupSort Then
'            l_rs!SORT_NAME = m_strGroupSortColumns(l_intSortCode)
'        Else
'            l_rs!SORT_NAME = m_strSortColumns(l_intSortCode)
'        End If
'        l_rs!RECORD_STATUS = ""
'        l_rs!LAST_MOD_EMP_NUM = m_rsTemplateData!LAST_MOD_EMP_NUM
'        l_rs!LAST_MOD_TIMESTAMP = m_rsTemplateData!LAST_MOD_TIMESTAMP
'    Next l_intI
'
'    Set BuildTemplateSortOrder = l_rs
'
'    Exit Function
'
'eh:
'    Set l_rs = Nothing
'    TrsRaiseError m_strCMODULENAME, "BuildTemplateSortOrder"
'
'End Function
'
'
'Private Function CreateTemplateSortRecordset() As ADODB.Recordset
'
'    Dim l_rs As ADODB.Recordset
'
'    On Error GoTo eh
'
'    Set l_rs = New ADODB.Recordset
'    With l_rs
'      .Fields.Append "TEMPLATE_NUM", adInteger, , adFldUpdatable
'      .Fields.Append "SORT_CD", adInteger, , adFldUpdatable
'      .Fields.Append "SORT_NAME", adChar, 30, adFldUpdatable
'      .Fields.Append "SORT_ORDER", adInteger, , adFldUpdatable
'      .Fields.Append "RECORD_STATUS", adChar, 1, adFldUpdatable
'      .Fields.Append "LAST_MOD_EMP_NUM", adInteger, , adFldUpdatable
'      .Fields.Append "LAST_MOD_TIMESTAMP", adChar, 26, adFldUpdatable
'
'      .CursorLocation = adUseClient
'      .CursorType = adOpenDynamic
'      .LockType = adLockOptimistic
'      .Open
'    End With
'
'
'    Set CreateTemplateSortRecordset = l_rs
'    Exit Function
'
'eh:
'    Set l_rs = Nothing
'    TrsRaiseError m_strCMODULENAME, "CreateTemplateSortRecordset"
'
'End Function


Private Function GetTemplateSortOrder( _
    ByVal v_lngTemplateNum As Long _
    ) As ADODB.Recordset

    Dim l_obj As dbOTTemplate.cdOTTemplate

On Error GoTo eh


    Set l_obj = CreateObject("dbOTTemplate.cdOTTemplate")
    Set GetTemplateSortOrder = l_obj.GetTemplateSort(v_lngTemplateNum, "C")
    Set l_obj = Nothing

    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "GetTemplateSortOrder"

End Function


Private Function GetTemplateForceSortOrder( _
    ByVal v_lngTemplateNum As Long _
    ) As ADODB.Recordset

    Dim l_obj As dbOTTemplate.cdOTTemplate

On Error GoTo eh


    Set l_obj = CreateObject("dbOTTemplate.cdOTTemplate")
    Set GetTemplateForceSortOrder = l_obj.GetTemplateSort(v_lngTemplateNum, "F")
    Set l_obj = Nothing

    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "GetTemplateForceSortOrder"

End Function


Private Function GetTemplateGroupSortOrder( _
    ByVal v_lngTemplateNum As Long _
    ) As ADODB.Recordset

    Dim l_obj As dbOTTemplate.cdOTTemplate

On Error GoTo eh

    Set l_obj = CreateObject("dbOTTemplate.cdOTTemplate")
    Set GetTemplateGroupSortOrder = l_obj.GetTemplateSort(v_lngTemplateNum, "G")
    Set l_obj = Nothing

    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "GetTemplateGroupSortOrder"

End Function


Private Function GetTemplateOTResponses( _
    ByVal v_lngTemplateNum As Long _
    ) As ADODB.Recordset

' Remove RESPONSE_IND from TTR_CANV_TEMPL

    Dim l_obj As dbOTTemplate.cdOTTemplate

On Error GoTo eh

    Set l_obj = CreateObject("dbOTTemplate.cdOTTemplate")
    Set GetTemplateOTResponses = l_obj.GetTemplateOTResponses(v_lngTemplateNum)
    Set l_obj = Nothing

    Exit Function
    
eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "GetTemplateOTResponses"

End Function


Private Function GetTemplateShifts( _
                    ByVal v_lngTemplateNum As Long _
                ) As ADODB.Recordset

    Dim l_obj As dbOTTemplate.cdOTTemplate

On Error GoTo eh

    Set l_obj = CreateObject("dbOTTemplate.cdOTTemplate")
    Set GetTemplateShifts = l_obj.GetTemplateShifts(v_lngTemplateNum)
    Set l_obj = Nothing

    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "GetTemplateShifts"

End Function


Private Function GetAvailableSortColumns() As ADODB.Recordset

    Dim l_obj As dbOTTemplate.cdOTTemplate

On Error GoTo eh

    Set l_obj = CreateObject("dbOTTemplate.cdOTTemplate")
    Set GetAvailableSortColumns = l_obj.GetAvailableSortColumns()
    Set l_obj = Nothing

    Exit Function
    
eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "GetAvailableSortColumns"

End Function


Private Function GetAvailableGroupColumns() As ADODB.Recordset

    Dim l_obj As dbOTTemplate.cdOTTemplate

On Error GoTo eh

    Set l_obj = CreateObject("dbOTTemplate.cdOTTemplate")
    Set GetAvailableGroupColumns = l_obj.GetAvailableGroupColumns()
    Set l_obj = Nothing

    Exit Function
    
eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "GetAvailableGroupColumns"

End Function


Private Function GetTemplateIneligibleRules( _
                    ByVal v_lngTemplateNum As Long _
                ) As ADODB.Recordset

    
    Dim l_obj As dbOTTemplate.cdOTTemplate

On Error GoTo eh

    Set l_obj = CreateObject("dbOTTemplate.cdOTTemplate")
    Set GetTemplateIneligibleRules = l_obj.GetTemplateIneligibleRules(v_lngTemplateNum)
    Set l_obj = Nothing
    
    Exit Function
    
eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "GetTemplateIneligibleRules"

End Function


'Private Function GetTemplateJobSelections( _
'    ByVal v_lngTemplateNum As Long _
'    ) As ADODB.Recordset
'
'    Dim l_strDataStore As String
'    Dim l_obj As dbOTTemplate.cdOTTemplate
'
'On Error GoTo eh
'
'    l_strDataStore = Trim(m_rsTemplateData!JOB_NUM_SELECT)
'    If l_strDataStore = "T" Then 'In Template
'        Set l_obj = CreateObject("dbOTTemplate.cdOTTemplate")
'        Set GetTemplateJobSelections = l_obj.GetTemplateJobSelections(v_lngTemplateNum)
'        Set l_obj = Nothing
'    Else
'        Set GetTemplateJobSelections = Nothing
'    End If
'
'    Exit Function
'
'eh:
'    Set l_obj = Nothing
'    TrsRaiseError m_strCMODULENAME, "GetTemplateJobSelections"
'
'End Function
'
'
'Private Function GetTemplateCraftSelections( _
'    ByVal v_lngTemplateNum As Long _
'    ) As ADODB.Recordset
'
'    Dim l_strDataStore As String
'    Dim l_obj As dbOTTemplate.cdOTTemplate
'
'On Error GoTo eh
'
'    l_strDataStore = Trim(m_rsTemplateData!JOB_NUM_SELECT)
'    If l_strDataStore = "T" Then 'In Template
'        Set l_obj = CreateObject("dbOTTemplate.cdOTTemplate")
'        Set GetTemplateCraftSelections = l_obj.GetTemplateCraftSelections(v_lngTemplateNum)
'        Set l_obj = Nothing
'    Else
'        Set GetTemplateCraftSelections = Nothing
'    End If
'
'    Exit Function
'
'eh:
'    Set l_obj = Nothing
'    TrsRaiseError m_strCMODULENAME, "GetTemplateCraftSelections"
'
'End Function
'
'
'Private Function GetTemplateLocationSelections( _
'    ByVal v_lngTemplateNum As Long _
'    ) As ADODB.Recordset
'
'    Dim l_strDataStore As String
'    Dim l_obj As dbOTTemplate.cdOTTemplate
'
'On Error GoTo eh
'
'    l_strDataStore = Trim(m_rsTemplateData!LOC_CD_SELECT)
'    If l_strDataStore = "T" Then 'In Template
'        Set l_obj = CreateObject("dbOTTemplate.cdOTTemplate")
'        Set GetTemplateLocationSelections = l_obj.GetTemplateLocationSelections(v_lngTemplateNum)
'        Set l_obj = Nothing
'    Else
'        Set GetTemplateLocationSelections = Nothing
'    End If
'
'    Exit Function
'
'eh:
'    Set l_obj = Nothing
'    TrsRaiseError m_strCMODULENAME, "GetTemplateLocationSelections"
'
'End Function
'
'
'Private Function GetTemplateWorkGroupSelections( _
'    ByVal v_lngTemplateNum As Long _
'    ) As ADODB.Recordset
'
'    Dim l_strDataStore As String
'    Dim l_obj As dbOTTemplate.cdOTTemplate
'
'On Error GoTo eh
'
'    l_strDataStore = Trim(m_rsTemplateData!LOC_CD_SELECT)
'    If l_strDataStore = "T" Then 'In Template
'        Set l_obj = CreateObject("dbOTTemplate.cdOTTemplate")
'        Set GetTemplateWorkGroupSelections = l_obj.GetTemplateWorkGroupSelections(v_lngTemplateNum)
'        Set l_obj = Nothing
'    Else
'        Set GetTemplateWorkGroupSelections = Nothing
'    End If
'
'    Exit Function
'
'eh:
'    Set l_obj = Nothing
'    TrsRaiseError m_strCMODULENAME, "GetTemplateWorkGroupSelections"
'
'End Function

'
'Private Function GetTemplateWorkLimitRules( _
'    ByVal v_lngTemplateNum As Long _
'    ) As ADODB.Recordset
'
'
'    Dim l_strWLRules As String
'    Dim l_obj As dbOTTemplate.cdOTTemplate
'    Dim l_intI As Integer
'    Dim l_intNumRules As Integer
'    Dim l_strWLRule As String
'    Dim l_intStartPos As Integer
'    Dim l_strWLRulesAry As Variant
'
'On Error GoTo eh
'
'    l_strWLRules = Trim(m_rsTemplateData!WORK_LIMITS)
'
'    If l_strWLRules = "" Then 'No Rules
'        Set GetTemplateWorkLimitRules = Nothing
'        Exit Function
'    End If
'
'    l_intNumRules = Len(l_strWLRules) / 2
'    ReDim l_strWLRulesAry(1 To l_intNumRules) As String
'    For l_intI = 1 To l_intNumRules
'        l_intStartPos = (l_intI - 1) * 2 + 1
'        l_strWLRule = Mid(l_strWLRules, l_intStartPos, 2)
'        l_strWLRulesAry(l_intI) = l_strWLRule
'    Next l_intI
'
'    Set l_obj = CreateObject("dbOTTemplate.cdOTTemplate")
'    Set GetTemplateWorkLimitRules = l_obj.GetTemplateWorkLimitRules(v_lngTemplateNum, l_strWLRulesAry)
'    Set l_obj = Nothing
'
'
'    Exit Function
'
'eh:
'    Set l_obj = Nothing
'    TrsRaiseError m_strCMODULENAME, "GetTemplateWorkLimitRules"
'
'End Function
'
'
'Private Function GetAvailSortColumns() As ADODB.Recordset
'
'    Dim l_intI As Integer
'    Dim l_rs As ADODB.Recordset
'
'
'    On Error GoTo eh
'
'    GetSortOrderColumns
'    Set l_rs = CreateAvailSortRecordset()
'    For l_intI = 1 To 24
'        l_rs.AddNew
'        l_rs!SORT_CD = l_intI
'        l_rs!SORT_NAME = m_strSortColumns(l_intI)
'    Next l_intI
'
'    Set GetAvailSortColumns = l_rs
'
'    Exit Function
'
'eh:
'    TrsRaiseError m_strCMODULENAME, "GetAvailSortColumns"
'
'End Function
'
'
'Private Function CreateAvailSortRecordset() As ADODB.Recordset
'
'    Dim l_rs As ADODB.Recordset
'
'    On Error GoTo eh
'
'    Set l_rs = New ADODB.Recordset
'    With l_rs
'      .Fields.Append "SORT_CD", adInteger, , adFldUpdatable
'      .Fields.Append "SORT_NAME", adChar, 30, adFldUpdatable
'      .CursorLocation = adUseClient
'      .CursorType = adOpenDynamic
'      .LockType = adLockOptimistic
'      .Open
'    End With
'
'
'    Set CreateAvailSortRecordset = l_rs
'    Exit Function
'
'eh:
'    Set l_rs = Nothing
'    TrsRaiseError m_strCMODULENAME, "CreateAvailSortRecordset"
'
'End Function
'
'
'Private Function GetAvailGroupSortColumns() As ADODB.Recordset
'
'    Dim l_intI As Integer
'    Dim l_rs As ADODB.Recordset
'
'
'    On Error GoTo eh
'
'    GetSortOrderColumns
'    Set l_rs = CreateAvailSortRecordset()
'    For l_intI = 1 To 4
'        l_rs.AddNew
'        l_rs!SORT_CD = l_intI
'        l_rs!SORT_NAME = m_strGroupSortColumns(l_intI)
'    Next l_intI
'
'    Set GetAvailGroupSortColumns = l_rs
'
'    Exit Function
'
'eh:
'    TrsRaiseError m_strCMODULENAME, "GetAvailGroupSortColumns"
'
'End Function
'
'
'Private Sub GetSortOrderColumns()
'
'    On Error GoTo eh
'
'    m_strSortColumns(1) = "OT Hours Low"
'    m_strSortColumns(2) = "OT Hours High"
'    m_strSortColumns(3) = "Plant Seniority Low"
'    m_strSortColumns(4) = "Plant Seniority High"
'    m_strSortColumns(5) = "Union Seniority Low"
'    m_strSortColumns(6) = "Union Seniority High"
'    m_strSortColumns(7) = "Corporate Seniority Low"
'    m_strSortColumns(8) = "Corporate Seniority High"
'    m_strSortColumns(9) = "Dept.Seniority Low"
'    m_strSortColumns(10) = "Dept.Seniority High"
'    m_strSortColumns(11) = "Employee No.Low"
'    m_strSortColumns(12) = "Employee No.High"
'    m_strSortColumns(13) = "Job Seniority Low"
'    m_strSortColumns(14) = "Job Seniority High"
'    m_strSortColumns(15) = "Age Low"
'    m_strSortColumns(16) = "Age High"
'    m_strSortColumns(17) = "No.of Opp.Low"
'    m_strSortColumns(18) = "No.of Opp.High"
'    m_strSortColumns(19) = "Craft Low"
'    m_strSortColumns(20) = "Craft High"
'    m_strSortColumns(21) = "LOP Seniority Low"
'    m_strSortColumns(22) = "LOP Seniority High"
'    m_strSortColumns(23) = "Company Seniority Low"
'    m_strSortColumns(24) = "Company Seniority High"
'
'    m_strGroupSortColumns(1) = "Resident Loc. First"
'    m_strGroupSortColumns(2) = "Craft"
'    m_strGroupSortColumns(3) = "Job Class"
'    m_strGroupSortColumns(4) = "Work Group"
'
'    Exit Sub
'
'eh:
'    TrsRaiseError m_strCMODULENAME, "GetSortOrderColumns"
'
'End Sub


