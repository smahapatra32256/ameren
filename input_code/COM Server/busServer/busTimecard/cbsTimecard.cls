VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "cbsTimecard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'//CRAC via client.comintegration.library
'//need to add a reference in the project file
'//Reference=*\G{E4D05D57-EFCF-4F0A-9F36-FA0E5A85D969}#1.0#0#..\..\..\..\TRISSandboxes\TRISSandbox\TRISMaster\Client\Ameren.TRIS.Client.COMIntegration.Library\bin\Debug\Ameren.TRIS.Client.COMIntegration.Library.tlb#Ameren_TRIS_Client_COMIntegration_Library

'//CRAC via client.comintegration.library
'Implements IObjectConstruct

'**********************************************************************
'Steps to add a new persistent property to cboTimecardEntry.
'dbTimeCard
'1.  Add the new property to the SQL in cdTimecard.Get_Timecard.
'  This provides a slot in data manager for the property value and
'  sets up the field in the recordset that will be used to collect
'  updates to pass along to dbSaveTimecard.
'
'dbSaveTimecard
'2.  There should not be anything to do here since it will build
'  the SQL statement based on the values and fields descripitons
'  in the recordset passed.  However, if the new persistent field
'  value is set by the database, eg. timestamp, you may need to
'  include the field name in the ignore list.  The ignore list
'  informs the SQL builder to ignore fields in the recordset that
'  match names in the list.
'
'busTimecard
'3.  Add the new property to the updated save recordset in
'  cbsTimecard.AddNewTimeEntry and UpdateTimeEntry.
'4.  Move the new propety values in the request to the new or
'  existing cboTimecardEntry.  This should be done in
'  cboTimecard.AddTimeEntry and UpdateTimecardEntry.
'5.  Ensure unchanged values are ignored by adding the new
'  field in the cboTimecard.SetUnchangedFieldsToNull preprocessing
'  for update requests.
'6.  If the new field should have a default assigned or derived
'  value when a value is not provided on new time entries,
'  provide support in cboTimecard.BuildMissingComponents.
'7.  Add property Get and Let procedures for the new field in
'  cboTimecardEntry.
'8.  Ensure the new cboTimecardEntry property is propagated
'  to a copy of itself in cboTimecardEntry.SplitSimple.
'9.  Ensure the new property is validated in cboTimecardEntry.
'**************************************************************************
'**************************************************************************
'***             CHANGE MANAGEMENT                                      ***
'***    Date        Programmer  Description                             ***
'***  07/25/2001       JFB      PTS 533 -- Allow Overtime entry overlap ***
'***                             a vacation entry                       ***
'***                                                                    ***
'**************************************************************************

Private Const m_strCMODULENAME = "cbsTimecard"

Public Enum TRSReturnOption
    TrActualTimeEntriesDetail = 1
    TrTotalHoursByActivitySummary = 2
    TrActualTimeEntriesDetailAndTotalHoursByActivitySummary = 3
End Enum

Private Enum TRSViewOption
    TrDept = 1
    TrWG = 2
End Enum

Private Enum TrsPrivateTimecardErrorEnum
    trLockedBypayroll = 30103 + vbObjectError
    trProcessedBypayroll = 30104
End Enum
    
Private m_strUserId As String   'employee number of user
Private m_datCollPerStart As Date
Private m_datCollPerEnd As Date

Private mobjTimecard As cboTimecard
Private mobjDepartmentManager As dbDepartment.cdDepartment
Private mobjJobManager As dbCodeTables.cdCodeTables
Private m_objErrors As ADODB.Recordset
Private mobjOtherPay As cboOtherPay

Private m_objContractRulesErrors As ADODB.Recordset
Private m_objOTTCConversionErrors As ADODB.Recordset
Private m_objUserEnteredDataForCRAC As ADODB.Recordset
'for older clients (6.0.0) that are not sending this column to the server
Private m_blnPreArrangedOTIndMissing As Boolean
Private m_rsHolidays As ADODB.Recordset

'//CRAC via client.comintegration.library
''//-----Begin Getting custom settings from COM+ rather than config or registry-----//
''in Component Services, TRIS Components, select component, context menu property, activation tab, check enable object construction
''set Constructor String to ...
''CRACInterface=2;CRACInterfaceLog=1
'
'Dim m_ConstructString As String
'Dim m_colConstructStringProperty As Collection
'Dim CRACInterface As Integer '1 - Legacy/COM+/In process;2 - WCF/Out of process
'Dim CRACInterfaceLog As Integer '0 - off;1 - on
'
'Private Sub Class_Initialize()
'    CRACInterface = 1 'Legacy/COM+
'    CRACInterfaceLog = 0 'Off
'End Sub
'
'Private Sub IObjectConstruct_Construct(ByVal pCtorObj As Object)
'    Dim lConstructString As IObjectConstructString
'
'    On Error GoTo eh
'    Set lConstructString = pCtorObj
'
'    m_ConstructString = lConstructString.ConstructString
'
'    Set m_colConstructStringProperty = ParseContructString(m_ConstructString)
'
'    PostConstruct
'    '//App.LogEvent MODULENAME & " ObjectConstructString = " & m_ConstructString, vbLogEventTypeInformation
'
'    Exit Sub
'eh:
'    TrsRaiseError m_strCMODULENAME, "IObjectConstruct_Construct", , , , DescArgsFrom(m_ConstructString), True
'End Sub
'
'Private Sub PostConstruct()
'    On Error GoTo eh
'
'    If Not m_colConstructStringProperty Is Nothing Then
'        If DoesPropertyExist(m_colConstructStringProperty, "CRACInterface") Then
'            If IsNumeric(m_colConstructStringProperty.Item("CRACInterface")) Then
'                CRACInterface = CInt(m_colConstructStringProperty.Item("CRACInterface"))
'            End If
'        End If
'
'        If DoesPropertyExist(m_colConstructStringProperty, "CRACInterfaceLog") Then
'            If IsNumeric(m_colConstructStringProperty.Item("CRACInterfaceLog")) Then
'                CRACInterfaceLog = CInt(m_colConstructStringProperty.Item("CRACInterfaceLog"))
'            End If
'        End If
'    End If
'
'    Exit Sub
'eh:
'    TrsRaiseError m_strCMODULENAME, "PostConstruct", , , , m_colConstructStringProperty, True
'End Sub
'
'Private Function DoesPropertyExist(ByVal colPropertyValue As Collection, ByVal Property As String) As Boolean
'    Dim oValue As Variant
'
'    On Error GoTo eh
'
'    oValue = colPropertyValue.Item(Property)
'
'    DoesPropertyExist = True
'
'    Exit Function
'eh:
'    DoesPropertyExist = False
'End Function
'
'Private Function ParseContructString(ByVal ConstructString As String) As Collection
'    Dim colConstructStringProperty As Collection
'    Dim PropertyValueArray() As String
'    Dim PropertyValue As Variant
'    Dim PropertyValueArray2() As String
'
'    On Error GoTo eh
'
'    Set colConstructStringProperty = New Collection
'    PropertyValueArray = Split(ConstructString, ";")
'
'    For Each PropertyValue In PropertyValueArray
'        PropertyValueArray2 = Split(PropertyValue, "=")
'        If UBound(PropertyValueArray2) >= 1 Then
'            colConstructStringProperty.Add PropertyValueArray2(1), PropertyValueArray2(0)
'        End If
'    Next
'
'    Set ParseContructString = colConstructStringProperty
'
'    Exit Function
'eh:
'    TrsRaiseError m_strCMODULENAME, "ParseContructString", , , , DescArgsFrom(ConstructString), True
'End Function
''//-----End Getting custom settings from COM+ rather than config or registry-----//

Private Property Get PreArrangedOTIndMissing() As Boolean
    PreArrangedOTIndMissing = m_blnPreArrangedOTIndMissing
End Property

Private Property Let PreArrangedOTIndMissing(blnPreArrangedOTIndMissing As Boolean)
    m_blnPreArrangedOTIndMissing = blnPreArrangedOTIndMissing
End Property

Public Function UserEmployeeNumber() As String
' made available when client provides, e.g. RecordTimeEntry
    UserEmployeeNumber = m_strUserId
End Function

Public Function ErrorSet() As ADODB.Recordset
    Set ErrorSet = m_objErrors
End Function

Public Function GetDepartmentManager() As dbDepartment.cdDepartment

On Error GoTo eh
    If mobjDepartmentManager Is Nothing Then
        Set mobjDepartmentManager = CreateObject("dbDepartment.cdDepartment")
    End If
    Set GetDepartmentManager = mobjDepartmentManager
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "GetDepartmentManager"
    
End Function

Public Function GetJobManager() As dbCodeTables.cdCodeTables

On Error GoTo eh
    If mobjJobManager Is Nothing Then
        Set mobjJobManager = CreateObject("dbCodeTables.cdCodeTables")
    End If
    Set GetJobManager = mobjJobManager
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "GetJobManager"
End Function

' Removed 3/15/2002
'Public Function GetDistributionManager() As busTimecard.cuDistributionManager

'//not sure where this is used, pick list description is 30 chars
Public Function GetActualTimeDetail( _
        ByVal v_lngEmpNum As Long, _
        ByVal v_datStartDate As Date, _
        ByVal v_datEndDate As Date) As ADODB.Recordset
    
    Dim l_objRSActualTimeDetail As ADODB.Recordset
    Dim l_objRSTotalHoursByActivitySummary As ADODB.Recordset

    On Error GoTo GetActualTimeDetail_EH
    
    Call RetrieveActualTimeEntriesForSummaryAndDetail(v_lngEmpNum, v_datStartDate, v_datEndDate, TrActualTimeEntriesDetail, l_objRSActualTimeDetail, l_objRSTotalHoursByActivitySummary, False)
    Set GetActualTimeDetail = l_objRSActualTimeDetail
    Exit Function

GetActualTimeDetail_EH:
    Set l_objRSActualTimeDetail = Nothing
    Set l_objRSTotalHoursByActivitySummary = Nothing
    TrsRaiseError m_strCMODULENAME, "GetActualTimeDetail", , , , , True
End Function

'//not sure where this is used, pick list description is 30 chars
Public Function GetActualTimeSummary( _
        ByVal v_lngEmpNum As Long, _
        ByVal v_datStartDate As Date, _
        ByVal v_datEndDate As Date) As ADODB.Recordset
    
    Dim l_objRSActualTimeDetail As ADODB.Recordset
    Dim l_objRSTotalHoursByActivitySummary As ADODB.Recordset

    On Error GoTo GetActualTimeSummary_EH
    
    Call RetrieveActualTimeEntriesForSummaryAndDetail(v_lngEmpNum, v_datStartDate, v_datEndDate, TrTotalHoursByActivitySummary, l_objRSActualTimeDetail, l_objRSTotalHoursByActivitySummary, False)
    
    Set GetActualTimeSummary = l_objRSTotalHoursByActivitySummary
    Exit Function

GetActualTimeSummary_EH:
    '2/2/01 rsg
    Set l_objRSActualTimeDetail = Nothing
    Set l_objRSTotalHoursByActivitySummary = Nothing
    TrsRaiseError m_strCMODULENAME, "GetActualTimeSummary", , , , , True
End Function

'''''Public Function GetActualTimeDetailAndSummary( _
'''''        ByVal v_lngEmpNum As Long, _
'''''        ByVal v_datStartDate As Date, _
'''''        ByVal v_datEndDate As Date, _
'''''        ByRef r_objRSActualTimeDetail As ADODB.Recordset, _
'''''        ByRef r_objRSTotalHoursByActivitySummary As ADODB.Recordset)
'''''' r_objRSActualTimeDetail and r_objRSTotalHoursByActivitySummary
'''''' are output parameters.
'''''    On Error GoTo GetActualTimeDetailAndSummary_EH
'''''
'''''    Call RetrieveActualTimeEntriesForSummaryAndDetail(v_lngEmpNum, v_datStartDate, v_datEndDate, TrActualTimeEntriesDetailAndTotalHoursByActivitySummary, r_objRSActualTimeDetail, r_objRSTotalHoursByActivitySummary, False)
'''''    Exit Function
'''''
'''''GetActualTimeDetailAndSummary_EH:
'''''    Set r_objRSActualTimeDetail = Nothing
'''''    Set r_objRSTotalHoursByActivitySummary = Nothing
'''''    TrsRaiseError m_strCMODULENAME, "GetActualTimeDetailAndSummary", , , , , True
'''''End Function
'''''
'''''Public Function GetOtherPayEntries( _
'''''        ByVal v_lngEmpNum As Long, _
'''''        ByVal v_datStartDate As Date, _
'''''        ByVal v_datEndDate As Date) As ADODB.Recordset
'''''
'''''    Dim l_objCtx As ObjectContext
'''''    Dim l_objdbTimecard As dbTimeCard.cdTimeCard
'''''
'''''On Error GoTo GetOtherPayEntries_EH
'''''
'''''    Set l_objCtx = GetObjectContext
'''''    Set l_objdbTimecard = CreateObject("dbTimeCard.cdTimeCard")
'''''    Set GetOtherPayEntries = l_objdbTimecard.GetOtherPayEntries(v_lngEmpNum, v_datStartDate, v_datEndDate)
'''''    Set l_objdbTimecard = Nothing
'''''    Set l_objCtx = Nothing
'''''    Exit Function
'''''
'''''GetOtherPayEntries_EH:
'''''    Set l_objCtx = Nothing
'''''    Set l_objdbTimecard = Nothing
'''''    TrsRaiseError m_strCMODULENAME, "GetOtherPayEntries"
'''''End Function


Private Function RetrieveActualTimeEntriesForSummaryAndDetail( _
        ByVal v_aryEmployees As Variant, _
        ByVal v_datStartDate As Date, _
        ByVal v_datEndDate As Date, _
        ByVal v_ReturnOption As TRSReturnOption, _
        ByRef r_objRSActualTimeDetail As ADODB.Recordset, _
        ByRef r_objRSTotalHoursByActivitySummary As ADODB.Recordset, _
        ByVal v_blnMultipleEmployees As Boolean, _
        Optional ByVal v_blnIncludeOtherPay As Boolean = False, _
        Optional ByVal v_intDescriptionLength As Integer = 39)
' r_objRSActualTimeDetail and r_objRSTotalHoursByActivitySummary
' are output parameters.
    Dim l_objGetOverExccdAccountDescription As ADODB.Recordset
    Dim l_objGetActualTimeRecord As ADODB.Recordset
    Dim l_objGetActualTimeEntries As dbTimeCard.cdTimeCard
    Dim l_objGetAccountDescription As busDistributionMgr.cbsDistributionMgr
    'Dim l_objGetAccountDescription As Object
    Dim l_objGetOverExccdDescription As dbTimeCard.cdTimeCard
    Dim l_objGetWGDistributionIDDescriptions As busDistributionMgr.cbsDistributionMgr
    Dim l_objRsPickListDistribution As ADODB.Recordset
    Dim l_strGetAccountDescription As String
    Dim l_objGetSubsetActualTimeRecord As New ADODB.Recordset
    Dim l_colWG As New Collection, lcolWorkgroup As New Collection
    Dim l_colOverExccd As New Collection
    Dim l_prevWG As Long
    Dim l_prevDeptNum As String
    Dim l_strFilter As String
    Dim l_lngEmpNum As Long
    Dim l_strWg As String
    Dim l_strOEC As String
    Dim l_objGetJobTaskDistributionidDesc As dbDistributionMgr.cdDistributionMgr
    Dim l_objRsJobTaskDistribution As ADODB.Recordset
    Dim l_strDeptNum As String
    '//Dim l_prevJobTaskMaximo As String
    Dim l_strJobTaskMaximo As String
    Dim l_strJobNumMaximo As String
    Dim l_strTaskNumMaximo As String
    Dim l_objRsJobTaskDistributionMaximo As ADODB.Recordset
    
    On Error GoTo RetrieveActualTimeEntriesForSummaryAndDetail_EH
    
    Set l_objRsJobTaskDistributionMaximo = Nothing

    If v_ReturnOption = TrActualTimeEntriesDetail Then 'Begin Clone subset for summary recordset request
    Else
        If v_blnMultipleEmployees Then
            l_objGetSubsetActualTimeRecord.Fields.Append "EMP_NUM", adInteger
        End If
        l_objGetSubsetActualTimeRecord.Fields.Append "DISTRIBUTION_ID", adChar, 39
        l_objGetSubsetActualTimeRecord.Fields.Append "EXCEPTION_CD", adChar, 3
        l_objGetSubsetActualTimeRecord.Fields.Append "OVER_EXC_CD", adChar, 3
        l_objGetSubsetActualTimeRecord.Fields.Append "CHARGE_DATE", adDate
        l_objGetSubsetActualTimeRecord.Fields.Append "HRS_PER_ACTIVITY", adDecimal
        l_objGetSubsetActualTimeRecord.Fields.Item("HRS_PER_ACTIVITY").Precision = 3
        l_objGetSubsetActualTimeRecord.Fields.Item("HRS_PER_ACTIVITY").NumericScale = 1
        l_objGetSubsetActualTimeRecord.Fields.Append "DISTRIBUTION_DESC", adChar, v_intDescriptionLength
        l_objGetSubsetActualTimeRecord.Fields.Append "DISTRIBUTION_TYPE", adChar, 1
        l_objGetSubsetActualTimeRecord.Fields.Append "JOB_TASK", adChar, 15
        l_objGetSubsetActualTimeRecord.Fields.Append "JOBTASK_STEP_NBR", adInteger
        l_objGetSubsetActualTimeRecord.Fields.Append "EMPRV_DATABASE_ID", adChar, 8
        l_objGetSubsetActualTimeRecord.Open
    End If  'Begin Clone subset for summary recordset request
    
    Set l_objGetActualTimeEntries = CreateObject("dbTimeCard.cdTimeCard")
'rsg 8/29/00 - GetActualTimeEntries should only return actual time,
' but for now it returns other pay for the client.  The client prefers
' actual time and other pay to be separate.  We will fix the method
' names when the recordsets are separated.

    If v_blnMultipleEmployees Then
        If v_blnIncludeOtherPay Then
            Set l_objGetActualTimeRecord = l_objGetActualTimeEntries.GetActualTimeEntries2(v_aryEmployees, v_datStartDate, v_datEndDate, v_intDescriptionLength)
        Else
            Set l_objGetActualTimeRecord = l_objGetActualTimeEntries.GetActualTimeEntriesWithoutOtherPay2(v_aryEmployees, v_datStartDate, v_datEndDate, v_intDescriptionLength)
        End If
    Else
        l_lngEmpNum = v_aryEmployees
        If v_blnIncludeOtherPay Then
            Set l_objGetActualTimeRecord = l_objGetActualTimeEntries.GetActualTimeEntries2(l_lngEmpNum, v_datStartDate, v_datEndDate, v_intDescriptionLength)
        Else
            Set l_objGetActualTimeRecord = l_objGetActualTimeEntries.GetActualTimeEntriesWithoutOtherPay2(l_lngEmpNum, v_datStartDate, v_datEndDate, v_intDescriptionLength)
        End If
    End If

    If (l_objGetActualTimeRecord.EOF And l_objGetActualTimeRecord.BOF) Then
       If v_ReturnOption = TrActualTimeEntriesDetail Then
          Set r_objRSActualTimeDetail = l_objGetActualTimeRecord
          Set r_objRSTotalHoursByActivitySummary = Nothing 'summary was not requested
       Else
          If v_ReturnOption = TrActualTimeEntriesDetailAndTotalHoursByActivitySummary Then
             Set r_objRSActualTimeDetail = l_objGetActualTimeRecord
             Set r_objRSTotalHoursByActivitySummary = l_objGetSubsetActualTimeRecord
          Else
             Set r_objRSActualTimeDetail = Nothing 'actual time was not requested
             Set r_objRSTotalHoursByActivitySummary = l_objGetSubsetActualTimeRecord
          End If
       End If
       Exit Function
    End If
    
    '//build the list of maximo job/task descriptions
    Set l_objGetJobTaskDistributionidDesc = CreateObject("dbDistributionMgr.cdDistributionMgr")
    Set l_objRsJobTaskDistributionMaximo = l_objGetJobTaskDistributionidDesc.GetDistJobTaskListMultiple(l_objGetActualTimeRecord)
    Set l_objGetJobTaskDistributionidDesc = Nothing
       
    'got all the actual time scope with bank code now create collection of wg and exc_cd to get descriptions from distribution_id or override_code
    Do Until l_objGetActualTimeRecord.EOF     '       process first pay period table (for splits) - overlay correct dep no effective date l_intSplitPeriod_flag = 1
        l_strWg = Trim(str(l_objGetActualTimeRecord!WG_NUM.Value))
        If Not IncludesString(l_strWg, l_colWG) Then
            l_colWG.Add l_strWg
        End If
        l_strOEC = Trim(l_objGetActualTimeRecord!OVER_EXC_CD.Value)
        If l_strOEC <> "" Then
            If Not IncludesString(l_strOEC, l_colOverExccd) Then
                l_colOverExccd.Add l_strOEC
            End If
        End If
       l_objGetActualTimeRecord.MoveNext
    Loop

    Set l_objGetWGDistributionIDDescriptions = CreateObject("busDistributionMgr.cbsDistributionMgr")
    Set l_objRsPickListDistribution = l_objGetWGDistributionIDDescriptions.GetWGDistributionIDDescriptions3(l_colWG, v_intDescriptionLength, "ActualTimeDesc") 'get internal picklist table

    If l_colOverExccd.Count < 1 Then 'if no override to process do not create override table
    Else
        Set l_objGetOverExccdDescription = CreateObject("dbTimeCard.cdTimeCard")
        Set l_objGetOverExccdAccountDescription = l_objGetOverExccdDescription.GetAccountOverrideDescription(l_colOverExccd)
    End If

    l_prevWG = 0
    l_prevDeptNum = ""

    l_objGetActualTimeRecord.MoveFirst

    Do Until l_objGetActualTimeRecord.EOF
       If Trim(l_objGetActualTimeRecord!DISTRIBUTION_TYPE) <> "" Then
           Set l_objGetAccountDescription = CreateObject("busDistributionMgr.cbsDistributionMgr")

           If Trim(l_objGetActualTimeRecord!DISTRIBUTION_TYPE) = "J" And _
              Trim(l_objGetActualTimeRecord!JOB_TASK) <> "" Then
                If l_prevWG <> l_objGetActualTimeRecord!WG_NUM Or _
                   l_prevDeptNum <> Trim(l_objGetActualTimeRecord!DEPT_NUM) Then
                    'Remove from Collection
                    If (lcolWorkgroup.Count > 0) Then
                        lcolWorkgroup.Remove (1)
                    End If

                    'Add to Collection
                    lcolWorkgroup.Add (Trim(l_objGetActualTimeRecord!WG_NUM))

                    '//use the department from the associations table rather than the employee table
                    l_strDeptNum = GetDepartmentOfWorkGroup(l_objGetActualTimeRecord!WG_NUM)
                    
                    'Get the list of Job Task descriptions
                    Set l_objGetJobTaskDistributionidDesc = CreateObject("dbDistributionMgr.cdDistributionMgr")
                    Set l_objRsJobTaskDistribution = l_objGetJobTaskDistributionidDesc.GetDistJobTaskList2(lcolWorkgroup, l_strDeptNum, True)
                    Set l_objGetJobTaskDistributionidDesc = Nothing

                    l_prevWG = l_objGetActualTimeRecord!WG_NUM
                    l_prevDeptNum = Trim(l_objGetActualTimeRecord!DEPT_NUM)

                    'Call local private method to get description for distribution/job task/EMPRV DB ID
                    l_strGetAccountDescription = GetAccountDescriptionJobTask(l_objRsJobTaskDistribution _
                                                    , l_objGetActualTimeRecord!DISTRIBUTION_ID _
                                                    , Mid(l_objGetActualTimeRecord!JOB_TASK, 1, 10) _
                                                    , l_objGetActualTimeRecord!EMPRV_DATABASE_ID, v_intDescriptionLength)
                Else
                    l_strGetAccountDescription = GetAccountDescriptionJobTask(l_objRsJobTaskDistribution _
                                                    , l_objGetActualTimeRecord!DISTRIBUTION_ID _
                                                    , Mid(l_objGetActualTimeRecord!JOB_TASK, 1, 10) _
                                                    , l_objGetActualTimeRecord!EMPRV_DATABASE_ID, v_intDescriptionLength)
                End If
           Else
            l_strJobTaskMaximo = Trim(l_objGetActualTimeRecord!JOB_TASK)
            If Trim(l_objGetActualTimeRecord!DISTRIBUTION_TYPE) = "M" And l_strJobTaskMaximo <> "" Then
                l_strJobNumMaximo = Left(l_strJobTaskMaximo, 7)
                l_strTaskNumMaximo = Right(l_strJobTaskMaximo, 3)
                l_strGetAccountDescription = GetAccountDescriptionJobTaskMaximo(l_objRsJobTaskDistributionMaximo, l_strJobNumMaximo, l_strTaskNumMaximo, l_objGetActualTimeRecord!EMPRV_DATABASE_ID, v_intDescriptionLength)
            Else
               l_strGetAccountDescription = l_objGetAccountDescription.GetAccountDescription2(l_objRsPickListDistribution, l_objGetActualTimeRecord!WG_NUM.Value, l_objGetActualTimeRecord!DISTRIBUTION_ID.Value, l_objGetActualTimeRecord!DISTRIBUTION_TYPE.Value, l_objGetActualTimeRecord!DISTR_DEPT_NUM.Value, v_intDescriptionLength)
            End If
           End If

           If Trim(l_strGetAccountDescription) <> "" Then
              l_objGetActualTimeRecord!DISTRIBUTION_DESC.Value = PadRight(l_strGetAccountDescription, v_intDescriptionLength)
           Else 'populate with the override description table if exist or default exception description 'l_strGetAccountDescription = from filter override desc
              If l_objGetOverExccdAccountDescription Is Nothing Then
                '
              Else
                l_objGetOverExccdAccountDescription.Filter = "EXCEPTION_CD= '" & l_objGetActualTimeRecord!OVER_EXC_CD.Value & " '"
                If Not (l_objGetOverExccdAccountDescription.EOF And l_objGetOverExccdAccountDescription.BOF) Then
                   l_strGetAccountDescription = PadRight(l_objGetOverExccdAccountDescription!EXCEPTION_DESC.Value, v_intDescriptionLength)
                   l_objGetActualTimeRecord!DISTRIBUTION_DESC.Value = l_strGetAccountDescription
                Else    'use exception description default
                    '
                End If
              End If
           End If
       Else 'populate with the override description table if exist or default exception description 'l_strGetAccountDescription = from filter override desc
           If l_objGetOverExccdAccountDescription Is Nothing Then
            '
           Else
               l_objGetOverExccdAccountDescription.Filter = "EXCEPTION_CD= '" & l_objGetActualTimeRecord!OVER_EXC_CD.Value & " '"

               If Not (l_objGetOverExccdAccountDescription.EOF And l_objGetOverExccdAccountDescription.BOF) Then
                  l_strGetAccountDescription = PadRight(l_objGetOverExccdAccountDescription!EXCEPTION_DESC.Value, v_intDescriptionLength)
                  l_objGetActualTimeRecord!DISTRIBUTION_DESC.Value = l_strGetAccountDescription
               Else     'use exception description default
                    '
               End If
           End If
       End If

       Set l_objGetAccountDescription = Nothing

       If Left$(l_objGetActualTimeRecord!START_TIME.Value, 2) = "99" Then
           l_objGetActualTimeRecord!TIME_ENTRY_TYPE.Value = 1
       Else
           If l_objGetActualTimeRecord!APPR_IND.Value = "Y" Then
                If Left$(l_objGetActualTimeRecord!EXCEPTION_BANK_CD.Value, 1) = "O" _
                Or l_objGetActualTimeRecord!EXCEPTION_BANK_CD.Value = "C+" Then
                    l_objGetActualTimeRecord!TIME_ENTRY_TYPE.Value = 2
                Else
                    If Trim(l_objGetActualTimeRecord!SCHED_CHG_EXC_CD.Value) <> "" Then
                        l_objGetActualTimeRecord!TIME_ENTRY_TYPE.Value = 3
                    Else
                        l_objGetActualTimeRecord!TIME_ENTRY_TYPE.Value = 4
                    End If
                End If
           Else
                If Left$(l_objGetActualTimeRecord!EXCEPTION_BANK_CD.Value, 1) = "O" _
                Or l_objGetActualTimeRecord!EXCEPTION_BANK_CD.Value = "C+" Then
                   l_objGetActualTimeRecord!TIME_ENTRY_TYPE.Value = 5
                Else
                    If Trim(l_objGetActualTimeRecord!SCHED_CHG_EXC_CD.Value) <> "" Then
                        l_objGetActualTimeRecord!TIME_ENTRY_TYPE.Value = 6
                    Else
                        l_objGetActualTimeRecord!TIME_ENTRY_TYPE.Value = 7
                    End If
                End If
           End If
       End If

       l_objGetActualTimeRecord.Update

       If v_ReturnOption = TrActualTimeEntriesDetail Then 'Add to Clone subset  for summary recordset request
        '
       Else 'building summary recordset from actual time recordset
           l_objGetSubsetActualTimeRecord.AddNew
           If v_blnMultipleEmployees Then
                l_objGetSubsetActualTimeRecord("EMP_NUM") = l_objGetActualTimeRecord!EMP_NUM.Value
           End If

           l_objGetSubsetActualTimeRecord("DISTRIBUTION_ID") = l_objGetActualTimeRecord!DISTRIBUTION_ID.Value
           l_objGetSubsetActualTimeRecord("EXCEPTION_CD") = l_objGetActualTimeRecord!EXCEPTION_CD.Value
           l_objGetSubsetActualTimeRecord("OVER_EXC_CD") = l_objGetActualTimeRecord!OVER_EXC_CD.Value
           l_objGetSubsetActualTimeRecord("CHARGE_DATE") = l_objGetActualTimeRecord!CHARGE_DATE.Value
           l_objGetSubsetActualTimeRecord("HRS_PER_ACTIVITY") = l_objGetActualTimeRecord!Hours.Value
           l_objGetSubsetActualTimeRecord("DISTRIBUTION_DESC") = PadRight(l_objGetActualTimeRecord!DISTRIBUTION_DESC.Value, v_intDescriptionLength)
           l_objGetSubsetActualTimeRecord("DISTRIBUTION_TYPE") = l_objGetActualTimeRecord!DISTRIBUTION_TYPE.Value
           l_objGetSubsetActualTimeRecord("JOB_TASK") = l_objGetActualTimeRecord!JOB_TASK
           l_objGetSubsetActualTimeRecord("JOBTASK_STEP_NBR") = l_objGetActualTimeRecord!JOBTASK_STEP_NBR.Value
           l_objGetSubsetActualTimeRecord("EMPRV_DATABASE_ID") = l_objGetActualTimeRecord!EMPRV_DATABASE_ID.Value
           l_objGetSubsetActualTimeRecord.Update
       End If       'Add to Clone subset  for summary recordset request

       l_objGetActualTimeRecord.MoveNext
    Loop
    
    If v_ReturnOption = TrActualTimeEntriesDetail Then
       l_objGetActualTimeRecord.MoveFirst
       Set r_objRSActualTimeDetail = l_objGetActualTimeRecord
       Set r_objRSTotalHoursByActivitySummary = Nothing
       Exit Function
    Else
        If v_ReturnOption = TrActualTimeEntriesDetailAndTotalHoursByActivitySummary Then
           l_objGetActualTimeRecord.MoveFirst
           Set r_objRSActualTimeDetail = l_objGetActualTimeRecord
        Else
           Set r_objRSActualTimeDetail = Nothing
       End If
       
       Set r_objRSTotalHoursByActivitySummary = RetrieveTotalHoursByActivitySummary(l_objGetSubsetActualTimeRecord, v_blnMultipleEmployees, v_intDescriptionLength)
    End If

    Exit Function

RetrieveActualTimeEntriesForSummaryAndDetail_EH:
    '2/2/01 rsg
    Set r_objRSActualTimeDetail = Nothing
    Set r_objRSTotalHoursByActivitySummary = Nothing
    TrsRaiseError m_strCMODULENAME, "RetrieveActualTimeEntriesForSummaryAndDetail"
End Function

'//-----
''//-----The following was created to allow the Callaway WebService to run faster
'Private Function RetrieveActualTimeEntriesForSummaryAndDetailForDept( _
'        ByVal v_aryDepartments As Variant, _
'        ByVal v_datStartDate As Date, _
'        ByVal v_datEndDate As Date, _
'        ByVal v_ReturnOption As TRSReturnOption, _
'        ByRef r_objRSActualTimeDetail As ADODB.Recordset, _
'        ByRef r_objRSTotalHoursByActivitySummary As ADODB.Recordset, _
'        ByVal v_blnMultipleEmployees As Boolean, _
'        Optional ByVal v_blnIncludeOtherPay As Boolean = False)
'' r_objRSActualTimeDetail and r_objRSTotalHoursByActivitySummary
'' are output parameters.
''    Dim l_objGetOverExccdAccountDescription As ADODB.Recordset
'    Dim l_objGetActualTimeRecord As ADODB.Recordset
'    Dim l_objGetActualTimeEntries As dbTimeCard.cdTimeCard
''    Dim l_objCtx As ObjectContext
''    Dim l_objGetAccountDescription As busDistributionMgr.cbsDistributionMgr
''    Dim l_objGetOverExccdDescription As dbTimeCard.cdTimeCard
''    Dim l_objGetWGDistributionIDDescriptions As busDistributionMgr.cbsDistributionMgr
''    Dim l_objRsPickListDistribution As ADODB.Recordset
''    Dim l_strGetAccountDescription As String
'    Dim l_objGetSubsetActualTimeRecord As New ADODB.Recordset
''    Dim l_strFilter As String
'    Dim l_strDept As String
'
'    On Error GoTo eh
'
'    If v_ReturnOption = TrActualTimeEntriesDetail Then 'Begin Clone subset for summary recordset request
'    Else
'        If v_blnMultipleEmployees Then
'            l_objGetSubsetActualTimeRecord.Fields.Append "EMP_NUM", adInteger
'        End If
'        l_objGetSubsetActualTimeRecord.Fields.Append "DISTRIBUTION_ID", adChar, 39
'        l_objGetSubsetActualTimeRecord.Fields.Append "EXCEPTION_CD", adChar, 3
'        l_objGetSubsetActualTimeRecord.Fields.Append "OVER_EXC_CD", adChar, 3
'        l_objGetSubsetActualTimeRecord.Fields.Append "CHARGE_DATE", adDate
'        l_objGetSubsetActualTimeRecord.Fields.Append "HRS_PER_ACTIVITY", adDecimal
'        l_objGetSubsetActualTimeRecord.Fields.Item("HRS_PER_ACTIVITY").Precision = 3
'        l_objGetSubsetActualTimeRecord.Fields.Item("HRS_PER_ACTIVITY").NumericScale = 1
'        l_objGetSubsetActualTimeRecord.Fields.Append "DISTRIBUTION_DESC", adChar, 30
'        l_objGetSubsetActualTimeRecord.Fields.Append "DISTRIBUTION_TYPE", adChar, 1
'        l_objGetSubsetActualTimeRecord.Open
'    End If  'Begin Clone subset for summary recordset request
'
'    Set l_objGetActualTimeEntries = CreateObject("dbTimeCard.cdTimeCard")
''rsg 8/29/00 - GetActualTimeEntries should only return actual time,
'' but for now it returns other pay for the client.  The client prefers
'' actual time and other pay to be separate.  We will fix the method
'' names when the recordsets are separated.
'
'    If v_blnMultipleEmployees Then
'        If v_blnIncludeOtherPay Then
'            Set l_objGetActualTimeRecord = l_objGetActualTimeEntries.GetActualTimeEntriesForDept(v_aryDepartments, v_datStartDate, v_datEndDate)
'        Else
''            Set l_objGetActualTimeRecord = l_objGetActualTimeEntries.GetActualTimeEntriesWithoutOtherPayForDept(v_aryDepartments, v_datStartDate, v_datEndDate)
'        End If
'    Else
'        l_strDept = v_aryDepartments
'        If v_blnIncludeOtherPay Then
'            Set l_objGetActualTimeRecord = l_objGetActualTimeEntries.GetActualTimeEntriesForDept(l_strDept, v_datStartDate, v_datEndDate)
'        Else
''            Set l_objGetActualTimeRecord = l_objGetActualTimeEntries.GetActualTimeEntriesWithoutOtherPayForDept(l_strDept, v_datStartDate, v_datEndDate)
'        End If
'    End If
'
'    If (l_objGetActualTimeRecord.EOF And l_objGetActualTimeRecord.BOF) Then
'       If v_ReturnOption = TrActualTimeEntriesDetail Then
'          Set r_objRSActualTimeDetail = l_objGetActualTimeRecord
'          Set r_objRSTotalHoursByActivitySummary = Nothing 'summary was not requested
'       Else
'          If v_ReturnOption = TrActualTimeEntriesDetailAndTotalHoursByActivitySummary Then
'             Set r_objRSActualTimeDetail = l_objGetActualTimeRecord
'             Set r_objRSTotalHoursByActivitySummary = l_objGetSubsetActualTimeRecord
'          Else
'             Set r_objRSActualTimeDetail = Nothing 'actual time was not requested
'             Set r_objRSTotalHoursByActivitySummary = l_objGetSubsetActualTimeRecord
'          End If
'       End If
'       Exit Function
'    End If
'
'    '//-----factor out???
'    ApplyOverridesToActualTimeEntriesForSummaryAndDetail l_objGetActualTimeRecord, _
'                                                        l_objGetSubsetActualTimeRecord, _
'                                                        v_ReturnOption, _
'                                                        v_blnMultipleEmployees
''    'got all the actual time scope with bank code now create collection of wg and exc_cd to get descriptions from distribution_id or override_code
''    Do Until l_objGetActualTimeRecord.EOF     '       process first pay period table (for splits) - overlay correct dep no effective date l_intSplitPeriod_flag = 1
''        l_strWg = Trim(str(l_objGetActualTimeRecord!WG_NUM.Value))
''        If Not IncludesString(l_strWg, l_colWG) Then
''            l_colWG.Add l_strWg
''        End If
''        l_strOEC = Trim(l_objGetActualTimeRecord!OVER_EXC_CD.Value)
''        If l_strOEC <> "" Then
''            If Not IncludesString(l_strOEC, l_colOverExccd) Then
''                l_colOverExccd.Add l_strOEC
''            End If
''        End If
''       l_objGetActualTimeRecord.MoveNext
''    Loop
''
''    Set l_objGetWGDistributionIDDescriptions = CreateObject("busDistributionMgr.cbsDistributionMgr")
''    Set l_objRsPickListDistribution = l_objGetWGDistributionIDDescriptions.GetWGDistributionIDDescriptions(l_colWG) 'get internal picklist table
''
''    If l_colOverExccd.Count < 1 Then 'if no override to process do not create override table
''    Else
''        Set l_objGetOverExccdDescription = CreateObject("dbTimeCard.cdTimeCard")
''        Set l_objGetOverExccdAccountDescription = l_objGetOverExccdDescription.GetAccountOverrideDescription(l_colOverExccd)
''    End If
''
''    l_objGetActualTimeRecord.MoveFirst
''
''    Do Until l_objGetActualTimeRecord.EOF
''       If Trim(l_objGetActualTimeRecord!DISTRIBUTION_TYPE.Value) <> "" Then
''           Set l_objGetAccountDescription = CreateObject("busDistributionMgr.cbsDistributionMgr")
''           l_strGetAccountDescription = l_objGetAccountDescription.GetAccountDescription(l_objRsPickListDistribution, l_objGetActualTimeRecord!WG_NUM.Value, l_objGetActualTimeRecord!DISTRIBUTION_ID.Value, l_objGetActualTimeRecord!DISTRIBUTION_TYPE.Value, l_objGetActualTimeRecord!DISTR_DEPT_NUM.Value)
''           If Trim(l_strGetAccountDescription) <> "" Then
''              l_objGetActualTimeRecord!DISTRIBUTION_DESC.Value = l_strGetAccountDescription
''           Else 'populate with the override description table if exist or default exception description 'l_strGetAccountDescription = from filter override desc
''              If l_objGetOverExccdAccountDescription Is Nothing Then
''              Else
''                l_strFilter = "EXCEPTION_CD= '" & l_objGetActualTimeRecord!OVER_EXC_CD.Value & " '"
''                l_objGetOverExccdAccountDescription.Filter = l_strFilter
''                If Not (l_objGetOverExccdAccountDescription.EOF And l_objGetOverExccdAccountDescription.BOF) Then
''                   l_strGetAccountDescription = l_objGetOverExccdAccountDescription!EXCEPTION_DESC.Value
''                   l_objGetActualTimeRecord!DISTRIBUTION_DESC.Value = l_strGetAccountDescription
''                Else    'use exception description default
''                End If
''              End If
''           End If
''       Else 'populate with the override description table if exist or default exception description 'l_strGetAccountDescription = from filter override desc
''           If l_objGetOverExccdAccountDescription Is Nothing Then
''           Else
''               l_strFilter = "EXCEPTION_CD= '" & l_objGetActualTimeRecord!OVER_EXC_CD.Value & " '"
''               l_objGetOverExccdAccountDescription.Filter = l_strFilter
''               If Not (l_objGetOverExccdAccountDescription.EOF And l_objGetOverExccdAccountDescription.BOF) Then
''                  l_strGetAccountDescription = l_objGetOverExccdAccountDescription!EXCEPTION_DESC.Value
''                  l_objGetActualTimeRecord!DISTRIBUTION_DESC.Value = l_strGetAccountDescription
''               Else     'use exception description default
''               End If
''           End If
''       End If
''
''       If Left$(l_objGetActualTimeRecord!START_TIME.Value, 2) = "99" Then
''           l_objGetActualTimeRecord!TIME_ENTRY_TYPE.Value = 1
''       Else
''           If l_objGetActualTimeRecord!APPR_IND.Value = "Y" Then
''                If Left$(l_objGetActualTimeRecord!EXCEPTION_BANK_CD.Value, 1) = "O" _
''                Or l_objGetActualTimeRecord!EXCEPTION_BANK_CD.Value = "C+" Then
''                    l_objGetActualTimeRecord!TIME_ENTRY_TYPE.Value = 2
''                Else
''                    If Trim(l_objGetActualTimeRecord!SCHED_CHG_EXC_CD.Value) <> "" Then
''                        l_objGetActualTimeRecord!TIME_ENTRY_TYPE.Value = 3
''                    Else
''                        l_objGetActualTimeRecord!TIME_ENTRY_TYPE.Value = 4
''                    End If
''                End If
''           Else
''                If Left$(l_objGetActualTimeRecord!EXCEPTION_BANK_CD.Value, 1) = "O" _
''                Or l_objGetActualTimeRecord!EXCEPTION_BANK_CD.Value = "C+" Then
''                   l_objGetActualTimeRecord!TIME_ENTRY_TYPE.Value = 5
''                Else
''                    If Trim(l_objGetActualTimeRecord!SCHED_CHG_EXC_CD.Value) <> "" Then
''                        l_objGetActualTimeRecord!TIME_ENTRY_TYPE.Value = 6
''                    Else
''                        l_objGetActualTimeRecord!TIME_ENTRY_TYPE.Value = 7
''                    End If
''                End If
''           End If
''       End If
''       l_objGetActualTimeRecord.Update
''
''       If v_ReturnOption = TrActualTimeEntriesDetail Then 'Add to Clone subset  for summary recordset request
''       Else 'building summary recordset from actual time recordset
''           l_objGetSubsetActualTimeRecord.AddNew
''           If v_blnMultipleEmployees Then
''                l_objGetSubsetActualTimeRecord("EMP_NUM") = l_objGetActualTimeRecord!EMP_NUM.Value
''           End If
''           l_objGetSubsetActualTimeRecord("DISTRIBUTION_ID") = l_objGetActualTimeRecord!DISTRIBUTION_ID.Value
''           l_objGetSubsetActualTimeRecord("EXCEPTION_CD") = l_objGetActualTimeRecord!EXCEPTION_CD.Value
''           l_objGetSubsetActualTimeRecord("OVER_EXC_CD") = l_objGetActualTimeRecord!OVER_EXC_CD.Value
''           l_objGetSubsetActualTimeRecord("CHARGE_DATE") = l_objGetActualTimeRecord!CHARGE_DATE.Value
''           l_objGetSubsetActualTimeRecord("HRS_PER_ACTIVITY") = l_objGetActualTimeRecord!Hours.Value
''           l_objGetSubsetActualTimeRecord("DISTRIBUTION_DESC") = l_objGetActualTimeRecord!DISTRIBUTION_DESC.Value
''           l_objGetSubsetActualTimeRecord("DISTRIBUTION_TYPE") = l_objGetActualTimeRecord!DISTRIBUTION_TYPE.Value
''           l_objGetSubsetActualTimeRecord.Update
''       End If       'Add to Clone subset  for summary recordset request
''       l_objGetActualTimeRecord.MoveNext
''    Loop
'
'        'end factoring
'
'    If v_ReturnOption = TrActualTimeEntriesDetail Then
'       l_objGetActualTimeRecord.MoveFirst
'       Set r_objRSActualTimeDetail = l_objGetActualTimeRecord
'       Set r_objRSTotalHoursByActivitySummary = Nothing
'       Exit Function
'    Else
'        If v_ReturnOption = TrActualTimeEntriesDetailAndTotalHoursByActivitySummary Then
'           l_objGetActualTimeRecord.MoveFirst
'           Set r_objRSActualTimeDetail = l_objGetActualTimeRecord
'        Else
'           Set r_objRSActualTimeDetail = Nothing
'       End If
'       Set r_objRSTotalHoursByActivitySummary = RetrieveTotalHoursByActivitySummary(l_objGetSubsetActualTimeRecord, v_blnMultipleEmployees)
'    End If
'    Exit Function
'
'eh:
'    '2/2/01 rsg
'    Set r_objRSActualTimeDetail = Nothing
'    Set r_objRSTotalHoursByActivitySummary = Nothing
'    TrsRaiseError m_strCMODULENAME, "RetrieveActualTimeEntriesForSummaryAndDetailForDept"
'End Function
'
'Private Function ApplyOverridesToActualTimeEntriesForSummaryAndDetail(ByVal l_objGetActualTimeRecord As ADODB.Recordset, _
'                                                                    ByVal l_objGetSubsetActualTimeRecord As ADODB.Recordset, _
'                                                                    ByVal v_ReturnOption As TRSReturnOption, _
'                                                                    ByVal v_blnMultipleEmployees As Boolean _
'                                                                )
'    Dim l_objGetAccountDescription As busDistributionMgr.cbsDistributionMgr
'    Dim l_objGetOverExccdDescription As dbTimeCard.cdTimeCard
'    Dim l_objGetWGDistributionIDDescriptions As busDistributionMgr.cbsDistributionMgr
'    Dim l_objRsPickListDistribution As ADODB.Recordset
'    Dim l_strGetAccountDescription As String
'    Dim l_objGetOverExccdAccountDescription As ADODB.Recordset
'    Dim l_colWG As New Collection
'    Dim l_colOverExccd As New Collection
'    Dim l_strWg As String
'    Dim l_strOEC As String
'    Dim l_strFilter As String
'
'    'got all the actual time scope with bank code now create collection of wg and exc_cd to get descriptions from distribution_id or override_code
'    Do Until l_objGetActualTimeRecord.EOF     '       process first pay period table (for splits) - overlay correct dep no effective date l_intSplitPeriod_flag = 1
'        l_strWg = Trim(str(l_objGetActualTimeRecord!WG_NUM.Value))
'        If Not IncludesString(l_strWg, l_colWG) Then
'            l_colWG.Add l_strWg
'        End If
'        l_strOEC = Trim(l_objGetActualTimeRecord!OVER_EXC_CD.Value)
'        If l_strOEC <> "" Then
'            If Not IncludesString(l_strOEC, l_colOverExccd) Then
'                l_colOverExccd.Add l_strOEC
'            End If
'        End If
'       l_objGetActualTimeRecord.MoveNext
'    Loop
'
'    Set l_objGetWGDistributionIDDescriptions = CreateObject("busDistributionMgr.cbsDistributionMgr")
'    Set l_objRsPickListDistribution = l_objGetWGDistributionIDDescriptions.GetWGDistributionIDDescriptions(l_colWG) 'get internal picklist table
'
'    If l_colOverExccd.Count < 1 Then 'if no override to process do not create override table
'    Else
'        Set l_objGetOverExccdDescription = CreateObject("dbTimeCard.cdTimeCard")
'        Set l_objGetOverExccdAccountDescription = l_objGetOverExccdDescription.GetAccountOverrideDescription(l_colOverExccd)
'    End If
'
'    l_objGetActualTimeRecord.MoveFirst
'
'    Do Until l_objGetActualTimeRecord.EOF
'       If Trim(l_objGetActualTimeRecord!DISTRIBUTION_TYPE.Value) <> "" Then
'           Set l_objGetAccountDescription = CreateObject("busDistributionMgr.cbsDistributionMgr")
'           l_strGetAccountDescription = l_objGetAccountDescription.GetAccountDescription(l_objRsPickListDistribution, l_objGetActualTimeRecord!WG_NUM.Value, l_objGetActualTimeRecord!DISTRIBUTION_ID.Value, l_objGetActualTimeRecord!DISTRIBUTION_TYPE.Value, l_objGetActualTimeRecord!DISTR_DEPT_NUM.Value)
'           If Trim(l_strGetAccountDescription) <> "" Then
'              l_objGetActualTimeRecord!DISTRIBUTION_DESC.Value = l_strGetAccountDescription
'           Else 'populate with the override description table if exist or default exception description 'l_strGetAccountDescription = from filter override desc
'              If l_objGetOverExccdAccountDescription Is Nothing Then
'              Else
'                l_strFilter = "EXCEPTION_CD= '" & l_objGetActualTimeRecord!OVER_EXC_CD.Value & " '"
'                l_objGetOverExccdAccountDescription.Filter = l_strFilter
'                If Not (l_objGetOverExccdAccountDescription.EOF And l_objGetOverExccdAccountDescription.BOF) Then
'                   l_strGetAccountDescription = l_objGetOverExccdAccountDescription!EXCEPTION_DESC.Value
'                   l_objGetActualTimeRecord!DISTRIBUTION_DESC.Value = l_strGetAccountDescription
'                Else    'use exception description default
'                End If
'              End If
'           End If
'       Else 'populate with the override description table if exist or default exception description 'l_strGetAccountDescription = from filter override desc
'           If l_objGetOverExccdAccountDescription Is Nothing Then
'           Else
'               l_strFilter = "EXCEPTION_CD= '" & l_objGetActualTimeRecord!OVER_EXC_CD.Value & " '"
'               l_objGetOverExccdAccountDescription.Filter = l_strFilter
'               If Not (l_objGetOverExccdAccountDescription.EOF And l_objGetOverExccdAccountDescription.BOF) Then
'                  l_strGetAccountDescription = l_objGetOverExccdAccountDescription!EXCEPTION_DESC.Value
'                  l_objGetActualTimeRecord!DISTRIBUTION_DESC.Value = l_strGetAccountDescription
'               Else     'use exception description default
'               End If
'           End If
'       End If
'
'       If Left$(l_objGetActualTimeRecord!START_TIME.Value, 2) = "99" Then
'           l_objGetActualTimeRecord!TIME_ENTRY_TYPE.Value = 1
'       Else
'           If l_objGetActualTimeRecord!APPR_IND.Value = "Y" Then
'                If Left$(l_objGetActualTimeRecord!EXCEPTION_BANK_CD.Value, 1) = "O" _
'                Or l_objGetActualTimeRecord!EXCEPTION_BANK_CD.Value = "C+" Then
'                    l_objGetActualTimeRecord!TIME_ENTRY_TYPE.Value = 2
'                Else
'                    If Trim(l_objGetActualTimeRecord!SCHED_CHG_EXC_CD.Value) <> "" Then
'                        l_objGetActualTimeRecord!TIME_ENTRY_TYPE.Value = 3
'                    Else
'                        l_objGetActualTimeRecord!TIME_ENTRY_TYPE.Value = 4
'                    End If
'                End If
'           Else
'                If Left$(l_objGetActualTimeRecord!EXCEPTION_BANK_CD.Value, 1) = "O" _
'                Or l_objGetActualTimeRecord!EXCEPTION_BANK_CD.Value = "C+" Then
'                   l_objGetActualTimeRecord!TIME_ENTRY_TYPE.Value = 5
'                Else
'                    If Trim(l_objGetActualTimeRecord!SCHED_CHG_EXC_CD.Value) <> "" Then
'                        l_objGetActualTimeRecord!TIME_ENTRY_TYPE.Value = 6
'                    Else
'                        l_objGetActualTimeRecord!TIME_ENTRY_TYPE.Value = 7
'                    End If
'                End If
'           End If
'       End If
'       l_objGetActualTimeRecord.Update
'
'       If v_ReturnOption = TrActualTimeEntriesDetail Then 'Add to Clone subset  for summary recordset request
'       Else 'building summary recordset from actual time recordset
'           l_objGetSubsetActualTimeRecord.AddNew
'           If v_blnMultipleEmployees Then
'                l_objGetSubsetActualTimeRecord("EMP_NUM") = l_objGetActualTimeRecord!EMP_NUM.Value
'           End If
'           l_objGetSubsetActualTimeRecord("DISTRIBUTION_ID") = l_objGetActualTimeRecord!DISTRIBUTION_ID.Value
'           l_objGetSubsetActualTimeRecord("EXCEPTION_CD") = l_objGetActualTimeRecord!EXCEPTION_CD.Value
'           l_objGetSubsetActualTimeRecord("OVER_EXC_CD") = l_objGetActualTimeRecord!OVER_EXC_CD.Value
'           l_objGetSubsetActualTimeRecord("CHARGE_DATE") = l_objGetActualTimeRecord!CHARGE_DATE.Value
'           l_objGetSubsetActualTimeRecord("HRS_PER_ACTIVITY") = l_objGetActualTimeRecord!Hours.Value
'           l_objGetSubsetActualTimeRecord("DISTRIBUTION_DESC") = l_objGetActualTimeRecord!DISTRIBUTION_DESC.Value
'           l_objGetSubsetActualTimeRecord("DISTRIBUTION_TYPE") = l_objGetActualTimeRecord!DISTRIBUTION_TYPE.Value
'           l_objGetSubsetActualTimeRecord.Update
'       End If       'Add to Clone subset  for summary recordset request
'       l_objGetActualTimeRecord.MoveNext
'    Loop
'
'End Function
'
'//-----
Private Function RetrieveTotalHoursByActivitySummary( _
        v_objGetSubsetActualTimeRecord As ADODB.Recordset, _
        ByVal v_blnMulipleEmployees As Boolean, _
        Optional ByVal v_intDescriptionLength As Integer = 39) As ADODB.Recordset
    
    Dim l_objRSTotalHoursByActivity As New ADODB.Recordset
    Dim l_strDistributionId As String
    Dim l_strExceptionCode As String
    Dim l_strOverExcCode As String
    Dim l_datChargeDate As Date
    Dim l_varHoursPerActivity As Variant
    Dim l_strDistributionDescription As String
    Dim l_strDistributionType As String
    Dim l_strJobTask As String
    Dim l_intStep As Integer
    Dim l_strEMPRVDBid As String
    Dim l_lngEmpNum As Long
    Dim l_lngCurrEmpNum As Long
    
    
    On Error GoTo RetrieveTotalHoursByActivitySummary_EH

    v_objGetSubsetActualTimeRecord.MoveFirst

    v_objGetSubsetActualTimeRecord.Sort = "DISTRIBUTION_ID ASC, JOB_TASK ASC, JOBTASK_STEP_NBR ASC, EMPRV_DATABASE_ID ASC, EXCEPTION_CD ASC, OVER_EXC_CD ASC, CHARGE_DATE ASC, DISTRIBUTION_TYPE ASC"
    
    If v_blnMulipleEmployees Then
        l_objRSTotalHoursByActivity.Fields.Append "EMP_NUM", adInteger
    End If
    
    l_objRSTotalHoursByActivity.Fields.Append "DISTRIBUTION_ID", adChar, 39
    l_objRSTotalHoursByActivity.Fields.Append "EXCEPTION_CD", adChar, 3
    l_objRSTotalHoursByActivity.Fields.Append "OVER_EXC_CD", adChar, 3
    l_objRSTotalHoursByActivity.Fields.Append "CHARGE_DATE", adDate
    l_objRSTotalHoursByActivity.Fields.Append "HRS_PER_ACTIVITY", adDecimal
    l_objRSTotalHoursByActivity.Fields.Item("HRS_PER_ACTIVITY").Precision = 3
    l_objRSTotalHoursByActivity.Fields.Item("HRS_PER_ACTIVITY").NumericScale = 1
    l_objRSTotalHoursByActivity.Fields.Append "DISTRIBUTION_DESC", adChar, v_intDescriptionLength
    l_objRSTotalHoursByActivity.Fields.Append "DISTRIBUTION_TYPE", adChar, 1
    l_objRSTotalHoursByActivity.Fields.Append "JOB_TASK", adChar, 15
    l_objRSTotalHoursByActivity.Fields.Append "JOBTASK_STEP_NBR", adInteger
    l_objRSTotalHoursByActivity.Fields.Append "EMPRV_DATABASE_ID", adChar, 8
    
    l_objRSTotalHoursByActivity.Open
    
    If v_blnMulipleEmployees Then
        l_lngEmpNum = v_objGetSubsetActualTimeRecord!EMP_NUM.Value
        l_lngCurrEmpNum = l_lngEmpNum
    Else
        l_lngEmpNum = 0
        l_lngCurrEmpNum = 0
    End If
    
    l_strDistributionId = v_objGetSubsetActualTimeRecord!DISTRIBUTION_ID.Value
    l_strExceptionCode = v_objGetSubsetActualTimeRecord!EXCEPTION_CD.Value
    l_datChargeDate = v_objGetSubsetActualTimeRecord!CHARGE_DATE.Value
    l_strOverExcCode = v_objGetSubsetActualTimeRecord!OVER_EXC_CD.Value
    l_varHoursPerActivity = CDec(v_objGetSubsetActualTimeRecord!HRS_PER_ACTIVITY.Value)
    l_strDistributionDescription = PadRight(v_objGetSubsetActualTimeRecord!DISTRIBUTION_DESC.Value, v_intDescriptionLength)
    l_strDistributionType = v_objGetSubsetActualTimeRecord!DISTRIBUTION_TYPE.Value
    l_strJobTask = v_objGetSubsetActualTimeRecord!JOB_TASK.Value
    l_intStep = v_objGetSubsetActualTimeRecord!JOBTASK_STEP_NBR.Value
    l_strEMPRVDBid = v_objGetSubsetActualTimeRecord!EMPRV_DATABASE_ID.Value
    
    v_objGetSubsetActualTimeRecord.MoveNext
    
    If Not (v_objGetSubsetActualTimeRecord.EOF And v_objGetSubsetActualTimeRecord.BOF) Then
    Else 'only 1 record to write because eof
       l_objRSTotalHoursByActivity.AddNew
       If v_blnMulipleEmployees Then
        l_objRSTotalHoursByActivity("EMP_NUM") = l_lngEmpNum
       End If
       l_objRSTotalHoursByActivity("DISTRIBUTION_ID") = l_strDistributionId
       l_objRSTotalHoursByActivity("EXCEPTION_CD") = l_strExceptionCode
       l_objRSTotalHoursByActivity("OVER_EXC_CD") = l_strOverExcCode
       l_objRSTotalHoursByActivity("CHARGE_DATE") = l_datChargeDate
       l_objRSTotalHoursByActivity("HRS_PER_ACTIVITY") = l_varHoursPerActivity
       l_objRSTotalHoursByActivity("DISTRIBUTION_DESC") = PadRight(l_strDistributionDescription, v_intDescriptionLength)
       l_objRSTotalHoursByActivity("DISTRIBUTION_TYPE") = l_strDistributionType
       l_objRSTotalHoursByActivity("JOB_TASK") = l_strJobTask
       l_objRSTotalHoursByActivity("JOBTASK_STEP_NBR") = l_intStep
       l_objRSTotalHoursByActivity("EMPRV_DATABASE_ID") = l_strEMPRVDBid
       l_objRSTotalHoursByActivity.Update
       Set RetrieveTotalHoursByActivitySummary = l_objRSTotalHoursByActivity
       Exit Function
    End If
    
    Do Until v_objGetSubsetActualTimeRecord.EOF
       If v_blnMulipleEmployees Then
             l_lngCurrEmpNum = v_objGetSubsetActualTimeRecord!EMP_NUM.Value
       End If
       If l_strDistributionId = v_objGetSubsetActualTimeRecord!DISTRIBUTION_ID.Value _
          And l_strExceptionCode = v_objGetSubsetActualTimeRecord!EXCEPTION_CD.Value _
          And l_strOverExcCode = v_objGetSubsetActualTimeRecord!OVER_EXC_CD.Value _
          And l_datChargeDate = v_objGetSubsetActualTimeRecord!CHARGE_DATE.Value _
          And l_strDistributionType = v_objGetSubsetActualTimeRecord!DISTRIBUTION_TYPE.Value _
          And l_strJobTask = v_objGetSubsetActualTimeRecord!JOB_TASK _
          And l_intStep = v_objGetSubsetActualTimeRecord!JOBTASK_STEP_NBR _
          And l_strEMPRVDBid = v_objGetSubsetActualTimeRecord!EMPRV_DATABASE_ID _
          And l_lngCurrEmpNum = l_lngEmpNum Then
'unique key is charge_date and distribution description to accumulate summary hours
          l_varHoursPerActivity = l_varHoursPerActivity + v_objGetSubsetActualTimeRecord!HRS_PER_ACTIVITY.Value               'add hours for 2nd
       Else
          l_objRSTotalHoursByActivity.AddNew   '1st
          If v_blnMulipleEmployees Then
             l_objRSTotalHoursByActivity("EMP_NUM") = l_lngEmpNum
          End If
          l_objRSTotalHoursByActivity("DISTRIBUTION_ID") = l_strDistributionId
          l_objRSTotalHoursByActivity("EXCEPTION_CD") = l_strExceptionCode
          l_objRSTotalHoursByActivity("OVER_EXC_CD") = l_strOverExcCode
          l_objRSTotalHoursByActivity("CHARGE_DATE") = l_datChargeDate
          l_objRSTotalHoursByActivity("HRS_PER_ACTIVITY") = l_varHoursPerActivity
          l_objRSTotalHoursByActivity("DISTRIBUTION_DESC") = PadRight(l_strDistributionDescription, v_intDescriptionLength)
          l_objRSTotalHoursByActivity("DISTRIBUTION_TYPE") = l_strDistributionType
          l_objRSTotalHoursByActivity("JOB_TASK") = l_strJobTask
          l_objRSTotalHoursByActivity("JOBTASK_STEP_NBR") = l_intStep
          l_objRSTotalHoursByActivity("EMPRV_DATABASE_ID") = l_strEMPRVDBid
          l_objRSTotalHoursByActivity.Update
          'write if break and populate new stuff in keys(charge_date and description) and fields
          If v_blnMulipleEmployees Then
             l_lngEmpNum = v_objGetSubsetActualTimeRecord!EMP_NUM.Value
          End If
          l_strDistributionId = v_objGetSubsetActualTimeRecord!DISTRIBUTION_ID.Value
          l_strExceptionCode = v_objGetSubsetActualTimeRecord!EXCEPTION_CD.Value
          l_strOverExcCode = v_objGetSubsetActualTimeRecord!OVER_EXC_CD.Value
          l_datChargeDate = v_objGetSubsetActualTimeRecord!CHARGE_DATE.Value
          l_varHoursPerActivity = v_objGetSubsetActualTimeRecord!HRS_PER_ACTIVITY.Value
          l_strDistributionDescription = PadRight(v_objGetSubsetActualTimeRecord!DISTRIBUTION_DESC.Value, v_intDescriptionLength) 'initialize all
          l_strDistributionType = v_objGetSubsetActualTimeRecord!DISTRIBUTION_TYPE.Value
          l_strJobTask = v_objGetSubsetActualTimeRecord!JOB_TASK.Value
          l_intStep = v_objGetSubsetActualTimeRecord!JOBTASK_STEP_NBR.Value
          l_strEMPRVDBid = v_objGetSubsetActualTimeRecord!EMPRV_DATABASE_ID.Value
       End If
       v_objGetSubsetActualTimeRecord.MoveNext
    Loop

    l_objRSTotalHoursByActivity.AddNew  '1st or 2nds
    If v_blnMulipleEmployees Then
       l_objRSTotalHoursByActivity("EMP_NUM") = l_lngEmpNum
    End If
    l_objRSTotalHoursByActivity("DISTRIBUTION_ID") = l_strDistributionId
    l_objRSTotalHoursByActivity("EXCEPTION_CD") = l_strExceptionCode
    l_objRSTotalHoursByActivity("OVER_EXC_CD") = l_strOverExcCode
    l_objRSTotalHoursByActivity("CHARGE_DATE") = l_datChargeDate
    l_objRSTotalHoursByActivity("HRS_PER_ACTIVITY") = l_varHoursPerActivity
    l_objRSTotalHoursByActivity("DISTRIBUTION_DESC") = PadRight(l_strDistributionDescription, v_intDescriptionLength)
    l_objRSTotalHoursByActivity("DISTRIBUTION_TYPE") = l_strDistributionType
    l_objRSTotalHoursByActivity("JOB_TASK") = l_strJobTask
    l_objRSTotalHoursByActivity("JOBTASK_STEP_NBR") = l_intStep
    l_objRSTotalHoursByActivity("EMPRV_DATABASE_ID") = l_strEMPRVDBid
    l_objRSTotalHoursByActivity.Update
    
    l_objRSTotalHoursByActivity.MoveFirst
    
    Set RetrieveTotalHoursByActivitySummary = l_objRSTotalHoursByActivity
    
    Exit Function

RetrieveTotalHoursByActivitySummary_EH:
    TrsRaiseError m_strCMODULENAME, "RetrieveTotalHoursByActivitySummary"
End Function

'//for older TRIS clients
Public Function GetPriorPeriodDistributionInformation(ByVal v_lngEmpNum, _
        ByVal v_datStartDate, ByVal v_datEndDate) As ADODB.Recordset
On Error GoTo eh

    Set GetPriorPeriodDistributionInformation = GetPriorPeriodDistributionInformationPrivate(v_lngEmpNum, v_datStartDate, v_datEndDate)
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "GetPriorPeriodDistributionInformation"
End Function

'//for newer TRIS clients (OTTC)
Public Function GetPriorPeriodDistributionInformation2(ByVal v_lngEmpNum, _
        ByVal v_datStartDate, ByVal v_datEndDate) As ADODB.Recordset
On Error GoTo eh

    Set GetPriorPeriodDistributionInformation2 = GetPriorPeriodDistributionInformationPrivate(v_lngEmpNum, v_datStartDate, v_datEndDate, 60)
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "GetPriorPeriodDistributionInformation2"
End Function

'//for newer TRIS clients DFE 48 chars, OTTC pick list description increased to 60 chars
Public Function GetPriorPeriodDistributionInformation2B(ByVal v_lngEmpNum As Long, _
                                                        ByVal v_datStartDate As Date, _
                                                        ByVal v_datEndDate As Date, _
                                                        ByVal v_intDescriptionLength As Integer) As ADODB.Recordset
On Error GoTo eh

    Set GetPriorPeriodDistributionInformation2B = GetPriorPeriodDistributionInformationPrivate(v_lngEmpNum, v_datStartDate, v_datEndDate, v_intDescriptionLength)
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "GetPriorPeriodDistributionInformation2B"
End Function

Private Function GetPriorPeriodDistributionInformationPrivate(ByVal v_lngEmpNum, _
        ByVal v_datStartDate, ByVal v_datEndDate, _
        Optional ByVal v_intDescriptionLength As Integer = 39) As ADODB.Recordset
    
    Dim l_strGetAccountDescription As String
    'Dim l_colWG As New Collection
    Dim l_colWG As utilTRISCommonLib.cuSet
    Dim l_colOverExccd As New Collection
    Dim l_strFilter As String
    Dim l_objGetOverExccdAccountDescription As ADODB.Recordset
    Dim l_objGetActualTimeRecord As ADODB.Recordset
    Dim l_objRsPickListDistribution As ADODB.Recordset
    Dim l_objRSDistributionIDsTable As New ADODB.Recordset
    Dim l_objRSEmployeefields As New ADODB.Recordset
    Dim l_objcdTimecard As dbTimeCard.cdTimeCard
    Dim l_objcbsDistributionMgr As busDistributionMgr.cbsDistributionMgr
    'Dim l_objcbsDistributionMgr As Object
    Dim l_objRSEmployee As busEmployee.cbsEmployee
    Dim l_strCmpyCd As String
    Dim l_strResponse As String
    Dim l_strJobTaskMaximo  As String
    Dim l_strJobNumMaximo As String
    Dim l_strTaskNumMaximo As String
    Dim l_objRsJobTaskDistributionMaximo As ADODB.Recordset
    Dim comint4 As Ameren_TRIS_Server_COMIntegration4.ValidateDistribution
    Dim rc As Boolean
    Dim blnDistributionIsValid As Boolean
    
    On Error GoTo eh
    Set l_objcdTimecard = CreateObject("dbTimeCard.cdTimeCard")
    Set l_objGetActualTimeRecord = l_objcdTimecard.GetActualTimeEntries2(v_lngEmpNum, v_datStartDate, v_datEndDate, v_intDescriptionLength)
    Set l_objRsJobTaskDistributionMaximo = Nothing

    'got all the actual time scope with bank code now create collection of wg and exc_cd to get descriptions from distribution_id or override_code
    Set l_objRSEmployee = CreateObject("busEmployee.cbsEmployee")
    Set l_objRSEmployeefields = l_objRSEmployee.GetEmployee(v_lngEmpNum)
    Set l_objRSEmployee = Nothing
    
    If Not (l_objGetActualTimeRecord.EOF And l_objGetActualTimeRecord.BOF) Then
    Else
       Set GetPriorPeriodDistributionInformationPrivate = l_objGetActualTimeRecord
       Exit Function
    End If
    
    Set l_colWG = New utilTRISCommonLib.cuSet
    Do Until l_objGetActualTimeRecord.EOF     '       process first pay period table (for splits) - overlay correct dep no effective date l_intSplitPeriod_flag = 1
       l_colWG.Add l_objGetActualTimeRecord!WG_NUM.Value
       l_colOverExccd.Add l_objGetActualTimeRecord!OVER_EXC_CD.Value
       l_objGetActualTimeRecord.MoveNext
    Loop
    
    l_objGetActualTimeRecord.MoveFirst
    
    Set l_objcbsDistributionMgr = CreateObject("busDistributionMgr.cbsDistributionMgr")
    Set l_objRsPickListDistribution = l_objcbsDistributionMgr.GetWGDistributionIDDescriptions3(l_colWG.Contents, v_intDescriptionLength, "LoadPriorPayPeriodDesc") 'get internal picklist table
    Set l_objGetOverExccdAccountDescription = l_objcdTimecard.GetAccountOverrideDescription(l_colOverExccd)
    Set l_objcdTimecard = Nothing
    
    l_objRSDistributionIDsTable.Fields.Append "WG_NUM", adInteger
    l_objRSDistributionIDsTable.Fields.Append "DISTRIBUTION_ID", adChar, 39
    l_objRSDistributionIDsTable.Fields.Append "DISTRIBUTION_TYPE", adChar, 1
    l_objRSDistributionIDsTable.Fields.Append "DISTRIBUTION_DESC", adChar, v_intDescriptionLength
    l_objRSDistributionIDsTable.Fields.Append "JOB_TASK", adVarChar, 15
    l_objRSDistributionIDsTable.Fields.Append "JOBTASK_STEP_NBR", adInteger
    l_objRSDistributionIDsTable.Fields.Append "EMPRV_DATABASE_ID", adVarChar, 8
    
    l_objRSDistributionIDsTable.Open
    
    Do Until l_objGetActualTimeRecord.EOF
'        If l_objcbsDistributionMgr.ValidateDistributionId( _
'                l_objGetActualTimeRecord!DISTRIBUTION_ID.Value, _
'                l_strResponse, _
'                l_objRSEmployeefields!CMPY_CD.Value, _
'                l_objGetActualTimeRecord!DISTRIBUTION_TYPE.Value) Then
'
'JIMB made dummy comparison so it will go through
'will need to call busDistributionMgr to validate
        blnDistributionIsValid = True
        If 1 = 1 Then
           l_strFilter = "WG_NUM= " & l_objGetActualTimeRecord!WG_NUM.Value & " AND DISTRIBUTION_ID= '" & l_objGetActualTimeRecord!DISTRIBUTION_ID.Value & " '"
           l_strFilter = MDAC25FilterFix(l_strFilter)
           l_objRSDistributionIDsTable.Filter = "" 'initialize filter
           l_objRSDistributionIDsTable.Filter = l_strFilter 'find duplicates using filter
           If Not (l_objRSDistributionIDsTable.EOF And l_objRSDistributionIDsTable.BOF) Then 'duplicates are bypassed
           Else
              If l_objGetActualTimeRecord!EXCEPTION_CD.Value = "ST " _
              Or Left$(l_objGetActualTimeRecord!EXCEPTION_CD.Value, 2) = "OT" Then
                If Trim(l_objGetActualTimeRecord!DISTRIBUTION_TYPE.Value) <> "" Then
                    Set l_objcbsDistributionMgr = CreateObject("busDistributionMgr.cbsDistributionMgr")
                    If Trim(l_objGetActualTimeRecord!DISTRIBUTION_TYPE.Value) = "J" Then
                        l_strGetAccountDescription = l_objcbsDistributionMgr.GetAccountDescriptionJobTask2(l_objRsPickListDistribution, l_objGetActualTimeRecord!WG_NUM.Value, l_objGetActualTimeRecord!DISTRIBUTION_ID.Value, l_objGetActualTimeRecord!DISTRIBUTION_TYPE.Value, Mid(l_objGetActualTimeRecord!JOB_TASK, 1, 10), l_objGetActualTimeRecord!EMPRV_DATABASE_ID, l_objGetActualTimeRecord!DISTR_DEPT_NUM.Value, v_intDescriptionLength)
                    ElseIf Trim(l_objGetActualTimeRecord!DISTRIBUTION_TYPE.Value) = "M" Then
                        l_strJobTaskMaximo = Trim(l_objGetActualTimeRecord!JOB_TASK)
                        l_strJobNumMaximo = Left(l_strJobTaskMaximo, 7)
                        l_strTaskNumMaximo = Right(l_strJobTaskMaximo, 3)

                        l_strGetAccountDescription = GetAccountDescriptionJobTaskMaximo(l_objRsJobTaskDistributionMaximo, l_strJobNumMaximo, l_strTaskNumMaximo, l_objGetActualTimeRecord!EMPRV_DATABASE_ID, v_intDescriptionLength)
                    Else
                        l_strGetAccountDescription = l_objcbsDistributionMgr.GetAccountDescription2(l_objRsPickListDistribution, l_objGetActualTimeRecord!WG_NUM.Value, l_objGetActualTimeRecord!DISTRIBUTION_ID.Value, l_objGetActualTimeRecord!DISTRIBUTION_TYPE.Value, l_objGetActualTimeRecord!DISTR_DEPT_NUM.Value, v_intDescriptionLength)
                        '//if accounting is no longer valid due to labor flag, active status, project finish date, etc,
                        '//then it will not be on the picklist (l_objRsPickListDistribution).
                        '//Then, GetAccountDescription2 will not be able to find the description for the actual time entry
                        '//so the desired effect is to NOT include the entry for "load prior" (GetPriorPeriodDistribution)
                        If Trim(l_strGetAccountDescription) = "" Then
                            blnDistributionIsValid = False
                        End If
                    End If
                    If Trim(l_strGetAccountDescription) <> "" Then
                       l_objGetActualTimeRecord!DISTRIBUTION_DESC.Value = PadRight(l_strGetAccountDescription, v_intDescriptionLength)
                    Else
                        'populate with the override description or default exception description
                        'l_strGetAccountDescription = from filter override desc
                        l_strFilter = "EXCEPTION_CD= '" & l_objGetActualTimeRecord!OVER_EXC_CD.Value & " '"
                        l_objGetOverExccdAccountDescription.Filter = l_strFilter
                        If Not (l_objGetOverExccdAccountDescription.EOF And l_objGetOverExccdAccountDescription.BOF) Then
                           l_strGetAccountDescription = PadRight(l_objGetOverExccdAccountDescription!EXCEPTION_DESC.Value, v_intDescriptionLength)
                           l_objGetActualTimeRecord!DISTRIBUTION_DESC.Value = l_strGetAccountDescription
                        Else
                           'leave for exception  (EXCEPTION_CD = 'ST' OR EXCEPTION_CD LIKE 'OT%')
                        End If
                    End If
                Else '//DISTRIBUTION_TYPE = ""
                    'populate with the override description or default exception description
                    'l_strGetAccountDescription = from filter override desc
                    l_strFilter = "EXCEPTION_CD= '" & l_objGetActualTimeRecord!OVER_EXC_CD.Value & " '"
                    l_objGetOverExccdAccountDescription.Filter = l_strFilter
                    If Not (l_objGetOverExccdAccountDescription.EOF And l_objGetOverExccdAccountDescription.BOF) Then
                       l_strGetAccountDescription = PadRight(l_objGetOverExccdAccountDescription!EXCEPTION_DESC.Value, v_intDescriptionLength)
                       l_objGetActualTimeRecord!DISTRIBUTION_DESC.Value = l_strGetAccountDescription
                    Else
                        'leave for exception
                    End If
                End If

                l_objGetActualTimeRecord.Update
                If blnDistributionIsValid Then
                    l_objRSDistributionIDsTable.AddNew
                    l_objRSDistributionIDsTable(0) = l_objGetActualTimeRecord!WG_NUM.Value
                    l_objRSDistributionIDsTable(1) = l_objGetActualTimeRecord!DISTRIBUTION_ID.Value
                    l_objRSDistributionIDsTable(2) = l_objGetActualTimeRecord!DISTRIBUTION_TYPE.Value
                    l_objRSDistributionIDsTable(3) = PadRight(l_objGetActualTimeRecord!DISTRIBUTION_DESC.Value, v_intDescriptionLength)
                    l_objRSDistributionIDsTable(4) = l_objGetActualTimeRecord!JOB_TASK.Value
                    l_objRSDistributionIDsTable(5) = l_objGetActualTimeRecord!JOBTASK_STEP_NBR
                    l_objRSDistributionIDsTable(6) = l_objGetActualTimeRecord!EMPRV_DATABASE_ID
                    l_objRSDistributionIDsTable.Update
                End If
              Else
                  'skip if not equal to "ST" or like "OT%"
              End If
           End If
        Else
                  'skip if not valid distribution Id
        End If
       l_objGetActualTimeRecord.MoveNext
    Loop
       
    Set l_objcbsDistributionMgr = Nothing
    
    If Not (l_objRSDistributionIDsTable.EOF And l_objRSDistributionIDsTable.BOF) Then
       l_objRSDistributionIDsTable.MoveFirst
    End If
    
    l_objRSDistributionIDsTable.Filter = "" 'initialize filter
   
    '//for OTTC cutover
    
    If IsOTTCEffective() And IsOTTCTransitionActive() Then
        l_strCmpyCd = "AMS"
        If RecordsExist(l_objRSEmployeefields) Then
            l_strCmpyCd = l_objRSEmployeefields!CMPY_CD.Value
        End If
        
        Set comint4 = New Ameren_TRIS_Server_COMIntegration4.ValidateDistribution
        
        rc = comint4.ConvertPriorPayPeriod_(l_objRSDistributionIDsTable, l_strCmpyCd)
        Set comint4 = Nothing
    End If

    Set GetPriorPeriodDistributionInformationPrivate = l_objRSDistributionIDsTable
    
    'after getting the vanilla table process it in business getaccountdescription
    'use the description table to populate the actual time scope with the description
    'return complete actual time scope
    
    Exit Function
eh:
    Set comint4 = Nothing
    TrsRaiseError m_strCMODULENAME, "GetPriorPeriodDistributionInformationPrivate"
End Function

'//allow the code to be deployed ahead of time and QA OTTC code to deploy to Production while still using GL Account Block
Private Function IsPOETActive() As Boolean
    Dim dbEnv As dbEnvironment.cdEnvironment
    Dim UseOTTC As String
    
    On Error GoTo eh
        
    Set dbEnv = CreateObject("dbEnvironment.cdEnvironment")
    
    IsPOETActive = False
    
    UseOTTC = UCase(dbEnv.GetSysParmsValue(50, 1))
    
    If UseOTTC = "Y" Or UseOTTC = "TRUE" Or UseOTTC = "1" Then
        IsPOETActive = True
    End If
        
    Set dbEnv = Nothing

    Exit Function
eh:
    Set dbEnv = Nothing
   
    TrsRaiseError m_strCMODULENAME, "IsPOETActive", , , , DescArgsFrom(UseOTTC), True
End Function

Private Function IsOTTCCutoverOutageActive() As Boolean
    Dim dbEnv As dbEnvironment.cdEnvironment
    Dim OTTCCutoverOutage As String
    
    On Error GoTo eh
        
    Set dbEnv = CreateObject("dbEnvironment.cdEnvironment")
    
    IsOTTCCutoverOutageActive = False
    
    OTTCCutoverOutage = UCase(dbEnv.GetSysParmsValue(52, 1))
    
    If OTTCCutoverOutage = "Y" Or OTTCCutoverOutage = "TRUE" Or OTTCCutoverOutage = "1" Then
        IsOTTCCutoverOutageActive = True
    End If
        
    Set dbEnv = Nothing

    Exit Function

eh:
    Set dbEnv = Nothing
   
    TrsRaiseError m_strCMODULENAME, "IsOTTCCutoverOutageActive", , , , DescArgsFrom(OTTCCutoverOutage), True
End Function

'//actual date to start using POET distributions after cutover
Private Function IsOTTCEffective() As Boolean
    Dim dbEnv As dbEnvironment.cdEnvironment
    Dim sEffDate As String
    Dim EffDateTime
    
    On Error GoTo eh
        
    Set dbEnv = CreateObject("dbEnvironment.cdEnvironment")
    
    IsOTTCEffective = False
    
    sEffDate = dbEnv.GetSysParmsValue(53, 1)
    
    EffDateTime = DateTimeFromDB2Timestamp(sEffDate)
    
    IsOTTCEffective = Now > EffDateTime
        
    Set dbEnv = Nothing

    Exit Function
eh:
    Set dbEnv = Nothing
   
    TrsRaiseError m_strCMODULENAME, "IsOTTCEffective", , , , DescArgsFrom(sEffDate), True
End Function

Private Function IsOTTCTransitionActive() As Boolean
    Dim dbEnv As dbEnvironment.cdEnvironment
    Dim OTTCTransition As String
    
    On Error GoTo eh
        
    Set dbEnv = CreateObject("dbEnvironment.cdEnvironment")
    
    IsOTTCTransitionActive = False
    
    OTTCTransition = UCase(dbEnv.GetSysParmsValue(54, 1))
    
    If OTTCTransition = "Y" Or OTTCTransition = "TRUE" Or OTTCTransition = "1" Then
        IsOTTCTransitionActive = True
    End If
        
    Set dbEnv = Nothing

    Exit Function
eh:
    Set dbEnv = Nothing
   
    TrsRaiseError m_strCMODULENAME, "IsOTTCTransitionActive", , , , DescArgsFrom(OTTCTransition), True
End Function

Private Function IsOTTCBypassAccountValidationActive() As Boolean
    Dim dbEnv As dbEnvironment.cdEnvironment
    Dim OTTCBypassAccountValidation As String
    
    On Error GoTo eh
        
    Set dbEnv = CreateObject("dbEnvironment.cdEnvironment")
    
    IsOTTCBypassAccountValidationActive = False
    
    OTTCBypassAccountValidation = UCase(dbEnv.GetSysParmsValue(55, 1))
    
    If OTTCBypassAccountValidation = "Y" Or OTTCBypassAccountValidation = "TRUE" Or OTTCBypassAccountValidation = "1" Then
        IsOTTCBypassAccountValidationActive = True
    End If
        
    Set dbEnv = Nothing

    Exit Function
eh:
    Set dbEnv = Nothing
   
    TrsRaiseError m_strCMODULENAME, "IsOTTCBypassAccountValidation", , , , DescArgsFrom(OTTCBypassAccountValidation), True
End Function

Public Function RecordTimeEntry( _
            ByVal v_rsTimeEntries As ADODB.Recordset, _
            ByVal v_rsOtherPay As ADODB.Recordset, _
            ByVal v_rsECDrawTime As ADODB.Recordset, _
            ByVal r_rsEmployees As ADODB.Recordset, _
            ByVal r_rsEmpWorkgroups As ADODB.Recordset, _
            ByRef r_rsTimecardEntries As ADODB.Recordset, _
            ByRef r_rsOtherpayEntries As ADODB.Recordset, _
            ByVal r_rsSchedules As ADODB.Recordset, _
            ByVal v_rsScheduleChgPmts As ADODB.Recordset, _
            ByVal v_startDate As Date, _
            ByVal v_endDate As Date, _
            ByRef r_rsAccrualBankUpdates As ADODB.Recordset, _
            ByRef r_rsOvertimeBankUpdates As ADODB.Recordset, _
            ByRef r_colEMPRVInterface As Collection, _
            ByRef r_rsErrors As ADODB.Recordset, _
            ByVal v_UserId As String) As Boolean
TrsRaiseError m_strCMODULENAME, "RecordTimeEntry", trErrObsoleteMethod, , , , True

'    Dim objCtx As ObjectContext
'    Dim l_lngEmployeeNumber As Long
'    Dim l_datChargeDate As Date
'    Dim l_blnResult As Boolean
'    Dim l_rsAccrualBankUpdates As ADODB.Recordset
'    Dim l_rsOvertimeBankUpdates As ADODB.Recordset
'    Dim l_colEMPRVInterface As Collection
'    Dim l_intI As Integer
'    Dim l_intCount As Integer
'
'On Error GoTo eh
'
''test only
'    #If TRACENEON Then
'        Dim lngTraceEmpNum As Long
'        lngTraceEmpNum = v_rsTimeEntries!KEY_EMP_NUM
'        TraceNeonMessage CStr(lngTraceEmpNum), "RECDTE", "BEGIN"
'    #End If
'
'    #If TRTRACE Then
'        TraceMessage "", "Begin RecordTimeEntry"
'    #End If
'
'    Set objCtx = GetObjectContext
'
'    m_strUserId = v_UserId
'    Set m_objErrors = r_rsErrors
'
'    If Not GetKeyEmployeeData( _
'                v_rsTimeEntries, _
'                v_rsOtherPay, _
'                l_lngEmployeeNumber, _
'                l_datChargeDate) Then
'        If RecordsExist(v_rsECDrawTime) Then
'            GoTo ecdraw
'        Else
'            RecordTimeEntry = False
'            Exit Function
'        End If
'    End If
'
'    m_datCollPerStart = v_startDate
'    m_datCollPerEnd = v_endDate
'
'    #If TRTRACE Then
'        TraceMessage "", "Begin RestoreTimecard"
'    #End If
'    RestoreTimecard r_rsTimecardEntries
'    #If TRTRACE Then
'        TraceMessage "", "End RestoreTimecard"
'    #End If
''    RestoreTimecard _
''        CLng(l_lngEmployeeNumber), m_datCollPerStart, m_datCollPerEnd
'    #If TRTRACE Then
'        TraceMessage "", "Begin RestoreEmployee"
'    #End If
'    RestoreEmployee r_rsEmployees, r_rsEmpWorkgroups
'    #If TRTRACE Then
'        TraceMessage "", "End RestoreEmployee"
'    #End If
''    RestoreEmployee _
''        CLng(l_lngEmployeeNumber), m_datCollPerStart, m_datCollPerEnd
'    #If TRTRACE Then
'        TraceMessage "", "Begin RestoreSchedule"
'    #End If
'    If r_rsSchedules Is Nothing Then
'        RestoreSchedule _
'            CLng(l_lngEmployeeNumber), m_datCollPerStart, m_datCollPerEnd
'    ElseIf r_rsSchedules.RecordCount = 0 Then
'        RestoreSchedule _
'            CLng(l_lngEmployeeNumber), m_datCollPerStart, m_datCollPerEnd
'    Else
'        #If TRTRACE Then
'            TraceMessage "", "RestoreSchedule state provided"
'        #End If
'        RestoreScheduleWith r_rsSchedules, v_rsScheduleChgPmts
'    End If
'    #If TRTRACE Then
'        TraceMessage "", "End RestoreSchedule"
'    #End If
'' Restore other pay here, instead of after save because
'' it will not cause a deadlock here even if the transaction
'' mode is not set to required.
'    If RecordsExist(v_rsOtherPay) Then
'        #If TRTRACE Then
'            TraceMessage "", "Begin RestoreOtherPay"
'        #End If
'        RestoreOtherPay r_rsOtherpayEntries, False
''        RestoreOtherPay l_lngEmployeeNumber, m_datCollPerStart, m_datCollPerEnd
'        #If TRTRACE Then
'            TraceMessage "", "End RestoreOtherPay"
'        #End If
'    End If
'
'    If RecordsExist(v_rsTimeEntries) Then
'        #If TRTRACE Then
'            TraceMessage "", "Begin RecordTimecardEntries"
'        #End If
'
'        l_blnResult = RecordTimecardEntries(v_rsTimeEntries)
'        If Not l_blnResult Then
'            objCtx.SetAbort
'            RecordTimeEntry = False
'            GoTo xit
'        End If
'
'        #If TRTRACE Then
'            TraceMessage "", "End RecordTimecardEntries"
'        #End If
''        if a recordset is passed then append to it,
''        otherwise simply return the updates.
'
'        #If TRTRACE Then
'            TraceMessage "", "Begin BuildAccrualBankUpdates"
'        #End If
'        Set l_rsAccrualBankUpdates = BuildAccrualBankUpdates
'        #If TRTRACE Then
'            TraceMessage "", "End BuildAccrualBankUpdates"
'        #End If
'
'        #If TRTRACE Then
'            TraceMessage "", "Begin BuildOvertimeBankUpdates"
'        #End If
'        Set l_rsOvertimeBankUpdates = BuildOvertimeBankUpdates
'        #If TRTRACE Then
'            TraceMessage "", "End BuildOvertimeBankUpdates"
'        #End If
'
'        #If TRTRACE Then
'            TraceMessage "", "Begin BuildEMPRVInterface"
'        #End If
'        Set l_colEMPRVInterface = BuildEMPRVInterface
'        #If TRTRACE Then
'            TraceMessage "", "End BuildEMPRVInterface"
'        #End If
'
'        If r_rsAccrualBankUpdates Is Nothing Then
'            Set r_rsAccrualBankUpdates = l_rsAccrualBankUpdates
'        Else
'            AppendRecordset l_rsAccrualBankUpdates, r_rsAccrualBankUpdates
'        End If
'        Set l_rsAccrualBankUpdates = Nothing
'        If r_rsOvertimeBankUpdates Is Nothing Then
'            Set r_rsOvertimeBankUpdates = l_rsOvertimeBankUpdates
'        Else
'            AppendRecordset l_rsOvertimeBankUpdates, r_rsOvertimeBankUpdates
'        End If
'        Set l_rsOvertimeBankUpdates = Nothing
'        If r_colEMPRVInterface Is Nothing Then
'            Set r_colEMPRVInterface = l_colEMPRVInterface
'        Else
'            For l_intI = 1 To l_colEMPRVInterface.Count
'                r_colEMPRVInterface.Add l_colEMPRVInterface(l_intI)
'            Next l_intI
'        End If
'        Set l_colEMPRVInterface = Nothing
'    End If
'
'    If RecordsExist(v_rsOtherPay) Then
'        #If TRTRACE Then
'            TraceMessage "", "Begin RecordOtherPayEntries"
'        #End If
''        RecordOtherPayEntries v_rsOtherPay, r_rsErrors
'        l_blnResult = RecordOtherPayEntries(v_rsOtherPay)
'        If Not l_blnResult Then
'            objCtx.SetAbort
'            RecordTimeEntry = False
'            GoTo xit
'        End If
'
'        #If TRTRACE Then
'            TraceMessage "", "End RecordOtherPayEntries"
'        #End If
'    End If
'
'    #If TRTRACE Then
'        TraceMessage "", "Begin ResetRecordStatus"
'    #End If
'    Timecard.ResetRecordStatus
'    If Not OtherPay Is Nothing Then
'        OtherPay.ResetRecordStatus
'    End If
'    #If TRTRACE Then
'        TraceMessage "", "End ResetRecordStatus"
'    #End If
'
'ecdraw:
'    If RecordsExist(v_rsECDrawTime) Then
'        l_blnResult = _
'            RecordECDrawTime(v_rsECDrawTime, v_UserId, r_rsErrors)
'        If Not l_blnResult Then
'            objCtx.SetAbort
'            GoTo xit
'        End If
'    End If
'
'    objCtx.SetComplete
'    RecordTimeEntry = True
'
'xit:
'
'    #If TRACENEON Then
'        TraceNeonMessage CStr(lngTraceEmpNum), "RECDTE", "EXIT"
'    #End If
'
''set byref regardless of rec count.
'    Set r_rsErrors = m_objErrors
'
'    Release
'    Set objCtx = Nothing
'
'    #If TRTRACE Then
'        TraceMessage "", "End RecordTimeEntry"
'    #End If
'    Exit Function
'
'eh:
'    objCtx.SetAbort
'    Release
'    Set m_objErrors = Nothing
'    Set r_rsErrors = Nothing
'    Set r_rsAccrualBankUpdates = Nothing
'    Set r_rsOvertimeBankUpdates = Nothing
'    Set r_rsTimecardEntries = Nothing
'    Set r_rsOtherpayEntries = Nothing
'    Set objCtx = Nothing
'    TrsRaiseError m_strCMODULENAME, "RecordTimeEntry", , , , , True

End Function

Public Function RecordTimeEntry2( _
            ByVal v_rsTimeEntries As ADODB.Recordset, _
            ByVal v_rsOtherPay As ADODB.Recordset, _
            ByVal v_rsECDrawTime As ADODB.Recordset, _
            ByVal r_rsEmployees As ADODB.Recordset, _
            ByVal r_rsEmpWorkgroups As ADODB.Recordset, _
            ByRef r_rsTimecardEntries As ADODB.Recordset, _
            ByRef r_rsOtherpayEntries As ADODB.Recordset, _
            ByVal r_rsSchedules As ADODB.Recordset, _
            ByVal v_rsScheduleChgPmts As ADODB.Recordset, _
            ByVal v_startDate As Date, _
            ByVal v_endDate As Date, _
            ByRef r_rsAccrualBankUpdates As ADODB.Recordset, _
            ByRef r_rsOvertimeBankUpdates As ADODB.Recordset, _
            ByRef r_colEMPRVInterface As Collection, _
            ByRef r_rsEMPRVJobTaskEntries As ADODB.Recordset, _
            ByRef r_rsErrors As ADODB.Recordset, _
            ByVal v_UserId As String) As Boolean
            
TrsRaiseError m_strCMODULENAME, "RecordTimeEntry2", trErrObsoleteMethod, , , , True

End Function

Public Function RecordTimeEntry3( _
            ByVal v_rsTimeEntries As ADODB.Recordset, _
            ByVal v_rsOtherPay As ADODB.Recordset, _
            ByVal r_rsEmployees As ADODB.Recordset, _
            ByVal r_rsEmpWorkgroups As ADODB.Recordset, _
            ByRef r_rsTimecardEntries As ADODB.Recordset, _
            ByRef r_rsOtherpayEntries As ADODB.Recordset, _
            ByVal r_rsSchedules As ADODB.Recordset, _
            ByVal v_rsScheduleChgPmts As ADODB.Recordset, _
            ByVal v_startDate As Date, _
            ByVal v_endDate As Date, _
            ByRef r_rsAccrualBankUpdates As ADODB.Recordset, _
            ByRef r_rsOvertimeBankUpdates As ADODB.Recordset, _
            ByRef r_colEMPRVInterface As Collection, _
            ByRef r_rsEMPRVJobTaskEntries As ADODB.Recordset, _
            ByRef r_rsErrors As ADODB.Recordset, _
            ByVal v_UserId As String) As Boolean
    Dim objCtx As ObjectContext
    Dim l_lngEmployeeNumber As Long
    Dim l_datChargeDate As Date
    Dim l_blnResult As Boolean
    Dim l_rsAccrualBankUpdates As ADODB.Recordset
    Dim l_rsOvertimeBankUpdates As ADODB.Recordset
    Dim l_colEMPRVInterface As Collection
    Dim l_rsEMPRVJobTaskEntries As ADODB.Recordset
    Dim l_intI As Integer
    Dim l_intCount As Integer
    Dim intFieldIndex As Integer
On Error GoTo eh
    
'test only
    #If TRACENEON Then
        Dim lngTraceEmpNum As Long
        lngTraceEmpNum = v_rsTimeEntries!KEY_EMP_NUM
        TraceNeonMessage CStr(lngTraceEmpNum), "RECDTE", "BEGIN"
    #End If
    
    #If TRTRACE Then
        TraceMessage "", "Begin RecordTimeEntry3"
    #End If

    Set objCtx = GetObjectContext
    
    m_strUserId = v_UserId
    Set m_objErrors = r_rsErrors

    'for older clients (6.0.0) that are not sending this column to the server
    PreArrangedOTIndMissing = Not FieldExists("PREARRANGED_OT_IND", v_rsTimeEntries)

    If Not GetKeyEmployeeData( _
                v_rsTimeEntries, _
                v_rsOtherPay, _
                l_lngEmployeeNumber, _
                l_datChargeDate) Then
'        If RecordsExist(v_rsECDrawTime) Then
'            GoTo ecdraw
'        Else
            RecordTimeEntry3 = False
            Exit Function
'        End If
    End If

    m_datCollPerStart = v_startDate
    m_datCollPerEnd = v_endDate
    
    #If TRTRACE Then
        TraceMessage "", "Begin RestoreTimecard"
    #End If
    RestoreTimecard r_rsTimecardEntries
    #If TRTRACE Then
        TraceMessage "", "End RestoreTimecard"
    #End If
'    RestoreTimecard _
'        CLng(l_lngEmployeeNumber), m_datCollPerStart, m_datCollPerEnd
    #If TRTRACE Then
        TraceMessage "", "Begin RestoreEmployee"
    #End If
    RestoreEmployee r_rsEmployees, r_rsEmpWorkgroups
    #If TRTRACE Then
        TraceMessage "", "End RestoreEmployee"
    #End If
'    RestoreEmployee _
'        CLng(l_lngEmployeeNumber), m_datCollPerStart, m_datCollPerEnd
    #If TRTRACE Then
        TraceMessage "", "Begin RestoreSchedule"
    #End If
    If r_rsSchedules Is Nothing Then
        RestoreSchedule _
            CLng(l_lngEmployeeNumber), m_datCollPerStart, m_datCollPerEnd
    ElseIf r_rsSchedules.RecordCount = 0 Then
        RestoreSchedule _
            CLng(l_lngEmployeeNumber), m_datCollPerStart, m_datCollPerEnd
    Else
        #If TRTRACE Then
            TraceMessage "", "RestoreSchedule state provided"
        #End If
        RestoreScheduleWith r_rsSchedules, v_rsScheduleChgPmts
    End If
    #If TRTRACE Then
        TraceMessage "", "End RestoreSchedule"
    #End If
' Restore other pay here, instead of after save because
' it will not cause a deadlock here even if the transaction
' mode is not set to required.
    If RecordsExist(v_rsOtherPay) Then
        #If TRTRACE Then
            TraceMessage "", "Begin RestoreOtherPay"
        #End If
        RestoreOtherPay r_rsOtherpayEntries, False
'        RestoreOtherPay l_lngEmployeeNumber, m_datCollPerStart, m_datCollPerEnd
        #If TRTRACE Then
            TraceMessage "", "End RestoreOtherPay"
        #End If
    End If

    If RecordsExist(v_rsTimeEntries) Then
        #If TRTRACE Then
            TraceMessage "", "Begin RecordTimecardEntries"
        #End If
    
        l_blnResult = RecordTimecardEntries(v_rsTimeEntries)
            If Not m_objContractRulesErrors Is Nothing Then
                If m_objContractRulesErrors.RecordCount > 0 Then
                    m_objContractRulesErrors.MoveFirst
                    While Not m_objContractRulesErrors.EOF
                        r_rsErrors.AddNew
                        
                        For intFieldIndex = 0 To r_rsErrors.Fields.Count - 1
                            r_rsErrors.Fields(intFieldIndex).Value = m_objContractRulesErrors.Fields(intFieldIndex).Value
                        Next
                        
                        r_rsErrors.Update
                        
                        m_objContractRulesErrors.MoveNext
                    Wend
                End If
            End If
            
            If Not m_objOTTCConversionErrors Is Nothing Then
                If m_objOTTCConversionErrors.RecordCount > 0 Then
                    m_objOTTCConversionErrors.MoveFirst
                    While Not m_objOTTCConversionErrors.EOF
                        r_rsErrors.AddNew
                        
                        For intFieldIndex = 0 To r_rsErrors.Fields.Count - 1
                            r_rsErrors.Fields(intFieldIndex).Value = m_objOTTCConversionErrors.Fields(intFieldIndex).Value
                        Next
                        
                        r_rsErrors.Update
                        
                        m_objOTTCConversionErrors.MoveNext
                    Wend
                End If
            End If
        
        If Not l_blnResult Then
            objCtx.SetAbort
            RecordTimeEntry3 = False
            GoTo xit
        End If
        
        #If TRTRACE Then
            TraceMessage "", "End RecordTimecardEntries"
        #End If
'        if a recordset is passed then append to it,
'        otherwise simply return the updates.

        #If TRTRACE Then
            TraceMessage "", "Begin BuildAccrualBankUpdates"
        #End If
        Set l_rsAccrualBankUpdates = BuildAccrualBankUpdates
        #If TRTRACE Then
            TraceMessage "", "End BuildAccrualBankUpdates"
        #End If
            
        #If TRTRACE Then
            TraceMessage "", "Begin BuildOvertimeBankUpdates"
        #End If
        Set l_rsOvertimeBankUpdates = BuildOvertimeBankUpdates
        #If TRTRACE Then
            TraceMessage "", "End BuildOvertimeBankUpdates"
        #End If
        
        #If TRTRACE Then
            TraceMessage "", "Begin BuildEMPRVInterface"
        #End If
        Set l_colEMPRVInterface = BuildEMPRVInterface
        #If TRTRACE Then
            TraceMessage "", "End BuildEMPRVInterface"
        #End If
        
        #If TRTRACE Then
            TraceMessage "", "Begin BuildEMPRVJobTasks"
        #End If
        Set l_rsEMPRVJobTaskEntries = BuildEMPRVJobTasks
        #If TRTRACE Then
            TraceMessage "", "End BuildEMPRVJobTasks"
        #End If
    
        If r_rsAccrualBankUpdates Is Nothing Then
            Set r_rsAccrualBankUpdates = l_rsAccrualBankUpdates
        Else
            AppendRecordset l_rsAccrualBankUpdates, r_rsAccrualBankUpdates
        End If
        Set l_rsAccrualBankUpdates = Nothing
        If r_rsOvertimeBankUpdates Is Nothing Then
            Set r_rsOvertimeBankUpdates = l_rsOvertimeBankUpdates
        Else
            AppendRecordset l_rsOvertimeBankUpdates, r_rsOvertimeBankUpdates
        End If
        Set l_rsOvertimeBankUpdates = Nothing
        If r_colEMPRVInterface Is Nothing Then
            Set r_colEMPRVInterface = l_colEMPRVInterface
        Else
            For l_intI = 1 To l_colEMPRVInterface.Count
                r_colEMPRVInterface.Add l_colEMPRVInterface(l_intI)
            Next l_intI
        End If
        Set l_colEMPRVInterface = Nothing
        If r_rsEMPRVJobTaskEntries Is Nothing Then
            Set r_rsEMPRVJobTaskEntries = l_rsEMPRVJobTaskEntries
        Else
            AppendRecordset l_rsEMPRVJobTaskEntries, r_rsEMPRVJobTaskEntries
        End If
        Set l_rsEMPRVJobTaskEntries = Nothing
    End If

    If RecordsExist(v_rsOtherPay) Then
        #If TRTRACE Then
            TraceMessage "", "Begin RecordOtherPayEntries"
        #End If
'        RecordOtherPayEntries v_rsOtherPay, r_rsErrors
        l_blnResult = RecordOtherPayEntries(v_rsOtherPay)
        If Not l_blnResult Then
            objCtx.SetAbort
            RecordTimeEntry3 = False
            GoTo xit
        End If
        
        #If TRTRACE Then
            TraceMessage "", "End RecordOtherPayEntries"
        #End If
    End If

    #If TRTRACE Then
        TraceMessage "", "Begin ResetRecordStatus"
    #End If
    Timecard.ResetRecordStatus
    If Not OtherPay Is Nothing Then
        OtherPay.ResetRecordStatus
    End If
    #If TRTRACE Then
        TraceMessage "", "End ResetRecordStatus"
    #End If
    
'ecdraw:
'    If RecordsExist(v_rsECDrawTime) Then
'        l_blnResult = _
'            RecordECDrawTime(v_rsECDrawTime, v_UserId, r_rsErrors)
'        If Not l_blnResult Then
'            objCtx.SetAbort
'            GoTo xit
'        End If
'    End If
    
    objCtx.SetComplete
    RecordTimeEntry3 = True
    
xit:

    #If TRACENEON Then
        TraceNeonMessage CStr(lngTraceEmpNum), "RECDTE", "EXIT"
    #End If
    
'set byref regardless of rec count.
    Set r_rsErrors = m_objErrors
    
    Release
    Set objCtx = Nothing
    
    #If TRTRACE Then
        TraceMessage "", "End RecordTimeEntry3"
    #End If
    Exit Function

eh:
    objCtx.SetAbort
    Release
    Set m_objErrors = Nothing
    Set r_rsErrors = Nothing
    Set r_rsAccrualBankUpdates = Nothing
    Set r_rsOvertimeBankUpdates = Nothing
    Set r_rsTimecardEntries = Nothing
    Set r_rsOtherpayEntries = Nothing
    Set r_rsEMPRVJobTaskEntries = Nothing
    Set objCtx = Nothing
    TrsRaiseError m_strCMODULENAME, "RecordTimeEntry3", , , , , True

End Function


Private Function GetKeyEmployeeData( _
            v_rsTimeEntries As ADODB.Recordset, _
            v_rsOtherPay As ADODB.Recordset, _
            r_lngEmployeeNumber As Long, _
            r_datChargeDate As Date) As Boolean
'get the empnum and chargedate from first request recordset found

On Error GoTo eh

    If RecordsExist(v_rsTimeEntries) Then
        v_rsTimeEntries.MoveFirst
        r_lngEmployeeNumber = v_rsTimeEntries!KEY_EMP_NUM
        r_datChargeDate = v_rsTimeEntries!KEY_CHARGE_DATE
        GetKeyEmployeeData = True
    Else
        If RecordsExist(v_rsOtherPay) Then
            v_rsOtherPay.MoveFirst
            r_lngEmployeeNumber = v_rsOtherPay!KEY_EMP_NUM
            r_datChargeDate = v_rsOtherPay!KEY_CHARGE_DATE
            GetKeyEmployeeData = True
        End If
    End If
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "GetKeyEmployeeData"

End Function

'Private Function RestoreOtherPay( _
'            ByVal v_lngEmpNum As Long, _
'            ByVal v_datStartDate As Date, _
'            ByVal v_datEndDate As Date) As Boolean
'    Dim objCtx As ObjectContext
'
'On Error GoTo RestoreOtherPay_EH
'
'    Set objCtx = GetObjectContext
'    Set mobjOtherPay = CreateObject("busTimecard.cboOtherPay")
'
'    mobjOtherPay.Restore v_lngEmpNum, v_datStartDate, v_datEndDate
'
'    OtherPay.SetBSTimecard SafeRef(Me)
'    OtherPay.SetSchedule Timecard.Schedule
'    OtherPay.SetEmployee Timecard.Employee
'
'Exit Function
'RestoreOtherPay_EH:
'
'      TrsRaiseError m_strCMODULENAME, "RestoreOtherPay"
'
'End Function

'Private Function RestoreTimecard( _
'            r_rsTimecardEntries As ADODB.Recordset) As Boolean
''Private Function RestoreTimecard( _
''            empId As Long, startDt As Date, endDt As Date) As Boolean
'    Dim objCtx As ObjectContext
'
'    Set objCtx = GetObjectContext
'    Set mobjTimecard = CreateObject("busTimecard.cboTimecard")
'    mobjTimecard.RestoreWithState r_rsTimecardEntries
''    mobjTimecard.Restore empId, startDt, endDt
'    mobjTimecard.SetBSTimecard SafeRef(Me)
'End Function

Private Function RestoreOtherPay( _
            r_rsOtherpayEntries As ADODB.Recordset, _
            Optional ByVal v_blnOvertime As Boolean = False) As Boolean
On Error GoTo eh

    Set mobjOtherPay = CreateObject("busTimecard.cboOtherPay")
    mobjOtherPay.RestoreWithState r_rsOtherpayEntries
    mobjOtherPay.SetBSTimecard SafeRef(Me)
    mobjOtherPay.SetSchedule Timecard.Schedule
    mobjOtherPay.SetEmployee Timecard.Employee
    mobjOtherPay.SetOvertime v_blnOvertime
    Exit Function
    
eh:
      TrsRaiseError m_strCMODULENAME, "RestoreOtherPay"
End Function


Private Function RecordTimecardEntries( _
            ByRef v_rsTimeEntries As ADODB.Recordset) As Boolean
' Answer false when a validation error occurs.  This will send back
' the brv recordset with all errors since a vb error was not raised.
    Dim l_varEmpChgDates As Variant     'array of empNum and ChgDate
    Dim lstrEmpNum As String
    Dim lcolErrArgs As Collection
    Dim lstrOverExc As String
    
'Release 07/25/2001 -- JFB  made changes in this module so vacations
'                            are processed with Holidays before
'                            all other time entries so validations
'                            work properly
    
On Error GoTo eh

    #If TRTRACE Then
        TraceMessage "", "Begin RecordTimecardEntries"
    #End If
    RecordTimecardEntries = True
    v_rsTimeEntries.MoveFirst
    lstrEmpNum = v_rsTimeEntries!KEY_EMP_NUM
    l_varEmpChgDates = GetEmployeeChargeDates(v_rsTimeEntries)

'Why is this filter here.  JIMB does not think it is needed.
    v_rsTimeEntries.Filter = "KEY_EMP_NUM = " & lstrEmpNum
    
' Process Deletes first
    #If TRTRACE Then
        TraceMessage "", "Begin RecordTimecardEntries DELETEs"
    #End If
        If RecordsExist(v_rsTimeEntries) Then
        v_rsTimeEntries.MoveFirst
    End If
    Do Until v_rsTimeEntries.EOF
        If v_rsTimeEntries!RECORD_STATUS = "D" Then
            Timecard.RecordTimeEntry v_rsTimeEntries
        End If
        v_rsTimeEntries.MoveNext
    Loop
    #If TRTRACE Then
        TraceMessage "", "END RecordTimecardEntries DELETEs"
    #End If


' Process holidays first
    #If TRTRACE Then
        TraceMessage "", "Begin RecordTimecardEntries holidays"
    #End If

'Process All Holiday Time First
    If RecordsExist(v_rsTimeEntries) Then
        v_rsTimeEntries.MoveFirst
    End If
    Do Until v_rsTimeEntries.EOF
        If v_rsTimeEntries!RECORD_STATUS <> "D" Then
            If IsNull(v_rsTimeEntries!EXCEPTION_CD) Then
            Else
                If IsAHolidayException(v_rsTimeEntries!EXCEPTION_CD) Then
                    Timecard.RecordTimeEntry v_rsTimeEntries
                End If
            End If
        End If
        v_rsTimeEntries.MoveNext
    Loop

    #If TRTRACE Then
        TraceMessage "", "End RecordTimecardEntries holidays"
    #End If
    
    
' Process vacations next
    #If TRTRACE Then
        TraceMessage "", "Begin RecordTimecardEntries vacations"
    #End If

    If RecordsExist(v_rsTimeEntries) Then
        v_rsTimeEntries.MoveFirst
    End If
    Do Until v_rsTimeEntries.EOF
        If v_rsTimeEntries!RECORD_STATUS <> "D" Then
            If IsNull(v_rsTimeEntries!EXCEPTION_CD) Then
            Else
                If IsAVacationException(v_rsTimeEntries!EXCEPTION_CD) Then
                    Timecard.RecordTimeEntry v_rsTimeEntries
                End If
            End If
        End If
        v_rsTimeEntries.MoveNext
    Loop
    
    ' Process special overrides next
    #If TRTRACE Then
        TraceMessage "", "Begin RecordTimecardEntries special overrides "
    #End If

    If RecordsExist(v_rsTimeEntries) Then
        v_rsTimeEntries.MoveFirst
    End If
    Do Until v_rsTimeEntries.EOF
        If v_rsTimeEntries!RECORD_STATUS <> "D" Then
            If IsNull(v_rsTimeEntries!OVER_EXC_CD) Then
                lstrOverExc = ""
            Else
                lstrOverExc = v_rsTimeEntries!OVER_EXC_CD
                If IsSpecialOverrideCode(lstrOverExc) Then
                    Timecard.RecordTimeEntry v_rsTimeEntries
                End If
            End If
        End If
        v_rsTimeEntries.MoveNext
    Loop

    #If TRTRACE Then
        TraceMessage "", "End RecordTimecardEntries vacations"
    #End If

'Process all other time (except holidays)
    #If TRTRACE Then
        TraceMessage "", "Begin RecordTimecardEntries non-holidays"
    #End If
    
    v_rsTimeEntries.MoveFirst
    Do Until v_rsTimeEntries.EOF
        If v_rsTimeEntries!RECORD_STATUS <> "D" Then
            If IsNull(v_rsTimeEntries!EXCEPTION_CD) Then
                Timecard.RecordTimeEntry v_rsTimeEntries
            Else
                If IsNull(v_rsTimeEntries!OVER_EXC_CD) Then
                    lstrOverExc = ""
                Else
                    lstrOverExc = v_rsTimeEntries!OVER_EXC_CD
                End If
                If IsAHolidayException(v_rsTimeEntries!EXCEPTION_CD) Or IsAVacationException(v_rsTimeEntries!EXCEPTION_CD) Or IsSpecialOverrideCode(lstrOverExc) Then
                    'Skip already processed
                Else
                     Timecard.RecordTimeEntry v_rsTimeEntries
                End If
            End If
        End If
        v_rsTimeEntries.MoveNext
    Loop

    #If TRTRACE Then
        TraceMessage "", "End RecordTimecardEntries non-holidays"
    #End If

    If Validate Then
        #If TRTRACE Then
            TraceMessage "", "Begin Save"
        #End If
        
        Set m_objUserEnteredDataForCRAC = v_rsTimeEntries
        If Not Save(l_varEmpChgDates, True) Then
            'New to support contact rules
            RecordTimecardEntries = False
            'Set m_objUserEnteredDataForCRAC = Nothing
            'Exit Function
        End If
        Set m_objUserEnteredDataForCRAC = Nothing
        
        #If TRTRACE Then
            TraceMessage "", "End Save"
        #End If
    Else
        RecordTimecardEntries = False
    End If
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "RecordTimecardEntries"

End Function

'Private Function RecordECDrawTime( _
'            ByRef v_rsECDrawTime As ADODB.Recordset, _
'            ByVal v_strUserId As String, _
'            ByRef r_rsErrors As ADODB.Recordset) As Boolean
'    Dim objECDraw As busECDrawTimeMgr.cbsECDrawTimeMgr
'    Dim l_strErrorMsg As String
'    Dim l_blnReturnVal As Boolean
'
'On Error GoTo eh
'
'    #If TRTRACE Then
'        TraceMessage "", "Begin RecordECDrawTime"
'    #End If
'
'    Set objECDraw = _
'        CreateObject("busECDrawTimeMgr.cbsECDrawTimeMgr")
'    l_blnReturnVal = _
'        objECDraw.Record(v_rsECDrawTime, v_strUserId, l_strErrorMsg)
'    If Not l_blnReturnVal Then
'        AppendError r_rsErrors, _
'                    0, _
'                    Date, _
'                    Date, _
'                    "", _
'                    "", _
'                    "0.0", _
'                    "", _
'                    "", _
'                    vbObjectError + 31000, _
'                    mstrCBUSINESSWAR, _
'                    l_strErrorMsg, _
'                    m_strCMODULENAME
'        Err.Raise 40002, m_strCMODULENAME, l_strErrorMsg
'    End If
'    RecordECDrawTime = l_blnReturnVal
'
'    #If TRTRACE Then
'        TraceMessage "", "End RecordECDrawTime"
'    #End If
'
'    Exit Function
'
'eh:
'    TrsRaiseError m_strCMODULENAME, "RecordECDrawTime"
'
'End Function

Private Function RecordOtherPayEntries( _
            ByRef v_rsOtherPay As ADODB.Recordset)
    
On Error GoTo eh

    RecordOtherPayEntries = True
    v_rsOtherPay.MoveFirst
    Do Until v_rsOtherPay.EOF
        OtherPay.RecordOtherPay v_rsOtherPay, Nothing
        v_rsOtherPay.MoveNext
    Loop

    If ValidateOtherPay Then
        SaveOtherPay
    Else
        RecordOtherPayEntries = False
    End If
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "RecordOtherPayEntries"
    
End Function

Private Function OtherPay() As cboOtherPay
    Set OtherPay = mobjOtherPay
End Function

Private Function GetEmployeeChargeDates( _
                    ByVal v_rsTimeEntries As ADODB.Recordset) As Variant
' Answer a distinct array of employee numbers and charge dates
' Dimension with the largest possible number of records for now.
    Dim l_varEmpChgDate()
    Dim l_varEmpChgDateOut()
    Dim l_lngEmployeeNumber As Long
    Dim l_datChargeDate As Date
    Dim l_intNumberOfElements As Integer
    Dim l_blnExists As Boolean
    Dim i As Integer
    
    On Error GoTo eh
    
    ReDim l_varEmpChgDate(v_rsTimeEntries.RecordCount, 1)
    'guard against and empty recordset
    If v_rsTimeEntries Is Nothing Then
        GetEmployeeChargeDates = Empty
        Exit Function
    End If
    
    If v_rsTimeEntries.BOF And v_rsTimeEntries.EOF Then
        GetEmployeeChargeDates = Empty
        Exit Function
    End If
    
    v_rsTimeEntries.MoveFirst
    Do Until v_rsTimeEntries.EOF
        If v_rsTimeEntries!SCHED_TIME_IND = "Y" Then
        'search the array for existence of empNum and chargedate
            l_lngEmployeeNumber = v_rsTimeEntries!KEY_EMP_NUM
            l_datChargeDate = v_rsTimeEntries!KEY_CHARGE_DATE
            l_blnExists = False
            For i = 0 To l_intNumberOfElements
                If l_varEmpChgDate(i, 0) = l_lngEmployeeNumber And _
                    l_varEmpChgDate(i, 1) = l_datChargeDate Then
                    l_blnExists = True
                    Exit For
                End If
            Next
            If Not l_blnExists Then
                l_varEmpChgDate(l_intNumberOfElements, 0) = l_lngEmployeeNumber
                l_varEmpChgDate(l_intNumberOfElements, 1) = l_datChargeDate
                l_intNumberOfElements = l_intNumberOfElements + 1
            End If
        End If
        v_rsTimeEntries.MoveNext
    Loop
    
'    redimension array and preserve the contents
    If l_intNumberOfElements = 0 Then
        GetEmployeeChargeDates = Empty
    Else
        ReDim l_varEmpChgDateOut(l_intNumberOfElements - 1, 1)
        For i = 0 To l_intNumberOfElements - 1
            l_varEmpChgDateOut(i, 0) = l_varEmpChgDate(i, 0)
            l_varEmpChgDateOut(i, 1) = l_varEmpChgDate(i, 1)
        Next
        GetEmployeeChargeDates = l_varEmpChgDateOut
    End If
' be sure to restore the position before leaving
    v_rsTimeEntries.MoveFirst
    Exit Function
                    
eh:
    TrsRaiseError m_strCMODULENAME, "GetEmployeeChargeDates"
                    
End Function

Private Function BuildAccrualBankUpdates() As ADODB.Recordset
'BLUE023

On Error GoTo eh
    
    Set BuildAccrualBankUpdates = Timecard.BuildAccrualBankUpdates
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "BuildAccrualBankUpdates"

End Function

Private Sub UpdateOtherPayEntry(rsOtherPay As ADODB.Recordset, OtherPayEntry As cboOtherPayEntry)
                
On Error GoTo UpdateOtherPayEntry_EH

    rsOtherPay.MoveFirst
    rsOtherPay.Find "INDEX_TIME = '" & OtherPayEntry.IndexTime & "'"
    
    rsOtherPay!RECORD_STATUS = "C"
    rsOtherPay!START_TIME = OtherPayEntry.startTime
    rsOtherPay!END_TIME = OtherPayEntry.EndTime
    rsOtherPay!JOB_NUM = OtherPayEntry.JobNumber
    rsOtherPay!WG_NUM = OtherPayEntry.WorkgroupNumber
    rsOtherPay!DEPT_NUM = OtherPayEntry.DepartmentNumber
    rsOtherPay!UNION_CD = OtherPayEntry.Union
    rsOtherPay!Hours = OtherPayEntry.hrs
    rsOtherPay!AMOUNT = OtherPayEntry.Amt
    rsOtherPay!START_LOC_CD = OtherPayEntry.StartLocation
    rsOtherPay!END_LOC_CD = OtherPayEntry.EndLocation
    rsOtherPay!LOC_CD = OtherPayEntry.Location
    rsOtherPay!EARNINGS_CD = OtherPayEntry.EarningsCode
    rsOtherPay!COMP_SHIFT_CD = OtherPayEntry.ShiftComponentCode
    rsOtherPay!COMP_SHIFT_CD_AMT = OtherPayEntry.ShiftComponentAmount
    rsOtherPay!DISTRIBUTION_ID = OtherPayEntry.DistributionId
    rsOtherPay!DISTRIBUTION_TYPE = OtherPayEntry.DistributionType
    rsOtherPay!DISTR_DEPT_NUM = OtherPayEntry.DistributionDeptNum
    rsOtherPay!EXCEPTION_CD = OtherPayEntry.ExceptionCode
    rsOtherPay!APPR_IND = OtherPayEntry.ApprovalInd
    rsOtherPay!APPR_BY_EMP_NUM = OtherPayEntry.ApproveByEmpNum
'***OP2***
    rsOtherPay!CHEM_MTL_TSTR_IND = OtherPayEntry.ChemTesterInd
    rsOtherPay!HALF_TRIP_IND = OtherPayEntry.HalfTripInd
'***OP2***
    rsOtherPay!APPR_DATE = OtherPayEntry.ApprovalDate
    rsOtherPay!LAST_MOD_EMP_NUM = CLng(m_strUserId)
    rsOtherPay.Update
    
    Exit Sub
    
UpdateOtherPayEntry_EH:
    TrsRaiseError m_strCMODULENAME, "UpdateOtherPayEntry"
    
End Sub


Private Function Validate() As Boolean
    Dim entry As cboTimecardEntry
    Dim answer As Boolean

On Error GoTo eh

    Validate = True
    For Each entry In Timecard.TimecardEntries
        If entry.IsNew Or entry.IsModified Then
            If Not entry.IsValid Then
                Validate = False
                Exit Function
            End If
        End If
    Next
    
    CheckReportedHours
    
    If CheckForOverlappingTime Then
        Validate = False
        Exit Function
    End If
'    Validate = Not CheckForOverlappingTime
    

    answer = ValidateCanNonCanOvertime
    If Not answer Then
        Validate = answer
        Exit Function
    End If

    answer = CheckShiftReportedHours
    If answer Then
        Validate = False
        Exit Function
    End If
    
    answer = CheckShiftReportedLunchHours
    If answer Then
        Validate = False
        Exit Function
    End If
    
    answer = CheckShiftReportedOvertimeHours
    If answer Then
        Validate = False
        Exit Function
    End If
    
    answer = CheckScheduleTimeStamp
    If answer Then
        Validate = False
        Exit Function
    End If

    GetHolidays
    answer = CheckHolidayDate
    If answer Then
        Validate = False
        Exit Function
    End If
    
    answer = CheckHolidayBanked
    If answer Then
        Validate = False
        Exit Function
    End If
    
    answer = CheckHolidayHours
    If answer Then
        Validate = False
        Exit Function
    End If
    
    ''jimbo check parental bonding
    answer = CheckHolidayParentalBonding
    If answer Then
        Validate = False
        Exit Function
    End If
    
    answer = CheckMAXIMO
    If answer Then
        Validate = False
        Exit Function
    End If
    
    answer = CheckGLAccountingAfterPOET
    If answer Then
        Validate = False
        Exit Function
    End If

    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "Validate"
    
End Function

Private Function ValidateOtherPay() As Boolean
    Dim answer As Boolean
    Dim l_objOtherPayEntry As cboOtherPayEntry

On Error GoTo ValidateOtherPay_EH

    ValidateOtherPay = True
    
    For Each l_objOtherPayEntry In OtherPay.OtherPayEntries
        If l_objOtherPayEntry.IsNew Or l_objOtherPayEntry.IsModified Then
            If Not l_objOtherPayEntry.IsValid Then
                ValidateOtherPay = False
                Exit Function
            End If
        End If
    Next

    ValidateTTA

    answer = CheckGLAccountingAfterPOETOtherPay
    If answer Then
        ValidateOtherPay = False
        Exit Function
    End If

Exit Function
ValidateOtherPay_EH:

      TrsRaiseError m_strCMODULENAME, "ValidateOtherPay"
      
End Function

Private Function Save(ByVal v_varEmpChgDates As Variant, v_blnAllowContractRulesApplication As Boolean) As Boolean
    'It seems like it would be safer to pass a recordset that
    'contained the save request records only, however the save
    'process is dependent on the original value of a field.
    'Try to do something better later.
    'Note: If dbSaveTimecard is out of process the filter is
    'removed so it is important that the record status is updated
    'for previous employee requests before calling saveTimecard.
    
    Dim dbSaveTimecard As dbSaveTimecard.cdSaveTimecard
    Dim rs As ADODB.Recordset
    Dim col As Collection
    Dim entry As cboTimecardEntry
    Dim ans As Boolean
    Dim rsSaveRequest As ADODB.Recordset
    Dim comint4 As Ameren_TRIS_Server_COMIntegration4.ValidateDistribution
    Dim rc As Boolean

On Error GoTo eh
    
    Save = False
    
    Set rs = Timecard.TimecardEntryRecordset
    Set col = Timecard.TimecardEntriesPrivate
    For Each entry In col
        If entry.IsNew Then
            AddNewTimeEntry rs, entry
        ElseIf entry.IsModified Then
            UpdateTimeEntry rs, entry
        ElseIf entry.IsDeleted Then
            DeleteTimeEntry rs, entry
        End If
    Next
    
    'Apply the contract rules here, since this is after the entries
    'are applied to the recordset.  Also, this is before OT banks
    'are processed.
    If v_blnAllowContractRulesApplication Then
        Dim CRACIntegration As Ameren_TRIS_Server_COMIntegration.Time
        Set CRACIntegration = New Ameren_TRIS_Server_COMIntegration.Time
        Set m_objContractRulesErrors = mstdErrorHandling.NewErrorset
        Dim HasErrors As Boolean
        CRACIntegration.ApplyContractRules rs, m_objUserEnteredDataForCRAC, _
            m_objContractRulesErrors, CLng(m_strUserId), HasErrors
        Set CRACIntegration = Nothing

'//CRAC via client.comintegration.library
'        Set m_objContractRulesErrors = mstdErrorHandling.NewErrorset
'        Dim HasErrors As Boolean
'
'        '//check SOME registry flag setting
'        '//utilRegAccessLib.cuRegistryAccess
'        '//DoesKeyExist HKEY_LOCAL_MACHINE, "Software\Ameren\TRIS\Trace???"
'        '//GetSetting(HKEY_LOCAL_MACHINE, "Software\Ameren\TRIS\Trace???", "SaveCRACRecordsets") '//registry value
'
'        '//see mADOExtensions, the data will be saved to %temp% in xml format, suffixes will be added to filenames to avoid duplicates
'        '//the call could come from OT Canvass, so try to be specific in the log name
'        If CRACInterfaceLog = 1 Then
'            Dim LogNamePrefix As String
'            Dim RowId As Double
'
'            LogNamePrefix = "CRAC" '//from time collection
'            RowId = 0
'On Error GoTo LogNamePrefixEH
'            If RecordsExist(m_objUserEnteredDataForCRAC) Then
'                LogNamePrefix = "CRAC" '//from time collection
'            Else '//don't think that the 2nd recordset comes over from ot canvass
'                If RecordsExist(rs) Then
'                    rs.MoveFirst
'                    If FieldExists("ROW_ID", rs) Then
'                        RowId = rs.Fields("ROW_ID").Value
'                    End If
'
'                    If IsNull(RowId) Or RowId = 0 Then
'                        LogNamePrefix = "CRAC_OTPost" '//from ot canvas post
'                    Else
'                        LogNamePrefix = "CRAC_OTUnPost" '//from ot canvas un-post
'                    End If
'                Else
'                    LogNamePrefix = "CRAC_OT"
'                End If
'            End If
'LogNamePrefixEH:
'On Error GoTo eh
'        End If
'
'        '//see mADOExtensions, the data will be saved to %temp% in xml format, suffixes will be added to filenames to avoid duplicates
'        If CRACInterfaceLog = 1 Then
'            SaveRecordset rs, LogNamePrefix & "_TimeData", False '//save using any filter
'            SaveRecordset m_objUserEnteredDataForCRAC, LogNamePrefix & "_RawUserEnteredData", False '//save using any filter
'        End If
'        '//
'
'        If CRACInterface = 2 Then
'            '//use the New process to load the CRAC into a different process/machine (via WCF)
'            Dim CRACIntegrationWCF As Ameren_TRIS_Client_COMIntegration_Library.ContractRuleAutomation
'            Set CRACIntegrationWCF = New Ameren_TRIS_Client_COMIntegration_Library.ContractRuleAutomation
'            CRACIntegrationWCF.ApplyContractRulesForTime rs, m_objUserEnteredDataForCRAC, _
'                m_objContractRulesErrors, CLng(m_strUserId), HasErrors
'            Set CRACIntegrationWCF = Nothing
'        Else '//CRACInterface = 1
'            '//use the Old process to load the CRAC into the COM+ process
'            Dim CRACIntegration As Ameren_TRIS_Server_COMIntegration.Time
'            Set CRACIntegration = New Ameren_TRIS_Server_COMIntegration.Time
'            CRACIntegration.ApplyContractRules rs, m_objUserEnteredDataForCRAC, _
'                m_objContractRulesErrors, CLng(m_strUserId), HasErrors
'            Set CRACIntegration = Nothing
'        End If
'
'        '//see mADOExtensions, the data will be saved to %temp% in xml format, suffixes will be added to filenames to avoid duplicates
'        If CRACInterfaceLog = 1 Then
'            SaveRecordset m_objContractRulesErrors, LogNamePrefix & "_Brvs", False '//save using any filter
'        End If
'        '//
        
        If HasErrors Then
            Save = False
            Exit Function
        End If
    End If
    
    '//for OTTC cutover
    If IsOTTCEffective() And IsOTTCTransitionActive() And Not IsOTTCCutoverOutageActive() And Not IsOTTCBypassAccountValidationActive() Then
        Set m_objOTTCConversionErrors = mstdErrorHandling.NewErrorset
        Dim HasErrors2 As Boolean
        Dim l_strCmpyCd As String
        
        l_strCmpyCd = "AMS"
        If Not mobjTimecard Is Nothing Then
            If Not mobjTimecard.Employee Is Nothing Then
                l_strCmpyCd = mobjTimecard.Employee.CompanyCode
            End If
        End If
        
'        If RecordsExist(rs) Then
'            l_strCmpyCd = rs!CMPY_CD.Value
'        End If
               
        Set comint4 = New Ameren_TRIS_Server_COMIntegration4.ValidateDistribution
        
        rc = comint4.ConvertTimecardForSave_(rs, l_strCmpyCd, m_objOTTCConversionErrors, HasErrors2)
        Set comint4 = Nothing
        
        If HasErrors2 Then
            Save = False
            Exit Function
        End If
    End If
    
    If Not (rs.BOF And rs.EOF) Then
        rs.MoveFirst
    End If
    
    Set dbSaveTimecard = CreateObject("dbSaveTimecard.cdSaveTimeCard")
        
    ans = dbSaveTimecard.SaveTimeCard(rs)
    
    If Not IsEmpty(v_varEmpChgDates) Then
        dbSaveTimecard.RemoveExceptionsFromSchedule v_varEmpChgDates, UserEmployeeNumber
    End If
    
    Save = True
    
    Exit Function
eh:
    Set comint4 = Nothing
    TrsRaiseError m_strCMODULENAME, "Save"
End Function

Private Function SaveOtherPay() As Boolean
    
    Dim dbSaveTimecard As dbSaveTimecard.cdSaveTimecard
    Dim rs As ADODB.Recordset
    Dim col As Collection
    Dim entry As cboOtherPayEntry
    
On Error GoTo SaveOtherPay_EH

    Set dbSaveTimecard = CreateObject("dbSaveTimecard.cdSaveTimeCard")
    
    Set rs = OtherPay.OtherPayEntryRecordset
    Set col = OtherPay.OtherPayEntriesPrivate
    
    '  Convert from Collection to Recordset
    For Each entry In col
        If entry.IsNew Then
            AddNewOtherPayEntry rs, entry
        ElseIf entry.IsModified Then
            UpdateOtherPayEntry rs, entry
        ElseIf entry.IsDeleted Then
            DeleteOtherPayEntry rs, entry
        End If
    Next

    If RecordsExist(rs) Then
        rs.MoveFirst
    End If
    
    dbSaveTimecard.SaveOtherPay rs
    
Exit Function
SaveOtherPay_EH:

    TrsRaiseError m_strCMODULENAME, "SaveOtherPay"

End Function


Private Sub AddNewTimeEntry(rs As ADODB.Recordset, _
                entry As cboTimecardEntry)
                
On Error GoTo eh

    rs.AddNew
    rs!RECORD_STATUS = "I"
    rs!SCHED_TIME_IND = entry.ScheduleTimeIndicator
    rs!EMP_NUM = entry.EmployeeNumber
    rs!CHARGE_DATE = entry.ChargeDate
    rs!DATE_WORKED = entry.DateWorked
    rs!DAY_NAME = entry.DateWorkedDayName
    rs!START_TIME = entry.startTime
    rs!END_TIME = entry.EndTime
    rs!Hours = entry.Hours
    
    rs!DEPT_NUM = entry.DepartmentNumber
    rs!UNION_CD = entry.UnionCode
    rs!LOC_CD = entry.LocationCode
    rs!WG_NUM = entry.WorkgroupNumber
    rs!EXCEPTION_CD = entry.ExceptionCode
    rs!OVER_EXC_CD = entry.OverrideExceptionCode
    rs!EARNINGS_CD = entry.EarningsCode
    rs!TRADE_EXC_CD = entry.TradeExceptionCode
    rs!JOB_NUM = entry.JobNumber
    rs!APPR_IND = entry.ApprovalIndicator
    rs!APPR_DATE = entry.ApprovedDate
    rs!APPR_BY_EMP_NUM = entry.ApprovedBy
    rs!FAMILY_LEAVE_IND = entry.FamilyLeaveIndicator
    rs!CASE_NUM = entry.CaseNumber
    rs!OT_OPP_NUM = entry.OvertimeOpportunityNumber
    rs!OT_OPP_DESC = entry.OvertimeOpportunityDescription
    rs!OT_REASON_CD = entry.OvertimeReasonCode
    If Not PreArrangedOTIndMissing Then
        rs!PREARRANGED_OT_IND = entry.PreArrangedOTIndicator
    End If
    rs!OT_BANK_CHRG_DT = entry.OvertimeBankChargeDate
    rs!DISTRIBUTION_ID = entry.DistributionId
    rs!DISTRIBUTION_TYPE = entry.DistributionType
    rs!JOB_NBR = entry.JobNbr
    rs!JOBTASK_NBR = entry.JobTaskNbr
    rs!JOBTASK_STEP_NBR = entry.JobTaskStepNbr
    rs!EMPRV_DATABASE_ID = entry.EMPRVDatabaseID
    rs!DISTR_DEPT_NUM = entry.DistributionDepartmentNumber
    rs!SCHED_CHG_EXC_CD = entry.ScheduleChangeExceptionCode
    rs!SCHED_CHG_EARN_CD = entry.ScheduleChangeEarningsCode
    rs!COMP_SHIFT_CD = entry.ShiftComponentCode
    rs!COMP_SHIFT_CD_AMT = entry.ShiftComponentAmount
    rs!REL_CD = entry.ReliefCode
    rs!COMP_CD = entry.ComponentCode
    rs!COMP_AMT = entry.ComponentAmount
    rs!LAST_MOD_EMP_NUM = CLng(m_strUserId)
    rs!TIME_STAMP = entry.GroupTimestamp
    rs!TRAVEL_PAY_HRS = entry.TravelPayHours
    rs!SCHED_CD = entry.ScheduleCode
    rs!SCHED_SRC_CD = entry.ScheduleSourceCode
    
    rs.Update
    Exit Sub

eh:

    TrsRaiseError m_strCMODULENAME, "AddNewTimeEntry"
    
End Sub

Private Sub AddNewOtherPayEntry(rsOtherPay As ADODB.Recordset, OtherPayEntry As cboOtherPayEntry)
                
On Error GoTo AddNewOtherPayEntry_EH

    rsOtherPay.AddNew
    rsOtherPay!RECORD_STATUS = "I"
    rsOtherPay!EMP_NUM = OtherPayEntry.EmployeeNumber
    rsOtherPay!CHARGE_DATE = OtherPayEntry.ChargeDate
    rsOtherPay!DATE_WORKED = OtherPayEntry.DateWorked
    rsOtherPay!START_TIME = OtherPayEntry.startTime
    rsOtherPay!END_TIME = OtherPayEntry.EndTime
    rsOtherPay!JOB_NUM = OtherPayEntry.JobNumber
    rsOtherPay!WG_NUM = OtherPayEntry.WorkgroupNumber
    rsOtherPay!DEPT_NUM = OtherPayEntry.DepartmentNumber
    rsOtherPay!UNION_CD = OtherPayEntry.Union
    rsOtherPay!Hours = OtherPayEntry.hrs
    rsOtherPay!AMOUNT = OtherPayEntry.Amt
    rsOtherPay!START_LOC_CD = OtherPayEntry.StartLocation
    rsOtherPay!END_LOC_CD = OtherPayEntry.EndLocation
    rsOtherPay!LOC_CD = OtherPayEntry.Location
    rsOtherPay!EARNINGS_CD = OtherPayEntry.EarningsCode
    rsOtherPay!COMP_SHIFT_CD = OtherPayEntry.ShiftComponentCode
    rsOtherPay!COMP_SHIFT_CD_AMT = OtherPayEntry.ShiftComponentAmount
    rsOtherPay!DISTRIBUTION_ID = OtherPayEntry.DistributionId
    rsOtherPay!DISTRIBUTION_TYPE = OtherPayEntry.DistributionType
    rsOtherPay!DISTR_DEPT_NUM = OtherPayEntry.DistributionDeptNum
    rsOtherPay!EXCEPTION_CD = OtherPayEntry.ExceptionCode
    rsOtherPay!APPR_IND = OtherPayEntry.ApprovalInd
    rsOtherPay!APPR_BY_EMP_NUM = OtherPayEntry.ApproveByEmpNum
    rsOtherPay!APPR_DATE = OtherPayEntry.ApprovalDate
'***OP2***
    rsOtherPay!CHEM_MTL_TSTR_IND = OtherPayEntry.ChemTesterInd
    rsOtherPay!HALF_TRIP_IND = OtherPayEntry.HalfTripInd
'***OP2***
    rsOtherPay!OT_OPP_NUM = OtherPayEntry.OvertimeOpportunityNumber
    rsOtherPay!DAY_NAME = OtherPayEntry.DayName
    rsOtherPay!LAST_MOD_EMP_NUM = CLng(m_strUserId)
    rsOtherPay.Update
    
Exit Sub
AddNewOtherPayEntry_EH:

    TrsRaiseError m_strCMODULENAME, "AddNewOtherPayEntry"
    
End Sub

Private Sub UpdateTimeEntry( _
                rs As ADODB.Recordset, _
                entry As cboTimecardEntry)
' Position rs on existing record for entry.
' Update the fields that have changed.
'
On Error GoTo eh
    
    rs.MoveFirst
    rs.Find "INDEX_TIME = '" & entry.IndexTime & "'"

    rs!RECORD_STATUS = "C"
    rs!DATE_WORKED = entry.DateWorked
    rs!SCHED_TIME_IND = entry.ScheduleTimeIndicator
    rs!OT_ADJUST_DESC = entry.OTAdjustDesc
    rs!START_TIME = entry.startTime
    rs!END_TIME = entry.EndTime
    rs!Hours = entry.Hours
    rs!DEPT_NUM = entry.DepartmentNumber
    rs!UNION_CD = entry.UnionCode
    rs!LOC_CD = entry.LocationCode
    rs!WG_NUM = entry.WorkgroupNumber
    rs!EXCEPTION_CD = entry.ExceptionCode
    rs!OVER_EXC_CD = entry.OverrideExceptionCode
    rs!EARNINGS_CD = entry.EarningsCode
    rs!TRADE_EXC_CD = entry.TradeExceptionCode
    rs!JOB_NUM = entry.JobNumber
    rs!APPR_IND = entry.ApprovalIndicator
    rs!APPR_BY_EMP_NUM = entry.ApprovedBy
    rs!APPR_DATE = entry.ApprovedDate
    rs!FAMILY_LEAVE_IND = entry.FamilyLeaveIndicator
    rs!CASE_NUM = entry.CaseNumber
    rs!OT_OPP_NUM = entry.OvertimeOpportunityNumber
    rs!OT_OPP_DESC = entry.OvertimeOpportunityDescription
    rs!OT_BANK_CHRG_DT = entry.OvertimeBankChargeDate
    rs!OT_REASON_CD = entry.OvertimeReasonCode
    If Not PreArrangedOTIndMissing Then
        rs!PREARRANGED_OT_IND = entry.PreArrangedOTIndicator
    End If
    rs!DISTRIBUTION_ID = entry.DistributionId
    rs!DISTRIBUTION_TYPE = entry.DistributionType
    rs!JOB_NBR = entry.JobNbr
    rs!JOBTASK_NBR = entry.JobTaskNbr
    rs!JOBTASK_STEP_NBR = entry.JobTaskStepNbr
    rs!EMPRV_DATABASE_ID = entry.EMPRVDatabaseID
    rs!DISTR_DEPT_NUM = entry.DistributionDepartmentNumber
    rs!SCHED_CHG_EXC_CD = entry.ScheduleChangeExceptionCode
    rs!SCHED_CHG_EARN_CD = entry.ScheduleChangeEarningsCode
    rs!COMP_SHIFT_CD = entry.ShiftComponentCode
    rs!COMP_SHIFT_CD_AMT = entry.ShiftComponentAmount
    rs!REL_CD = entry.ReliefCode
    rs!COMP_CD = entry.ComponentCode
    rs!COMP_AMT = entry.ComponentAmount
    rs!LAST_MOD_EMP_NUM = CLng(m_strUserId)
    rs!TIME_STAMP = entry.GroupTimestamp
    rs!TRAVEL_PAY_HRS = entry.TravelPayHours
    rs!SCHED_CD = entry.ScheduleCode
    rs!SCHED_SRC_CD = entry.ScheduleSourceCode
    rs.Update
    Exit Sub
    
eh:

    TrsRaiseError m_strCMODULENAME, "UpdateTimeEntry"

End Sub

Private Sub DeleteTimeEntry( _
                rs As ADODB.Recordset, _
                entry As cboTimecardEntry)
' Position rs on existing record for entry.
' Mark the record for deletion.
'
On Error GoTo eh
    
    rs.MoveFirst
    rs.Find "INDEX_TIME = '" & entry.IndexTime & "'"

    rs!RECORD_STATUS = "D"
    rs!SCHED_TIME_IND = entry.ScheduleTimeIndicator
    rs!OT_ADJUST_DESC = entry.OTAdjustDesc
    rs.Update
    Exit Sub
    
eh:

    TrsRaiseError m_strCMODULENAME, "DeleteTimeEntry"

End Sub

Private Sub DeleteOtherPayEntry(rsOtherPay As ADODB.Recordset, OtherPayEntry As cboOtherPayEntry)

On Error GoTo DeleteOtherPayEntry_EH
    
    rsOtherPay.MoveFirst
    rsOtherPay.Find "INDEX_TIME = '" & OtherPayEntry.IndexTime & "'"

    rsOtherPay!RECORD_STATUS = "D"
    rsOtherPay.Update
    
Exit Sub
DeleteOtherPayEntry_EH:

    TrsRaiseError m_strCMODULENAME, "DeleteOtherPayEntry"

End Sub

 Private Function Timecard() As cboTimecard
    Set Timecard = mobjTimecard
End Function

Private Sub GetCollectionPeriod( _
                ByVal v_lngEmployeeNumber As Long, _
                ByVal v_datChargeDate As Date, _
                ByRef r_datCollPerStart As Date, _
                ByRef r_datCollPerEnd As Date)
' get the collection period start and end date based on the
' first charge date found for this employee
' therefore, this method will only handle a set of entries
' within the same collection period.
    Dim l_objCollPeriod As busCollectionPeriod.cbsCollectionPeriod
    Dim l_rsCollPeriods As ADODB.Recordset
    
    Set l_objCollPeriod = _
        CreateObject("busCollectionPeriod.cbsCollectionPeriod")
    Set l_rsCollPeriods = _
        l_objCollPeriod.GetCollectionPeriodStatusByEmployee( _
                            v_lngEmployeeNumber, v_datChargeDate, 1)
    
'    this error handling needs work - write a brv
    If l_rsCollPeriods.BOF And l_rsCollPeriods.EOF Then
        Err.Raise 30000 + vbObjectError, , "Collection period not found for charge date"
    Else
        r_datCollPerStart = l_rsCollPeriods!PAY_PER_START_DATE
        r_datCollPerEnd = l_rsCollPeriods!PAY_PER_END_DATE
    End If

    Set l_objCollPeriod = Nothing

End Sub

Private Function RestoreTimecard( _
            r_rsTimecardEntries As ADODB.Recordset) As Boolean
'Private Function RestoreTimecard( _
'            empId As Long, startDt As Date, endDt As Date) As Boolean
   
    Set mobjTimecard = CreateObject("busTimecard.cboTimecard")
    'for older clients (6.0.0) that are not sending this column to the server
    mobjTimecard.PreArrangedOTIndMissing = PreArrangedOTIndMissing

    mobjTimecard.RestoreWithState r_rsTimecardEntries
'    mobjTimecard.Restore empId, startDt, endDt
    mobjTimecard.SetBSTimecard SafeRef(Me)
    
End Function

Private Sub RestoreEmployee( _
                r_rsEmployee As ADODB.Recordset, _
                r_rsEmpWorkgroups As ADODB.Recordset)
'Private Sub RestoreEmployee( _
'                empId As Long, _
'                datCollPerStart As Date, _
'                datCollPerEnd As Date)
                
    Dim mobjEmployee As boCommonLib.cboEmployee
   
    Set mobjEmployee = CreateObject("boCommonLib.cboEmployee")
    
'    mobjEmployee.Restore empId, datCollPerStart, datCollPerEnd
    mobjEmployee.RestoreWithState r_rsEmployee, r_rsEmpWorkgroups
    Timecard.SetEmployee mobjEmployee

End Sub

Private Sub RestoreScheduleWith( _
                r_rsSchedules As ADODB.Recordset, _
                r_rsScheduleChgPmts As ADODB.Recordset)
    Dim mobjSchedule As cboSchedule
    
    Set mobjSchedule = CreateObject("busTimecard.cboSchedule")
    mobjSchedule.RestoreWithState r_rsSchedules, r_rsScheduleChgPmts
    Timecard.SetSchedule mobjSchedule
    
End Sub


Private Sub RestoreSchedule( _
            empId As Long, schedStartDate As Date, schedEndDate As Date)
    Dim mobjSchedule As cboSchedule
    
    Set mobjSchedule = CreateObject("busTimecard.cboSchedule")
    mobjSchedule.Restore empId, schedStartDate, schedEndDate
    Timecard.SetSchedule mobjSchedule
    
End Sub

Private Sub Release()

'Donot enable an error handler since this is called from an error handler.
    
    ReleaseTimecard
    ReleaseOtherPay
    Set mobjDepartmentManager = Nothing
    Set mobjJobManager = Nothing
    Set m_objErrors = Nothing
  
End Sub

Private Sub ReleaseTimecard()

'Donot enable an error handler since this is called from an error handler.
    
    If Not mobjTimecard Is Nothing Then
        mobjTimecard.Release
        Set mobjTimecard = Nothing
    End If
  
End Sub

Private Sub ReleaseOtherPay()

'Donot enable an error handler since this is called from an error handler.
    
    If Not mobjOtherPay Is Nothing Then
        mobjOtherPay.Release
        Set mobjOtherPay = Nothing
    End If
  
End Sub

Public Function CheckReportedHours() As Boolean
'  Will Check Schedule Hours against Reported Hours for each day time exists.
'  If Reported Hours are less than Scheduled Hours Warn User
'  This routine only produces Warnings
'  WHITE104
    Dim l_datChargeDate As Variant
    Dim l_datLastDate As Date
    Dim l_decReportedHours As Variant
    Dim l_decScheduledHours As Variant
    Dim lcolErrDescArgs As New Collection
    Dim lstrErrDesc As String
    
On Error GoTo CheckReportedHours_EH

    
    l_datChargeDate = m_datCollPerStart
    If m_datCollPerEnd > Date Then
        l_datLastDate = Date
    Else
        l_datLastDate = m_datCollPerEnd
    End If
    Do Until l_datChargeDate > l_datLastDate
        l_decReportedHours = Timecard.SumHoursInSchedule(l_datChargeDate)
        l_decScheduledHours = Timecard.Schedule.GetScheduledHours(l_datChargeDate, Now)
        If l_decReportedHours < l_decScheduledHours Then
                lcolErrDescArgs.Add Format(l_datChargeDate, "mm/dd/yyyy")
                lstrErrDesc = _
                    TrsGetErrDescription(LongAppErr(trWarTimecardHoursLessSched), lcolErrDescArgs)
                Set lcolErrDescArgs = Nothing
                AppendError ErrorSet, _
                    Timecard.Employee.EmployeeNumber, _
                    l_datChargeDate, _
                    l_datChargeDate, _
                    "", _
                    "", _
                    CStr(l_decReportedHours), _
                    "", _
                    "", _
                    trWarTimecardHoursLessSched, _
                    mstrCBUSINESSWAR, _
                    lstrErrDesc, _
                    m_strCMODULENAME
        End If
        l_datChargeDate = DateAdd("d", 1, l_datChargeDate)
    Loop
    CheckReportedHours = True
    Exit Function
    
CheckReportedHours_EH:
    TrsRaiseError m_strCMODULENAME, "CheckReportedHours"

End Function


Public Function CheckForOverlappingTime() As Boolean
' This routine will check for overlapping time
' Answer True when there is an overlapping time error; answer
' False when only a warning or no overlapping time.
    Dim aryError(2) As Variant
    Dim aryErrors() As Variant
    Dim intErrors As Integer
    Dim colEntries As Collection
    Dim colCurrentOverlap As Collection
    Dim colTopOverlap As Collection
    Dim lCurrentEntry As cboTimecardEntry
    Dim lOverlapEntry As cboTimecardEntry
    Dim lTopEntry As cboTimecardEntry
    Dim lLowEntry As cboTimecardEntry
    Dim l_Entry As cboTimecardEntry
    Dim llngErrorNumber As Long
    Dim i As Integer
    Dim lstrErrDesc As String
    Dim lstrErrorType As String
    Dim lcolModifiedDates As Collection
    Dim ldatChgDate As Variant
    
On Error GoTo eh
    ReDim aryErrors(1 To 1) As Variant      'declare now to avoid out of range errors

    Set lcolModifiedDates = Timecard.GetChargeDatesWithModifiedTime
    For Each ldatChgDate In lcolModifiedDates
        Set colEntries = Timecard.TimecardEntriesOnChargeDate(ldatChgDate)
        For Each lCurrentEntry In colEntries
            Set colCurrentOverlap = FindOverlappingEntries(lCurrentEntry, Timecard.TimecardEntries)
            If colCurrentOverlap.Count > 0 Then
                For Each lOverlapEntry In colCurrentOverlap
                    FindTopmostEntry lCurrentEntry, lOverlapEntry, lTopEntry, lLowEntry, Timecard.TimecardEntries
                    If IsOverlappingTimeAllowed( _
                       lTopEntry.ExceptionCode, lLowEntry.ExceptionCode, _
                       lTopEntry.OverrideExceptionCode, lLowEntry.OverrideExceptionCode) Then
                        llngErrorNumber = trWarTimecardOverlappingTime
                    Else
                        llngErrorNumber = trErrbusTimecardOverlappingTime
                    End If
                    If Not IsErrorStored(lTopEntry, lLowEntry, aryErrors, llngErrorNumber) Then
                        ' create new error
                        Set aryError(0) = lTopEntry
                        aryError(1) = llngErrorNumber
                        If lTopEntry Is lOverlapEntry Then
                        'if topmost is overlaping;
                        ' create new collection and add lCurrentEntry when first encounter;
                        ' otherwise add lCurrentEntry to overlapping collection avoiding duplicates.
                            Set colTopOverlap = _
                                TopmostOverlappingEntries(lTopEntry, aryErrors)
                            If colTopOverlap Is Nothing Then
                                Set colTopOverlap = New Collection
                                colTopOverlap.Add lCurrentEntry
                                Set aryError(2) = colTopOverlap
                            Else
                                If Not IsIncluded(lCurrentEntry, colTopOverlap) Then
                                    colTopOverlap.Add lCurrentEntry
                                End If
                            End If
                        Else
                            'use current entry overlapping collection
                            Set aryError(2) = colCurrentOverlap
                        End If
                        'add new error to errors array
                        If intErrors > 0 Then   'first error already has empty slot available
                            ReDim Preserve aryErrors(1 To intErrors + 1) As Variant
                        End If
                        intErrors = intErrors + 1
                        aryErrors(intErrors) = aryError
                    Else
                        'is already stored
                        'ensure overlapping entry is stored for topmost entry
                        If lTopEntry Is lOverlapEntry Then
                            Set colTopOverlap = _
                                TopmostOverlappingEntries(lTopEntry, aryErrors)
                            'colTopOverlap will be nothing when the top was already stored
                            ' as overlapping; in this case, nothing more to do.
                            If Not colTopOverlap Is Nothing Then
                                If Not IsIncluded(lCurrentEntry, colTopOverlap) Then
                                    colTopOverlap.Add lCurrentEntry
                                End If
                            End If
                        End If
                    End If
                Next
            End If
        Next
    Next

    CheckForOverlappingTime = False
' append all errors stored to error set.
    For i = LBound(aryErrors) To UBound(aryErrors)
        If Not IsEmpty(aryErrors(i)) Then
            Set l_Entry = aryErrors(i)(0)
            llngErrorNumber = aryErrors(i)(1)
            If llngErrorNumber = trWarTimecardOverlappingTime Then
                lstrErrorType = mstrCBUSINESSWAR
            Else
                lstrErrorType = mstrCBUSINESSERR
                CheckForOverlappingTime = True
            End If
            lstrErrDesc = _
                TrsGetErrDescription(LongAppErr(llngErrorNumber), Nothing)
            AppendError ErrorSet, l_Entry.EmployeeNumber, l_Entry.ChargeDate, _
                l_Entry.DateWorked, l_Entry.startTime, l_Entry.EndTime, _
                CStr(l_Entry.Hours), l_Entry.ExceptionCode, _
                l_Entry.OverrideExceptionCode, llngErrorNumber, _
                lstrErrorType, lstrErrDesc, m_strCMODULENAME
        End If
    Next
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "CheckForOverlappingTime"

End Function

Private Function CheckShiftReportedHours() As Boolean
'  Will Check Schedule Hours against Reported Hours for each shift time exists.
'  If Reported Hours are greater than Scheduled Hours, report error.
    Dim ldatChargeDate As Date
    Dim ldatLastDate As Date
    Dim ldecReportedHours As Variant
    Dim lcolSchedShifts As Collection
    Dim eShift As cboScheduledShift
    Dim lcolEntries As Collection
    Dim lcolErrDescArgs As New Collection
    Dim lstrErrDesc As String
    Dim lblnError As Boolean
    Dim l_datDatePlus1 As Date
    Dim eEntry As cboTimecardEntry
    Dim ldecReportedParentalBondingHours As Variant
    
    
On Error GoTo eh

    ldatChargeDate = m_datCollPerStart
    l_datDatePlus1 = DateAdd("d", 1, Date)
    If m_datCollPerEnd > l_datDatePlus1 Then
        ldatLastDate = l_datDatePlus1
    Else
        ldatLastDate = m_datCollPerEnd
    End If
    Do Until ldatChargeDate > ldatLastDate
        Set lcolSchedShifts = Timecard.Schedule.ScheduledShiftsOn(ldatChargeDate)
        Set lcolEntries = Timecard.TimecardEntriesOnChargeDate(ldatChargeDate)
        For Each eShift In lcolSchedShifts  'assign entries to shift
            ldecReportedHours = SumReportedHoursByShift(eShift, lcolEntries)
            If ldecReportedHours > eShift.ShiftHours Then
                lcolErrDescArgs.Add Format(ldatChargeDate, "mm/dd/yyyy")
                lstrErrDesc = _
                    TrsGetErrDescription(LongAppErr(trErrTimecardHoursMoreSched), lcolErrDescArgs)
                Set lcolErrDescArgs = Nothing
                AppendError ErrorSet, _
                    Timecard.Employee.EmployeeNumber, _
                    ldatChargeDate, ldatChargeDate, _
                    "", "", CStr(ldecReportedHours), _
                    "", "", _
                    trErrTimecardHoursMoreSched, mstrCBUSINESSERR, _
                    lstrErrDesc, m_strCMODULENAME
                lblnError = True
            Else
                If ldecReportedHours <= eShift.ShiftHours Then
                    ldecReportedParentalBondingHours = SumReportedParentalBondingHoursByShift(eShift, lcolEntries)
                    If ldecReportedParentalBondingHours <> 0 And ldecReportedParentalBondingHours < eShift.ShiftHours Then
                    
                    
                        lcolErrDescArgs.Add Format(ldatChargeDate, "mm/dd/yyyy")
                        lstrErrDesc = "Parental Bonding hours must entered as a full shift."
                        Set lcolErrDescArgs = Nothing
                        AppendError ErrorSet, _
                            Timecard.Employee.EmployeeNumber, _
                            ldatChargeDate, ldatChargeDate, _
                            "", "", CStr(ldecReportedParentalBondingHours), _
                            "", "", _
                            vbObjectError + 30000, mstrCBUSINESSERR, _
                            lstrErrDesc, m_strCMODULENAME
                        lblnError = True
                    End If
                End If
            End If
        Next
        ldatChargeDate = DateAdd("d", 1, ldatChargeDate)
    Loop
        
    Set eShift = Nothing
    Set lcolSchedShifts = Nothing
    Set lcolEntries = Nothing
    CheckShiftReportedHours = lblnError
    Exit Function
    
eh:
    Set eShift = Nothing
    Set lcolSchedShifts = Nothing
    Set lcolEntries = Nothing
    TrsRaiseError m_strCMODULENAME, "CheckShiftReportedHours"

End Function

Private Function SumReportedHoursByShift( _
                    ByRef robjSchedShift As cboScheduledShift, _
                    ByRef rcolEntries As Collection) As Variant
    Dim eEntry As cboTimecardEntry
    Dim ldecShiftHours As Variant
    Dim lcolShiftEntries As Collection
    
On Error GoTo eh

    Set lcolShiftEntries = _
        GetEntriesWithinShift(robjSchedShift, rcolEntries)
    ldecShiftHours = CDec(0)
    For Each eEntry In lcolShiftEntries
        If eEntry.OvertimeOpportunityNumber = 0 And _
            eEntry.ExceptionCode <> "ONP" And _
            eEntry.ExceptionCode <> "LU" And _
            eEntry.ExceptionCode <> "LP" And _
            eEntry.ExceptionCode <> "HB" And _
            Not IsOverrideThatAllowsOverlap(eEntry.OverrideExceptionCode) Then
            ldecShiftHours = ldecShiftHours + eEntry.Hours
        End If
    Next

    SumReportedHoursByShift = ldecShiftHours
    Set eEntry = Nothing
    Set lcolShiftEntries = Nothing
    Exit Function
    
eh:
    Set eEntry = Nothing
    Set lcolShiftEntries = Nothing
    TrsRaiseError m_strCMODULENAME, "SumReportedHoursByShift"

End Function

Private Function GetEntriesWithinShift( _
                    ByRef robjSchedShift As cboScheduledShift, _
                    ByRef rcolEntries As Collection) As Collection
    Dim lcolEntriesThisShift As Collection
    Dim eEntry As cboTimecardEntry

On Error GoTo eh
    Set lcolEntriesThisShift = New Collection
    For Each eEntry In rcolEntries
        If robjSchedShift.IsIntervalWithinShift( _
            eEntry.StartTimeAsDate, eEntry.EndTimeAsDate) Then
            lcolEntriesThisShift.Add eEntry
        End If
    Next
    Set GetEntriesWithinShift = lcolEntriesThisShift
    Set lcolEntriesThisShift = Nothing
    Set eEntry = Nothing
    Exit Function
    
eh:
    Set lcolEntriesThisShift = Nothing
    Set eEntry = Nothing
    TrsRaiseError m_strCMODULENAME, "GetEntriesWithinShift"

End Function


Private Function ValidateCanNonCanOvertime() As Boolean
' Check for any modified canvass or noncanvass overtime.
' Split time entries are allowed, however changing the start time,
' end time, or total hours for an opportunity is not allowed.
' Answer True when any/all changes are allowed, otherwise False.
    Dim lvarOrigOTOppNumDetails As Variant
    Dim lvarCurOTOppNumDetails As Variant
    Dim intOrigIndex As Integer
    Dim intCurIndex As Integer
    Dim llngOrigOTOppNum As Long
    Dim llngOTOppNum As Long
    Dim lcolEntries As Collection
    Dim lEntry As cboTimecardEntry
    Dim llngErrorNumber As Long
    Dim lstrErrDesc As String

On Error GoTo eh

    lvarOrigOTOppNumDetails = Timecard.GetOriginalOvertimeOpportunities
    lvarCurOTOppNumDetails = Timecard.GetAllOvertimeOpportunityDetails
    If IsEmpty(lvarOrigOTOppNumDetails) Then
        If IsEmpty(lvarCurOTOppNumDetails) Then
            ValidateCanNonCanOvertime = True
            Exit Function
        Else
            'write error - cannot add canvass/noncanvass overtime
            ' in time collection, just get the first entry you find.
            llngOTOppNum = lvarCurOTOppNumDetails(1)(1)
            Set lcolEntries = Timecard.TimecardEntriesWithOpportunity(llngOTOppNum)
            If lcolEntries.Count > 0 Then
                llngErrorNumber = trErrAddOTCanvass
                lstrErrDesc = _
                    TrsGetErrDescription(LongAppErr(llngErrorNumber), Nothing)
                Set lEntry = lcolEntries(1)
                AppendError ErrorSet, lEntry.EmployeeNumber, lEntry.ChargeDate, _
                    lEntry.DateWorked, lEntry.startTime, lEntry.EndTime, _
                    CStr(lEntry.Hours), lEntry.ExceptionCode, _
                    lEntry.OverrideExceptionCode, llngErrorNumber, _
                    mstrCBUSINESSERR, lstrErrDesc, m_strCMODULENAME
                ValidateCanNonCanOvertime = False
                Exit Function
            End If
        End If
    Else
        If IsEmpty(lvarCurOTOppNumDetails) Then
            'write error - cannot delete canvass/noncanvass overtime
            ' in time collection
            llngOTOppNum = lvarOrigOTOppNumDetails(1)(1)
            Set lcolEntries = _
                Timecard.TimecardEntriesWithOpportunity(llngOTOppNum, True)
            If lcolEntries.Count > 0 Then
                llngErrorNumber = trErrDeleteOTCanvass
                lstrErrDesc = _
                    TrsGetErrDescription(LongAppErr(llngErrorNumber), Nothing)
                Set lEntry = lcolEntries(1)
                AppendError ErrorSet, lEntry.EmployeeNumber, lEntry.ChargeDate, _
                    lEntry.DateWorked, lEntry.startTime, lEntry.EndTime, _
                    CStr(lEntry.Hours), lEntry.ExceptionCode, _
                    lEntry.OverrideExceptionCode, llngErrorNumber, _
                    mstrCBUSINESSERR, lstrErrDesc, m_strCMODULENAME
                ValidateCanNonCanOvertime = False
                Exit Function
            End If
        End If
    End If

'   check each opportunity in the original array for a matching
'   array in the current array.
    For intOrigIndex = LBound(lvarOrigOTOppNumDetails) To UBound(lvarOrigOTOppNumDetails)
        llngOrigOTOppNum = lvarOrigOTOppNumDetails(intOrigIndex)(1)
        ' position on matching current ot opportunity details
        For intCurIndex = LBound(lvarCurOTOppNumDetails) To UBound(lvarCurOTOppNumDetails)
            If llngOrigOTOppNum = lvarCurOTOppNumDetails(intCurIndex)(1) Then
                Exit For
            End If
        Next
        ValidateCanNonCanOvertime = True
        If intCurIndex > UBound(lvarCurOTOppNumDetails) Then
            'error no matching ot opportunity
            llngOTOppNum = llngOrigOTOppNum
            Set lcolEntries = _
                Timecard.TimecardEntriesWithOpportunity(llngOTOppNum, True)
            If lcolEntries.Count > 0 Then
                llngErrorNumber = trErrDeleteOTCanvass
                lstrErrDesc = _
                    TrsGetErrDescription(LongAppErr(llngErrorNumber), Nothing)
                Set lEntry = lcolEntries(1)
                AppendError ErrorSet, lEntry.EmployeeNumber, lEntry.ChargeDate, _
                    lEntry.DateWorked, lEntry.startTime, lEntry.EndTime, _
                    CStr(lEntry.Hours), lEntry.ExceptionCode, _
                    lEntry.OverrideExceptionCode, llngErrorNumber, _
                    mstrCBUSINESSERR, lstrErrDesc, m_strCMODULENAME
                ValidateCanNonCanOvertime = False
                Exit Function
            End If
        End If
        If Not lvarOrigOTOppNumDetails(intOrigIndex)(2) = lvarCurOTOppNumDetails(intCurIndex)(2) Then
            'error start time change
            ValidateCanNonCanOvertime = False
        End If
        If Not lvarOrigOTOppNumDetails(intOrigIndex)(3) = lvarCurOTOppNumDetails(intCurIndex)(3) Then
            'error end time change
            ValidateCanNonCanOvertime = False
        End If
        If Not lvarOrigOTOppNumDetails(intOrigIndex)(4) = lvarCurOTOppNumDetails(intCurIndex)(4) Then
            'error hours change
            ValidateCanNonCanOvertime = False
        End If
        If Not ValidateCanNonCanOvertime Then
            llngOTOppNum = llngOrigOTOppNum
            Set lcolEntries = _
                Timecard.TimecardEntriesWithOpportunity(llngOTOppNum, False)
            If lcolEntries.Count > 0 Then
                llngErrorNumber = trErrChangeOTCanvass
                lstrErrDesc = _
                    TrsGetErrDescription(LongAppErr(llngErrorNumber), Nothing)
                Set lEntry = lcolEntries(1)
                AppendError ErrorSet, lEntry.EmployeeNumber, lEntry.ChargeDate, _
                    lEntry.DateWorked, lEntry.startTime, lEntry.EndTime, _
                    CStr(lEntry.Hours), lEntry.ExceptionCode, _
                    lEntry.OverrideExceptionCode, llngErrorNumber, _
                    mstrCBUSINESSERR, lstrErrDesc, m_strCMODULENAME
                Exit Function
            End If
        End If
    Next

    ValidateCanNonCanOvertime = True
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "ValidateCanNonCanOvertime"
    
End Function


Private Sub ValidateTTA()
' Check each new or modified TTA for associated time entry
' on the charge date.  If there are not any reported hours
' on the charge date, append a warning brv.
    Dim l_colChgDates As Collection
    Dim l_datChargeDate As Variant
    Dim l_colTimeEntries As Collection
    Dim lcolErrDescArgs As Collection
    Dim lstrErrDesc As String
    Dim lobjOtherPayEntry As cboOtherPayEntry
    
On Error GoTo eh

    Set l_colChgDates = GetChargeDatesWithModifiedTTA
    For Each l_datChargeDate In l_colChgDates
        If Timecard.TimecardEntriesOnChargeDate(l_datChargeDate).Count = 0 Then
            Set lobjOtherPayEntry = FindFirstModifiedTTAOn(l_datChargeDate)
            Set lcolErrDescArgs = New Collection
            lcolErrDescArgs.Add Format(l_datChargeDate, "mm/dd/yyyy")
            lstrErrDesc = _
                TrsGetErrDescription(LongAppErr(trWarTTAWithoutTimeEntry), lcolErrDescArgs)
            Set lcolErrDescArgs = Nothing
            AppendError ErrorSet, _
                OtherPay.Employee.EmployeeNumber, _
                l_datChargeDate, l_datChargeDate, _
                lobjOtherPayEntry.startTime, _
                lobjOtherPayEntry.EndTime, _
                CStr(lobjOtherPayEntry.hrs), _
                lobjOtherPayEntry.ExceptionCode, _
                "", _
                trWarTTAWithoutTimeEntry, mstrCBUSINESSWAR, _
                lstrErrDesc, m_strCMODULENAME
        End If
    Next
    
    Set l_colChgDates = Nothing
    Set l_colTimeEntries = Nothing
    Set lcolErrDescArgs = Nothing
    Exit Sub
    
eh:
    Set l_colChgDates = Nothing
    Set l_colTimeEntries = Nothing
    Set lcolErrDescArgs = Nothing
    TrsRaiseError m_strCMODULENAME, "ValidateTTA"
End Sub


Private Function GetChargeDatesWithModifiedTTA() As Collection
' Answer a collection of dates that represents the charge dates
' that TTA have been added or changed.
    Dim l_colChgDatesWithModTime As cuSet
    Dim l_objOtherPayEntry As cboOtherPayEntry

On Error GoTo eh
    If OtherPay Is Nothing Then
        Set GetChargeDatesWithModifiedTTA = New Collection
        Exit Function
    End If

    Set l_colChgDatesWithModTime = New cuSet
    For Each l_objOtherPayEntry In OtherPay.OtherPayEntriesPrivate
        If Trim(l_objOtherPayEntry.ExceptionCode) = "TTA" Then
            If l_objOtherPayEntry.IsNew Or l_objOtherPayEntry.IsModified Then
                l_colChgDatesWithModTime.Add l_objOtherPayEntry.ChargeDate
            End If
        End If
    Next

    Set GetChargeDatesWithModifiedTTA = _
            l_colChgDatesWithModTime.Contents
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "GetChargeDatesWithModifiedTTA"

End Function

Private Function FindFirstModifiedTTAOn( _
            ByVal vdatChargeDate As Date) As cboOtherPayEntry
' Answer an cboOtherPayEntry.
    Dim l_objOtherPayEntry As cboOtherPayEntry

On Error GoTo eh
    If OtherPay Is Nothing Then
        Set FindFirstModifiedTTAOn = Nothing
        Exit Function
    End If

    For Each l_objOtherPayEntry In OtherPay.OtherPayEntriesPrivate
        If Trim(l_objOtherPayEntry.ExceptionCode) = "TTA" Then
            If l_objOtherPayEntry.IsNew Or l_objOtherPayEntry.IsModified Then
                If l_objOtherPayEntry.ChargeDate = vdatChargeDate Then
                    Set FindFirstModifiedTTAOn = l_objOtherPayEntry
                    Exit Function
                End If
            End If
        End If
    Next
    Set FindFirstModifiedTTAOn = Nothing
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "FindFirstModifiedTTAOn"

End Function


Private Sub FindTopmostEntry( _
                ByRef rCurrentEntry As cboTimecardEntry, _
                ByRef rOverlapEntry As cboTimecardEntry, _
                ByRef rTopmostEntry As cboTimecardEntry, _
                ByRef rLowerEntry As cboTimecardEntry, _
                ByRef rcolEntries As Collection)
'First rule, modified entry is the topmost.
'Second rule, overtime is the topmost.
'Third rule, entry with earlier start.
'Fourth rule, entry with greatest hours.
    Dim lcolOverlapEntries As Collection

On Error GoTo eh

    If rCurrentEntry.IsNew Or rCurrentEntry.IsModified Then
        If Not (rOverlapEntry.IsNew Or rOverlapEntry.IsModified) Then
            Set rTopmostEntry = rCurrentEntry
            Set rLowerEntry = rOverlapEntry
            Exit Sub
        End If
    ElseIf rOverlapEntry.IsNew Or rOverlapEntry.IsModified Then
        Set rTopmostEntry = rOverlapEntry
        Set rLowerEntry = rCurrentEntry
        Exit Sub
    End If
    
    If IsExceptionOvertimeRelated(rCurrentEntry.ExceptionCode) Then      'needs to overtime related
        If Not IsExceptionOvertimeRelated(rOverlapEntry.ExceptionCode) Then
            Set rTopmostEntry = rCurrentEntry
            Set rLowerEntry = rOverlapEntry
            Exit Sub
        End If
    ElseIf IsExceptionOvertimeRelated(rOverlapEntry.ExceptionCode) Then
        Set rTopmostEntry = rOverlapEntry
        Set rLowerEntry = rCurrentEntry
        Exit Sub
    End If
    
    If Not rCurrentEntry.StartTimeAsDate = rOverlapEntry.StartTimeAsDate Then
        If rCurrentEntry.StartTimeAsDate < rOverlapEntry.StartTimeAsDate Then
            'override starttime rule when the overlapping entry
            'end time is greater than the current end time, and
            'overlapEntry overlaps more than the current entry.
            If rOverlapEntry.EndTime > rCurrentEntry.EndTime Then
                Set lcolOverlapEntries = _
                    FindOverlappingEntries(rOverlapEntry, rcolEntries)
                If lcolOverlapEntries.Count > 1 Then
                    Set rTopmostEntry = rOverlapEntry
                    Set rLowerEntry = rCurrentEntry
                Else
                    Set rTopmostEntry = rCurrentEntry
                    Set rLowerEntry = rOverlapEntry
                End If
            Else
                Set rTopmostEntry = rCurrentEntry
                Set rLowerEntry = rOverlapEntry
            End If
        Else
            Set rTopmostEntry = rOverlapEntry
            Set rLowerEntry = rCurrentEntry
        End If
        Exit Sub
    End If

    If Not rCurrentEntry.Hours = rOverlapEntry.Hours Then
        If rCurrentEntry.Hours > rOverlapEntry.Hours Then
            Set rTopmostEntry = rCurrentEntry
            Set rLowerEntry = rOverlapEntry
        Else
            Set rTopmostEntry = rOverlapEntry
            Set rLowerEntry = rCurrentEntry
        End If
        Exit Sub
    Else
        'default, assign current to top
        Set rTopmostEntry = rCurrentEntry
        Set rLowerEntry = rOverlapEntry
    End If
    Exit Sub
    
eh:
    TrsRaiseError m_strCMODULENAME, "FindTopmostEntry"

End Sub

Private Function IsErrorStored( _
            ByRef rTopEntry As cboTimecardEntry, _
            ByRef rLowEntry As cboTimecardEntry, _
            ByRef rErrors() As Variant, _
            ByVal vlngErrorNumber As Long) As Boolean
'Need the error to determine if duplicate, because
' offending entry could have a warning and an error with another entry.

    Dim lOverlapEntry As cboTimecardEntry
    Dim i As Integer

On Error GoTo eh
' Determine if topmost entry is already stored
    For i = LBound(rErrors) To UBound(rErrors)
        If Not IsEmpty(rErrors(i)) Then
            If rErrors(i)(0) Is rTopEntry _
                And rErrors(i)(1) = vlngErrorNumber Then
                IsErrorStored = True
                Exit Function
            End If
        End If
    Next
    ' if topmost entry not stored yet, check
    ' overlapping entries.
    For i = LBound(rErrors) To UBound(rErrors)
        If Not IsEmpty(rErrors(i)) Then
            If rErrors(i)(0) Is rLowEntry Then
                'check overlapping entries for offender
                For Each lOverlapEntry In rErrors(i)(2)
                    If lOverlapEntry Is rTopEntry _
                        And rErrors(i)(1) = vlngErrorNumber Then
                        IsErrorStored = True
                        Exit Function
                    End If
                Next
            End If
        End If
    Next
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "IsErrorStored"

End Function


Private Function TopmostOverlappingEntries( _
            ByRef rTopEntry As cboTimecardEntry, _
            ByRef rErrors() As Variant) As Collection
            
    Dim i As Integer
    
    On Error GoTo eh

    For i = LBound(rErrors) To UBound(rErrors)
        If Not IsEmpty(rErrors(i)) Then     'guard against empty slot
            If rErrors(i)(0) Is rTopEntry Then
                Set TopmostOverlappingEntries = rErrors(i)(2)
                Exit Function
            End If
        End If
    Next
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "TopmostOverlappingEntries"
        
End Function


Private Function IsOverlappingTimeAllowed( _
            ByVal v_strExcCode1 As String, _
            ByVal v_strExcCode2 As String, _
            ByVal v_strOverExcCode1 As String, _
            ByVal v_strOverExcCode2 As String _
            ) As Boolean
            
' The system allows duplicate time if the exception codes of the
' 2 pieces of time are different

' The system will allow duplicate time for the same exception code
' if either of the 2 pieces of time has the call out stand by (CS)
' override execption code
' JimB 03/10/2005
'Add First Responder (FR) override exception code and it behaves like CS

On Error GoTo eh
    If Trim(v_strExcCode1) = Trim(v_strExcCode2) Then
        If IsOverrideThatAllowsOverlap(v_strOverExcCode1) _
            Or IsOverrideThatAllowsOverlap(v_strOverExcCode2) Then
            IsOverlappingTimeAllowed = True
        Else
            IsOverlappingTimeAllowed = False
        End If
    Else
        If Trim(v_strExcCode1) = "HB" Or Trim(v_strExcCode2) = "HB" Then
            If Trim(v_strExcCode1) = "ST" Or Trim(v_strExcCode2) = "ST" _
              Or IsExceptionOvertimeRelated(v_strExcCode1) Or IsExceptionOvertimeRelated(v_strExcCode2) _
              Or Trim(v_strExcCode1) = "ST5" Or Trim(v_strExcCode2) = "ST5" Then
                IsOverlappingTimeAllowed = True
            Else
                IsOverlappingTimeAllowed = False
            End If
        Else
            IsOverlappingTimeAllowed = True
        End If
    End If
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "IsOverlappingTimeAllowed"

End Function

Public Function FindOverlappingEntries( _
            ByRef v_Entry As cboTimecardEntry, _
            ByRef r_colEntries As Collection _
            ) As Collection
'Release 07/25/2001 -- JFB  Changed method from private to public
'                            so it can be called from cboTimeCard


    Dim l_OverlapEntry As cboTimecardEntry
    Dim l_colEntries As Collection
    Dim l_blnOverlapFound As Boolean
    
On Error GoTo eh
    Set l_colEntries = New Collection
    l_blnOverlapFound = False
    For Each l_OverlapEntry In r_colEntries
        If Not l_OverlapEntry Is v_Entry Then
            If TimePeriodsIntersect(v_Entry.StartTimeAsDate, v_Entry.EndTimeAsDate, l_OverlapEntry.StartTimeAsDate, l_OverlapEntry.EndTimeAsDate) Then
                l_colEntries.Add l_OverlapEntry
                l_blnOverlapFound = True
            End If
        End If
    Next
    
    Set FindOverlappingEntries = l_colEntries
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "FindOverlappingEntries"

End Function

Public Function TimeEntriesByStartTime(x As Collection) As Collection
'
'Replace with ordered collection
'
    Dim tmpCol As Collection, teColCopy As Collection
    Dim te As cboTimecardEntry
    Dim te2 As cboTimecardEntry
    Dim lowTimeEntry As cboTimecardEntry
    Dim lowStartTime As Date, teStartTime As Date
    Dim i As Integer, lowEntryIndex As Integer
    Set tmpCol = New Collection
    Set teColCopy = CopyObjects(x)
    While teColCopy.Count > 0
        'seed with a higher date than possible
        lowStartTime = CDate(Format(#6/23/2100#, "mm/dd/yyyy") & " " & "07:30")
        For i = 1 To teColCopy.Count
            Set te = teColCopy.Item(i)
            teStartTime = te.StartTimeAsDate
            If teStartTime <= lowStartTime Then
                Set lowTimeEntry = te
                lowStartTime = teStartTime
                lowEntryIndex = i
            End If
        Next
        tmpCol.Add lowTimeEntry
        teColCopy.Remove lowEntryIndex
    Wend
    Set TimeEntriesByStartTime = tmpCol
    
End Function


Private Function BuildOvertimeBankUpdates() As ADODB.Recordset

On Error GoTo eh
    Set BuildOvertimeBankUpdates = Timecard.BuildOvertimeBankUpdates
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "BuildOvertimeBankUpdates"

End Function

Private Function BuildEMPRVInterface() As Collection

On Error GoTo eh
    Set BuildEMPRVInterface = Timecard.BuildEMPRVInterface
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "BuildEMPRVInterface"

End Function
Private Function BuildEMPRVJobTasks() As ADODB.Recordset

On Error GoTo eh
    Set BuildEMPRVJobTasks = Timecard.BuildEMPRVJobTasks
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "BuildEMPRVJobTasks"

End Function

Public Function GetViewTimeSummaryByEmployee( _
        ByVal v_aryEmployees As Variant, _
        ByVal v_datStartDate As Date, _
        ByVal v_datEndDate As Date) As ADODB.Recordset

    Dim l_objCtx As ObjectContext
    Dim l_rsSummaryData As ADODB.Recordset
    Dim l_rsEmployeeData As ADODB.Recordset

On Error GoTo GetViewTimeSummaryByEmployee_EH

    Set l_objCtx = GetObjectContext
    
    If UBound(v_aryEmployees) > 444 Then 'maximum number of employees reached for view
        Err.Raise trErrMaxNoEmployeeLimit444
    End If

    Set l_rsEmployeeData = GetViewTimeEmployeeData(v_aryEmployees)
    If l_rsEmployeeData.EOF And l_rsEmployeeData.BOF Then
        Set GetViewTimeSummaryByEmployee = Nothing
        l_objCtx.SetComplete
        Exit Function
    End If
    
    Set l_rsSummaryData = GetViewTimeSummaryData(v_aryEmployees, v_datStartDate, v_datEndDate)
    Set GetViewTimeSummaryByEmployee = BuildViewTimeSummary(l_rsSummaryData)
    l_objCtx.SetComplete

Exit Function
GetViewTimeSummaryByEmployee_EH:
    
    l_objCtx.SetAbort
    
    Select Case Err.Number
        Case trErrMaxNoEmployeeLimit444
             Err.Source = "1;;;;;;;;"
             TrsRaiseError m_strCMODULENAME, "GetViewTimeSummaryByEmployee", trErrMaxNoEmployeeLimit444
        Case Else
             TrsRaiseError m_strCMODULENAME, "GetViewTimeSummaryByEmployee"
    End Select
     

End Function

Private Function GetViewTimeSummaryData( _
            ByVal v_aryEmployees As Variant, _
            ByVal v_datStartDate As Date, _
            ByVal v_datEndDate As Date _
            ) As ADODB.Recordset

    Dim l_objCtx As ObjectContext
    Dim l_objdbTimecard As dbTimeCard.cdTimeCard
    Dim l_rsSummaryData As ADODB.Recordset

On Error GoTo GetViewTimeSummaryData_EH

    Set l_objCtx = GetObjectContext
    Set l_objdbTimecard = CreateObject("dbTimeCard.cdTimeCard")

    Set GetViewTimeSummaryData = l_objdbTimecard.GetTimeSummaryByEmployee(v_aryEmployees, v_datStartDate, v_datEndDate)

    Set l_objdbTimecard = Nothing
    Set l_objCtx = Nothing

Exit Function
GetViewTimeSummaryData_EH:
    l_objCtx.SetAbort
    Set l_objdbTimecard = Nothing
    Set l_objCtx = Nothing

    TrsRaiseError m_strCMODULENAME, "GetViewTimeSummaryData"

End Function



Private Function GetViewTimeEmployeeData( _
            ByVal v_aryEmployees As Variant _
            ) As ADODB.Recordset

    Dim l_objCtx As ObjectContext
    Dim l_objdbEmployee As dbEmployee.cdEmployee
    Dim l_rsSummaryData As ADODB.Recordset

On Error GoTo GetViewTimeEmployeeData_EH

    Set l_objCtx = GetObjectContext
    Set l_objdbEmployee = CreateObject("dbEmployee.cdEmployee")

    Set GetViewTimeEmployeeData = l_objdbEmployee.GetEmployee(v_aryEmployees, trGetBasicEmployee)

    Set l_objdbEmployee = Nothing
    Set l_objCtx = Nothing

Exit Function
GetViewTimeEmployeeData_EH:
    l_objCtx.SetAbort
    Set l_objdbEmployee = Nothing
    Set l_objCtx = Nothing

    TrsRaiseError m_strCMODULENAME, "GetViewTimeEmployeeData"

End Function

Private Function BuildViewTimeSummary( _
            ByVal v_rsSummaryData As ADODB.Recordset _
            ) As ADODB.Recordset

    Dim l_rsViewTimeSummary As ADODB.Recordset

    On Error GoTo BuildViewTimeSummary_EH

    Set l_rsViewTimeSummary = CreateViewTimeSummaryRecordSet

    If v_rsSummaryData.EOF And v_rsSummaryData.BOF Then
        Set BuildViewTimeSummary = Nothing
        Exit Function
    End If

    SetViewTimeSummaryFields l_rsViewTimeSummary, v_rsSummaryData

    Set BuildViewTimeSummary = l_rsViewTimeSummary

    Exit Function

BuildViewTimeSummary_EH:
        Set BuildViewTimeSummary = Nothing
        TrsRaiseError m_strCMODULENAME, "BuildViewTimeSummary"

End Function




Private Function CreateViewTimeSummaryRecordSet(Optional ByVal v_ViewOption As TRSViewOption) As ADODB.Recordset

    Dim l_rsViewTimeSummary As ADODB.Recordset

    On Error GoTo CreateViewTimeSummaryRecordSet_EH

    Set l_rsViewTimeSummary = New ADODB.Recordset

    With l_rsViewTimeSummary
      .Fields.Append "EMP_NUM", adInteger, , adFldUpdatable
      .Fields.Append "EMP_NAME", adChar, 30, adFldUpdatable
      .Fields.Append "ST_HRS_UNAPPROVED", adDecimal, , adFldUpdatable
      .Fields.Item("ST_HRS_UNAPPROVED").Precision = 7
      .Fields.Item("ST_HRS_UNAPPROVED").NumericScale = 1
      .Fields.Append "OT_HRS_UNAPPROVED", adDecimal, , adFldUpdatable
      .Fields.Item("OT_HRS_UNAPPROVED").Precision = 7
      .Fields.Item("OT_HRS_UNAPPROVED").NumericScale = 1
      .Fields.Append "SYM_HRS_UNAPPROVED", adDecimal, , adFldUpdatable
      .Fields.Item("SYM_HRS_UNAPPROVED").Precision = 7
      .Fields.Item("SYM_HRS_UNAPPROVED").NumericScale = 1
      .Fields.Append "TOT_HRS_UNAPPROVED", adDecimal, , adFldUpdatable
      .Fields.Item("TOT_HRS_UNAPPROVED").Precision = 9
      .Fields.Item("TOT_HRS_UNAPPROVED").NumericScale = 1
      .Fields.Append "ST_HRS_APPROVED", adDecimal, , adFldUpdatable
      .Fields.Item("ST_HRS_APPROVED").Precision = 7
      .Fields.Item("ST_HRS_APPROVED").NumericScale = 1
      .Fields.Append "OT_HRS_APPROVED", adDecimal, , adFldUpdatable
      .Fields.Item("OT_HRS_APPROVED").Precision = 7
      .Fields.Item("OT_HRS_APPROVED").NumericScale = 1
      .Fields.Append "SYM_HRS_APPROVED", adDecimal, , adFldUpdatable
      .Fields.Item("SYM_HRS_APPROVED").Precision = 7
      .Fields.Item("SYM_HRS_APPROVED").NumericScale = 1
      .Fields.Append "TOT_HRS_APPROVED", adDecimal, , adFldUpdatable
      .Fields.Item("TOT_HRS_APPROVED").Precision = 9
      .Fields.Item("TOT_HRS_APPROVED").NumericScale = 1
      .Fields.Append "OTHER_PAY_IND_UNAPPROVED", adChar, 1, adFldUpdatable
      .Fields.Append "OTHER_PAY_IND_APPROVED", adChar, 1, adFldUpdatable
      If v_ViewOption = TrDept Then 'use department
        .Fields.Append "DEPT_NUM", adChar, 3, adFldUpdatable
      Else
          If v_ViewOption = TrWG Then 'use WG
            .Fields.Append "WG_NUM", adInteger, adFldUpdatable
          Else
          End If
      End If
      .CursorLocation = adUseClient
      .CursorType = adOpenDynamic
      .LockType = adLockOptimistic
      
      .Open
    End With

    Set CreateViewTimeSummaryRecordSet = l_rsViewTimeSummary

    Exit Function

CreateViewTimeSummaryRecordSet_EH:

    Set CreateViewTimeSummaryRecordSet = Nothing

    TrsRaiseError m_strCMODULENAME, "CreateViewTimeSummaryRecordSet"

End Function
Private Sub InitailizeViewTimeSummaryFields( _
                ByRef r_rsViewTimeSummary As ADODB.Recordset, _
                ByVal v_rsEmployeeData As ADODB.Recordset, _
                Optional ByVal v_ViewOption As TRSViewOption)

    On Error GoTo InitailizeViewTimeSummaryFields_EH

    r_rsViewTimeSummary!EMP_NUM = v_rsEmployeeData!EMP_NUM
    r_rsViewTimeSummary!EMP_NAME = v_rsEmployeeData!EMP_NAME
    r_rsViewTimeSummary!ST_HRS_UNAPPROVED = 0
    r_rsViewTimeSummary!OT_HRS_UNAPPROVED = 0
    r_rsViewTimeSummary!SYM_HRS_UNAPPROVED = 0
    r_rsViewTimeSummary!TOT_HRS_UNAPPROVED = 0
    r_rsViewTimeSummary!ST_HRS_APPROVED = 0
    r_rsViewTimeSummary!OT_HRS_APPROVED = 0
    r_rsViewTimeSummary!SYM_HRS_APPROVED = 0
    r_rsViewTimeSummary!TOT_HRS_APPROVED = 0
    
    If v_ViewOption = TrDept Then 'use department
        r_rsViewTimeSummary!DEPT_NUM = v_rsEmployeeData!DEPT_NUM
    Else
        If v_ViewOption = TrWG Then 'use WG
            r_rsViewTimeSummary!WG_NUM = v_rsEmployeeData!WG_NUM
        Else
        End If
    End If

    Exit Sub
InitailizeViewTimeSummaryFields_EH:
    TrsRaiseError m_strCMODULENAME, "InitailizeViewTimeSummaryFields"

End Sub


Private Sub SetViewTimeSummaryFields( _
                ByRef r_rsViewTimeSummary As ADODB.Recordset, _
                ByVal v_rsSummaryData As ADODB.Recordset)

    Dim l_strTimeType As String
    ' ('N','Y','O') -- Straight Time, Symbol Time, Overtime
    Dim l_strApprovedInd As String  ' Value is 'Y' or 'N'
    Dim l_strOTHERPAYIND_UNAPP As String  ' Value is 'Y' or 'N'
    Dim l_strOTHERPAYIND_APP As String  ' Value is 'Y' or 'N'
    Dim l_decHours As Variant
    Dim l_decTotalUnapprovedHours As Variant
    Dim l_decTotalApprovedHours As Variant
    Dim l_strInClause As String
    Dim l_strInClause2 As String
    On Error GoTo SetViewTimeSummaryFields_EH

    If v_rsSummaryData.EOF And v_rsSummaryData.BOF Then
    'No employees have actual time
        Exit Sub
    End If
    
   

    v_rsSummaryData.MoveFirst
    While Not v_rsSummaryData.EOF
       r_rsViewTimeSummary.AddNew
       If v_rsSummaryData!OTHERPAYIND_UNAPP = 0 Then
           r_rsViewTimeSummary!OTHER_PAY_IND_UNAPPROVED = "N"
       Else
           r_rsViewTimeSummary!OTHER_PAY_IND_UNAPPROVED = "Y"
       End If
       If v_rsSummaryData!OTHERPAYIND_APP = 0 Then
           r_rsViewTimeSummary!OTHER_PAY_IND_APPROVED = "N"
       Else
           r_rsViewTimeSummary!OTHER_PAY_IND_APPROVED = "Y"
       End If
       
       r_rsViewTimeSummary!EMP_NUM = v_rsSummaryData!EMP_NUM
       r_rsViewTimeSummary!EMP_NAME = v_rsSummaryData!EMP_NAME
       r_rsViewTimeSummary!ST_HRS_APPROVED = v_rsSummaryData!ST_HRS_APPROVED
       r_rsViewTimeSummary!ST_HRS_UNAPPROVED = v_rsSummaryData!ST_HRS_UNAPPROVED
       r_rsViewTimeSummary!OT_HRS_APPROVED = v_rsSummaryData!OT_HRS_APPROVED
       r_rsViewTimeSummary!OT_HRS_UNAPPROVED = v_rsSummaryData!OT_HRS_UNAPPROVED
       r_rsViewTimeSummary!SYM_HRS_APPROVED = v_rsSummaryData!SYM_HRS_APPROVED
       r_rsViewTimeSummary!SYM_HRS_UNAPPROVED = v_rsSummaryData!SYM_HRS_UNAPPROVED
       r_rsViewTimeSummary!TOT_HRS_UNAPPROVED = v_rsSummaryData!TOT_HRS_UNAPPROVED
       r_rsViewTimeSummary!TOT_HRS_APPROVED = v_rsSummaryData!TOT_HRS_APPROVED
       v_rsSummaryData.MoveNext

    Wend

     Exit Sub

SetViewTimeSummaryFields_EH:
    TrsRaiseError m_strCMODULENAME, "SetViewTimeSummaryFields"

End Sub


Public Function GetViewTimeSummaryByWorkGroup( _
        ByVal v_aryWorkGroups As Variant, _
        ByVal v_datStartDate As Date, _
        ByVal v_datEndDate As Date, _
        Optional ByVal v_bolActiveEmployeesOnly As Boolean = False) _
        As ADODB.Recordset

    Dim l_objCtx As ObjectContext
    Dim l_rsSummaryData As ADODB.Recordset
    Dim l_rsEmployeeData As ADODB.Recordset
    Dim l_aryEmployees As Variant



On Error GoTo GetViewTimeSummaryByWorkGroup_EH

    Set l_objCtx = GetObjectContext

    Set l_rsEmployeeData = GetViewTimeEmployeesInWorkGroup(v_aryWorkGroups, v_datStartDate, v_datEndDate, v_bolActiveEmployeesOnly)
    If l_rsEmployeeData.EOF And l_rsEmployeeData.BOF Then
        Set GetViewTimeSummaryByWorkGroup = Nothing
        l_objCtx.SetComplete
        Exit Function
    End If
    l_aryEmployees = BuildViewTimeEmployeeList(l_rsEmployeeData)
    Set l_rsSummaryData = GetViewTimeSummaryData(l_aryEmployees, v_datStartDate, v_datEndDate)
    Set GetViewTimeSummaryByWorkGroup = BuildViewTimeSummary(l_rsSummaryData)
    l_objCtx.SetComplete

Exit Function
GetViewTimeSummaryByWorkGroup_EH:
    l_objCtx.SetAbort
    TrsRaiseError m_strCMODULENAME, "GetViewTimeSummaryByWorkGroup"

End Function

Private Function GetViewTimeEmployeesInWorkGroup( _
            ByVal v_aryWorkGroups As Variant, _
            ByVal v_datStartDate As Date, _
            ByVal v_datEndDate As Date, _
            Optional ByVal v_bolActiveEmployeesOnly As Boolean = False) _
            As ADODB.Recordset
    Dim l_objCtx As ObjectContext
    Dim l_objdbWorkGroup As dbWorkgroup.cdWorkGroup
    Dim l_rsSummaryData As ADODB.Recordset

On Error GoTo GetViewTimeEmployeesInWorkGroup_EH

    Set l_objCtx = GetObjectContext
    
    If UBound(v_aryWorkGroups) > 1500 Then 'maximum number of employees reached
        Err.Raise trErrMaxNoWGLimitInView1500
    End If

    Set l_objdbWorkGroup = CreateObject("dbWorkGroup.cdWorkGroup")

    Set GetViewTimeEmployeesInWorkGroup = l_objdbWorkGroup.GetEmployeesInWorkGroup(v_aryWorkGroups, v_datStartDate, v_datEndDate, v_bolActiveEmployeesOnly)

    Set l_objdbWorkGroup = Nothing
    Set l_objCtx = Nothing

Exit Function
GetViewTimeEmployeesInWorkGroup_EH:
    
    l_objCtx.SetAbort
    
    Set l_objdbWorkGroup = Nothing
    Set l_objCtx = Nothing

    Select Case Err.Number
        Case trErrMaxNoWGLimitInView1500
             Err.Source = "1;;;;;;;;"
             TrsRaiseError m_strCMODULENAME, "GetViewTimeEmployeesInWorkGroup", trErrMaxNoWGLimitInView1500
        Case Else
             TrsRaiseError m_strCMODULENAME, "GetViewTimeEmployeesInWorkGroup"
    End Select


End Function


Private Function BuildViewTimeEmployeeList( _
            ByVal v_rsEmployeeData As ADODB.Recordset _
            ) As Variant

    Dim l_aryEmployees() As String
    Dim l_lngIndex As Long

    On Error GoTo BuildViewTimeEmployeeList_EH

    If v_rsEmployeeData.EOF And v_rsEmployeeData.BOF Then
        l_lngIndex = 0
        ReDim Preserve l_aryEmployees(l_lngIndex)
        l_aryEmployees(l_lngIndex) = "0"
        BuildViewTimeEmployeeList = l_aryEmployees
        Exit Function
    End If

    If v_rsEmployeeData.RecordCount > 444 Then 'maximum number of employees reached
        Err.Raise trErrMaxNoEmpTimeSummaryLimit444
    End If


    v_rsEmployeeData.MoveFirst

    l_lngIndex = 0
    While Not v_rsEmployeeData.EOF
        ReDim Preserve l_aryEmployees(l_lngIndex)
        l_aryEmployees(l_lngIndex) = str(v_rsEmployeeData!EMP_NUM)
        v_rsEmployeeData.MoveNext
        l_lngIndex = l_lngIndex + 1
    Wend

    BuildViewTimeEmployeeList = l_aryEmployees

    Exit Function

BuildViewTimeEmployeeList_EH:

    Select Case Err.Number
        Case trErrMaxNoEmpTimeSummaryLimit444
             Err.Source = "1;;;;;;;;"
             TrsRaiseError m_strCMODULENAME, "BuildViewTimeEmployeeList", trErrMaxNoEmpTimeSummaryLimit444
        Case Else
             TrsRaiseError m_strCMODULENAME, "BuildViewTimeEmployeeList"
    End Select

End Function

'//for older TRIS clients, pick list description is 30 chars
Public Function GetActualTimeDetailAndSummary( _
        ByVal v_aryEmployees As Variant, _
        ByVal v_datStartDate As Date, _
        ByVal v_datEndDate As Date, _
        ByRef r_objRSActualTimeDetail As ADODB.Recordset, _
        ByRef r_objRSTotalHoursByActivitySummary As ADODB.Recordset)
' r_objRSActualTimeDetail and r_objRSTotalHoursByActivitySummary
' are output parameters.
    On Error GoTo eh
    
    Call RetrieveActualTimeEntriesForSummaryAndDetail(v_aryEmployees, v_datStartDate, v_datEndDate, TrActualTimeEntriesDetailAndTotalHoursByActivitySummary, r_objRSActualTimeDetail, r_objRSTotalHoursByActivitySummary, True)
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "GetActualTimeDetailAndSummary", , , , , True
End Function

'//interim use, keep for compatability
Public Function GetActualTimeDetailAndSummary2( _
        ByVal v_aryEmployees As Variant, _
        ByVal v_datStartDate As Date, _
        ByVal v_datEndDate As Date, _
        ByRef r_objRSActualTimeDetail As ADODB.Recordset, _
        ByRef r_objRSTotalHoursByActivitySummary As ADODB.Recordset)
' r_objRSActualTimeDetail and r_objRSTotalHoursByActivitySummary
' are output parameters.
    On Error GoTo eh
    
    Call RetrieveActualTimeEntriesForSummaryAndDetail(v_aryEmployees, v_datStartDate, v_datEndDate, TrActualTimeEntriesDetailAndTotalHoursByActivitySummary, r_objRSActualTimeDetail, r_objRSTotalHoursByActivitySummary, True, , 60)
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "GetActualTimeDetailAndSummary2", , , , , True
End Function

'//for newer TRIS clients DFE 48 chars, OTTC pick list description increased to 60 chars
Public Function GetActualTimeDetailAndSummary2B( _
        ByVal v_aryEmployees As Variant, _
        ByVal v_datStartDate As Date, _
        ByVal v_datEndDate As Date, _
        ByRef r_objRSActualTimeDetail As ADODB.Recordset, _
        ByRef r_objRSTotalHoursByActivitySummary As ADODB.Recordset, _
        ByVal v_intDescriptionLength As Integer)
' r_objRSActualTimeDetail and r_objRSTotalHoursByActivitySummary
' are output parameters.
    On Error GoTo eh
    
    Call RetrieveActualTimeEntriesForSummaryAndDetail(v_aryEmployees, v_datStartDate, v_datEndDate, TrActualTimeEntriesDetailAndTotalHoursByActivitySummary, r_objRSActualTimeDetail, r_objRSTotalHoursByActivitySummary, True, , v_intDescriptionLength)
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "GetActualTimeDetailAndSummary2B", , , , , True
End Function

Public Function GetOtherPayEntries( _
        ByVal v_aryEmployees As Variant, _
        ByVal v_datStartDate As Date, _
        ByVal v_datEndDate As Date) As ADODB.Recordset
    
    Dim l_objdbTimecard As dbTimeCard.cdTimeCard

On Error GoTo GetOtherPayEntries_EH
    
    Set l_objdbTimecard = CreateObject("dbTimeCard.cdTimeCard")
    
    Set GetOtherPayEntries = l_objdbTimecard.GetOtherPayEntries(v_aryEmployees, v_datStartDate, v_datEndDate)
    
    Set l_objdbTimecard = Nothing
    
Exit Function
GetOtherPayEntries_EH:
    TrsRaiseError m_strCMODULENAME, "GetOtherPayEntries", , , , , True
End Function

'//not sure where this is used, pick list description is 30 chars
Public Sub GetActualTimeDetailAndSummaryAndOtherPay( _
        ByVal v_aryEmployees As Variant, _
        ByVal v_datStartDate As Date, _
        ByVal v_datEndDate As Date, _
        ByRef r_objRSActualTimeDetail As ADODB.Recordset, _
        ByRef r_objRSTotalHoursByActivitySummary As ADODB.Recordset, _
        ByRef r_objRSOtherPay As ADODB.Recordset)
' r_objRSActualTimeDetail and r_objRSTotalHoursByActivitySummary
' are output parameters.
    
    Dim lr_objRSActualTimeDetail As ADODB.Recordset
    Dim lr_objRSTotalHoursByActivitySummary As ADODB.Recordset
    Dim lr_objRSOtherPay As ADODB.Recordset
    
    On Error GoTo GetActualTimeDetailAndSummaryAndOtherPay_EH
    
    Call RetrieveActualTimeEntriesForSummaryAndDetail(v_aryEmployees, v_datStartDate, v_datEndDate, TrActualTimeEntriesDetailAndTotalHoursByActivitySummary, lr_objRSActualTimeDetail, lr_objRSTotalHoursByActivitySummary, True)
    Set lr_objRSOtherPay = GetOtherPayEntries(v_aryEmployees, v_datStartDate, v_datEndDate)
    
        
    Set r_objRSActualTimeDetail = lr_objRSActualTimeDetail
    Set r_objRSTotalHoursByActivitySummary = lr_objRSTotalHoursByActivitySummary
    Set r_objRSOtherPay = lr_objRSOtherPay
    
    Exit Sub

GetActualTimeDetailAndSummaryAndOtherPay_EH:

    TrsRaiseError m_strCMODULENAME, "GetActualTimeDetailAndSummaryAndOtherPay"
End Sub

'//for older TRIS clients, pick list description is 30 chars
Public Function GetActualTimeDetailWithOtherPay( _
        ByVal v_aryEmployees As Variant, _
        ByVal v_datStartDate As Date, _
        ByVal v_datEndDate As Date) As ADODB.Recordset
    
    Dim l_objRSActualTimeDetail As ADODB.Recordset
    Dim l_objRSTotalHoursByActivitySummary As ADODB.Recordset

    On Error GoTo GetActualTimeDetailWithOtherPay_EH
    
    Call RetrieveActualTimeEntriesForSummaryAndDetail(v_aryEmployees, v_datStartDate, v_datEndDate, TrActualTimeEntriesDetail, l_objRSActualTimeDetail, l_objRSTotalHoursByActivitySummary, True, True)
    Set GetActualTimeDetailWithOtherPay = l_objRSActualTimeDetail
    Exit Function

GetActualTimeDetailWithOtherPay_EH:
    TrsRaiseError m_strCMODULENAME, "GetActualTimeDetailWithOtherPay"
End Function

'//interim use, keep for compatability
Public Function GetActualTimeDetailWithOtherPay2( _
        ByVal v_aryEmployees As Variant, _
        ByVal v_datStartDate As Date, _
        ByVal v_datEndDate As Date) As ADODB.Recordset
    
    Dim l_objRSActualTimeDetail As ADODB.Recordset
    Dim l_objRSTotalHoursByActivitySummary As ADODB.Recordset

    On Error GoTo eh
    
    Call RetrieveActualTimeEntriesForSummaryAndDetail(v_aryEmployees, v_datStartDate, v_datEndDate, TrActualTimeEntriesDetail, l_objRSActualTimeDetail, l_objRSTotalHoursByActivitySummary, True, True, 60)
    Set GetActualTimeDetailWithOtherPay2 = l_objRSActualTimeDetail
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "GetActualTimeDetailWithOtherPay2"
End Function

'//for newer TRIS clients DFE 48 chars, OTTC pick list description increased to 60 chars
Public Function GetActualTimeDetailWithOtherPay2B( _
        ByVal v_aryEmployees As Variant, _
        ByVal v_datStartDate As Date, _
        ByVal v_datEndDate As Date, _
        ByVal v_intDescriptionLength As Integer) As ADODB.Recordset
    
    Dim l_objRSActualTimeDetail As ADODB.Recordset
    Dim l_objRSTotalHoursByActivitySummary As ADODB.Recordset

    On Error GoTo eh
    
    Call RetrieveActualTimeEntriesForSummaryAndDetail(v_aryEmployees, v_datStartDate, v_datEndDate, TrActualTimeEntriesDetail, l_objRSActualTimeDetail, l_objRSTotalHoursByActivitySummary, True, True, v_intDescriptionLength)
    Set GetActualTimeDetailWithOtherPay2B = l_objRSActualTimeDetail
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "GetActualTimeDetailWithOtherPay2B"
End Function

''//-----The following was created to allow the Callaway WebService to run faster
'Public Function GetActualTimeDetailWithOtherPayForDept( _
'        ByVal v_aryDepartments As Variant, _
'        ByVal v_datStartDate As Date, _
'        ByVal v_datEndDate As Date) As ADODB.Recordset
'
'    Dim l_objRSActualTimeDetail As ADODB.Recordset
'    Dim l_objRSTotalHoursByActivitySummary As ADODB.Recordset
'
'    On Error GoTo eh
'
'    Call RetrieveActualTimeEntriesForSummaryAndDetailForDept(v_aryDepartments, v_datStartDate, v_datEndDate, TrActualTimeEntriesDetail, l_objRSActualTimeDetail, l_objRSTotalHoursByActivitySummary, True, True)
'    Set GetActualTimeDetailWithOtherPayForDept = l_objRSActualTimeDetail
'    Exit Function
'
'eh:
'    TrsRaiseError m_strCMODULENAME, "GetActualTimeDetailWithOtherPayByDept"
'End Function

Public Function GetApproveTimeSummaryByWorkGroup( _
        ByVal v_aryWorkGroups As Variant, _
        ByVal v_datStartDate As Date, _
        ByVal v_datEndDate As Date, _
        Optional ByVal v_bolActiveEmployeesOnly As Boolean = False) _
        As ADODB.Recordset

    Dim l_objCtx As ObjectContext
    Dim l_rsSummaryData As ADODB.Recordset
    Dim l_rsEmployeeData As ADODB.Recordset
    Dim l_aryEmployees As Variant

On Error GoTo GetApproveTimeSummaryByWorkGroup_EH

    Set l_objCtx = GetObjectContext
    
    Set l_rsEmployeeData = GetEmployeesInWorkGroupTimeCharged(v_aryWorkGroups, v_datStartDate, v_datEndDate, v_bolActiveEmployeesOnly)
    If l_rsEmployeeData.EOF And l_rsEmployeeData.BOF Then
        Set GetApproveTimeSummaryByWorkGroup = Nothing
        l_objCtx.SetComplete
        Exit Function
    End If
    l_aryEmployees = BuildViewTimeEmployeeList(l_rsEmployeeData)

    Set l_rsSummaryData = GetApprovedTimeSummaryData(l_aryEmployees, v_datStartDate, v_datEndDate)
    Set GetApproveTimeSummaryByWorkGroup = BuildApprovedTimeSummary(l_rsSummaryData, l_rsEmployeeData, v_aryWorkGroups)
    l_objCtx.SetComplete

Exit Function
GetApproveTimeSummaryByWorkGroup_EH:
    l_objCtx.SetAbort
    TrsRaiseError m_strCMODULENAME, "GetApproveTimeSummaryByWorkGroup"

End Function


Private Function GetEmployeesInWorkGroupTimeCharged( _
            ByVal v_aryWorkGroups As Variant, _
            ByVal v_datStartDate As Date, _
            ByVal v_datEndDate As Date, _
            Optional ByVal v_bolActiveEmployeesOnly As Boolean = False _
            ) As ADODB.Recordset
    Dim l_objCtx As ObjectContext
    Dim l_objdbWorkGroup As dbWorkgroup.cdWorkGroup
    Dim l_rsSummaryData As ADODB.Recordset

On Error GoTo GetEmployeesInWorkGroupTimeCharged_EH

    Set l_objCtx = GetObjectContext
    Set l_objdbWorkGroup = CreateObject("dbWorkGroup.cdWorkGroup")

    Set GetEmployeesInWorkGroupTimeCharged = l_objdbWorkGroup.GetEmployeesInWorkGroupTimeCharged(v_aryWorkGroups, v_datStartDate, v_datEndDate, v_bolActiveEmployeesOnly)

    Set l_objdbWorkGroup = Nothing
    Set l_objCtx = Nothing

Exit Function
GetEmployeesInWorkGroupTimeCharged_EH:
    l_objCtx.SetAbort
    Set l_objdbWorkGroup = Nothing
    Set l_objCtx = Nothing

    TrsRaiseError m_strCMODULENAME, "GetEmployeesInWorkGroupTimeCharged"

End Function


Private Sub SetApprovedTimeSummaryFields( _
                ByRef r_rsViewTimeSummary As ADODB.Recordset, _
                ByVal v_rsSummaryData As ADODB.Recordset, _
                ByVal v_lngEmployee As Long, _
                ByVal v_aryWorkGroups As Variant)

    Dim l_decST_HRS_APPROVED As Variant
    Dim l_decST_HRS_UNAPPROVED As Variant
    Dim l_decOT_HRS_APPROVED As Variant
    Dim l_decOT_HRS_UNAPPROVED As Variant
    Dim l_decSYM_HRS_APPROVED As Variant
    Dim l_decSYM_HRS_UNAPPROVED As Variant
    Dim l_decTOT_HRS_UNAPPROVED As Variant
    Dim l_decTOT_HRS_APPROVED As Variant
    Dim l_strOTHER_PAY_IND_UNAPPROVED As String
    Dim l_strOTHER_PAY_IND_APPROVED As String
    Dim l_strfilteredEmployee As String

    On Error GoTo SetApprovedTimeSummaryFields_EH

    If v_rsSummaryData.EOF And v_rsSummaryData.BOF Then
    'No employees have actual time
        Exit Sub
    End If

    
    l_strfilteredEmployee = "EMP_NUM = " & Trim(str(v_lngEmployee))
    v_rsSummaryData.Filter = BuildWGFilterClause(v_aryWorkGroups, l_strfilteredEmployee)

    If v_rsSummaryData.EOF And v_rsSummaryData.BOF Then
    'This employee does not have actual time
        v_rsSummaryData.Filter = ""
        Exit Sub
    End If
    
    
    v_rsSummaryData.MoveFirst
    
    l_decST_HRS_APPROVED = 0
    l_decST_HRS_UNAPPROVED = 0
    l_decOT_HRS_APPROVED = 0
    l_decOT_HRS_UNAPPROVED = 0
    l_decSYM_HRS_APPROVED = 0
    l_decSYM_HRS_UNAPPROVED = 0
    l_decTOT_HRS_UNAPPROVED = 0
    l_decTOT_HRS_APPROVED = 0
    l_strOTHER_PAY_IND_UNAPPROVED = "N"
    l_strOTHER_PAY_IND_APPROVED = "N"
    
    While Not v_rsSummaryData.EOF
       If v_rsSummaryData!OTHERPAYIND_UNAPP > 0 Then
           l_strOTHER_PAY_IND_UNAPPROVED = "Y"
       End If
       If v_rsSummaryData!OTHERPAYIND_APP > 0 Then
           l_strOTHER_PAY_IND_APPROVED = "Y"
       End If
       l_decST_HRS_APPROVED = l_decST_HRS_APPROVED + CDec(v_rsSummaryData!ST_HRS_APPROVED)
       l_decST_HRS_UNAPPROVED = l_decST_HRS_UNAPPROVED + CDec(v_rsSummaryData!ST_HRS_UNAPPROVED)
       l_decOT_HRS_APPROVED = l_decOT_HRS_APPROVED + CDec(v_rsSummaryData!OT_HRS_APPROVED)
       l_decOT_HRS_UNAPPROVED = l_decOT_HRS_UNAPPROVED + CDec(v_rsSummaryData!OT_HRS_UNAPPROVED)
       l_decSYM_HRS_APPROVED = l_decSYM_HRS_APPROVED + CDec(v_rsSummaryData!SYM_HRS_APPROVED)
       l_decSYM_HRS_UNAPPROVED = l_decSYM_HRS_UNAPPROVED + CDec(v_rsSummaryData!SYM_HRS_UNAPPROVED)
       l_decTOT_HRS_UNAPPROVED = l_decTOT_HRS_UNAPPROVED + CDec(v_rsSummaryData!TOT_HRS_UNAPPROVED)
       l_decTOT_HRS_APPROVED = l_decTOT_HRS_APPROVED + CDec(v_rsSummaryData!TOT_HRS_APPROVED)
       v_rsSummaryData.MoveNext
    Wend


    r_rsViewTimeSummary!ST_HRS_APPROVED = l_decST_HRS_APPROVED
    r_rsViewTimeSummary!ST_HRS_UNAPPROVED = l_decST_HRS_UNAPPROVED
    r_rsViewTimeSummary!OT_HRS_APPROVED = l_decOT_HRS_APPROVED
    r_rsViewTimeSummary!OT_HRS_UNAPPROVED = l_decOT_HRS_UNAPPROVED
    r_rsViewTimeSummary!SYM_HRS_APPROVED = l_decSYM_HRS_APPROVED
    r_rsViewTimeSummary!SYM_HRS_UNAPPROVED = l_decSYM_HRS_UNAPPROVED
    r_rsViewTimeSummary!TOT_HRS_UNAPPROVED = l_decTOT_HRS_UNAPPROVED
    r_rsViewTimeSummary!TOT_HRS_APPROVED = l_decTOT_HRS_APPROVED
    r_rsViewTimeSummary!OTHER_PAY_IND_UNAPPROVED = l_strOTHER_PAY_IND_UNAPPROVED
    r_rsViewTimeSummary!OTHER_PAY_IND_APPROVED = l_strOTHER_PAY_IND_APPROVED

    v_rsSummaryData.Filter = ""

     Exit Sub

SetApprovedTimeSummaryFields_EH:
    TrsRaiseError m_strCMODULENAME, "SetApprovedTimeSummaryFields"

End Sub
Private Function BuildApprovedTimeSummary( _
            ByVal v_rsSummaryData As ADODB.Recordset, _
            ByVal v_rsEmployeeData As ADODB.Recordset, _
            ByVal v_aryWorkGroups As Variant _
            ) As ADODB.Recordset

    Dim l_rsViewTimeSummary As ADODB.Recordset

    On Error GoTo BuildApprovedTimeSummary_EH

    Set l_rsViewTimeSummary = CreateViewTimeSummaryRecordSet(TrWG)

    If v_rsEmployeeData.EOF And v_rsEmployeeData.BOF Then
    ' No employees
        Set BuildApprovedTimeSummary = Nothing
        Exit Function
    End If
    
    Dim l_strdupclause As String
    Dim l_strfilteredEmployee As String

    v_rsEmployeeData.MoveFirst

    While Not v_rsEmployeeData.EOF
        l_strfilteredEmployee = "EMP_NUM = " & Trim(str(v_rsEmployeeData!EMP_NUM))
        l_strdupclause = BuildFilterDupClause(l_strfilteredEmployee)
        l_rsViewTimeSummary.Filter = l_strdupclause
        If l_rsViewTimeSummary.EOF And l_rsViewTimeSummary.BOF Then
        'This employee has not been processed
            l_rsViewTimeSummary.Filter = ""
            l_rsViewTimeSummary.AddNew
            InitailizeViewTimeSummaryFields l_rsViewTimeSummary, v_rsEmployeeData, TrWG
            SetApprovedTimeSummaryFields l_rsViewTimeSummary, v_rsSummaryData, v_rsEmployeeData!EMP_NUM, v_aryWorkGroups
        Else 'bypass duplicate employee
            l_rsViewTimeSummary.Filter = ""
        End If
        v_rsEmployeeData.MoveNext
    Wend
    
    l_rsViewTimeSummary.MoveFirst

    Set BuildApprovedTimeSummary = l_rsViewTimeSummary

    Exit Function

BuildApprovedTimeSummary_EH:
    Set BuildApprovedTimeSummary = Nothing
    TrsRaiseError m_strCMODULENAME, "BuildApprovedTimeSummary"

End Function

Private Function GetApprovedTimeSummaryData( _
            ByVal v_aryEmployees As Variant, _
            ByVal v_datStartDate As Date, _
            ByVal v_datEndDate As Date _
            ) As ADODB.Recordset

    Dim l_objCtx As ObjectContext
    Dim l_objdbTimecard As dbTimeCard.cdTimeCard
    Dim l_rsSummaryData As ADODB.Recordset

On Error GoTo GetApprovedTimeSummaryData_EH

    Set l_objCtx = GetObjectContext
    Set l_objdbTimecard = CreateObject("dbTimeCard.cdTimeCard")

    Set GetApprovedTimeSummaryData = l_objdbTimecard.GetApproveTimeSummaryByEmployee(v_aryEmployees, v_datStartDate, v_datEndDate)

    Set l_objdbTimecard = Nothing
    Set l_objCtx = Nothing

Exit Function
GetApprovedTimeSummaryData_EH:
    l_objCtx.SetAbort
    Set l_objdbTimecard = Nothing
    Set l_objCtx = Nothing

    TrsRaiseError m_strCMODULENAME, "GetApprovedTimeSummaryData"

End Function


Private Function BuildWGFilterClause( _
            ByVal v_aryValues As Variant, _
            ByVal v_strfilteredEmployee As String _
            ) As String
            
'"(EMP_NUM = 27284 and WG_NUM = 14189) OR (EMP_NUM = 27284 and WG_NUM = 12211)"
'(EMP_NUM = 27158 and WG_NUM = 14189) OR (EMP_NUM = 27158 and WG_NUM = 3152)

    Dim str As String
    Dim i As Long
    Dim l_lngStart As Long
    Dim l_lngEnd As Long
    
    l_lngStart = LBound(v_aryValues)
    l_lngEnd = UBound(v_aryValues)
    str = "(" & v_strfilteredEmployee & " and WG_NUM = "
    For i = l_lngStart To l_lngEnd
        str = str & Trim(v_aryValues(i))
        If Not (i = l_lngEnd) Then
            str = str & ") OR (" & v_strfilteredEmployee & " and WG_NUM = "
        Else
            str = str & ")"
        End If
    Next
    
    BuildWGFilterClause = MDAC25FilterFix(str)
            

End Function


Private Function BuildFilterDupClause( _
            ByVal v_strfilteredEmployee As String _
            ) As String
            
'This function will Build a string for filtering of processed employees
'    '"(EMP_NUM = 27284 and OTHER_PAY_IND_APPROVED = "N") OR (EMP_NUM = 27284 and OTHER_PAY_IND_APPROVED = "Y")"

    Dim str As String
    str = "(" & v_strfilteredEmployee & " and OTHER_PAY_IND_APPROVED = 'N'"
    str = str & ") OR (" & v_strfilteredEmployee & " and OTHER_PAY_IND_APPROVED = 'Y')"
    
    BuildFilterDupClause = MDAC25FilterFix(str)
            
End Function


Public Function GetViewTimeSummaryByDepartment( _
        ByVal v_aryDepartments As Variant, _
        ByVal v_datStartDate As Date, _
        ByVal v_datEndDate As Date, _
        Optional ByVal v_bolActiveEmployeesOnly As Boolean = False) _
        As ADODB.Recordset

    Dim l_objCtx As ObjectContext
    Dim l_rsSummaryData As ADODB.Recordset
    Dim l_rsEmployeeData As ADODB.Recordset
    Dim l_aryEmployees As Variant

On Error GoTo GetViewTimeSummaryByDepartment_EH

    Set l_objCtx = GetObjectContext
    
    Set l_rsEmployeeData = GetEmployeesInDepartmentTimeCharged(v_aryDepartments, v_datStartDate, v_datEndDate, v_bolActiveEmployeesOnly)
    If l_rsEmployeeData.EOF And l_rsEmployeeData.BOF Then
        Set GetViewTimeSummaryByDepartment = Nothing
        l_objCtx.SetComplete
        Exit Function
    End If
    l_aryEmployees = BuildViewTimeEmployeeList(l_rsEmployeeData)

    Set l_rsSummaryData = GetDeptTimeSummaryData(l_aryEmployees, v_datStartDate, v_datEndDate)
    Set GetViewTimeSummaryByDepartment = BuildDeptTimeSummary(l_rsSummaryData, l_rsEmployeeData, v_aryDepartments)
    l_objCtx.SetComplete

Exit Function
GetViewTimeSummaryByDepartment_EH:
    l_objCtx.SetAbort
    TrsRaiseError m_strCMODULENAME, "GetViewTimeSummaryByDepartment"

End Function



Private Function GetEmployeesInDepartmentTimeCharged( _
            ByVal v_aryDepartments As Variant, _
            ByVal v_datStartDate As Date, _
            ByVal v_datEndDate As Date, _
            Optional ByVal v_bolActiveEmployeesOnly As Boolean = False _
            ) As ADODB.Recordset
    Dim l_objCtx As ObjectContext
    Dim l_objdbWorkGroup As dbWorkgroup.cdWorkGroup
    Dim l_rsSummaryData As ADODB.Recordset

On Error GoTo GetEmployeesInDepartmentTimeCharged_EH

    Set l_objCtx = GetObjectContext
    If UBound(v_aryDepartments) > 250 Then 'maximum number of employees reached
        Err.Raise trErrMaxNoDEPTLimitInView250
    End If

    Set l_objdbWorkGroup = CreateObject("dbWorkGroup.cdWorkGroup")

    Set GetEmployeesInDepartmentTimeCharged = l_objdbWorkGroup.GetEmployeesInDepartmentTimeCharged(v_aryDepartments, v_datStartDate, v_datEndDate, v_bolActiveEmployeesOnly)

    Set l_objdbWorkGroup = Nothing
    Set l_objCtx = Nothing

Exit Function

GetEmployeesInDepartmentTimeCharged_EH:
    l_objCtx.SetAbort
    Set l_objdbWorkGroup = Nothing
    Set l_objCtx = Nothing

    Select Case Err.Number
        Case trErrMaxNoDEPTLimitInView250
             Err.Source = "1;;;;;;;;"
             TrsRaiseError m_strCMODULENAME, "GetEmployeesInDepartmentTimeCharged", trErrMaxNoDEPTLimitInView250
        Case Else
             TrsRaiseError m_strCMODULENAME, "GetEmployeesInDepartmentTimeCharged"
    End Select

End Function

'copied from busDistributionMgr.cbsDistributionMgr with edits to not cache the cache, will only be called prior to calling EMPRV to get distribution/job/task list
Private Function GetDepartmentOfWorkGroup( _
            ByVal v_lngChargeWGNum As Long _
            ) As String
            
    Dim l_strDept As String
    Dim l_obj As dbWorkgroup.cdWorkGroup
'    Dim l_obj2 As dbDepartment.cdDepartment
    
'    Dim l_intPosition As Integer
   
On Error GoTo eh:
    
    '//cache the cache ?
'    If IsStringInList2(str(v_lngChargeWGNum), l_varWorkGroupDepartmentCache, l_lngWorkGroupDepartmentCacheCount, l_intPosition) Then
'        l_strDept = l_varWorkGroupDepartmentCache(l_intPosition, 1)
'    Else
        
        Set l_obj = CreateObject("dbWorkGroup.cdWorkGroup")
        l_strDept = l_obj.DepartmentNumber(v_lngChargeWGNum)
        Set l_obj = Nothing
        '//cache the cache ?
'        l_varWorkGroupDepartmentCache(l_lngWorkGroupDepartmentCacheCount, 0) = str(v_lngChargeWGNum)
'        l_varWorkGroupDepartmentCache(l_lngWorkGroupDepartmentCacheCount, 1) = l_strDept
'        l_lngWorkGroupDepartmentCacheCount = l_lngWorkGroupDepartmentCacheCount + 1
'    End If
    
    GetDepartmentOfWorkGroup = l_strDept
    
    Exit Function
    
eh:
    Set l_obj = Nothing
'    Set l_obj2 = Nothing
    TrsRaiseError m_strCMODULENAME, "GetDepartmentOfWorkGroup"

            
End Function

Private Function GetDeptTimeSummaryData( _
            ByVal v_aryEmployees As Variant, _
            ByVal v_datStartDate As Date, _
            ByVal v_datEndDate As Date _
            ) As ADODB.Recordset

    Dim l_objCtx As ObjectContext
    Dim l_objdbTimecard As dbTimeCard.cdTimeCard
    Dim l_rsSummaryData As ADODB.Recordset

On Error GoTo GetDeptTimeSummaryData_EH

    Set l_objCtx = GetObjectContext
    Set l_objdbTimecard = CreateObject("dbTimeCard.cdTimeCard")

    Set GetDeptTimeSummaryData = l_objdbTimecard.GetDeptTimeSummaryByEmployee(v_aryEmployees, v_datStartDate, v_datEndDate)

    Set l_objdbTimecard = Nothing
    Set l_objCtx = Nothing

Exit Function
GetDeptTimeSummaryData_EH:
    l_objCtx.SetAbort
    Set l_objdbTimecard = Nothing
    Set l_objCtx = Nothing

    TrsRaiseError m_strCMODULENAME, "GetDeptTimeSummaryData"

End Function

Private Function BuildDeptTimeSummary( _
            ByVal v_rsSummaryData As ADODB.Recordset, _
            ByVal v_rsEmployeeData As ADODB.Recordset, _
            ByVal v_aryDepartments As Variant _
            ) As ADODB.Recordset

    Dim l_rsViewTimeSummary As ADODB.Recordset

    On Error GoTo BuildDeptTimeSummary_EH

    Set l_rsViewTimeSummary = CreateViewTimeSummaryRecordSet(TrDept)

    If v_rsEmployeeData.EOF And v_rsEmployeeData.BOF Then
    ' No employees
        Set BuildDeptTimeSummary = Nothing
        Exit Function
    End If
    
    Dim l_strdupclause As String
    Dim l_strfilteredEmployee As String

    v_rsEmployeeData.MoveFirst

    While Not v_rsEmployeeData.EOF
        l_strfilteredEmployee = "EMP_NUM = " & Trim(str(v_rsEmployeeData!EMP_NUM))
        l_strdupclause = BuildFilterDupClause(l_strfilteredEmployee)
        l_rsViewTimeSummary.Filter = l_strdupclause
        If l_rsViewTimeSummary.EOF And l_rsViewTimeSummary.BOF Then
        'This employee has not been processed
            l_rsViewTimeSummary.Filter = ""
            l_rsViewTimeSummary.AddNew
            InitailizeViewTimeSummaryFields l_rsViewTimeSummary, v_rsEmployeeData, TrDept
            SetDeptTimeSummaryFields l_rsViewTimeSummary, v_rsSummaryData, v_rsEmployeeData!EMP_NUM, v_aryDepartments
        Else 'bypass duplicate employee
            l_rsViewTimeSummary.Filter = ""
        End If
        v_rsEmployeeData.MoveNext
    Wend
    
    l_rsViewTimeSummary.MoveFirst

    Set BuildDeptTimeSummary = l_rsViewTimeSummary

    Exit Function

BuildDeptTimeSummary_EH:
    Set BuildDeptTimeSummary = Nothing
    TrsRaiseError m_strCMODULENAME, "BuildDeptTimeSummary"

End Function


Private Sub SetDeptTimeSummaryFields( _
                ByRef r_rsViewTimeSummary As ADODB.Recordset, _
                ByVal v_rsSummaryData As ADODB.Recordset, _
                ByVal v_lngEmployee As Long, _
                ByVal v_aryDepartments As Variant)

    Dim l_decST_HRS_APPROVED As Variant
    Dim l_decST_HRS_UNAPPROVED As Variant
    Dim l_decOT_HRS_APPROVED As Variant
    Dim l_decOT_HRS_UNAPPROVED As Variant
    Dim l_decSYM_HRS_APPROVED As Variant
    Dim l_decSYM_HRS_UNAPPROVED As Variant
    Dim l_decTOT_HRS_UNAPPROVED As Variant
    Dim l_decTOT_HRS_APPROVED As Variant
    Dim l_strOTHER_PAY_IND_UNAPPROVED As String
    Dim l_strOTHER_PAY_IND_APPROVED As String
    Dim l_strfilteredEmployee As String

    On Error GoTo SetDeptTimeSummaryFields_EH

    If v_rsSummaryData.EOF And v_rsSummaryData.BOF Then
    'No employees have actual time
        Exit Sub
    End If

    
    l_strfilteredEmployee = "EMP_NUM = " & Trim(str(v_lngEmployee))
    v_rsSummaryData.Filter = BuildDEPTFilterClause(v_aryDepartments, l_strfilteredEmployee)

    If v_rsSummaryData.EOF And v_rsSummaryData.BOF Then
    'This employee does not have actual time
        v_rsSummaryData.Filter = ""
        Exit Sub
    End If
    
    
    v_rsSummaryData.MoveFirst
    
    l_decST_HRS_APPROVED = 0
    l_decST_HRS_UNAPPROVED = 0
    l_decOT_HRS_APPROVED = 0
    l_decOT_HRS_UNAPPROVED = 0
    l_decSYM_HRS_APPROVED = 0
    l_decSYM_HRS_UNAPPROVED = 0
    l_decTOT_HRS_UNAPPROVED = 0
    l_decTOT_HRS_APPROVED = 0
    l_strOTHER_PAY_IND_UNAPPROVED = "N"
    l_strOTHER_PAY_IND_APPROVED = "N"
    
    While Not v_rsSummaryData.EOF
       If v_rsSummaryData!OTHERPAYIND_UNAPP > 0 Then
           l_strOTHER_PAY_IND_UNAPPROVED = "Y"
       End If
       If v_rsSummaryData!OTHERPAYIND_APP > 0 Then
           l_strOTHER_PAY_IND_APPROVED = "Y"
       End If
       l_decST_HRS_APPROVED = l_decST_HRS_APPROVED + CDec(v_rsSummaryData!ST_HRS_APPROVED)
       l_decST_HRS_UNAPPROVED = l_decST_HRS_UNAPPROVED + CDec(v_rsSummaryData!ST_HRS_UNAPPROVED)
       l_decOT_HRS_APPROVED = l_decOT_HRS_APPROVED + CDec(v_rsSummaryData!OT_HRS_APPROVED)
       l_decOT_HRS_UNAPPROVED = l_decOT_HRS_UNAPPROVED + CDec(v_rsSummaryData!OT_HRS_UNAPPROVED)
       l_decSYM_HRS_APPROVED = l_decSYM_HRS_APPROVED + CDec(v_rsSummaryData!SYM_HRS_APPROVED)
       l_decSYM_HRS_UNAPPROVED = l_decSYM_HRS_UNAPPROVED + CDec(v_rsSummaryData!SYM_HRS_UNAPPROVED)
       l_decTOT_HRS_UNAPPROVED = l_decTOT_HRS_UNAPPROVED + CDec(v_rsSummaryData!TOT_HRS_UNAPPROVED)
       l_decTOT_HRS_APPROVED = l_decTOT_HRS_APPROVED + CDec(v_rsSummaryData!TOT_HRS_APPROVED)
       v_rsSummaryData.MoveNext
    Wend


    r_rsViewTimeSummary!ST_HRS_APPROVED = l_decST_HRS_APPROVED
    r_rsViewTimeSummary!ST_HRS_UNAPPROVED = l_decST_HRS_UNAPPROVED
    r_rsViewTimeSummary!OT_HRS_APPROVED = l_decOT_HRS_APPROVED
    r_rsViewTimeSummary!OT_HRS_UNAPPROVED = l_decOT_HRS_UNAPPROVED
    r_rsViewTimeSummary!SYM_HRS_APPROVED = l_decSYM_HRS_APPROVED
    r_rsViewTimeSummary!SYM_HRS_UNAPPROVED = l_decSYM_HRS_UNAPPROVED
    r_rsViewTimeSummary!TOT_HRS_UNAPPROVED = l_decTOT_HRS_UNAPPROVED
    r_rsViewTimeSummary!TOT_HRS_APPROVED = l_decTOT_HRS_APPROVED
    r_rsViewTimeSummary!OTHER_PAY_IND_UNAPPROVED = l_strOTHER_PAY_IND_UNAPPROVED
    r_rsViewTimeSummary!OTHER_PAY_IND_APPROVED = l_strOTHER_PAY_IND_APPROVED

    v_rsSummaryData.Filter = ""

     Exit Sub

SetDeptTimeSummaryFields_EH:
    TrsRaiseError m_strCMODULENAME, "SetDeptTimeSummaryFields"

End Sub

Private Function BuildDEPTFilterClause( _
            ByVal v_aryValues As Variant, _
            ByVal v_strfilteredEmployee As String _
            ) As String
            
'"(EMP_NUM = 27284 and DEPT_NUM = '050') OR (EMP_NUM = 27284 and DEPT_NUM = '050')"
'(EMP_NUM = 27158 and DEPT_NUM = '050') OR (EMP_NUM = 27158 and DEPT_NUM = '050')

    Dim str As String
    Dim i As Long
    Dim l_lngStart As Long
    Dim l_lngEnd As Long
    
    l_lngStart = LBound(v_aryValues)
    l_lngEnd = UBound(v_aryValues)
    str = "(" & v_strfilteredEmployee & " and DEPT_NUM = "
    For i = l_lngStart To l_lngEnd
        str = str & "'" & Trim(v_aryValues(i)) & "'"
        If Not (i = l_lngEnd) Then
            str = str & ") OR (" & v_strfilteredEmployee & " and DEPT_NUM = "
        Else
            str = str & ")"
        End If
    Next
    
    BuildDEPTFilterClause = MDAC25FilterFix(str)

End Function



Private Function CheckShiftReportedLunchHours() As Boolean
'  Will Check Schedule Hours against Reported Hours for each shift time exists.
'  If Reported Hours are greater than Scheduled Hours, report error.
    Dim ldatChargeDate As Date
    Dim ldatLastDate As Date
    Dim ldecReportedLUHours As Variant
    Dim lcolSchedShifts As Collection
    Dim eShift As cboScheduledShift
    Dim lcolEntries As Collection
    Dim lcolErrDescArgs As New Collection
    Dim lstrErrDesc As String
    Dim lblnError As Boolean
    Dim l_datDatePlus1 As Date
    
On Error GoTo eh

    ldatChargeDate = m_datCollPerStart
    l_datDatePlus1 = DateAdd("d", 1, Date)
    If m_datCollPerEnd > l_datDatePlus1 Then
        ldatLastDate = l_datDatePlus1
    Else
        ldatLastDate = m_datCollPerEnd
    End If
    Do Until ldatChargeDate > ldatLastDate
        Set lcolSchedShifts = Timecard.Schedule.ScheduledShiftsOn(ldatChargeDate)
        Set lcolEntries = Timecard.TimecardEntriesOnChargeDate(ldatChargeDate)
        For Each eShift In lcolSchedShifts  'assign entries to shift
            ldecReportedLUHours = SumReportedLunchHoursByShift(eShift, lcolEntries)
            If ldecReportedLUHours > eShift.ShiftLunchHours Then
                lcolErrDescArgs.Add Format(ldatChargeDate, "mm/dd/yyyy")
'                lstrErrDesc = _
'                    TrsGetErrDescription(LongAppErr(trErrTimecardHoursMoreSched), lcolErrDescArgs)
                lstrErrDesc = "Lunch hours exceed the hours of the normally scheduled lunch."
                Set lcolErrDescArgs = Nothing
                AppendError ErrorSet, _
                    Timecard.Employee.EmployeeNumber, _
                    ldatChargeDate, ldatChargeDate, _
                    "", "", CStr(ldecReportedLUHours), _
                    "", "", _
                    vbObjectError + 20000, mstrCBUSINESSWAR, _
                    lstrErrDesc, m_strCMODULENAME
'                lblnError = True
            End If
        Next
        ldatChargeDate = DateAdd("d", 1, ldatChargeDate)
    Loop
        
    Set eShift = Nothing
    Set lcolSchedShifts = Nothing
    Set lcolEntries = Nothing
    CheckShiftReportedLunchHours = lblnError
    Exit Function
    
eh:
    Set eShift = Nothing
    Set lcolSchedShifts = Nothing
    Set lcolEntries = Nothing
    TrsRaiseError m_strCMODULENAME, "CheckShiftReportedLunchHours"

End Function


Private Function SumReportedLunchHoursByShift( _
                    ByRef robjSchedShift As cboScheduledShift, _
                    ByRef rcolEntries As Collection) As Variant
' 8/2/01 RSG
' Include Lunch Unpaid(LU) and Lunch Premium(LP) in the result.
'
    Dim eEntry As cboTimecardEntry
    Dim ldecLunchHours As Variant
    Dim lcolShiftEntries As Collection
    
On Error GoTo eh

    Set lcolShiftEntries = _
        GetEntriesWithinShift(robjSchedShift, rcolEntries)
    ldecLunchHours = CDec(0)
    For Each eEntry In lcolShiftEntries
        If eEntry.ExceptionCode = "LU" Or eEntry.ExceptionCode = "LP" Then
            ldecLunchHours = ldecLunchHours + eEntry.Hours
        End If
    Next

    SumReportedLunchHoursByShift = ldecLunchHours
    Set eEntry = Nothing
    Set lcolShiftEntries = Nothing
    Exit Function
    
eh:
    Set eEntry = Nothing
    Set lcolShiftEntries = Nothing
    TrsRaiseError m_strCMODULENAME, "SumReportedLunchHoursByShift"

End Function


Private Function CheckShiftReportedOvertimeHours() As Boolean
'  Will Check Overtime Hours against Reported Hours for each shift time exists.
'  If Reported Hours are greater than Scheduled Hours, report error.

'Release 07/25/2001 -- JFB      Made changes to not add overtime
'                                 within shift if a vacation was
'                                 reported on that day


    Dim ldatChargeDate As Date
    Dim ldatLastDate As Date
    Dim ldecReportedOTHours As Variant
    Dim ldecReportedLUHours As Variant
    Dim lcolSchedShifts As Collection
    Dim eShift As cboScheduledShift
    Dim lcolEntries As Collection
    Dim lcolErrDescArgs As New Collection
    Dim lstrErrDesc As String
    Dim lblnError As Boolean
    Dim l_datDatePlus1 As Date
    
On Error GoTo eh

    ldatChargeDate = m_datCollPerStart
    l_datDatePlus1 = DateAdd("d", 1, Date)
    If m_datCollPerEnd > l_datDatePlus1 Then
        ldatLastDate = l_datDatePlus1
    Else
        ldatLastDate = m_datCollPerEnd
    End If
    Do Until ldatChargeDate > ldatLastDate
        If Not Timecard.IsDateAHoliday(ldatChargeDate) _
            And Not Timecard.IsDateAVacation(ldatChargeDate) _
            And Not Timecard.IsDateASpecialOverride(ldatChargeDate) _
            And Not Timecard.IsDateAHolidayBanked(ldatChargeDate) Then
            Set lcolSchedShifts = Timecard.Schedule.ScheduledShiftsOn(ldatChargeDate)
            Set lcolEntries = Timecard.TimecardEntriesOnChargeDate(ldatChargeDate)
            For Each eShift In lcolSchedShifts  'assign entries to shift
                ldecReportedLUHours = SumReportedLunchHoursByShift(eShift, lcolEntries)
                ldecReportedOTHours = SumReportedOvertimeHoursByShift(eShift, lcolEntries)
                If ldecReportedOTHours > eShift.ShiftLunchHours Then
                    lcolErrDescArgs.Add Format(ldatChargeDate, "mm/dd/yyyy")
    '                lstrErrDesc = _
    '                    TrsGetErrDescription(LongAppErr(trErrTimecardHoursMoreSched), lcolErrDescArgs)
                    lstrErrDesc = "Overtime hours within scheduled shift exceed the hours of the normally scheduled lunch."
                    Set lcolErrDescArgs = Nothing
                    AppendError ErrorSet, _
                        Timecard.Employee.EmployeeNumber, _
                        ldatChargeDate, ldatChargeDate, _
                        "", "", CStr(ldecReportedLUHours), _
                        "", "", _
                        vbObjectError + 30000, mstrCBUSINESSERR, _
                        lstrErrDesc, m_strCMODULENAME
                    lblnError = True
                Else
                    If (ldecReportedOTHours + ldecReportedLUHours) > eShift.ShiftLunchHours _
                        And ldecReportedOTHours > 0 Then
                        lcolErrDescArgs.Add Format(ldatChargeDate, "mm/dd/yyyy")
        '                lstrErrDesc = _
        '                    TrsGetErrDescription(LongAppErr(trErrTimecardHoursMoreSched), lcolErrDescArgs)
                        lstrErrDesc = "Overtime hours and Lunch Unpaid hours within scheduled shift exceed the hours of the normally scheduled lunch."
                        Set lcolErrDescArgs = Nothing
                        AppendError ErrorSet, _
                            Timecard.Employee.EmployeeNumber, _
                            ldatChargeDate, ldatChargeDate, _
                            "", "", CStr(ldecReportedLUHours), _
                            "", "", _
                            vbObjectError + 30000, mstrCBUSINESSERR, _
                            lstrErrDesc, m_strCMODULENAME
                        lblnError = True
                    End If
                End If
            Next
        End If
        ldatChargeDate = DateAdd("d", 1, ldatChargeDate)
    Loop
        
    Set eShift = Nothing
    Set lcolSchedShifts = Nothing
    Set lcolEntries = Nothing
    CheckShiftReportedOvertimeHours = lblnError
    Exit Function
    
eh:
    Set eShift = Nothing
    Set lcolSchedShifts = Nothing
    Set lcolEntries = Nothing
    TrsRaiseError m_strCMODULENAME, "CheckShiftReportedOvertimeHours"

End Function


Private Function SumReportedOvertimeHoursByShift( _
                    ByRef robjSchedShift As cboScheduledShift, _
                    ByRef rcolEntries As Collection) As Variant
    Dim eEntry As cboTimecardEntry
    Dim ldecOvertimeHours As Variant
    Dim lcolShiftEntries As Collection
    
On Error GoTo eh

    Set lcolShiftEntries = _
        GetEntriesWithinShift(robjSchedShift, rcolEntries)
    ldecOvertimeHours = CDec(0)
    For Each eEntry In lcolShiftEntries
        If eEntry.OvertimeOpportunityNumber > 0 _
          And Not IsOverrideThatAllowsOverlap(eEntry.OverrideExceptionCode) Then
            ldecOvertimeHours = ldecOvertimeHours + eEntry.Hours
        End If
    Next

    SumReportedOvertimeHoursByShift = ldecOvertimeHours
    Set eEntry = Nothing
    Set lcolShiftEntries = Nothing
    Exit Function
    
eh:
    Set eEntry = Nothing
    Set lcolShiftEntries = Nothing
    TrsRaiseError m_strCMODULENAME, "SumReportedOvertimeHoursByShift"

End Function

Private Function CheckScheduleTimeStamp() As Boolean
'  Will Check Schedule to see if the schedule has changed
' before save
' Returns True if the schedule has changed
    Dim strScheduleTimeStamp As String
    Dim strEmpModScheduleTimeStamp As String
    Dim lstrErrDesc As String
    
On Error GoTo eh

    strEmpModScheduleTimeStamp = GetScheduleTimeStamp()
    strScheduleTimeStamp = Timecard.Schedule.ScheduleTimeStamp
    lstrErrDesc = "The time can not be saved because the schedule has changed.  " & _
        "Please exit Time Collection and come back in for this employee. "
    If strScheduleTimeStamp <> strEmpModScheduleTimeStamp Then
        AppendError ErrorSet, _
        Timecard.Employee.EmployeeNumber, _
        m_datCollPerStart, m_datCollPerStart, _
        "", "", "", _
        "", "", _
        vbObjectError + 30000, mstrCBUSINESSERR, _
        lstrErrDesc, m_strCMODULENAME
        CheckScheduleTimeStamp = True
    Else
        CheckScheduleTimeStamp = False
    End If

    Exit Function
    
eh:

    TrsRaiseError m_strCMODULENAME, "CheckScheduleTimeStamp"

End Function

Private Function GetScheduleTimeStamp() As String

    Dim l_lngEmpNum As Long
    Dim l_obj As dbSaveTimecard.cdSaveTimecard
    
    On Error GoTo eh
    
    l_lngEmpNum = Timecard.Employee.EmployeeNumber
    Set l_obj = CreateObject("dbSaveTimecard.cdSaveTimeCard")
    GetScheduleTimeStamp = Trim(l_obj.GetScheduleTimeStamp(l_lngEmpNum))

    Set l_obj = Nothing
    Exit Function
    
eh:
    Set l_obj = Nothing
    TrsRaiseError m_strCMODULENAME, "GetScheduleTimeStamp"
End Function
Public Function RecordEqualizedOvertime( _
            ByVal v_rsTimeEntries As ADODB.Recordset, _
            ByVal r_rsEmployees As ADODB.Recordset, _
            ByVal r_rsEmpWorkgroups As ADODB.Recordset, _
            ByRef r_rsTimecardEntries As ADODB.Recordset, _
            ByVal r_rsSchedules As ADODB.Recordset, _
            ByVal v_startDate As Date, _
            ByVal v_endDate As Date, _
            ByRef r_rsErrors As ADODB.Recordset, _
            ByVal v_UserId As String) As Boolean
    Dim objCtx As ObjectContext
    Dim l_lngEmployeeNumber As Long
    Dim l_datChargeDate As Date
    Dim l_blnResult As Boolean

On Error GoTo eh
    
    Set objCtx = GetObjectContext
    
    m_strUserId = v_UserId
    Set m_objErrors = r_rsErrors

    'for server based interfaces
    PreArrangedOTIndMissing = Not FieldExists("PREARRANGED_OT_IND", v_rsTimeEntries)

    If Not GetKeyEmployeeData( _
                v_rsTimeEntries, _
                Nothing, _
                l_lngEmployeeNumber, _
                l_datChargeDate) Then
            RecordEqualizedOvertime = False
            Exit Function
    End If

    m_datCollPerStart = v_startDate
    m_datCollPerEnd = v_endDate
    
    RestoreTimecard r_rsTimecardEntries
    RestoreEmployee r_rsEmployees, r_rsEmpWorkgroups
    If r_rsSchedules Is Nothing Then
        RestoreSchedule _
            CLng(l_lngEmployeeNumber), m_datCollPerStart, m_datCollPerEnd
    ElseIf r_rsSchedules.RecordCount = 0 Then
        RestoreSchedule _
            CLng(l_lngEmployeeNumber), m_datCollPerStart, m_datCollPerEnd
    Else
        RestoreScheduleWith r_rsSchedules, Nothing
    End If


    If RecordsExist(v_rsTimeEntries) Then
        l_blnResult = RecordOvertimeEntries(v_rsTimeEntries)
        If Not l_blnResult Then
            objCtx.SetAbort
            RecordEqualizedOvertime = False
            GoTo xit
        End If
    End If

'?
    Timecard.ResetRecordStatus
    
    objCtx.SetComplete
    RecordEqualizedOvertime = True
    
xit:

    Set r_rsErrors = m_objErrors
    Release
    Set objCtx = Nothing
    Exit Function

eh:
    objCtx.SetAbort
    Release
    Set m_objErrors = Nothing
    Set r_rsErrors = Nothing
    Set r_rsTimecardEntries = Nothing
    Set objCtx = Nothing
    TrsRaiseError m_strCMODULENAME, "RecordEqualizedOvertime", , , , , True

End Function


Public Function RecordEqualizedOvertime2( _
            ByVal v_rsTimecardRequest As ADODB.Recordset, _
            ByVal v_rsOtherPayRequest As ADODB.Recordset, _
            ByVal r_rsEmployees As ADODB.Recordset, _
            ByVal r_rsEmpWorkgroups As ADODB.Recordset, _
            ByRef r_rsTimecardEntries As ADODB.Recordset, _
            ByRef r_rsOtherpayEntries As ADODB.Recordset, _
            ByVal r_rsSchedules As ADODB.Recordset, _
            ByVal v_startDate As Date, _
            ByVal v_endDate As Date, _
            ByRef r_rsErrors As ADODB.Recordset, _
            ByVal v_UserId As String) As Boolean
    Dim objCtx As ObjectContext
    Dim l_lngEmployeeNumber As Long
    Dim l_datChargeDate As Date
    Dim l_blnResult As Boolean

On Error GoTo eh
    
    Set objCtx = GetObjectContext
    
    m_strUserId = v_UserId
    Set m_objErrors = r_rsErrors

    'for server based interfaces
    PreArrangedOTIndMissing = Not FieldExists("PREARRANGED_OT_IND", v_rsTimecardRequest)

    If Not GetKeyEmployeeData( _
                v_rsTimecardRequest, _
                v_rsOtherPayRequest, _
                l_lngEmployeeNumber, _
                l_datChargeDate) Then
        RecordEqualizedOvertime2 = False
        Exit Function
    End If

    m_datCollPerStart = v_startDate
    m_datCollPerEnd = v_endDate
    
    RestoreTimecard r_rsTimecardEntries
    Timecard.SetCallerOvertime
    
    RestoreEmployee r_rsEmployees, r_rsEmpWorkgroups
    If r_rsSchedules Is Nothing Then
        RestoreSchedule _
            CLng(l_lngEmployeeNumber), m_datCollPerStart, m_datCollPerEnd
    ElseIf r_rsSchedules.RecordCount = 0 Then
        RestoreSchedule _
            CLng(l_lngEmployeeNumber), m_datCollPerStart, m_datCollPerEnd
    Else
        RestoreScheduleWith r_rsSchedules, Nothing
    End If
    If RecordsExist(v_rsOtherPayRequest) Then
        RestoreOtherPay r_rsOtherpayEntries, True
    End If

    If RecordsExist(v_rsTimecardRequest) Then
        l_blnResult = RecordOvertimeEntries(v_rsTimecardRequest)
        If Not l_blnResult Then
            objCtx.SetAbort
            RecordEqualizedOvertime2 = False
            GoTo xit
        End If
    End If

    If RecordsExist(v_rsOtherPayRequest) Then
        l_blnResult = RecordOtherPayEntries(v_rsOtherPayRequest)
        If Not l_blnResult Then
            objCtx.SetAbort
            RecordEqualizedOvertime2 = False
            GoTo xit
        End If
    End If

    Timecard.ResetRecordStatus
    
    objCtx.SetComplete
    RecordEqualizedOvertime2 = True
    
xit:

    Set r_rsErrors = m_objErrors
    Release
    Set objCtx = Nothing
    Exit Function

eh:
    objCtx.SetAbort
    Release
    Set m_objErrors = Nothing
    Set r_rsErrors = Nothing
    Set r_rsTimecardEntries = Nothing
    Set objCtx = Nothing
    TrsRaiseError m_strCMODULENAME, "RecordEqualizedOvertime2", , , , , True

End Function



Private Function RecordOvertimeEntries( _
            ByRef v_rsTimeEntries As ADODB.Recordset) As Boolean
' Answer false when a validation error occurs.  This will send back
' the brv recordset with all errors since a vb error was not raised.
    Dim l_varEmpChgDates As Variant     'array of empNum and ChgDate
    Dim lstrEmpNum As String
    Dim lcolErrArgs As Collection
    Dim lstrOverExc As String
    
    
On Error GoTo eh

    RecordOvertimeEntries = True
    v_rsTimeEntries.MoveFirst
    lstrEmpNum = v_rsTimeEntries!KEY_EMP_NUM
    l_varEmpChgDates = GetEmployeeChargeDates(v_rsTimeEntries)
    
' Process Deletes first
    If RecordsExist(v_rsTimeEntries) Then
        v_rsTimeEntries.MoveFirst
    End If
    Do Until v_rsTimeEntries.EOF
        If v_rsTimeEntries!RECORD_STATUS = "D" Then
            Timecard.RecordTimeEntry v_rsTimeEntries
        End If
        v_rsTimeEntries.MoveNext
    Loop


' Process holidays first

    If RecordsExist(v_rsTimeEntries) Then
        v_rsTimeEntries.MoveFirst
    End If
    Do Until v_rsTimeEntries.EOF
        If v_rsTimeEntries!RECORD_STATUS <> "D" Then
            If IsNull(v_rsTimeEntries!EXCEPTION_CD) Then
            Else
                If IsAHolidayException(v_rsTimeEntries!EXCEPTION_CD) Then
                    Timecard.RecordTimeEntry v_rsTimeEntries
                End If
            End If
        End If
        v_rsTimeEntries.MoveNext
    Loop
    
    
' Process vacations next

    If RecordsExist(v_rsTimeEntries) Then
        v_rsTimeEntries.MoveFirst
    End If
    Do Until v_rsTimeEntries.EOF
        If v_rsTimeEntries!RECORD_STATUS <> "D" Then
            If IsNull(v_rsTimeEntries!EXCEPTION_CD) Then
            Else
                If IsAVacationException(v_rsTimeEntries!EXCEPTION_CD) Then
                    Timecard.RecordTimeEntry v_rsTimeEntries
                End If
            End If
        End If
        v_rsTimeEntries.MoveNext
    Loop

' Process Special Override Codes next
    If RecordsExist(v_rsTimeEntries) Then
        v_rsTimeEntries.MoveFirst
    End If
    Do Until v_rsTimeEntries.EOF
        If v_rsTimeEntries!RECORD_STATUS <> "D" Then
            If IsNull(v_rsTimeEntries!OVER_EXC_CD) Then
                lstrOverExc = ""
            Else
                lstrOverExc = v_rsTimeEntries!OVER_EXC_CD
                If IsSpecialOverrideCode(lstrOverExc) Then
                    Timecard.RecordTimeEntry v_rsTimeEntries
                End If
            End If
        End If
        v_rsTimeEntries.MoveNext
    Loop

'Process all other time (except holidays)
    
    v_rsTimeEntries.MoveFirst
    Do Until v_rsTimeEntries.EOF
        If v_rsTimeEntries!RECORD_STATUS <> "D" Then
            If IsNull(v_rsTimeEntries!EXCEPTION_CD) Then
                Timecard.RecordTimeEntry v_rsTimeEntries
            Else
                If IsNull(v_rsTimeEntries!OVER_EXC_CD) Then
                    lstrOverExc = ""
                Else
                    lstrOverExc = v_rsTimeEntries!OVER_EXC_CD
                End If
                If IsAHolidayException(v_rsTimeEntries!EXCEPTION_CD) Or IsAVacationException(v_rsTimeEntries!EXCEPTION_CD) Or IsSpecialOverrideCode(lstrOverExc) Then
                    'Skip already processed
                Else
                     Timecard.RecordTimeEntry v_rsTimeEntries
                End If
            End If
        End If
        v_rsTimeEntries.MoveNext
    Loop

    If ValidateEqualizedOvertime Then
        If Save(l_varEmpChgDates, True) Then
            RecordOvertimeEntries = True
        Else
            RecordOvertimeEntries = False
            RecordOTEntrySaveBRV v_rsTimeEntries
        End If
    Else
        RecordOvertimeEntries = False
    End If
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "RecordOvertimeEntries"

End Function

Private Function ValidateEqualizedOvertime() As Boolean
    Dim entry As cboTimecardEntry
    Dim answer As Boolean

On Error GoTo eh

    ValidateEqualizedOvertime = True
    For Each entry In Timecard.TimecardEntries
        If entry.IsNew Or entry.IsModified Then
            If Not entry.IsValid Then
                ValidateEqualizedOvertime = False
                Exit Function
            End If
        End If
    Next
    
    CheckReportedHours
    
    If CheckForOverlappingTime Then
        ValidateEqualizedOvertime = False
        Exit Function
    End If

'    answer = ValidateCanNonCanOvertime
'    If Not answer Then
'        ValidateEqualizedOvertime = answer
'        Exit Function
'    End If

    answer = CheckShiftReportedHours
    If answer Then
        ValidateEqualizedOvertime = False
        Exit Function
    End If
    
    answer = CheckShiftReportedLunchHours
    If answer Then
        ValidateEqualizedOvertime = False
        Exit Function
    End If
    
    answer = CheckShiftReportedOvertimeHours
    If answer Then
        ValidateEqualizedOvertime = False
        Exit Function
    End If
    
    answer = CheckScheduleTimeStamp
    If answer Then
        ValidateEqualizedOvertime = False
        Exit Function
    End If
    
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "ValidateEqualizedOvertime"
    
End Function
'June 11, 2007
Private Function GetAccountDescriptionJobTask(ByVal vadrsJobTaskDescription As ADODB.Recordset _
                      , ByVal v_strDistributionId As String _
                      , ByVal v_strJobTask As String _
                      , ByVal v_strEMPRVDBid As String _
                      , Optional ByVal v_intDescriptionLength As Integer = 39) As String
On Error GoTo GetAccountDescriptionJobTask_EH

    Dim sBaseFilter As String
    GetAccountDescriptionJobTask = Trim(v_strJobTask)

    If RecordsExist(vadrsJobTaskDescription) Then
        'exact match
        sBaseFilter = "JOB_NBR = '" & Mid(v_strJobTask, 1, 8) & "'" & _
            " AND TASK_NBR = '" & Mid(v_strJobTask, 9, 2) & "'" & _
            " AND UTLTY_CD = '" & Mid(v_strDistributionId, 1, 1) & "'" & _
            " AND BUS_DV_CD = '" & Mid(v_strDistributionId, 2, 2) & "'" & _
            " AND MAJ_CD = '" & Mid(v_strDistributionId, 4, 3) & "'" & _
            " AND MIN_CD = '" & Mid(v_strDistributionId, 7, 3) & "'" & _
            " AND FMC_CD = '" & Mid(v_strDistributionId, 10, 3) & "'" & _
            " AND RMC_CD = '" & Mid(v_strDistributionId, 13, 3) & "'" & _
            " AND TRN_TP_CD = '" & Mid(v_strDistributionId, 16, 1) & "'" & _
            " AND PROJECT_ID = '" & Mid(v_strDistributionId, 17, 5) & "'" & _
            " AND PROD_CD =  '" & Mid(v_strDistributionId, 22, 2) & "'" & _
            " AND ACTIV_CD = '" & Mid(v_strDistributionId, 24, 4) & "'" & _
            " AND RES_TP_CD = '" & Mid(v_strDistributionId, 28, 2) & "'"
        vadrsJobTaskDescription.Filter = GetAccountDescriptionJobTask_AddFilterForDBId(sBaseFilter, v_strEMPRVDBid)
        If Not (vadrsJobTaskDescription.EOF And vadrsJobTaskDescription.BOF) Then
            vadrsJobTaskDescription.MoveFirst

            GetAccountDescriptionJobTask = PadRight(vadrsJobTaskDescription!DESC60, v_intDescriptionLength)
            Exit Function
        End If

        vadrsJobTaskDescription.Filter = adFilterNone

        'remove rmc and rt
        sBaseFilter = "JOB_NBR = '" & Mid(v_strJobTask, 1, 8) & "'" & _
            " AND TASK_NBR = '" & Mid(v_strJobTask, 9, 2) & "'" & _
            " AND UTLTY_CD = '" & Mid(v_strDistributionId, 1, 1) & "'" & _
            " AND BUS_DV_CD = '" & Mid(v_strDistributionId, 2, 2) & "'" & _
            " AND MAJ_CD = '" & Mid(v_strDistributionId, 4, 3) & "'" & _
            " AND MIN_CD = '" & Mid(v_strDistributionId, 7, 3) & "'" & _
            " AND FMC_CD = '" & Mid(v_strDistributionId, 10, 3) & "'" & _
            " AND TRN_TP_CD = '" & Mid(v_strDistributionId, 16, 1) & "'" & _
            " AND PROJECT_ID = '" & Mid(v_strDistributionId, 17, 5) & "'" & _
            " AND PROD_CD =  '" & Mid(v_strDistributionId, 22, 2) & "'" & _
            " AND ACTIV_CD = '" & Mid(v_strDistributionId, 24, 4) & "'"
        vadrsJobTaskDescription.Filter = GetAccountDescriptionJobTask_AddFilterForDBId(sBaseFilter, v_strEMPRVDBid)
        If Not (vadrsJobTaskDescription.EOF And vadrsJobTaskDescription.BOF) Then
            vadrsJobTaskDescription.MoveFirst

            GetAccountDescriptionJobTask = PadRight(vadrsJobTaskDescription!DESC60, v_intDescriptionLength)
            Exit Function
        End If
        
        vadrsJobTaskDescription.Filter = adFilterNone
        'remove rmc and fmc and rt
        sBaseFilter = "JOB_NBR = '" & Mid(v_strJobTask, 1, 8) & "'" & _
            " AND TASK_NBR = '" & Mid(v_strJobTask, 9, 2) & "'" & _
            " AND UTLTY_CD = '" & Mid(v_strDistributionId, 1, 1) & "'" & _
            " AND BUS_DV_CD = '" & Mid(v_strDistributionId, 2, 2) & "'" & _
            " AND MAJ_CD = '" & Mid(v_strDistributionId, 4, 3) & "'" & _
            " AND MIN_CD = '" & Mid(v_strDistributionId, 7, 3) & "'" & _
            " AND TRN_TP_CD = '" & Mid(v_strDistributionId, 16, 1) & "'" & _
            " AND PROJECT_ID = '" & Mid(v_strDistributionId, 17, 5) & "'" & _
            " AND PROD_CD =  '" & Mid(v_strDistributionId, 22, 2) & "'" & _
            " AND ACTIV_CD = '" & Mid(v_strDistributionId, 24, 4) & "'"
            vadrsJobTaskDescription.Filter = GetAccountDescriptionJobTask_AddFilterForDBId(sBaseFilter, v_strEMPRVDBid)
        If Not (vadrsJobTaskDescription.EOF And vadrsJobTaskDescription.BOF) Then
            vadrsJobTaskDescription.MoveFirst

            GetAccountDescriptionJobTask = PadRight(vadrsJobTaskDescription!DESC60, v_intDescriptionLength)
            Exit Function
        End If

        vadrsJobTaskDescription.Filter = adFilterNone
        'remove rmc and rt and actv
        sBaseFilter = "JOB_NBR = '" & Mid(v_strJobTask, 1, 8) & "'" & _
            " AND TASK_NBR = '" & Mid(v_strJobTask, 9, 2) & "'" & _
            " AND UTLTY_CD = '" & Mid(v_strDistributionId, 1, 1) & "'" & _
            " AND BUS_DV_CD = '" & Mid(v_strDistributionId, 2, 2) & "'" & _
            " AND MAJ_CD = '" & Mid(v_strDistributionId, 4, 3) & "'" & _
            " AND MIN_CD = '" & Mid(v_strDistributionId, 7, 3) & "'" & _
            " AND FMC_CD = '" & Mid(v_strDistributionId, 10, 3) & "'" & _
            " AND TRN_TP_CD = '" & Mid(v_strDistributionId, 16, 1) & "'" & _
            " AND PROJECT_ID = '" & Mid(v_strDistributionId, 17, 5) & "'" & _
            " AND PROD_CD =  '" & Mid(v_strDistributionId, 22, 2) & "'"
        vadrsJobTaskDescription.Filter = GetAccountDescriptionJobTask_AddFilterForDBId(sBaseFilter, v_strEMPRVDBid)
        If Not (vadrsJobTaskDescription.EOF And vadrsJobTaskDescription.BOF) Then
            vadrsJobTaskDescription.MoveFirst

            GetAccountDescriptionJobTask = PadRight(vadrsJobTaskDescription!DESC60, v_intDescriptionLength)
            Exit Function
        End If

        vadrsJobTaskDescription.Filter = adFilterNone
        'remove rmc and rt and fmc and actv
        sBaseFilter = "JOB_NBR = '" & Mid(v_strJobTask, 1, 8) & "'" & _
            " AND TASK_NBR = '" & Mid(v_strJobTask, 9, 2) & "'" & _
            " AND UTLTY_CD = '" & Mid(v_strDistributionId, 1, 1) & "'" & _
            " AND BUS_DV_CD = '" & Mid(v_strDistributionId, 2, 2) & "'" & _
            " AND MAJ_CD = '" & Mid(v_strDistributionId, 4, 3) & "'" & _
            " AND MIN_CD = '" & Mid(v_strDistributionId, 7, 3) & "'" & _
            " AND TRN_TP_CD = '" & Mid(v_strDistributionId, 16, 1) & "'" & _
            " AND PROJECT_ID = '" & Mid(v_strDistributionId, 17, 5) & "'" & _
            " AND PROD_CD =  '" & Mid(v_strDistributionId, 22, 2) & "'"
        vadrsJobTaskDescription.Filter = GetAccountDescriptionJobTask_AddFilterForDBId(sBaseFilter, v_strEMPRVDBid)
        If Not (vadrsJobTaskDescription.EOF And vadrsJobTaskDescription.BOF) Then
            vadrsJobTaskDescription.MoveFirst

            GetAccountDescriptionJobTask = PadRight(vadrsJobTaskDescription!DESC60, v_intDescriptionLength)
            Exit Function
        End If

        vadrsJobTaskDescription.Filter = adFilterNone
        'remove fmc and actv
        sBaseFilter = "JOB_NBR = '" & Mid(v_strJobTask, 1, 8) & "'" & _
            " AND TASK_NBR = '" & Mid(v_strJobTask, 9, 2) & "'" & _
            " AND UTLTY_CD = '" & Mid(v_strDistributionId, 1, 1) & "'" & _
            " AND BUS_DV_CD = '" & Mid(v_strDistributionId, 2, 2) & "'" & _
            " AND MAJ_CD = '" & Mid(v_strDistributionId, 4, 3) & "'" & _
            " AND MIN_CD = '" & Mid(v_strDistributionId, 7, 3) & "'" & _
            " AND RMC_CD = '" & Mid(v_strDistributionId, 13, 3) & "'" & _
            " AND TRN_TP_CD = '" & Mid(v_strDistributionId, 16, 1) & "'" & _
            " AND PROJECT_ID = '" & Mid(v_strDistributionId, 17, 5) & "'" & _
            " AND PROD_CD =  '" & Mid(v_strDistributionId, 22, 2) & "'" & _
            " AND RES_TP_CD = '" & Mid(v_strDistributionId, 28, 2) & "'"
        vadrsJobTaskDescription.Filter = GetAccountDescriptionJobTask_AddFilterForDBId(sBaseFilter, v_strEMPRVDBid)
        If Not (vadrsJobTaskDescription.EOF And vadrsJobTaskDescription.BOF) Then
            vadrsJobTaskDescription.MoveFirst

            GetAccountDescriptionJobTask = PadRight(vadrsJobTaskDescription!DESC60, v_intDescriptionLength)
            Exit Function
        End If
        
        vadrsJobTaskDescription.Filter = adFilterNone
        'remove actv
        sBaseFilter = "JOB_NBR = '" & Mid(v_strJobTask, 1, 8) & "'" & _
            " AND TASK_NBR = '" & Mid(v_strJobTask, 9, 2) & "'" & _
            " AND UTLTY_CD = '" & Mid(v_strDistributionId, 1, 1) & "'" & _
            " AND BUS_DV_CD = '" & Mid(v_strDistributionId, 2, 2) & "'" & _
            " AND MAJ_CD = '" & Mid(v_strDistributionId, 4, 3) & "'" & _
            " AND MIN_CD = '" & Mid(v_strDistributionId, 7, 3) & "'" & _
            " AND FMC_CD = '" & Mid(v_strDistributionId, 10, 3) & "'" & _
            " AND RMC_CD = '" & Mid(v_strDistributionId, 13, 3) & "'" & _
            " AND TRN_TP_CD = '" & Mid(v_strDistributionId, 16, 1) & "'" & _
            " AND PROJECT_ID = '" & Mid(v_strDistributionId, 17, 5) & "'" & _
            " AND PROD_CD =  '" & Mid(v_strDistributionId, 22, 2) & "'" & _
            " AND RES_TP_CD = '" & Mid(v_strDistributionId, 28, 2) & "'"
        vadrsJobTaskDescription.Filter = GetAccountDescriptionJobTask_AddFilterForDBId(sBaseFilter, v_strEMPRVDBid)
        If Not (vadrsJobTaskDescription.EOF And vadrsJobTaskDescription.BOF) Then
            vadrsJobTaskDescription.MoveFirst

            GetAccountDescriptionJobTask = PadRight(vadrsJobTaskDescription!DESC60, v_intDescriptionLength)
            Exit Function
        End If

        vadrsJobTaskDescription.Filter = adFilterNone
    End If

    Exit Function
GetAccountDescriptionJobTask_EH:
    TrsRaiseError "cbsTimeCard", "GetAccountDescriptionJobTask" _
        , , , , , True

End Function

'//also used in busDistributionMgr.cbsDistributionMgr
Private Function GetAccountDescriptionJobTask_AddFilterForDBId(ByVal v_strBaseFilter As String, ByVal v_strEMPRVDBid As String) As String
On Error GoTo eh
            '//can only use OR at top level of filter
            '//so the following will not work, Arguments are of the wrong type, are out of acceptable range, or are in conflict with one another.
'            vadrsJobTaskDescription.Filter = _
'                v_strBaseFilter & _
'            " AND (FK_EMPRVDBCODE = '" & v_strEMPRVDBid & "')" & _
'            "     OR FK_EMPRVDBCODE = '')" & _
'            "     OR " & IIf(v_strEMPRVDBid = "", "True", "False") & ")"
    '//logical equivalent
    '//match on db id, or allow emprv's db id to be blank, or allow tris's db id to be blank
    GetAccountDescriptionJobTask_AddFilterForDBId = _
            "(" & v_strBaseFilter & _
            "    AND FK_EMPRVDBCODE = '" & v_strEMPRVDBid & "')" & _
            " OR ( " & v_strBaseFilter & _
            "    AND FK_EMPRVDBCODE = '')" & _
            " OR ( " & v_strBaseFilter & _
            "    AND " & IIf(v_strEMPRVDBid = "", "FK_EMPRVDBCODE<>'NotFound(Result=True)'", "FK_EMPRVDBCODE='NotFound(Result=False)'") & ")"
Exit Function
eh:
    TrsRaiseError "cbsTimeCard", "GetAccountDescriptionJobTask_AddFilterForDBId", , , , , True
End Function

Private Sub RecordOTEntrySaveBRV(ByVal v_rsTimeEntries As ADODB.Recordset)

    On Error GoTo eh

    Dim l_strErrorDesc As String

        
        
        If Not m_objContractRulesErrors Is Nothing Then
            If m_objContractRulesErrors.RecordCount > 0 Then
                m_objContractRulesErrors.MoveFirst
                While Not m_objContractRulesErrors.EOF
                    ErrorSet.AddNew
                    Dim intFieldIndex As Integer
                    For intFieldIndex = 0 To ErrorSet.Fields.Count - 1
                        ErrorSet.Fields(intFieldIndex).Value = m_objContractRulesErrors.Fields(intFieldIndex).Value
                    Next
                    ErrorSet.Update
                    m_objContractRulesErrors.MoveNext
                Wend
                l_strErrorDesc = "The Contract Rule Automation Code found an error"
            Else
                l_strErrorDesc = "Problem Saving Overtime in Time Collection. Please contact TRIS team"
            End If
        End If
        
        v_rsTimeEntries.MoveFirst
        AppendError ErrorSet, _
                    Timecard.Employee.EmployeeNumber, _
                    v_rsTimeEntries!CHARGE_DATE, _
                    v_rsTimeEntries!DATE_WORKED, _
                    "", _
                    "", _
                    "", _
                    "", _
                    "", _
                    30000, _
                    mstrCAPPLICATIONERR, _
                    l_strErrorDesc, _
                    m_strCMODULENAME


    Exit Sub
eh:
        TrsRaiseError "cbsTimeCard", "RecordOTEntrySaveBRV" _
        , , , , , True

End Sub

Private Function CheckHolidayDate() As Boolean

    Dim entry As cboTimecardEntry
    Dim answer As Boolean

    answer = True
    For Each entry In Timecard.TimecardEntries
        If entry.IsNew Or entry.IsModified Then
            answer = IsValidHolidayDate(entry)
        End If
    Next

    CheckHolidayDate = Not answer

End Function

Private Function CheckHolidayBanked() As Boolean

    Dim entry As cboTimecardEntry
    Dim answer As Boolean

    answer = False
    For Each entry In Timecard.TimecardEntries
        If entry.IsNew Or entry.IsModified Then
            If Not IsValidHolidayDate(entry, True) Then
                answer = True
            End If
        End If
    Next

    CheckHolidayBanked = answer

End Function



Private Function CheckHolidayHours() As Boolean

    Dim answer As Boolean
    Dim lstrErrDesc As String
    
    Dim l_datChargeDate As Variant
    Dim l_datLastDate As Date
    Dim l_decReportedHours As Variant
    Dim lcolErrDescArgs As New Collection

    
On Error GoTo eh

    answer = False
    
    If Not LimitEightHoursOfHoliday() Then
        CheckHolidayHours = False
        Exit Function
    End If
    
    l_datChargeDate = m_datCollPerStart
    If m_datCollPerEnd > Date Then
        l_datLastDate = Date
    Else
        l_datLastDate = m_datCollPerEnd
    End If
    Do Until l_datChargeDate > l_datLastDate
        l_decReportedHours = SumHolidayHours(l_datChargeDate)
        If l_decReportedHours > 8 Then
            lstrErrDesc = "The time can not be saved because the date " & str(l_datChargeDate) & " because more the 8 hours was reported."
            AppendError ErrorSet, Timecard.Employee.EmployeeNumber, l_datChargeDate, l_datChargeDate, _
                "", "", l_decReportedHours, "H", "", _
                vbObjectError + 30000, mstrCBUSINESSERR, lstrErrDesc, m_strCMODULENAME
            answer = True
        End If
        l_datChargeDate = DateAdd("d", 1, l_datChargeDate)
    Loop
   

    CheckHolidayHours = answer
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "CheckHolidayHours"
    

End Function

Private Function LimitEightHoursOfHoliday() As Boolean

    Dim l_obj As dbRules.cdRules
    Dim l_strOTExclusion As String
    Dim l_strDepartmentCode As String
    
    On Error GoTo eh
    
    Set l_obj = CreateObject("dbRules.cdRules")
    
    If Not Timecard.Employee.IsExecutive() Then
        LimitEightHoursOfHoliday = False
        Exit Function
    End If
    
    
    If l_obj.DoesRuleApply(600, trLocationCode, Trim(str(Timecard.Employee.LocationCode))) Then
        LimitEightHoursOfHoliday = False
        Exit Function
    End If
    
    l_strDepartmentCode = Trim(Timecard.Employee.DepartmentNumber)
    l_strOTExclusion = Trim(Timecard.Employee.OTExclusn)
    
    If l_obj.DoesRuleApply(601, trDepartmentNumber, l_strDepartmentCode) And l_strOTExclusion = "0000" Then
        LimitEightHoursOfHoliday = False
        Exit Function
    End If
    
    If l_obj.DoesRuleApply(602, trDepartmentNumber, l_strDepartmentCode) And l_strOTExclusion = "0000" Then
        LimitEightHoursOfHoliday = False
        Exit Function
    End If
    
    
    If l_obj.DoesRuleApply(600, trDepartmentNumber, l_strDepartmentCode) Then
        LimitEightHoursOfHoliday = False
        Exit Function
    End If
        
    LimitEightHoursOfHoliday = True
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "LimitEightHoursOfHoliday"
        
End Function
Private Function SumHolidayHours(ByVal v_datChargeDate As Date) As Variant
' Only sum time within schedule, unless it is callout standby or lunch.
' May also add an optional parameter for sum option to designate exceptions
' to summing all.
' May use IsAnyPeriodWithinSchedule since all time is split on schedule.
    Dim eEntry As cboTimecardEntry
    Dim l_decHours As Variant

On Error GoTo eh

    l_decHours = CDec(0)
    For Each eEntry In Timecard.TimecardEntries
        If Not eEntry.IsDeleted Then
            If eEntry.ChargeDate = v_datChargeDate Then
                If RequiresHolidayCalander(Trim(eEntry.ExceptionCode)) Then
                    l_decHours = l_decHours + eEntry.Hours
                End If
            End If
        End If
    Next
    
    SumHolidayHours = l_decHours
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "SumHolidayHours"

End Function
Private Sub GetHolidays()

On Error GoTo eh

    Dim l_obj As dbCodeTables.cdCodeTables
'    Dim l_col As Collection
'    Dim l_strUnion
    
'    l_strUnion = Timecard.Employee.UnionCode()
'    Set l_col = New Collection
'    l_col.Add l_strUnion

    Set l_obj = GetJobManager()
    
    Set m_rsHolidays = l_obj.GetHolidaysBetween(m_datCollPerStart, m_datCollPerEnd)

    Exit Sub

eh:
    TrsRaiseError m_strCMODULENAME, "GetHolidays"
    
End Sub


Private Function IsValidHolidayDate(ByVal entry As cboTimecardEntry, Optional vblnHB As Boolean = False) As Boolean

On Error GoTo eh

    Dim l_blnIsValid As Boolean
    Dim l_lngEmpNum As Long
    Dim l_datChargeDate As Date
    Dim l_strExceptionCode As String
    Dim l_strEmployeeUnion As String
    Dim l_decHours As Variant
    Dim lstrErrDesc As String
    Dim l_blnFlexEmp As Boolean

    l_strExceptionCode = Trim(entry.ExceptionCode)
    If Not RequiresHolidayCalander(l_strExceptionCode) Then
        IsValidHolidayDate = True
        Exit Function
    End If
    
    l_lngEmpNum = entry.EmployeeNumber
    l_datChargeDate = entry.ChargeDate
    l_strEmployeeUnion = entry.UnionCode
    l_decHours = entry.Hours
    lstrErrDesc = ""

    If LimitEightHoursOfHoliday() Then
        l_blnFlexEmp = True
    Else
        l_blnFlexEmp = False
    End If

    l_blnIsValid = True
    If m_rsHolidays Is Nothing Then
        lstrErrDesc = "Date " & l_datChargeDate & " is not on the Holiday Calendar."
    Else
        m_rsHolidays.Filter = "HOLIDAY_DATE = " & str(l_datChargeDate) & " and UNION_CD = '" & l_strEmployeeUnion & "' "
        If m_rsHolidays.BOF And m_rsHolidays.EOF Then
            lstrErrDesc = "Date " & l_datChargeDate & " is not on the Holiday Calendar."
        End If
        m_rsHolidays.Filter = ""
    End If
    
    
    If lstrErrDesc <> "" Then
'        If vblnHB And l_strExceptionCode = "PBH" Then
'            lstrErrDesc = "The PBH time can not be saved because the date " & str(entry.ChargeDate) & "  is not a Holiday."
'                AppendError ErrorSet, Timecard.Employee.EmployeeNumber, entry.ChargeDate, entry.DateWorked, _
'                    entry.startTime, entry.EndTime, entry.Hours, entry.ExceptionCode, entry.OverrideExceptionCode, _
'                    vbObjectError + 30000, mstrCBUSINESSERR, lstrErrDesc, m_strCMODULENAME
'            l_blnIsValid = False
'        End If
        If vblnHB And l_strExceptionCode = "HB" Then
            lstrErrDesc = "The time can not be saved because the date " & str(entry.ChargeDate) & " of the Holiday Banked is not on the Holiday Calendar."
                AppendError ErrorSet, Timecard.Employee.EmployeeNumber, entry.ChargeDate, entry.DateWorked, _
                    entry.startTime, entry.EndTime, entry.Hours, entry.ExceptionCode, entry.OverrideExceptionCode, _
                    vbObjectError + 30000, mstrCBUSINESSERR, lstrErrDesc, m_strCMODULENAME
            l_blnIsValid = False
        Else
            If l_blnFlexEmp Then
                AppendError ErrorSet, Timecard.Employee.EmployeeNumber, entry.ChargeDate, entry.DateWorked, _
                    entry.startTime, entry.EndTime, entry.Hours, entry.ExceptionCode, entry.OverrideExceptionCode, _
                    vbObjectError + 30000, mstrCBUSINESSERR, lstrErrDesc, m_strCMODULENAME
                l_blnIsValid = False
            Else
                AppendError ErrorSet, l_lngEmpNum, l_datChargeDate, l_datChargeDate, "", "", CStr(l_decHours), "", "", _
                    20000, mstrCBUSINESSWAR, lstrErrDesc, m_strCMODULENAME
                l_blnIsValid = True
            End If
        End If
    Else
        If l_blnFlexEmp And l_strExceptionCode = "HR" Then
            lstrErrDesc = "Holiday on a Day of Rest (HR) was changed to Holiday Banked (HB)."
            AppendError ErrorSet, l_lngEmpNum, l_datChargeDate, l_datChargeDate, "", "", CStr(l_decHours), "", "", _
                    20000, mstrCBUSINESSWAR, lstrErrDesc, m_strCMODULENAME
            entry.ExceptionCode = "HB"
            entry.DistributionId = "Holiday Banked"
            entry.DistributionType = ""
            l_blnIsValid = True
        Else
            l_blnIsValid = True
        End If
    End If

    IsValidHolidayDate = l_blnIsValid
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "IsValidHolidayDate"

End Function

Private Function RequiresHolidayCalander(ByVal strExc) As Boolean

    Select Case strExc
        Case "H", "HR", "HB"
            RequiresHolidayCalander = True
        Case Else
            RequiresHolidayCalander = False
    End Select

End Function

Private Function GetAccountDescriptionJobTaskMaximo(ByRef radrsJobTaskDescription As ADODB.Recordset _
                      , ByVal v_strJobNum As String _
                      , ByVal v_strTaskNum As String _
                      , ByVal v_strSiteId As String, _
                      Optional ByVal v_intDescriptionLength As Integer = 39) As String
On Error GoTo eh

    Dim l_obj As dbDistributionMgr.cdDistributionMgr
    Dim l_rs As ADODB.Recordset
    Dim l_sTaskDesc As String
    Dim ProjectLength As Integer
            
    ProjectLength = (v_intDescriptionLength - Len(DISTRIBUTION_DESCRIPTION_SEPARATOR) + 1) / 2
        
    If RecordsExist(radrsJobTaskDescription) Then
        If v_strSiteId <> "NULL" And v_strSiteId <> "" Then
            radrsJobTaskDescription.Filter = "JOB_NBR = '" & v_strJobNum & "' AND TASK_NBR = '" & v_strTaskNum & "' AND FK_EMPRVDBCODE = '" & v_strSiteId & "'"
        Else
            radrsJobTaskDescription.Filter = "JOB_NBR = '" & v_strJobNum & "' AND TASK_NBR = '" & v_strTaskNum & "'"
        End If
    End If
    
    If RecordsExist(radrsJobTaskDescription) Then
        radrsJobTaskDescription.MoveFirst
        '//GetAccountDescriptionJobTaskMaximo = PadRight(radrsJobTaskDescription!DESC60, v_intDescriptionLength)
        l_sTaskDesc = Mid(radrsJobTaskDescription!JOBTASK_AND_DESC, 12)
        If v_intDescriptionLength = 30 Then '//old TRIS client 6.9.4
            GetAccountDescriptionJobTaskMaximo = PadRight(l_sTaskDesc, v_intDescriptionLength)
'        ElseIf v_intDescriptionLength = 48 Then '//new TRIS client 6.9.7
'            '//Alt-186 double vertical bars
'            GetAccountDescriptionJobTaskMaximo = PadRight(Left(radrsJobTaskDescription!DESC60, 24) & "¦" & l_sTaskDesc, v_intDescriptionLength)
'        Else '//new OTTC client
'            GetAccountDescriptionJobTaskMaximo = PadRight(Left(radrsJobTaskDescription!DESC60, 30) & "¦" & l_sTaskDesc, v_intDescriptionLength)
        Else
            GetAccountDescriptionJobTaskMaximo = PadRight(PadRight(radrsJobTaskDescription!DESC60, ProjectLength) & DISTRIBUTION_DESCRIPTION_SEPARATOR & l_sTaskDesc, v_intDescriptionLength)
        End If
        radrsJobTaskDescription.Filter = adFilterNone
        Exit Function
    End If
    
    If Not radrsJobTaskDescription Is Nothing Then
        radrsJobTaskDescription.Filter = adFilterNone
    End If
        
    Set l_obj = CreateObject("dbDistributionMgr.cdDistributionMgr")
    Set l_rs = l_obj.GetDistJobTaskListbyJob2(v_strJobNum, v_strTaskNum, v_strSiteId)
    Set l_obj = Nothing
    
    If RecordsExist(l_rs) Then
        '//GetAccountDescriptionJobTaskMaximo = PadRight(l_rs!DESC60, v_intDescriptionLength)
        l_sTaskDesc = Mid(l_rs!JOBTASK_AND_DESC, 12)
        If v_intDescriptionLength = 30 Then '//old TRIS client
            GetAccountDescriptionJobTaskMaximo = PadRight(l_sTaskDesc, v_intDescriptionLength)
'        ElseIf v_intDescriptionLength = 48 Then '//new TRIS client 6.9.7
'            '//Alt-186 double vertical bars
'            GetAccountDescriptionJobTaskMaximo = PadRight(Left(l_rs!DESC60, 24) & "¦" & l_sTaskDesc, v_intDescriptionLength)
'        Else '//new OTTC client
'            GetAccountDescriptionJobTaskMaximo = PadRight(Left(l_rs!DESC60, 30) & "¦" & l_sTaskDesc, v_intDescriptionLength)
        Else
            GetAccountDescriptionJobTaskMaximo = PadRight(PadRight(l_rs!DESC60, ProjectLength) & DISTRIBUTION_DESCRIPTION_SEPARATOR & l_sTaskDesc, v_intDescriptionLength)
        End If
        If RecordsExist(radrsJobTaskDescription) Then
            l_rs.MoveFirst
            radrsJobTaskDescription.AddNew
            radrsJobTaskDescription!JOB_NBR = l_rs!JOB_NBR
            radrsJobTaskDescription!TASK_NBR = l_rs!TASK_NBR
            radrsJobTaskDescription!FK_EMPRVDBCODE = l_rs!FK_EMPRVDBCODE
            radrsJobTaskDescription!DESC60 = l_rs!DESC60
            radrsJobTaskDescription!JOBTASK_AND_DESC = l_rs!JOBTASK_AND_DESC
            radrsJobTaskDescription!CORP = l_rs!CORP
            radrsJobTaskDescription!UTLTY_CD = l_rs!UTLTY_CD
            radrsJobTaskDescription!BUS_DV_CD = l_rs!BUS_DV_CD
            radrsJobTaskDescription!MAJ_CD = l_rs!MAJ_CD
            radrsJobTaskDescription!MIN_CD = l_rs!MIN_CD
            radrsJobTaskDescription!FMC_CD = l_rs!FMC_CD
            radrsJobTaskDescription!RMC_CD = l_rs!RMC_CD
            radrsJobTaskDescription!TRN_TP_CD = l_rs!TRN_TP_CD
            radrsJobTaskDescription!PROJECT_ID = l_rs!PROJECT_ID
            radrsJobTaskDescription!PROD_CD = l_rs!PROD_CD
            radrsJobTaskDescription!ACTIV_CD = l_rs!ACTIV_CD
            radrsJobTaskDescription!RES_TP_CD = l_rs!RES_TP_CD
            radrsJobTaskDescription!FEEDER_REF = l_rs!FEEDER_REF
            radrsJobTaskDescription!WG_NUM = l_rs!WG_NUM
            radrsJobTaskDescription!COSTCNTRCODE = ""
            radrsJobTaskDescription!CCDESC = ""
            radrsJobTaskDescription!JOBSTCODE = ""
            radrsJobTaskDescription.Update
        Else
            Set radrsJobTaskDescription = l_rs
        End If
    Else
        GetAccountDescriptionJobTaskMaximo = PadRight("CAN NOT FIND DESCRIPTION IN MAXIMO", v_intDescriptionLength)
    End If
    
    '//radrsJobTaskDescription.Filter = adFilterNone

    Exit Function
eh:
    TrsRaiseError "cbsTimeCard", "GetAccountDescriptionJobTaskMaximo" _
        , , , , , True

End Function


Private Function SumReportedParentalBondingHoursByShift( _
                    ByRef robjSchedShift As cboScheduledShift, _
                    ByRef rcolEntries As Collection) As Variant
    Dim eEntry As cboTimecardEntry
    Dim ldecShiftHours As Variant
    Dim lcolShiftEntries As Collection
    
On Error GoTo eh

    Set lcolShiftEntries = _
        GetEntriesWithinShift(robjSchedShift, rcolEntries)
    ldecShiftHours = CDec(0)
    For Each eEntry In lcolShiftEntries
        If eEntry.OvertimeOpportunityNumber = 0 And _
            Trim(eEntry.ExceptionCode) = "PBL" Then
            ldecShiftHours = ldecShiftHours + eEntry.Hours
        End If
    Next

    SumReportedParentalBondingHoursByShift = ldecShiftHours
    Set eEntry = Nothing
    Set lcolShiftEntries = Nothing
    Exit Function
    
eh:
    Set eEntry = Nothing
    Set lcolShiftEntries = Nothing
    TrsRaiseError m_strCMODULENAME, "SumReportedParentalBondingHoursByShift"

End Function

Private Function BuildMAXIMOJobTasks() As ADODB.Recordset

On Error GoTo eh
    Set BuildMAXIMOJobTasks = Timecard.BuildMAXIMOJobTasks
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "BuildMAXIMOJobTasks"

End Function


Public Function RecordTimeEntry4( _
            ByVal v_rsTimeEntries As ADODB.Recordset, _
            ByVal v_rsOtherPay As ADODB.Recordset, _
            ByVal r_rsEmployees As ADODB.Recordset, _
            ByVal r_rsEmpWorkgroups As ADODB.Recordset, _
            ByRef r_rsTimecardEntries As ADODB.Recordset, _
            ByRef r_rsOtherpayEntries As ADODB.Recordset, _
            ByVal r_rsSchedules As ADODB.Recordset, _
            ByVal v_rsScheduleChgPmts As ADODB.Recordset, _
            ByVal v_startDate As Date, _
            ByVal v_endDate As Date, _
            ByRef r_rsAccrualBankUpdates As ADODB.Recordset, _
            ByRef r_rsOvertimeBankUpdates As ADODB.Recordset, _
            ByRef r_colEMPRVInterface As Collection, _
            ByRef r_rsEMPRVJobTaskEntries As ADODB.Recordset, _
            ByRef r_rsMAXIMOJobTaskEntries As ADODB.Recordset, _
            ByRef r_rsErrors As ADODB.Recordset, _
            ByVal v_UserId As String) As Boolean
    Dim objCtx As ObjectContext
    Dim l_lngEmployeeNumber As Long
    Dim l_datChargeDate As Date
    Dim l_blnResult As Boolean
    Dim l_rsAccrualBankUpdates As ADODB.Recordset
    Dim l_rsOvertimeBankUpdates As ADODB.Recordset
    Dim l_colEMPRVInterface As Collection
    Dim l_rsEMPRVJobTaskEntries As ADODB.Recordset
    Dim l_rsMAXIMOJobTaskEntries As ADODB.Recordset
    Dim l_intI As Integer
    Dim l_intCount As Integer

On Error GoTo eh
    
'test only
    #If TRACENEON Then
        Dim lngTraceEmpNum As Long
        lngTraceEmpNum = v_rsTimeEntries!KEY_EMP_NUM
        TraceNeonMessage CStr(lngTraceEmpNum), "RECDTE", "BEGIN"
    #End If
    
    #If TRTRACE Then
        TraceMessage "", "Begin RecordTimeEntry3"
    #End If

    Set objCtx = GetObjectContext
    
    m_strUserId = v_UserId
    Set m_objErrors = r_rsErrors

    'for older clients (6.0.0) that are not sending this column to the server
    PreArrangedOTIndMissing = Not FieldExists("PREARRANGED_OT_IND", v_rsTimeEntries)

    If Not GetKeyEmployeeData( _
                v_rsTimeEntries, _
                v_rsOtherPay, _
                l_lngEmployeeNumber, _
                l_datChargeDate) Then
'        If RecordsExist(v_rsECDrawTime) Then
'            GoTo ecdraw
'        Else
            RecordTimeEntry4 = False
            Exit Function
'        End If
    End If

    m_datCollPerStart = v_startDate
    m_datCollPerEnd = v_endDate
    
    #If TRTRACE Then
        TraceMessage "", "Begin RestoreTimecard"
    #End If
    RestoreTimecard r_rsTimecardEntries
    #If TRTRACE Then
        TraceMessage "", "End RestoreTimecard"
    #End If
'    RestoreTimecard _
'        CLng(l_lngEmployeeNumber), m_datCollPerStart, m_datCollPerEnd
    #If TRTRACE Then
        TraceMessage "", "Begin RestoreEmployee"
    #End If
    RestoreEmployee r_rsEmployees, r_rsEmpWorkgroups
    #If TRTRACE Then
        TraceMessage "", "End RestoreEmployee"
    #End If
'    RestoreEmployee _
'        CLng(l_lngEmployeeNumber), m_datCollPerStart, m_datCollPerEnd
    #If TRTRACE Then
        TraceMessage "", "Begin RestoreSchedule"
    #End If
    If r_rsSchedules Is Nothing Then
        RestoreSchedule _
            CLng(l_lngEmployeeNumber), m_datCollPerStart, m_datCollPerEnd
    ElseIf r_rsSchedules.RecordCount = 0 Then
        RestoreSchedule _
            CLng(l_lngEmployeeNumber), m_datCollPerStart, m_datCollPerEnd
    Else
        #If TRTRACE Then
            TraceMessage "", "RestoreSchedule state provided"
        #End If
        RestoreScheduleWith r_rsSchedules, v_rsScheduleChgPmts
    End If
    #If TRTRACE Then
        TraceMessage "", "End RestoreSchedule"
    #End If
' Restore other pay here, instead of after save because
' it will not cause a deadlock here even if the transaction
' mode is not set to required.
    If RecordsExist(v_rsOtherPay) Then
        #If TRTRACE Then
            TraceMessage "", "Begin RestoreOtherPay"
        #End If
        RestoreOtherPay r_rsOtherpayEntries, False
'        RestoreOtherPay l_lngEmployeeNumber, m_datCollPerStart, m_datCollPerEnd
        #If TRTRACE Then
            TraceMessage "", "End RestoreOtherPay"
        #End If
    End If

    If RecordsExist(v_rsTimeEntries) Then
        #If TRTRACE Then
            TraceMessage "", "Begin RecordTimecardEntries"
        #End If
    
        l_blnResult = RecordTimecardEntries(v_rsTimeEntries)
            If Not m_objContractRulesErrors Is Nothing Then
                If m_objContractRulesErrors.RecordCount > 0 Then
                    m_objContractRulesErrors.MoveFirst
                    While Not m_objContractRulesErrors.EOF
                        r_rsErrors.AddNew
                        
                        Dim intFieldIndex As Integer
                        For intFieldIndex = 0 To r_rsErrors.Fields.Count - 1
                            r_rsErrors.Fields(intFieldIndex).Value = m_objContractRulesErrors.Fields(intFieldIndex).Value
                        Next
                        
                        r_rsErrors.Update
                        
                        m_objContractRulesErrors.MoveNext
                    Wend
                End If
            End If
            
            If Not m_objOTTCConversionErrors Is Nothing Then
                If m_objOTTCConversionErrors.RecordCount > 0 Then
                    m_objOTTCConversionErrors.MoveFirst
                    While Not m_objOTTCConversionErrors.EOF
                        r_rsErrors.AddNew
                        
                        For intFieldIndex = 0 To r_rsErrors.Fields.Count - 1
                            r_rsErrors.Fields(intFieldIndex).Value = m_objOTTCConversionErrors.Fields(intFieldIndex).Value
                        Next
                        
                        r_rsErrors.Update
                        
                        m_objOTTCConversionErrors.MoveNext
                    Wend
                End If
            End If

        
        If Not l_blnResult Then
            objCtx.SetAbort
            RecordTimeEntry4 = False
            GoTo xit
        End If
        
        #If TRTRACE Then
            TraceMessage "", "End RecordTimecardEntries"
        #End If
'        if a recordset is passed then append to it,
'        otherwise simply return the updates.

        #If TRTRACE Then
            TraceMessage "", "Begin BuildAccrualBankUpdates"
        #End If
        Set l_rsAccrualBankUpdates = BuildAccrualBankUpdates
        #If TRTRACE Then
            TraceMessage "", "End BuildAccrualBankUpdates"
        #End If
            
        #If TRTRACE Then
            TraceMessage "", "Begin BuildOvertimeBankUpdates"
        #End If
        Set l_rsOvertimeBankUpdates = BuildOvertimeBankUpdates
        #If TRTRACE Then
            TraceMessage "", "End BuildOvertimeBankUpdates"
        #End If
        
        #If TRTRACE Then
            TraceMessage "", "Begin BuildEMPRVInterface"
        #End If
        Set l_colEMPRVInterface = BuildEMPRVInterface
        #If TRTRACE Then
            TraceMessage "", "End BuildEMPRVInterface"
        #End If
        
        #If TRTRACE Then
            TraceMessage "", "Begin BuildEMPRVJobTasks"
        #End If
        Set l_rsEMPRVJobTaskEntries = BuildEMPRVJobTasks
        #If TRTRACE Then
            TraceMessage "", "End BuildEMPRVJobTasks"
        #End If
                
        #If TRTRACE Then
            TraceMessage "", "Begin BuildMAXIMOJobTasks"
        #End If
        Set l_rsMAXIMOJobTaskEntries = BuildMAXIMOJobTasks
        #If TRTRACE Then
            TraceMessage "", "End BuildMAXIMOJobTasks"
        #End If
    
        If r_rsAccrualBankUpdates Is Nothing Then
            Set r_rsAccrualBankUpdates = l_rsAccrualBankUpdates
        Else
            AppendRecordset l_rsAccrualBankUpdates, r_rsAccrualBankUpdates
        End If
        Set l_rsAccrualBankUpdates = Nothing
        If r_rsOvertimeBankUpdates Is Nothing Then
            Set r_rsOvertimeBankUpdates = l_rsOvertimeBankUpdates
        Else
            AppendRecordset l_rsOvertimeBankUpdates, r_rsOvertimeBankUpdates
        End If
        Set l_rsOvertimeBankUpdates = Nothing
        If r_colEMPRVInterface Is Nothing Then
            Set r_colEMPRVInterface = l_colEMPRVInterface
        Else
            For l_intI = 1 To l_colEMPRVInterface.Count
                r_colEMPRVInterface.Add l_colEMPRVInterface(l_intI)
            Next l_intI
        End If
        Set l_colEMPRVInterface = Nothing
        If r_rsEMPRVJobTaskEntries Is Nothing Then
            Set r_rsEMPRVJobTaskEntries = l_rsEMPRVJobTaskEntries
        Else
            AppendRecordset l_rsEMPRVJobTaskEntries, r_rsEMPRVJobTaskEntries
        End If
        Set l_rsEMPRVJobTaskEntries = Nothing
        
        If r_rsMAXIMOJobTaskEntries Is Nothing Then
            Set r_rsMAXIMOJobTaskEntries = l_rsMAXIMOJobTaskEntries
        Else
            AppendRecordset l_rsMAXIMOJobTaskEntries, r_rsMAXIMOJobTaskEntries
        End If
        Set l_rsEMPRVJobTaskEntries = Nothing
        
    End If

    If RecordsExist(v_rsOtherPay) Then
        #If TRTRACE Then
            TraceMessage "", "Begin RecordOtherPayEntries"
        #End If
'        RecordOtherPayEntries v_rsOtherPay, r_rsErrors
        l_blnResult = RecordOtherPayEntries(v_rsOtherPay)
        If Not l_blnResult Then
            objCtx.SetAbort
            RecordTimeEntry4 = False
            GoTo xit
        End If
        
        #If TRTRACE Then
            TraceMessage "", "End RecordOtherPayEntries"
        #End If
    End If

    #If TRTRACE Then
        TraceMessage "", "Begin ResetRecordStatus"
    #End If
    Timecard.ResetRecordStatus
    If Not OtherPay Is Nothing Then
        OtherPay.ResetRecordStatus
    End If
    #If TRTRACE Then
        TraceMessage "", "End ResetRecordStatus"
    #End If
    
'ecdraw:
'    If RecordsExist(v_rsECDrawTime) Then
'        l_blnResult = _
'            RecordECDrawTime(v_rsECDrawTime, v_UserId, r_rsErrors)
'        If Not l_blnResult Then
'            objCtx.SetAbort
'            GoTo xit
'        End If
'    End If
    
    objCtx.SetComplete
    RecordTimeEntry4 = True
    
xit:

    #If TRACENEON Then
        TraceNeonMessage CStr(lngTraceEmpNum), "RECDTE", "EXIT"
    #End If
    
'set byref regardless of rec count.
    Set r_rsErrors = m_objErrors
    
    Release
    Set objCtx = Nothing
    
    #If TRTRACE Then
        TraceMessage "", "End RecordTimeEntry4"
    #End If
    Exit Function

eh:
    objCtx.SetAbort
    Release
    Set m_objErrors = Nothing
    Set r_rsErrors = Nothing
    Set r_rsAccrualBankUpdates = Nothing
    Set r_rsOvertimeBankUpdates = Nothing
    Set r_rsTimecardEntries = Nothing
    Set r_rsOtherpayEntries = Nothing
    Set r_rsEMPRVJobTaskEntries = Nothing
    Set r_rsMAXIMOJobTaskEntries = Nothing
    Set objCtx = Nothing
    TrsRaiseError m_strCMODULENAME, "RecordTimeEntry4", , , , , True

End Function



Private Function CheckHolidayParentalBonding() As Boolean

    Dim entry As cboTimecardEntry
    Dim answer As Boolean
    Dim lstrErrDesc As String
    

    answer = False
'''''    For Each entry In Timecard.TimecardEntries
'''''        If entry.IsNew Or entry.IsModified Then
'''''            If Not IsValidHolidayDate(entry, True) Then
'''''                answer = True
'''''            End If
'''''        End If
'''''    Next
    
    '        If vblnHB And l_strExceptionCode = "PBH" Then
'            lstrErrDesc = "The PBH time can not be saved because the date " & str(entry.ChargeDate) & "  is not a Holiday."
'                AppendError ErrorSet, Timecard.Employee.EmployeeNumber, entry.ChargeDate, entry.DateWorked, _
'                    entry.startTime, entry.EndTime, entry.Hours, entry.ExceptionCode, entry.OverrideExceptionCode, _
'                    vbObjectError + 30000, mstrCBUSINESSERR, lstrErrDesc, m_strCMODULENAME
'            l_blnIsValid = False
'        End If
    
    If answer = False Then
        For Each entry In Timecard.TimecardEntries
            If entry.IsNew Or entry.IsModified Then
                If Trim(entry.ExceptionCode) = "PBL" And IsOnHolidayCalendar(entry) Then
                    lstrErrDesc = "PBL can't be entered on a Holiday " & str(entry.ChargeDate) & "."
                    AppendError ErrorSet, Timecard.Employee.EmployeeNumber, entry.ChargeDate, entry.DateWorked, _
                        entry.startTime, entry.EndTime, entry.Hours, entry.ExceptionCode, entry.OverrideExceptionCode, _
                        vbObjectError + 30000, mstrCBUSINESSERR, lstrErrDesc, m_strCMODULENAME
                    answer = True
                End If
                If Trim(entry.ExceptionCode) = "PBH" And Not IsOnHolidayCalendar(entry) Then
                    lstrErrDesc = "PBH can only be entered on a Holiday " & str(entry.ChargeDate) & "."
                    AppendError ErrorSet, Timecard.Employee.EmployeeNumber, entry.ChargeDate, entry.DateWorked, _
                        entry.startTime, entry.EndTime, entry.Hours, entry.ExceptionCode, entry.OverrideExceptionCode, _
                        vbObjectError + 30000, mstrCBUSINESSERR, lstrErrDesc, m_strCMODULENAME
                    answer = True
                End If
            End If
        Next
    End If

    CheckHolidayParentalBonding = answer

End Function



Private Function CheckMAXIMO() As Boolean

    Dim entry As cboTimecardEntry
    Dim answer As Boolean
    Dim lstrErrDesc As String
    

    If answer = False Then
        For Each entry In Timecard.TimecardEntries
            If entry.IsNew Or entry.IsModified Then
                If Trim(entry.DistributionType) = "M" Then
                    If Trim(entry.JobNbr) = "" Or Trim(entry.JobTaskNbr) = "" Or Trim(entry.EMPRVDatabaseID) = "" Then
                        lstrErrDesc = "Please delete and re-enter.  Issue with MAXIMO data"
                        AppendError ErrorSet, Timecard.Employee.EmployeeNumber, entry.ChargeDate, entry.DateWorked, _
                            entry.startTime, entry.EndTime, entry.Hours, entry.ExceptionCode, entry.OverrideExceptionCode, _
                            vbObjectError + 30000, mstrCBUSINESSERR, lstrErrDesc, m_strCMODULENAME
                        answer = True
                    End If
                End If
            End If
        Next
    End If

    CheckMAXIMO = answer

End Function

Private Function CheckGLAccountingAfterPOET() As Boolean

    Dim entry As cboTimecardEntry
    Dim answer As Boolean
    Dim lstrErrDesc As String
    

    If answer = False Then
        If IsOTTCEffective And Not IsOTTCTransitionActive Then
            For Each entry In Timecard.TimecardEntries
                If entry.IsNew Or entry.IsModified Then
                    If (Trim(entry.DistributionType) = "A" And Not IsPOET39(entry.DistributionId)) Or Trim(entry.DistributionType) = "S" Then
                        lstrErrDesc = "Unable to save/convert GL accounting, please remove and select POET accounting instead."
                        AppendError ErrorSet, Timecard.Employee.EmployeeNumber, entry.ChargeDate, entry.DateWorked, _
                            entry.startTime, entry.EndTime, entry.Hours, entry.ExceptionCode, entry.OverrideExceptionCode, _
                            vbObjectError + 30000, mstrCBUSINESSERR, lstrErrDesc, m_strCMODULENAME
                        answer = True
                    End If
                End If
            Next
        End If
    End If

    CheckGLAccountingAfterPOET = answer

End Function

Private Function CheckGLAccountingAfterPOETOtherPay() As Boolean

    Dim entry As cboOtherPayEntry
    Dim answer As Boolean
    Dim lstrErrDesc As String
    

    If answer = False Then
        If IsOTTCEffective And Not IsOTTCTransitionActive Then
            For Each entry In OtherPay.OtherPayEntriesPrivate
                If entry.IsNew Or entry.IsModified Then
                    If (Trim(entry.DistributionType) = "A" And Not IsPOET39(entry.DistributionId)) Or Trim(entry.DistributionType) = "S" Then
                        lstrErrDesc = "Unable to save/convert GL accounting, please remove and select POET accounting instead. (other pay)"
                        AppendError ErrorSet, Timecard.Employee.EmployeeNumber, entry.ChargeDate, entry.DateWorked, _
                            entry.startTime, entry.EndTime, CStr(entry.hrs), entry.ExceptionCode, "", _
                            vbObjectError + 30000, mstrCBUSINESSERR, lstrErrDesc, m_strCMODULENAME
                        answer = True
                    End If
                End If
            Next
        End If
    End If

    CheckGLAccountingAfterPOETOtherPay = answer

End Function

Private Function IsOnHolidayCalendar(ByVal entry As cboTimecardEntry) As Boolean


On Error GoTo eh

    Dim l_blnIsHoliday As Boolean
    Dim l_lngEmpNum As Long
    Dim l_datChargeDate As Date
    Dim l_strExceptionCode As String
    Dim l_strEmployeeUnion As String

    If m_rsHolidays Is Nothing Then
        l_blnIsHoliday = False
    Else
        m_rsHolidays.Filter = "HOLIDAY_DATE = " & str(entry.ChargeDate) & " and UNION_CD = '" & entry.UnionCode & "' "
        If m_rsHolidays.BOF And m_rsHolidays.EOF Then
            l_blnIsHoliday = False
        Else
            l_blnIsHoliday = True
        End If
        m_rsHolidays.Filter = ""
    End If
    
    IsOnHolidayCalendar = l_blnIsHoliday

    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "IsOnHolidayCalendar"
    
End Function
