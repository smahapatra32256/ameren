VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "cbsSchedCodeMaintTran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const ADOFIELDATTR = adFldUpdatable Or adFldIsNullable
Private Const mstrCMODULENAME As String = "busSchedCodeMaintTrans"

Public Function SaveSchedCode(ByRef r_adrsSchedCodeEntries As ADODB.Recordset _
                        , ByVal v_lTRISUser As Long _
                        , ByRef r_adrsBRV As ADODB.Recordset) As Boolean
                        
On Error GoTo SaveSchedCode_Err

Dim l_objSchedCode As dbSchedCodeMaint.cdSchedCodeMaint
Dim l_objCodeTables As dbCodeTables.cdCodeTables
Dim l_sFilterString As String
Dim l_adrsSaveRecordset As ADODB.Recordset
Dim l_iCounter As Integer
    'Define the BRV recordset
    Set r_adrsBRV = CreateNewBusinessErrorSet
    
    'Validate schedule change entries
    If ValidInput(r_adrsSchedCodeEntries, r_adrsBRV) Then
        SaveSchedCode = True
    Else
        SaveSchedCode = False
        Exit Function
    End If
    
    Set l_adrsSaveRecordset = DefineRecordset(r_adrsSchedCodeEntries)
    
    'Convert change records to delete and insert (delete will utilize the RI)
    With r_adrsSchedCodeEntries
        .MoveFirst
        Do While Not .EOF
            l_adrsSaveRecordset.AddNew
            For l_iCounter = 0 To .Fields.Count - 1
                l_adrsSaveRecordset.Fields(l_iCounter) = .Fields(l_iCounter)
            Next l_iCounter
            If !RECORD_STATUS = "C" Then
                l_adrsSaveRecordset!RECORD_STATUS = "I"
                l_adrsSaveRecordset.Update
                l_adrsSaveRecordset.AddNew
                For l_iCounter = 0 To .Fields.Count - 1
                    l_adrsSaveRecordset.Fields(l_iCounter) = .Fields(l_iCounter)
                Next l_iCounter
                l_adrsSaveRecordset!RECORD_STATUS = "D"
                l_adrsSaveRecordset.Update
            Else
                l_adrsSaveRecordset.Update
            End If
            .MoveNext
        Loop
    End With
    
    'Apply schedule change updates to the database and the results of cascade
    Set l_objSchedCode = CreateObject("dbSchedCodeMaint.cdSchedCodeMaint")
    l_adrsSaveRecordset.MoveFirst
    If l_objSchedCode.SaveSchedCode(l_adrsSaveRecordset _
             , r_adrsBRV _
             , v_lTRISUser) Then
        Set l_objSchedCode = Nothing
        If IsRecordsetPopulated(l_adrsSaveRecordset) Then
            Set l_objCodeTables = CreateObject("dbCodeTables.cdCodeTables")
            With l_adrsSaveRecordset
                .MoveFirst
                Do While Not .EOF
                    l_objCodeTables.InvalidateAvailScheduleCache !SCHED_CD
                    .MoveNext
                Loop
            End With
            Set l_objCodeTables = Nothing
        End If
    Else
        GoTo SaveSchedCode_Err
    End If
    
    Set l_adrsSaveRecordset = Nothing
    
    Exit Function

SaveSchedCode_Err:

    SaveSchedCode = False
    Set l_objCodeTables = Nothing
    Set l_objSchedCode = Nothing
   
    TrsRaiseError mstrCMODULENAME, "SaveSchedCode", , , , , True

End Function
Public Function SchedCodeInUse(ByVal v_lSchedCode As Long) As Boolean

    On Error GoTo SchedCodeInUse_Err
    
    Dim l_objScheduleCodes As dbSchedCodeMaint.cdSchedCodeMaint
    
    'Determine if the schedule code is reference in the TRIS database
    Set l_objScheduleCodes = CreateObject("dbSchedCodeMaint.cdSchedCodeMaint")
    
    If l_objScheduleCodes.SchedCodeInUse(v_lSchedCode) Then
        SchedCodeInUse = True
    Else
        SchedCodeInUse = False
    End If
    
    Set l_objScheduleCodes = Nothing
    GetObjectContext.SetComplete
    
    Exit Function
    
SchedCodeInUse_Err:
    Set l_objScheduleCodes = Nothing
    GetObjectContext.SetAbort
    SchedCodeInUse = True
    TrsRaiseError mstrCMODULENAME, "SchedCodeInUse", , , , , True
    
End Function

Private Function IsRecordsetPopulated(ByRef r_adrsRecordset) As Boolean

    'Determine if the given recordset exists and contains data records
    IsRecordsetPopulated = False
    If Not r_adrsRecordset Is Nothing Then
        If Not (r_adrsRecordset.EOF And r_adrsRecordset.BOF) Then
            IsRecordsetPopulated = True
        End If
    End If
    
End Function

Private Function ValidInput(ByRef r_adrsSchedCodeEntries As ADODB.Recordset, _
        ByRef r_adrsBRV As ADODB.Recordset) As Boolean
        
    On Error GoTo ValidInput_Err
    'Validate the user's schedule code maintenance input
        
    Dim l_objSchedCode As dbSchedCodeMaint.cdSchedCodeMaint
    Dim l_bReturn As Boolean
    Dim l_sSchedDesc  As String
    
    ValidInput = True
    
    'Verify that there are records to save
    If Not IsRecordsetPopulated(r_adrsSchedCodeEntries) Then
        AppendBusinessError r_adrsBRV, LongAppErr(trErrbusValidationError), mstrCBUSINESSERR, _
            "Schedule code request is empty.", mstrCMODULENAME, 1, 0
        ValidInput = False
        Exit Function
    End If
    
    With r_adrsSchedCodeEntries
        .MoveFirst
        Do While Not .EOF

            Select Case !RECORD_STATUS
                Case "I"
                    If Not ValidNewSchedCode(r_adrsSchedCodeEntries, r_adrsBRV) Then
                        ValidInput = False
                    End If
                Case "D"
                    If Not ValidDeleteSchedCode(r_adrsSchedCodeEntries, r_adrsBRV) Then
                        ValidInput = False
                    End If
                Case "C"
                    If Not ValidChangeSchedCode(r_adrsSchedCodeEntries, r_adrsBRV) Then
                        ValidInput = False
                    End If
                Case Else
                    AppendSchedCodeEntryError r_adrsBRV, r_adrsSchedCodeEntries, _
                                LongAppErr(trErrbusValidationError), mstrCBUSINESSERR, "Invalid record status.  Record status must be I, C or D", "SCHED_CD"
                    ValidInput = False
            End Select
            .MoveNext
        Loop
    End With
    
    Exit Function
    
ValidInput_Err:

    TrsRaiseError mstrCMODULENAME, "ValidInput"

End Function

Private Function ValidDeleteSchedCode( _
        ByRef r_adrsSchedCodeEntries As ADODB.Recordset, _
        ByRef r_adrsBRV As ADODB.Recordset) As Boolean
        
On Error GoTo ValidDeleteSchedCode_Err

    ValidDeleteSchedCode = True
    
    'Schedule code cannot be deleted if it is referenced in the TRIS database
    If SchedCodeInUse(r_adrsSchedCodeEntries!KEY_SCHED_CD) Then
        AppendBusinessError r_adrsBRV, LongAppErr(trErrbusValidationError), mstrCBUSINESSERR, _
            "Schedule code is in use by the TRIS system.  Delete is not allowed. ", mstrCMODULENAME, 1, _
                    r_adrsSchedCodeEntries!RECORD_ID, "SCHED_DESC", , r_adrsSchedCodeEntries!SCHED_CD, , , "Y"
        ValidDeleteSchedCode = False
   End If
   
   Exit Function
   
ValidDeleteSchedCode_Err:
    TrsRaiseError mstrCMODULENAME, "ValidDeleteSchedCode"

End Function
                    
Private Function ValidNewSchedCode( _
        ByRef r_adrsSchedCodeEntries As ADODB.Recordset, _
        ByRef r_adrsBRV As ADODB.Recordset) As Boolean
        
    On Error GoTo ValidNewSchedCode_Err
    
    ValidNewSchedCode = True
        
    With r_adrsSchedCodeEntries
        If ValidSchedCodeInfo(r_adrsSchedCodeEntries, r_adrsBRV) Then
            ValidNewSchedCode = True
        Else
            ValidNewSchedCode = False
        End If
    End With
    
    Exit Function
    
ValidNewSchedCode_Err:
    TrsRaiseError mstrCMODULENAME, "ValidNewSchedCode"

End Function
Private Function ValidChangeSchedCode( _
        ByRef r_adrsSchedCodeEntries As ADODB.Recordset, _
        ByRef r_adrsBRV As ADODB.Recordset) As Boolean
        
On Error GoTo ValidChangeSchedCode_Err

    ValidChangeSchedCode = True
    
    'Schedule code cannot be changed if it is referenced in the TRIS database
    If SchedCodeInUse(r_adrsSchedCodeEntries!SCHED_CD) Then
        AppendBusinessError r_adrsBRV, LongAppErr(trErrbusValidationError), mstrCBUSINESSERR, _
            "Schedule code is in use by the TRIS system.  Schedule code cannot be changed. ", mstrCMODULENAME, 1, _
                    r_adrsSchedCodeEntries!RECORD_ID, "SCHED_DESC", , r_adrsSchedCodeEntries!SCHED_CD, , , "Y"
        ValidChangeSchedCode = False
   Else
        With r_adrsSchedCodeEntries
             If ValidSchedCodeInfo(r_adrsSchedCodeEntries, r_adrsBRV) Then
                 ValidChangeSchedCode = True
             Else
                 ValidChangeSchedCode = False
             End If
         End With
   End If
   
   Exit Function
   
ValidChangeSchedCode_Err:
    TrsRaiseError mstrCMODULENAME, "ValidChangeSchedCode"

End Function

Private Function ValidSchedCodeInfo(ByRef r_adrsSchedCodeEntries As ADODB.Recordset, _
        ByRef r_adrsBRV As ADODB.Recordset) As Boolean
    On Error GoTo ValidSchedCodeInfo_Err
        
    Dim l_dStraightTimeHrs As Double, l_dOvertimeHrs As Double
    Dim l_lApostropheCount  As Long, l_iCounter As Integer
    Dim l_sMessage As String
    
    Dim l_objCodeTables As dbCodeTables.cdCodeTables
    
    ValidSchedCodeInfo = True
    
    With r_adrsSchedCodeEntries
        'Must have a Schedule code description
        If Trim(!SCHED_DESC) = "" Then
            AppendSchedCodeEntryError r_adrsBRV, r_adrsSchedCodeEntries, _
                        trErrbusValidationError, mstrCBUSINESSERR, "Schedule code description is blank", "SCHED_DESC"
            ValidSchedCodeInfo = False
        End If
        'Schedule Code cannot contains an "&"
        If InStr(1, !SCHED_DESC, "&") > 0 Then
            AppendSchedCodeEntryError r_adrsBRV, r_adrsSchedCodeEntries, _
                        trErrbusValidationError, mstrCBUSINESSERR, "Schedule description cannot contain an ampersand (&)", "SCHED_DESC"
            ValidSchedCodeInfo = False
        End If
        
        l_lApostropheCount = 0
        For l_iCounter = 1 To Len(Trim(!SCHED_DESC))
            If Mid(!SCHED_DESC, l_iCounter, 1) = "'" Then
                l_lApostropheCount = l_lApostropheCount + 1
            End If
        Next l_iCounter
        
        'Comment length is limited to 30 characters
        If Len(Trim(!SCHED_DESC)) > (30 - l_lApostropheCount) Then
            If l_lApostropheCount > 0 Then
                l_sMessage = "Schedule description cannot exceed 30 characters. (Note: Each apostrophe counts as two characters.)"
            Else
                l_sMessage = "Schedule description cannot exceed 30 characters."
            End If
            AppendSchedCodeEntryError r_adrsBRV, r_adrsSchedCodeEntries, _
                        trErrbusValidationError, mstrCBUSINESSERR, _
                        l_sMessage, "SCHED_DESC"
            ValidSchedCodeInfo = False
        End If
        
        'Either straight time or overtime hours must be entered
        l_dStraightTimeHrs = !SUN_STR_HRS + !MON_STR_HRS + !TUE_STR_HRS + !WED_STR_HRS _
                                + !THU_STR_HRS + !FRI_STR_HRS + !SAT_STR_HRS
                                
        l_dOvertimeHrs = !SUN_OT_HOURS + !MON_OT_HOURS + !TUE_OT_HOURS + !WED_OT_HOURS _
                                + !THU_OT_HOURS + !FRI_OT_HOURS + !SAT_OT_HOURS
                                
        If l_dStraightTimeHrs = 0 And l_dOvertimeHrs = 0 Then
            AppendSchedCodeEntryError r_adrsBRV, r_adrsSchedCodeEntries, _
                        trErrbusValidationError, mstrCBUSINESSERR, "Straight time or overtime hours required.", "SUN_STR_HRS"
            ValidSchedCodeInfo = False
        End If
        
        'Start Day equal the charge day or one day prior
        If Trim(!SUN_ACT_DAY) <> "" Then
            If UCase(!SUN_ACT_DAY) <> "SUN" And UCase(!SUN_ACT_DAY) <> "SAT" Then
                AppendSchedCodeEntryError r_adrsBRV, r_adrsSchedCodeEntries, _
                            trErrbusValidationError, mstrCBUSINESSERR, _
                            "Start Day must be a valid weekday abbreviation and not more than one prior to the schedule day.", _
                            "SUN_ACT_DAY"
                ValidSchedCodeInfo = False
            End If
        End If
        If Trim(!MON_ACT_DAY) <> "" Then
            If UCase(!MON_ACT_DAY) <> "MON" And UCase(!MON_ACT_DAY) <> "SUN" Then
                AppendSchedCodeEntryError r_adrsBRV, r_adrsSchedCodeEntries, _
                            trErrbusValidationError, mstrCBUSINESSERR, _
                            "Start Day must be a valid weekday abbreviation and not more than one prior to the schedule day.", _
                            "MON_ACT_DAY"
                ValidSchedCodeInfo = False
            End If
        End If
        If Trim(!TUE_ACT_DAY) <> "" Then
            If UCase(!TUE_ACT_DAY) <> "TUE" And UCase(!TUE_ACT_DAY) <> "MON" Then
                AppendSchedCodeEntryError r_adrsBRV, r_adrsSchedCodeEntries, _
                            trErrbusValidationError, mstrCBUSINESSERR, _
                            "Start Day must be a valid weekday abbreviation and not more than one prior to the schedule day.", _
                            "TUE_ACT_DAY"
                ValidSchedCodeInfo = False
            End If
        End If
        If Trim(!WED_ACT_DAY) <> "" Then
            If UCase(!WED_ACT_DAY) <> "WED" And UCase(!WED_ACT_DAY) <> "TUE" Then
                AppendSchedCodeEntryError r_adrsBRV, r_adrsSchedCodeEntries, _
                            trErrbusValidationError, mstrCBUSINESSERR, _
                            "Start Day must be a valid weekday abbreviation and not more than one prior to the schedule day.", _
                            "WED_ACT_DAY"
                ValidSchedCodeInfo = False
            End If
        End If
        If Trim(!THU_ACT_DAY) <> "" Then
            If UCase(!THU_ACT_DAY) <> "THU" And UCase(!THU_ACT_DAY) <> "WED" Then
                AppendSchedCodeEntryError r_adrsBRV, r_adrsSchedCodeEntries, _
                            trErrbusValidationError, mstrCBUSINESSERR, _
                            "Start Day must be a valid weekday abbreviation and not more than one prior to the schedule day.", _
                            "THU_ACT_DAY"
                ValidSchedCodeInfo = False
            End If
        End If
        If Trim(!FRI_ACT_DAY) <> "" Then
            If UCase(!FRI_ACT_DAY) <> "FRI" And UCase(!FRI_ACT_DAY) <> "THU" Then
                AppendSchedCodeEntryError r_adrsBRV, r_adrsSchedCodeEntries, _
                            trErrbusValidationError, mstrCBUSINESSERR, _
                            "Start Day must be a valid weekday abbreviation and not more than one prior to the schedule day.", _
                            "FRI_ACT_DAY"
                ValidSchedCodeInfo = False
            End If
        End If
        If Trim(!SAT_ACT_DAY) <> "" Then
            If UCase(!SAT_ACT_DAY) <> "SAT" And UCase(!SAT_ACT_DAY) <> "FRI" Then
                AppendSchedCodeEntryError r_adrsBRV, r_adrsSchedCodeEntries, _
                            trErrbusValidationError, mstrCBUSINESSERR, _
                            "Start Day must be a valid weekday abbreviation and not more than one prior to the schedule day.", _
                            "SAT_ACT_DAY"
                ValidSchedCodeInfo = False
            End If
        End If
        
        'If schedule code entry is invalid, do not proceed with additional edits
        If Not ValidSchedCodeInfo Then
            Exit Function
        End If
        
        'Start day required if hours are entered for a charge day
        If !SUN_STR_HRS > 0 Or !SUN_OT_HOURS > 0 Then
            If Trim(!SUN_ACT_DAY) = "" Then
                AppendSchedCodeEntryError r_adrsBRV, r_adrsSchedCodeEntries, _
                        trErrbusValidationError, mstrCBUSINESSERR, _
                        "Straight time and/or overtime time hour(s) require a valid start day", _
                        "SUN_ACT_DAY"
                ValidSchedCodeInfo = False
            End If
        End If
        
        If !MON_STR_HRS > 0 Or !MON_OT_HOURS Then
            If Trim(!MON_ACT_DAY) = "" Then
                AppendSchedCodeEntryError r_adrsBRV, r_adrsSchedCodeEntries, _
                        trErrbusValidationError, mstrCBUSINESSERR, _
                        "Straight time and/or overtime time hour(s) require a valid start day", _
                        "MON_ACT_DAY"
                ValidSchedCodeInfo = False
            End If
        End If
        
        If !TUE_STR_HRS > 0 Or !TUE_OT_HOURS Then
            If Trim(!TUE_ACT_DAY) = "" Then
                AppendSchedCodeEntryError r_adrsBRV, r_adrsSchedCodeEntries, _
                        trErrbusValidationError, mstrCBUSINESSERR, _
                        "Straight time and/or overtime time hour(s) require a valid start day", _
                        "TUE_ACT_DAY"
                ValidSchedCodeInfo = False
            End If
        End If
        
        If !WED_STR_HRS > 0 Or !WED_OT_HOURS Then
            If Trim(!WED_ACT_DAY) = "" Then
                AppendSchedCodeEntryError r_adrsBRV, r_adrsSchedCodeEntries, _
                        trErrbusValidationError, mstrCBUSINESSERR, _
                        "Straight time and/or overtime time hour(s) require a valid start day", _
                        "WED_ACT_DAY"
                ValidSchedCodeInfo = False
            End If
        End If
        If !THU_STR_HRS > 0 Or !THU_OT_HOURS Then
            If Trim(!THU_ACT_DAY) = "" Then
                AppendSchedCodeEntryError r_adrsBRV, r_adrsSchedCodeEntries, _
                        trErrbusValidationError, mstrCBUSINESSERR, _
                        "Straight time and/or overtime time hour(s) require a valid start day", _
                        "THU_ACT_DAY"
                ValidSchedCodeInfo = False
            End If
        End If
        If !FRI_STR_HRS > 0 Or !FRI_OT_HOURS Then
            If Trim(!FRI_ACT_DAY) = "" Then
                AppendSchedCodeEntryError r_adrsBRV, r_adrsSchedCodeEntries, _
                        trErrbusValidationError, mstrCBUSINESSERR, _
                        "Straight time and/or overtime time hour(s) require a valid start day", _
                        "FRI_ACT_DAY"
                ValidSchedCodeInfo = False
            End If
        End If
        
        If !SAT_STR_HRS > 0 Or !SAT_OT_HOURS Then
            If Trim(!SAT_ACT_DAY) = "" Then
                AppendSchedCodeEntryError r_adrsBRV, r_adrsSchedCodeEntries, _
                        trErrbusValidationError, mstrCBUSINESSERR, _
                        "Straight time and/or overtime time hour(s) require a valid start day", _
                        "SAT_ACT_DAY"
                ValidSchedCodeInfo = False
            End If
        End If
        
        'If schedule code entry is invalid, do not proceed with additional edits
        If Not ValidSchedCodeInfo Then
            Exit Function
        End If
        
        If Not ValidTimeEntries(r_adrsSchedCodeEntries, r_adrsBRV, "SUN", !SUN_START_TIME, !SUN_END_TIME, _
                !SUN_LUNCH_START, !SUN_LUNCH_END, _
                !SUN_OT_START_TIME, !SUN_OT_END_TIME) Then
            ValidSchedCodeInfo = False
        End If
        If Not ValidTimeEntries(r_adrsSchedCodeEntries, r_adrsBRV, "MON", !MON_START_TIME, !MON_END_TIME, _
                !MON_LUNCH_START, !MON_LUNCH_END, _
                !MON_OT_START_TIME, !MON_OT_END_TIME) Then
            ValidSchedCodeInfo = False
        End If
        If Not ValidTimeEntries(r_adrsSchedCodeEntries, r_adrsBRV, "TUE", !TUE_START_TIME, !TUE_END_TIME, _
                !TUE_LUNCH_START, !TUE_LUNCH_END, _
                !TUE_OT_START_TIME, !TUE_OT_END_TIME) Then
            ValidSchedCodeInfo = False
        End If
        If Not ValidTimeEntries(r_adrsSchedCodeEntries, r_adrsBRV, "WED", !WED_START_TIME, !WED_END_TIME, _
                !WED_LUNCH_START, !WED_LUNCH_END, _
                !WED_OT_START_TIME, !WED_OT_END_TIME) Then
            ValidSchedCodeInfo = False
        End If
        If Not ValidTimeEntries(r_adrsSchedCodeEntries, r_adrsBRV, "THU", !THU_START_TIME, !THU_END_TIME, _
                !THU_LUNCH_START, !THU_LUNCH_END, _
                !THU_OT_START_TIME, !THU_OT_END_TIME) Then
            ValidSchedCodeInfo = False
        End If
        If Not ValidTimeEntries(r_adrsSchedCodeEntries, r_adrsBRV, "FRI", !FRI_START_TIME, !FRI_END_TIME, _
                !FRI_LUNCH_START, !FRI_LUNCH_END, _
                !FRI_OT_START_TIME, !FRI_OT_END_TIME) Then
            ValidSchedCodeInfo = False
        End If
        If Not ValidTimeEntries(r_adrsSchedCodeEntries, r_adrsBRV, "SAT", !SAT_START_TIME, !SAT_END_TIME, _
                !SAT_LUNCH_START, !SAT_LUNCH_END, _
                !SAT_OT_START_TIME, !SAT_OT_END_TIME) Then
            ValidSchedCodeInfo = False
        End If
        
        'Valid shift codes are "A","P","O","N","D" or blank
        If Not ValidShiftCode(!SUN_SHIFT_CD) Then
            AppendSchedCodeEntryError r_adrsBRV, r_adrsSchedCodeEntries, _
                    trErrbusValidationError, mstrCBUSINESSERR, _
                    "Invalid shift code for SUN.  Valid shift codes are 'A', 'P', 'O', 'N', 'D' or blank", _
                    "SUN_SHIFT_CD"
            ValidSchedCodeInfo = False
        End If
        If Not ValidShiftCode(!MON_SHIFT_CD) Then
            AppendSchedCodeEntryError r_adrsBRV, r_adrsSchedCodeEntries, _
                    trErrbusValidationError, mstrCBUSINESSERR, _
                    "Invalid shift code for MON.  Valid shift codes are 'A', 'P', 'O', 'N', 'D' or blank", _
                    "MON_SHIFT_CD"
            ValidSchedCodeInfo = False
        End If
        If Not ValidShiftCode(!TUE_SHIFT_CD) Then
            AppendSchedCodeEntryError r_adrsBRV, r_adrsSchedCodeEntries, _
                    trErrbusValidationError, mstrCBUSINESSERR, _
                    "Invalid shift code TUE.  Valid shift codes are 'A', 'P', 'O', 'N', 'D' or blank", _
                    "TUE_SHIFT_CD"
            ValidSchedCodeInfo = False
        End If
        If Not ValidShiftCode(!WED_SHIFT_CD) Then
            AppendSchedCodeEntryError r_adrsBRV, r_adrsSchedCodeEntries, _
                    trErrbusValidationError, mstrCBUSINESSERR, _
                    "Invalid shift code for WED.  Valid shift codes are 'A', 'P', 'O', 'N', 'D' or blank", _
                    "WED_SHIFT_CD"
            ValidSchedCodeInfo = False
        End If
        If Not ValidShiftCode(!THU_SHIFT_CD) Then
            AppendSchedCodeEntryError r_adrsBRV, r_adrsSchedCodeEntries, _
                    trErrbusValidationError, mstrCBUSINESSERR, _
                    "Invalid shift code for THU.  Valid shift codes are 'A', 'P', 'O', 'N', 'D' or blank", _
                    "THU_SHIFT_CD"
            ValidSchedCodeInfo = False
        End If
        If Not ValidShiftCode(!FRI_SHIFT_CD) Then
            AppendSchedCodeEntryError r_adrsBRV, r_adrsSchedCodeEntries, _
                    trErrbusValidationError, mstrCBUSINESSERR, _
                    "Invalid shift code for FRI.  Valid shift codes are 'A', 'P', 'O', 'N', 'D' or blank", _
                    "FRI_SHIFT_CD"
            ValidSchedCodeInfo = False
        End If
        
        If Not ValidShiftCode(!SAT_SHIFT_CD) Then
            AppendSchedCodeEntryError r_adrsBRV, r_adrsSchedCodeEntries, _
                    trErrbusValidationError, mstrCBUSINESSERR, _
                    "Invalid shift code for SAT.  Valid shift codes are 'A', 'P', 'O', 'N', 'D' or blank", _
                    "SAT_SHIFT_CD"
            ValidSchedCodeInfo = False
        End If
        
        'Component shift codes are validated against the component shift codes stored in the SPMCache
        'The object is defined here, so that the object is only created once for the week.
        
        Set l_objCodeTables = CreateObject("dbCodeTables.cdCodeTables")
        
        If Trim(!SUN_COMP_SHIFT_CD) <> "" Then
            If Not l_objCodeTables.ComponentCodeIsValid _
                (Trim(!SUN_COMP_SHIFT_CD), trShiftComponent) Then
                ValidSchedCodeInfo = False
                AppendSchedCodeEntryError r_adrsBRV, r_adrsSchedCodeEntries, _
                        trErrbusValidationError, mstrCBUSINESSERR, _
                        "Invalid shift component code for SUN. ", _
                        "SUN_COMP_SHIFT_CD"
                ValidSchedCodeInfo = False
            End If
        End If
        If Trim(!MON_COMP_SHIFT_CD) <> "" Then
            If Not l_objCodeTables.ComponentCodeIsValid _
                (Trim(!MON_COMP_SHIFT_CD), trShiftComponent) Then
                ValidSchedCodeInfo = False
                AppendSchedCodeEntryError r_adrsBRV, r_adrsSchedCodeEntries, _
                        trErrbusValidationError, mstrCBUSINESSERR, _
                        "Invalid shift component code for MON. ", _
                        "MON_COMP_SHIFT_CD"
            End If
        End If
        If Trim(!TUE_COMP_SHIFT_CD) <> "" Then
            If Not l_objCodeTables.ComponentCodeIsValid _
                (Trim(!TUE_COMP_SHIFT_CD), trShiftComponent) Then
                ValidSchedCodeInfo = False
                AppendSchedCodeEntryError r_adrsBRV, r_adrsSchedCodeEntries, _
                        trErrbusValidationError, mstrCBUSINESSERR, _
                        "Invalid shift component code for TUE. ", _
                        "TUE_COMP_SHIFT_CD"
            End If
        End If
        If Trim(!WED_COMP_SHIFT_CD) <> "" Then
            If Not l_objCodeTables.ComponentCodeIsValid _
                (Trim(!WED_COMP_SHIFT_CD), trShiftComponent) Then
                ValidSchedCodeInfo = False
                AppendSchedCodeEntryError r_adrsBRV, r_adrsSchedCodeEntries, _
                        trErrbusValidationError, mstrCBUSINESSERR, _
                        "Invalid shift component code for WED. ", _
                        "WED_COMP_SHIFT_CD"
            End If
        End If
        If Trim(!THU_COMP_SHIFT_CD) <> "" Then
            If Not l_objCodeTables.ComponentCodeIsValid _
                (Trim(!THU_COMP_SHIFT_CD), trShiftComponent) Then
                ValidSchedCodeInfo = False
                AppendSchedCodeEntryError r_adrsBRV, r_adrsSchedCodeEntries, _
                        trErrbusValidationError, mstrCBUSINESSERR, _
                        "Invalid shift component code for THU. ", _
                        "THU_COMP_SHIFT_CD"
            End If
        End If
        If Trim(!FRI_COMP_SHIFT_CD) <> "" Then
            If Not l_objCodeTables.ComponentCodeIsValid _
                (Trim(!FRI_COMP_SHIFT_CD), trShiftComponent) Then
                ValidSchedCodeInfo = False
                AppendSchedCodeEntryError r_adrsBRV, r_adrsSchedCodeEntries, _
                        trErrbusValidationError, mstrCBUSINESSERR, _
                        "Invalid shift component code for FRI. ", _
                        "FRI_COMP_SHIFT_CD"
            End If
        End If
        If Trim(!SAT_COMP_SHIFT_CD) <> "" Then
            If Not l_objCodeTables.ComponentCodeIsValid _
                (Trim(!SAT_COMP_SHIFT_CD), trShiftComponent) Then
                ValidSchedCodeInfo = False
                AppendSchedCodeEntryError r_adrsBRV, r_adrsSchedCodeEntries, _
                        trErrbusValidationError, mstrCBUSINESSERR, _
                        "Invalid shift component code for SAT. ", _
                        "SAT_COMP_SHIFT_CD"
            End If
        End If
        Set l_objCodeTables = Nothing
       
        'If schedule code entry is invalid, do not proceed with additional edits
        If Not ValidSchedCodeInfo Then
            ValidSchedCodeInfo = False
            Exit Function
        End If
        
        If Not ValidTimeIntervals(r_adrsSchedCodeEntries, r_adrsBRV) Then
            ValidSchedCodeInfo = False
            Exit Function
        End If
    End With
    
    Exit Function
    
ValidSchedCodeInfo_Err:
    Set l_objCodeTables = Nothing
        
    ValidSchedCodeInfo = False
    
    TrsRaiseError mstrCMODULENAME, "ValidSchedCodeInfo", , , , , True
    
End Function
Private Function ValidTimeIntervals(ByRef r_adrsSchedCodeEntries As ADODB.Recordset, _
    ByRef r_adrsBRV As ADODB.Recordset) As Boolean
    
    On Error GoTo ValidTimeIntervals_Err
    
    Dim l_dtCurrentDate As Date
    Dim l_sDayOfWeek As String
    Dim l_sStartDay As String, l_sChargeDay As String
    Dim l_iDayCounter As Integer
    Dim l_dtShiftStartDateTime As Date, l_dtShiftEndDateTime As Date
    Dim l_dtOTStartDateTime As Date, l_dtOTEndDateTime As Date
    Dim l_sStartTime As String, l_sEndTime As String
    Dim l_sOTStartTime As String, l_sOTEndTime As String
    Dim l_lStartMinutes As Long, l_lEndMinutes As Long
    Dim l_dtStartDate As Date
    Dim l_adrsShiftTimes As ADODB.Recordset
    Dim l_adrsShiftTimesClone As ADODB.Recordset
    Dim l_dHourCounter As Double
    Dim l_dtCalculatedEndDateTime As Date
    Dim l_lOTStartMinutes As Long, l_lOTEndMinutes As Long
    Dim l_lOTMinutes As Long, l_dGapMinutes As Double
    Dim l_dtTempDate As Date
    Dim l_bOTCrossMidnight As Boolean, l_bShiftCrossMidnight As Boolean
    Dim l_sFilterString As String
    
    Set l_adrsShiftTimes = New ADODB.Recordset
    
    ValidTimeIntervals = True
    
    With l_adrsShiftTimes
        'KEY INFORMATION
        .Fields.Append "CHARGE_DAY", adChar, 3, ADOFIELDATTR
        .Fields.Append "START_DATE_TIME", adChar, 16, ADOFIELDATTR
        .Fields.Append "END_DATE_TIME", adChar, 16, ADOFIELDATTR
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .Open
    End With

    l_dtCurrentDate = Format(Now(), "mm/dd/yyyy")

    With r_adrsSchedCodeEntries
        For l_iDayCounter = 1 To 7
            l_sDayOfWeek = DayOfWeek(l_iDayCounter)
            l_sStartDay = ""
            l_sChargeDay = ""
            l_sStartTime = ""
            l_sEndTime = ""
            l_sOTStartTime = ""
            l_sOTEndTime = ""
            Select Case l_sDayOfWeek
                Case "SUN"
                    If Trim(!SUN_ACT_DAY) <> "" Then
                        l_sStartDay = !SUN_ACT_DAY
                        l_sChargeDay = "SUN"
                        l_sStartTime = !SUN_START_TIME
                        l_sEndTime = !SUN_END_TIME
                        l_sOTStartTime = !SUN_OT_START_TIME
                        l_sOTEndTime = !SUN_OT_END_TIME
                    End If
                Case "MON"
                    If Trim(!MON_ACT_DAY) <> "" Then
                        l_sStartDay = !MON_ACT_DAY
                        l_sChargeDay = "MON"
                        l_sStartTime = !MON_START_TIME
                        l_sEndTime = !MON_END_TIME
                        l_sOTStartTime = !MON_OT_START_TIME
                        l_sOTEndTime = !MON_OT_END_TIME
                    End If
                Case "TUE"
                    If Trim(!TUE_ACT_DAY) <> "" Then
                        l_sStartDay = !TUE_ACT_DAY
                        l_sChargeDay = "TUE"
                        l_sStartTime = !TUE_START_TIME
                        l_sEndTime = !TUE_END_TIME
                        l_sOTStartTime = !TUE_OT_START_TIME
                        l_sOTEndTime = !TUE_OT_END_TIME
                    End If
                Case "WED"
                    If Trim(!WED_ACT_DAY) <> "" Then
                        l_sStartDay = !WED_ACT_DAY
                        l_sChargeDay = "WED"
                        l_sStartTime = !WED_START_TIME
                        l_sEndTime = !WED_END_TIME
                        l_sOTStartTime = !WED_OT_START_TIME
                        l_sOTEndTime = !WED_OT_END_TIME
                    End If
                Case "THU"
                    If Trim(!THU_ACT_DAY) <> "" Then
                        l_sStartDay = !THU_ACT_DAY
                        l_sChargeDay = "THU"
                        l_sStartTime = !THU_START_TIME
                        l_sEndTime = !THU_END_TIME
                        l_sOTStartTime = !THU_OT_START_TIME
                        l_sOTEndTime = !THU_OT_END_TIME
                    End If
                Case "FRI"
                    If Trim(!FRI_ACT_DAY) <> "" Then
                        l_sStartDay = !FRI_ACT_DAY
                        l_sChargeDay = "FRI"
                        l_sStartTime = !FRI_START_TIME
                        l_sEndTime = !FRI_END_TIME
                        l_sOTStartTime = !FRI_OT_START_TIME
                        l_sOTEndTime = !FRI_OT_END_TIME
                    End If
                Case "SAT"
                    If Trim(!SAT_ACT_DAY) <> "" Then
                        l_sStartDay = !SAT_ACT_DAY
                        l_sChargeDay = "SAT"
                        l_sStartTime = !SAT_START_TIME
                        l_sEndTime = !SAT_END_TIME
                        l_sOTStartTime = !SAT_OT_START_TIME
                        l_sOTEndTime = !SAT_OT_END_TIME
                    End If
            End Select
            If Trim(l_sStartDay) <> "" Then
                If Trim(l_sStartTime) <> "" Then
                    l_lStartMinutes = Hour(l_sStartTime) * 60 + Minute(l_sStartTime)
                    l_lEndMinutes = Hour(l_sEndTime) * 60 + Minute(l_sEndTime)
                     'The current day will be used to calculated shift start and end times for all of the week's shifts.
                    '(i.e. current day + counter,  where the counter is 1 thru 7)
                    If l_sStartDay = l_sChargeDay Then
                        l_dtShiftStartDateTime = CDate(DateAdd("d", l_iDayCounter, l_dtCurrentDate))
                        l_dtShiftStartDateTime = CDate(DateAdd("n", l_lStartMinutes, l_dtShiftStartDateTime))
                    Else
                        'If the start day is one day prior to the charge day, subtract 1 day from the start date.
                        l_dtShiftStartDateTime = CDate(DateAdd("d", l_iDayCounter, l_dtCurrentDate))
                        l_dtShiftStartDateTime = CDate(DateAdd("d", -1, l_dtShiftStartDateTime))
                        l_dtShiftStartDateTime = CDate(DateAdd("n", l_lStartMinutes, l_dtShiftStartDateTime))
                    End If
                    l_dtStartDate = Format(l_dtShiftStartDateTime, "mm/dd/yyyy")
                    'Determine if the shift crosses midnight (end minutes < start minutes)
                    If l_lEndMinutes < l_lStartMinutes Then
                        'If the shift crosses midnight, add one to the calculated date (current date + day counter)
                        l_dtShiftEndDateTime = CDate(DateAdd("d", 1, l_dtStartDate))
                        l_dtShiftEndDateTime = CDate(DateAdd("n", l_lEndMinutes, l_dtShiftEndDateTime))
                    Else
                        l_dtShiftEndDateTime = CDate(DateAdd("n", l_lEndMinutes, l_dtStartDate))
                    End If
                    'Store the shift in an ADO recordset
                    With l_adrsShiftTimes
                        .AddNew
                        !CHARGE_DAY = l_sChargeDay
                        !START_DATE_TIME = Format(l_dtShiftStartDateTime, "mm/dd/yyyy hh:mm")
                        !END_DATE_TIME = Format(l_dtShiftEndDateTime, "mm/dd/yyyy hh:mm")
                        .Update
                    End With
                End If
                If Trim(l_sOTStartTime) <> "" Then
                    l_lOTStartMinutes = Hour(l_sOTStartTime) * 60 + Minute(l_sOTStartTime)
                    l_lOTEndMinutes = Hour(l_sOTEndTime) * 60 + Minute(l_sOTEndTime)
                    If Trim(l_sStartTime) <> "" Then
                        l_lStartMinutes = Hour(l_sStartTime) * 60 + Minute(l_sStartTime)
                    Else
                        l_lStartMinutes = 0
                    End If
                    If Trim(l_sEndTime) <> "" Then
                        l_lEndMinutes = Hour(l_sEndTime) * 60 + Minute(l_sEndTime)
                    Else
                        l_lEndMinutes = 0
                    End If
                    
                    If l_lOTStartMinutes < l_lOTEndMinutes Then
                        l_bOTCrossMidnight = False
                    Else
                        l_bOTCrossMidnight = True
                    End If
                    If l_lStartMinutes < l_lEndMinutes Then
                        l_bShiftCrossMidnight = False
                    Else
                        l_bShiftCrossMidnight = True
                    End If
                    If Not l_bShiftCrossMidnight Or _
                        Not l_bOTCrossMidnight Then
                        If l_lOTStartMinutes > 0 And l_lOTEndMinutes > 0 And _
                            l_lStartMinutes > 0 And l_lEndMinutes > 0 Then
                            If l_lOTStartMinutes < l_lOTEndMinutes Then
                                l_lOTMinutes = l_lOTEndMinutes - l_lOTStartMinutes
                            Else
                                l_lOTMinutes = 1440 - l_lOTStartMinutes + l_lOTEndMinutes
                            End If
                            If l_lOTEndMinutes = 0 And l_lStartMinutes = 0 Then
                                l_dtOTEndDateTime = CDate(DateAdd("n", 0, Format(l_dtShiftStartDateTime, "mm/dd/yyyy")))
                                l_dtOTStartDateTime = CDate(DateAdd("n", -1 * l_lOTMinutes, l_dtOTEndDateTime))
                            ElseIf l_lEndMinutes = 0 And l_lOTStartMinutes = 0 Then
                                l_dtTempDate = CDate(Format(DateAdd("d", 1, l_dtShiftEndDateTime), "mm/dd/yyyy"))
                                l_dtOTStartDateTime = CDate(DateAdd("n", 0, l_dtTempDate))
                                l_dtOTEndDateTime = CDate(DateAdd("n", l_lOTMinutes, l_dtOTStartDateTime))
                            ElseIf l_lOTEndMinutes = l_lStartMinutes Then
                                l_dtOTEndDateTime = CDate(l_dtShiftStartDateTime)
                                l_dtOTStartDateTime = CDate(DateAdd("n", -1 * l_lOTMinutes, l_dtOTEndDateTime))
                            Else
                                If l_lOTStartMinutes < l_lEndMinutes Then
                                    l_dtOTStartDateTime = CDate(DateAdd("d", 1, Format(l_dtShiftEndDateTime, "mm/dd/yyyy")))
                                Else
                                    l_dtOTStartDateTime = Format(l_dtShiftEndDateTime, "mm/dd/yyyy")
                                End If
                                    l_dtOTStartDateTime = CDate(DateAdd("n", l_lOTStartMinutes, l_dtOTStartDateTime))
                                    l_dtOTEndDateTime = DateAdd("n", l_lOTMinutes, l_dtOTStartDateTime)
                            End If
                            'Store the OT hours in an ADO recordset
                            With l_adrsShiftTimes
                                .AddNew
                                !CHARGE_DAY = l_sChargeDay
                                !START_DATE_TIME = Format(l_dtOTStartDateTime, "mm/dd/yyyy hh:mm")
                                !END_DATE_TIME = Format(l_dtOTEndDateTime, "mm/dd/yyyy hh:mm")
                                .Update
                            End With
                        End If
                    Else
                        AppendSchedCodeEntryError r_adrsBRV, r_adrsSchedCodeEntries, _
                            trErrbusValidationError, mstrCBUSINESSERR, _
                            "Overtime hours must be outside of schedule hours.", _
                            l_sDayOfWeek & "_OT_START_TIME"
                        ValidTimeIntervals = False
                        Exit Function
                    End If
                End If
                If Trim(l_sOTStartTime) <> "" And Trim(l_sStartTime) <> "" Then
                    If (CDate(l_dtOTStartDateTime) < CDate(l_dtShiftStartDateTime) And _
                         CDate(l_dtOTEndDateTime) <= CDate(l_dtShiftStartDateTime)) Or _
                        (CDate(l_dtOTStartDateTime) >= CDate(l_dtShiftEndDateTime) And _
                         CDate(l_dtOTEndDateTime) > CDate(l_dtShiftEndDateTime)) Then
                         'OT hours are outside of the shift hours
                    Else
                        AppendSchedCodeEntryError r_adrsBRV, r_adrsSchedCodeEntries, _
                            trErrbusValidationError, mstrCBUSINESSERR, _
                            "Overtime hours must be outside of schedule hours.", _
                             l_sDayOfWeek & "_OT_START_TIME"
                        ValidTimeIntervals = False
                        Exit Function
                    End If
                End If
            End If
        Next l_iDayCounter
        Set l_adrsShiftTimesClone = l_adrsShiftTimes.Clone
        With l_adrsShiftTimes
            .MoveFirst
            Do While Not .EOF
                l_adrsShiftTimesClone.MoveFirst
                'Determine if there is an exact duplicate to the time entry
                l_sFilterString = "CHARGE_DAY = '" & !CHARGE_DAY & "' " & _
                                    " AND START_DATE_TIME = '" & !START_DATE_TIME & "'" & _
                                    " AND END_DATE_TIME = '" & !END_DATE_TIME & "' "
                l_adrsShiftTimesClone.Filter = MDAC25FilterFix(l_sFilterString)
                If l_adrsShiftTimesClone.RecordCount > 1 Then
                    ValidTimeIntervals = False
                    AppendSchedCodeEntryError r_adrsBRV, r_adrsSchedCodeEntries, _
                            trErrbusValidationError, mstrCBUSINESSERR, _
                            "Overlapping time entries are not allowed.", _
                            !CHARGE_DAY & "_ACT_DAY"
                End If
                l_adrsShiftTimesClone.Filter = adFilterNone
                
                Do While Not l_adrsShiftTimesClone.EOF
                    If Trim(l_adrsShiftTimesClone!CHARGE_DAY) <> Trim(!CHARGE_DAY) Or _
                        CDate(l_adrsShiftTimesClone!START_DATE_TIME) <> CDate(!START_DATE_TIME) Or _
                        CDate(l_adrsShiftTimesClone!END_DATE_TIME) <> CDate(!END_DATE_TIME) Then
                        If (CDate(l_adrsShiftTimesClone!START_DATE_TIME) < CDate(!START_DATE_TIME) And _
                           CDate(l_adrsShiftTimesClone!END_DATE_TIME) > CDate(!START_DATE_TIME)) Or _
                          (CDate(l_adrsShiftTimesClone!START_DATE_TIME) < CDate(!END_DATE_TIME) And _
                           CDate(l_adrsShiftTimesClone!END_DATE_TIME) > CDate(!END_DATE_TIME)) Or _
                          (CDate(l_adrsShiftTimesClone!START_DATE_TIME) >= CDate(!START_DATE_TIME) And _
                           CDate(l_adrsShiftTimesClone!END_DATE_TIME) <= CDate(!END_DATE_TIME)) Then
                            AppendSchedCodeEntryError r_adrsBRV, r_adrsSchedCodeEntries, _
                                trErrbusValidationError, mstrCBUSINESSERR, _
                                 "Hours cannot overlap.  Hours entered for " & !CHARGE_DAY & " overlap other time entries.", _
                                Trim(!CHARGE_DAY) & "_ACT_DAY"
                            ValidTimeIntervals = False
                            Exit Do
                        End If
                    End If
                    l_adrsShiftTimesClone.MoveNext
                Loop
                'Check for 8 hour rest period within 24 hours
                'Determine the 24 hours.  START TIME + 24 hours.
                l_dtCalculatedEndDateTime = CDate(DateAdd("d", 1, !START_DATE_TIME))
                l_adrsShiftTimesClone.MoveFirst
                l_dHourCounter = DateDiff("h", !START_DATE_TIME, !END_DATE_TIME)
                Do While Not l_adrsShiftTimesClone.EOF
                    If Trim(l_adrsShiftTimesClone!CHARGE_DAY) <> Trim(!CHARGE_DAY) Or _
                        CDate(l_adrsShiftTimesClone!START_DATE_TIME) <> CDate(!START_DATE_TIME) Or _
                        CDate(l_adrsShiftTimesClone!END_DATE_TIME) <> CDate(!END_DATE_TIME) Then
                        'Sum of the hours of the scheduled shifts which intersect with the calculated 24 hours time period
                        'If the start date time crosses the beginning of the 24 hour period, count the hours within the calculated 24 hour period.
                        If CDate(l_adrsShiftTimesClone!START_DATE_TIME) <= CDate(!START_DATE_TIME) And _
                            CDate(l_adrsShiftTimesClone!END_DATE_TIME) > CDate(!START_DATE_TIME) Then
                            If CDate(l_adrsShiftTimesClone!END_DATE_TIME) <= CDate(l_dtCalculatedEndDateTime) Then
                                l_dHourCounter = l_dHourCounter + _
                                    DateDiff("h", !START_DATE_TIME, l_adrsShiftTimesClone!END_DATE_TIME)
                            Else
                                l_dHourCounter = l_dHourCounter + _
                                    DateDiff("h", !STARTDATETIME, l_dtCalculatedEndDateTime)
                            End If
                        'If the end time crosses the end time of the calculated 24 hour period, count the hours within the calculated 24 hour period
                        ElseIf CDate(l_adrsShiftTimesClone!START_DATE_TIME) < CDate(l_dtCalculatedEndDateTime) And _
                            CDate(l_adrsShiftTimesClone!END_DATE_TIME) >= CDate(l_dtCalculatedEndDateTime) Then
                            If CDate(l_adrsShiftTimesClone!START_DATE_TIME) < CDate(!START_DATE_TIME) Then
                                l_dHourCounter = l_dHourCounter + _
                                    DateDiff("h", !START_DATE_TIME, l_dtCalculatedEndDateTime)
                            Else
                                l_dHourCounter = l_dHourCounter + _
                                    DateDiff("h", l_adrsShiftTimesClone!START_DATE_TIME, l_dtCalculatedEndDateTime)
                            End If
                        'If the time is inclusive of the calculated 24 hour period, count the hours.
                        ElseIf CDate(l_adrsShiftTimesClone!START_DATE_TIME) >= CDate(!START_DATE_TIME) And _
                            CDate(l_adrsShiftTimesClone!END_DATE_TIME) <= CDate(l_dtCalculatedEndDateTime) Then
                            l_dHourCounter = l_dHourCounter + _
                                    DateDiff("h", l_adrsShiftTimesClone!START_DATE_TIME, l_adrsShiftTimesClone!END_DATE_TIME)
                        End If
                    End If
                    'If the Sum of the scheduled shifts within the 24 hour time period exceed 16, the 8 hour rest requirement
                    'is not being met and an error is identified
                    If l_dHourCounter > 16 Then
                        ValidTimeIntervals = False
                        AppendSchedCodeEntryError r_adrsBRV, r_adrsSchedCodeEntries, _
                                trErrbusValidationError, mstrCBUSINESSERR, _
                                "Schedule hours on " & !CHARGE_DAY & " do not allow for the required 8 hour rest period within a 24 hour time period.", _
                                !CHARGE_DAY & "_ACT_DAY"
                        Exit Do
                    End If
                    l_adrsShiftTimesClone.MoveNext
                Loop
                .MoveNext
            Loop
        End With
    End With
    
    Set l_adrsShiftTimesClone = Nothing
    Set l_adrsShiftTimes = Nothing
    
    Exit Function
    
ValidTimeIntervals_Err:
    Set l_adrsShiftTimesClone = Nothing
    Set l_adrsShiftTimes = Nothing
    
    ValidTimeIntervals = False
    
    TrsRaiseError mstrCMODULENAME, "ValidTimeIntervals", , , , , True
 
End Function
Private Function DayOfWeek(ByVal v_lDayOfWeek As Long) As String

    'Determine the numerical day of the week (1-7) for an abbreviated day name(SUN-SAT)
    DayOfWeek = ""
    Select Case v_lDayOfWeek
        Case 1
            DayOfWeek = "SUN"
        Case 2
            DayOfWeek = "MON"
        Case 3
            DayOfWeek = "TUE"
        Case 4
            DayOfWeek = "WED"
        Case 5
            DayOfWeek = "THU"
        Case 6
            DayOfWeek = "FRI"
        Case 7
            DayOfWeek = "SAT"
    End Select
    
End Function
Private Function ValidShiftCode(ByVal v_sShiftCode As String) As Boolean

    ValidShiftCode = False
    
    If Trim(v_sShiftCode) = "A" Or _
        Trim(v_sShiftCode) = "P" Or _
        Trim(v_sShiftCode) = "O" Or _
        Trim(v_sShiftCode) = "N" Or _
        Trim(v_sShiftCode) = "D" Or _
        Trim(v_sShiftCode) = "T" Or _
        Trim(v_sShiftCode) = "C" Or _
        Trim(v_sShiftCode) = "" Then
        ValidShiftCode = True
    End If

End Function



Private Function ValidTimeEntries(ByRef r_adrsSchedCodeEntries As ADODB.Recordset, _
        ByRef r_adrsBRV As ADODB.Recordset, _
        ByVal v_sActualDay As String, ByVal v_sStartTime As String, _
        ByVal v_sEndTime As String, ByVal v_sLunchStartTime, _
        ByVal v_sLunchEndTime, ByVal v_sOTStartTime As String, ByVal v_sOTEndTime As String) As Boolean
        
    On Error GoTo ValidTimeEntries_Err
    
    Dim l_dLunchMinutes As Double, l_dShiftMinutes As Double
    Dim l_dLunchStartMinutes As Double, l_dLunchEndMinutes As Double
    Dim l_dShiftStartMinutes As Double, l_dShiftEndMinutes As Double
    Dim l_dAdjustMinutes As Double
    
    ValidTimeEntries = True
    
    'If the start time is not blank, the end time is required.
    'If the end time is not blank, the start time is required
    If (Trim(v_sStartTime) <> "" And Trim(v_sEndTime) = "") Or _
        (Trim(v_sEndTime) <> "" And Trim(v_sStartTime) = "") Then
        AppendSchedCodeEntryError r_adrsBRV, r_adrsSchedCodeEntries, _
                trErrbusValidationError, mstrCBUSINESSERR, _
                "Shift Start and End Times are required", _
                v_sActualDay & "_START_TIME"
        ValidTimeEntries = False
        Exit Function
    End If
    'Shift start and end times cannot be equal
    If Trim(v_sStartTime) <> "" Then
        If v_sStartTime = v_sEndTime Then
            AppendSchedCodeEntryError r_adrsBRV, r_adrsSchedCodeEntries, _
                trErrbusValidationError, mstrCBUSINESSERR, _
                "Shift Start and End times cannot be equal", _
                v_sActualDay & "_START_TIME"
            ValidTimeEntries = False
            Exit Function
        End If
    End If
    'If the overtime start time is not blank, the overtime end time is required.
    If Trim(v_sOTStartTime) <> "" And Trim(v_sOTEndTime) = "" Then
        AppendSchedCodeEntryError r_adrsBRV, r_adrsSchedCodeEntries, _
                trErrbusValidationError, mstrCBUSINESSERR, _
                "Overtime end time is required if the overtime start time is entered.", _
                v_sActualDay & "_OT_START_TIME"
        ValidTimeEntries = False
        Exit Function
    End If
    
    'If the overtime end time is not blank, the overtime start time is required
    If Trim(v_sOTEndTime) <> "" And Trim(v_sOTStartTime) = "" Then
        AppendSchedCodeEntryError r_adrsBRV, r_adrsSchedCodeEntries, _
                trErrbusValidationError, mstrCBUSINESSERR, _
                "Overtime start time is required if the overtime end time is entered.", _
                v_sActualDay & "_OT_START_TIME"
        ValidTimeEntries = False
        Exit Function
    End If

    'OT start and end times cannot be equal
    If Trim(v_sOTStartTime) <> "" Then
        If Trim(v_sOTStartTime) = Trim(v_sOTEndTime) Then
            AppendSchedCodeEntryError r_adrsBRV, r_adrsSchedCodeEntries, _
                trErrbusValidationError, mstrCBUSINESSERR, _
                "Overtime start and end times cannot be equal", _
                v_sActualDay & "_OT_START_TIME"
            ValidTimeEntries = False
            Exit Function
        End If
    End If
        
    'Validate the start time format and value
    If Trim(v_sStartTime) <> "" Then
        If ValidTimeFormat(v_sStartTime) Then
            l_dShiftStartMinutes = Hour(v_sStartTime) * 60 + Minute(v_sStartTime)
        Else
            AppendSchedCodeEntryError r_adrsBRV, r_adrsSchedCodeEntries, _
                trErrbusValidationError, mstrCBUSINESSERR, _
                "Invalid start time value", _
                v_sActualDay & "_START_TIME"
            ValidTimeEntries = False
            Exit Function
        End If
    End If
    
    'Validate the end time format and value
    If Trim(v_sEndTime) <> "" Then
        If ValidTimeFormat(v_sEndTime) Then
            l_dShiftEndMinutes = Hour(v_sEndTime) * 60 + Minute(v_sEndTime)
        Else
            AppendSchedCodeEntryError r_adrsBRV, r_adrsSchedCodeEntries, _
                trErrbusValidationError, mstrCBUSINESSERR, _
                "Invalid end time value", _
                v_sActualDay & "_END_TIME"
            ValidTimeEntries = False
            Exit Function
        End If
    End If
    'Validate the overtime start time format and value
    If Trim(v_sOTStartTime) <> "" Then
        If Not ValidTimeFormat(v_sOTStartTime) Then
            AppendSchedCodeEntryError r_adrsBRV, r_adrsSchedCodeEntries, _
                trErrbusValidationError, mstrCBUSINESSERR, _
                "Invalid overtime start value", _
                v_sActualDay & "_OT_START_TIME"
            ValidTimeEntries = False
            Exit Function
        End If
    End If
    'Validate the overtime end time format and value
    If Trim(v_sOTEndTime) <> "" Then
        If Not ValidTimeFormat(v_sOTEndTime) Then
            AppendSchedCodeEntryError r_adrsBRV, r_adrsSchedCodeEntries, _
                trErrbusValidationError, mstrCBUSINESSERR, _
                "Invalid overtime end value", _
                v_sActualDay & "_OT_END_TIME"
            ValidTimeEntries = False
            Exit Function
        End If
    End If
    'Validate the lunch start time format and value
    If Trim(v_sLunchStartTime) <> "" Then
        If Not ValidTimeFormat(v_sLunchStartTime) Then
            AppendSchedCodeEntryError r_adrsBRV, r_adrsSchedCodeEntries, _
                trErrbusValidationError, mstrCBUSINESSERR, _
                "Invalid lunch start time value", _
                v_sActualDay & "_LUNCH_START"
            ValidTimeEntries = False
            Exit Function
        End If
    End If
    'Validate the lunch end time format and value
    If Trim(v_sLunchEndTime) <> "" Then
        If Not ValidTimeFormat(v_sLunchEndTime) Then
            AppendSchedCodeEntryError r_adrsBRV, r_adrsSchedCodeEntries, _
                trErrbusValidationError, mstrCBUSINESSERR, _
                "Invalid lunch end time value", _
                v_sActualDay & "_LUNCH_END"
            ValidTimeEntries = False
            Exit Function
        End If
    End If
    'Validate that lunch is entered in six minute increments
    If Trim(v_sLunchStartTime) <> "" And Trim(v_sLunchEndTime) <> "" Then
        l_dLunchStartMinutes = Hour(v_sLunchStartTime) * 60 + Minute(v_sLunchStartTime)
        l_dLunchEndMinutes = Hour(v_sLunchEndTime) * 60 + Minute(v_sLunchEndTime)
        If l_dLunchEndMinutes < l_dLunchStartMinutes Then
            l_dLunchMinutes = 1440 - l_dLunchStartMinutes + l_dLunchEndMinutes
        Else
            l_dLunchMinutes = l_dLunchEndMinutes - l_dLunchStartMinutes
        End If
        If Round(l_dLunchMinutes / 6, 0) <> (l_dLunchMinutes / 6) Then
            AppendSchedCodeEntryError r_adrsBRV, r_adrsSchedCodeEntries, _
                trErrbusValidationError, mstrCBUSINESSERR, _
                "Lunch must be in six minute increments", _
                v_sActualDay & "_LUNCH_START"
            ValidTimeEntries = False
            Exit Function
        End If
        'Validate that lunch falls within the shift hours.
        If l_dShiftEndMinutes < l_dShiftStartMinutes Then
            'If the shift crosses midnight, adjust the time entries as if the shift started
            'at 0 minutes, then compare the sequence of the time entries.
            l_dAdjustMinutes = 1440 - l_dShiftStartMinutes
            If l_dLunchStartMinutes > l_dShiftStartMinutes Then
                l_dLunchStartMinutes = l_dLunchStartMinutes - l_dShiftStartMinutes
            Else
                l_dLunchStartMinutes = l_dLunchStartMinutes + l_dAdjustMinutes
            End If
            If l_dLunchEndMinutes > l_dShiftStartMinutes Then
                l_dLunchEndMinutes = l_dLunchEndMinutes - l_dShiftStartMinutes
            Else
                l_dLunchEndMinutes = l_dLunchEndMinutes + l_dAdjustMinutes
            End If
            l_dShiftEndMinutes = l_dShiftEndMinutes + l_dAdjustMinutes
            l_dShiftStartMinutes = 0
        End If
        
        If l_dLunchStartMinutes < l_dShiftStartMinutes Or _
            l_dLunchStartMinutes > l_dShiftEndMinutes Or _
            l_dLunchEndMinutes < l_dShiftStartMinutes Or _
            l_dLunchEndMinutes > l_dShiftEndMinutes Or _
            l_dLunchStartMinutes > l_dLunchEndMinutes Then
            AppendSchedCodeEntryError r_adrsBRV, r_adrsSchedCodeEntries, _
                trErrbusValidationError, mstrCBUSINESSERR, _
                "Lunch must be within the time period defined by the shift start and end times", _
                v_sActualDay & "_LUNCH_START"
            ValidTimeEntries = False
            Exit Function
        End If
        'Lunch hours must be < the shift hours
        l_dShiftMinutes = (l_dShiftEndMinutes - l_dShiftStartMinutes) - l_dLunchMinutes
        If l_dLunchMinutes >= l_dShiftMinutes Then
            AppendSchedCodeEntryError r_adrsBRV, r_adrsSchedCodeEntries, _
                trErrbusValidationError, mstrCBUSINESSERR, _
                "Lunch time must be less than the shift time", _
                v_sActualDay & "_LUNCH_START"
            ValidTimeEntries = False
            Exit Function
        End If
    End If
    Exit Function
    
ValidTimeEntries_Err:
    ValidTimeEntries = False
    TrsRaiseError mstrCMODULENAME, "ValidTimeEntries", , , , , True
    
End Function
    
Private Function ValidTimeFormat(ByVal v_sTimeEntry As String) As Boolean

    Dim l_sHours As String, l_sMinutes As String
    Dim l_sHoursMinutes As String
    Dim l_iCounter As Integer

    ValidTimeFormat = True
    'Time entry must be 5 characters
    If Len(Trim(v_sTimeEntry)) <> 5 Then
        ValidTimeFormat = False
    End If
    'Time entry must have a ":" in the third position
    If Not Mid(v_sTimeEntry, 3, 1) = ":" Then
        ValidTimeFormat = False
        Exit Function
    End If
    
    l_sHours = Left(v_sTimeEntry, 2)
    l_sMinutes = Right(v_sTimeEntry, 2)
    l_sHoursMinutes = l_sHours & l_sMinutes
    
    'Hours and minutes must be numeric
    For l_iCounter = 1 To Len(l_sHoursMinutes)
        If Not IsNumeric(Mid(l_sHoursMinutes, l_iCounter, 1)) Then
            ValidTimeFormat = False
            Exit Function
        End If
    Next
    
   'Hours must be between 00 and 23
    If CInt(l_sHours) < 0 Or CInt(l_sHours) > 24 Then
        ValidTimeFormat = False
        Exit Function
    End If
    
    'Minutes be between 00 and 59
    If CInt(l_sMinutes) < 0 Or CInt(l_sHours) > 59 Then
        ValidTimeFormat = False
        Exit Function
    End If
    
    '24:00 is not a valid time entry
    If CInt(l_sHours) = 24 And CInt(l_sMinutes) <> "00" Then
        ValidTimeFormat = False
    End If
    
End Function


Private Sub AppendSchedCodeEntryError( _
            ByRef r_adrsBRV As ADODB.Recordset, _
            ByRef r_adrsSchedChgEntries As ADODB.Recordset, _
            ByVal v_lErrorNumber As Long, _
            ByVal v_sErrorType As String, _
            ByVal v_sErrorDescription As String, _
            Optional ByVal v_sErrorField As String, _
            Optional ByVal v_sErrorSource As String)

    Dim l_vEmpNum As Long
    Dim l_vChargeDate As Date
    Dim l_vDateWorked As Date
    Dim l_vErrorSource As Variant
    Dim l_vARFields(10) As Variant
    Dim l_vErrorField As String
    Dim l_vScheduleCode As Variant
    
    On Error GoTo AppendSchedCodeEntryError_Err
    
    l_vARFields(0) = "SCHED_CD"
    l_vARFields(1) = "SCHED_DESC"
    l_vARFields(2) = "TIME_STAMP"
    l_vARFields(3) = "RESET_FLAG"
    
    
    'Determine literal for error source
    l_vErrorSource = IIf(v_sErrorSource = "", mstrCMODULENAME, v_sErrorSource)
    l_vErrorField = CVar(Trim(v_sErrorField))
    
    If r_adrsSchedChgEntries.BOF Or r_adrsSchedChgEntries.EOF Then
        AppendBusinessError _
            r_adrsBRV, v_lErrorNumber, v_sErrorType, v_sErrorDescription, _
            v_sErrorSource, 1, 0
        Exit Sub
    End If
    
    AppendBusinessError _
        r_adrsBRV, _
        v_lErrorNumber, _
        v_sErrorType, _
        v_sErrorDescription, _
        l_vErrorSource, _
        1, _
        r_adrsSchedChgEntries!RECORD_ID, _
        l_vErrorField, _
        l_vARFields, _
        r_adrsSchedChgEntries!SCHED_CD, _
        r_adrsSchedChgEntries!SCHED_DESC, _
        r_adrsSchedChgEntries!TIME_STAMP, _
        "N"
    Exit Sub
    
AppendSchedCodeEntryError_Err:
    TrsRaiseError mstrCMODULENAME, "AppendSchedCodeEntryError", _
        trErrbusAppendEntryError, True, , , True
        
End Sub
Private Function DefineRecordset(ByVal v_adrsInputRecordset As ADODB.Recordset) As ADODB.Recordset
    On Error GoTo DefineRecordset_Err
    
    Dim l_iCounter As Integer
    Dim l_adrsTemp As ADODB.Recordset
    
    Set l_adrsTemp = New ADODB.Recordset
                                
    With v_adrsInputRecordset
        For l_iCounter = 0 To .Fields.Count - 1
            l_adrsTemp.Fields.Append .Fields(l_iCounter).Name, .Fields(l_iCounter).Type _
                    , .Fields(l_iCounter).DefinedSize _
                    , ADOFIELDATTR
            l_adrsTemp.Fields.Item(l_iCounter).Precision = .Fields(l_iCounter).Precision
            l_adrsTemp.Fields.Item(l_iCounter).NumericScale = .Fields(l_iCounter).NumericScale
        Next l_iCounter
        l_adrsTemp.CursorLocation = adUseClient
        l_adrsTemp.CursorType = adOpenStatic
        l_adrsTemp.LockType = adLockBatchOptimistic
        l_adrsTemp.Open
    End With

    Set DefineRecordset = l_adrsTemp
    
    Set l_adrsTemp = Nothing
    
    Exit Function
    
DefineRecordset_Err:
    Set l_adrsTemp = Nothing
    GetObjectContext.SetAbort
    TrsRaiseError mstrCMODULENAME, "DefineRecordset"
    
End Function

