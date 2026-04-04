VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "cbsEnvironmentTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const MODULENAME As String = "cbsFunSecurityTrans"
Private Const m_strCMODULENAME As String = MODULENAME
Private Const m_lngKeyForDefault As Long = -999991
'//-----------------------------------------------------------------------------------------------------------------------//
'//-----this class has been changed to Non-Transactional, cbsEnvironmentTrans2 is now the transactional class
'//-----------------------------------------------------------------------------------------------------------------------//
Public Enum PreferenceTypeCode
    TRISOptions = 0
    JobList
    ScheduleList
    RotationList
    ExceptionList
    OTReasonList
    WorkGroupList
    Other
    TRISPersistence
End Enum

'//-----for migrating to new way of storage
Public Function DoesUserPreferenceExist(ByVal v_lngEmpNum As Long, _
                                        ByVal v_eTypeCd As PreferenceTypeCode, _
                                        ByVal v_strOtherTypeCd As String) As Boolean
                                    
    On Error GoTo eh
    
    Dim dbEnvTrans As dbEnvironment.cdEnvironmentTrans
    Dim strTypeCd As String
    
    Set dbEnvTrans = CreateObject("dbEnvironment.cdEnvironmentTrans")
    
    If v_eTypeCd <> PreferenceTypeCode.Other Then
        strTypeCd = PreferenceTypeCodeToString(v_eTypeCd)
    Else
        strTypeCd = v_strOtherTypeCd
    End If
    
    DoesUserPreferenceExist = dbEnvTrans.DoesUserPreferenceExist(v_lngEmpNum, strTypeCd)
                  
    Set dbEnvTrans = Nothing
    
    GetObjectContext.SetComplete
    Exit Function
eh:
    'Abort Transaction
    GetObjectContext.SetAbort
    
    TrsRaiseError MODULENAME _
        , "DoesUserPreferenceExist" _
        , _
        , , DescArgsFrom(CStr(v_lngEmpNum), PreferenceTypeCodeToString(v_eTypeCd)), _
        , True
End Function

Public Function GetUserPreference(ByVal v_lngEmpNum As Long, _
                                    ByVal v_eTypeCd As PreferenceTypeCode, _
                                    ByVal v_strOtherTypeCd As String) As ADODB.Recordset
                                    
    On Error GoTo eh
    
    Dim dbEnvTrans As dbEnvironment.cdEnvironmentTrans
    Dim strTypeCd As String
    
    Set dbEnvTrans = CreateObject("dbEnvironment.cdEnvironmentTrans")
    
    If v_eTypeCd <> PreferenceTypeCode.Other Then
        strTypeCd = PreferenceTypeCodeToString(v_eTypeCd)
    Else
        strTypeCd = v_strOtherTypeCd
    End If
    
    Set GetUserPreference = dbEnvTrans.GetUserPreference(v_lngEmpNum, strTypeCd)
                  
    Set dbEnvTrans = Nothing
'Begin Prem TRIS 5.9
    RemoveCaseNumberPreference GetUserPreference
'End Prem TRIS 5.9
    GetObjectContext.SetComplete
    Exit Function
eh:
    'Abort Transaction
    GetObjectContext.SetAbort
    
    TrsRaiseError MODULENAME _
        , "GetUserPreference" _
        , _
        , , DescArgsFrom(CStr(v_lngEmpNum), PreferenceTypeCodeToString(v_eTypeCd)), _
        , True
End Function

'//-----use with OPT for setting the initial or default TRIS Options for the client
'could store empty recordsets for the other types instead of creating them with code ???
Public Function GetUserPreferenceDefault(ByVal v_eTypeCd As PreferenceTypeCode, _
                                        ByVal v_strOtherTypeCd As String) As ADODB.Recordset
                                    
    On Error GoTo eh
    
    Dim dbEnvTrans As dbEnvironment.cdEnvironmentTrans
    Dim strTypeCd As String
    
    Set dbEnvTrans = CreateObject("dbEnvironment.cdEnvironmentTrans")
    
    If v_eTypeCd <> PreferenceTypeCode.Other Then
        strTypeCd = PreferenceTypeCodeToString(v_eTypeCd)
    Else
        strTypeCd = v_strOtherTypeCd
    End If
    
    Set GetUserPreferenceDefault = dbEnvTrans.GetUserPreference(m_lngKeyForDefault, strTypeCd)
                  
    Set dbEnvTrans = Nothing
'Begin Prem TRIS 5.9
    RemoveCaseNumberPreference GetUserPreferenceDefault
'End Prem TRIS 5.9
    GetObjectContext.SetComplete
    Exit Function
eh:
    'Abort Transaction
    GetObjectContext.SetAbort
    
    TrsRaiseError MODULENAME _
        , "GetUserPreferenceDefault" _
        , _
        , , DescArgsFrom(CStr(m_lngKeyForDefault), PreferenceTypeCodeToString(v_eTypeCd)), _
        , True
End Function
'Begin Prem TRIS 5.9
Private Sub RemoveCaseNumberPreference(ByRef radrs As ADODB.Recordset)
On Error GoTo RemoveCaseNumberPreference_EH
Dim llIndex As Long, liPivotColOrder As Integer
Dim lsPreviousSection As String, lbSectionExists As Boolean
Dim lbKeyExists As Boolean

    If Not (radrs Is Nothing) Then
        With radrs
            If Not (.BOF And .EOF) Then
                lbSectionExists = False

                For llIndex = 0 To .Fields.Count - 1
                    If (UCase(.Fields(llIndex).Name) = "SECTION") Then
                        lbSectionExists = True
                    End If
                    If (UCase(.Fields(llIndex).Name) = "KEY") Then
                        lbKeyExists = True
                    End If
                Next llIndex

                If (lbKeyExists And lbSectionExists) Then
                    .Filter = adFilterNone

                    .Filter = "(SECTION = 'TRIS\Schedule\Exceptions\Column Hide Show' AND KEY = 'Case Num.')" & _
                        " OR (SECTION = 'TRIS\Schedule\Exceptions\Details Column Hide Show' AND KEY = 'Case Number')" & _
                        " OR (SECTION = 'TRIS\Schedule\Exceptions\Column Width' AND KEY = 'CASE_NUM')" & _
                        " OR (SECTION = 'TRIS\Time\Approve\Column Hide Show' AND KEY = 'Case Num.')" & _
                        " OR (SECTION = 'TRIS\Time\Approve\Column Width' AND KEY = 'CASE_NUM')" & _
                        " OR (SECTION = 'TRIS\Time\Collection\Day\Time Table\Columns' AND KEY = 'Case Number')" & _
                        " OR (SECTION = 'TRIS\Time\Collection\Pay Period\Time Table\Columns' AND KEY = 'Case Number')" & _
                        " OR (SECTION = 'TRIS\View\Calendar\Column Hide Show' AND KEY = 'Case Num.')" & _
                        " OR (SECTION = 'TRIS\View\Calendar\Column Width' AND KEY = 'CASE_NUM')" & _
                        " OR (SECTION = 'TRIS\View\Calendar Multiple\Column Hide Show' AND KEY = 'Case Num.')" & _
                        " OR (SECTION = 'TRIS\View\Calendar Multiple\Column Width' AND KEY = 'CASE_NUM')" & _
                        " OR (SECTION = 'TRIS\View\Schedule\Column Hide Show' AND KEY = 'Case Num.')" & _
                        " OR (SECTION = 'TRIS\View\Schedule\Column Width' AND KEY = 'CASE_NUM')" & _
                        " OR (SECTION = 'TRIS\View\Time\Column Hide Show' AND KEY = 'Case Num.')" & _
                        " OR (SECTION = 'TRIS\View\Time\Column Width' AND KEY = 'CASE_NUM')"
                    If Not (.EOF) Then
                        .Sort = "SECTION"
                        .MoveLast

                        For llIndex = .RecordCount To 1 Step -1
                            .Delete adAffectCurrent
                            .Update

                            If Not (.BOF) Then
                                .MovePrevious

                                If (.BOF) Then
                                    Exit For
                                End If
                            Else
                                Exit For
                            End If
                        Next llIndex
                    End If

                    .Filter = adFilterNone
                    .Filter = "(SECTION = 'TRIS\Schedule\Exceptions\Column Order')" & _
                        " OR (SECTION = 'TRIS\Time\Approve\Column Order')" & _
                        " OR (SECTION = 'TRIS\View\Calendar\Column Order')" & _
                        " OR (SECTION = 'TRIS\View\Calendar Multiple\Column Order')" & _
                        " OR (SECTION = 'TRIS\View\Schedule\Column Order')" & _
                        " OR (SECTION = 'TRIS\View\Time\Column Order')"
                    If Not (.EOF) Then
                        .Sort = "SECTION, SETTING"

                        For llIndex = 1 To .RecordCount
                            If (!SECTION <> lsPreviousSection) Then
                                liPivotColOrder = -1
                                lsPreviousSection = !SECTION
                            End If

                            If (!SECTION = lsPreviousSection) Then
                                If (!Key = "CASE_NUM") Then
                                    liPivotColOrder = CInt(!SETTING)
                                    .Delete adAffectCurrent
                                    .Update
                                Else
                                    If (liPivotColOrder > -1) Then
                                        If (CInt(!SETTING) > liPivotColOrder) Then
                                            !SETTING = CStr(CInt(!SETTING) - 1)
                                            .Update
                                        End If
                                    End If
                                End If
                            End If

                            If Not (.EOF) Then
                                .MoveNext

                                If (.EOF) Then
                                    Exit For
                                End If
                            Else
                                Exit For
                            End If
                        Next llIndex
                    End If

                    .Filter = adFilterNone
                End If
            End If
        End With
    End If

    Exit Sub
RemoveCaseNumberPreference_EH:
    TrsRaiseError MODULENAME _
        , "RemoveCaseNumberPreference"

End Sub
'End Prem TRIS 5.9

Public Function SetUserPreference(ByVal v_lngEmpNum As Long, _
                                    ByVal v_eTypeCd As PreferenceTypeCode, _
                                    ByVal v_strOtherTypeCd As String, _
                                    ByVal v_rsPreference As ADODB.Recordset) As Boolean
    On Error GoTo eh
    
    Dim busEnvTrans As busEnvironment.cbsEnvironmentTrans2
        
    Set busEnvTrans = CreateObject("busEnvironment.cbsEnvironmentTrans2")

    SetUserPreference = busEnvTrans.SetUserPreference(v_lngEmpNum, v_eTypeCd, v_strOtherTypeCd, v_rsPreference)
    
    Set busEnvTrans = Nothing

'    Dim dbEnvTrans As dbEnvironment.cdEnvironmentTrans
'    Dim strTypeCd As String
'
'    Set dbEnvTrans = CreateObject("dbEnvironment.cdEnvironmentTrans")
'
'    If v_eTypeCd <> PreferenceTypeCode.Other Then
'        strTypeCd = PreferenceTypeCodeToString(v_eTypeCd)
'    Else
'        strTypeCd = v_strOtherTypeCd
'    End If
'
'    Set dbEnvTrans = CreateObject("dbEnvironment.cdEnvironmentTrans")
'
'    SetUserPreference = dbEnvTrans.SetUserPreference(v_lngEmpNum, strTypeCd, v_rsPreference)
'
'    Set dbEnvTrans = Nothing
    
'    GetObjectContext.SetComplete
    Exit Function
eh:
    'Abort Transaction
'    GetObjectContext.SetAbort
    Set busEnvTrans = Nothing
    TrsRaiseError MODULENAME _
        , "SetUserPreference" _
        , _
        , , DescArgsFrom(CStr(v_lngEmpNum), PreferenceTypeCodeToString(v_eTypeCd)), _
        , True
End Function

