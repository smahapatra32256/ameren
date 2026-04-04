Attribute VB_Name = "mbusEnvironment"
Option Explicit

Private Const MODULENAME As String = "mbusEnvironment"
Private Const m_strCMODULENAME As String = MODULENAME

Public Function PreferenceTypeCodeToString(ByVal v_eTypeCd As PreferenceTypeCode) As String
    Select Case v_eTypeCd
    Case TRISOptions
        PreferenceTypeCodeToString = "OPT"
    Case JobList
        PreferenceTypeCodeToString = "JOB"
    Case ScheduleList
        PreferenceTypeCodeToString = "SCH"
    Case RotationList
        PreferenceTypeCodeToString = "ROT"
    Case ExceptionList
        PreferenceTypeCodeToString = "EXC"
    Case OTReasonList
        PreferenceTypeCodeToString = "OTR"
    Case WorkGroupList
        PreferenceTypeCodeToString = "WG"
    Case TRISPersistence
        PreferenceTypeCodeToString = "PST"
    End Select
End Function

Public Function StringToPreferenceTypeCode(ByVal v_strTypeCd As String) As PreferenceTypeCode
    Select Case v_strTypeCd
    Case "OPT"
        StringToPreferenceTypeCode = TRISOptions
    Case "JOB"
        StringToPreferenceTypeCode = JobList
    Case "SCH"
        StringToPreferenceTypeCode = ScheduleList
    Case "ROT"
        StringToPreferenceTypeCode = RotationList
    Case "EXC"
        StringToPreferenceTypeCode = ExceptionList
    Case "OTR"
        StringToPreferenceTypeCode = OTReasonList
    Case "WG"
        StringToPreferenceTypeCode = WorkGroupList
    Case "PST"
        StringToPreferenceTypeCode = TRISPersistence
    End Select
End Function
