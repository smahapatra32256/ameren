VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "cbsOTGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'AMEREN Project: TRIS Conversion Phase 3
'June 2004
'Author: Prem M. Chander
'Module Description: This class will provide the user the ability
'to get Overtime groups and related information

Option Explicit

Private Const MODULENAME = "busOTGroup.cbsOTGroup"
Public Sub GetOvertimeGroupInformation(ByVal vlOTGroupNumber As Long _
                , ByRef radrsOTGroupBase As ADODB.Recordset _
                , ByRef radrsOTGroupBanks As ADODB.Recordset _
                , ByRef radrsOTGroupWorkgroups As ADODB.Recordset)
On Error GoTo GetOvertimeGroupInformation_EH
Dim lobjOTGroup As dbOTGroup.cdOTGroup

    'Get OT Group Information
    Set lobjOTGroup = CreateObject( _
                        "dbOTGroup.cdOTGroup")

    Set radrsOTGroupBase = lobjOTGroup.GetOvertimeGroupBase( _
                            vlOTGroupNumber)
    Set radrsOTGroupBanks = lobjOTGroup.GetOvertimeGroupBanks( _
                            vlOTGroupNumber)
    Set radrsOTGroupWorkgroups = lobjOTGroup.GetOvertimeGroupWorkgroups( _
                            vlOTGroupNumber)

    Set lobjOTGroup = Nothing

    Exit Sub
GetOvertimeGroupInformation_EH:
    Set lobjOTGroup = Nothing

    TrsRaiseError MODULENAME _
        , "GetOvertimeGroupInformation" _
        , , , , _
        , True

End Sub
Public Function GetAllOvertimeGroups() As ADODB.Recordset
On Error GoTo GetAllOvertimeGroups_EH
Dim lobjOTGroup As dbOTGroup.cdOTGroup

    'Get OT Group Information
    Set lobjOTGroup = CreateObject( _
                        "dbOTGroup.cdOTGroup")
    Set GetAllOvertimeGroups = lobjOTGroup.GetAllOvertimeGroups()
    Set lobjOTGroup = Nothing

    Exit Function
GetAllOvertimeGroups_EH:
    Set lobjOTGroup = Nothing

    TrsRaiseError MODULENAME _
        , "GetAllOvertimeGroups" _
        , , , , _
        , True

End Function
Public Function GetEmployeesInOTGroup(ByVal vlOTGroupNumber As Long _
                    , Optional ByVal vdtDate As Date = #1/1/1900# _
                    ) As ADODB.Recordset
On Error GoTo GetEmployeesInOTGroup_EH
Dim lobjOTGroup As dbOTGroup.cdOTGroup

    If (vdtDate = #1/1/1900#) Then
        vdtDate = Date
    End If

    'Get OT Group Information
    Set lobjOTGroup = CreateObject( _
                        "dbOTGroup.cdOTGroup")
    Set GetEmployeesInOTGroup = lobjOTGroup.GetEmployeesInOTGroup( _
                                    vlOTGroupNumber _
                                    , vdtDate)
    Set lobjOTGroup = Nothing

    Exit Function
GetEmployeesInOTGroup_EH:
    Set lobjOTGroup = Nothing

    TrsRaiseError MODULENAME _
        , "GetEmployeesInOTGroup" _
        , , , , _
        , True

End Function
