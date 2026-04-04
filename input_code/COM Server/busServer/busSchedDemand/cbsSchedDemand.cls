VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "cbsSchedDemand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const m_strCMODULENAME As String = "cbsSchedDemand"
Private Const m_datCOPTIONALSTARTDATE As Date = #1/1/100#
Private Const m_datCOPTIONALENDDATE As Date = #12/31/9999#

'//-----Blue251
Public Function GetAssignments(ByVal v_lngWgNum As Long, ByVal v_datStartDate As Date, Optional ByVal v_datEndDate As Date = m_datCOPTIONALENDDATE, _
                                Optional ByVal v_lngStartLine As Long = -999, Optional v_lngEndLine As Long = -999) As ADODB.Recordset
    Dim l_objSchDmd As dbSchedDemand.cdSchedDemand
    Dim l_objRS As Recordset

    On Error GoTo eh

    Set l_objSchDmd = CreateObject("dbSchedDemand.cdSchedDemand")

    Set l_objRS = l_objSchDmd.GetAssignments(v_lngWgNum, v_datStartDate, v_datEndDate, v_lngStartLine, v_lngEndLine)
    Set GetAssignments = l_objRS
    Set l_objRS = Nothing
    Set l_objSchDmd = Nothing
    Exit Function

eh:
    Set l_objSchDmd = Nothing
    TrsRaiseError m_strCMODULENAME, "GetAssignments", , , , , True
End Function

Public Function GetEmployeesInWorkGroup(ByVal v_lngWgNum As Long, ByVal v_datStartDate As Date, ByVal v_datEndDate As Date) As ADODB.Recordset
    Dim l_objSchDmd As dbSchedDemand.cdSchedDemand
    Dim l_objRS As Recordset

    On Error GoTo eh

    Set l_objSchDmd = CreateObject("dbSchedDemand.cdSchedDemand")

    Set l_objRS = l_objSchDmd.GetEmployeesInWorkGroup(v_lngWgNum, v_datStartDate, v_datEndDate)
    Set GetEmployeesInWorkGroup = l_objRS
    Set l_objRS = Nothing
    Set l_objSchDmd = Nothing
    Exit Function
eh:
    Set l_objSchDmd = Nothing
    TrsRaiseError m_strCMODULENAME, "GetEmployeesInWorkGroup", , , , , True
End Function
'BEGIN PREM TRIS CONVERSION PHASE 3
Public Function GetDemandJobPositions(ByVal vlWorkgroup As Long _
        ) As ADODB.Recordset
On Error GoTo GetDemandJobPositions_EH
Dim lobjdb As dbSchedDemand.cdSchedDemand

    Set lobjdb = CreateObject( _
                    "dbSchedDemand.cdSchedDemand")
    Set GetDemandJobPositions = lobjdb.GetDemandJobPositions(vlWorkgroup)
    Set lobjdb = Nothing

    Exit Function
GetDemandJobPositions_EH:
    Set lobjdb = Nothing

    TrsRaiseError m_strCMODULENAME _
        , "GetDemandJobPositions" _
        , , , , _
        , True

End Function
Public Function GetDemandJobPositionsByDate(ByVal vlWorkgroup As Long _
                , ByVal vdtStartDate As Date _
                , ByVal vdtEndDate As Date _
                ) As ADODB.Recordset
On Error GoTo GetDemandJobPositionsByDate_EH
Dim lobjdb As dbSchedDemand.cdSchedDemand

    Set lobjdb = CreateObject( _
                    "dbSchedDemand.cdSchedDemand")
    Set GetDemandJobPositionsByDate = lobjdb.GetDemandJobPositions( _
                                        vlWorkgroup _
                                        , vdtStartDate _
                                        , vdtEndDate)
    Set lobjdb = Nothing

    Exit Function
GetDemandJobPositionsByDate_EH:
    Set lobjdb = Nothing

    TrsRaiseError m_strCMODULENAME _
        , "GetDemandJobPositionsByDate" _
        , , , , _
        , True

End Function
'END PREM TRIS CONVERSION PHASE 3
