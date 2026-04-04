VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "cbsSchedPersonal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const mstrCMODULENAME As String = "busSchedPersonal"
'Private Const mstrCVARFIELDNAMES As String = "EMP_NUM;START_DATE;END_DATE"
Private Const m_datCOPTIONALSTARTDATE As Date = #1/1/100#
Private Const m_datCOPTIONALENDDATE As Date = #12/31/9999#

'Private m_lngRecordId As Long
'Private m_lngUserId As Long
'Private m_lngErrorNo As Long


'Get Personal schedules for multiple employees
Public Function GetPersonalSchedForMultiEmp _
                (ByVal v_aryEmployees As Variant _
                 , Optional ByVal v_datStartDate As Date = m_datCOPTIONALSTARTDATE _
                 , Optional ByVal v_datEndDate As Date = m_datCOPTIONALENDDATE _
                ) As ADODB.Recordset


On Error GoTo GetPersonalSchedForMultiEmp_EH

    Dim l_dbSchedPersonal As dbSchedPersonal.cdSchedPersonal

    Set l_dbSchedPersonal = CreateObject("dbSchedPersonal.cdSchedPersonal")

    Set GetPersonalSchedForMultiEmp = l_dbSchedPersonal.GetPersonalSchedForMultiEmp(v_aryEmployees, v_datStartDate, v_datEndDate)

    Set l_dbSchedPersonal = Nothing
Exit Function

GetPersonalSchedForMultiEmp_EH:
    Set l_dbSchedPersonal = Nothing
    TrsRaiseError mstrCMODULENAME, "GetPersonalSchedForMultiEmp"

End Function


'Get Personal schedules for an employee
Public Function GetPersonalSchedules _
                (ByVal v_lngEmployee As Long _
                 , Optional ByVal v_datStartDate As Date = m_datCOPTIONALSTARTDATE _
                 , Optional ByVal v_datEndDate As Date = m_datCOPTIONALENDDATE _
                ) As ADODB.Recordset


On Error GoTo GetPersonalSchedules_EH

    Dim l_dbSchedPersonal As dbSchedPersonal.cdSchedPersonal

    Set l_dbSchedPersonal = CreateObject("dbSchedPersonal.cdSchedPersonal")

    Set GetPersonalSchedules = l_dbSchedPersonal.GetPersonalSchedules(v_lngEmployee, v_datStartDate, v_datEndDate)

    Set l_dbSchedPersonal = Nothing
Exit Function

GetPersonalSchedules_EH:
    Set l_dbSchedPersonal = Nothing
    TrsRaiseError mstrCMODULENAME, "GetPersonalSchedules"

End Function

