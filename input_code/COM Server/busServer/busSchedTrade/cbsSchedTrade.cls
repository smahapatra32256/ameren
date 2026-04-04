VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "cbsSchedTrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const m_strCMODULENAME = "cbsSchedTrade"

Public Function GetTradesByEmployee( _
        ByVal v_lngEmployees As Variant, _
        ByVal v_datStartDate As Date, _
        ByVal v_datEndDate As Date) _
            As ADODB.Recordset
                    
'This Function call dbSchedTrade.GetTradesByEmployee
'This gets all trades that is on the selected
' employee's schedule.  In the case of a demand line it is
' the employee that is honoring the trade.
'This may be different than the employee that requested the trade.
'Also if an employee made an Emp Trade prior to being placed on
' a demand line, that trade will not be retrieved
                    

    Dim l_obj As dbSchedTrade.cdSchedTrade
    Dim l_objRS As Recordset
    
    On Error GoTo eh
    
    
    Set l_obj = CreateObject("dbSchedTrade.cdSchedTrade")

    Set l_objRS = l_obj.GetTradesByEmployee(v_lngEmployees, v_datStartDate, v_datEndDate)
    Set GetTradesByEmployee = l_objRS
    Set l_objRS = Nothing
    Set l_obj = Nothing
    
    Exit Function

eh:
    Set l_objRS = Nothing
    Set l_obj = Nothing
    
    
    TrsRaiseError m_strCMODULENAME, "GetTradesByEmployee", , , , , True
    
End Function

Public Function GetTradesByEmployeeAndLine( _
        ByVal v_lngEmployees As Variant, _
        ByVal v_aryWorkGroups As Variant, _
        ByVal v_aryLineNums As Variant, _
        ByVal v_datStartDate As Date, _
        ByVal v_datEndDate As Date) _
            As ADODB.Recordset
                    
'This function retrieves all trades initiated by the employee
' or from an employee on the line.
'In the case of a demand line changing employees, the employee
' that requested the trade will be displayed.
'The employee currenly on the line will not be displayed

    Dim l_obj As dbSchedTrade.cdSchedTrade
    Dim l_objRS As Recordset
    
    On Error GoTo eh
    
    
    Set l_obj = CreateObject("dbSchedTrade.cdSchedTrade")

    Set l_objRS = l_obj.GetTradesByEmployeeAndLine(v_lngEmployees, v_aryWorkGroups, v_aryLineNums, v_datStartDate, v_datEndDate)
    Set GetTradesByEmployeeAndLine = l_objRS
    Set l_objRS = Nothing
    Set l_obj = Nothing
    
    Exit Function

eh:
    Set l_objRS = Nothing
    Set l_obj = Nothing
    
    
    TrsRaiseError m_strCMODULENAME, "GetTradesByEmployeeAndLine", , , , , True
    
End Function


Public Function GetTradesByLine( _
        ByVal v_aryWorkGroups As Variant, _
        ByVal v_aryLineNums As Variant, _
        ByVal v_datStartDate As Date, _
        ByVal v_datEndDate As Date) _
            As ADODB.Recordset
                    
'This function retrieves all trades initiated from an employee
' on the line.
'In the case of line changes employees, the employee
' that requested the trade will be displayed.
'The employee currenly on the line will not be displayed
                    

    Dim l_obj As dbSchedTrade.cdSchedTrade
    Dim l_objRS As Recordset
    
    On Error GoTo eh
    
    
    Set l_obj = CreateObject("dbSchedTrade.cdSchedTrade")

    Set l_objRS = l_obj.GetTradesByLine(v_aryWorkGroups, v_aryLineNums, v_datStartDate, v_datEndDate)
    Set GetTradesByLine = l_objRS
    Set l_objRS = Nothing
    Set l_obj = Nothing
    
    Exit Function

eh:
    Set l_objRS = Nothing
    Set l_obj = Nothing
    
    
    TrsRaiseError m_strCMODULENAME, "GetTradesByLine", , , , , True
    
End Function


