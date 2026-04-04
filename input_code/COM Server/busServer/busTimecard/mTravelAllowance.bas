Attribute VB_Name = "mTravelAllowance"
Option Explicit

Private Const mstrCMODULENAME As String = "mstTravelAllowance"
Public g_rsTravelAllowance As ADODB.Recordset


Public Function DetermineDistIdandAmount( _
                    ByVal v_strChemTester As String, _
                    ByVal v_intFromLoc As Integer, _
                    ByVal v_intToLoc As Integer, _
                    ByVal v_strTripType As String, _
                    ByVal v_datChargeDate As Date, _
                    ByRef r_strDistId As String, _
                    ByRef r_curAmount As Variant _
                    ) As Boolean

    Dim l_lngYear As Long
    Dim l_strChemTester As String
    
On Error GoTo DetermineDistIdandAmount_EH
    
    DetermineDistIdandAmount = False
    
    '  Is recordset available, Unfilter and Sort the recordset
    IsTravelAllowance
    ResetTravelAllowance
    
    '  Convert ChargeDate to Year
    l_lngYear = Year(v_datChargeDate)
    
    '  Convert Chem Tester
    If v_strChemTester = "Y" Then
        l_strChemTester = "C"
    Else
        l_strChemTester = ""
    End If
    
    '  Filter the recordset with the given arguments
        g_rsTravelAllowance.Filter = MDAC25FilterFix("START_LOC_CD = " & Trim(v_intFromLoc) & _
                                                " AND END_LOC_CD = " & Trim(v_intToLoc) & _
                                                " AND USER_TYPE = '" & Trim(l_strChemTester) & "'" & _
                                                " AND EFF_YEAR = " & l_lngYear)
                                                
    '  Find DistId and Amount
    If Not g_rsTravelAllowance.EOF And Not g_rsTravelAllowance.BOF Then
            r_strDistId = Trim(g_rsTravelAllowance("DISTRIBUTION_ID"))
            r_curAmount = Trim(g_rsTravelAllowance("ALLOWANCE_AMT"))
    Else
            DetermineDistIdandAmount = False
    End If
    
    If UCase(v_strTripType) = "H" Then
        r_curAmount = r_curAmount / 2
    End If

    ResetTravelAllowance
    DetermineDistIdandAmount = True

Exit Function
DetermineDistIdandAmount_EH:

        DetermineDistIdandAmount = False
        TrsRaiseError mstrCMODULENAME, "DetermineDistIdandAmount", Err.Number, , , , True

End Function

Public Function ReleaseTravelAllowance() As Boolean

 On Error GoTo ReleaseTravelAllowance_EH

    If Not g_rsTravelAllowance Is Nothing Then
        Set g_rsTravelAllowance = Nothing
    End If

Exit Function
ReleaseTravelAllowance_EH:

    TrsRaiseError mstrCMODULENAME, "ReleaseTravelAllowance", Err.Number

End Function

Private Sub InitializeTravelAllowance()

    Dim l_objTravelAllowTable As dbCodeTables.cdCodeTables
    
On Error GoTo InitializeTravelAllowance_EH

    Set l_objTravelAllowTable = CreateObject("dbCodeTables.cdCodeTables")
    Set g_rsTravelAllowance = l_objTravelAllowTable.GetTravelAllowances
    
    Set l_objTravelAllowTable = Nothing
    Exit Sub
    
InitializeTravelAllowance_EH:

    TrsRaiseError mstrCMODULENAME, "InitializeTravelAllowance", Err.Number

End Sub

Private Sub IsTravelAllowance()
    
On Error GoTo IsTravelAllowance_EH

    If g_rsTravelAllowance Is Nothing Then
        InitializeTravelAllowance
    End If

Exit Sub
IsTravelAllowance_EH:

    TrsRaiseError mstrCMODULENAME, "IsTravelAllowance", Err.Number

End Sub

Private Sub ResetTravelAllowance()

On Error GoTo ResetTravelAllowance_EH

    g_rsTravelAllowance.Filter = ""
    g_rsTravelAllowance.Sort = ""
    
Exit Sub
ResetTravelAllowance_EH:

    TrsRaiseError mstrCMODULENAME, "ResetTravelAllowance", Err.Number, , , , True

End Sub


