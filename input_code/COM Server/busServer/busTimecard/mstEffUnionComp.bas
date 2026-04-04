Attribute VB_Name = "mstEffUnionComp"
Option Explicit

Private Const mstrCMODULENAME As String = "mstEffUnionComp"

Public Function IsVacationReliefCode( _
        ByVal v_strUnionCode As String, _
        ByVal v_datEffDate As Date, _
        ByVal v_strCompCode As String) As Boolean

On Error GoTo eh
    
    IsVacationReliefCode = _
        GetEffComponentBusManager.IsVacationReliefCode( _
                v_strUnionCode, v_datEffDate, v_strCompCode)
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "IsVacationReliefCode"

End Function

Public Function IsShiftCompCode( _
        ByVal v_strUnionCode As String, _
        ByVal v_datEffDate As Date, _
        ByVal v_strCompCode As String) As Boolean

On Error GoTo eh

    IsShiftCompCode = _
        GetEffComponentBusManager.IsShiftCompCode( _
                v_strUnionCode, v_datEffDate, v_strCompCode)
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "IsShiftCompCode"

End Function

Public Function IsComponentCode( _
        ByVal v_strUnionCode As String, _
        ByVal v_datEffDate As Date, _
        ByVal v_strCompCode As String) As Boolean

On Error GoTo eh
    
    IsComponentCode = _
        GetEffComponentBusManager.IsComponentCode( _
                v_strUnionCode, v_datEffDate, v_strCompCode)
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "IsComponentCode"

End Function

Public Function ComponentAmountFor( _
            ByVal v_strUnionCode As String, _
            ByVal v_datEffDate As Date, _
            ByVal v_strCompCode As String) As Variant
' if found, answer a decimal amount, otherwise answer Null
' WHITE069
    Dim str As String
    
On Error GoTo eh

    ComponentAmountFor = _
        GetEffComponentDataManager.UnionComponentAmount( _
                v_strUnionCode, v_datEffDate, v_strCompCode)
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "ComponentAmountFor"

End Function

Public Function ShiftComponentAmountFor( _
            ByVal v_strUnionCode As String, _
            ByVal v_datEffDate As Date, _
            ByVal v_strCompCode As String) As Variant
' if found, answer a decimal amount, otherwise answer Null
    Dim str As String
    
On Error GoTo eh

    ShiftComponentAmountFor = _
        GetEffComponentDataManager.UnionComponentShiftAmount( _
                v_strUnionCode, v_datEffDate, v_strCompCode)
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "ShiftComponentAmountFor"

End Function

Public Function VacationReliefComponentAmountFor( _
        ByVal v_strUnionCode As String, _
        ByVal v_datEffDate As Date, _
        ByVal v_strCompCode As String) As Variant
' if found, answer a decimal amount, otherwise answer Null
    Dim str As String
    
On Error GoTo eh

    VacationReliefComponentAmountFor = _
        GetEffComponentDataManager.UnionComponentVacationReliefAmount( _
                v_strUnionCode, v_datEffDate, v_strCompCode)
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "VacationReliefComponentAmountFor"

End Function
   
Private Function GetEffComponentDataManager() As dbCodeTables.cdCodeTables
    
    Set GetEffComponentDataManager = _
        CreateObject("dbCodeTables.cdCodeTables")

End Function

Private Function GetEffComponentBusManager() As busCodeTables.cbsCodeTables
    
    Set GetEffComponentBusManager = _
        CreateObject("busCodeTables.cbsCodeTables")

End Function

