Attribute VB_Name = "mbusOTTemplate"



Public Function NewTemplateErrorset() As ADODB.Recordset
    Dim l_rsErrors As ADODB.Recordset
' Err_Source should be used to capture run-time error object
' source property.
'    Const l_lngCFIELDATTR = adFldUpdatable Or adFldIsNullable

    On Error GoTo eh
    Set l_rsErrors = New ADODB.Recordset
    With l_rsErrors
        .Fields.Append "EMP_NUM", adInteger, , adFldUpdatable
        .Fields.Append "CHARGE_DATE", adDate, , adFldUpdatable
        .Fields.Append "DATE_WORKED", adDate, , adFldUpdatable
        .Fields.Append "START_TIME", adChar, 5, adFldUpdatable
        .Fields.Append "END_TIME", adChar, 5, adFldUpdatable
        .Fields.Append "ELAPSED_HOURS", adDecimal, , adFldUpdatable
            .Fields.Item("ELAPSED_HOURS").Precision = 3
            .Fields.Item("ELAPSED_HOURS").NumericScale = 1
'        .Fields.Append "ELAPSED_HOURS", adChar, 5, adFldUpdatable
        .Fields.Append "EXCEPTION_CD", adChar, 3, adFldUpdatable
        .Fields.Append "OVER_EXC_CD", adChar, 3, adFldUpdatable
        .Fields.Append "ERR_NUMBER", adInteger, , adFldUpdatable
        .Fields.Append "ERR_TYPE", adChar, 2, adFldUpdatable
        .Fields.Append "ERR_DESCRIPTION", adChar, 300, adFldUpdatable
        .Fields.Append "ERR_SOURCE", adChar, 1000, adFldUpdatable
        .Open
    End With
    
    Set NewTemplateErrorset = l_rsErrors
    Exit Function
    
eh:
    Err.Raise Err.Number
End Function
