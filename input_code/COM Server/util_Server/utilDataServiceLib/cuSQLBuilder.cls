VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cuSQLBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum TrActualTimeTable
    TrTable_Actual = 1
    TrTable_History = 2
    TrTable_ActualAndHistory = 3
End Enum

'Private Enum TrSetCurrentTimestampEnum
'    TrSetTSDisable = 0
'    TrSetTSAlways = 1
'    TrSetTSInsertOnly = 2
'    TrSetTSInsertAndEmpty = 3
'    TrSetTSTouch = 4
'End Enum

' Public Const CKEYCUTOFFDATE = "HS_CUTOFF_DATE" 'use to get date using spm GetActualTimeHSCutoffDate from codetables


Public Function SQLStatement( _
            ByRef v_rsData As ADODB.Recordset, _
            ByRef r_cexRecordset As CExRecordset, _
            Optional ByVal v_varAdditional As Variant) As String
' Answer a SQL statement for the current record.
    If v_rsData.Status = adRecNew Then
        SQLStatement = InsertStatement(v_rsData, r_cexRecordset, v_varAdditional)
    ElseIf v_rsData.Status = adRecDeleted Then
        SQLStatement = DeleteStatement(v_rsData, r_cexRecordset)
    ElseIf v_rsData.Status = adRecModified Then
        SQLStatement = UpdateStatement(v_rsData, r_cexRecordset)
    End If
End Function


Public Function InsertStatement( _
            ByRef v_rsData As ADODB.Recordset, _
            ByRef r_cexRecordset As CExRecordset, _
            Optional ByVal v_varAdditional As Variant) As String
' v_varAdditional should be a 2 dim array with column names and values.
' this is used to add to the insert statement fields that are not in the
' v_rsData.
            
    Dim strSQL As String
    Dim strTableClause As String
    Dim colList As String
    Dim valList As String
    Dim numberNewFields As Integer
    Dim fieldsProcessed As Integer
    Dim l_aryColumns As Variant
    Dim i As Integer
    Dim l_strEND_WORK_DT As String
    Dim l_strOTHER_PAY_SEQ_NUM As String
'    Dim l_strLAST_MOD_VER_NUM As String
    Dim l_strSTART_WORK_TIME As String
    Dim l_strEND_WORK_TIME As String
    Dim addTris30DefaultFields() As Variant
    Dim l_strTableName As String
   
    On Error GoTo eh
'
'    If IsMissing(v_varIgnore) Then
'        v_varIgnore = Array("")
'    End If
    
    
    l_aryColumns = FieldsChanged(v_rsData, r_cexRecordset, False)
    numberNewFields = SizeOfColumnArray(l_aryColumns)

    If numberNewFields = 0 Then
        InsertStatement = ""
        Exit Function
    End If
    
    l_strTableName = r_cexRecordset.FullTableName
    
    strTableClause = "insert into " & l_strTableName & " ("
    valList = " Values("
    For i = 1 To numberNewFields
        fieldsProcessed = fieldsProcessed + 1
        If l_strTableName = "METHDBA.TTR_ACTUAL_TIME" Then 'changes for end_time new fields
            If l_aryColumns(i, 1) = "START_TIME" Then
               l_strSTART_WORK_TIME = Replace(l_aryColumns(i, 2), "'", "", , , vbTextCompare) 'eliminate single quote
               If Left(l_strSTART_WORK_TIME, 2) = "99" Then 'other pay
                   l_strOTHER_PAY_SEQ_NUM = CStr(Right(l_strSTART_WORK_TIME, 2))
                   l_strSTART_WORK_TIME = "'00:00'"
               Else
                   l_strSTART_WORK_TIME = l_aryColumns(i, 2) 'START_WORK_TIME equal START_TIME
                   l_strOTHER_PAY_SEQ_NUM = "00"  'other pay ind is initialized to zero
               End If
               colList = colList & "START_WORK_TIME"
               valList = valList & l_strSTART_WORK_TIME
            Else
              If l_aryColumns(i, 1) = "END_TIME" Then
                   l_strEND_WORK_TIME = Replace(l_aryColumns(i, 2), "'", "", , , vbTextCompare) 'eliminate single quote
                   If Left(l_strEND_WORK_TIME, 2) = "99" Then 'other pay
                       l_strEND_WORK_TIME = "'00:00'"
                   Else
                       If l_strEND_WORK_TIME = "00:00" Then
                          l_strEND_WORK_TIME = "'24:00'"
                       Else
                          l_strEND_WORK_TIME = l_aryColumns(i, 2) 'END_WORK_TIME = END_TIME
                       End If
                   End If
                   colList = colList & "END_WORK_TIME"
                   valList = valList & l_strEND_WORK_TIME
               Else
                  If l_aryColumns(i, 1) = "DATE_WORKED" Then
                     l_strEND_WORK_DT = CStr(l_aryColumns(i, 2)) 'END_WORK_DT = DATE_WORKED
                  End If
                  colList = colList & l_aryColumns(i, 1)
                  valList = valList & l_aryColumns(i, 2)
               End If
            End If
        Else
            colList = colList & l_aryColumns(i, 1)
            valList = valList & l_aryColumns(i, 2)
        End If
        If fieldsProcessed < numberNewFields Then
            colList = colList & ", "
            valList = valList & ", "
        Else
            colList = colList & " "
            valList = valList & " "
        End If
    Next
    
    If l_strTableName = "METHDBA.TTR_ACTUAL_TIME" Then 'changes for end_time new fields
        ReDim addTris30DefaultFields(2, 1)
        addTris30DefaultFields(0, 0) = "END_WORK_DT"
        addTris30DefaultFields(1, 0) = "OTHER_PAY_SEQ_NUM"
        addTris30DefaultFields(2, 0) = "LAST_MOD_VER_NUM"
        addTris30DefaultFields(0, 1) = l_strEND_WORK_DT
        addTris30DefaultFields(1, 1) = l_strOTHER_PAY_SEQ_NUM
        addTris30DefaultFields(2, 1) = "3.0"
        For i = LBound(addTris30DefaultFields) To UBound(addTris30DefaultFields)
            colList = colList & ", "
            valList = valList & ", "
            colList = colList & addTris30DefaultFields(i, 0)
            valList = valList & addTris30DefaultFields(i, 1)
        Next
    End If
    
    
    
    If Not IsMissing(v_varAdditional) Then
        For i = LBound(v_varAdditional) To UBound(v_varAdditional)
            colList = colList & ", "
            valList = valList & ", "
            colList = colList & v_varAdditional(i, 0)
            valList = valList & v_varAdditional(i, 1)
        Next
    End If
    
    colList = colList & ") "
    valList = valList & ") "

    strSQL = strTableClause & colList & valList
    
    InsertStatement = strSQL
    Exit Function
    
eh:
    
    TrsRaiseError "mDataServiceLib", "InsertStatement", trErrdbDataService, True

End Function


Public Function UpdateStatement( _
            ByRef v_rsData As ADODB.Recordset, _
            ByRef r_cexRecordset As CExRecordset) As String
    Dim strSQL As String
    Dim strTableClause As String
    Dim strSetList As String
    Dim strWhereClause As String
    Dim l_aryColumns As Variant
    Dim intKeyIndex As Integer
   
    Dim l_strTableName As String
    Dim l_varPrimaryKeyNames As Variant
   

    l_aryColumns = FieldsChanged(v_rsData, r_cexRecordset, True)
    If SizeOfColumnArray(l_aryColumns) = 0 Then
        UpdateStatement = ""
        Exit Function
    End If

' get these items from the dictionary or other container.
    l_strTableName = r_cexRecordset.FullTableName
    l_varPrimaryKeyNames = r_cexRecordset.PrimaryKeys

    strTableClause = "update " & l_strTableName & " "
    strSetList = BuildSetStatement(l_aryColumns, l_strTableName)
'    strWhereClause = BuildWhereClause(v_rsData, l_varPrimaryKeyNames)
    strWhereClause = BuildWhereClause(v_rsData, r_cexRecordset)
    strSQL = strTableClause & strSetList & strWhereClause
    
    UpdateStatement = strSQL
    
End Function


Public Function DeleteStatement( _
            ByRef v_rsData As ADODB.Recordset, _
            ByRef r_cexRecordset As CExRecordset) As String
    Dim strSQL As String
    Dim strTableClause As String
    Dim strWhereClause As String
    Dim intKeyIndex As Integer
   
    strTableClause = "delete from " & r_cexRecordset.FullTableName
'    strWhereClause = BuildWhereClause(v_rsData, r_cexRecordset.PrimaryKeys)
    strWhereClause = BuildWhereClause(v_rsData, r_cexRecordset)
    strSQL = strTableClause & strWhereClause
    
    DeleteStatement = strSQL
    
End Function


Public Function FormatSetStatement(ByVal v_rsData As ADODB.Recordset) As String
    Dim l_aryColumns As Variant
    
    l_aryColumns = FieldsChanged(v_rsData)
    FormatSetStatement = BuildSetStatement(l_aryColumns)
    
End Function


'Private Function BuildWhereClause( _
'            ByVal v_rsData As ADODB.Recordset, _
'            ByVal v_varPrimaryKeyNames As Variant) As String
'    Dim strWhereClause As String
'    Dim intKeyIndex As Integer
'    Dim i As Integer
'
'    intKeyIndex = 1
'    strWhereClause = " where "
'    For i = LBound(v_varPrimaryKeyNames) To UBound(v_varPrimaryKeyNames)
'        If i > LBound(v_varPrimaryKeyNames) Then
'            strWhereClause = strWhereClause & " and "
'        End If
'        strWhereClause = strWhereClause & v_varPrimaryKeyNames(i) & " = " _
'            & DelimitedDataValue(v_rsData.Fields(v_varPrimaryKeyNames(i)).OriginalValue)
'        intKeyIndex = intKeyIndex + 1
'    Next
'
'    BuildWhereClause = strWhereClause
'
'End Function


Private Function BuildWhereClause( _
            ByVal v_rsData As ADODB.Recordset, _
            ByRef r_cexRecordset As CExRecordset) As String
    Dim strWhereClause As String
    Dim intKeyIndex As Integer
    Dim i As Integer
    Dim v_varPrimaryKeyNames As Variant
    Dim v_varConcurrencyKeyNames As Variant
    Dim l_varFieldPropValue As Variant
    Dim l_colKeyNames As Collection
    Dim l_strKeyName As String
   
    v_varPrimaryKeyNames = r_cexRecordset.PrimaryKeys
    v_varConcurrencyKeyNames = r_cexRecordset.ConcurrencyKeys
    Set l_colKeyNames = New Collection
   
    intKeyIndex = 1
    strWhereClause = " where "
    For i = LBound(v_varPrimaryKeyNames) To UBound(v_varPrimaryKeyNames)
        If i > LBound(v_varPrimaryKeyNames) Then
            strWhereClause = strWhereClause & " and "
        End If
        strWhereClause = strWhereClause & v_varPrimaryKeyNames(i) & " = " _
            & DelimitedDataValue(v_rsData.Fields(v_varPrimaryKeyNames(i)).OriginalValue)
        intKeyIndex = intKeyIndex + 1
        l_colKeyNames.Add v_varPrimaryKeyNames(i)
    Next
    
    If Not IsEmpty(v_varConcurrencyKeyNames) Then
        For i = LBound(v_varConcurrencyKeyNames) To UBound(v_varConcurrencyKeyNames)
            l_strKeyName = v_varConcurrencyKeyNames(i)
            If Not IncludesString(l_strKeyName, l_colKeyNames) Then
                If intKeyIndex > 1 Then     'at least 1 primary key
                    strWhereClause = strWhereClause & " and "
                End If
                If r_cexRecordset.UseClientConcurrencyKey(v_varConcurrencyKeyNames(i)) Then
                    strWhereClause = strWhereClause & v_varConcurrencyKeyNames(i) & " = " _
                        & DelimitedDataValue(v_rsData.Fields(v_varConcurrencyKeyNames(i)).Value)
                Else
                    strWhereClause = strWhereClause & v_varConcurrencyKeyNames(i) & " = " _
                        & DelimitedDataValue(v_rsData.Fields(v_varConcurrencyKeyNames(i)).OriginalValue)
                End If
                intKeyIndex = intKeyIndex + 1
                l_colKeyNames.Add v_varConcurrencyKeyNames(i)
            End If
        Next
    End If
    
    BuildWhereClause = strWhereClause
    
End Function


Private Function FieldsChanged( _
                    ByRef v_rsData As ADODB.Recordset, _
                    Optional ByRef r_cexRecordset As CExRecordset, _
                    Optional ByVal v_blnUpdate As Boolean = True) As Variant
'Answer a 1-based, 2 dim string array. Sub 1 represents column name and Sub 2
' represents a column value formatted for sql string.
    Dim l_objField As ADODB.Field
    Dim l_aryColumns(1 To 100, 1 To 2) As String
    Dim l_aryColumnsEmptyReturn(1 To 1, 1 To 2) As String
    Dim l_intNumberofColumns As Integer
    Dim l_strValue As String
    Dim l_strFormat As String
    Dim l_intPrecision As Integer
    Dim l_intScale As Integer
    Dim l_intI As Integer
    Dim l_intFieldsNotRequired As Integer
    
    On Error GoTo eh
    
    l_intNumberofColumns = 0
    For Each l_objField In v_rsData.Fields
        If IsFieldValueCandidate(l_objField, r_cexRecordset, v_blnUpdate) Then
            If IsFieldValueSet(l_objField) Then
'            Or Not v_blnUpdate Then       'i.e. an insert       still needed?
                l_intNumberofColumns = l_intNumberofColumns + 1
                l_aryColumns(l_intNumberofColumns, 1) = l_objField.Name
                Select Case l_objField.Type
                
                Case adChar, adVarChar
                     l_aryColumns(l_intNumberofColumns, 2) = "'" & ReplaceString(l_objField.Value, "'", "") & "'"
                        
                Case adDate, adDBDate
                    l_strValue = Format(l_objField.Value, "yyyy-mm-dd")
                    If l_strValue = "0100-01-01" Then
                        l_strValue = "'0001-01-01'"
                    Else
                        l_strValue = "'" & l_strValue & "'"
                    End If
                    l_aryColumns(l_intNumberofColumns, 2) = l_strValue
                              
                Case adDecimal, adNumeric
                    l_intScale = l_objField.NumericScale
                    l_intPrecision = l_objField.Precision
                    l_intI = l_intPrecision - l_intScale
                    l_strFormat = ""
                    If l_intScale > 0 Then
                        l_strFormat = "." & String(l_intScale, "0")
                    End If
                    If l_intI > 0 Then
                        l_strFormat = String(l_intI - 1, "#") & "0" & l_strFormat
                    End If
                    l_strValue = Format(l_objField.Value, l_strFormat)
                    l_aryColumns(l_intNumberofColumns, 2) = l_strValue
                          
                Case adInteger, adSmallInt
                    l_aryColumns(l_intNumberofColumns, 2) = CStr(l_objField.Value)
                    
                Case Else
                  ' below are the other possible values that are covered in the else
                  ' adBigInt , adBinary, adBoolean, adBSTR, adChapter, adCurrency
                  ' adDBFileTime, adDBTime, adDBTimeStamp, adDouble, adError, adFileTime
                  ' adGUID, adIDispatch, adIUnknown, adLongVarBinary, adLongVarChar
                  ' adLongVarWChar, adPropVariant, adSingle, adTinyInt, adUnsignedBigInt
                  ' adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt, adUserDefined
                  ' adVarBinary, adVarNumeric, adVarWChar, adWChar
                            
                End Select
            End If
        'Anything set here should answer false IsFieldValueCandidate
        'Make it the negative test?
        'At least make this a generic method that sets to appropriate db2 special register value.
        ElseIf IsSetCurrentTimestamp(l_objField, r_cexRecordset, v_blnUpdate) Then
            l_intNumberofColumns = l_intNumberofColumns + 1
            l_aryColumns(l_intNumberofColumns, 1) = l_objField.Name
            l_aryColumns(l_intNumberofColumns, 2) = "CURRENT TIMESTAMP"
        End If
        If FieldIgnoredWhenSoleUpdate(r_cexRecordset, l_objField) Then
            l_intFieldsNotRequired = l_intFieldsNotRequired + 1
        End If
    Next l_objField
    
    If l_intNumberofColumns - l_intFieldsNotRequired > 0 Then
        FieldsChanged = l_aryColumns
    Else
        FieldsChanged = l_aryColumnsEmptyReturn
    End If
    Exit Function
    
eh:
    TrsRaiseError "mDataServiceLib", "FieldsChanged", trErrdbDataService, True, , , True
End Function


Private Function BuildSetStatement(l_aryColumns As Variant, _
        Optional ByVal v_strTableName As String _
) As String
    Dim l_intNumberofColumns As Integer
    Dim l_intI As Integer
    Dim i As Integer
    Dim l_strSetStatement As String
    
    Dim l_strEND_WORK_DT As String
    Dim l_strOTHER_PAY_SEQ_NUM As String
'    Dim l_strLAST_MOD_VER_NUM As String
    Dim l_strSTART_WORK_TIME As String
    Dim l_strEND_WORK_TIME As String
    Dim addTris30DefaultFields() As Variant
    Dim l_strTableName As String
    
    If IsMissing(v_strTableName) Then
        addTris30DefaultFields = Array("")
    Else
        l_strTableName = v_strTableName
    End If
    
    l_intNumberofColumns = SizeOfColumnArray(l_aryColumns)
    
    l_strSetStatement = ""
    For l_intI = 1 To l_intNumberofColumns
        If l_intI = 1 Then
            If l_strTableName = "METHDBA.TTR_ACTUAL_TIME" Then 'changes for end_time new fields
                If l_aryColumns(l_intI, 1) = "START_TIME" Then
                   l_strSTART_WORK_TIME = Replace(l_aryColumns(l_intI, 2), "'", "", , , vbTextCompare) 'eliminate single quote
                   If Left(l_strSTART_WORK_TIME, 2) = "99" Then 'other pay
                       l_strOTHER_PAY_SEQ_NUM = CStr(Right(l_strSTART_WORK_TIME, 2))
                       l_strSTART_WORK_TIME = "'00:00'"
                   Else
                       l_strSTART_WORK_TIME = l_aryColumns(l_intI, 2) 'START_WORK_TIME equal START_TIME
                       l_strOTHER_PAY_SEQ_NUM = "00"  'other pay ind is initialized to zero
                   End If
'                   l_strLAST_MOD_VER_NUM = "3.0"
                   l_strSetStatement = "SET " & "START_WORK_TIME" & " = " & l_strSTART_WORK_TIME
                Else
                  If l_aryColumns(l_intI, 1) = "END_TIME" Then
                       l_strEND_WORK_TIME = Replace(l_aryColumns(l_intI, 2), "'", "", , , vbTextCompare) 'eliminate single quote
                       If Left(l_strEND_WORK_TIME, 2) = "99" Then 'other pay
                           l_strEND_WORK_TIME = "'00:00'"
                       Else
                           If l_strEND_WORK_TIME = "00:00" Then
                              l_strEND_WORK_TIME = "'24:00'"
                           Else
                              l_strEND_WORK_TIME = l_aryColumns(l_intI, 2) 'END_WORK_TIME = END_TIME
                           End If
                       End If
'                       l_strLAST_MOD_VER_NUM = "3.0"
                       l_strSetStatement = "SET " & "END_WORK_TIME" & " = " & l_strEND_WORK_TIME
                  Else
                      If l_aryColumns(l_intI, 1) = "DATE_WORKED" Then
                         l_strEND_WORK_DT = CStr(l_aryColumns(l_intI, 2)) 'END_WORK_DT = DATE_WORKED
                      End If
                      l_strSetStatement = "SET " & l_aryColumns(l_intI, 1) & " = " & l_aryColumns(l_intI, 2)
                  End If
                End If
            Else
                l_strSetStatement = "SET " & l_aryColumns(l_intI, 1) & " = " & l_aryColumns(l_intI, 2)
            End If
        Else
            If l_strTableName = "METHDBA.TTR_ACTUAL_TIME" Then 'changes for end_time new fields
                If l_aryColumns(l_intI, 1) = "START_TIME" Then
                   l_strSTART_WORK_TIME = Replace(l_aryColumns(l_intI, 2), "'", "", , , vbTextCompare) 'eliminate single quote
                   If Left(l_strSTART_WORK_TIME, 2) = "99" Then 'other pay
                       l_strOTHER_PAY_SEQ_NUM = CStr(Right(l_strSTART_WORK_TIME, 2))
                       l_strSTART_WORK_TIME = "'00:00'"
                   Else
                       l_strSTART_WORK_TIME = l_aryColumns(l_intI, 2) 'START_WORK_TIME equal START_TIME
                       l_strOTHER_PAY_SEQ_NUM = "00"  'other pay ind is initialized to zero
                   End If
'                   l_strLAST_MOD_VER_NUM = "3.0"
                   l_strSetStatement = l_strSetStatement & " , " & "START_WORK_TIME" & " = " & l_strSTART_WORK_TIME
                Else
                  If l_aryColumns(l_intI, 1) = "END_TIME" Then
                       l_strEND_WORK_TIME = Replace(l_aryColumns(l_intI, 2), "'", "", , , vbTextCompare) 'eliminate single quote
                       If Left(l_strEND_WORK_TIME, 2) = "99" Then 'other pay
                           l_strEND_WORK_TIME = "'00:00'"
                       Else
                           If l_strEND_WORK_TIME = "00:00" Then
                              l_strEND_WORK_TIME = "'24:00'"
                           Else
                              l_strEND_WORK_TIME = l_aryColumns(l_intI, 2) 'END_WORK_TIME = END_TIME
                           End If
                       End If
'                      l_strLAST_MOD_VER_NUM = "3.0"
                      l_strSetStatement = l_strSetStatement & " , " & "END_WORK_TIME" & " = " & l_strEND_WORK_TIME
                  Else
                      If l_aryColumns(l_intI, 1) = "DATE_WORKED" Then
                         l_strEND_WORK_DT = CStr(l_aryColumns(l_intI, 2)) 'END_WORK_DT = DATE_WORKED
                      End If
                      l_strSetStatement = l_strSetStatement & " , " & l_aryColumns(l_intI, 1) & " = " & l_aryColumns(l_intI, 2)
                  End If
                End If
            Else
                l_strSetStatement = l_strSetStatement & " , " & l_aryColumns(l_intI, 1) & " = " & l_aryColumns(l_intI, 2)
            End If
        
        End If
    Next l_intI
    
    If l_strTableName = "METHDBA.TTR_ACTUAL_TIME" Then 'changes for end_time new fields
        ReDim addTris30DefaultFields(2, 1)
        If l_strEND_WORK_DT <> "" Then
            addTris30DefaultFields(0, 0) = "END_WORK_DT"
            addTris30DefaultFields(0, 1) = l_strEND_WORK_DT
'            l_strLAST_MOD_VER_NUM = "3.0"
        End If
        If l_strOTHER_PAY_SEQ_NUM <> "" Then
            addTris30DefaultFields(1, 0) = "OTHER_PAY_SEQ_NUM"
            addTris30DefaultFields(1, 1) = l_strOTHER_PAY_SEQ_NUM
'            l_strLAST_MOD_VER_NUM = "3.0"
        End If
'        If l_strLAST_MOD_VER_NUM <> "" Then
        addTris30DefaultFields(2, 0) = "LAST_MOD_VER_NUM"
        addTris30DefaultFields(2, 1) = "3.0"
'        End If
        For i = LBound(addTris30DefaultFields) To UBound(addTris30DefaultFields)
            If addTris30DefaultFields(i, 0) <> "" Then
                l_strSetStatement = l_strSetStatement & " , " & addTris30DefaultFields(i, 0) & " = " & addTris30DefaultFields(i, 1)
            End If
        Next
    End If

    BuildSetStatement = l_strSetStatement

End Function


Private Function SizeOfColumnArray(r_aryColumns As Variant) As Integer
'Count the number of elements before the first empty string
    Dim i As Integer
    For i = 1 To UBound(r_aryColumns)
        If r_aryColumns(i, 1) = "" Then
            SizeOfColumnArray = i - 1
            Exit For
        End If
    Next
End Function


Private Function DelimitedDataValue(dv As Variant) As String
    Dim ddv As String, fieldType As Integer
    fieldType = VarType(dv)
    Select Case fieldType
        Case vbString
            ddv = "'" & dv & "'"
        Case vbInteger, vbLong
            ddv = CStr(dv)
        Case vbDate
            ddv = "'" & Format(dv, "mm/dd/yyyy") & "'"
        Case Else
            ddv = Format(dv)
    End Select
    DelimitedDataValue = ddv
    
End Function


Public Function BuildInClause( _
            ByVal v_ColumnName As String, _
            ByVal v_aryValues As Variant, _
            ByVal v_bolDelimit As Boolean _
            ) As String
            
'This function will Build an SQL in clause
'The array of values is a string array - One Dimensional
'The boolean parameter is to signify that the
'field the values need delimited with a
' single quote. Characters and dates need
'delimited.

    Dim str As String
    Dim i As Long
    Dim l_lngStart As Long
    Dim l_lngEnd As Long
    
    l_lngStart = LBound(v_aryValues)
    l_lngEnd = UBound(v_aryValues)
    str = v_ColumnName & " IN ( "
    For i = l_lngStart To l_lngEnd
        If v_bolDelimit Then
            str = str & "'" & v_aryValues(i) & "'"
        Else
            str = str & v_aryValues(i)
        End If
        If Not (i = l_lngEnd) Then
            str = str & ","
        Else
            str = str & " ) "
        End If
    Next
    
    BuildInClause = str
            

End Function


Private Function IsFieldValueSet( _
                ByRef r_adoField As ADODB.Field) As Boolean
                        
    Dim l_blnChanged As Boolean
            
            
On Error GoTo eh

    IsFieldValueSet = True
        
    If r_adoField.Type <> adEmpty And Not IsEmpty(r_adoField.Value) Then
        'continue
    Else
        IsFieldValueSet = False
        GoTo xit
    End If
    
    If IsNull(r_adoField.Value) And IsNull(r_adoField.OriginalValue) Then
        IsFieldValueSet = False
        GoTo xit
    End If
    
'   If r_adoField.OriginalValue <> r_adoField.Value Then 'fix to allow zero for nondefault fields TRIS 4.0 - empty original value equated to zero
    If IsNull(r_adoField.OriginalValue) Then
        l_blnChanged = True
    Else
        l_blnChanged = False
        If IsNull(r_adoField.Value) Then
            IsFieldValueSet = False
            GoTo xit
        End If
    End If
    
    If l_blnChanged Then
    Else
        If r_adoField.OriginalValue <> r_adoField.Value _
            Or IsEmpty(r_adoField.OriginalValue) Then
        Else
            IsFieldValueSet = False
            GoTo xit
        End If
    End If
    
xit:
    Exit Function
    
eh:
    TrsRaiseError "mDataServiceLib", "IsFieldValueSet", trErrdbDataService, True
    
End Function


Private Function IsFieldValueCandidate( _
        ByRef r_adoField As ADODB.Field, _
        ByRef r_cexRecordset As CExRecordset, _
        ByVal v_blnUpdate As Boolean) As Boolean
' Properties IGNOREFIELD and SETCURRENTTIMESTAMP should not be set with
' recordset value.
    Dim l_cexField As CExField
    Dim l_varProperty As Variant
    Dim l_strFieldName As String
    
On Error GoTo eh

    IsFieldValueCandidate = True
    If r_cexRecordset Is Nothing Then
        Exit Function
    End If

    l_strFieldName = r_adoField.Name
    
    Set l_cexField = r_cexRecordset.Field(Trim(l_strFieldName))
    If Not l_cexField Is Nothing Then
        l_varProperty = l_cexField.Properties("IGNOREFIELD")
        If Not IsEmpty(l_varProperty) Then
            If l_varProperty Then
                IsFieldValueCandidate = False
                Exit Function
            End If
        End If
        l_varProperty = l_cexField.Properties("SETCURRENTTIMESTAMP")   '(0,1,2,3,4)
        If Not IsEmpty(l_varProperty) Then
            'always set the current timestamp(1), insert only(2), or just when insert and no value(3)
            'touch timestamp(4) - regardless of any other updates.
            If l_varProperty = 0 Then
                Exit Function
            ElseIf l_varProperty = TrSetTSAlways Then
                IsFieldValueCandidate = False
                Exit Function
            ElseIf l_varProperty = TrSetTSInsertOnly And Not v_blnUpdate Then
                IsFieldValueCandidate = False
                Exit Function
            ElseIf l_varProperty = TrSetTSInsertAndEmpty And (Not v_blnUpdate) And IsFieldValueEmpty(r_adoField) Then
                IsFieldValueCandidate = False
                Exit Function
            ElseIf l_varProperty = TrSetTSTouch Then
                IsFieldValueCandidate = False
                Exit Function
            End If
        End If
    End If
    
    Exit Function
    
eh:
    TrsRaiseError "mDataServiceLib", "IsFieldValueCandidate", trErrdbDataService, True
End Function


Private Function FieldIgnoredWhenSoleUpdate( _
        ByRef r_cexRecordset As CExRecordset, _
        ByRef r_adoField As ADODB.Field) As Boolean
' Properties IGNOREFIELD and SETCURRENTTIMESTAMP should not be set with
' recordset value.
'IGNORESOLEUPDATE
    Dim l_cexField As CExField
    Dim l_varProperty As Variant
    Dim l_strFieldName As String
    
On Error GoTo eh

    If r_cexRecordset Is Nothing Then
        Exit Function
    End If

    l_strFieldName = r_adoField.Name
    
    Set l_cexField = r_cexRecordset.Field(Trim(l_strFieldName))
    If l_cexField Is Nothing Then
        FieldIgnoredWhenSoleUpdate = False
        Exit Function
    End If
    
' TOUCH takes precedence over IGNORESOLEUPDATE, which is set by system
    l_varProperty = l_cexField.Properties("TOUCH")
    If Not IsEmpty(l_varProperty) Then
        If l_varProperty Then
            FieldIgnoredWhenSoleUpdate = False
            GoTo xit
        End If
    End If
    
    l_varProperty = l_cexField.Properties("IGNORESOLEUPDATE")
    If Not IsEmpty(l_varProperty) Then
        FieldIgnoredWhenSoleUpdate = l_varProperty
    End If
    
xit:
    Exit Function
    
eh:
    TrsRaiseError "mDataServiceLib", "FieldIgnoredWhenSoleUpdate", trErrdbDataService, True
End Function


Private Function IsFieldValueEmpty( _
                     ByRef r_adoField As ADODB.Field) As Boolean
            
    If IsNull(r_adoField.Value) Then
        IsFieldValueEmpty = True
    ElseIf IsEmpty(r_adoField.Value) Then
        IsFieldValueEmpty = True
    Else
        Select Case r_adoField.Type
            Case adChar, adVarChar
                If Trim(r_adoField.Value) = "" Then
                    IsFieldValueEmpty = True
                End If
        End Select
    End If
    Exit Function
    
eh:
    TrsRaiseError "mDataServiceLib", "IsFieldValueEmpty", trErrdbDataService, True
End Function


Private Function IsSetCurrentTimestamp( _
        ByRef r_adoField As ADODB.Field, _
        ByRef r_cexRecordset As CExRecordset, _
        ByVal v_blnUpdate As Boolean) As Boolean
' Properties SETCURRENTTIMESTAMP should not be set with recordset value.
    Dim l_cexField As CExField
    Dim l_varProperty As Variant
    Dim l_strFieldName As String
    
On Error GoTo eh
    If r_cexRecordset Is Nothing Then
        IsSetCurrentTimestamp = False
        Exit Function
    End If

    l_strFieldName = r_adoField.Name
    Set l_cexField = r_cexRecordset.Field(Trim(l_strFieldName))
    If Not l_cexField Is Nothing Then
        l_varProperty = l_cexField.Properties("SETCURRENTTIMESTAMP")

'This should be same code as IsFieldCandidate
        If Not IsEmpty(l_varProperty) Then
            If l_varProperty = 0 Then
                IsSetCurrentTimestamp = False
            ElseIf l_varProperty = TrSetTSAlways Then
                IsSetCurrentTimestamp = True
            ElseIf l_varProperty = TrSetTSInsertOnly And Not v_blnUpdate Then
                IsSetCurrentTimestamp = True
            ElseIf l_varProperty = TrSetTSInsertAndEmpty And (Not v_blnUpdate) And IsFieldValueEmpty(r_adoField) Then
                IsSetCurrentTimestamp = True
            ElseIf l_varProperty = TrSetTSTouch Then   'like 1, except do not ignore sole update
                IsSetCurrentTimestamp = True
            End If
        End If
    End If
    
    Exit Function
    
eh:
    TrsRaiseError "mDataServiceLib", "IsSetCurrentTimestamp", trErrdbDataService, True
End Function



Public Function BuildArrayPartValueKeys( _
            ByVal v_datStartDate As Date, _
            ByVal v_datEndDate As Date _
            ) As Variant

    Dim l_aryPartValueKeys(1) As Variant
    
    l_aryPartValueKeys(0) = Right(str(DatePart("yyyy", v_datStartDate)), 1)
    l_aryPartValueKeys(1) = Right(str(DatePart("yyyy", v_datEndDate)), 1)

    BuildArrayPartValueKeys = l_aryPartValueKeys

End Function



'Please note:
'Function IsStringInTheList was added to remove the
'dependancy of having to include mCollectionLib.bas
'whenever the developer needed mDataServiceLib.bas
'The function IsStringInTheList is IsStringInList
' in mCollectionLib.bas
'
' Move to common function library and use it here.
'

'
' MTS lib is included for the standard error handling bas module.
'
'



Public Function ReplaceString(varString, varFind, varReplace) As String
    ' this function will replace search strString and replace all occurences of
    ' strFind with strReplace
    ' strFind and strReplace can be of different lenghts. strReplace can be an empty string
    Dim strString As String
    Dim strFind As String
    Dim strReplace As String
    Dim strTemp As String
    Dim i As Integer

    strString = CStr(varString)
    strFind = CStr(varFind)
    strReplace = CStr(varReplace)
    i = 1
    strTemp = strString
    Do
        ' search for the find string
        i = InStr(i, strTemp, strFind)
        ' if it is not there, we are done
        If i = 0 Then Exit Do
        ' replace the find string with the replace string
        strTemp = Left(strTemp, i - 1) & strReplace & Mid(strTemp, i + Len(strFind))
        ' start next search after the string that we just inserted
        i = i + Len(strReplace)
    Loop
        
    ReplaceString = strTemp

End Function


Public Function IsStringInTheList( _
        ByVal v_strValue As String, _
        ByVal v_strList As Variant) As Boolean
'is v_strValue equal to any of values in v_strList
'v_strList should be a string array or variant string array.
    
    Dim i As Integer
    For i = LBound(v_strList) To UBound(v_strList)
        If UCase(v_strValue) = UCase(v_strList(i)) Then
            IsStringInTheList = True
            Exit Function
        End If
    Next

End Function

Public Function DateAsEnclosedString(ByVal v_datDate As Date) As String
    DateAsEnclosedString = "'" & Format(v_datDate, "mm/dd/yyyy") & "'"
End Function

Public Function SinqleQuotedString(ByVal v_strData As String) As String
    SinqleQuotedString = StringEnclosedBy(v_strData, "'")
End Function

Public Function StringEnclosedBy(ByVal v_strData As String, _
                    ByVal v_strEncloseChar As String) As String
    StringEnclosedBy = v_strEncloseChar & v_strData & v_strEncloseChar
End Function

