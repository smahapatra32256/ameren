VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cuUtility"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const MODULENAME = "utilADOHelper.cuUtility"

'//from TnvLib.Utility
Public Function PadL(sSource As String, nPad As Long, Optional sPadChar As String = " ") As String
    Dim sPad As String

    If Len(sSource) >= nPad Then
        sPad = Right(sSource, nPad)
    Else
        sPad = String(nPad - Len(sSource), sPadChar) & sSource
    End If
    
    PadL = sPad
    Exit Function
End Function

Public Function PadR(sSource As String, nPad As Long, Optional sPadChar As String = " ") As String
    Dim sPad As String

    If Len(sSource) >= nPad Then
        sPad = Left(sSource, nPad)
    Else
        sPad = sSource & String(nPad - Len(sSource), sPadChar)
    End If
    
    PadR = sPad
    Exit Function
End Function

Public Function rsCompare(ByVal rsSourceLeft As ADODB.Recordset, _
                            ByVal rsSourceRight As ADODB.Recordset, _
                            ByRef sKeyColumns() As String, _
                            ByRef sIgnoreColumns() As String, _
                            ByRef rsDiffLeft As ADODB.Recordset, _
                            ByRef rsDiffRight As ADODB.Recordset, _
                            ByRef rsDiffSummary As ADODB.Recordset) As Long
    Dim nRow1 As Long, nRow2 As Long, nColCount As Long, nCol As Long, bIgnoreCase As Boolean
    Dim bError As Boolean, nErrors As Long, bWarning As Boolean, nWarnings As Long
    Dim sKeyLeft As String, sKeyRight As String
    Dim nComp As Integer, sFlag As String, sRowFlag As String
    Dim sDiffMask As String, xx As Long, sLeft As String, sRight As String
    Dim bErrorOnce As Boolean
    
    nRow1 = 1
    nRow2 = 1
    bIgnoreCase = False
    nColCount = rsSourceLeft.Fields.Count
    nErrors = 0
    nWarnings = 0
    
    rsCreateDiffRecordSets rsSourceLeft, rsSourceRight, rsDiffLeft, rsDiffRight, rsDiffSummary

    If Not rsSourceLeft Is Nothing Then
        If rsSourceLeft.State > 0 Then
            If Not (rsSourceLeft.BOF And rsSourceLeft.EOF) Then
                rsSourceLeft.MoveFirst
            End If
        End If
    End If
    If Not rsSourceRight Is Nothing Then
        If rsSourceRight.State > 0 Then
            If Not (rsSourceRight.BOF And rsSourceRight.EOF) Then
                rsSourceRight.MoveFirst
            End If
        End If
    End If
    
    If Not rsSourceLeft Is Nothing And Not rsSourceRight Is Nothing Then
    If rsSourceLeft.State > 0 And rsSourceRight.State > 0 Then
        Do While Not rsSourceLeft.EOF And Not rsSourceRight.EOF
            bError = False
            bWarning = False
            '//-----compare columns in current row-----//
            sKeyLeft = rsGetKey(rsSourceLeft, sKeyColumns)
            sKeyRight = rsGetKey(rsSourceRight, sKeyColumns)
            nComp = StrComp(IIf(bIgnoreCase, LCase(sKeyLeft), sKeyLeft), IIf(bIgnoreCase, LCase(sKeyRight), sKeyRight), vbTextCompare)
            Select Case nComp
            Case 0 '//-----equal-----//
                sFlag = ""
                sRowFlag = ""
                bErrorOnce = True
                For nCol = 0 To nColCount - 1
                    If rsSourceLeft.Fields(nCol) <> rsSourceRight.Fields(nCol) Then
                        If ArrayFindString(sIgnoreColumns, rsSourceLeft.Fields(nCol).Name) >= 0 Then '//Or (bIgnoreCase And LCase(hfgLeft.TextMatrix(nRow1, nCol)) = LCase(hfgRight.TextMatrix(nRow2, nCol))) Then
                            bWarning = True
                            sFlag = "W"
                        Else
                            bError = True
                            sFlag = "E"
    
                            If bErrorOnce Then
                                nErrors = nErrors + 1
                                rsAppendRowToSummary nRow1, "L", sFlag, rsSourceLeft, rsDiffSummary
                                rsAppendRowToSummary nRow2, "R", sFlag, rsSourceRight, rsDiffSummary
                '                rsAppendRowToSummary 0, "", sDiffMask, rsSourceRight, rsDiffSummary
    
                                rsDiffSummary.AddNew
                                rsDiffSummary!DiffError = "E"
                                rsDiffSummary.Update
                                bErrorOnce = False
                            End If
                            sDiffMask = ""
                            sLeft = rsSourceLeft.Fields(nCol)
                            sRight = rsSourceRight.Fields(nCol)
                            For xx = 1 To Max(Len(sLeft), Len(sRight))
                                If Mid(sLeft, xx, 1) <> Mid(sRight, xx, 1) Then
                                    sDiffMask = sDiffMask & "*"
                                Else
                                    sDiffMask = sDiffMask & " "
                                End If
                            Next xx
                            rsDiffSummary.Fields(nCol + 3).Value = sDiffMask
                            rsDiffSummary.Update
                        End If
                        If Trim(sRowFlag) = "" Or Trim(sRowFlag) = "W" Then '//don't overwrite if already "E"
                            sRowFlag = sFlag
                        End If
                    End If
                Next nCol
                
                rsAppendRowToDiff nRow1, sRowFlag, rsSourceLeft, rsDiffLeft
                rsAppendRowToDiff nRow2, sRowFlag, rsSourceRight, rsDiffRight
                
                If bWarning Then
                    nWarnings = nWarnings + 1
                End If
                rsSourceLeft.MoveNext
                rsSourceRight.MoveNext
    
                nRow1 = nRow1 + 1
                nRow2 = nRow2 + 1
            '//-----Left side is smaller than Right side-----//
            Case -1
                rsAppendRowToDiff nRow1, "I", rsSourceLeft, rsDiffLeft
                rsAppendRowToSummary nRow1, "L", "I", rsSourceLeft, rsDiffSummary
                
                rsSourceLeft.MoveNext
                rsDiffRight.AddNew '//-----push right side down 1 row-----//
                rsDiffRight!DiffError = " "
                rsDiffRight.Update
    '            rsSourceRight.MoveNext
                nErrors = nErrors + 1
                nRow1 = nRow1 + 1
            '//-----Right side is smaller than Left side-----//
            Case 1
                rsAppendRowToDiff nRow2, "I", rsSourceRight, rsDiffRight
                rsAppendRowToSummary nRow2, "R", "I", rsSourceRight, rsDiffSummary
                
                rsSourceRight.MoveNext
                rsDiffLeft.AddNew '//-----push left side down 1 row-----//
                rsDiffLeft!DiffError = " "
                rsDiffLeft.Update
    '            rsSourceLeft.MoveNext
                nErrors = nErrors + 1
                nRow2 = nRow2 + 1
            End Select
        Loop
    End If
    End If
    
    If Not rsSourceLeft Is Nothing And rsSourceLeft.State > 0 Then
        Do While Not rsSourceLeft.EOF
            nErrors = nErrors + 1
            
            rsAppendRowToDiff nRow1, "*", rsSourceLeft, rsDiffLeft
            rsAppendRowToSummary nRow1, "L", "I", rsSourceLeft, rsDiffSummary
            
            rsSourceLeft.MoveNext
            nRow1 = nRow1 + 1
        Loop
    End If
    
    If Not rsSourceRight Is Nothing And rsSourceRight.State > 0 Then
        Do While Not rsSourceRight.EOF
            nErrors = nErrors + 1
    
            rsAppendRowToDiff nRow2, "*", rsSourceRight, rsDiffRight
            rsAppendRowToSummary nRow2, "R", "I", rsSourceRight, rsDiffSummary
            
            rsSourceRight.MoveNext
            nRow2 = nRow2 + 1
        Loop
    End If
    
    rsCompare = nErrors
            
    Exit Function
eh:
End Function

Private Function rsGetKey(ByRef rsDataSource As Variant, sKeyColumns() As String) As String
    Dim xx As Long, sKey As String
    Dim sCell As String
    Dim f As ADODB.Field
    
    For xx = LBound(sKeyColumns) To UBound(sKeyColumns)
        If Trim(sKeyColumns(xx)) <> "" Then
            If Not rsDataSource Is Nothing Then
                Set f = rsDataSource.Fields(sKeyColumns(xx))
                sCell = dbCastToString(f.Value, f)
            End If
            sKey = sKey & sCell
        End If
    Next xx
    
    rsGetKey = sKey
End Function
'//------

Private Function rsCreateDiffRecordSets(ByVal rsSourceLeft As ADODB.Recordset, _
                            ByVal rsSourceRight As ADODB.Recordset, _
                            ByRef rsDiffLeft As ADODB.Recordset, _
                            ByRef rsDiffRight As ADODB.Recordset, _
                            ByRef rsDiffSummary As ADODB.Recordset)
    Dim rsLeft As New ADODB.Recordset, rsRight As New ADODB.Recordset, rsSummary As New ADODB.Recordset
    Dim xx As Long
    Dim f As ADODB.Field, f2 As ADODB.Field, nSize As Long
    
    If Not rsSourceLeft Is Nothing Then
        With rsLeft
            .CursorLocation = adUseClient
            .Fields.Append "DiffRow", adInteger
            .Fields.Append "DiffError", adChar, 1
                    
            For xx = 0 To rsSourceLeft.Fields.Count - 1
                rsAppendField rsSourceLeft, rsLeft, xx
'                Set f = rsSourceLeft.Fields(xx)
'                .Fields.Append f.Name, f.Type, f.DefinedSize
            Next xx
            
            .CursorLocation = adUseClient '//3
            .CursorType = adOpenStatic '//3
            .LockType = adLockBatchOptimistic '//4
            .Open
        End With
    End If
    
    If Not rsSourceRight Is Nothing Then
        With rsRight
            .CursorLocation = adUseClient
            .Fields.Append "DiffRow", adInteger
            .Fields.Append "DiffError", adChar, 1
                    
            For xx = 0 To rsSourceRight.Fields.Count - 1
                rsAppendField rsSourceRight, rsRight, xx
'                Set f = rsSourceRight.Fields(xx)
'                .Fields.Append f.Name, f.Type, f.DefinedSize
            Next xx
            
            .CursorLocation = adUseClient '//3
            .CursorType = adOpenStatic '//3
            .LockType = adLockBatchOptimistic '//4
            .Open
        End With
    End If
    
    With rsSummary
        .CursorLocation = adUseClient
        .Fields.Append "DiffSource", adChar, 1
        .Fields.Append "DiffRow", adInteger
        .Fields.Append "DiffError", adChar, 1
        
        If rsSourceLeft.Fields.Count > 0 Then
            For xx = 0 To rsSourceLeft.Fields.Count - 1
                rsAppendField rsSourceLeft, rsSummary, xx
    '            Set f = rsSourceLeft.Fields(xx)
    '            Select Case f.Type
    '            Case adChar, adVarChar
    '                nSize = f.DefinedSize
    '            Case adDate, adDBDate
    '                nSize = 10
    '            Case adInteger, adSmallInt
    '                nSize = 5
    '            Case adNumeric
    '                nSize = 5
    '            End Select
    '
    '            Set f2 = .Fields.Append(f.Name, adChar, nSize)
    '
    '            If f.Type = adNumeric Then
    '                f2.Precision = f.Precision
    '            End If
            Next xx
        ElseIf rsSourceRight.Fields.Count > 0 Then
            For xx = 0 To rsSourceRight.Fields.Count - 1
                rsAppendField rsSourceRight, rsSummary, xx
            Next
        End If
        
        .CursorLocation = adUseClient '//3
        .CursorType = adOpenStatic '//3
        .LockType = adLockBatchOptimistic '//4
        .Open
    End With
    
    Set rsDiffLeft = rsLeft
    Set rsDiffRight = rsRight
    Set rsDiffSummary = rsSummary
End Function

Private Function rsAppendRowToDiff(ByVal nRow As Long, _
                                        ByVal sError As String, _
                                        ByVal rsSource As ADODB.Recordset, _
                                        ByRef rsDiff As ADODB.Recordset) As Long
    Dim nCol As Long

    With rsDiff
        .AddNew
        !DiffRow = nRow
        !DiffError = sError
    
        For nCol = 0 To rsSource.Fields.Count - 1
            If Not IsNull(rsSource.Fields(nCol).Value) Then
                .Fields(nCol + 2).Value = CStr(rsSource.Fields(nCol).Value)
            End If
        Next nCol
    
        .Update
    End With
    
    rsAppendRowToDiff = 0
End Function

Private Function rsAppendRowToSummary(ByVal nRow As Long, _
                                        ByVal sSource As String, _
                                        ByVal sError As String, _
                                        ByVal rsSource As ADODB.Recordset, _
                                        ByRef rsSummary As ADODB.Recordset) As Long
    Dim nCol As Long

    With rsSummary
        .AddNew
        !DiffSource = sSource
        !DiffRow = nRow
        !DiffError = sError
    
        For nCol = 0 To rsSource.Fields.Count - 1
            If Not IsNull(rsSource.Fields(nCol).Value) Then
                .Fields(nCol + 3).Value = CStr(rsSource.Fields(nCol).Value)
            End If
        Next nCol
    
        .Update
    End With
    
    rsAppendRowToSummary = 0
End Function

Private Function rsAppendField(ByVal rsSource As ADODB.Recordset, ByRef rsDestination As ADODB.Recordset, ByVal nField As Long)
    Dim f As ADODB.Field, f2 As ADODB.Field, nSize As Long

    Set f = rsSource.Fields(nField)
    Select Case f.Type
    Case adChar, adVarChar
        nSize = f.DefinedSize
    Case adDate, adDBDate
        nSize = 10
    Case adInteger
        nSize = 10
    Case adSmallInt
        nSize = 5
    Case adBigInt
        nSize = 20
    Case adNumeric
        nSize = 5
    Case Else
        nSize = 20
    End Select

    rsDestination.Fields.Append f.Name, adChar, nSize

'//    Set f2 = rsDestination.Fields.Append(f.Name, adChar, nSize)
    
'    If f.Type = adNumeric Then
'        f2.Precision = f.Precision
'    End If
            
End Function

Public Function ArrayFind(nArray() As Long, nFind) As Long
    Dim nIndex As Long, xx As Long
    
    On Error GoTo eh
    
    nIndex = -1
    
    For xx = LBound(nArray) To UBound(nArray)
        If nArray(xx) = nFind Then
            nIndex = xx
            Exit For
        End If
    Next xx
    
    ArrayFind = nIndex
    Exit Function
eh:
    ArrayFind = nIndex
End Function

Public Function ArrayFindString(ByRef sArray() As String, ByVal sFind As String) As Long
    Dim nIndex As Long, xx As Long
    
    On Error GoTo eh
    
    nIndex = -1
    
    For xx = LBound(sArray) To UBound(sArray)
        If sArray(xx) = sFind Then
            nIndex = xx
            Exit For
        End If
    Next xx
    
    ArrayFindString = nIndex
    Exit Function
eh:
    ArrayFindString = nIndex
End Function

Public Function Min(vArg1 As Variant, vArg2 As Variant) As Variant
    Min = IIf(vArg1 <= vArg2, vArg1, vArg2)
End Function

Public Function Max(vArg1 As Variant, vArg2 As Variant) As Variant
    Max = IIf(vArg1 >= vArg2, vArg1, vArg2)
End Function

Public Function dbFormat(ByVal v_varData As Variant) As String
    Dim l_strFormated As String, fieldType As Integer
    
    On Error GoTo eh
    
    fieldType = VarType(v_varData)
    Select Case fieldType
        Case vbString
            l_strFormated = "'" & Replace(v_varData, "'", "''") & "'"
        Case vbInteger, vbLong
            l_strFormated = CStr(v_varData)
        Case vbDate
            l_strFormated = "'" & Format(v_varData, "mm/dd/yyyy") & "'"
        Case Else
            l_strFormated = Format(v_varData)
    End Select
    dbFormat = l_strFormated
    Exit Function
eh:
    Err.Raise Err.Number
End Function


Public Function dbTypeCast(ByVal v_strData As String, ByVal v_fADOField As Field) As Variant
    Dim l_varCast As Variant
    
    On Error GoTo eh

    Select Case v_fADOField.Type
        Case adChar
            l_varCast = v_strData
        Case vbInteger, vbLong
            l_varCast = CStr(v_strData)
        Case vbDate, adDBDate
            l_varCast = CDate(v_strData)
        Case Else
            l_varCast = v_strData
    End Select
    dbTypeCast = l_varCast
    Exit Function
eh:
    Err.Raise Err.Number
End Function

Public Function dbCastToString(ByVal v_strData As String, ByVal v_fADOField As Field) As String
    Dim l_strCast As String
    
    On Error GoTo eh

    Select Case v_fADOField.Type
        Case adChar
            l_strCast = v_strData
        Case adInteger, adSmallInt, vbInteger, vbLong
            l_strCast = PadL(v_strData, 5, "0")
'        Case vbInteger, vbLong
'            l_strCast = PadL(v_strData, v_fADOField.DefinedSize, "0")
        Case vbDate, adDate, adDBDate
'            l_strCast = CDate(v_strData)
            l_strCast = Format(CDate(v_strData), "yyyymmdd")
        Case Else
            l_strCast = v_strData
    End Select
    dbCastToString = l_strCast
    Exit Function
eh:
    Err.Raise Err.Number
End Function

