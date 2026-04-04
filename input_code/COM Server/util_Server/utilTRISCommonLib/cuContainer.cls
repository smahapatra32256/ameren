VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cuContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public Function IsCollection(vData As Variant) As Boolean
    If UCase(TypeName(vData)) = "COLLECTION" Then
        IsCollection = True
    Else
        IsCollection = False
    End If
End Function

Public Function PrintArray(vData As Variant)
    Dim i As Integer
    For i = 0 To UBound(vData)
        Debug.Print vData(i)
    Next

End Function

Public Function CopyObjects(colOriginal As Collection) As Collection
    Dim colCopy As New Collection, obj As Object
    For Each obj In colOriginal
        colCopy.Add obj
    Next
    Set CopyObjects = colCopy
End Function

Public Sub AddObjects(colSource As Collection, _
               colTarget As Collection, _
               Optional blnCheckDuplicates As Boolean = True)
    Dim obj As Object
    For Each obj In colSource
        If blnCheckDuplicates Then
            If Not IsIncluded(obj, colTarget) Then
                colTarget.Add obj
            End If
        Else
            colTarget.Add obj
        End If
    Next
End Sub

Public Function Includes(objIn As Object, col As Collection) As Boolean
    Includes = IsIncluded(objIn, col)
End Function

Public Function IsIncluded(objIn As Object, col As Collection) As Boolean
    Dim obj As Object
    For Each obj In col
        If obj Is objIn Then
            IsIncluded = True
            Exit Function
        End If
    Next
    IsIncluded = False
End Function

Public Function RemoveObjectFrom(anObject As Object, _
            aCollection As Collection) As Boolean
'Answer true is anObject was removed, otherwise false.
'Loop to determine index of anObject
'Use this when anObject was not added with a key.
    Dim ndx As Integer, ndxToRemove As Integer, obj As Object
    ndx = 1
    For Each obj In aCollection
        If obj Is anObject Then
            ndxToRemove = ndx
            Exit For
        Else
            ndx = ndx + 1
        End If
    Next
    If ndx > aCollection.Count Then
        RemoveObjectFrom = False
    Else
        aCollection.Remove ndxToRemove
        RemoveObjectFrom = True
    End If
End Function


Public Function ColReverse(colOriginal As Collection) As Collection

    Dim colReversed As New Collection, i As Integer
    For i = colOriginal.Count To 1 Step -1
        colReversed.Add colOriginal.Item(i)
    Next
    Set ColReverse = colReversed

End Function

Public Function CollectionFromArray(vArray As Variant, _
        Optional startIndex As Integer = 1, _
        Optional endIndex As Integer, _
        Optional withKey As Boolean = False) As Collection
    
    Dim colAnswer As New Collection, i As Integer
    If endIndex = 0 Then endIndex = UBound(vArray)
    For i = startIndex To endIndex
        If withKey Then
            colAnswer.Add vArray(i), vArray(i)
        Else
            colAnswer.Add vArray(i)
        End If
    Next
    Set CollectionFromArray = colAnswer
    
End Function

Public Function ArrayFromCollection( _
            r_objCollection As Collection, _
            Optional v_intBase As Integer = 0) As Variant
' v_intBase should be 0 or 1, if not assume 0
' return empty when collection is empty
    Dim aryReturn() As Variant
    Dim colElement As Variant
    Dim i As Integer
    
On Error GoTo eh
    
    If r_objCollection.Count = 0 Then
        ArrayFromCollection = Empty
        Exit Function
    End If
    
    If v_intBase = 1 Then
        i = 1
        ReDim aryReturn(i To r_objCollection.Count)
    Else
        i = 0
        ReDim aryReturn(i To r_objCollection.Count - 1)
    End If
    For Each colElement In r_objCollection
        aryReturn(i) = colElement
        i = i + 1
    Next
    ArrayFromCollection = aryReturn
    Exit Function
    
eh:
    Err.Raise Err.Number
    
End Function

Public Function ArraySize(ByVal v_varArray As Variant) As Long
    
    If Not IsArray(v_varArray) Then
        ArraySize = 0
        Exit Function
    End If
    
    ArraySize = UBound(v_varArray) - LBound(v_varArray) + 1
    Exit Function

End Function

Public Function DistinctStringCollectionFromArray( _
        vArray As Variant, _
        Optional startIndex As Integer = 1, _
        Optional endIndex As Integer) As Collection
'
'Not tested yet
    Dim i As Integer
    Dim colAnswer As New Collection
    
    If endIndex = 0 Then endIndex = UBound(vArray)
    For i = startIndex To endIndex
'        colAnswer.Add vArray(i)
        AddElement vArray(i), colAnswer, True
    Next
    Set DistinctStringCollectionFromArray = colAnswer
    
End Function

Private Sub AddElement(ByVal anElement As String, _
                r_colNew As Collection, _
                Optional ByVal v_blnDistinct As Boolean)

    If v_blnDistinct Then
        If Not IncludesString(anElement, r_colNew) Then
            r_colNew.Add anElement
        End If
    Else
        r_colNew.Add anElement
    End If

End Sub


Public Function IncludesString(anElement As String, _
            r_colNew As Collection) As Boolean
    Dim lngIndex As Long
    lngIndex = IndexOfString(anElement, r_colNew)
    If lngIndex = 0 Then
        IncludesString = False
    Else
        IncludesString = True
    End If
End Function

Private Function IndexOfString(anElement As String, _
        r_colNew As Collection) As Long
    Dim lngIndex As Long
    For lngIndex = 1 To r_colNew.Count
        If r_colNew(lngIndex) = anElement Then
            IndexOfString = lngIndex
            Exit Function
        End If
    Next
    IndexOfString = 0
End Function

Public Function AsDelimitedString(col As Collection, _
            Optional singleQuote As Boolean = True, _
            Optional delimiter As String = ",", _
            Optional enclosedBy As String = "()") As String
    Dim str As String, i As Integer
    str = Left(enclosedBy, 1)
    For i = 1 To col.Count
        If singleQuote Then
            str = str & EncloseWithSingleQuote(col.Item(i))
        Else
            str = str & col.Item(i)
        End If
        If Not (i = col.Count) Then
            str = str & delimiter
        End If
    Next
    AsDelimitedString = str & Right(enclosedBy, 1)

End Function

Public Function EncloseWithSingleQuote(ByVal v_data As Variant) As String
    EncloseWithSingleQuote = "'" & CStr(v_data) & "'"

End Function

Public Function CollectionPosition(anObject As Object, aCol As Collection) As Integer
    Dim i As Integer
    Dim pos As Integer
    Dim enumObject As Object
    
    For i = 1 To aCol.Count
        Set enumObject = aCol.Item(i)
        If enumObject Is anObject Then
            CollectionPosition = i
            Exit Function
        End If
    Next
    
End Function

Public Function IsStringInList( _
        ByVal v_strValue As String, _
        ByVal v_strList As Variant) As Boolean
'is v_strValue equal to any of values in v_strList
'v_strList should be a string array or variant string array.
    
    Dim i As Integer
    For i = LBound(v_strList) To UBound(v_strList)
        If UCase(v_strValue) = UCase(v_strList(i)) Then
            IsStringInList = True
            Exit Function
        End If
    Next

End Function


Public Function IsStringInList2( _
        ByVal v_strValue As String, _
        ByVal v_strList As Variant, _
        ByVal v_intNumofEntries As Integer, _
        ByRef r_intPosition As Integer) As Boolean
'is v_strValue equal to any of values in v_strList
'v_strList should be a string array or variant string array.
    
    Dim i As Integer
    For i = 0 To v_intNumofEntries - 1
        If UCase(v_strValue) = UCase(v_strList(i, 0)) Then
            IsStringInList2 = True
            r_intPosition = i
            Exit Function
        End If
    Next

End Function


