VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cuSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'******************************************************
' Purpose:
'   Collect distinct objects, i.e. no duplicates set.
'   This could be more efficient if used keys when
'   adding, but keep more flexible for now.
'
'   Make this collection sorted in ascending order.
'   Make this more flexible later.
'******************************************************
Private Const ElementNotIncluded As Long = 0
Private Const NoElementHigher As Long = 65535
Private mcolContents As Collection

Public Function Item(ndx As Long) As Variant
'index must be a long.
'answer an object or nothing
'raise error if not found
    On Error GoTo notfound
    If IsObject(mcolContents(ndx)) Then
        Set Item = mcolContents(ndx)
    Else
        Item = mcolContents(ndx)
    End If
    Exit Function
notfound:
    Set Item = Nothing
End Function

Public Property Get Count() As Long
    Count = mcolContents.Count
End Property

Public Function Contents() As Collection
'Answer a copy of my contents
    Dim l_colCopy As Collection
    Dim l_varItem As Variant
    
    Set l_colCopy = New Collection
    For Each l_varItem In mcolContents
        l_colCopy.Add l_varItem
    Next
    Set Contents = l_colCopy

End Function

Public Sub Add(anElement As Variant)
    Dim lowIndex As Long
    
    If Not Includes(anElement) Then
    ' find first element higher than anElement
        lowIndex = FirstHigherElement(anElement)
        If mcolContents.Count = 0 Then
            mcolContents.Add anElement
        ElseIf lowIndex <= mcolContents.Count Then
            mcolContents.Add anElement, , lowIndex
        Else
            mcolContents.Add anElement
        End If
    End If
End Sub

Public Sub Remove(ndx As Long)
'Must be a long, 1 to count, until use keys.
    mcolContents.Remove ndx
End Sub

Public Sub RemoveObject(anElement As Variant)
' raise an error when not found
' Find the index of anItem then remove
    Dim lngPosition As Long
    lngPosition = IndexOf(anElement)
    If Not lngPosition = ElementNotIncluded Then
       Remove lngPosition
    End If
End Sub

Public Function Includes(anElement As Variant) As Boolean
    Dim lngIndex As Long
    lngIndex = IndexOf(anElement)
    If lngIndex = ElementNotIncluded Then
        Includes = False
    Else
        Includes = True
    End If
End Function

Private Function FirstHigherElement(anElement As Variant) As Long
' Answer the element index to be used as Add/Before argument
    If mcolContents.Count = 0 Then
        FirstHigherElement = 1
        Exit Function
    End If
    
' this does not handle objects yet
    If IsObject(anElement) Then
        FirstHigherElement = NoElementHigher
        Exit Function
    End If

    Dim lngIndex As Long
    For lngIndex = 1 To Me.Count
        If mcolContents(lngIndex) >= anElement Then
            FirstHigherElement = lngIndex
            Exit Function
        End If
    Next

    FirstHigherElement = NoElementHigher

End Function

Public Function IndexOf(anElement As Variant) As Long
'mcolContents is a 1-based collection
'Ensure Empty is only equal to Empty, same for Null.
    Dim lngIndex As Long
    For lngIndex = 1 To Me.Count
        If IsObject(anElement) Then
            If IsObject(mcolContents(lngIndex)) Then
                If mcolContents(lngIndex) Is anElement Then
                    IndexOf = lngIndex
                    Exit Function
                End If
            End If
        Else
            If Not IsObject(mcolContents(lngIndex)) Then
                If IsNull(mcolContents(lngIndex)) Or IsNull(anElement) Then
                    If IsNull(mcolContents(lngIndex)) And IsNull(anElement) Then
                        IndexOf = lngIndex
                        Exit Function
                    End If
                ElseIf IsEmpty(mcolContents(lngIndex)) Or IsEmpty(anElement) Then
                    If IsEmpty(mcolContents(lngIndex)) And IsEmpty(anElement) Then
                        IndexOf = lngIndex
                        Exit Function
                    End If
                ElseIf mcolContents(lngIndex) = anElement Then
                    IndexOf = lngIndex
                    Exit Function
                End If
            End If
        End If
    Next
    IndexOf = ElementNotIncluded
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mcolContents.[_NewEnum]
End Function

Private Sub Class_Initialize()
    Set mcolContents = New Collection
End Sub

