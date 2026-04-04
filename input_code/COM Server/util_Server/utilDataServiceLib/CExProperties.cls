VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CExProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Only use index access for enumeration.  Index may change if adding an
'association that already existed.  This can be avoided by using an association
'object internally or add using before or after.

Private mcolContents As Collection

Public Sub Add(ByVal v_strKey As String, ByVal v_varValue As Variant)
    Dim l_strKey As String

    l_strKey = Trim(v_strKey)
    If Includes(l_strKey) Then
        Remove l_strKey
        Contents.Add Array(v_strKey, v_varValue), l_strKey
    Else
        Contents.Add Array(v_strKey, v_varValue), l_strKey
    End If
    
End Sub

Public Function Item(ByVal v_varIndex As Variant) As Variant
Attribute Item.VB_UserMemId = 0
    
    If Includes(v_varIndex) Then
        Item = Contents.Item(v_varIndex)(1)
    Else
        Item = Empty
    End If
    
End Function

Public Function Association(ByVal v_varIndex As Variant) As Variant

    Association = Contents.Item(v_varIndex)
    
End Function

Public Function Remove(ByVal v_varIndex As Variant)

    Contents.Remove v_varIndex
    
End Function

Public Function Includes(ByVal v_varIndex As Variant) As Boolean
    
On Error GoTo eh
    Contents.Item v_varIndex
    Includes = True
    Exit Function
    
eh:
    Includes = False
End Function

Public Property Get Count() As Long
    Count = mcolContents.Count
End Property

Public Function Contents() As Collection
    Set Contents = mcolContents
End Function

Public Function GetString() As String
'Answer all the property names and values
    Dim l_strReturn As String
    Dim l_varData As Variant

    For Each l_varData In Contents
        l_strReturn = l_strReturn & l_varData(0) & "=" & l_varData(1) & vbCrLf
    Next
    GetString = l_strReturn
    
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
   Set NewEnum = mcolContents.[_NewEnum]
End Function

Private Sub Class_Initialize()
    Set mcolContents = New Collection
End Sub
