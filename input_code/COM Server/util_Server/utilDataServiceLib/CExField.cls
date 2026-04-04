VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CExField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mstrName As String
Private mcolProperties As CExProperties

Public Property Get Name() As String

    Name = mstrName

End Property

Public Property Let Name(ByVal v_strData As String)

    mstrName = Trim(v_strData)
    
End Property

Public Function Properties() As CExProperties

    Set Properties = mcolProperties

End Function

Public Function GetString() As String

    GetString = Properties.GetString
    
End Function

Public Function IsPrimaryKey() As Boolean
    Dim l_varValue As Variant

    l_varValue = Properties.Item("PRIMARYKEY")
    If Not IsEmpty(l_varValue) Then
        IsPrimaryKey = l_varValue
    Else
        IsPrimaryKey = False
    End If
    
End Function


Public Function IsConcurrencyKey() As Boolean
    Dim l_varValue As Variant

    l_varValue = Properties.Item("CONCURRENCYKEY")
    If Not IsEmpty(l_varValue) Then
        IsConcurrencyKey = l_varValue
    Else
        IsConcurrencyKey = False
    End If
    
End Function


Public Function UseClientConcurrencyKey() As Boolean
    'Answer the UseClientConcurrencyKey property value for this field.
    'Answer False if the property does not exist for this field.
    'Ignore this property and answer false if not a concurrency key field.
    Dim l_varValue As Variant

    If Me.IsConcurrencyKey Then
        l_varValue = Properties.Item("USECLIENTCONCURRENCYKEY")
        If Not IsEmpty(l_varValue) Then
            UseClientConcurrencyKey = l_varValue
        Else
            UseClientConcurrencyKey = False
        End If
    Else
        UseClientConcurrencyKey = False
    End If

End Function


Private Sub Class_Initialize()
    Set mcolProperties = New CExProperties
End Sub
