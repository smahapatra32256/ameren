VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CExRecordset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const CMODULENAME As String = "CExRecordset"
'Const CIGNOREFIELD As String = "IGNOREFIELD"
Private Const CDEFAULTOWNER As String = "METHDBA"

'Private Enum TrSetCurrentTimestampEnum
'    TrSetTSDisable = 0
'    TrSetTSAlways = 1
'    TrSetTSInsertOnly = 2
'    TrSetTSInsertAndEmpty = 3
'    TrSetTSTouch = 4
'End Enum

Private mcolFields As Collection
Private mcolProperties As CExProperties
Private mvarPrimaryKeys As Variant
Private mblnOptimisticLocking As Boolean

Public Function Properties() As CExProperties
    Set Properties = mcolProperties
End Function


Public Sub SetTableName(ByVal v_strData As String)

    Properties.Add "TABLENAME", v_strData
    
End Sub

Public Sub SetTableOwner(ByVal v_strData As String)

    Properties.Add "TABLEOWNER", v_strData
    
End Sub

Public Function GetTableName() As String
    Dim l_varName As Variant

    l_varName = Properties.Item("TABLENAME")
    If Not IsEmpty(l_varName) Then
        GetTableName = l_varName
    End If

End Function

Public Function GetTableOwner() As String
    Dim l_varName As Variant

    l_varName = Properties.Item("TABLEOWNER")
    If Not IsEmpty(l_varName) Then
        GetTableOwner = l_varName
    End If

End Function


Public Function FullTableName() As String
    Dim l_varOwner As Variant
    Dim l_varName As Variant
    Dim l_strReturn As String

    l_varOwner = Properties.Item("TABLEOWNER")
    l_varName = Properties.Item("TABLENAME")
    
    If IsEmpty(l_varOwner) Then
        l_strReturn = CDEFAULTOWNER & "."
    Else
        l_strReturn = l_varOwner & "."
    End If
    
    If Not IsEmpty(l_varName) Then
        l_strReturn = l_strReturn & l_varName
    End If
    FullTableName = l_strReturn

End Function


Public Sub SetFullTableName(ByVal v_strData As String)
    Dim l_varOwnerName As Variant
    Dim l_intSize As Integer
    Dim l_objCommonLib As utilTRISCommonLib.cuContainer

    l_varOwnerName = Split(v_strData, ".")
    Set l_objCommonLib = New cuContainer
    l_intSize = l_objCommonLib.ArraySize(l_varOwnerName)
    Set l_objCommonLib = Nothing
    Select Case l_intSize
        Case 0
            Err.Raise vbObjectError + 3000, CMODULENAME, "Table Name is empty string"
        Case 1
            'assume table name only, default owner
            SetTableOwner CDEFAULTOWNER
            SetTableName l_varOwnerName(LBound(l_varOwnerName))
        Case 2
            SetTableOwner l_varOwnerName(LBound(l_varOwnerName))
            SetTableName l_varOwnerName(UBound(l_varOwnerName))
        Case Else
            Err.Raise vbObjectError + 3001, CMODULENAME, "Table Name has more than 2 levels"
    End Select

End Sub

Public Function PrimaryKeys() As Variant

    If Not IsPrimaryKeyCacheBuilt Then
        BuildPrimaryKeyCache
    End If
    PrimaryKeys = mvarPrimaryKeys

End Function

Public Function ConcurrencyKeys() As Variant

    Dim l_objField As CExField
    Dim l_colKeys As Collection
    Dim l_objCommonLib As utilTRISCommonLib.cuContainer

    If Not mblnOptimisticLocking Then
        ConcurrencyKeys = Empty
        Exit Function
    End If
    
    Set l_colKeys = New Collection
    For Each l_objField In Fields
        If l_objField.IsConcurrencyKey Then
            l_colKeys.Add l_objField.Name
        End If
    Next
    Set l_objCommonLib = New cuContainer
    ConcurrencyKeys = l_objCommonLib.ArrayFromCollection(l_colKeys)
    Set l_objCommonLib = Nothing
    Set l_colKeys = Nothing

End Function


Public Function UseClientConcurrencyKey(ByVal v_strFieldName As String) As Boolean
    'Answer the UseClientConcurrencyKey property value for a field name.
    'Answer False if the property does not exist for a field.
    Dim l_objField As CExField
    
    Set l_objField = Field(v_strFieldName)
    If Not l_objField Is Nothing Then
        UseClientConcurrencyKey = l_objField.UseClientConcurrencyKey
    Else
        UseClientConcurrencyKey = False
    End If

End Function


Public Function Fields() As Collection

    Set Fields = mcolFields
    
End Function

Public Function Field(ByVal v_strFieldName As String) As CExField
    Dim l_objField As CExField

    For Each l_objField In Fields
        If l_objField.Name = v_strFieldName Then
            Set Field = l_objField
            Exit For
        End If
    Next
    
End Function

Public Sub AddField(ByVal v_strFieldName As String, _
                ByVal v_varPropertyNames As Variant, _
                ByVal v_varPropertyValues As Variant)
    Dim l_objField As CExField
    Dim l_objProperties As CExProperties
    Dim i As Integer

    Set l_objField = AddFieldBasic(v_strFieldName)
    Set l_objProperties = l_objField.Properties
    For i = LBound(v_varPropertyNames) To UBound(v_varPropertyNames)
        l_objProperties.Add v_varPropertyNames(i), v_varPropertyValues(i)
        If v_varPropertyNames(i) = "SETCURRENTTIMESTAMP" Then
            If v_varPropertyValues(i) = TrSetTSAlways Then
                l_objProperties.Add "IGNORESOLEUPDATE", True
            ElseIf v_varPropertyValues(i) = TrSetTSTouch Then
                l_objProperties.Add "IGNORESOLEUPDATE", False
            End If
        End If
        If v_varPropertyNames(i) = "PRIMARYKEY" Then
            InitializePrimaryKeyCache
        End If
        If v_varPropertyNames(i) = "CONCURRENCYKEY" Then
            mblnOptimisticLocking = True
            l_objProperties.Add "SETCURRENTTIMESTAMP", TrSetTSAlways
            l_objProperties.Add "IGNORESOLEUPDATE", True
        End If
        'Move this to properties container
        If v_varPropertyNames(i) = "USECLIENTCONCURRENCYKEY" Then
            If v_varPropertyValues(i) Then
                l_objProperties.Add "CONCURRENCYKEY", True
            End If
        End If
    Next

End Sub


Public Sub AddFields( _
                ByVal v_strPropertyName As String, _
                ByVal v_varPropertyValue As Variant, _
                ByVal v_varFieldNames As Variant)
    Dim l_objField As CExField
    Dim l_objProperties As CExProperties
    Dim i As Integer

    For i = LBound(v_varFieldNames) To UBound(v_varFieldNames)
        Set l_objField = AddFieldBasic(v_varFieldNames(i))
        Set l_objProperties = l_objField.Properties
        l_objProperties.Add v_strPropertyName, v_varPropertyValue
        If v_strPropertyName = "SETCURRENTTIMESTAMP" Then
            If v_varPropertyValue = TrSetTSAlways Then
                l_objProperties.Add "IGNORESOLEUPDATE", True
            ElseIf v_varPropertyValue = TrSetTSTouch Then
                l_objProperties.Add "IGNORESOLEUPDATE", False
            End If
        End If

        If v_strPropertyName = "USECLIENTCONCURRENCYKEY" Then
            If v_varPropertyValue Then
                l_objProperties.Add "CONCURRENCYKEY", True
            End If
        End If
        If v_strPropertyName = "CONCURRENCYKEY" Then
            l_objProperties.Add "SETCURRENTTIMESTAMP", TrSetTSAlways
        End If
    Next
    If v_strPropertyName = "PRIMARYKEY" Then
        InitializePrimaryKeyCache
    End If
    If v_strPropertyName = "CONCURRENCYKEY" Then
        mblnOptimisticLocking = True
    End If

End Sub

Private Function AddFieldBasic(ByVal v_strFieldName As String) As CExField
    Dim l_objField As CExField
    
    Set l_objField = Field(v_strFieldName)
    If l_objField Is Nothing Then
        Set l_objField = New CExField
        l_objField.Name = v_strFieldName
        Fields.Add l_objField
    End If
    Set AddFieldBasic = l_objField
    
End Function

Private Function IsPrimaryKeyCacheBuilt()

    IsPrimaryKeyCacheBuilt = Not IsEmpty(mvarPrimaryKeys)

End Function

Private Sub InitializePrimaryKeyCache()

    mvarPrimaryKeys = Empty

End Sub

Private Sub BuildPrimaryKeyCache()
    Dim l_objField As CExField
    Dim l_colPrimKeys As Collection
    Dim l_objCommonLib As utilTRISCommonLib.cuContainer


'Debug.Print "Building primary key cache"
    Set l_colPrimKeys = New Collection
    For Each l_objField In Fields
        If l_objField.IsPrimaryKey Then
            l_colPrimKeys.Add l_objField.Name
        End If
    Next
    Set l_objCommonLib = New cuContainer
    mvarPrimaryKeys = l_objCommonLib.ArrayFromCollection(l_colPrimKeys)
    Set l_objCommonLib = Nothing
    Set l_colPrimKeys = Nothing

End Sub


Private Sub Class_Initialize()
    Set mcolProperties = New CExProperties
    Set mcolFields = New Collection
    mblnOptimisticLocking = False
End Sub
