VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cuADO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const MODULENAME = "utilADOHelper.cuADOCommand"
Private m_ADOConnection As ADODB.Connection
Private m_ADOCommand As ADODB.Command

Public Property Set Connection(ByVal v_ADOConnection As ADODB.Connection)
    Set m_ADOConnection = v_ADOConnection
End Property
Public Property Get Connection() As ADODB.Connection
    Set Connection = m_ADOConnection
End Property

Public Property Set Command(ByVal v_ADOCommand As ADODB.Command)
    Set m_ADOCommand = v_ADOCommand
End Property
Public Property Get Command() As ADODB.Command
    Set Command = m_ADOCommand
End Property

Public Function CreateConnection() As Integer
    Set m_ADOConnection = OpenConnection()
End Function

Public Function CreateCommand() As Integer
    Set m_ADOCommand = New ADODB.Command
End Function

Public Function Destroy() As Integer
    DestroyCommand
    DestroyConnection
End Function

Public Function DestroyConnection() As Integer
    If Not m_ADOConnection Is Nothing Then
        If m_ADOConnection.State = adStateOpen Then
            m_ADOConnection.Close
        End If
        Set m_ADOConnection = Nothing
    End If
End Function

Public Function DestroyCommand() As Integer
    If Not m_ADOCommand Is Nothing Then
        If Not m_ADOCommand.ActiveConnection Is Nothing Then
            If m_ADOCommand.ActiveConnection.State = adStateOpen Then
                m_ADOCommand.ActiveConnection.Close
            End If
            Set m_ADOCommand.ActiveConnection = Nothing
        End If
        Set m_ADOCommand = Nothing
    End If
End Function

Public Function CreateCharacterParameter(ByVal ParameterName As String, _
                                        ByVal DataType As ADODB.DataTypeEnum, _
                                        ByVal Length As Long, _
                                        ByVal Value As Variant) As Integer
On Error GoTo eh
    Dim ladParm As ADODB.Parameter

    If m_ADOCommand Is Nothing Then
        Set m_ADOCommand = New ADODB.Command
    End If

    Set ladParm = m_ADOCommand.CreateParameter(ParameterName _
                         , DataType _
                         , adParamInput _
                         , Length _
                         , Value)
    
    m_ADOCommand.Parameters.Append ladParm
    CreateCharacterParameter = 0
    Exit Function
eh:
    Set ladParm = Nothing
    TrsRaiseError MODULENAME, "CreateCharacterParameter", , , , DescArgsFrom("ParameterName = " & ParameterName), True
End Function

Public Function CreateNumericParameter(ByVal ParameterName As String, _
                                        ByVal DataType As ADODB.DataTypeEnum, _
                                        ByVal Length As Long, _
                                        ByVal nScale As Long, _
                                        ByVal Value As Variant) As Integer
On Error GoTo eh
    Dim ladParm As ADODB.Parameter

    If m_ADOCommand Is Nothing Then
        Set m_ADOCommand = New ADODB.Command
    End If

    Set ladParm = m_ADOCommand.CreateParameter(ParameterName _
                             , DataType _
                             , adParamInput _
                             , _
                             , Value)
'                If sDataType = "DECIMAL" Then
    ladParm.Precision = Length
    ladParm.NumericScale = nScale
'                End If

    m_ADOCommand.Parameters.Append ladParm

    CreateNumericParameter = 0
    Exit Function
eh:
    Set ladParm = Nothing
    TrsRaiseError MODULENAME, "CreateNumericParameter", , , , DescArgsFrom("ParameterName = " & ParameterName), True
End Function

Public Function DeriveSQL(Optional ByVal v_ADOCommand As ADODB.Command) As String
    Dim SQL As String, SQLA As String, SQLB As String
    Dim dbParm As ADODB.Parameter
    Dim ParmValue As String, ParameterMarkerIndex As Integer
    
    If v_ADOCommand Is Nothing Then
        Set v_ADOCommand = Command
    End If
    
    If Not v_ADOCommand Is Nothing Then
        SQL = v_ADOCommand.CommandText
        
        ParameterMarkerIndex = 1
        For Each dbParm In v_ADOCommand.Parameters
            ParmValue = ParmToString(dbParm)
            
            ParameterMarkerIndex = InStr(ParameterMarkerIndex, SQL, "?")
            If ParameterMarkerIndex > 1 Then
                'SQL = SQLA & Replace(SQL, "?", ParmValue, ParameterMarkerIndex, 1)
                SQLA = Left(SQL, ParameterMarkerIndex - 1)
                SQLB = Mid(SQL, ParameterMarkerIndex)
                SQL = SQLA & Replace(SQLB, "?", ParmValue, 1, 1)
                
                ParameterMarkerIndex = ParameterMarkerIndex + Len(ParmValue)
            End If
        Next
    End If
    
    DeriveSQL = SQL
End Function

Private Function ParmToString(ByVal dbParm As ADODB.Parameter) As String
    Dim ParmValue As Variant
    Dim strValue As String

    ParmValue = dbParm.Value
    
    Select Case dbParm.Type
    Case ADODB.DataTypeEnum.adChar, ADODB.DataTypeEnum.adVarChar ', ADODB.DataTypeEnum.adWChar, ADODB.DataTypeEnum.adVarWChar
        If Not IsNull(ParmValue) Then
            strValue = "'" & Replace(CStr(ParmValue), "'", "''") & "'"
        Else
            strValue = "''"
        End If
    Case ADODB.DataTypeEnum.adDBTimeStamp '2008-08-28-11.47.29.744946
        If Not IsNull(ParmValue) Then
            strValue = "'" & Format(CDate(CStr(ParmValue)), "YYYY-MM-DD-HH.NN.SS") & "." & GetMilliseconds(ParmValue) & "'"
        Else
            strValue = "NULL"
        End If
    Case ADODB.DataTypeEnum.adDate, ADODB.DataTypeEnum.adDBDate
        If Not IsNull(ParmValue) Then
            strValue = "'" & Format(CDate(CStr(ParmValue)), "MM/DD/YYYY") & "'"
        Else
            strValue = "NULL"
        End If
    Case ADODB.DataTypeEnum.adDBTime '12.35.58
        If Not IsNull(ParmValue) Then
            strValue = "'" & Format(CDate(CStr(ParmValue)), "HH.NN.SS") & "'"
        Else
            strValue = "NULL"
        End If
'        Case ADODB.DataTypeEnum.adBinary, ADODB.DataTypeEnum.adVarBinary
'            If Not IsNull(ParmValue) Then
'                strValue = "'" & Conversion.ConvertBytesToString(CType(ParmValue, Byte())) & "'"
'            Else
'                strValue = ""
'            End If
    Case Else
        If Not IsNull(ParmValue) Then
            strValue = CStr(ParmValue)
        Else
            strValue = "NULL"
        End If
    End Select
    
    ParmToString = strValue
End Function

Private Function GetMilliseconds(ByVal varDateTime As Variant) As Long

    '  The Decimal datatype can store decimal values exactly.
    '  Variables cannot be directly declared as Decimal, so
    '     create a Variant then use CDec( ) to convert to Decimal.
      Dim decConversionFactor As Variant
      Dim decTime As Variant

    'K is used to convert a VB time unit back to seconds
    'K = 86400000 milliseconds per day
      decConversionFactor = CDec(86400000)

    'Store the DateTime value in an exact decimal value called decTime
      decTime = CDec(varDateTime)

    'Make sure the date/time value is positive
      decTime = Abs(decTime)

    'Get rid of the date (whole number), leaving time (decimal)
      decTime = decTime - Int(decTime)

    'Convert to time to seconds
      decTime = (decTime * decConversionFactor)

    'Return the milliseconds
      GetMilliseconds = decTime Mod 1000

End Function
   
Private Sub Class_Terminate()
    Destroy
End Sub
