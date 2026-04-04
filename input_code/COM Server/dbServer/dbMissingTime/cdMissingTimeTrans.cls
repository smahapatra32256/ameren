VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "cdMissingTimeTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Project:   TRIS Time to Pay
'Author:    Prem Chander
'Date:      August 2006
'Description: This module will insert schedules in the Brief Sch Table

Option Explicit

Implements IObjectConstruct
Dim m_ConstructString As String

'#Const DEBUGGING = 0 '1 - ON;0 - OFF

Private Const MODULENAME As String = "dbMissingTime.cdMissingTimeTrans"

Dim mlOffset As Long
Dim m_colConstructStringProperty As Collection
Dim DEBUGGING As Integer '1 - ON;0 - OFF

''test
'Private Sub Class_Initialize()
'    DEBUGGING = 0
'End Sub

Private Sub IObjectConstruct_Construct(ByVal pCtorObj As Object)
    Dim mConstructString As IObjectConstructString
    
    Set mConstructString = pCtorObj
    
    m_ConstructString = mConstructString.ConstructString
    
    Set m_colConstructStringProperty = ParseContructString(m_ConstructString)
    
    PostConstruct
    App.LogEvent MODULENAME & " ObjectConstructString = " & m_ConstructString, vbLogEventTypeInformation
End Sub

Private Sub PostConstruct()
    If DoesPropertyExists(m_colConstructStringProperty, "DEBUGGING") Then
        If IsNumeric(m_colConstructStringProperty.Item("DEBUGGING")) Then
            DEBUGGING = CInt(m_colConstructStringProperty.Item("DEBUGGING"))
        End If
    End If
End Sub
    
Private Function DoesPropertyExists(ByVal colPropertyValue As Collection, ByVal Property As String) As Boolean
    Dim oValue As Variant

    On Error GoTo eh
    
    oValue = colPropertyValue.Item(Property)
    
    DoesPropertyExists = True
    Exit Function
eh:
    DoesPropertyExists = False
End Function

Private Function ParseContructString(ByVal ConstructString As String) As Collection
    Dim colConstructStringProperty As Collection
    Dim PropertyValueArray() As String
    Dim PropertyValue As Variant
    Dim PropertyValueArray2() As String

    Set colConstructStringProperty = New Collection
    PropertyValueArray = Split(ConstructString, ";")
    
    For Each PropertyValue In PropertyValueArray
        PropertyValueArray2 = Split(PropertyValue, "=")
        If UBound(PropertyValueArray2) >= 1 Then
            colConstructStringProperty.Add PropertyValueArray2(1), PropertyValueArray2(0)
        End If
    Next
       
    Set ParseContructString = colConstructStringProperty
End Function

Public Sub InsertEmployeeBriefSchedule(ByVal vadrsBriefSchedule As ADODB.Recordset)
On Error GoTo InsertEmployeeBriefSchedule_EH
'Note: By preallocating string size (through Space), I am able to concatenate
'the string faster than normal concatenation. Maximum VB String Length (@ 4bytes)
'= 65,535 but I have given soom room
Dim lsSQL As String, llIndex As Long
Dim ladcn As ADODB.Connection
Dim llExecuteCounter As Long

    If Not (vadrsBriefSchedule Is Nothing) Then
        With vadrsBriefSchedule
            If Not (.EOF And .BOF) Then
                .MoveFirst

                lsSQL = ""
                mlOffset = 0

                Set ladcn = OpenConnection

                If (DEBUGGING = 1) Then
                    llExecuteCounter = 0
                End If

                For llIndex = 1 To .RecordCount
                    FastConcatString lsSQL, _
                        " INSERT INTO METHDBA.TTR_EMPL_SCH" & _
                        " VALUES " & _
                        "(" & CStr(!EMP_NUM) & _
                        ", '" & CStr(!CHARGE_DATE) & _
                        "', " & CStr(!HOURS) & _
                        ", '" & Trim(!OT_IND) & "');" _
                        , mlOffset _
                        , 64588
                    If (Len(Trim(lsSQL)) * 4 > 64860) Or _
                        (llIndex = .RecordCount) Then
                        mlOffset = 0

                        ExecuteOTHERSQL Trim(lsSQL), ladcn _
                            , adCmdText

                        If (DEBUGGING = 1) Then
                            llExecuteCounter = llExecuteCounter + 1
                        End If

                        lsSQL = ""
                    End If

                    .MoveNext
                Next llIndex

                Set ladcn = Nothing
            End If
        End With
    End If

    If (DEBUGGING = 1) Then
        Debug.Print CStr(llExecuteCounter)
    End If

    GetObjectContext.SetComplete
    Exit Sub
InsertEmployeeBriefSchedule_EH:
    GetObjectContext.SetAbort

    Set ladcn = Nothing

    Select Case (Err.Number)
        Case trErrutOpenConnection
            TrsRaiseError MODULENAME _
                , "InsertEmployeeBriefSchedule"
        Case trErrdb_ExecuteSQL_F
            TrsRaiseError MODULENAME _
                , "InsertEmployeeBriefSchedule" _
                , trErrdb_ExecuteSQL_F _
                , , _
                , DescArgsFrom(lsSQL) _
                , True
        Case Else
            TrsRaiseError MODULENAME _
                , "InsertEmployeeBriefSchedule" _
                , Err.Number _
                , , _
                , DescArgsFrom(lsSQL) _
                , True
    End Select

End Sub
Public Sub DeleteEmployeeBriefSchedule()
On Error GoTo DeleteEmployeeBriefSchedule_EH
Dim ladcn As ADODB.Connection

    Set ladcn = OpenConnection
    ExecuteOTHERSQL "DELETE FROM METHDBA.TTR_EMPL_SCH" _
        , ladcn _
        , adCmdText
    Set ladcn = Nothing

    GetObjectContext.SetComplete
    Exit Sub
DeleteEmployeeBriefSchedule_EH:
    GetObjectContext.SetAbort

    Set ladcn = Nothing

    Select Case (Err.Number)
        Case trErrutOpenConnection
            TrsRaiseError MODULENAME _
                , "DeleteEmployeeBriefSchedule"
        Case trErrdb_ExecuteSQL_F
            TrsRaiseError MODULENAME _
                , "DeleteEmployeeBriefSchedule" _
                , trErrdb_ExecuteSQL_F _
                , , _
                , DescArgsFrom("DELETE FROM METHDBA.TTR_EMPL_SCH") _
                , True
        Case Else
            TrsRaiseError MODULENAME _
                , "DeleteEmployeeBriefSchedule" _
                , Err.Number _
                , , _
                , DescArgsFrom("DELETE FROM METHDBA.TTR_EMPL_SCH") _
                , True
    End Select
End Sub
