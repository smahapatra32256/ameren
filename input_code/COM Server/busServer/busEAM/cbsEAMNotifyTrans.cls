VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "cbsEAMNotifyTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Project:   Post TRIS 5.5 Employee Absence Management
'Author:    Prem Chander
'Date:      October 2005
'Description:

Option Explicit

Private Const MODULENAME As String = "busEAM.cbsEAMNotifyTrans"
Public Function SetEAMNotifyPreference(ByVal vadrsPreference As ADODB.Recordset _
                , ByVal vlTRISUserNumber As Long _
                ) As Boolean
On Error GoTo SetEAMNotifyPreference_EH
Dim lobjdb As dbEAM.cdEAMNotifyTrans
Dim llIndex As Long

    SetEAMNotifyPreference = False

    If (RecordsExist(vadrsPreference)) Then
        With vadrsPreference
            .Filter = adFilterNone
            .MoveFirst

            For llIndex = 1 To .RecordCount
                Select Case (Len(Trim(!USER_NUM_AS_STR)))
                    Case 0
                        SetEAMNotifyPreference = True
                        GetObjectContext.SetComplete
                        Exit Function

                    Case 1
                        !USER_NUM_AS_STR = "00000000" & Trim(!USER_NUM_AS_STR)
                        .Update
                    Case 2
                        !USER_NUM_AS_STR = "0000000" & Trim(!USER_NUM_AS_STR)
                        .Update
                    Case 3
                        !USER_NUM_AS_STR = "000000" & Trim(!USER_NUM_AS_STR)
                        .Update
                    Case 4
                        !USER_NUM_AS_STR = "00000" & Trim(!USER_NUM_AS_STR)
                        .Update
                    Case 5
                        !USER_NUM_AS_STR = "0000" & Trim(!USER_NUM_AS_STR)
                        .Update
                    Case 6
                        !USER_NUM_AS_STR = "000" & Trim(!USER_NUM_AS_STR)
                        .Update
                    Case 7
                        !USER_NUM_AS_STR = "00" & Trim(!USER_NUM_AS_STR)
                        .Update
                    Case 8
                        !USER_NUM_AS_STR = "0" & Trim(!USER_NUM_AS_STR)
                        .Update
                End Select

                Select Case (!RECORD_TYPE)
                    Case "CNTCT"
                        Select Case (!NOTIFY_IND)
                            Case "Y", "YES", "1"
                                If (!ELAPSED_DAYS <= 0) Then
                                    !NOTIFY_IND = "N"
                                    .Update
                                End If

                            Case Else
                                !NOTIFY_IND = "N"
                                .Update
                        End Select
                    Case "RTW"
                        Select Case (!NOTIFY_IND)
                            Case "Y", "YES", "1"
                                !NOTIFY_IND = "Y"
                                .Update

                            Case Else
                                !NOTIFY_IND = "N"
                                .Update
                        End Select
                End Select

                .MoveNext
            Next llIndex
        End With
    End If

    Set lobjdb = CreateObject("dbEAM.cdEAMNotifyTrans")
    SetEAMNotifyPreference = lobjdb.dbSaveNotify(vadrsPreference _
                                , vlTRISUserNumber)
    Set lobjdb = Nothing

    GetObjectContext.SetComplete

    Exit Function
SetEAMNotifyPreference_EH:
    GetObjectContext.SetAbort

    Set lobjdb = Nothing

    TrsRaiseError MODULENAME _
        , "SetEAMNotifyPreference" _
        , , , , _
        , True
End Function
