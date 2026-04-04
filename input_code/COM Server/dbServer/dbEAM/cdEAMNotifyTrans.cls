VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "cdEAMNotifyTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Project:   Post TRIS 5.5 Employee Absence Management
'Author:    Prem Chander
'Date:      October 2005
'Description:
Option Explicit
Private Const MODULENAME As String = "dbEAM.cdEAMNotifyTrans"
Public Function dbSaveNotify(ByVal vadrsPreference As ADODB.Recordset _
                , ByVal vlTRISUserNumber As Long _
                ) As Boolean
On Error GoTo dbSaveNotify_EH
Dim lsSQL As String, ladcn As ADODB.Connection

    dbSaveNotify = False

    lsSQL = ""

    If (RecordsExist(vadrsPreference)) Then
        With vadrsPreference
            .MoveFirst
            .Find "RECORD_TYPE = 'CNTCT'"

            If Not (.EOF) Then
                Select Case (!RECORD_STATUS)
                    Case "I"
                        lsSQL = lsSQL & _
                            " INSERT INTO METHDBA.TTR_RG_GROUP" & _
                            " (RG_CD, RG_SUB_CD" & _
                            " , BUS_TYPE_CD, RG_VALUE_CD" & _
                            " , SEQ_NUM, APPLY_IND" & _
                            " , CREATED_EMP_NUM, LAST_MOD_EMP_NUM" & _
                            " , CREATED_TIMESTAMP, LAST_MOD_TIMESTAMP) " & _
                            " VALUES" & _
                            " (303, ''" & _
                            " , 40, '" & Trim(!USER_NUM_AS_STR) & ":" & !NOTIFY_IND & ":" & CStr(!ELAPSED_DAYS) & "'" & _
                            " , 1, 1" & _
                            " , " & CStr(vlTRISUserNumber) & ", " & CStr(vlTRISUserNumber) & _
                            " , CURRENT TIMESTAMP, CURRENT TIMESTAMP);"

                    Case "C"
                        lsSQL = lsSQL & _
                            " UPDATE METHDBA.TTR_RG_GROUP" & _
                            " SET   RG_VALUE_CD = '" & Trim(!USER_NUM_AS_STR) & ":" & !NOTIFY_IND & ":" & CStr(!ELAPSED_DAYS) & "'" & _
                            "       , LAST_MOD_EMP_NUM = " & CStr(vlTRISUserNumber) & _
                            "       , LAST_MOD_TIMESTAMP = CURRENT TIMESTAMP" & _
                            " WHERE RG_CD = 303" & _
                            " AND   BUS_TYPE_CD = 40" & _
                            " AND   SUBSTR(RG_VALUE_CD, 1, 9) = '" & Trim(!USER_NUM_AS_STR) & "';"
                End Select
            End If

            .MoveFirst
            .Find "RECORD_TYPE = 'RTW'"

            If Not (.EOF) Then
                Select Case (!RECORD_STATUS)
                    Case "I"
                        lsSQL = lsSQL & _
                            " INSERT INTO METHDBA.TTR_RG_GROUP" & _
                            " (RG_CD, RG_SUB_CD" & _
                            " , BUS_TYPE_CD, RG_VALUE_CD" & _
                            " , SEQ_NUM, APPLY_IND" & _
                            " , CREATED_EMP_NUM, LAST_MOD_EMP_NUM" & _
                            " , CREATED_TIMESTAMP, LAST_MOD_TIMESTAMP) " & _
                            " VALUES" & _
                            " (307, ''" & _
                            " , 41, '" & Trim(!USER_NUM_AS_STR) & ":" & !NOTIFY_IND & "'" & _
                            " , 1, 1" & _
                            " , " & CStr(vlTRISUserNumber) & ", " & CStr(vlTRISUserNumber) & _
                            " , CURRENT TIMESTAMP, CURRENT TIMESTAMP);"

                    Case "C"
                        lsSQL = lsSQL & _
                            " UPDATE METHDBA.TTR_RG_GROUP" & _
                            " SET   RG_VALUE_CD = '" & Trim(!USER_NUM_AS_STR) & ":" & !NOTIFY_IND & "'" & _
                            "       , LAST_MOD_EMP_NUM = " & CStr(vlTRISUserNumber) & _
                            "       , LAST_MOD_TIMESTAMP = CURRENT TIMESTAMP" & _
                            " WHERE RG_CD = 307" & _
                            " AND   BUS_TYPE_CD = 41" & _
                            " AND   SUBSTR(RG_VALUE_CD, 1, 9) = '" & Trim(!USER_NUM_AS_STR) & "';"
                End Select
            End If
        End With
    End If

    If (Trim(lsSQL) <> "") Then
        Set ladcn = OpenConnection

        ExecuteOTHERSQL lsSQL _
            , ladcn _
            , adCmdText

        If Not (ladcn Is Nothing) Then
            ladcn.Close
        End If
        Set ladcn = Nothing
    End If

    dbSaveNotify = True

    Exit Function
dbSaveNotify_EH:
    GetObjectContext.SetAbort

    Set ladcn = Nothing

    Select Case (Err.Number)
        Case trErrutOpenConnection
            TrsRaiseError MODULENAME _
                , "dbSaveNotify"
        Case trErrdb_ExecuteSQL_F
            TrsRaiseError MODULENAME _
                , "dbSaveNotify" _
                , trErrdb_ExecuteSQL_F _
                , , _
                , DescArgsFrom(lsSQL) _
                , True
        Case Else
            TrsRaiseError MODULENAME _
                , "dbSaveNotify" _
                , trErrdbSaveEAM _
                , , _
                , DescArgsFrom(lsSQL) _
                , True
    End Select

End Function
