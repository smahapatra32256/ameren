VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "cbsJCL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'AMEREN Project: TRIS Conversion Phase 3
'October 2004
'Author:    Prem M. Chander
'Advisor:   John R. Stricklin
'Module Description: This class will provide the user the ability
'to retrieve TRIS 16 bit report request information from
'TTR_JCL%% tables. It may also call TR427

'COMMENTS: OCT, 2005 - CHANGE CreateObject
'TO CREATEOBJECT -- PREM

Option Explicit

Private Const MODULENAME = "busJCL.cbsJCL"
Public Function Start427() As Boolean
On Error GoTo Start427_EH
Dim lobjDBJCL As dbJCL.cdJCL

    Start427 = False

    Set lobjDBJCL = CreateObject("dbJCL.cdJCL")
    Start427 = lobjDBJCL.Start427()
    Set lobjDBJCL = Nothing

    Exit Function
Start427_EH:
    Set lobjDBJCL = Nothing

    TrsRaiseError MODULENAME, "Start427" _
        , , , , , True

End Function
Public Function CheckJCLStatus(ByVal vlRequestNumber As Long _
                    , Optional ByVal vsProgramID As String = "TR427" _
                    ) As String
On Error GoTo CheckJCLStatus_EH
Dim lobjDBJCL As dbJCL.cdJCL

    CheckJCLStatus = ""

    Set lobjDBJCL = CreateObject("dbJCL.cdJCL")
    CheckJCLStatus = lobjDBJCL.CheckJCLStatus(vlRequestNumber _
                        , vsProgramID)
    Set lobjDBJCL = Nothing

    Exit Function
CheckJCLStatus_EH:
    Set lobjDBJCL = Nothing

    TrsRaiseError MODULENAME, "CheckJCLStatus" _
        , , , , , True

End Function
Public Function GetListOfPrinters() As ADODB.Recordset
On Error GoTo GetListOfPrinters_EH
Dim lobjDBJCL As dbJCL.cdJCL

    Set GetListOfPrinters = Nothing

    Set lobjDBJCL = CreateObject("dbJCL.cdJCL")
    Set GetListOfPrinters = _
        lobjDBJCL.GetListOfPrinters()
    Set lobjDBJCL = Nothing

    Exit Function
GetListOfPrinters_EH:
    Set lobjDBJCL = Nothing

    TrsRaiseError MODULENAME, "GetListOfPrinters" _
        , , , , , True

End Function
