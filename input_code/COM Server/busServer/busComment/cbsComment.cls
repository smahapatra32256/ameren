VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "cbsComment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const m_strCMODULENAME As String = "cbsComment"
Public Function GetComment(ByVal v_aryEmpNum As Variant, _
                                ByVal v_dtStartDate As Date, _
                                ByVal v_dtEndDate As Date) As ADODB.Recordset
                                
    On Error GoTo GetComment_Err
    
    Dim l_objComment As dbComment.cdComment
    Dim l_adrsComment As ADODB.Recordset
    Dim l_strFilter As String
    
    Set l_objComment = CreateObject("dbComment.cdComment")
    Set l_adrsComment = l_objComment.GetComment(v_aryEmpNum, v_dtStartDate, v_dtEndDate)
    Set l_objComment = Nothing
    
    Set GetComment = l_adrsComment
    Set l_objComment = Nothing
    Set l_adrsComment = Nothing

    Exit Function
    
GetComment_Err:
    Set GetComment = Nothing
    Set l_objComment = Nothing
    Set l_adrsComment = Nothing
    
    TrsRaiseError m_strCMODULENAME, "GetComment", , , , , True
End Function
