VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cdiNewFunSecurity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Project:   TRIS Conversion Phase 1
'Author:    Prem M. Chander
'Date:      July 2001 (Firstcut)
'Description: This class module is an abstract class Public NotCreatable
'This is used for defining new interface.
'NOTE: Client version number will be included in the new interface

Public Function dbInsCustomerInfo2(ByVal v_sCustomerID As String _
                , ByVal v_lCustomerNumber As Long _
                , ByVal v_sClientMachineID As String _
                , ByVal v_sRemoteServerName As String _
                , ByVal v_sClientVersionNumber As String _
                ) As Boolean

End Function

Public Function dbUpdCustomerInfo2(ByVal v_sCustomerID As String _
                , ByVal v_lCustomerNumber As Long _
                , ByVal v_sClientMachineID As String _
                , ByVal v_sRemoteServerName As String _
                , ByVal v_sPrevSessTimeStamp As String _
                , ByVal v_sPrevSessMachineID As String _
                , ByVal v_sPrevSessRSN As String _
                , ByVal v_sClientVersionNumber As String _
                ) As Boolean

End Function
