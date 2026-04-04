VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "cbsSysAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Project:   Post TRIS 5.5 Enhancement
'Author:    Prem Chander
'Date:      June 2005
'Description:   This module provides the ability to query and modify
'TRIS Data.

Option Explicit

Private Const MODULENAME = "busSysAdmin.cbsSysAdmin"
Public Function ProcessSysAdminRequest(ByVal vsRequest As String _
                    , ByVal vlUserNumber As Long _
                    , ByRef radrsMulti() As ADODB.Recordset _
                    , ByRef rlRecordsAffected As Long _
                    , ByRef rsErrorString As String _
                    , Optional ByVal vsUserID As String = "" _
                    ) As Long
On Error GoTo ProcessSysAdminRequest_EH
Dim lobjdb As dbSysAdmin.cdSysAdmin
Dim ladrs As ADODB.Recordset

    ProcessSysAdminRequest = -1

    Select Case (vsUserID)
        'Greg Barks, Ratan Shabbir
        'Tim Vincent, John Stricklin, Jim Barnett
        'Herb Medley, Lisa Bertke, Prem Chander
        'Rick Guffey
        Case "E26649", "E26791" _
             , "E26895", "E27158", "E33724" _
             , "E34684", "E33642", "E41802" _
             , "E84531"
            Set lobjdb = CreateObject("dbSysAdmin.cdSysAdmin")
            'Set lobjdb = CreateObject("dbSysAdmin.cdSysAdmin")
            Set ladrs = lobjdb.ExecuteSysAdminRequest(vsRequest _
                            , vlUserNumber _
                            , rlRecordsAffected _
                            , rsErrorString)
            Set lobjdb = Nothing

            ProcessSysAdminRequest = DeepCopyRecordSet(ladrs, radrsMulti())

        Case Else
            '
    End Select

    Set ladrs = Nothing

    Exit Function
ProcessSysAdminRequest_EH:
    Set lobjdb = Nothing
    Set ladrs = Nothing

    TrsRaiseError MODULENAME _
        , "ProcessSysAdminRequest" _
        , , True _
        , , , True _
        , CStr(vlUserNumber)
End Function
Public Function ProcessSysAdminSELECT(ByVal vsRequest As String _
                    , ByVal vlUserNumber As Long _
                    , ByRef radrsMulti() As ADODB.Recordset _
                    , ByRef rsErrorString As String _
                    , Optional ByVal vsUserID As String = "" _
                    ) As Long
On Error GoTo ProcessSysAdminSELECT_EH
Dim lobjdb As dbSysAdmin.cdSysAdmin
Dim ladrs As ADODB.Recordset

    ProcessSysAdminSELECT = -1

    Select Case (vsUserID)
        'Greg Barks, Ratan Shabbir
        'Tim Vincent, John Stricklin, Jim Barnett
        'Herb Medley, Lisa Bertke, Prem Chander
        'Rick Guffey
        Case "E26649", "E26791" _
             , "E26895", "E27158", "E33724" _
             , "E34684", "E33642", "E41802" _
             , "E84531"
            Set lobjdb = CreateObject("dbSysAdmin.cdSysAdmin")
            'Set lobjdb = CreateObject("dbSysAdmin.cdSysAdmin")
            Set ladrs = lobjdb.ExecuteSysAdminSELECT(vsRequest _
                            , vlUserNumber _
                            , rsErrorString)
            Set lobjdb = Nothing

            ProcessSysAdminSELECT = DeepCopyRecordSet(ladrs, radrsMulti())

        Case Else
            '
    End Select

    Set ladrs = Nothing

    Exit Function
ProcessSysAdminSELECT_EH:
    Set lobjdb = Nothing
    Set ladrs = Nothing

    TrsRaiseError MODULENAME _
        , "ProcessSysAdminSELECT" _
        , , True _
        , , , True _
        , CStr(vlUserNumber)
End Function
Private Function DeepCopyRecordSet(ByVal vadrsSource As ADODB.Recordset _
            , ByRef radrsCopy() As ADODB.Recordset _
            ) As Long
On Error GoTo DeepCopyRecordSet_EH
Dim llIndex As Long, lsFldName As String
Dim llFldSize As Long, ladFldType As ADODB.DataTypeEnum
Dim ladFldAttrib As ADODB.FieldAttributeEnum
Dim llFldIndex As Long, llCount As Long

    ReDim Preserve radrsCopy(0)
    Set radrsCopy(0) = New ADODB.Recordset

    llCount = 0

    Do
        ReDim Preserve radrsCopy(llCount)

        Set radrsCopy(llCount) = New ADODB.Recordset

        With vadrsSource
            For llIndex = 0 To .Fields.Count - 1
                'Get the fld name, size
                lsFldName = .Fields.Item(llIndex).Name
                ladFldType = .Fields.Item(llIndex).Type
                llFldSize = .Fields.Item(llIndex).DefinedSize
                ladFldAttrib = .Fields.Item(llIndex).Attributes

                If (llFldSize > 0) Then
                    radrsCopy(llCount).Fields.Append lsFldName, ladFldType _
                        , llFldSize, ladFldAttrib
                Else
                    radrsCopy(llCount).Fields.Append lsFldName, ladFldType _
                        , , ladFldAttrib
                End If
            Next llIndex
        End With

        With radrsCopy(llCount)
            'Open the recordset
            .CursorLocation = adUseClient
            .LockType = adLockBatchOptimistic
            .CursorType = adOpenStatic
            .Open
            .ActiveConnection = Nothing

            If Not (vadrsSource.BOF And vadrsSource.EOF) Then
                Do
                    .AddNew

                    On Error Resume Next
                    For llFldIndex = 0 To vadrsSource.Fields.Count - 1
                        .Fields.Item(llFldIndex).Value _
                            = vadrsSource.Fields.Item(llFldIndex).Value
                    Next llFldIndex

                    .Update

                    vadrsSource.MoveNext
                Loop Until (vadrsSource.EOF)
            End If

            If Not (.BOF And .EOF) Then
                .MoveFirst
            End If
        End With

        Set vadrsSource = vadrsSource.NextRecordset
        llCount = llCount + 1
'        ReDim Preserve radrsCopy(llCount)
'
'        Set radrsCopy(llCount) = New ADODB.Recordset
    Loop Until (vadrsSource Is Nothing)

    DeepCopyRecordSet = llCount

    Exit Function
DeepCopyRecordSet_EH:
    Err.Raise Err.Number, "DeepCopyRecordSet"

End Function
Public Function ProcessSysAdminINSUPDDEL(ByVal vsRequest As String _
                    , ByVal vlUserNumber As Long _
                    , ByRef rsErrorString As String _
                    , Optional ByVal vsUserID As String = "" _
                    ) As Long
On Error GoTo ProcessSysAdminINSUPDDEL_EH
Dim lobjdb As dbSysAdmin.cdSysAdmin

    Select Case (vsUserID)
        'Greg Barks, Ratan Shabbir
        'Tim Vincent, John Stricklin, Jim Barnett
        'Herb Medley, Lisa Bertke, Prem Chander
        'Rick Guffey
        Case "E26649", "E26791" _
             , "E26895", "E27158", "E33724" _
             , "E34684", "E33642", "E41802" _
             , "E84531"
            Set lobjdb = CreateObject("dbSysAdmin.cdSysAdmin")
            'Set lobjdb = CreateObject("dbSysAdmin.cdSysAdmin")
            ProcessSysAdminINSUPDDEL = lobjdb.ExecuteSysAdminINSUPDDEL(vsRequest _
                                            , vlUserNumber _
                                            , rsErrorString)
            Set lobjdb = Nothing

        Case Else
            '
    End Select

    Exit Function
ProcessSysAdminINSUPDDEL_EH:
    Set lobjdb = Nothing

    TrsRaiseError MODULENAME _
        , "ProcessSysAdminINSUPDDEL" _
        , , True _
        , , , True _
        , CStr(vlUserNumber)
End Function
Public Sub GetAllCatalogInformation(ByRef radrsTables As ADODB.Recordset _
            , ByRef radrsColumns As ADODB.Recordset _
            , ByRef radrsIndices As ADODB.Recordset _
            , ByRef radrsKeys As ADODB.Recordset _
            , ByVal vlUserNumber As Long _
            , Optional ByVal vsUserID As String)
On Error GoTo GetAllCatalogInformation_EH

    Select Case (vsUserID)
        'Greg Barks, Ratan Shabbir
        'Tim Vincent, John Stricklin, Jim Barnett
        'Herb Medley, Lisa Bertke, Prem Chander
        'Rick Guffey
        Case "E26649", "E26791" _
             , "E26895", "E27158", "E33724" _
             , "E34684", "E33642", "E41802" _
             , "E84531"
            Set radrsTables = GetTableCatalogInformation()
            Set radrsColumns = GetColumnCatalogInformation()
            Set radrsIndices = GetIndexCatalogInformation()
            Set radrsKeys = GetIndexKeysCatalogInformation()

        Case Else
            '
    End Select

    Exit Sub
GetAllCatalogInformation_EH:
    TrsRaiseError MODULENAME _
        , "GetAllCatalogInformation" _
        , , True _
        , , , True _
        , CStr(vlUserNumber)
End Sub
Public Function GetTableCatalogInformation() As ADODB.Recordset
On Error GoTo GetTableCatalogInformation_EH
Dim lobjdb As dbSysAdmin.cdSysAdmin

    Set lobjdb = CreateObject("dbSysAdmin.cdSysAdmin")
    Set GetTableCatalogInformation = lobjdb.GetTableCatalogInformation
    Set lobjdb = Nothing

    Exit Function
GetTableCatalogInformation_EH:
    Set lobjdb = Nothing

    TrsRaiseError MODULENAME _
        , "GetTableCatalogInformation" _
        , , True _
        , , , True

End Function
Public Function GetColumnCatalogInformation() As ADODB.Recordset
On Error GoTo GetColumnCatalogInformation_EH
Dim lobjdb As dbSysAdmin.cdSysAdmin

    Set lobjdb = CreateObject("dbSysAdmin.cdSysAdmin")
    Set GetColumnCatalogInformation = lobjdb.GetColumnCatalogInformation()
    Set lobjdb = Nothing

    Exit Function
GetColumnCatalogInformation_EH:
    Set lobjdb = Nothing

    TrsRaiseError MODULENAME _
        , "GetColumnCatalogInformation" _
        , , True _
        , , , True

End Function
Public Function GetIndexCatalogInformation() As ADODB.Recordset
On Error GoTo GetIndexCatalogInformation_EH
Dim lobjdb As dbSysAdmin.cdSysAdmin

    Set lobjdb = CreateObject("dbSysAdmin.cdSysAdmin")
    Set GetIndexCatalogInformation = lobjdb.GetIndexCatalogInformation()
    Set lobjdb = Nothing

    Exit Function
GetIndexCatalogInformation_EH:
    Set lobjdb = Nothing

    TrsRaiseError MODULENAME _
        , "GetIndexKeysCatalogInformation" _
        , , True _
        , , , True

End Function
Public Function GetIndexKeysCatalogInformation() As ADODB.Recordset
On Error GoTo GetIndexKeysCatalogInformation_EH
Dim lobjdb As dbSysAdmin.cdSysAdmin

    Set lobjdb = CreateObject("dbSysAdmin.cdSysAdmin")
    Set GetIndexKeysCatalogInformation = lobjdb.GetIndexKeysCatalogInformation()
    Set lobjdb = Nothing

    Exit Function
GetIndexKeysCatalogInformation_EH:
    Set lobjdb = Nothing

    TrsRaiseError MODULENAME _
        , "GetIndexKeysCatalogInformation" _
        , , True _
        , , , True

End Function
