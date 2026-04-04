Attribute VB_Name = "mTest"

'Public Function rsBrowse(v_rsBrowse1 As ADODB.Recordset, Optional v_rsBrowse2 As ADODB.Recordset, _
'                Optional v_rsBrowse3 As ADODB.Recordset, Optional v_rsBrowse4 As ADODB.Recordset, _
'                Optional v_rsBrowse5 As ADODB.Recordset, Optional v_rsBrowse6 As ADODB.Recordset)
'    Dim o As New TnvLib.Debug
'
'    o.AddRecordSet "rsBrowse1", v_rsBrowse1
'
'    If Not IsMissing(v_rsBrowse2) Then
'        o.AddRecordSet "rsBrowse2", v_rsBrowse2
'    End If
'    If Not IsMissing(v_rsBrowse3) Then
'        o.AddRecordSet "rsBrowse3", v_rsBrowse3
'    End If
'    If Not IsMissing(v_rsBrowse4) Then
'        o.AddRecordSet "rsBrowse4", v_rsBrowse4
'    End If
'    If Not IsMissing(v_rsBrowse5) Then
'        o.AddRecordSet "rsBrowse5", v_rsBrowse5
'    End If
'    If Not IsMissing(v_rsBrowse6) Then
'        o.AddRecordSet "rsBrowse6", v_rsBrowse6
'    End If
'
'    o.Display
'
'    Set o = Nothing
'End Function

Public Function rsSave(rs As ADODB.Recordset, sFile As String)
    Dim varSave As Variant
    Dim bRestore As Boolean
    
    bRestore = False
    
    If InStr(1, sFile, ".") = 0 Then
        sFile = sFile & ".adtg"
    End If
    
    If Dir(sFile) <> "" Then
        Kill sFile
    End If
    
    If Not rs Is Nothing Then
        If Not (rs.BOF And rs.EOF) Then
            varSave = rs.Bookmark
            rs.MoveFirst
            bRestore = True
        End If
    End If
    
    rs.Save sFile
    
    If bRestore Then
        rs.Bookmark = varSave
    End If

End Function
