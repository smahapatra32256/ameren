Attribute VB_Name = "mdbSecurity"
Option Explicit

Private Const MODULENAME = "mdbSecurity"

Public m_blnStaticVarsLoaded As Boolean
Public m_strDATASCOPE_EMP_ACTIVE_rebuild_sql As String
Public m_strDATASCOPE_EMP_ALL_rebuild_sql As String
Public m_strDATASCOPE_EMP_rebuild_changes_sql As String
Public m_strDATASCOPE_EMP_check_sql As String
Public m_strDATASCOPE_EMP_select_sql As String
Public m_strDATASCOPE_WG_rebuild_sql As String
Public m_strDATASCOPE_WG_check_sql As String
Public m_strDATASCOPE_WG_select_sql As String
Public m_strDATASCOPE_DEPT_rebuild_sql As String
Public m_strDATASCOPE_DEPT_check_sql As String
Public m_strDATASCOPE_DEPT_select_sql As String
Public m_strDATASCOPE_FUN_rebuild_sql As String
Public m_strDATASCOPE_FUN_check_sql As String
Public m_strDATASCOPE_FUN_select_sql As String
Public m_strDATASCOPE_DIV_rebuild_sql As String
Public m_strDATASCOPE_DIV_check_sql As String
Public m_strDATASCOPE_DIV_select_sql As String

Public m_strDATASCOPE_AEEMP_select_sql As String
Public m_strDATASCOPE_AEWG_select_sql As String
Public m_strDATASCOPE_AEDEPT_select_sql As String
Public m_strDATASCOPE_AEFUN_select_sql As String
Public m_strDATASCOPE_AEDIV_select_sql As String

Public m_strDATASCOPE_DWG_select_sql As String
Public m_strDATASCOPE_NWG_select_sql As String

Public m_strINVALIDATE_WG_By_Eff_Emp_sql As String
Public m_strINVALIDATE_Dept_By_Eff_Emp_sql As String
Public m_strINVALIDATE_Fun_By_Eff_Emp_sql As String
Public m_strINVALIDATE_Div_By_Eff_Emp_sql As String
'Public m_strINVALIDATE_WG_By_Ref_Emp_sql As String
'Public m_strINVALIDATE_Dept_By_Ref_Emp_sql As String
'Public m_strINVALIDATE_Div_By_Ref_Emp_sql As String
'Public m_strINVALIDATE_Fun_By_Ref_Emp_sql As String
'Public m_strINVALIDATE_Emp_By_Ref_WG_sql As String
'Public m_strINVALIDATE_Emp_By_Ref_Dept_sql As String
'Public m_strINVALIDATE_Emp_By_Ref_Fun_sql As String
'Public m_strINVALIDATE_Emp_By_Ref_Div_sql As String
'Public m_strINVALIDATE_WG_By_Ref_WG_sql As String
'Public m_strINVALIDATE_WG_By_Ref_Dept_sql As String
'Public m_strINVALIDATE_WG_By_Ref_Fun_sql As String
'Public m_strINVALIDATE_WG_By_Ref_Div_sql As String
'Public m_strINVALIDATE_Dept_By_Ref_Dept_sql As String
'Public m_strINVALIDATE_Dept_By_Ref_Fun_sql As String
'Public m_strINVALIDATE_Dept_By_Ref_Div_sql As String
'Public m_strINVALIDATE_Fun_By_Ref_Fun_sql As String
'Public m_strINVALIDATE_Fun_By_Ref_Div_sql As String
'Public m_strINVALIDATE_Div_By_Ref_Div_sql As String
Public m_strINVALIDATE_Emp_By_Ref_Emp_sql As String
Public m_strINVALIDATE_Emp_By_Eff_Emp_sql As String

Public m_strINVALIDATE_DS_By_Ref_Div_sql As String
Public m_strINVALIDATE_DS_By_Ref_Fun_sql As String
Public m_strINVALIDATE_DS_By_Ref_Dept_sql As String
Public m_strINVALIDATE_DS_By_Ref_WG_sql As String

Public m_strBATCH_TO_INVALIDATE_EMP_sql As String
Public m_strBATCH_TO_INVALIDATE_WG_sql As String
Public m_strBATCH_TO_INVALIDATE_DEPT_sql As String
Public m_strBATCH_TO_INVALIDATE_DEPT2_sql As String
Public m_strBATCH_TO_INVALIDATE_FUN2_sql As String
Public m_strBATCH_TO_INVALIDATE_DIV2_sql As String
Public m_strBATCH_TO_INVALIDATE_EMP3_sql As String
Public m_strBATCH_TO_INVALIDATE_WG3_sql As String
Public m_strBATCH_TO_INVALIDATE_DEPT3_sql As String
Public m_strBATCH_TO_INVALIDATE_FUN3_sql As String
Public m_strBATCH_TO_INVALIDATE_DIV3_sql As String

Public m_strBATCH_NEEDING_REBUILT_EMP_sql As String
Public m_strBATCH_NEEDING_REBUILT_WG_sql As String
Public m_strBATCH_NEEDING_REBUILT_DEPT_sql As String
Public m_strBATCH_NEEDING_REBUILT_FUN_sql As String
Public m_strBATCH_NEEDING_REBUILT_DIV_sql As String

Public m_strBATCH_REMOVE_sql As String

Public m_strBATCH_SYNC_ASSOCIATIONS_TO_AUDIT2_sql As String

Public m_strBATCH_NEEDING_REBUILT_TRACE_sql As String
Public m_strBATCH_SYNC_ASSOCIATIONS_TO_AUDIT2_TRACE_sql As String
Public m_strBATCH_REMOVE_TRACE_sql As String
Public m_strBATCH_TO_INVALIDATE_TRACE_sql As String

'Public m_strDATASCOPE_EMP_delete_WG_ASSGN_sql As String
'Public m_strDATASCOPE_EMP_insert_WG_ASSGN_sql As String

Public Sub LoadStaticVars()
    If Not m_blnStaticVarsLoaded Then '//-----do once-----//
        m_strDATASCOPE_EMP_ACTIVE_rebuild_sql = GetStringResource("DATASCOPE_EMP_ACTIVE_rebuild.sql")
        m_strDATASCOPE_EMP_ALL_rebuild_sql = GetStringResource("DATASCOPE_EMP_ALL_rebuild.sql")
        m_strDATASCOPE_EMP_rebuild_changes_sql = GetStringResource("DATASCOPE_EMP_rebuild_changes.sql")
        m_strDATASCOPE_EMP_check_sql = GetStringResource("DATASCOPE_EMP_check.sql")
        m_strDATASCOPE_EMP_select_sql = GetStringResource("DATASCOPE_EMP_select.sql")
        m_strDATASCOPE_WG_rebuild_sql = GetStringResource("DATASCOPE_WG_rebuild.sql")
        m_strDATASCOPE_WG_check_sql = GetStringResource("DATASCOPE_WG_check.sql")
        m_strDATASCOPE_WG_select_sql = GetStringResource("DATASCOPE_WG_select.sql")
        m_strDATASCOPE_DEPT_rebuild_sql = GetStringResource("DATASCOPE_DEPT_rebuild.sql")
        m_strDATASCOPE_DEPT_check_sql = GetStringResource("DATASCOPE_DEPT_check.sql")
        m_strDATASCOPE_DEPT_select_sql = GetStringResource("DATASCOPE_DEPT_select.sql")
        m_strDATASCOPE_FUN_rebuild_sql = GetStringResource("DATASCOPE_FUN_rebuild.sql")
        m_strDATASCOPE_FUN_check_sql = GetStringResource("DATASCOPE_FUN_check.sql")
        m_strDATASCOPE_FUN_select_sql = GetStringResource("DATASCOPE_FUN_select.sql")
        m_strDATASCOPE_DIV_rebuild_sql = GetStringResource("DATASCOPE_DIV_rebuild.sql")
        m_strDATASCOPE_DIV_check_sql = GetStringResource("DATASCOPE_DIV_check.sql")
        m_strDATASCOPE_DIV_select_sql = GetStringResource("DATASCOPE_DIV_select.sql")
        
        m_strDATASCOPE_AEEMP_select_sql = GetStringResource("DATASCOPE_AEEMP_select.sql")
        m_strDATASCOPE_AEWG_select_sql = GetStringResource("DATASCOPE_AEWG_select.sql")
        m_strDATASCOPE_AEDEPT_select_sql = GetStringResource("DATASCOPE_AEDEPT_select.sql")
        m_strDATASCOPE_AEFUN_select_sql = GetStringResource("DATASCOPE_AEFUN_select.sql")
        m_strDATASCOPE_AEDIV_select_sql = GetStringResource("DATASCOPE_AEDIV_select.sql")
                
        m_strDATASCOPE_DWG_select_sql = GetStringResource("DATASCOPE_DWG_select.sql")
        m_strDATASCOPE_NWG_select_sql = GetStringResource("DATASCOPE_NWG_select.sql")
'//
        m_strDATASCOPE_NWG_select_sql = GetStringResource("DATASCOPE_NWG_select.sql")

        m_strINVALIDATE_WG_By_Eff_Emp_sql = GetStringResource("INVALIDATE_WG_By_Eff_Emp.sql")
        m_strINVALIDATE_Dept_By_Eff_Emp_sql = GetStringResource("INVALIDATE_Dept_By_Eff_Emp.sql")
        m_strINVALIDATE_Fun_By_Eff_Emp_sql = GetStringResource("INVALIDATE_Fun_By_Eff_Emp.sql")
        m_strINVALIDATE_Div_By_Eff_Emp_sql = GetStringResource("INVALIDATE_Div_By_Eff_Emp.sql")
'        m_strINVALIDATE_WG_By_Ref_Emp_sql = GetStringResource("INVALIDATE_WG_By_Ref_Emp.sql")
'        m_strINVALIDATE_Dept_By_Ref_Emp_sql = GetStringResource("INVALIDATE_Dept_By_Ref_Emp.sql")
'        m_strINVALIDATE_Div_By_Ref_Emp_sql = GetStringResource("INVALIDATE_Div_By_Ref_Emp.sql")
'        m_strINVALIDATE_Fun_By_Ref_Emp_sql = GetStringResource("INVALIDATE_Fun_By_Ref_Emp.sql")
'        m_strINVALIDATE_Emp_By_Ref_WG_sql = GetStringResource("INVALIDATE_Emp_By_Ref_WG.sql")
'        m_strINVALIDATE_Emp_By_Ref_Dept_sql = GetStringResource("INVALIDATE_Emp_By_Ref_Dept.sql")
'        m_strINVALIDATE_Emp_By_Ref_Fun_sql = GetStringResource("INVALIDATE_Emp_By_Ref_Fun.sql")
'        m_strINVALIDATE_Emp_By_Ref_Div_sql = GetStringResource("INVALIDATE_Emp_By_Ref_Div.sql")
'        m_strINVALIDATE_WG_By_Ref_WG_sql = GetStringResource("INVALIDATE_WG_By_Ref_WG.sql")
'        m_strINVALIDATE_WG_By_Ref_Dept_sql = GetStringResource("INVALIDATE_WG_By_Ref_Dept.sql")
'        m_strINVALIDATE_WG_By_Ref_Fun_sql = GetStringResource("INVALIDATE_WG_By_Ref_Fun.sql")
'        m_strINVALIDATE_WG_By_Ref_Div_sql = GetStringResource("INVALIDATE_WG_By_Ref_Div.sql")
'        m_strINVALIDATE_Dept_By_Ref_Dept_sql = GetStringResource("INVALIDATE_Dept_By_Ref_Dept.sql")
'        m_strINVALIDATE_Dept_By_Ref_Fun_sql = GetStringResource("INVALIDATE_Dept_By_Ref_Fun.sql")
'        m_strINVALIDATE_Dept_By_Ref_Div_sql = GetStringResource("INVALIDATE_Dept_By_Ref_Div.sql")
'        m_strINVALIDATE_Fun_By_Ref_Fun_sql = GetStringResource("INVALIDATE_Fun_By_Ref_Fun.sql")
'        m_strINVALIDATE_Fun_By_Ref_Div_sql = GetStringResource("INVALIDATE_Fun_By_Ref_Div.sql")
'        m_strINVALIDATE_Div_By_Ref_Div_sql = GetStringResource("INVALIDATE_Div_By_Ref_Div.sql")
        m_strINVALIDATE_Emp_By_Ref_Emp_sql = GetStringResource("INVALIDATE_Emp_By_Ref_Emp.sql")
        m_strINVALIDATE_Emp_By_Eff_Emp_sql = GetStringResource("INVALIDATE_Emp_By_Eff_Emp.sql")
                
        m_strINVALIDATE_DS_By_Ref_Div_sql = GetStringResource("INVALIDATE_DS_By_Ref_Div.sql")
        m_strINVALIDATE_DS_By_Ref_Fun_sql = GetStringResource("INVALIDATE_DS_By_Ref_Fun.sql")
        m_strINVALIDATE_DS_By_Ref_Dept_sql = GetStringResource("INVALIDATE_DS_By_Ref_Dept.sql")
        m_strINVALIDATE_DS_By_Ref_WG_sql = GetStringResource("INVALIDATE_DS_By_Ref_WG.sql")
        
        m_strBATCH_TO_INVALIDATE_EMP_sql = GetStringResource("BATCH_TO_INVALIDATE_EMP.sql")
        m_strBATCH_TO_INVALIDATE_WG_sql = GetStringResource("BATCH_TO_INVALIDATE_WG.sql")
        m_strBATCH_TO_INVALIDATE_DEPT_sql = GetStringResource("BATCH_TO_INVALIDATE_DEPT.sql")
        m_strBATCH_TO_INVALIDATE_DEPT2_sql = GetStringResource("BATCH_TO_INVALIDATE_DEPT2.sql")
        m_strBATCH_TO_INVALIDATE_FUN2_sql = GetStringResource("BATCH_TO_INVALIDATE_FUN2.sql")
        m_strBATCH_TO_INVALIDATE_DIV2_sql = GetStringResource("BATCH_TO_INVALIDATE_DIV2.sql")

        m_strBATCH_TO_INVALIDATE_EMP3_sql = GetStringResource("BATCH_TO_INVALIDATE_EMP3.sql")
        m_strBATCH_TO_INVALIDATE_WG3_sql = GetStringResource("BATCH_TO_INVALIDATE_WG3.sql")
        m_strBATCH_TO_INVALIDATE_DEPT3_sql = GetStringResource("BATCH_TO_INVALIDATE_DEPT3.sql")
        m_strBATCH_TO_INVALIDATE_FUN3_sql = GetStringResource("BATCH_TO_INVALIDATE_FUN3.sql")
        m_strBATCH_TO_INVALIDATE_DIV3_sql = GetStringResource("BATCH_TO_INVALIDATE_DIV3.sql")
        
        m_strBATCH_NEEDING_REBUILT_EMP_sql = GetStringResource("BATCH_NEEDING_REBUILT_EMP.sql")
        m_strBATCH_NEEDING_REBUILT_WG_sql = GetStringResource("BATCH_NEEDING_REBUILT_WG.sql")
        m_strBATCH_NEEDING_REBUILT_DEPT_sql = GetStringResource("BATCH_NEEDING_REBUILT_DEPT.sql")
        m_strBATCH_NEEDING_REBUILT_FUN_sql = GetStringResource("BATCH_NEEDING_REBUILT_FUN.sql")
        m_strBATCH_NEEDING_REBUILT_DIV_sql = GetStringResource("BATCH_NEEDING_REBUILT_DIV.sql")

        m_strBATCH_REMOVE_sql = GetStringResource("BATCH_REMOVE.sql")
        
        m_strBATCH_SYNC_ASSOCIATIONS_TO_AUDIT2_sql = GetStringResource("BATCH_SYNC_ASSOCIATIONS_TO_AUDIT2.sql")

        m_strBATCH_NEEDING_REBUILT_TRACE_sql = GetStringResource("BATCH_NEEDING_REBUILT_TRACE.sql")
        m_strBATCH_SYNC_ASSOCIATIONS_TO_AUDIT2_TRACE_sql = GetStringResource("BATCH_SYNC_ASSOCIATIONS_TO_AUDIT2_TRACE.sql")
        m_strBATCH_REMOVE_TRACE_sql = GetStringResource("BATCH_REMOVE_TRACE.sql")
        m_strBATCH_TO_INVALIDATE_TRACE_sql = GetStringResource("BATCH_TO_INVALIDATE_TRACE.sql")
        
        'm_strDATASCOPE_EMP_delete_WG_ASSGN_sql = GetStringResource("DATASCOPE_EMP_delete_WG_ASSGN.sql")
        'm_strDATASCOPE_EMP_insert_WG_ASSGN_sql = GetStringResource("DATASCOPE_EMP_insert_WG_ASSGN.sql")
        
        m_blnStaticVarsLoaded = True
    End If
End Sub

Private Function GetStringResource(ByVal v_strFilename As String) As String
    Dim oFS As New FileSystemObject
    Dim ts As TextStream
    Dim sFilename As String, sFileContents As String
    
On Error GoTo eh
    sFilename = oFS.BuildPath("C:\TRIS SOLUTION\dbScripts\dbSecurity", v_strFilename)
    
    If oFS.FileExists(sFilename) Then
        Set ts = oFS.OpenTextFile(sFilename, ForReading, False)
        sFileContents = ts.ReadAll
    Else
        sFileContents = ""
        Err.Raise trErrbusService, , "Unable to load string resource for file" & v_strFilename
    End If
    
    Set oFS = Nothing
    Set ts = Nothing
    GetStringResource = sFileContents
    
    Exit Function
eh:
    TrsRaiseError MODULENAME _
        , "GetStringResource" _
        , , , _
        , DescArgsFrom("v_strFilename = " & v_strFilename, "sFilename = " & sFilename) _
        , True
End Function
