VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cdEmpFMLA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const m_strCMODULENAME As String = "cdEmpFMLA"
Private Const m_strCProcessID As String = "PS2TRIS:FMLA"
Private m_lngStep As Long
Private m_blnAudit As Boolean
Private m_objPSCN As New ADODB.Connection
Private m_objTRCN As New ADODB.Connection


Public Sub Process_TTR_EMP_FMLA(Optional ByVal siOption_Num As Integer = 0, _
                                Optional ByVal strStart_Dt As String = "01-01-1994", _
                                Optional ByVal strEnd_Dt As String = "12-31-9999", _
                                Optional blnAudit As Boolean = False)

On Error GoTo Err_Process_TTR_EMP_FMLA

    Dim l_strProcessId As String
    Dim l_strSQL As String
    Dim l_lngRecAffected As Long
    Dim l_lngRow As Long
    Dim m_lngStep As Long
    
    Dim l_objCM As util_DBConnectMgr.clsuDBConnectMgr
    
'    Dim m_objPSCN As New ADODB.Connection
'    Dim m_objTRCN As New ADODB.Connection
    
    Dim l_objRs0 As ADODB.Recordset
    Dim l_objRs1 As ADODB.Recordset
    Dim l_objRs2 As ADODB.Recordset
    Dim l_objRs4 As ADODB.Recordset
    
    Set l_objCM = New clsuDBConnectMgr
    m_objTRCN.ConnectionString = l_objCM.GetConnectionString("SHADOWDIRECT")
    m_objTRCN.CursorLocation = adUseClient
    m_objTRCN.Open
    Set l_objCM = Nothing
    
    m_blnAudit = blnAudit
    m_lngStep = 0
    
    l_strSQL = " UPDATE METHDBA.TTR_PROCESS_CNTL" & _
               " SET (START_TIMESTAMP,CURRENT_STEP,STATUS_CD,RETURN_CD,SQL_CD,ERROR_CD,LAST_MOD_AUTH_ID,LAST_MOD_TIMESTAMP,END_TIMESTAMP,PROGRAM_ID) " & _
               "   = (CURRENT TIMESTAMP," & m_lngStep & ",-1,-1,-1,-1,USER,CURRENT TIMESTAMP,'9999-12-31-24.00.00.000000','" & m_strCMODULENAME & "')" & _
               " WHERE PROCESS_ID = '" & m_strCProcessID & "';"
               
    m_objTRCN.Execute l_strSQL, l_lngRecAffected
    
    Select Case siOption_Num
        Case 0
            l_strSQL = " SELECT" & _
                       "  CASE" & _
                       "  WHEN DAYOFWEEK(DATE(PROCESS_TIMESTAMP)) = 2" & _
                       "  THEN DATE(PROCESS_TIMESTAMP) - 14 DAYS" & _
                       "  WHEN DAYOFWEEK(DATE(PROCESS_TIMESTAMP)) = 3" & _
                       "  THEN DATE(PROCESS_TIMESTAMP) - 5 DAYS" & _
                       "  ELSE DATE(PROCESS_TIMESTAMP) - 3 DAYS" & _
                       "  END As START_DT" & _
                       " ,CURRENT DATE AS END_DT" & _
                       " FROM" & _
                       " METHDBA.TTR_PROCESS_CNTL" & _
                       " WHERE PROCESS_ID = 'PS2TRIS:FMLA'"
              
            Set l_objRs0 = CreateObject("ADODB.Recordset")
            Set l_objRs0 = New ADODB.Recordset
            l_objRs0.Open l_strSQL, m_objTRCN, adOpenStatic, adLockReadOnly, adCmdText
            
            InsertProcessAudit "S", "TTR_PROCESS_CNTL", l_objRs0.RecordCount
            
            strStart_Dt = l_objRs0.Fields("START_DT").Value
            strEnd_Dt = l_objRs0.Fields("END_DT").Value
        Case 1
            'Continue
    End Select
    
    Set l_objRs0 = Nothing
    
    Set l_objCM = New clsuDBConnectMgr
    
    m_objPSCN.ConnectionString = l_objCM.GetConnectionString("PeopleSoft")
    m_objPSCN.CursorLocation = adUseClient
    m_objPSCN.Open
    Set l_objCM = Nothing
    
    l_strSQL = " DELETE FROM SYSADM.PS_TAR_GTT_FMLA_01"
    m_objPSCN.Execute l_strSQL, l_lngRecAffected
           
    InsertProcessAudit "D", "GTT_FMLA_01", l_lngRecAffected
    
    l_strSQL = " DELETE FROM SYSADM.PS_TAR_GTT_FMLA_02"
    m_objPSCN.Execute l_strSQL, l_lngRecAffected
           
    InsertProcessAudit "D", "GTT_FMLA_02", l_lngRecAffected
    
    l_strSQL = " INSERT INTO SYSADM.PS_TAR_GTT_FMLA_01 (DAYOFWEEK) VALUES(0)"
    m_objPSCN.Execute l_strSQL, l_lngRecAffected
    
    InsertProcessAudit "I", "GTT_FMLA_01", l_lngRecAffected
    
    l_strSQL = " INSERT INTO SYSADM.PS_TAR_GTT_FMLA_01 (DAYOFWEEK) VALUES(1)"
    m_objPSCN.Execute l_strSQL, l_lngRecAffected
    
    l_strSQL = " INSERT INTO SYSADM.PS_TAR_GTT_FMLA_01 (DAYOFWEEK) VALUES(2)"
    m_objPSCN.Execute l_strSQL, l_lngRecAffected
    
    l_strSQL = " INSERT INTO SYSADM.PS_TAR_GTT_FMLA_01 (DAYOFWEEK) VALUES(3)"
    m_objPSCN.Execute l_strSQL, l_lngRecAffected
    
    l_strSQL = " INSERT INTO SYSADM.PS_TAR_GTT_FMLA_01 (DAYOFWEEK) VALUES(4)"
    m_objPSCN.Execute l_strSQL, l_lngRecAffected
    
    l_strSQL = " INSERT INTO SYSADM.PS_TAR_GTT_FMLA_01 (DAYOFWEEK) VALUES(5)"
    m_objPSCN.Execute l_strSQL, l_lngRecAffected
    
    l_strSQL = " INSERT INTO SYSADM.PS_TAR_GTT_FMLA_01 (DAYOFWEEK) VALUES(6)"
    m_objPSCN.Execute l_strSQL, l_lngRecAffected

    l_strSQL = "COMMIT"
    m_objPSCN.Execute l_strSQL, l_lngRecAffected
    
'    l_strSQL = " INSERT INTO SYSADM.PS_TAR_GTT_FMLA_02" & _
'             " (EMP_NUM,FMLA_DT,FMLA_HRS,FMLA_REASON_CD,CREATED_EMP_NUM,CREATED_DT)" & _
'             " SELECT" & _
'             "  TO_NUMBER(A.EMPLID) AS EMP_NUM" & _
'             " ,B.WEEK_BEGIN_DT + DAYOFWEEK AS FMLA_DATE" & _
'             " ,CASE" & _
'             "    WHEN C.DAYOFWEEK = 0" & _
'             "      THEN B.SUN_HRS" & _
'             "    WHEN C.DAYOFWEEK = 1" & _
'             "      THEN B.MON_HRS" & _
'             "    WHEN C.DAYOFWEEK = 2" & _
'             "      THEN B.TUES_HRS" & _
'             "    WHEN C.DAYOFWEEK = 3" & _
'             "      THEN B.WED_HRS" & _
'             "    WHEN C.DAYOFWEEK = 4" & _
'             "      THEN B.THURS_HRS" & _
'             "    WHEN C.DAYOFWEEK = 5" & _
'             "      THEN B.FRI_HRS" & _
'             "    ELSE   B.SAT_HRS" & _
'             "  END AS FMLA_HRS" & _
'             " ,A.FMLA_LEAVE_REASON AS FMLA_REASON_CD" & _
'             " ,TO_NUMBER(A.APPROVER_ID) AS CREATED_EMP_NUM" & _
'             " ,A.APPRVL_DT AS CREATED_DT"
'
'    l_strSQL = l_strSQL & _
'             " FROM" & _
'             " SYSADM.PS_TAR_GTT_FMLA_01 C" & _
'             " ," & _
'             " SYSADM.PS_FMLA_LV_REQUEST A" & _
'             " INNER JOIN" & _
'             " SYSADM.PS_FMLA_LV_HISTORY B" & _
'             " ON  B.EMPLID = A.EMPLID" & _
'             " AND B.FMLA_REQUEST_ID = A.FMLA_REQUEST_ID" & _
'             " WHERE A.FMLA_LV_APPROVED = 'A'" & _
'             "   AND A.APPRVL_DT BETWEEN TO_DATE('" & Format(strStart_Dt, "YYYY-MM-DD") & "','YYYY-MM-DD') and TO_DATE('" & Format(strEnd_Dt, "YYYY-MM-DD") & "','YYYY-MM-DD')"
    
    l_strSQL = " INSERT INTO SYSADM.PS_TAR_GTT_FMLA_02 (EMP_NUM,FMLA_DT,FMLA_HRS,FMLA_REASON_CD,CREATED_EMP_NUM,CREATED_DT)" & _
             " SELECT" & _
             "  TO_NUMBER(D.EMPLID) AS EMP_NUM" & _
             " ,C.WEEK_BEGIN_DT + E.DAYOFWEEK AS FMLA_DATE" & _
             " ,CASE    WHEN E.DAYOFWEEK = 0      THEN C.SUN_HRS    WHEN E.DAYOFWEEK = 1      THEN C.MON_HRS    WHEN E.DAYOFWEEK = 2      THEN C.TUES_HRS    WHEN E.DAYOFWEEK = 3      THEN C.WED_HRS    WHEN E.DAYOFWEEK = 4      THEN C.THURS_HRS    WHEN E.DAYOFWEEK = 5      THEN C.FRI_HRS    ELSE   C.SAT_HRS  END AS FMLA_HRS" & _
             " ,D.FMLA_LEAVE_REASON AS FMLA_REASON_CD" & _
             " ,TO_NUMBER(D.APPROVER_ID) AS CREATED_EMP_NUM" & _
             " ,D.APPRVL_DT AS CREATED_DT"
    
    l_strSQL = l_strSQL & _
             " FROM" & _
             " SYSADM.PS_TAR_GTT_FMLA_01 E" & _
             " ," & _
             " SYSADM.PS_FMLA_LV_REQUEST A" & _
             " INNER JOIN" & _
             " SYSADM.PS_FMLA_LV_HISTORY B" & _
             " ON  B.EMPLID = A.EMPLID" & _
             " AND B.FMLA_REQUEST_ID = A.FMLA_REQUEST_ID" & _
             " INNER JOIN" & _
             " SYSADM.PS_FMLA_LV_HISTORY C" & _
             " ON  C.EMPLID = B.EMPLID" & _
             " AND C.WEEK_BEGIN_DT = B.WEEK_BEGIN_DT" & _
             " INNER JOIN" & _
             " SYSADM.PS_FMLA_LV_REQUEST D" & _
             " ON  D.EMPLID = C.EMPLID" & _
             " AND D.FMLA_REQUEST_ID = C.FMLA_REQUEST_ID" & _
             " AND D.FMLA_LV_APPROVED = 'A'" & _
             " WHERE A.FMLA_LV_APPROVED = 'A'" & _
             " AND A.APPRVL_DT BETWEEN TO_DATE('" & Format(strStart_Dt, "YYYY-MM-DD") & "','YYYY-MM-DD') and TO_DATE('" & Format(strEnd_Dt, "YYYY-MM-DD") & "','YYYY-MM-DD')"
    
    m_objPSCN.Execute l_strSQL, l_lngRecAffected
    
    InsertProcessAudit "I", "GTT_FMLA_02", l_lngRecAffected
    
    If l_lngRecAffected = 0 Then
        l_strSQL = "COMMIT"
        m_objPSCN.Execute l_strSQL, l_lngRecAffected
        GoTo Cleanup
    End If
    
    l_strSQL = "COMMIT"
    m_objPSCN.Execute l_strSQL, l_lngRecAffected

    l_strSQL = " SELECT DISTINCT * FROM (SELECT" & _
              " ' INSERT INTO SESSION.GTT_FMLA_DEL_00 (EMP_NUM,START_DT,END_DT) VALUES('" & _
              " ||TRIM(EMP_NUM)" & _
              " ||','''" & _
              " ||TO_CHAR(FMLA_DT,'YYYY-MM-DD')" & _
              " ||''','''" & _
              " ||TO_CHAR(FMLA_DT + 6,'YYYY-MM-DD')" & _
              " ||''''" & _
              " ||');' AS FMLA_SQL" & _
              " FROM SYSADM.PS_TAR_GTT_FMLA_02" & _
              " WHERE TRIM(TO_CHAR(FMLA_DT,'DAY')) = 'SUNDAY'" & _
              " ORDER BY" & _
              "  EMP_NUM" & _
              " ,FMLA_DT) NTE0"
              
    Set l_objRs1 = CreateObject("ADODB.Recordset")
    Set l_objRs1 = New ADODB.Recordset
    
    l_objRs1.Open l_strSQL, m_objPSCN, adOpenStatic, adLockReadOnly, adCmdText
              
    InsertProcessAudit "S", "GTT_FMLA_02", l_objRs1.RecordCount
    
    l_strSQL = "COMMIT"
    m_objPSCN.Execute l_strSQL, l_lngRecAffected
    
    m_objTRCN.BeginTrans

    l_strSQL = " DECLARE GLOBAL TEMPORARY TABLE SESSION.GTT_FMLA_DEL_00" & _
              " (EMP_NUM         INTEGER NOT NULL" & _
              " ,START_DT        DATE NOT NULL" & _
              " ,END_DT          DATE NOT NULL" & _
              " )" & _
              " CCSID EBCDIC" & _
              " ON COMMIT PRESERVE ROWS;"
    
    m_objTRCN.Execute l_strSQL, l_lngRecAffected
    
    InsertProcessAudit "G", "GTT_FMLA_DEL_00", l_lngRecAffected
    
    l_objRs1.MoveFirst
    
    l_strSQL = ""

    For l_lngRow = 1 To l_objRs1.RecordCount
        If Len(l_strSQL) + l_objRs1.Fields("FMLA_SQL").ActualSize < 64000 Then
            l_strSQL = l_strSQL & l_objRs1.Fields("FMLA_SQL").Value
        Else
            m_objTRCN.Execute l_strSQL, l_lngRecAffected
            m_objTRCN.CommitTrans
            m_objTRCN.BeginTrans
            l_strSQL = l_objRs1.Fields("FMLA_SQL").Value
        End If
        
        l_objRs1.MoveNext
    Next
    
    If l_strSQL <> "" Then
        m_objTRCN.Execute l_strSQL, l_lngRecAffected
        m_objTRCN.CommitTrans
        m_objTRCN.BeginTrans
    End If
    
    Set l_objRs4 = New ADODB.Recordset
    l_objRs4.Open "select count(*) as record_cnt from session.gtt_fmla_del_00", m_objTRCN, adOpenStatic, adLockReadOnly, adCmdText
    l_objRs4.MoveFirst
    InsertProcessAudit "I", "GTT_FMLA_DEL_00", l_objRs4.Fields("RECORD_CNT").Value
    
    l_strSQL = " DECLARE GLOBAL TEMPORARY TABLE SESSION.GTT_FMLA_DEL_01" & _
              " (EMP_NUM         INTEGER NOT NULL" & _
              " ,START_DT        DATE NOT NULL" & _
              " ,END_DT          DATE NOT NULL" & _
              " )" & _
              " CCSID EBCDIC" & _
              " ON COMMIT PRESERVE ROWS;"

    m_objTRCN.Execute l_strSQL, l_lngRecAffected
    
    InsertProcessAudit "G", "GTT_FMLA_DEL_01", l_lngRecAffected
    
    l_strSQL = " CREATE INDEX SESSION.XGTTFMLADEL01 ON SESSION.GTT_FMLA_DEL_01" & _
              " (EMP_NUM  ASC" & _
              " ,START_DT ASC" & _
              " ,END_DT   ASC" & _
              " );"

    m_objTRCN.Execute l_strSQL, l_lngRecAffected
    
    InsertProcessAudit "G", "XGTTFMLADEL01", l_lngRecAffected
    
    l_strSQL = " INSERT INTO SESSION.GTT_FMLA_DEL_01" & _
              " (EMP_NUM,START_DT,END_DT)" & _
              " SELECT" & _
              "  A.EMP_NUM" & _
              " ,B.START_DT" & _
              " ,B.END_DT" & _
              " FROM" & _
              " SESSION.GTT_FMLA_DEL_00 A" & _
              " INNER JOIN" & _
              " METHDBA.TTR_EMP_FMLA B" & _
              " ON  B.EMP_NUM = A.EMP_NUM" & _
              " AND B.START_DT <= A.END_DT" & _
              " AND B.END_DT >= A.START_DT" & _
              " AND B.LAST_MOD_EMP_NUM = 999999999" & _
              " GROUP BY" & _
              "  A.EMP_NUM" & _
              " ,B.START_DT" & _
              " ,B.END_DT;"

    m_objTRCN.Execute l_strSQL, l_lngRecAffected
    
    InsertProcessAudit "I", "GTT_FMLA_DEL_01", l_lngRecAffected
    
    If l_lngRecAffected > 0 Then
        l_strSQL = " DELETE FROM METHDBA.TTR_EMP_FMLA A" & _
                   " WHERE EXISTS" & _
                   "      (SELECT 1" & _
                   "       FROM" & _
                   "       SESSION.GTT_FMLA_DEL_01 B" & _
                   "       WHERE B.EMP_NUM = A.EMP_NUM" & _
                   "         AND B.START_DT = A.START_DT" & _
                   "         AND B.END_DT = A.END_DT);"
    
        m_objTRCN.Execute l_strSQL, l_lngRecAffected
        InsertProcessAudit "D", "TTR_EMP_FMLA", l_lngRecAffected
        m_objTRCN.CommitTrans
        m_objTRCN.BeginTrans
    End If
    
'    l_strSQL = " SELECT" & _
'               " ' INSERT INTO SESSION.GTT_FMLA_STAGING (EMP_NUM,FMLA_DT,FMLA_HRS,FMLA_REASON_CD,CREATED_EMP_NUM,CREATED_DT) VALUES('" & _
'               " ||TRIM(EMP_NUM)" & _
'               " ||','''" & _
'               " ||TO_CHAR(FMLA_DT,'YYYY-MM-DD')" & _
'               " ||''','" & _
'               " ||FMLA_HRS" & _
'               " ||','''" & _
'               " ||FMLA_REASON_CD" & _
'               " ||''','" & _
'               " ||TRIM(CREATED_EMP_NUM)" & _
'               " ||','''" & _
'               " ||TO_CHAR(CREATED_DT,'YYYY-MM-DD')" & _
'               " ||''''" & _
'               " ||');' AS FMLA_SQL" & _
'               " FROM SYSADM.PS_TAR_GTT_FMLA_02" & _
'               " WHERE FMLA_HRS > 0" & _
'               " ORDER BY" & _
'               "  EMP_NUM" & _
'               " ,FMLA_DT"
              
    l_strSQL = " SELECT DISTINCT *" & _
             " FROM" & _
             " (" & _
             " SELECT ' INSERT INTO SESSION.GTT_FMLA_STAGING (EMP_NUM,FMLA_DT,FMLA_HRS,FMLA_REASON_CD,CREATED_EMP_NUM,CREATED_DT) VALUES(' ||TRIM(EMP_NUM) ||',''' ||TO_CHAR(FMLA_DT,'YYYY-MM-DD') ||''',' ||FMLA_HRS ||',''' ||FMLA_REASON_CD ||''',' ||TRIM(CREATED_EMP_NUM) ||',''' ||TO_CHAR(CREATED_DT,'YYYY-MM-DD') ||'''' ||');' AS FMLA_SQL" & _
             " FROM" & _
             " (" & _
             " SELECT" & _
             "  NTE0.EMP_NUM" & _
             " ,NTE0.FMLA_DT" & _
             " ,NTE0.FMLA_HRS" & _
             " ,NTE1.FMLA_REASON_CD" & _
             " ,NTE1.CREATED_EMP_NUM" & _
             " ,NTE0.CREATED_DT"

    l_strSQL = l_strSQL & _
             " FROM" & _
             " (" & _
             " SELECT" & _
             "  EMP_NUM" & _
             " ,FMLA_DT" & _
             " ,SUM(FMLA_HRS) AS FMLA_HRS" & _
             " ,MAX(CREATED_DT) AS CREATED_DT" & _
             " FROM" & _
             " SYSADM.PS_TAR_GTT_FMLA_02" & _
             " WHERE FMLA_HRS <> 0" & _
             " GROUP BY" & _
             "  EMP_NUM" & _
             " ,FMLA_DT" & _
             " ) NTE0"

    l_strSQL = l_strSQL & _
             " INNER JOIN" & _
             " (" & _
             " SELECT" & _
             "  EMP_NUM" & _
             " ,FMLA_DT" & _
             " ,CREATED_DT" & _
             " ,FMLA_REASON_CD" & _
             " ,MIN(CREATED_EMP_NUM) AS CREATED_EMP_NUM" & _
             " FROM" & _
             " SYSADM.PS_TAR_GTT_FMLA_02" & _
             " WHERE FMLA_HRS <> 0" & _
             " GROUP BY" & _
             "  EMP_NUM" & _
             " ,FMLA_DT" & _
             " ,CREATED_DT" & _
             " ,FMLA_REASON_CD" & _
             " ) NTE1"
             
    l_strSQL = l_strSQL & _
             " ON  NTE1.EMP_NUM = NTE0.EMP_NUM" & _
             " AND NTE1.FMLA_DT = NTE0.FMLA_DT" & _
             " AND NTE1.CREATED_DT = NTE0.CREATED_DT" & _
             " ) NTE2" & _
             " ORDER BY" & _
             "  EMP_NUM" & _
             " ,FMLA_DT" & _
             " ) NTE3"
              
    Set l_objRs2 = CreateObject("ADODB.Recordset")
    Set l_objRs2 = New ADODB.Recordset
    l_objRs2.Open l_strSQL, m_objPSCN, adOpenStatic, adLockReadOnly, adCmdText
    
    InsertProcessAudit "S", "GTT_FMLA_02", l_objRs2.RecordCount
    
    If l_objRs2.RecordCount = 0 Then
        l_strSQL = "COMMIT"
        m_objPSCN.Execute l_strSQL, l_lngRecAffected
        GoTo Cleanup
    End If
    
    l_strSQL = "COMMIT"
    m_objPSCN.Execute l_strSQL, l_lngRecAffected
    
    l_strSQL = " DECLARE GLOBAL TEMPORARY TABLE SESSION.GTT_FMLA_STAGING" & _
              " (EMP_NUM         INTEGER NOT NULL" & _
              " ,FMLA_DT         DATE NOT NULL" & _
              " ,FMLA_HRS        DECIMAL(10,1) NOT NULL" & _
              " ,FMLA_REASON_CD  CHAR(2) NOT NULL" & _
              " ,CREATED_EMP_NUM INTEGER NOT NULL" & _
              " ,CREATED_DT      DATE NOT NULL" & _
              " )" & _
              " CCSID EBCDIC" & _
              " ON COMMIT PRESERVE ROWS;"
    
    m_objTRCN.Execute l_strSQL, l_lngRecAffected
    
    InsertProcessAudit "G", "GTT_FMLA_STAGING", l_lngRecAffected
    
    l_objRs2.MoveFirst
    
    l_strSQL = ""
    
    For l_lngRow = 1 To l_objRs2.RecordCount
        If Len(l_strSQL) + l_objRs2.Fields("FMLA_SQL").ActualSize < 64000 Then
            l_strSQL = l_strSQL & l_objRs2.Fields("FMLA_SQL").Value
        Else
            m_objTRCN.Execute l_strSQL, l_lngRecAffected
            m_objTRCN.CommitTrans
            m_objTRCN.BeginTrans
            l_strSQL = l_objRs2.Fields("FMLA_SQL").Value
        End If
        
        l_objRs2.MoveNext
    Next
    
    If l_strSQL <> "" Then
        m_objTRCN.Execute l_strSQL, l_lngRecAffected
        m_objTRCN.CommitTrans
        m_objTRCN.BeginTrans
    End If
    
    Set l_objRs4 = New ADODB.Recordset
    l_objRs4.Open "select count(*) as record_cnt from session.gtt_fmla_staging", m_objTRCN, adOpenStatic, adLockReadOnly, adCmdText
    l_objRs4.MoveFirst
    InsertProcessAudit "I", "GTT_FMLA_STAGING", l_objRs4.Fields("RECORD_CNT").Value
    
    l_strSQL = " CREATE INDEX SESSION.XGTTFMLASTAGING" & _
              " ON SESSION.GTT_FMLA_STAGING" & _
              " (EMP_NUM ASC" & _
              " ,FMLA_DT ASC" & _
              " ,FMLA_REASON_CD ASC" & _
              " )"

    m_objTRCN.Execute l_strSQL, l_lngRecAffected
    
    InsertProcessAudit "G", "XGTTFMLASTAGING", l_lngRecAffected
    
    l_strSQL = " INSERT INTO METHDBA.TTR_EMP_FMLA" & _
              " (EMP_NUM,START_DT,END_DT,FMLA_HRS,FMLA_REASON_CD,CREATED_EMP_NUM,CREATED_TIMESTAMP,LAST_MOD_EMP_NUM,LAST_MOD_TIMESTAMP)" & _
              " SELECT" & _
              "  A.EMP_NUM" & _
              " ,A.FMLA_DT AS START_DT" & _
              " ,A.FMLA_DT AS END_DT" & _
              " ,A.FMLA_HRS" & _
              " ,A.FMLA_REASON_CD" & _
              " ,A.CREATED_EMP_NUM" & _
              " ,TIMESTAMP(A.CREATED_DT,'00:00:00') AS CREATED_TIMESTAMP" & _
              " ,999999999 AS LAST_MOD_EMP_NUM" & _
              " ,CURRENT TIMESTAMP AS LAST_MOD_TIMESTAMP" & _
              " FROM" & _
              " SESSION.GTT_FMLA_STAGING A" & _
              " LEFT OUTER JOIN" & _
              " METHDBA.TTR_EMP_FMLA B" & _
              " ON  B.EMP_NUM = A.EMP_NUM" & _
              " AND B.START_DT <= A.FMLA_DT" & _
              " AND B.END_DT >= A.FMLA_DT"
    
    l_strSQL = l_strSQL & _
              " WHERE A.FMLA_HRS > 0" & _
              "   AND B.EMP_NUM IS NULL" & _
              " GROUP BY" & _
              "  A.EMP_NUM" & _
              " ,A.FMLA_DT" & _
              " ,A.FMLA_HRS" & _
              " ,A.FMLA_REASON_CD" & _
              " ,A.CREATED_EMP_NUM" & _
              " ,A.CREATED_DT;"

    m_objTRCN.Execute l_strSQL, l_lngRecAffected
              
    InsertProcessAudit "I", "TTR_EMP_FMLA", l_lngRecAffected
    
    m_objTRCN.CommitTrans
    
Cleanup:
    l_strSQL = " UPDATE METHDBA.TTR_PROCESS_CNTL" & _
               " SET (PROCESS_TIMESTAMP,STATUS_CD,RETURN_CD,LAST_MOD_AUTH_ID,LAST_MOD_TIMESTAMP,END_TIMESTAMP,PROGRAM_ID)" & _
               "   = (CURRENT TIMESTAMP,0,0,USER,CURRENT TIMESTAMP,CURRENT TIMESTAMP,'" & m_strCMODULENAME & "')" & _
               " WHERE PROCESS_ID = '" & m_strCProcessID & "';"

    m_objTRCN.Execute l_strSQL, l_lngRecAffected

    Set l_objRs0 = Nothing
    Set l_objRs1 = Nothing
    Set l_objRs2 = Nothing
    Set l_objRs4 = Nothing
    
    m_objPSCN.Close
    Set m_objPSCN = Nothing
    
    m_objTRCN.Close
    Set m_objTRCN = Nothing
    
    Exit Sub
    
Err_Process_TTR_EMP_FMLA:
    Set l_objRs0 = Nothing
    Set l_objRs1 = Nothing
    Set l_objRs2 = Nothing
    Set l_objRs4 = Nothing
    
    If m_objPSCN.Errors.Count > 0 Then
        TrsRaiseError m_strCMODULENAME, "Process_TTR_EMP_FMLA:PeopleSoft Connection", , , , DescArgsFrom(l_strSQL), True
        Set m_objPSCN = Nothing
        Set m_objTRCN = Nothing
        Exit Sub
    End If
    
    If m_objTRCN.Errors.Count > 0 Then
        TrsRaiseError m_strCMODULENAME, "Process_TTR_EMP_FMLA:TRIS Connection", , , , DescArgsFrom(l_strSQL), True
        Set m_objPSCN = Nothing
        Set m_objTRCN = Nothing
        Exit Sub
    End If
    
    TrsRaiseError m_strCMODULENAME, "Process_TTR_EMP_FMLA:Other", , , , DescArgsFrom(l_strSQL), True
End Sub
Private Sub UpdateProcessCntl()
On Error GoTo Err_UpdateProcessCntl

    Dim l_strSQL As String
    Dim l_lngRecAffected As Long
    
    m_lngStep = m_lngStep + 1
    
    l_strSQL = " UPDATE METHDBA.TTR_PROCESS_CNTL" & _
               " SET (CURRENT_STEP,STATUS_CD,LAST_MOD_AUTH_ID,LAST_MOD_TIMESTAMP,PROGRAM_ID) " & _
               "   = (" & m_lngStep & ",0,USER,CURRENT TIMESTAMP,'" & m_strCMODULENAME & "')" & _
               " WHERE PROCESS_ID = '" & m_strCProcessID & "';"
    
    m_objTRCN.Execute l_strSQL, l_lngRecAffected

    Exit Sub
    
Err_UpdateProcessCntl:
    
    TrsRaiseError m_strCMODULENAME, "UpdateProcessCntl", , , , DescArgsFrom(l_strSQL), True
End Sub
Private Sub InsertProcessAudit(strOperation_Cd As String, strTableNm As String, lngRecordsAffected As Long)
On Error GoTo Err_InsertProcessAudit

    Dim l_strSQL As String
    Dim l_lngRecAffected As Long
    
    If m_blnAudit = False Then
       Exit Sub
    End If
    
    l_strSQL = " INSERT INTO METHDBA.TTR_PROCESS_AUDIT" & _
               " (PROCESS_ID,LAST_MOD_AUTH_ID,STEP,OPERATION_CD,TABLE_NM,RECORDS_AFFECTED,PROGRAM_ID)" & _
               " VALUES('" & m_strCProcessID & "',USER," & m_lngStep & ",'" & strOperation_Cd & "','" & strTableNm & "'," & lngRecordsAffected & ",'" & m_strCMODULENAME & "')"
    
    m_objTRCN.Execute l_strSQL, l_lngRecAffected

    UpdateProcessCntl
    
    Exit Sub
    
Err_InsertProcessAudit:
    
    TrsRaiseError m_strCMODULENAME, "InsertProcessAudit", , , , DescArgsFrom(l_strSQL), True
End Sub




