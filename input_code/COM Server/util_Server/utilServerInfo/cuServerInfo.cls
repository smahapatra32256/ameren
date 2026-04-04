VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "cuServerInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const m_strCMODULENAME As String = "cuServerInfo"
Private Const MODULENAME As String = "cuServerInfo"

' NLB Cluster/Cluster Members Group constants
Private Const NLBCLUSTER_CACHE = "NLBCLUSTERGROUP"
Private Const NLBCLUSTER_CACHE_KEY = 0
Private Const NLBCLUSTER_CACHE_ROWVALUES = 1
Private Const NLBCLUSTER_CACHE_LAST_SLOT = 1
Private Const NLBCLUSTER_CACHE_CLUSTERNAME = 0
Private Const NLBCLUSTER_CACHE_CLUSTERMEMBER = 1

Public Function GetServerInfo(Optional ByVal v_strCategory As String = "") As ADODB.Recordset
    Dim l_rsInfo As ADODB.Recordset
    Dim lobjCodeTables As dbCodeTables.cdCodeTables
    Dim lobjDistMgr As dbDistributionMgr.cdDistributionMgr
    Dim lobjDBConnectMgr As util_DBConnectMgr.clsuDBConnectMgr
    Dim lobjClientRelease As dbClientRelease.cdClientRelease

    Dim l_rsClientVersion As ADODB.Recordset
    Dim l_strDB2Subsys As String
    Dim sCategory As String
    Dim rsGetDB2Info As ADODB.Recordset
On Error GoTo eh

    v_strCategory = LCase(Trim(v_strCategory))
    
    Set l_rsInfo = CreateInfoRecordSet
'    Set l_rsInfo = New ADODB.Recordset
    With l_rsInfo
'        .Fields.Append "Category", adVarChar, 20
'        .Fields.Append "Property", adVarChar, 40
'        .Fields.Append "Value", adVarChar, 255
'
'        .CursorLocation = adUseClient
'        .CursorType = adOpenDynamic
'        .LockType = adLockBatchOptimistic
'        .Open
        
        If v_strCategory = "" Or v_strCategory = "all" Or v_strCategory = "db" Then
            Set lobjDBConnectMgr = CreateObject("util_DBConnectMgr.clsuDBConnectMgr")
            
            .AddNew
            .Fields("Category").Value = "Database"
            .Fields("Key").Value = "DSN"
            .Fields("Property").Value = .Fields("Key").Value
            .Fields("Value").Value = lobjDBConnectMgr.GetDataSourceName '//"DB2BLOC"
            .Update

            .AddNew
            .Fields("Category").Value = "Database"
            .Fields("Key").Value = "User Id"
            .Fields("Property").Value = .Fields("Key").Value
            .Fields("Value").Value = lobjDBConnectMgr.GetUserId '//"TR1STEST"
            .Update
            
            .AddNew
            .Fields("Category").Value = "Database"
            .Fields("Key").Value = "DSN2"
            .Fields("Property").Value = .Fields("Key").Value
            .Fields("Value").Value = lobjDBConnectMgr.GetDataSourceName("SHADOWDIRECT2") '//"DB2BLOC"
            .Update

            .AddNew
            .Fields("Category").Value = "Database"
            .Fields("Key").Value = "User Id2"
            .Fields("Property").Value = .Fields("Key").Value
            .Fields("Value").Value = lobjDBConnectMgr.GetUserId("SHADOWDIRECT2") '//"TR1STEST"
            .Update
            
            .AddNew
            .Fields("Category").Value = "Database"
            .Fields("Key").Value = "PeopleSoftDSN"
            .Fields("Property").Value = .Fields("Key").Value
            .Fields("Value").Value = lobjDBConnectMgr.GetDataSourceName("PeopleSoft")
            .Update

            .AddNew
            .Fields("Category").Value = "Database"
            .Fields("Key").Value = "PeopleSoftUserId"
            .Fields("Property").Value = .Fields("Key").Value
            .Fields("Value").Value = lobjDBConnectMgr.GetUserId("PeopleSoft") '//"TR1STEST"
            .Update
            
            .AddNew
            .Fields("Category").Value = "Database"
            .Fields("Key").Value = "EMPRV_DSN"
            .Fields("Property").Value = .Fields("Key").Value
            .Fields("Value").Value = lobjDBConnectMgr.GetDataSourceName("EMPRV")
            .Update

            .AddNew
            .Fields("Category").Value = "Database"
            .Fields("Key").Value = "EMPRV_UserId"
            .Fields("Property").Value = .Fields("Key").Value
            .Fields("Value").Value = lobjDBConnectMgr.GetUserId("EMPRV") '//"TR1STEST"
            .Update
            
            Set lobjDBConnectMgr = Nothing

            sCategory = "Database"
            Set rsGetDB2Info = GetDB2Info()
            
            If RecordsExist(rsGetDB2Info) Then
                rsGetDB2Info.MoveFirst
                
                Do While Not rsGetDB2Info.EOF
                    AddServerInfo l_rsInfo, sCategory, "", "", rsGetDB2Info.Fields("Property").Value, rsGetDB2Info.Fields("Value").Value
                    rsGetDB2Info.MoveNext
                Loop
            End If
                        
            Set lobjDistMgr = CreateObject("dbDistributionMgr.cdDistributionMgr")
            
            .AddNew
            .Fields("Category").Value = "Database"
            .Fields("Key").Value = "DB2 Subsystem"
            .Fields("Property").Value = .Fields("Key").Value
            l_strDB2Subsys = lobjDistMgr.GetDB2Subsystem("SHADOWDIRECT2")
            .Fields("Value").Value = l_strDB2Subsys '//"DB2B"
            .Update
            
            .AddNew
            .Fields("Category").Value = "Database"
            .Fields("Key").Value = "Edit Module Region2"
            .Fields("Property").Value = .Fields("Key").Value
            .Fields("Value").Value = lobjDistMgr.GetCICSRegion(l_strDB2Subsys) '//"SY7821"
            .Update
            
            Set lobjDistMgr = Nothing
            
            Set lobjClientRelease = CreateObject("dbClientRelease.cdClientRelease")
            
            Set l_rsClientVersion = lobjClientRelease.db_GetClientVersionInfo
            If Not l_rsClientVersion Is Nothing Then
                If Not (l_rsClientVersion.BOF And l_rsClientVersion.EOF) Then
                    l_rsClientVersion.MoveFirst
                    Do While Not l_rsClientVersion.EOF
                        .AddNew
                        .Fields("Category").Value = "Database"
                        .Fields("Key").Value = l_rsClientVersion.Fields("PARM_DESC")
                        .Fields("Property").Value = .Fields("Key").Value
                        .Fields("Value").Value = l_rsClientVersion.Fields("PARM_VALUE")
                        .Update
                        
                        l_rsClientVersion.MoveNext
                    Loop
                End If
            End If
            Set lobjClientRelease = Nothing
        End If

        If v_strCategory = "" Or v_strCategory = "all" Or v_strCategory = "com+" Then
            Dim oCatalog As COMAdmin.COMAdminCatalog
            Set oCatalog = CreateObject("COMAdmin.COMAdminCatalog")
            
            Dim oApplications As COMAdmin.COMAdminCatalogCollection
            Dim oApplication As COMAdmin.COMAdminCatalogObject

            oCatalog.Connect ("")

            Set oApplications = oCatalog.GetCollection("Applications")
            oApplications.Populate

            For Each oApplication In oApplications
                If oApplication.Name = "TRIS" Then
                    .AddNew
                    .Fields("Category").Value = "COM+"
                    .Fields("Key").Value = "COM+ Application Identity"
                    .Fields("Property").Value = .Fields("Key").Value
                    .Fields("Value").Value = oApplication.Value("Identity")
                    .Update
                End If
            Next

'            Dim oApplicationInstances As COMAdmin.COMAdminCatalogCollection
'            Dim oApplicationInstance As COMAdmin.COMAdminCatalogObject
'
'            Set oApplicationInstances = oCatalog.GetCollection("ApplicationInstances")
'            oApplicationInstances.Populate
'
'            For Each oApplicationInstance In oApplicationInstances
''                    .AddNew
''                    .Fields("Category").Value = "COM+"
''                    .Fields("Property").Value = "COM+ Process Id"
''                    .Fields("Value").Value = oApplicationInstance.Value("ProcessId")
''                    .Update
''
''                    .AddNew
''                    .Fields("Category").Value = "COM+"
''                    .Fields("Property").Value = "COM+ Partition Name"
''                    .Fields("Value").Value = oApplicationInstance.Value("PartitionId")
''                    .Update
'''oCatalog.GetPartitionName
'                    .AddNew
'                    .Fields("Category").Value = "COM+"
'                    .Fields("Property").Value = "COM+ Application Id"
'                    .Fields("Value").Value = oApplicationInstance.Value("Application")
'                    .Update
'            Next

            .AddNew
            .Fields("Category").Value = "COM+"
            .Fields("Key").Value = "Current Partition ID"
            .Fields("Property").Value = .Fields("Key").Value
            .Fields("Value").Value = oCatalog.CurrentPartitionID
            .Update
            
            .AddNew
            .Fields("Category").Value = "COM+"
            .Fields("Key").Value = "Current Partition Name"
            .Fields("Property").Value = .Fields("Key").Value
            .Fields("Value").Value = oCatalog.CurrentPartitionName
            .Update
        End If
                
        If v_strCategory = "" Or v_strCategory = "all" Or v_strCategory = "date" Then
            .AddNew
            .Fields("Category").Value = "Date"
            .Fields("Key").Value = "Server Date Time"
            .Fields("Property").Value = .Fields("Key").Value
            .Fields("Value").Value = Now()
            .Update
        End If
        
        If v_strCategory = "" Or v_strCategory = "all" Or v_strCategory = "cache" Then
            Set lobjCodeTables = CreateObject("dbCodeTables.cdCodeTables")
            
            .AddNew
            .Fields("Category").Value = "Cache"
            .Fields("Key").Value = "History Cutoff Date"
            .Fields("Property").Value = .Fields("Key").Value
'//            .Fields("Value").Value = CDate(DateAdd("m", -6, Now))
            .Fields("Value").Value = lobjCodeTables.GetActualTimeHSCutoffDate(CKEYCUTOFFDATE)
            .Update
            
            Set lobjCodeTables = Nothing
        End If
        
        If v_strCategory = "" Or v_strCategory = "all" Or v_strCategory = "server" Then
'            .AddNew
'            .Fields("Category").Value = "Server"
'            .Fields("Key").Value = "Computer Name"
'            .Fields("Property").Value = .Fields("Key").Value
'            .Fields("Value").Value = Environ("COMPUTERNAME")
'            .Update
            
            sCategory = "Server"
            AddServerInfo l_rsInfo, sCategory, "", "", "Computer Name", Environ("COMPUTERNAME")
            AddServerInfo l_rsInfo, sCategory, "", "", "Computer IP", GetIPAddress(Environ("COMPUTERNAME"))
            'AddServerInfo l_rsInfo, sCategory, "", "", "Computer Ping", PingServer(Environ("COMPUTERNAME"))
            
            .AddNew
            .Fields("Category").Value = "Server"
            .Fields("Key").Value = "Number of Processors"
            .Fields("Property").Value = .Fields("Key").Value
            .Fields("Value").Value = Environ("NUMBER_OF_PROCESSORS")
            .Update
            
'            .AddNew
'            .Fields("Category").Value = "Server"
'            .Fields("Property").Value = "User Name"
'            .Fields("Value").Value = Environ("USERNAME")
'            .Update
'
'            .AddNew
'            .Fields("Category").Value = "Server"
'            .Fields("Property").Value = "User Domain"
'            .Fields("Value").Value = Environ("USERDOMAIN")
'            .Update
            
'//could find useful items to return from WMI
'//C:\Projects\vbs\wmi
            GetNLBNodeInfo l_rsInfo
            
            GetCRACInfo l_rsInfo
        End If
 
         If v_strCategory = "dbconnectinfo" Then
            Set lobjDBConnectMgr = CreateObject("util_DBConnectMgr.clsuDBConnectMgr")
            
            .AddNew
            .Fields("Category").Value = "Database"
            .Fields("Key").Value = "SHADOWDIRECT"
            .Fields("Property").Value = .Fields("Key").Value
            .Fields("Value").Value = lobjDBConnectMgr.GetConnectionString("SHADOWDIRECT")
            .Update
            
            .AddNew
            .Fields("Category").Value = "Database"
            .Fields("Key").Value = "SHADOWDIRECT2"
            .Fields("Property").Value = .Fields("Key").Value
            .Fields("Value").Value = lobjDBConnectMgr.GetConnectionString("SHADOWDIRECT2")
            .Update
            
            .AddNew
            .Fields("Category").Value = "Database"
            .Fields("Key").Value = "PeopleSoft"
            .Fields("Property").Value = .Fields("Key").Value
            .Fields("Value").Value = lobjDBConnectMgr.GetConnectionString("PEOPLESOFT")
            .Update
            
            .AddNew
            .Fields("Category").Value = "Database"
            .Fields("Key").Value = "CorpAcctEdit"
            .Fields("Property").Value = .Fields("Key").Value
            .Fields("Value").Value = lobjDBConnectMgr.GetConnectionString("CorpAcctEdit")
            .Update
            
            Set lobjDBConnectMgr = Nothing
        End If
        
    End With
    
    Set GetServerInfo = l_rsInfo
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "GetServerInfo", , , , , True
End Function

Private Function AddServerInfo(ByRef rs As ADODB.Recordset, _
                                ByVal sCategory As String, _
                                ByVal sParentKey As String, _
                                ByVal sKey As String, _
                                ByVal sProperty As String, _
                                ByVal sValue As String _
                                )
    On Error Resume Next
    
    rs.AddNew
    rs.Fields("Category").Value = sCategory
'    rs.Fields("ParentKey").Value = sParentKey
    If sKey <> "" Then
        rs.Fields("Key").Value = sKey
    Else
        rs.Fields("Key").Value = sProperty
    End If
    rs.Fields("Property").Value = sProperty
    rs.Fields("Value").Value = sValue '//IIf(sValue <> "", sValue, "null")
    rs.Update
    
    On Error GoTo 0
    Exit Function
eh:
    Err.Raise Err.Number _
        , m_strCMODULENAME & ": AddServerInfo " & sCategory & "  " & sProperty & "  " & sValue, Err.Description
End Function

'//-----for utilRecordCache.InvalidateCache
Public Function GetServersInMyCluster(ByVal v_strServerName As String, _
                                    Optional ByVal v_bolAvailableOnly As Boolean = True) As Variant
    Dim aryCluster As Variant
    Dim xx As Long, yy As Long

On Error GoTo eh
    Dim sClusterName As String, aryClusterMembers As Variant

    yy = 0
    
    '//-----most of the time this first ping should not be needed
    If PingServer(v_strServerName) Then '//-----check for powered off
        sClusterName = GetNLBClusterName(v_strServerName)
        If Not (sClusterName = "" Or IsEmpty(sClusterName)) Then
            aryClusterMembers = GetNLBClusterMembers(sClusterName)
            
            If Not IsEmpty(aryClusterMembers) Then
                ReDim aryCluster(UBound(aryClusterMembers))
        
                For xx = LBound(aryClusterMembers) To UBound(aryClusterMembers)
                    If PingServer(aryClusterMembers(xx)) Then '//-----check for powered off
                        '//If IsNLBServerAvailable(aryClusterMembers(xx)) Then
                            aryCluster(yy) = aryClusterMembers(xx)
                            yy = yy + 1
                        '//End If
                    End If
                Next xx
            End If
        End If
    End If
    
    If yy > 0 Then
        ReDim Preserve aryCluster(yy - 1) '//-----exclude empty elements at the bottom of the array-----//
        GetServersInMyCluster = aryCluster
    Else
        GetServersInMyCluster = Empty
    End If
    
    Exit Function
eh:
    GetServersInMyCluster = Empty
    '//-----check for invalid namespace or not using NLB or powered off or preparing to shut down-----//
    If Err.Number = -2147217394 Or Err.Number = -2147217389 Or Err.Number = -2147023174 Or Err.Number = -2147023781 Then
    Else
        TrsTraceError m_strCMODULENAME, "GetServersInMyCluster", , , , DescArgsFrom(v_strServerName), True
    End If
End Function

Private Function PingServer(ByVal v_strServerName As String) As Boolean
    Dim objWMI As Object 'WbemScripting.SWbemServicesEx
    Dim objPing As Object 'WbemScripting.SWbemObjectSet
    Dim objStatus As Object 'WbemScripting.SWbemObjectEx

On Error GoTo eh
    
    PingServer = False

    Set objWMI = GetObject("winmgmts:{impersonationLevel=impersonate}")
    Set objPing = objWMI.ExecQuery("select * from Win32_PingStatus where address = '" & v_strServerName & "'")
    For Each objStatus In objPing
        If IsNull(objStatus.StatusCode) Or objStatus.StatusCode <> 0 Then
            PingServer = False
        Else
            PingServer = True
        End If
    Next
    
    Exit Function
eh:
    PingServer = False
    '//-----check for invalid namespace or not using NLB or powered off or preparing to shut down-----//
    If Err.Number = -2147217394 Or Err.Number = -2147217389 Or Err.Number = -2147023174 Or Err.Number = -2147023781 Then
    Else
        TrsTraceError m_strCMODULENAME, "PingServer", , , , DescArgsFrom(v_strServerName), True
    End If
End Function

Private Function GetIPAddress(ByVal v_strServerName As String) As String
    Dim objWMI As Object 'WbemScripting.SWbemServicesEx
    Dim objPing As Object 'WbemScripting.SWbemObjectSet
    Dim objStatus As Object 'WbemScripting.SWbemObjectEx

On Error GoTo eh
    
    GetIPAddress = "127.0.0.1"

    Set objWMI = GetObject("winmgmts:{impersonationLevel=impersonate}")
    Set objPing = objWMI.ExecQuery("select * from Win32_PingStatus where address = '" & v_strServerName & "'")
    For Each objStatus In objPing
        If IsNull(objStatus.StatusCode) Or objStatus.StatusCode <> 0 Then
        Else
            GetIPAddress = objStatus.ProtocolAddress
        End If
    Next
    
    Exit Function
eh:
    '//-----check for invalid namespace or not using NLB or powered off or preparing to shut down-----//
    If Err.Number = -2147217394 Or Err.Number = -2147217389 Or Err.Number = -2147023174 Or Err.Number = -2147023781 Then
    Else
        TrsTraceError m_strCMODULENAME, "GetIPAddress", , , , DescArgsFrom(v_strServerName), True
    End If
End Function

Private Function IsNLBServerAvailable(ByVal v_strServerName As String) As Boolean
    Dim xx As Long, l_strThisComputerName As String

On Error GoTo eh
   
    IsNLBServerAvailable = False
    
    Dim SWbemLocator As WbemScripting.SWbemLocator
    Set SWbemLocator = CreateObject("WbemScripting.SWbemLocator")

    Dim SWbemServices As WbemScripting.SWbemServices
    '//-----receive wmi error cannot supply user credentials when running locally???-----//
    l_strThisComputerName = Trim(stdmTRISCommon.ThisComputerName)
    If UCase(v_strServerName) = UCase(l_strThisComputerName) Then
        v_strServerName = "."
    End If
'    If v_strServerName = "." Then
        Set SWbemServices = SWbemLocator.ConnectServer(v_strServerName, "root\MicrosoftNLB")
'    Else
'        Set SWbemServices = SWbemLocator.ConnectServer(v_strServerName, "root\MicrosoftNLB", v_strServerName & "\db2admin", "db2admin")
'    End If

    Dim objNodes As WbemScripting.SWbemObjectSet
    Set objNodes = SWbemServices.InstancesOf("MicrosoftNLB_Node")
    
    Dim objNode As WbemScripting.SWbemObject
    
    If Not objNodes Is Nothing Then
        xx = 0
        For Each objNode In objNodes
            IsNLBServerAvailable = IsNLBNodeActive(objNode)
            xx = xx + 1
        Next
    End If

    Exit Function
eh:
    IsNLBServerAvailable = False
    '//-----check for invalid namespace or not using NLB or powered off or preparing to shut down-----//
    If Err.Number = -2147217394 Or Err.Number = -2147217389 Or Err.Number = -2147023174 Or Err.Number = -2147023781 Then
    Else
        TrsTraceError m_strCMODULENAME, "IsNLBServerAvailable", , , , DescArgsFrom(v_strServerName), True
    End If
End Function

'Private Function CreateWBEMServices(ByVal v_strServerName As String) As WbemScripting.SWbemServices
'    Dim l_strThisComputerName As String
'    Dim SWbemLocator As WbemScripting.SWbemLocator
'    Set SWbemLocator = CreateObject("WbemScripting.SWbemLocator")
'
'    Dim SWbemServices As WbemScripting.SWbemServices
'    '//-----receive wmi error cannot supply user credentials when running locally???-----//
'    l_strThisComputerName = Trim(ThisComputerName)
'    If UCase(v_strServerName) = UCase(l_strThisComputerName) Then
'        v_strServerName = "."
'    End If
'    If v_strServerName = "." Then
'        Set SWbemServices = SWbemLocator.ConnectServer(v_strServerName, "root\MicrosoftNLB")
'    Else
'        Set SWbemServices = SWbemLocator.ConnectServer(v_strServerName, "root\MicrosoftNLB", v_strServerName & "\db2admin", "db2admin")
'    End If
'    Exit Function
'eh:
'    Set GetWBEMServices = Nothing
'    '//-----check for invalid namespace or not using NLB-----//
'    If Err.Number = -2147217394 Or Err.Number = -2147217389 Then
'    Else
'        TrsRaiseError m_strCMODULENAME, "IsNLBServerAvailable", , , , , True
'    End If
'
'End Function

Private Function GetComputerNameFromFullDNSName(ByVal v_strFullName As String) As String
    Dim nDelim As Long
    
    nDelim = InStr(1, v_strFullName, ".")
    
    If nDelim > 1 Then
        GetComputerNameFromFullDNSName = Left(v_strFullName, nDelim - 1)
    Else
        GetComputerNameFromFullDNSName = v_strFullName
    End If

End Function

Private Function GetCRACInfo(ByVal v_rsServerInfo As ADODB.Recordset) As Long
    Dim sCategory  As String
    Dim l_rsCRACInfo As ADODB.Recordset
    Dim Time As Ameren_TRIS_Server_COMIntegration.Time
    
    sCategory = "Server"
    
    Set Time = New Ameren_TRIS_Server_COMIntegration.Time
    Set l_rsCRACInfo = Time.GetServerInfo()
    
    If RecordsExist(l_rsCRACInfo) Then
        l_rsCRACInfo.MoveFirst
        
        Do While Not l_rsCRACInfo.EOF
            AddServerInfo v_rsServerInfo, l_rsCRACInfo.Fields("Category").Value, "", _
                                        l_rsCRACInfo.Fields("Key").Value, l_rsCRACInfo.Fields("Property").Value, l_rsCRACInfo.Fields("Value").Value
            
            l_rsCRACInfo.MoveNext
        Loop
    End If
    
    Set Time = Nothing
    Set l_rsCRACInfo = Nothing
End Function

'//-----for display
Public Function GetNLBNodeInfo(ByRef v_rsServerInfo As ADODB.Recordset) As Long

On Error GoTo eh
    Dim sServer As String, sNodeName As String, sClusterName As String
    sServer = "."

    Dim SWbemLocator As WbemScripting.SWbemLocator
    Set SWbemLocator = CreateObject("WbemScripting.SWbemLocator")

    Dim SWbemServices As WbemScripting.SWbemServices
    Set SWbemServices = SWbemLocator.ConnectServer(sServer, "root\MicrosoftNLB")

    If v_rsServerInfo Is Nothing Then
        Set v_rsServerInfo = CreateInfoRecordSet
    End If

    Dim objNodes As WbemScripting.SWbemObjectSet
    Set objNodes = SWbemServices.InstancesOf("MicrosoftNLB_Node")

    Dim objNode As WbemScripting.SWbemObject
    If Not objNodes Is Nothing Then
        For Each objNode In objNodes
            sNodeName = GetComputerNameFromFullDNSName(objNode.ComputerName)
            With v_rsServerInfo
                .AddNew
                .Fields("Category").Value = "MicrosoftNLB_Node"
                .Fields("ParentKey").Value = sNodeName
                .Fields("Key").Value = "ComputerName"
                .Fields("Property").Value = .Fields("Key").Value
                .Fields("Value").Value = objNode.ComputerName
                .Update
    
                .AddNew
                .Fields("Category").Value = "MicrosoftNLB_Node"
                .Fields("ParentKey").Value = sNodeName
                .Fields("Key").Value = "DedicatedIPAddress"
                .Fields("Property").Value = .Fields("Key").Value
                .Fields("Value").Value = objNode.DedicatedIPAddress
                .Update
    
                .AddNew
                .Fields("Category").Value = "MicrosoftNLB_Node"
                .Fields("ParentKey").Value = sNodeName
                .Fields("Key").Value = "Name"
                .Fields("Property").Value = .Fields("Key").Value
                .Fields("Value").Value = objNode.Name
                .Update
                
                .AddNew
                .Fields("Category").Value = "MicrosoftNLB_Node"
                .Fields("ParentKey").Value = sNodeName
                .Fields("Key").Value = "Status"
                .Fields("Property").Value = .Fields("Key").Value
                .Fields("Value").Value = SetNullToEmpty(objNode.Status, "(null)")
                .Update
                
                .AddNew
                .Fields("Category").Value = "MicrosoftNLB_Node"
                .Fields("ParentKey").Value = sNodeName
                .Fields("Key").Value = "StatusCode"
                .Fields("Property").Value = .Fields("Key").Value
                .Fields("Value").Value = objNode.StatusCode & ":" & StatusCodeDesc(objNode.StatusCode)
                .Update
            End With
        Next
    End If

    Dim objClusters As WbemScripting.SWbemObjectSet
    Set objClusters = SWbemServices.InstancesOf("MicrosoftNLB_ClusterSetting")

    Dim objCluster As WbemScripting.SWbemObject
    If Not objClusters Is Nothing Then
        For Each objCluster In objClusters
            sClusterName = GetComputerNameFromFullDNSName(objCluster.ClusterName)
            With v_rsServerInfo
                .AddNew
                .Fields("Category").Value = "MicrosoftNLB_ClusterSetting"
                .Fields("ParentKey").Value = sClusterName
                .Fields("Key").Value = "ClusterName"
                .Fields("Property").Value = .Fields("Key").Value
                .Fields("Value").Value = objCluster.ClusterName
                .Update
    
                .AddNew
                .Fields("Category").Value = "MicrosoftNLB_ClusterSetting"
                .Fields("ParentKey").Value = sClusterName
                .Fields("Key").Value = "ClusterIPAddress"
                .Fields("Property").Value = .Fields("Key").Value
                .Fields("Value").Value = objCluster.ClusterIPAddress
                .Update
    
                .AddNew
                .Fields("Category").Value = "MicrosoftNLB_ClusterSetting"
                .Fields("ParentKey").Value = sClusterName
                .Fields("Key").Value = "Name"
                .Fields("Property").Value = .Fields("Key").Value
                .Fields("Value").Value = objCluster.Name
                .Update
            End With
        Next
    End If
    
    Exit Function
eh:
    '//-----check for invalid namespace or not using NLB-----//
    If Err.Number = -2147217394 Or Err.Number = -2147217389 Then
        With v_rsServerInfo
            .AddNew
            .Fields("Category").Value = "MicrosoftNLB_Node"
            .Fields("ParentKey").Value = ""
            .Fields("Key").Value = "Error"
            .Fields("Property").Value = Err.Number
            .Fields("Value").Value = Err.Description
            .Update
        End With
 
    Else
        TrsRaiseError m_strCMODULENAME, "GetNLBNodeInfo", , , , , True
    End If
End Function

Private Function StatusCodeDesc(ByVal iStatusCode As Integer) As String

    Select Case iStatusCode
    Case 1005
       StatusCodeDesc = "WLBS_STOPPED"
    Case 1006
       StatusCodeDesc = "WLBS_CONVERGING"
    Case 1007
       StatusCodeDesc = "WLBS_CONVERGED"
    Case 1008
       StatusCodeDesc = "WLBS_DEFAULT"
    Case 1009
       StatusCodeDesc = "WLBS_DRAINING"
    Case 1013
       StatusCodeDesc = "WLBS_SUSPENDED"
    End Select
    
    Exit Function
End Function

Private Function IsNLBNodeActive(ByVal oNLBNode As WbemScripting.SWbemObject) As Boolean
    Dim rc As Boolean
    
    Select Case oNLBNode.StatusCode
    Case 1005
       rc = False '"WLBS_STOPPED"
    Case 1006
       rc = True '"WLBS_CONVERGING"
    Case 1007
       rc = True '"WLBS_CONVERGED"
    Case 1008
       rc = True '"WLBS_DEFAULT"
    Case 1009
       rc = False '"WLBS_DRAINING"
    Case 1013
       rc = False '"WLBS_SUSPENDED"
    Case Else
       rc = False
    End Select
    
    IsNLBNodeActive = rc
    
    Exit Function
End Function

Private Function CreateInfoRecordSet() As ADODB.Recordset
    Dim l_rsInfo As ADODB.Recordset
    
    On Error GoTo eh
    
    Set l_rsInfo = New ADODB.Recordset
    With l_rsInfo
        .Fields.Append "Category", adVarChar, 50
        .Fields.Append "ParentKey", adVarChar, 50, adFldMayBeNull
        .Fields.Append "Key", adVarChar, 50
        .Fields.Append "Property", adVarChar, 50
        .Fields.Append "Value", adVarChar, 255, adFldMayBeNull
        
        .CursorLocation = adUseClient
        .CursorType = adOpenDynamic
        .LockType = adLockBatchOptimistic
        .Open
    End With
    Set CreateInfoRecordSet = l_rsInfo
    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "CreateInfoRecordSet", , , , , True
End Function

Private Function GetNLBClusterName(Optional ByVal v_strServerName As String = ".") As String
'    Dim sNLBClusterName As String, l_strThisComputerName As String
'    Dim util As utilRecordCache.cuRecordCache
'    Dim varClusters As Variant
'    Dim l_varRecord As Variant
'    Dim l_varRowValues As Variant
'    Dim i As Integer, j As Integer
'    Dim lblnExists As Boolean
    Dim GetCOMCluster As Recordset
    
    GetNLBClusterName = ""
On Error GoTo eh:
'    l_strThisComputerName = Trim(stdmTRISCommon.ThisComputerName)
'    If UCase(v_strServerName) = UCase(l_strThisComputerName) Then
'        v_strServerName = "."
'    End If
    
    Dim l_sSQLStatement As String
    l_sSQLStatement = "select trim(parm_value) as COMCluster " & _
            "From methdba.ttr_sysparms " & _
            "Where parm_num = 47 " & _
            "and parm_suffix = 1 " & _
            "and GETVARIABLE('SYSIBM.SYSTEM_NAME') = 'DB2P'"

    Set GetCOMCluster = ExecuteSQL(l_sSQLStatement, adOpenStatic, adLockReadOnly)
    If RecordsExist(GetCOMCluster) Then
        GetCOMCluster.MoveFirst
        GetNLBClusterName = GetCOMCluster.Fields("COMCluster").Value
    End If
'    '//force retrieval for cache
'    l_varRecord = GetRecord(Array("TRISPRODNLB"), NLBCLUSTER_CACHE, lblnExists)
'
'    Set util = CreateObject("utilRecordCache.cuRecordCache")
'    varClusters = util.GetKeys(NLBCLUSTER_CACHE)
'//keys not being returned so cannot use
'    If Not IsEmpty(varClusters) Then
'        For i = LBound(varClusters) To UBound(varClusters)
'            sNLBClusterName = varClusters(i)
'
'            l_varRecord = GetRecord(Array(sNLBClusterName), NLBCLUSTER_CACHE, lblnExists)
'
'            If Not IsEmpty(l_varRecord) Then
'                l_varRowValues = l_varRecord(NLBCLUSTER_CACHE_ROWVALUES)
'                'ReDim l_varGetNLBClusterMembers(UBound(l_varRowValues, 2))
'                For j = LBound(l_varRowValues, 2) To UBound(l_varRowValues, 2)
'                    'l_varGetNLBClusterMembers(i) = l_varRowValues(NLBCLUSTER_CACHE_CLUSTERMEMBER, i)
'                    If l_varRowValues(NLBCLUSTER_CACHE_CLUSTERMEMBER, j) = v_strServerName Then
'                        GetNLBClusterName = UCase(sNLBClusterName)
'                        Exit Function
'                    End If
'                Next
'            End If
'        Next
'    End If
'
'    GetNLBClusterName = ""
    
    Exit Function
eh:
    GetNLBClusterName = ""
    TrsRaiseError m_strCMODULENAME, "GetNLBClusterName", , , , DescArgsFrom(v_strServerName), True
End Function

Private Function GetNLBClusterMembers(ByVal v_sClusterName As String) As Variant
    Dim l_varRecord As Variant
    Dim lblnExists As Boolean
    Dim i As Long, l_varRowValues As Variant, l_varGetNLBClusterMembers As Variant
    
On Error GoTo eh
    
    l_varRecord = GetRecord(Array(v_sClusterName), NLBCLUSTER_CACHE, lblnExists)
    
    If Not IsEmpty(l_varRecord) Then
        l_varRowValues = l_varRecord(NLBCLUSTER_CACHE_ROWVALUES)
        ReDim l_varGetNLBClusterMembers(UBound(l_varRowValues, 2))
        For i = LBound(l_varRowValues, 2) To UBound(l_varRowValues, 2)
            l_varGetNLBClusterMembers(i) = l_varRowValues(NLBCLUSTER_CACHE_CLUSTERMEMBER, i)
        Next
    End If
    
    GetNLBClusterMembers = l_varGetNLBClusterMembers

    Exit Function
eh:
    TrsRaiseError m_strCMODULENAME, "GetNLBClusterMembers"
End Function

'*********************************************************
' Begin Public Cache management methods
'*********************************************************

Public Function GetGroupRecords( _
                    ByVal v_strGroupName As String, _
                    ByVal v_varKeys As Variant, _
                    ByRef r_varKeyArray As Variant, _
                    ByRef r_varColumnArray As Variant) As ADODB.Recordset
    Dim l_rsRecords As ADODB.Recordset
    Dim l_colErrArgs As Collection
On Error GoTo eh

    Select Case v_strGroupName
        Case NLBCLUSTER_CACHE
            Set l_rsRecords = GetNLBClusterRecords(v_varKeys)
            r_varKeyArray = Array("CLUSTERNAME")
            r_varColumnArray = Empty
    End Select
    Set GetGroupRecords = l_rsRecords

    Set l_rsRecords = Nothing
    Exit Function
eh:
    Set l_colErrArgs = New Collection
    l_colErrArgs.Add "v_strGroupName = " & v_strGroupName
    l_colErrArgs.Add "v_varKeys = " & v_varKeys
    l_colErrArgs.Add "l_rsRecords.Source = " & l_rsRecords.Source
    
    Set l_rsRecords = Nothing
    TrsRaiseError m_strCMODULENAME, "GetGroupRecords", , , , l_colErrArgs, True
End Function

'*********************************************************
' Begin Private Caching methods
'*********************************************************

Private Function GetRecord( _
            ByVal v_varKey As Variant, _
            ByVal v_strGroupName As String, _
            ByRef r_blnExists As Boolean) As Variant
' Answer an variant array, representing a record, or empty.
' This method should not invoke InitializeGroup unless the
' group is setup for bulk initialization.  If setup for lazy
' initialization it should only get the single record and add it
' to the cache.  If the record is not found, using lazy, it should
' add a value to the cache to denote invalid so that it does not
' attempt to retrieve it from the database again.  Currently, a string
' constant is added for record not found and the spmCache component
' will answer that it exists, but the property value is a
' RECORDNOTFOUND string.
'
    Dim recCache As utilRecordCache.cuRecordCache
    Dim l_colErrArgs As Collection
           
On Error GoTo eh

'//--------------------------------------------------------------------------------------------------------------//
'//-----http://msdn.microsoft.com/library/default.asp?url=/library/en-us/cossdk/htm/imtxas_2bs6.asp
'//-----Do not use the SafeRef method in COM+. In MTS, SafeRef was used by an object to obtain
'//-----a reference to itself that could be passed outside its context wrapper.
'//-----With contexts and interception fully integrated into COM+, this is no longer necessary.
'//-----
'//-----[C/C++]
'//-----Note   The C++ function SafeRef should be used instead of this method.
'//--------------------------------------------------------------------------------------------------------------//

        Set recCache = CreateObject("utilRecordCache.cuRecordCache")
    
        If IsArray(v_varKey) Then
            GetRecord = recCache.GetRecord2(v_varKey, v_strGroupName, _
                                IsLazyInitialized(v_strGroupName), _
                                GroupSupportsLoadRecord(v_strGroupName), _
                                OneRowForEachKey(v_strGroupName), _
                                r_blnExists, Me)
        Else
            GetRecord = recCache.GetRecord(CStr(v_varKey), v_strGroupName, _
                                IsLazyInitialized(v_strGroupName), _
                                GroupSupportsLoadRecord(v_strGroupName), _
                                OneRowForEachKey(v_strGroupName), _
                                r_blnExists, Me)
        End If
        
        Set recCache = Nothing
    
    Exit Function
eh:
    Set recCache = Nothing

    Set l_colErrArgs = New Collection
    If IsArray(v_varKey) Then
        l_colErrArgs.Add "v_varKey() = (" & Join(v_varKey, ", ") & ")"
    Else
        l_colErrArgs.Add "v_varKey = " & CStr(v_varKey)
    End If
    l_colErrArgs.Add "v_strGroupName = " & v_strGroupName
    
    TrsRaiseError m_strCMODULENAME, "GetRecord", , , , l_colErrArgs, True
End Function

''//-----code has been factored into utilRecordCache-----//
'Private Sub InvalidateCache( _
'                ByVal v_varKey As Variant, _
'                ByVal v_strGroupName As String)
'
'    Dim recCache As utilRecordCache.cuRecordCache
'    Dim l_colErrArgs As Collection
'    Dim l_blnExists As Boolean
'On Error GoTo eh
'
'    Set recCache = CreateObject("utilRecordCache.cuRecordCache")
'
'    If IsArray(v_varKey) Then '//-----uses metadata to format key, may need to retrieve data to get metadata
'        recCache.InvalidateCache2 v_varKey, v_strGroupName, _
'                                IsLazyInitialized(v_strGroupName), _
'                                GroupSupportsLoadRecord(v_strGroupName), _
'                                OneRowForEachKey(v_strGroupName), _
'                                l_blnExists, Me
'    Else
'        recCache.InvalidateCache CStr(v_varKey), v_strGroupName
'    End If
'
'    Set recCache = Nothing
'
'    Exit Sub
'eh:
'    Set recCache = Nothing
'
'    Set l_colErrArgs = New Collection
'    If IsArray(v_varKey) Then
'        l_colErrArgs.Add "v_varKey() = (" & Join(v_varKey, ", ") & ")"
'    Else
'        l_colErrArgs.Add "v_varKey = " & CStr(v_varKey)
'    End If
'    l_colErrArgs.Add "v_strGroupName = " & v_strGroupName
'
'    TrsRaiseError m_strCMODULENAME, "InvalidateCache", , , , l_colErrArgs, True
'End Sub

'//-----used by RecordCache-----//

Private Function IsLazyInitialized(ByVal v_strGroupName As String) As Boolean
' If group is not lazy initialized, then it is bulk initialized.
On Error GoTo eh
'    Select Case Trim(v_strGroupName)
'    Case CJB_GROUPNAME
'        IsLazyInitialized = True
'    Case Else
        IsLazyInitialized = False
'    End Select
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "IsLazyInitialized"
End Function

Private Function GroupSupportsLoadRecord(ByVal v_strGroupName As String) As Boolean
'All group caches that support Invalidate must use parameterized SQL in corresponding GetxxxRecords (GetGroupRecords)
On Error GoTo eh
'    Select Case Trim(v_strGroupName)
'    Case NLBCLUSTER_CACHE
        GroupSupportsLoadRecord = True
'    Case Else
'        GroupSupportsLoadRecord = False
'    End Select
    Exit Function

eh:
    TrsRaiseError m_strCMODULENAME, "GroupSupportsLoadRecord"
End Function

Private Function OneRowForEachKey(ByVal v_strGroupName As String) As Boolean
'//-----normally each key/row is cached into a 1-dimensional array that represents each column-----//
'//-----union components will have multiple rows per key and will be cached into a 2-dimensional array-----//
On Error GoTo eh
'    Select Case Trim(v_strGroupName)
'    Case NLBCLUSTER_CACHE
        OneRowForEachKey = False
'    Case Else
'        OneRowForEachKey = True
'    End Select
    Exit Function
    
eh:
    TrsRaiseError m_strCMODULENAME, "OneRowForEachKey"
End Function

''//-----use RULE 127 to find user membership in group and add to user recordset-----//
'Private Function GetNLBClusterMemberRecords(ByVal v_sClusterName As String, _
'                                    ByRef r_rsUsers As ADODB.Recordset) As ADODB.Recordset
'    Dim rsCluster As ADODB.Recordset
'    Dim l_colErrArgs As Collection
'On Error GoTo eh
'
'        Set rsCluster = GetNLBClusterRecords(v_sClusterName)
'
'    Exit Function
'eh:
'    Set l_colErrArgs = New Collection
'    l_colErrArgs.Add "v_sClusterName = " & v_sClusterName
'    If RecordsExist(rsCluster) Then
'        rsCluster.MoveFirst
'        l_colErrArgs.Add "rsCluster = " & rsCluster.GetString
'    End If
'
'    TrsRaiseError m_strCMODULENAME, "GetNLBClusterMemberRecords", , , , l_colErrArgs, True
'End Function

'//-----Determine NLB Cluster/Cluster Members by using Rule 127
Private Function GetNLBClusterRecords(Optional ByVal v_varKey As Variant) As ADODB.Recordset
    On Error GoTo eh
    
    Dim l_sSQLStatement As String

    'Set SQL SELECT statment
    l_sSQLStatement = _
        "SELECT " & _
                "CASE WHEN LOCATE(':',RG_VALUE_CD) > 0  " & _
                    "  THEN RTRIM(SUBSTR(RG_VALUE_CD,1,LOCATE(':',RG_VALUE_CD,1)-1)) " & _
                    "  ELSE '' END AS CLUSTERNAME " & _
                ",CASE WHEN LOCATE(':',RG_VALUE_CD) > 0  " & _
                    "  THEN RTRIM(SUBSTR(RG_VALUE_CD,LOCATE(':',RG_VALUE_CD,1)+1,LENGTH(RG_VALUE_CD))) " & _
                "  ELSE '' END AS CLUSTERMEMBER " & _
        " FROM METHDBA.TTR_RG_GROUP " & _
        " WHERE RG_CD = 127 "
        
    If IsArray(v_varKey) Then
    ElseIf IsEmpty(v_varKey) Then
    Else
        l_sSQLStatement = l_sSQLStatement & _
                "AND CASE WHEN LOCATE(':',RG_VALUE_CD) > 0  " & _
                "  THEN RTRIM(SUBSTR(RG_VALUE_CD,1,LOCATE(':',RG_VALUE_CD,1)-1)) " & _
                "  ELSE '' END = '" & CStr(v_varKey) & "'"
    End If
    
    Set GetNLBClusterRecords = ExecuteSQL(l_sSQLStatement, adOpenStatic, adLockReadOnly)
                                
    Exit Function
eh:
    'Call the standard module function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME, "GetNLBClusterRecords"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "GetNLBClusterRecords" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(l_sSQLStatement) _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "GetNLBClusterRecords" _
                    , trErrdb_TRISAvailable_F _
                    , , , _
                    , True
    End Select
End Function

'//-----Get info about db2
Private Function GetDB2Info() As ADODB.Recordset
    On Error GoTo eh
    
    Dim l_sSQLStatement As String

    l_sSQLStatement = GetStringResource("GetDB2Info.sql")
    
    Set GetDB2Info = ExecuteSQL(l_sSQLStatement, adOpenStatic, adLockReadOnly)
                                
    Exit Function
eh:
    'Call the standard module function
    'Raise the appropriate error message depending on the error number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME, "GetDB2Info"
            Case trErrdb_ExecuteSQL_F
                TrsRaiseError MODULENAME _
                    , "GetDB2Info" _
                    , trErrdb_ExecuteSQL_F _
                    , , _
                    , DescArgsFrom(l_sSQLStatement) _
                    , True
            Case Else
                TrsRaiseError MODULENAME _
                    , "GetDB2Info" _
                    , trErrdb_TRISAvailable_F _
                    , , , _
                    , True
    End Select
End Function

'copied
Private Function GetStringResource(ByVal v_strFilename As String) As String
    Dim oFS As New FileSystemObject
    Dim ts As TextStream
    Dim sFilename As String, sFileContents As String
    
On Error GoTo eh
    sFilename = oFS.BuildPath("C:\TRIS SOLUTION\dbScripts\utilServerInfo", v_strFilename)
    
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

'//-----for testing errors being raised back through the Gateway/WebService
Public Function RaiseError(ByVal v_lngErrorNumber As Long, _
                            ByVal v_strSource As String, _
                            ByVal v_strDescription As String) As Long
'//On Error GoTo eh

    Err.Raise v_lngErrorNumber, v_strSource, v_strDescription

Exit Function
'//eh:
    '//TrsRaiseError MODULENAME, "RaiseError", trErrdb_TRISAvailable_F, , , , True
End Function

