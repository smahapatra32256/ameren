VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cuClientInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const m_strCMODULENAME As String = "cuClientInfo"


Public Function GetClientInfo(Optional ByVal v_strCategory As String = "") As ADODB.Recordset
    Dim l_rsInfo As ADODB.Recordset
    Dim strComputer As String
    Dim objWMIService As Object
    Dim sCategory As String
    Dim colSettings As Object
    Dim objComputer As Object
    Dim objOperatingSystem As Object
    Dim colItems As Object
    Dim objItem As Object
    Dim colChassis As Object
    Dim objChassis As Object
    Dim colInstalledPrinters As Object
    Dim objPrinter As Object
    
'On Error GoTo eh
On Error Resume Next

    v_strCategory = LCase(Trim(v_strCategory))
    
    'Set l_rsInfo = New ADODB.Recordset
    Set l_rsInfo = CreateInfoRecordSet
    With l_rsInfo
'        .Fields.Append "Category", adVarChar, 20
'        .Fields.Append "Property", adVarChar, 40
'        .Fields.Append "Value", adVarChar, 255
'
'        .CursorLocation = adUseClient
'        .CursorType = adOpenDynamic
'        .LockType = adLockBatchOptimistic
'        .Open

        If v_strCategory = "" Or v_strCategory = "all" Then
    
            strComputer = "."
            
            Set objWMIService = GetObject("winmgmts:" _
                & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
    
            sCategory = "Client Computer"
            Set colSettings = objWMIService.ExecQuery("Select * from Win32_ComputerSystem")
            For Each objComputer In colSettings
                AddClientInfo l_rsInfo, sCategory, "", "", "Computer Domain", objComputer.Domain
                AddClientInfo l_rsInfo, sCategory, "", "", "Computer Name", objComputer.Name
                AddClientInfo l_rsInfo, sCategory, "", "", "Computer IP", GetIPAddress(objComputer.Name)
                AddClientInfo l_rsInfo, sCategory, "", "", "Computer Manufacturer", objComputer.Manufacturer
                AddClientInfo l_rsInfo, sCategory, "", "", "Computer Model", objComputer.Model
'                AddClientInfo l_rsInfo, sCategory, "","", objComputer.UserName
                AddClientInfo l_rsInfo, sCategory, "", "", "Computer UserName", IIf(IsNull(objComputer.UserName), "null", objComputer.UserName)
            Next
            AddClientInfo l_rsInfo, "Client runtime", "", "", "Command Line", Command
        End If
        
        If v_strCategory = "all" Then
            sCategory = "Client OS"
            Set colSettings = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
            For Each objOperatingSystem In colSettings
                AddClientInfo l_rsInfo, sCategory, "", "", "OS Caption", objOperatingSystem.Caption
                AddClientInfo l_rsInfo, sCategory, "", "", "OS Version", objOperatingSystem.Version
                AddClientInfo l_rsInfo, sCategory, "", "", "OS Service Pack", objOperatingSystem.ServicePackMajorVersion _
                                                                 & "." & objOperatingSystem.ServicePackMinorVersion
                AddClientInfo l_rsInfo, sCategory, "", "", "Free Physical Memory", objOperatingSystem.FreePhysicalMemory
                AddClientInfo l_rsInfo, sCategory, "", "", "Total Visible Memory Size", objOperatingSystem.TotalVisibleMemorySize
            Next
            
            sCategory = "Client Processor"
            Set colItems = objWMIService.ExecQuery("Select * from Win32_Processor")
            For Each objItem In colItems
                AddClientInfo l_rsInfo, sCategory, Trim(objItem.Name), "", "CurrentClockSpeed", objItem.CurrentClockSpeed
                AddClientInfo l_rsInfo, sCategory, Trim(objItem.Name), "", "Ext Clock", objItem.ExtClock
                AddClientInfo l_rsInfo, sCategory, Trim(objItem.Name), "", "Maximum Clock Speed", objItem.MaxClockSpeed
                AddClientInfo l_rsInfo, sCategory, Trim(objItem.Name), "", "Processer Name", Trim(objItem.Name)
            Next
    
            sCategory = "Client Hardware"
            Set colChassis = objWMIService.ExecQuery("Select * from Win32_SystemEnclosure")
            For Each objChassis In colChassis
                AddClientInfo l_rsInfo, sCategory, "", "", "Chassis Type", objChassis.ChassisTypes(0)
'                For Each objItem In objChassis.ChassisTypes
'                    AddClientInfo l_rsInfo, sCategory, "Chassis Type", objItem.ChassisTypes
'                Next
            Next
    
            Set colItems = objWMIService.ExecQuery("Select * from Win32_DesktopMonitor")
            For Each objItem In colItems
                AddClientInfo l_rsInfo, sCategory, "", "", "Screen Height", objItem.ScreenHeight
                AddClientInfo l_rsInfo, sCategory, "", "", "Screen Width", objItem.ScreenWidth
            Next
    
            Set colInstalledPrinters = objWMIService.ExecQuery("Select * from Win32_Printer Where Default = True")
            For Each objPrinter In colInstalledPrinters
                AddClientInfo l_rsInfo, sCategory, objPrinter.Name, "", "Printer Name", objPrinter.Name
                AddClientInfo l_rsInfo, sCategory, objPrinter.Name, "", "Printer Driver Name", objPrinter.DriverName
                AddClientInfo l_rsInfo, sCategory, objPrinter.Name, "", "Printer Processor", objPrinter.PrintProcessor
            Next
        End If
    End With
    
    Set GetClientInfo = l_rsInfo
    On Error GoTo 0
    Exit Function
eh:
'    MsgBox Err.Number
'    If Err.Number = 94 Or Err.Number = -2147217385 Or Err.Number = 91 Then
'        Resume Next
'    End If
    Err.Raise Err.Number _
        , m_strCMODULENAME & ": GetClientInfo " & sCategory, Err.Description
End Function

Private Function AddClientInfo(ByRef rs As ADODB.Recordset, _
                                ByVal sCategory As String, _
                                ByVal sParentKey As String, _
                                ByVal sKey As String, _
                                ByVal sProperty As String, _
                                ByVal sValue As String _
                                )
    On Error Resume Next
    
    rs.AddNew
    rs.Fields("Category").Value = sCategory
    rs.Fields("ParentKey").Value = sParentKey
    If sKey <> "" Then
        rs.Fields("Key").Value = sKey
    Else
        rs.Fields("Key").Value = sProperty
    End If
    rs.Fields("Property").Value = sProperty
    
'    MsgBox sCategory & vbTab & sProperty, , "AddClientInfo"
'    If IsNull(sValue) Then
'        MsgBox "sValue IsNull" & sCategory & vbTab & sProperty, , "AddClientInfo"
'    End If
    
    rs.Fields("Value").Value = sValue '//IIf(sValue <> "", sValue, "null")
    rs.Update
    
    On Error GoTo 0
    Exit Function
eh:
'    MsgBox m_strCMODULENAME & ": AddClientInfo " & sCategory & "  " & sProperty & "  " & sValue
    
    Err.Raise Err.Number _
        , m_strCMODULENAME & ": AddClientInfo " & sCategory & "  " & sProperty & "  " & sValue, Err.Description
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
    Err.Raise Err.Number, m_strCMODULENAME & ": CreateInfoRecordSet"
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
'        TrsTraceError m_strCMODULENAME, "GetIPAddress", , , , DescArgsFrom(v_strServerName), True
    End If
End Function

