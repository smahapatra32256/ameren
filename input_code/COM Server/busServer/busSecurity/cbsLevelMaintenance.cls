VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "cbsLevelMaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'AMEREN Project: TRIS Conversion Phase 2
'March 2002
'Author: Prem M. Chander
'Module Description: This class will provide the user the ability
'to access security level information

Option Explicit

'Private Const ADOFIELDATTR = adFldUpdatable Or adFldIsNullable
Private Const MODULENAME As String = "busSecurity.cbsLevelMaintenance"
'30206, 30207, 30208, 30209, 30210
'The following retrieves all (level, employee, ABF info) information
'for a specific security level
Public Function GetAllSecurityLevelInformation(ByVal v_lSecurityLevel As Long _
                    , ByVal v_bActiveEmployeesOnly As Boolean _
                    , ByRef r_adrsSecurityLevelInfo As ADODB.Recordset _
                    , ByRef r_adrsEmployeesInLevel As ADODB.Recordset _
                    , ByRef r_adrsAuthorizedBusFuns As ADODB.Recordset _
                    , ByVal vlTRISUserNumber As Long _
                    ) As Boolean
On Error GoTo GetAllSecurityLevelInformation_EH
Dim lobjDataAccess As dbSecurity.cdLevelMaintenance
Dim lobjdbFunSecurity As dbSecurity.cdFunSecurity
Dim ldtLoadFromDate As Date, lbSystemAdmin As Boolean

    GetAllSecurityLevelInformation = False

    'Determine system admin
    lbSystemAdmin = IsASystemAdministrator(vlTRISUserNumber)

    If (v_bActiveEmployeesOnly) Then
        ldtLoadFromDate = CDate(Format(DateAdd("m", -6, Date) _
                                , "mm/dd/yyyy"))
    Else
        ldtLoadFromDate = #1/1/1900#
    End If

    'Construct the Data Access Object
    Set lobjDataAccess = CreateObject( _
                            "dbSecurity.cdLevelMaintenance")
    Set r_adrsSecurityLevelInfo = lobjDataAccess.db_GetLevels( _
                                    GETSECLVLINFOBYLEVEL _
                                    , v_lSecurityLevel _
                                    , , lbSystemAdmin)
    Set r_adrsEmployeesInLevel = lobjDataAccess.db_GetEmployeesForLevel( _
                                    v_lSecurityLevel _
                                    , v_bActiveEmployeesOnly _
                                    , ldtLoadFromDate _
                                    , lbSystemAdmin)
    Set r_adrsAuthorizedBusFuns = lobjDataAccess.db_GetAuthrzBusFuncsForLevel( _
                                    v_lSecurityLevel _
                                    , lbSystemAdmin)
    'destroy
    Set lobjDataAccess = Nothing

    GetAllSecurityLevelInformation = True

    Exit Function
GetAllSecurityLevelInformation_EH:
    'destroy
    Set lobjDataAccess = Nothing

    'Raise Error depending on the Error Number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "GetAllSecurityLevelInformation"
            Case Else
                TrsRaiseError MODULENAME _
                    , "GetAllSecurityLevelInformation" _
                    , trErrAllSecLvlInfo_F _
                    , , , , _
                    True
    End Select

End Function
'The following retrieves all security levels or for a
'specific division or level
Public Function GetSecurityLevels(ByVal v_iToGet As Integer _
                    , ByVal vlTRISUserNumber As Long _
                    , Optional ByVal v_lSecurityLevel As Long _
                    , Optional ByVal v_sDivision As String _
                    ) As ADODB.Recordset
On Error GoTo GetSecurityLevels_EH
Dim lobjDataAccess As dbSecurity.cdLevelMaintenance
Dim lbSystemAdmin As Boolean

    'Determine system admin
    lbSystemAdmin = IsASystemAdministrator(vlTRISUserNumber)

    'Construct the Data Access Object
    Set lobjDataAccess = CreateObject( _
                            "dbSecurity.cdLevelMaintenance")
    'Call data access method
    Select Case (v_iToGet)
            Case GETALLSECLVLINFO
                Set GetSecurityLevels = lobjDataAccess.db_GetLevels( _
                                        GETALLSECLVLINFO _
                                        , , , lbSystemAdmin)
            Case GETSECLVLINFOBYLEVEL
                Set GetSecurityLevels = lobjDataAccess.db_GetLevels( _
                                        GETSECLVLINFOBYLEVEL _
                                        , v_lSecurityLevel _
                                        , , lbSystemAdmin)
            Case GETALLSECLVLINFOBYDIVISION
                Set GetSecurityLevels = lobjDataAccess.db_GetLevels( _
                                        GETALLSECLVLINFOBYDIVISION _
                                        , _
                                        , v_sDivision _
                                        , lbSystemAdmin)
    End Select
    'destroy
    Set lobjDataAccess = Nothing

    Exit Function
GetSecurityLevels_EH:
    'destroy
    Set lobjDataAccess = Nothing

    'Raise Error depending on the Error Number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "GetSecurityLevels"
            Case Else
                TrsRaiseError MODULENAME _
                    , "GetSecurityLevels" _
                    , trErrSecLvl_F _
                    , , , , _
                    True
    End Select

End Function
'The following retrieves employees assigned to a security level
Public Function GetEmployeesAssignedToLevel(ByVal v_lSecurityLevel As Long _
                    , ByVal v_bActiveOnly As Boolean _
                    , ByVal vlTRISUserNumber As Long _
                    ) As ADODB.Recordset
On Error GoTo GetEmployeesAssignedToLevel_EH
Dim lobjDataAccess As dbSecurity.cdLevelMaintenance
Dim ldtLoadFromDate As Date, lbSystemAdmin As Boolean

    'Determine system admin
    lbSystemAdmin = IsASystemAdministrator(vlTRISUserNumber)

    If (v_bActiveOnly) Then
        ldtLoadFromDate = CDate(Format(DateAdd("m", -6, Date) _
                                , "mm/dd/yyyy"))
    Else
        ldtLoadFromDate = #1/1/1900#
    End If

    'Construct the Data Access Object
    Set lobjDataAccess = CreateObject( _
                            "dbSecurity.cdLevelMaintenance")
    Set GetEmployeesAssignedToLevel = lobjDataAccess.db_GetEmployeesForLevel( _
                                        v_lSecurityLevel _
                                        , v_bActiveOnly _
                                        , ldtLoadFromDate _
                                        , lbSystemAdmin)
    'destroy
    Set lobjDataAccess = Nothing

    Exit Function
GetEmployeesAssignedToLevel_EH:
    'destroy
    Set lobjDataAccess = Nothing

    'Raise Error depending on the Error Number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "GetEmployeesAssignedToLevel"
            Case Else
                TrsRaiseError MODULENAME _
                    , "GetEmployeesAssignedToLevel" _
                    , trErrEmpsAssignToLvl_F _
                    , , , , _
                    True
    End Select

End Function
'The following retrieves authorized business functions for a security level
Public Function GetAuthorizedBusinessFunctionsForLevel(ByVal v_lSecurityLevel As Long _
                , ByVal vlTRISUserNumber As Long _
                ) As ADODB.Recordset
On Error GoTo GetAuthorizedBusinessFunctionsForLevel_EH
Dim lobjDataAccess As dbSecurity.cdLevelMaintenance
Dim lbSystemAdmin As Boolean

    'Determine system admin
    lbSystemAdmin = IsASystemAdministrator(vlTRISUserNumber)

    'Construct the Data Access Object
    Set lobjDataAccess = CreateObject( _
                            "dbSecurity.cdLevelMaintenance")
    Set GetAuthorizedBusinessFunctionsForLevel = lobjDataAccess.db_GetAuthrzBusFuncsForLevel( _
                                        v_lSecurityLevel _
                                        , lbSystemAdmin)
    'destroy
    Set lobjDataAccess = Nothing

    Exit Function
GetAuthorizedBusinessFunctionsForLevel_EH:
    'destroy
    Set lobjDataAccess = Nothing

    'Raise Error depending on the Error Number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "GetAuthorizedBusinessFunctionsForLevel"
            Case Else
                TrsRaiseError MODULENAME _
                    , "GetAuthorizedBusinessFunctionsForLevel" _
                    , trErrBusFuncAssignToLvl_F _
                    , , , , _
                    True
    End Select

End Function
Public Function GetAllAvailableBusFuncs(ByVal vlTRISUserNumber As Long _
                    ) As ADODB.Recordset
On Error GoTo GetAllAvailableBusFuncs_EH
Dim lobjDataAccess As dbSecurity.cdLevelMaintenance
Dim lbSystemAdmin As Boolean

    'Determine system admin
    lbSystemAdmin = IsASystemAdministrator(vlTRISUserNumber)

    'Construct the Data Access Object
    Set lobjDataAccess = CreateObject( _
                            "dbSecurity.cdLevelMaintenance")
    Set GetAllAvailableBusFuncs = lobjDataAccess.db_GetAvailableBusFuncs( _
                                    lbSystemAdmin)
    'destroy
    Set lobjDataAccess = Nothing

    Exit Function
GetAllAvailableBusFuncs_EH:
    'destroy
    Set lobjDataAccess = Nothing

    'Raise Error depending on the Error Number
    Select Case (Err.Number)
            Case trErrutOpenConnection
                TrsRaiseError MODULENAME _
                    , "GetEmployeesAssignedToLevel"
            Case Else
                TrsRaiseError MODULENAME _
                    , "GetEmployeesAssignedToLevel" _
                    , trErrAvailBusFuncs_F _
                    , , , , _
                    True
    End Select

End Function
'Begin TRIS SOX July 2004
Private Function IsASystemAdministrator(ByVal vlTRISUserNumber As Long _
                    ) As Boolean
On Error GoTo IsASystemAdministrator_EH
Dim ladrsEmployee As ADODB.Recordset
Dim lobjdbFunSecurity As dbSecurity.cdFunSecurity

    IsASystemAdministrator = False

    If Not (CheckForValidLong(vlTRISUserNumber, True)) Then
        Err.Raise trErrAllSecLvlInfo_F
        Exit Function
    Else
        Set lobjdbFunSecurity = CreateObject( _
                                "dbSecurity.cdFunSecurity")
        Set ladrsEmployee = lobjdbFunSecurity.db_CheckEmployee( _
                                vlTRISUserNumber)
        Set lobjdbFunSecurity = Nothing

        If Not (ladrsEmployee Is Nothing) Then
            With ladrsEmployee
                If Not (.EOF And .BOF) Then
                    If (!SEC_LEVEL = 10) Then
                        IsASystemAdministrator = True
                    End If
                End If
            End With
        End If
    End If

    Set ladrsEmployee = Nothing
    Set lobjdbFunSecurity = Nothing

    Exit Function
IsASystemAdministrator_EH:
    Set ladrsEmployee = Nothing
    Set lobjdbFunSecurity = Nothing

    TrsRaiseError MODULENAME _
        , "IsASystemAdministrator"

End Function
'End TRIS SOX July 2004

'FROM TRIS 16-BIT
''Select existing levels for a given Division
'SELECT  SEC_LEVEL, SEC_LEVEL_DESC
'        , TIME_STAMP
'FROM   METHDBA.TTR_SECURITY_LEVEL
'WHERE  DIVISION_NUM = Char
'ORDER BY SEC_LEVEL

''Insert new level
'INSERT INTO METHDBA.TTR_SECURITY_LEVEL
'(SEC_LEVEL, SEC_LEVEL_DESC, LAST_MOD_EMP_NUM, DIVISION_NUM)
'VALUES
'(long, char, long, char)

''Update description
'UPDATE METHDBA.TTR_SECURITY_LEVEL
'SET    SEC_LEVEL_DESC = Char
'       , LAST_MOD_EMP_NUM = long
'       , TIME_STAMP = char
'WHERE  SEC_LEVEL = long
'AND    TIME_STAMP = char

''Check for emps assigned to sec level
'SELECT COUNT(*)
'FROM   METHDBA.TTR_EMPLOYEE
'WHERE  SEC_LEVEL = long

''Check for restrictions applied to the Level
'SELECT COUNT(*)
'FROM   METHDBA.TTR_SECURITY_RESTR
'WHERE  SEC_LEVEL = long

''Delete
'DELETE FROM    METHDBA.TTR_SECURITY
'WHERE  SEC_LEVEL = long
'AND    DIVISION_NUM = char
'AND    TIME_STAMP = char

'existing system allows user to determine the new
'SEC_LEVEL number.

'ASSIGN RESTRICTIONS
'SELECT DISTINCT    APPL_NAME
'FROM               METHDBA.TTR_SECURITY_COMP
'ORDER BY APPL_NAME

