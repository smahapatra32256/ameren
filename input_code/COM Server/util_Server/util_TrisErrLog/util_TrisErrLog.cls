VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsuTrisErrLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const gsModule_Name As String = "util_TrisErrLog.clsuTrisErrLog" & "-"

'==========================================================================
'
'       mstrHldAppExeName Name: util_TrisErrLog
'mstrHldAppExeName Description: Routines for debug and log file output handling
'
'===========================================================================
'trmsMSG_FIRSTAPPEND is used as an id to indicate if this is the first or the next append
'trmsMSG_FIRSTAPPEND size is 17 =    12345678901234567
Private Const trmsMSG_FIRSTAPPEND = "========>BEGIN LOGMSG: "
Private Const trmsMSG_NEXTAPPEND = "========>APPEND "
Private Const trmsMSG_WARNRESFILEMISSING = "(1.Resource File may be missing in utilTrisErrLog-cuTrisErrLog-BuildDescriptionFromResFile_EH): "
Private Const trmsMSG_WARNFILEFAILED = "(2.Open/Write/Close File may have failed in utilTrisErrLog-cuTrisErrLog-LogFile_EH): "
Private Const trmsMSG_WARNRUNERROR = "(3.Error in utilTrisErrLog-cuTrisErrLog-ErrTrace_EH): "
Private Const trmsMSG_WARNRUNERRORP = "(4.Error in utilTrisErrLog-cuTrisErrLog-RecycleBackupLogFile_EH): "
Private Const trmsMSG_WARNRUNERRORK = "(5.Error in utilTrisErrLog-cuTrisErrLog-KillFile_EH): "
Private Const trmsMSG_WARNRUNERRORL = "(6.Error in utilTrisErrLog-cuTrisErrLog-ServerEventLog_EH): "
Private Const trmsMSG_WARNDESCRUNERROR = "(7.Error in utilTrisErrLog-cuTrisErrLog-ProcessResourceDesc_EH): "
Private Const trmsMSG_WARNGETFILEERROR = "(8.Error in utilTrisErrLog-cuTrisErrLog-GetDayOfWeekLogFiles_EH): "
Private Const trmsMSG_WARNCALLHELP = " Please provide this information to the help desk: "
Private Const trmiEXPIREDAYS = 6
Private Const trmsEXPIREDAYTYPE = "d"
Private mstrHldLbkFilePath As String    'Path and name for backup of expired log file
Private mstrHldFilePath As String       'Path and name for log file
Private mstrHldAppExeName As String     'mstrHldAppExeName name to log
Private mstrHldSourceRoutine As String 'mstrHldAppExeName mstrHldSourceRoutineto log
Private mstrHldUserId As String         'vstrUserId name to log
Private mstrHldErrDescription As String 'errdescription to log
Private mstrHldFormatDescription As String 'errdescription to log
Private mstrHldErrSource As String      'errsource to log
Private mlngHldErrNumber As Long        'long error number to log
Private mvntHldErrArg As Variant
Private mvntHldDescriptionArg As Variant
Private mstrHldResDelimiter As String
Private mintHldDescNdx As Integer
Private mstrHldLine As String
Private mstrHldArgLine As String
Private mstrHldWriteLine As String
Private mintHldFileNum As Integer
Private mstrHldMsg As String
Private miostrErrDescription As String

Public Function GetErrDescription( _
        ByVal vlngAppErr As Long, _
        Optional ByVal vcolDescriptionArgs As Collection) As String
    
'This routine provides the error description build from RES file and optional parameters
'It initializes all the module variables
'Builds the error description which may include multiple optional parameters


    Initializemvars
    
    ProcessResourceDesc vlngAppErr, vcolDescriptionArgs
    
    GetErrDescription = mstrHldErrDescription
    
End Function

Public Function ErrTrace _
   (ByVal vlngAppErr As Long, _
    ByVal vlngVBErr As Long, _
    iostrErrDescription As String, _
    iostrErrSource As String, _
    ByVal vstrUserId As String, _
    ByVal vstrSourceRoutine As String, _
    Optional ByVal vblnLog As Boolean = False, _
    Optional ByVal vcolErrArgs As Collection, _
    Optional ByVal vcolDescriptionArgs As Collection, _
    Optional ByVal vblnAppendErrToDescription As Boolean = False) As Boolean
    
'This routine provides the error information build from RES file (description with optional parameters)
'if it the ProcessResourceDesc fails (blank), it preserves back the current input/output description
'It initializes all the module variables
'Builds the error description which may include multiple optional parameters
'Display for error messages only the run time error (mlngHldErrNumber = vlngVBErr)
'The vlngAppErr is used to retrieve error descriptions from the res file
'iostrErrDescription is the run time error description
'iostrErrSource is the container of the log
'vstrUserId is the NT employee id
'vstrSourceRoutine is the name of the VB routine
'vblnLog is the flag to write the error log
'vcolErrArgs are the additional arguments (i.e.SQL)
'vcolDescriptionArgs are the positional parameters for the error description in the res file
'vblnAppendErrToDescription is used to append the run time error description to the res error description

On Error GoTo ErrTrace_EH


    Initializemvars
        
    If ProcessResourceDesc(vlngAppErr, vcolDescriptionArgs) = True Then
    Else
       mstrHldErrDescription = iostrErrDescription
    End If


    mlngHldErrNumber = vlngVBErr
    mstrHldErrSource = iostrErrSource
    mstrHldAppExeName = App.EXEName
    miostrErrDescription = iostrErrDescription
    
    If vstrSourceRoutine = "" Then
       mstrHldSourceRoutine = "blank"
    Else
       mstrHldSourceRoutine = vstrSourceRoutine
    End If
    
    If vstrUserId = "" Then
       mstrHldUserId = "blank"
    Else
       mstrHldUserId = vstrUserId
    End If
    
    
    If vblnLog = False Then    'Do not write to the file when input parm vblnLog is false
        If InStr(mstrHldErrSource, trmsMSG_FIRSTAPPEND) <> 0 Then 'append to if trmsMSG_FIRSTAPPEND has been already been appended
           mstrHldLine = vbCrLf & trmsMSG_NEXTAPPEND
           BuildmstrHldLine vlngAppErr, vcolErrArgs, False
           mstrHldMsg = mstrHldErrDescription
           iostrErrSource = mstrHldWriteLine
           iostrErrDescription = mstrHldErrDescription
           ErrTrace = True
           Exit Function
        Else                    'Begin or first append - use trmsMSG_FIRSTAPPEND as indicator for searches using InStr
           mstrHldLine = vbCrLf & trmsMSG_FIRSTAPPEND
           mstrHldLine = mstrHldLine & Format(Now, "dd-mmm-yy hh:mm:ss")
           BuildmstrHldLine vlngAppErr, vcolErrArgs, True
           mstrHldMsg = mstrHldErrDescription
            iostrErrSource = mstrHldWriteLine
           If vblnAppendErrToDescription = True Then
'              mstrHldErrDescription = mstrHldArgLine & mstrHldMsg
              mstrHldErrDescription = mstrHldMsg & " - " & iostrErrDescription
           Else
              mstrHldErrDescription = mstrHldMsg
           End If
           iostrErrDescription = mstrHldErrDescription
           ErrTrace = True
           'add event viewer only for Begin Message
           ServerEventLog (mstrHldWriteLine)  'Log in the server
           Exit Function
        End If
    Else                'write to the file when input parm vblnLog is true
        If InStr(mstrHldErrSource, trmsMSG_FIRSTAPPEND) <> 0 Then  'append to log
           mstrHldLine = vbCrLf & trmsMSG_NEXTAPPEND
           BuildmstrHldLine vlngAppErr, vcolErrArgs, False
        Else
           mstrHldLine = vbCrLf & trmsMSG_FIRSTAPPEND  'create a new line to log
           mstrHldLine = mstrHldLine & Format(Now, "dd-mmm-yy hh:mm:ss")
           BuildmstrHldLine vlngAppErr, vcolErrArgs, True
           'add event viewer only for Begin Message
           ServerEventLog (mstrHldWriteLine)  'Log in the server
        End If
    End If

    
'    mstrHldFilePath = App.Path & "\" & mstrHldAppExeName & ".log"
'    mstrHldLbkFilePath = App.Path & "\" & mstrHldAppExeName & ".lbk"
'get the file names that will be used for the daily log

    GetDayOfWeekLogFiles mstrHldFilePath, mstrHldLbkFilePath

'errors are written to a daily flat file in the server

    LogFile mstrHldFilePath, mstrHldWriteLine, mstrHldLbkFilePath

'publish error without args (ie sql) in event viewer
'    ServerEventLog (mstrHldLine)  'Log in the server

'Building the i/output error source from what is written to the logfile

    mstrHldMsg = mstrHldErrDescription

    iostrErrSource = mstrHldWriteLine
    
'Building the i/output error description
'We use to build argument line of sql errors and the description
'But now we need to move the error description and the io error description
'The description generated by the res file (mstrHldMsg=mstrHldErrDescription)
'is appended to the run time description (iostrErrDescription)


    If vblnAppendErrToDescription = True Then
'       mstrHldErrDescription = mstrHldArgLine & mstrHldMsg
       mstrHldErrDescription = mstrHldMsg & " - " & iostrErrDescription
    Else
       mstrHldErrDescription = mstrHldMsg
    End If
    
    iostrErrDescription = mstrHldErrDescription
    
    ErrTrace = True

    Exit Function
    
ErrTrace_EH:

    iostrErrSource = mstrHldWriteLine
    If vblnAppendErrToDescription = True Then
'       mstrHldErrDescription = mstrHldArgLine & mstrHldMsg
       mstrHldErrDescription = mstrHldMsg & " - " & iostrErrDescription
    Else
       mstrHldErrDescription = mstrHldMsg
    End If
    
    iostrErrDescription = mstrHldErrDescription & trmsMSG_WARNCALLHELP & vbCrLf & mstrHldWriteLine & vbCrLf & trmsMSG_WARNRUNERROR & Err.Number & " - " & Err.Source & " : " & Err.Description & vbCrLf
    
    ErrTrace = False
    
    'publish error in the event viewer if error
    
    ServerEventLog (iostrErrDescription & vbCrLf & "*" & iostrErrSource & "*")  'Log in the server
    
    Exit Function
    
End Function

Private Function BuildDescriptionFromResFile(ByVal vlngAppErr As Long) As Boolean
'This routine extracts the description from the res file using the application error number

On Error GoTo BuildDescriptionFromResFile_EH

    mstrHldErrDescription = LoadResString(vlngAppErr)
    BuildDescriptionFromResFile = True
    Exit Function

BuildDescriptionFromResFile_EH:
    
    mstrHldErrDescription = mstrHldErrDescription & trmsMSG_WARNRESFILEMISSING & Err.Number & " - " & Err.Source & " : " & Err.Description & vbCrLf
    
    BuildDescriptionFromResFile = False
    

End Function

Private Function LogFile(ByVal vstrHldFilePath As String, ByVal vstrHldWriteLine As String, ByVal vstrHldLbkFilePath As String) As Boolean
'This routine creates a record log in the server:
'writing it in the server event viewer or
'writing it in the NT machine writing it to a flat file with extension .log
'This routine also recycles .log files that are older than 7 days

On Error GoTo LogFile_EH
    
'    If ServerEventLog(vstrHldWriteLine) = True Then  'Log in the server
'        Exit Function
'    Else
'    End If
    
    RecycleBackupLogFile vstrHldFilePath, vstrHldLbkFilePath 'clean log files 7 days old and create lbk backups
    
    mintHldFileNum = FreeFile

    Open vstrHldFilePath For Append Shared As #mintHldFileNum
    
    Print #mintHldFileNum, vstrHldWriteLine          'Log in the Client
    
    Close #mintHldFileNum
    LogFile = True
    
    Exit Function

LogFile_EH:
    
    mstrHldErrDescription = mstrHldErrDescription & trmsMSG_WARNFILEFAILED & Err.Number & " - " & Err.Source & " : " & Err.Description & vbCrLf
    
    LogFile = False

End Function

Private Sub BuildmstrHldLine(ByVal vlngAppErr As Long, _
                            ByVal vcolErrArgs As Collection, _
                            ByVal v_blnFirstAppend As Boolean)

'This routine uses the application error number and error optional arguments
'to build the global holdline error message (part I) which includes:
'application error number
'global run time VB error number
'global source routine name
'global NT employee id
'to build the global input parameters (part II) contained in the global argument line (io. SQL statements or anything you want to log)
'build the final global writeline (mstrHldWriteLine) which includes part I and part II
'mstrHldArgLine is appended to the written line by default
'if a res file error description is used
'(identical to the run time error description(miostrErrDescription)), (mstrHldArgLine=spaces)
'use only the error args (sql) since the description is already included
'for the res file.  Only include the run time error description when
'it is available

    mstrHldLine = mstrHldLine & vbCrLf & "AppErr: " & vlngAppErr
    mstrHldLine = mstrHldLine & vbCrLf & "VBErr: " & mlngHldErrNumber
    mstrHldLine = mstrHldLine & vbCrLf & "ErrDescription: " & mstrHldErrDescription
    If mstrHldErrDescription <> miostrErrDescription Then
       mstrHldLine = mstrHldLine & vbCrLf & "OrigErrDescription: " & miostrErrDescription
    End If
    mstrHldLine = mstrHldLine & vbCrLf & "SourceRoutine: " & mstrHldSourceRoutine
    mstrHldLine = mstrHldLine & vbCrLf & "UserId: " & mstrHldUserId
    '//-----COM+ application partition (mainly for test servers)-----//
    '//-----all partitions share the same log file-----//
    If v_blnFirstAppend Then '//-----output only once (sortof)
        mstrHldLine = mstrHldLine & vbCrLf & "COMPartition: " & GetCOMPartition
    End If
    mstrHldLine = mstrHldLine & vbCrLf & "Args: "      'adding optional input arguments
    
    If vcolErrArgs Is Nothing Then
    Else
        For Each mvntHldErrArg In vcolErrArgs
            mstrHldArgLine = mstrHldArgLine & vbCrLf & mvntHldErrArg
        Next
    End If

'    If mstrHldErrDescription = miostrErrDescription Then
'    Else
'       mstrHldArgLine = mstrHldArgLine & vbCrLf & "OrigErrDescription: " & miostrErrDescription
'    End If
    
    mstrHldWriteLine = vbCrLf & mstrHldErrSource & mstrHldLine & " " & mstrHldArgLine
    
End Sub

Private Function GetCOMPartition() As String
    Dim o As COMAdmin.COMAdminCatalog
    Dim sCOMPartition As String

    On Error GoTo eh
    
    Set o = CreateObject("COMAdmin.COMAdminCatalog")
    
    sCOMPartition = o.CurrentPartitionName
       
    GetCOMPartition = sCOMPartition
    
    Set o = Nothing
    Exit Function
eh:
    Resume Next
End Function

Private Sub Initializemvars()
'This routine initializes the module variables

    mstrHldLbkFilePath = ""
    mstrHldFilePath = ""
    mstrHldAppExeName = ""
    mstrHldSourceRoutine = ""
    mstrHldUserId = ""
    mstrHldErrDescription = ""
    mstrHldFormatDescription = ""
    mstrHldErrSource = ""
    mlngHldErrNumber = 0
    mvntHldErrArg = ""
    mvntHldDescriptionArg = ""
    mstrHldResDelimiter = ""
    mintHldDescNdx = 0
    mstrHldLine = ""
    mstrHldArgLine = ""
    mstrHldWriteLine = ""
    mintHldFileNum = 0
    mstrHldMsg = ""


End Sub

Private Function ProcessResourceDesc( _
        ByVal vlngAppErr As Long, _
        Optional ByVal vcolDescriptionArgs As Collection) As Boolean
    
'This routine uses the application error number (ie.30162) and optional parameters (ie.26649,2000) to
'build the description extracting the description from the Res file (TRISERRLOGMSG.RES)
'This routine replaces the brackets ([input parameter]) in the description with optional
'parameters such as employee number, department name
'30162 Employee [EmpNumber] does not have an accrual bank for bank year [BankYear]
'In this example it may result as follows:
'30162 Employee 30162 does not have an accrual bank for bank year 2000

    Dim lintInstrStart As Integer
    Dim lintInstrEnd As Integer
    Dim lintMidLen As Integer
    Dim lstrMidString As String
    
On Error GoTo ProcessResourceDesc_EH
    
    If vlngAppErr > 0 Then
       If BuildDescriptionFromResFile(vlngAppErr) = True Then
            ProcessResourceDesc = True
            If vcolDescriptionArgs Is Nothing Then
            Else
              For Each mvntHldDescriptionArg In vcolDescriptionArgs
                  lintInstrStart = InStr(mstrHldErrDescription, "[")                        'do the first special char of pair
                  lintInstrEnd = InStr(mstrHldErrDescription, "]")                          'do the second special char of pair
                  If lintInstrStart > 0 Then                                                'if a square bracket is found
                     lintMidLen = (lintInstrEnd - lintInstrStart) + 1                       'calculate length of parameter in resource file including square brackets []
                     lstrMidString = Mid(mstrHldErrDescription, lintInstrStart, lintMidLen) 'pull the parameter value so it can be replace in the next statement
                     mstrHldErrDescription = Replace(mstrHldErrDescription, lstrMidString, mvntHldDescriptionArg, , 1)  'replace first occurrence of parameter value with an argument
                     lintInstrStart = 0                                                     'initialize
                  Else
                     Exit Function
                  End If
              Next
            End If
       Else
           mstrHldErrDescription = "" 'return nothing
           ProcessResourceDesc = False
           Exit Function
       End If
    Else
       mstrHldErrDescription = ""     'return nothing
       ProcessResourceDesc = False
       Exit Function
    End If
        
    Exit Function
    
ProcessResourceDesc_EH:
    
    mstrHldErrDescription = mstrHldErrDescription & trmsMSG_WARNDESCRUNERROR & Err.Number & " - " & Err.Source & " : " & Err.Description & vbCrLf
       
    ProcessResourceDesc = False

End Function

Private Function RecycleBackupLogFile(ByVal vstrHldFilePath As String, ByVal vstrHldLbkFilePath As String) As Boolean
'This routine gets the existing log file calculated date (Creation date + 7: older than 7 days) and compares it to today's date
'Create a backup file with extension .lbk
'    App.Path & "\" & App.EXEName & ".log"
'    App.Path & "\" & App.EXEName & ".lbk"


    Dim lFileFSO As New FileSystemObject 'use reference Microsoft Scripting Runtime
    Dim lFileAttr As File
    Dim lFileLogDate As Date 'calc date of current log file
    Dim ldatLogDate As Date 'calc date of current log file
    Dim ldatTodayDate As Date 'today's date

On Error GoTo RecycleBackupLogFile_EH

    If Len(Dir(vstrHldFilePath)) Then
        Set lFileAttr = lFileFSO.GetFile(vstrHldFilePath)
        lFileLogDate = DateAdd(trmsEXPIREDAYTYPE, trmiEXPIREDAYS, lFileAttr.DateLastModified) 'use datelastmodified instead of create date
        ldatLogDate = CDate(DatePart("m", lFileLogDate) & "/" & DatePart("d", lFileLogDate) & "/" & DatePart("yyyy", lFileLogDate))
        ldatTodayDate = CDate(DatePart("m", Now) & "/" & DatePart("d", Now) & "/" & DatePart("yyyy", Now))
        If ldatLogDate < ldatTodayDate Then
           lFileFSO.CopyFile vstrHldFilePath, vstrHldLbkFilePath, True  'copy existing .log file to .lbk backup file
           KillFile (vstrHldFilePath)                                  'delete existing .log file
        Else
        End If
        RecycleBackupLogFile = True
    Else
        RecycleBackupLogFile = False
    End If
    
    Exit Function

RecycleBackupLogFile_EH:

    mstrHldErrDescription = mstrHldErrDescription & trmsMSG_WARNCALLHELP & vbCrLf & mstrHldWriteLine & vbCrLf & trmsMSG_WARNRUNERRORP & Err.Number & " - " & Err.Source & " : " & Err.Description & vbCrLf
    
    RecycleBackupLogFile = False

End Function

Private Function KillFile(ByVal vstrHldFilePath As String) As Boolean
'This routine deletes the existing log file calculated date (Creation date + 7: older than 7 days) is less than today's date
'A new log file will be created when a new error is encountered

On Error GoTo KillFile_EH

    If Len(Dir(vstrHldFilePath)) Then
        Kill vstrHldFilePath
        KillFile = True
    Else
        KillFile = False
    End If
    
    Exit Function

KillFile_EH:
    
    mstrHldErrDescription = mstrHldErrDescription & trmsMSG_WARNCALLHELP & vbCrLf & mstrHldWriteLine & vbCrLf & trmsMSG_WARNRUNERRORK & Err.Number & " - " & Err.Source & " : " & Err.Description & vbCrLf

    KillFile = False

End Function

Private Function ServerEventLog(ByVal vstrHldWriteLine As String) As Boolean
'This routine sends the error message () to the event viewer
'In order for this record to show up in the event viewer, the server needs this registry entry:
'My Computer\HKEY_LOCAL_MACHINE\SOFTWARE\Ameren\TRIS\NTEventLog

On Error GoTo ServerEventLog_EH

    Dim objRegAccess As New REGTool5.Registry
    Dim strValueData As String
    If objRegAccess.GetKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Ameren\TRIS\NTEventLog", "", strValueData) Then   'all this is required
        App.LogEvent ("TRISNTEventLog " & vstrHldWriteLine), vbLogEventTypeError
        ServerEventLog = True
    Else
        ServerEventLog = False
    End If
    
    Set objRegAccess = Nothing
   
    Exit Function

ServerEventLog_EH:
    
    mstrHldErrDescription = mstrHldErrDescription & trmsMSG_WARNCALLHELP & vbCrLf & vstrHldWriteLine & vbCrLf & trmsMSG_WARNRUNERRORL & Err.Number & " - " & Err.Source & " : " & Err.Description & vbCrLf

    ServerEventLog = False

End Function

Private Function GetDayOfWeekLogFiles(ByRef vstrHldFilePath As String, ByRef vstrHldLbkFilePath As String)
'This routine gets the day of week file for the error log
'    App.Path & "\" & App.EXEName & ".log"
'    App.Path & "\" & App.EXEName & ".lbk"
    Dim lstrTimeStamp As String 'put timestamp for all backup files only
    
On Error GoTo GetDayOfWeekLogFiles_EH
    
    lstrTimeStamp = Trim(CStr(DatePart("m", Now) & DatePart("d", Now) & DatePart("yyyy", Now) & DatePart("h", Now) & DatePart("n", Now) & DatePart("s", Now)))

    Select Case Weekday(Date)
    Case vbSunday
        vstrHldFilePath = App.Path & "\" & mstrHldAppExeName & ".sun"
        vstrHldLbkFilePath = App.Path & "\" & mstrHldAppExeName & lstrTimeStamp & ".snk"
    Case vbMonday
        vstrHldFilePath = App.Path & "\" & mstrHldAppExeName & ".mon"
        vstrHldLbkFilePath = App.Path & "\" & mstrHldAppExeName & lstrTimeStamp & ".mnk"
    Case vbTuesday
        vstrHldFilePath = App.Path & "\" & mstrHldAppExeName & ".tue"
        vstrHldLbkFilePath = App.Path & "\" & mstrHldAppExeName & lstrTimeStamp & ".tek"
    Case vbWednesday
        vstrHldFilePath = App.Path & "\" & mstrHldAppExeName & ".wed"
        vstrHldLbkFilePath = App.Path & "\" & mstrHldAppExeName & lstrTimeStamp & ".wdk"
    Case vbThursday
        vstrHldFilePath = App.Path & "\" & mstrHldAppExeName & ".thu"
        vstrHldLbkFilePath = App.Path & "\" & mstrHldAppExeName & lstrTimeStamp & ".thk"
    Case vbFriday
        vstrHldFilePath = App.Path & "\" & mstrHldAppExeName & ".fri"
        vstrHldLbkFilePath = App.Path & "\" & mstrHldAppExeName & lstrTimeStamp & ".frk"
    Case vbSaturday
        vstrHldFilePath = App.Path & "\" & mstrHldAppExeName & ".sat"
        vstrHldLbkFilePath = App.Path & "\" & mstrHldAppExeName & lstrTimeStamp & ".stk"
    Case Else
        vstrHldFilePath = App.Path & "\" & mstrHldAppExeName & ".log"
        vstrHldLbkFilePath = App.Path & "\" & mstrHldAppExeName & lstrTimeStamp & ".lgk"
    End Select

    
    Exit Function

GetDayOfWeekLogFiles_EH:

    mstrHldErrDescription = mstrHldErrDescription & trmsMSG_WARNCALLHELP & vbCrLf & mstrHldWriteLine & vbCrLf & trmsMSG_WARNGETFILEERROR & Err.Number & " - " & Err.Source & " : " & Err.Description & vbCrLf
    
End Function




