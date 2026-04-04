VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "cuSPMCache"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' 8/7/2002 RSG
' Changed Invalidate, Count, Keys so that a group is not created when it
' does not already exist.  This will ensure that the group is bulk initialized
' on the first Item call, even if a previous attempt to invalidate a property
' occurred before the group was initialized.
' This leaves Add, AddAll, and Remove that create the group when it did not
' already exist.  This is fine, as long as the Add functions are called as a
' result of initializing.  If Add/AddAll will be called outside of initializing,
' the group must be initialized first.  Otherwise, GetRecord will not attempt to
' initialize a bulk group.
' 5/29/2002 RSG
' While database response times were really slow we encountered a scenario
' where we needed LockMethod.  Initially, we believed that the worst case of
' using LockGetSet was an erroneous validation error.  However, since processing
' is driven by some indicators cached, the worst case is erroneous processing,
' e.g. not updating a bank.  Therefore, no erroneous cached information is acceptable,
' no matter how remote the possibility.
'
' It seems to me that 3 changes are necessary.
' 1.  Avoid creating the property group when checking for an item.
' I originally coded it this way to avoid a runtime error,
' Invalid procedure call or argument, when the group did not exist yet.
' It appears that I must use the Group getter method and handle the runtime error.
' For an Item request, this will answer False for group existence and cause each
' user that requests an item before the cache is initialized to request the
' data from the DB and load again.  Basically, this is the same as
' lazy initiallized cache.  This will cause multiple requests and initialization,
' but no inconsistent results since the user is blocked until the
' initialization is complete.  This is where the biggest window for a problem
' to arise, usually no more than 1/2 second, but could be longer when the
' database response time is slow.
' 2.  Use LockMethod isolation mode.  I previously recommended LockGetSet,
' but now, considering the worst case scenario, we need this to guard against
' a scenario where user A starts the load process, creating the group,
' then user B requests an item that hasn't been loaded yet.
' Since all the data is already retrieved from the database, I don't believe
' this window is bigger than 1/10 second, but I don't think we want
' to take the chance.  Considering the lock duration is not the entire
' causality, I believe the performance degradation will be negligle.
' 3.  Avoid instance data altogether.  We did get some confusing results
' when using the IsolationMode letter that needed to be fixed.
' The property letter simply sets an instance variable that will be
' used by the next method invocation that results in creating the group.
' If the group is already created, it will be ignored.  Having instance data
' is a bad idea in this case since it does not warn that it is ignored and
' it may be lost depending on whether a reference to the object is released
' or not before the next method invocation.  Suggest changing the interface
' to allow setting all the mode letters in a single call and create the
' group at that time.  This will remove all the instance state from the component.
' Also, return a boolean from this method indicating if successful or not.
' If group already existed, then unsuccessful.
' ******************************************************
' ALERT!!!      ALERT!!!        ALERT!!!
' BE SURE TO DISABLE TRDEBUG BEFORE MOVING TO PRODUCTION
' ******************************************************
'
' An instance of this class will be the shared property
' manager cache wrapper.
' All keys must be a string.
' Optional:
' Build entire cache on the first request.
' Load records as requested.
'
' The isolation mode, release mode, and key support are initialized
' on instantiation.
' The client can only specify these values when the group is created.
' Subsequent access to an existing group will ignore these values.
' I LockMethod isolation mode is supported. LockMethod is
' required to support the Count and Keys shared properties.
'
' Currently, coded to support Keys and Count if property is
' set when the group is created.  May need more work here
' to ensure no duplicates since an error may result when
' attempting to add a duplicate to keys collection, or attempt
' to remove a key that does not exist.  It seems this can only
' occur for 2 reasons:
' 1) if add/remove requests are inconsistent regarding the
' supportKeys property.  This problably should be made a property
' of the group itself and set when the group is initialized.
' I am not certain this can happen while Createproperty is
' being used.
' 2) LockMethod was not selected with key support.  This could
' result in the keys array becoming corrupted.
'
' I believe both of these have been addressed by making the
' SupportKeys property similar to Lock and Release mode, where
' the property is used when the group is created, but ignored for
' each access after that.  Also, key support is only enabled when
' LockMethod has been selected.
'
' DEBUGGING NOTE:
'
' Must enable conditional compilation constant TRDEBUG for debugging.
' This will set the isolation mode to LockGetSet without having to
' change the source code.
'
' Isolation mode must LockGetSet while debugging in VB. Otherwise,
' a fail fast will result with following message in event log:
' Event: 4149   Source: Transaction Server
' The Shared Property Manager (SPM) requires a context for
' method level locking.  Make sure that the component using SPM
' is a Transaction Server component.
' (Microsoft Transaction Server Internals Information:
' File: d:\viper\src\runtime\mtxspm\spgobj.cpp, Line: 64)
'
' This is caused because LockMethod requires access to
' object context and it is not really available while debugging,
' only the special debug version.
'
' 11/19/01 RSG - Added Invalidate method.
' Client will call this after changing a database value for
' a cached property.  A client cannot update the property directly
' because such an update cannot be backed out when a transaction is
' aborted.  Therefore, it can only invalidate the cached entry and
' a subsequent lookup request will return CSPINVALID,
' which will cause a SQL request to be processed to update
' the cache.
'
'
' ******************************************************
' ALERT!!!      ALERT!!!        ALERT!!!
' BE SURE TO DISABLE TRDEBUG BEFORE MOVING TO PRODUCTION
' ******************************************************
'Enable trace info by providing conditional constant
' TRDEBUG = -1 here or project properties dialog.
' Disable with TRDEBUG = 0.
'#Const TRDEBUG = -1

Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Const mstrCMODULENAME = "cuSPMCache"
Private Const CDOESNOTEXIST = "SPVALUENOTEXIST"
Private Const CITEMCOUNTKEY = "SPGITEMCOUNT"
Private Const CITEMKEYS = "SPGKEYS"
Private Const CSUPPORTKEYS = "SPGSUPPORTKEYS"
Private Const CSPINVALID = "SPINVALID"

Private Const CDEBUGISOLATIONMODE As Long = LockSetGet
Private Const CDEFAULTISOLATIONMODE As Long = LockMethod
Private Const CDEFAULTRELEASEMODE As Long = Process
Private Const CDEFAULTKEYSUPPORT As Boolean = False

' Set true in initialize
'Private mblnDebugSession As Boolean

'These module variables are i/o parameters in Create
'Private mlngIsolationMode As Long
'Private mlngReleaseMode As Long
'Private mblnSupportKeys As Boolean

'******************************************************
' Public properties
'******************************************************
'Group name is not required here because these properties
'may only be set once before the first call that creates the
'group.

'
' Letters are obsolete, use CreateGroup.
'
'Public Property Let IsolationMode(ByVal v_lngIsoMode As Long)
''    mlngIsolationMode = v_lngIsoMode
'    Err.Raise 30000 + vbojbecterror, mstrCMODULENAME, "IsolationMode is obsolete, use CreateGroup"
'End Property
'
'Public Property Let ReleaseMode(ByVal v_lngRelMode As Long)
''    mlngReleaseMode = v_lngRelMode
'    Err.Raise 30000 + vbojbecterror, mstrCMODULENAME, "ReleaseMode is obsolete, use CreateGroup"
'End Property
'
'Public Property Let SupportKeys(ByVal v_blnSupKeys As Boolean)
''    mblnSupportKeys = v_blnSupKeys
'    Err.Raise 30000 + vbojbecterror, mstrCMODULENAME, "SupportKeys is obsolete, use CreateGroup"
'End Property

'******************************************************
' Public procedures
'******************************************************
'
' Preserve compatibility with GetIsolationMode, GetReleaseMode, and GetSupportsKeys
'
' These getters will create the group by default.  This is the only way to get
' the isolation and release mode for an existing group.  If this is not desired,
' use GroupExists first, then CreateGroup as desired.
'
'Public Function GetIsolationMode(ByVal v_strGroupName As String) As Long
'    Dim lblnExists As Boolean
'
'On Error GoTo eh
'    GetIsolationMode = IsolationMode(Trim(v_strGroupName), lblnExists)
'    Exit Function
'
'eh
'    TrsRaiseError mstrCMODULENAME, "GetIsolationMode"
'End Function

Public Function GetIsolationMode( _
                    ByVal v_strGroupName As String, _
                    ByRef r_blnExists As Boolean) As Long
'    Answer the isolation mode for an existing group.  If it does not exist,
'    answer the default isolation mode.  Client must check r_blnExists to know
'    if it answered the actual or default.  This will avoid creating the group
'    with default parameters when it does not exist.
    Dim spg As SharedPropertyGroup
    Dim llngIsolationMode As Long
    Dim llngReleaseMode As Long
    Dim lblnKeySupport As Boolean
    
On Error GoTo eh

    Set spg = GetSharedPropertyGroup(v_strGroupName, r_blnExists, False)
    If Not r_blnExists Then
        GetIsolationMode = CDEFAULTISOLATIONMODE
        GoTo xit
    End If
    
    InitializeDefaults llngIsolationMode, llngReleaseMode, lblnKeySupport
    Set spg = CreateSharedPropertyGroup(Trim(v_strGroupName), llngIsolationMode, llngReleaseMode, lblnKeySupport, r_blnExists)
    GetIsolationMode = llngIsolationMode
    
xit:
    Set spg = Nothing
    Exit Function
eh:
    Set spg = Nothing
    TrsRaiseError mstrCMODULENAME, "GetIsolationMode"
End Function

'Public Function GetReleaseMode(ByVal v_strGroupName As String) As Long
'    Dim lblnExists As Boolean
'
'On Error GoTo eh
'    GetReleaseMode = ReleaseMode(Trim(v_strGroupName), lblnExists)
'    Exit Function
'
'eh
'    TrsRaiseError mstrCMODULENAME, "GetReleaseMode"
'End Function

Public Function GetReleaseMode( _
                    ByVal v_strGroupName As String, _
                    ByRef r_blnExists As Boolean) As Long
    Dim spg As SharedPropertyGroup
    Dim llngIsolationMode As Long
    Dim llngReleaseMode As Long
    Dim lblnKeySupport As Boolean
    
On Error GoTo eh
    Set spg = GetSharedPropertyGroup(v_strGroupName, r_blnExists, False)
    If Not r_blnExists Then
        GetReleaseMode = CDEFAULTRELEASEMODE
        GoTo xit
    End If

    InitializeDefaults llngIsolationMode, llngReleaseMode, lblnKeySupport
    Set spg = CreateSharedPropertyGroup(Trim(v_strGroupName), llngIsolationMode, llngReleaseMode, lblnKeySupport, r_blnExists)
    GetReleaseMode = llngReleaseMode
    
xit:
    Set spg = Nothing

    Exit Function
eh:
    Set spg = Nothing
    TrsRaiseError mstrCMODULENAME, "GetReleaseMode"
End Function

'Public Function GetSupportsKeys(ByVal v_strGroupName As String) As Boolean
'    Dim lblnExists As Boolean
'On Error GoTo eh
'
'    GetSupportsKeys = SupportsKeys(Trim(v_strGroupName), lblnExists)
'    Exit Function
'
'eh
'    TrsRaiseError mstrCMODULENAME, "GetSupportsKeys"
'End Function

Public Function GetSupportsKeys( _
                    ByVal v_strGroupName As String, _
                    ByRef r_blnExists As Boolean) As Boolean
    Dim spg As SharedPropertyGroup
    Dim llngIsolationMode As Long
    Dim llngReleaseMode As Long
    Dim lblnKeySupport As Boolean
    
On Error GoTo eh
    Set spg = GetSharedPropertyGroup(v_strGroupName, r_blnExists, False)
    If Not r_blnExists Then
        GetSupportsKeys = CDEFAULTKEYSUPPORT
        GoTo xit
    End If

    InitializeDefaults llngIsolationMode, llngReleaseMode, lblnKeySupport
    Set spg = CreateSharedPropertyGroup(Trim(v_strGroupName), llngIsolationMode, llngReleaseMode, lblnKeySupport, r_blnExists)
    GetSupportsKeys = lblnKeySupport
    
xit:
    Set spg = Nothing

    Exit Function
eh:
    Set spg = Nothing
    TrsRaiseError mstrCMODULENAME, "GetSupportsKeys"
End Function


Public Function CreateGroup( _
                    ByVal v_strGroupName As String, _
                    ByRef r_blnExists As Boolean, _
                    Optional ByRef r_lngIsolationMode As Long = CDEFAULTISOLATIONMODE, _
                    Optional ByRef r_lngReleaseMode As Long = CDEFAULTRELEASEMODE, _
                    Optional ByRef r_blnKeySupport As Boolean = CDEFAULTKEYSUPPORT)
' Provide this method to provide isolation mode, release mode,
' and key support, only when the group is created.  This will
' avoid the instance variables and return this wrapper to
' stateless methods.
' If r_blnExists is True, the isolation mode, release mode, and
' key support were ignored.
'
' DO NOT set Release mode to Standard, or non default unless you know what you are doing.
' Use caution when setting Release mode to Standard.  The group will be discarded
' when all references are released, which is standard procedure for the TRIS application.
' I dont know when this release mode is appropriate.
'
    Dim spg As SharedPropertyGroup
    Dim lblnExists As Boolean
    
On Error GoTo eh
    Set spg = CreateSharedPropertyGroup(Trim(v_strGroupName), r_lngIsolationMode, r_lngReleaseMode, r_blnKeySupport, r_blnExists)

    Set spg = Nothing

    Exit Function
eh:
    Set spg = Nothing
    TrsRaiseError mstrCMODULENAME, "CreateGroup"
End Function


Public Function GroupExists( _
        ByVal v_strGroupName As String) As Boolean
'Answer True when the group already existed.
    
    Dim spg As SharedPropertyGroup
    Dim lblnExist As Boolean
    Dim lstrGroupName As String
    
On Error GoTo eh
    
    lstrGroupName = Trim(v_strGroupName)
    Set spg = GetSharedPropertyGroup(lstrGroupName, lblnExist, False)
    GroupExists = lblnExist
    
    Set spg = Nothing
    Exit Function
eh:
    Set spg = Nothing
    TrsRaiseError mstrCMODULENAME, "GroupExists"
End Function

Public Function Item( _
        ByVal v_strKey As String, _
        ByVal v_strGroupName As String, _
        ByRef r_blnExists As Boolean, _
        ByRef r_blnGroupExists As Boolean) As Variant
'Answer a shared property value.
'If r_blnPropExists and return value is Empty, this would indicate to
'caller that previous lookup has been occured but no value set.
'Useful for lazy materializing caller, indicating no need to
'attempt additional data access.
'If r_blnGroupExists is false, it indicates that the group
'has not been initialized.
    
    Dim spg As SharedPropertyGroup
    Dim sp As SharedProperty
    Dim lstrGroupName As String
    Dim lstrKey As String
    Dim lvarValue As Variant
    Dim lblnExists As Boolean
    
    Dim stTime As Long
    Dim ndTime As Long
    Dim dfTime As Long
    
    Dim l_colErrArgs As Collection
    Dim sLine As String
On Error GoTo eh
    

'    #If TRDEBUG Then
'        stTime = GetTickCount()
'    #End If

    lstrKey = Trim(v_strKey)
    lstrGroupName = Trim(v_strGroupName)
    
    sLine = "Set spg = GetSharedPropertyGroup(lstrGroupName, lblnExists, False)"
    Set spg = GetSharedPropertyGroup(lstrGroupName, lblnExists, False)
    r_blnGroupExists = lblnExists
    If Not lblnExists Then
        Item = Empty
        GoTo xit
    End If
    
    '//-----workaround for Windows 2003 Server-----//
    If Trim(lstrKey) = "" Then
        lstrKey = "<blank>"
    End If
   
    sLine = "Set sp = spg.CreateProperty(lstrKey, r_blnExists)"
    Set sp = spg.CreateProperty(lstrKey, r_blnExists)
    
    sLine = "If r_blnExists Then"
    If r_blnExists Then
        If DoesNotExist(sp.Value) Then
            r_blnExists = False
            sLine = "Item = Empty"
            Item = Empty
        Else
            sLine = "Item = sp.Value"
            Item = sp.Value
        End If
    Else
    ' initialize it
        sp.Value = CDOESNOTEXIST
        Item = Empty
    End If
    
'    #If TRDEBUG Then
'        ndTime = GetTickCount()
'        dfTime = ndTime - stTime
'        Debug.Print "Get Item: " & lstrKey & " in seconds: " & CStr(dfTime / 1000)
'    #End If
    
xit:
    Set spg = Nothing
    Set sp = Nothing
    Exit Function
    
eh:
    Set l_colErrArgs = New Collection
    l_colErrArgs.Add "v_strKey = """ & v_strKey & """"
    l_colErrArgs.Add "v_strGroupName = " & v_strGroupName
    l_colErrArgs.Add "r_blnExists = " & r_blnExists
    l_colErrArgs.Add "r_blnGroupExists = " & r_blnGroupExists
    l_colErrArgs.Add "sLine = """ & sLine & """"
    
    Set spg = Nothing
    Set sp = Nothing
    TrsRaiseError mstrCMODULENAME, "Item", , , , l_colErrArgs, True
End Function

Public Function Add( _
        ByVal v_strKey As String, _
        ByVal v_varValue As Variant, _
        ByVal v_strGroupName As String) As Long
'Add a shared property value.
    
    Dim sp As SharedProperty
    Dim spg As SharedPropertyGroup
'    Dim spgm As SharedPropertyGroupManager
    Dim lblnExists As Boolean
    Dim lvarKeysArray As Variant
    Dim lstrGroupName As String
    Dim lstrKey As String
    Dim colItems As Collection
    Dim llngKeyCount As Long
    Dim i As Integer
    
On Error GoTo eh
    
    lstrKey = Trim(v_strKey)
    lstrGroupName = Trim(v_strGroupName)

    Set spg = GetSharedPropertyGroup(lstrGroupName, lblnExists)
    If spg Is Nothing Then
        Add = 1     'Invalid group name - cannot be empty string
        GoTo xit
    End If
    
    lvarKeysArray = Array()
    
    '//-----workaround for Windows 2003 Server-----//
    If Trim(lstrKey) = "" Then
        lstrKey = "<blank>"
    End If
        
    Set sp = spg.CreateProperty(lstrKey, lblnExists)
' Only increment if existing property has value of CDOESNOTEXIST,
' otherwise it is a replace.
    If lblnExists Then
        If DoesNotExist(sp.Value) Then
            lvarKeysArray = Array(lstrKey)
        End If
    Else
        lvarKeysArray = Array(lstrKey)
    End If
    sp.Value = v_varValue
    
    If ArraySize(lvarKeysArray) < 1 _
        Or Not IsKeySupportEnabled(spg) Then
        GoTo xit
    End If
    
    ' update the keys and count
    Set sp = spg.CreateProperty(CITEMKEYS, lblnExists)
    If lblnExists Then
        Set colItems = _
            CollectionFromArray(sp.Value, LBound(sp.Value), UBound(sp.Value), True)
    Else
        Set colItems = New Collection
    End If
    For i = LBound(lvarKeysArray) To UBound(lvarKeysArray)
        colItems.Add lvarKeysArray(i)
    Next
    llngKeyCount = colItems.Count
    If llngKeyCount = 0 Then
        sp.Value = Array()
    Else
        sp.Value = ArrayFromCollection(colItems)
    End If
    Set sp = spg.CreateProperty(CITEMCOUNTKEY, lblnExists)
    sp.Value = llngKeyCount

xit:
    Set sp = Nothing
    Set spg = Nothing
    
    Exit Function
eh:
    Set sp = Nothing
    Set spg = Nothing
    
    TrsRaiseError mstrCMODULENAME, "Add"
End Function

Public Function AddAll( _
        ByVal v_varKeyValueArray As Variant, _
        ByVal v_strGroupName As String) As Long
'Add all shared property values.
'v_varKeyValuePairs must be a variant array with 2 dim.
    
    Dim sp As SharedProperty
    Dim spg As SharedPropertyGroup
'    Dim spgm As SharedPropertyGroupManager
    Dim lblnExists As Boolean
    Dim lintLBoundDim1 As Integer
    Dim lintUBoundDim1 As Integer
    Dim i As Integer
    Dim lvarKeysArray As Variant
    Dim lstrKey As String
    Dim lstrGroupName As String
    Dim lcolNewKeys As Collection
    Dim lcolKeys As Collection
    Dim llngKeyCount As Long
    Dim lblnKeySupportEnabled As Boolean
    
On Error GoTo eh
    
'Assume it is a 2 dim array
    If Not IsArray(v_varKeyValueArray) Then
        AddAll = 2     'Invalid array - keyValue pairs is not array
        GoTo xit
    End If
    lintLBoundDim1 = LBound(v_varKeyValueArray, 1)
    lintUBoundDim1 = UBound(v_varKeyValueArray, 1)

    lstrGroupName = Trim(v_strGroupName)
    Set spg = GetSharedPropertyGroup(lstrGroupName, lblnExists)
    If spg Is Nothing Then
        AddAll = 1     'Invalid group name - cannot be empty string
        GoTo xit
    End If
    
    lblnKeySupportEnabled = IsKeySupportEnabled(spg)
' Build collect keys and convert to array to avoid redim preserve.
    Set lcolNewKeys = New Collection
    For i = LBound(v_varKeyValueArray, 2) To UBound(v_varKeyValueArray, 2)
        lstrKey = Trim(v_varKeyValueArray(lintLBoundDim1, i))
        
        '//-----workaround for Windows 2003 Server-----//
        If Trim(lstrKey) = "" Then
            lstrKey = "<blank>"
        End If

        Set sp = spg.CreateProperty(lstrKey, lblnExists)
        If lblnExists Then
            If DoesNotExist(sp.Value) Then
                If lblnKeySupportEnabled Then
                    lcolNewKeys.Add lstrKey
                End If
            End If
        Else
            If lblnKeySupportEnabled Then
                lcolNewKeys.Add lstrKey
            End If
        End If
        ' reassign value if already existed
        sp.Value = v_varKeyValueArray(lintUBoundDim1, i)
    Next
    
    If lcolNewKeys.Count < 1 _
        Or Not lblnKeySupportEnabled Then
        GoTo xit
    End If

    ' update the keys and count
    lvarKeysArray = ArrayFromCollection(lcolNewKeys)
    Set sp = spg.CreateProperty(CITEMKEYS, lblnExists)
    If lblnExists Then
        Set lcolKeys = _
            CollectionFromArray(sp.Value, LBound(sp.Value), UBound(sp.Value), True)
    Else
        Set lcolKeys = New Collection
    End If
    For i = LBound(lvarKeysArray) To UBound(lvarKeysArray)
        lcolKeys.Add lvarKeysArray(i)
    Next
    llngKeyCount = lcolKeys.Count
    If llngKeyCount = 0 Then
        sp.Value = Array()
    Else
        sp.Value = ArrayFromCollection(lcolKeys)
    End If
    Set sp = spg.CreateProperty(CITEMCOUNTKEY, lblnExists)
    sp.Value = llngKeyCount
    
xit:
    Set sp = Nothing
    Set spg = Nothing
    
    Exit Function
eh:
    Set sp = Nothing
    Set spg = Nothing
    
    TrsRaiseError mstrCMODULENAME, "AddAll"
End Function

Public Function Count( _
        ByVal v_strGroupName As String) As Long
'Answer a count of the valid properties, ie. not including CDOESNOTEXIST
    
    Dim spg As SharedPropertyGroup
    Dim sp As SharedProperty
    Dim lblnExists As Boolean
    Dim lstrGroupName As String
    
On Error GoTo eh
    
    lstrGroupName = Trim(v_strGroupName)
    Set spg = GetSharedPropertyGroup(lstrGroupName, lblnExists, False)
    If spg Is Nothing Then
        Count = 0
        GoTo xit
    End If
    
    Set sp = spg.CreateProperty(CITEMCOUNTKEY, lblnExists)
    If lblnExists Then
        Count = sp.Value
    Else
        sp.Value = 0
        Count = 0
    End If
    
xit:
    Set sp = Nothing
    Set spg = Nothing
    
    Exit Function
eh:
    Set spg = Nothing
    Set sp = Nothing
    
    TrsRaiseError mstrCMODULENAME, "Count"
End Function

Public Function Remove( _
        ByVal v_strKey As String, _
        ByVal v_strGroupName As String) As Long
'Logically remove a shared property value.
    
    
On Error GoTo eh
    
    Remove = RemoveBasic(v_strKey, v_strGroupName)
    
    Exit Function
eh:
    TrsRaiseError mstrCMODULENAME, "Remove"
End Function

Public Function RemoveAll( _
        ByVal v_varKeysArray As Variant, _
        ByVal v_strGroupName As String) As Long
'Logically remove a shared property value.
'Not very efficient, but dont expect much use.
    
    Dim sp As SharedProperty
    Dim i As Integer
    Dim llngReturn As Long
    Dim lblnError As Boolean
    Dim lstrGroupName As String
    
On Error GoTo eh
    
    lstrGroupName = Trim(v_strGroupName)
    If Not IsArray(v_varKeysArray) Then
        RemoveAll = 1     'Invalid array - keys is not array
        GoTo xit
    End If

    For i = LBound(v_varKeysArray) To UBound(v_varKeysArray)
        llngReturn = RemoveBasic(v_varKeysArray(i), lstrGroupName)
        If llngReturn <> 0 Then
            lblnError = True
        End If
    Next
    RemoveAll = 1   'One or more items where not removed
    
xit:
    Set sp = Nothing
    
    Exit Function
eh:
    Set sp = Nothing
    
    TrsRaiseError mstrCMODULENAME, "RemoveAll"
End Function

Public Function Keys( _
        ByVal v_strGroupName As String) As Variant
'Answer the variant array of keys.
    
    Dim spg As SharedPropertyGroup
    Dim sp As SharedProperty
    Dim lblnExists As Boolean
    Dim lstrGroupName As String
    Dim i As Integer
    
On Error GoTo eh
    
    lstrGroupName = Trim(v_strGroupName)

    Set spg = GetSharedPropertyGroup(lstrGroupName, lblnExists, False)
    If spg Is Nothing Then
        Keys = Array()
        GoTo xit
    End If
    
    Set sp = spg.CreateProperty(CITEMKEYS, lblnExists)
    Keys = sp.Value

xit:
    Set spg = Nothing
    Set sp = Nothing
    
    Exit Function
eh:
    Set spg = Nothing
    Set sp = Nothing
    
    TrsRaiseError mstrCMODULENAME, "Keys"
End Function


Public Function Invalidate( _
        ByVal v_strKey As String, _
        ByVal v_strGroupName As String) As Long
' Invalidate a shared property value.
' Client will call this after changing a database value for a cached property.
' A subsequent lookup request will return CRECORDINVALID,
' which will cause a SQL request to be processed to update the cache.
' Do not update keys or count when invalidating.  If client deleted it,
' the subsequent lookup will remove it.
    Dim sp As SharedProperty
    Dim spg As SharedPropertyGroup
    Dim lblnExists As Boolean
    Dim lstrKey As String
    Dim lstrGroupName As String
    
On Error GoTo eh
    lstrKey = Trim(v_strKey)
    lstrGroupName = Trim(v_strGroupName)
' Do not create the group if it does not exist.
    Set spg = GetSharedPropertyGroup(lstrGroupName, lblnExists, False)
' If it does not exist, nothing to invalidate.
    If lblnExists Then
    
        '//-----workaround for Windows 2003 Server-----//
        If Trim(lstrKey) = "" Then
            lstrKey = "<blank>"
        End If

        Set sp = spg.CreateProperty(lstrKey, lblnExists)
    ' Always set to CSPINVALID.  This will ensure that a newly added rotation
    ' will be loaded on the next validation.
        sp.Value = CSPINVALID
    End If
    
    Set sp = Nothing
    Set spg = Nothing

    Exit Function
eh:
    Set sp = Nothing
    Set spg = Nothing
    TrsRaiseError mstrCMODULENAME, "Invalidate"
End Function


'******************************************************
' Private procedures
'******************************************************
Private Function IsKeySupportEnabled( _
                    ByRef r_spg As SharedPropertyGroup) As Boolean
    Dim sp As SharedProperty
    Dim lblnExists As Boolean
    
On Error GoTo eh

    Set sp = r_spg.CreateProperty(CSUPPORTKEYS, lblnExists)
'    Guard against an uninitialized property
    If VarType(sp.Value) = vbBoolean Then
        IsKeySupportEnabled = sp.Value
    Else
        IsKeySupportEnabled = False
    End If

    Exit Function

eh:
    TrsRaiseError mstrCMODULENAME, "IsKeySupportEnabled"
End Function


'Private Function IsKeySupportEnabled() As Boolean
'
'On Error GoTo eh
'    If mblnSupportKeys Then
'        If (mlngIsolationMode = LockMethod) Or mblnDebugSession Then
'            IsKeySupportEnabled = True
'            Exit Function
'        End If
'    End If
'    IsKeySupportEnabled = False
'    Exit Function
'
'eh:
'    TrsRaiseError mstrCMODULENAME, "IsKeySupportEnabled"
'End Function

Private Function GetSharedPropertyGroup( _
                    ByVal v_strGroupName As String, _
                    ByRef r_blnExists As Boolean, _
                    Optional ByVal v_blnCreate As Boolean = True) _
                As SharedPropertyGroup
'Ok to initialize here with both LockMethod and LockGetSet.
    Dim spgm As SharedPropertyGroupManager
    Dim spg As SharedPropertyGroup
    Dim sp As SharedProperty
    Dim llngIsolationMode As Long
    Dim llngReleaseMode As Long
    Dim lblnKeySupport As Boolean
    
On Error GoTo eh
    If v_strGroupName = "" Then
        r_blnExists = False
        Set GetSharedPropertyGroup = Nothing
        Exit Function
    End If
    
    If v_blnCreate Then
        InitializeDefaults llngIsolationMode, llngReleaseMode, lblnKeySupport
        Set spg = _
            CreateSharedPropertyGroup(v_strGroupName, _
                llngIsolationMode, llngReleaseMode, lblnKeySupport, r_blnExists)
    Else
        Set spgm = GetSharedPropertyGroupManager
        On Error Resume Next
        Set spg = spgm.Group(v_strGroupName)
        If Err.Number <> 0 Then
            If Err.Number = 5 Then
                r_blnExists = False
                Err.Clear
                Exit Function
            Else
                GoTo eh
            End If
        Else
            r_blnExists = True
        End If
        On Error GoTo eh
    End If
    
    Set GetSharedPropertyGroup = spg
    Set spg = Nothing
    Set sp = Nothing
    Set spgm = Nothing
    Exit Function
    
eh:
    Set spg = Nothing
    Set sp = Nothing
    Set spgm = Nothing
    TrsRaiseError mstrCMODULENAME, "GetSharedPropertyGroup"
End Function

Private Function InitializeDefaults( _
                    ByRef r_lngIsolationMode As Long, _
                    ByRef r_lngReleaseMode As Long, _
                    ByRef r_blnSupportKeys As Boolean)

On Error GoTo eh

    r_lngIsolationMode = CDEFAULTISOLATIONMODE
    r_lngReleaseMode = CDEFAULTRELEASEMODE
    r_blnSupportKeys = CDEFAULTKEYSUPPORT

    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "InitializeDefaults"
End Function


Private Function CreateSharedPropertyGroup( _
                    ByVal v_strGroupName As String, _
                    ByRef r_lngIsolationMode As Long, _
                    ByRef r_lngReleaseMode As Long, _
                    ByRef r_blnSupportKeys As Boolean, _
                    ByRef r_blnExists As Boolean) _
                As SharedPropertyGroup
'Ok to initialize here with both LockMethod and LockGetSet.
    Dim spgm As SharedPropertyGroupManager
    Dim spg As SharedPropertyGroup
    Dim sp As SharedProperty
    Dim lblnExists As Boolean
    
On Error GoTo eh
    
    If v_strGroupName = "" Then
        r_blnExists = False
        Set CreateSharedPropertyGroup = Nothing
        Exit Function
    End If
    
    #If TRDEBUG Then
        r_lngIsolationMode = CDEBUGISOLATIONMODE
    #End If
    
    Set spgm = GetSharedPropertyGroupManager
    Set spg = _
        spgm.CreatePropertyGroup(v_strGroupName, r_lngIsolationMode, r_lngReleaseMode, r_blnExists)
    
'    If it did exist, return the previously initialized key support,
'    otherwise initialize the properties.
    If r_blnExists Then
        r_blnSupportKeys = IsKeySupportEnabled(spg)
    Else
        Set sp = spg.CreateProperty(CITEMCOUNTKEY, lblnExists)
        sp.Value = 0
        Set sp = spg.CreateProperty(CITEMKEYS, lblnExists)
        sp.Value = Array()
        Set sp = spg.CreateProperty(CSUPPORTKEYS, lblnExists)
        sp.Value = r_blnSupportKeys
    End If
    Set CreateSharedPropertyGroup = spg
    Set spg = Nothing
    Set sp = Nothing
    Set spgm = Nothing
    Exit Function
    
eh:
    Set spg = Nothing
    Set sp = Nothing
    Set spgm = Nothing
    TrsRaiseError mstrCMODULENAME, "CreateSharedPropertyGroup"
End Function


Private Function GetSharedPropertyGroupManager() As SharedPropertyGroupManager
    
On Error GoTo eh
    Set GetSharedPropertyGroupManager = _
        CreateObject("mtxspm.SharedPropertyGroupManager.1")
    Exit Function
eh:
    TrsRaiseError mstrCMODULENAME, "GetSharedPropertyGroupManager"
End Function

Private Function RemoveBasic( _
        ByVal v_strKey As String, _
        ByVal v_strGroupName As String) As Long
'Logically remove a shared property value.
    Dim sp As SharedProperty
    Dim spg As SharedPropertyGroup
    Dim lblnExists As Boolean
    Dim lvarKeysArray As Variant
    Dim lstrGroupName As String
    Dim lstrKey As String
    Dim colItems As Collection
    Dim i As Integer
    Dim llngKeyCount As Long
    Dim lblnKeySupportEnabled As Boolean
    
On Error GoTo eh
    lstrKey = Trim(v_strKey)
    lstrGroupName = Trim(v_strGroupName)

    Set spg = GetSharedPropertyGroup(lstrGroupName, lblnExists)
    lblnKeySupportEnabled = IsKeySupportEnabled(spg)
    lvarKeysArray = Array()
    Set sp = spg.CreateProperty(lstrKey, lblnExists)
    If lblnExists Then
        If Not DoesNotExist(sp.Value) Then
            lvarKeysArray = Array(lstrKey)
        End If
    End If
    sp.Value = CDOESNOTEXIST
    
    If ArraySize(lvarKeysArray) < 1 _
        Or Not lblnKeySupportEnabled Then
        GoTo xit
    End If
    
    ' update the keys and count
    ' get the keys
    Set sp = spg.CreateProperty(CITEMKEYS, lblnExists)
    If lblnExists Then
        Set colItems = _
            CollectionFromArray(sp.Value, LBound(sp.Value), UBound(sp.Value), True)
    Else
        Set colItems = New Collection
    End If
    ' update the keys
    For i = LBound(lvarKeysArray) To UBound(lvarKeysArray)
        colItems.Remove lvarKeysArray(i)
    Next
    llngKeyCount = colItems.Count
    If llngKeyCount = 0 Then
        sp.Value = Array()
    Else
        sp.Value = ArrayFromCollection(colItems)
    End If
    ' update the count
    Set sp = spg.CreateProperty(CITEMCOUNTKEY, lblnExists)
    sp.Value = llngKeyCount

xit:
    Set sp = Nothing
    Set spg = Nothing
    Exit Function
    
eh:
    Set sp = Nothing
    Set spg = Nothing
    TrsRaiseError mstrCMODULENAME, "RemoveBasic"
End Function

Private Function DoesNotExist(ByVal v_varData As Variant) As Boolean

On Error GoTo eh
    If VarType(v_varData) = vbString Then
        If v_varData = CDOESNOTEXIST Then
            DoesNotExist = True
        End If
    End If
    Exit Function
    
eh:
    TrsRaiseError mstrCMODULENAME, "DoesNotExist"
End Function

Private Sub Class_Initialize()
'Must LockSetGet when debugging.
'    mlngIsolationMode = LockSetGet
'    mlngIsolationMode = LockMethod
'    mlngReleaseMode = Process
'    mblnSupportKeys = False
'    #If TRDEBUG Then
'        mblnDebugSession = True
'    #End If
End Sub


