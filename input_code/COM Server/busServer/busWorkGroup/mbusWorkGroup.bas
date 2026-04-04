Attribute VB_Name = "mbusWorkGroup"
Option Explicit


Private Const m_strCMODULENAME As String = "mbusWorkGroup"
Private Const m_lngTransferWG As Long = 99997
Private Const m_lngLeftServiceWG_Contract As Long = 99998
Private Const m_lngLeftServiceWG_Executive As Long = 99999
Private Const m_lngLeftServiceWG_Archive_Contract As Long = 99995
Private Const m_lngLeftServiceWG_Archive_Executive As Long = 99996



'The AmerenUE power plants and the DOJM regions have a specific range of work group numbers each of them use.  Notes copied from code and altered to reflect actual code.
'
'
'    Following is a list of the pre-assigned work group numbers:
'      Meramec   (location 55) --> 3000 - 3999         Sioux       (location 56) --> 4000 - 4999
'      Labadie   (location 57) --> 5000 - 5999         Rush Island (location 58) --> 6000 - 6999
'      Venice    (location 54) --> 7000 - 7999         Osage       (location 51) --> 8000 - 8499
'      Taum Sauk (location 83) --> 8500 - 8999         Keokuk      (location 50) --> 9000 - 9499
'      PPM       (depts 117, 180, 640,648,654, 655, 670) --> 9500 - 10999
'
'    NOTE:  Ron Belcher from T&D requested a pre-assigned block of numbers for DOJM work groups.  This
'    is handled by looking at the interface_to_cd on the dept table.  If equal to "D", set the type to
'    "DO" and retrieve the proper work group number.  These numbers range from 100 - 1999.
'
'
'    NOTE:  Ron Belcher from T&D requested a pre-assigned block of numbers for CIPS DOJM work groups.
'    This is handled by looking at the Department Number selected on the window.
'    These numbers range from 1000 - 1899.  There are not any department number for the 1000 - 1099
'    range.  Ron is reserving these but will not be using yet.
'    Following is a list of the pre-assigned work group numbers:
'      Operations Support        (Depts none)     --> 1000 - 1099
'      Northern Prairie Region       (Depts 311, 312) --> 1100 - 1199
'      Heritage Region           (Depts 341, 342) --> 1200 - 1299
'      Wabash Region         (Depts 391, 392) --> 1300 - 1399
'      Shawnee Region            (Depts 441, 442) --> 1400 - 1499
'      Southern Hill Region      (Depts 471, 472) --> 1500 - 1599
'      Midland Region            (Depts 421, 422) --> 1600 - 1699
'      Eagle View Region     (Depts 481, 482) --> 1700 - 1799
'      Four River Region     (Depts 491, 492) --> 1800 - 1899

Public Function IsPPM(ByVal v_strDept As String) As Boolean
        
'RED220

    Select Case v_strDept
    
        Case "117", "180", "554", "640", "648", "654", "655", "670"
            IsPPM = True
        Case Else
            IsPPM = False
    
    End Select
        
End Function

Public Function IsCIPSDojmWG( _
            ByVal v_lngWgNum As Long _
            ) As Boolean
 'RED222
            
    If v_lngWgNum >= 1000 And v_lngWgNum <= 1999 Then
        IsCIPSDojmWG = True
    Else
        IsCIPSDojmWG = False
    End If
            
End Function


Public Function GetWGSequenceType_UEPlant( _
            ByVal v_lngLoc As Long _
            ) As String
            
    On Error GoTo eh
    
        If IsKeokuk(v_lngLoc) Then
            GetWGSequenceType_UEPlant = "KP"
            Exit Function
        End If
        
        If IsOsage(v_lngLoc) Then
            GetWGSequenceType_UEPlant = "OP"
            Exit Function
        End If
        
        If IsVenice(v_lngLoc) Then
            GetWGSequenceType_UEPlant = "VP"
            Exit Function
        End If
        
        If IsMeramec(v_lngLoc) Then
            GetWGSequenceType_UEPlant = "MP"
            Exit Function
        End If
        If IsSioux(v_lngLoc) Then
            GetWGSequenceType_UEPlant = "SP"
            Exit Function
        End If
        If IsLabadie(v_lngLoc) Then
            GetWGSequenceType_UEPlant = "LP"
            Exit Function
        End If
        If IsRushIsland(v_lngLoc) Then
            GetWGSequenceType_UEPlant = "RP"
            Exit Function
        End If
        If IsTaumSauk(v_lngLoc) Then
            GetWGSequenceType_UEPlant = "SP"
            Exit Function
        End If
        
        Exit Function
    
eh:
    
  TrsRaiseError m_strCMODULENAME, "GetWGSequenceType_UEPlant"

End Function


'''''
'''''Public Function GetWGSequenceType_DOJM( _
'''''            ByVal v_strDept As String)
'''''
'''''
'''''    GetWGSequenceType_DOJM = "DO"
'''''
'''''    If IsNorthernPraireRegion(v_strDept) Then
'''''        GetWGSequenceType_DOJM = "NR"
'''''        Exit Function
'''''    End If
'''''
'''''    If IsHeritageRegion(v_strDept) Then
'''''        GetWGSequenceType_DOJM = "HR"
'''''        Exit Function
'''''    End If
'''''
'''''
'''''    If IsWalbashRegion(v_strDept) Then
'''''        GetWGSequenceType_DOJM = "WR"
'''''        Exit Function
'''''    End If
'''''
'''''    If IsShawneeRegion(v_strDept) Then
'''''        GetWGSequenceType_DOJM = "SR"
'''''        Exit Function
'''''    End If
'''''
'''''    If IsSouthernHillRegion(v_strDept) Then
'''''        GetWGSequenceType_DOJM = "CC"
'''''        Exit Function
'''''    End If
'''''
'''''    If IsMidlandRegion(v_strDept) Then
'''''        GetWGSequenceType_DOJM = "MR"
'''''        Exit Function
'''''    End If
'''''
'''''    If IsEagleViewRegion(v_strDept) Then
'''''        GetWGSequenceType_DOJM = "ER"
'''''        Exit Function
'''''    End If
'''''
'''''    If IsFourRiversRegion(v_strDept) Then
'''''        GetWGSequenceType_DOJM = "FR"
'''''        Exit Function
'''''    End If
'''''
'''''
'''''End Function


Public Function CanSuperviseAnyWorkGroup( _
            ByVal v_strDept As String _
            ) As Boolean
'RED217

    On Error GoTo eh

    Select Case v_strDept
        Case "500", "535", "545", "555", _
             "745", "800", "850", "900", _
             "71L", "72E"
            CanSuperviseAnyWorkGroup = True
        Case Else
            CanSuperviseAnyWorkGroup = False
    End Select
    
    Exit Function
    
eh:
    
  TrsRaiseError m_strCMODULENAME, "CanSuperviseAnyWorkGroup"
            
End Function

Public Function IsMeramec( _
        ByVal v_lngLocation As Long _
        ) As Boolean
        
    If v_lngLocation = 55 Then
        IsMeramec = True
    Else
        IsMeramec = False
    End If
        
End Function

Public Function IsSioux( _
        ByVal v_lngLocation As Long _
        ) As Boolean
        
    If v_lngLocation = 56 Then
        IsSioux = True
    Else
        IsSioux = False
    End If
        
End Function

Public Function IsLabadie( _
        ByVal v_lngLocation As Long _
        ) As Boolean
        
    If v_lngLocation = 57 Then
        IsLabadie = True
    Else
        IsLabadie = False
    End If
        
End Function

Public Function IsRushIsland( _
        ByVal v_lngLocation As Long _
        ) As Boolean
        
    If v_lngLocation = 58 Then
        IsRushIsland = True
    Else
        IsRushIsland = False
    End If
        
End Function

Public Function IsCallaway( _
        ByVal v_lngLocation As Long _
        ) As Boolean
        
    If v_lngLocation = 59 Then
        IsCallaway = True
    Else
        IsCallaway = False
    End If
        
End Function

Public Function IsVenice( _
        ByVal v_lngLocation As Long _
        ) As Boolean
        
    If v_lngLocation = 54 Then
        IsVenice = True
    Else
        IsVenice = False
    End If
        
End Function

Public Function IsOsage( _
        ByVal v_lngLocation As Long _
        ) As Boolean
        
    If v_lngLocation = 51 Then
        IsOsage = True
    Else
        IsOsage = False
    End If
        
End Function


Public Function IsKeokuk( _
        ByVal v_lngLocation As Long _
        ) As Boolean
        
    If v_lngLocation = 50 Then
        IsKeokuk = True
    Else
        IsKeokuk = False
    End If
        
End Function

Public Function IsTaumSauk( _
        ByVal v_lngLocation As Long _
        ) As Boolean
        
    If v_lngLocation = 83 Then
        IsTaumSauk = True
    Else
        IsTaumSauk = False
    End If
        
End Function

Public Function IsAmerenUEPlant( _
        ByVal v_lngLocation As Long _
    ) As Boolean
    
    
    If IsMeramec(v_lngLocation) Or _
       IsSioux(v_lngLocation) Or _
       IsLabadie(v_lngLocation) Or _
       IsRushIsland(v_lngLocation) Or _
       IsVenice(v_lngLocation) Or _
       IsOsage(v_lngLocation) Or _
       IsTaumSauk(v_lngLocation) Or _
       IsKeokuk(v_lngLocation) Then

        IsAmerenUEPlant = True
    Else
        IsAmerenUEPlant = False
    End If

End Function


Public Function GetTransferWG() As Long

    GetTransferWG = m_lngTransferWG

End Function

Public Function IsSystemDefinedWorkGroup( _
        ByVal v_lngWgNum As Long _
        ) As Boolean
        
    Select Case v_lngWgNum
    
        Case m_lngTransferWG, _
             m_lngLeftServiceWG_Contract, _
             m_lngLeftServiceWG_Executive, _
             m_lngLeftServiceWG_Archive_Contract, _
             m_lngLeftServiceWG_Archive_Executive
            IsSystemDefinedWorkGroup = True
        Case Else
            IsSystemDefinedWorkGroup = False
    End Select
        
End Function


Public Function IsUEDojmWG( _
            ByVal v_lngWgNum As Long _
            ) As Boolean
 'RED222
            
    If v_lngWgNum >= 1 And v_lngWgNum <= 999 Then
        IsUEDojmWG = True
    Else
        IsUEDojmWG = False
    End If
            
End Function



