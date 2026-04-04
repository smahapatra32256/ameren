VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cuString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'
' move the sql builder functions, at bottom, here.
'
'Please note:
'Function IsStringInTheList was added to remove the
'dependancy of having to include mCollectionLib.bas
'whenever the developer needed mDataServiceLib.bas
'The function IsStringInTheList is IsStringInList
' in mCollectionLib.bas
'
' Move to common function library and use it here.
'

'
' MTS lib is included for the standard error handling bas module.
'
'



Public Function ReplaceString(varString, varFind, varReplace) As String
    ' this function will replace search strString and replace all occurences of
    ' strFind with strReplace
    ' strFind and strReplace can be of different lenghts. strReplace can be an empty string
    Dim strString As String
    Dim strFind As String
    Dim strReplace As String
    Dim strTemp As String
    Dim i As Integer

    strString = CStr(varString)
    strFind = CStr(varFind)
    strReplace = CStr(varReplace)
    i = 1
    strTemp = strString
    Do
        ' search for the find string
        i = InStr(i, strTemp, strFind)
        ' if it is not there, we are done
        If i = 0 Then Exit Do
        ' replace the find string with the replace string
        strTemp = Left(strTemp, i - 1) & strReplace & Mid(strTemp, i + Len(strFind))
        ' start next search after the string that we just inserted
        i = i + Len(strReplace)
    Loop
        
    ReplaceString = strTemp

End Function


Public Function IsStringInTheList( _
        ByVal v_strValue As String, _
        ByVal v_strList As Variant) As Boolean
'is v_strValue equal to any of values in v_strList
'v_strList should be a string array or variant string array.
    
    Dim i As Integer
    For i = LBound(v_strList) To UBound(v_strList)
        If UCase(v_strValue) = UCase(v_strList(i)) Then
            IsStringInTheList = True
            Exit Function
        End If
    Next

End Function

Public Function DateAsEnclosedString(ByVal v_datDate As Date) As String
    DateAsEnclosedString = "'" & Format(v_datDate, "mm/dd/yyyy") & "'"
End Function

Public Function SinqleQuotedString(ByVal v_strData As String) As String
    SinqleQuotedString = StringEnclosedBy(v_strData, "'")
End Function

Public Function StringEnclosedBy(ByVal v_strData As String, _
                    ByVal v_strEncloseChar As String) As String
    StringEnclosedBy = v_strEncloseChar & v_strData & v_strEncloseChar
End Function

