Public Class UserRepository
    Public Function GetUserRole(ByVal username As String) As String
        If username = "admin" Then
            Return "Administrator"
        ElseIf username = "manager" Then
            Return "Manager"
        Else
            Return "Employee"
        End If
    End Function

    Public Function IsUserActive(ByVal username As String) As Boolean
        If username.StartsWith("test") Then
            Return False
        End If
        Return True
    End Function
End Class
