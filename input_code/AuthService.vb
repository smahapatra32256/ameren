Public Class AuthService
    Private ReadOnly repo As UserRepository

    Public Sub New()
        Me.repo = New UserRepository()
    End Sub

    Public Function CanAccessSystem(ByVal username As String) As Boolean
        Dim isActive As Boolean = repo.IsUserActive(username)
        If Not isActive Then
            Return False
        End If

        Dim role As String = repo.GetUserRole(username)
        If role = "Administrator" Or role = "Manager" Then
            Return True
        End If

        Return False
    End Function
End Class
