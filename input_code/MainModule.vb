Public Module MainModule
    Public Sub ProcessLogin(ByVal user As String)
        Dim auth As New AuthService()
        If auth.CanAccessSystem(user) Then
            Console.WriteLine("Access Granted for " & user)
            LoadDashboard(user)
        Else
            Console.WriteLine("Access Denied for " & user)
        End If
    End Sub

    Private Sub LoadDashboard(ByVal user As String)
        ' Initialize the Dashboard UI
        Console.WriteLine("Loading specialized UI components...")
    End Sub
End Module
