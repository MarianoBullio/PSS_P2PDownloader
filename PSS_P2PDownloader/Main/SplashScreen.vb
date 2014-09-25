Public NotInheritable Class SplashScreen

    Private Sub SplashScreen1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        ApplicationTitle.Text = Project_Name
        lb_version.Text = Project_Version
    End Sub
End Class
