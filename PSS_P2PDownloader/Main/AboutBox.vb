Public NotInheritable Class AboutBox

    Private Sub AboutBox1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        TextBoxDescription.Text = Project_Version
        Me.LabelVersion.Text = Project_Version
        Me.LabelProductName.Text = Project_Name
        Me.Text = Project_Name
    End Sub

    Private Sub OKButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OKButton.Click
        Me.Close()
    End Sub


End Class
