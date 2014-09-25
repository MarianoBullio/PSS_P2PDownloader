
Public Class EPO_Notification

    Dim SQL_F As New SQL_Functions
    Dim SF As New System_Functions

    Private Sub ConfirmationGenerator_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim p() As Process
        p = Process.GetProcessesByName("PSS Variants DB")
        If p.Count > 1 Then
            CloseMe()
        Else
            BackgroundWorker1.RunWorkerAsync()
        End If

    End Sub

    Private Sub Run_Worker()

        Dim Worker As New Emails_Notifications

        Try

            'Email Notification to EPO
            BackgroundWorker1.ReportProgress(25)
            System.Threading.Thread.Sleep(1)

            'STR
            Worker.Email_NotificationToEPO_AddDelete("STR", "Storeroom")
            Worker.Email_NotificationToInfosys_Change("STR", "Storeroom")

            'SS
            Worker.Email_NotificationToEPO_AddDelete("SS", "SelfService")
            Worker.Email_NotificationToInfosys_Change("SS", "SelfService")

            'Direct
            Worker.Email_NotificationToEPO_AddDelete("Direct", "Direct")
            Worker.Email_NotificationToInfosys_Change("Direct", "Direct")

            'Custom
            Worker.Email_NotificationToEPO_AddDelete("Custom", "Customization")
            Worker.Email_NotificationToInfosys_Change("Custom", "Customization")

            'LogIndNAPD
            Worker.Email_NotificationToEPO_AddDelete("LogIndNAPD", "Logistics Ind NAPD")
            Worker.Email_NotificationToInfosys_Change("LogIndNAPD", "Logistics Ind NAPD")

            'LogTMS
            Worker.Email_NotificationToEPO_AddDelete("LogTMS", "Logistics TMS")
            Worker.Email_NotificationToInfosys_Change("LogTMS", "Logistics TMS")

            'LogImpLA
            Worker.Email_NotificationToEPO_AddDelete("LogImpLA", "Logistics Imp LA")
            Worker.Email_NotificationToInfosys_Change("LogImpLA", "Logistics Imp LA")

            'LogImpNA
            Worker.Email_NotificationToEPO_AddDelete("LogImpNA", "Logistics Imp NA")
            Worker.Email_NotificationToInfosys_Change("LogImpNA", "Logistics Imp NA")

            'LogDirect
            Worker.Email_NotificationToEPO_AddDelete("LogDirect", "Logistics Direct")
            Worker.Email_NotificationToInfosys_Change("LogDirect", "Logistics Direct")

        Catch ex As Exception
            'F.SetLogInfo("Error: Sending Pending PTE Confirmation Email)")
        End Try

    End Sub

    Private Sub CloseMe()
        Me.Close()
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        BackgroundWorker1.ReportProgress(0)
        System.Threading.Thread.Sleep(1)
        Run_Worker()
    End Sub

    Private Sub BackgroundWorker1_ProgressChanged(ByVal sender As System.Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged
        ProgressBar1.Value = (e.ProgressPercentage)

        Select Case (e.ProgressPercentage)
            Case 0
                TextLabel.Text = "Starting " & (e.ProgressPercentage) & "%"
            Case 25
                TextLabel.Text = "Sending PTE Confirmations " & (e.ProgressPercentage) & "%"
         
            Case 100
                TextLabel.Text = "Done " & (e.ProgressPercentage) & "%"
            Case Else
        End Select
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As System.Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        CloseMe()
    End Sub

End Class
