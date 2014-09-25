
Public Class Data_Downloader

    Dim SQL_F As New SQL_Functions
    Dim DT_Passwords As New DataTable
    Dim SF As New System_Functions

    Dim SAP_Login As String = ""
    Dim SAP_Password As String = ""
    Dim SAP_Box As String = ""
    Dim SAP_Client As String = ""

    Private Sub Data_Downloader_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        BackgroundWorker1.RunWorkerAsync()
    End Sub

    Private Sub Run_Reports()

        DownloadBI_Report("G4P") 'AP Trade

        DownloadS161_Report("G4P") 'ZFI2
        DownloadS161_Report("GBP") 'ZFI2
        DownloadS161_Report("L6P") 'ZFI2
        DownloadS161_Report("L7P") 'ZFI2
        DownloadS161_Report("N6P") 'ZFI2

        DownloadIS_Report("GRP", "NA") 'BW-IFR
        DownloadIS_Report("GRP", "LA") 'BW-IFR

        'Done
        BackgroundWorker1.ReportProgress(100)
        System.Threading.Thread.Sleep(1)

    End Sub

    Private Sub DownloadBI_Report(ByVal Box As String)
        Try
            SAP_Box = Box
            SAP_Login = DT_Passwords.Select("Box = '" & SAP_Box & "'")(0)("Login")
            SAP_Password = DT_Passwords.Select("Box = '" & SAP_Box & "'")(0)("Password")
            SAP_Client = DT_Passwords.Select("Box = '" & SAP_Box & "'")(0)("Client")

            BackgroundWorker1.ReportProgress(10)
            System.Threading.Thread.Sleep(1)

            Dim BI_Worker As New Get_Reports(SAP_Box, SAP_Client, SAP_Login, SAP_Password)

            BI_Worker.Download_BI_Reports()
            BI_Worker.Read_BI_Reports()
            BI_Worker.DownloadMaterialGroup_BI(DT_Passwords)
            BI_Worker.Distribute_BI_Items()

        Catch ex As Exception
            SF.SetLogInfo("Error Downloading BI Report From " & Box & " (Logic)")
        End Try
    End Sub

    Private Sub DownloadS161_Report(ByVal Box As String)
        Try
            SAP_Box = Box
            SAP_Login = DT_Passwords.Select("Box = '" & SAP_Box & "'")(0)("Login")
            SAP_Password = DT_Passwords.Select("Box = '" & SAP_Box & "'")(0)("Password")
            SAP_Client = DT_Passwords.Select("Box = '" & SAP_Box & "'")(0)("Client")

            BackgroundWorker1.ReportProgress(30)
            System.Threading.Thread.Sleep(1)

            Dim S161_Worker As New Get_Reports(SAP_Box, SAP_Client, SAP_Login, SAP_Password)

            S161_Worker.Download_S161_Report()
            S161_Worker.Read_S161_Report()
            S161_Worker.DownloadMaterialGroup_S161(DT_Passwords)
            S161_Worker.Distribute_S161_Items()

        Catch ex As Exception
            SF.SetLogInfo("Error Downloading S161 Report From " & Box & " (Logic)")
        End Try
    End Sub

    Private Sub DownloadIS_Report(ByVal Box As String, ByRef Region As String)
        Try

            SAP_Box = Box
            SAP_Login = DT_Passwords.Select("Box = '" & SAP_Box & "'")(0)("Login")
            SAP_Password = DT_Passwords.Select("Box = '" & SAP_Box & "'")(0)("Password")
            SAP_Client = DT_Passwords.Select("Box = '" & SAP_Box & "'")(0)("Client")

            Dim IS_Worker As New Get_Reports(SAP_Box, SAP_Client, SAP_Login, SAP_Password)

            BackgroundWorker1.ReportProgress(50)
            System.Threading.Thread.Sleep(1)

            IS_Worker.KillInternetExplorer()

            Dim DT_BW As New DataTable
            DT_BW = IS_Worker.Download_IS_Report(Region)

            If DT_BW.Rows.Count > 0 Then
                IS_Worker.Read_IS_Report(Region, DT_BW)
                IS_Worker.DownloadMaterialGroup_IS(DT_Passwords, Region)
                IS_Worker.Distribute_IS_Items()
            Else
                SF.SetLogInfo("Error Downloading IS Report From BW-IFR " & Region & "(Logic)")
                DownloadIS_Report(Box, Region)
            End If

        Catch ex As Exception
            SF.SetLogInfo("Error Downloading IS Report From BW-IFR " & Region & " (Logic)")
            DownloadIS_Report(Box, Region)
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork

        BackgroundWorker1.ReportProgress(0)
        System.Threading.Thread.Sleep(1)

        'Delete Current Log
        SQL_F.SQL_Execute_NQ(CS, "Delete From P2P_Download_Log")

        'Delete Temporal Files
        If Not Debugger.IsAttached Then
            SF.DeleteDirectoryFiles(EXCELPath)
        End If

        'Register Log
        SF.SetLogInfo("******** Starting Process(" & Today.Month & "-" & Today.Day & "-" & Today.Year & ")********)")

        'Read SAP Passwords
        DT_Passwords = SQL_F.GetDataTable(CS, "Select * From P2P_SAP_Boxes")

        'Test SAP Passwords
        SF.SetLogInfo("******** Testing (SAP Passwords) ********")
        If Test_SAPBoxes(DT_Passwords) Then
            Run_Reports()
        Else
            'Register Log
            SQL_F.SQL_Execute_NQ(CS, "Insert Into P2P_Download_Log(Date,Description) Values(GetDate(),'******** Error (Check SAP Passwords) ********')")
        End If
        SF.SetLogInfo("******** END ********")
    End Sub

    Private Sub BackgroundWorker1_ProgressChanged(ByVal sender As System.Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged
        ProgressBar1.Value = (e.ProgressPercentage)

        Select Case (e.ProgressPercentage)
            Case 0
                TextLabel.Text = "Starting " & (e.ProgressPercentage) & "%"
            Case 10
                TextLabel.Text = "Downloading BI Reports (AP-Trade)" & (e.ProgressPercentage) & "%"
            Case 30
                TextLabel.Text = "Downloading S161 Reports (ZFI2)" & (e.ProgressPercentage) & "%"
            Case 50
                TextLabel.Text = "Downloading IS Report (BW-IFR)" & (e.ProgressPercentage) & "%"
                'Case 70
                '    TextLabel.Text = "Downloading Records (L6P) " & (e.ProgressPercentage) & "%"
                'Case 90
                '    TextLabel.Text = "Downloading Records (L7P) " & (e.ProgressPercentage) & "%"
            Case 100
                TextLabel.Text = "Done " & (e.ProgressPercentage) & "%"
            Case Else
        End Select
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As System.Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        CloseMe()
    End Sub

    Public Function Test_SAPBoxes(ByRef DT_Boxes As DataTable) As Boolean
        Dim Test As Boolean = True

        Try

            For Each R_Box As DataRow In DT_Boxes.Rows

                'Test Connection
                Dim D As New SAPCOM.ConnectionData
                If R_Box("Box") = "ANP" Then
                    D.Box = R_Box("Box") & "_" & R_Box("Client")
                Else
                    D.Box = R_Box("Box")
                End If

                D.Login = R_Box("Login")
                D.Password = R_Box("Password")
                D.SSO = False

                Dim SC As New SAPCOM.SAPConnector
                Dim Con = SC.TestConnection(D)

                If Con Then
                    Dim Update As String = "Update P2P_SAP_Boxes set Test_Date=GetDate(), Test_Done=1 Where Box='" & R_Box("Box") & "' And Client=" & R_Box("Client")
                    Dim Exc As String = SQL_F.SQL_Execute_NQ(CS, Update)
                Else
                    Dim Update As String = "Update P2P_SAP_Boxes set Test_Date=GetDate(), Test_Done=0 Where Box='" & R_Box("Box") & "' And Client=" & R_Box("Client")
                    Dim Exc As String = SQL_F.SQL_Execute_NQ(CS, Update)
                    Test = False
                End If

            Next

            Test_SAPBoxes = Test
        Catch ex As Exception
            Test_SAPBoxes = False
        End Try

    End Function

    Private Sub CloseMe()
        Me.Close()
    End Sub

End Class
