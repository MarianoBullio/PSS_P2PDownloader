
Public Class Data_DownloaderBIDB

    Dim SQL_F As New SQL_Functions
    Dim DT_Passwords As New DataTable
    Dim SF As New System_Functions

    Dim SAP_Login As String = ""
    Dim SAP_Password As String = ""
    Dim SAP_Box As String = ""
    Dim SAP_Client As String = ""
    Dim CSBIDB As String = "Data Source=BDC-SQLD033.na.pg.com\DEVNT3310;Initial Catalog=PSSD_G_BI;Persist Security Info=True;User ID=PSSD_Admin;Password=procter" 'HP Server Connection String

    'This Class Depends on 
    '-SQL DB: Variants_DB/Table: P2P_SAP_Boxes
    '-Classes: Global.vb, GetSAPTables.vb
    'The rest of Tables are located on BI SQL Server

    Private Sub Data_DownloaderBIDB_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        BackgroundWorker1.RunWorkerAsync()
    End Sub

    Private Sub Run_Reports()
        BackgroundWorker1.ReportProgress(10)
        System.Threading.Thread.Sleep(1)
        DownloadBI_PaymentMethod("G11") 'AP Trade

        'Done
        BackgroundWorker1.ReportProgress(100)
        System.Threading.Thread.Sleep(1)

    End Sub

    Private Sub DownloadBI_PaymentMethod(ByVal Box As String)
        SQL_F.SQL_Execute_NQ(CSBIDB, "Insert Into Job_Download_Log(Date,Description) Values(GetDate(),'Downloading Payment Method (LFB1)')")
        Dim GetTables As New GetSAPTables

        Try
            SAP_Box = Box
            SAP_Login = DT_Passwords.Select("Box = '" & SAP_Box & "'")(0)("Login")
            SAP_Password = DT_Passwords.Select("Box = '" & SAP_Box & "'")(0)("Password")
            SAP_Client = DT_Passwords.Select("Box = '" & SAP_Box & "'")(0)("Client")

            Dim DT_Vendors As New DataTable
            DT_Vendors = SQL_F.GetDataTable(CSBIDB, "Select Vendor,CompanyCode From Blocked_Invoices Where (ProcessLevel in (0,1,2)) AND ((Payment_Method IS NULL) OR (Payment_Method ='')) Group by Vendor,CompanyCode")

            If Not DT_Vendors Is Nothing Then

                'New Columns
                Dim NewColumnLE As New DataColumn("LE", GetType(String))
                NewColumnLE.DefaultValue = "000"
                DT_Vendors.Columns.Add(NewColumnLE)

                If DT_Vendors.Rows.Count > 0 Then

                    For Each R As DataRow In DT_Vendors.Rows
                        Try
                            If R("CompanyCode").ToString.Trim.Length = 1 Then
                                R("LE") = "00" & R("CompanyCode")
                            ElseIf R("CompanyCode").ToString.Trim.Length = 2 Then
                                R("LE") = "0" & R("CompanyCode")
                            Else
                                R("LE") = R("CompanyCode")
                            End If
                        Catch ex As Exception
                            R("LE") = "000"
                        End Try
                    Next

                    Dim DT_LFB1 As New DataTable
                    DT_LFB1 = GetTables.Get_LFB1(SAP_Box, SAP_Login, SAP_Password, DT_Vendors)


                    If Not DT_LFB1 Is Nothing Then

                        For Each R2 As DataRow In DT_LFB1.Rows
                            Try
                                Dim Update As String = " Update Blocked_Invoices set Payment_Method='" & R2("Payment_Method").ToString.Trim & "' Where (Vendor=" & R2("Vendor").ToString.Trim & ") AND (CompanyCode=" & R2("LE").ToString.Trim & ")"
                                Dim Exc As String = SQL_F.SQL_Execute_NQ(CSBIDB, Update)
                                If Exc <> Nothing Then
                                    SQL_F.SQL_Execute_NQ(CSBIDB, "Insert Into Job_Download_Log(Date,Description) Values(GetDate(),'Error: Vendor-LE: Data no ingresada:" & Exc & " SQL: " & Update.Replace("'", "?") & "')")
                                End If
                            Catch ex As Exception
                                SQL_F.SQL_Execute_NQ(CSBIDB, "Insert Into Job_Download_Log(Date,Description) Values(GetDate(),'Error Critico: [Box: " & Box & "][LFB1/Vendor-CompanyCode]')")
                            End Try
                        Next
                    End If

                End If
            End If

        Catch ex As Exception
            SQL_F.SQL_Execute_NQ(CSBIDB, "Insert Into Job_Download_Log(Date,Description) Values(GetDate(),'Error Downloading BI Report From " & Box & " (Logic)')")
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork

        BackgroundWorker1.ReportProgress(0)
        System.Threading.Thread.Sleep(1)

        'Delete Current Log
        SQL_F.SQL_Execute_NQ(CSBIDB, "Delete From Job_Download_Log")

        'Register Log
        SQL_F.SQL_Execute_NQ(CSBIDB, "Insert Into Job_Download_Log(Date,Description) Values(GetDate(),'******** Starting Process(" & Today.Month & "-" & Today.Day & "-" & Today.Year & ")********')")

        'Read SAP Passwords
        DT_Passwords = SQL_F.GetDataTable(CS, "Select * From P2P_SAP_Boxes")

        'Test SAP Passwords
        SQL_F.SQL_Execute_NQ(CSBIDB, "Insert Into Job_Download_Log(Date,Description) Values(GetDate(),'******** Testing (SAP Passwords) ********')")
        If Test_SAPBoxes(DT_Passwords) Then
            Run_Reports()
        Else
            'Register Log
            SQL_F.SQL_Execute_NQ(CSBIDB, "Insert Into Job_Download_Log(Date,Description) Values(GetDate(),'******** Error (Check SAP Passwords) ********')")
        End If
        SQL_F.SQL_Execute_NQ(CSBIDB, "Insert Into Job_Download_Log(Date,Description) Values(GetDate(),'******** END ********')")
    End Sub

    Private Sub BackgroundWorker1_ProgressChanged(ByVal sender As System.Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged
        ProgressBar1.Value = (e.ProgressPercentage)

        Select Case (e.ProgressPercentage)
            Case 0
                TextLabel.Text = "Starting " & (e.ProgressPercentage) & "%"
            Case 10
                TextLabel.Text = "Downloading Payment Method (LFB1)" & (e.ProgressPercentage) & "%"
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
