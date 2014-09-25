'************************************
' Module: Main - PSS Variants DB
' Created By: Mariano Bullio, bullio.m@pg.com
'**************************************


Public Class Main
    Inherits System.Windows.Forms.Form
    Private Control = Nothing
    Private SF As New System_Functions
    Private SQL_F As New SQL_Functions

    Private Sub Main_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        'System Titlem
        Me.Text = Project_Name

        'Return User TNumber
        Login = SF.ReturnUserTNumber()
        Login_Name = SF.ReturnUserName(Login)

        'Delete Temporal Files
        If Not Debugger.IsAttached Then
            SF.DeleteDirectoryFiles(EXCELPath)
        End If

        'Review User Credentials
        If Login <> Nothing Then

            'If User has access
            If Not SF.ValidateUser(Login) Then
                Lock_Controls()
                lb_User.Text = "Unknown"
                lb_TNumber.Text = "Unknown"
                MsgBox("User " & Login & " is not Registered!" & vbCr & Standart_MessageError, MsgBoxStyle.Exclamation)
            Else

                Login_IsAdministrator = SF.IsAdministrator(Login)
                Login_IsVariantsSPOC = SF.IsVariantsSPOC(Login)
                Login_IsInfosysEmployee = SF.IsInfosysEmployee(Login)
                Login_IsReadOnly = SF.IsReadOnly(Login)

                'If User is not Administrator
                If Not Login_IsAdministrator Then

                    PSSVariantsToolStripMenuItem.Enabled = False

                    If (Login_IsAdministrator = False) And (Login_IsVariantsSPOC = False) And (Login_IsInfosysEmployee = False) And (Login_IsReadOnly = False) Then
                        Lock_Controls()
                        lb_User.Text = "Unknown"
                        lb_TNumber.Text = "Unknown"
                        MsgBox("User " & Login & " is not registered [Chek User Role]!" & vbCr & Standart_MessageError, MsgBoxStyle.Exclamation)
                    End If

                    'If Login_IsVariantsSPOC Then
                    ''Review Access to Variants
                    'If Not SF.ValidateAccessToVariant(Login, "Var_STR") Then
                    '    LK_Americas_STR.Enabled = False
                    'End If

                    'If Not SF.ValidateAccessToVariant(Login, "Var_SS") Then
                    '    LK_Americas_SS.Enabled = False
                    'End If

                    'If Not SF.ValidateAccessToVariant(Login, "Var_Direct") Then
                    '    LK_Americas_Direct.Enabled = False
                    'End If

                    'If Not SF.ValidateAccessToVariant(Login, "Var_Custom") Then
                    '    LK_Americas_Custom.Enabled = False
                    'End If

                    'If Not SF.ValidateAccessToVariant(Login, "Var_LogIndNAPD") Then
                    '    LK_LogIndNAPD.Enabled = False
                    'End If

                    'If Not SF.ValidateAccessToVariant(Login, "Var_LogTMS") Then
                    '    LK_LogTMS.Enabled = False
                    'End If

                    'If Not SF.ValidateAccessToVariant(Login, "Var_LogImpLA") Then
                    '    LK_LogImpLA.Enabled = False
                    'End If

                    'If Not SF.ValidateAccessToVariant(Login, "Var_LogImpNA") Then
                    '    LK_LogImpNA.Enabled = False
                    'End If

                    'If Not SF.ValidateAccessToVariant(Login, "Var_LogDirect") Then
                    '    LK_LogDirect.Enabled = False
                    'End If
                    'End If

            End If
            lb_User.Text = Login_Name
            lb_TNumber.Text = Login

            End If

        Else
            lb_User.Text = "Unknown"
            lb_TNumber.Text = "Unknown"
            MsgBox("User " & Login & " is not registered!" & vbCr & "Please contact your system administrator.", MsgBoxStyle.Exclamation)
        End If

        'System Version
        Main_ToolStripStatusLabel.Text = Project_Version
        Me.WindowState = FormWindowState.Normal
    End Sub

    Private Sub Main_Shown(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Shown

        'Review New Version  
        'If Not Debugger.IsAttached Then
        '    Dim U As New UpdateClass("PSSI_BI", My.Application.Info.Title, My.Application.Info.Version)
        '    U.Update()
        'End If

      
    End Sub

    '"Control Subs"

    Private Sub Lock_Controls()
        MenuStrip.Enabled = False
        GB_Variants.Enabled = False
        GB_Maintenance.Enabled = False
        GB_Reports.Enabled = False
    End Sub

    Private Sub Wait_Mode(ByVal Wait As Boolean)
        If Wait Then
            Main_ToolStripStatusLabel.Text = "Please wait..."
            Main_ToolStripProgressBar.Style = ProgressBarStyle.Marquee
            Me.Refresh()
        Else
            Main_ToolStripProgressBar.Style = ProgressBarStyle.Blocks
            Main_ToolStripStatusLabel.Text = ""
            Me.Refresh()
        End If

    End Sub

    Private Sub Show_Progress(ByVal Msg As String, ByVal Percent As Integer)
        Main_ToolStripStatusLabel.Text = Msg
        Main_ToolStripProgressBar.Value = Percent
        Me.Refresh()
    End Sub

    Private Sub LoadControl()

        Control.Dock = DockStyle.Fill
        Me.Panel1.Controls.Add(Control)
        Control.BringTofront()

    End Sub

    Public Sub UnloadControl()

        Me.Panel1.Controls.Remove(Control)
        Control = Nothing
        GC.Collect()
        Me.Refresh()

    End Sub

    '"Open Modules"

    Private Sub OpenMenu()
        Try
            'Return User TNumber
            Login = SF.ReturnUserTNumber()

            Panel2.Visible = True
            Panel3.Visible = True
            Panel6.Visible = True
            Wait_Mode(True)
            UnloadControl()

            Wait_Mode(False)
        Catch ex As Exception
            MsgBox("Error to Open Menu: " & ex.Message)
        End Try
    End Sub

    Private Sub Show_Menu(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MainMenuToolStripMenuItem.Click
        OpenMenu()
    End Sub

    Private Sub AmericasSTR_MenuItem(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AmericasSTRToolStripMenuItem.Click
        Open_Americas_STR()
    End Sub

    Private Sub AmericasSS_MenuItem(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AmericasSSToolStripMenuItem.Click
        Open_Americas_SS()
    End Sub

    Private Sub AmericasSS_SubVariants_MenuItem(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SubVariantsToolStripMenuItem.Click
        Americas_SS_SubVariants.Show()
    End Sub

    Private Sub AmericasDirect_MenuItem(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AmericasDirectToolStripMenuItem.Click
        Open_Americas_Direct()
    End Sub

    Private Sub AmericasDirect_SubVariants_MenuItem(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SubVariantsToolStripMenuItem1.Click
        Americas_Direct_SubVariants.Show()
    End Sub

    Private Sub AmericasCustomization_MenuItem(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AmericasCustomizationToolStripMenuItem.Click
        Open_Americas_Custom()
    End Sub

    Private Sub AmericasCustomization_SubVariantsMenuItem(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SubVarianToolStripMenuItem.Click
        Americas_Custom_SubVariants.Show()
    End Sub

    Private Sub AmericasLogIndNAPD_MenuItem(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles IndirectNAPDToolStripMenuItem.Click
        Open_Americas_LogIndNAPD()
    End Sub

    Private Sub AmericasLogTMS_MenuItem(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TMSNALAToolStripMenuItem.Click
        Open_Americas_LogTMS()
    End Sub

    Private Sub AmericasLogImpLA_MenuItem(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ImportsLAToolStripMenuItem.Click
        Open_Americas_LogImpLA()
    End Sub

    Private Sub AmericasLogImpNA_MenuItem(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ImportsNAToolStripMenuItem.Click
        Open_Americas_LogImpNA()
    End Sub

    Private Sub AmericasLogDirect_MenuItem(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DirectToolStripMenuItem.Click
        Open_Americas_LogDirect()
    End Sub

    Private Sub AmericasLogDirect_SubVariantsMenuItem(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SubVariantsToolStripMenuItem3.Click
        Americas_LogDirect_SubVariants.Show()
    End Sub

    Private Sub Show_About(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AboutToolStripMenuItem.Click
        AboutBox.Show()
    End Sub

    Private Sub Close_DB(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CloseToolStripMenuItem.Click

        If SF.ShowConfirmationMessage Then
            End
        End If

    End Sub

    Private Sub Open_Americas_STR()
        Try
            Wait_Mode(True)
            Dim C As New Americas_STR
            If C.Initialize() Then
                UnloadControl()
                Control = C
                LoadControl()
                Panel2.Visible = False
                Panel3.Visible = False
                Panel5.Visible = False
                Panel6.Visible = False
            End If
            Wait_Mode(False)
            'End If
        Catch ex As Exception
            MsgBox("Error to Open Americas STR: " & ex.Message, MsgBoxStyle.Exclamation)
        End Try
    End Sub

    Private Sub Open_Americas_SS()
        Try
            Wait_Mode(True)
            Dim C As New Americas_SS
            If C.Initialize() Then
                UnloadControl()
                Control = C
                LoadControl()
                Panel2.Visible = False
                Panel3.Visible = False
                Panel5.Visible = False
                Panel6.Visible = False
            End If
            Wait_Mode(False)
            'End If
        Catch ex As Exception
            MsgBox("Error to Open Upload Americas SS: " & ex.Message, MsgBoxStyle.Exclamation)
        End Try
    End Sub

    Private Sub Open_Americas_Direct()
        Try
            Wait_Mode(True)
            Dim C As New Americas_Direct
            If C.Initialize() Then
                UnloadControl()
                Control = C
                LoadControl()
                Panel2.Visible = False
                Panel3.Visible = False
                Panel5.Visible = False
                Panel6.Visible = False
            End If
            Wait_Mode(False)
            'End If
        Catch ex As Exception
            MsgBox("Error to Open Americas Direct: " & ex.Message, MsgBoxStyle.Exclamation)
        End Try
    End Sub

    Private Sub Open_Americas_Custom()
        Try
            Wait_Mode(True)
            Dim C As New Americas_Custom
            If C.Initialize() Then
                UnloadControl()
                Control = C
                LoadControl()
                Panel2.Visible = False
                Panel3.Visible = False
                Panel5.Visible = False
                Panel6.Visible = False
            End If
            Wait_Mode(False)
            'End If
        Catch ex As Exception
            MsgBox("Error to Open Americas Custom: " & ex.Message, MsgBoxStyle.Exclamation)
        End Try
    End Sub

    Private Sub Open_Americas_LogIndNAPD()
        Try
            Wait_Mode(True)
            Dim C As New Americas_LogIndNAPD
            If C.Initialize() Then
                UnloadControl()
                Control = C
                LoadControl()
                Panel2.Visible = False
                Panel3.Visible = False
                Panel5.Visible = False
                Panel6.Visible = False
            End If
            Wait_Mode(False)
            'End If
        Catch ex As Exception
            MsgBox("Error to Open Logistics Indirect NAPD: " & ex.Message, MsgBoxStyle.Exclamation)
        End Try
    End Sub

    Private Sub Open_Americas_LogTMS()
        Try
            Wait_Mode(True)
            Dim C As New Americas_LogTMS
            If C.Initialize() Then
                UnloadControl()
                Control = C
                LoadControl()
                Panel2.Visible = False
                Panel3.Visible = False
                Panel5.Visible = False
                Panel6.Visible = False
            End If
            Wait_Mode(False)
            'End If
        Catch ex As Exception
            MsgBox("Error to Open Logistics TMS NALA: " & ex.Message, MsgBoxStyle.Exclamation)
        End Try
    End Sub

    Private Sub Open_Americas_LogImpLA()
        Try
            Wait_Mode(True)
            Dim C As New Americas_LogImpLA
            If C.Initialize() Then
                UnloadControl()
                Control = C
                LoadControl()
                Panel2.Visible = False
                Panel3.Visible = False
                Panel5.Visible = False
                Panel6.Visible = False
            End If
            Wait_Mode(False)
            'End If
        Catch ex As Exception
            MsgBox("Error to Open Logistics Imports LA: " & ex.Message, MsgBoxStyle.Exclamation)
        End Try
    End Sub

    Private Sub Open_Americas_LogImpNA()
        Try
            Wait_Mode(True)
            Dim C As New Americas_LogImpNA
            If C.Initialize() Then
                UnloadControl()
                Control = C
                LoadControl()
                Panel2.Visible = False
                Panel3.Visible = False
                Panel5.Visible = False
                Panel6.Visible = False
            End If
            Wait_Mode(False)
            'End If
        Catch ex As Exception
            MsgBox("Error to Open Logistics Imports NA: " & ex.Message, MsgBoxStyle.Exclamation)
        End Try
    End Sub

    Private Sub Open_Americas_LogDirect()
        Try
            Wait_Mode(True)
            Dim C As New Americas_LogDirect
            If C.Initialize() Then
                UnloadControl()
                Control = C
                LoadControl()
                Panel2.Visible = False
                Panel3.Visible = False
                Panel5.Visible = False
                Panel6.Visible = False
            End If
            Wait_Mode(False)
            'End If
        Catch ex As Exception
            MsgBox("Error to Open Logistics Direct " & ex.Message, MsgBoxStyle.Exclamation)
        End Try
    End Sub


    '"Links"
    Private Sub LK_Americas_STR_(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LK_Americas_STR.LinkClicked
        Open_Americas_STR()
    End Sub

    Private Sub LK_Americas_SS_(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LK_Americas_SS.LinkClicked
        Open_Americas_SS()
    End Sub

    Private Sub LK_Americas_Direct_(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LK_Americas_Direct.LinkClicked
        Open_Americas_Direct()
    End Sub

    Private Sub LK_Americas_Custom_(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LK_Americas_Custom.LinkClicked
        Open_Americas_Custom()
    End Sub

    Private Sub LK_LogIndNAPD_(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LK_LogIndNAPD.LinkClicked
        Open_Americas_LogIndNAPD()
    End Sub

    Private Sub LK_LogTMS_(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LK_LogTMS.LinkClicked
        Open_Americas_LogTMS()
    End Sub

    Private Sub LK_LogImpLA_(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LK_LogImpLA.LinkClicked
        Open_Americas_LogImpLA()
    End Sub

    Private Sub LK_LogImpNA_(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LK_LogImpNA.LinkClicked
        Open_Americas_LogImpNA()
    End Sub

    Private Sub LK_LogDirect_(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LK_LogDirect.LinkClicked
        Open_Americas_LogDirect()
    End Sub

    Private Sub LK_ExportVariants_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LK_ExportVariants.LinkClicked
        If SF.ExportVariantToExcel("ALL", My.Computer.FileSystem.SpecialDirectories.Desktop & "\PSS Variants (" & Date.Now.Month & "-" & Date.Now.Day & "-" & Date.Now.Year & ").xlsx") Then
            MsgBox("Done: The Report has been created on your desk", MsgBoxStyle.Information)
        Else
            MsgBox(Standart_MessageError, MsgBoxStyle.Exclamation)
        End If
    End Sub

End Class

