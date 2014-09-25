Imports System.ComponentModel
Imports System.Threading

Public Class SAPGUI

    Private SAPApp As Object = Nothing
    Private Connection As Object = Nothing
    Private Session As Object = Nothing

    Private LI As Boolean = False

    Public Property LoggedIn() As Boolean

        Get
            LoggedIn = LI
        End Get
        Set(ByVal value As Boolean)
            LI = value
        End Set

    End Property

    Public ReadOnly Property StatusBarText() As String

        Get
            StatusBarText = Session.findById("wnd[0]/sbar").text
        End Get

    End Property

    Public Sub New(ByVal Box As String, ByVal Client As String, ByVal User As String, ByVal Password As String, Optional ByRef NewPass As String = Nothing)

        Dim CS As String = GetConnString(Box)

        Try
            Enable_GUI_Theme()
            SAPApp = CreateObject("Sapgui.ScriptingCtrl.1")
            Connection = SAPApp.OpenConnectionByConnectionString(CS)
            Session = Connection.Children(0)
            Session.findById("wnd[0]/usr/txtRSYST-MANDT").Text = Client
            Session.findById("wnd[0]/usr/txtRSYST-BNAME").Text = User
            Session.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = Password
            Session.findById("wnd[0]").sendVKey(0)

            If Not Session.findById("wnd[1]/usr/txtMULTI_LOGON_TEXT", False) Is Nothing Then
                If Not Session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2", False) Is Nothing Then
                    Session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").Select()
                Else
                    Session.findById("wnd[1]/usr/radMULTI_LOGON_OPT1").Select()
                End If
                Session.findById("wnd[1]").sendVKey(0)
            End If

            If Not Session.findById("wnd[1]/usr/pwdRSYST-NCODE", False) Is Nothing Then
                If Not NewPass Is Nothing Then
                    Session.findById("wnd[1]/usr/pwdRSYST-NCODE").Text = NewPass
                    Session.findById("wnd[1]/usr/pwdRSYST-NCOD2").Text = NewPass
                    Session.findById("wnd[1]/tbar[0]/btn[0]").Press()
                Else
                    Session.findById("wnd[1]/tbar[0]/btn[12]").Press()
                    Exit Sub
                End If
            Else
                If Not NewPass Is Nothing Then NewPass = Nothing
            End If

            If Not Session.findById("wnd[1]/usr/txtMULTI_LOGON_TEXT", False) Is Nothing Then
                If Not Session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2", False) Is Nothing Then
                    Session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").Select()
                Else
                    Session.findById("wnd[1]/usr/radMULTI_LOGON_OPT1").Select()
                End If
                Session.findById("wnd[1]").sendVKey(0)
            End If

            Do While Not Session.findById("wnd[1]/tbar[0]/btn[0]", False) Is Nothing
                Session.findById("wnd[1]/tbar[0]/btn[0]").Press()
            Loop

            If Session.ActiveWindow.Text Like "SAP Easy Access*" Then
                LI = True
            End If

        Catch ex As Exception

        End Try

    End Sub

    Public Sub New(ByVal Box As String)

        Try
            'Enable_GUI_Theme()
            SAPApp = CreateObject("Sapgui.ScriptingCtrl.1")
            Connection = SAPApp.OpenConnection(GetSSOConnString(Box), True)
            Session = Connection.Children(0)
            Do While Not Session.findById("wnd[1]", False) Is Nothing
                If Session.ActiveWindow.Text = "SAP" Then
                    Session.findById("wnd[1]/tbar[0]/btn[12]").Press()
                    Exit Sub
                End If
                If Not Session.findById("wnd[1]/usr/txtMULTI_LOGON_TEXT", False) Is Nothing Then
                    If Not Session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2", False) Is Nothing Then
                        Session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").Select()
                    Else
                        Session.findById("wnd[1]/usr/radMULTI_LOGON_OPT1").Select()
                    End If
                End If
                Session.findById("wnd[1]").sendVKey(0)
                If Session.ActiveWindow.Text = "SAP" Then
                    Session.findById("wnd[1]/tbar[0]/btn[12]").Press()
                    Exit Sub
                End If
            Loop
            If Session.ActiveWindow.Text Like "SAP Easy Access*" Then
                LI = True
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub Enable_GUI_Theme()

        Dim PN As String = System.Diagnostics.Process.GetCurrentProcess.ProcessName
        Dim RV = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\SAP\General\Applications\" & PN, "Enjoy", Nothing)
        If RV Is Nothing Then
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\SAP\General\Applications\" & PN, "Enjoy", "On")
        End If

    End Sub

    Private Function GetSSOConnString(ByVal Box As String) As String

        GetSSOConnString = ""
        Select Case Box
            Case "L6A"
                GetSSOConnString = "L6A LA SC Acc - SSO"
            Case "L6P"
                GetSSOConnString = "L6P LA SC  Prod - SSO"
            Case "N6P"
                GetSSOConnString = "N6P NA Prod- SSO"
            Case "N6A"
                GetSSOConnString = "N6A NA SC Acc - SSO"
            Case "F6P"
                GetSSOConnString = "F6P EU SC Prod - SSO"
            Case "ANP"
                GetSSOConnString = "ANP NEA Prod(JP) - SSO"
            Case "A6P"
                GetSSOConnString = "A6P SC Prod(EN) - SSO"
            Case "G4P"
                GetSSOConnString = "G4P GCF/Cons Prod- SSO"
            Case "G4A"
                GetSSOConnString = "G4A GCF Acc - SSO"
        End Select

    End Function

    Public Sub Close()

        If LI Then
            ActiveWindow.Close()
            If Not FindByNameEx("SPOP-OPTION1", 40) Is Nothing Then
                FindByNameEx("SPOP-OPTION1", 40).press()
            End If
        End If

    End Sub

    Public Function FindById(ByVal ID As String) As Object

        Try
            FindById = Session.findById(ID, False)
        Catch ex As Exception
            FindById = Nothing
        End Try

    End Function

    Public Function FindByNameEx(ByVal Name As String, ByVal Type As Long) As Object

        Try
            FindByNameEx = Session.ActiveWindow.FindByNameEx(Name, Type)
        Catch ex As Exception
            FindByNameEx = Nothing
        End Try

    End Function

    Public Function FindAllByNameEx(ByVal Name As String, ByVal Type As Long) As Object

        Try
            FindAllByNameEx = Session.ActiveWindow.FindAllByNameEx(Name, Type)
        Catch ex As Exception
            FindAllByNameEx = Nothing
        End Try

    End Function

    Public Function ActiveWindow() As Object

        ActiveWindow = Session.ActiveWindow

    End Function

    Public Function GuiFocus() As Object

        GuiFocus = ActiveWindow.GuiFocus

    End Function

    Public Sub SendVKey(ByVal Code As Integer)

        Session.ActiveWindow.SendVKey(Code)

    End Sub

    Public Sub StartTransaction(ByVal Code As String)

        Session.StartTransaction(Code)

    End Sub

    Public Sub SendCommand(ByVal Code As String)

        Session.SendCommand(Code)

    End Sub

    Public Function DisplayPO(ByVal PO As String) As Boolean


        If Not LI Then
            DisplayPO = False
            Exit Function
        End If

        DisplayPO = True
        StartTransaction("me23n")
        FindByNameEx("btn[17]", 40).Press()
        FindByNameEx("MEPO_SELECT-EBELN", 32).Text = PO
        FindByNameEx("btn[0]", 40).Press()
        If StatusBarText <> "" Then
            DisplayPO = False
        End If

    End Function

    Public Function ChangePO(ByVal PO As String) As Boolean

        ChangePO = True
        If Not DisplayPO(PO) Then
            ChangePO = False
        Else
            FindByNameEx("btn[7]", 40).Press()
            If StatusBarText <> "" AndAlso StatusBarText <> "Text contains formatting -> SAPscript editor" Then
                ChangePO = False
            End If
        End If

    End Function

    Public Function GetConnString(ByVal Box As String) As String

        Dim R As String = "/R/*/G/SPACE"
        GetConnString = R.Replace("*", Box)

    End Function



    'Public Function GetConnString(ByVal Box As String) As String

    '    Dim R As String = Nothing

    '    Select Case Box
    '        Case "L6A"
    '            R = "L6A LA SC Acc - SSO"
    '        Case "L6P"
    '            R = "L6P LA SC  Prod - SSO"
    '        Case "N6P"
    '            R = "N6P NA Prod- SSO"
    '        Case "N6A"
    '            R = "N6A NA SC Acc - SSO"
    '        Case "F6P"
    '            R = "F6P EU SC Prod - SSO"
    '        Case "ANP"
    '            R = "ANP NEA Prod(JP) - SSO"
    '        Case "A6P"
    '            R = "A6P SC Prod(EN) - SSO"
    '        Case "L7P"
    '            R = "L7P LA TS Prod - SSO"
    '        Case "G4P"
    '            R = "G4P GCF/Cons Prod- SSO"
    '        Case "GBP"
    '            R = "GBP GCM Production- SSO"
    '    End Select

    '    GetConnString = R

    'End Function

    Public Sub ArrayToClipboard(ByVal A() As String)

        Dim S As String = Nothing
        Dim C As String = Nothing
        My.Computer.Clipboard.Clear()
        For Each S In A
            C = C & S & Chr(13) & Chr(10)
        Next
        My.Computer.Clipboard.SetText(C)

    End Sub

    Public Sub TableToClipboard(ByVal DT As DataTable, ByVal ColumnIndex As Integer)

        Dim S As String = Nothing
        Dim DR As DataRow

        My.Computer.Clipboard.Clear()
        For Each DR In DT.Rows
            S = S & DR(ColumnIndex) & Chr(13) & Chr(10)
        Next
        My.Computer.Clipboard.SetText(S)

    End Sub

    Public Sub DRArrayToClipboard(ByVal DRA() As DataRow, ByVal ColumnIndex As Object)

        Dim S As String = Nothing
        Dim DR As DataRow

        My.Computer.Clipboard.Clear()
        For Each DR In DRA
            S = S & DR(ColumnIndex) & Chr(13) & Chr(10)
        Next
        My.Computer.Clipboard.SetText(S)

    End Sub

End Class
