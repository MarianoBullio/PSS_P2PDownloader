
Imports System.Windows.Forms
Imports System.Data
Imports OfficeOpenXml

'**************************************
' Module: Americas_Direct
' Created By: Mariano Bullio, bullio.m@pg.com
'**************************************

Public Class Americas_Direct

    Private MBS1 As New BindingSource

    Private SQL_F As New SQL_Functions
    Private SF As New System_Functions
    Private Variant_Code As String = "Direct"


    Public Function Initialize() As Boolean
        Initialize = False
        Fill_Matrix1()
        WebBrowser1.Navigate(HTMLPath & "P&G_Logo.html")
        WebBrowser1.ScrollBarsEnabled = False
        Initialize = True
    End Function

#Region "Functions Grid 1"

    Private Sub Fill_Matrix1()

        Try

            HSBarPos = SF.GetHSBarPos(dgv_Matrix1)
            VSBarPos = SF.GetVSBarPos(dgv_Matrix1)

            Dim C As DataGridViewColumn
            dgv_Matrix1.DataSource = Nothing
            dgv_Matrix1.Columns.Clear()
            MBS1.DataSource = SQL_F.GetDataTable(CS, "Select * From View_Americas_Direct")

            dgv_Matrix1.DataSource = MBS1
            BindingNavigator1.BindingSource = MBS1

            'Define Write Columns
            For Each C In dgv_Matrix1.Columns
                If C.Name <> "Flag" And C.Name <> "Owner" And C.Name <> "Enabled" And C.Name <> "Analyst" And C.Name <> "Pending_Approval" Then
                    C.ReadOnly = True
                End If
            Next

            'Styles
            dgv_Matrix1.EnableHeadersVisualStyles = False

            Dim Key_Column As New DataGridViewCellStyle
            Key_Column.BackColor = Color.MidnightBlue
            Key_Column.ForeColor = Color.White

            Dim Editable_Column As New DataGridViewCellStyle
            Editable_Column.BackColor = Color.SkyBlue
            Editable_Column.ForeColor = Color.Black

            'Key Columns
            dgv_Matrix1.Columns("PGrp").HeaderCell.Style = Key_Column
            dgv_Matrix1.Columns("SAP_Box").HeaderCell.Style = Key_Column
            dgv_Matrix1.Columns("POrg").HeaderCell.Style = Key_Column
            dgv_Matrix1.Columns("Conca_BoxPGrpPOrg").HeaderCell.Style = Key_Column

            'PickLists
            Dim I As Integer
            Dim CC As DataGridViewComboBoxColumn

            dgv_Matrix1.Columns("Flag").HeaderCell.Style = Editable_Column

            'Owner
            I = dgv_Matrix1.Columns("Owner").DisplayIndex
            dgv_Matrix1.Columns.Remove("Owner")
            CC = New DataGridViewComboBoxColumn
            CC.Name = "Owner"
            CC.DataPropertyName = "Owner"
            CC.HeaderText = "Owner"
            CC.DataSource = SQL_F.GetDataTable(CS, "SELECT TNumber, Name FROM Variant_Users Where (Employee = 1) AND (SPS =1) Order By Name Desc")
            CC.DisplayMember = "Name"
            CC.ValueMember = "TNumber"
            CC.DisplayIndex = I
            dgv_Matrix1.Columns.Add(CC)
            dgv_Matrix1.Columns("Owner").AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
            dgv_Matrix1.Columns("Owner").HeaderCell.Style = Editable_Column
            dgv_Matrix1.Columns("Owner").ReadOnly = True

            dgv_Matrix1.Columns("Enabled").HeaderCell.Style = Editable_Column

            'Analyst
            I = dgv_Matrix1.Columns("Analyst").DisplayIndex
            dgv_Matrix1.Columns.Remove("Analyst")
            CC = New DataGridViewComboBoxColumn
            CC.Name = "Analyst"
            CC.DataPropertyName = "Analyst"
            CC.HeaderText = "Analyst"
            CC.DataSource = SQL_F.GetDataTable(CS, "SELECT TNumber, Name FROM Variant_Users Where (Employee) = 0 Order By Name Desc")
            CC.DisplayMember = "Name"
            CC.ValueMember = "TNumber"
            CC.DisplayIndex = I
            dgv_Matrix1.Columns.Add(CC)
            dgv_Matrix1.Columns("Analyst").AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
            dgv_Matrix1.Columns("Analyst").HeaderCell.Style = Editable_Column

            Dim Change_Column As New DataGridViewCellStyle
            Change_Column.BackColor = Color.Orange
            Change_Column.ForeColor = Color.Black

            'Suggested_Owner
            I = dgv_Matrix1.Columns("Suggested_Owner").DisplayIndex
            dgv_Matrix1.Columns.Remove("Suggested_Owner")
            CC = New DataGridViewComboBoxColumn
            CC.Name = "Suggested_Owner"
            CC.DataPropertyName = "Suggested_Owner"
            CC.HeaderText = "Suggested_Owner"
            CC.DataSource = SQL_F.GetDataTable(CS, "SELECT TNumber, Name FROM Variant_Users Where (((Employee = 1) AND (SPS =1)) or (TNumber='TBD')) Order By Name Desc")
            CC.DisplayMember = "Name"
            CC.ValueMember = "TNumber"
            CC.DisplayIndex = I
            dgv_Matrix1.Columns.Add(CC)
            dgv_Matrix1.Columns("Suggested_Owner").AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
            dgv_Matrix1.Columns("Suggested_Owner").HeaderCell.Style = Change_Column

            dgv_Matrix1.Columns("Change_Type").HeaderCell.Style = Change_Column
            dgv_Matrix1.Columns("Email_Date").HeaderCell.Style = Change_Column
            dgv_Matrix1.Columns("Pending_Approval").HeaderCell.Style = Change_Column

            CheckUserPermission()

        Catch ex As Exception
            MsgBox("Error: " & ex.Message)
        End Try
    End Sub

    Private Sub CheckUserPermission()
        Dim Disable_Column As New DataGridViewCellStyle
        Disable_Column.BackColor = Color.Gray
        Disable_Column.ForeColor = Color.Black

        If Login_IsInfosysEmployee Then 'Infosys
            For Each C In dgv_Matrix1.Columns
                If C.Name <> "Flag" And C.Name <> "Infosys_ReqToPO" And C.Name <> "Infosys_OpenOrders" Then
                    C.ReadOnly = True
                End If
            Next
            btn_New.Enabled = False
            btn_Delete.Enabled = False
            btn_ModifySubVariants.Enabled = False
            dgv_Matrix1.Columns("Suggested_Owner").HeaderCell.Style = Disable_Column
            dgv_Matrix1.Columns("Change_Type").HeaderCell.Style = Disable_Column
            dgv_Matrix1.Columns("Email_Date").HeaderCell.Style = Disable_Column
            dgv_Matrix1.Columns("Pending_Approval").HeaderCell.Style = Disable_Column
            GroupBox1.Text = "Contractor Access"

            btn_MassiveUpdate.Enabled = False
            btn_MassiveRejection.Enabled = False
            btn_MassiveApproval.Enabled = False

        ElseIf Login_IsReadOnly Then 'ReadOnly
            For Each C In dgv_Matrix1.Columns
                C.ReadOnly = True
            Next
            btn_New.Enabled = False
            btn_Delete.Enabled = False
            ToolStripDropDownButton1.Enabled = False
            btn_ModifySubVariants.Enabled = False
            dgv_Matrix1.Columns("Suggested_Owner").HeaderCell.Style = Disable_Column
            dgv_Matrix1.Columns("Change_Type").HeaderCell.Style = Disable_Column
            dgv_Matrix1.Columns("Email_Date").HeaderCell.Style = Disable_Column
            dgv_Matrix1.Columns("Pending_Approval").HeaderCell.Style = Disable_Column
            GroupBox1.Text = "ReadOnly Access"

            btn_MassiveUpdate.Enabled = False
            btn_MassiveRejection.Enabled = False
            btn_MassiveApproval.Enabled = False

        ElseIf Login_IsVariantsSPOC Then 'VariantsSPOC
            If Not SF.ValidateAccessToVariant(Login, "Var_Direct") Then
                For Each C In dgv_Matrix1.Columns
                    C.ReadOnly = True
                Next
                btn_New.Enabled = False
                btn_Delete.Enabled = False
                ToolStripDropDownButton1.Enabled = False
                btn_ModifySubVariants.Enabled = False
                dgv_Matrix1.Columns("Suggested_Owner").HeaderCell.Style = Disable_Column
                dgv_Matrix1.Columns("Change_Type").HeaderCell.Style = Disable_Column
                dgv_Matrix1.Columns("Email_Date").HeaderCell.Style = Disable_Column
                dgv_Matrix1.Columns("Pending_Approval").HeaderCell.Style = Disable_Column
                GroupBox1.Text = "ReadOnly Access"

                btn_MassiveUpdate.Enabled = False
                btn_MassiveRejection.Enabled = False
                btn_MassiveApproval.Enabled = False
            Else
                GroupBox1.Text = "Variants SPOC Access"

                btn_MassiveUpdate.Enabled = True
                btn_MassiveRejection.Enabled = False
                btn_MassiveApproval.Enabled = False
            End If
        End If

    End Sub

    Private Sub dgv_Matrix1_RowLeave(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv_Matrix1.RowLeave

        Try
            If dgv_Matrix1.IsCurrentRowDirty Then

                Dim R As DataRow
                Dim PGrp As String = dgv_Matrix1.CurrentRow.Cells("PGrp").Value
                Dim SAP_Box As String = dgv_Matrix1.CurrentRow.Cells("SAP_Box").Value
                Dim POrg As String = dgv_Matrix1.CurrentRow.Cells("POrg").Value

                Dim DT As DataTable = SQL_F.GetSQLTable(CS, "Var_Americas_Direct", "(PGrp='" & PGrp & "') AND (SAP_Box='" & SAP_Box & "') AND (POrg=" & POrg & ")")
                If DT.Rows.Count = 1 Then
                    R = DT.Rows(0)
                    dgv_Matrix1.EndEdit()

                    Dim Enabled As String = R("Enabled")
                    Dim Suggested_Owner As String = R("Suggested_Owner")
                    Dim Pending_Approval As String = R("Pending_Approval")

                    R("Flag") = dgv_Matrix1.CurrentRow.Cells("Flag").Value
                    R("Owner") = dgv_Matrix1.CurrentRow.Cells("Owner").Value
                    R("Enabled") = dgv_Matrix1.CurrentRow.Cells("Enabled").Value
                    R("Analyst") = dgv_Matrix1.CurrentRow.Cells("Analyst").Value
                    R("Suggested_Owner") = dgv_Matrix1.CurrentRow.Cells("Suggested_Owner").Value
                    R("Pending_Approval") = dgv_Matrix1.CurrentRow.Cells("Pending_Approval").Value

                    'Set TRUE Enabled
                    If (Enabled = "False") And (R("Enabled") <> Enabled) Then
                        If (R("Change_Type") = 2) Then
                            MsgBox("This Combination is waiting for an approval from EPO, You can't Enable this Item")
                            R("Enabled") = Enabled
                            dgv_Matrix1.CurrentRow.Cells("Enabled").Value = "False"
                        Else
                            If SF.ShowConfirmationMessage("Set Enabled = True?", "") Then
                            Else
                                R("Enabled") = Enabled
                                dgv_Matrix1.CurrentRow.Cells("Enabled").Value = "False"
                            End If
                        End If

                    End If

                    'Set FALSE Enabled
                    If (Enabled = "True") And (R("Enabled") <> Enabled) Then
                        If SF.ShowConfirmationMessage("Set Enabled = False?", "") Then
                        Else
                            R("Enabled") = Enabled
                            dgv_Matrix1.CurrentRow.Cells("Enabled").Value = "True"
                        End If
                    End If

                    '*** Suggested Owner
                    If (R("Suggested_Owner") <> Suggested_Owner) And ((R("Change_Type") = 2) Or (R("Change_Type") = 3)) Then 'Si intentan cambiar el Owner en Status New Addition/Request for Deletion
                        MsgBox("This Combination is waiting for an approval from EPO, You can't suggest another Owner")
                        R("Suggested_Owner") = Suggested_Owner
                        dgv_Matrix1.CurrentRow.Cells("Suggested_Owner").Value = Suggested_Owner
                    Else

                        'Case 1 --->(TBD -> Suggestion)
                        If (Suggested_Owner = "TBD") And (R("Suggested_Owner") <> Suggested_Owner) Then

                            If Not DBNull.Value.Equals(R("Email_Date")) Then
                                MsgBox("This Combination is waiting for an approval from EPO, You can't suggest another Owner")
                                R("Suggested_Owner") = Suggested_Owner
                                dgv_Matrix1.CurrentRow.Cells("Suggested_Owner").Value = Suggested_Owner
                            Else

                                If SF.ShowConfirmationMessage("Would you like to Suggest a New Owner ?", Project_Name) Then

                                    R("Change_Type") = 1
                                    dgv_Matrix1.CurrentRow.Cells("Change_Type").Value = "Pending Owner Change"

                                Else
                                    R("Suggested_Owner") = Suggested_Owner
                                    dgv_Matrix1.CurrentRow.Cells("Suggested_Owner").Value = Suggested_Owner
                                End If

                            End If

                        End If

                        'Case 2 --->(Suggestion-> Suggestion)
                        If (Suggested_Owner <> "TBD") And (R("Suggested_Owner") <> Suggested_Owner) And (R("Suggested_Owner") <> "TBD") Then

                            If Not DBNull.Value.Equals(R("Email_Date")) Then
                                MsgBox("This Combination is waiting for an approval from EPO, You can't suggest another Owner")
                                R("Suggested_Owner") = Suggested_Owner
                                dgv_Matrix1.CurrentRow.Cells("Suggested_Owner").Value = Suggested_Owner
                            Else

                                If SF.ShowConfirmationMessage("Would you like to Suggest a New Owner ?", Project_Name) Then

                                    R("Change_Type") = 1
                                    dgv_Matrix1.CurrentRow.Cells("Change_Type").Value = "Pending Owner Change"

                                Else
                                    R("Suggested_Owner") = Suggested_Owner
                                    dgv_Matrix1.CurrentRow.Cells("Suggested_Owner").Value = Suggested_Owner
                                End If

                            End If

                        End If

                        'Case 3 --->(Suggestion-> TBD)
                        If (Suggested_Owner <> "TBD") And (R("Suggested_Owner") = "TBD") Then

                            If Not DBNull.Value.Equals(R("Email_Date")) Then
                                MsgBox("This Combination is waiting for an approval from EPO, You can't suggest another Owner")
                                R("Suggested_Owner") = Suggested_Owner
                                dgv_Matrix1.CurrentRow.Cells("Suggested_Owner").Value = Suggested_Owner
                            Else

                                If SF.ShowConfirmationMessage("Would you like to Delete the suggested Owner ?", Project_Name) Then

                                    R("Change_Type") = 0
                                    dgv_Matrix1.CurrentRow.Cells("Change_Type").Value = "No pending Changes"

                                Else
                                    R("Suggested_Owner") = Suggested_Owner
                                    dgv_Matrix1.CurrentRow.Cells("Suggested_Owner").Value = Suggested_Owner
                                End If

                            End If

                        End If

                    End If

                    '*** Pending_Approval
                    'Set TRUE
                    If (Pending_Approval = "False") And (R("Pending_Approval") <> Pending_Approval) Then
                        MsgBox("You can't Check this Column", MsgBoxStyle.Information)
                        R("Pending_Approval") = Pending_Approval
                        dgv_Matrix1.CurrentRow.Cells("Pending_Approval").Value = "False"
                    End If

                    'Set FALSE
                    If (Pending_Approval = "True") And (R("Pending_Approval") <> Pending_Approval) Then
                        If SF.ShowConfirmationMessage("Would you like to Approve this Change? " & vbNewLine & vbNewLine & "Change Type: " & dgv_Matrix1.CurrentRow.Cells("Change_Type").Value, "") Then

                            Select Case R("Change_Type")

                                Case 1 'Pending Owner Change
                                    R("Owner") = R("Suggested_Owner")
                                    dgv_Matrix1.CurrentRow.Cells("Owner").Value = R("Suggested_Owner")

                                    R("Suggested_Owner") = "TBD"
                                    dgv_Matrix1.CurrentRow.Cells("Suggested_Owner").Value = "TBD"

                                    R("Change_Type") = 0
                                    dgv_Matrix1.CurrentRow.Cells("Change_Type").Value = "No pending Changes"

                                    Delete_Email_Date(R)
                                    'dgv_Matrix1.CurrentRow.Cells("Email_Date").Value = System.DBNull.Value

                                    'Save Historical Data


                                Case 2 'New Addition
                                    R("Owner") = R("Suggested_Owner")
                                    dgv_Matrix1.CurrentRow.Cells("Owner").Value = R("Suggested_Owner")

                                    R("Suggested_Owner") = "TBD"
                                    dgv_Matrix1.CurrentRow.Cells("Suggested_Owner").Value = "TBD"

                                    R("Change_Type") = 0
                                    dgv_Matrix1.CurrentRow.Cells("Change_Type").Value = "No pending Changes"

                                    Delete_Email_Date(R)
                                    'dgv_Matrix1.CurrentRow.Cells("Email_Date").Value = System.DBNull.Value

                                    R("Enabled") = "True"
                                    dgv_Matrix1.CurrentRow.Cells("Enabled").Value = "True"

                                    'Save Historical Data

                                Case 3 'Request for Deletion

                                    If SF.ShowConfirmationMessage("This Combination will be Permanently Deleted", Project_Name) Then
                                        Delete_Combination(R)
                                        dgv_Matrix1.CurrentRow.DefaultCellStyle.BackColor = Color.Red
                                        'Save Historical Data

                                    Else
                                        R("Pending_Approval") = Pending_Approval
                                        dgv_Matrix1.CurrentRow.Cells("Pending_Approval").Value = "True"

                                    End If
                            End Select

                        Else
                            R("Pending_Approval") = Pending_Approval
                            dgv_Matrix1.CurrentRow.Cells("Pending_Approval").Value = "True"
                        End If
                    End If


                    Dim Update As String = "Update Var_Americas_Direct Set " & _
                    "Flag=" & SF.BooleanConvertion(R("Flag")) & ", " & _
                    "Owner='" & R("Owner") & "'," & _
                    "Enabled='" & R("Enabled") & "'," & _
                    "Analyst='" & R("Analyst") & "'," & _
                    "Updated_Date=GetDate()," & _
                    "Updated_By='" & Login & "'," & _
                    "Suggested_Owner='" & R("Suggested_Owner").ToString.Trim & "', " & _
                    "Change_Type=" & R("Change_Type") & "," & _
                    "Pending_Approval=" & SF.BooleanConvertion(R("Pending_Approval")) & " " & _
                    "Where (PGrp='" & PGrp & "') AND (SAP_Box='" & SAP_Box & "') AND (POrg=" & POrg & ")"

                    Dim Exec As String = SQL_F.SQL_Execute_NQ(CS, Update)
                    If Exec <> Nothing Then
                        MsgBox("Error: " & Exec, MsgBoxStyle.Exclamation, Project_Name)
                    Else
                        dgv_Matrix1.CurrentRow.Cells("Updated_Date").Value = SF.GetDateString
                        dgv_Matrix1.CurrentRow.Cells("Updated_By").Value = Login
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, MsgBoxStyle.Critical, "Update Error")
        End Try
    End Sub

    Private Sub Delete_Email_Date(ByRef R As DataRow)
        Try
            Dim Update As String = "Update Var_Americas_Direct Set Email_Date=Null Where (PGrp='" & R("PGrp") & "') AND (SAP_Box='" & R("SAP_Box") & "') AND (POrg=" & R("POrg") & ")"
            Dim Exc As String = SQL_F.SQL_Execute_NQ(CS, Update)
            If Exc <> Nothing Then
                'MsgBox("Error")
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub Delete_Combination(ByRef R As DataRow)
        Try
            Dim Delete As String = "Delete From Var_Americas_Direct Where (PGrp='" & R("PGrp") & "') AND (SAP_Box='" & R("SAP_Box") & "') AND (POrg=" & R("POrg") & ")"
            Dim Exc As String = SQL_F.SQL_Execute_NQ(CS, Delete)
            If Exc <> Nothing Then
                'MsgBox("Error")
            Else
                MsgBox("Combination has been Deleted.", MsgBoxStyle.Information, Project_Name)
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub dgv_Matrix1_CellMouseDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgv_Matrix1.CellMouseDoubleClick
        'Try

        '    If Not dgv_Matrix1.CurrentRow.IsNewRow Then
        '        Dim FRM_Small As New ViewCommentsLE
        '        FRM_Small.Comment = ""
        '        Row_Global = dgv_Matrix1.CurrentRow
        '        If (FRM_Small.ShowDialog() = Windows.Forms.DialogResult.OK) Then
        '            dgv_Matrix1.CurrentRow.Cells("Comments").Value = FRM_Small.Comment
        '        End If
        '    End If

        'Catch
        'End Try
    End Sub

    Private Sub dgv_Matrix1_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dgv_Matrix1.KeyDown

        'Ctr + F 
        If e.KeyCode = Keys.F AndAlso e.Control Then
            Dim D As New Shared_Functions.BS_Find
            D.Initialize(dgv_Matrix1)
            D.Show()
        End If

        'Delete
        If e.KeyCode = Keys.Delete Then
            If Not dgv_Matrix1.CurrentCell Is Nothing Then
                dgv_Matrix1.CurrentCell.Value = DBNull.Value
            End If
        End If

        If dgv_Matrix1.AreAllCellsSelected(True) Or ((Not dgv_Matrix1.CurrentRow Is Nothing) AndAlso dgv_Matrix1.CurrentRow.Selected) Then
            dgv_Matrix1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText
        Else
            dgv_Matrix1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText
        End If

    End Sub

    Private Sub dgv_Matrix1_PaintRows() Handles dgv_Matrix1.RowPostPaint
        For Each row As DataGridViewRow In dgv_Matrix1.Rows
            If Not row.IsNewRow Then

                Try
                    Select Case row.Cells("Change_Type").Value.ToString
                        Case "Pending Owner Change"
                            row.DefaultCellStyle.BackColor = Color.Orange
                        Case "New Addition"
                            row.DefaultCellStyle.BackColor = Color.LightGreen
                        Case "Request for Deletion"
                            row.DefaultCellStyle.BackColor = Color.LightSalmon
                    End Select
                Catch ex As Exception
                End Try

            End If
        Next
    End Sub

    Private Sub dgv_Matrix1_CellClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv_Matrix1.CellClick
        dgv_Matrix1.BeginEdit(True)
    End Sub

#End Region

#Region "Functions Tool Bar"

    Private Sub btn_Refresh(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        Fill_Matrix1()
    End Sub

    Private Sub btn_Filter_BySelection_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Filter_BySelection.Click

        Select Case TabControl.SelectedIndex
            Case 0

                Try
                    Dim FV As String
                    If Not DBNull.Value.Equals(dgv_Matrix1.CurrentCell.Value) Then
                        If dgv_Matrix1.CurrentCell.ValueType.Equals(GetType(Date)) Then
                            FV = " >= '" & String.Format("{0:yyyy-MM-dd}", dgv_Matrix1.CurrentCell.Value) & " 00:00:00.000' AND " & dgv_Matrix1.CurrentCell.OwningColumn.Name & " <= '" & String.Format("{0:yyyy-MM-dd}", dgv_Matrix1.CurrentCell.Value) & " 23:59:00.000'"
                        Else
                            FV = " = '" & dgv_Matrix1.CurrentCell.Value & "'"
                        End If
                    Else
                        FV = " Is Null"
                    End If
                    Dim FE As String = dgv_Matrix1.CurrentCell.OwningColumn.Name & FV
                    If MBS1.Filter <> Nothing Then
                        MBS1.Filter = MBS1.Filter & " AND " & FE
                    Else
                        MBS1.Filter = FE
                    End If
                    btn_Filter_Clear.CheckOnClick = True
                    btn_Filter_Clear.Checked = True
                Catch
                End Try

        End Select

    End Sub

    Private Sub btn_Filter_ExcSelection_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Filter_ExcSelection.Click
        Select Case TabControl.SelectedIndex

            Case 0

                Try
                    Dim FV As String
                    If Not DBNull.Value.Equals(dgv_Matrix1.CurrentCell.Value) Then
                        If dgv_Matrix1.CurrentCell.ValueType.Equals(GetType(Date)) Then
                            FV = " >= '" & String.Format("{0:yyyy-MM-dd}", dgv_Matrix1.CurrentCell.Value) & " 00:00:00.000' AND " & dgv_Matrix1.CurrentCell.OwningColumn.Name & " <= '" & String.Format("{0:yyyy-MM-dd}", dgv_Matrix1.CurrentCell.Value) & " 23:59:00.000'"
                        Else
                            FV = " = '" & dgv_Matrix1.CurrentCell.Value & "'"
                        End If
                    Else
                        FV = " Is Null"
                    End If
                    Dim FE As String = "Not (" & dgv_Matrix1.CurrentCell.OwningColumn.Name & FV & ")"
                    If MBS1.Filter <> Nothing Then
                        MBS1.Filter = MBS1.Filter & " AND " & FE
                    Else
                        MBS1.Filter = FE
                    End If
                    btn_Filter_Clear.CheckOnClick = True
                    btn_Filter_Clear.Checked = True
                Catch ex As Exception
                    MsgBox("Error: " & ex.Message)
                End Try

        End Select
    End Sub

    Private Sub btn_Filter_Clear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Filter_Clear.Click

        btn_Filter_Clear.CheckOnClick = False
        MBS1.Filter = Nothing

    End Sub

    Private Sub btn_FlagAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Flag_All.Click
        FlagCheck("Flag=", 1)
    End Sub

    Private Sub btn_FlagNone_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Flag_None.Click
        FlagCheck("Flag=", 0)
    End Sub

    Private Sub FlagCheck(ByVal Column As String, ByVal Value As Integer)
        Dim View As String = ""
        Dim Filter As String = ""

        Select Case TabControl.SelectedIndex
            Case 0
                View = "View_Americas_Direct"
                Filter = MBS1.Filter
        End Select

        Try
            Dim R As DataRow

            Dim PGrp As String = Nothing
            Dim SAP_Box As String = Nothing
            Dim POrg As String = Nothing
            Dim Selec As String


            If Filter <> "Not (Flag = 'True')" And Filter <> Nothing Then
                Selec = "Select * from " & View & " where " & Filter '(Plant='" & Plant & "') AND (SAP_Box='" & SAP_Box & "') AND (POrg=" & POrg & ") AND " & Filter
            Else
                Selec = "Select * from " & View
            End If

            Dim DT As DataTable = SQL_F.GetDataTable(CS, Selec)
            For i = 0 To DT.Rows.Count - 1

                R = DT.Rows(i)
                PGrp = R("PGrp")
                SAP_Box = R("SAP_Box")
                POrg = R("POrg")

                Dim Update As String = "Update Var_Americas_Direct Set " & _
                Column & Value & " " & _
                "Where (PGrp='" & PGrp & "') AND (SAP_Box='" & SAP_Box & "') AND (POrg=" & POrg & ")"

                Dim Exec As String = SQL_F.SQL_Execute_NQ(CS, Update)
            Next

            Select Case TabControl.SelectedIndex
                Case 0
                    Fill_Matrix1()
            End Select

        Catch ex As Exception
            MsgBox("Error to Flag = True/False", MsgBoxStyle.Critical, "Error!")
        End Try

    End Sub

#End Region

#Region "Buttons"

    Private Sub btn_New_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_New.Click

        Dim FRM_Small As New Americas_Direct_NewRow
        Dim PGrp As String = ""
        Dim SAP_Box As String = ""
        Dim POrg As String = ""
        Dim Owner As String = ""

        Try
            While Not FRM_Small.Result
                If (FRM_Small.ShowDialog() = Windows.Forms.DialogResult.OK) Then
                    If FRM_Small.Result Then

                        PGrp = FRM_Small.PGrp_
                        SAP_Box = FRM_Small.SAP_Box_
                        POrg = FRM_Small.POrg_
                        Owner = FRM_Small.Owner_

                    End If
                End If

                If Not FRM_Small.Result Then 'If Cancel
                    Exit Sub
                End If

                Dim Insert As String = "Insert Into Var_Americas_Direct(PGrp,SAP_Box,POrg,Suggested_Owner,Change_Type,Enabled,Owner,Analyst,Created_Date,Created_By,Updated_Date,Updated_By) Values"
                Insert += "("
                Insert += "'" & PGrp.Trim & "',"
                Insert += "'" & SAP_Box.Trim & "',"
                Insert += POrg.Trim & ","
                Insert += "'" & Owner.Trim & "'," 'Suggested_Owner
                Insert += 2 & "," 'Change_Type
                Insert += 0 & "," 'Enabled
                Insert += "'" & Owner.Trim & "',"
                Insert += "'TBD',"
                Insert += "GetDate(),"
                Insert += "'" & Login.Trim & "',"
                Insert += "GetDate(),"
                Insert += "'" & Login.Trim & "'"
                Insert += ")"

                Dim Exc As String = SQL_F.SQL_Execute_NQ(CS, Insert)
                If Exc <> Nothing Then
                    If Exc.Contains("Violation") Then
                        MsgBox("This Combination Already exists [PGrp,SAP_Box,POrg]", MsgBoxStyle.Information, Project_Name)
                    Else
                        MsgBox(Standart_MessageError, MsgBoxStyle.Exclamation, Project_Name)
                    End If
                Else
                    Fill_Matrix1()
                    MsgBox("New Row has been Created:  It will be enabled as soon as the EPO Owner approves the New Addition", MsgBoxStyle.Information, Project_Name)
                End If

            End While

        Catch ex As Exception
            MsgBox(Standart_MessageError, MsgBoxStyle.Exclamation, Project_Name)
        End Try

    End Sub

    Private Sub btn_Delete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Delete.Click
        Try
            Dim DT_ToBeDeleted As New DataTable
            DT_ToBeDeleted = SQL_F.GetDataTable(CS, "Select * From View_Americas_Direct Where (Flag=1)")
            Dim Count As Integer = 0

            If DT_ToBeDeleted.Rows.Count > 0 Then
                If SF.ShowConfirmationMessage("Would you like to Delete: [" & DT_ToBeDeleted.Rows.Count & "] Combinations? ", Project_Name) Then
                    For Each R As DataRow In DT_ToBeDeleted.Rows
                        Dim Delete As String = "Update Var_Americas_Direct Set Change_Type=3 Where (PGrp='" & R("PGrp") & "') AND (SAP_Box='" & R("SAP_Box") & "') AND (POrg=" & R("POrg") & ")"
                        Dim Exc As String = SQL_F.SQL_Execute_NQ(CS, Delete)
                        If Exc <> Nothing Then
                            MsgBox("Error: " & Exc, MsgBoxStyle.Exclamation, Project_Name)
                        Else
                            Count += 1
                        End If
                    Next
                    Fill_Matrix1()
                    MsgBox("Done: [" & Count & "] Combinations have been set as 'Request for Deletion'. Please wait for EPO Owner's Approval", MsgBoxStyle.Information, Project_Name)
                End If
            Else
                MsgBox("No Combinations Selected", MsgBoxStyle.Exclamation, Project_Name)
            End If

        Catch ex As Exception
            MsgBox(Standart_MessageError, MsgBoxStyle.Exclamation)
        End Try
    End Sub

    Private Sub btn_ModifySubVariants_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_ModifySubVariants.Click
        Americas_Direct_SubVariants.Show()
    End Sub

    Private Sub btn_Export_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Export.Click
        If SF.ExportVariantToExcel("Direct", My.Computer.FileSystem.SpecialDirectories.Desktop & "\Variant - Americas Direct (" & Date.Now.Month & "-" & Date.Now.Day & "-" & Date.Now.Year & ").xlsx") Then
            MsgBox("Done: The Report has been created on your desk", MsgBoxStyle.Information)
        Else
            MsgBox(Standart_MessageError, MsgBoxStyle.Exclamation)
        End If

    End Sub

    Private Sub btn_MassiveUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_MassiveUpdate.Click
        Dim CountOfCombinations As Integer = SF.MassiveUpdate(Me.Variant_Code)
        If CountOfCombinations > 0 Then
            Fill_Matrix1()
            MsgBox("Done: [" & CountOfCombinations & "] Combinations have been Modified", MsgBoxStyle.Information)
        End If
    End Sub

    Private Sub btn_MassiveRejection_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_MassiveRejection.Click
        Dim CountOfCombinations As Integer = SF.MassiveRejection(Me.Variant_Code)
        If CountOfCombinations > 0 Then
            Fill_Matrix1()
            MsgBox("Done: [" & CountOfCombinations & "] Combinations have been Rejected", MsgBoxStyle.Information)
        End If
    End Sub

    Private Sub btn_MassiveApproval_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_MassiveApproval.Click
        Dim CountOfCombinations As Integer = SF.MassiveApproval(Me.Variant_Code)
        If CountOfCombinations > 0 Then
            Fill_Matrix1()
            MsgBox("Done: [" & CountOfCombinations & "] Combinations have been Approved", MsgBoxStyle.Information)
        End If
    End Sub
#End Region





End Class
