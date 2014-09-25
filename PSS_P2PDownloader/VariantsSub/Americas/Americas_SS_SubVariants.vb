Imports Shared_Functions
Imports System.IO
Imports System.Globalization
Imports System.Threading

Public Class Americas_SS_SubVariants

    Private SQL_F As New SQL_Functions
    Private SF As New System_Functions

    Private MBS As New BindingSource

    Private Sub Americas_SS_SubVariants_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Try
            SF.Load_ComboBox(List_Variant, CS, "SELECT Variant, Description FROM Var_Americas_SS_Variants Order By Description Asc")
            Fill_DGV(List_Variant.SelectedValue)
        Catch ex As Exception

        End Try


    End Sub

    Private Sub Fill_DGV(ByVal Variant_Number As Object)
        Try

            If Variant_Number < 10 Then
                Variant_Number = "0" & Variant_Number
            Else
                Variant_Number = Variant_Number
            End If

            Dim C As DataGridViewColumn
            dgv_Main.DataSource = Nothing
            dgv_Main.Columns.Clear()

            If Variant_Number <> "01" Then
                MBS.DataSource = SQL_F.GetDataTable(CS, "Select * From Var_Americas_SS_Var" & Variant_Number)

                dgv_Main.DataSource = MBS
                BindingNavigator1.BindingSource = MBS
                Dim dts As DataTable = MBS.DataSource

                'Definimos los campos solo lectura
                For Each C In dgv_Main.Columns
                    C.ReadOnly = True
                Next
            End If

          
        Catch ex As Exception
            'MsgBox("Error: " & ex.Message)
        End Try
    End Sub

    Private Sub dgv_Main_RowLeave(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgv_Main.RowLeave

    End Sub

    Private Sub dgv_Matrix1_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dgv_Main.KeyDown

        'Ctr + F 
        If e.KeyCode = Keys.F AndAlso e.Control Then
            Dim D As New Shared_Functions.BS_Find
            D.Initialize(dgv_Main)
            D.Show()
        End If

        'Delete
        If e.KeyCode = Keys.Delete Then
            If Not dgv_Main.CurrentCell Is Nothing Then
                dgv_Main.CurrentCell.Value = DBNull.Value
            End If
        End If

        If dgv_Main.AreAllCellsSelected(True) Or ((Not dgv_Main.CurrentRow Is Nothing) AndAlso dgv_Main.CurrentRow.Selected) Then
            dgv_Main.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText
        Else
            dgv_Main.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText
        End If

    End Sub

    Private Sub List_Variant_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles List_Variant.SelectedIndexChanged
        Try
            Fill_DGV(List_Variant.SelectedValue)
        Catch ex As Exception

        End Try

    End Sub

#Region "Functions Tool Bar"

    Private Sub btn_Filter_BySelection_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Filter_BySelection.Click
        Try


            Dim FV As String
            If Not DBNull.Value.Equals(dgv_Main.CurrentCell.Value) Then
                If dgv_Main.CurrentCell.ValueType.Equals(GetType(Date)) Then
                    FV = " >= '" & String.Format("{0:yyyy-MM-dd}", dgv_Main.CurrentCell.Value) & " 00:00:00.000' AND " & dgv_Main.CurrentCell.OwningColumn.Name & " <= '" & String.Format("{0:yyyy-MM-dd}", dgv_Main.CurrentCell.Value) & " 23:59:00.000'"
                Else
                    FV = " = '" & dgv_Main.CurrentCell.Value & "'"
                End If
            Else
                FV = " Is Null"
            End If
            Dim FE As String = dgv_Main.CurrentCell.OwningColumn.Name & FV
            If MBS.Filter <> Nothing Then
                MBS.Filter = MBS.Filter & " AND " & FE
            Else
                MBS.Filter = FE
            End If
            btn_Filter_Clear.CheckOnClick = True
            btn_Filter_Clear.Checked = True
        Catch ex As Exception

        End Try
    End Sub

    Private Sub btn_Filter_ExcSelection_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Filter_ExcSelection.Click
        Try
            Dim FV As String
            If Not DBNull.Value.Equals(dgv_Main.CurrentCell.Value) Then
                If dgv_Main.CurrentCell.ValueType.Equals(GetType(Date)) Then
                    FV = " >= '" & String.Format("{0:yyyy-MM-dd}", dgv_Main.CurrentCell.Value) & " 00:00:00.000' AND " & dgv_Main.CurrentCell.OwningColumn.Name & " <= '" & String.Format("{0:yyyy-MM-dd}", dgv_Main.CurrentCell.Value) & " 23:59:00.000'"
                Else
                    FV = " = '" & dgv_Main.CurrentCell.Value & "'"
                End If
            Else
                FV = " Is Null"
            End If
            Dim FE As String = "Not (" & dgv_Main.CurrentCell.OwningColumn.Name & FV & ")"
            If MBS.Filter <> Nothing Then
                MBS.Filter = MBS.Filter & " AND " & FE
            Else
                MBS.Filter = FE
            End If
            btn_Filter_Clear.CheckOnClick = True
            btn_Filter_Clear.Checked = True
        Catch ex As Exception
            MsgBox("Error: " & ex.Message)
        End Try
    End Sub

    Private Sub btn_Filter_Clear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Filter_Clear.Click
        Try


            btn_Filter_Clear.CheckOnClick = False
            MBS.Filter = Nothing

        Catch
        End Try
    End Sub

    Private Sub btn_Refresh(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        Fill_DGV(List_Variant.SelectedValue)
    End Sub

    Private Sub btn_Add_Rows_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Add_Rows.Click
        Try

            Dim Variant_Number As Object
            Variant_Number = List_Variant.SelectedValue

            If Variant_Number < 10 Then
                Variant_Number = "0" & Variant_Number
            Else
                Variant_Number = Variant_Number
            End If

            Select Case Variant_Number

                Case "01"
                    MsgBox("This List does not contain Items", MsgBoxStyle.Information, Project_Name)
                Case "02", "03", "04", "05", "06", "07", "10", "11", "12", "13", "16"

                    Dim Material_Group As String
                    Material_Group = InputBox("Material Group [Required Length 9]:", Project_Name, "")

                    If Material_Group = "" Then
                        Exit Sub
                    End If

                    If Material_Group.Length <> 9 Then
                        MsgBox("Error: Review 'Material Group' Value [Length must be = 9]", MsgBoxStyle.Exclamation, Project_Name)
                    Else
                        Dim Insert As String = "Insert Into Var_Americas_SS_Var" & Variant_Number & " Values ('" & Material_Group & "')"
                        Dim Exc As String = SQL_F.SQL_Execute_NQ(CS, Insert)
                        If Exc <> Nothing Then
                            If Exc.Contains("Violation") Then
                                MsgBox("This Combination Already exists [Material Group]", MsgBoxStyle.Information, Project_Name)
                            Else
                                MsgBox(Standart_MessageError, MsgBoxStyle.Exclamation, Project_Name)
                            End If
                        Else
                            Fill_DGV(List_Variant.SelectedValue)
                            MsgBox("New Item has been Created", MsgBoxStyle.Information, Project_Name)
                        End If
                    End If

                Case "08", "09", "14", "15"

                    Dim PGrp As String
                    PGrp = InputBox("PGrp [Required Lengt 3]:", Project_Name, "")

                    If PGrp = "" Then
                        Exit Sub
                    End If

                    If PGrp.Length <> 3 Then
                        MsgBox("Error: Review 'PGrp' Value [Length must be = 3]", MsgBoxStyle.Exclamation, Project_Name)
                    Else
                        Dim Insert As String = "Insert Into Var_Americas_SS_Var" & Variant_Number & " Values ('" & PGrp & "')"
                        Dim Exc As String = SQL_F.SQL_Execute_NQ(CS, Insert)
                        If Exc <> Nothing Then
                            If Exc.Contains("Violation") Then
                                MsgBox("This Combination Already exists [PGrp]", MsgBoxStyle.Information, Project_Name)
                            Else
                                MsgBox(Standart_MessageError, MsgBoxStyle.Exclamation, Project_Name)
                            End If
                        Else
                            Fill_DGV(List_Variant.SelectedValue)
                            MsgBox("New Item has been Created", MsgBoxStyle.Information, Project_Name)
                        End If
                    End If

                Case "17"

                    Dim Material_Group As String
                    Material_Group = InputBox("Material Group [Required Length 9]:", Project_Name, "")

                    If Material_Group = "" Then
                        Exit Sub
                    End If

                    Dim ServiceLine As String
                    ServiceLine = InputBox("ServiceLine [Customization/Logistics]:", Project_Name, "")

                    If ServiceLine = "" Then
                        Exit Sub
                    End If

                    If (Material_Group.Length <> 9) Then
                        MsgBox("Error: Review 'Material Group' Value [Length must be = 9]", MsgBoxStyle.Exclamation, Project_Name)
                    Else
                        Dim Insert As String = "Insert Into Var_Americas_SS_Var" & Variant_Number & " Values ('" & Material_Group & "','" & ServiceLine & "')"
                        Dim Exc As String = SQL_F.SQL_Execute_NQ(CS, Insert)
                        If Exc <> Nothing Then
                            If Exc.Contains("Violation") Then
                                MsgBox("This Combination Already exists [Material Group]", MsgBoxStyle.Information, Project_Name)
                            Else
                                MsgBox(Standart_MessageError, MsgBoxStyle.Exclamation, Project_Name)
                            End If
                        Else
                            Fill_DGV(List_Variant.SelectedValue)
                            MsgBox("New Item has been Created", MsgBoxStyle.Information, Project_Name)
                        End If
                    End If
            End Select

        Catch ex As Exception

        End Try
    End Sub

    Private Sub btn_Delete_Row_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Delete_Row.Click
        Try
            If SF.ShowConfirmationMessage() Then

                Dim Variant_Number As Object
                Variant_Number = List_Variant.SelectedValue

                If Variant_Number < 10 Then
                    Variant_Number = "0" & Variant_Number
                Else
                    Variant_Number = Variant_Number
                End If

                Select Case Variant_Number

                    Case "01"
                        MsgBox("This List does not contain Items", MsgBoxStyle.Information, Project_Name)
                    Case "02", "03", "04", "05", "06", "07", "10", "11", "12", "13", "16", "17"

                        Dim Material_Group As String
                        Material_Group = dgv_Main.CurrentRow.Cells("Material_Group").Value

                        Dim Delete As String = "Delete From Var_Americas_SS_Var" & Variant_Number & " Where(Material_Group='" & Material_Group & "')"
                        Dim Exc As String = SQL_F.SQL_Execute_NQ(CS, Delete)
                        If Exc <> Nothing Then
                            MsgBox(Standart_MessageError, MsgBoxStyle.Exclamation, Project_Name)
                        Else
                            Fill_DGV(List_Variant.SelectedValue)
                            MsgBox("Item has been Deleted [" & Material_Group & "]", MsgBoxStyle.Information, Project_Name)
                        End If


                    Case "08", "09", "14", "15"

                        Dim PGrp As String
                        PGrp = dgv_Main.CurrentRow.Cells("PGrp").Value

                        Dim Delete As String = "Delete From Var_Americas_SS_Var" & Variant_Number & " Where(PGrp='" & PGrp & "')"
                        Dim Exc As String = SQL_F.SQL_Execute_NQ(CS, Delete)
                        If Exc <> Nothing Then
                            MsgBox(Standart_MessageError, MsgBoxStyle.Exclamation, Project_Name)
                        Else
                            Fill_DGV(List_Variant.SelectedValue)
                            MsgBox("Item has been Deleted [" & PGrp & "]", MsgBoxStyle.Information, Project_Name)
                        End If
                End Select
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub btn_Copy_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Copy.Click
        Try

            Dim Variant_Number As Object
            Variant_Number = List_Variant.SelectedValue

            If Variant_Number < 10 Then
                Variant_Number = "0" & Variant_Number
            Else
                Variant_Number = Variant_Number
            End If

            If Variant_Number <> "01" Then
                SQL_F.TableToClipboard(SQL_F.GetDataTable(CS, "Select * From Var_Americas_SS_Var" & Variant_Number))
            End If

        Catch ex As Exception

        End Try
    End Sub

#End Region






   

 
End Class