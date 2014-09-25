'**************************************
' Module: Popup New Row
' Created By: Mariano Bullio, bullio.m@pg.com
'**************************************

Public Class Massive_OwnerChange

    Private SQL_F As New SQL_Functions
    Private SF As New System_Functions

    Friend Result As Boolean = False
    Friend Owner_ As String = ""
    Friend CountOfCombinations_ As Integer = 0
    Friend Variant_Code As String = ""

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try
            SF.Load_ListBox(List_Owner, CS, "SELECT TNumber, Name FROM Variant_Users Where ((Active = 1) AND (Employee = 1) AND (SPS =1)) OR (TNumber ='TBD') Order By Name Asc")
        Catch ex As Exception
            MessageBox.Show(Standart_MessageError, Project_Name, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Me.Close()
            Me.Result = False
        End Try

    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            Me.Close()
            Me.Result = False
        Catch ex As Exception
            MessageBox.Show(Me.Name + " - btnCancel_Click: " + ex.Message, Project_Name, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOk.Click
        Try
            Dim DT As New DataTable
            DT = SQL_F.GetSQLTable(CS, "Var_Americas_" & Variant_Code, " (Flag = 1) AND (Change_Type IN (0,1)) AND (Email_Date Is NULL)")
            Me.CountOfCombinations_ = DT.Rows.Count

            If Me.CountOfCombinations_ > 0 Then
                If SF.ShowConfirmationMessage("Would you like to Suggest a New Owner for [" & DT.Rows.Count & "] Combinations ?", "Massive Owner Change") Then

                    For Each R As DataRow In DT.Rows

                        Dim Update As String = ""
                        If Not List_Owner.SelectedValue = "TBD" Then
                            Update = "Update Var_Americas_" & Variant_Code & " Set Suggested_Owner='" & List_Owner.SelectedValue & "',Change_Type=1 " & SF.CreateWhereCondition(R, Variant_Code)
                        Else
                            Update = "Update Var_Americas_" & Variant_Code & " Set Suggested_Owner='" & List_Owner.SelectedValue & "',Change_Type=0 " & SF.CreateWhereCondition(R, Variant_Code)
                        End If
                        Dim Exc As String = SQL_F.SQL_Execute_NQ(CS, Update)
                        If Exc <> Nothing Then
                            MsgBox("Error:" & Standart_MessageError, Project_Name)
                        End If

                    Next

                    Result = Send_Info()
                    If Result Then
                        Me.DialogResult = Windows.Forms.DialogResult.OK
                        Me.Close()
                    Else
                        Exit Sub
                    End If
                End If
            Else
                MsgBox("No Combinations Selected", MsgBoxStyle.Exclamation, Project_Name)
            End If

           
        Catch ex As Exception
        End Try
    End Sub

    Private Function Send_Info() As Boolean

        Try
            Send_Info = True
            Me.Owner_ = List_Owner.SelectedValue
        Catch ex As Exception
            Send_Info = False
        End Try
        

    End Function



End Class