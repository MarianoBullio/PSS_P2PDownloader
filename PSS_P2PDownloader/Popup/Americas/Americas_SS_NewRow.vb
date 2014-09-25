'**************************************
' Module: Popup New Row
' Created By: Mariano Bullio, bullio.m@pg.com
'**************************************

Public Class Americas_SS_NewRow

    Private SQL_F As New SQL_Functions
    Private SF As New System_Functions

    Friend Result As Boolean = False

    Friend Region_ As String = ""
    Friend Plant_ As String = ""
    Friend SAP_Box_ As String = ""
    Friend POrg_ As String = ""
    Friend Variant_ As String = ""
    Friend Variant_To_Apply_ As String = ""
    Friend Owner_ As String = ""

    Private ServiceLine As String = "SS"

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try
            SF.Load_ComboBox(List_Region, CS, "SELECT ID, Description FROM Regions Order By Description Desc")
            SF.Load_ComboBox(List_SAP_Box, CS, "SELECT SAP_Box, Description FROM SAP_Boxes Where (" & ServiceLine & " = 1) Order By Description Desc")
            SF.Load_ListBox(List_Owner, CS, "SELECT TNumber, Name FROM Variant_Users Where (Active = 1) AND (Employee = 1) AND (SPS =1) Order By Name Asc")
            SF.Load_ComboBox(List_Variant, CS, "SELECT Variant, Description FROM Var_Americas_SS_Variants Order By Description Asc")
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
            If SF.ShowConfirmationMessage() Then

                If List_Region.Text.Length <> 2 Then
                    MsgBox("Error: Review 'Region' Value", MsgBoxStyle.Exclamation, Project_Name)
                    Exit Sub
                End If

                If txt_Plant.Text.Trim.Length <> 4 Then
                    MsgBox("Error: Review 'Plant' Value [Length must be = 4]", MsgBoxStyle.Exclamation, Project_Name)
                    Exit Sub
                Else
                    If Not SF.ValidatePlant(txt_Plant.Text.Trim) Then
                        MsgBox("Error: 'Plant' Code does not exist", MsgBoxStyle.Exclamation, Project_Name)
                        Exit Sub
                    End If
                End If

                If Not IsNumeric(txt_POrg.Text.Trim) Then
                    MsgBox("Error: Review 'POrg' Value [It must be Numeric]", MsgBoxStyle.Exclamation, Project_Name)
                    Exit Sub
                Else
                    If Not SF.ValidatePOrg(txt_POrg.Text.Trim, ServiceLine) Then
                        MsgBox("Error: 'POrg' Code does not exist for this Service Line", MsgBoxStyle.Exclamation, Project_Name)
                        Exit Sub
                    End If
                End If

                If txt_VariantToApply.Text.Replace(" ", "").Length = 0 Then
                    MsgBox("Error: Review 'Variant To Apply' Value [Length must be > 0]", MsgBoxStyle.Exclamation, Project_Name)
                    Exit Sub
                End If


                Result = Send_Info()

                If Result Then
                    Me.DialogResult = Windows.Forms.DialogResult.OK
                    Me.Close()
                Else
                    Exit Sub
                End If

            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Function Send_Info() As Boolean

        Try
            Send_Info = True

            Me.Region_ = List_Region.SelectedValue
            Me.Plant_ = txt_Plant.Text.Trim.ToUpper
            Me.SAP_Box_ = List_SAP_Box.SelectedValue
            Me.POrg_ = txt_POrg.Text.Trim
            Me.Variant_ = List_Variant.SelectedValue
            Me.Variant_To_Apply_ = txt_VariantToApply.Text
            Me.Owner_ = List_Owner.SelectedValue
        Catch ex As Exception
            Send_Info = False
        End Try
        

    End Function



End Class