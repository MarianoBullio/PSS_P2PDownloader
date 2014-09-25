'**************************************
' Module: Popup New Row
' Created By: Mariano Bullio, bullio.m@pg.com
'**************************************

Public Class Americas_LogImpLA_NewRow

    Private SQL_F As New SQL_Functions
    Private SF As New System_Functions

    Friend Result As Boolean = False

    Friend Plant_ As String = ""
    Friend SAP_Box_ As String = ""
    Friend POrg_ As String = ""
    Friend Spend_ As String = ""
    Friend Owner_ As String = ""


    Private ServiceLine As String = "LogImpLA"


    Private Sub Form_Load() Handles Me.Load

        Try

            SF.Load_ComboBox(List_SAP_Box, CS, "SELECT SAP_Box, Description FROM SAP_Boxes Where (" & ServiceLine & " = 1) Order By Description Desc")
            SF.Load_ListBox(List_Owner, CS, "SELECT TNumber, Name FROM Variant_Users Where (Active = 1) AND (Employee = 1) AND (SPS =1) Order By Name Asc")
            SF.Load_ComboBox(List_Spend, CS, "SELECT Spend, Description From Var_Americas_LogImpLA_Spends Order By Description Desc")
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
            Me.Plant_ = txt_Plant.Text.Trim.ToUpper
            Me.SAP_Box_ = List_SAP_Box.SelectedValue
            Me.POrg_ = txt_POrg.Text.Trim
            Me.Spend_ = List_Spend.SelectedValue
            Me.Owner_ = List_Owner.SelectedValue
        Catch ex As Exception
            Send_Info = False
        End Try


    End Function





End Class