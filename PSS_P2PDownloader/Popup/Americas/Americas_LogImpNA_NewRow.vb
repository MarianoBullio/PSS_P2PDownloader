'**************************************
' Module: Popup New Row
' Created By: Mariano Bullio, bullio.m@pg.com
'**************************************

Public Class Americas_LogImpNA_NewRow

    Private SQL_F As New SQL_Functions
    Private SF As New System_Functions

    Friend Result As Boolean = False

    Friend Vendor_ As String = ""
    Friend Vendor_Name_ As String = ""
    Friend Owner_ As String = ""

    Private ServiceLine As String = "STR"

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try
            SF.Load_ListBox(List_Owner, CS, "SELECT TNumber, Name FROM Variant_Users Where (Active = 1) AND (Employee = 1) AND (SPS =1) Order By Name Asc")
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


                If txt_Vendor.Text.Trim.Length < 8 Then
                    MsgBox("Error: Review 'Vendor' Value [Length must be = 8]", MsgBoxStyle.Exclamation, Project_Name)
                    Exit Sub
                Else
                    If Not IsNumeric(txt_Vendor.Text.Trim) Then
                        MsgBox("Error: Review 'Vendor' Value [It must be Numeric]", MsgBoxStyle.Exclamation, Project_Name)
                        Exit Sub
                    End If
                End If

                If txt_VendorName.Text.Trim.Length = 0 Then
                    MsgBox("Error: Review 'Vendor Name' Value [Length must be > 0]", MsgBoxStyle.Exclamation, Project_Name)
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

            Me.Vendor_ = txt_Vendor.Text.Trim
            Me.Vendor_Name_ = txt_VendorName.Text.Trim
            Me.Owner_ = List_Owner.SelectedValue
        Catch ex As Exception
            Send_Info = False
        End Try


    End Function



End Class