Imports System.Threading
Imports System.Globalization

Public Class Emails_Notifications

    Private SF As New System_Functions
    Private SQL_F As New SQL_Functions



    Public Sub Email_NotificationToEPO_AddDelete(ByVal Variant_Code As String, ByVal ServiceLine As String)
        Dim CP As String = Nothing
        Dim CC As String = Nothing

        Dim Subject As String = Nothing
        Dim HTMLBody As String = Nothing

        Dim FileName As String = Nothing
        Dim FilePath As String = Nothing

        Dim Attachment() As String = {"Bullio"}
        Dim Tracking_List As String = ""
        Dim Contains_Lines As Boolean = True

        '*********************************************************************
        Try
            Dim DT As DataTable
            DT = SQL_F.GetDataTable(CS, "Select * From Report_Americas_" & Variant_Code & " Where (Email_Date IS NULL) and (Change_Type in ('New Addition','Request for Deletion'))")

            If DT.Rows.Count > 0 Then

                FileName = "PSS Variant - " & ServiceLine & " (" & Date.Now.Month & "-" & Date.Now.Day & "-" & Date.Now.Year & ").xlsx"
                FilePath = EXCELPath & FileName

                'Delete The Last File
                Try
                    Kill(FilePath)
                Catch
                End Try

                'Create The New File
                Dim DT_Variant As DataTable
                DT_Variant = SQL_F.GetDataTable(CS, "Select * From Report_Americas_" & Variant_Code)
                If Not SF.ExportVariantToExcel(Variant_Code, FilePath) Then
                    Exit Sub
                End If


                'Define Contacts
                CP = SF.ReturnEmailContact("EPO_Contact")
                CC = SF.ReturnEmailContact("Infosys_Contact") & " ; " & SF.ReturnEmailContact(Variant_Code)

                'Validate Contact
                If Not SF.Validate_Email_List(CP) Then
                    CP = DefaultEmailContact
                    CC = Nothing
                End If

                'Define Body
                Subject = "PSS Variant DB - Pending Requests for Add-Delete Combinations (" & Date.Now.Month & "-" & Date.Now.Day & "-" & Date.Now.Year & ") [" & ServiceLine & "]"
                HTMLBody = SF.GetHTMLcode(HTMLPath & "EPO_Notification_AddDeletion.htm")
                HTMLBody = HTMLBody.Replace("XXXX", ServiceLine)

                'Attachment
                Attachment(0) = FilePath

                'Send Email
                If SF.Send_eMail(CP, CC, Subject, HTMLBody, Attachment, Nothing, DraftEmails) Then

                    For Each R As DataRow In DT.Rows

                        Dim Update As String = "Update Var_Americas_" & Variant_Code & " Set Email_Date = GetDate(),Pending_Approval=1  "
                        Update += SF.CreateWhereCondition(R, Variant_Code)

                        Dim Exc As String = SQL_F.SQL_Execute_NQ(CS, Update)
                        If Exc <> Nothing Then
                            Dim a As String = "STOP"
                        End If

                    Next

                Else
                    MsgBox("Error/Confirmation to Requester", MsgBoxStyle.Critical)
                End If

            End If

        Catch ex As Exception
            MsgBox("Error/Letter Generation to Requester:" & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Public Sub Email_NotificationToInfosys_Change(ByVal Variant_Code As String, ByVal ServiceLine As String)
        Dim CP As String = Nothing
        Dim CC As String = Nothing

        Dim Subject As String = Nothing
        Dim HTMLBody As String = Nothing

        Dim FileName As String = Nothing
        Dim FilePath As String = Nothing

        Dim Attachment() As String = {"Bullio"}
        Dim Tracking_List As String = ""
        Dim Contains_Lines As Boolean = True

        '*********************************************************************
        Try
            Dim DT As DataTable
            DT = SQL_F.GetDataTable(CS, "Select * From Report_Americas_" & Variant_Code & " Where (Email_Date IS NULL) and (Change_Type in ('Pending Owner Change'))")

            If DT.Rows.Count > 0 Then

                FileName = "PSS Variant - " & ServiceLine & " (" & Date.Now.Month & "-" & Date.Now.Day & "-" & Date.Now.Year & ").xlsx"
                FilePath = EXCELPath & FileName

                'Delete The Last File
                Try
                    Kill(FilePath)
                Catch
                End Try

                'Create The New File
                Dim DT_Variant As DataTable
                DT_Variant = SQL_F.GetDataTable(CS, "Select * From Report_Americas_" & Variant_Code)
                If Not SF.ExportVariantToExcel(Variant_Code, FilePath) Then
                    Exit Sub
                End If


                'Define Contacts
                CP = SF.ReturnEmailContact("Infosys_Contact")
                CC = SF.ReturnEmailContact("EPO_Contact") & " ; " & SF.ReturnEmailContact(Variant_Code)

                'Validate Contact
                If Not SF.Validate_Email_List(CP) Then
                    CP = DefaultEmailContact
                    CC = Nothing
                End If

                'Define Body
                Subject = "PSS Variant DB - Pending Requests for SPS Change (" & Date.Now.Month & "-" & Date.Now.Day & "-" & Date.Now.Year & ") [" & ServiceLine & "]"
                HTMLBody = SF.GetHTMLcode(HTMLPath & "EPO_Notification_Change.htm")
                HTMLBody = HTMLBody.Replace("XXXX", ServiceLine)

                'Attachment
                Attachment(0) = FilePath

                'Send Email
                If SF.Send_eMail(CP, CC, Subject, HTMLBody, Attachment, Nothing, DraftEmails) Then

                    For Each R As DataRow In DT.Rows

                        Dim Update As String = "Update Var_Americas_" & Variant_Code & " Set Email_Date = GetDate(),Pending_Approval=1  "
                        Update += SF.CreateWhereCondition(R, Variant_Code)

                        Dim Exc As String = SQL_F.SQL_Execute_NQ(CS, Update)
                        If Exc <> Nothing Then
                            Dim a As String = "STOP"
                        End If

                    Next

                Else
                    MsgBox("Error/Confirmation to Requester", MsgBoxStyle.Critical)
                End If

            End If

        Catch ex As Exception
            MsgBox("Error/Letter Generation to Requester:" & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

 


End Class
