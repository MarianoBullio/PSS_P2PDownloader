
'**************************************
' Project: GetSAPTables
' Created By: Mariano Bullio, bullio.m@pg.com
'**************************************M

Imports Microsoft.VisualBasic
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.Data
Imports System.IO
Imports System.Globalization
Imports System.Threading

Public Class GetSAPTables

    'Document Lines
    Public Function Get_EKPO(ByVal Box As String, ByVal Login As String, ByVal Password As String, ByRef DT As DataTable) As DataTable

        Dim D As New SAPCOM.ConnectionData
        D.Box = Box
        D.Login = Login
        D.Password = Password
        D.SSO = False

        Dim SC As New SAPCOM.SAPConnector
        Dim Con = SC.GetSAPConnection(D)

        Dim EKPO As New SAPCOM.EKPO_Report(Con)
        Dim DT_Result As New DataTable

        Try
            EKPO.AddCustomField("AFNAM", "AFNAM")
            For Each R As DataRow In DT.Rows
                EKPO.IncludeDocument(R("Purch_Doc").ToString.Trim)
            Next
            '

            EKPO.Execute()
            DT_Result = EKPO.Data

            Dim A As String = EKPO.ErrMessage
            DT_Result.Columns("Doc Number").ColumnName = "Purch_Doc"
            DT_Result.Columns("Mat Group").ColumnName = "Material_Group"
            DT_Result.Columns("Item Number").ColumnName = "LineItem"
            Get_EKPO = DT_Result
        Catch ex As Exception
            Get_EKPO = Nothing
        End Try


    End Function

    'Vendor Master (Vendor-CompanyCode/LE)
    Public Function Get_LFB1(ByVal Box As String, ByVal Login As String, ByVal Password As String, ByRef DT As DataTable) As DataTable
        Dim D As New SAPCOM.ConnectionData
        D.Box = Box
        D.Login = Login
        D.Password = Password
        D.SSO = False

        Dim SC As New SAPCOM.SAPConnector
        Dim Con = SC.GetSAPConnection(D)

        Dim LFB1 As New SAPCOM.LFB1_Report(Con)
        Dim DT_Result As New DataTable

        Try
            For Each R As DataRow In DT.Rows
                LFB1.IncludeVendor(R("Vendor").ToString.Trim)
            Next

            For Each R As DataRow In DT.Rows
                LFB1.Include_CCode(R("LE").ToString.Trim)
            Next

            LFB1.AddCustomField("ZWELS", "Payment_Method")

            LFB1.Execute()
            DT_Result = LFB1.Data

            Dim A As String = LFB1.ErrMessage
            DT_Result.Columns("Pmnt Terms").ColumnName = "PTerms"
            DT_Result.Columns("CCode").ColumnName = "LE"
            DT_Result.Columns("Created On").ColumnName = "Created_On"
            DT_Result.Columns("Created By").ColumnName = "Created_By"
            DT_Result.Columns("Block").ColumnName = "Blocked"
            DT_Result.Columns("Delete").ColumnName = "Deleted"

            Get_LFB1 = DT_Result
        Catch ex As Exception
            Get_LFB1 = Nothing
        End Try

    End Function


End Class
