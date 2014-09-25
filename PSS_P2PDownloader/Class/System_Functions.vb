'**************************************
' Module: System_Functions - P2P Downloader System
' Created By: Mariano Bullio, bullio.m@pg.com
'**************************************

Imports Microsoft.VisualBasic
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.Data
Imports System.IO
Imports OfficeOpenXml
Imports System.Globalization
Imports System.Threading
Imports System.Text.RegularExpressions
Imports SAPCOM
Imports SAPCOM.SAPTextIDs
Imports Microsoft.Office.Interop

Public Class System_Functions

    Dim SQL_F As New SQL_Functions
    '"User Functions"

    Public Function ReturnUserTNumber() As String
        Try
            ReturnUserTNumber = Environ("USERID")
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Public Function ReturnUserName(ByVal Login As String) As String
        Try
            Dim Name As String = SQL_F.ReturnDescription(CS, "Variant_Users", "Name", "TNumber ='" & Login & "'")
            If Name <> "" Then
                ReturnUserName = Name
            Else
                ReturnUserName = "Unknown"
            End If
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Public Function ReturnUserNameByEmail(ByVal Email As String) As String
        Try
            Dim Name As String = SQL_F.ReturnDescription(CS, "Variant_Users", "Name", "Email ='" & Email & "'")
            If Name <> "" Then
                ReturnUserNameByEmail = Name
            Else
                ReturnUserNameByEmail = "Unknown"
            End If
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Public Function ValidateUser(ByVal User As String) As Boolean
        Dim P As String = SQL_F.SQL_Execute_SC(CS, "Select TNumber From Variant_Users Where (TNumber='" & Login & "') and (Access=1)")
        If P Is Nothing Then
            ValidateUser = False
            Exit Function
        Else
            ValidateUser = True
        End If
    End Function

    Public Function IsAdministrator(ByVal TNumber As String) As Boolean
        Dim response As String = SQL_F.ReturnDescription(CS, "Variant_Users", "Administrator", "TNumber='" & TNumber & "'")
        If response = "False" Then
            IsAdministrator = False
        Else
            IsAdministrator = True
        End If
    End Function

    Public Function IsVariantsSPOC(ByVal TNumber As String) As Boolean
        Dim response As String = SQL_F.ReturnDescription(CS, "Variant_Users", "VariantsSPOC", "TNumber='" & TNumber & "'")
        If response = "False" Then
            IsVariantsSPOC = False
        Else
            IsVariantsSPOC = True
        End If
    End Function

    Public Function IsInfosysEmployee(ByVal TNumber As String) As Boolean
        Dim response As String = SQL_F.ReturnDescription(CS, "Variant_Users", "InfosysEmployee", "TNumber='" & TNumber & "'")
        If response = "False" Then
            IsInfosysEmployee = False
        Else
            IsInfosysEmployee = True
        End If
    End Function

    Public Function IsReadOnly(ByVal TNumber As String) As Boolean
        Dim response As String = SQL_F.ReturnDescription(CS, "Variant_Users", "ReadOnly", "TNumber='" & TNumber & "'")
        If response = "False" Then
            IsReadOnly = False
        Else
            IsReadOnly = True
        End If
    End Function

    Public Function ValidateAccessToVariant(ByVal User As String, ByVal VariantCode As String) As Boolean
        Dim P As String = SQL_F.SQL_Execute_SC(CS, "Select TNumber From Variant_Users Where (TNumber='" & Login & "') and (" & VariantCode & "=1)")
        If P Is Nothing Then
            ValidateAccessToVariant = False
            Exit Function
        Else
            ValidateAccessToVariant = True
        End If
    End Function

    Public Function ReturnEmailContact(ByVal Email_Code As String) As String
        Try
            Dim Email As String = SQL_F.ReturnDescription(CS, "Email_Contacts", "Email_Contact", "Email_Code ='" & Email_Code & "'")
            If Email <> "" Then
                ReturnEmailContact = Email
            Else
                ReturnEmailContact = DefaultEmailContact
            End If
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    'Public Function ReturnUserBackup(ByVal Login As String) As String
    '    Try
    '        Dim Name As String = SQL_F.ReturnDescription(CS, "P2P_Users", "BackupTNumber", "TNumber ='" & Login & "'")
    '        If Name <> "" Then
    '            ReturnUserBackup = Name
    '        Else
    '            ReturnUserBackup = "Unknown"
    '        End If
    '    Catch ex As Exception
    '        Return Nothing
    '    End Try
    'End Function

    '"HTML Functions"

    Public Function GetHTMLcode(ByVal path As String) As String
        Dim HTMLcode As String = ""
        If System.IO.File.Exists(path) = True Then
            Dim objReader As New System.IO.StreamReader(path)
            HTMLcode = objReader.ReadToEnd
            objReader.Close()
        Else
            MsgBox("File: " & path & " Does Not Exist")
        End If
        Return HTMLcode
    End Function

    '"File Functions"

    Public Function DeleteDirectoryFiles(ByVal Directory_Path As String) As Boolean
        Try
            Dim dir As New IO.DirectoryInfo(Directory_Path)
            Dim fa() As IO.FileInfo
            Dim f As IO.FileInfo
            fa = dir.GetFiles

            For Each f In fa
                f.Delete()
            Next

            Dim da() As IO.DirectoryInfo
            Dim d1 As IO.DirectoryInfo
            da = dir.GetDirectories

            For Each d1 In da
                DeleteDirectoryFiles(d1.FullName)
            Next

            DeleteDirectoryFiles = True
        Catch ex As Exception
            DeleteDirectoryFiles = False
        End Try
    End Function

    Public Function ReturnExcelFilePath(ByVal OpenFileDialog As OpenFileDialog, ByVal Ext As String) As String
        Dim FileName As String = ""

        OpenFileDialog.Filter = "Excel Worksheets|*." & Ext
        OpenFileDialog.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.Desktop
        'While Not FileName.Contains(".xls")
        OpenFileDialog.FileName = "." & Ext
        OpenFileDialog.ShowDialog()
        FileName = OpenFileDialog.FileName
        'End While

        If OpenFileDialog.ValidateNames And FileName.Contains(".xlsx") Then
            ReturnExcelFilePath = FileName
        Else
            ReturnExcelFilePath = Nothing
        End If
    End Function

    Public Sub CreateExcelFile(ByVal FilePath As String, ByVal View As String, ByVal Sheet As String)
        Try
            Dim DT As DataTable = SQL_F.GetDataTable(CS, "Select * From " & View)

            DataTableToExcel(DT, FilePath, Sheet)
        Catch ex As Exception
            MsgBox("Error To Create Report, Please Contact The System Administrator", MsgBoxStyle.Critical)
        End Try
    End Sub

    Public Function GetFileText(ByVal path As String) As String
        Dim Text As String = ""
        If System.IO.File.Exists(path) = True Then
            Dim objReader As New System.IO.StreamReader(path)
            Text = objReader.ReadToEnd
            objReader.Close()
        Else
            MsgBox("File: " & path & " Does Not Exist")
        End If
        Return Text
    End Function

    Public Function DataTableToExcel(ByVal dt As DataTable, ByVal ExcelFilePath As String, ByVal SheetName As String) As Boolean

        Dim Exc As Excel.Application = Nothing
        Dim WB As Excel.Workbook
        Dim WS As Excel.Worksheet = Nothing
        Dim ExcConstans As Excel.Constants
        Dim misValue As Object = System.Reflection.Missing.Value
        Dim Dir As Excel.XlDirection

        Thread.CurrentThread.CurrentCulture = New CultureInfo("en-US", False)
        Dim aAA As String = My.Application.Culture.ToString

        Try
            Exc = New Excel.ApplicationClass()
            WB = Exc.Workbooks.Add(misValue)
            WS = WB.Sheets("sheet1")
            WS.Name = SheetName

            Dim ColumnIndex As Integer = 1
            Dim Rowindex As Integer = 2
            Dim ColumnNumber As Integer = dt.Columns.Count

            'Escribimos las columnas
            For Each dc As DataColumn In dt.Columns
                Exc.Cells(1, ColumnIndex) = dc.ToString

                'Cambiamos el Color de los encabezados
                Try
                    Exc.Cells(1, ColumnIndex).Select()
                    With Exc.Selection.Interior
                        .Pattern = ExcConstans.xlSolid
                        .PatternColorIndex = ExcConstans.xlAutomatic
                        '.ThemeColor = xlThemeColorDark1
                        .TintAndShade = -0.349986266670736
                        .PatternTintAndShade = 0
                    End With
                Catch
                End Try

                'Acomodamos el tamano de las Columnas a sus valores
                Try
                    Exc.Cells(1, ColumnIndex).EntireColumn.AutoFit()
                    ColumnIndex = ColumnIndex + 1
                Catch
                End Try
            Next

            'Escribimos las Filas
            For Each dr As DataRow In dt.Rows
                For c = 1 To ColumnNumber
                    Try
                        Exc.Cells(Rowindex, c) = dr(c - 1).ToString
                    Catch
                        Exc.Cells(Rowindex, c) = Nothing
                    End Try
                Next
                Rowindex = Rowindex + 1
            Next

            Exc.Columns("A:BF").Select()
            Exc.Columns("A:BF").EntireColumn.AutoFit()
            Exc.ActiveWindow.ScrollColumn = 1

            Exc.Rows("1:1").Select()
            Exc.Range(Exc.Selection, Exc.Selection.End(Dir.xlDown)).Select()
            Exc.Selection.RowHeight = 15

            'Guardamos el Archivo de Excel
            'WB.SaveAs(ExcelFilePath)
            Try
                'http://msdn.microsoft.com/en-us/library/microsoft.office.interop.excel.xlfileformat%28office.11%29.aspx
                WB.SaveAs(ExcelFilePath, 56)
            Catch ex As Exception
                WB.SaveAs(ExcelFilePath)
            End Try

            WB.Close()
            Exc.Quit()
            releaseObject(Exc)
            releaseObject(WB)
            releaseObject(WS)
            DataTableToExcel = True
        Catch
            DataTableToExcel = False
        End Try
    End Function

    Public Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    Public Function GetDirectoryFilesCount(ByVal Folder_Path As String) As Integer
        Try
            Dim dir As New IO.DirectoryInfo(Folder_Path)
            Dim fa() As IO.FileInfo
            fa = dir.GetFiles
            GetDirectoryFilesCount = fa.Count()
        Catch ex As Exception
            GetDirectoryFilesCount = 0
        End Try
    End Function


    '"Other Fuctions"

    Public Function ShowConfirmationMessage(ByVal message As String, ByVal Title As String) As Boolean

        Try
            If MsgBox(message, MsgBoxStyle.OkCancel, Title) = MsgBoxResult.Ok Then
                ShowConfirmationMessage = True
            Else
                ShowConfirmationMessage = False
            End If
        Catch ex As Exception
            ShowConfirmationMessage = False
        End Try
    End Function

    Public Function ShowConfirmationMessage() As Boolean
        Try
            If MsgBox("Please Confirm To Continue...", MsgBoxStyle.OkCancel, "Confirmation Message") = MsgBoxResult.Ok Then
                ShowConfirmationMessage = True
            Else
                ShowConfirmationMessage = False
            End If
        Catch ex As Exception
            ShowConfirmationMessage = False
        End Try
    End Function

    Public Sub Load_ComboBox(ByVal ComboBox As ComboBox, ByVal CS As String, ByVal sql_select As String)
        Dim cn As New SqlConnection(CS)
        Try
            cn.Open()
            Dim cmd As New SqlCommand(sql_select, cn)
            Dim da As New SqlDataAdapter(cmd)
            Dim ds As New DataSet
            da.Fill(ds)
            ComboBox.DataSource = ds.Tables(0)
            ComboBox.ValueMember = ds.Tables(0).Columns(0).Caption.ToString
            ComboBox.DisplayMember = ds.Tables(0).Columns(1).Caption.ToString
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString, "error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If cn.State = ConnectionState.Open Then
                cn.Close()
            End If
        End Try
    End Sub

    Public Sub Load_ListBox(ByVal ComboBox As ListBox, ByVal CS As String, ByVal sql_select As String)
        Dim cn As New SqlConnection(CS)
        Try
            cn.Open()
            Dim cmd As New SqlCommand(sql_select, cn)
            Dim da As New SqlDataAdapter(cmd)
            Dim ds As New DataSet
            da.Fill(ds)
            ComboBox.DataSource = ds.Tables(0)
            ComboBox.ValueMember = ds.Tables(0).Columns(0).Caption.ToString
            ComboBox.DisplayMember = ds.Tables(0).Columns(1).Caption.ToString
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString, "error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If cn.State = ConnectionState.Open Then
                cn.Close()
            End If
        End Try
    End Sub

    Public Function BooleanConvertion(ByVal val As String) As Integer
        BooleanConvertion = 1
        If (val = "False") Then
            BooleanConvertion = 0
        End If
    End Function

    Public Function GetHSBarPos(ByVal dg As DataGridView) As Integer
        For Each c As Control In dg.Controls
            If TypeOf (c) Is HScrollBar Then
                Dim sBar As ScrollBar = CType(c, ScrollBar)
                GetHSBarPos = sBar.Value
            End If
        Next
    End Function

    Public Function GetVSBarPos(ByVal dg As DataGridView) As Integer
        For Each c As Control In dg.Controls
            If TypeOf (c) Is VScrollBar Then
                Dim sBar As ScrollBar = CType(c, ScrollBar)
                GetVSBarPos = sBar.Value
            End If
        Next
    End Function

    Public Sub SetSBarPos(ByVal dg As DataGridView, ByVal HSBarPos As Integer, ByVal VSBarPos As Integer)

        For Each c As Control In dg.Controls
            If TypeOf (c) Is HScrollBar Then
                Dim sBar As ScrollBar = CType(c, ScrollBar)
                sBar.Value = HSBarPos
            End If

            If TypeOf (c) Is VScrollBar Then
                Dim sBar As ScrollBar = CType(c, ScrollBar)
                sBar.Value = VSBarPos
            End If

        Next
    End Sub

    Public Sub CreateCellCalendar(ByRef DataGridView_ As DataGridView, ByVal ColumnName As String)
        Dim I As Integer
        Dim Calendar As New CalendarColumn()
        I = DataGridView_.Columns(ColumnName).DisplayIndex
        DataGridView_.Columns.Remove(ColumnName)
        Calendar.Name = ColumnName
        Calendar.DataPropertyName = ColumnName
        Calendar.HeaderText = ColumnName
        Calendar.DisplayIndex = I
        DataGridView_.Columns.Add(Calendar)
    End Sub

    Public Function ReturnMonth(ByVal m As String) As String
        Dim return_month As String = ""
        Try
            Select Case m
                Case "1"
                    return_month = "Jan"
                Case "2"
                    return_month = "Feb"
                Case "3"
                    return_month = "Mar"
                Case "4"
                    return_month = "Apr"
                Case "5"
                    return_month = "May"
                Case "6"
                    return_month = "Jun"
                Case "7"
                    return_month = "Jul"
                Case "8"
                    return_month = "Aug"
                Case "9"
                    return_month = "Sep"
                Case "10"
                    return_month = "Oct"
                Case "11"
                    return_month = "Nov"
                Case "12"
                    return_month = "Dec"
                Case Else
                    return_month = ""
            End Select
        Catch
            return_month = ""
        End Try
        Return return_month
    End Function

    Public Function ReturnExcelFilePathXSLX(ByVal OpenFileDialog As OpenFileDialog) As String
        Dim FileName As String = ""

        OpenFileDialog.Filter = "Excel Worksheets|*.xlsx"
        OpenFileDialog.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.Desktop
        'While Not FileName.Contains(".xls")
        OpenFileDialog.FileName = ".xlsx"
        OpenFileDialog.ShowDialog()
        FileName = OpenFileDialog.FileName
        'End While

        If OpenFileDialog.ValidateNames And FileName.Contains(".xlsx") Then
            ReturnExcelFilePathXSLX = FileName
        Else
            ReturnExcelFilePathXSLX = Nothing
        End If
    End Function

    Public Function CreateWhereCondition(ByRef R As DataRow, ByVal Variant_Code As String) As String
        Try

            Dim DT_Variants_PrimaryKeys As New DataTable
            DT_Variants_PrimaryKeys = SQL_F.GetDataTable(CS, "Select * From Variants_PrimaryKeys Where Variant_Code='" & Variant_Code & "'")

            Dim Where As String = ""
            Where += "Where "

            For Each Key As DataRow In DT_Variants_PrimaryKeys.Rows
                Where += " ("

                Dim a As String = Key("IsNumeric").ToString.Trim
                If Key("IsNumeric").ToString.Trim = "True" Then
                    Where += Key("Column_Key").ToString.Trim & " = " & R(Key("Column_Key").ToString.Trim)
                Else
                    Where += Key("Column_Key").ToString.Trim & " = '" & R(Key("Column_Key").ToString.Trim) & "'"
                End If

                Where += ") AND"
            Next

            Where = Where.Remove((Where.Length - 3), 3) 'Delete last "AND"

            CreateWhereCondition = Where
        Catch ex As Exception
            CreateWhereCondition = ""
        End Try
    End Function

    'Email Functions"

    Public Function Validate_Email(ByVal sMail As String) As Boolean
        Return Regex.IsMatch(sMail, "^([\w-]+\.)*?[\w-]+@[\w-]+\.([\w-]+\.)*?[\w]+$")
    End Function

    Public Function Validate_Email_List(ByVal sMail As String) As Boolean
        'Example = "collins.jl.1@pg.com;collins.ra.1@pg.com;whitaker.dw@pg.com;collins.ra.1@pg.com"
        Dim emailExpression As New Regex("\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*([,;]\s*\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*)*")
        Return emailExpression.IsMatch(sMail)
    End Function

    Public Function Send_eMail(ByVal Recipient As String, ByVal CopyTo As String, ByVal Subject As String, ByVal Body As String, ByVal Attachments() As String, ByVal OnBehalfOf As String, Optional ByVal Draft As Boolean = False) As Boolean
        Dim MyolApp
        Dim myNameSpace
        Dim objMail
        Dim A As String
        Try
            MyolApp = CreateObject("Outlook.Application")
            myNameSpace = MyolApp.GetNamespace("MAPI")
            objMail = MyolApp.CreateItem(0)

            With objMail
                MyolApp = .GetInspector()
                .Subject = Subject
                .HTMLBody = Body & .HTMLBody
                .To = Recipient
                .CC = CopyTo
                .SentOnBehalfOfName = OnBehalfOf

                If Not Attachments Is Nothing Then
                    For Each A In Attachments
                        If A <> "" Then
                            .Attachments.Add(CStr(A))
                        End If
                    Next
                End If
            End With

            'SMI = CreateObject("Redemption.SafeMailItem")
            'SMI.Item = objMail
            'SMI.Send()

            If Not Draft Then
                objMail.Send()
            Else
                objMail.Save()
            End If

            Send_eMail = True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Outlook Error")
            Send_eMail = False
        End Try
    End Function

    'System Functions

    Public Function GetDateString() As String
        Thread.CurrentThread.CurrentCulture = New CultureInfo("en-US", False)
        Dim Time As DateTime = DateTime.Now
        Dim format As String = "MM/dd/yyyy HH:mm:ss"
        Dim d As String = Time.ToString(format)
        Return d
    End Function

    Public Function ValidatePlant(ByVal Plant As String) As Boolean
        Dim P As String = SQL_F.SQL_Execute_SC(CS, "Select Plant From Plants Where (Plant='" & Plant & "')")
        If P Is Nothing Then
            ValidatePlant = False
            Exit Function
        Else
            ValidatePlant = True
        End If
    End Function

    Public Function ValidatePOrg(ByVal POrg As String, ByVal ServiceLine As String) As Boolean
        Dim P As String = SQL_F.SQL_Execute_SC(CS, "Select POrg From POrgs Where (POrg=" & POrg & ") AND (" & ServiceLine & "=1) AND(Enabled=1)")
        If P Is Nothing Then
            ValidatePOrg = False
            Exit Function
        Else
            ValidatePOrg = True
        End If
    End Function

    Public Function ValidatePGrp(ByVal PGrp As String) As Boolean
        Dim P As String = SQL_F.SQL_Execute_SC(CS, "Select PGrp From PGrps Where (PGrp=" & PGrp & ") AND(Enabled=1)")
        If P Is Nothing Then
            ValidatePGrp = False
            Exit Function
        Else
            ValidatePGrp = True
        End If
    End Function

    Public Function MassiveUpdate(ByVal Variant_Code As String) As Integer
        Dim FRM_Small As New Massive_OwnerChange
        Dim Owner As String = ""
        Dim CountOfCombinations As String = 0
        FRM_Small.Variant_Code = Variant_Code

        Try
            While Not FRM_Small.Result
                If (FRM_Small.ShowDialog() = Windows.Forms.DialogResult.OK) Then
                    If FRM_Small.Result Then
                        Owner = FRM_Small.Owner_
                        CountOfCombinations = FRM_Small.CountOfCombinations_
                    End If
                End If

                If Not FRM_Small.Result Then 'If Cancel
                    MassiveUpdate = 0
                    Exit Function
                End If

                MassiveUpdate = CountOfCombinations

            End While

        Catch ex As Exception
            MassiveUpdate = 0
        End Try
    End Function

    Public Function MassiveRejection(ByVal Variant_Code As String) As Integer

        Dim CountRejectedItems As Integer = 0
        Try
            Dim DT_Update As New DataTable
            DT_Update = SQL_F.GetSQLTable(CS, "Var_Americas_" & Variant_code, " (Flag = 1) AND (Change_Type IN (1)) AND (Email_Date Is Not NULL) AND (Pending_Approval=1)")

            Dim DT_Addition As New DataTable
            DT_Addition = SQL_F.GetSQLTable(CS, "Var_Americas_" & Variant_code, " (Flag = 1) AND (Change_Type IN (2)) AND (Email_Date Is Not NULL) AND (Pending_Approval=1)")

            Dim DT_Deletion As New DataTable
            DT_Deletion = SQL_F.GetSQLTable(CS, "Var_Americas_" & Variant_code, " (Flag = 1) AND (Change_Type IN (3)) AND (Email_Date Is Not NULL) AND (Pending_Approval=1)")

            If ShowConfirmationMessage("Would you like to Reject the Following Requests?:" & vbNewLine & "[" & DT_Update.Rows.Count & "] Pending Owner Changes" & vbNewLine & "[" & DT_Addition.Rows.Count & "] New Additions" & vbNewLine & "[" & DT_Deletion.Rows.Count & "] Request for Deletions", "Massive Rejection") Then

                'Reject Pending Updates (Change_Type 1)
                For Each R As DataRow In DT_Update.Rows
                    Try
                        Dim Update As String = "Update Var_Americas_" & Variant_code
                        Update += " Set"
                        Update += " Suggested_Owner = 'TBD',"
                        Update += " Change_Type = 0,"
                        Update += " Email_Date = NULL,"
                        Update += " Pending_Approval = 0 "
                        Update += CreateWhereCondition(R, Variant_code)
                        Dim Exc As String = SQL_F.SQL_Execute_NQ(CS, Update)
                        If Exc <> Nothing Then
                            MsgBox("Error:" & Standart_MessageError)
                        Else
                            CountRejectedItems += 1
                        End If
                    Catch ex As Exception
                        MsgBox("Error:" & Standart_MessageError)
                    End Try
                Next

                'Reject Pending Additions (Change_Type 2)
                For Each R As DataRow In DT_Addition.Rows
                    Try
                        Dim Delete As String = "Delete From Var_Americas_" & Variant_code & " "
                        Delete += CreateWhereCondition(R, Variant_code)
                        Dim Exc As String = SQL_F.SQL_Execute_NQ(CS, Delete)
                        If Exc <> Nothing Then
                            MsgBox("Error:" & Standart_MessageError)
                        Else
                            CountRejectedItems += 1
                        End If
                    Catch ex As Exception
                        MsgBox("Error:" & Standart_MessageError)
                    End Try

                Next

                'Reject Pending Deletes (Change_Type 3)
                For Each R As DataRow In DT_Deletion.Rows
                    Try
                        Dim Update As String = "Update Var_Americas_" & Variant_code
                        Update += " Set"
                        Update += " Suggested_Owner = 'TBD',"
                        Update += " Change_Type = 0,"
                        Update += " Email_Date = NULL,"
                        Update += " Pending_Approval = 0 "
                        Update += CreateWhereCondition(R, Variant_code)
                        Dim Exc As String = SQL_F.SQL_Execute_NQ(CS, Update)
                        If Exc <> Nothing Then
                            MsgBox("Error:" & Standart_MessageError)
                        Else
                            CountRejectedItems += 1
                        End If
                    Catch ex As Exception
                        MsgBox("Error:" & Standart_MessageError)
                    End Try
                Next

                MassiveRejection = CountRejectedItems
            Else
                MassiveRejection = 0
            End If

        Catch ex As Exception
            MassiveRejection = 0
        End Try
    End Function

    Public Function MassiveApproval(ByVal Variant_Code As String) As Integer
        Dim CountApprovedItems As Integer = 0
        Try
            Dim DT_Update As New DataTable
            DT_Update = SQL_F.GetSQLTable(CS, "Var_Americas_" & Variant_Code, " (Flag = 1) AND (Change_Type IN (1)) AND (Email_Date Is Not NULL) AND (Pending_Approval=1)")

            Dim DT_Addition As New DataTable
            DT_Addition = SQL_F.GetSQLTable(CS, "Var_Americas_" & Variant_Code, " (Flag = 1) AND (Change_Type IN (2)) AND (Email_Date Is Not NULL) AND (Pending_Approval=1)")

            Dim DT_Deletion As New DataTable
            DT_Deletion = SQL_F.GetSQLTable(CS, "Var_Americas_" & Variant_Code, " (Flag = 1) AND (Change_Type IN (3)) AND (Email_Date Is Not NULL) AND (Pending_Approval=1)")

            If ShowConfirmationMessage("Would you like to Approve the Following Requests?:" & vbNewLine & "[" & DT_Update.Rows.Count & "] Pending Owner Changes" & vbNewLine & "[" & DT_Addition.Rows.Count & "] New Additions" & vbNewLine & "[" & DT_Deletion.Rows.Count & "] Request for Deletions", "Massive Approval") Then

                'Approve Pending Updates (Change_Type 1)
                For Each R As DataRow In DT_Update.Rows
                    Try
                        Dim Update As String = "Update Var_Americas_" & Variant_Code
                        Update += " Set"
                        Update += " Owner = '" & R("Suggested_Owner") & "',"
                        Update += " Suggested_Owner = 'TBD',"
                        Update += " Change_Type = 0,"
                        Update += " Email_Date = NULL,"
                        Update += " Pending_Approval = 0 "
                        Update += CreateWhereCondition(R, Variant_Code)
                        Dim Exc As String = SQL_F.SQL_Execute_NQ(CS, Update)
                        If Exc <> Nothing Then
                            MsgBox("Error:" & Standart_MessageError)
                        Else
                            CountApprovedItems += 1
                        End If
                    Catch ex As Exception
                        MsgBox("Error:" & Standart_MessageError)
                    End Try
                Next

                'Approve Pending Additions (Change_Type 2)
                For Each R As DataRow In DT_Addition.Rows
                    Try
                        Dim Update As String = "Update Var_Americas_" & Variant_Code
                        Update += " Set"
                        Update += " Enabled = 1,"
                        Update += " Owner = '" & R("Suggested_Owner") & "',"
                        Update += " Suggested_Owner = 'TBD',"
                        Update += " Change_Type = 0,"
                        Update += " Email_Date = NULL,"
                        Update += " Pending_Approval = 0 "
                        Update += CreateWhereCondition(R, Variant_Code)
                        Dim Exc As String = SQL_F.SQL_Execute_NQ(CS, Update)
                        If Exc <> Nothing Then
                            MsgBox("Error:" & Standart_MessageError)
                        Else
                            CountApprovedItems += 1
                        End If
                    Catch ex As Exception
                        MsgBox("Error:" & Standart_MessageError)
                    End Try
                Next

                'Approve Pending Deletes (Change_Type 3)
                For Each R As DataRow In DT_Deletion.Rows
                    Try
                        Dim Delete As String = "Delete From Var_Americas_" & Variant_Code & " "
                        Delete += CreateWhereCondition(R, Variant_Code)
                        Dim Exc As String = SQL_F.SQL_Execute_NQ(CS, Delete)
                        If Exc <> Nothing Then
                            MsgBox("Error:" & Standart_MessageError)
                        Else
                            CountApprovedItems += 1
                        End If
                    Catch ex As Exception
                        MsgBox("Error:" & Standart_MessageError)
                    End Try
                Next

                MassiveApproval = CountApprovedItems
            Else
                MassiveApproval = 0
            End If

        Catch ex As Exception
            MassiveApproval = 0
        End Try
    End Function


    Public Sub SetLogInfo(ByVal Description As String)
        Dim Exc As String = SQL_F.SQL_Execute_NQ(CS, "Insert Into P2P_Download_Log(Date,Description) Values(GetDate(),'" & Description & "')")
    End Sub

    'SAP Functions

    Public Function Test_Box_Connection(ByVal Box As String, ByVal Login As String, ByVal Password As String) As Boolean
        'Test Connection
        Try
            Dim D As New SAPCOM.ConnectionData

            D.Box = Box
            D.Login = Login
            D.Password = Password
            D.SSO = False

            Dim SC As New SAPCOM.SAPConnector
            Dim Con = SC.TestConnection(D)

            If Con Then
                Test_Box_Connection = True
            Else
                Test_Box_Connection = False
            End If
        Catch ex As Exception
            Test_Box_Connection = False
        End Try

    End Function

    Public Function Return_SAP_Password(ByVal Box As String, ByVal Client As Integer, ByVal Login As String) As String
        Return SQL_F.ReturnDescription(CS, "P2P_Users_SAP", "Password", "Box='" & Box & "' and Client=" & Client & " and Login='" & Login & "'")
    End Function

    Public Function Set_SAP_Password(ByVal Box As String, ByVal Client As String, ByVal Login As String, ByVal Password As String)
        Try
            If Return_If_SAP_Password_Exist(Box, Client, Login) Then
                Dim Update As String = "Update P2P_Users_SAP set Password='" & Password & "' Where Box='" & Box & "' and Client=" & Client & " and Login='" & Login & "'"

                If SQL_F.SQL_Execute_NQ(CS, Update) <> Nothing Then
                    Set_SAP_Password = False
                Else
                    Set_SAP_Password = True
                End If
            Else
                Dim Insert As String = "Insert into SCF_Users_SAP(Login,Box,Client,Password) Values('" & Login & "','" & Box & "'," & Client & ",'" & Password & "')"

                If SQL_F.SQL_Execute_NQ(CS, Insert) <> Nothing Then
                    Set_SAP_Password = False
                Else
                    Set_SAP_Password = True
                End If

            End If


        Catch ex As Exception
            Set_SAP_Password = False
        End Try
    End Function

    Public Function Return_If_SAP_Password_Exist(ByVal Box As String, ByVal Client As String, ByVal Login As String) As Boolean
        Dim Pass As String = Return_SAP_Password(Box, Client, Login)

        If Pass = "" Then
            Return_If_SAP_Password_Exist = False
        Else
            Return_If_SAP_Password_Exist = True
        End If
    End Function

    Public Sub Save_SAP_Box_Test_Status(ByVal Login As String, ByVal Box As String, ByVal Client As String, ByVal TestStatus As Integer)
        Dim Update As String = "Update P2P_SAP_Boxes set Test_Date=GetDate(), Test_Done=" & TestStatus & " Where (Box='" & Box & "') And (Client=" & Client & ") And (Login='" & Login & "')"
        Dim Exc As String = SQL_F.SQL_Execute_NQ(CS, Update)
    End Sub

    'Excel
    Public Function ExportDataTableToExcel(ByRef DT As DataTable, ByVal FilePath As String, ByVal Sheet As String, ByVal SheetLocation As String) As Boolean
        Try
            Dim WB_EEPLUS As New ExcelPackage(New IO.FileInfo(FilePath))
            WB_EEPLUS.Workbook.Worksheets.Add(Sheet)
            WB_EEPLUS.Workbook.Worksheets(Sheet).Cells(SheetLocation).LoadFromDataTable(DT, True)

            'Set Column Format for DateTime (When Location = A1)
            Dim ColumIndex As Integer = 0
            For Each C As DataColumn In DT.Columns
                Try
                    If C.DataType.FullName = "System.DateTime" Then
                        WB_EEPLUS.Workbook.Worksheets(Sheet).Column(ColumIndex + 1).Style.Numberformat.Format = "mm-dd-yy"
                    End If
                    ColumIndex += 1
                Catch ex As Exception
                End Try
            Next

            'Set Header Color
            For I = 1 To DT.Columns.Count
                WB_EEPLUS.Workbook.Worksheets(Sheet).Cells(1, I).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                WB_EEPLUS.Workbook.Worksheets(Sheet).Cells(1, I).Style.Fill.BackgroundColor.SetColor(Color.MidnightBlue)
                WB_EEPLUS.Workbook.Worksheets(Sheet).Cells(1, I).Style.Font.Color.SetColor(Color.White)
            Next


            WB_EEPLUS.Save()
            ExportDataTableToExcel = True
        Catch ex As Exception
            ExportDataTableToExcel = False
        End Try
    End Function

    Public Function ExportVariantToExcel(ByVal Variant_Code As String, ByVal FilePath As String) As Boolean
        Select Case Variant_Code

            Case "STR"
                ExportVariantToExcel = Export_Variant_STR(FilePath)
            Case "SS"
                ExportVariantToExcel = Export_Variant_SS(FilePath)
            Case "Direct"
                ExportVariantToExcel = Export_Variant_Direct(FilePath)
            Case "Custom"
                ExportVariantToExcel = Export_Variant_Custom(FilePath)
            Case "LogIndNAPD"
                ExportVariantToExcel = Export_Variant_LogIndNAPD(FilePath)
            Case "LogTMS"
                ExportVariantToExcel = Export_Variant_LogTMS(FilePath)
            Case "LogImpLA"
                ExportVariantToExcel = Export_Variant_LogImpLA(FilePath)
            Case "LogImpNA"
                ExportVariantToExcel = Export_Variant_LogImpNA(FilePath)
            Case "LogDirect"
                ExportVariantToExcel = Export_Variant_LogDirect(FilePath)
            Case Else
                ExportVariantToExcel = Export_AllVariants(FilePath)
        End Select
    End Function

    Private Function Export_Variant_STR(ByVal FilePath) As Boolean
        Try

            'Data Source
            Dim DT As New DataTable
            DT = SQL_F.GetDataTable(CS, "Select * From Report_Americas_STR")

            Try
                Kill(FilePath)
            Catch ex As Exception
            End Try

            'Create Excel File
            Dim WB_EEPLUS As New ExcelPackage(New IO.FileInfo(FilePath))
            WB_EEPLUS.Workbook.Worksheets.Add("Variant_STR")
            WB_EEPLUS.Workbook.Worksheets("Variant_STR").Cells("A1").LoadFromDataTable(DT, True)

            Dim ColumIndex As Integer = 0
            For Each C As DataColumn In DT.Columns
                Try
                    If C.DataType.FullName = "System.DateTime" Then
                        WB_EEPLUS.Workbook.Worksheets("Variant_STR").Column(ColumIndex + 1).Style.Numberformat.Format = "mm-dd-yy"
                    End If
                    ColumIndex += 1
                Catch ex As Exception
                End Try
            Next

            'Set Header Color
            For I = 1 To DT.Columns.Count
                WB_EEPLUS.Workbook.Worksheets("Variant_STR").Cells(1, I).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                WB_EEPLUS.Workbook.Worksheets("Variant_STR").Cells(1, I).Style.Fill.BackgroundColor.SetColor(Color.MidnightBlue)
                WB_EEPLUS.Workbook.Worksheets("Variant_STR").Cells(1, I).Style.Font.Color.SetColor(Color.White)
            Next

            'Color Coding for Change Type
            Dim Row As Integer = 2
            For Each R As DataRow In DT.Rows

                Select Case R("Change_Type")

                    Case "New Addition"
                        For I = 1 To DT.Columns.Count
                            WB_EEPLUS.Workbook.Worksheets("Variant_STR").Cells(Row, I).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                            WB_EEPLUS.Workbook.Worksheets("Variant_STR").Cells(Row, I).Style.Fill.BackgroundColor.SetColor(Color.Green)
                            WB_EEPLUS.Workbook.Worksheets("Variant_STR").Cells(Row, I).Style.Font.Color.SetColor(Color.White)
                        Next
                    Case "Request for Deletion"
                        For I = 1 To DT.Columns.Count
                            WB_EEPLUS.Workbook.Worksheets("Variant_STR").Cells(Row, I).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                            WB_EEPLUS.Workbook.Worksheets("Variant_STR").Cells(Row, I).Style.Fill.BackgroundColor.SetColor(Color.Red)
                            WB_EEPLUS.Workbook.Worksheets("Variant_STR").Cells(Row, I).Style.Font.Color.SetColor(Color.Black)
                        Next
                    Case "Pending Owner Change"
                        For I = 1 To DT.Columns.Count
                            WB_EEPLUS.Workbook.Worksheets("Variant_STR").Cells(Row, I).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                            WB_EEPLUS.Workbook.Worksheets("Variant_STR").Cells(Row, I).Style.Fill.BackgroundColor.SetColor(Color.Orange)
                            WB_EEPLUS.Workbook.Worksheets("Variant_STR").Cells(Row, I).Style.Font.Color.SetColor(Color.Black)
                        Next
                End Select
                Row += 1
            Next

            WB_EEPLUS.Save()
            Export_Variant_STR = True
        Catch ex As Exception
            Export_Variant_STR = False
        End Try
    End Function

    Private Function Export_Variant_SS(ByVal FilePath) As Boolean
        Try

            'Data Source
            Dim DT As New DataTable
            DT = SQL_F.GetDataTable(CS, "Select * From Report_Americas_SS")

            Try
                Kill(FilePath)
            Catch ex As Exception
            End Try

            'Create Excel File
            Dim WB_EEPLUS As New ExcelPackage(New IO.FileInfo(FilePath))
            WB_EEPLUS.Workbook.Worksheets.Add("Variant_SS")
            WB_EEPLUS.Workbook.Worksheets("Variant_SS").Cells("A1").LoadFromDataTable(DT, True)

            Dim ColumIndex As Integer = 0
            For Each C As DataColumn In DT.Columns
                Try
                    If C.DataType.FullName = "System.DateTime" Then
                        WB_EEPLUS.Workbook.Worksheets("Variant_SS").Column(ColumIndex + 1).Style.Numberformat.Format = "mm-dd-yy"
                    End If
                    ColumIndex += 1
                Catch ex As Exception
                End Try
            Next

            'Set Header Color
            For I = 1 To DT.Columns.Count
                WB_EEPLUS.Workbook.Worksheets("Variant_SS").Cells(1, I).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                WB_EEPLUS.Workbook.Worksheets("Variant_SS").Cells(1, I).Style.Fill.BackgroundColor.SetColor(Color.MidnightBlue)
                WB_EEPLUS.Workbook.Worksheets("Variant_SS").Cells(1, I).Style.Font.Color.SetColor(Color.White)
            Next

            'Color Coding for Change Type
            Dim Row As Integer = 2
            For Each R As DataRow In DT.Rows

                Select Case R("Change_Type")

                    Case "New Addition"
                        For I = 1 To DT.Columns.Count
                            WB_EEPLUS.Workbook.Worksheets("Variant_SS").Cells(Row, I).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                            WB_EEPLUS.Workbook.Worksheets("Variant_SS").Cells(Row, I).Style.Fill.BackgroundColor.SetColor(Color.Green)
                            WB_EEPLUS.Workbook.Worksheets("Variant_SS").Cells(Row, I).Style.Font.Color.SetColor(Color.White)
                        Next
                    Case "Request for Deletion"
                        For I = 1 To DT.Columns.Count
                            WB_EEPLUS.Workbook.Worksheets("Variant_SS").Cells(Row, I).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                            WB_EEPLUS.Workbook.Worksheets("Variant_SS").Cells(Row, I).Style.Fill.BackgroundColor.SetColor(Color.Red)
                            WB_EEPLUS.Workbook.Worksheets("Variant_SS").Cells(Row, I).Style.Font.Color.SetColor(Color.Black)
                        Next
                    Case "Pending Owner Change"
                        For I = 1 To DT.Columns.Count
                            WB_EEPLUS.Workbook.Worksheets("Variant_SS").Cells(Row, I).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                            WB_EEPLUS.Workbook.Worksheets("Variant_SS").Cells(Row, I).Style.Fill.BackgroundColor.SetColor(Color.Orange)
                            WB_EEPLUS.Workbook.Worksheets("Variant_SS").Cells(Row, I).Style.Font.Color.SetColor(Color.Black)
                        Next
                End Select
                Row += 1
            Next


            'Add Sub-Variants
            WB_EEPLUS.Workbook.Worksheets.Add("SubVariants_SS")

            Dim DT_SubVariants As DataTable
            DT_SubVariants = SQL_F.GetDataTable(CS, "Select * From Var_Americas_SS_Variants order by Variant")

            Dim ExcelColumn As Integer = 1

            For Each R As DataRow In DT_SubVariants.Rows

                Dim Variant_Number As Object
                Variant_Number = R("Variant")

                If Variant_Number < 10 Then
                    Variant_Number = "0" & Variant_Number
                Else
                    Variant_Number = Variant_Number
                End If

                If Variant_Number <> "01" Then
                    Dim DT_Variant As New DataTable
                    DT_Variant = SQL_F.GetDataTable(CS, "Select * From Var_Americas_SS_Var" & Variant_Number)

                    If DT_Variant.Rows.Count > 0 Then
                        WB_EEPLUS.Workbook.Worksheets("SubVariants_SS").Cells(1, ExcelColumn).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                        WB_EEPLUS.Workbook.Worksheets("SubVariants_SS").Cells(1, ExcelColumn).Style.Fill.BackgroundColor.SetColor(Color.MidnightBlue)
                        WB_EEPLUS.Workbook.Worksheets("SubVariants_SS").Cells(1, ExcelColumn).Style.Font.Color.SetColor(Color.White)

                        WB_EEPLUS.Workbook.Worksheets("SubVariants_SS").Cells(1, ExcelColumn).Value = R("Description")
                        WB_EEPLUS.Workbook.Worksheets("SubVariants_SS").Cells(3, ExcelColumn).LoadFromDataTable(DT_Variant, False)

                        Select Case Variant_Number
                            Case "02", "03", "04", "05", "06", "07", "10", "11", "12", "13", "16"
                                WB_EEPLUS.Workbook.Worksheets("SubVariants_SS").Cells(2, ExcelColumn).Value = "Material Group"

                                WB_EEPLUS.Workbook.Worksheets("SubVariants_SS").Cells(2, ExcelColumn).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                                WB_EEPLUS.Workbook.Worksheets("SubVariants_SS").Cells(2, ExcelColumn).Style.Fill.BackgroundColor.SetColor(Color.SkyBlue)
                                WB_EEPLUS.Workbook.Worksheets("SubVariants_SS").Cells(2, ExcelColumn).Style.Font.Color.SetColor(Color.Black)

                            Case "08", "09", "14", "15"
                                WB_EEPLUS.Workbook.Worksheets("SubVariants_SS").Cells(2, ExcelColumn).Value = "PGrp"

                                WB_EEPLUS.Workbook.Worksheets("SubVariants_SS").Cells(2, ExcelColumn).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                                WB_EEPLUS.Workbook.Worksheets("SubVariants_SS").Cells(2, ExcelColumn).Style.Fill.BackgroundColor.SetColor(Color.SkyBlue)
                                WB_EEPLUS.Workbook.Worksheets("SubVariants_SS").Cells(2, ExcelColumn).Style.Font.Color.SetColor(Color.Black)
                            Case "17"
                                WB_EEPLUS.Workbook.Worksheets("SubVariants_SS").Cells(2, ExcelColumn).Value = "Material Group"
                                WB_EEPLUS.Workbook.Worksheets("SubVariants_SS").Cells(2, ExcelColumn + 1).Value = "ServiceLine"

                                WB_EEPLUS.Workbook.Worksheets("SubVariants_SS").Cells(2, ExcelColumn).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                                WB_EEPLUS.Workbook.Worksheets("SubVariants_SS").Cells(2, ExcelColumn).Style.Fill.BackgroundColor.SetColor(Color.SkyBlue)
                                WB_EEPLUS.Workbook.Worksheets("SubVariants_SS").Cells(2, ExcelColumn).Style.Font.Color.SetColor(Color.Black)
                                ExcelColumn += 1
                        End Select
                        ExcelColumn += 1

                    End If
                End If
                ExcelColumn += 1
            Next

            WB_EEPLUS.Save()
            Export_Variant_SS = True
        Catch ex As Exception
            Export_Variant_SS = True
        End Try
    End Function

    Private Function Export_Variant_Direct(ByVal FilePath) As Boolean
        Try

            'Data Source
            Dim DT As New DataTable
            DT = SQL_F.GetDataTable(CS, "Select * From Report_Americas_Direct")

            Try
                Kill(FilePath)
            Catch ex As Exception
            End Try

            'Create Excel File
            Dim WB_EEPLUS As New ExcelPackage(New IO.FileInfo(FilePath))
            WB_EEPLUS.Workbook.Worksheets.Add("Variant_Direct")
            WB_EEPLUS.Workbook.Worksheets("Variant_Direct").Cells("A1").LoadFromDataTable(DT, True)

            Dim ColumIndex As Integer = 0
            For Each C As DataColumn In DT.Columns
                Try
                    If C.DataType.FullName = "System.DateTime" Then
                        WB_EEPLUS.Workbook.Worksheets("Variant_Direct").Column(ColumIndex + 1).Style.Numberformat.Format = "mm-dd-yy"
                    End If
                    ColumIndex += 1
                Catch ex As Exception
                End Try
            Next

            'Set Header Color
            For I = 1 To DT.Columns.Count
                WB_EEPLUS.Workbook.Worksheets("Variant_Direct").Cells(1, I).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                WB_EEPLUS.Workbook.Worksheets("Variant_Direct").Cells(1, I).Style.Fill.BackgroundColor.SetColor(Color.MidnightBlue)
                WB_EEPLUS.Workbook.Worksheets("Variant_Direct").Cells(1, I).Style.Font.Color.SetColor(Color.White)
            Next

            'Color Coding for Change Type
            Dim Row As Integer = 2
            For Each R As DataRow In DT.Rows

                Select Case R("Change_Type")

                    Case "New Addition"
                        For I = 1 To DT.Columns.Count
                            WB_EEPLUS.Workbook.Worksheets("Variant_Direct").Cells(Row, I).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                            WB_EEPLUS.Workbook.Worksheets("Variant_Direct").Cells(Row, I).Style.Fill.BackgroundColor.SetColor(Color.Green)
                            WB_EEPLUS.Workbook.Worksheets("Variant_Direct").Cells(Row, I).Style.Font.Color.SetColor(Color.White)
                        Next
                    Case "Request for Deletion"
                        For I = 1 To DT.Columns.Count
                            WB_EEPLUS.Workbook.Worksheets("Variant_Direct").Cells(Row, I).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                            WB_EEPLUS.Workbook.Worksheets("Variant_Direct").Cells(Row, I).Style.Fill.BackgroundColor.SetColor(Color.Red)
                            WB_EEPLUS.Workbook.Worksheets("Variant_Direct").Cells(Row, I).Style.Font.Color.SetColor(Color.Black)
                        Next
                    Case "Pending Owner Change"
                        For I = 1 To DT.Columns.Count
                            WB_EEPLUS.Workbook.Worksheets("Variant_Direct").Cells(Row, I).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                            WB_EEPLUS.Workbook.Worksheets("Variant_Direct").Cells(Row, I).Style.Fill.BackgroundColor.SetColor(Color.Orange)
                            WB_EEPLUS.Workbook.Worksheets("Variant_Direct").Cells(Row, I).Style.Font.Color.SetColor(Color.Black)
                        Next
                End Select
                Row += 1
            Next


            'Add Sub-Variants
            WB_EEPLUS.Workbook.Worksheets.Add("SubVariants_Direct")

            Dim DT_SubVariants As DataTable
            DT_SubVariants = SQL_F.GetDataTable(CS, "Select * From Var_Americas_Direct_Variants order by Variant")

            Dim ExcelColumn As Integer = 1

            For Each R As DataRow In DT_SubVariants.Rows

                Dim Variant_Number As Object
                Variant_Number = R("Variant")

                If Variant_Number < 10 Then
                    Variant_Number = "0" & Variant_Number
                Else
                    Variant_Number = Variant_Number
                End If

                ''If Variant_Number <> "01" Then
                Dim DT_Variant As New DataTable
                DT_Variant = SQL_F.GetDataTable(CS, "Select * From Var_Americas_Direct_Var" & Variant_Number)

                If DT_Variant.Rows.Count > 0 Then
                    WB_EEPLUS.Workbook.Worksheets("SubVariants_Direct").Cells(1, ExcelColumn).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                    WB_EEPLUS.Workbook.Worksheets("SubVariants_Direct").Cells(1, ExcelColumn).Style.Fill.BackgroundColor.SetColor(Color.MidnightBlue)
                    WB_EEPLUS.Workbook.Worksheets("SubVariants_Direct").Cells(1, ExcelColumn).Style.Font.Color.SetColor(Color.White)

                    WB_EEPLUS.Workbook.Worksheets("SubVariants_Direct").Cells(1, ExcelColumn).Value = R("Description")
                    WB_EEPLUS.Workbook.Worksheets("SubVariants_Direct").Cells(3, ExcelColumn).LoadFromDataTable(DT_Variant, False)

                    Select Case Variant_Number
                        Case "01"
                            WB_EEPLUS.Workbook.Worksheets("SubVariants_Direct").Cells(2, ExcelColumn).Value = "Vendor"
                            WB_EEPLUS.Workbook.Worksheets("SubVariants_Direct").Cells(2, ExcelColumn).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                            WB_EEPLUS.Workbook.Worksheets("SubVariants_Direct").Cells(2, ExcelColumn).Style.Fill.BackgroundColor.SetColor(Color.SkyBlue)
                            WB_EEPLUS.Workbook.Worksheets("SubVariants_Direct").Cells(2, ExcelColumn).Style.Font.Color.SetColor(Color.Black)

                            WB_EEPLUS.Workbook.Worksheets("SubVariants_Direct").Cells(2, ExcelColumn + 1).Value = "Vendor  Name"
                            WB_EEPLUS.Workbook.Worksheets("SubVariants_Direct").Cells(2, ExcelColumn + 1).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                            WB_EEPLUS.Workbook.Worksheets("SubVariants_Direct").Cells(2, ExcelColumn + 1).Style.Fill.BackgroundColor.SetColor(Color.SkyBlue)
                            WB_EEPLUS.Workbook.Worksheets("SubVariants_Direct").Cells(2, ExcelColumn + 1).Style.Font.Color.SetColor(Color.Black)

                            WB_EEPLUS.Workbook.Worksheets("SubVariants_Direct").Cells(2, ExcelColumn + 2).Value = "Scope"
                            WB_EEPLUS.Workbook.Worksheets("SubVariants_Direct").Cells(2, ExcelColumn + 2).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                            WB_EEPLUS.Workbook.Worksheets("SubVariants_Direct").Cells(2, ExcelColumn + 2).Style.Fill.BackgroundColor.SetColor(Color.SkyBlue)
                            WB_EEPLUS.Workbook.Worksheets("SubVariants_Direct").Cells(2, ExcelColumn + 2).Style.Font.Color.SetColor(Color.Black)

                    End Select
                    ExcelColumn += 1

                End If
                ''End If

            Next

            WB_EEPLUS.Save()
            Export_Variant_Direct = True
        Catch ex As Exception
            Export_Variant_Direct = False
        End Try
    End Function

    Private Function Export_Variant_Custom(ByVal FilePath) As Boolean
        Try

            'Data Source
            Dim DT As New DataTable
            DT = SQL_F.GetDataTable(CS, "Select * From Report_Americas_Custom")

            Try
                Kill(FilePath)
            Catch ex As Exception
            End Try

            'Create Excel File
            Dim WB_EEPLUS As New ExcelPackage(New IO.FileInfo(FilePath))
            WB_EEPLUS.Workbook.Worksheets.Add("Variant_Custom")
            WB_EEPLUS.Workbook.Worksheets("Variant_Custom").Cells("A1").LoadFromDataTable(DT, True)

            Dim ColumIndex As Integer = 0
            For Each C As DataColumn In DT.Columns
                Try
                    If C.DataType.FullName = "System.DateTime" Then
                        WB_EEPLUS.Workbook.Worksheets("Variant_Custom").Column(ColumIndex + 1).Style.Numberformat.Format = "mm-dd-yy"
                    End If
                    ColumIndex += 1
                Catch ex As Exception
                End Try
            Next

            'Set Header Color
            For I = 1 To DT.Columns.Count
                WB_EEPLUS.Workbook.Worksheets("Variant_Custom").Cells(1, I).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                WB_EEPLUS.Workbook.Worksheets("Variant_Custom").Cells(1, I).Style.Fill.BackgroundColor.SetColor(Color.MidnightBlue)
                WB_EEPLUS.Workbook.Worksheets("Variant_Custom").Cells(1, I).Style.Font.Color.SetColor(Color.White)
            Next

            'Color Coding for Change Type
            Dim Row As Integer = 2
            For Each R As DataRow In DT.Rows

                Select Case R("Change_Type")

                    Case "New Addition"
                        For I = 1 To DT.Columns.Count
                            WB_EEPLUS.Workbook.Worksheets("Variant_Custom").Cells(Row, I).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                            WB_EEPLUS.Workbook.Worksheets("Variant_Custom").Cells(Row, I).Style.Fill.BackgroundColor.SetColor(Color.Green)
                            WB_EEPLUS.Workbook.Worksheets("Variant_Custom").Cells(Row, I).Style.Font.Color.SetColor(Color.White)
                        Next
                    Case "Request for Deletion"
                        For I = 1 To DT.Columns.Count
                            WB_EEPLUS.Workbook.Worksheets("Variant_Custom").Cells(Row, I).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                            WB_EEPLUS.Workbook.Worksheets("Variant_Custom").Cells(Row, I).Style.Fill.BackgroundColor.SetColor(Color.Red)
                            WB_EEPLUS.Workbook.Worksheets("Variant_Custom").Cells(Row, I).Style.Font.Color.SetColor(Color.Black)
                        Next
                    Case "Pending Owner Change"
                        For I = 1 To DT.Columns.Count
                            WB_EEPLUS.Workbook.Worksheets("Variant_Custom").Cells(Row, I).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                            WB_EEPLUS.Workbook.Worksheets("Variant_Custom").Cells(Row, I).Style.Fill.BackgroundColor.SetColor(Color.Orange)
                            WB_EEPLUS.Workbook.Worksheets("Variant_Custom").Cells(Row, I).Style.Font.Color.SetColor(Color.Black)
                        Next
                End Select
                Row += 1
            Next


            'Add Sub-Variants
            WB_EEPLUS.Workbook.Worksheets.Add("SubVariants_Custom")

            Dim DT_SubVariants As DataTable
            DT_SubVariants = SQL_F.GetDataTable(CS, "Select * From Var_Americas_Custom_Variants order by Variant")

            Dim ExcelColumn As Integer = 1

            For Each R As DataRow In DT_SubVariants.Rows

                Dim Variant_Number As Object
                Variant_Number = R("Variant")

                If Variant_Number < 10 Then
                    Variant_Number = "0" & Variant_Number
                Else
                    Variant_Number = Variant_Number
                End If

                If Variant_Number <> "01" Then
                    Dim DT_Variant As New DataTable
                    DT_Variant = SQL_F.GetDataTable(CS, "Select * From Var_Americas_SS_Var" & Variant_Number)

                    If DT_Variant.Rows.Count > 0 Then
                        WB_EEPLUS.Workbook.Worksheets("SubVariants_Custom").Cells(1, ExcelColumn).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                        WB_EEPLUS.Workbook.Worksheets("SubVariants_Custom").Cells(1, ExcelColumn).Style.Fill.BackgroundColor.SetColor(Color.MidnightBlue)
                        WB_EEPLUS.Workbook.Worksheets("SubVariants_Custom").Cells(1, ExcelColumn).Style.Font.Color.SetColor(Color.White)

                        WB_EEPLUS.Workbook.Worksheets("SubVariants_Custom").Cells(1, ExcelColumn).Value = R("Description")
                        WB_EEPLUS.Workbook.Worksheets("SubVariants_Custom").Cells(3, ExcelColumn).LoadFromDataTable(DT_Variant, False)

                        Select Case Variant_Number
                            Case "11"
                                WB_EEPLUS.Workbook.Worksheets("SubVariants_Custom").Cells(2, ExcelColumn).Value = "Material Group"

                                WB_EEPLUS.Workbook.Worksheets("SubVariants_Custom").Cells(2, ExcelColumn).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                                WB_EEPLUS.Workbook.Worksheets("SubVariants_Custom").Cells(2, ExcelColumn).Style.Fill.BackgroundColor.SetColor(Color.SkyBlue)
                                WB_EEPLUS.Workbook.Worksheets("SubVariants_Custom").Cells(2, ExcelColumn).Style.Font.Color.SetColor(Color.Black)

                                'Case "09"
                                '    WB_EEPLUS.Workbook.Worksheets("SubVariants_Custom").Cells(2, ExcelColumn).Value = "PGrp"

                                '    WB_EEPLUS.Workbook.Worksheets("SubVariants_Custom").Cells(2, ExcelColumn).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                                '    WB_EEPLUS.Workbook.Worksheets("SubVariants_Custom").Cells(2, ExcelColumn).Style.Fill.BackgroundColor.SetColor(Color.SkyBlue)
                                '    WB_EEPLUS.Workbook.Worksheets("SubVariants_Custom").Cells(2, ExcelColumn).Style.Font.Color.SetColor(Color.Black)
                        End Select
                        ExcelColumn += 1

                    End If
                End If
                ExcelColumn += 1
            Next

            WB_EEPLUS.Save()
            Export_Variant_Custom = True
        Catch ex As Exception
            Export_Variant_Custom = False
        End Try
    End Function

    Private Function Export_Variant_LogIndNAPD(ByVal FilePath) As Boolean
        Try

            'Data Source
            Dim DT As New DataTable
            DT = SQL_F.GetDataTable(CS, "Select * From Report_Americas_LogIndNAPD")

            Try
                Kill(FilePath)
            Catch ex As Exception
            End Try

            'Create Excel File
            Dim WB_EEPLUS As New ExcelPackage(New IO.FileInfo(FilePath))
            WB_EEPLUS.Workbook.Worksheets.Add("Variant_LogIndNAPD")
            WB_EEPLUS.Workbook.Worksheets("Variant_LogIndNAPD").Cells("A1").LoadFromDataTable(DT, True)

            Dim ColumIndex As Integer = 0
            For Each C As DataColumn In DT.Columns
                Try
                    If C.DataType.FullName = "System.DateTime" Then
                        WB_EEPLUS.Workbook.Worksheets("Variant_LogIndNAPD").Column(ColumIndex + 1).Style.Numberformat.Format = "mm-dd-yy"
                    End If
                    ColumIndex += 1
                Catch ex As Exception
                End Try
            Next

            'Set Header Color
            For I = 1 To DT.Columns.Count
                WB_EEPLUS.Workbook.Worksheets("Variant_LogIndNAPD").Cells(1, I).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                WB_EEPLUS.Workbook.Worksheets("Variant_LogIndNAPD").Cells(1, I).Style.Fill.BackgroundColor.SetColor(Color.MidnightBlue)
                WB_EEPLUS.Workbook.Worksheets("Variant_LogIndNAPD").Cells(1, I).Style.Font.Color.SetColor(Color.White)
            Next

            'Color Coding for Change Type
            Dim Row As Integer = 2
            For Each R As DataRow In DT.Rows

                Select Case R("Change_Type")

                    Case "New Addition"
                        For I = 1 To DT.Columns.Count
                            WB_EEPLUS.Workbook.Worksheets("Variant_LogIndNAPD").Cells(Row, I).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                            WB_EEPLUS.Workbook.Worksheets("Variant_LogIndNAPD").Cells(Row, I).Style.Fill.BackgroundColor.SetColor(Color.Green)
                            WB_EEPLUS.Workbook.Worksheets("Variant_LogIndNAPD").Cells(Row, I).Style.Font.Color.SetColor(Color.White)
                        Next
                    Case "Request for Deletion"
                        For I = 1 To DT.Columns.Count
                            WB_EEPLUS.Workbook.Worksheets("Variant_LogIndNAPD").Cells(Row, I).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                            WB_EEPLUS.Workbook.Worksheets("Variant_LogIndNAPD").Cells(Row, I).Style.Fill.BackgroundColor.SetColor(Color.Red)
                            WB_EEPLUS.Workbook.Worksheets("Variant_LogIndNAPD").Cells(Row, I).Style.Font.Color.SetColor(Color.Black)
                        Next
                    Case "Pending Owner Change"
                        For I = 1 To DT.Columns.Count
                            WB_EEPLUS.Workbook.Worksheets("Variant_LogIndNAPD").Cells(Row, I).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                            WB_EEPLUS.Workbook.Worksheets("Variant_LogIndNAPD").Cells(Row, I).Style.Fill.BackgroundColor.SetColor(Color.Orange)
                            WB_EEPLUS.Workbook.Worksheets("Variant_LogIndNAPD").Cells(Row, I).Style.Font.Color.SetColor(Color.Black)
                        Next
                End Select
                Row += 1
            Next

            ''Add Sub-Variants
            'WB_EEPLUS.Workbook.Worksheets.Add("SubVariants_LogIndNAPD")

            'Dim DT_SubVariants As DataTable
            'DT_SubVariants = SQL_F.GetDataTable(CS, "Select * From Var_Americas_LogIndNAPD_Variants order by Variant")

            'Dim ExcelColumn As Integer = 1

            'For Each R As DataRow In DT_SubVariants.Rows

            '    Dim Variant_Number As Object
            '    Variant_Number = R("Variant")

            '    If Variant_Number < 10 Then
            '        Variant_Number = "0" & Variant_Number
            '    Else
            '        Variant_Number = Variant_Number
            '    End If

            '    If Variant_Number <> "01" Then
            '        Dim DT_Variant As New DataTable
            '        DT_Variant = SQL_F.GetDataTable(CS, "Select * From Var_Americas_SS_Var" & Variant_Number)

            '        If DT_Variant.Rows.Count > 0 Then
            '            WB_EEPLUS.Workbook.Worksheets("SubVariants_LogIndNAPD").Cells(1, ExcelColumn).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
            '            WB_EEPLUS.Workbook.Worksheets("SubVariants_LogIndNAPD").Cells(1, ExcelColumn).Style.Fill.BackgroundColor.SetColor(Color.MidnightBlue)
            '            WB_EEPLUS.Workbook.Worksheets("SubVariants_LogIndNAPD").Cells(1, ExcelColumn).Style.Font.Color.SetColor(Color.White)

            '            WB_EEPLUS.Workbook.Worksheets("SubVariants_LogIndNAPD").Cells(1, ExcelColumn).Value = R("Description")
            '            WB_EEPLUS.Workbook.Worksheets("SubVariants_LogIndNAPD").Cells(3, ExcelColumn).LoadFromDataTable(DT_Variant, False)

            '            Select Case Variant_Number
            '                Case "10"
            '                    WB_EEPLUS.Workbook.Worksheets("SubVariants_LogIndNAPD").Cells(2, ExcelColumn).Value = "Material Group"

            '                    WB_EEPLUS.Workbook.Worksheets("SubVariants_LogIndNAPD").Cells(2, ExcelColumn).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
            '                    WB_EEPLUS.Workbook.Worksheets("SubVariants_LogIndNAPD").Cells(2, ExcelColumn).Style.Fill.BackgroundColor.SetColor(Color.SkyBlue)
            '                    WB_EEPLUS.Workbook.Worksheets("SubVariants_LogIndNAPD").Cells(2, ExcelColumn).Style.Font.Color.SetColor(Color.Black)

            '            End Select
            '            ExcelColumn += 1

            '        End If
            '    End If
            '    ExcelColumn += 1
            'Next

            WB_EEPLUS.Save()
            Export_Variant_LogIndNAPD = True
        Catch ex As Exception
            Export_Variant_LogIndNAPD = False
        End Try
    End Function

    Private Function Export_Variant_LogTMS(ByVal FilePath) As Boolean
        Try

            'Data Source
            Dim DT As New DataTable
            DT = SQL_F.GetDataTable(CS, "Select * From Report_Americas_LogTMS")

            Try
                Kill(FilePath)
            Catch ex As Exception
            End Try

            'Create Excel File
            Dim WB_EEPLUS As New ExcelPackage(New IO.FileInfo(FilePath))
            WB_EEPLUS.Workbook.Worksheets.Add("Variant_LogTMS")
            WB_EEPLUS.Workbook.Worksheets("Variant_LogTMS").Cells("A1").LoadFromDataTable(DT, True)

            Dim ColumIndex As Integer = 0
            For Each C As DataColumn In DT.Columns
                Try
                    If C.DataType.FullName = "System.DateTime" Then
                        WB_EEPLUS.Workbook.Worksheets("Variant_LogTMS").Column(ColumIndex + 1).Style.Numberformat.Format = "mm-dd-yy"
                    End If
                    ColumIndex += 1
                Catch ex As Exception
                End Try
            Next

            'Set Header Color
            For I = 1 To DT.Columns.Count
                WB_EEPLUS.Workbook.Worksheets("Variant_LogTMS").Cells(1, I).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                WB_EEPLUS.Workbook.Worksheets("Variant_LogTMS").Cells(1, I).Style.Fill.BackgroundColor.SetColor(Color.MidnightBlue)
                WB_EEPLUS.Workbook.Worksheets("Variant_LogTMS").Cells(1, I).Style.Font.Color.SetColor(Color.White)
            Next

            'Color Coding for Change Type
            Dim Row As Integer = 2
            For Each R As DataRow In DT.Rows

                Select Case R("Change_Type")

                    Case "New Addition"
                        For I = 1 To DT.Columns.Count
                            WB_EEPLUS.Workbook.Worksheets("Variant_LogTMS").Cells(Row, I).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                            WB_EEPLUS.Workbook.Worksheets("Variant_LogTMS").Cells(Row, I).Style.Fill.BackgroundColor.SetColor(Color.Green)
                            WB_EEPLUS.Workbook.Worksheets("Variant_LogTMS").Cells(Row, I).Style.Font.Color.SetColor(Color.White)
                        Next
                    Case "Request for Deletion"
                        For I = 1 To DT.Columns.Count
                            WB_EEPLUS.Workbook.Worksheets("Variant_LogTMS").Cells(Row, I).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                            WB_EEPLUS.Workbook.Worksheets("Variant_LogTMS").Cells(Row, I).Style.Fill.BackgroundColor.SetColor(Color.Red)
                            WB_EEPLUS.Workbook.Worksheets("Variant_LogTMS").Cells(Row, I).Style.Font.Color.SetColor(Color.Black)
                        Next
                    Case "Pending Owner Change"
                        For I = 1 To DT.Columns.Count
                            WB_EEPLUS.Workbook.Worksheets("Variant_LogTMS").Cells(Row, I).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                            WB_EEPLUS.Workbook.Worksheets("Variant_LogTMS").Cells(Row, I).Style.Fill.BackgroundColor.SetColor(Color.Orange)
                            WB_EEPLUS.Workbook.Worksheets("Variant_LogTMS").Cells(Row, I).Style.Font.Color.SetColor(Color.Black)
                        Next
                End Select
                Row += 1
            Next

            WB_EEPLUS.Save()
            Export_Variant_LogTMS = True
        Catch ex As Exception
            Export_Variant_LogTMS = False
        End Try
    End Function

    Private Function Export_Variant_LogImpLA(ByVal FilePath) As Boolean
        Try

            'Data Source
            Dim DT As New DataTable
            DT = SQL_F.GetDataTable(CS, "Select * From Report_Americas_LogImpLA")

            Try
                Kill(FilePath)
            Catch ex As Exception
            End Try

            'Create Excel File
            Dim WB_EEPLUS As New ExcelPackage(New IO.FileInfo(FilePath))
            WB_EEPLUS.Workbook.Worksheets.Add("Variant_LogImpLA")
            WB_EEPLUS.Workbook.Worksheets("Variant_LogImpLA").Cells("A1").LoadFromDataTable(DT, True)

            Dim ColumIndex As Integer = 0
            For Each C As DataColumn In DT.Columns
                Try
                    If C.DataType.FullName = "System.DateTime" Then
                        WB_EEPLUS.Workbook.Worksheets("Variant_LogImpLA").Column(ColumIndex + 1).Style.Numberformat.Format = "mm-dd-yy"
                    End If
                    ColumIndex += 1
                Catch ex As Exception
                End Try
            Next

            'Set Header Color
            For I = 1 To DT.Columns.Count
                WB_EEPLUS.Workbook.Worksheets("Variant_LogImpLA").Cells(1, I).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                WB_EEPLUS.Workbook.Worksheets("Variant_LogImpLA").Cells(1, I).Style.Fill.BackgroundColor.SetColor(Color.MidnightBlue)
                WB_EEPLUS.Workbook.Worksheets("Variant_LogImpLA").Cells(1, I).Style.Font.Color.SetColor(Color.White)
            Next

            'Color Coding for Change Type
            Dim Row As Integer = 2
            For Each R As DataRow In DT.Rows

                Select Case R("Change_Type")

                    Case "New Addition"
                        For I = 1 To DT.Columns.Count
                            WB_EEPLUS.Workbook.Worksheets("Variant_LogImpLA").Cells(Row, I).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                            WB_EEPLUS.Workbook.Worksheets("Variant_LogImpLA").Cells(Row, I).Style.Fill.BackgroundColor.SetColor(Color.Green)
                            WB_EEPLUS.Workbook.Worksheets("Variant_LogImpLA").Cells(Row, I).Style.Font.Color.SetColor(Color.White)
                        Next
                    Case "Request for Deletion"
                        For I = 1 To DT.Columns.Count
                            WB_EEPLUS.Workbook.Worksheets("Variant_LogImpLA").Cells(Row, I).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                            WB_EEPLUS.Workbook.Worksheets("Variant_LogImpLA").Cells(Row, I).Style.Fill.BackgroundColor.SetColor(Color.Red)
                            WB_EEPLUS.Workbook.Worksheets("Variant_LogImpLA").Cells(Row, I).Style.Font.Color.SetColor(Color.Black)
                        Next
                    Case "Pending Owner Change"
                        For I = 1 To DT.Columns.Count
                            WB_EEPLUS.Workbook.Worksheets("Variant_LogImpLA").Cells(Row, I).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                            WB_EEPLUS.Workbook.Worksheets("Variant_LogImpLA").Cells(Row, I).Style.Fill.BackgroundColor.SetColor(Color.Orange)
                            WB_EEPLUS.Workbook.Worksheets("Variant_LogImpLA").Cells(Row, I).Style.Font.Color.SetColor(Color.Black)
                        Next
                End Select
                Row += 1
            Next


            WB_EEPLUS.Save()
            Export_Variant_LogImpLA = True
        Catch ex As Exception
            Export_Variant_LogImpLA = False
        End Try
    End Function

    Private Function Export_Variant_LogImpNA(ByVal FilePath) As Boolean
        Try

            'Data Source
            Dim DT As New DataTable
            DT = SQL_F.GetDataTable(CS, "Select * From Report_Americas_LogImpNA")

            Try
                Kill(FilePath)
            Catch ex As Exception
            End Try

            'Create Excel File
            Dim WB_EEPLUS As New ExcelPackage(New IO.FileInfo(FilePath))
            WB_EEPLUS.Workbook.Worksheets.Add("Variant_LogImpNA")
            WB_EEPLUS.Workbook.Worksheets("Variant_LogImpNA").Cells("A1").LoadFromDataTable(DT, True)

            Dim ColumIndex As Integer = 0
            For Each C As DataColumn In DT.Columns
                Try
                    If C.DataType.FullName = "System.DateTime" Then
                        WB_EEPLUS.Workbook.Worksheets("Variant_LogImpNA").Column(ColumIndex + 1).Style.Numberformat.Format = "mm-dd-yy"
                    End If
                    ColumIndex += 1
                Catch ex As Exception
                End Try
            Next

            'Set Header Color
            For I = 1 To DT.Columns.Count
                WB_EEPLUS.Workbook.Worksheets("Variant_LogImpNA").Cells(1, I).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                WB_EEPLUS.Workbook.Worksheets("Variant_LogImpNA").Cells(1, I).Style.Fill.BackgroundColor.SetColor(Color.MidnightBlue)
                WB_EEPLUS.Workbook.Worksheets("Variant_LogImpNA").Cells(1, I).Style.Font.Color.SetColor(Color.White)
            Next


            'Color Coding for Change Type
            Dim Row As Integer = 2
            For Each R As DataRow In DT.Rows

                Select Case R("Change_Type")

                    Case "New Addition"
                        For I = 1 To DT.Columns.Count
                            WB_EEPLUS.Workbook.Worksheets("Variant_LogImpNA").Cells(Row, I).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                            WB_EEPLUS.Workbook.Worksheets("Variant_LogImpNA").Cells(Row, I).Style.Fill.BackgroundColor.SetColor(Color.Green)
                            WB_EEPLUS.Workbook.Worksheets("Variant_LogImpNA").Cells(Row, I).Style.Font.Color.SetColor(Color.White)
                        Next
                    Case "Request for Deletion"
                        For I = 1 To DT.Columns.Count
                            WB_EEPLUS.Workbook.Worksheets("Variant_LogImpNA").Cells(Row, I).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                            WB_EEPLUS.Workbook.Worksheets("Variant_LogImpNA").Cells(Row, I).Style.Fill.BackgroundColor.SetColor(Color.Red)
                            WB_EEPLUS.Workbook.Worksheets("Variant_LogImpNA").Cells(Row, I).Style.Font.Color.SetColor(Color.Black)
                        Next
                    Case "Pending Owner Change"
                        For I = 1 To DT.Columns.Count
                            WB_EEPLUS.Workbook.Worksheets("Variant_LogImpNA").Cells(Row, I).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                            WB_EEPLUS.Workbook.Worksheets("Variant_LogImpNA").Cells(Row, I).Style.Fill.BackgroundColor.SetColor(Color.Orange)
                            WB_EEPLUS.Workbook.Worksheets("Variant_LogImpNA").Cells(Row, I).Style.Font.Color.SetColor(Color.Black)
                        Next
                End Select
                Row += 1
            Next


            WB_EEPLUS.Save()
            Export_Variant_LogImpNA = True
        Catch ex As Exception
            Export_Variant_LogImpNA = False
        End Try
    End Function

    Private Function Export_Variant_LogDirect(ByVal FilePath) As Boolean
        Try

            'Data Source
            Dim DT As New DataTable
            DT = SQL_F.GetDataTable(CS, "Select * From Report_Americas_LogDirect")

            Try
                Kill(FilePath)
            Catch ex As Exception
            End Try

            'Create Excel File
            Dim WB_EEPLUS As New ExcelPackage(New IO.FileInfo(FilePath))
            WB_EEPLUS.Workbook.Worksheets.Add("Variant_LogDirect")
            WB_EEPLUS.Workbook.Worksheets("Variant_LogDirect").Cells("A1").LoadFromDataTable(DT, True)

            Dim ColumIndex As Integer = 0
            For Each C As DataColumn In DT.Columns
                Try
                    If C.DataType.FullName = "System.DateTime" Then
                        WB_EEPLUS.Workbook.Worksheets("Variant_LogDirect").Column(ColumIndex + 1).Style.Numberformat.Format = "mm-dd-yy"
                    End If
                    ColumIndex += 1
                Catch ex As Exception
                End Try
            Next

            'Set Header Color
            For I = 1 To DT.Columns.Count
                WB_EEPLUS.Workbook.Worksheets("Variant_LogDirect").Cells(1, I).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                WB_EEPLUS.Workbook.Worksheets("Variant_LogDirect").Cells(1, I).Style.Fill.BackgroundColor.SetColor(Color.MidnightBlue)
                WB_EEPLUS.Workbook.Worksheets("Variant_LogDirect").Cells(1, I).Style.Font.Color.SetColor(Color.White)
            Next

            'Color Coding for Change Type
            Dim Row As Integer = 2
            For Each R As DataRow In DT.Rows

                Select Case R("Change_Type")

                    Case "New Addition"
                        For I = 1 To DT.Columns.Count
                            WB_EEPLUS.Workbook.Worksheets("Variant_LogDirect").Cells(Row, I).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                            WB_EEPLUS.Workbook.Worksheets("Variant_LogDirect").Cells(Row, I).Style.Fill.BackgroundColor.SetColor(Color.Green)
                            WB_EEPLUS.Workbook.Worksheets("Variant_LogDirect").Cells(Row, I).Style.Font.Color.SetColor(Color.White)
                        Next
                    Case "Request for Deletion"
                        For I = 1 To DT.Columns.Count
                            WB_EEPLUS.Workbook.Worksheets("Variant_LogDirect").Cells(Row, I).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                            WB_EEPLUS.Workbook.Worksheets("Variant_LogDirect").Cells(Row, I).Style.Fill.BackgroundColor.SetColor(Color.Red)
                            WB_EEPLUS.Workbook.Worksheets("Variant_LogDirect").Cells(Row, I).Style.Font.Color.SetColor(Color.Black)
                        Next
                    Case "Pending Owner Change"
                        For I = 1 To DT.Columns.Count
                            WB_EEPLUS.Workbook.Worksheets("Variant_LogDirect").Cells(Row, I).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                            WB_EEPLUS.Workbook.Worksheets("Variant_LogDirect").Cells(Row, I).Style.Fill.BackgroundColor.SetColor(Color.Orange)
                            WB_EEPLUS.Workbook.Worksheets("Variant_LogDirect").Cells(Row, I).Style.Font.Color.SetColor(Color.Black)
                        Next
                End Select
                Row += 1
            Next


            WB_EEPLUS.Save()
            Export_Variant_LogDirect = True
        Catch ex As Exception
            Export_Variant_LogDirect = False
        End Try
    End Function

    Private Function Export_AllVariants(ByVal FilePath) As Boolean
        Try


            Try
                Kill(FilePath)
            Catch ex As Exception
            End Try

            'Create Excel File
            Dim WB_EEPLUS As New ExcelPackage(New IO.FileInfo(FilePath))

            'STR
            Try

                'Data Source
                Dim DT As New DataTable
                DT = SQL_F.GetDataTable(CS, "Select * From Report_Americas_STR")

                'Create Excel File
                WB_EEPLUS.Workbook.Worksheets.Add("Variant_STR")
                WB_EEPLUS.Workbook.Worksheets("Variant_STR").Cells("A1").LoadFromDataTable(DT, True)

                Dim ColumIndex As Integer = 0
                For Each C As DataColumn In DT.Columns
                    Try
                        If C.DataType.FullName = "System.DateTime" Then
                            WB_EEPLUS.Workbook.Worksheets("Variant_STR").Column(ColumIndex + 1).Style.Numberformat.Format = "mm-dd-yy"
                        End If
                        ColumIndex += 1
                    Catch ex As Exception
                    End Try
                Next

                'Set Header Color
                For I = 1 To DT.Columns.Count
                    WB_EEPLUS.Workbook.Worksheets("Variant_STR").Cells(1, I).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                    WB_EEPLUS.Workbook.Worksheets("Variant_STR").Cells(1, I).Style.Fill.BackgroundColor.SetColor(Color.MidnightBlue)
                    WB_EEPLUS.Workbook.Worksheets("Variant_STR").Cells(1, I).Style.Font.Color.SetColor(Color.White)
                Next

                'Color Coding for Change Type
                Dim Row As Integer = 2
                For Each R As DataRow In DT.Rows

                    Select Case R("Change_Type")

                        Case "New Addition"
                            For I = 1 To DT.Columns.Count
                                WB_EEPLUS.Workbook.Worksheets("Variant_STR").Cells(Row, I).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                                WB_EEPLUS.Workbook.Worksheets("Variant_STR").Cells(Row, I).Style.Fill.BackgroundColor.SetColor(Color.Green)
                                WB_EEPLUS.Workbook.Worksheets("Variant_STR").Cells(Row, I).Style.Font.Color.SetColor(Color.White)
                            Next
                        Case "Request for Deletion"
                            For I = 1 To DT.Columns.Count
                                WB_EEPLUS.Workbook.Worksheets("Variant_STR").Cells(Row, I).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                                WB_EEPLUS.Workbook.Worksheets("Variant_STR").Cells(Row, I).Style.Fill.BackgroundColor.SetColor(Color.Red)
                                WB_EEPLUS.Workbook.Worksheets("Variant_STR").Cells(Row, I).Style.Font.Color.SetColor(Color.Black)
                            Next
                        Case "Pending Owner Change"
                            For I = 1 To DT.Columns.Count
                                WB_EEPLUS.Workbook.Worksheets("Variant_STR").Cells(Row, I).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                                WB_EEPLUS.Workbook.Worksheets("Variant_STR").Cells(Row, I).Style.Fill.BackgroundColor.SetColor(Color.Orange)
                                WB_EEPLUS.Workbook.Worksheets("Variant_STR").Cells(Row, I).Style.Font.Color.SetColor(Color.Black)
                            Next
                    End Select
                    Row += 1
                Next
            Catch ex As Exception
            End Try

            'SS
            Try

                'Data Source
                Dim DT As New DataTable
                DT = SQL_F.GetDataTable(CS, "Select * From Report_Americas_SS")

                'Create Excel File
                WB_EEPLUS.Workbook.Worksheets.Add("Variant_SS")
                WB_EEPLUS.Workbook.Worksheets("Variant_SS").Cells("A1").LoadFromDataTable(DT, True)

                Dim ColumIndex As Integer = 0
                For Each C As DataColumn In DT.Columns
                    Try
                        If C.DataType.FullName = "System.DateTime" Then
                            WB_EEPLUS.Workbook.Worksheets("Variant_SS").Column(ColumIndex + 1).Style.Numberformat.Format = "mm-dd-yy"
                        End If
                        ColumIndex += 1
                    Catch ex As Exception
                    End Try
                Next

                'Set Header Color
                For I = 1 To DT.Columns.Count
                    WB_EEPLUS.Workbook.Worksheets("Variant_SS").Cells(1, I).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                    WB_EEPLUS.Workbook.Worksheets("Variant_SS").Cells(1, I).Style.Fill.BackgroundColor.SetColor(Color.MidnightBlue)
                    WB_EEPLUS.Workbook.Worksheets("Variant_SS").Cells(1, I).Style.Font.Color.SetColor(Color.White)
                Next

                'Color Coding for Change Type
                Dim Row As Integer = 2
                For Each R As DataRow In DT.Rows

                    Select Case R("Change_Type")

                        Case "New Addition"
                            For I = 1 To DT.Columns.Count
                                WB_EEPLUS.Workbook.Worksheets("Variant_SS").Cells(Row, I).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                                WB_EEPLUS.Workbook.Worksheets("Variant_SS").Cells(Row, I).Style.Fill.BackgroundColor.SetColor(Color.Green)
                                WB_EEPLUS.Workbook.Worksheets("Variant_SS").Cells(Row, I).Style.Font.Color.SetColor(Color.White)
                            Next
                        Case "Request for Deletion"
                            For I = 1 To DT.Columns.Count
                                WB_EEPLUS.Workbook.Worksheets("Variant_SS").Cells(Row, I).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                                WB_EEPLUS.Workbook.Worksheets("Variant_SS").Cells(Row, I).Style.Fill.BackgroundColor.SetColor(Color.Red)
                                WB_EEPLUS.Workbook.Worksheets("Variant_SS").Cells(Row, I).Style.Font.Color.SetColor(Color.Black)
                            Next
                        Case "Pending Owner Change"
                            For I = 1 To DT.Columns.Count
                                WB_EEPLUS.Workbook.Worksheets("Variant_SS").Cells(Row, I).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                                WB_EEPLUS.Workbook.Worksheets("Variant_SS").Cells(Row, I).Style.Fill.BackgroundColor.SetColor(Color.Orange)
                                WB_EEPLUS.Workbook.Worksheets("Variant_SS").Cells(Row, I).Style.Font.Color.SetColor(Color.Black)
                            Next
                    End Select
                    Row += 1
                Next


                'Add Sub-Variants
                WB_EEPLUS.Workbook.Worksheets.Add("SubVariants_SS")

                Dim DT_SubVariants As DataTable
                DT_SubVariants = SQL_F.GetDataTable(CS, "Select * From Var_Americas_SS_Variants order by Variant")

                Dim ExcelColumn As Integer = 1

                For Each R As DataRow In DT_SubVariants.Rows

                    Dim Variant_Number As Object
                    Variant_Number = R("Variant")

                    If Variant_Number < 10 Then
                        Variant_Number = "0" & Variant_Number
                    Else
                        Variant_Number = Variant_Number
                    End If

                    If Variant_Number <> "01" Then
                        Dim DT_Variant As New DataTable
                        DT_Variant = SQL_F.GetDataTable(CS, "Select * From Var_Americas_SS_Var" & Variant_Number)

                        If DT_Variant.Rows.Count > 0 Then
                            WB_EEPLUS.Workbook.Worksheets("SubVariants_SS").Cells(1, ExcelColumn).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                            WB_EEPLUS.Workbook.Worksheets("SubVariants_SS").Cells(1, ExcelColumn).Style.Fill.BackgroundColor.SetColor(Color.MidnightBlue)
                            WB_EEPLUS.Workbook.Worksheets("SubVariants_SS").Cells(1, ExcelColumn).Style.Font.Color.SetColor(Color.White)

                            WB_EEPLUS.Workbook.Worksheets("SubVariants_SS").Cells(1, ExcelColumn).Value = R("Description")
                            WB_EEPLUS.Workbook.Worksheets("SubVariants_SS").Cells(3, ExcelColumn).LoadFromDataTable(DT_Variant, False)

                            Select Case Variant_Number
                                Case "02", "03", "04", "05", "06", "07", "10", "11", "12", "13", "16"
                                    WB_EEPLUS.Workbook.Worksheets("SubVariants_SS").Cells(2, ExcelColumn).Value = "Material Group"

                                    WB_EEPLUS.Workbook.Worksheets("SubVariants_SS").Cells(2, ExcelColumn).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                                    WB_EEPLUS.Workbook.Worksheets("SubVariants_SS").Cells(2, ExcelColumn).Style.Fill.BackgroundColor.SetColor(Color.SkyBlue)
                                    WB_EEPLUS.Workbook.Worksheets("SubVariants_SS").Cells(2, ExcelColumn).Style.Font.Color.SetColor(Color.Black)

                                Case "08", "09", "14", "15"
                                    WB_EEPLUS.Workbook.Worksheets("SubVariants_SS").Cells(2, ExcelColumn).Value = "PGrp"

                                    WB_EEPLUS.Workbook.Worksheets("SubVariants_SS").Cells(2, ExcelColumn).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                                    WB_EEPLUS.Workbook.Worksheets("SubVariants_SS").Cells(2, ExcelColumn).Style.Fill.BackgroundColor.SetColor(Color.SkyBlue)
                                    WB_EEPLUS.Workbook.Worksheets("SubVariants_SS").Cells(2, ExcelColumn).Style.Font.Color.SetColor(Color.Black)
                                Case "17"
                                    WB_EEPLUS.Workbook.Worksheets("SubVariants_SS").Cells(2, ExcelColumn).Value = "Material Group"
                                    WB_EEPLUS.Workbook.Worksheets("SubVariants_SS").Cells(2, ExcelColumn + 1).Value = "ServiceLine"

                                    WB_EEPLUS.Workbook.Worksheets("SubVariants_SS").Cells(2, ExcelColumn).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                                    WB_EEPLUS.Workbook.Worksheets("SubVariants_SS").Cells(2, ExcelColumn).Style.Fill.BackgroundColor.SetColor(Color.SkyBlue)
                                    WB_EEPLUS.Workbook.Worksheets("SubVariants_SS").Cells(2, ExcelColumn).Style.Font.Color.SetColor(Color.Black)
                                    ExcelColumn += 1
                            End Select
                            ExcelColumn += 1

                        End If
                    End If
                    ExcelColumn += 1
                Next

            Catch ex As Exception
            End Try
            'Direct
            Try

                'Data Source
                Dim DT As New DataTable
                DT = SQL_F.GetDataTable(CS, "Select * From Report_Americas_Direct")

                WB_EEPLUS.Workbook.Worksheets.Add("Variant_Direct")
                WB_EEPLUS.Workbook.Worksheets("Variant_Direct").Cells("A1").LoadFromDataTable(DT, True)

                Dim ColumIndex As Integer = 0
                For Each C As DataColumn In DT.Columns
                    Try
                        If C.DataType.FullName = "System.DateTime" Then
                            WB_EEPLUS.Workbook.Worksheets("Variant_Direct").Column(ColumIndex + 1).Style.Numberformat.Format = "mm-dd-yy"
                        End If
                        ColumIndex += 1
                    Catch ex As Exception
                    End Try
                Next

                'Set Header Color
                For I = 1 To DT.Columns.Count
                    WB_EEPLUS.Workbook.Worksheets("Variant_Direct").Cells(1, I).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                    WB_EEPLUS.Workbook.Worksheets("Variant_Direct").Cells(1, I).Style.Fill.BackgroundColor.SetColor(Color.MidnightBlue)
                    WB_EEPLUS.Workbook.Worksheets("Variant_Direct").Cells(1, I).Style.Font.Color.SetColor(Color.White)
                Next

                'Color Coding for Change Type
                Dim Row As Integer = 2
                For Each R As DataRow In DT.Rows

                    Select Case R("Change_Type")

                        Case "New Addition"
                            For I = 1 To DT.Columns.Count
                                WB_EEPLUS.Workbook.Worksheets("Variant_Direct").Cells(Row, I).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                                WB_EEPLUS.Workbook.Worksheets("Variant_Direct").Cells(Row, I).Style.Fill.BackgroundColor.SetColor(Color.Green)
                                WB_EEPLUS.Workbook.Worksheets("Variant_Direct").Cells(Row, I).Style.Font.Color.SetColor(Color.White)
                            Next
                        Case "Request for Deletion"
                            For I = 1 To DT.Columns.Count
                                WB_EEPLUS.Workbook.Worksheets("Variant_Direct").Cells(Row, I).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                                WB_EEPLUS.Workbook.Worksheets("Variant_Direct").Cells(Row, I).Style.Fill.BackgroundColor.SetColor(Color.Red)
                                WB_EEPLUS.Workbook.Worksheets("Variant_Direct").Cells(Row, I).Style.Font.Color.SetColor(Color.Black)
                            Next
                        Case "Pending Owner Change"
                            For I = 1 To DT.Columns.Count
                                WB_EEPLUS.Workbook.Worksheets("Variant_Direct").Cells(Row, I).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                                WB_EEPLUS.Workbook.Worksheets("Variant_Direct").Cells(Row, I).Style.Fill.BackgroundColor.SetColor(Color.Orange)
                                WB_EEPLUS.Workbook.Worksheets("Variant_Direct").Cells(Row, I).Style.Font.Color.SetColor(Color.Black)
                            Next
                    End Select
                    Row += 1
                Next


                'Add Sub-Variants
                WB_EEPLUS.Workbook.Worksheets.Add("SubVariants_Direct")

                Dim DT_SubVariants As DataTable
                DT_SubVariants = SQL_F.GetDataTable(CS, "Select * From Var_Americas_Direct_Variants order by Variant")

                Dim ExcelColumn As Integer = 1

                For Each R As DataRow In DT_SubVariants.Rows

                    Dim Variant_Number As Object
                    Variant_Number = R("Variant")

                    If Variant_Number < 10 Then
                        Variant_Number = "0" & Variant_Number
                    Else
                        Variant_Number = Variant_Number
                    End If

                    ''If Variant_Number <> "01" Then
                    Dim DT_Variant As New DataTable
                    DT_Variant = SQL_F.GetDataTable(CS, "Select * From Var_Americas_Direct_Var" & Variant_Number)

                    If DT_Variant.Rows.Count > 0 Then
                        WB_EEPLUS.Workbook.Worksheets("SubVariants_Direct").Cells(1, ExcelColumn).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                        WB_EEPLUS.Workbook.Worksheets("SubVariants_Direct").Cells(1, ExcelColumn).Style.Fill.BackgroundColor.SetColor(Color.MidnightBlue)
                        WB_EEPLUS.Workbook.Worksheets("SubVariants_Direct").Cells(1, ExcelColumn).Style.Font.Color.SetColor(Color.White)

                        WB_EEPLUS.Workbook.Worksheets("SubVariants_Direct").Cells(1, ExcelColumn).Value = R("Description")
                        WB_EEPLUS.Workbook.Worksheets("SubVariants_Direct").Cells(3, ExcelColumn).LoadFromDataTable(DT_Variant, False)

                        Select Case Variant_Number
                            Case "01"
                                WB_EEPLUS.Workbook.Worksheets("SubVariants_Direct").Cells(2, ExcelColumn).Value = "Vendor"
                                WB_EEPLUS.Workbook.Worksheets("SubVariants_Direct").Cells(2, ExcelColumn).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                                WB_EEPLUS.Workbook.Worksheets("SubVariants_Direct").Cells(2, ExcelColumn).Style.Fill.BackgroundColor.SetColor(Color.SkyBlue)
                                WB_EEPLUS.Workbook.Worksheets("SubVariants_Direct").Cells(2, ExcelColumn).Style.Font.Color.SetColor(Color.Black)

                                WB_EEPLUS.Workbook.Worksheets("SubVariants_Direct").Cells(2, ExcelColumn + 1).Value = "Vendor  Name"
                                WB_EEPLUS.Workbook.Worksheets("SubVariants_Direct").Cells(2, ExcelColumn + 1).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                                WB_EEPLUS.Workbook.Worksheets("SubVariants_Direct").Cells(2, ExcelColumn + 1).Style.Fill.BackgroundColor.SetColor(Color.SkyBlue)
                                WB_EEPLUS.Workbook.Worksheets("SubVariants_Direct").Cells(2, ExcelColumn + 1).Style.Font.Color.SetColor(Color.Black)

                                WB_EEPLUS.Workbook.Worksheets("SubVariants_Direct").Cells(2, ExcelColumn + 2).Value = "Scope"
                                WB_EEPLUS.Workbook.Worksheets("SubVariants_Direct").Cells(2, ExcelColumn + 2).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                                WB_EEPLUS.Workbook.Worksheets("SubVariants_Direct").Cells(2, ExcelColumn + 2).Style.Fill.BackgroundColor.SetColor(Color.SkyBlue)
                                WB_EEPLUS.Workbook.Worksheets("SubVariants_Direct").Cells(2, ExcelColumn + 2).Style.Font.Color.SetColor(Color.Black)

                        End Select
                        ExcelColumn += 1

                    End If
                    ''End If

                Next

            Catch ex As Exception
            End Try
            'Customization
            Try

                'Data Source
                Dim DT As New DataTable
                DT = SQL_F.GetDataTable(CS, "Select * From Report_Americas_Custom")
                WB_EEPLUS.Workbook.Worksheets.Add("Variant_Custom")
                WB_EEPLUS.Workbook.Worksheets("Variant_Custom").Cells("A1").LoadFromDataTable(DT, True)

                Dim ColumIndex As Integer = 0
                For Each C As DataColumn In DT.Columns
                    Try
                        If C.DataType.FullName = "System.DateTime" Then
                            WB_EEPLUS.Workbook.Worksheets("Variant_Custom").Column(ColumIndex + 1).Style.Numberformat.Format = "mm-dd-yy"
                        End If
                        ColumIndex += 1
                    Catch ex As Exception
                    End Try
                Next

                'Set Header Color
                For I = 1 To DT.Columns.Count
                    WB_EEPLUS.Workbook.Worksheets("Variant_Custom").Cells(1, I).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                    WB_EEPLUS.Workbook.Worksheets("Variant_Custom").Cells(1, I).Style.Fill.BackgroundColor.SetColor(Color.MidnightBlue)
                    WB_EEPLUS.Workbook.Worksheets("Variant_Custom").Cells(1, I).Style.Font.Color.SetColor(Color.White)
                Next

                'Color Coding for Change Type
                Dim Row As Integer = 2
                For Each R As DataRow In DT.Rows

                    Select Case R("Change_Type")

                        Case "New Addition"
                            For I = 1 To DT.Columns.Count
                                WB_EEPLUS.Workbook.Worksheets("Variant_Custom").Cells(Row, I).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                                WB_EEPLUS.Workbook.Worksheets("Variant_Custom").Cells(Row, I).Style.Fill.BackgroundColor.SetColor(Color.Green)
                                WB_EEPLUS.Workbook.Worksheets("Variant_Custom").Cells(Row, I).Style.Font.Color.SetColor(Color.White)
                            Next
                        Case "Request for Deletion"
                            For I = 1 To DT.Columns.Count
                                WB_EEPLUS.Workbook.Worksheets("Variant_Custom").Cells(Row, I).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                                WB_EEPLUS.Workbook.Worksheets("Variant_Custom").Cells(Row, I).Style.Fill.BackgroundColor.SetColor(Color.Red)
                                WB_EEPLUS.Workbook.Worksheets("Variant_Custom").Cells(Row, I).Style.Font.Color.SetColor(Color.Black)
                            Next
                        Case "Pending Owner Change"
                            For I = 1 To DT.Columns.Count
                                WB_EEPLUS.Workbook.Worksheets("Variant_Custom").Cells(Row, I).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                                WB_EEPLUS.Workbook.Worksheets("Variant_Custom").Cells(Row, I).Style.Fill.BackgroundColor.SetColor(Color.Orange)
                                WB_EEPLUS.Workbook.Worksheets("Variant_Custom").Cells(Row, I).Style.Font.Color.SetColor(Color.Black)
                            Next
                    End Select
                    Row += 1
                Next


                'Add Sub-Variants
                WB_EEPLUS.Workbook.Worksheets.Add("SubVariants_Custom")

                Dim DT_SubVariants As DataTable
                DT_SubVariants = SQL_F.GetDataTable(CS, "Select * From Var_Americas_Custom_Variants order by Variant")

                Dim ExcelColumn As Integer = 1

                For Each R As DataRow In DT_SubVariants.Rows

                    Dim Variant_Number As Object
                    Variant_Number = R("Variant")

                    If Variant_Number < 10 Then
                        Variant_Number = "0" & Variant_Number
                    Else
                        Variant_Number = Variant_Number
                    End If

                    If Variant_Number <> "01" Then
                        Dim DT_Variant As New DataTable
                        DT_Variant = SQL_F.GetDataTable(CS, "Select * From Var_Americas_SS_Var" & Variant_Number)

                        If DT_Variant.Rows.Count > 0 Then
                            WB_EEPLUS.Workbook.Worksheets("SubVariants_Custom").Cells(1, ExcelColumn).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                            WB_EEPLUS.Workbook.Worksheets("SubVariants_Custom").Cells(1, ExcelColumn).Style.Fill.BackgroundColor.SetColor(Color.MidnightBlue)
                            WB_EEPLUS.Workbook.Worksheets("SubVariants_Custom").Cells(1, ExcelColumn).Style.Font.Color.SetColor(Color.White)

                            WB_EEPLUS.Workbook.Worksheets("SubVariants_Custom").Cells(1, ExcelColumn).Value = R("Description")
                            WB_EEPLUS.Workbook.Worksheets("SubVariants_Custom").Cells(3, ExcelColumn).LoadFromDataTable(DT_Variant, False)

                            Select Case Variant_Number
                                Case "11"
                                    WB_EEPLUS.Workbook.Worksheets("SubVariants_Custom").Cells(2, ExcelColumn).Value = "Material Group"

                                    WB_EEPLUS.Workbook.Worksheets("SubVariants_Custom").Cells(2, ExcelColumn).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                                    WB_EEPLUS.Workbook.Worksheets("SubVariants_Custom").Cells(2, ExcelColumn).Style.Fill.BackgroundColor.SetColor(Color.SkyBlue)
                                    WB_EEPLUS.Workbook.Worksheets("SubVariants_Custom").Cells(2, ExcelColumn).Style.Font.Color.SetColor(Color.Black)

                                    'Case "09"
                                    '    WB_EEPLUS.Workbook.Worksheets("SubVariants_Custom").Cells(2, ExcelColumn).Value = "PGrp"

                                    '    WB_EEPLUS.Workbook.Worksheets("SubVariants_Custom").Cells(2, ExcelColumn).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                                    '    WB_EEPLUS.Workbook.Worksheets("SubVariants_Custom").Cells(2, ExcelColumn).Style.Fill.BackgroundColor.SetColor(Color.SkyBlue)
                                    '    WB_EEPLUS.Workbook.Worksheets("SubVariants_Custom").Cells(2, ExcelColumn).Style.Font.Color.SetColor(Color.Black)
                            End Select
                            ExcelColumn += 1

                        End If
                    End If
                    ExcelColumn += 1
                Next
            Catch ex As Exception
            End Try
            'Log Ind NAPD
            Try

                'Data Source
                Dim DT As New DataTable
                DT = SQL_F.GetDataTable(CS, "Select * From Report_Americas_LogIndNAPD")
                WB_EEPLUS.Workbook.Worksheets.Add("Variant_LogIndNAPD")
                WB_EEPLUS.Workbook.Worksheets("Variant_LogIndNAPD").Cells("A1").LoadFromDataTable(DT, True)

                Dim ColumIndex As Integer = 0
                For Each C As DataColumn In DT.Columns
                    Try
                        If C.DataType.FullName = "System.DateTime" Then
                            WB_EEPLUS.Workbook.Worksheets("Variant_LogIndNAPD").Column(ColumIndex + 1).Style.Numberformat.Format = "mm-dd-yy"
                        End If
                        ColumIndex += 1
                    Catch ex As Exception
                    End Try
                Next

                'Set Header Color
                For I = 1 To DT.Columns.Count
                    WB_EEPLUS.Workbook.Worksheets("Variant_LogIndNAPD").Cells(1, I).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                    WB_EEPLUS.Workbook.Worksheets("Variant_LogIndNAPD").Cells(1, I).Style.Fill.BackgroundColor.SetColor(Color.MidnightBlue)
                    WB_EEPLUS.Workbook.Worksheets("Variant_LogIndNAPD").Cells(1, I).Style.Font.Color.SetColor(Color.White)
                Next

                'Color Coding for Change Type
                Dim Row As Integer = 2
                For Each R As DataRow In DT.Rows

                    Select Case R("Change_Type")

                        Case "New Addition"
                            For I = 1 To DT.Columns.Count
                                WB_EEPLUS.Workbook.Worksheets("Variant_LogIndNAPD").Cells(Row, I).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                                WB_EEPLUS.Workbook.Worksheets("Variant_LogIndNAPD").Cells(Row, I).Style.Fill.BackgroundColor.SetColor(Color.Green)
                                WB_EEPLUS.Workbook.Worksheets("Variant_LogIndNAPD").Cells(Row, I).Style.Font.Color.SetColor(Color.White)
                            Next
                        Case "Request for Deletion"
                            For I = 1 To DT.Columns.Count
                                WB_EEPLUS.Workbook.Worksheets("Variant_LogIndNAPD").Cells(Row, I).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                                WB_EEPLUS.Workbook.Worksheets("Variant_LogIndNAPD").Cells(Row, I).Style.Fill.BackgroundColor.SetColor(Color.Red)
                                WB_EEPLUS.Workbook.Worksheets("Variant_LogIndNAPD").Cells(Row, I).Style.Font.Color.SetColor(Color.Black)
                            Next
                        Case "Pending Owner Change"
                            For I = 1 To DT.Columns.Count
                                WB_EEPLUS.Workbook.Worksheets("Variant_LogIndNAPD").Cells(Row, I).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                                WB_EEPLUS.Workbook.Worksheets("Variant_LogIndNAPD").Cells(Row, I).Style.Fill.BackgroundColor.SetColor(Color.Orange)
                                WB_EEPLUS.Workbook.Worksheets("Variant_LogIndNAPD").Cells(Row, I).Style.Font.Color.SetColor(Color.Black)
                            Next
                    End Select
                    Row += 1
                Next

                ''Add Sub-Variants
                'WB_EEPLUS.Workbook.Worksheets.Add("SubVariants_LogIndNAPD")

                'Dim DT_SubVariants As DataTable
                'DT_SubVariants = SQL_F.GetDataTable(CS, "Select * From Var_Americas_LogIndNAPD_Variants order by Variant")

                'Dim ExcelColumn As Integer = 1

                'For Each R As DataRow In DT_SubVariants.Rows

                '    Dim Variant_Number As Object
                '    Variant_Number = R("Variant")

                '    If Variant_Number < 10 Then
                '        Variant_Number = "0" & Variant_Number
                '    Else
                '        Variant_Number = Variant_Number
                '    End If

                '    If Variant_Number <> "01" Then
                '        Dim DT_Variant As New DataTable
                '        DT_Variant = SQL_F.GetDataTable(CS, "Select * From Var_Americas_SS_Var" & Variant_Number)

                '        If DT_Variant.Rows.Count > 0 Then
                '            WB_EEPLUS.Workbook.Worksheets("SubVariants_LogIndNAPD").Cells(1, ExcelColumn).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                '            WB_EEPLUS.Workbook.Worksheets("SubVariants_LogIndNAPD").Cells(1, ExcelColumn).Style.Fill.BackgroundColor.SetColor(Color.MidnightBlue)
                '            WB_EEPLUS.Workbook.Worksheets("SubVariants_LogIndNAPD").Cells(1, ExcelColumn).Style.Font.Color.SetColor(Color.White)

                '            WB_EEPLUS.Workbook.Worksheets("SubVariants_LogIndNAPD").Cells(1, ExcelColumn).Value = R("Description")
                '            WB_EEPLUS.Workbook.Worksheets("SubVariants_LogIndNAPD").Cells(3, ExcelColumn).LoadFromDataTable(DT_Variant, False)

                '            Select Case Variant_Number
                '                Case "10"
                '                    WB_EEPLUS.Workbook.Worksheets("SubVariants_LogIndNAPD").Cells(2, ExcelColumn).Value = "Material Group"

                '                    WB_EEPLUS.Workbook.Worksheets("SubVariants_LogIndNAPD").Cells(2, ExcelColumn).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                '                    WB_EEPLUS.Workbook.Worksheets("SubVariants_LogIndNAPD").Cells(2, ExcelColumn).Style.Fill.BackgroundColor.SetColor(Color.SkyBlue)
                '                    WB_EEPLUS.Workbook.Worksheets("SubVariants_LogIndNAPD").Cells(2, ExcelColumn).Style.Font.Color.SetColor(Color.Black)

                '            End Select
                '            ExcelColumn += 1

                '        End If
                '    End If
                '    ExcelColumn += 1
                'Next

            Catch ex As Exception
            End Try
            'Log TMS
            Try

                'Data Source
                Dim DT As New DataTable
                DT = SQL_F.GetDataTable(CS, "Select * From Report_Americas_LogTMS")
                WB_EEPLUS.Workbook.Worksheets.Add("Variant_LogTMS")
                WB_EEPLUS.Workbook.Worksheets("Variant_LogTMS").Cells("A1").LoadFromDataTable(DT, True)

                Dim ColumIndex As Integer = 0
                For Each C As DataColumn In DT.Columns
                    Try
                        If C.DataType.FullName = "System.DateTime" Then
                            WB_EEPLUS.Workbook.Worksheets("Variant_LogTMS").Column(ColumIndex + 1).Style.Numberformat.Format = "mm-dd-yy"
                        End If
                        ColumIndex += 1
                    Catch ex As Exception
                    End Try
                Next

                'Set Header Color
                For I = 1 To DT.Columns.Count
                    WB_EEPLUS.Workbook.Worksheets("Variant_LogTMS").Cells(1, I).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                    WB_EEPLUS.Workbook.Worksheets("Variant_LogTMS").Cells(1, I).Style.Fill.BackgroundColor.SetColor(Color.MidnightBlue)
                    WB_EEPLUS.Workbook.Worksheets("Variant_LogTMS").Cells(1, I).Style.Font.Color.SetColor(Color.White)
                Next

                'Color Coding for Change Type
                Dim Row As Integer = 2
                For Each R As DataRow In DT.Rows

                    Select Case R("Change_Type")

                        Case "New Addition"
                            For I = 1 To DT.Columns.Count
                                WB_EEPLUS.Workbook.Worksheets("Variant_LogTMS").Cells(Row, I).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                                WB_EEPLUS.Workbook.Worksheets("Variant_LogTMS").Cells(Row, I).Style.Fill.BackgroundColor.SetColor(Color.Green)
                                WB_EEPLUS.Workbook.Worksheets("Variant_LogTMS").Cells(Row, I).Style.Font.Color.SetColor(Color.White)
                            Next
                        Case "Request for Deletion"
                            For I = 1 To DT.Columns.Count
                                WB_EEPLUS.Workbook.Worksheets("Variant_LogTMS").Cells(Row, I).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                                WB_EEPLUS.Workbook.Worksheets("Variant_LogTMS").Cells(Row, I).Style.Fill.BackgroundColor.SetColor(Color.Red)
                                WB_EEPLUS.Workbook.Worksheets("Variant_LogTMS").Cells(Row, I).Style.Font.Color.SetColor(Color.Black)
                            Next
                        Case "Pending Owner Change"
                            For I = 1 To DT.Columns.Count
                                WB_EEPLUS.Workbook.Worksheets("Variant_LogTMS").Cells(Row, I).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                                WB_EEPLUS.Workbook.Worksheets("Variant_LogTMS").Cells(Row, I).Style.Fill.BackgroundColor.SetColor(Color.Orange)
                                WB_EEPLUS.Workbook.Worksheets("Variant_LogTMS").Cells(Row, I).Style.Font.Color.SetColor(Color.Black)
                            Next
                    End Select
                    Row += 1
                Next
            Catch ex As Exception
            End Try
            'Log Imp LA
            Try

                'Data Source
                Dim DT As New DataTable
                DT = SQL_F.GetDataTable(CS, "Select * From Report_Americas_LogImpLA")
                WB_EEPLUS.Workbook.Worksheets.Add("Variant_LogImpLA")
                WB_EEPLUS.Workbook.Worksheets("Variant_LogImpLA").Cells("A1").LoadFromDataTable(DT, True)

                Dim ColumIndex As Integer = 0
                For Each C As DataColumn In DT.Columns
                    Try
                        If C.DataType.FullName = "System.DateTime" Then
                            WB_EEPLUS.Workbook.Worksheets("Variant_LogImpLA").Column(ColumIndex + 1).Style.Numberformat.Format = "mm-dd-yy"
                        End If
                        ColumIndex += 1
                    Catch ex As Exception
                    End Try
                Next

                'Set Header Color
                For I = 1 To DT.Columns.Count
                    WB_EEPLUS.Workbook.Worksheets("Variant_LogImpLA").Cells(1, I).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                    WB_EEPLUS.Workbook.Worksheets("Variant_LogImpLA").Cells(1, I).Style.Fill.BackgroundColor.SetColor(Color.MidnightBlue)
                    WB_EEPLUS.Workbook.Worksheets("Variant_LogImpLA").Cells(1, I).Style.Font.Color.SetColor(Color.White)
                Next

                'Color Coding for Change Type
                Dim Row As Integer = 2
                For Each R As DataRow In DT.Rows

                    Select Case R("Change_Type")

                        Case "New Addition"
                            For I = 1 To DT.Columns.Count
                                WB_EEPLUS.Workbook.Worksheets("Variant_LogImpLA").Cells(Row, I).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                                WB_EEPLUS.Workbook.Worksheets("Variant_LogImpLA").Cells(Row, I).Style.Fill.BackgroundColor.SetColor(Color.Green)
                                WB_EEPLUS.Workbook.Worksheets("Variant_LogImpLA").Cells(Row, I).Style.Font.Color.SetColor(Color.White)
                            Next
                        Case "Request for Deletion"
                            For I = 1 To DT.Columns.Count
                                WB_EEPLUS.Workbook.Worksheets("Variant_LogImpLA").Cells(Row, I).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                                WB_EEPLUS.Workbook.Worksheets("Variant_LogImpLA").Cells(Row, I).Style.Fill.BackgroundColor.SetColor(Color.Red)
                                WB_EEPLUS.Workbook.Worksheets("Variant_LogImpLA").Cells(Row, I).Style.Font.Color.SetColor(Color.Black)
                            Next
                        Case "Pending Owner Change"
                            For I = 1 To DT.Columns.Count
                                WB_EEPLUS.Workbook.Worksheets("Variant_LogImpLA").Cells(Row, I).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                                WB_EEPLUS.Workbook.Worksheets("Variant_LogImpLA").Cells(Row, I).Style.Fill.BackgroundColor.SetColor(Color.Orange)
                                WB_EEPLUS.Workbook.Worksheets("Variant_LogImpLA").Cells(Row, I).Style.Font.Color.SetColor(Color.Black)
                            Next
                    End Select
                    Row += 1
                Next
            Catch ex As Exception
            End Try
            'Log Imp NA
            Try

                'Data Source
                Dim DT As New DataTable
                DT = SQL_F.GetDataTable(CS, "Select * From Report_Americas_LogImpNA")
                WB_EEPLUS.Workbook.Worksheets.Add("Variant_LogImpNA")
                WB_EEPLUS.Workbook.Worksheets("Variant_LogImpNA").Cells("A1").LoadFromDataTable(DT, True)

                Dim ColumIndex As Integer = 0
                For Each C As DataColumn In DT.Columns
                    Try
                        If C.DataType.FullName = "System.DateTime" Then
                            WB_EEPLUS.Workbook.Worksheets("Variant_LogImpNA").Column(ColumIndex + 1).Style.Numberformat.Format = "mm-dd-yy"
                        End If
                        ColumIndex += 1
                    Catch ex As Exception
                    End Try
                Next

                'Set Header Color
                For I = 1 To DT.Columns.Count
                    WB_EEPLUS.Workbook.Worksheets("Variant_LogImpNA").Cells(1, I).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                    WB_EEPLUS.Workbook.Worksheets("Variant_LogImpNA").Cells(1, I).Style.Fill.BackgroundColor.SetColor(Color.MidnightBlue)
                    WB_EEPLUS.Workbook.Worksheets("Variant_LogImpNA").Cells(1, I).Style.Font.Color.SetColor(Color.White)
                Next


                'Color Coding for Change Type
                Dim Row As Integer = 2
                For Each R As DataRow In DT.Rows

                    Select Case R("Change_Type")

                        Case "New Addition"
                            For I = 1 To DT.Columns.Count
                                WB_EEPLUS.Workbook.Worksheets("Variant_LogImpNA").Cells(Row, I).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                                WB_EEPLUS.Workbook.Worksheets("Variant_LogImpNA").Cells(Row, I).Style.Fill.BackgroundColor.SetColor(Color.Green)
                                WB_EEPLUS.Workbook.Worksheets("Variant_LogImpNA").Cells(Row, I).Style.Font.Color.SetColor(Color.White)
                            Next
                        Case "Request for Deletion"
                            For I = 1 To DT.Columns.Count
                                WB_EEPLUS.Workbook.Worksheets("Variant_LogImpNA").Cells(Row, I).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                                WB_EEPLUS.Workbook.Worksheets("Variant_LogImpNA").Cells(Row, I).Style.Fill.BackgroundColor.SetColor(Color.Red)
                                WB_EEPLUS.Workbook.Worksheets("Variant_LogImpNA").Cells(Row, I).Style.Font.Color.SetColor(Color.Black)
                            Next
                        Case "Pending Owner Change"
                            For I = 1 To DT.Columns.Count
                                WB_EEPLUS.Workbook.Worksheets("Variant_LogImpNA").Cells(Row, I).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                                WB_EEPLUS.Workbook.Worksheets("Variant_LogImpNA").Cells(Row, I).Style.Fill.BackgroundColor.SetColor(Color.Orange)
                                WB_EEPLUS.Workbook.Worksheets("Variant_LogImpNA").Cells(Row, I).Style.Font.Color.SetColor(Color.Black)
                            Next
                    End Select
                    Row += 1
                Next
            Catch ex As Exception
            End Try
            'Log Direct
            Try

                'Data Source
                Dim DT As New DataTable
                DT = SQL_F.GetDataTable(CS, "Select * From Report_Americas_LogDirect")
                WB_EEPLUS.Workbook.Worksheets.Add("Variant_LogDirect")
                WB_EEPLUS.Workbook.Worksheets("Variant_LogDirect").Cells("A1").LoadFromDataTable(DT, True)

                Dim ColumIndex As Integer = 0
                For Each C As DataColumn In DT.Columns
                    Try
                        If C.DataType.FullName = "System.DateTime" Then
                            WB_EEPLUS.Workbook.Worksheets("Variant_LogDirect").Column(ColumIndex + 1).Style.Numberformat.Format = "mm-dd-yy"
                        End If
                        ColumIndex += 1
                    Catch ex As Exception
                    End Try
                Next

                'Set Header Color
                For I = 1 To DT.Columns.Count
                    WB_EEPLUS.Workbook.Worksheets("Variant_LogDirect").Cells(1, I).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                    WB_EEPLUS.Workbook.Worksheets("Variant_LogDirect").Cells(1, I).Style.Fill.BackgroundColor.SetColor(Color.MidnightBlue)
                    WB_EEPLUS.Workbook.Worksheets("Variant_LogDirect").Cells(1, I).Style.Font.Color.SetColor(Color.White)
                Next

                'Color Coding for Change Type
                Dim Row As Integer = 2
                For Each R As DataRow In DT.Rows

                    Select Case R("Change_Type")

                        Case "New Addition"
                            For I = 1 To DT.Columns.Count
                                WB_EEPLUS.Workbook.Worksheets("Variant_LogDirect").Cells(Row, I).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                                WB_EEPLUS.Workbook.Worksheets("Variant_LogDirect").Cells(Row, I).Style.Fill.BackgroundColor.SetColor(Color.Green)
                                WB_EEPLUS.Workbook.Worksheets("Variant_LogDirect").Cells(Row, I).Style.Font.Color.SetColor(Color.White)
                            Next
                        Case "Request for Deletion"
                            For I = 1 To DT.Columns.Count
                                WB_EEPLUS.Workbook.Worksheets("Variant_LogDirect").Cells(Row, I).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                                WB_EEPLUS.Workbook.Worksheets("Variant_LogDirect").Cells(Row, I).Style.Fill.BackgroundColor.SetColor(Color.Red)
                                WB_EEPLUS.Workbook.Worksheets("Variant_LogDirect").Cells(Row, I).Style.Font.Color.SetColor(Color.Black)
                            Next
                        Case "Pending Owner Change"
                            For I = 1 To DT.Columns.Count
                                WB_EEPLUS.Workbook.Worksheets("Variant_LogDirect").Cells(Row, I).Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                                WB_EEPLUS.Workbook.Worksheets("Variant_LogDirect").Cells(Row, I).Style.Fill.BackgroundColor.SetColor(Color.Orange)
                                WB_EEPLUS.Workbook.Worksheets("Variant_LogDirect").Cells(Row, I).Style.Font.Color.SetColor(Color.Black)
                            Next
                    End Select
                    Row += 1
                Next
            Catch ex As Exception
            End Try

            WB_EEPLUS.Save()
            Export_AllVariants = True
        Catch ex As Exception
            Export_AllVariants = False
        End Try
    End Function








End Class
