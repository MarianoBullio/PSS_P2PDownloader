
'**************************************
' Project: Get_Reports
' Created By: Mariano Bullio, bullio.m@pg.com
'**************************************

Imports Microsoft.VisualBasic
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.Data
Imports System.IO
Imports System.Globalization
Imports System.Threading

Public Class Get_Reports

    Private Session As SAPGUI = Nothing
    Private Box As String = ""
    Private Client As String = ""
    Private Login As String = ""
    Private Password As String = ""

    Private message_SAP As String = Nothing
    Private message_SAP_number As Integer = 0
    Private message_SAP_type As String = Nothing
    Private SF As New System_Functions
    Private SQL_F As New SQL_Functions

    Sub New(ByVal Box As String, ByVal Client As String, ByVal Login As String, ByVal Password As String)
        Me.Box = Box
        Me.Client = Client
        Me.Login = Login
        Me.Password = Password
    End Sub

    'SAP ****************************************************************************************

    Private Sub Open_SAP()
        Session = New SAPGUI(Box, Client, Login, Password)
    End Sub

    Private Sub Close_SAP()
        Session.Close()
        Session = Nothing
    End Sub

    Private Function Read_SAP_Message(ByVal Session As Object) As Boolean
        Try
            message_SAP = ""
            message_SAP_number = 0
            message_SAP_type = ""
            If Session.findById("wnd[0]/sbar").text <> "" Then
                message_SAP = Session.findById("wnd[0]/sbar").text
                message_SAP_number = Session.findById("wnd[0]/sbar").MessageNumber
                message_SAP_type = Session.findById("wnd[0]/sbar").MessageType
            End If
            Read_SAP_Message = True
        Catch
            Read_SAP_Message = False
        End Try
    End Function

    Private Function Read_SAP_Report(ByVal Path As String, Optional ByVal SkipTopLines As Integer = 0) As DataTable

        Read_SAP_Report = Nothing
        Try
            Dim F As New Microsoft.VisualBasic.FileIO.TextFieldParser(Path)
            F.TextFieldType = FileIO.FieldType.Delimited
            F.SetDelimiters(Chr(9))
            Dim R As String()
            Dim D As New DataTable
            Dim CI As Integer

            For I As Integer = 1 To SkipTopLines
                R = F.ReadFields
            Next

            R = F.ReadFields
            If R.Length > 0 Then
                CI = 1
                For Each CN As String In R
                    Do While Not D.Columns(CN) Is Nothing
                        CN = CN & CI
                        CI += 1
                    Loop
                    D.Columns.Add(CN, Type.GetType("System.String"))
                Next
            End If

            While Not F.EndOfData
                Try
                    R = F.ReadFields
                    D.LoadDataRow(R, True)
                Catch ex As Exception
                End Try
            End While

            Read_SAP_Report = D

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Function

    'BI Report **********************************************************************************

    Public Sub Download_BI_Reports()
        Open_SAP()
        SF.SetLogInfo("Downloading BI Report (NA)")
        DownloadAPTradeReport("NA")
        SF.SetLogInfo("Downloading BI Report (LA)")
        DownloadAPTradeReport("LA")
        Close_SAP()
    End Sub

    Public Sub Read_BI_Reports()
        SF.SetLogInfo("Reading BI Report (NA)")
        ReadAPTradeReport("NA")
        SF.SetLogInfo("Reading BI Report (LA)")
        ReadAPTradeReport("LA")
    End Sub

    Private Sub DownloadAPTradeReport(ByVal Region As String)

        Dim FileName As String = "BI Report " & Region & " (" & Date.Now.Month & "-" & Date.Now.Day & "-" & Date.Now.Year & ").xls"
        Dim FilePath As String = EXCELPath & FileName

        Try
            'Delete The Last File
            Try
                Kill(FilePath)
            Catch
            End Try

            ''USING F.98
            ''Go to Transaction F.98
            'Session.StartTransaction("F.98")
            'Session.FindById("wnd[0]").maximize()

            ''Select Monthly Reports
            'Session.FindById("wnd[0]/usr/lbl[5,7]").setFocus()
            'Session.FindById("wnd[0]/usr/lbl[5,7]").caretPosition = 0 '1
            'Session.FindById("wnd[0]").sendVKey(2)

            ''Select Automatic Trade BSR tool for Accounts Payable Store Only option
            'Session.FindById("wnd[0]/usr/lbl[12,22]").setFocus()
            'Session.FindById("wnd[0]/usr/lbl[12,22]").caretPosition = 38 '40
            'Session.FindById("wnd[0]").sendVKey(2)

            'Display SAP Menu
            Session.FindById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").expandNode("0000000100")
            'P&G Report Menu
            Session.FindById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").expandNode("0000000106")
            'P2P Area Menu
            'Daily Reports
            Session.FindById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").expandNode("0000000125")
            'Automatic Trade BSR tool for payble accounts DISPLAY
            Session.FindById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").selectedNode = "0000000133"
            Session.FindById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").doubleClickNode("0000000133")

            'Enter Region
            Session.FindById("wnd[0]/usr/ctxtP_REGION").text = Region

            'Run
            Session.FindById("wnd[0]/usr/ctxtP_REGION").caretPosition = 2
            Session.FindById("wnd[0]/tbar[1]/btn[8]").press()

            'Export to Excel
            Session.FindById("wnd[0]/tbar[0]/okcd").text = "%PC"
            Session.FindById("wnd[0]").sendVKey(0)
            Session.FindById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
            Session.FindById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus()
            Session.FindById("wnd[1]/tbar[0]/btn[0]").press()
            Session.FindById("wnd[1]/usr/ctxtDY_PATH").text = EXCELPath
            Session.FindById("wnd[1]/usr/ctxtDY_FILENAME").text = FileName
            Session.FindById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 10
            Session.FindById("wnd[1]/tbar[0]/btn[0]").press()

            Read_SAP_Message(Session)

            If message_SAP_type = "S" Then
                If message_SAP.Contains("transmitted") Then
                    Session.FindById("wnd[0]/tbar[0]/btn[3]").press() ' back
                    Session.FindById("wnd[0]/tbar[0]/btn[3]").press() ' back
                    Session.FindById("wnd[0]/tbar[0]/btn[3]").press() ' back
                End If
            End If

        Catch ex As Exception
            SF.SetLogInfo("Error Downloading BI Report " & Region & " Check SAP Script")
        End Try
    End Sub

    Private Sub ReadAPTradeReport(ByVal Region As String)
        Dim FileName As String = "BI Report " & Region & " (" & Date.Now.Month & "-" & Date.Now.Day & "-" & Date.Now.Year & ").xls"
        Dim FilePath As String = EXCELPath & FileName

        Try
            'Read Report
            Dim DT As DataTable = Read_SAP_Report(FilePath)
            If DT Is Nothing Then Exit Sub

            'Delete Old Invoices
            SQL_F.SQL_Execute_NQ(CS, "Delete From P2P_BI Where Region='" & Region & "'")

            'Delete Upload Today
            SQL_F.SQL_Execute_NQ(CS, "Delete From P2P_BI_UploadToday Where Region='" & Region & "'")

            'Customazing BI Report
            If Not DT.Columns("Column1") Is Nothing Then DT.Columns.Remove("Column1")
            If Not DT.Columns("Column2") Is Nothing Then DT.Columns.Remove("Column2") '
            If Not DT.Columns("Column3") Is Nothing Then DT.Columns.Remove("Column3") '
            If Not DT.Columns("Column4") Is Nothing Then DT.Columns.Remove("Column4") '
            DT.Columns("Log.System").ColumnName = "Box"
            DT.Columns("CoCd").ColumnName = "LE"
            If Not DT.Columns("Year") Is Nothing Then DT.Columns.Remove("Year")
            DT.Columns("DocumentNo").ColumnName = "Doc_Number"
            If Not DT.Columns("Itm") Is Nothing Then DT.Columns.Remove("Itm")
            If Not DT.Columns("Reference Key") Is Nothing Then DT.Columns.Remove("Reference Key")
            If Not DT.Columns("Region") Is Nothing Then DT.Columns.Remove("Region")
            If Not DT.Columns("Status light") Is Nothing Then DT.Columns.Remove("Status light")
            If Not DT.Columns("Accty") Is Nothing Then DT.Columns.Remove("Accty")
            DT.Columns("Name 1").ColumnName = "Vendor_Name"
            DT.Columns("Vend. Cty").ColumnName = "Vendor_Cty"
            If Not DT.Columns("Group") Is Nothing Then DT.Columns.Remove("Group")
            If Not DT.Columns("SG") Is Nothing Then DT.Columns.Remove("SG")
            DT.Columns("Pstng Date").ColumnName = "Posting_Date"
            DT.Columns("Doc. Date").ColumnName = "Doc_Date"
            DT.Columns("Entry Date").ColumnName = "Entry_Date"
            DT.Columns("Crcy").ColumnName = "Currency"
            DT.Columns("Type").ColumnName = "Doc_Type"
            If Not DT.Columns("D/C") Is Nothing Then DT.Columns.Remove("D/C")
            If Not DT.Columns("BusA") Is Nothing Then DT.Columns.Remove("BusA")
            If Not DT.Columns("Tx") Is Nothing Then DT.Columns.Remove("Tx")
            If Not DT.Columns("Amount in LC") Is Nothing Then DT.Columns.Remove("Amount in LC")
            If Not DT.Columns("Crcy1") Is Nothing Then DT.Columns.Remove("Crcy1")
            If Not DT.Columns("Am.Doc.Curr") Is Nothing Then DT.Columns.Remove("Am.Doc.Curr")
            If Not DT.Columns("G/L") Is Nothing Then DT.Columns.Remove("G/L")
            If Not DT.Columns("Recon.acct") Is Nothing Then DT.Columns.Remove("Recon.acct")
            DT.Columns("Bline Date").ColumnName = "BaseLine_Date"
            DT.Columns("PayT").ColumnName = "PTerms"
            If Not DT.Columns("Net") Is Nothing Then DT.Columns.Remove("Net")
            If Not DT.Columns("Disc.1") Is Nothing Then DT.Columns.Remove("Disc.1")
            If Not DT.Columns("Disc.2") Is Nothing Then DT.Columns.Remove("Disc.2")
            If Not DT.Columns("Disc.date") Is Nothing Then DT.Columns.Remove("Disc.date")
            DT.Columns("Due on").ColumnName = "Due_Date"
            DT.Columns("Days o/due").ColumnName = "Due_Days"
            If Not DT.Columns("Discount base") Is Nothing Then DT.Columns.Remove("Discount base")
            If Not DT.Columns("Discount amount") Is Nothing Then DT.Columns.Remove("Discount amount")
            If Not DT.Columns("Discount amount2") Is Nothing Then DT.Columns.Remove("Discount amount2")
            If Not DT.Columns("House Bk") Is Nothing Then DT.Columns.Remove("House Bk")
            If Not DT.Columns("S") Is Nothing Then DT.Columns.Remove("S")
            If Not DT.Columns("Tr.Prt") Is Nothing Then DT.Columns.Remove("Tr.Prt")
            DT.Columns("Amnt.Group").ColumnName = "Amount_Group"
            DT.Columns("ABS Amount").ColumnName = "ABS_Amount"
            If Not DT.Columns("LCur2") Is Nothing Then DT.Columns.Remove("LCur2")
            If Not DT.Columns("Amount Hard Curr") Is Nothing Then DT.Columns.Remove("Amount Hard Curr")
            If Not DT.Columns("Reverse clearing") Is Nothing Then DT.Columns.Remove("Reverse clearing")
            If Not DT.Columns("PmtMthSu") Is Nothing Then DT.Columns.Remove("PmtMthSu")
            If Not DT.Columns("Reversal flag") Is Nothing Then DT.Columns.Remove("Reversal flag")
            If Not DT.Columns("Ref.key 1") Is Nothing Then DT.Columns.Remove("Ref.key 1")
            If Not DT.Columns("Ref.key 2") Is Nothing Then DT.Columns.Remove("Ref.key 2")
            If Not DT.Columns("Reference key 3") Is Nothing Then DT.Columns.Remove("Reference key 3")
            DT.Columns("MM Doc.no.").ColumnName = "MM_Doc_Number"
            DT.Columns("Purch.Doc.").ColumnName = "Purch_Doc"
            DT.Columns("Plnt").ColumnName = "Plant"
            DT.Columns("Plant description").ColumnName = "Plant_Description"
            DT.Columns("PGr").ColumnName = "PGrp"
            DT.Columns("Description").ColumnName = "PGrp_Description"
            DT.Columns("Purch. Org. Descr.").ColumnName = "POrg_Description"
            DT.Columns("Prc").ColumnName = "Block_Price"
            DT.Columns("Qty").ColumnName = "Block_Quantity"
            DT.Columns("Block ind.").ColumnName = "Block_Ind"
            If Not DT.Columns("Err") Is Nothing Then DT.Columns.Remove("Err")
            DT.Columns("Bl.Reason").ColumnName = "BI_Reason"
            DT.Columns("Blocking Reason description").ColumnName = "Pending_Reason"
            If Not DT.Columns("More R.Bl?") Is Nothing Then DT.Columns.Remove("More R.Bl?")
            If Not DT.Columns("Aging") Is Nothing Then DT.Columns.Remove("Aging")
            DT.Columns("User Name").ColumnName = "User_Name"

            'Upload Today Table
            Dim DT_UploadToday As New DataTable
            DT_UploadToday.Columns.Add("Box", GetType(String))
            DT_UploadToday.Columns.Add("LE", GetType(Double))
            DT_UploadToday.Columns.Add("Doc_Number", GetType(Double))
            DT_UploadToday.Columns.Add("MM_Doc_Number", GetType(Double))
            DT_UploadToday.Columns.Add("Region", GetType(String))

            'Validate Invoices
            Try

                For Each R As DataRow In DT.Rows
                    Dim NewUploadTodayRow As DataRow = DT_UploadToday.NewRow

                    'Box
                    Try
                        R("Box") = R("Box").ToString.Trim.Substring(0, 3)
                    Catch ex As Exception
                    End Try

                    'LE
                    Try
                        R("LE") = CDbl(R("LE"))
                    Catch ex As Exception
                        R("LE") = 0
                    End Try

                    'Doc_Number
                    Try
                        R("Doc_Number") = CDbl(R("Doc_Number"))
                    Catch ex As Exception
                        R("Doc_Number") = 0
                    End Try


                    'Vendor
                    Try
                        R("Vendor") = CDbl(R("Vendor"))
                    Catch ex As Exception
                        R("Vendor") = 0
                    End Try

                    'Amount_Group
                    Try
                        R("Amount_Group") = CDbl(R("Amount_Group"))
                    Catch ex As Exception
                        R("Amount_Group") = 0
                    End Try

                    'ABS_Amount
                    Try
                        R("ABS_Amount") = CDbl(R("ABS_Amount"))
                    Catch ex As Exception
                        R("ABS_Amount") = 0
                    End Try

                    'MM_Doc_Number
                    If Not IsNumeric(R("MM_Doc_Number")) Then
                        R("MM_Doc_Number") = 0
                    End If

                    'Purch_Doc
                    If Not IsNumeric(R("Purch_Doc")) Then
                        R("Purch_Doc") = 0
                    End If

                    'POrg
                    If Not IsNumeric(R("POrg")) Then
                        R("POrg") = 0
                    End If

                    If Not Validate_BI(Region, R) Then
                        R.Delete()
                    Else
                        Try
                            NewUploadTodayRow("Box") = R("Box")
                            NewUploadTodayRow("LE") = R("LE")
                            NewUploadTodayRow("Doc_Number") = R("Doc_Number")
                            NewUploadTodayRow("MM_Doc_Number") = R("MM_Doc_Number")
                            NewUploadTodayRow("Region") = Region
                            DT_UploadToday.Rows.Add(NewUploadTodayRow)
                        Catch ex As Exception
                        End Try
                     
                    End If
                Next
                DT.AcceptChanges()
            Catch ex As Exception
                SF.SetLogInfo("Error Validating Invoices" & Region)
            End Try

            'New Columns
            Dim NewColumnScope As New DataColumn("Scope", GetType(String))
            NewColumnScope.DefaultValue = "TBD"
            DT.Columns.Add(NewColumnScope)
            Dim NewColumnOwner As New DataColumn("Owner", GetType(String))
            NewColumnOwner.DefaultValue = "TBD"
            DT.Columns.Add(NewColumnOwner)
            DT.Columns.Add("Material_Group", GetType(String))
            DT.Columns.Add("Upload_Date", GetType(Date))
            Dim NewColumn As New DataColumn("Region", GetType(String))
            NewColumn.DefaultValue = Region
            DT.Columns.Add(NewColumn)
            DT.Columns.Add("Material", GetType(String))

            'Save Invoices in SQL
            Try
                Dim BulkInsert As String = SQL_F.Bulk_Insert(CS, "P2P_BI", DT)
                If BulkInsert <> Nothing Then
                    SF.SetLogInfo("Error Saving BI Report in SQL " & Region)
                End If
            Catch ex As Exception
                SF.SetLogInfo("Error Saving BI Report in SQL " & Region)
            End Try

            'Save Upload Today Invoices in SQL
            Try
                Dim BulkInsert As String = SQL_F.Bulk_Insert(CS, "P2P_BI_UploadToday", DT_UploadToday)
                If BulkInsert <> Nothing Then
                    SF.SetLogInfo("Error Saving BI UploadToday in SQL " & Region)
                End If
            Catch ex As Exception
                SF.SetLogInfo("Error Saving BI UploadToday in SQL " & Region)
            End Try


        Catch ex As Exception
            SF.SetLogInfo("Error Reading BI Report " & Region)
        End Try
    End Sub

    Private Function Validate_BI(ByVal Region As String, ByVal DR As DataRow) As Boolean

        Validate_BI = False
        If Region = "NA" Then
            If (DR("PBk") = "D") Or (DR("PBk") = "R" And (DR("POrg") = "1129" Or DR("POrg") = "1201" Or DR("POrg") = "1345" Or DR("POrg") = "1346" Or DR("POrg") = "1485" Or DR("POrg") = "1522" Or DR("POrg") = "1382")) _
            Or (DR("PBk") = "R" And DR("POrg") <> "1129" And DR("POrg") <> "1201" And DR("Block_Price") = "X") Then
                Validate_BI = True
            End If
            If (DR("POrg") = "1345" Or DR("POrg") = "1346") And DR("Block_Quantity") = "X" And DR("Block_Price") = "" Then
                If DR("Plant") = "9950" Or DR("Plant") = "9951" Or DR("Plant") = "9954" Or (DR("Plant") = "7209" And DR("Box") = "G4P") Then
                    Validate_BI = True
                Else
                    Validate_BI = False
                End If
            End If
        End If
        If Region = "LA" Then
            If (DR("LE") <> "480" And DR("LE") <> "814" And DR("LE") <> "830") And ((DR("PBk") = "D") Or (DR("PBk") = "R" And DR("Block_Price") = "X") Or (DR("PBk") = "R" And DR("Block_Price") <> "X" And DR("Block_Quantity") <> "X")) Then
                Validate_BI = True
            End If
            If DR("PBk") = "D" And DR("Block_Price") <> "X" And DR("Block_Quantity") = "X" Then
                Validate_BI = False
            End If
        End If

    End Function

    Public Function DownloadMaterialGroup_BI(ByRef DT_Passwords As DataTable) As Boolean
        Try

            SF.SetLogInfo("Downloading Material Groups - BI ")

            'Disctict SAP Boxes
            Dim DT_SAP_Box As New DataTable
            DT_SAP_Box = SQL_F.GetDataTable(CS, "Select * From P2P_SAP_Boxes where Download_MatGroup =1")

            'For each SAP Box
            For Each R_SAP_Box As DataRow In DT_SAP_Box.Rows


                If R_SAP_Box("Box") = "ANP" Then
                    R_SAP_Box("Box") = R_SAP_Box("Box") & "_" & R_SAP_Box("Client")
                End If

                'List of Documents to find Material Group Codes
                Dim DT_Documents As New DataTable
                DT_Documents = SQL_F.GetDataTable(CS, "SELECT * FROM P2P_BI WHERE (Box='" & R_SAP_Box("Box") & "') AND ((Material_Group IS NULL) or (Material_Group = '')) ")

                If Not DT_Documents Is Nothing Then

                    If DT_Documents.Rows.Count > 0 Then

                        Dim GetTables As New GetSAPTables
                        Dim DT_EKPO As DataTable

                        DT_EKPO = GetTables.Get_EKPO(R_SAP_Box("Box"), R_SAP_Box("Login"), R_SAP_Box("Password"), DT_Documents)

                        If Not DT_EKPO Is Nothing Then

                            Dim EKPO_Lines As DataRow()
                            EKPO_Lines = DT_EKPO.Select("(LineItem = 1) or (LineItem = 10)")

                            For Each Line As DataRow In EKPO_Lines
                                Dim Update As String
                                Dim Material_Group As String = Line("Material_Group")
                                Dim Material As String = Line("Material")
                                Dim User_Name As String = Line("AFNAM")
                                Try
                                    Material_Group = Material_Group.Substring(0, 9)
                                Catch ex As Exception
                                End Try
                                Try
                                    Material = Material.Substring(0, 8)
                                Catch ex As Exception
                                End Try
                                Try
                                    User_Name = User_Name.Substring(0, 6)
                                Catch ex As Exception
                                End Try
                                Update = "Update P2P_BI set Material_Group ='" & Material_Group & "', Material='" & Material & "', User_Name='" & User_Name & "' "
                                Update += " Where (Box='" & R_SAP_Box("Box") & "') AND (Purch_Doc=" & Line("Purch_Doc") & ")"
                                Dim Exc As String = SQL_F.SQL_Execute_NQ(CS, Update)
                                If Exc <> Nothing Then
                                End If
                            Next
                        End If
                    End If

                End If
            Next

        Catch ex As Exception
            SF.SetLogInfo("Error Downloading Material/Material_Group BI")
        End Try
    End Function

    Public Function Distribute_BI_Items() As Boolean

        SF.SetLogInfo("Applying Variants - BI ")

        'Customization
        If SQL_F.SQL_Execute_NQ(CS, "Exec ReAssign_Americas_BI_1_Custom") <> Nothing Then
            SF.SetLogInfo("Error: BI Distribution: Customization [Exec ReAssign_Americas_BI_1_Custom]")
        End If

        'LogIndNAPD
        If SQL_F.SQL_Execute_NQ(CS, "Exec ReAssign_Americas_BI_2_LogIndNAPD") <> Nothing Then
            SF.SetLogInfo("Error: BI Distribution: LogIndNAPD [Exec ReAssign_Americas_BI_2_LogIndNAPD]")
        End If

        'SS 01
        If SQL_F.SQL_Execute_NQ(CS, "Exec ReAssign_Americas_BI_3_SS_01") <> Nothing Then
            SF.SetLogInfo("Error: BI Distribution: SS 01 [Exec ReAssign_Americas_BI_3_SS_01]")
        End If
        'SS 02
        If SQL_F.SQL_Execute_NQ(CS, "Exec ReAssign_Americas_BI_3_SS_02") <> Nothing Then
            SF.SetLogInfo("Error: BI Distribution: SS 02 [Exec ReAssign_Americas_BI_3_SS_02]")
        End If
        'SS 03
        If SQL_F.SQL_Execute_NQ(CS, "Exec ReAssign_Americas_BI_3_SS_03") <> Nothing Then
            SF.SetLogInfo("Error: BI Distribution: SS 03 [Exec ReAssign_Americas_BI_3_SS_03]")
        End If
        'SS 04
        If SQL_F.SQL_Execute_NQ(CS, "Exec ReAssign_Americas_BI_3_SS_04") <> Nothing Then
            SF.SetLogInfo("Error: BI Distribution: SS 04 [Exec ReAssign_Americas_BI_3_SS_04]")
        End If
        'SS 05
        If SQL_F.SQL_Execute_NQ(CS, "Exec ReAssign_Americas_BI_3_SS_05") <> Nothing Then
            SF.SetLogInfo("Error: BI Distribution: SS 05 [Exec ReAssign_Americas_BI_3_SS_05]")
        End If
        'SS 06
        If SQL_F.SQL_Execute_NQ(CS, "Exec ReAssign_Americas_BI_3_SS_06") <> Nothing Then
            SF.SetLogInfo("Error: BI Distribution: SS 06 [Exec ReAssign_Americas_BI_3_SS_06]")
        End If
        'SS 07
        If SQL_F.SQL_Execute_NQ(CS, "Exec ReAssign_Americas_BI_3_SS_07") <> Nothing Then
            SF.SetLogInfo("Error: BI Distribution: SS 07 [Exec ReAssign_Americas_BI_3_SS_07]")
        End If
        'SS 08
        If SQL_F.SQL_Execute_NQ(CS, "Exec ReAssign_Americas_BI_3_SS_08") <> Nothing Then
            SF.SetLogInfo("Error: BI Distribution: SS 08 [Exec ReAssign_Americas_BI_3_SS_08]")
        End If
        'SS 12
        If SQL_F.SQL_Execute_NQ(CS, "Exec ReAssign_Americas_BI_3_SS_12") <> Nothing Then
            SF.SetLogInfo("Error: BI Distribution: SS 12 [Exec ReAssign_Americas_BI_3_SS_12]")
        End If
        'SS 13
        If SQL_F.SQL_Execute_NQ(CS, "Exec ReAssign_Americas_BI_3_SS_13") <> Nothing Then
            SF.SetLogInfo("Error: BI Distribution: SS 13 [Exec ReAssign_Americas_BI_3_SS_13]")
        End If
        'SS 14
        If SQL_F.SQL_Execute_NQ(CS, "Exec ReAssign_Americas_BI_3_SS_14") <> Nothing Then
            SF.SetLogInfo("Error: BI Distribution: SS 14 [Exec ReAssign_Americas_BI_3_SS_14]")
        End If
        'SS 15
        If SQL_F.SQL_Execute_NQ(CS, "Exec ReAssign_Americas_BI_3_SS_15") <> Nothing Then
            SF.SetLogInfo("Error: BI Distribution: SS 15 [Exec ReAssign_Americas_BI_3_SS_15]")
        End If
        'SS 16
        If SQL_F.SQL_Execute_NQ(CS, "Exec ReAssign_Americas_BI_3_SS_16") <> Nothing Then
            SF.SetLogInfo("Error: BI Distribution: SS 16 [Exec ReAssign_Americas_BI_3_SS_16]")
        End If
        'SS 17
        If SQL_F.SQL_Execute_NQ(CS, "Exec ReAssign_Americas_BI_3_SS_17") <> Nothing Then
            SF.SetLogInfo("Error: BI Distribution: SS 17 [Exec ReAssign_Americas_BI_3_SS_17]")
        End If

        'STR
        If SQL_F.SQL_Execute_NQ(CS, "Exec ReAssign_Americas_BI_4_STR") <> Nothing Then
            SF.SetLogInfo("Error: BI Distribution: STR [Exec ReAssign_Americas_BI_4_STR]")
        End If

        'Direct Exec ReAssign_Americas_BI_5_Direct
        If SQL_F.SQL_Execute_NQ(CS, "Exec ReAssign_Americas_BI_5_Direct") <> Nothing Then
            SF.SetLogInfo("Error: BI Distribution: Direct [Exec ReAssign_Americas_BI_5_Direct]")
        End If

        'LogTMS
        If SQL_F.SQL_Execute_NQ(CS, "Exec ReAssign_Americas_BI_6_LogTMS") <> Nothing Then
            SF.SetLogInfo("Error: BI Distribution: LogTMS [Exec ReAssign_Americas_BI_6_LogTMS]")
        End If

        'LogImpLA (SS) Exec ReAssign_Americas_BI_7_LogImpLA_SS
        If SQL_F.SQL_Execute_NQ(CS, "Exec ReAssign_Americas_BI_7_LogImpLA_SS") <> Nothing Then
            SF.SetLogInfo("Error: BI Distribution: LogImpLA (SS) [Exec ReAssign_Americas_BI_7_LogImpLA_SS]")
        End If

        'LogImpLA (STR) Exec ReAssign_Americas_BI_7_LogImpLA_STR
        If SQL_F.SQL_Execute_NQ(CS, "Exec ReAssign_Americas_BI_7_LogImpLA_STR") <> Nothing Then
            SF.SetLogInfo("Error: BI Distribution: LogImpLA (STR) [Exec ReAssign_Americas_BI_7_LogImpLA_STR]")
        End If

        'LogImpNA
        If SQL_F.SQL_Execute_NQ(CS, "Exec ReAssign_Americas_BI_8_LogImpNA") <> Nothing Then
            SF.SetLogInfo("Error: BI Distribution: LogImpNA [Exec ReAssign_Americas_BI_8_LogImpNA]")
        End If

        'LogDirect
        If SQL_F.SQL_Execute_NQ(CS, "Exec ReAssign_Americas_BI_9_LogDirect") <> Nothing Then
            SF.SetLogInfo("Error: BI Distribution: LogImpNA [Exec ReAssign_Americas_BI_9_LogDirect]")
        End If

        'SS EndUser
        If SQL_F.SQL_Execute_NQ(CS, "Exec ReAssign_Americas_BI_10_SS_EndUser") <> Nothing Then
            SF.SetLogInfo("Error: BI Distribution: SS EndUser [Exec ReAssign_Americas_BI_10_SS_EndUser]")
        End If

        'STR TopSuppliers
        If SQL_F.SQL_Execute_NQ(CS, "Exec ReAssign_Americas_BI_11_STR_TopSuppliers") <> Nothing Then
            SF.SetLogInfo("Error: BI Distribution: STR TopSuppliers [Exec ReAssign_Americas_BI_11_STR_TopSuppliers]")
        End If

        'OutOfScope
        If SQL_F.SQL_Execute_NQ(CS, "Exec ReAssign_Americas_BI_12_OutOfScope") <> Nothing Then
            SF.SetLogInfo("Error: BI Distribution: POrgs OutOfScope [Exec ReAssign_Americas_BI_12_OutOfScope]")
        End If


        'Delete Excluded POrgs
        If SQL_F.SQL_Execute_NQ(CS, "Delete From P2P_BI Where POrg in (Select POrg From P2P_BI_ExcludedPOrgs)") <> Nothing Then
            SF.SetLogInfo("Error: BI Distribution: Delete Excluded POrgs [Delete From P2P_BI Where POrg in (Select POrg From P2P_BI_ExcludedPOrgs)]")
        End If

    End Function

    'S161 Report **********************************************************************************

    Public Sub Download_S161_Report()
        Open_SAP()
        SF.SetLogInfo("Downloading S161 Report (" & Box & ")")
        DownloadZFI2Report()
        Close_SAP()
    End Sub

    Public Sub Read_S161_Report()
        SF.SetLogInfo("Reading S161 Report (" & Box & ")")
        ReadZFI2Report("Americas-" & Box)
    End Sub

    Private Sub DownloadZFI2Report()

        Dim FileName As String = "S161 Report " & Box & " (" & Date.Now.Month & "-" & Date.Now.Day & "-" & Date.Now.Year & ").xls"
        Dim FilePath As String = EXCELPath & FileName

        Try
            'Delete The Last File
            Try
                Kill(FilePath)
            Catch
            End Try

            'Go to Transaction ZFI2
            Session.StartTransaction("ZFI2")
            Session.FindById("wnd[0]").maximize()

            'Select Status 161
            Session.FindById("wnd[0]/usr/ctxtS_STATUS-LOW").text = "161"
            Session.FindById("wnd[0]/usr/txtP_MXHITS").text = "9999"
            Session.FindById("wnd[0]/usr/ctxtS_STATUS-LOW").setFocus()
            Session.FindById("wnd[0]/usr/ctxtS_STATUS-LOW").caretPosition = 3

            'Run
            Session.FindById("wnd[0]/tbar[1]/btn[8]").press()

            'Export to Excel
            Session.FindById("wnd[0]/tbar[0]/okcd").text = "%PC"
            Session.FindById("wnd[0]").sendVKey(0)
            Session.FindById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
            Session.FindById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus()
            Session.FindById("wnd[1]/tbar[0]/btn[0]").press()
            Session.FindById("wnd[1]/usr/ctxtDY_PATH").text = EXCELPath
            Session.FindById("wnd[1]/usr/ctxtDY_FILENAME").text = FileName
            Session.FindById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 10
            Session.FindById("wnd[1]/tbar[0]/btn[0]").press()

            Read_SAP_Message(Session)

            If message_SAP_type = "S" Then
                If message_SAP.Contains("transmitted") Then
                    Session.FindById("wnd[0]/tbar[0]/btn[3]").press() ' back
                    Session.FindById("wnd[0]/tbar[0]/btn[3]").press() ' back
                    Session.FindById("wnd[0]/tbar[0]/btn[3]").press() ' back
                End If
            End If

        Catch ex As Exception
            SF.SetLogInfo("Error Downloading S161 Report " & Box)
        End Try
    End Sub

    Private Sub ReadZFI2Report(ByVal Region As String)
        Dim FileName As String = "S161 Report " & Box & " (" & Date.Now.Month & "-" & Date.Now.Day & "-" & Date.Now.Year & ").xls"
        Dim FilePath As String = EXCELPath & FileName

        Try
            'Read Report
            Dim DT As DataTable = Read_SAP_Report(FilePath)
            If DT Is Nothing Then Exit Sub

            'Delete Old Invoices
            SQL_F.SQL_Execute_NQ(CS, "Delete From P2P_S161 Where Region='" & Region & "'")

            'Delete Upload Today
            SQL_F.SQL_Execute_NQ(CS, "Delete From P2P_S161_UploadToday Where Region='" & Region & "'")

            'Customazing BI Report
            If Not DT.Columns("Column1") Is Nothing Then DT.Columns.Remove("Column1")
            If Not DT.Columns("Date") Is Nothing Then DT.Columns.Remove("Date")
            If Not DT.Columns("Time") Is Nothing Then DT.Columns.Remove("Time")
            DT.Columns("Cty").ColumnName = "Country"
            'If Not DT.Columns("Cty") Is Nothing Then DT.Columns.Remove("Cty")
            If Not DT.Columns("Status") Is Nothing Then DT.Columns.Remove("Status")
            DT.Columns("Account").ColumnName = "Vendor"
            DT.Columns("Vendor name").ColumnName = "Vendor_Name"
            DT.Columns("Code").ColumnName = "LE"
            DT.Columns("Doc. no.").ColumnName = "Doc_Number"
            If Not DT.Columns("Year") Is Nothing Then DT.Columns.Remove("Year")
            DT.Columns("Doc Ty").ColumnName = "Doc_Type"
            If Not DT.Columns("Pst Date") Is Nothing Then DT.Columns.Remove("Pst Date")
            If Not DT.Columns("Clg Date") Is Nothing Then DT.Columns.Remove("Clg Date")
            DT.Columns("Cur").ColumnName = "Currency"
            If Not DT.Columns("LC") Is Nothing Then DT.Columns.Remove("LC")
            If Not DT.Columns("Amt in LC") Is Nothing Then DT.Columns.Remove("Amt in LC")
            If Not DT.Columns("Amt in USD") Is Nothing Then DT.Columns.Remove("Amt in USD")
            If Not DT.Columns("Payments Terms") Is Nothing Then DT.Columns.Remove("Payments Terms")
            If Not DT.Columns("Baseline Date") Is Nothing Then DT.Columns.Remove("Baseline Date")
            If Not DT.Columns("Disc Due Date") Is Nothing Then DT.Columns.Remove("Disc Due Date")
            DT.Columns("Due Date").ColumnName = "Due_Date"
            If Not DT.Columns("Disc Amt") Is Nothing Then DT.Columns.Remove("Disc Amt")
            DT.Columns("Doc. Date").ColumnName = "Doc_Date"
            DT.Columns("PO no.").ColumnName = "Purch_Doc"
            DT.Columns("Purchasing Group").ColumnName = "PGrp"
            DT.Columns("Purchasing Group Name").ColumnName = "PGrp_Description"
            DT.Columns("Purchasing  Org.").ColumnName = "POrg"
            DT.Columns("Plan").ColumnName = "Plant"
            If Not DT.Columns("BOL Number") Is Nothing Then DT.Columns.Remove("BOL Number")
            If Not DT.Columns("MM Document Number") Is Nothing Then DT.Columns.Remove("MM Document Number")
            If Not DT.Columns("N/P") Is Nothing Then DT.Columns.Remove("N/P")
            If Not DT.Columns("I/C") Is Nothing Then DT.Columns.Remove("I/C")
            If Not DT.Columns("Exp") Is Nothing Then DT.Columns.Remove("Exp")
            If Not DT.Columns("mySAP") Is Nothing Then DT.Columns.Remove("mySAP")
            If Not DT.Columns("MM Scan Index") Is Nothing Then DT.Columns.Remove("MM Scan Index")
            If Not DT.Columns("Scan text") Is Nothing Then DT.Columns.Remove("Scan text")
            If Not DT.Columns("Ledger text") Is Nothing Then DT.Columns.Remove("Ledger text")
            If Not DT.Columns("Action Req By") Is Nothing Then DT.Columns.Remove("Action Req By")
            If Not DT.Columns("Reason code") Is Nothing Then DT.Columns.Remove("Reason code")
            If Not DT.Columns("Reason Description") Is Nothing Then DT.Columns.Remove("Reason Description")
            If Not DT.Columns("Multiple R") Is Nothing Then DT.Columns.Remove("Multiple R")
            If Not DT.Columns("TapDate") Is Nothing Then DT.Columns.Remove("TapDate")
            If Not DT.Columns("Modified by") Is Nothing Then DT.Columns.Remove("Modified by")
            If Not DT.Columns("Num  Batch") Is Nothing Then DT.Columns.Remove("Num  Batch")
            If Not DT.Columns("User associate") Is Nothing Then DT.Columns.Remove("User associate")
            If Not DT.Columns("S") Is Nothing Then DT.Columns.Remove("S")
            If Not DT.Columns("SCF sent d") Is Nothing Then DT.Columns.Remove("SCF sent d")
            If Not DT.Columns("S1") Is Nothing Then DT.Columns.Remove("S1")
            If Not DT.Columns("SCF block") Is Nothing Then DT.Columns.Remove("SCF block")
            If Not DT.Columns("Invoice Source") Is Nothing Then DT.Columns.Remove("Invoice Source")
            If Not DT.Columns("Unique Identifier from BES") Is Nothing Then DT.Columns.Remove("Unique Identifier from BES")
            If Not DT.Columns("Supp. Doc") Is Nothing Then DT.Columns.Remove("Supp. Doc")

            'New Columns
            Dim NewColumnBox As New DataColumn("Box", GetType(String))
            NewColumnBox.DefaultValue = Box
            DT.Columns.Add(NewColumnBox)
            DT.Columns.Add("User_Name", GetType(String))
            Dim NewColumnScope As New DataColumn("Scope", GetType(String))
            NewColumnScope.DefaultValue = "TBD"
            DT.Columns.Add(NewColumnScope)
            Dim NewColumnOwner As New DataColumn("Owner", GetType(String))
            NewColumnOwner.DefaultValue = "TBD"
            DT.Columns.Add(NewColumnOwner)
            DT.Columns.Add("Material_Group", GetType(String))
            DT.Columns.Add("Upload_Date", GetType(Date))
            Dim NewColumn As New DataColumn("Region", GetType(String))
            NewColumn.DefaultValue = Region
            DT.Columns.Add(NewColumn)
            DT.Columns.Add("Material", GetType(String))

            'Upload Today Table
            Dim DT_UploadToday As New DataTable
            DT_UploadToday.Columns.Add("Box", GetType(String))
            DT_UploadToday.Columns.Add("Record", GetType(Double))
            DT_UploadToday.Columns.Add("Region", GetType(String))

            'Validate Items
            Try

                For Each R As DataRow In DT.Rows
                    Dim NewUploadTodayRow As DataRow = DT_UploadToday.NewRow

                    'Country
                    Try
                        R("Country") = R("Country").ToString.Trim
                    Catch ex As Exception
                    End Try

                    'Vendor
                    Try
                        R("Vendor") = CDbl(R("Vendor"))
                    Catch ex As Exception
                        R("Vendor") = 0
                    End Try


                    'LE
                    Try
                        R("LE") = CDbl(R("LE"))
                    Catch ex As Exception
                        R("LE") = 0
                    End Try

                    'Doc_Number
                    Try
                        R("Doc_Number") = CDbl(R("Doc_Number"))
                    Catch ex As Exception
                        R("Doc_Number") = 0
                    End Try

                    'Amount
                    Try
                        R("Amount") = CDbl(R("Amount"))
                    Catch ex As Exception
                        R("Amount") = 0
                    End Try


                    'Purch_Doc
                    If Not IsNumeric(R("Purch_Doc")) Then
                        R("Purch_Doc") = 0
                    End If

                    'POrg
                    If Not IsNumeric(R("POrg")) Then
                        R("POrg") = 0
                    End If


                    'Box
                    Try
                        R("Box") = Box
                    Catch ex As Exception
                    End Try

                    'Due_Date
                    If Not IsDate(R("Due_Date")) Then
                        R("Due_Date") = R("User_Name")
                    Else
                        Try
                            Dim Due_Date As Date = CDate(R("Due_Date"))
                            R("Due_Date") = Due_Date.Month & "/" & Due_Date.Day & "/" & Due_Date.Year
                        Catch ex As Exception
                            R("Due_Date") = R("User_Name")
                        End Try

                    End If
              
                    Try
                        NewUploadTodayRow("Box") = R("Box")
                        NewUploadTodayRow("Record") = R("Record")
                        NewUploadTodayRow("Region") = Region
                        DT_UploadToday.Rows.Add(NewUploadTodayRow)
                    Catch ex As Exception
                    End Try

                Next
                DT.AcceptChanges()
            Catch ex As Exception
                SF.SetLogInfo("Error Validating S161 Cases" & Region)
            End Try

          

            'Save Items in SQL
            Try
                Dim BulkInsert As String = SQL_F.Bulk_Insert(CS, "P2P_S161", DT)
                If BulkInsert <> Nothing Then
                    SF.SetLogInfo("Error Saving S161 Report in SQL " & Region)
                End If
            Catch ex As Exception
                SF.SetLogInfo("Error Saving S161 Report in SQL " & Region)
            End Try

            'Save Upload Today Invoices in SQL
            Try
                Dim BulkInsert As String = SQL_F.Bulk_Insert(CS, "P2P_S161_UploadToday", DT_UploadToday)
                If BulkInsert <> Nothing Then
                    SF.SetLogInfo("Error Saving S161 UploadToday in SQL " & Region)
                End If
            Catch ex As Exception
                SF.SetLogInfo("Error Saving S161 UploadToday in SQL " & Region)
            End Try

        Catch ex As Exception
            SF.SetLogInfo("Error Reading S161 Report " & Box)
        End Try
    End Sub

    Public Function DownloadMaterialGroup_S161(ByRef DT_Passwords As DataTable) As Boolean
        Try

            SF.SetLogInfo("Downloading Material Groups - S161 (" & Box & ")")

            'List of Documents to find Material Group Codes
            Dim DT_Documents As New DataTable
            DT_Documents = SQL_F.GetDataTable(CS, "SELECT * FROM P2P_S161 WHERE (Box='" & Box & "') AND ((Material_Group IS NULL) or (Material_Group = '')) ")

            If Not DT_Documents Is Nothing Then

                If DT_Documents.Rows.Count > 0 Then

                    Dim GetTables As New GetSAPTables
                    Dim DT_EKPO As DataTable

                    DT_EKPO = GetTables.Get_EKPO(Box, Login, Password, DT_Documents)

                    If Not DT_EKPO Is Nothing Then

                        Dim EKPO_Lines As DataRow()
                        EKPO_Lines = DT_EKPO.Select("(LineItem = 1) or (LineItem = 10)")

                        For Each Line As DataRow In EKPO_Lines
                            Dim Update As String
                            Dim Material_Group As String = Line("Material_Group")
                            Dim Material As String = Line("Material")
                            Dim User_Name As String = Line("AFNAM")
                            Try
                                Material_Group = Material_Group.Substring(0, 9)
                            Catch ex As Exception
                            End Try
                            Try
                                Material = Material.Substring(0, 8)
                            Catch ex As Exception
                            End Try
                            Try
                                User_Name = User_Name.Substring(0, 6)
                            Catch ex As Exception
                            End Try
                            Update = "Update P2P_S161 set Material_Group ='" & Material_Group & "', Material='" & Material & "', User_Name='" & User_Name & "' "
                            Update += " Where (Box='" & Box & "') AND (Purch_Doc=" & Line("Purch_Doc") & ")"
                            Dim Exc As String = SQL_F.SQL_Execute_NQ(CS, Update)
                            If Exc <> Nothing Then
                            End If
                        Next
                    End If
                End If

            End If

        Catch ex As Exception
            SF.SetLogInfo("Error Downloading Material/Material_Group S161 (" & Box & ")")
        End Try
    End Function

    Public Function Distribute_S161_Items() As Boolean

        SF.SetLogInfo("Applying Variants - S161 (" & Box & ")")

        'Customization
        If SQL_F.SQL_Execute_NQ(CS, "Exec ReAssign_Americas_S161_1_Custom") <> Nothing Then
            SF.SetLogInfo("Error: S161 Distribution: Customization [Exec ReAssign_Americas_S161_1_Custom]")
        End If

        'LogIndNAPD
        If SQL_F.SQL_Execute_NQ(CS, "Exec ReAssign_Americas_S161_2_LogIndNAPD") <> Nothing Then
            SF.SetLogInfo("Error: S161 Distribution: LogIndNAPD [Exec ReAssign_Americas_S161_2_LogIndNAPD]")
        End If

        'SS 01
        If SQL_F.SQL_Execute_NQ(CS, "Exec ReAssign_Americas_S161_3_SS_01") <> Nothing Then
            SF.SetLogInfo("Error: S161 Distribution: SS 01 [Exec ReAssign_Americas_S161_3_SS_01]")
        End If
        'SS 02
        If SQL_F.SQL_Execute_NQ(CS, "Exec ReAssign_Americas_S161_3_SS_02") <> Nothing Then
            SF.SetLogInfo("Error: S161 Distribution: SS 02 [Exec ReAssign_Americas_S161_3_SS_02]")
        End If
        'SS 03
        If SQL_F.SQL_Execute_NQ(CS, "Exec ReAssign_Americas_S161_3_SS_03") <> Nothing Then
            SF.SetLogInfo("Error: S161 Distribution: SS 03 [Exec ReAssign_Americas_S161_3_SS_03]")
        End If
        'SS 04
        If SQL_F.SQL_Execute_NQ(CS, "Exec ReAssign_Americas_S161_3_SS_04") <> Nothing Then
            SF.SetLogInfo("Error: S161 Distribution: SS 04 [Exec ReAssign_Americas_S161_3_SS_04]")
        End If
        'SS 05
        If SQL_F.SQL_Execute_NQ(CS, "Exec ReAssign_Americas_S161_3_SS_05") <> Nothing Then
            SF.SetLogInfo("Error: S161 Distribution: SS 05 [Exec ReAssign_Americas_S161_3_SS_05]")
        End If
        'SS 06
        If SQL_F.SQL_Execute_NQ(CS, "Exec ReAssign_Americas_S161_3_SS_06") <> Nothing Then
            SF.SetLogInfo("Error: S161 Distribution: SS 06 [Exec ReAssign_Americas_S161_3_SS_06]")
        End If
        'SS 07
        If SQL_F.SQL_Execute_NQ(CS, "Exec ReAssign_Americas_S161_3_SS_07") <> Nothing Then
            SF.SetLogInfo("Error: S161 Distribution: SS 07 [Exec ReAssign_Americas_S161_3_SS_07]")
        End If
        'SS 08
        If SQL_F.SQL_Execute_NQ(CS, "Exec ReAssign_Americas_S161_3_SS_08") <> Nothing Then
            SF.SetLogInfo("Error: S161 Distribution: SS 08 [Exec ReAssign_Americas_S161_3_SS_08]")
        End If
        'SS 12
        If SQL_F.SQL_Execute_NQ(CS, "Exec ReAssign_Americas_S161_3_SS_12") <> Nothing Then
            SF.SetLogInfo("Error: S161 Distribution: SS 12 [Exec ReAssign_Americas_S161_3_SS_12]")
        End If
        'SS 13
        If SQL_F.SQL_Execute_NQ(CS, "Exec ReAssign_Americas_S161_3_SS_13") <> Nothing Then
            SF.SetLogInfo("Error: S161 Distribution: SS 13 [Exec ReAssign_Americas_S161_3_SS_13]")
        End If
        'SS 14
        If SQL_F.SQL_Execute_NQ(CS, "Exec ReAssign_Americas_S161_3_SS_14") <> Nothing Then
            SF.SetLogInfo("Error: S161 Distribution: SS 14 [Exec ReAssign_Americas_S161_3_SS_14]")
        End If
        'SS 15
        If SQL_F.SQL_Execute_NQ(CS, "Exec ReAssign_Americas_S161_3_SS_15") <> Nothing Then
            SF.SetLogInfo("Error: S161 Distribution: SS 15 [Exec ReAssign_Americas_S161_3_SS_15]")
        End If
        'SS 16
        If SQL_F.SQL_Execute_NQ(CS, "Exec ReAssign_Americas_S161_3_SS_16") <> Nothing Then
            SF.SetLogInfo("Error: S161 Distribution: SS 16 [Exec ReAssign_Americas_S161_3_SS_16]")
        End If
        'SS 17
        If SQL_F.SQL_Execute_NQ(CS, "Exec ReAssign_Americas_S161_3_SS_17") <> Nothing Then
            SF.SetLogInfo("Error: S161 Distribution: SS 17 [Exec ReAssign_Americas_S161_3_SS_17]")
        End If

        'STR
        If SQL_F.SQL_Execute_NQ(CS, "Exec ReAssign_Americas_S161_4_STR") <> Nothing Then
            SF.SetLogInfo("Error: S161 Distribution: STR [Exec ReAssign_Americas_S161_4_STR]")
        End If

        'Direct Exec ReAssign_Americas_S161_5_Direct
        If SQL_F.SQL_Execute_NQ(CS, "Exec ReAssign_Americas_S161_5_Direct") <> Nothing Then
            SF.SetLogInfo("Error: S161 Distribution: Direct [Exec ReAssign_Americas_S161_5_Direct]")
        End If

        'LogTMS
        If SQL_F.SQL_Execute_NQ(CS, "Exec ReAssign_Americas_S161_6_LogTMS") <> Nothing Then
            SF.SetLogInfo("Error: S161 Distribution: LogTMS [Exec ReAssign_Americas_S161_6_LogTMS]")
        End If

        'LogImpLA (SS) Exec ReAssign_Americas_S161_7_LogImpLA_SS
        If SQL_F.SQL_Execute_NQ(CS, "Exec ReAssign_Americas_S161_7_LogImpLA_SS") <> Nothing Then
            SF.SetLogInfo("Error: BI Distribution: LogImpLA (SS) [Exec ReAssign_Americas_S161_7_LogImpLA_SS]")
        End If

        'LogImpLA (STR) Exec ReAssign_Americas_S161_7_LogImpLA_STR
        If SQL_F.SQL_Execute_NQ(CS, "Exec ReAssign_Americas_S161_7_LogImpLA_STR") <> Nothing Then
            SF.SetLogInfo("Error: S161 Distribution: LogImpLA (STR) [Exec ReAssign_Americas_S161_7_LogImpLA_STR]")
        End If

        'LogImpNA
        If SQL_F.SQL_Execute_NQ(CS, "Exec ReAssign_Americas_S161_8_LogImpNA") <> Nothing Then
            SF.SetLogInfo("Error: S161 Distribution: LogImpNA [Exec ReAssign_Americas_S161_8_LogImpNA]")
        End If

        'LogDirect
        If SQL_F.SQL_Execute_NQ(CS, "Exec ReAssign_Americas_S161_9_LogDirect") <> Nothing Then
            SF.SetLogInfo("Error: S161 Distribution: LogImpNA [Exec ReAssign_Americas_S161_9_LogDirect]")
        End If

        'SS EndUser
        If SQL_F.SQL_Execute_NQ(CS, "Exec ReAssign_Americas_S161_10_SS_EndUser") <> Nothing Then
            SF.SetLogInfo("Error: S161 Distribution: SS EndUser [Exec ReAssign_Americas_S161_10_SS_EndUser]")
        End If

        'STR TopSuppliers
        If SQL_F.SQL_Execute_NQ(CS, "Exec ReAssign_Americas_S161_11_STR_TopSuppliers") <> Nothing Then
            SF.SetLogInfo("Error: S161 Distribution: STR TopSuppliers [Exec ReAssign_Americas_S161_11_STR_TopSuppliers]")
        End If

        'OutOfScope
        If SQL_F.SQL_Execute_NQ(CS, "Exec ReAssign_Americas_S161_12_OutOfScope") <> Nothing Then
            SF.SetLogInfo("Error: S161 Distribution: POrgs OutOfScope [Exec ReAssign_Americas_S161_12_OutOfScope]")
        End If

        'Delete Excluded POrgs
        If SQL_F.SQL_Execute_NQ(CS, "Delete From P2P_S161 Where POrg in (Select POrg From P2P_S161_ExcludedPOrgs)") <> Nothing Then
            SF.SetLogInfo("Error: S161 Distribution: Delete Excluded POrgs [Delete From P2P_S161 Where POrg in (Select POrg From P2P_S161_ExcludedPOrgs)]")
        End If

        'Delete Excluded Countries
        If SQL_F.SQL_Execute_NQ(CS, "DELETE FROM P2P_S161 WHERE (Country NOT IN (SELECT Country FROM P2P_S161_IncludedCountries))") <> Nothing Then
            SF.SetLogInfo("Error: S161 Distribution: Delete Excluded Countries [DELETE FROM P2P_S161 WHERE (Country NOT IN (SELECT Country FROM P2P_S161_IncludedCountries))]")
        End If

        Distribute_S161_Items = True
    End Function

    'IS Report **********************************************************************************

    Public Sub KillInternetExplorer()
        Try
            Dim InternetExplorer() As Process = Process.GetProcessesByName("iexplore")
            '[KILL ALL INTERNET EXPLORER PROCESSES]
            For Each Process As Process In InternetExplorer
                Process.Kill()
            Next
        Catch ex As Exception

        End Try
    End Sub

    Public Function Download_IS_Report(ByVal Region As String) As DataTable

        SF.SetLogInfo("Downloading IS Report (" & Region & ")")

        Try
            Dim BW As BWConnect.BW
            Dim DT As New DataTable

            BW = New BWConnect.BW(Return_Bookmark(Region), Me.Login, Me.Password)
            DT = bw.DownloadData()
            BW.close()

            Download_IS_Report = DT
        Catch ex As Exception
            Download_IS_Report = Nothing
        End Try

    End Function

    Private Function Return_Bookmark(ByVal ID As String) As String
        Try
            Dim Bookmark As String = SQL_F.ReturnDescription(CS, "P2P_S161_Bookmarks", "Bookmark", "ID ='" & ID & "'")
            If Bookmark <> "" Then
                Return_Bookmark = Bookmark
            Else
                Return_Bookmark = ""
            End If
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Public Sub Read_IS_Report(ByVal Region As String, ByRef DT As DataTable)

        SF.SetLogInfo("Reading IS Report (" & Region & ")")

        Try

            'Delete Old Invoices
            SQL_F.SQL_Execute_NQ(CS, "Delete From P2P_IS Where Region='" & Region & "'")

            'Delete Upload Today
            SQL_F.SQL_Execute_NQ(CS, "Delete From P2P_IS_UploadToday Where Region='" & Region & "'")

            'Customazing IS Report
            DT.Columns("Document Number").ColumnName = "Doc_Number"
            DT.Columns("Ledger number").ColumnName = "Ledger_Number"
            DT.Columns("Vendor name").ColumnName = "Vendor_Name"
            DT.Columns("Workflow Status").ColumnName = "Workflow_Status"
            DT.Columns("Days Untouched").ColumnName = "Days_Untouched"
            DT.Columns("Country Name").ColumnName = "Country_Name"
            DT.Columns("Plant Name").ColumnName = "Plant_Name"
            DT.Columns("Purchase Order").ColumnName = "Purch_Doc"
            DT.Columns("PO requisitioner*New*").ColumnName = "User_Name"
            DT.Columns("Overdue / Not Overdue").ColumnName = "Status"
            DT.Columns("Workflow Status Category").ColumnName = "Workflow_Status_Category"
            DT.Columns("Purchasing Org").ColumnName = "POrg"
            DT.Columns("Purchasing Org Name").ColumnName = "POrg_Description"
            DT.Columns("Invoice Type").ColumnName = "Doc_Type"
            DT.Columns("Purchasing Group").ColumnName = "PGrp"
            DT.Columns("Purchasing Group Name").ColumnName = "PGrp_Description"
            If Not DT.Columns("OI Aging cluster") Is Nothing Then DT.Columns.Remove("OI Aging cluster")
            DT.Columns("Due Date").ColumnName = "Due_Date"
            DT.Columns("Amount  in LC").ColumnName = "Amount"

            'New Columns
            Dim NewColumnBox As New DataColumn("Box", GetType(String))
            NewColumnBox.DefaultValue = Box
            DT.Columns.Add(NewColumnBox)
            Dim NewColumnLE As New DataColumn("LE", GetType(Double))
            NewColumnLE.DefaultValue = 0
            DT.Columns.Add(NewColumnLE)
            Dim NewColumnScope As New DataColumn("Scope", GetType(String))
            NewColumnScope.DefaultValue = "TBD"
            DT.Columns.Add(NewColumnScope)
            Dim NewColumnOwner As New DataColumn("Owner", GetType(String))
            NewColumnOwner.DefaultValue = "TBD"
            DT.Columns.Add(NewColumnOwner)
            DT.Columns.Add("Material_Group", GetType(String))
            DT.Columns.Add("Upload_Date", GetType(Date))
            Dim NewColumn As New DataColumn("Region", GetType(String))
            NewColumn.DefaultValue = Region
            DT.Columns.Add(NewColumn)
            DT.Columns.Add("Material", GetType(String))

            'Upload Today Table
            Dim DT_UploadToday As New DataTable
            DT_UploadToday.Columns.Add("Box", GetType(String))
            DT_UploadToday.Columns.Add("LE", GetType(Double))
            DT_UploadToday.Columns.Add("Purch_Doc", GetType(Double))
            DT_UploadToday.Columns.Add("Doc_Number", GetType(Double))
            DT_UploadToday.Columns.Add("Ledger_Number", GetType(Double))
            DT_UploadToday.Columns.Add("Region", GetType(String))

            'Validate Items
            Try

                For Each R As DataRow In DT.Rows

                    Dim NewUploadTodayRow As DataRow = DT_UploadToday.NewRow
                    Dim Split_DocNumber As String() = R("Doc_Number").ToString.Split(" ")
                    Dim Split_PurchDoc As String() = R("Purch_Doc").ToString.Split("/")
                    Dim Split_DueDate As String() = R("Due_Date").ToString.Split(".")

                    'Doc_Number
                    Try
                        R("Doc_Number") = Split_DocNumber(2)
                        R("Doc_Number") = CDbl(R("Doc_Number"))
                    Catch ex As Exception
                        R("Doc_Number") = 0
                    End Try

                    'Ledger_Number
                    Try
                        R("Ledger_Number") = CDbl(R("Ledger_Number"))
                    Catch ex As Exception
                        R("Ledger_Number") = 0
                    End Try

                    'Vendor
                    Try
                        R("Vendor") = CDbl(R("Vendor"))
                    Catch ex As Exception
                        R("Vendor") = 0
                    End Try

                    'Workflow_Status
                    Try
                        R("Workflow_Status") = CDbl(R("Workflow_Status"))
                    Catch ex As Exception
                        R("Workflow_Status") = 0
                    End Try

                    'Days_Untouched
                    Try
                        R("Days_Untouched") = CDbl(R("Days_Untouched"))
                    Catch ex As Exception
                        R("Days_Untouched") = 0
                    End Try

                    'Country
                    Try
                        R("Country") = R("Country").ToString.Trim
                    Catch ex As Exception
                    End Try

                    'Purch_Doc
                    Try
                        R("Purch_Doc") = Split_PurchDoc(1)
                        R("Purch_Doc") = CDbl(R("Purch_Doc"))
                    Catch ex As Exception
                        R("Purch_Doc") = 0
                    End Try

                    'POrg
                    Try
                        If Not IsNumeric(R("POrg")) Then
                            R("POrg") = 0
                        End If
                    Catch ex As Exception
                        R("POrg") = 0
                    End Try

                    'PGrp
                    Try
                        R("PGrp") = R("PGrp").ToString.Trim.Substring(0, 3)
                    Catch ex As Exception
                        R("PGrp") = ""
                    End Try

                    'Due_Date

                    Try
                        R("Due_Date") = Split_DueDate(1) & "/" & Split_DueDate(0) & "/" & Split_DueDate(2) & " 00:00:00"
                        R("Due_Date") = CDate(R("Due_Date"))
                    Catch ex As Exception
                        R("Due_Date") = R("Material_Group")
                    End Try

                    'Box
                    Try
                        R("Box") = Split_DocNumber(0).Substring(0, 3)
                    Catch ex As Exception
                        R("Box") = ""
                    End Try

                    'LE
                    Try
                        R("LE") = Split_DocNumber(1)
                        R("LE") = CDbl(R("LE"))
                    Catch ex As Exception
                        R("LE") = 0
                    End Try

                    Try
                        NewUploadTodayRow("Box") = R("Box")
                        NewUploadTodayRow("LE") = R("LE")
                        NewUploadTodayRow("Purch_Doc") = R("Purch_Doc")
                        NewUploadTodayRow("Doc_Number") = R("Doc_Number")
                        NewUploadTodayRow("Ledger_Number") = R("Ledger_Number")
                        NewUploadTodayRow("Region") = Region
                        DT_UploadToday.Rows.Add(NewUploadTodayRow)
                    Catch ex As Exception
                    End Try

                Next

                'Delete Last Lines
                DT.Rows((DT.Rows.Count) - 1).Delete()
                DT.AcceptChanges()
                DT_UploadToday.Rows((DT_UploadToday.Rows.Count) - 1).Delete()
                DT_UploadToday.AcceptChanges()

            Catch ex As Exception
                SF.SetLogInfo("Error Validating S161 Cases" & Region)
            End Try

            'Save Items in SQL
            Try
                Dim BulkInsert As String = SQL_F.Bulk_Insert(CS, "P2P_IS", DT)
                If BulkInsert <> Nothing Then
                    SF.SetLogInfo("Error Saving IS Report in SQL " & Region)
                End If
            Catch ex As Exception
                SF.SetLogInfo("Error Saving IS Report in SQL " & Region)
            End Try

            'Save Upload Today Invoices in SQL
            Try
                Dim BulkInsert As String = SQL_F.Bulk_Insert(CS, "P2P_IS_UploadToday", DT_UploadToday)
                If BulkInsert <> Nothing Then
                    SF.SetLogInfo("Error Saving IS UploadToday in SQL " & Region)
                End If
            Catch ex As Exception
                SF.SetLogInfo("Error Saving IS UploadToday in SQL " & Region)
            End Try


        Catch ex As Exception
            SF.SetLogInfo("Error Reading IS Report " & Region)
        End Try

    End Sub

    Public Function DownloadMaterialGroup_IS(ByRef DT_Passwords As DataTable, ByVal Region As String) As Boolean
        Try

            SF.SetLogInfo("Downloading Material Groups - IS (" & Region & ")")

            'Disctict SAP Boxes
            Dim DT_SAP_Box As New DataTable
            DT_SAP_Box = SQL_F.GetDataTable(CS, "Select * From P2P_SAP_Boxes where Download_MatGroup =1")

            'For each SAP Box
            For Each R_SAP_Box As DataRow In DT_SAP_Box.Rows


                If R_SAP_Box("Box") = "ANP" Then
                    R_SAP_Box("Box") = R_SAP_Box("Box") & "_" & R_SAP_Box("Client")
                End If

                'List of Documents to find Material Group Codes
                Dim DT_Documents As New DataTable
                DT_Documents = SQL_F.GetDataTable(CS, "SELECT * FROM P2P_IS WHERE (Box='" & R_SAP_Box("Box") & "') AND ((Material_Group IS NULL) or (Material_Group = '')) AND (Region='" & Region & "') ")

                If Not DT_Documents Is Nothing Then

                    If DT_Documents.Rows.Count > 0 Then

                        Dim GetTables As New GetSAPTables
                        Dim DT_EKPO As DataTable

                        DT_EKPO = GetTables.Get_EKPO(R_SAP_Box("Box"), R_SAP_Box("Login"), R_SAP_Box("Password"), DT_Documents)

                        If Not DT_EKPO Is Nothing Then

                            Dim EKPO_Lines As DataRow()
                            EKPO_Lines = DT_EKPO.Select("(LineItem = 1) or (LineItem = 10)")

                            For Each Line As DataRow In EKPO_Lines
                                Dim Update As String
                                Dim Material_Group As String = Line("Material_Group")
                                Dim Material As String = Line("Material")
                                Dim User_Name As String = Line("AFNAM")
                                Try
                                    Material_Group = Material_Group.Substring(0, 9)
                                Catch ex As Exception
                                End Try
                                Try
                                    Material = Material.Substring(0, 8)
                                Catch ex As Exception
                                End Try
                                Try
                                    User_Name = User_Name.Substring(0, 6)
                                Catch ex As Exception
                                End Try
                                Update = "Update P2P_IS set Material_Group ='" & Material_Group & "', Material='" & Material & "', User_Name='" & User_Name & "' "
                                Update += " Where (Box='" & R_SAP_Box("Box") & "') AND (Purch_Doc=" & Line("Purch_Doc") & ")"
                                Dim Exc As String = SQL_F.SQL_Execute_NQ(CS, Update)
                                If Exc <> Nothing Then
                                End If
                            Next
                        End If
                    End If

                End If
            Next

        Catch ex As Exception
            SF.SetLogInfo("Error Downloading Material/Material_Group BI")
        End Try
    End Function

    Public Function Distribute_IS_Items() As Boolean

        SF.SetLogInfo("Applying Variants - IS (" & Box & ")")

        'Customization
        If SQL_F.SQL_Execute_NQ(CS, "Exec ReAssign_Americas_IS_1_Custom") <> Nothing Then
            SF.SetLogInfo("Error: IS Distribution: Customization [Exec ReAssign_Americas_IS_1_Custom]")
        End If

        'LogIndNAPD
        If SQL_F.SQL_Execute_NQ(CS, "Exec ReAssign_Americas_IS_2_LogIndNAPD") <> Nothing Then
            SF.SetLogInfo("Error: IS Distribution: LogIndNAPD [Exec ReAssign_Americas_IS_2_LogIndNAPD]")
        End If

        'SS 01
        If SQL_F.SQL_Execute_NQ(CS, "Exec ReAssign_Americas_IS_3_SS_01") <> Nothing Then
            SF.SetLogInfo("Error: IS Distribution: SS 01 [Exec ReAssign_Americas_IS_3_SS_01]")
        End If
        'SS 02
        If SQL_F.SQL_Execute_NQ(CS, "Exec ReAssign_Americas_IS_3_SS_02") <> Nothing Then
            SF.SetLogInfo("Error: IS Distribution: SS 02 [Exec ReAssign_Americas_IS_3_SS_02]")
        End If
        'SS 03
        If SQL_F.SQL_Execute_NQ(CS, "Exec ReAssign_Americas_IS_3_SS_03") <> Nothing Then
            SF.SetLogInfo("Error: IS Distribution: SS 03 [Exec ReAssign_Americas_IS_3_SS_03]")
        End If
        'SS 04
        If SQL_F.SQL_Execute_NQ(CS, "Exec ReAssign_Americas_IS_3_SS_04") <> Nothing Then
            SF.SetLogInfo("Error: IS Distribution: SS 04 [Exec ReAssign_Americas_IS_3_SS_04]")
        End If
        'SS 05
        If SQL_F.SQL_Execute_NQ(CS, "Exec ReAssign_Americas_IS_3_SS_05") <> Nothing Then
            SF.SetLogInfo("Error: IS Distribution: SS 05 [Exec ReAssign_Americas_IS_3_SS_05]")
        End If
        'SS 06
        If SQL_F.SQL_Execute_NQ(CS, "Exec ReAssign_Americas_IS_3_SS_06") <> Nothing Then
            SF.SetLogInfo("Error: IS Distribution: SS 06 [Exec ReAssign_Americas_IS_3_SS_06]")
        End If
        'SS 07
        If SQL_F.SQL_Execute_NQ(CS, "Exec ReAssign_Americas_IS_3_SS_07") <> Nothing Then
            SF.SetLogInfo("Error: IS Distribution: SS 07 [Exec ReAssign_Americas_IS_3_SS_07]")
        End If
        'SS 08
        If SQL_F.SQL_Execute_NQ(CS, "Exec ReAssign_Americas_IS_3_SS_08") <> Nothing Then
            SF.SetLogInfo("Error: IS Distribution: SS 08 [Exec ReAssign_Americas_IS_3_SS_08]")
        End If
        'SS 12
        If SQL_F.SQL_Execute_NQ(CS, "Exec ReAssign_Americas_IS_3_SS_12") <> Nothing Then
            SF.SetLogInfo("Error: IS Distribution: SS 12 [Exec ReAssign_Americas_IS_3_SS_12]")
        End If
        'SS 13
        If SQL_F.SQL_Execute_NQ(CS, "Exec ReAssign_Americas_IS_3_SS_13") <> Nothing Then
            SF.SetLogInfo("Error: IS Distribution: SS 13 [Exec ReAssign_Americas_IS_3_SS_13]")
        End If
        'SS 14
        If SQL_F.SQL_Execute_NQ(CS, "Exec ReAssign_Americas_IS_3_SS_14") <> Nothing Then
            SF.SetLogInfo("Error: IS Distribution: SS 14 [Exec ReAssign_Americas_IS_3_SS_14]")
        End If
        'SS 15
        If SQL_F.SQL_Execute_NQ(CS, "Exec ReAssign_Americas_IS_3_SS_15") <> Nothing Then
            SF.SetLogInfo("Error: IS Distribution: SS 15 [Exec ReAssign_Americas_IS_3_SS_15]")
        End If
        'SS 16
        If SQL_F.SQL_Execute_NQ(CS, "Exec ReAssign_Americas_IS_3_SS_16") <> Nothing Then
            SF.SetLogInfo("Error: IS Distribution: SS 16 [Exec ReAssign_Americas_IS_3_SS_16]")
        End If
        'SS 17
        If SQL_F.SQL_Execute_NQ(CS, "Exec ReAssign_Americas_IS_3_SS_17") <> Nothing Then
            SF.SetLogInfo("Error: IS Distribution: SS 17 [Exec ReAssign_Americas_IS_3_SS_17]")
        End If

        'STR
        If SQL_F.SQL_Execute_NQ(CS, "Exec ReAssign_Americas_IS_4_STR") <> Nothing Then
            SF.SetLogInfo("Error: IS Distribution: STR [Exec ReAssign_Americas_IS_4_STR]")
        End If

        'Direct Exec ReAssign_Americas_IS_5_Direct
        If SQL_F.SQL_Execute_NQ(CS, "Exec ReAssign_Americas_IS_5_Direct") <> Nothing Then
            SF.SetLogInfo("Error: IS Distribution: Direct [Exec ReAssign_Americas_IS_5_Direct]")
        End If

        'LogTMS
        If SQL_F.SQL_Execute_NQ(CS, "Exec ReAssign_Americas_IS_6_LogTMS") <> Nothing Then
            SF.SetLogInfo("Error: IS Distribution: LogTMS [Exec ReAssign_Americas_IS_6_LogTMS]")
        End If

        'LogImpLA (SS) Exec ReAssign_Americas_IS_7_LogImpLA_SS
        If SQL_F.SQL_Execute_NQ(CS, "Exec ReAssign_Americas_IS_7_LogImpLA_SS") <> Nothing Then
            SF.SetLogInfo("Error: BI Distribution: LogImpLA (SS) [Exec ReAssign_Americas_IS_7_LogImpLA_SS]")
        End If

        'LogImpLA (STR) Exec ReAssign_Americas_IS_7_LogImpLA_STR
        If SQL_F.SQL_Execute_NQ(CS, "Exec ReAssign_Americas_IS_7_LogImpLA_STR") <> Nothing Then
            SF.SetLogInfo("Error: IS Distribution: LogImpLA (STR) [Exec ReAssign_Americas_IS_7_LogImpLA_STR]")
        End If

        'LogImpNA
        If SQL_F.SQL_Execute_NQ(CS, "Exec ReAssign_Americas_IS_8_LogImpNA") <> Nothing Then
            SF.SetLogInfo("Error: IS Distribution: LogImpNA [Exec ReAssign_Americas_IS_8_LogImpNA]")
        End If

        'LogDirect
        If SQL_F.SQL_Execute_NQ(CS, "Exec ReAssign_Americas_IS_9_LogDirect") <> Nothing Then
            SF.SetLogInfo("Error: IS Distribution: LogImpNA [Exec ReAssign_Americas_IS_9_LogDirect]")
        End If

        'SS EndUser
        If SQL_F.SQL_Execute_NQ(CS, "Exec ReAssign_Americas_IS_10_SS_EndUser") <> Nothing Then
            SF.SetLogInfo("Error: IS Distribution: SS EndUser [Exec ReAssign_Americas_IS_10_SS_EndUser]")
        End If

        'STR TopSuppliers
        If SQL_F.SQL_Execute_NQ(CS, "Exec ReAssign_Americas_IS_11_STR_TopSuppliers") <> Nothing Then
            SF.SetLogInfo("Error: IS Distribution: STR TopSuppliers [Exec ReAssign_Americas_IS_11_STR_TopSuppliers]")
        End If

        'OutOfScope
        If SQL_F.SQL_Execute_NQ(CS, "Exec ReAssign_Americas_IS_12_OutOfScope") <> Nothing Then
            SF.SetLogInfo("Error: IS Distribution: POrgs OutOfScope [Exec ReAssign_Americas_IS_12_OutOfScope]")
        End If

        'Delete Invoices in BI RawData
        If SQL_F.SQL_Execute_NQ(CS, "Delete From P2P_BI Where Doc_Number in (Select Doc_Number from P2P_IS)") <> Nothing Then
            SF.SetLogInfo("Error: IS Distribution: Delete Invoices in BI RawData [Delete From P2P_BI Where Doc_Number in (Select Doc_Number from P2P_IS)]")
        End If

     

        Distribute_IS_Items = True
    End Function

End Class
