Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop
Imports System.Data


Public Class MSExcel

    Public Function ExportToExcel(ByVal pTable As System.Data.DataTable) As Boolean
        Dim xlApp As Excel.Application
        Dim xlBook As Excel.Workbook
        Dim xlSheet As Excel.Worksheet
        Dim oQueryTable As Excel.QueryTable
        Dim rs As ADODB.Recordset

        Try


            xlApp = CreateObject("Excel.Application")

            xlApp.UserControl = True
            xlBook = xlApp.Workbooks.Add
            xlSheet = xlBook.Worksheets(1)

            rs = ConvertToRecordset(pTable)

            oQueryTable = xlSheet.QueryTables.Add(rs, xlSheet.Cells(1, 1))
            oQueryTable.Refresh()

            xlApp.Visible = True
            ' rs.Close()
            rs = Nothing

            Return True

        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Function ExportToExcel(ByVal pTable As System.Data.DataTable, ByVal pFilePath As String) As Boolean
        Dim xlApp As Excel.Application
        Dim xlBook As Excel.Workbook
        Dim xlSheet As Excel.Worksheet
        Dim oQueryTable As Excel.QueryTable
        Dim rs As ADODB.Recordset


        Try

            If Len(Dir(pFilePath)) > 0 Then
                Kill(pFilePath)
            End If

            xlApp = CreateObject("Excel.Application")

            xlApp.UserControl = True
            xlBook = xlApp.Workbooks.Add
            xlSheet = xlBook.Worksheets(1)

            rs = ConvertToRecordset(pTable)

            oQueryTable = xlSheet.QueryTables.Add(rs, xlSheet.Cells(1, 1))
            oQueryTable.Refresh()

            xlApp.ActiveWorkbook.SaveAs(pFilePath)

            xlApp.ActiveWorkbook.Close()
            xlApp.Quit()
            'xlApp.Visible = True
            ' rs.Close()
            Return True
        Catch ex As Exception

            Return False
        Finally

            rs = Nothing
        End Try


    End Function

    Public Shared Sub DataTableToRange(ByVal anchorCell As Excel.Range, _
    ByVal tableToCopy As System.Data.DataTable, _
    Optional ByVal tableHeader As String = "")

        If tableHeader <> "" Then
            Try
                anchorCell.Value = tableHeader
                anchorCell = anchorCell.Offset(1, 0)
            Catch ex As Exception
            End Try
        End If

        Dim tableHeaderOffset As Integer = 0

        For Each loopHeaders As DataColumn In tableToCopy.Columns
            Try
                anchorCell.Offset(0, tableHeaderOffset).Value = loopHeaders.ColumnName
            Catch ex As Exception
            End Try

            tableHeaderOffset += 1

        Next

        anchorCell.Offset(1, 0).CopyFromRecordset(ConvertToRecordset(tableToCopy))

    End Sub

    Public Shared Function ConvertToRecordset(ByVal inTable As System.Data.DataTable) As ADODB.Recordset

        Dim result As ADODB.Recordset = New ADODB.Recordset()
        result.CursorLocation = ADODB.CursorLocationEnum.adUseClient

        Dim resultFields As ADODB.Fields = result.Fields
        Dim inColumns As System.Data.DataColumnCollection = inTable.Columns

        For Each inColumn As DataColumn In inColumns
            resultFields.Append(inColumn.ColumnName, _
                TranslateType(inColumn.DataType), _
                inColumn.MaxLength, _
                ADODB.FieldAttributeEnum.adFldIsNullable, _
                Nothing)
        Next

        result.Open(System.Reflection.Missing.Value _
                , System.Reflection.Missing.Value _
                , ADODB.CursorTypeEnum.adOpenStatic _
                , ADODB.LockTypeEnum.adLockOptimistic)

        For Each dr As DataRow In inTable.Rows
            result.AddNew(System.Reflection.Missing.Value, _
                      System.Reflection.Missing.Value)

            For columnIndex As Integer = 0 To inColumns.Count - 1
                If Not DBNull.Value.Equals(dr(columnIndex)) Then
                    resultFields(columnIndex).Value = Replace(dr(columnIndex), "�", "")
                Else
                    resultFields(columnIndex).Value = dr(columnIndex)
                End If
            Next
        Next

        Return result
    End Function

    Shared Function TranslateType(ByVal columnType As Type) As ADODB.DataTypeEnum
        Select Case columnType.UnderlyingSystemType.ToString()

            Case "System.Boolean"
                Return ADODB.DataTypeEnum.adBoolean

            Case "System.Byte"
                Return ADODB.DataTypeEnum.adUnsignedTinyInt

            Case "System.Char"
                Return ADODB.DataTypeEnum.adChar

            Case "System.DateTime"
                Return ADODB.DataTypeEnum.adDate

            Case "System.Decimal"
                Return ADODB.DataTypeEnum.adCurrency

            Case "System.Double"
                Return ADODB.DataTypeEnum.adDouble

            Case "System.Int16"
                Return ADODB.DataTypeEnum.adSmallInt

            Case "System.Int32"
                Return ADODB.DataTypeEnum.adInteger

            Case "System.Int64"
                Return ADODB.DataTypeEnum.adBigInt

            Case "System.SByte"
                Return ADODB.DataTypeEnum.adTinyInt

            Case "System.Single"
                Return ADODB.DataTypeEnum.adSingle

            Case "System.UInt16"
                Return ADODB.DataTypeEnum.adUnsignedSmallInt

            Case "System.UInt32"
                Return ADODB.DataTypeEnum.adUnsignedInt

            Case "System.UInt64"
                Return ADODB.DataTypeEnum.adUnsignedBigInt

        End Select

        'Note Strings are not cased and will return here:
        Return ADODB.DataTypeEnum.adVarChar

    End Function

End Class

'Add Reference : On the COM tab, select Microsoft ActiveX Data Objects 2.5 Library. 
'Example:
'Try
'Dim Exc As New MSExcel
'            If Exc.ExportToExcel(MBS.DataSource, My.Computer.FileSystem.SpecialDirectories.Desktop & "\BI Report (" & Date.Now.Month & "-" & Date.Now.Day & "-" & Date.Now.Year & ").xlsx") Then
'                MsgBox("Done: The Report has been created on your desk", MsgBoxStyle.Information)
'            Else
'                MsgBox("Error To Create Report, Please Contact The System Administrator", MsgBoxStyle.Critical)
'            End If
'        Catch ex As Exception

'        End Try
