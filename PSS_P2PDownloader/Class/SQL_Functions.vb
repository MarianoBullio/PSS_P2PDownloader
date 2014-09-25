'**************************************
' Module: SQL_Functions
' Created By: Mariano Bullio, bullio.m@pg.com
'**************************************

Imports Microsoft.VisualBasic
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.Windows.Forms
Imports System.Drawing
Imports System.Data
Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Reflection

Imports System.Globalization
Imports System.Threading


Public Class SQL_Functions

    Public Function SQL_Execute_SC(ByVal ConnString As String, ByVal sqlCommand As String) As Object

        Dim Connection As SqlConnection = New SqlConnection(ConnString)
        Dim Command As New SqlCommand(sqlCommand, Connection)
        SQL_Execute_SC = Nothing
        Try
            Connection.Open()
            SQL_Execute_SC = Command.ExecuteScalar()
            Connection.Close()
        Catch ex As Exception
            SQL_Execute_SC = Nothing
        End Try

    End Function

    Public Function SQL_Execute_NQ(ByVal ConnString As String, ByVal sqlCommand As String) As String

        Dim Connection As SqlConnection = New SqlConnection(ConnString)
        Dim Command As New SqlCommand(sqlCommand, Connection)

        SQL_Execute_NQ = Nothing
        Try
            Connection.Open()
            Command.ExecuteNonQuery()
            Connection.Close()
        Catch ex As Exception
            If ex.Message.Contains("Time") Then
                Dim a As String = ""
                SQL_Execute_NQ(ConnString, sqlCommand)
            End If
            SQL_Execute_NQ = ex.Message
        End Try

    End Function

    Public Function Update_SQL_Table(ByRef DT As DataTable, ByVal CnString As String) As Boolean

        Dim DA As New SqlDataAdapter
        Dim DS As New DataSet
        Dim Con As New SqlConnection(CnString)
        Dim TN As String = DT.TableName
        DA.SelectCommand = New SqlCommand("SELECT * FROM " & DT.TableName, Con)
        Update_SQL_Table = False
        Try
            Con.Open()
            DA.Fill(DS)
            DA.FillSchema(DS, SchemaType.Source)
            Dim B As New SqlCommandBuilder(DA)
            B.GetUpdateCommand()
            B.GetDeleteCommand()
            B.GetInsertCommand()
            DS.Tables(0).Merge(DT)
            DA.Update(DS)
            DT.AcceptChanges()
            Update_SQL_Table = True
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Function

    Public Function GetSQLTable(ByVal ConnString As String, ByVal TableName As String, Optional ByVal FilterExp As String = Nothing) As DataTable

        GetSQLTable = Nothing

        Dim Connection As SqlConnection = New SqlConnection(ConnString)
        Dim Command As New SqlCommand("Select * From " & TableName, Connection)

        If Not FilterExp Is Nothing Then
            If FilterExp <> "" Then
                Command.CommandText = Command.CommandText & " WHERE " & FilterExp
            End If
        End If

        Dim Adapter As SqlDataAdapter = New SqlDataAdapter()
        Dim Table As New DataTable
        Adapter.SelectCommand = Command
        Try
            Adapter.Fill(Table)
            Adapter.FillSchema(Table, SchemaType.Source)
            Table.TableName = TableName
            GetSQLTable = Table
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Function

    Public Function GetDataTable(ByVal ConnString As String, ByVal sqlCommand As String) As DataTable

        Dim Connection As SqlConnection = New SqlConnection(ConnString)
        Dim Command As New SqlCommand(sqlCommand, Connection)
        Dim Adapter As SqlDataAdapter = New SqlDataAdapter()
        Dim Table As New DataTable

        Adapter.SelectCommand = Command
        Try
            Adapter.Fill(Table)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Return Table

    End Function

    Public Function LinQToDataTable(Of T)(ByVal source As IEnumerable(Of T)) As DataTable
        Return New ObjectShredder(Of T)().Shred(source, Nothing, Nothing)
    End Function

    Public Function ReturnDescription(ByVal CS As String, ByVal table As String, ByVal column As String, ByVal Filter As String) As String
        Dim Description As String = ""
        Dim SF As New SQL_Functions
        Try
            Dim SE As String = Filter & " and " & column & " is not null"
            Dim DT As DataTable = SF.GetSQLTable(CS, table, SE)
            Dim R As DataRow

            If DT.Rows.Count = 1 Then
                R = DT.Rows(0)
                Description = R(column).ToString()
            End If
        Catch
            Description = ""
        End Try
        Return Description
    End Function

    Public Function ReturnDescriptionError(ByVal CS As String, ByVal table As String, ByVal column As String, ByVal Filter As String) As String
        Dim Description As String = ""
        Dim SF As New SQL_Functions
        Try
            Dim SE As String = Filter & " and " & column & " is not null"
            Dim DT As DataTable = SF.GetSQLTable(CS, table, SE)
            Dim R As DataRow

            If DT.Rows.Count = 1 Then
                R = DT.Rows(0)
                Description = R(column).ToString()
            End If
        Catch
            Description = "(E99) Undefined Issue"
        End Try
        Return Description
    End Function

    Public Function Return_TotalRows(ByVal DT As DataTable) As String

        Dim Count As String = "?"
        Try

            For Each DR As DataRow In DT.Rows
                Count = DR(0)
            Next
            Return_TotalRows = Count
        Catch ex As Exception
            Return_TotalRows = "?"
        End Try

    End Function

    Public Function Return_TotalRows(ByVal table As String) As String

        Dim DT As DataTable
        Dim Count As String = "?"
        Try
            DT = GetDataTable(CS, "Select Total From " & table)

            For Each DR As DataRow In DT.Rows
                Count = DR(0)
            Next
            Return_TotalRows = Count
        Catch ex As Exception
            Return_TotalRows = "?"
        End Try

    End Function

    Public Sub TableToClipboard(ByVal DT As DataTable)

        Dim S As String = Nothing
        Dim DR As DataRow

        My.Computer.Clipboard.Clear()
        For Each DR In DT.Rows
            S = S & DR(0) & Chr(13) & Chr(10)
        Next
        My.Computer.Clipboard.SetText(S)

    End Sub

    Public Sub ClipboardToTable(ByRef DT As DataTable)

        Dim sText As String
        Dim sLines() As String
        Dim sColValues() As String
        Dim DR As DataRow
        Dim I As Integer

        sText = My.Computer.Clipboard.GetText()
        sLines = sText.Split(New String() {vbCrLf}, StringSplitOptions.RemoveEmptyEntries)
        For Each sLine As String In sLines
            sColValues = sLine.Split(vbTab)
            DR = DT.NewRow
            Try
                For I = 0 To sColValues.Length - 1
                    DR(I) = sColValues(I)
                Next
                DT.Rows.Add(DR)
            Catch ex As Exception
            End Try
        Next

    End Sub

    Public Function Bulk_Insert(ByVal CS As String, ByVal DestinationTable As String, ByVal Data As DataTable) As String
        Bulk_Insert = Nothing
        Using Connection As New SqlConnection(CS)
            Dim BI As New SqlBulkCopy(Connection)

            'Dim SF As New System_Functions
            'SF.ExportDataTableToExcel(Data, "C:\Users\bullio.m\Desktop\New folder\algo.xlsx", "test", "A1")

            Try
                Connection.Open()
                BI.BulkCopyTimeout = 0
                BI.DestinationTableName = DestinationTable
                BI.WriteToServer(Data)
                Connection.Close()
            Catch ex As Exception
                If ex.Message.Contains("Time") Then
                    Bulk_Insert(CS, DestinationTable, Data)
                End If
                Bulk_Insert = ex.Message
            End Try
        End Using

    End Function

End Class

Friend Class ObjectShredder(Of T)
    ' Fields
    Private _fi As FieldInfo()
    Private _ordinalMap As Dictionary(Of String, Integer)
    Private _pi As PropertyInfo()
    Private _type As Type

    ' Constructor 
    Public Sub New()
        Me._type = GetType(T)
        Me._fi = Me._type.GetFields
        Me._pi = Me._type.GetProperties
        Me._ordinalMap = New Dictionary(Of String, Integer)
    End Sub

    Public Function ShredObject(ByVal table As DataTable, ByVal instance As T) As Object()
        Dim fi As FieldInfo() = Me._fi
        Dim pi As PropertyInfo() = Me._pi
        If (Not instance.GetType Is GetType(T)) Then
            ' If the instance is derived from T, extend the table schema
            ' and get the properties and fields.
            Me.ExtendTable(table, instance.GetType)
            fi = instance.GetType.GetFields
            pi = instance.GetType.GetProperties
        End If

        ' Add the property and field values of the instance to an array.
        Dim values As Object() = New Object(table.Columns.Count - 1) {}
        Dim f As FieldInfo
        For Each f In fi
            values(Me._ordinalMap.Item(f.Name)) = f.GetValue(instance)
        Next
        Dim p As PropertyInfo
        For Each p In pi
            values(Me._ordinalMap.Item(p.Name)) = p.GetValue(instance, Nothing)
        Next

        ' Return the property and field values of the instance.
        Return values
    End Function


    ' Summary:           Loads a DataTable from a sequence of objects.
    ' source parameter:  The sequence of objects to load into the DataTable.</param>
    ' table parameter:   The input table. The schema of the table must match that 
    '                    the type T.  If the table is null, a new table is created  
    '                    with a schema created from the public properties and fields 
    '                    of the type T.
    ' options parameter: Specifies how values from the source sequence will be applied to 
    '                    existing rows in the table.
    ' Returns:           A DataTable created from the source sequence.

    Public Function Shred(ByVal source As IEnumerable(Of T), ByVal table As DataTable, ByVal options As LoadOption?) As DataTable

        ' Load the table from the scalar sequence if T is a primitive type.
        If GetType(T).IsPrimitive Then
            Return Me.ShredPrimitive(source, table, options)
        End If

        ' Create a new table if the input table is null.
        If (table Is Nothing) Then
            table = New DataTable(GetType(T).Name)
        End If

        ' Initialize the ordinal map and extend the table schema based on type T.
        table = Me.ExtendTable(table, GetType(T))

        ' Enumerate the source sequence and load the object values into rows.
        table.BeginLoadData()
        Using e As IEnumerator(Of T) = source.GetEnumerator
            Do While e.MoveNext
                If options.HasValue Then
                    table.LoadDataRow(Me.ShredObject(table, e.Current), options.Value)
                Else
                    table.LoadDataRow(Me.ShredObject(table, e.Current), True)
                End If
            Loop
        End Using
        table.EndLoadData()

        ' Return the table.
        Return table
    End Function

    Public Function ShredPrimitive(ByVal source As IEnumerable(Of T), ByVal table As DataTable, ByVal options As LoadOption?) As DataTable
        ' Create a new table if the input table is null.
        If (table Is Nothing) Then
            table = New DataTable(GetType(T).Name)
        End If
        If Not table.Columns.Contains("Value") Then
            table.Columns.Add("Value", GetType(T))
        End If

        ' Enumerate the source sequence and load the scalar values into rows.
        table.BeginLoadData()
        Using e As IEnumerator(Of T) = source.GetEnumerator
            Dim values As Object() = New Object(table.Columns.Count - 1) {}
            Do While e.MoveNext
                values(table.Columns.Item("Value").Ordinal) = e.Current
                If options.HasValue Then
                    table.LoadDataRow(values, options.Value)
                Else
                    table.LoadDataRow(values, True)
                End If
            Loop
        End Using
        table.EndLoadData()

        ' Return the table.
        Return table
    End Function

    Public Function ExtendTable(ByVal table As DataTable, ByVal type As Type) As DataTable
        ' Extend the table schema if the input table was null or if the value 
        ' in the sequence is derived from type T.
        Dim f As FieldInfo
        Dim p As PropertyInfo

        For Each f In type.GetFields
            If Not Me._ordinalMap.ContainsKey(f.Name) Then
                Dim dc As DataColumn

                ' Add the field as a column in the table if it doesn't exist
                ' already.
                dc = IIf(table.Columns.Contains(f.Name), table.Columns.Item(f.Name), table.Columns.Add(f.Name, f.FieldType))

                ' Add the field to the ordinal map.
                Me._ordinalMap.Add(f.Name, dc.Ordinal)
            End If

        Next

        For Each p In type.GetProperties
            If Not Me._ordinalMap.ContainsKey(p.Name) Then
                ' Add the property as a column in the table if it doesn't exist
                ' already.
                Dim dc As DataColumn
                If table.Columns.Contains(p.Name) Then
                    dc = table.Columns.Item(p.Name)
                Else
                    dc = table.Columns.Add(p.Name, p.PropertyType)
                End If
                ' Add the property to the ordinal map.
                Me._ordinalMap.Add(p.Name, dc.Ordinal)
            End If
        Next

        ' Return the table.
        Return table
    End Function

End Class
