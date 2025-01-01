Imports System.IO
Imports System.Security.Cryptography
Imports System.Text.RegularExpressions
Imports Npgsql
Imports NpgsqlTypes

Module DataManagement
    Public connection As NpgsqlConnection
    Public username As String
    Public pwrd As String
    Public host As String
    Public port As String
    Public database As String
    Public connect_string As String
    Public Edit_Mode As Boolean = False
    Public Add_Mode As Boolean = False
    Public ind1 As Integer = -1
    Public dst As DataSet
    Public dst1 As DataSet
    Public dstbank As DataSet
    Public totsql As String
    Public nocat As String
    Public _overhead As String
    Public report_year As Integer
    Public db As String


    Sub Connect(ByVal SQL)
connstart:
        Dim ans = ""
        totsql &= SQL & vbCrLf

        Try
            connection = New NpgsqlConnection(connect_string)
            If connection.State = ConnectionState.Closed Then connection.Open()
        Catch ex As Exception
            ans = MsgBox("Er kon geen verbinding gemaakt worden met de database. Wilt u het nog een keer proberen?", vbYesNo)
            If ans = vbYes Then
                GoTo connstart
            Else
                Application.Exit()
            End If
        End Try

    End Sub
    Public Sub RunQuery(ByVal Qname As String)

        Dim sql = QuerySQL("Select sql from query where category ilike 'Transaction' and name='" & Qname & "'")
        Try
            RunSQL(sql, "NULL", Qname)
            'Debug.Print("Success: RunQuery " & Qname)
        Catch ex As Exception
            MsgBox(ex.ToString)
            ToClipboard("sql", True)
            Debug.Print("Error: RunQuery " & Qname)
        End Try

    End Sub

    Public Sub ToClipboard(t As String, v As Boolean)
        If IsDBNull(t) Then Exit Sub
        If t = "" Then Exit Sub
        If Strings.Right(connect_string, 4) = "PROD" Or v = False Then Exit Sub
        Clipboard.Clear()
        Clipboard.SetText(t)

    End Sub

    Public Sub RunSQL(ByVal sql As String, jpg As String, msg As String)
        Try
            Connect(sql)
            Dim cmd As New NpgsqlCommand
            If jpg <> "NULL" Then cmd.Parameters.Add(ImageToBlob("@image", jpg))
            cmd.Connection = connection
            cmd.CommandType = CommandType.Text
            cmd.CommandText = sql
            cmd.ExecuteNonQuery()
            cmd.Dispose()
            connection.Close()

        Catch ex As Exception
            Dim e1 As Integer = ex.ToString.IndexOf("UNIQUE constraint failed")
            If e1 > 0 Then MessageBox.Show("Name already exists.") Else MsgBox("RunSQL error while running procedure " & msg & vbCrLf & vbCrLf & Left(ex.ToString, 1000))

        End Try
    End Sub

    Public Sub Load_Listbox(ByVal ls As ListBox, SQLstr As String)
        Connect(SQLstr)
        Dim da = New NpgsqlDataAdapter(SQLstr, connection)
        Dim ds = New DataTable

        da.Fill(ds)
        ls.DataSource = ds
        ls.ValueMember = "id"

        Try
            ls.DisplayMember = "name"
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        connection.Close()
    End Sub
    Public Sub Load_Combobox(ByVal cb As ComboBox, vm As String, dm As String, SQLstr As String)
        'cb.DataSource = Nothing
        Connect(SQLstr)
        Try
            Dim da = New NpgsqlDataAdapter(SQLstr, connection)
            Dim ds = New DataTable
            da.Fill(ds)
            cb.DataSource = ds
            cb.ValueMember = vm
            cb.DisplayMember = dm
        Catch
        End Try
        connection.Close()

    End Sub
    Sub Collect_data(ByVal sql As String)
        Try
            Connect(sql)
            Dim da = New NpgsqlDataAdapter(sql, connection)
            connection.Close()
            Dim ds = New DataSet
            da.Fill(ds, "Table")
            dst = ds
        Catch
            MsgBox("De data kon niet opgehaald worden.")
            Clipboard.Clear()
            Clipboard.SetText(sql)
        End Try

    End Sub
    Sub Collect_data1(ByVal sql As String)
        Try
            Connect(sql)
            Dim da = New NpgsqlDataAdapter(sql, connection)
            connection.Close()
            Dim ds = New DataSet
            da.Fill(ds, "Table")
            dst1 = ds
        Catch
            MsgBox("De data kon niet opgehaald worden.")
            Clipboard.Clear()
            Clipboard.SetText(sql)
        End Try

    End Sub


    Sub Collect_bankdata(ByVal sql As String)
        Try
            Connect(sql)
            Dim da = New NpgsqlDataAdapter(sql, connection)
            connection.Close()
            Dim ds = New DataSet
            da.Fill(ds, "Table")
            dstbank = ds
        Catch
            MsgBox("De data kon niet opgehaald worden.")
            Clipboard.Clear()
            Clipboard.SetText(sql)
        End Try

    End Sub
    Function QuerySQL(ByVal sql As String)
        Try
            Connect(sql)
            Dim cmd As New NpgsqlCommand
            cmd.Connection = connection
            cmd.CommandType = CommandType.Text
            cmd.CommandText = sql
            QuerySQL = cmd.ExecuteScalar()
            cmd.Dispose()
            connection.Close()
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
            QuerySQL = ""
        End Try
    End Function

    Public Sub Load_Listview(ByVal lv As ListView, SQLstr As String)
        lv.Clear()
        lv.View = View.Details
        lv.GridLines = True
        Dim ItemColl(100) As String
        Connect(SQLstr)
        Dim da = New NpgsqlDataAdapter(SQLstr, connection)
        Dim ds = New DataSet
        da.Fill(ds, "Table")

        For i = 0 To ds.Tables(0).Columns.Count - 1
            lv.Columns.Add(ds.Tables(0).Columns(i).ColumnName.ToString())
        Next

        'Now adding the Items in Listview
        For i = 0 To ds.Tables(0).Rows.Count - 1
            For j = 0 To ds.Tables(0).Columns.Count - 1
                ItemColl(j) = ds.Tables(0).Rows(i)(j).ToString()
            Next
            Dim lvi As New ListViewItem(ItemColl)
            lv.Items.Add(lvi)
        Next
        connection.Close() 'later toegevoegd
    End Sub
    Sub Load_Datagridview(ByVal dgv As DataGridView, sql As String, errmsg As String)

        dgv.DataSource = Nothing
        Connect(sql)
        Dim ds = New DataSet
        Dim da = New NpgsqlDataAdapter()
        Dim ItemColl(1000) As String
        Dim col As New DataGridViewTextBoxColumn

        dgv.CellBorderStyle = DataGridViewCellBorderStyle.None
        'dgv.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

        Try
            da.SelectCommand = New NpgsqlCommand(sql, connection)
            da.Fill(ds, sql)
            dgv.DataSource = ds.Tables(sql)
            ds.Tables.Add()

        Catch ex As Exception
            MsgBox("er is een fout opgetreden in de datagridview, module: " & errmsg & vbCrLf & vbCrLf & ex.ToString)
        End Try
        connection.Close()

    End Sub

    Sub Save_Gridview_to_sql(ByVal dgv As DataGridView, tabl As String)
        Connect(dgv.Name & "-" & tabl)
        Dim name, version, type, SQLstr As String
        Dim datum As Date = Format(CDate("01-01-1900"), "dd-MM-yyyy")
        SQLstr = ""
        Dim id As Integer

        For x As Integer = 0 To dgv.Rows.Count - 2
            name = dgv.Rows(x).Cells(1).Value
            If IsDBNull(dgv.Rows(x).Cells(2).Value) Then version = "" Else version = dgv.Rows(x).Cells(2).Value
            If Not IsDBNull(dgv.Rows(x).Cells(3).Value) Then datum = dgv.Rows(x).Cells(3).Value
            If IsDBNull(dgv.Rows(x).Cells(4).Value) Then type = "" Else type = dgv.Rows(x).Cells(4).Value

            If name <> "" Then
                If IsDBNull(dgv.Rows(x).Cells(0).Value) Then 'not yet an id available
                    SQLstr = "INSERT INTO " & tabl & "(name, version, date, type) VALUES 
                ('" & name & "','" & version & "','" & datum & "','" & type & "')"
                Else 'update existing record
                    id = dgv.Rows(x).Cells(0).Value
                    SQLstr = "UPDATE " & tabl & " SET name='" & name & "',version='" & version & "',date='" & datum & "',
                    type='" & type & "' WHERE id ='" & id & "'"
                End If
            End If
            Dim Comm As New NpgsqlCommand(SQLstr, connection)
            Comm.ExecuteNonQuery()
            Comm.Dispose()
        Next
        connection.Close()
    End Sub
    Public Function ImageToBlob(ByVal id As String, ByVal filePath As String)   'Overloads
        Dim fs As FileStream = New FileStream(filePath, FileMode.Open, FileAccess.Read)
        Dim br As BinaryReader = New BinaryReader(fs)
        Dim bm() As Byte = br.ReadBytes(fs.Length)
        br.Close()
        fs.Close()
        'Create ParmYes
        Dim photo() As Byte = bm
        Dim SQLparm As New NpgsqlParameter("@image", photo)
        SQLparm.DbType = DbType.Binary
        SQLparm.Value = photo
        Return SQLparm
    End Function
    Public Function BlobToImage(ByVal blob)

        Dim mStream As New System.IO.MemoryStream
        Dim pData() As Byte = DirectCast(blob, Byte())
        mStream.Write(pData, 0, Convert.ToInt32(pData.Length))
        Dim bm As Bitmap = New Bitmap(mStream, False)
        mStream.Dispose()
        Return bm

    End Function

    Sub Export_2_Excel(ByVal dgv As DataGridView)
        Dim cols As Boolean = MsgBox("Wilt u kleuren in de export meenemen (duurt iets langer)?", vbYesNo) = vbYes
        Dim ExcelApp As Object, ExcelBook As Object
        Dim ExcelSheet As Object
        Dim i As Integer
        Dim j As Integer

        ' Create object of Excel
        ExcelApp = CreateObject("Excel.Application")
        ExcelBook = ExcelApp.WorkBooks.Add
        ExcelSheet = ExcelBook.WorkSheets(1)

        With ExcelSheet
            ' Export column headers
            For Each column As DataGridViewColumn In dgv.Columns
                .Cells(1, column.Index + 1) = column.HeaderText
            Next

            ' Export rows and apply formatting
            For i = 1 To dgv.RowCount
                For j = 0 To dgv.Columns.Count - 1
                    ' Export cell value
                    .Cells(i + 1, j + 1) = dgv.Rows(i - 1).Cells(j).Value

                    If cols Then
                        ' Inside the For loops
                        Dim cellStyle = dgv.Rows(i - 1).Cells(j).Style
                        Dim backColor = If(Not cellStyle.BackColor.IsEmpty, cellStyle.BackColor, dgv.Rows(i - 1).DefaultCellStyle.BackColor)
                        'Dim foreColor = If(Not cellStyle.ForeColor.IsEmpty, cellStyle.ForeColor, dgv.Rows(i - 1).DefaultCellStyle.ForeColor)

                        ' Check for fallback to DataGridView default styles
                        If backColor.IsEmpty Then backColor = dgv.DefaultCellStyle.BackColor
                        'If foreColor.IsEmpty Then foreColor = dgv.DefaultCellStyle.ForeColor

                        ' Apply colors to Excel
                        .Cells(i + 1, j + 1).Interior.Color = RGB(backColor.R, backColor.G, backColor.B)
                        '.Cells(i + 1, j + 1).Font.Color = RGB(foreColor.R, foreColor.G, foreColor.B)
                    End If
                Next
            Next
        End With

        ExcelApp.Visible = True

        ' Release Excel objects
        ExcelSheet = Nothing
        ExcelBook = Nothing
        ExcelApp = Nothing



    End Sub
    Public Sub Run_SQL_Journal(ByVal caller As String, operation As String, id As Integer, name As String, datum As Date, status As String,
        amt1 As Decimal, amt2 As Decimal, description As String, source As String, fk_account As Integer,
        fk_relation As Integer, fk_bank As Integer, type As String, cpinfo As String, iban As String)

        Dim sql As String = ""
        Dim upd_qt As Integer = 0
        Dim ins_qt As Integer = 0
        Dim operationDetails As String = $"id:'{id}';name:'{name}';date:'{datum}', status:'{status}' ,
        amt1:'{amt1}';amt2:'{amt2}', description:'{description}';source:'{source}';fk_account:'{fk_account}',
        fk_relation:'{fk_relation}';fk_bank:'{fk_bank}';type:'{type}';cpinfo:'{cpinfo}';iban:'{iban}'"

        Select Case operation
            Case "UPDATE"
                sql = "UPDATE public.journal SET name=@name,Date=@datum,status=@status,amt1=@amt1,amt2=@amt2, 
                description=@description,source=@source,fk_account=@fk_account,fk_relation=@fk_relation, 
                fk_bank=@fk_bank, type=@type,cpinfo=@cpinfo,iban=@cpinfo WHERE id=@id;"

            Case "INSERT"
                sql = "INSERT INTO public.journal (name,Date,status,amt1,amt2
                ,description,source,fk_account, fk_relation, fk_bank,type, cpinfo, iban) 
                VALUES(@name,@datum,@status,@amt1,@amt2,@description,@source,@fk_account
                ,@fk_relation,@fk_bank,@type,@cpinfo,@iban);"

            Case "DELETE"
                sql = "DELETE FROM public.journal WHERE id=@id;"
        End Select

        Try
            Connect(sql)
            Dim cmd As New NpgsqlCommand
            cmd.Parameters.AddWithValue("@id", id)
            cmd.Parameters.AddWithValue("@name", name)
            cmd.Parameters.AddWithValue("@datum", datum)
            cmd.Parameters.AddWithValue("@status", status)
            cmd.Parameters.AddWithValue("@amt1", amt1)
            cmd.Parameters.AddWithValue("@amt2", amt2)
            cmd.Parameters.AddWithValue("@description", description)
            cmd.Parameters.AddWithValue("@source", source)
            cmd.Parameters.AddWithValue("@fk_account", fk_account)
            cmd.Parameters.AddWithValue("@fk_relation", fk_relation)
            cmd.Parameters.AddWithValue("@fk_bank", fk_bank)
            cmd.Parameters.AddWithValue("@type", type)
            cmd.Parameters.AddWithValue("@cpinfo", cpinfo)
            cmd.Parameters.AddWithValue("@iban", iban)

            cmd.Connection = connection
            cmd.CommandType = CommandType.Text
            cmd.CommandText = sql
            cmd.ExecuteNonQuery()
            cmd.Dispose()
            'connection.Close()

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        MsgBox(db)
        Dim logSql As String = "INSERT INTO public.operation_logs (db_name, caller, operation_type, username, operation_details) " &
                       "VALUES (@db_name, @caller, @operation_type, @username, @operation_details);"
        Try

            Dim logCmd As New NpgsqlCommand(logSql, connection)
            logCmd.Connection = connection
            logCmd.Parameters.AddWithValue("@db_name", db)
            logCmd.Parameters.AddWithValue("@caller", caller)
            logCmd.Parameters.AddWithValue("@operation_type", operation)
            logCmd.Parameters.AddWithValue("@username", username)
            logCmd.Parameters.AddWithValue("@operation_details", operationDetails)
            logCmd.ExecuteNonQuery()
            logCmd.Dispose()
        Catch ex As Exception
            Console.WriteLine("Logging failed: " & ex.Message)
            MsgBox(ex.Message)
        End Try

    End Sub

    Sub PopulateDataGridView()
        ' Define your list of tables
        Dim tableNames As List(Of String) = New List(Of String) From {
            "account", "accgroup", "bank", "bankacc", "contract",
            "cp", "journal", "relation", "Settings", "target"
        }


        ' Create a DataTable to hold the table name and record count
        Dim dataTable As New DataTable()
        dataTable.Columns.Add("TableName", GetType(String))
        dataTable.Columns.Add("TotalRecords", GetType(Integer))

        ' Open the PostgreSQL connection
        Using connection As New NpgsqlConnection(connect_string)
            connection.Open()

            ' Loop through each table and get the record count
            For Each tableName As String In tableNames
                Dim query As String = $"SELECT COUNT(*) FROM {tableName};"
                Using cmd As New NpgsqlCommand(query, connection)
                    Dim totalRecords As Integer = Convert.ToInt32(cmd.ExecuteScalar())
                    ' Add the table name and total records to the DataTable
                    dataTable.Rows.Add(tableName, totalRecords)
                End Using
            Next
        End Using

        ' Bind the DataTable to the DataGridView
        SPAS.Dgv_Mgnt_Tables.DataSource = dataTable

        ' Optionally, set the DataGridView column headers
        SPAS.Dgv_Mgnt_Tables.Columns(0).HeaderText = "Table Name"
        SPAS.Dgv_Mgnt_Tables.Columns(1).HeaderText = "Total Records"
    End Sub

    Sub Populate_DataTree()
        Collect_data("select name from query where sql ilike '%[year]%'")
        Dim conn = New NpgsqlConnection(connect_string)
        conn.Open()
        Dim cmd As New NpgsqlCommand($"SELECT module, name from query where category = 'Overzicht' order by module, name;", conn)
        Dim cat As String
        Dim name1 As String
        Dim cat2 As String = ""


        Dim dt As NpgsqlDataReader = cmd.ExecuteReader


        While dt.Read
            cat = dt.Item("module")
            name1 = dt.Item("name")
            Dim node1 As New TreeNode

            If cat2 = cat Then
                Dim childNode As New TreeNode(name1)
                childNode.Tag = "Child"
                SPAS.BankTree.SelectedNode.Nodes.Add(childNode) '

                For i = 0 To dst.Tables(0).Rows.Count - 1
                    If name1 = dst.Tables(0).Rows(i)(0) Then
                        AddLevel2Nodes(childNode)
                    End If
                Next

            Else
                node1 = SPAS.BankTree.Nodes.Add(cat)
                cat2 = cat
                SPAS.BankTree.SelectedNode = node1
                Dim childNode As New TreeNode(name1)
                childNode.Tag = "Child"
                SPAS.BankTree.SelectedNode.Nodes.Add(childNode)

                For i = 0 To dst.Tables(0).Rows.Count - 1
                    'MsgBox($"Overzicht: '{dst.Tables(0).Rows(i)(0)}'")
                    If name1 = dst.Tables(0).Rows(i)(0) Then
                        AddLevel2Nodes(childNode)
                    End If
                Next
            End If

        End While

    End Sub

    Private Sub AddLevel2Nodes(parentNode As TreeNode)
        Dim level2Values As New List(Of String) From {"2023", "2022", "2021"}

        For Each value As String In level2Values
            Dim level2Node As New TreeNode(value)
            level2Node.Tag = "Level2"
            parentNode.Nodes.Add(level2Node)
        Next
    End Sub
End Module
