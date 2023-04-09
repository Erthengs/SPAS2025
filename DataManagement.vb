Imports System.IO
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
            Clipboard.Clear()
            Clipboard.SetText(sql)

        End Try
    End Sub
    Public Sub RunSQL2(ByVal sql As String, p1 As String, p2 As String, msg As String)
        Try
            Connect(sql)

            Dim cmd As New NpgsqlCommand
            cmd.Parameters.Add(":id", NpgsqlDbType.Text).Value = p1
            cmd.Parameters.Add(":type", NpgsqlDbType.Text).Value = p2

            cmd.Connection = connection
            cmd.CommandType = CommandType.Text
            cmd.CommandText = sql
            cmd.ExecuteNonQuery()
            cmd.Dispose()
            connection.Close()

        Catch ex As Exception
            Dim e1 As Integer = ex.ToString.IndexOf("UNIQUE constraint failed")
            If e1 > 0 Then MessageBox.Show("Name already exists.") Else MsgBox("RunSQL error while running procedure " & msg & vbCrLf & vbCrLf & Left(ex.ToString, 1000))
            Clipboard.Clear()
            Clipboard.SetText(sql)

        End Try
    End Sub
    Sub Load_Datagridview2(ByVal dgv As DataGridView, sql As String, p1 As String, errmsg As String)

        dgv.DataSource = Nothing
        Connect(sql)

        'cmd.Parameters.Add(":type", NpgsqlDbType.Text).Value = p2

        Dim ds = New DataSet
        Dim da = New NpgsqlDataAdapter()
        Dim ItemColl(1000) As String
        Dim col As New DataGridViewTextBoxColumn

        dgv.CellBorderStyle = DataGridViewCellBorderStyle.None
        'dgv.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

        Try

            da.SelectCommand = New NpgsqlCommand(sql, connection)
            da.SelectCommand.Parameters.Add(":id", NpgsqlDbType.Text).Value = p1
            da.Fill(ds, sql)
            dgv.DataSource = ds.Tables(sql)
            ds.Tables.Add()

        Catch ex As Exception
            MsgBox("er is een fout opgetreden in de datagridview, module: " & errmsg & vbCrLf & vbCrLf & ex.ToString)
        End Try
        connection.Close()

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


        Dim ExcelApp As Object, ExcelBook As Object
        Dim ExcelSheet As Object
        Dim i As Integer
        Dim j As Integer

        'create object of excel
        ExcelApp = CreateObject("Excel.Application")
        ExcelBook = ExcelApp.WorkBooks.Add
        ExcelSheet = ExcelBook.WorkSheets(1)

        With ExcelSheet
            For Each column As DataGridViewColumn In dgv.Columns
                .cells(1, column.Index + 1) = column.HeaderText
            Next
            For i = 1 To dgv.RowCount
                '.cells(i + 1, 1) = dgv.Rows(i - 1).Cells("id").Value
                For j = 0 To dgv.Columns.Count - 1
                    .cells(i + 1, j + 1) = dgv.Rows(i - 1).Cells(j).Value
                Next
            Next
        End With

        ExcelApp.Visible = True
        '
        ExcelSheet = Nothing
        ExcelBook = Nothing
        ExcelApp = Nothing




    End Sub


End Module
