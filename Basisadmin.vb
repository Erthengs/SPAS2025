Imports System.IO
Imports System.Runtime.Remoting.Channels
Imports Npgsql
Imports PdfSharp.Pdf.Content.Objects
Module Basisadmin
    Public group As String
    Public reload As Boolean = False

    Public Class ContractModel
        Public Property Id As Integer
        Public Property Name As String ' Het logische contractnummer (bijv. K00001)
        Public Property FkTargetId As Integer
        Public Property FkRelationId As Integer
        Public Property FkAccountId As Integer
        Public Property Donation As Decimal
        Public Property Overhead As Decimal
        Public Property Term As Integer
        Public Property StartDate As Date
        Public Property EndDate As Date
        Public Property Description As String
        Public Property Autcol As Boolean
        Public Property Active As Boolean
        Public Property Intern As Boolean

        ' Helper functie om te controleren of financiën/incasso zijn gewijzigd
        Public Function RequiresNewVersion(other As ContractModel) As Boolean
            Return Me.Donation <> other.Donation OrElse
               Me.Overhead <> other.Overhead OrElse
               Me.Autcol <> other.Autcol OrElse
               Me.Term <> other.Term
        End Function
    End Class

    ' 1. Controleer of er een toekomstige versie van dit contract bestaat
    Public Function HasFutureVersion(contractName As String, currentStartDate As Date) As Boolean
        Dim sql As String = "SELECT COUNT(id) FROM contract WHERE name = @name AND startdate > @startdate"

        Using conn As New NpgsqlConnection(connect_string)
            conn.Open()
            Using cmd As New NpgsqlCommand(sql, conn)
                cmd.Parameters.AddWithValue("@name", contractName)
                cmd.Parameters.AddWithValue("@startdate", currentStartDate)

                Dim count As Integer = Convert.ToInt32(cmd.ExecuteScalar())
                Return count > 0
            End Using
        End Using
    End Function

    ' 2. Haal één contract op via ID
    Public Function GetContractById(id As Integer) As ContractModel
        Dim contract As ContractModel = Nothing
        Dim sql As String = "SELECT * FROM contract WHERE id = @id"

        Using conn As New NpgsqlConnection(connect_string)
            conn.Open()
            Using cmd As New NpgsqlCommand(sql, conn)
                cmd.Parameters.AddWithValue("@id", id)

                Using reader As NpgsqlDataReader = cmd.ExecuteReader()
                    If reader.Read() Then
                        contract = New ContractModel()

                        contract.Id = Convert.ToInt32(reader("id"))
                        contract.Name = reader("name").ToString()

                        contract.FkTargetId = If(IsDBNull(reader("fk_target_id")), 0, Convert.ToInt32(reader("fk_target_id")))
                        contract.FkRelationId = If(IsDBNull(reader("fk_relation_id")), 0, Convert.ToInt32(reader("fk_relation_id")))
                        contract.FkAccountId = If(IsDBNull(reader("fk_account_id")), 0, Convert.ToInt32(reader("fk_account_id")))

                        contract.Donation = If(IsDBNull(reader("donation")), 0D, Convert.ToDecimal(reader("donation")))
                        contract.Overhead = If(IsDBNull(reader("overhead")), 0D, Convert.ToDecimal(reader("overhead")))

                        contract.Term = If(IsDBNull(reader("term")), 12, Convert.ToInt32(reader("term")))
                        contract.StartDate = Convert.ToDateTime(reader("startdate"))
                        contract.EndDate = Convert.ToDateTime(reader("enddate"))
                        contract.Description = reader("description").ToString()

                        contract.Autcol = If(IsDBNull(reader("autcol")), False, Convert.ToBoolean(reader("autcol")))
                        contract.Active = If(IsDBNull(reader("active")), False, Convert.ToBoolean(reader("active")))
                        contract.Intern = If(IsDBNull(reader("intern")), False, Convert.ToBoolean(reader("intern")))
                    End If
                End Using
            End Using
        End Using

        Return contract
    End Function

    ' 3. Update alleen de omschrijving (dit leidt nooit tot een nieuwe versie)

    ' 3. Update omschrijving en einddatum (dit leidt nooit tot een nieuwe versie)
    Public Sub UpdateContractBasicInfo(contract As ContractModel)
        Dim sql As String = "UPDATE contract SET description = @desc, enddate = @enddate, active = @active WHERE id = @id"

        Using conn As New NpgsqlConnection(connect_string)
            conn.Open()
            Using cmd As New NpgsqlCommand(sql, conn)
                If String.IsNullOrEmpty(contract.Description) Then
                    cmd.Parameters.AddWithValue("@desc", DBNull.Value)
                Else
                    cmd.Parameters.AddWithValue("@desc", contract.Description)
                End If

                cmd.Parameters.AddWithValue("@enddate", contract.EndDate)
                cmd.Parameters.AddWithValue("@active", contract.Active)
                cmd.Parameters.AddWithValue("@id", contract.Id)
                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    ' 4. Maak een nieuwe versie aan (Sluit de oude, voeg de nieuwe in)
    Public Sub CreateNewContractVersion(oldContractId As Integer, newContract As ContractModel)
        Dim newOldEndDate As Date = newContract.StartDate.AddDays(-1)
        Dim isActive As Boolean = (newOldEndDate >= Date.Today)

        Using conn As New NpgsqlConnection(connect_string)
            conn.Open()

            ' Transacties moeten expliciet aan het command worden doorgegeven bij Npgsql
            Using trans = conn.BeginTransaction()
                Try
                    ' --- OUDE CONTRACT UPDATEN ---
                    Dim sqlUpdateOld As String = "UPDATE contract SET enddate = @enddate, active = @active WHERE id = @oldId"
                    Using cmdUpdate As New NpgsqlCommand(sqlUpdateOld, conn, trans)
                        cmdUpdate.Parameters.AddWithValue("@enddate", newOldEndDate)
                        cmdUpdate.Parameters.AddWithValue("@active", isActive)
                        cmdUpdate.Parameters.AddWithValue("@oldId", oldContractId)
                        cmdUpdate.ExecuteNonQuery()
                    End Using

                    ' --- NIEUWE CONTRACT TOEVOEGEN ---
                    Dim sqlInsertNew As String = "INSERT INTO contract (name, fk_target_id, fk_relation_id, fk_account_id, donation, overhead, term, startdate, enddate, description, autcol, active, intern) " &
                                             "VALUES (@name, @target, @relation, @account, @donation, @overhead, @term, @startdate, '2999-12-31', @desc, @autcol, @active, @intern)"

                    Using cmdInsert As New NpgsqlCommand(sqlInsertNew, conn, trans)
                        cmdInsert.Parameters.AddWithValue("@name", newContract.Name)
                        cmdInsert.Parameters.AddWithValue("@target", newContract.FkTargetId)
                        cmdInsert.Parameters.AddWithValue("@relation", newContract.FkRelationId)

                        If newContract.FkAccountId > 0 Then
                            cmdInsert.Parameters.AddWithValue("@account", newContract.FkAccountId)
                        Else
                            cmdInsert.Parameters.AddWithValue("@account", DBNull.Value)
                        End If

                        cmdInsert.Parameters.AddWithValue("@donation", newContract.Donation)
                        cmdInsert.Parameters.AddWithValue("@overhead", newContract.Overhead)
                        cmdInsert.Parameters.AddWithValue("@term", newContract.Term)
                        cmdInsert.Parameters.AddWithValue("@startdate", newContract.StartDate)

                        If String.IsNullOrEmpty(newContract.Description) Then
                            cmdInsert.Parameters.AddWithValue("@desc", DBNull.Value)
                        Else
                            cmdInsert.Parameters.AddWithValue("@desc", newContract.Description)
                        End If

                        cmdInsert.Parameters.AddWithValue("@autcol", newContract.Autcol)
                        cmdInsert.Parameters.AddWithValue("@active", newContract.Active)
                        cmdInsert.Parameters.AddWithValue("@intern", newContract.Intern)

                        cmdInsert.ExecuteNonQuery()
                    End Using

                    trans.Commit()
                Catch ex As Exception
                    trans.Rollback()
                    Throw
                End Try
            End Using
        End Using
    End Sub

    ' 5. Standaard insert voor een compleet nieuw contract
    Public Function InsertNewContract(newContract As ContractModel) As Integer
        Dim sql As String = "INSERT INTO contract (name, fk_target_id, fk_relation_id, fk_account_id, donation, overhead, term, startdate, enddate, description, autcol, active, intern) " &
                        "VALUES (@name, @target, @relation, @account, @donation, @overhead, @term, @startdate, '2999-12-31', @desc, @autcol, @active, @intern) RETURNING id"

        Using conn As New NpgsqlConnection(connect_string)
            conn.Open()
            Using cmd As New NpgsqlCommand(sql, conn)
                cmd.Parameters.AddWithValue("@name", newContract.Name)
                cmd.Parameters.AddWithValue("@target", newContract.FkTargetId)
                cmd.Parameters.AddWithValue("@relation", newContract.FkRelationId)

                If newContract.FkAccountId > 0 Then
                    cmd.Parameters.AddWithValue("@account", newContract.FkAccountId)
                Else
                    cmd.Parameters.AddWithValue("@account", DBNull.Value)
                End If

                cmd.Parameters.AddWithValue("@donation", newContract.Donation)
                cmd.Parameters.AddWithValue("@overhead", newContract.Overhead)
                cmd.Parameters.AddWithValue("@term", newContract.Term)
                cmd.Parameters.AddWithValue("@startdate", newContract.StartDate)

                If String.IsNullOrEmpty(newContract.Description) Then
                    cmd.Parameters.AddWithValue("@desc", DBNull.Value)
                Else
                    cmd.Parameters.AddWithValue("@desc", newContract.Description)
                End If

                cmd.Parameters.AddWithValue("@autcol", newContract.Autcol)
                cmd.Parameters.AddWithValue("@active", newContract.Active)
                cmd.Parameters.AddWithValue("@intern", newContract.Intern)

                Dim newId As Integer = Convert.ToInt32(cmd.ExecuteScalar())
                Return newId
            End Using
        End Using
    End Function

    ' 6. Ophalen overlappend contract ter validatie
    Public Function GetOverlappingContract(targetId As Integer, relationId As Integer, startDate As Date) As String
        Dim sql As String = "SELECT c.name FROM contract c " &
                        "WHERE c.fk_target_id = @targetId AND c.fk_relation_id = @relationId " &
                        "AND c.enddate >= @startDate " &
                        "LIMIT 1"

        Using conn As New NpgsqlConnection(connect_string)
            conn.Open()
            Using cmd As New NpgsqlCommand(sql, conn)
                cmd.Parameters.AddWithValue("@targetId", targetId)
                cmd.Parameters.AddWithValue("@relationId", relationId)
                cmd.Parameters.AddWithValue("@startDate", startDate)

                Dim result As Object = cmd.ExecuteScalar()
                If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                    Return result.ToString()
                End If
            End Using
        End Using

        Return String.Empty
    End Function

    ' 7. Het veilig inladen van de contractlijst met filters
    Public Function GetContractList(searchTerm As String, lifeCycle As String) As DataTable
        Dim dt As New DataTable()
        Dim sql As String = "SELECT contract.id, CONCAT(relation.name, ', ', relation.name_add, ' ---> ', target.name, ', ', target.name_add) as name " &
                        "FROM contract " &
                        "JOIN target ON contract.fk_target_id = target.id " &
                        "JOIN relation ON contract.fk_relation_id = relation.id " &
                        "WHERE 1=1 "

        If lifeCycle = "Actief" Then
            sql &= "AND contract.active = True "
        ElseIf lifeCycle = "Inactief" Then
            sql &= "AND contract.active = False "
        End If

        If Not String.IsNullOrWhiteSpace(searchTerm) Then
            sql &= "AND (contract.name ILIKE @search OR target.name ILIKE @search OR relation.name ILIKE @search OR target.name_add ILIKE @search OR relation.name_add ILIKE @search) "
        End If

        sql &= "ORDER BY relation.name, target.name"

        Using conn As New NpgsqlConnection(connect_string)
            conn.Open()
            Using cmd As New NpgsqlCommand(sql, conn)
                If Not String.IsNullOrWhiteSpace(searchTerm) Then
                    cmd.Parameters.AddWithValue("@search", "%" & searchTerm & "%")
                End If

                Using adapter As New NpgsqlDataAdapter(cmd)
                    adapter.Fill(dt)
                End Using
            End Using
        End Using

        Return dt
    End Function

    ' 8. Het veilig verwijderen van een toekomstige wijziging
    Public Sub DeleteFutureContract(contractId As Integer, contractName As String)
        Using conn As New NpgsqlConnection(connect_string)
            conn.Open()

            Using trans = conn.BeginTransaction()
                Try
                    ' Deletion
                    Dim sqlDel As String = "DELETE FROM contract WHERE id = @id"
                    Using cmdDel As New NpgsqlCommand(sqlDel, conn, trans)
                        cmdDel.Parameters.AddWithValue("@id", contractId)
                        cmdDel.ExecuteNonQuery()
                    End Using

                    ' Restore End Date
                    Dim sqlUpd As String = "UPDATE contract SET enddate = '2999-12-31', active = True " &
                                       "WHERE name = @name AND id = (SELECT MAX(id) FROM contract WHERE name = @name)"
                    Using cmdUpd As New NpgsqlCommand(sqlUpd, conn, trans)
                        cmdUpd.Parameters.AddWithValue("@name", contractName)
                        cmdUpd.ExecuteNonQuery()
                    End Using

                    trans.Commit()
                Catch ex As Exception
                    trans.Rollback()
                    Throw
                End Try
            End Using
        End Using
    End Sub

    '================================================================================================================
    '==
    '==
    '===
    '================ G E N E R I C =================================================================================
    'field codes:
    '---_x-_ format (0 = undetermined, 1 = currency, 2 = integer, 3 = date)
    '---_-x_ obligatory (0 = optional, 1 = Not NULL)

    Function Cur(ByVal amt As String)
        'amt = Convert.ToDecimal(amt)
        Dim ovh As Boolean = (SPAS.Text <> "SPAS LOKALE TESTDATABASE")
        Dim curamt As String
        If amt = "" Then
            curamt = "'0'"
        Else
            curamt = "'" & IIf(Not ovh, amt, Replace(amt, ",", ".")) & "'"
        End If
        Return curamt


    End Function
    Function Cur2(ByVal amt As String)
        'amt = Convert.ToDecimal(amt)
        Dim ovh As Boolean = (SPAS.Text <> "SPAS LOKALE TESTDATABASE")
        Dim curamt As String
        If amt = "" Then
            curamt = "0"
        Else
            curamt = IIf(Not ovh, amt, Replace(amt, ",", "."))

        End If
        Return curamt


    End Function

    Sub Count_Occurences()
        Dim qty As Integer
        PopulateDataGridView()
        Exit Sub

        Dim metadata = Collect_data2("
                select 
                (select count(*) from account) As Accounts,
                (select count(*) from accgroup) As Accountgroepen,
                (select count(*) from bank) As Banktransacties,
                (select count(*) from bankacc) As Bankrekeningen,
                (select count(*) from contract) As Contracten,
                (select count(*) from cp) As Contactpersonen,
                (select count(*) from journal) As Journaalposten,
                (select count(*) from relation) As Relaties,
                (select count(*) from settings) As Settings,
                (select count(*) from target) As Doel
                ")

        For i = 0 To metadata.Columns.Count - 1

            qty = IIf(IsDBNull(metadata.Rows(0)(i)), 0, metadata.Rows(0)(i))

            SPAS.Dgv_Mgnt_Tables.ColumnCount = 2
            SPAS.Dgv_Mgnt_Tables.Columns(0).Name = "Tabel"
            SPAS.Dgv_Mgnt_Tables.Columns(1).Name = "Aantal records"
            SPAS.Dgv_Mgnt_Tables.Rows.Add(metadata.Columns(i).ColumnName)
            SPAS.Dgv_Mgnt_Tables.Rows(SPAS.Dgv_Mgnt_Tables.Rows.Count - 1).Cells(1).Value = qty
        Next i
        MsgBox(SPAS.Dgv_Mgnt_Tables.Rows(3).Cells(1).Value)
    End Sub


    Sub Empty_Tabpage()

        Dim tb As Integer = SPAS.TC_Object.SelectedIndex
        For Each ctl In SPAS.TC_Object.TabPages(tb).Controls
            If TypeOf ctl Is TextBox Or TypeOf ctl Is ComboBox Then
                If Mid(ctl.Name, 5, 1) = "0" Then ctl.Text = "" Else ctl.Text = 0
                'ctl.SelectedIndex = -1
            End If
            If TypeOf ctl Is CheckBox Then
                If InStr(ctl.Name, "__active") > 0 Then
                    ctl.Checked = True
                Else
                    ctl.Checked = False
                End If
            End If
            If TypeOf ctl Is Label And Strings.InStr(ctl.Name, "__") > 0 Then ctl.Text = ""
            If TypeOf ctl Is DateTimePicker Then ctl.Value = "31-12-2999"
            If TypeOf ctl Is PictureBox Then ctl.Image = Nothing
        Next
        'If SPAS.Lbx_Basis.Items.Count > 0 Then Select_Obj2()
    End Sub

    Sub Select_Obj2(sender As String)

        SPAS.isManualChange = False

        '@@@deze module moet nog verbeterd worden via gebruik van een dataset en het kunnen hanteren van 0-waarden


        'A GENERIC PART ===========================================================================
        Dim fld, fk_tbl As String, id, fk_id, pos, pos1, pos2 As Integer
        Dim tb As Integer = SPAS.TC_Object.SelectedIndex
        Dim tbl As String = SPAS.TC_Object.TabPages(tb).Name
        Dim tmp
        Dim col As Integer = -1


        Try
            id = SPAS.Lbx_Basis.Items(SPAS.Lbx_Basis.SelectedIndex)(SPAS.Lbx_Basis.ValueMember)
        Catch ex As Exception
            Exit Sub
        End Try
        Dim objectdata = Collect_data2("SELECT * FROM " & tbl & " WHERE id='" & id & "'")

        Empty_Tabpage()

        For Each ctl In SPAS.TC_Object.TabPages(tb).Controls

            If Strings.InStr(ctl.Name, "_pkid") > 0 Then ctl.Text = id

            pos = Strings.InStr(ctl.Name, "__")
            If pos > 0 Then
                fld = Mid(ctl.Name, pos + 2, Len(ctl.Name) - pos)
                'retrieve the name of accompanying columns
                For i = 0 To objectdata.Columns.Count - 1
                    If fld = objectdata.Columns(i).ColumnName Then
                        col = i
                        Exit For
                    End If
                Next
                If col = -1 Then Exit Sub

                If TypeOf ctl Is TextBox Or TypeOf ctl Is Label Then
                    Select Case Strings.Mid(ctl.Name, 5, 1)

                        Case 1
                            tmp = objectdata.Rows(0)(col)
                            If IsDBNull(tmp) Then
                                ctl.Text = 0
                            Else
                                ctl.Text = tmp
                                '@@@currency_converter
                            End If
                        Case Else
                            If IsDBNull(objectdata.Rows(0)(col)) Then ctl.Text = "" Else ctl.Text = objectdata.Rows(0)(col)

                    End Select
                ElseIf TypeOf ctl Is CheckBox Then

                    ctl.Checked = objectdata.Rows(0)(col)
                ElseIf TypeOf ctl Is PictureBox Then

                    Dim img As Image
                    Try
                        Dim photo = objectdata.Rows(0)(col) 'QuerySQL("SELECT " & fld & " FROM " & tbl & " WHERE id='" & id & "'")
                        img = BlobToImage(photo)
                        ctl.Image = img
                    Catch ex As Exception
                        ctl.Image = Nothing
                    End Try
                ElseIf TypeOf ctl Is ComboBox Then

                    '1) get fk_id from data base
                    pos1 = Strings.InStr(ctl.Name, "fk_")
                    If pos1 > 0 Then
                        pos2 = Strings.InStr(ctl.Name, "_id")
                        fk_id = objectdata.Rows(0)(col) 'QuerySQL("SELECT " & fld & " FROM " & tbl & " WHERE id='" & id & "'")
                        fk_tbl = Mid(fld, 4, Len(fld) - 6) ', pos2 - pos1
                        If fk_tbl = "bank" Then fk_tbl = "bankacc"
                        If fk_tbl = "acco" Then fk_tbl = "account"
                        If fk_tbl = "bankacc" Or fk_tbl = "account" Or fk_tbl = "accgroup" Then  '@@@ workaround
                            ctl.Text = QuerySQL("SELECT name FROM " & fk_tbl & " WHERE id='" & fk_id & "'")
                        Else
                            Dim sqltext = $"SELECT Concat(name, ', ', name_add) as name FROM {fk_tbl} WHERE id='{fk_id}'"
                            'Clipboard.SetText(sqltext)
                            ctl.Text = QuerySQL(sqltext)

                            If ctl.Name = "Cmx_00_contract_fk_relation_id" Then SPAS.Lbl_11_contract__fk_relation_id.Text = fk_id
                        End If
                    Else
                        ctl.Text = objectdata.Rows(0)(col).ToString
                    End If

                ElseIf TypeOf ctl Is DateTimePicker Then
                    ctl.Value = objectdata.Rows(0)(col)
                End If
            Else

                If ctl.Name = "Cmx_00_contract_fk_relation_id" Then
                    Dim relid As String = QuerySQL($"Select fk_relation_id from contract where id = {objectdata.Rows(0)(0)}")
                    SPAS.Lbl_11_contract__fk_relation_id.Text = relid
                    SPAS.Cmx_00_contract_fk_relation_id.Text = QuerySQL($"Select concat(name,', ',name_add) from relation where id = {relid}")
                End If
                If ctl.Name = "Cmx_01_contract_fk_target_id" Then
                    Dim tarid As String = QuerySQL($"Select fk_target_id from contract where id = {objectdata.Rows(0)(0)}")
                    SPAS.Lbl_11_contract__fk_target_id.Text = tarid
                    SPAS.Cmx_01_contract_fk_target_id.Text = QuerySQL($"Select concat(name,', ',name_add) from target where id = {tarid}")
                    Try
                        SPAS.Pic_Contract_Target_photo.Image = BlobToImage(QuerySQL("SELECT photo FROM target WHERE id='" & id & "'"))
                    Catch ex As Exception
                        SPAS.Pic_Contract_Target_photo.Image = Nothing
                    End Try
                End If

            End If
        Next

        'B OBJECT SPECIFIC PART ========================================================================
        'addition for contract
        If tb = 0 Then

            Dim sqlstr = "
                        SELECT ta.ttype, r.iban, ba.id
                        FROM contract co
                        LEFT join target ta ON co.fk_target_id = ta.id
                        LEFT join relation r ON co.fk_relation_id = r.id
                        LEFT join bankacc ba ON ba.accountno = r.iban 
                        WHERE co.id = '" & id & "'
                        "
            Dim contractdata = Collect_data2(sqlstr)
            'SPAS.Tbx_Contract_ttype.Text = contractdata.Rows(0)(0)
            If contractdata.Rows(0)(0) = "Kind" Then SPAS.Rbn_00_contract_child.Checked = True
            If contractdata.Rows(0)(0) = "Oudere" Then SPAS.Rbn_00_contract_elder.Checked = True
            If contractdata.Rows(0)(0) = "Overig" Then SPAS.Rbn_00_contract_other.Checked = True

            'SPAS.Dtp_30_Contract_Change.Value = Date.Today
            SPAS.Lbl_Contract_Bronaccount.Visible = Not IsDBNull(contractdata.Rows(0)(2))
            SPAS.Cmx_Contract_fk_account_id.Visible = Not IsDBNull(contractdata.Rows(0)(2))
            SPAS.Lbl_Contract_tgt.Text = SPAS.Cmx_01_contract_fk_target_id.Text
            'Cmx_01_Target__fk_cp_id

        End If
        If tb = 1 Then
            'SPAS.Cmx_01_Target__fk_cp_id.Text = "Marchitan"

        End If

        If tb = 2 Then 'RELATION
            SPAS.Dtp_00_relation__date1.Enabled = (SPAS.Tbx_00_Relation__iban.Text <> "")
            SPAS.Dtp_00_relation__date2.Enabled = (SPAS.Tbx_00_Relation__iban.Text <> "")
            SPAS.Dtp_00_relation__date3.Enabled = (SPAS.Tbx_00_Relation__iban.Text <> "")
            'vul giftenoverzicht
            Dim sql As String = "
		    select r.name, j.date, 'Overschrijving', amt1 from relation r
		    left join bank b on b.iban2 = r.iban
		    left join journal j on j.fk_bank = b.id
		    where b.code in ('cb', 'ei') and r.id =" & id & "
		    union select r.name,j.date, 'Incasso', sum(amt1) from relation r 
		    left join journal j on r.id = j.fk_relation
		    where source='Incasso' and r.id =" & id & "
		    group by r.name, j.date"
            Load_Datagridview(SPAS.Dgv_relation_giften, sql, "Select_Obj2")
            With SPAS.Dgv_relation_giften
                .Columns(0).Visible = False
                .Columns(1).HeaderText = "Datum"
                .Columns(1).Width = 75
                .Columns(2).HeaderText = "Betaling"
                .Columns(3).HeaderText = "Bedrag"
                .Columns(3).DefaultCellStyle.Format = "N2"
                .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                .Columns(3).Width = 70
            End With
        End If
        If tb = 4 Then 'ACCOUNT

            SPAS.Cbx_00_Account__active.Enabled = (SPAS.Lbl_00_Account__source.Text = "cat")
            If SPAS.Lbl_00_Account__source.Text = "cat" Then SPAS.Tbx_01_Account__name.Enabled = True

        End If
        If tb = 6 Then 'bankACCOUNT

            SPAS.Tbx_BankAcc_startbalance.Text = QuerySQL("select credit-debit from bank b left join bankacc c on c.accountno=b.iban 
                                                where b.name='_startsaldo_' and b.iban ='" & SPAS.Tbx_01_BankAcc__accountno.Text & "'")

        End If
        SPAS.isManualChange = True
        SPAS.StoreInitialValues(SPAS.Controls)
    End Sub

    Sub Load_Table()
        Dim tb As Integer = SPAS.TC_Object.SelectedIndex
        Dim tbl As String = SPAS.TC_Object.TabPages(tb).Name
        Dim SQLstr, SQLstr1, SQLstr2 As String

        ' Dim arg = SPAS.Tbx_Basis_Filter.Text.ToUpper
        'Dim arg = SPAS.Searchbox.Text.ToUpper
        Dim arg = SPAS.Searchbox2.Text.ToUpper
        Dim sel_act As String = ""
        If SPAS.Cbx_LifeCycle2.Text = "Actief" Then
            sel_act = " AND active=True"
        End If
        If SPAS.Cbx_LifeCycle2.Text = "Inactief" Then
            sel_act = " AND active=False"
        End If


        Dim filtersql As String = ""


        If tb = 0 Then
            If arg <> "" Then
                filtersql = "And (contract.name Like '%" & arg & "%' 
                              Or target.name iLike '%" & arg & "%' 
                              Or relation.name iLike '%" & arg & "%'
                              Or target.name_add iLike '%" & arg & "%' 
                              Or relation.name_add iLike '%" & arg & "%')"
            Else
                filtersql = ""
            End If
            SQLstr = "SELECT contract.id, CONCAT(relation.name, ',', relation.name_add, ' - ', target.name, ',', target.name_add) as name FROM contract 
                          JOIN target ON contract.fk_target_id = target.id 
                          JOIN relation ON contract.fk_relation_id = relation.id 
                          WHERE contract.active=" & IIf(SPAS.Cbx_LifeCycle2.Text = "Inactief", False, True) & "
                          " & filtersql & "
                          ORDER BY relation.name, target.name"
            Load_Listbox(SPAS.Lbx_Basis, SQLstr)
            'Clipboard.Clear()
            'Clipboard.SetText(SQLstr)

        ElseIf tb = 1 Then
            If arg <> "" Then
                filtersql = "
                        AND (t.name ILIKE '%" & arg & "%' 
                        OR t.name_add ILIKE '%" & arg & "%' 
                        OR t.ttype ILIKE '%" & arg & "%'
                        OR cp.name ILIKE  '%" & arg & "%')
                              "
            Else
                filtersql = ""
            End If


            SQLstr = "SELECT t.id, CONCAT(t.ttype,' / ', t.name, ', ', t.name_add,' / ', 
                        cp.name) as name 
                        FROM " & tbl & " t 
                        LEFT JOIN cp ON cp.id = t.fk_cp_id
                        WHERE t.active=" & IIf(SPAS.Cbx_LifeCycle2.Text = "Inactief", False, True) & " 
                         " & filtersql & "
                        ORDER BY t.name"

            Load_Listbox(SPAS.Lbx_Basis, SQLstr)

        ElseIf tb = 4 Then
            SQLstr1 = "SELECT id, CONCAT(source,' ',name,' (',accgroup,')') as name FROM " & tbl & " 
                       WHERE (name iLike '%" & arg & "%'" & sel_act & " 
                       OR accgroup iLike '%" & arg & "%' 
                       OR source iLike '%" & arg & "%')
                       AND (active=" & IIf(SPAS.Cbx_LifeCycle2.Text = "Inactief", False, True) & ") 
                       ORDER BY source, accgroup, name"
            Load_Listbox(SPAS.Lbx_Basis, SQLstr1)


        ElseIf tb = 6 Then
            SQLstr2 = "SELECT id, name FROM " & LCase(tbl) & " WHERE name ILIKE '%" & arg & "%'" & sel_act & " ORDER BY name"

            Load_Listbox(SPAS.Lbx_Basis, SQLstr2)

        Else
            SQLstr1 = "SELECT id, CONCAT(name, ', ', name_add) as name FROM " & tbl & " WHERE UPPER(name) Like '%" & arg & "%'" & sel_act & " ORDER BY name"
            SQLstr2 = "SELECT id, name FROM " & tbl & " WHERE UPPER(name) Like '%" & arg & "%'" & sel_act & " ORDER BY name"
            Try
                If SPAS.Chbx_test.Checked = True Then MsgBox(SQLstr1)
                Load_Listbox(SPAS.Lbx_Basis, SQLstr1)

            Catch ex As Exception
                Load_Listbox(SPAS.Lbx_Basis, SQLstr2)

            End Try

        End If

        SPAS.isManualChange = True
    End Sub

    Sub Locate_Listbox_Position(ByVal valit1 As String)
        Dim rowit1 As Int32

        For rowit1 = 0 To SPAS.Lbx_Basis.Items.Count - 1
            If SPAS.Lbx_Basis.Items(rowit1)(SPAS.Lbx_Basis.ValueMember) = valit1 Then
                SPAS.Lbx_Basis.SetSelected(rowit1, True)

                Exit For
            End If
        Next

    End Sub
    Function Handle_errors(ByVal errmsg As String)
        Dim tb As Integer = SPAS.TC_Object.SelectedIndex
        Dim tbl As String = SPAS.TC_Object.TabPages(tb).Name

        Dim pos, cnt, ix As Integer
        Dim nm, nma As String
        Dim errmsg1 = ""
        For Each f In SPAS.TC_Object.TabPages(tb).Controls
            If Strings.Mid(f.Name, 6, 1) = "1" And f.Text = "" Then
                pos = Strings.InStr(f.Name, "__")
                errmsg1 &= "- " & f.Tag & " (" & f.name & ") mag niet leeg zijn" & vbCrLf   '' Mid(f.Tag, pos + 2)
            End If

        Next
        ix = SPAS.TC_Object.SelectedIndex
        Select Case ix

            Case 0
                'contract: control that either sponsor or intern account is selected
                If SPAS.Cmx_Contract_fk_account_id.Text = "" And SPAS.Cmx_00_contract_fk_relation_id.Text = "" Then
                    errmsg1 &= "- Kies ofwel een externe sponsor ofwel een intern fondsaccount." & vbCrLf
                End If
                'check whether there is an active contract with the same sponsor and sponsoree
                If Add_Mode Then
                    Dim startdate As String = SPAS.Dtp_31_contract__startdate.Value.Year & "-" & SPAS.Dtp_31_contract__startdate.Value.Month & "-" &
                        SPAS.Dtp_31_contract__startdate.Value.Day
                    Dim sqlstr As String = "
                    SELECT t.name||','||t.name_add||' ('||r.name||','||r.name_add||') tot '||c.enddate FROM contract c
                    LEFT JOIN target t on t.id = c.fk_target_id 
                    LEFT JOIN relation r on r.id = c.fk_relation_id 
                    WHERE '" & startdate & "' < enddate
                    AND fk_target_id ='" & SPAS.Cmx_01_contract_fk_target_id.SelectedValue & "'
                    AND fk_relation_id ='" & SPAS.Cmx_00_contract_fk_relation_id.SelectedValue & "'"

                    Clipboard.Clear()
                    Clipboard.SetText(sqlstr)

                    Dim res As String = QuerySQL(sqlstr)
                    If res <> "" Then errmsg1 &= "Er loopt al een contract voor " & res & "." & vbCrLf &
                        "Dit contract mag daarmee niet overlappen. Beëindig deze eerst alvorens dit contract af te sluiten."
                End If

            Case 1, 2, 3
                'control on unique names
                If Add_Mode Then
                    nm = Strings.Trim(IIf(ix = 1, SPAS.Tbx_01_Target__name.Text, IIf(ix = 2, SPAS.Tbx_01_relation__name.Text, SPAS.Tbx_01_CP__name.Text)))
                    nma = Strings.Trim(IIf(ix = 1, SPAS.Tbx_01_Target__name_add.Text, IIf(ix = 2, SPAS.Tbx_01_Relation__name_add.Text, SPAS.Tbx_01_CP__name_add.Text)))

                    cnt = QuerySQL("SELECT count(*) from " & tbl &
                                   " WHERE name='" & nm & "' AND name_add='" & nma & "'")

                    If cnt > 0 Then
                        errmsg1 &= "- De naam " & nm & ", " & nma & " komt al voor in de administratie" & vbCrLf
                    End If

                    If ix = 3 And LCase(Strings.Left(nm, 5)) = "nieuw" Then
                        errmsg1 &= "- de naam van een contactpersoon mag niet met 'nieuw' beginnen."
                    End If
                End If
            Case 4
                If Tbx2Dec(SPAS.Lbl_Account_Budget_Difference.Text) <> 0 Then
                    errmsg1 &= "- som van aangepaste maandbudgetten Is ongelijk aan jaarbudget" & vbCrLf
                End If
        End Select

        If errmsg1 <> "" Then errmsg = "Invoerfouten: " & vbCrLf & errmsg1
        Return errmsg
    End Function

    Sub Update_table()
        'Dim id As Integer = SPAS.Lbx_Basis.Items(SPAS.Lbx_Basis.SelectedIndex)(SPAS.Lbx_Basis.ValueMember)
        Dim tb As Integer = SPAS.TC_Object.SelectedIndex
        Dim tbl As String = SPAS.TC_Object.TabPages(tb).Name
        Dim id1 As Integer
        Dim fld As String
        Dim v
        Dim SQLStr As String = "UPDATE " & tbl.ToLower & " SET "

        For Each f In SPAS.TC_Object.TabPages(tb).Controls
            If Strings.InStr(f.Name, "_pkid") > 0 Then
                id1 = Convert.ToInt32(f.Text)  'retrieve proprietary key
            Else
                Dim pos = Strings.InStr(f.Name, "__")
                If pos > 0 And TypeOf f IsNot PictureBox Then
                    If TypeOf f Is CheckBox Then
                        v = f.Checked
                    ElseIf TypeOf f Is DateTimePicker Then
                        'v = f.Text
                        v = "'" & Format(CDate(f.Text), "yyyy-MM-dd") & "'"
                    ElseIf TypeOf f Is ComboBox And Strings.Right(f.Name, 3) = "_id" Then
                        v = f.SelectedValue
                        If Len(v) = 0 Then v = 0
                        If SPAS.Chbx_test.Checked = True Then MsgBox(f.Name & "->" & Mid(f.Name, pos + 2, Len(f.Name) - pos) & "-->" & v)
                    ElseIf Mid(f.Name, 5, 1) = 1 Then 'currency value
                        v = Cur(f.Text)
                    ElseIf Mid(f.Name, 5, 1) = 2 Then 'integer
                        v = IIf(f.Text = "", 0, f.Text)

                    Else
                        f.Text = Replace(f.Text, "'", "´")
                        v = "'" & f.Text & "'"
                    End If
                    fld = Mid(f.Name, pos + 2, Len(f.Name) - pos)
                    SQLStr &= fld & "= " & v & ","
                End If
            End If
        Next
        SQLStr = Left(SQLStr, Strings.Len(SQLStr) - 1) & " WHERE id=" & id1 & ";"  'remove final komma
        If SPAS.Chbx_test.Checked = True Then MsgBox(SQLStr)

        RunSQL(SQLStr, "NULL", "Update table " & tbl)

    End Sub
    Sub Insert_into_table()

        'Dim id As Integer = SPAS.Lbx_Basis.Items(SPAS.Lbx_Basis.SelectedIndex)(SPAS.Lbx_Basis.ValueMember)
        Dim tb As Integer = SPAS.TC_Object.SelectedIndex
        Dim tbl As String = SPAS.TC_Object.TabPages(tb).Name
        Dim pos, new_id As Integer
        Dim ImgFile As String = "NULL"
        Dim fld, SQLstr, name As String
        Dim v
        Dim d As Date


        Dim s1 As String = "INSERT INTO " & tbl.ToLower & "("
        Dim s2 As String = " VALUES("

        For Each f In SPAS.TC_Object.TabPages(tb).Controls

            'If Strings.InStr(f.Name, "__id") > 0 Then
            If Strings.InStr(f.Name, "_pkid") > 0 Then
                'do nothing, id is generated
            Else
                pos = Strings.InStr(f.Name, "__")
                If pos > 0 Then
                    If TypeOf f Is CheckBox Then
                        v = f.Checked
                    ElseIf TypeOf f Is ComboBox And Strings.Right(f.Name, 3) = "_id" Then
                        v = f.SelectedValue
                        If Strings.Len(f.SelectedValue) = 0 Then v = "0"
                    ElseIf TypeOf f Is TextBox And Mid(f.Name, 5, 1) = "0" Then
                        f.Text = Replace(f.Text, "'", "´")
                        v = "'" & f.Text & "'"
                    ElseIf TypeOf f Is TextBox And Mid(f.Name, 5, 1) = "1" Then
                        'currency
                        v = Cur(f.Text)
                        'v = v.Replace(".00", "")
                    ElseIf TypeOf f Is Label And Mid(f.Name, 5, 1) = "0" Then
                        f.Text = Replace(f.Text, "'", "´")
                        v = "'" & f.Text & "'"
                    ElseIf TypeOf f Is Label And Mid(f.Name, 5, 1) = "1" Then
                        v = Cur(f.Text)
                    ElseIf TypeOf f Is DateTimePicker Then
                        'date
                        d = f.Text
                        v = "'" & d.Year & "-" & d.Month & "-" & d.Day & "'"
                    ElseIf TypeOf f Is TextBox And Mid(f.Name, 5, 1) <> "2" Then
                        f.Text = Replace(f.Text, "'", "´")
                        If f.Text = "" Then v = "0" Else v = f.Text
                    Else
                        v = "'" & f.Text & "'"
                    End If
                    fld = Mid(f.Name, pos + 2, Len(f.Name) - pos)
                    s1 &= fld & ","
                    s2 &= v & ","
                End If
            End If
        Next
        SQLstr = Left(s1, Strings.Len(s1) - 1) & ") " & Left(s2, Strings.Len(s2) - 1) & ");"
        If SPAS.Chbx_test.Checked = True Then MsgBox(SQLstr)
        Clipboard.SetText(SQLstr)

        RunSQL(SQLstr, "NULL", "Insert into table " & tbl)

        'addition for contract;  target  - --
        Select Case tb
            Case 0
                SPAS.Pan_contract_select_target.Enabled = False
                Dim Source_Account = QuerySQL("SELECT id FROM account WHERE f_key='" & SPAS.Cmx_01_contract_fk_target_id.SelectedValue & "'")
                Calculate_Budget(Source_Account)
                SPAS.Cmx_00_contract_fk_relation_id.Enabled = False
                SPAS.Cmx_01_contract_fk_target_id.Enabled = False
            Case 2
                Load_Combobox(SPAS.Cmx_00_contract_fk_relation_id, "id", "name", "SELECT r.id, CONCAT(r.name, ', ', r.name_add) as name FROM relation r WHERE r.active=TRUE ORDER BY r.name")
            Case 1, 3 'creating an account for target or cp...
                Load_Combobox(SPAS.Cmx_01_Target__fk_cp_id, "id", "name", "SELECT cp.id, CONCAT(cp.name, ', ', cp.name_add) as name FROM cp WHERE cp.active=True ORDER BY cp.name")
                new_id = QuerySQL("Select Max(id) From " & tbl)
                Dim tbtxt As String = SPAS.TC_Object.TabPages(tb).Tag
                SQLstr = "SELECT CONCAT(name,',', name_add) FROM " & tbl & " WHERE id=" & new_id
                If SPAS.Chbx_test.Checked Then MsgBox(SQLstr)
                name = QuerySQL(SQLstr)


                Create_Account(tbtxt.ToLower, name, SPAS.Tbx_01_Target__ttype.Text, new_id, "Specifiek (doel)")
        End Select

    End Sub
    Sub Save_Image(pic As PictureBox)

        Dim ImgFile, SQLstr As String
        Dim id As Integer
        Dim tb As Integer = SPAS.TC_Object.SelectedIndex
        Dim tbl As String = SPAS.TC_Object.TabPages(tb).Name


        For Each f In SPAS.TC_Object.TabPages(tb).Controls
            If Strings.InStr(f.Name, "_pkid") > 0 Then id = Convert.ToInt32(f.Text)
        Next

        If pic.Image Is Nothing Then
            pic.Image = Clipboard.GetImage()
            If pic.Image Is Nothing Then
                MsgBox("U heeft geen afbeeling op het klembord. Druk op Shift+Windowstoets+S om een afbeelding van het scherm te selecteren")
                Exit Sub
            End If
            Try
                If Not pic.Image Is Nothing Then
                    pic.Image.Save(IO.Path.Combine(My.Computer.FileSystem.SpecialDirectories.MyPictures, "SPAStmp_pic.jpg"))
                    ImgFile = IO.Path.Combine(My.Computer.FileSystem.SpecialDirectories.MyPictures, "SPAStmp_pic.jpg")
                Else
                    ImgFile = "NULL"
                End If

            Catch
                ImgFile = "NULL"
                MsgBox("Niets op het klembord")
            End Try
            If Add_Mode Then  'Or Edit_Mode
                MsgBox("U kunt pas een foto toevoegen als " & tbl & "opgeslagen is.")
            Else

            End If
            SQLstr = "UPDATE " & tbl & " SET photo=@image WHERE id=" & id
            If SPAS.Chbx_test.Checked = True Then MsgBox(SQLstr)
            RunSQL(SQLstr, ImgFile, "")
        Else
            pic.Image = Nothing
            SQLstr = "UPDATE " & tbl & " SET photo=null WHERE id=" & id
            If SPAS.Chbx_test.Checked = True Then MsgBox(SQLstr)
            RunSQL(SQLstr, "NULL", "Save_Image")


        End If

    End Sub
    'TARGET MODULES ===============================================================================================

    Sub Calculate_Target_Totals()
        SPAS.Lbl_Target_Total_Income.Text = GetDouble(SPAS.Tbx_10_Target__allowance.Text) + GetDouble(SPAS.Tbx_10_Target__otherincome.Text) +
             GetDouble(SPAS.Tbx_10_Target__benefit.Text) + GetDouble(SPAS.Tbx_10_Target__pension.Text) + GetDouble(SPAS.Tbx_10_Target__income.Text)

        SPAS.Lbl_Target_Total_Expenses.Text = GetDouble(SPAS.Tbx_10_Target__rent.Text) + GetDouble(SPAS.Tbx_10_Target__gaselectra.Text) +
             GetDouble(SPAS.Tbx_10_Target__medicine.Text) + GetDouble(SPAS.Tbx_10_Target__food.Text) + GetDouble(SPAS.Tbx_10_Target__heating.Text) +
             GetDouble(SPAS.Tbx_10_Target__water.Text)

        SPAS.Lbl_Target_Total_Income.Text = Tbx2Dec(SPAS.Lbl_Target_Total_Income.Text)
        SPAS.Lbl_Target_Total_Expenses.Text = Tbx2Dec(SPAS.Lbl_Target_Total_Expenses.Text)
    End Sub

    Sub Save_Target()
        Dim dat As String = Convert.ToDateTime(SPAS.Dtp_00_Target__birthday.Value).ToString("dd-MM-yyyy")
    End Sub

    'END TARGET MODULES =============================================================================================


    'START RELATION  MODULES ===============================================================================================
    'to do: edit mode voor velden cmbx accountno
    'account no moet fkey worden
    'image functionaliteit toevoegen
    Sub Generate_Reference()

        Dim name = Strings.Left(SPAS.Tbx_01_relation__name.Text, 3)
        If Strings.Len(name) > 1 Then
            Dim amt As Integer = QuerySQL("SELECT COUNT(*) FROM relation WHERE name LIKE '" & name & "%'") + 1
            SPAS.Lbl_00_relation__reference.Text = name.ToLower & Strings.Left("0" & amt.ToString, 2)
        End If

    End Sub

    Sub CheckActive(ByVal chbx As CheckBox, id1 As Label, relatedobj As String)
        Dim n As String = chbx.Name
        Dim obj = Strings.Left(Mid(n, InStr(n, "00_") + 3), Len(n) - 15)

        If chbx.Checked Then
            RunSQL("UPDATE " & obj & " SET active=True WHERE id=" & id1.Text, "NULL", "")
        Else
            Dim SQLstr = "SELECT count(id) FROM " & relatedobj & " WHERE fk_" & obj & "_id=" & CInt(id1.Text) & " AND active=true"

            If QuerySQL(SQLstr) > 0 Then
                MsgBox("Deactivatie is niet mogelijk, er zijn nog één of meer relaties met " & relatedobj)
                chbx.Checked = True
            Else
                RunSQL("UPDATE " & obj & " SET active=False WHERE id=" & id1.Text, "NULL", "")
                MsgBox("Deactivatie uitgevoerd: kan niet meer gekozen worden in een contract of uitkeringsformulier.")
            End If
        End If

    End Sub


    'END RELATION MODULES =============================================================================================

    'ACCOUNT MODULES =============================================================================================

    Sub Create_Account(ByVal source As String, name As String, accgroup As String, fk As Integer, acctype As String)

        Dim SQLstr As String = "INSERT INTO account(name,source,type,f_key,active, fk_accgroup_id) 
                                VALUES('" & name & "','" & source & "','" & acctype & "','" & fk & "',true
                                , (select id from accgroup where subtype='" & accgroup & "'))"
        RunSQL(SQLstr, "NULL", "")

    End Sub


    'END ACCOUNT MODULES =============================================================================================

    'START CONTRACT MODULES ===========================================================================================
    'to do
    '1) prevent contracts when there is already an active contract between the combination of relation and target
    '2) error: wrong image after adding a new 

    Sub Handle_Contract_Fields()
        'SPAS.Pan_contract_select_target.Visible = Add_Mode
        '
        SPAS.Cmx_01_contract_fk_target_id.Enabled = Add_Mode
        SPAS.Cmx_00_contract_fk_relation_id.Enabled = Add_Mode
        SPAS.Cmx_Contract_fk_account_id.Enabled = Add_Mode

    End Sub
    Sub Create_Contract_Version()
        'SQLstr = "INSERT INTO contract"

    End Sub
    Sub Get_Sponsor_data()


        Dim rel_id = QuerySQL($"Select id from relation where concat(name,', ',name_add)='{SPAS.Cmx_00_contract_fk_relation_id.Text}' ")
        Dim d As String = $"date{IIf(SPAS.Rbn_00_contract_child.Checked, "1", IIf(SPAS.Rbn_00_contract_child.Checked, "2", "3"))}"
        Dim intern_contract As Boolean = (QuerySQL($"Select iban from relation where id={rel_id}") = "RekeningStichting")

        With SPAS
            .Lbl_11_contract__fk_relation_id.Text = rel_id
            .Lbl_00_contract_autcol.Text = QuerySQL("SELECT reference FROM relation WHERE id=" & rel_id)
            .dtp_contract_relation_date.Value = QuerySQL("SELECT " & d & " FROM relation WHERE id=" & rel_id)
            .Chx_00_contract__autcol.Checked =
            .dtp_contract_relation_date.Value < SPAS.Dtp_31_contract__startdate.Value
            .Lbl_Contract_Bronaccount.Visible = intern_contract
            .Cmx_Contract_fk_account_id.Visible = intern_contract
            .Lbl_10_Contract__fk_account_id.Visible = intern_contract
            .Chx_00_contract__autcol.Enabled = Not intern_contract
            .Lbl_10_Contract__fk_account_id.Text = Strings.Trim(Strings.Left(.Cmx_Contract_fk_account_id.Text, 4))
        End With



    End Sub

    Sub Calculate_contract_amounts()

        SPAS.Tbx_01_contract_yeartotal.Text = (GetDouble(SPAS.Tbx_11_Contract__donation.Text) _
           + GetDouble(SPAS.Tbx_11_contract__overhead.Text))
        SPAS.Tbx_contract_period_amt.Text = (GetDouble(SPAS.Tbx_01_contract_yeartotal.Text) /
            GetDouble(SPAS.Cmx_02_Contract__term.Text))

    End Sub


    Function Contract_number(ByVal prefix As String)
        Dim cnt = QuerySQL("SELECT COUNT(name) FROM Contract WHERE name Like '%" & prefix & "%'")
        Contract_number = prefix & Strings.Right("000000" & cnt + 1, 5)
        Return Contract_number
    End Function

    'END CONTRACT MODULES =============================================================================================

    'START JOURNAL MODULES ===============================================================================================
    'Lifecycle Journal -- per transaction or per individual posting? 
    'automatic generated posting as 'undesignated': new
    'generated without linked banktransaction: new
    'manually assigned category: open
    'automatically assigned category: open
    'year close: posted

    Sub Add_Journal_Post(ByVal _dat As Date, stat As String, descr As String, sour As String, name As String,
                         amt1 As Double, amt2 As Double,
                         fkac As Integer, fkba As Integer, fkre As Integer)
        Dim SQLstr As String

        'generate name

        SQLstr = "INSERT INTO journal(name, date, status, amt1, amt2, 
                  description, source, fk_account, fk_bank, fk_relation)
                  VALUES('" & name & "','" & _dat & "','" & stat & "'," & amt1 & "," & amt2 & ",'" &
                  descr & "','" & sour & "'," & fkac & "," & fkba & "," & fkre & ");"
        If SPAS.Chbx_test.Checked Then MsgBox(SQLstr)
        RunSQL(SQLstr, "NULL", "")

    End Sub

    '=================================================================================
    'incasso
    '=================================================================================

    Function Display_Incasso()
        Dim SQLstr = "
            SELECT distinct Concat(r.name, ', ', r.name_add), fk_account, ta.ttype,
            (Select sum(amt1) from journal where fk_account = '100169'  AND fk_relation = r.id) As ovd,
            (Select sum(amt1) from journal where fk_account != '100169' AND fk_relation = r.id  As don
            FROM journal j
            LEFT JOIN relation r ON j.fk_relation = r.id
            LEFT join account ac ON j.fk_account = ac.id
            LEFT JOIN target ta ON ac.f_key = ta.id
            WHERE
            j.source = 'Incasso' AND 
            j.date='01-01-2021' 
            Group by  j.amt1, Concat(r.name, ', ', r.name_add), j.fk_relation, r.id, fk_account, ta.ttype
"
        Return SQLstr
    End Function

    Function Create_Incasso_Totals(date_start As String)

        Dim SQLstr As String = "
            Select 'Kind' As Doel,  count (distinct r.id) As Aantal,sum((co.donation+co.overhead)/term) As Totaal 
            From contract co LEFT Join Target ta ON co.fk_target_id = ta.id LEFT Join Relation r ON co.fk_relation_id = r.id 
            Where co.autcol = True And co.startdate <= '" & date_start & "' AND co.enddate > '" & date_start & "' AND r.date1 <='" & date_start & "' AND ta.ttype = 'Kind'
            union
            Select 'Oudere',  count (distinct r.id),sum((co.donation+co.overhead)/term)
            From contract co  LEFT Join Target ta ON co.fk_target_id = ta.id LEFT Join Relation r ON co.fk_relation_id = r.id
            Where co.autcol = True And co.startdate <= '" & date_start & "' AND co.enddate > '" & date_start & "' AND  r.date2 <='" & date_start & "' AND ta.ttype = 'Oudere'
            union
            Select 'Overig',  count (distinct r.id),sum((co.donation+co.overhead)/term)
            From contract co LEFT Join Target ta ON co.fk_target_id = ta.id LEFT Join Relation r ON co.fk_relation_id = r.id
            Where co.autcol = True And co.startdate <= '" & date_start & "' AND co.enddate > '" & date_start & "' AND  r.date3 <='" & date_start & "' AND ta.ttype = 'Overig'

            union
            Select 'Totaal',
			(SELECT count (distinct r.id)
            FROM contract co LEFT JOIN Target ta ON co.fk_target_id = ta.id LEFT JOIN Relation r ON co.fk_relation_id = r.id 
            WHERE co.autcol = True AND co.startdate <= '" & date_start & "' AND co.enddate > '" & date_start & "' AND r.date1 <='" & date_start & "' AND ta.ttype = 'Kind'
			) + 
			(SELECT count (distinct r.id)
            FROM contract co  LEFT JOIN Target ta ON co.fk_target_id = ta.id LEFT JOIN Relation r ON co.fk_relation_id = r.id
            WHERE co.autcol = True AND co.startdate <= '" & date_start & "' AND co.enddate > '" & date_start & "' AND  r.date2 <='" & date_start & "' AND ta.ttype = 'Oudere'
			) + 
			(SELECT count (distinct r.id)
            FROM contract co  LEFT JOIN Target ta ON co.fk_target_id = ta.id LEFT JOIN Relation r ON co.fk_relation_id = r.id
            WHERE co.autcol = True AND co.startdate <= '" & date_start & "' AND co.enddate > '" & date_start & "' AND  r.date3 <='" & date_start & "' AND ta.ttype = 'Overig'
			),
			(SELECT sum((co.donation+co.overhead)/term)
            FROM contract co  LEFT JOIN Target ta ON co.fk_target_id = ta.id LEFT JOIN Relation r ON co.fk_relation_id = r.id
            WHERE co.autcol = True AND co.startdate <= '" & date_start & "' AND co.enddate > '" & date_start & "') 

        "
        Return SQLstr


    End Function


    Function Existing_Excasso(ByVal exnam As String)

        Dim overhead As String = QuerySQL("SELECT value FROM settings WHERE label='overhead'")


        Dim SQLstr = "
            SELECT 
            ac.id, ac.name, 
            (SELECT SUM(amt1) FROM journal WHERE fk_account=ac.id  AND type ILIKE 'Contract%') + 
            (SELECT SUM(amt1) FROM journal WHERE fk_account=ac.id  AND type ILIKE 'Contract%' 
                AND name ='" & exnam & "')*-1 As Contract, 
            (SELECT SUM(amt1) FROM journal WHERE fk_account=ac.id  AND type ILIKE 'Extra%') + 
            (SELECT SUM(amt1) FROM journal WHERE fk_account=ac.id  AND type ILIKE 'Extra%' 
                AND name ='" & exnam & "')*-1 As Extra,
            (SELECT SUM(amt1) FROM journal WHERE fk_account=ac.id  AND type ILIKE 'Intern%') +
            (SELECT SUM(amt1) FROM journal WHERE fk_account=ac.id  AND type ILIKE 'Intern%' 
                AND name ='" & exnam & "')*-1 As Intern,
            SUM(j.amt1)*-1 As Eur, 
            SUM(j.amt2)*-1 As MDL,
            j.type
            FROM journal j
            LEFT JOIN account ac ON ac.id = fk_account
            WHERE j.name ='" & exnam & "'
            AND ac.id != '" & overhead & "'
            GROUP BY ac.id, ac.name, j.type
            ORDER BY ac.name ASC

"
        Return SQLstr




    End Function


    Function Create_Excasso(ByVal CP As String, t1 As String, t2 As String, t3 As String, d1 As String, d2 As String)

        Dim SQLstr As String = "

        SELECT 
            distinct ac.id, ac.name,
	        CASE 
				WHEN " & d2 & " = 1 Then ac.b_jan
				WHEN " & d2 & " = 2 Then ac.b_feb 
				WHEN " & d2 & " = 3 Then ac.b_mar
				WHEN " & d2 & " = 4 Then ac.b_apr
				WHEN " & d2 & " = 5 Then ac.b_may 
				WHEN " & d2 & " = 6 Then ac.b_jun
				WHEN " & d2 & " = 7 Then ac.b_jul
				WHEN " & d2 & " = 8 Then ac.b_aug
				WHEN " & d2 & " = 9 Then ac.b_sep
				WHEN " & d2 & " = 10 Then ac.b_oct
				WHEN " & d2 & " = 11 Then ac.b_nov
				WHEN " & d2 & " = 12 Then ac.b_dec
            END As MndBdt,

			(Select sum(amt1) from journal where type = 'Contract' 
                AND journal.fk_account = ac.id 
                AND journal.date <='" & d1 & "') As Contract,
	        (Select sum(amt1) from journal where type =  'Extra' 
                AND journal.fk_account = ac.id) As Extra,
	        (Select sum(amt1) from journal where type = 'Internal' 
                AND journal.fk_account = ac.id) As Intern,0,0

        FROM  
            Account ac
            ---LEFT JOIN journal j ON j.fk_account = ac.id AND  j.name LIKE 'Contract%'  
            ---LEFT JOIN journal j2 ON j2.fk_account = ac.id AND j2.name LIKE 'Extra%'
            ---LEFT JOIN journal j3 ON j3.fk_account = ac.id  AND j3.name LIKE 'Intern%'
            LEFT JOIN target ta ON ta.id = ac.f_key
            LEFT JOIN cp ON cp.id = ta.fk_cp_id
            WHERE cp.id='" & CP & "'
			---AND 
                ---j.date <= '" & d1 & "'::date
            AND
				(ta.ttype='" & t1 & "' OR
                ta.ttype='" & t2 & "' OR
                ta.ttype='" & t3 & "')
            AND ta.active=true
			GROUP BY ac.id, ac.name ---, j.name, j.amt1
            ORDER BY ac.name ASC

"

        Return SQLstr


    End Function

    Sub Basis_Delete()
        Dim id As Integer
        Dim sqlstr As String = ""
        Dim t As Integer = SPAS.TC_Object.SelectedIndex
        If SPAS.Lbx_Basis.SelectedIndex <> -1 Then id = SPAS.Lbx_Basis.SelectedItem(SPAS.Lbx_Basis.ValueMember) Else Exit Sub

        Select Case t
            Case 0
                If SPAS.Dtp_31_contract__startdate.Value <= Date.Today Then
                    MsgBox("Alleen contracten die nog niet zijn ingegaan kunnen verwijderd worden.")
                    Exit Sub
                Else
                    If MsgBox("Weet u zeker dat u dit contract wilt verwijderen (vergeet niet eventueel de einddatum van eerdere versie van dit contract terug te zetten)?", vbYesNo) = vbNo Then
                        Exit Sub
                    Else

                        QuerySQL("Update account set b_jan=0, b_feb=0, b_mar=0, b_apr=0, b_may=0, b_jun=0, b_jul=0, b_aug=0, b_sep=0, b_oct=0, b_nov=0, b_dec=0 
                        where source ilike 'Doel' and f_key=" & SPAS.Cmx_01_contract_fk_target_id.SelectedValue)

                        sqlstr = "DELETE FROM contract WHERE id=" & id

                    End If
                End If

            Case 1
                Dim targetdata = Collect_data2("SELECT t.id, t.name, t.active, ac.name, ac.id, j.id As journal, c.id As Contract
                                From target t
                                LEFT join account ac on t.id= ac.f_key
                                LEFT join journal j on j.fk_account = ac.id
                                LEFT join contract c on c.fk_target_id = t.id
                                WHERE (j.id is null or c.id is null)
                                AND t.id =" & id)
                If targetdata.Rows.Count = 0 Then
                    MsgBox("Dit doel maakt nog onderdeel uit van een contract waarop transacties hebben plaatsgevonden." & vbCrLf &
                           "U kunt het niet verwijderen, maar wel inactief maken zodat er geen contract meer voor kan worden afgesloten of giften aan gegeven.")
                    Exit Sub
                End If
                Dim account_id = targetdata.Rows(0)(4)
                Dim journal_id = targetdata.Rows(0)(5)
                Dim contract_id = targetdata.Rows(0)(6)

                Dim Msg As String = "Dit doel: "
                If Not IsDBNull(contract_id) Then Msg &= vbCrLf & "- maakt onderdeel uit van contract " & contract_id
                If Not IsDBNull(journal_id) Then Msg &= vbCrLf & "- komt voor in journaalposten"
                If Len(Msg) > 10 Then
                    Msg &= vbCrLf & "en kan daarom niet verwijderd worden. U kunt het wel als [inactief] markeren."
                    MsgBox(Msg)
                Else
                    If MsgBox("Weet u zeker dat u het doel " & SPAS.Tbx_01_Target__name.Text & "," & SPAS.Tbx_01_Target__name_add.Text &
                        " wilt verwijderen?") Then
                        sqlstr = "Delete from target where id=" & id
                        'verwijderd totdat er ook in journal_archive een check plaatsvindt'  [;  'DELETE from account WHERE id=" & account_id]
                    End If

                End If

            Case 2
                If QuerySQL("select count(id) from contract where fk_relation_id = " & id) > 0 Then
                    MsgBox("Deze relatie staat geregistreerd bij contracten; deze moeten eerst verwijderd worden")
                Else
                    If MsgBox("Weet u zeker dat u deze relatie wilt verwijderen?", vbYesNo) = vbNo Then
                        Exit Sub
                    Else
                        sqlstr = "delete from relation where id = " & id
                    End If
                End If
            Case 3
                Dim cpdata = Collect_data2("SELECT cp.name, ac.name, j.name As journal, t.name FROM CP
                                LEFT join account ac on cp.id = ac.f_key
                                LEFT JOIN journal j on ac.id = j.fk_account
                                LEFT JOIN target t on t.fk_cp_id = cp.id 
                                WHERE ac.id is not distinct from null or j.id is not distinct from null 
                                or cp.id is not distinct from null AND cp.id =" & id)
                Dim account_id = cpdata.Rows(0)(1)
                If cpdata.Rows.Count = 0 Then
                    MsgBox("Deze staat nog geregistreerd bij doel(en) en/of journaalposten." & vbCrLf &
                           "U kunt het niet verwijderen, maar wel inactief maken zodat deze niet mmer gebruikt kan worden.")
                    Exit Sub
                Else
                    sqlstr = "Delete from cp where id=" & id & ";DELETE from account WHERE id=" & account_id
                End If
            Case Else
                MsgBox("Deze functie Is nog niet voor dit object gedefinieerd")

        End Select

        If sqlstr <> "" Then
            RunSQL(sqlstr, "NULL", "Menu_Delete_Click")
            Load_Table()
            MsgBox("Het object is verwijderd.")
        End If

    End Sub

    Public Sub Populate_BasisTree(ByVal tvBasis As TreeView, ByVal lifecycleStatus As String, Optional ByVal searchword2 As String = "")
        ' 1. Fetch the base SQL query from the public.query table
        Dim sql As String = QuerySQL("SELECT sql FROM query WHERE name = 'BasisTree_Load'")

        If String.IsNullOrEmpty(sql) Then
            MsgBox("The query 'BasisTree_Load' was not found in the database.")
            Exit Sub
        End If

        ' 2. Process the Lifecycle Filter (Actief, Inactief, Beide)
        Dim activeFilter As String = ""
        Select Case lifecycleStatus.ToLower()
            Case "actief"
                activeFilter = " AND c.active = true "
            Case "inactief"
                activeFilter = " AND c.active = false "
            Case "beide"
                activeFilter = "" ' Leave empty so it doesn't restrict by active status
        End Select

        ' 3. Process the Searchword Filter
        Dim searchFilter As String = ""
        Dim safeSearch As String = searchword2.Replace("'", "''") ' Prevent basic syntax errors
        If Not String.IsNullOrWhiteSpace(safeSearch) Then
            searchFilter = $" AND (c.name ILIKE '%{safeSearch}%' OR r.name ILIKE '%{safeSearch}%' OR t.name ILIKE '%{safeSearch}%') "
        End If

        ' 4. Inject the generated filters into the SQL template
        sql = sql.Replace("{activeFilter}", activeFilter)
        sql = sql.Replace("{searchFilter}", searchFilter)

        ' 5. Fetch the data and build the tree
        Dim dtBasis As DataTable = Collect_data2(sql)

        If dtBasis IsNot Nothing AndAlso dtBasis.Rows.Count > 0 Then
            TreeViewMapper.Populate3LevelTree(tvBasis, dtBasis, "TargetTypeNode", "ContractNode", "DetailsNode")
        Else
            tvBasis.Nodes.Clear() ' Clear the tree if no matches are found
        End If
        ' 6. Build the tree
        If dtBasis IsNot Nothing AndAlso dtBasis.Rows.Count > 0 Then
            TreeViewMapper.Populate3LevelTree(tvBasis, dtBasis, "TargetTypeNode", "ContractNode", "DetailsNode")

            ' NEW: Expand the tree automatically if a searchword was used
            If Not String.IsNullOrWhiteSpace(searchword2) Then
                tvBasis.ExpandAll()
            End If
        Else
            tvBasis.Nodes.Clear()
        End If
    End Sub


End Module
