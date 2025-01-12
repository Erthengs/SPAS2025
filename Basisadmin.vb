Imports System.IO
Imports Npgsql
Module Basisadmin
    Public group As String
    Public reload As Boolean = False


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

        Collect_data("
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

        For i = 0 To dst.Tables(0).Columns.Count - 1

            qty = IIf(IsDBNull(dst.Tables(0).Rows(0)(i)), 0, dst.Tables(0).Rows(0)(i))

            SPAS.Dgv_Mgnt_Tables.ColumnCount = 2
            SPAS.Dgv_Mgnt_Tables.Columns(0).Name = "Tabel"
            SPAS.Dgv_Mgnt_Tables.Columns(1).Name = "Aantal records"
            SPAS.Dgv_Mgnt_Tables.Rows.Add(dst.Tables(0).Columns(i).ColumnName)
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

        '@@@deze module moet nog verbeterd worden via gebruik van een dataset en het kunnen hanteren van 0-waarden


        'A GENERIC PART ===========================================================================
        Dim fld, fk_tbl As String, id, fk_id, pos, pos1, pos2 As Integer
        Dim tb As Integer = SPAS.TC_Object.SelectedIndex
        Dim tbl As String = SPAS.TC_Object.TabPages(tb).Name
        Dim tmp
        Dim col As Integer = -1
        Edit_Mode = False
        'SPAS.Manage_Buttons_Target(True, True, True, False, False, "Select_Obj2")

        Try
            id = SPAS.Lbx_Basis.Items(SPAS.Lbx_Basis.SelectedIndex)(SPAS.Lbx_Basis.ValueMember)
        Catch ex As Exception
            Exit Sub
        End Try
        Collect_data("SELECT * FROM " & tbl & " WHERE id='" & id & "'")

        Empty_Tabpage()

        For Each ctl In SPAS.TC_Object.TabPages(tb).Controls

            If Strings.InStr(ctl.Name, "_pkid") > 0 Then ctl.Text = id

            pos = Strings.InStr(ctl.Name, "__")
            If pos > 0 Then
                fld = Mid(ctl.Name, pos + 2, Len(ctl.Name) - pos)
                'retrieve the accompanying columns
                For i = 0 To dst.Tables(0).Columns.Count - 1
                    If fld = dst.Tables(0).Columns(i).ColumnName Then
                        col = i
                        Exit For
                    End If
                Next
                If col = -1 Then Exit Sub

                If TypeOf ctl Is TextBox Or TypeOf ctl Is Label Then
                    Select Case Strings.Mid(ctl.Name, 5, 1)
                        Case 1
                            tmp = dst.Tables(0).Rows(0)(col)
                            If IsDBNull(tmp) Then
                                ctl.Text = 0
                            Else
                                ctl.Text = tmp
                                '@@@currency_converter
                            End If
                        Case Else

                            If IsDBNull(dst.Tables(0).Rows(0)(col)) Then ctl.Text = "" Else ctl.Text = dst.Tables(0).Rows(0)(col)

                            'ctl.Text = dst.Tables(0).Rows(0)(col)
                            'If IsDBNull(ctl.Text) Then ctl.Text = ""
                    End Select
                ElseIf TypeOf ctl Is CheckBox Then
                    'Clipboard.Clear()
                    'Clipboard.SetText(ctl.Name)
                    ctl.Checked = dst.Tables(0).Rows(0)(col)
                ElseIf TypeOf ctl Is PictureBox Then

                    Dim img As Image
                    Try
                        Dim photo = dst.Tables(0).Rows(0)(col) 'QuerySQL("SELECT " & fld & " FROM " & tbl & " WHERE id='" & id & "'")
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
                        fk_id = dst.Tables(0).Rows(0)(col) 'QuerySQL("SELECT " & fld & " FROM " & tbl & " WHERE id='" & id & "'")
                        fk_tbl = Mid(fld, 4, Len(fld) - 6) ', pos2 - pos1

                        If fk_tbl = "bank" Then fk_tbl = "bankacc"
                        If fk_tbl = "acco" Then fk_tbl = "account"

                        If fk_tbl = "bankacc" Or fk_tbl = "account" Or fk_tbl = "accgroup" Then  '@@@ workaround

                            ctl.Text = QuerySQL("SELECT name FROM " & fk_tbl & " WHERE id='" & fk_id & "'")
                        Else
                            ctl.Text = QuerySQL("SELECT Concat(name, ', ', name_add) as name 
                            FROM " & fk_tbl & " WHERE id='" & fk_id & "'")
                        End If
                    Else
                        ctl.Text = dst.Tables(0).Rows(0)(col).ToString
                    End If

                ElseIf TypeOf ctl Is DateTimePicker Then
                    ctl.Value = dst.Tables(0).Rows(0)(col)
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
            Collect_data(sqlstr)
            'SPAS.Tbx_Contract_ttype.Text = dst.Tables(0).Rows(0)(0)
            If dst.Tables(0).Rows(0)(0) = "Kind" Then SPAS.Rbn_00_contract_child.Checked = True
            If dst.Tables(0).Rows(0)(0) = "Oudere" Then SPAS.Rbn_00_contract_elder.Checked = True
            If dst.Tables(0).Rows(0)(0) = "Overig" Then SPAS.Rbn_00_contract_other.Checked = True

            'SPAS.Dtp_30_Contract_Change.Value = Date.Today
            SPAS.Lbl_Contract_Bronaccount.Visible = Not IsDBNull(dst.Tables(0).Rows(0)(2))
            SPAS.Cmx_00_Contract__fk_account_id.Visible = Not IsDBNull(dst.Tables(0).Rows(0)(2))
            SPAS.Lbl_Contract_tgt.Text = SPAS.Cmx_01_contract__fk_target_id.Text
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
        If tb = 5 Then 'bankACCOUNT

            SPAS.Tbx_BankAcc_startbalance.Text = QuerySQL("select credit-debit from bank b left join bankacc c on c.accountno=b.iban 
                                                where b.name='_startsaldo_' and b.iban ='" & SPAS.Tbx_01_BankAcc__accountno.Text & "'")

        End If

    End Sub

    Sub Load_Table()
        Dim tb As Integer = SPAS.TC_Object.SelectedIndex
        Dim tbl As String = SPAS.TC_Object.TabPages(tb).Name
        Dim SQLstr, SQLstr1, SQLstr2 As String

        ' Dim arg = SPAS.Tbx_Basis_Filter.Text.ToUpper
        Dim arg = SPAS.Searchbox.Text.ToUpper
        Dim sel_act As String = ""
        If SPAS.Cbx_LifeCycle.Text = "Actief" Then
            sel_act = " AND active=True"
        End If
        If SPAS.Cbx_LifeCycle.Text = "Inactief" Then
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
                          WHERE contract.active=" & IIf(SPAS.Cbx_LifeCycle.Text = "Inactief", False, True) & "
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
                        WHERE t.active=" & IIf(SPAS.Cbx_LifeCycle.Text = "Inactief", False, True) & " 
                         " & filtersql & "
                        ORDER BY t.name"

            Load_Listbox(SPAS.Lbx_Basis, SQLstr)

        ElseIf tb = 4 Then
            SQLstr1 = "SELECT id, CONCAT(source,' ',name,' (',accgroup,')') as name FROM " & tbl & " 
                       WHERE (name iLike '%" & arg & "%'" & sel_act & " 
                       OR accgroup iLike '%" & arg & "%' 
                       OR source iLike '%" & arg & "%')
                       AND (active=" & IIf(SPAS.Cbx_LifeCycle.Text = "Inactief", False, True) & ") 
                       ORDER BY source, accgroup, name"
            Load_Listbox(SPAS.Lbx_Basis, SQLstr1)


        ElseIf tb = 5 Then
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
                errmsg1 &= "- " & f.Tag & " mag niet leeg zijn" & vbCrLf   '' Mid(f.Tag, pos + 2)
            End If

        Next
        ix = SPAS.TC_Object.SelectedIndex
        Select Case ix

            Case 0
                'contract: control that either sponsor or intern account is selected
                If SPAS.Cmx_00_Contract__fk_account_id.Text = "" And SPAS.Cmx_00_contract__fk_relation_id.Text = "" Then
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
                    AND fk_target_id ='" & SPAS.Cmx_01_contract__fk_target_id.SelectedValue & "'
                    AND fk_relation_id ='" & SPAS.Cmx_00_contract__fk_relation_id.SelectedValue & "'"

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
                    'Clipboard.Clear()
                    'Clipboard.SetText("SELECT count(*) from " & tbl &
                    '" WHERE name='" & nm & "' AND name_add='" & nma & "'")
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

        RunSQL(SQLstr, "NULL", "Insert into table " & tbl)

        'addition for contract;  target  - --
        Select Case tb
            Case 0
                SPAS.Pan_contract_select_target.Enabled = False
                Dim Source_Account = QuerySQL("SELECT id FROM account WHERE f_key='" & SPAS.Cmx_01_contract__fk_target_id.SelectedValue & "'")
                Calculate_Budget(Source_Account)
            Case 2
                Load_Combobox(SPAS.Cmx_00_contract__fk_relation_id, "id", "name", "SELECT id, CONCAT(name, ', ', name_add) as name FROM relation WHERE active=TRUE ORDER BY name")
            Case 1, 3 'creating an account for target or cp...
                Load_Combobox(SPAS.Cmx_01_Target__fk_cp_id, "id", "name", "SELECT id, CONCAT(name, ', ', name_add) as name FROM cp WHERE active=True ORDER BY name")
                new_id = QuerySQL("Select Max(id) From " & tbl)
                Dim tbtxt As String = SPAS.TC_Object.TabPages(tb).Tag
                SQLstr = "SELECT CONCAT(name,',', name_add) FROM " & tbl & " WHERE id=" & new_id
                If SPAS.Chbx_test.Checked Then MsgBox(SQLstr)
                name = QuerySQL(SQLstr)
                Dim accgroup As String
                'Dim tt As String = QuerySQL("SELECT ttype " & tbl & " WHERE id=" & new_id)

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
        SPAS.Cmx_01_contract__fk_target_id.Enabled = Add_Mode
        SPAS.Cmx_00_contract__fk_relation_id.Enabled = Add_Mode
        SPAS.Cmx_00_Contract__fk_account_id.Enabled = Add_Mode

    End Sub
    Sub Create_Contract_Version()
        'SQLstr = "INSERT INTO contract"

    End Sub
    Sub Get_Sponsor_data()

        Dim rel_id = SPAS.Cmx_00_contract__fk_relation_id.SelectedValue
        Dim d As String = "date"

        If SPAS.Rbn_00_contract_child.Checked Then
            d &= "1"
        ElseIf SPAS.Rbn_00_contract_elder.Checked Then
            d &= "2"
        Else
            d &= "3"
        End If

        SPAS.Lbl_00_contract_autcol.Text = QuerySQL("SELECT reference FROM relation WHERE id=" & rel_id)
        SPAS.dtp_contract_relation_date.Value = QuerySQL("SELECT " & d & " FROM relation WHERE id=" & rel_id)
        SPAS.Chx_00_contract__autcol.Checked =
        SPAS.dtp_contract_relation_date.Value < SPAS.Dtp_31_contract__startdate.Value

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
            SELECT ta.ttype, count (distinct r.id), 
            sum((co.donation+co.overhead)/term)
            FROM contract co 
            LEFT JOIN Target ta ON co.fk_target_id = ta.id
            LEFT JOIN Relation r ON co.fk_relation_id = r.id
            WHERE co.autcol = True 
            AND co.startdate <= '" & date_start & "' 
            AND co.enddate > '" & date_start & "'
            AND 
            ((r.date1 <='" & date_start & "' AND ta.ttype = 'Kind') OR
            (r.date2 <='" & date_start & "' AND ta.ttype = 'Oudere') OR
            (r.date3 <='" & date_start & "' AND ta.ttype = 'Overig'))
            GROUP BY ta.ttype
"
        '        Return SQLstr

        SQLstr = "
            Select 'Kind',  count (distinct r.id),sum((co.donation+co.overhead)/term)
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

    Function Create_Incasso_Bookings(date_start As String)
        Dim SQLstr As String = "
            SELECT 
                Concat(r.name, ', ',r.name_add) As Sponsor, 
                ta.name||', '||ta.name_add As Doel, 
                co.name As Contractnr, 
                ta.ttype As Doeltype, 
                sum(co.donation/co.term) As Donatie,
                sum(co.overhead/co.term) As overhead,
                ac.id As Accountid, 
                r.id As Sponsorid
            FROM contract co 
                LEFT JOIN Target ta ON co.fk_target_id = ta.id
                LEFT JOIN Relation r ON co.fk_relation_id = r.id
                LEFT JOIN Account ac ON ac.f_key = ta.id
            WHERE co.autcol = True 
            AND co.startdate <= '" & date_start & "' 
            AND co.enddate > '" & date_start & "'
                AND 
            ((r.date1 <='" & date_start & "' AND ta.ttype = 'Kind') OR
            (r.date2 <='" & date_start & "' AND ta.ttype = 'Oudere') OR
            (r.date3 <='" & date_start & "' AND ta.ttype = 'Overig'))

            GROUP BY  ac.id,r.id,ta.name,ta.name_add, co.name, r.reference, r.name, r.name_add, r.iban, ta.ttype, r.date1, r.date2
            ORDER by  ta.ttype, r.reference
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
        'Clipboard.Clear()
        'Clipboard.SetText(SQLstr)
        Return SQLstr


    End Function
    'concat(ta.name, ta.name_add)

    Sub Set_Budgets()
        Dim SQLstr = "


        update account Set 
        b_jan = 0, b_feb=0, b_mar=0, b_apr=0,b_may = 0, b_jun=0, b_jul=0, b_aug=0,b_sep = 0, b_oct=0, b_nov=0, b_dec=0;
        UPDATE account c1   
        SET 
        b_jan = (co.donation+co.overhead)/co.term,
        b_feb = (co.donation+co.overhead)/co.term,
        b_mar = (co.donation+co.overhead)/co.term,
        b_apr = (co.donation+co.overhead)/co.term,
        b_may = (co.donation+co.overhead)/co.term,
        b_jun = (co.donation+co.overhead)/co.term,
        b_jul = (co.donation+co.overhead)/co.term,
        b_aug = (co.donation+co.overhead)/co.term,
        b_sep = (co.donation+co.overhead)/co.term,
        b_oct = (co.donation+co.overhead)/co.term,
        b_nov = (co.donation+co.overhead)/co.term,
        b_dec = (co.donation+co.overhead)/co.term
        FROM account 
        LEFT JOIN contract co ON co.fk_target_id = account.f_key
        LEFT JOIN account c2 ON co.fk_target_id = c2.f_key
        WHERE c1.f_key = c2.f_key
        AND c1.f_key = co.fk_target_id;
"


    End Sub

    Function E2D(ByVal amt)
        Dim camt = FormatCurrency(amt, 2)

        Return camt
    End Function

End Module
