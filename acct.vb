Module acct


    '"Known error: 
    'bij het aanvinken van selecteer alle worden alle accounts getoond, niet alleen de subselectie. 



    Sub Fill_Cmx_Journal_List()

        For Each Ctl In SPAS.Gbx_Journal_Totals.Controls
            If Strings.Left(Ctl.Name, 11) = "Lbl_Journal" Then Ctl.Text = 0
        Next

        Dim k As String = "", lf As String = ""
        Dim t As String = SPAS.Cmx_Journal_List.Text
        Dim f As String = SPAS.Tbx_Journal_Filter.Text
        Dim act As Boolean = Not SPAS.Chbx_Journal_Inactive.Checked
        Dim stat As String = ""


        Load_Datagridview(SPAS.Dgv_Journal_items, "SELECT * FROM account WHERE name = 'xxxxxxxx'", "Fill_Cmx_Journal_List")

        Select Case t
            Case "Journaalnaam"
                Load_Listview(SPAS.Lv_Journal_List, "SELECT DISTINCT name, name FROM journal 
                                                WHERE name ILIKE '%" & f & "%' 
                                                ORDER BY name")

            Case "Kind", "Oudere", "Overig"

                Load_Listview(SPAS.Lv_Journal_List, "SELECT ac.id, ac.name, 
                                                CASE 
                                                WHEN Sum(j.amt1) is not distinct from null Then ac.startsaldo
                                                WHEN Sum(j.amt1) is  distinct from null Then Sum(j.amt1)
                                                End, 
                                                ac.startsaldo, ac.accgroup 
                                                From account ac
												LEFT JOIN journal j ON j.fk_account = ac.id
                                                LEFT JOIN target ta ON ta.id = ac.f_key 
                                                WHERE ac.f_key = ta.id 
                                                AND ta.ttype ILIKE '%" & t & "%' 
												AND ac.name ILIKE '%" & f & "%' 
                                                AND (ta.active = 'True' OR ta.active = '" & act & "')  
                                                Group by ac.id
                                                ORDER BY ac.name
                                                ")
            Case "Alle accounts"
                Load_Listview(SPAS.Lv_Journal_List, "SELECT ac.id, ac.name, 
                                                CASE 
                                                WHEN Sum(j.amt1) is not distinct from null Then ac.startsaldo
                                                WHEN Sum(j.amt1) is  distinct from null Then Sum(j.amt1)
                                                End, 
                                                ac.startsaldo, ac.accgroup
                                                FROM account ac
												LEFT JOIN journal j ON j.fk_account = ac.id
  												WHERE ac.name ILIKE '%" & f & "%' 
                                                AND (ac.active = 'True' OR ac.active = '" & act & "') 
                                                Group by ac.id
                                                ORDER BY ac.name
                                                ")
            Case "CP"
                Dim sqlstr = "
                                                SELECT ac.id, ac.name,  
                                                CASE 
                                                WHEN Sum(j.amt1) is not distinct from null Then ac.startsaldo
                                                WHEN Sum(j.amt1) is  distinct from null Then Sum(j.amt1)
                                                End, 
                                                ac.startsaldo, ac.accgroup From account ac
                                                Left JOIN journal j ON j.fk_account = ac.id
                                                LEFT JOIN cp ON cp.id = ac.f_key 
                                                Left Join target ta on fk_cp_id = cp.id
                                                Left Join contract co on fk_target_id = ta.id
                                                WHERE ac.f_key = cp.id 
       											AND ac.name ILIKE '%" & f & "%'
                                                AND 
                                                (cp.active = 'True' 
                                                OR cp.active = '" & act & "'
                                                OR co.enddate > '" & Date.Now & "' 
                                                ) 
                                                Group by ac.id
                                                ORDER BY ac.name"

                'Clipboard.SetText(sqlstr)
                Load_Listview(SPAS.Lv_Journal_List, sqlstr)


                ', Sum(j.amt1), ac.startsaldo

            Case "Categoriën"
                Load_Listview(SPAS.Lv_Journal_List, "SELECT ac.id, ac.name, 
                                                CASE 
                                                WHEN Sum(j.amt1) is not distinct from null Then ac.startsaldo
                                                WHEN Sum(j.amt1) is  distinct from null Then Sum(j.amt1) + ac.startsaldo
                                                End, 
                                                ac.startsaldo, ac.accgroup From account ac
												LEFT JOIN journal j ON j.fk_account = ac.id
                                                WHERE ac.source = 'cat' 
 												AND ac.name ILIKE '%" & f & "%'
                                                AND (ac.active = 'True' OR ac.active = '" & act & "') 
                                                Group by ac.id
                                                ORDER BY ac.name
                                                ")

            Case "Relation"
                MsgBox("nog niet geïmplementeerd")
            Case "Accountgroep"
                Load_Listview(SPAS.Lv_Journal_List, "SELECT ac.id, ac.name, 
                                                CASE 
                                                WHEN Sum(j.amt1) is not distinct from null Then ac.startsaldo
                                                WHEN Sum(j.amt1) is  distinct from null Then Sum(j.amt1)
                                                End, 
                                                ac.startsaldo, ac.accgroup
                                                FROM account ac
												LEFT JOIN journal j ON j.fk_account = ac.id
  												WHERE ac.accgroup ILIKE '%" & f & "%'
                                                AND (ac.active = 'True' OR ac.active = '" & act & "') 
                                                Group by ac.id
                                                ORDER BY ac.name
                                                ")
        End Select

        SPAS.Cbx_Journal_DeSelect_All.Checked = False
        SPAS.Cbx_Journal_Select_All.Checked = False

        With SPAS.Lv_Journal_List
            .Columns.Item(0).Width = 0
            .Columns(1).Text = "Naam"

            If t = "Journaalnaam" Then
                .Columns.Item(1).Width = 180
            Else
                .Columns(2).Text = "Saldo"
                .Columns.Item(1).Width = 150
                .Columns.Item(2).Width = 70
                .Columns.Item(3).Width = 0
                .Columns.Item(4).Width = 120
            End If

        End With


    End Sub

    '===============================================================================================
    Sub Select_Deselect_Accounts(ByVal sel As Boolean)
        For i = 0 To SPAS.Lv_Journal_List.Items.Count - 1
            SPAS.Lv_Journal_List.Items(i).Selected = sel
        Next
    End Sub

    Sub Select_Source_Account()
        Dim ErrMsg As String = ""
        Dim sel As Integer = 0
        Dim cs As Decimal = 0

        For i = 0 To SPAS.Lv_Journal_List.Items.Count - 1
            If SPAS.Lv_Journal_List.Items(i).Selected Then sel = sel + 1
        Next

        If SPAS.Cmx_Journal_List.Text = "Journaalnaam" Then ErrMsg &= vbCrLf & "Selecteer een account i.p.v. een journaalitem"
        If sel <> 1 Then ErrMsg &= vbCrLf & "Selecteer één account als bronaccount."

        If sel = 1 Then
            cs = Tbx2Dec(SPAS.Lv_Journal_List.FocusedItem.SubItems(2).Text)
        End If
        If cs = 0 And sel = 1 And SPAS.Cmx_Journal_List.Text <> "Journaalnaam" Then
            ErrMsg &= vbCrLf & "Het saldo van een bronaccount moet positief zijn."
        End If

        If ErrMsg <> "" Then
            MsgBox("Selectie van bronaccount is mislukt: " & ErrMsg)
            Exit Sub
        End If


        SPAS.Lbl_Journal_Source_id.Text = SPAS.Lv_Journal_List.FocusedItem.SubItems(0).Text
        SPAS.Lbl_Journal_Source_Name.Text = SPAS.Lv_Journal_List.FocusedItem.SubItems(1).Text
        SPAS.Lbl_Journal_Source_Saldo.Text = Tbx2Dec(cs)
        SPAS.Tbx_Journal_Source_Amt.Text = Tbx2Dec(cs)

    End Sub

    Sub test()


    End Sub

    Sub Select_Target_Account()

        Dim ErrMsg As String = ""
        Dim i As Integer
        Dim sel As Integer = 0
        Dim id
        Dim amt As Integer = 0
        Dim tgt_tot As Decimal = 0

        'what part of the amount is already allocated
        For i = 0 To SPAS.Dgv_Journal_Intern.Rows.Count - 1
            tgt_tot = tgt_tot + SPAS.Dgv_Journal_Intern.Rows(i).Cells(2).Value
        Next



        For i = 0 To SPAS.Lv_Journal_List.Items.Count - 1
            If SPAS.Lv_Journal_List.Items(i).Selected Then sel = sel + 1
        Next

        If SPAS.Cmx_Journal_List.Text = "Journaalnaam" Then ErrMsg &= vbCrLf & "Selecteer een account i.p.v. een journaalitem"
        If sel = 0 Then ErrMsg &= vbCrLf & "Selecteer minimaal één account als doelaccount."
        If Tbx2Dec(SPAS.Tbx_Journal_Source_Amt.Text) = 0 Then ErrMsg &= vbCrLf & "Selecteer eerst een bronaccount."

        If ErrMsg <> "" Then
            MsgBox("Selectie van doelaccount(s) is mislukt: " & ErrMsg)
            Exit Sub
        End If

        For i = 0 To SPAS.Lv_Journal_List.Items.Count - 1
            With SPAS.Lv_Journal_List.Items(i)

                If (.Selected) Then
                    amt = (Tbx2Dec(SPAS.Tbx_Journal_Source_Amt.Text) - tgt_tot) / sel
                    id = SPAS.Lv_Journal_List.Items(SPAS.Lv_Journal_List.Items(i).Index).SubItems(1).Text
                    SPAS.Dgv_Journal_Intern.Rows.Add(SPAS.Lv_Journal_List.Items(i).Text)
                    SPAS.Dgv_Journal_Intern.Rows(SPAS.Dgv_Journal_Intern.Rows.Count - 1).Cells(1).Value = id
                    SPAS.Dgv_Journal_Intern.Rows(SPAS.Dgv_Journal_Intern.Rows.Count - 1).Cells(2).Value = Tbx2Dec(amt)

                End If

            End With
        Next

        With SPAS.Dgv_Journal_Intern
            .Columns(0).Visible = False
            .Columns(1).Width = 130
            .Columns(1).ReadOnly = True
            .Columns(2).Width = 70
            .Columns(2).DefaultCellStyle.Format = "N2"
            .Columns(2).DefaultCellStyle.ForeColor = Color.Blue
            .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

        End With
        Calculate_Journal_Booking_Data()
    End Sub

    Sub Divide_among_targets()
        Dim cnt As Integer = 0
        Dim amt As Decimal = Tbx2Dec(SPAS.Tbx_Journal_Source_Amt.Text)

        'count number of selected target accounts
        For i = 0 To SPAS.Dgv_Journal_Intern.Rows.Count - 1
            cnt = cnt + 1
        Next

        For i = 0 To SPAS.Dgv_Journal_Intern.Rows.Count - 1
            SPAS.Dgv_Journal_Intern.Rows(i).Cells(2).Value = Int(amt / cnt)
        Next
        Calculate_Journal_Booking_Data()

    End Sub

    Sub Save_Internal_Booking()
        'checks
        'restbedrag mag niet negatief zijn
        'indien er een positief restbedrag is: waarschuwen
        Dim rst = Tbx2Int(SPAS.Lbl_Journal_Source_Restamt.Text)
        Dim act_cnt, act_cnt2 As Integer
        If rst < 0 Then
            MsgBox("Het te verdelen bedrag is hoger dan het saldo van de bronaccount.")
            Exit Sub
        ElseIf rst > 0 Then
            Dim answ = MsgBox("Een bedrag ad €" & rst & " is nog niet verdeeld. Wilt u doorgaan met bewaren?", vbYesNo)
            If answ = vbNo Then Exit Sub
        End If


        Dim SQLstr As String = ""
        Dim SQLroot As String = ""
        Dim name As String = "Intern " & SPAS.Lbl_Journal_Source_Name.Text


        act_cnt = QuerySQL("SELECT COUNT(Distinct(name)) FROM journal WHERE name ILIKE '%" & name & "%'")
        name &= "_" & act_cnt + 1
        act_cnt2 = QuerySQL("SELECT COUNT(Distinct(name)) FROM journal WHERE name ILIKE '%" & name & "%'")
        If act_cnt2 > 0 Then name &= "_" & DateTime.Now.Second
        'in theorie zou een naam niet uniek hoeven te zijn;
        Dim dat As Date = SPAS.Dtp_Journal_intern.Value
        Dim dat1 As String = dat.Year & "-" & dat.Month & "-" & dat.Day
        Dim src_amt As Integer = Cur2(Tbx2Int(SPAS.Tbx_Journal_Source_Amt.Text) - rst)
        Dim desc As String = SPAS.Tbx_Journal_Description.Text
        Dim fka As String = SPAS.Lbl_Journal_Source_id.Text

        'save source 
        SQLroot = "INSERT INTO journal(name,date,status,type,source,description,amt1,fk_account)
                   VALUES('" & name & "','" & dat1 & "'::date,'Verwerkt','Internal','Intern','" & desc & "','"

        SQLstr &= SQLroot & -src_amt & "','" & fka & "');"

        For i = 0 To SPAS.Dgv_Journal_Intern.Rows.Count - 1

            If SPAS.Dgv_Journal_Intern.Rows(i).Cells(2).Value > 0 Then
                SQLstr &= SQLroot & Cur2(SPAS.Dgv_Journal_Intern.Rows(i).Cells(2).Value) & "','" &
                SPAS.Dgv_Journal_Intern.Rows(i).Cells(0).Value & "');"
                '@@@hier gaat het fout met currency-conversie
            End If
            'nulwaarden overslaan
        Next i
        'Clipboard.SetText(SQLstr)
        RunSQL(SQLstr, "NULL", "Save_Internal_Booking")
        MsgBox("Deze interne boeking is opgeslagen met de naam " & name & ".")

        SPAS.Lbl_Journal_Source_Saldo.Text = 0
        SPAS.Lbl_Journal_Source_Name.Text = ""
        SPAS.Tbx_Journal_Source_Amt.Text = 0
        SPAS.Dgv_Journal_Intern.Rows.Clear()
        SPAS.Lbl_Journal_Source_Restamt.Text = 0

        SPAS.Cmx_Journal_List.Text = "Journaalnaam"
        SPAS.Tbx_Journal_Filter.Text = name


    End Sub
    Sub Add_Internal_Contract_Bookings(ByVal accnt1 As String, accnt2 As String, amt1 As Decimal,
                                       name As String, rel As String)
        'calculate the internal bookings date
        Dim d As Date, m_add As Integer, SQLstr As String
        If Strings.Len(accnt1) = 0 Then accnt1 = 0
        m_add = IIf(SPAS.Dtp_31_contract__startdate.Value.Day > 1, 1, 0)
        SQLstr = "INSERT INTO journal (amt1, fk_account,date, status,
                                    description,source,fk_relation,name,type) VALUES"

        For m = m_add + SPAS.Dtp_31_contract__startdate.Value.Month To 12

            d = CDate("01-" & m & "-" & SPAS.Dtp_Incasso_start.Value.Year)
            SQLstr &= "('" & -amt1 & "','" & accnt1 & "','" & d & "','Verwerkt','Gegenereerde interne contractboeking " & name & "','Internal','
            " & rel & "','Intern contract " & name & "','Contract'),"

            SQLstr &= "('" & amt1 & "','" & accnt2 & "','" & d & "','Verwerkt','Gegenereerde interne contractboeking" & name & "','Internal','
            " & rel & "','Intern contract " & name & "','Contract'),"

        Next m
        SQLstr = Strings.Left(SQLstr, Strings.Len(SQLstr) - 1)  'remove last comma

        'Clipboard.Clear()
        'Clipboard.SetText(SQLstr)
        RunSQL(SQLstr, "NULL", "Add_Internal_Contract_Bookings")

    End Sub


    Sub Calculate_Journal_Booking_Data()
        'calculate values of target accounts
        Dim tgt_tot As Decimal = 0
        For i = 0 To SPAS.Dgv_Journal_Intern.Rows.Count - 1
            tgt_tot = tgt_tot + SPAS.Dgv_Journal_Intern.Rows(i).Cells(2).Value
        Next

        SPAS.Lbl_Journal_Source_Restamt.Text = Tbx2Dec(Tbx2Dec(SPAS.Tbx_Journal_Source_Amt.Text) - tgt_tot)

    End Sub

    Sub Calculate_Journal_Overview()

        Dim tot_in As Integer
        Dim amt_in As Decimal = 0
        Dim amt_out As Decimal = 0
        Dim _amt1, _amt2
        Dim startsaldo As Decimal
        If SPAS.Lbl_Journal_Sum_Amt_Start.Text = "" Then
            startsaldo = 0
        Else
            startsaldo = Tbx2Dec(SPAS.Lbl_Journal_Sum_Amt_Start.Text)
        End If

        For i = 0 To SPAS.Dgv_Journal_items.Rows.Count - 1
            _amt1 = SPAS.Dgv_Journal_items.Rows(i).Cells(3).Value
            _amt2 = SPAS.Dgv_Journal_items.Rows(i).Cells(4).Value
            If IsDBNull(_amt1) Then _amt1 = 0  'prevent error on null values
            If IsDBNull(_amt2) Then _amt2 = 0
            amt_in = amt_in + IIf(Strings.Left(SPAS.Dgv_Journal_items.Rows(i).Cells(1).Value, 10) <> "Startsaldo",
                                 Tbx2Dec(_amt1), 0)
            amt_out = amt_out + Tbx2Dec(_amt2)
            If Tbx2Dec(_amt1) > 0 Then tot_in = tot_in + 1
            's = Me.Dgv_Journal_items.Rows(i).Cells(5).Value + Me.Dgv_Journal_items.Rows(i).Cells(4).Value
        Next

        SPAS.Lbl_Journal_Sum_Item_In.Text = tot_in
        SPAS.Lbl_Journal_Sum_Item_Out.Text = Tbx2Int(SPAS.Dgv_Journal_items.Rows.Count) - tot_in
        SPAS.Lbl_Journal_Sum_Item_Saldo.Text = Tbx2Int(SPAS.Dgv_Journal_items.Rows.Count)

        SPAS.Lbl_Journal_Sum_Amt_In.Text = Tbx2Dec(amt_in)
        SPAS.Lbl_Journal_Sum_Amt_Out.Text = Tbx2Dec(amt_out)
        SPAS.Lbl_Journal_Sum_Amt_Saldo.Text =
                Tbx2Dec(startsaldo + Tbx2Dec(amt_in) - Tbx2Dec(amt_out))  'Format(, "#.##")

    End Sub

    Function Create_Journal_SQL()

        Dim i As Integer
        Dim id
        Dim tbl, SQL_Where, SQLStr As String
        Dim stat As String = ""
        tbl = ""
        SQL_Where = ""
        SQLStr = ""

        Select Case SPAS.Cmx_Journal_List.Text
            Case "Relatie"
            Case "Incasso"
            Case "Uitkering"
            Case "Journaalnaam" : tbl = "j.name"
            Case Else : tbl = "ac.id"
        End Select

        If SPAS.Cbx_Journal_Status_Open.Checked And Not SPAS.Cbx_Journal_Status_Verwerkt.Checked Then
            stat = " AND j.status = 'Open' "
        ElseIf Not SPAS.Cbx_Journal_Status_Open.Checked And SPAS.Cbx_Journal_Status_Verwerkt.Checked Then
            stat = " AND j.status = 'Verwerkt' "
        End If


        SQLStr = "
             SELECT j.date::date As Datum, j.name As Name, j.status As Status,
             CASE 
 	            WHEN j.amt1::decimal > 0.00 THEN j.amt1
 	            WHEN j.amt1::decimal < 0.00 THEN '0'
             END As Bij,
              CASE 
 	            WHEN j.amt1::decimal < 0.00 THEN j.amt1::decimal * -1
 	            WHEN j.amt1::decimal > 0.00 THEN '0'
             END As Af,
             j.description As Omschrijving,
             ac.name As Account
             FROM journal j 
             LEFT JOIN account ac ON j.fk_account = ac.id
             LEFT JOIN relation r ON j.fk_relation = r.id
             LEFT JOIN target ta ON ta.id = ac.id  
             "


        Dim saldo As Decimal = 0.00
        For i = 0 To SPAS.Lv_Journal_List.Items.Count - 1
            With SPAS.Lv_Journal_List.Items(i)

                If (.Selected) Then
                    id = SPAS.Lv_Journal_List.Items(SPAS.Lv_Journal_List.Items(i).Index).SubItems(0).Text
                    SQL_Where &= IIf(SQL_Where = "", " WHERE ", " OR ") & tbl & "='" & id & "'" & stat
                    saldo = saldo +
                    Tbx2Dec(SPAS.Lv_Journal_List.Items(SPAS.Lv_Journal_List.Items(i).Index).SubItems(3).Text) +
                    Tbx2Dec(SPAS.Lv_Journal_List.Items(SPAS.Lv_Journal_List.Items(i).Index).SubItems(4).Text)
                End If
                '
                ' Dgv_Journal_Intern.DataSource = boundSet.Tables(0)
            End With
        Next
        SQLStr &= SQL_Where & " ORDER BY date"
        Clipboard.Clear()
        Clipboard.SetText(SQLStr)
        SPAS.Lbl_Journal_Sum_Amt_Start.Text = Tbx2Dec(saldo)
        Return SQLStr

    End Function

    Sub Fill_Journal_List()
        Dim jrnl As Boolean
        jrnl = SPAS.Cmx_Journal_List.Text = "Journaalnaam"


        Load_Datagridview(SPAS.Dgv_Journal_items, Create_Journal_SQL, "Lv_Journal_List_Click")
        With SPAS.Dgv_Journal_items
            .Columns(0).Width = 60
            .Columns(1).Width = 120
            .Columns(2).Width = 60
            .Columns(3).Width = 60
            .Columns(4).Width = 60
            .Columns(5).Width = 250 + IIf(jrnl, 70, 0)
            .Columns(6).Width = 110
            .Columns(0).HeaderText = "Datum"
            .Columns(1).HeaderText = "Naam"
            .Columns(2).HeaderText = "Status"
            .Columns(3).HeaderText = "Bij"
            .Columns(4).HeaderText = "Af"
            .Columns(5).HeaderText = "Omschrijving"
            .Columns(6).HeaderText = "Account"

            .Columns(0).DefaultCellStyle.Format = "dd-MM"
            .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            .Columns(3).DefaultCellStyle.Format = "N2"
            .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            .Columns(4).DefaultCellStyle.Format = "N2"

            .Columns(1).Visible = Not jrnl
            .Columns(6).Visible = jrnl
            .Columns(2).Visible = Not jrnl
        End With

        'Calculate_Journal_Totals()
        Calculate_Journal_Overview()
    End Sub

    Sub Calculate_Budget(ByVal id As String)


        Dim SQLStr As String = ""
        Dim sd As String
        Dim mon As String = ""
        Dim where As String = ""
        If id <> "" Then where = "WHERE ac1.id='" & id & "'"


        For m = 1 To 12
            sd = Date.Now.Year & "-" & m & "-01"            'ed = CDate("01-" & m + 1 & "-" & Date.Now.Year).AddDays(-1)
            Select Case m
                Case 1 : mon = "b_jan"
                Case 2 : mon = "b_feb"
                Case 3 : mon = "b_mar"
                Case 4 : mon = "b_apr"
                Case 5 : mon = "b_may"
                Case 6 : mon = "b_jun"
                Case 7 : mon = "b_jul"
                Case 8 : mon = "b_aug"
                Case 9 : mon = "b_sep"
                Case 10 : mon = "b_oct"
                Case 11 : mon = "b_nov"
                Case 12 : mon = "b_dec"
            End Select

            SQLStr &= "
                    UPDATE account ac1
                    SET " & mon & "=
                    (
                    SELECT  co.donation/co.term
                    FROM contract co
                    LEFT JOIN target ta ON co.fk_target_id = ta.id
                    LEFT JOIN account ac ON ac.f_key = ta.id
                    WHERE co.startdate <='" & sd & "' 
                    AND co.enddate > '" & sd & "'
                    AND ac1.f_key = ta.id
                    )
                    " & where & ";
"

        Next

        'MsgBox(SQLStr)
        Clipboard.Clear()
        Clipboard.SetText(SQLStr)
        RunSQL(SQLStr, "NULL", "Calculate_Budget")
        RunSQL(Budget_Year_Totals, "NULL", "Calculate_Budget/Budget_Year_Totals")

    End Sub

    Function Budget_Year_Totals()
        Dim SQLStr As String = "
                    UPDATE account
                    SET b_year=
                    (
                    Select 
							(CASE
								WHEN b_jan IS NOT NULL Then b_jan
								WHEN b_jan IS NULL Then '0'
							END) +
							(CASE
								WHEN b_feb IS NOT NULL Then b_feb
								WHEN b_feb IS NULL Then '0'
							END) +
							(CASE
								WHEN b_mar IS NOT NULL Then b_mar
								WHEN b_mar IS NULL Then '0'
							END) +
							(CASE
								WHEN b_apr IS NOT NULL Then b_apr
								WHEN b_apr IS NULL Then '0'
							END)+
							(CASE
								WHEN b_may IS NOT NULL Then b_may
								WHEN b_may IS NULL Then '0'
							END) +
							(CASE
								WHEN b_jun IS NOT NULL Then b_jun
								WHEN b_jun IS NULL Then '0'
							END) +
							(CASE
								WHEN b_jul IS NOT NULL Then b_jul
								WHEN b_jul IS NULL Then '0'
							END) +
							(CASE
								WHEN b_aug IS NOT NULL Then b_aug
								WHEN b_aug IS NULL Then '0'
							END) +
							(CASE
								WHEN b_sep IS NOT NULL Then b_sep
								WHEN b_sep IS NULL Then '0'
							END) +
							(CASE
								WHEN b_oct IS NOT NULL Then b_oct
								WHEN b_oct IS NULL Then '0'
							END) +
							(CASE
								WHEN b_nov IS NOT NULL Then b_nov
								WHEN b_nov IS NULL Then '0'
							END) +
							(CASE
								WHEN b_dec IS NOT NULL Then b_dec
								WHEN b_dec IS NULL Then '0'
							END)
                            ); 

"
        Return SQLStr


    End Function

    Sub Calculate_Manual_Budgets()


        SPAS.Lbl_Account_Budget_Difference.Text = Tbx2Dec(
                    GetDouble(SPAS.Tbx_10_Account__b_year.Text) -
                    (
                    GetDouble(SPAS.Tbx_10_Account__b_jan.Text) +
                    GetDouble(SPAS.Tbx_10_Account__b_feb.Text) +
                    GetDouble(SPAS.Tbx_10_Account__b_mar.Text) +
                    GetDouble(SPAS.Tbx_10_Account__b_apr.Text) +
                    GetDouble(SPAS.Tbx_10_Account__b_may.Text) +
                    GetDouble(SPAS.Tbx_10_Account__b_jun.Text) +
                    GetDouble(SPAS.Tbx_10_Account__b_jul.Text) +
                    GetDouble(SPAS.Tbx_10_Account__b_aug.Text) +
                    GetDouble(SPAS.Tbx_10_Account__b_sep.Text) +
                    GetDouble(SPAS.Tbx_10_Account__b_oct.Text) +
                    GetDouble(SPAS.Tbx_10_Account__b_nov.Text) +
                    GetDouble(SPAS.Tbx_10_Account__b_dec.Text))
        )

        'SPAS.Lbl_Account_Budget_Difference.Text Then SPAS.Lbl_Account_Budget_Difference.ForeColor = Color.Red

    End Sub

End Module
