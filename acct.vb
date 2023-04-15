Imports System.ComponentModel.DataAnnotations

Module acct


    '"Known error: 
    'bij het aanvinken van selecteer alle worden alle accounts getoond, niet alleen de subselectie. 



    Sub Fill_Cmx_Journal_List()

        Dim k As String = "", lf As String = ""
        Dim t As String = SPAS.Cmx_Journal_List.Text
        Dim f As String = SPAS.Searchbox.Text 'SPAS.Tbx_Journal_Filter.Text
        Dim act As Boolean = (SPAS.Cbx_LifeCycle.Text = "Actief")  'Not SPAS.Chbx_Journal_Inactive.Checked
        Dim verwerkt As Boolean = SPAS.Cbx_Journal_Status_Verwerkt.Checked
        Dim open As Boolean = SPAS.Cbx_Journal_Status_Open.Checked
        Dim sqlstr As String
        Dim ttype As String = ""
        Dim nulsaldo As String = ""

        Dim st As String = " AND (j.status "
        If open And verwerkt Then st &= "IN ('Open','Verwerkt') or j.status isnull)"
        If open And Not verwerkt Then st &= "IN ('Open') or j.status isnull)"
        If Not open And verwerkt Then st &= "IN ('Verwerkt') or j.status isnull)"
        If Not open And Not verwerkt Then st &= "Not IN ('Open','Verwerkt') or j.status isnull)"
        If SPAS.Cbx_Journal_Saldo_Open.Checked Then nulsaldo = "having sum(amt1) !=0::money "


        Load_Datagridview(SPAS.Dgv_Journal_items, "SELECT * FROM account WHERE name = 'xxxxxxxx'", "Fill_Cmx_Journal_List")

        Select Case t
            Case "Journaalnaam"
                Load_Listview(SPAS.Lv_Journal_List, "SELECT DISTINCT name, name, date FROM journal 
                                                WHERE name ILIKE '%" & f & "%' 
                                                ORDER BY date desc, name")

            Case "Alle accounts", "Kind", "Oudere", "Overig"

                ttype = "AND ta.ttype ILIKE '%" & t & "%' "
                If t = "Alle accounts" Then ttype = ""

                sqlstr = "
                                          SELECT ac.id, ac.name As Accountnaam, 
                                                CASE WHEN Sum(j.amt1) is  distinct from null Then Sum(j.amt1) else 0::money End
                                                ,(select sum(amt1) from journal j2 where j2.source='Closing' and j2.fk_account = ac.id) As Startsaldo
                                                ,ac.accgroup As Group
                                                From account ac
												LEFT JOIN journal j ON j.fk_account = ac.id
                                                LEFT JOIN target ta ON ta.id = ac.f_key 
                                                WHERE ac.name ILIKE '%" & f & "%' " & ttype & st & "
                                                Group by ac.id " & nulsaldo & "
                                                ORDER BY ac.name
                                               "
                Clipboard.SetText(sqlstr)

                Load_Listview(SPAS.Lv_Journal_List, sqlstr)

            Case "CP"
                sqlstr = "
                                                SELECT ac.id, ac.name,  
                                                CASE WHEN Sum(j.amt1) is not distinct from null Then ac.startsaldo
                                                WHEN Sum(j.amt1) is  distinct from null Then Sum(j.amt1) End, 
                                                ac.startsaldo, ac.accgroup From account ac
                                                Left JOIN journal j ON j.fk_account = ac.id
                                                LEFT JOIN cp ON cp.id = ac.f_key 
                                                Left Join target ta on fk_cp_id = cp.id
                                                Left Join contract co on fk_target_id = ta.id
                                                WHERE ac.f_key = cp.id 
       											AND ac.name ILIKE '%" & f & "%'" & st & "
                                                AND 
                                                (cp.active = 'True' 
                                                --OR cp.active = '" & act & "'
                                                --OR co.enddate > CURRENT_DATE
                                                ) 
                                                Group by ac.id
                                                ORDER BY ac.name"


                Load_Listview(SPAS.Lv_Journal_List, sqlstr)


                ', Sum(j.amt1), ac.startsaldo

            Case "Categoriën"
                Load_Listview(SPAS.Lv_Journal_List, "
                                                SELECT ac.id, ac.name, 
                                                CASE WHEN Sum(j.amt1) is  distinct from null Then Sum(j.amt1) else 0::money End 
                                                ,(select sum(amt1) from journal j2 where j2.source='Closing' and j2.fk_account = ac.id)
                                                , ac.accgroup
                                                FROM account ac
												LEFT JOIN journal j ON j.fk_account = ac.id
                                                WHERE ac.source = 'cat' 
 												AND ac.name ILIKE '%" & f & "%'
                                                AND (ac.active = 'True' OR ac.active = '" & act & "') " & st & "
                                                Group by ac.id
                                                ORDER BY ac.name
                                                ")

            Case "Relaties"
                'MsgBox("nog niet geïmplementeerd")
                Load_Listview(SPAS.Lv_Journal_List, "
                                                Select r.id, r.name||', '||r.name_add, 
                                                --CASE WHEN Sum(j.amt1) is not distinct from null Then sum(j.amt1) ELSE '$0.00' END,
                                                sum(amt1),null, null
                                                FROM relation r 
                                                LEFT Join journal j on r.id=j.fk_relation
                                                WHERE j.source NOT in ('Uitkering', 'Intern')
                                                group by r.id, r.name_add, r.name
                                                order by r.name
")

            Case "Accountgroep"
                Load_Listview(SPAS.Lv_Journal_List, "
                                                SELECT ac.id, ac.name, 
                                                CASE WHEN Sum(j.amt1) is  distinct from null Then Sum(j.amt1) else 0::money End 
                                                ,(select sum(amt1) from journal j2 where j2.source='Closing' and j2.fk_account = ac.id)
                                                , ac.accgroup
                                                FROM account ac
												LEFT JOIN journal j ON j.fk_account = ac.id
  												WHERE ac.accgroup ILIKE '%" & f & "%'
                                                AND (ac.active = 'True' OR ac.active = '" & act & "') " & st & "
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
                .Columns.Item(1).Width = 150
                .Columns.Item(2).Width = 100
                .Columns(2).Text = "Date"

            Else
                .Columns(2).Text = "Saldo"
                .Columns.Item(1).Width = 150
                .Columns.Item(2).Width = 70
                .Columns.Item(3).Width = 70
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
        SPAS.Tbx_Journal_Name.Text = SPAS.Lbl_Journal_Source_Name.Text & ">"

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
            .Columns(1).Width = 160
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
        'Dim act_cnt, act_cnt2 As Integer
        Dim err As String = ""
        Dim rst = SPAS.Lbl_Journal_Source_Restamt.Text
        Dim name As String = SPAS.Tbx_Journal_Name.Text
        If QuerySQL("SELECT COUNT(Distinct(name)) FROM journal WHERE name ILIKE '%" & name & "%'") > 0 Then
            name = name & "_" & DateTime.Now.Day & "." & Date.Now.Month & "." & Right(DateTime.Now.Year, 2) & ":" & DateTime.Now.Second '"Intern " & SPAS.Lbl_Journal_Source_Name.Text
        End If

        If SPAS.Dgv_Journal_Intern.RowCount = 0 Then err &= "Er is geen doelaccount geselecteerd" & vbCr
        If rst < 0 Then err = "Het te verdelen bedrag is hoger dan het saldo van de bronaccount." & vbCr

        If err <> "" Then
            MsgBox(err)
            Exit Sub
        End If

        If rst > 0 Then
            Dim answ = MsgBox("Een bedrag ad €" & rst & " is nog niet verdeeld. Wilt u doorgaan met bewaren?", vbYesNo)
            If answ = vbNo Then Exit Sub
        End If

        Dim SQLstr As String = ""
        Dim SQLroot As String = ""

        Dim dat As Date = SPAS.Dtp_Journal_intern.Value
        Dim dat1 As String = dat.Year & "-" & dat.Month & "-" & dat.Day
        Dim src_amt As Integer '= Cur2(Tbx2Int(SPAS.Tbx_Journal_Source_Amt.Text) - rst)
        Dim desc As String = SPAS.Tbx_Journal_Description.Text
        Dim fka As String = SPAS.Lbl_Journal_Source_id.Text
        Dim type As String = IIf(SPAS.Rbn_Journal_Intern.Checked, "'Internal'", IIf(SPAS.Rbn_Journal_Contract.Checked, "'Contract'", "'Extra'"))
        'save source 
        SQLroot = "INSERT INTO journal(name,date,status,type,source,description,amt1,fk_account)
                   VALUES('" & name & "','" & dat1 & "'::date,'Verwerkt'," & type & ",'Intern','" & desc & "','"



        For i = 0 To SPAS.Dgv_Journal_Intern.Rows.Count - 1

            If SPAS.Dgv_Journal_Intern.Rows(i).Cells(2).Value > 0 Then
                SQLstr &= SQLroot & Cur2(CLng(SPAS.Dgv_Journal_Intern.Rows(i).Cells(2).Value)) & "','" &
                SPAS.Dgv_Journal_Intern.Rows(i).Cells(0).Value & "');"

            End If
            'nulwaarden overslaan
            src_amt = src_amt + Cur2(CLng(SPAS.Dgv_Journal_Intern.Rows(i).Cells(2).Value))
        Next i
        SQLstr &= SQLroot & -Cur2(Tbx2Int(src_amt)) & "','" & fka & "');"
        Clipboard.Clear()
        Clipboard.SetText(SQLstr)

        RunSQL(SQLstr, "NULL", "Save_Internal_Booking")
        MsgBox("Deze interne boeking is opgeslagen met de naam " & name & ".")

        SPAS.Lbl_Journal_Source_Saldo.Text = 0
        SPAS.Lbl_Journal_Source_Name.Text = ""
        SPAS.Tbx_Journal_Source_Amt.Text = 0
        SPAS.Dgv_Journal_Intern.Rows.Clear()
        SPAS.Lbl_Journal_Source_Restamt.Text = 0
        SPAS.Tbx_Journal_Description.Text = ""
        SPAS.Dtp_Journal_intern.Value = Date.Today

        SPAS.Cmx_Journal_List.Text = "Journaalnaam"
        SPAS.Searchbox.Text = name
        SPAS.TC_Boeking.SelectedIndex = 0


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
        Dim tname As String
        For i = 0 To SPAS.Dgv_Journal_Intern.Rows.Count - 1
            tgt_tot = tgt_tot + SPAS.Dgv_Journal_Intern.Rows(i).Cells(2).Value
            tname = SPAS.Dgv_Journal_Intern.Rows(i).Cells(1).Value
        Next

        SPAS.Lbl_Journal_Source_Restamt.Text = Tbx2Dec(Tbx2Dec(SPAS.Tbx_Journal_Source_Amt.Text) - tgt_tot)


        If SPAS.Dgv_Journal_Intern.Rows.Count > 1 Then
            SPAS.Tbx_Journal_Name.Text = SPAS.Lbl_Journal_Source_Name.Text & ">" & tname & "+" & SPAS.Dgv_Journal_Intern.Rows.Count - 1
        Else
            SPAS.Tbx_Journal_Name.Text = SPAS.Lbl_Journal_Source_Name.Text & ">" & tname
        End If


    End Sub


    Function Create_Journal_SQL()

        Dim i As Integer
        Dim id
        Dim dat
        Dim tbl, SQL_Where, SQLStr, dateselect As String
        Dim stat As String = ""
        tbl = ""
        SQL_Where = ""
        SQLStr = ""
        Dim jrnl = SPAS.Cmx_Journal_List.Text = "Journaalnaam"

        Select Case SPAS.Cmx_Journal_List.Text
            Case "Relaties" : tbl = "j.fk_relation"
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
             SELECT j.date::date As Datum, j.name As Name, 
             CASE 
 	            WHEN j.amt1::decimal > 0.00 THEN j.amt1
 	            WHEN j.amt1::decimal < 0.00 THEN '0'
             END As Bij,
              CASE 
 	            WHEN j.amt1::decimal < 0.00 THEN j.amt1::decimal * -1
 	            WHEN j.amt1::decimal > 0.00 THEN '0'
             END As Af,
             TRIM(j.description) As Omschrijving,
             ac.name As Account, 
             j.status As Status,
             j.source As Bron,
            substring(j.iban,5,3)||substring(j.iban,15,4) As IBAN,
            j.type As Soort,
             --,cp.name As CP
             --,r.name||','||r.name_add As Relatie,
             --,b.name||'/'||b.description||b.batchid As Bankinfo
             j.id As Id
             FROM journal j 
             LEFT JOIN account ac ON j.fk_account = ac.id
             LEFT JOIN relation r ON j.fk_relation = r.id
             LEFT JOIN target ta ON ta.id = ac.id  
             LEFT JOIN bank b ON b.id = j.fk_bank
             LEFT JOIN cp ON cp.id = ta.fk_cp_id
             "


        'Dim saldo As Decimal = 0.00
        For i = 0 To SPAS.Lv_Journal_List.Items.Count - 1
            With SPAS.Lv_Journal_List.Items(i)

                If (.Selected) Then
                    id = SPAS.Lv_Journal_List.Items(SPAS.Lv_Journal_List.Items(i).Index).SubItems(0).Text

                    If jrnl Then
                        Dim _dat As Date
                        _dat = SPAS.Lv_Journal_List.Items(SPAS.Lv_Journal_List.Items(i).Index).SubItems(2).Text
                        dat = _dat.Year & "-" & _dat.Month & "-" & _dat.Day
                        dateselect = " and j.date ='" & dat & "'::date"
                    Else
                        dateselect = ""
                    End If
                    SQL_Where &= IIf(SQL_Where = "", " WHERE ", " OR ") & tbl & "='" & id & "'" & stat & dateselect
                    'saldo = saldo +
                    'Tbx2Dec(SPAS.Lv_Journal_List.Items(SPAS.Lv_Journal_List.Items(i).Index).SubItems(3).Text) +
                    'Tbx2Dec(SPAS.Lv_Journal_List.Items(SPAS.Lv_Journal_List.Items(i).Index).SubItems(4).Text)
                End If
                '
                ' Dgv_Journal_Intern.DataSource = boundSet.Tables(0)
            End With
        Next
        SQLStr &= SQL_Where & " ORDER BY j.date, j.name"
        SPAS.ToClipboard(SQLStr, True)

        Return SQLStr

    End Function

    Sub Fill_Journal_List()
        Dim jrnl As Boolean
        jrnl = SPAS.Cmx_Journal_List.Text = "Journaalnaam"
        Dim cred, deb As Decimal

        Load_Datagridview(SPAS.Dgv_Journal_items, Create_Journal_SQL, "Lv_Journal_List_Click")
        'If jrnl And SPAS.Dgv_Journal_items.RowCount > 0 Then

        Try
            SPAS.Tbx_.Text = SPAS.Dgv_Journal_items.Rows(0).Cells(1).Value
            SPAS.Tbx_Journal_Descr.Text = SPAS.Dgv_Journal_items.Rows(0).Cells(4).Value

        Catch ex As Exception
            'MsgBox(ex.ToString)
        End Try


        'End If
        With SPAS.Dgv_Journal_items

            .Columns(0).Width = 60
            .Columns(0).HeaderText = "Dat"
            .Columns(0).DefaultCellStyle.Format = "dd-MM"

            .Columns(1).Width = 160
            .Columns(1).HeaderText = "Naam"

            .Columns(2).Width = 70
            .Columns(3).Width = 70
            .Columns(2).HeaderText = "Bij"
            .Columns(3).HeaderText = "Af"
            .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            .Columns(2).DefaultCellStyle.Format = "N2"
            .Columns(3).DefaultCellStyle.Format = "N2"
            .Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

            .Columns(4).Width = IIf(jrnl, 0, 250)
            .Columns(4).HeaderText = "Omschrijving"


            .Columns(5).Visible = True
            .Columns(5).Width = 250
            .Columns(5).HeaderText = "Account"

            .Columns(6).Width = 80
            .Columns(6).HeaderText = "Status"

            .Columns(7).Width = 80
            .Columns(7).HeaderText = "Bron"

            .Columns(8).Width = 70
            .Columns(8).HeaderText = "IBAN"

            .Columns(9).Width = 80
            .Columns(9).HeaderText = "Soort"

            'determine visibility 


            .Columns(1).Visible = Not jrnl
            .Columns(8).Visible = Not jrnl
            .Columns(6).Visible = True 'jrnl
            '.Columns(2).Visible = Not jrnl
        End With
        cred = 0
        deb = 0
        Try
            For r = 0 To SPAS.Dgv_Journal_items.RowCount - 1
                cred += SPAS.Dgv_Journal_items.Rows(r).Cells(2).Value
                deb += SPAS.Dgv_Journal_items.Rows(r).Cells(3).Value
            Next
            SPAS.Tbx_Journal_Credit.Text = cred.ToString("#0.00")
            SPAS.Tbx_Journal_Debit.Text = deb
            SPAS.Tbx_Journal_Saldo.Text = cred - deb
        Catch ex As Exception

        End Try
        'Calculate_Journal_Totals()
        'Calculate_Journal_Overview()
    End Sub

    Sub Calculate_Budget(ByVal id As String)

        Dim SQLStr As String = ""
        Dim sd As String
        Dim mon As String = ""
        Dim where As String = ""
        If id <> "" Then where = "WHERE ac1.id=" & id


        For m = 1 To 12
            sd = Date.Today.Year & "-" & m & "-01"            'ed = CDate("01-" & m + 1 & "-" & Date.Now.Year).AddDays(-1)
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
                    SELECT  sum(co.donation/co.term)
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
    Sub Get_Settings_Data()
        Collect_data("SELECT * FROM settings where label ilike '%kind%' or label ilike '%oudere%' order by label")
        SPAS.Tbx_Settings_Banktext_Kind.Text = dst.Tables(0).Rows(0)(1)
        SPAS.Tbx_Settings_Banktext_Oudere.Text = dst.Tables(0).Rows(1)(1)
        SPAS.Tbx_Settings_Bedrag_Kind.Text = dst.Tables(0).Rows(2)(1)
        SPAS.Tbx_Settings_Bedrag_Oudere.Text = dst.Tables(0).Rows(3)(1)
        SPAS.Tbx_Settings_Overhead_Kind.Text = dst.Tables(0).Rows(4)(1)
        SPAS.Tbx_Settings_Overhead_Oudere.Text = dst.Tables(0).Rows(5)(1)


    End Sub

    Sub Load_Account_Settings()

        Load_Combobox(SPAS.Cmx_Settings_Overhead, "id", "name", "SELECT id, name FROM accgroup WHERE active=TRUE ORDER BY name")
        Load_Combobox(SPAS.Cmx_Settings_No_Cat, "id", "name", "SELECT id, name FROM accgroup WHERE active=TRUE ORDER BY name")
        Load_Combobox(SPAS.Cmx_Settings_Euro_Account, "id", "name", "SELECT id, name FROM accgroup WHERE active=TRUE ORDER BY name")
        Load_Combobox(SPAS.Cmx_Settings_ExchangeRate, "id", "name", "SELECT id, name FROM accgroup WHERE active=TRUE ORDER BY name")
        Load_Combobox(SPAS.Cmx_Settings_Bankkosten, "id", "name", "SELECT id, name FROM accgroup WHERE active=TRUE  ORDER BY name")
        Load_Combobox(SPAS.Cmx_Settings_Banktransactiekosten, "id", "name", "SELECT id, name FROM accgroup WHERE active=TRUE ORDER BY name")
        Load_Combobox(SPAS.Cmx_Settings_Transitoria, "id", "name", "SELECT id, name FROM accgroup WHERE active=TRUE ORDER BY name")
        Load_Combobox(SPAS.Cmx_Settings_Saldosteun, "id", "name", "SELECT id, name FROM accgroup WHERE active=TRUE ORDER BY name")

        Collect_data("
                    select s.value, s.label, ag.name from settings s
                    left join accgroup ag on ag.id = s.value::integer
                    where s.value ~ '^-?\d*\.?\d+$'
                    group by s.value, s.label,ag.name
                    having ag.name is distinct from null
                    order by  ag.name
        ")

        For i = 0 To dst.Tables(0).Rows.Count - 1
            Select Case dst.Tables(0).Rows(i)(1)
                Case "overhead" : SPAS.Cmx_Settings_Overhead.Text = dst.Tables(0).Rows(i)(2)
                Case "eurotegenwaarde" : SPAS.Cmx_Settings_Euro_Account.Text = dst.Tables(0).Rows(i)(2)
                Case "nocat" : SPAS.Cmx_Settings_No_Cat.Text = dst.Tables(0).Rows(i)(2)
                Case "bank_transactie_kosten" : SPAS.Cmx_Settings_Banktransactiekosten.Text = dst.Tables(0).Rows(i)(2)
                Case "wisselkoersverschil" : SPAS.Cmx_Settings_ExchangeRate.Text = dst.Tables(0).Rows(i)(2)
                Case "bank_kosten" : SPAS.Cmx_Settings_Bankkosten.Text = dst.Tables(0).Rows(i)(2)
                Case "transitoria" : SPAS.Cmx_Settings_Transitoria.Text = dst.Tables(0).Rows(i)(2)
                Case "saldosteun" : SPAS.Cmx_Settings_Saldosteun.Text = dst.Tables(0).Rows(i)(2)

            End Select


        Next i

    End Sub
    Sub Save_Settings()
        Dim sqlstr As String
        sqlstr = "
        Update public.settings SET value = " & Tbx2Int(SPAS.Tbx_Settings_Overhead_Oudere.Text) & " WHERE label = 'standaard_overhead_oudere';
        Update public.settings SET value = " & Tbx2Int(SPAS.Tbx_Settings_Overhead_Kind.Text) & " WHERE label = 'standaard_overhead_kind';
        Update public.settings SET value = " & Tbx2Int(SPAS.Tbx_Settings_Bedrag_Oudere.Text) & " WHERE label = 'standaard_bedrag_oudere';
        Update public.settings SET value = " & Tbx2Int(SPAS.Tbx_Settings_Bedrag_Kind.Text) & " WHERE label = 'standaard_bedrag_kind';
        Update public.settings SET value = '" & SPAS.Tbx_Settings_Banktext_Kind.Text & "' WHERE label = 'bank_kind';
        Update public.settings SET value = '" & SPAS.Tbx_Settings_Banktext_Oudere.Text & "' WHERE label = 'bank_oudere';

        Update public.settings SET value = '" & SPAS.Cmx_Settings_Overhead.SelectedValue & "' WHERE label = 'overhead';
        Update public.settings SET value = '" & SPAS.Cmx_Settings_No_Cat.SelectedValue & "' WHERE label = 'nocat';
        Update public.settings SET value = '" & SPAS.Cmx_Settings_Euro_Account.SelectedValue & "' WHERE label = 'eurotegenwaarde';
        Update public.settings SET value = '" & SPAS.Cmx_Settings_Bankkosten.SelectedValue & "' WHERE label = 'bank_kosten';
        Update public.settings SET value = '" & SPAS.Cmx_Settings_Banktransactiekosten.SelectedValue & "' WHERE label = 'bank_transactie_kosten';
        Update public.settings SET value = '" & SPAS.Cmx_Settings_ExchangeRate.SelectedValue & "' WHERE label = 'wisselkoersverschil';
        Update public.settings SET value = '" & SPAS.Cmx_Settings_Transitoria.SelectedValue & "' WHERE label = 'transitoria';
        Update public.settings SET value = '" & SPAS.Cmx_Settings_Saldosteun.SelectedValue & "' WHERE label = 'saldosteun';
"
        RunSQL(sqlstr, "NULL", "Btn_Settings_Save")
        Clipboard.Clear()
        Clipboard.SetText(sqlstr)

    End Sub

End Module
