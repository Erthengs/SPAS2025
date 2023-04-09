Imports System.Windows.Forms.VisualStyles

Module report
    Sub Report_overview()

        Dim sqlstr As String = ""
        Dim un As String = ""
        Dim gtype() As String = {"Inkomsten", "Uitgaven", "Transit"}
        Dim rl As Integer = 0

        For i = 0 To 2

            un = IIf(i > 0, "union ", "")
            sqlstr &= un &
        "	select " & rl & ", '" & gtype(i) & "', null, null, null, null, null, null, null 
            union select " & rl + 1 & ", ag.name
            ,(select sum(amt1) from journal j2 left join account a2 on a2.id = j2.fk_account where a2.fk_accgroup_id = a.fk_accgroup_id and j2.status != 'Open' and j2.source = 'Closing')
            ,(select sum(amt1) from journal j2 left join account a2 on a2.id = j2.fk_account where a2.fk_accgroup_id = a.fk_accgroup_id and j2.status != 'Open' and j2.source = 'Incasso')
            ,(select sum(amt1) from journal j2 left join account a2 on a2.id = j2.fk_account where a2.fk_accgroup_id = a.fk_accgroup_id and j2.status != 'Open' and j2.source = 'Bank')
            ,(select sum(amt1) from journal j2 left join account a2 on a2.id = j2.fk_account where a2.fk_accgroup_id = a.fk_accgroup_id and j2.status != 'Open' and j2.source = 'Intern')
            ,(select sum(amt1) from journal j2 left join account a2 on a2.id = j2.fk_account where a2.fk_accgroup_id = a.fk_accgroup_id and j2.status != 'Open' and j2.source = 'Uitkering')
            ,(select sum(amt1) from journal j2 left join account a2 on a2.id = j2.fk_account where a2.fk_accgroup_id = a.fk_accgroup_id and j2.status != 'Open'),null
            from journal j left join account a on a.id = fk_account left join accgroup ag on ag.id = a.fk_accgroup_id
            where status != 'Open' and ag.type = '" & gtype(i) & "' group by ag.name, a.fk_accgroup_id
            --having (select sum(amt1) from journal j2 left join account a2 on a2.id = j2.fk_account where a2.fk_accgroup_id = a.fk_accgroup_id) != '0.00'
            union select " & rl + 2 & ", 'Totaal' 
            ,(select sum(amt1) from journal j left join account a on a.id = j.fk_account where status != 'Open' and a.fk_accgroup_id in (select id from accgroup where type = '" & gtype(i) & "') and j.source = 'Closing')
            ,(select sum(amt1) from journal j left join account a on a.id = j.fk_account where status != 'Open' and a.fk_accgroup_id in (select id from accgroup where type = '" & gtype(i) & "') and j.source = 'Incasso')
            ,(select sum(amt1) from journal j left join account a on a.id = j.fk_account where status != 'Open' and a.fk_accgroup_id in (select id from accgroup where type = '" & gtype(i) & "') and j.source = 'Bank')
            ,(select sum(amt1) from journal j left join account a on a.id = j.fk_account where status != 'Open' and a.fk_accgroup_id in (select id from accgroup where type = '" & gtype(i) & "') and j.source = 'Intern')
            ,(select sum(amt1) from journal j left join account a on a.id = j.fk_account where status != 'Open' and a.fk_accgroup_id in (select id from accgroup where type = '" & gtype(i) & "') and j.source = 'Uitkering')
            ,(select sum(amt1) from journal j left join account a on a.id = j.fk_account where status != 'Open' and a.fk_accgroup_id in (select id from accgroup where type = '" & gtype(i) & "') and j.status != 'Open'),null
            union select " & rl + 3 & ", null, null, null, null, null, null, null,null "
            rl = rl + 4
        Next i

        sqlstr &= "
            --union select 12, null, null, null, null, null, null, null,null 
            union select 12, 'Totaal generaal' 
            ,(select sum(amt1) from journal j left join account a2 on a2.id = j.fk_account where j.source = 'Closing' and status != 'Open')
            ,(select sum(amt1) from journal j left join account a2 on a2.id = j.fk_account where j.source = 'Incasso' and status != 'Open')
            ,(select sum(amt1) from journal j left join account a2 on a2.id = j.fk_account where j.source = 'Bank' and status != 'Open')
            ,(select sum(amt1) from journal j left join account a2 on a2.id = j.fk_account where j.source = 'Intern' and status != 'Open')
            ,(select sum(amt1) from journal j left join account a2 on a2.id = j.fk_account where j.source = 'Uitkering' and status != 'Open')
            ,(select sum(amt1) from journal j left join account a2 on a2.id = j.fk_account where j.status != 'Open')
            ,null  
        "
        Clipboard.Clear()
        Clipboard.SetText(sqlstr)


        Load_Datagridview(SPAS.Dgv_Rapportage_Overzicht, sqlstr, "rapportagefout report_1")
        'FORMATTING REPORT
        Try
            With SPAS.Dgv_Rapportage_Overzicht
                .Columns(0).Visible = False
                .Columns(1).HeaderText = "Categorie"
                .Columns(2).HeaderText = "Saldo 1/1"
                .Columns(3).HeaderText = "Automatische Incasso"
                .Columns(4).HeaderText = "Overige Banktrans."
                .Columns(5).HeaderText = "Interne boekingen"
                .Columns(6).HeaderText = "Uitkering"
                .Columns(7).HeaderText = "Saldo 31/12"
                .Columns(8).HeaderText = "Saldo 1/1 nieuw jaar"
                .Columns(1).Width = 230

                For k = 2 To 8
                    .Columns(k).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    .Columns(k).Width = 100
                    .Columns(k).DefaultCellStyle.Format = "N2"
                    .Columns(k).ReadOnly = True

                Next
                Dim spc As Integer
                For r = 0 To .Rows.Count - 1
                    spc = .Rows(r).Cells(0).Value
                    If spc = 0 Or spc = 2 Or spc = 4 Or spc = 6 Or spc = 8 Or spc = 10 Or spc = 12 Then
                        .Rows(r).DefaultCellStyle.ForeColor = Color.Blue
                        .Rows(r).DefaultCellStyle.Font = New Font("Calibri", 12, FontStyle.Bold)
                        If spc = 12 Then
                            .Rows(r).DefaultCellStyle.BackColor = Color.Yellow
                        End If
                        If spc = 0 Or spc = 4 Or spc = 8 Then
                            .Rows(r).DefaultCellStyle.BackColor = Color.PaleGreen
                        End If

                    End If
                Next r

            End With
        Catch ex As Exception

        End Try
    End Sub

    Sub Drill_down_Report_overview(ByVal i As Integer, ByVal j As Integer)

        Dim bedrag As Integer = SPAS.Dgv_Rapportage_Overzicht.CurrentCell.Value
        Dim source As String = ""
        Dim accgroup As String

        Select Case j
            Case 2 : source = "Closing"
            Case 3 : source = "Incasso"
            Case 4 : source = "Bank"
            Case 5 : source = "Intern"
            Case 6 : source = "Uitkering"
        End Select

        accgroup = SPAS.Dgv_Rapportage_Overzicht.Rows(i).Cells(1).Value
        'MsgBox(accgroup & " " & source & " " & bedrag)

        Dim sql As String = "select j.date, a.name,j.amt1,j.name, j.type, j.description, j.iban,  ag.name,  j.fk_bank, j.id 
                             from journal j left join account a on a.id = j.fk_account  left join accgroup ag on ag.id = a.fk_accgroup_id
                             where j.source='" & source & "' and ag.name='" & accgroup & "' and j.status != 'Open' order by j.date desc"

        Clipboard.Clear()
        Clipboard.SetText(sql)

        Load_Datagridview(SPAS.Dgv_Report_6, sql, "boekingen")
        Format_drill_down()
        SPAS.TC_Rapportage.SelectedTab = SPAS.TC_Boekingen

    End Sub



    Sub Report_Bank_overview()
        Dim sql As String = "
        select 0, 'BANKREKENINGEN', null, NULL, null,null, NULL      
        UNION  SELECT '1',(SELECT'Bankrekening '::text || bc.name AS name FROM bankacc bc  WHERE b.iban::text = bc.accountno) AS account,
        ( SELECT bc.accountno FROM bankacc bc WHERE b.iban::text = bc.accountno) AS account_detail,
        ( SELECT bc.startbalance FROM bankacc bc WHERE b.iban::text = bc.accountno) AS startsaldo,
        sum(b.credit) AS Bij,
        sum(b.debit) AS Af,
        sum(b.credit) - sum(b.debit) + (( SELECT bc.startbalance FROM bankacc bc WHERE b.iban::text = bc.accountno)) AS ActueelSaldo
        FROM bank b GROUP BY b.iban
        --union  select '18', null, NULL, null,null, NULL, null
        UNION select 19, 'Banksaldi totaal'::text AS account, null,
        (SELECT sum(bc.startbalance) AS sum FROM bankacc bc) AS startsaldo,
        sum(bank.credit) AS Bij,
        sum(bank.debit) AS Af,
        (( SELECT sum(bc.startbalance) AS sum FROM bankacc bc)) +
            CASE
                WHEN sum(bank.credit) IS NULL THEN 0::money ELSE sum(bank.credit)
            END -
            CASE
                WHEN sum(bank.debit) IS NULL THEN 0::money ELSE sum(bank.debit)
            END AS ActueelSaldo
        FROM bank
        union select 20,  null, NULL, null,null, NULL, null                     
        union select 30, 'BANK/INTERNE BOEKINGEN', '', NULL, null,null, NULL 
        union select 30, 'bron', 'doel', null,null,null,null 
        union select 32, b.iban, b.iban2,null, sum(debit), sum(credit), sum(debit) - sum(credit) from bank b 
        where iban2 in (select accountno from bankacc where expense = True)
        and iban in (select accountno from bankacc where expense = False)
        group by iban, iban2
        union select 39,'Totaal', null, null, sum(debit), sum(credit), sum(debit) - sum(credit) as totaal from bank b 
        where iban2 in (select accountno from bankacc where expense = True) and iban in (select accountno from bankacc where expense = False)
"

        Load_Datagridview(SPAS.Dgv_Rapportage, sql, "rapportagefout Report_Bank_overview")
        For r = 0 To SPAS.Dgv_Rapportage.Rows.Count - 1
            Select Case SPAS.Dgv_Rapportage.Rows(r).Cells(0).Value
                Case 0, 19, 30, 39
                    SPAS.Dgv_Rapportage.Rows(r).DefaultCellStyle.ForeColor = Color.Blue
                    SPAS.Dgv_Rapportage.Rows(r).DefaultCellStyle.Font = New Font("Calibri", 12, FontStyle.Bold)
                Case Else
            End Select
        Next

        Try
            With SPAS.Dgv_Rapportage

                .Columns(0).Visible = False
                .Columns(1).HeaderText = "Rekeningnaam"
                .Columns(2).HeaderText = "IBAN"
                .Columns(3).HeaderText = "Startsaldo"
                .Columns(4).HeaderText = "Bij"
                .Columns(5).HeaderText = "Af"
                .Columns(6).HeaderText = "Saldo"
                .Columns(1).Width = 270
                .Columns(2).Width = 200

                For k = 3 To 6
                    .Columns(k).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    .Columns(k).Width = 100
                    .Columns(k).DefaultCellStyle.Format = "N2"
                    .Columns(k).ReadOnly = False
                Next

            End With
        Catch
        End Try

    End Sub

    Sub Report_bank_analysis()

        Dim sql As String

        sql = "
                  select 0, 'DETAIL CP-REKENINGEN', null, null, null, null, null, null, null, null, null, null, null
                   UNION select 1, TRIM(j1.iban)
                  ,(select sum(amt1) from journal j2 left join account a on j2.fk_account = a.id left join accgroup ag1 on ag1.id = a.fk_accgroup_id where status!='Open'
                  			and j2.iban = j1.iban and amt1 > '$0.00' and j2.type != 'Internal' and 
                            ag1.id=(select value::integer from settings where label='eurotegenwaarde'))      
                  ,(select sum(amt1) from journal j2 left join account a on j2.fk_account = a.id left join accgroup ag1 on ag1.id = a.fk_accgroup_id where status!='Open'
                  			and j2.iban = j1.iban and amt1 > '$0.00' and ag1.id=(select value::integer from settings where label='transitoria'))
                  ,(select sum(amt1) from journal j2 left join account a on j2.fk_account = a.id left join accgroup ag1 on ag1.id = a.fk_accgroup_id where status!='Open'
                  			and j2.iban = j1.iban and j2.type = 'Internal' and amt1 > '$0.00' and ag1.id=(select value::integer from settings where label='eurotegenwaarde'))
                  ,(select sum(amt1) from journal j2 left join account a on j2.fk_account = a.id left join accgroup ag1 on ag1.id = a.fk_accgroup_id where status!='Open' 
                  			and j2.iban = j1.iban and amt1 > '$0.00' and ag1.id!=(select value::integer from settings where label='wisselkoersverschil'))
                  ,(select sum(amt1) from journal j2 left join account a on j2.fk_account = a.id left join accgroup ag1 on ag1.id = a.fk_accgroup_id where status!='Open' 
                            and j2.iban = j1.iban  and amt1 < '$0.00' and ag1.id=(select value::integer from settings where label='eurotegenwaarde') and j2.type != 'Internal')  as Uitkering
                  ,(select sum(amt1) from journal j2 left join account a on j2.fk_account = a.id left join accgroup ag1 on ag1.id = a.fk_accgroup_id where status!='Open' 
                            and j2.iban = j1.iban  and amt1 < '$0.00' and ag1.id=(select value::integer from settings where label='transitoria'))  as transit_uit
                  ,(select sum(amt1) from journal j2 left join account a on j2.fk_account = a.id left join accgroup ag1 on ag1.id = a.fk_accgroup_id where status!='Open' 
                            and j2.iban = j1.iban  and amt1 < '$0.00' and ag1.id=(select value::integer from settings where label='eurotegenwaarde') and j2.type = 'Internal')
                  ,(select sum(amt1) from journal j2 left join account a on j2.fk_account = a.id left join accgroup ag1 on ag1.id = a.fk_accgroup_id where status!='Open' 
                            and j2.iban = j1.iban  and amt1 < '$0.00' and ag1.id=(select value::integer from settings where label='bank_kosten'))
                  ,(select sum(amt1) from journal j2 left join account a on j2.fk_account = a.id left join accgroup ag1 on ag1.id = a.fk_accgroup_id where status!='Open' 
                            and j2.iban = j1.iban  and amt1 < '$0.00' and ag1.id=(select value::integer from settings where label='bank_transactie_kosten'))
                  ,(select case when (select sum(amt1) from journal j2 left join account a on j2.fk_account = a.id left join accgroup ag1 on ag1.id = a.fk_accgroup_id where status!='Open' 
                            and j2.iban = j1.iban and ag1.id=(select value::integer from settings where label='wisselkoersverschil')) is not distinct from null then '$0.00'
                            else (select sum(amt1) from journal j2 left join account a on j2.fk_account = a.id left join accgroup ag1 on ag1.id = a.fk_accgroup_id where status!='Open' 
                            and j2.iban = j1.iban and ag1.id=(select value::integer from settings where label='wisselkoersverschil')) end)

                   ,(select sum(amt1) from journal j2 left join account a on j2.fk_account = a.id left join accgroup ag1 on ag1.id = a.fk_accgroup_id where status!='Open' 
                            and j2.iban = j1.iban and amt1 < '$0.00') 
                   +(select case when (select sum(amt1) from journal j2 left join account a on j2.fk_account = a.id left join accgroup ag1 on ag1.id = a.fk_accgroup_id where status!='Open' 
                            and j2.iban = j1.iban and ag1.id=(select value::integer from settings where label='wisselkoersverschil')and amt1 > '$0.00') is not distinct from null then '$0.00'
                            else (select sum(amt1) from journal j2 left join account a on j2.fk_account = a.id left join accgroup ag1 on ag1.id = a.fk_accgroup_id where status!='Open' 
                            and j2.iban = j1.iban and ag1.id=(select value::integer from settings where label='wisselkoersverschil')and amt1 > '$0.00') end) 

                  from journal j1 left join bankacc ba on ba.accountno = j1.iban
                  where ba.expense = true  group by j1.iban
                  
                  union
                  select 2, 'Totaal'
                  ,(select sum(amt1) from journal j 
                            where amt1 > '$0.00' and status != 'Open' and type != 'Internal' and 
                            fk_account =443 ) 
                  ,(select sum(amt1) from journal j 
                            where j.iban in (select accountno from bankacc where expense=true) and status != 'Open'and amt1 > '$0.00' and 
                            fk_account =440)
                  ,(select sum(amt1) from journal j 
                            where amt1 > '$0.00' and status != 'Open'and type = 'Internal' and 
                            fk_account =443)  
                  ,(select sum(amt1) from journal j 
                            where iban in (select accountno from bankacc where expense=true) and status != 'Open'and amt1 > '$0.00' and 
                            fk_account not in (448)) 
                  ,(select sum(amt1) from journal j 
                            where amt1 < '$0.00' and status != 'Open' and type != 'Internal' and 
                            fk_account =443) 
                  ,(select sum(amt1) from journal j 
                            where j.iban in (select accountno from bankacc where expense=true) and amt1 < '$0.00' and 
                            fk_account =440)
                  ,(select sum(amt1) from journal j 
                            where amt1 < '$0.00' and status != 'Open'and type = 'Internal' and 
                            fk_account =443)
                  ,(select sum(amt1) from journal j 
                            where j.iban in (select accountno from bankacc where expense=true) and status != 'Open'and amt1 < '$0.00' and 
                            fk_account =442)
                  ,(select sum(amt1) from journal j 
                            where j.iban in (select accountno from bankacc where expense=true) and amt1 < '$0.00' and 
                            fk_account =441) 
                  ,(select case when (select sum(amt1) from journal j where j.iban in (select accountno from bankacc where expense=true)  and 
                                    fk_account =448) is not distinct from null then '$0.00' 
                            else (select sum(amt1) from journal j where j.iban in 
                                (select accountno from bankacc where expense=true)and status != 'Open'  
                                 and fk_account =448) 
                            end)
                  ,(select sum(amt1) from journal j where iban in (select accountno from bankacc where expense=true)and status != 'Open' and amt1 > '$0.00' and 
                            fk_account not in (448))                 
        "

        Load_Datagridview(SPAS.Dgv_Rapportage_Details, sql, "report details")

        Try
            With SPAS.Dgv_Rapportage_Details
                '.Columns(0).HeaderText = "Nr"
                .Columns(1).HeaderText = "IBAN"
                .Columns(2).HeaderText = "Giften"
                .Columns(3).HeaderText = "Transit inkomend"
                .Columns(4).HeaderText = "Extra inkomend"
                .Columns(5).HeaderText = "Totaal inkomend"
                .Columns(6).HeaderText = "Kas opname"
                .Columns(7).HeaderText = "Transit uitgaand"
                .Columns(8).HeaderText = "Extra uitgaand"
                .Columns(9).HeaderText = "Bank kosten"
                .Columns(10).HeaderText = "Opname kosten"
                .Columns(11).HeaderText = "Koers verschil"
                .Columns(12).HeaderText = "Totaal uitgaand"


                .Columns(0).Visible = False
                .Columns(1).Width = 145
                '.Columns(2).Visible = False
                '.Columns(3).Width = 200

                For k = 2 To 12
                    .Columns(k).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    .Columns(k).Width = IIf(k = 5 Or k = 6 Or k = 12, 79, 76)
                    .Columns(k).DefaultCellStyle.Format = "N2"
                    .Columns(k).ReadOnly = True

                    Select Case k
                        Case < 5 : .Columns(k).DefaultCellStyle.ForeColor = Color.DarkOliveGreen
                        Case 5 : .Columns(k).DefaultCellStyle.ForeColor = Color.Green
                        Case 6, 7, 8, 9, 10 : .Columns(k).DefaultCellStyle.ForeColor = Color.DarkRed
                        Case 11 : .Columns(k).DefaultCellStyle.ForeColor = Color.Blue
                        Case 12
                            .Columns(k).DefaultCellStyle.ForeColor = Color.Red
                            .Columns(k).DefaultCellStyle.Font = New Font("Calibri", 12, FontStyle.Bold)
                    End Select
                Next
            End With

            For r = 0 To SPAS.Dgv_Rapportage_Details.Rows.Count
                If r = 0 Or r = 5 Then
                    If r = 0 Then SPAS.Dgv_Rapportage_Details.Rows(r).DefaultCellStyle.ForeColor = Color.Blue
                    SPAS.Dgv_Rapportage_Details.Rows(r).DefaultCellStyle.Font = New Font("Calibri", 12, FontStyle.Bold)
                End If
            Next

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub
    Sub Drill_down_Report_Bank_Analysis(ByVal i As Integer, ByVal j As Integer)
        Dim sql, sql_where As String
        Dim iban As String = SPAS.Dgv_Rapportage_Details.Rows(i).Cells(1).Value
        If iban = "Totaal" Then iban = "%"

        Select Case j
            Case 2 : sql_where = " AND amt1 > '$0.00' and fk_account =443 and j2.type != 'Internal' "
            Case 3 : sql_where = " AND amt1 > '$0.00' and fk_account =440"
            Case 4 : sql_where = " AND amt1 > '$0.00' and fk_account =443 and j2.type = 'Internal' "
            Case 5 : sql_where = "and amt1 > '$0.00' and fk_account not in (448,4401)"
            Case 2 : sql_where = " AND amt1 < '$0.00' and fk_account =443 and j2.type != 'Internal' "
            Case 3 : sql_where = " AND amt1 < '$0.00' and fk_account =440"
            Case 4 : sql_where = " AND amt1 < '$0.00' and fk_account =443 and j2.type = 'Internal' "
            Case 5 : sql_where = "and amt1 , '$0.00' and fk_account not in (448,4401)"
            Case 10 : sql_where = ""
            Case 11 : sql_where = ""
            Case 12 : sql_where = ""

            Case Else
                sql_where = ""
        End Select

        sql = "select j2.date, a2.name,j2.amt1,j2.name, j2.description, j2.iban,  a2.accgroup,  j2.fk_bank, j2.type from journal j2 left join account a2 on a2.id = j2.fk_account  
        where j2.iban='" & iban & "' and status != 'Open'" & sql_where

        Load_Datagridview(SPAS.Dgv_Report_6, sql, "boekingen")
        Format_drill_down()
        SPAS.TC_Rapportage.SelectedTab = SPAS.TC_Boekingen

    End Sub

    Sub Report_checks()




    End Sub
    Sub Report_Closing()

        Dim sql2 As String = ""
        Dim sql3 = "INSERT INTO journal(name, date, status, amt1, description, source, fk_account, type) VALUES" & vbCrLf
        Dim sql1 As String
        sql1 = "
drop table if exists postings;
select 'Startsaldo 1-1-2023' as name, '2023-01-01'as date, 'Verwerkt' as Status, sum(amt1) as amt1, 'Startsaldo '||a.name as Description, 'Opening' as source, fk_account, 'Internal' as type  
into temp table postings
from journal j
left join account a on a.id=j.fk_account
left join accgroup g on g.id= a.fk_accgroup_id 
where g.type != 'Uitgaven'
group by a.name, j.fk_account
having sum(amt1) !=0::money
union
select 'Startsaldo 1-1-2023' as name, '2023-01-01' as date, 'Verwerkt' as Status, sum(amt1) as amt1, 'Totaal uitgaven' as Description, 'Opening' as source, null as fk_account, 'Internal' as type  
from journal j
left join account a on a.id=j.fk_account
left join accgroup g on g.id= a.fk_accgroup_id 
where g.type = 'Uitgaven'
union 
select b.iban,  '2023-01-01', 'Bank', sum(b.credit-b.debit) + (select startbalance from bankacc a where a.accountno = b.iban), null, null, 99999, null
from bank b
group by iban;

update postings set amt1 = (select amt1 from postings where fk_account=439)::money + (select amt1 from postings where fk_account is not distinct from null)::money + '0.35'
where fk_account=439;
update postings set amt1 = (select amt1 from postings where fk_account=701)::money + (select amt1 from postings where fk_account =439)::money
where fk_account=701;
delete from postings  where fk_account=439;
delete from postings where fk_account is not distinct from null;


--select sum(amt1) from postings;
select * from postings
"

        'RunSQL(sql1, "NULL", "Report_Closing")

        Load_Datagridview(SPAS.Dgv_Report_Year_Closing, sql1, "Report_Closing")

        Dim amt As String
        For r = 0 To SPAS.Dgv_Report_Year_Closing.Rows.Count - 1
            amt = Replace(SPAS.Dgv_Report_Year_Closing.Rows(r).Cells(3).Value, ".", "")
            amt = Replace(amt, ",", ".")
            If SPAS.Dgv_Report_Year_Closing.Rows(r).Cells(2).Value = "Bank" Then
                sql2 = sql2 & "UPDATE bankacc SET startbalance='" & amt & "' WHERE accountno='" & SPAS.Dgv_Report_Year_Closing.Rows(r).Cells(0).Value & "';" & vbCrLf
            Else
                sql3 = sql3 & "('" & SPAS.Dgv_Report_Year_Closing.Rows(r).Cells(0).Value & "','" &
                SPAS.Dgv_Report_Year_Closing.Rows(r).Cells(1).Value & "','" &
                SPAS.Dgv_Report_Year_Closing.Rows(r).Cells(2).Value & "','" &
                amt & "','" &
                SPAS.Dgv_Report_Year_Closing.Rows(r).Cells(4).Value & "','" &
                SPAS.Dgv_Report_Year_Closing.Rows(r).Cells(5).Value & "','" &
                SPAS.Dgv_Report_Year_Closing.Rows(r).Cells(6).Value & "','" &
                SPAS.Dgv_Report_Year_Closing.Rows(r).Cells(7).Value & "')," & vbCrLf
            End If

        Next

        Clipboard.Clear()
        Clipboard.SetText(sql2 & Strings.Left(sql3, Strings.Len(sql3) - 3) & ";")
        MsgBox(sql2 & Strings.Left(sql3, Strings.Len(sql3) - 3) & ";" & vbCrLf & vbCrLf & " is gekopieerd naar het klembord.")

        Exit Sub

        SPAS.Lbl_Report_total.Text = QuerySQL("select sum(amt1) from journal 
            where status != 'Open' and extract(year from date) = (select extract(year from min(date)) from journal)")

        Try
            With SPAS.Dgv_Report_Year_Closing
                .Columns(0).HeaderText = "Omschrijving"
                .Columns(1).HeaderText = "Bedrag"
                .Columns(2).HeaderText = "Journaalnaam"
                .Columns(3).HeaderText = "Boekdatum"
                .Columns(4).HeaderText = "Status"
                .Columns(5).HeaderText = "Bron"
                .Columns(6).HeaderText = "Ac.nr"
                .Columns(7).HeaderText = "Ac.groep"
                .Columns(0).Width = 230
                .Columns(2).Visible = False

                For k = 1 To 6
                    .Columns(k).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    .Columns(k).Width = 90
                    .Columns(k).DefaultCellStyle.Format = "N2"
                    .Columns(k).ReadOnly = True
                Next

            End With
        Catch ex As Exception

        End Try
    End Sub

    Sub Report_collection2()
        Dim sql As String = "
        select 20, j.source,
        (select sum(amt1) from journal j2 left join account a on a.id = j2.fk_account left join accgroup ag on ag.id = a.fk_accgroup_id 
        where j2.amt1 > '$0.00' and j2.source = j.source  and j2.status != 'Open' and ag.id = (select value from settings where label='overhead')::integer),
        (select sum(amt1) from journal j2 left join account a on a.id = j2.fk_account left join accgroup ag on ag.id = a.fk_accgroup_id  
        where j2.amt1 < '$0.00' and j2.source = j.source  and j2.status != 'Open' and ag.id = (select value from settings where label='overhead')::integer),
        (select sum(amt1) from journal j2 left join account a on a.id = j2.fk_account left join accgroup ag on ag.id = a.fk_accgroup_id  
        where j2.source = j.source  and j2.status != 'Open' and ag.id = (select value from settings where label='overhead')::integer)
        --,'Toelichting'
        from journal j 
        where status != 'Open' group by j.source
        having  (select sum(amt1) from journal j2 left join account a on a.id = j2.fk_account left join accgroup ag on ag.id = a.fk_accgroup_id  
        where j2.source = j.source  and j2.status != 'Open' and ag.id = (select value from settings where label='overhead')::integer) != '$0.00'

        union select 21, 'Totaal',
        (select sum(amt1) from journal j2 left join account a on a.id = j2.fk_account left join accgroup ag on ag.id = a.fk_accgroup_id  
        where j2.amt1 > '$0.00' and j2.status != 'Open' and  ag.id = (select value from settings where label='overhead')::integer),
        (select sum(amt1) from journal j2 left join account a on a.id = j2.fk_account left join accgroup ag on ag.id = a.fk_accgroup_id  
        where j2.amt1 < '$0.00' and j2.status != 'Open' and ag.id = (select value from settings where label='overhead')::integer),
        (select sum(amt1) from journal j2 left join account a on a.id = j2.fk_account left join accgroup ag on ag.id = a.fk_accgroup_id  
        where j2.status != 'Open' and ag.id = (select value from settings where label='overhead')::integer)

    "
        Load_Datagridview(SPAS.Dgv_Report_7, sql, "report 7 overhead")
        Try
            With SPAS.Dgv_Report_7
                .Columns(0).Visible = False
                .Columns(1).HeaderText = "Bron"
                .Columns(2).HeaderText = "Bij"
                .Columns(3).HeaderText = "Af"
                .Columns(4).HeaderText = "Saldo"

                For k = 2 To 4
                    .Columns(k).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    .Columns(k).Width = 90
                    .Columns(k).DefaultCellStyle.Format = "N2"
                    .Columns(k).ReadOnly = True
                Next
                .Rows(3).DefaultCellStyle.ForeColor = Color.Blue
                .Rows(3).DefaultCellStyle.Font = New Font("Calibri", 12, FontStyle.Bold)
            End With
        Catch ex As Exception
        End Try


    End Sub
    Sub Drill_down_Overhead_Detail(ByVal i As Integer, ByVal j As Integer)
        Dim amt, source As String

        Select Case j
            Case 2
                amt = "and j2.amt1 > '$0.00'"
            Case 3
                amt = "and j2.amt1 < '$0.00'"
            Case 4
                amt = ""
        End Select
        source = SPAS.Dgv_Report_7.Rows(i).Cells(1).Value
        If source = "Totaal" Then source = "%"

        Dim sql As String = "
                select j2.date, a2.name,j2.amt1,j2.name, j2.description, j2.iban,  a2.accgroup,  j2.fk_bank from journal j2 left join account a2 on a2.id = j2.fk_account  
            	where j2.source = '" & source & "' and j2.status != 'Open' and j2.fk_account = 439" & amt

        Clipboard.SetText(sql)
        Load_Datagridview(SPAS.Dgv_Report_6, sql, "boekingen")
        Format_drill_down()
        SPAS.TC_Rapportage.SelectedTab = SPAS.TC_Boekingen
    End Sub
    Sub Format_drill_down()
        'select j.date, a.name,j.amt1,j.name, j.type, j.description, j.iban,  ag.name,  j.fk_bank 
        Try
            With SPAS.Dgv_Report_6

                .Columns(0).HeaderText = "Datum"
                .Columns(0).Width = 80
                .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                .Columns(1).HeaderText = "Account"
                .Columns(1).Width = 140
                .Columns(2).HeaderText = "Bedrag"
                .Columns(2).Width = 70
                .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                .Columns(2).DefaultCellStyle.Format = "N2"
                .Columns(2).ReadOnly = True
                .Columns(3).HeaderText = "Naam transactie"
                .Columns(3).Width = 150
                .Columns(4).HeaderText = "Type"
                .Columns(4).Visible = False
                .Columns(5).HeaderText = "Omschrijving transactie"
                .Columns(5).Width = 400
                .Columns(6).HeaderText = "IBAN"
                .Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                .Columns(6).Visible = False
                .Columns(7).HeaderText = "Accountgroep"
                .Columns(7).Width = 125
                .Columns(8).Visible = False
                .Columns(9).Visible = False

            End With
        Catch ex As Exception
        End Try
    End Sub


End Module
