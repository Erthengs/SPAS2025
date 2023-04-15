Imports System.Windows.Forms.VisualStyles
Imports Microsoft.EntityFrameworkCore.Metadata

Module report
    Function Report_table(report_year)
        If CInt(report_year) >= CInt(QuerySQL("select min(extract (year from date)) from journal")) Then
            Return "journal"

        Else
            Return "journal_archive"
        End If

    End Function
    Function Bank_table(report_year)
        If CInt(report_year) >= CInt(QuerySQL("select min(extract (year from date)) from journal")) Then
            Return "bank"
        Else
            Return "bank_archive"
        End If
    End Function


    Sub Report_overview()
        report_year = SPAS.Cmx_Report_Year.Text

        Dim jtable, btable, yearcheck_j2 As String
        jtable = Report_table(report_year)
        btable = Bank_table(report_year)
        yearcheck_j2 = "and extract(year from date)=" & report_year
        Dim tabname As String
        For x = 0 To 1
            tabname = SPAS.TC_Rapportage.TabPages(x).Text
            If InStr(1, tabname, " (") > 0 Then SPAS.TC_Rapportage.TabPages(x).Text = Trim(Strings.Left(tabname, InStr(1, tabname, " (")))
            SPAS.TC_Rapportage.TabPages(x).Text &= " (" & report_year & ")"


        Next x



        Dim sqlstr As String = ""
        Dim select_j2 = "select sum(amt1) from " & jtable & " j2 left join account a2 on a2.id = j2.fk_account where a2.fk_accgroup_id = a.fk_accgroup_id and j2.status != 'Open' "
        Dim select_j = "select sum(amt1) from " & jtable & " j left join account a on a.id = j.fk_account where extract(year from j.date)=" & report_year & " and status != 'Open' and a.fk_accgroup_id in (select id from accgroup where type = '"
        Dim select_sum = "(select sum(amt1) from " & jtable & " j left join account a2 on a2.id = j.fk_account where extract(year from j.date)=" & report_year
        Dim un As String = ""
        Dim gtype() As String = {"Inkomsten", "Uitgaven", "Transit"}
        Dim rl As Integer = 0

        For i = 0 To 2

            un = IIf(i > 0, "union ", "")
            sqlstr &= un &
        "	select " & rl & ", '" & gtype(i) & "', null, null, null, null, null, null
            union select " & rl + 1 & ", ag.name
            ,(" & select_j2 & yearcheck_j2 & " and j2.source = 'Closing')
            ,(" & select_j2 & yearcheck_j2 & " and j2.source = 'Incasso')
            ,(" & select_j2 & yearcheck_j2 & " and j2.source = 'Bank')
            ,(" & select_j2 & yearcheck_j2 & " and j2.source = 'Intern')
            ,(" & select_j2 & yearcheck_j2 & " and j2.source = 'Uitkering')
            ,(" & select_j2 & yearcheck_j2 & ")
            from " & jtable & " j left join account a on a.id = fk_account left join accgroup ag on ag.id = a.fk_accgroup_id
            where status != 'Open' and ag.type = '" & gtype(i) & "' and extract(year from date)=" & CInt(report_year) & " group by ag.name, a.fk_accgroup_id
            --having (select sum(amt1) from " & jtable & " j2 left join account a2 on a2.id = j2.fk_account where a2.fk_accgroup_id = a.fk_accgroup_id) != '0.00'
            union select " & rl + 2 & ", 'Totaal' 
            ,(" & select_j & gtype(i) & "') and j.source = 'Closing')
            ,(" & select_j & gtype(i) & "') and j.source = 'Incasso')
            ,(" & select_j & gtype(i) & "') and j.source = 'Bank')
            ,(" & select_j & gtype(i) & "') and j.source = 'Intern')
            ,(" & select_j & gtype(i) & "') and j.source = 'Uitkering')
            ,(" & select_j & gtype(i) & "') and j.status != 'Open')
            union select " & rl + 3 & ", null, null, null, null, null, null, null "
            rl = rl + 4
        Next i

        sqlstr &= "
            --union select 12, null, null, null, null, null, null, null 
            union select 12, 'Totaal generaal' 
            ," & select_sum & " and j.source = 'Closing' and status != 'Open')
            ," & select_sum & " and j.source = 'Incasso' and status != 'Open')
            ," & select_sum & " and j.source = 'Bank' and status != 'Open')
            ," & select_sum & " and j.source = 'Intern' and status != 'Open')
            ," & select_sum & " and j.source = 'Uitkering' and status != 'Open')
            ," & select_sum & " and j.status != 'Open')

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
                .Columns(1).Width = 230

                For k = 2 To 7
                    .Columns(k).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    .Columns(k).Width = 120
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
                            .Rows(r).DefaultCellStyle.BackColor = IIf(btable = "bank_archive", Color.Silver, Color.Yellow)

                        End If
                        If spc = 0 Or spc = 4 Or spc = 8 Then
                            .Rows(r).DefaultCellStyle.BackColor = IIf(btable = "bank_archive", Color.Gainsboro, Color.PaleGreen)
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
                             from " & Report_table(report_year) & " j left join account a on a.id = j.fk_account  left join accgroup ag on ag.id = a.fk_accgroup_id
                             where extract(year from j.date)=" & report_year & "and j.source='" & source & "' and ag.name='" & accgroup & "' and j.status != 'Open' order by j.date desc"

        Clipboard.Clear()
        Clipboard.SetText(sql)

        Load_Datagridview(SPAS.Dgv_Report_6, sql, "boekingen")
        Format_drill_down()
        SPAS.TC_Rapportage.SelectedTab = SPAS.TC_Boekingen

    End Sub



    Sub Report_Bank_overview()

        Dim yearcheck = " and b2.iban = ba.accountno and extract(year from date)=" & report_year
        Dim sql As String = "
        select 0, 'BANKREKENINGEN', null, null, null,null, null,null,null   
        union select 1, b.iban,ba.name,
        (select credit-debit from " & Bank_table(report_year) & " b2 where name = '_startsaldo_'" & yearcheck & "),  
        (select sum(credit) from " & Bank_table(report_year) & " b2 where name != '_startsaldo_'" & yearcheck & "),
        (select sum(debit) from " & Bank_table(report_year) & " b2 where name != '_startsaldo_'" & yearcheck & "),
        (select sum(credit)-sum(debit) from " & Bank_table(report_year) & " b2 where b2.iban = ba.accountno and extract(year from date)=" & report_year & "),
        (select sum(credit)-sum(debit) from " & Bank_table(report_year) & " b2 where b2.iban = ba.accountno and extract(year from date)=" & report_year & "and iban2 in (select accountno from bankacc)),
        (select sum(credit)-sum(debit) from " & Bank_table(report_year) & " b2 where name != '_startsaldo_'" & yearcheck & ")
        from " & Bank_table(report_year) & " b left join bankacc ba on b.iban=ba.accountno 
        union 
        select 2,'Banksaldi totalen', null,  
        (select sum(credit)-sum(debit) from " & Bank_table(report_year) & " b2 where name = '_startsaldo_' and extract(year from date)=" & report_year & "),  
        (select sum(credit) from " & Bank_table(report_year) & " b2 where name != '_startsaldo_' and extract(year from date)=" & report_year & "),  
        (select sum(debit) from " & Bank_table(report_year) & " b2 where name != '_startsaldo_' and extract(year from date)=" & report_year & "),
        (select sum(credit)-sum(debit) from " & Bank_table(report_year) & " b2  where extract(year from date)=" & report_year & "),
        (select sum(credit)-sum(debit) from " & Bank_table(report_year) & " b2 where iban2 in (select accountno from bankacc) and extract(year from date)=" & report_year & "),
        (select sum(credit)-sum(debit) from " & Bank_table(report_year) & " b2 where name != '_startsaldo_' and extract(year from date)=" & report_year & ")
        from " & Bank_table(report_year) & " b 


        --union select 20,  null, NULL, null,null, NULL, null                     
        --union select 30, 'BANK/INTERNE BOEKINGEN', '', NULL, null,null, NULL 
        --union select 30, 'bron', 'doel', null,null,null,null 
        --union select 32, b.iban, b.iban2,null, sum(debit), sum(credit), sum(debit) - sum(credit) from bank b 
        --where iban2 in (select accountno from bankacc where expense = True)
        --and iban in (select accountno from bankacc where expense = False)
        --group by iban, iban2
        --union select 39,'Totaal', null, null, sum(debit), sum(credit), sum(debit) - sum(credit) as totaal from bank b 
        --where iban2 in (select accountno from bankacc where expense = True) and iban in (select accountno from bankacc where expense = False)
"
        SPAS.ToClipboard(sql, True)

        Load_Datagridview(SPAS.Dgv_Rapportage, sql, "rapportagefout Report_Bank_overview")


        For r = 0 To SPAS.Dgv_Rapportage.Rows.Count - 1
            Select Case r
                Case 0, 6 ', 30, 39
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
                .Columns(7).HeaderText = "Interne overboeking"
                .Columns(8).HeaderText = "Mutatie"
                .Columns(1).Width = 200
                .Columns(2).Width = 200

                For k = 3 To 8
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

        Dim saldo_date_start As String = "'Startsaldo 1-1-" & report_year + 1 & "' as name,'" & report_year + 1 & "-01-01' as date,'Verwerkt',"
        Dim tables As String = "journal j left join account a on a.id=j.fk_account left join accgroup g on g.id= a.fk_accgroup_id "
        Dim sql2 As String = ""
        Dim sql3 = "INSERT INTO journal(name, date, status, amt1, description, source, fk_account, type) VALUES" & vbCrLf
        Dim sql1 As String
        sql1 = "
    drop table if exists postings;
    select a.type as acnttype, " & saldo_date_start & " 
    'Startsaldo '||a.name as Description, 'Closing' as source, fk_account, j.type as jrntype, sum(amt1) as amt1, null As Verrekening, null As Eindtotaal 
    into temp table postings
    from " & tables & "
    where a.type = 'Specifiek (doel)' and extract (year from date)=2023 --g.type = 'Inkomsten' and j.status != 'Open' and
    group by a.type, a.name, j.fk_account, j.type
    having sum(amt1) !=0::money
    --order by a.type
    ------------------------------------------------------------------------------------------------------------
    -- 2 haal de totaal van de kosten op

    union ------------------uitgaven
    select 'Uitgaven'," & saldo_date_start & "
    'Totaal uitgaven', 'Closing',null, null,
    (select sum(amt1) from  " & tables & " where extract (year from date)=2023 and g.type = 'Uitgaven' and j.status != 'Open'), 
    null As Verrekening, null As Eindtotaal 

    union ----------------------------overhead
    select 'Generiek (overhead)', 'Startsaldo 1-1-2024', '2024-01-01','Verwerkt',
    'Overhead', 'Closing',null, null,
     (select sum(amt1) from  " & tables & " where a.type = 'Anders' and g.type != 'Uitgaven' and g.name = 'Overhead'),
     null As Verrekening, null As Eindtotaal 


    union----- fondsen--------------------------
    select a.type as acnttype, " & saldo_date_start & " 
    'Startsaldo '||a.name as Description, 'Closing' as source, fk_account, null,sum(amt1) as amt1, null As Verrekening, null As Eindtotaal 
    from " & tables & "
    where  extract (year from date)=2023 and a.type = 'Generiek (fonds)' and j.status != 'Open' and g.type = 'Inkomsten'
    group by a.type, a.name, j.fk_account
    having sum(amt1) !='0'::money

    union ----------------------------transit

    select 'Transit', " & saldo_date_start & "
    'Startsaldo '||a.name as Description, 'Closing' as source, fk_account, null,sum(amt1) as amt1, null As Verrekening, null As Eindtotaal 
    from " & tables & "
    where extract (year from date)=2023  and j.status != 'Open' and g.type = 'Transit'
    group by a.type, a.name, j.fk_account
    having sum(amt1) !='0'::money
    order by acnttype;


    select * from postings
"

        'RunSQL(sql1, "NULL", "Report_Closing")
        SPAS.ToClipboard(sql1, True)
        Load_Datagridview(SPAS.Dgv_Report_Year_Closing, sql1, "Report_Closing")
        For r = 0 To SPAS.Dgv_Report_Year_Closing.Rows.Count - 1
            If SPAS.Dgv_Report_Year_Closing.Rows(r).Cells(1).Value = "Transitoria" Then
            End If

        Next r




        Dim amt = 0.00
        Dim af As Integer = 0
        For r = 0 To SPAS.Dgv_Report_Year_Closing.Rows.Count - 1
            'amt = Replace(SPAS.Dgv_Report_Year_Closing.Rows(r).Cells(3).Value, ".", "")
            'amt = Replace(amt, ",", ".")
            If InStr(SPAS.Dgv_Report_Year_Closing.Rows(r).Cells(5).Value, "Startsaldo Algemeen,Fonds") Then af = r

            If InStr(SPAS.Dgv_Report_Year_Closing.Rows(r).Cells(0).Value, "v") > 0 Then 'oVerhead, uitgaVen
                SPAS.Dgv_Report_Year_Closing.Rows(r).Cells(9).Value = Format(SPAS.Dgv_Report_Year_Closing.Rows(r).Cells(8).Value, "N2")
                SPAS.Dgv_Report_Year_Closing.Rows(r).Cells(10).Value = 0
                amt += SPAS.Dgv_Report_Year_Closing.Rows(r).Cells(9).Value
            Else
                SPAS.Dgv_Report_Year_Closing.Rows(r).Cells(9).Value = 0
                SPAS.Dgv_Report_Year_Closing.Rows(r).Cells(10).Value = Format(SPAS.Dgv_Report_Year_Closing.Rows(r).Cells(8).Value, "N2")
            End If
        Next
        SPAS.Dgv_Report_Year_Closing.Rows(af).Cells(9).Value = amt
        SPAS.Dgv_Report_Year_Closing.Rows(af).Cells(10).Value = Format(SPAS.Dgv_Report_Year_Closing.Rows(af).Cells(10).Value + amt - 30000, "N2")
        SPAS.Lbl_Report_total.Text = SPAS.Dgv_Report_Year_Closing.Rows(af).Cells(10).Value


        Try
            With SPAS.Dgv_Report_Year_Closing



                .Columns(0).HeaderText = "Type"
                .Columns(0).Width = 135
                .Columns(1).Visible = False
                .Columns(2).Visible = False
                .Columns(3).Visible = False
                .Columns(4).HeaderText = "Omschrijving"
                .Columns(4).Width = 300
                .Columns(5).Visible = False
                .Columns(6).HeaderText = "Accnt"
                .Columns(7).HeaderText = "Journaltype"
                .Columns(8).HeaderText = "Bedrag"
                .Columns(9).HeaderText = "Verrekenen"
                .Columns(10).HeaderText = "Overdracht"
                .Columns(10).DefaultCellStyle.ForeColor = Color.DarkGreen
                .Columns(10).DefaultCellStyle.Format = "N2"

                For k = 0 To 10
                    .Columns(k).ReadOnly = True
                    If k > 7 Then
                        .Columns(k).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        .Columns(k).DefaultCellStyle.Format = "N2"
                    End If
                Next

            End With
        Catch ex As Exception

        End Try
        If SPAS.Dgv_Report_Year_Closing.Rows(af).Cells(10).Value < 0 Then
            SPAS.Dgv_Report_Year_Closing.Rows(af).DefaultCellStyle.ForeColor = Color.Red
            MsgBox("Het algemeen fonds bevat onvoldoende middelen om te overhead en uitgaven te verdisconteren, vul deze s.v.p. aan.")

        End If


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
