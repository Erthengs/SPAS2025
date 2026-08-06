Imports System.Text.RegularExpressions
Imports System.Windows.Forms.VisualStyles
Imports System.Xml
Imports Microsoft.EntityFrameworkCore.Metadata

Module report
    Sub Run_ReportTree(ByVal rep As String)
        Dim arr_format() As String = Nothing
        Dim sql As String = ""
        Dim formatting As String


        sql = QuerySQL($"Select sql from query where category ilike 'Overzicht%' and name='{rep}'")
        If IsNothing(sql) Then Exit Sub

        formatting = QuerySQL($"Select formatting from query where category ilike 'Overzicht%' and name='{rep}'")
        SPAS.LbL_Formatting.Text = formatting
        If Not IsNothing(SPAS.LbL_Formatting.Text) Then arr_format = SPAS.LbL_Formatting.Text.Split(","c)

        sql = Replace(sql, "[year]", report_year)
        If SPAS.Cmbx_Reporting_Year.SelectedIndex > 0 Then
            sql = sql.Replace("from bank ", "from bank_archive ")
            sql = sql.Replace("from journal ", "from journal_archive ")
        End If
        'Load_Datagridview(SPAS.Dgv_Rapportage_Overzicht, sql, "ReportTree.NodeMouseClick-level2")
        '
        SPAS.Prepare_Datagridview(SPAS.Dgv_Rapportage_Overzicht, sql, arr_format)
        Call SPAS.ApplyFilter(SPAS.Dgv_Rapportage_Overzicht.DataSource)
    End Sub


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


    Sub Drill_down_Report_overview(ByVal i As Integer, ByVal j As Integer)


        Dim source As String = ""
        Dim accgroup As String

        Select Case j
            Case 2 : source = "Closing"
            Case 3 : source = "Incasso"
            Case 4 : source = "Bank"
            Case 5 : source = "Intern"
            Case 6 : source = "Uitkering"
            Case Else
                Exit Sub
        End Select
        Dim bedrag As Integer = SPAS.Dgv_Rapportage_Overzicht.CurrentCell.Value

        accgroup = SPAS.Dgv_Rapportage_Overzicht.Rows(i).Cells(1).Value

        Dim sql As String = "
                select j.date As Datum, a.name Account,j.amt1 As Bedrag,j.name As Journaalnaam, j.type As Journaaltype, j.description As Omschrijving, j.iban As Iban,  ag.name as Accountgroep,  j.fk_bank, j.id 
                from " & Report_table(report_year) & " j left join account a on a.id = j.fk_account  left join accgroup ag on ag.id = a.fk_accgroup_id
                where extract(year from j.date)=" & report_year & "and j.source='" & source & "' and ag.name='" & accgroup & "' and j.status != 'Open' order by j.date desc;
"
        'Load_Datagridview(SPAS.Dgv_Report_6, sql, "boekingen")
        SPAS.Prepare_Datagridview(SPAS.Dgv_Report_6, sql, {"DZ080", "TZ140", "NG070", "TZ220", "HZ070", "TZ150", "HZ500", "TZ200", "HZ080", "HZ080"})
    End Sub

    Sub Drill_down_Bank_overview(ByVal i As Integer, ByVal j As Integer)


        Dim sqlpart1 As String = ""

        Select Case j
            Case 1 : sqlpart1 = " and b.name ilike '%startsaldo%'"
            Case 2 : sqlpart1 = " and b.credit >0::money"
            Case 3 : sqlpart1 = " and b.debit >0::money"
            Case 4
            Case 5
            Case Else
                Exit Sub
        End Select

        Dim bedrag As Integer = SPAS.Dgv_Rapportage_Overzicht.CurrentCell.Value
        Dim sql As String = "select b.date As Datum, b.seqorder As Afschrift, b.name As Naam, b.credit As Bij, b.debit As Af, b.code, b.description As omschrijving, fk_journal_name
                             from " & Bank_table(report_year) & " b 
                             where iban='" & Trim(SPAS.Dgv_Rapportage_Overzicht.Rows(i).Cells(0).Value) & "' and extract(year from b.date)= " & report_year & sqlpart1 &
                             " order by b.seqorder desc"

        ToClipboard(sql, True)

        Load_Datagridview(SPAS.Dgv_Report_6, sql, "drilldown banktransacties")
        SPAS.Prepare_Datagridview(SPAS.Dgv_Report_6, sql, {"DZ080", "TZ050", "TZ140", "NZ070", "NZ070", "TG040", "TZ500", "HZ050"})


    End Sub

    Sub Report_Closing()


        Dim Sqlc = QuerySQL("Select sql from query where category = 'Overzicht' and name='Transitieposten'")
        If IsNothing(Sqlc) Then Exit Sub
        Sqlc = Sqlc.Replace("[year]", report_year)

        RunSQL(Sqlc, "NULL", "Report Closing")

        Dim formatting As String = QuerySQL("select formatting from query where name='Transitieposten'")
        Dim arr_format() As String = Nothing
        If Not IsNothing(formatting) Then arr_format = formatting.Split(",")

        'Load_Datagridview(SPAS.Dgv_Report_Year_Closing, Sqlc, "...")
        SPAS.Prepare_Datagridview(SPAS.Dgv_Report_Year_Closing, Sqlc, arr_format)

    End Sub

End Module
