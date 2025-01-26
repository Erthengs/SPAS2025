Imports System.Windows.Forms.VisualStyles
Imports System.Xml
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
        'MsgBox(accgroup & " " & source & " " & bedrag)

        Dim sql As String = "
                select j.date As Datum, a.name Account,j.amt1 As Bedrag,j.name As Journaalnaam, j.type As Journaaltype, j.description As Omschrijving, j.iban As Iban,  ag.name as Accountgroep,  j.fk_bank, j.id 
                from " & Report_table(report_year) & " j left join account a on a.id = j.fk_account  left join accgroup ag on ag.id = a.fk_accgroup_id
                where extract(year from j.date)=" & report_year & "and j.source='" & source & "' and ag.name='" & accgroup & "' and j.status != 'Open' order by j.date desc;
"
        Load_Datagridview(SPAS.Dgv_Report_6, sql, "boekingen")
        Format_drill_down()
        'SPAS.TC_Main.SelectedIndex = 5
        'SPAS.TC_Rapportage.SelectedTab = SPAS.TC_Boekingen

    End Sub

    Sub Drill_down_Bank_overview(ByVal i As Integer, ByVal j As Integer)

        MsgBox("drilldown_bank")
        Dim sqlpart1 As String = ""

        Select Case j
            Case 3 : sqlpart1 = " and b.name='_startsaldo_'"
            Case 4, 5, 6
            Case 7 : sqlpart1 = " and iban2 in (select accountno from bankacc)"
            Case Else
                Exit Sub
        End Select

        Dim bedrag As Integer = SPAS.Dgv_Rapportage_Overzicht.CurrentCell.Value
        Dim sql As String = "select b.date, b.seqorder, b.name, b.credit, b.debit, b.code, b.description, fk_journal_name
                             from " & Bank_table(report_year) & " b 
                             where iban='" & Trim(SPAS.Dgv_Rapportage_Overzicht.Rows(i).Cells(1).Value) & "' and extract(year from b.date)= " & report_year & sqlpart1 &
                             " order by b.seqorder desc"



        Load_Datagridview(SPAS.Dgv_Report_6, sql, "drilldown banktransacties")
        Format_drill_down_bank()
        'SPAS.TC_Rapportage.SelectedTab = SPAS.TC_Boekingen

    End Sub

    Sub Report_Closing()


        Dim Sqlc = QuerySQL("Select sql from query where category = 'Overzicht' and name='Transitieposten'")
        If IsNothing(Sqlc) Then Exit Sub
        Sqlc = Sqlc.Replace("[year]", report_year)

        RunSQL(Sqlc, "NULL", "Report Closing")

        Dim formatting As String = QuerySQL("select formatting from query where name='Transitieposten'")
        Dim arr_format() As String = Nothing
        If Not IsNothing(formatting) Then arr_format = formatting.Split(",")

        Load_Datagridview(SPAS.Dgv_Report_Year_Closing, Sqlc, "...")
        SPAS.Format_Datagridview(SPAS.Dgv_Report_Year_Closing, arr_format, False)

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

    Sub Format_drill_down_bank()
        'select j.date, a.name,j.amt1,j.name, j.type, j.description, j.iban,  ag.name,  j.fk_bank 
        Try
            With SPAS.Dgv_Report_6

                .Columns(0).HeaderText = "Datum"
                .Columns(0).Width = 80
                .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                .Columns(1).HeaderText = "Afschrift"
                .Columns(1).Width = 50
                .Columns(2).HeaderText = "Naam"
                .Columns(2).Width = 140
                .Columns(3).HeaderText = "Bij"
                .Columns(4).HeaderText = "Bij"
                For k = 3 To 4
                    .Columns(k).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    .Columns(k).DefaultCellStyle.Format = "N2"
                    .Columns(k).ReadOnly = True
                    .Columns(k).Width = 70
                Next k
                .Columns(5).HeaderText = "code"
                .Columns(5).Width = 40
                .Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                .Columns(5).DefaultCellStyle.ForeColor = Color.Gray
                .Columns(6).HeaderText = "Omschrijving"
                .Columns(6).Width = 500
                .Columns(7).Visible = False


            End With
        Catch ex As Exception
        End Try
    End Sub

End Module
