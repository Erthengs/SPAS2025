Imports System.Xml
Imports System.IO
Imports System.Net
Imports System.Text.RegularExpressions

Module In_excasso
    Sub Create_Incassolist()

        Dim d As DateTime
        Dim t1 As String
        Dim t2 As String
        Dim newDate As Date = Date.Now.AddMonths(1)
        Dim maxDate As Date = Date.Now.AddMonths(2)
        Dim minDate1 As Date = Date.Now.AddMonths(-1)

        SPAS.Dtp_Incasso_start.MinDate = CDate("01-" & minDate1.Month & "-" & minDate1.Year)

        SPAS.Dtp_Incasso_start.Value = CDate("01-" & SPAS.Dtp_Incasso_start.Value.Month & "-" & SPAS.Dtp_Incasso_start.Value.Year)
        If SPAS.Dtp_Incasso_start.Value.Year <> Date.Today.Year Then
            SPAS.Dtp_Incasso_start.Value = CDate("01-" & newDate.Month & "-" & newDate.Year)

        End If


        d = SPAS.Dtp_Incasso_start.Value.AddMonths(1)
        SPAS.Dtp_Incasso_end.Value = New DateTime(d.Year, d.Month, 1).AddDays(-1)
        'SPAS.Dtp_Incasso_start.MinDate = New Date(minDate1.Year, 1, 1)
        SPAS.Dtp_Incasso_start.MaxDate = New Date(maxDate.Year, maxDate.Month, 1)

        Dim isd As Date = SPAS.Dtp_Incasso_start.Value
        Dim MsgId = "Contract incasso " & Month(isd) & "-" & Year(isd)
        SPAS.Lbl_Incasso_job_name.Text = MsgId
        Dim qtopen, qtverwerkt As Integer

        t1 = Year(SPAS.Dtp_Incasso_start.Value) & "-" & Month(SPAS.Dtp_Incasso_start.Value) & "-01"
        t2 = Year(SPAS.Dtp_Incasso_end.Value) & "-" &
            Month(SPAS.Dtp_Incasso_end.Value) & "-" & SPAS.Dtp_Incasso_end.Value.Day

        'load lists and overview
        If SPAS.Rbn_Incasso_SEPA.Checked Then
            Load_Datagridview(SPAS.Dgv_Incasso, Create_Incasso(t1), "Me.Dtp_Incasso_start.ValueChanged")
        Else
            Load_Datagridview(SPAS.Dgv_Incasso, Create_Incasso_Bookings(t1), "Me.Dtp_Incasso_start.ValueChanged")
        End If

        'Load_Datagridview(SPAS.Dgv_incasso_totals, Create_Incasso_Totals(t1), "Create_Incassolist")
        SPAS.Prepare_Datagridview(SPAS.Dgv_incasso_totals, Create_Incasso_Totals(t1), {"TZ100", "TZ060", "NB080"})


        Dim Tot_amt = QuerySQL($"SELECT sum((co.donation+co.overhead)/term)
            FROM contract co  LEFT JOIN Target ta ON co.fk_target_id = ta.id LEFT JOIN Relation r ON co.fk_relation_id = r.id
            WHERE co.autcol = True AND co.startdate <= '{t1}' AND co.enddate > '{t1}'")

        Dim sql = QuerySQL($"select sql from query where name='Check_incasso'")
        sql = sql.replace("[date]", $"'{Year(SPAS.Dtp_Incasso_start.Value)}-{Month(SPAS.Dtp_Incasso_start.Value)}-01'")


        'Check_Existing_Incasso()
        SPAS.Lbl_Incasso_Error.Visible = False
        Dim journal_name As String = SPAS.Lbl_Incasso_job_name.Text
        qtopen = QuerySQL("select count(id) from journal where status = 'Open' and name ='" & journal_name & "'")
        qtverwerkt = QuerySQL("select count(id) from journal where status = 'Verwerkt' and name ='" & journal_name & "'")

        If qtopen > 0 Then
            SPAS.Lbl_Incasso_Status.Text = "Open"
            SPAS.Menu_Print.Enabled = True

            Dim Checksum = QuerySQL("Select Sum(amt1) from journal where name ='" & journal_name & "'")
            If Tot_amt <> Checksum Then
                Dim msg = $"Het totaalbedrag ({Tot_amt}) verschilt van de eerder gecreëerde incassojob ({Checksum}). De details zijn te zien via de radiobutton 'Verschillen' op deze pagina."
                SPAS.Lbl_Incasso_Error.Text = msg
                SPAS.Lbl_Incasso_Error.Visible = True
                SPAS.Rbn_Incasso_Verschillen.BackColor = Color.MistyRose
            Else
                SPAS.Rbn_Incasso_Verschillen.BackColor = Color.Transparent
            End If
        ElseIf qtverwerkt > 0 Then
            SPAS.Lbl_Incasso_Status.Text = "Verwerkt"
            SPAS.Menu_Print.Enabled = True

            Dim Checksum = QuerySQL("SELECT Sum(amt1) from journal where name ='" & journal_name & "'")
            If Tot_amt <> Checksum Then
                SPAS.Lbl_Incasso_Error.Text = "Opgeslagen incassojob is niet in lijn met contractdata"
            End If
        Else
            SPAS.Lbl_Incasso_Status.Text = "Nieuw"
            SPAS.Menu_Print.Enabled = False
        End If
        SPAS.Enable_Buttons(False, True)


    End Sub

    Sub Create_Incasso_Journals()
        'goed nadenken over het genereren van een naam voor een (groep) journaaltransactie
        'Dgv_incasso vervangen door dst
        Dim _isd As Date = SPAS.Dtp_Incasso_start.Value
        Dim isd As String = _isd.Year & "-" & _isd.Month & "-" & _isd.Day
        Dim s1 = Year(isd) & "-" & Month(isd) & "-01"
        Dim overhead As Integer
        Dim iban As String = Trim(SPAS.Cmx_Incasso_Bankaccount.Text)
        overhead = QuerySQL("SELECT value FROM settings WHERE label = 'overhead'")

        Dim journal_name = "Contract incasso " & Month(isd) & "-" & Year(isd)
        If QuerySQL("Select count(*) FROM journal WHERE name='" & journal_name & "'") > 0 Then
            MsgBox(journal_name & " bestaat al, graag eerst verwijderen alvorens een nieuwe aan te maken.")
            Exit Sub
        End If

        Dim incassodata = Collect_data2(Create_Incasso_Bookings(s1))

        Dim SQLstr = "INSERT INTO journal (date,status,type,amt1,description,source, fk_account,fk_relation,name,iban) VALUES "

        For x As Integer = 0 To incassodata.Rows.Count - 1

            SQLstr &= "('" &
                isd & "','Open','Contract','" & 'date/status
                Cur2(incassodata.Rows(x)(4)) & "','" & 'donation->amt1
                incassodata.Rows(x)(1) & "','Incasso','" & 'description/source
                incassodata.Rows(x)(6) & "','" & 'fk_account
                incassodata.Rows(x)(7) & "','" &
                journal_name & "','" & iban & "')," 'fk_relation/name

            If incassodata.Rows(x)(5) > 0 Then
                SQLstr &= "('" &
                isd & "','Open','Contract','" & 'date/status
                Cur2(incassodata.Rows(x)(5)) & "','" & 'overhead->amt1
                incassodata.Rows(x)(1) & "','Incasso','" & 'description/source
                overhead & "','" &   'incasso
                incassodata.Rows(x)(7) & "','" &
                journal_name & "','" & iban & "')," 'fk_relation/name
            End If
        Next

        RunSQL(Left(SQLstr, Strings.Len(SQLstr) - 1), "NULL", "Create_Incasso_Journals")
    End Sub
    Sub Create_SEPA_XML()


        Dim isd As Date = SPAS.Dtp_Incasso_start.Value
        Dim s1 = Year(isd) & "-" & Month(isd) & "-01"
        Dim MsgId = "Contract incasso " & Month(isd) & "-" & Year(isd)
        Dim f As System.IO.StreamWriter
        Dim filename = "Incassojob_" & Month(isd) & "_" & Year(isd) & ".xml"

        Dim incassodata = Collect_data2(Create_Incasso_Totals(s1))
        Dim nr As Integer = incassodata.Rows(0)(1) + incassodata.Rows(1)(1) + incassodata.Rows(2)(1)
        Dim amt = Replace(CDbl(incassodata.Rows(0)(2) + incassodata.Rows(1)(2) + incassodata.Rows(2)(2)).ToString("F2"), ",", ".")

        '@@@ moet gewijzigd worden naar nieuwe tabel
        Dim pi = MsgId
        Dim Inc_date As Date = Format(isd, "yyyy-MM-dd")
        Dim text_child = QuerySQL("Select value From settings WHERE label='text_bank_kind'")
        Dim text_elder = QuerySQL("Select value From settings WHERE label='text_bank_oudere'")
        Dim text_other = QuerySQL("Select value From settings WHERE label='text_bank_overig'")
        'retrieve account data

        Dim bankaccountdata = Collect_data2($"SELECT owner,accountno,bic,id2 FROM bankacc WHERE accountno='{SPAS.Cmx_Incasso_Bankaccount.Text}'")
        If IsDBNull(bankaccountdata.Rows(0)(2)) Or IsDBNull(bankaccountdata.Rows(0)(3)) Then
            MsgBox("Van een incassorekening moet de BIC en bank id ingevuld zijn.")
            Exit Sub
        End If

        Dim fnd As String = bankaccountdata.Rows(0)(0)
        Dim iban As String = bankaccountdata.Rows(0)(1)
        Dim bic = Strings.Trim(bankaccountdata.Rows(0)(2))
        Dim id2 = Strings.Trim(bankaccountdata.Rows(0)(3))


        Dim incassodata2 = Collect_data2(Create_Incasso(s1))

        Dim SelectFolder As New FolderBrowserDialog
        With SelectFolder
            .SelectedPath = My.Settings._excassopath
            .ShowNewFolderButton = True
        End With

        If (SelectFolder.ShowDialog() = DialogResult.OK) Then
            filename = SelectFolder.SelectedPath & "\" & filename
            My.Settings._excassopath = SelectFolder.SelectedPath
        End If


        f = My.Computer.FileSystem.OpenTextFileWriter(filename, False)

        'H E A D E R ====================

        f.WriteLine("<?xml version=""1.0"" encoding=""UTF-8"" ?>")
        f.WriteLine("<Document xmlns=""urn:iso:std:iso:20022:tech:xsd:pain.008.001.02"" xmlns:xsi=""http://www.w3.org/2001/xmlSchema-instance"">")
        f.WriteLine("<!-- HOET -->")
        f.WriteLine("<CstmrDrctDbtInitn>")
        f.WriteLine(Tabs(1) & "<GrpHdr>")
        f.WriteLine(Tabs(2) & "<MsgId>" & MsgId & "</MsgId>")
        f.WriteLine(Tabs(2) & "<CreDtTm>" & Format(Date.Now, "yyyy-MM-ddTHH:mm:ss") & "</CreDtTm>")
        f.WriteLine(Tabs(2) & "<NbOfTxs>" & nr.ToString & "</NbOfTxs>")
        f.WriteLine(Tabs(2) & "<CtrlSum>" & amt & "</CtrlSum>")
        f.WriteLine(Tabs(2) & "<InitgPty>")
        f.WriteLine(Tabs(3) & "<Nm>" & fnd & "</Nm>")
        f.WriteLine(Tabs(2) & "</InitgPty>")
        f.WriteLine(Tabs(1) & "</GrpHdr>")

        'payment info
        f.WriteLine(Tabs(1) & "<PmtInf>")

        f.WriteLine(Tabs(2) & "<PmtInfId>" & pi & "</PmtInfId>")
        f.WriteLine(Tabs(2) & "<PmtMtd>DD</PmtMtd>")
        f.WriteLine(Tabs(2) & "<BtchBookg>true</BtchBookg>")
        f.WriteLine(Tabs(2) & "<PmtTpInf>")
        f.WriteLine(Tabs(3) & "<SvcLvl>")
        f.WriteLine(Tabs(4) & "<Cd>SEPA</Cd>")
        f.WriteLine(Tabs(3) & "</SvcLvl>")
        f.WriteLine(Tabs(3) & "<LclInstrm>")
        f.WriteLine(Tabs(4) & "<Cd>CORE</Cd>")
        f.WriteLine(Tabs(3) & "</LclInstrm>")
        f.WriteLine(Tabs(4) & "<SeqTp>RCUR</SeqTp>")
        f.WriteLine(Tabs(2) & "</PmtTpInf>")

        f.WriteLine(Tabs(2) & "<ReqdColltnDt>" & Format(Inc_date, "yyyy-MM-dd") & "</ReqdColltnDt>")
        f.WriteLine(Tabs(2) & "<Cdtr>")
        f.WriteLine(Tabs(3) & "<Nm>" & fnd & "</Nm>")
        f.WriteLine(Tabs(2) & "</Cdtr>")
        f.WriteLine(Tabs(2) & "<CdtrAcct>")
        f.WriteLine(Tabs(3) & "<Id>")
        f.WriteLine(Tabs(4) & "<IBAN>" & iban & "</IBAN>")
        f.WriteLine(Tabs(3) & "</Id>")
        f.WriteLine(Tabs(2) & "</CdtrAcct>")
        f.WriteLine(Tabs(2) & "<CdtrAgt>")
        f.WriteLine(Tabs(3) & "<FinInstnId>")
        f.WriteLine(Tabs(4) & "<BIC>" & bic & "</BIC>")
        f.WriteLine(Tabs(3) & "</FinInstnId>")
        f.WriteLine(Tabs(2) & "</CdtrAgt>")
        f.WriteLine(Tabs(2) & "<ChrgBr>SLEV</ChrgBr>")
        f.WriteLine(Tabs(2) & "<CdtrSchmeId>")
        f.WriteLine(Tabs(3) & "<Id>")
        f.WriteLine(Tabs(4) & "<PrvtId>")
        f.WriteLine(Tabs(5) & "<Othr>")
        f.WriteLine(Tabs(6) & "<Id>" & id2 & "</Id>")
        f.WriteLine(Tabs(6) & "<SchmeNm>")
        f.WriteLine(Tabs(7) & "<Prtry>SEPA</Prtry>")
        f.WriteLine(Tabs(6) & "</SchmeNm>")
        f.WriteLine(Tabs(5) & "</Othr>")
        f.WriteLine(Tabs(4) & "</PrvtId>")
        f.WriteLine(Tabs(3) & "</Id>")
        f.WriteLine(Tabs(2) & "</CdtrSchmeId>")

        'individual payments
        For i = 0 To nr - 1
            'Dim ttype = IIf(incassodata2.Rows(i)(3) = "Kind", "KINDEREN", "OUDEREN")
            Dim relmsg = IIf(incassodata2.Rows(i)(3) = "Kind", text_child, IIf(incassodata2.Rows(i)(3) = "Oudere", text_elder, text_other))
            Dim relnam = incassodata2.Rows(i)(0)
            Dim iban2 = incassodata2.Rows(i)(2)
            Dim mancod = incassodata2.Rows(i)(4)
            Dim mandat = Format(CDate(incassodata2.Rows(i)(5)), "yyyy-MM-dd")
            Dim gift = Replace(incassodata2.Rows(i)(1).ToString, ",", ".")


            f.WriteLine(Tabs(2) & "<DrctDbtTxInf>")
            f.WriteLine(Tabs(3) & "<PmtId>")
            f.WriteLine(Tabs(4) & "<EndToEndId>" & Format(Date.Today, "yyyy-MM-dd") & "-" & Strings.Right("-0000" & i + 1, 6) & "</EndToEndId>")
            f.WriteLine(Tabs(3) & "</PmtId>")
            f.WriteLine(Tabs(4) & "<InstdAmt Ccy=""EUR"">" & gift & "</InstdAmt>")
            f.WriteLine(Tabs(3) & "<DrctDbtTx>")
            f.WriteLine(Tabs(4) & "<MndtRltdInf>")

            f.WriteLine(Tabs(5) & "<MndtId>" & mancod & "</MndtId>")
            f.WriteLine(Tabs(5) & "<DtOfSgntr>" & mandat & "</DtOfSgntr>")
            f.WriteLine(Tabs(5) & "<AmdmntInd>false</AmdmntInd>")

            f.WriteLine(Tabs(4) & "</MndtRltdInf>")
            f.WriteLine(Tabs(3) & "</DrctDbtTx>")
            f.WriteLine(Tabs(3) & "<DbtrAgt>")
            f.WriteLine(Tabs(4) & "<FinInstnId></FinInstnId>")
            f.WriteLine(Tabs(3) & "</DbtrAgt>")
            f.WriteLine(Tabs(3) & "<Dbtr>")
            f.WriteLine(Tabs(4) & "<Nm>" & relnam & "</Nm>")
            f.WriteLine(Tabs(3) & "<PstlAdr>")
            f.WriteLine(Tabs(4) & "<Ctry>NL</Ctry>")
            f.WriteLine(Tabs(3) & "</PstlAdr>")
            f.WriteLine(Tabs(3) & "</Dbtr>")
            f.WriteLine(Tabs(3) & "<DbtrAcct>")
            f.WriteLine(Tabs(4) & "<Id>")
            f.WriteLine(Tabs(5) & "<IBAN>" & iban2 & "</IBAN>")
            f.WriteLine(Tabs(4) & "</Id>")
            f.WriteLine(Tabs(3) & "</DbtrAcct>")
            f.WriteLine(Tabs(3) & "<Purp>")
            f.WriteLine(Tabs(4) & "<Cd>OTHR</Cd>")
            f.WriteLine(Tabs(3) & "</Purp>")
            f.WriteLine(Tabs(3) & "<RmtInf>")
            f.WriteLine(Tabs(4) & "<Ustrd>" & relmsg & "</Ustrd>")
            f.WriteLine(Tabs(3) & "</RmtInf>")
            f.WriteLine(Tabs(2) & "</DrctDbtTxInf>")
        Next

        f.WriteLine(Tabs(1) & "</PmtInf>")
        f.WriteLine("</CstmrDrctDbtInitn>")
        f.WriteLine("</Document>")

        f.Close()

        MsgBox("De incassojob is gecreëerd en beschikbaar.")

    End Sub

    Function Create_Incasso(date_start As String)
        Dim SQLstr = $"
            SELECT Concat(r.name, ', ', r.name_add) As Donateur
            ,sum((co.donation+co.overhead)/co.term) As Bedrag
            ,r.iban, ta.ttype As Doeltype
            ,CASE 
	            WHEN ta.ttype = 'Kind' Then Concat('k', r.reference)
	            WHEN ta.ttype = 'Oudere' Then Concat('o',r.reference)
                WHEN ta.ttype = 'Overig' Then Concat('v',r.reference)
            END As Mandaatcode
            ,CASE 
	            WHEN ta.ttype = 'Kind' Then r.date1
	            WHEN ta.ttype = 'Oudere' Then r.date2
                WHEN ta.ttype = 'Overig' Then r.date3
            END As Mandaatdatum
            FROM contract co 
            LEFT JOIN Target ta ON co.fk_target_id = ta.id
            LEFT JOIN Relation r ON co.fk_relation_id = r.id
            LEFT JOIN Account ac ON ac.f_key = ta.id
            WHERE co.autcol = True 
            AND co.startdate <= '{date_start}' 
            AND co.enddate > '{date_start}'
            AND ac.active = True
            AND 
            ((r.date1 <='{date_start}' AND ta.ttype = 'Kind') OR
            (r.date2 <='{date_start}' AND ta.ttype = 'Oudere') OR
            (r.date3 <='{date_start}' AND ta.ttype = 'Overig'))

            GROUP BY  r.reference, r.name, r.name_add, r.iban, ta.ttype, r.date1, r.date2, r.date3
            ORDER by  ta.ttype, r.reference

"

        Return SQLstr


    End Function
    Function Create_Incasso_Bookings(date_start As String)
        Dim SQLstr As String = $"
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
            AND co.startdate <= '{date_start}' 
            AND co.enddate > '{date_start}'
            AND ac.active = True
            AND 
            ((r.date1 <='{date_start}' AND ta.ttype = 'Kind') OR
            (r.date2 <='{date_start}' AND ta.ttype = 'Oudere') OR
            (r.date3 <='{date_start}' AND ta.ttype = 'Overig'))

            GROUP BY  ac.id,r.id,ta.name,ta.name_add, co.name, r.reference, r.name, r.name_add, r.iban, ta.ttype, r.date1, r.date2
            ORDER by  ta.ttype, r.reference
"
        Return SQLstr

    End Function


    '=============================================================================================================
    '==============  E   X   C   A   S   S   O   ================================================================= 
    '=============================================================================================================
    '@@@ ERROR 1: bij overmaken wordt het bedrag niet van de juiste post afgetrokken (eerst intern, dan extra, dan 
    'contract
    'ERROR 2: CP naam nog niet in omschrijving van journaalpost 
    'ERROR 3: schiet naar budget view ook al is "nulwaarden" geselecteerd
    'ERROR 4: bij meerdere excasso's op een dag is de naam niet uniek/Jobnummer wordt niet opgehoogd na creatie van een nieuwe job
    'ERROR 5: Geen toets of bedragen hoger zijn dan saldo als deze niet individueel bewerkt worden
    'ERROR 6: 
    'OPEN 5: selecteren van bestaande excasso's: moeten gepresenteerd worden 

    Sub Fill_Cmx_Excasso_Select_Combined()
        'this module combines existing excasso jobs and potential new ones (based on cp) in one combobox

        SPAS.Cmx_Excasso_Select.Items.Clear()

        Dim journaaldata = Collect_data2("SELECT distinct(name) FROM journal WHERE name ILIKE 'Excasso%' AND status = 'Open' GROUP By name, status")

        For x As Integer = 0 To journaaldata.Rows.Count - 1
            SPAS.Cmx_Excasso_Select.Items.Add(journaaldata.Rows(x)(0))
        Next

        Dim cpdata = Collect_data2("SELECT DISTINCT(cp.name), cp.name_add, cp.id FROM cp
                    LEFT JOIN target ta on fk_cp_id = cp.id LEFT JOIN contract co on fk_target_id = ta.id
                    WHERE co.enddate > current_date AND cp.active = 'True'")

        For x As Integer = 0 To cpdata.Rows.Count - 1
            SPAS.Cmx_Excasso_Select.Items.Add($"Nieuwe lijst {cpdata.Rows(x)(0)}, {cpdata.Rows(x)(1)} [{cpdata.Rows(x)(2)}]")
        Next


    End Sub




    Function Get_Excasso_data2(ByVal cp As String, type1 As String, type2 As String, type3 As String, naam1 As String, naam2 As String, dat As String)

        Dim Sqlstr As String =
            "
    select distinct(ac.id), ac.name, 

    CASE 
        WHEN extract(month from timestamp '" & dat & "')=1 Then case when round(max(ac.b_jan)::numeric,0) is distinct from null 
        THEN round(max(ac.b_jan)::numeric,0) else 0::numeric end
	    WHEN extract(month from timestamp '" & dat & "')=2 Then case when round(max(ac.b_feb)::numeric,0) is distinct from null 
	    THEN round(max(ac.b_feb)::numeric,0) else 0::numeric end 
	    WHEN extract(month from timestamp '" & dat & "')=3 Then case when round(max(ac.b_mar)::numeric,0) is distinct from null 
	    THEN round(max(ac.b_mar)::numeric,0) else 0::numeric end 
	    WHEN extract(month from timestamp '" & dat & "')=4 Then case when round(max(ac.b_apr)::numeric,0) is distinct from null 
	    THEN round(max(ac.b_apr)::numeric,0) else 0::numeric end 
	    WHEN extract(month from timestamp '" & dat & "')=5 Then case when round(max(ac.b_may)::numeric,0) is distinct from null 
	    THEN round(max(ac.b_may)::numeric,0) else 0::numeric end 
	    WHEN extract(month from timestamp '" & dat & "')=6 Then case when round(max(ac.b_jun)::numeric,0) is distinct from null 
	    THEN round(max(ac.b_jun)::numeric,0) else 0::numeric end 
	    WHEN extract(month from timestamp '" & dat & "')=7 Then case when round(max(ac.b_jul)::numeric,0) is distinct from null 
	    THEN round(max(ac.b_jul)::numeric,0) else 0::numeric end 
	    WHEN extract(month from timestamp '" & dat & "')=8 Then case when round(max(ac.b_aug)::numeric,0) is distinct from null 
	    THEN round(max(ac.b_aug)::numeric,0) else 0::numeric end 
	    WHEN extract(month from timestamp '" & dat & "')=9 Then case when round(max(ac.b_sep)::numeric,0) is distinct from null 
	    THEN round(max(ac.b_sep)::numeric,0) else 0::numeric end 
	    WHEN extract(month from timestamp '" & dat & "')=10 Then case when round(max(ac.b_oct)::numeric,0) is distinct from null 
	    THEN round(max(ac.b_oct)::numeric,0) else 0::numeric end 
	    WHEN extract(month from timestamp '" & dat & "')=11 Then case when round(max(ac.b_nov)::numeric,0) is distinct from null 
	    THEN round(max(ac.b_nov)::numeric,0) else 0::numeric end 
	    WHEN extract(month from timestamp '" & dat & "')=12 Then case when round(max(ac.b_dec)::numeric,0) is distinct from null 
	    THEN round(max(ac.b_dec)::numeric,0) else 0::numeric end 
    end as plAN,
-- calculated values: new form, then based on calculation of all up to given date; existing: than 
case 
 when (select round(sum(amt1)::numeric) from journal j where j.fk_account = ac.id and j.type = 'Contract' and j.name not ilike '" & naam1 & "' and j.date <='" & dat & "') is not distinct from null then 0::numeric
 else (select round(sum(amt1)::numeric) from journal j where j.fk_account = ac.id and j.type = 'Contract' and j.name not ilike '" & naam1 & "'and j.date <='" & dat & "')
end as saldo,
case 
 when (select round(sum(amt1)::numeric) from journal j where j.fk_account = ac.id and j.type = 'Extra' and j.name not ilike '" & naam1 & "'and j.date <='" & dat & "') is not distinct from null then 0::numeric
 else (select round(sum(amt1)::numeric) from journal j where j.fk_account = ac.id and j.type = 'Extra' and j.name not ilike '" & naam1 & "'and j.date <='" & dat & "')
end as extra,
case 
 when (select round(sum(amt1)::numeric) from journal j where j.fk_account = ac.id and j.type = 'Internal' and j.name not ilike '" & naam1 & "'and j.date <='" & dat & "') is not distinct from null then 0::numeric
 else (select round(sum(amt1)::numeric) from journal j where j.fk_account = ac.id and j.type = 'Internal' and j.name not ilike '" & naam1 & "'and j.date <='" & dat & "')
end as intern,

-- derived values:
case 
when (select -round(sum(amt1)::numeric) from journal j where j.fk_account = ac.id and j.type = 'Contract' and j.name ilike '%" & naam2 & "%'and j.date <='" & dat & "') is not distinct from null 
--or (select round(sum(amt1)::numeric) from journal j where j.fk_account = ac.id and j.type = 'Contract' and j.name ilike '%" & naam2 & "%'and j.date <='" & dat & "') < 0
then 0::numeric
 else (select -round(sum(amt1)::numeric) from journal j where j.fk_account = ac.id and j.type = 'Contract' and j.name ilike '%" & naam2 & "%'and j.date <='" & dat & "')
end as ""plan "",
case 
 when (select -round(sum(amt1)::numeric) from journal j where j.fk_account = ac.id and j.type = 'Extra' and j.name ilike '" & naam1 & "'and j.date <='" & dat & "') is not distinct from null then 0::numeric
 else (select -round(sum(amt1)::numeric) from journal j where j.fk_account = ac.id and j.type = 'Extra' and j.name ilike '" & naam1 & "'and j.date <='" & dat & "')
end as ""extra "",
case 
 when (select -round(sum(amt1)::numeric) from journal j where j.fk_account = ac.id and j.type = 'Internal' and j.name ilike '" & naam1 & "'and j.date <='" & dat & "') is not distinct from null then 0::numeric
 else (select -round(sum(amt1)::numeric) from journal j where j.fk_account = ac.id and j.type = 'Internal' and j.name ilike '" & naam1 & "'and j.date <='" & dat & "')
end as ""intern "",
    0::numeric as ∑Eur,
    0::numeric as ∑MLD

    from account ac
    left join target ta on ac.f_key = ta.id
    left join (select * from journal where date <='" & dat & "') j on j.fk_account = ac.f_key

    where ac.source ilike 'doel' 
    and ta.fk_cp_id = " & cp & " 
    and (ta.ttype = '" & type1 & "' or  ta.ttype='" & type2 & "' or ta.ttype='" & type3 & "')
    AND ta.active=true

    group by AC.ID
    order by ac.name

"
        Return Sqlstr

    End Function




    Sub Save_Excasso_job()
        '--------------Controles ------------------------------------------------------------------------------------------------------------:
        Dim exch As Decimal = CDec(SPAS.Tbx_Excasso_Exchange_rate.Text)
        Dim errmsg As String = ""
        Dim overhead = QuerySQL("SELECT value FROM settings WHERE label = 'overhead'")
        'If SPAS.Lbl_Excasso_Status.Text = "Verwerkt" Then errmsg &= "- verwerkte jobs kunnen niet verwijderd worden." & vbCrLf
        'If SPAS.Lbl_Excasso_Totaal.Text = "0" And SPAS.Lbl_Excasso_CP_Totaal.Text = "0" Then errmsg &= "- het totaalbedrag is 0. " & vbCrLf
        'If SPAS.Btn_Excasso_CP_Calculate.Enabled = True Then errmsg &= "- CP berekening is niet ververst." & vbCrLf
        'If SPAS.Btn_Excasso_Exchrate.Enabled Then errmsg &= "- Wijziging in wisselkoers is niet doorgevoerd." & vbCrLf
        'If CDec(SPAS.Tbx_Excasso_Exchange_rate.Text) = 0 Then errmsg &= "- Wisselkoers mag niet 0 zijn."
        If errmsg <> "" Then
            MsgBox("Er zijn de volgende fouten geconstateerd:" & vbCrLf & errmsg)
            Exit Sub
        End If

        'If SPAS.Btn_Excasso_CP_Calculate.Enabled = True Then
        If MsgBox("Heeft u de CP-bijdrage gecalculeerd?", vbYesNo) = vbNo Then Exit Sub
        'End If
        If CDec(SPAS.Tbx_Excasso_Exchange_rate.Text) = 0 Then
            If MsgBox("De wisselkoers is 0 of nog niet ververst. Wilt u doorgaan?", vbYesNo) = vbNo Then Exit Sub
        End If
        'Dim cntold As Integer = 0
        Dim j_name As String
        If Strings.Left(SPAS.Cmx_Excasso_Select.SelectedItem, 13) <> "Nieuwe lijst " Then
            j_name = SPAS.Cmx_Excasso_Select.SelectedItem
            'cntold = QuerySQL("SELECT count(*) FROM journal WHERE name ILIKE '%" & j_name & "'")
            RunSQL("DELETE FROM journal WHERE name ILIKE '%" & j_name & "'", "NULL", "Save_Excasso_job 1")

        Else
            ' j_name = "Excasso-" &
            '       IIf(SPAS.Cbx_Uitkering_Kind.Checked, "K", "") &
            '   IIf(SPAS.Cbx_Uitkering_Oudere.Checked, "O", "") &
            '   IIf(SPAS.Cbx_Uitkering_Overig.Checked, "V", "") & "-" &
            'Left(Mid(SPAS.Cmx_Excasso_Select.Text, 14), 4) & "-" &
            '       SPAS.Dtp_Excasso_Start.Value
            Dim cnt = QuerySQL("SELECT count(distinct name) FROM journal WHERE name LIKE '" & j_name & "%'")
            j_name &= "-" & (cnt + 1).ToString

        End If




        'determine values that are valid for all journalpost within this job:
        Dim j_contr, j_extr, j_inte
        Dim dif1 As Decimal = 0
        Dim j_fkac As Integer
        'Dim j_name = SPAS.Lbl_Excasso_Job_Name.Text
        Dim j_date = SPAS.Dtp_Excasso_Start.Value.Year & "-" & SPAS.Dtp_Excasso_Start.Value.Month &
            "-" & SPAS.Dtp_Excasso_Start.Value.Day
        Dim SQLstr = "INSERT INTO journal(name, date,status,source,amt1,amt2,fk_account,description,type,cpinfo,iban) VALUES"
        Dim j_desc As String = "", j_desc2 As String = ""
        'cpinfo: cpid-Tbx_Excasso_Norm1- ..2-..3-Tbx_Excasso_CP1-..2-..3
        'Dim unit1 As String = IIf(SPAS.Btn_Excasso_Base1.Text = "€", "1", "0")
        'Dim unit2 As String = IIf(SPAS.Btn_Excasso_Base2.Text = "€", "1", "0")
        'Dim unit3 As String = IIf(SPAS.Btn_Excasso_Base3.Text = "€", "1", "0")


        'Dim j_cpinfo As String = SPAS.Lbl_Excasso_CPid.Text & "-" &
        'SPAS.Tbx_Excasso_Norm1.Text & "-" & SPAS.Tbx_Excasso_Norm2.Text & "-" & SPAS.Tbx_Excasso_Norm3.Text & "-" &
        '  SPAS.Tbx_Excasso_CP1.Text & "-" & SPAS.Tbx_Excasso_CP2.Text & "-" & SPAS.Tbx_Excasso_CP3.Text & "-" &
        '   unit1 & "-" & unit2 & "-" & unit3
        Dim j_cp_fk = QuerySQL("Select id From account where f_key='" & SPAS.Lbl_Excasso_CPid.Text & "'")
        'Dim j_iban = Strings.Trim(QuerySQL("
        'Select Case bankacc.accountno FROM cp LEFT JOIN bankacc ON bankacc.id = cp.fk_bankacc_id WHERE cp.id='" & SPAS.Lbl_Excasso_CPid.Text & "'"))
        Dim j_iban = "NL66RABO0177491310"  '@@@tijdelijke workaround



        For x As Integer = 0 To SPAS.Dgv_Excasso2.Rows.Count - 1
            j_contr = CDec(SPAS.Dgv_Excasso2.Rows(x).Cells(6).Value)
            j_extr = CDec(SPAS.Dgv_Excasso2.Rows(x).Cells(7).Value)
            j_inte = CDec(SPAS.Dgv_Excasso2.Rows(x).Cells(8).Value)
            j_fkac = SPAS.Dgv_Excasso2.Rows(x).Cells(0).Value
            j_desc = "Uitkering aan " & SPAS.Dgv_Excasso2.Rows(x).Cells(1).Value
            j_desc2 = "Distribution costs " & SPAS.Cmx_Excasso_Select.SelectedText


            If j_contr > 0 Then
                'SQLstr &= $"('{j_name}','{j_date}','Open','Uitkering','{-j_contr}','{-CInt(j_contr * exch)}','{j_fkac}','{j_desc}','Contract','{j_cpinfo}','{j_iban}'),"
            End If
            If j_extr > 0 Then
                'SQLstr &= $"('{j_name}','{j_date}','Open','Uitkering','{-j_extr}','{-CInt(j_extr * exch)}','{j_fkac}','{j_desc}','Extra','{j_cpinfo}','{j_iban}'),"
            End If
            If j_inte > 0 Then
                'SQLstr &= $"('{j_name}','{j_date}','Open','Uitkering','{-j_inte}','{-CInt(j_inte * exch)}','{j_fkac}','{j_desc}','Internal','{j_cpinfo}','{j_iban}'),"
            End If
        Next
        'cp transactie toevoegen
        ' Dim j_cp = CDec(SPAS.Lbl_Excasso_CP_Totaal.Text)
        'If j_cp > 0 Then
        'from overhead
        'SQLstr &= $"('Intern tbv CP {j_name}','{j_date}','Verwerkt','Intern','{Cur2(j_cp) * -1}','{-CInt(j_cp * exch)}','{overhead}','{j_desc2}', 'CP','{j_cpinfo}',''),"
        'to CP account
        'SQLstr &= $"('Intern tbv CP {j_name}','{j_date}','Verwerkt','Intern','{Cur2(j_cp)}','{CInt(j_cp * exch)}','{j_cp_fk}','{j_desc2}', 'CP','{j_cpinfo}',''),"
        'from CP account 
        'SQLstr &= $"('{j_name}','{j_date}','Open','Uitkering','{Cur2(j_cp) * -1}','{-CInt(j_cp * exch)}','{j_cp_fk}','{j_desc2}', 'CP','{j_cpinfo}','{j_iban}'),"
        'End If
        'Clipboard.Clear()
        'Clipboard.SetText(SQLstr)

        SQLstr = Strings.Left(SQLstr, Strings.Len(SQLstr) - 1) 'remove the last comma
        RunSQL(SQLstr, "NULL", "Save Excasso job 2")
        'If cntold > 0 Then
        If Strings.Left(SPAS.Cmx_Excasso_Select.SelectedItem, 13) = "Nieuwe lijst " Then
            SPAS.Cmx_Excasso_Select.Items.Add(j_name)
            SPAS.Cmx_Excasso_Select.SelectedIndex = SPAS.Cmx_Excasso_Select.Items.Count - 1
        End If


    End Sub




    Sub Load_Existing_Excasso()
        '' Loads an excasso form that has already been saved in the past, but is noet yet posted to the G/L.
        '''
        If SPAS.Cmx_Excasso_Select.SelectedItem = "" Then Exit Sub 'alleen lijst genereren als gekozen is voor een bestaande of nieuwe lijst

        Dim str1() As String = Split(QuerySQL($"SELECT cpinfo FROM journal WHERE name ='{ SPAS.Cmx_Excasso_Select.SelectedItem}'"), "-")
        Dim str2() As String = Split(SPAS.Cmx_Excasso_Select.SelectedItem, "-")
        Dim cp As String = str1(0)
        Dim dat As String = ""

        With SPAS
            .Btn_Excasso_Delete.Enabled = True
            .Btn_Excasso_Print.Enabled = True

            'calculate actual exchange rate
            Dim exr = QuerySQL($"SELECT sum(amt2)/sum(amt1) FROM journal WHERE name ='{ .Cmx_Excasso_Select.SelectedItem}'")
            If IsDBNull(exr) Then exr = 0
            .Tbx_Excasso_Exchange_rate.Text = Math.Round(GetDouble(exr), 2)

            ' determine date
            .Dtp_Excasso_Start.Value = CDate(QuerySQL($"SELECT date FROM journal WHERE name='{ .Cmx_Excasso_Select.SelectedItem}'"))
            .Dtp_Excasso_Start.Enabled = False
            dat = .Dtp_Excasso_Start.Value.Year & "-" & .Dtp_Excasso_Start.Value.Month & "-" & .Dtp_Excasso_Start.Value.Day
            'Dim cp_amount = QuerySQL($"Select sum(amt1) FROM journal WHERE name ='{ .Cmx_Excasso_Select.SelectedItem}' AND type='CP' AND amt1<='0.00'")
        End With

        Dim s2 As String = Get_Excasso_data2(cp, "Kind", "Oudere", "Overig", SPAS.Cmx_Excasso_Select.SelectedItem, SPAS.Cmx_Excasso_Select.SelectedItem, dat)
        If s2 = "" Then Exit Sub

        'Load_Datagridview(SPAS.Dgv_Excasso2, s2, "Call_Excasso_form2")

        SPAS.Dgv_Excasso2.DataSource = Collect_data2(s2)
        Prepare_Excasso()
        With SPAS.Dgv_Excasso_numbers

            .Columns(1).ReadOnly = True
            .Columns(2).Visible = False
            .Rows(0).Cells(1).Value = InStr(str2(1), "K") > 0

            .Rows(0).Cells(6).Value = str1(1)
            .Rows(0).Cells(7).Value = IIf(str1(7) = "1", "€", "%")
            .Rows(0).Cells(8).Value = str1(4)

            .Rows(1).Cells(1).Value = InStr(str2(1), "O") > 0
            .Rows(1).Cells(6).Value = str1(2)
            .Rows(1).Cells(7).Value = IIf(str1(8) = "1", "€", "%")
            .Rows(1).Cells(8).Value = str1(5)

            .Rows(1).Cells(10).Value = CInt(str1(4)) + CInt(str1(5)) + CInt(str1(6))
            .Rows(1).Cells(11).Value = .Rows(1).Cells(10).Value * CInt(SPAS.Tbx_Excasso_Exchange_rate.Text)
            .Rows(2).Cells(1).Value = InStr(str2(1), "V") > 0

            .Rows(2).Cells(6).Value = str1(3)
            .Rows(2).Cells(7).Value = IIf(str1(9) = "1", "€", "%")
            .Rows(2).Cells(8).Value = str1(6)

        End With
        Calculate_Excasso_Totals(2)


    End Sub

    Sub Load_New_Excasso(ByVal modus As Boolean)
        '' Loads an excasso form that has already been saved in the past, but is noet yet posted to the G/L.
        '''
        If SPAS.Cmx_Excasso_Select.SelectedItem = "" Then Exit Sub 'alleen lijst genereren als gekozen is voor een bestaande of nieuwe lijst

        'Dim str1() As String = Split(QuerySQL($"SELECT cpinfo FROM journal WHERE name ='{ SPAS.Cmx_Excasso_Select.SelectedItem}'"), "-")
        'Dim str2() As String = Split(SPAS.Cmx_Excasso_Select.SelectedItem, "-")
        Dim pos1 As Integer = Strings.InStr(SPAS.Cmx_Excasso_Select.SelectedItem, "[")
        Dim cp = Strings.Mid(SPAS.Cmx_Excasso_Select.SelectedItem, pos1 + 1, Len(SPAS.Cmx_Excasso_Select.SelectedItem) - pos1 - 1)
        Dim dat As String = ""
        Dim kind As String = "--"
        Dim oudere As String = "--"
        Dim overig As String = "--"

        With SPAS
            .Btn_Excasso_Delete.Enabled = False
            .Btn_Excasso_Print.Enabled = True

            'determine exchange rate based on previous stored value
            Dim exr = Tbx2Dec(My.Settings._exrate)
            If IsDBNull(exr) Then exr = 0
            .Tbx_Excasso_Exchange_rate.Text = Math.Round(GetDouble(exr), 2)

            ' determine date
            .Dtp_Excasso_Start.Value = Date.Today
            .Dtp_Excasso_Start.Enabled = True
            dat = .Dtp_Excasso_Start.Value.Year & "-" & .Dtp_Excasso_Start.Value.Month & "-" & .Dtp_Excasso_Start.Value.Day

            'bepaal defaultwaarde voor doeltype
            If modus = True Then
                kind = IIf(QuerySQL($"Select count(*) from target where fk_cp_id={cp} and ttype= 'Kind' and active") > 0, "Kind", "--")
                oudere = IIf(QuerySQL($"Select count(*) from target where fk_cp_id={cp} and ttype= 'Oudere' and active") > 0, "Oudere", "--")
                overig = IIf(QuerySQL($"Select count(*) from target where fk_cp_id={cp} and ttype= 'Overig' and active") > 0, "Overig", "--")
            Else
                kind = IIf(.Dgv_Excasso_numbers.Rows(0).Cells(1).Value, "Kind", "--")
                oudere = IIf(.Dgv_Excasso_numbers.Rows(1).Cells(1).Value, "Oudere", "--")
                overig = IIf(.Dgv_Excasso_numbers.Rows(2).Cells(1).Value, "Overig", "--")
            End If

        End With


        Dim s2 As String = Get_Excasso_data2(cp, kind, oudere, overig, SPAS.Cmx_Excasso_Select.SelectedItem, "", dat)
        If s2 = "" Then Exit Sub

        SPAS.Dgv_Excasso2.DataSource = Collect_data2(s2)

        Prepare_Excasso()
        With SPAS.Dgv_Excasso_numbers
            .Rows(0).Cells(2).Style.BackColor = Color.CornflowerBlue
            .Columns(1).ReadOnly = False
            .Columns(2).Visible = True
            .Rows(0).Cells(1).Value = (kind <> "--")

            .Rows(0).Cells(6).Value = 4
            .Rows(0).Cells(7).Value = "€"
            '.Rows(0).Cells(8).Value = str1(4)

            .Rows(1).Cells(1).Value = (oudere <> "--")
            .Rows(1).Cells(6).Value = 3
            .Rows(1).Cells(7).Value = "€"
            '.Rows(1).Cells(8).Value = str1(5)

            '.Rows(1).Cells(10).Value = CInt(str1(4)) + CInt(str1(5)) + CInt(str1(6))
            .Rows(1).Cells(11).Value = .Rows(1).Cells(10).Value * CInt(SPAS.Tbx_Excasso_Exchange_rate.Text)
            .Rows(2).Cells(1).Value = (overig <> "--")

            .Rows(2).Cells(6).Value = 3
            .Rows(2).Cells(7).Value = "€"
            '.Rows(2).Cells(8).Value = str1(6)

        End With
        SPAS.Cbx_CP_Automatisch.Checked = True
        Calculate_Excasso_Totals(2)


    End Sub
    Sub Prepare_Excasso()
        With SPAS.Dgv_Excasso_numbers
            .RowCount = 3
            For r = 0 To 2
                .Rows(r).Height = 24
            Next
            .Rows(0).Cells(0).Value = "Kind"
            .Rows(0).Cells(2).Value = "Gepland"
            .Rows(0).Cells(3).Value = "Contract"
            .Rows(0).Cells(9).Value = "Uitkering"

            .Rows(1).Cells(0).Value = "Oudere"
            .Rows(1).Cells(2).Value = "Saldo's"
            .Rows(1).Cells(3).Value = "Extra"
            .Rows(1).Cells(9).Value = "Naar CP"

            .Rows(2).Cells(0).Value = "Overig"
            .Rows(2).Cells(2).Value = "Nulwaarden"
            .Rows(2).Cells(3).Value = "Intern"
            .Rows(2).Cells(9).Value = "Totaal"



        End With
    End Sub

    Sub Calculate_Excasso_Totals(ByVal mode As Integer)

        If SPAS.Dgv_Excasso2.Rows.Count = 0 Then Exit Sub
        If IsNothing(SPAS.Tbx_Excasso_Exchange_rate.Text) Or SPAS.Tbx_Excasso_Exchange_rate.Text = "" Then SPAS.Tbx_Excasso_Exchange_rate.Text = 0

        '1 = nulwaarden, 2 = budgetwaarden, 3= saldowaarden
        If SPAS.Dgv_Excasso_numbers.Rows(0).Cells(2).Style.BackColor = Color.CornflowerBlue Then
            mode = 2
        ElseIf SPAS.Dgv_Excasso_numbers.Rows(1).Cells(2).Style.BackColor = Color.CornflowerBlue Then
            mode = 3
        Else
            mode = 1
        End If



        'after changing value
        Dim str1() As String = Split(QuerySQL($"SELECT cpinfo FROM journal WHERE name ='{SPAS.Cmx_Excasso_Select.SelectedItem}'"), "-")
 

        With SPAS.Dgv_Excasso2
            For x As Integer = 0 To .Rows.Count - 1
                If mode > 1 Then
                    .Rows(x).Cells(6).Value = IIf(.Rows(x).Cells(mode).Value > 0, .Rows(x).Cells(mode).Value, 0)
                    .Rows(x).Cells(7).Value = IIf(.Rows(x).Cells(4).Value > 0, .Rows(x).Cells(4).Value, 0)
                    .Rows(x).Cells(8).Value = IIf(.Rows(x).Cells(5).Value > 0, .Rows(x).Cells(5).Value, 0)
                Else
                    .Rows(x).Cells(6).Value = 0
                    .Rows(x).Cells(7).Value = 0
                    .Rows(x).Cells(8).Value = 0
                End If

                .Rows(x).Cells(9).Value = CInt(.Rows(x).Cells(6).Value) + CInt(.Rows(x).Cells(7).Value) + CInt(.Rows(x).Cells(8).Value)
                .Rows(x).Cells(10).Value = Math.Round(.Rows(x).Cells(9).Value * CInt(SPAS.Tbx_Excasso_Exchange_rate.Text), 0)



            Next x

        End With


        With SPAS.Dgv_Excasso_numbers
            '.Rows(1).Cells(10).Value = CInt(str1(4)) + CInt(str1(5)) + CInt(str1(6))
            .Rows(0).Cells(4).Value = SPAS.CalculateColumnCount(SPAS.Dgv_Excasso2, 6)
            .Rows(0).Cells(5).Value = SPAS.CalculateColumnSum(SPAS.Dgv_Excasso2, 6)
            .Rows(0).Cells(10).Value = SPAS.CalculateColumnSum(SPAS.Dgv_Excasso2, 9)
            .Rows(0).Cells(11).Value = SPAS.CalculateColumnSum(SPAS.Dgv_Excasso2, 10)

            .Rows(1).Cells(4).Value = SPAS.CalculateColumnCount(SPAS.Dgv_Excasso2, 7)
            .Rows(1).Cells(5).Value = SPAS.CalculateColumnSum(SPAS.Dgv_Excasso2, 7)
            '.Rows(1).Cells(11).Value = .Rows(1).Cells(10).Value * CInt(SPAS.Tbx_Excasso_Exchange_rate.Text)

            .Rows(2).Cells(4).Value = SPAS.CalculateColumnCount(SPAS.Dgv_Excasso2, 8)
            .Rows(2).Cells(5).Value = SPAS.CalculateColumnSum(SPAS.Dgv_Excasso2, 8)
            .Rows(2).Cells(10).Value = .Rows(0).Cells(10).Value + .Rows(1).Cells(10).Value
            .Rows(2).Cells(11).Value = .Rows(0).Cells(11).Value + .Rows(1).Cells(11).Value

            If SPAS.Cbx_CP_Automatisch.Checked Then
                .Rows(0).Cells(8).Value = CInt(.Rows(0).Cells(4).Value) * CInt(.Rows(0).Cells(6).Value)
                .Rows(1).Cells(8).Value = CInt(.Rows(1).Cells(4).Value) * CInt(.Rows(1).Cells(6).Value)
                .Rows(2).Cells(8).Value = CInt(.Rows(2).Cells(4).Value) * CInt(.Rows(2).Cells(6).Value)
                .Rows(1).Cells(10).Value = CInt(.Rows(0).Cells(8).Value) + CInt(.Rows(1).Cells(8).Value) + CInt(.Rows(2).Cells(8).Value)
                .Rows(1).Cells(11).Value = .Rows(1).Cells(10).Value * CInt(SPAS.Tbx_Excasso_Exchange_rate.Text)
                .Rows(2).Cells(10).Value = CInt(.Rows(0).Cells(10).Value) + CInt(.Rows(1).Cells(10).Value)
                .Rows(2).Cells(11).Value = .Rows(2).Cells(10).Value * CInt(SPAS.Tbx_Excasso_Exchange_rate.Text)
            End If

        End With

        SPAS.Prepare_Datagridview(SPAS.Dgv_Excasso2, Nothing, {"HZ010", "TZ122", "JZ052", "JZ052", "JZ052", "JZ052", "JB052", "JB053", "JB052", "IG052", "IG054"})
    End Sub

    Sub Calculate_CP_Values1()



        If Not SPAS.isManualChange Then Exit Sub
        If SPAS.Dgv_Excasso2.Rows.Count = 0 Then Exit Sub
        If IsDBNull(SPAS.Dgv_Excasso_numbers.CurrentCell.Value) Then Exit Sub

        Dim colindex As Integer = SPAS.Dgv_Excasso_numbers.CurrentCell.ColumnIndex
        Dim rowindex As Integer = SPAS.Dgv_Excasso_numbers.CurrentCell.RowIndex
        If colindex <> 8 And colindex <> 10 Then Exit Sub

        Try
            If colindex = 8 Then SPAS.Dgv_Excasso_numbers.Rows(rowindex).Cells(8).Value = SPAS.Dgv_Excasso_numbers.CurrentCell.Value * 10
        Catch ex As Exception

        End Try

    End Sub

End Module
