
Imports System.Windows
Imports Npgsql
'Imports PdfSharp.Pdf
Imports System.IO
Imports System.Diagnostics.Tracing
Imports Microsoft.EntityFrameworkCore.Update.Internal

Module bank

    Sub Download_Bank_Transactions()
        'initialize variables
        Dim csv As String = ""

        'get csv
        SPAS.OpenFileDialog1.Title = "Selecteer een bankafschrift"
        SPAS.OpenFileDialog1.FileName = ""
        SPAS.OpenFileDialog1.InitialDirectory = "" 'My.Settings._bankpath
        SPAS.OpenFileDialog1.Filter = "ING/Rabo bestanden|*.csv"

        If SPAS.OpenFileDialog1.ShowDialog() = DialogResult.OK Then csv = SPAS.OpenFileDialog1.FileName
        If csv = "" Then Exit Sub

        Upload_CSV(csv)

        SPAS.Fill_bank_transactions("Download_Bank_Transactions")
        Categorize_Bank_Transactions()
    End Sub
    Sub Load_Bank_csv_from_folder()
        Dim SelectFolder As New FolderBrowserDialog
        Dim fold As String = ""


        With SelectFolder
            .SelectedPath = My.Settings._bankpath
            .ShowNewFolderButton = False
        End With

        If (SelectFolder.ShowDialog() = DialogResult.OK) Then
            fold = SelectFolder.SelectedPath
            My.Settings._bankpath = SelectFolder.SelectedPath
        Else
            Exit Sub
        End If



        Dim dir As New DirectoryInfo(fold)
        'Dim newdir As String

        For Each f In dir.GetFiles()
            'If Left(f.Name, 3) = "CSV" Then newdir = fold + Mid(f.Name, 11, 2) Else newdir = fold + Mid(f.Name, 23, 2) & "\"
            'If (Not System.IO.Directory.Exists(fold & "\nieuw\")) Then
            'System.IO.Directory.CreateDirectory(fold & "\nieuw\")
            'End If

            If Strings.Right(f.Name, 4) = ".csv" Then

                Upload_CSV(SelectFolder.SelectedPath & "\" & f.Name)


            End If
        Next
        'Categorize_Bank_Transactions()
        SPAS.Fill_bank_transactions("Load_Bank_csv_from_folder")

    End Sub


    Sub Upload_CSV(ByVal csv As String)


        Dim bank, _dat, _date, des As String
        Dim amt
        Dim _exch As Decimal
        Dim cost As Decimal
        Dim lastdate As Date
        Dim amt2 As Decimal
        Dim filename As String = "testfilename"
        Dim SQLstr As String = ""
        Dim SQLstr2 As String = "INSERT INTO journal(name,date,status,description,source,amt1,fk_account,
                               fk_bank,fk_relation,iban) VALUES "
        Dim delimiter As String
        Dim content = IO.File.ReadAllText(csv)  'My.Computer.FileSystem.ReadAllText(csv)

        delimiter = IIf(content.Contains(""";"""), """;""", """,""")

        'delimiter = IIf(InStr(csv, "ING") > 0, """;""", """,""")
        IO.File.WriteAllText(csv, IO.File.ReadAllText(csv).Replace(delimiter, """|"""))
        Dim items = (From line In IO.File.ReadAllLines(csv)
                     Select Array.ConvertAll(line.Split("|"c), Function(v) _
                     v.ToString.TrimStart(""" ".ToCharArray).TrimEnd(""" ".ToCharArray))).ToArray

        'store output array in datatable
        Dim Bank_DT As New DataTable
        For x As Integer = 0 To 50 'items(0).GetUpperBound(0)
            Bank_DT.Columns.Add()
        Next

        For Each a In items
            Dim dr As DataRow = Bank_DT.NewRow
            dr.ItemArray = a
            Bank_DT.Rows.Add(dr)
        Next

        Dim csv1 = StrReverse(csv)
        filename = StrReverse(Strings.Left(csv1, InStr(csv1, "\") - 1))

        'determine Bank
        bank = IIf(InStr(Bank_DT.Rows(1)(2), "RABO") > 0, "rabo", "ing")

        'check if there are already bank transactions in database / bank_id not used elsewhere
        Dim bank_id = QuerySQL("SELECT MAX(id) FROM bank")
        If IsDBNull(bank_id) Then bank_id = 0  '@@@gaat mogelijk fout als id niet gereset wordt

        '------------------------------------------'specific for RABO
        'Check on bank account number + sequence number

        If bank = "rabo" Then
            '=====check is rabo csv is the following one
            Dim Last_seq_order = QuerySQL("Select MAX(seqorder) FROM bank WHERE iban='" & Bank_DT.Rows(1)(0) & "'")
            'MsgBox(CInt(Bank_DT.Rows(1)(3)).ToString & " - " & Last_seq_order.ToString)
            '1) checks on import file
            If Not IsDBNull(Last_seq_order) Then
                If CInt(Bank_DT.Rows(1)(3)) <> Last_seq_order + 1 Then
                    MsgBox("Er zijn nog niet ingeladen tussenliggende banktransacties. Doe dit s.v.p eerst. ")
                    Exit Sub
                End If
            End If

            'load contents of datatable into database through composing sql statements
            For i = 1 To Bank_DT.Rows.Count - 1
                '==== Transformations =====
                'a) determine in which column the amount must be stored
                Dim debit = Cur2(IIf(Bank_DT.Rows(i)(6) < 0, Bank_DT.Rows(i)(6) * -1, 0))
                Dim credit = Cur2(IIf(Bank_DT.Rows(i)(6) > 0, Bank_DT.Rows(i)(6), 0))
                'debit = Replace(debit, ",", ".")
                'credit = Replace(credit, ",", ".")
                Dim descr = Strings.Trim(Bank_DT.Rows(i)(19)) & " " & Strings.Trim(Bank_DT.Rows(i)(20)) & " " _
                            & Strings.Trim(Bank_DT.Rows(i)(21)) & Strings.Trim(Bank_DT.Rows(i)(14))
                descr = Replace(descr, "'", "")
                'b) retrieve relation based on bankacc
                Dim iban2 = Bank_DT.Rows(i)(8)
                Dim relation_name As String = QuerySQL("SELECT CONCAT(name, ', ',name_add) FROM relation WHERE iban='" & iban2 & "'")
                If relation_name = "" Then relation_name = Bank_DT.Rows(i)(9)

                '==== Creating SQL strings =====
                SQLstr &=
                    "INSERT INTO bank(iban, currency, seqorder, date, debit,credit,iban2,
                    name, code, batchid,description,exch_rate, fk_journal_name, filename) VALUES('" &
                    Trim(Bank_DT.Rows(i)(0).ToString()) & "','" &      'iban
                    Bank_DT.Rows(i)(1).ToString() & "','" &      'currency
                    Bank_DT.Rows(i)(3) & "','" &  'seqorder
                    Bank_DT.Rows(i)(4) & "','" &    'date
                    debit & "','" &    'debit
                    credit & "','" &    'credit
                    Bank_DT.Rows(i)(8) & "','" &    'iban2
                    relation_name & "','" &    'name   
                    Trim(Bank_DT.Rows(i)(13)) & "','" &    'code
                    Bank_DT.Rows(i)(14) & "','" &    'batchid
                    Trim(descr) &    'description
                     "',1,'" & Left(filename, 4) & "." & i.ToString & "','" & filename & "');" & vbCrLf  'exchange rate

                SQLstr2 &= "('" & Left(filename, 4) & "." & i.ToString & "','" & Bank_DT.Rows(i)(4) & "','Verwerkt','" & descr & "','Bank','" &
                           Cur2(Bank_DT.Rows(i)(6)) & "','" & nocat & "',0,0,
                           '" & Bank_DT.Rows(i)(0).ToString() & "'),"
            Next

        Else 'bank is ING
            '========= check on right file
            'Check on bank account number, date interval between last and current download
            Dim csvdate() As String = Split(filename, "_")
            Dim startdate = CDate(csvdate(1))
            '1 check if the file has already been uploaded...
            If QuerySQL("SELECT COUNT(id) FROM bank WHERE filename='" & filename & "'") > 0 Then
                MsgBox("Dit bankbestand is al geladen")
                Exit Sub
            End If
            Dim ld = QuerySQL("SELECT MAX(date)::date FROM bank WHERE iban='" & Bank_DT.Rows(1)(2).ToString() & "'")
            If Not IsDBNull(ld) Then
                lastdate = CDate(ld)
                'MsgBox(DateDiff(DateInterval.Day, lastdate, startdate)).ToString()
                If DateDiff(DateInterval.Day, lastdate, startdate) > 30 Then
                    MsgBox("De laatste banktransactie van deze rekening dateert van " &
                              lastdate.ToString & ". De startdatum van dit bankbestand is " &
                              startdate & ". Er zit dus minimaal een maand tussen. " &
                              "Upload s.v.p. eerst de tussenliggende banktransacties.")
                    Exit Sub
                End If
            End If

            For i = 1 To Bank_DT.Rows.Count - 1
                '==== Transformations =====
                amt = Cur2(Bank_DT.Rows(i)(6))
                'amt = Replace(amt, ",", ".")
                des = Bank_DT.Rows(i)(8)
                _dat = Bank_DT.Rows(i)(0)
                _exch = 1
                amt2 = 0
                cost = 0
                _date = Strings.Left(_dat, 4) & "-" & Mid(_dat, 5, 2) & "-" & Strings.Right(_dat, 2)


                If InStr(des, "MDL Koers: ") > 0 Then
                    _exch = 1 / Mid(des, Strings.Left(InStr(des, "MDL Koers: ") + 11, 10), 8)
                    amt2 = CInt(Mid(des, InStr(des, "Valuta: ") + 8, InStr(des, "MDL Koers: ") - InStr(des, "Valuta: ") - 9))
                End If
                If InStr(des, "Kosten: ") > 0 Then
                    cost = Tbx2Dec(Mid(des, InStr(des, "Kosten: ") + 8, InStr(des, "EUR Valutadatum: ") - InStr(des, "Kosten: ") - 9))
                    'MsgBox(Mid(des, InStr(des, "Kosten: ") + 8, 4))
                End If

                '==== Creating SQL strings =====
                SQLstr &=
                    "INSERT INTO bank(iban, currency, seqorder, date, debit,credit,iban2,
                    name, code, batchid,description,exch_rate,amt_cur, fk_journal_name,filename,cost) VALUES('" &
                    Bank_DT.Rows(i)(2).ToString() & "','" &      'iban
                    "EUR','" &      'currency
                    "0','" &  'seqorder
                    _date & "','" &    'date
                    IIf(Bank_DT.Rows(i)(5) = "Af", amt, 0) & "','" &    'debit
                    IIf(Bank_DT.Rows(i)(5) = "Bij", amt, 0) & "','" &    'credit
                    Bank_DT.Rows(i)(3) & "','" &    'iban2
                    Bank_DT.Rows(i)(1) & "','" &    'name                             
                    Bank_DT.Rows(i)(4) & "','" &    'code
                    "','" &    'batchid
                    Bank_DT.Rows(i)(7) & " " & Strings.RTrim(Bank_DT.Rows(i)(8)) &   'description
                    "','" & Replace(_exch, ",", ".") & "','" &  'exchange rate 
                    amt2 & "','" & Left(filename, 4) & "." & i.ToString & "','" &
                    filename & "','" & Cur2(cost) & "');" & vbCrLf

                SQLstr2 &= "('" & Left(filename, 4) & "." & i.ToString & "','" & _date & "','Verwerkt','" &
                    Bank_DT.Rows(i)(7) & " " & Bank_DT.Rows(i)(8) & "','Bank','" &
                    IIf(Bank_DT.Rows(i)(5) = "Af", Cur2(-Bank_DT.Rows(i)(6)), Cur2(Bank_DT.Rows(i)(6))) &
                    "','" & nocat & "',0,0,'" & Bank_DT.Rows(i)(2).ToString() & "'),"
            Next

        End If
        'Clipboard.Clear()
        'Clipboard.SetText(SQLstr)
        If SPAS.Chbx_test.Checked Then MsgBox(SQLstr)
        RunSQL(SQLstr, "NULL", "Download_Bank_Transactions/SQLstr")


        SQLstr2 = Strings.Left(SQLstr2, Strings.Len(SQLstr2) - 1) 'remove the last comma
        'Clipboard.SetText(SQLstr2)
        If SPAS.Chbx_test.Checked Then MsgBox(SQLstr2)
        RunSQL(SQLstr2, "NULL", "Download_Bank_Transactions/SQLstr2")
        'MsgBox("Wacht")

        'link journal postings to bank transaction based on temporary link in name/reference field, as bankid
        'is not yet known before the records are inserted. 
        Dim SQLstr3 = "UPDATE journal SET 
                            fk_bank=bank.id, 
                            name='nog te bepalen',
                            iban=bank.iban,
                            amt2=0::money
                       FROM bank 
                       WHERE 
                            bank.fk_journal_name=journal.name AND 
                            journal.name !='nog te bepalen';
                       UPDATE bank SET fk_journal_name='nog te bepalen' WHERE fk_journal_name ilike 'NL%' or fk_journal_name ilike 'CSV_.%';"
        If SPAS.Chbx_test.Checked Then MsgBox(SQLstr3)
        RunSQL(SQLstr3, "NULL", "Download_Bank_Transactions/SQLstr3")
        'Clipboard.Clear()
        'Clipboard.SetText(SQLstr3)
        'Dim sql = QuerySQL("Select sql from query where category ilike 'Transaction' and name='Vervang null door 0/lege string'")
        'RunSQL(sql, "NULL", "Upload_CSV")

    End Sub

    Sub Categorize_Bank_Transactions()
        ' ==========================   P R E P A R A T I O N   =====================
        Dim runstr As String = ""
        Dim sqlupd As String = ""
        Dim sqlins As String = ""
        Dim upsert As String = ""
        Dim delstr As String = ""
        Dim bankid As Integer
        Dim overhead As String = QuerySQL("SELECT value FROM settings WHERE label='overhead'")
        Dim SQLinsert_hdr = "
            INSERT INTO journal
            (amt1,amt2, fk_account,date, status,description,source,fk_relation,fk_bank,name,type,iban) 
            VALUES 
"
        '/////////////// TOEVOEGEN: IBAN AAN SQLINSERT_HDR en SQLINS////////////////////////

        'A ============== CREATE JOURNAL POSTINGS FOR CONTRACTS WITHOUT AUTOMATIC COLLECTION 
        '1) create journal posts for contracts without automated collections
        '[ok]

        Dim nocat As String = QuerySQL("Select value from settings where label='nocat'")
        Dim SQLstrA As String = "
            SELECT 
                b.id, b.date, co.donation/co.term As amtd, co.overhead/co.term As amto, 
                b.description, ac.id As acc, r.id As rel, ta.id As tar, 
                co.donation, co.overhead, co.term, J.amt1, b.credit As bcred, co.name, r.name,r.name_add, b.iban
            FROM bank B 
                LEFT JOIN relation R ON r.iban = b.iban2  
                LEFT JOIN contract CO  ON fk_relation_id = r.id
                LEFT JOIN target TA ON TA.id = CO.fk_target_id
                LEFT JOIN journal J ON fk_bank = b.id
                LEFT JOIN account AC ON f_key = ta.id
            WHERE co.autcol = False 
                AND co.active = True
                AND j.fk_account='" & nocat & "'
                AND j.fk_bank=b.id
                AND amt1 = b.credit 
                AND donation + overhead = credit * term
                AND Position ('extra' IN b.description)='0'
            GROUP BY b.id, r.id, co.id, j.id, ta.id, ac.id;
"


        Collect_data(SQLstrA)

        '2) Create journal posts based on output in dataset (dst)

        For x As Integer = 0 To dst.Tables(0).Rows.Count - 1
            bankid = dst.Tables(0).Rows(x)(0)
            sqlupd &=
                    "UPDATE journal SET     
                        amt1='" & Cur2(dst.Tables(0).Rows(x)(2)) & "',
                        fk_relation='" & dst.Tables(0).Rows(x)(6) & "', 
                        fk_account='" & dst.Tables(0).Rows(x)(5) & "', 
                        name='Contract " & dst.Tables(0).Rows(x)(13) & " " & dst.Tables(0).Rows(x)(14) & "," & Strings.Left(dst.Tables(0).Rows(x)(15), 1) & "',
                        type='Contract',
                        iban='" & dst.Tables(0).Rows(x)(16) & "'
                    WHERE fk_bank='" & bankid & "';"  'donation amount to be booked
            '@@@here journal name must be set with contract or (even better?) source = contract 


            If dst.Tables(0).Rows(x)(3) > 0 Then 'create extra journal post for overhead
                'MsgBox(dst.Tables(0).Rows(x)(3))
                sqlins &=
                        "('" & Cur2(dst.Tables(0).Rows(x)(3)) & "',0,'" & overhead & "','" &
                        Format(CDate(dst.Tables(0).Rows(x)(1)), "yyyy-MM-dd") & "','Verwerkt','" &  'date, status
                        dst.Tables(0).Rows(x)(4) & "','Bank','" & _   'description', source
                        dst.Tables(0).Rows(x)(6) & "','" & _ 'relation id
                        bankid & "','Contract " & dst.Tables(0).Rows(x)(13) & "','Contract','" & dst.Tables(0).Rows(x)(16) & "')," 'bank id
            End If

        Next

        'posting contracts
        If sqlins <> "" Then upsert = SQLinsert_hdr & Strings.Left(sqlins, Strings.Len(sqlins) - 1)
        If sqlupd <> "" Then RunSQL(sqlupd & upsert, "NULL", " Categorize banktransactions A")

        sqlupd = ""
        sqlins = ""
        'upsert = ""

        'B ============== CATEGORIZE BANK TRANSACTIONS BASED ON CODE ==========================
        '[NOK]

        Dim SQLstr2 = "
            SELECT b.id, b.date, b.debit, b.credit,b.description, b.code,
                   (SELECT ac.id FROM account ac WHERE POSITION (b.code IN searchword)> 0)
            FROM bank b
			JOIN journal j ON j.fk_bank=b.id
            WHERE (Select ac.id FROM account ac WHERE POSITION (b.code IN searchword)> 0) is distinct From Null
                   AND b.code!='st'
				   AND j.fk_account='" & nocat & "' 
            GROUP BY b.id
            ORDER BY b.code
            "
        Collect_data(SQLstr2)
        For x As Integer = 0 To dst.Tables(0).Rows.Count - 1
            bankid = dst.Tables(0).Rows(x)(0)
            sqlupd &=
                "UPDATE journal SET fk_account='" & dst.Tables(0).Rows(x)(6) & "' 
                                    WHERE fk_bank='" & bankid & "';"  'category to be booked
        Next


        'C ============== CATEGORIZE BANK TRANSACTIONS BASED ON DESCRIPTION AND SEARCH WORD ====================


        'For sw = 1 To 1


        Dim SQLstr4 = "
            SELECT B.id, ac.id, B.description, ac.searchword 
            FROM bank B
            JOIN journal j ON j.fk_bank=b.id
			JOIN account ac
            On Position(searchword in B.description)>0
	        WHERE j.fk_account='" & nocat & "' 
            GROUP BY b.id, ac.id
            ORDER BY b.id;
            "
        Collect_data(SQLstr4)

        'MsgBox(sw)
        'MsgBox(dst.Tables(0).Rows.Count - 1)
        For x As Integer = 0 To -1 'dst.Tables(0).Rows.Count - 1
            'if there is two candidate categories, these two are both in the output but should be ignored
            'if v(x) = V(x + 1) Then x = x + 2 
            bankid = dst.Tables(0).Rows(x)(0)

            sqlupd &=
                    "UPDATE journal SET 
                    fk_account='" & dst.Tables(0).Rows(x)(6) & "',
                    name='" & dst.Tables(0).Rows(x)(3) & "',
                    date='" & dst.Tables(0).Rows(x)(1) & "  
                    WHERE fk_bank='" & bankid & "';"  'category to be booked
        Next

        'Clipboard.Clear()
        'Clipboard.SetText(sqlupd & SQLinsert_hdr)
        'If sqlins <> "" Then upsert = SQLinsert_hdr & Strings.Left(sqlins, Strings.Len(sqlins) - 1)
        'RunSQL(sqlupd & upsert, "NULL", " Categorize banktransactions A")

        'Next sw



        'D ============== CATEGORIZE AUTOMATIC INCASSO ========================================
        Dim sql = QuerySQL("Select sql from query where category ilike 'Transaction' and name='Match contractincasso'")
        RunSQL(sql, "NULL", "Match contractincasso")

        'E ============== CATEGORIZE EXCASSO JOBS ====================================================
        Collect_data("
        Select (b.debit *-1)- Sum(j.amt1), b.id, j.date, j.name, sum(j.amt2)/(sum(j.amt1)),b.debit, sum(j.amt1), b.iban
                        FROM bank b
                        JOIN journal j ON b.description  ilike '%'||j.name||'%' 
                        WHERE b.code = 'bg'
                        AND j.status = 'Open'
                        AND j.source = 'Uitkering'
                        GROUP BY b.credit, b.id, j.date, j.name
")

        Dim _dat As Date
        Dim dat As String
        Dim calc As Boolean = False
        Dim tr = QuerySQL("Select value from settings where label='tussenrekening_uitk'")
        If IsDBNull(tr) Then
            MsgBox("Tussenrekening is niet geconfigureerd in Instellingen, categorisering wordt afgebroken")
            Exit Sub
        End If

        'If dst.Tables(0).Rows.Count > 0 Then  'excasso available


        For x = 0 To dst.Tables(0).Rows.Count - 1

            'Delete previous standard journal post on bank transaction
            RunSQL("DELETE FROM journal WHERE fk_bank='" & dst.Tables(0).Rows(x)(1) & "'", "NULL",
               "Categorize banktransactions E1") 'remove the journal entry entered originally on importing bank transactions
            'Update journal posts with bankid, set type on contract
            RunSQL("Update Journal 
                    SET status='Verwerkt',fk_bank='" & dst.Tables(0).Rows(x)(1) & "' 
                    WHERE name ilike '%" & dst.Tables(0).Rows(x)(3) & "%'",
                   "NULL", "Categorize banktransactions E2")
            'Set fk_journal_name
            Dim updsql As String = "UPDATE bank set fk_journal_name='" & dst.Tables(0).Rows(x)(3) & "' where id=" & dst.Tables(0).Rows(x)(1)
            RunSQL(updsql, "NULL", "Categorize banktransactions E4")
            MsgBox(updsql)

            If dst.Tables(0).Rows(x)(0) <> 0 Then
                Dim msg = MsgBox("Er is een verschil tussen de bankafschrijving en het uitkeringsformulier" & vbCrLf _
                       & "van " & dst.Tables(0).Rows(x)(0) & " euro. Wilt u doorgaan (het verschil wordt" &
                       vbCrLf & "dan naar een tussenrekening geboekt)?", vbYesNo)
                If msg = vbNo Then Exit Sub
                _dat = dst.Tables(0).Rows(x)(2)
                dat = _dat.Year & "-" & _dat.Month & "-" & _dat.Day
                'create an extra journal item 'nocat' for the difference
                '(amt1,fk_account,date, status,description,source,fk_relation,fk_bank,name,type) 
                sqlins =
                 "('" & Cur2(dst.Tables(0).Rows(x)(0)) & "','" & 'amt1
                 Tbx2Dec(dst.Tables(0).Rows(x)(4)) * Cur2(dst.Tables(0).Rows(x)(0)) & "','" & 'amt2
                 tr & "','" & 'fk_account
                 dat & 'date
                 "','Verwerkt','Verschil tussen bankbetaling en excassoboeking','Bank',null,'" &  'status/descr/source/relat
                 dst.Tables(0).Rows(x)(1) & "','Tussenrek. nav " & 'fk_bank
                 dst.Tables(0).Rows(x)(3) & "','Internal','" & dst.Tables(0).Rows(x)(7) & "');" 'name/type
                RunSQL(SQLinsert_hdr & sqlins, "NULL", "Categorize banktransactions E3")
                calc = True
            End If
            'Clipboard.Clear()
            'Clipboard.SetText(SQLinsert_hdr & sqlins)

        Next x
        'RunSQL("Update Bank set fk_journal_name='Excasso' where id=")
        'End If
        'F ============== CATEGORIZE EXCHANGE RATES ====================================================
        'If MsgBox("Calculate exchange rates?", vbYesNo) = vbYes Then
        If calc Then
            Fill_Cmx_Excasso_Select_Combined()
        End If
        RunSQL("update bank b set fk_journal_name = j.source from journal J where j.fk_bank = b.id and b.fk_journal_name='nog te bepalen' and j.fk_account !='" & nocat & "'", "NULL", "Categorize bank transactions")
        calc = False
        'Calculate_Exchange_Rates()
    End Sub
    Sub Calculate_Exchange_Rates()
        'F ============== CALCULATE EXCHANGE DIFFERENCE ====================================================
        Dim payment_rate As Decimal = -1
        Dim euro_tegenwaarde, diff, trans_cost As Decimal
        Dim sql1 As String = ""
        Dim sql2 As String = ""
        Dim iban_old As String = "xxx"
        Dim cod, iban, acc_euro_tegenwaarde, acc_wisselkoers_verschil, bank_transactie_kosten, bank_kosten, dat As String
        Dim id As Integer
        Dim _dat As Date
        Dim transtype, bname As String

        '1 retrieve account names
        Collect_data("SELECT * FROM settings")
        For x As Integer = 0 To dst.Tables(0).Rows.Count - 1
            If dst.Tables(0).Rows(x)(0) = "eurotegenwaarde" Then acc_euro_tegenwaarde = dst.Tables(0).Rows(x)(1)
            If dst.Tables(0).Rows(x)(0) = "wisselkoersverschil" Then acc_wisselkoers_verschil = dst.Tables(0).Rows(x)(1)
            If dst.Tables(0).Rows(x)(0) = "bank_transactie_kosten" Then bank_transactie_kosten = dst.Tables(0).Rows(x)(1)
            If dst.Tables(0).Rows(x)(0) = "bank_kosten" Then bank_kosten = dst.Tables(0).Rows(x)(1)
        Next x

        Collect_data("
        Select b.iban, b.iban2, b.Date, b.debit, b.credit, 
                       b.exch_rate, b.amt_cur, b.cost, b.id, b.fk_journal_name, 
                    (
				select amt2/amt1 from journal j11
				left join cp cp1 on cp1.id = substring(j11.cpinfo,1,2)::integer
				left join bankacc ba1 on ba1.id = cp1.fk_bankacc_id 
				 where source = 'Uitkering'
				 and date <= b.date
				 and ba1.accountno = b.iban
				 and j11.amt1 < 0::money
				 and b.code='GM'
				 order by j11.date desc
				 limit 1
                ), b.code, b.description
                FROM bank b
                LEFT join bankacc a ON a.accountno=b.iban
                WHERE a.expense='True'
                ORDER BY b.iban, b.date asc
")

        For x As Integer = 0 To dst.Tables(0).Rows.Count - 1
            iban = Trim(dst.Tables(0).Rows(x)(0))

            'bepaal ING rekening ------------------------------------------------------
            If iban <> iban_old Then
                payment_rate = 1 '@@@ gaat fout als eerste transactie van het jaar geen storting is
                iban_old = iban
            End If

            'onthoud wisselkoers
            If Not IsDBNull(dst.Tables(0).Rows(x)(11)) Then cod = Trim(dst.Tables(0).Rows(x)(11))
            id = dst.Tables(0).Rows(x)(8)
            _dat = dst.Tables(0).Rows(x)(2)
            dat = _dat.Year & "-" & _dat.Month & "-" & _dat.Day

            If cod = "GM" Then
                If Not IsDBNull(dst.Tables(0).Rows(x)(10)) Then
                    payment_rate = Decimal.Round(dst.Tables(0).Rows(x)(10), 2)
                Else payment_rate = -1
                End If
            End If

            If IsDBNull(dst.Tables(0).Rows(x)(9)) Then dst.Tables(0).Rows(x)(9) = "nog te bepalen"
            If dst.Tables(0).Rows(x)(9) = "" Then dst.Tables(0).Rows(x)(9) = "nog te bepalen"
            If IsDBNull(dst.Tables(0).Rows(x)(7)) Then dst.Tables(0).Rows(x)(7) = 0
            If IsDBNull(dst.Tables(0).Rows(x)(6)) Then dst.Tables(0).Rows(x)(6) = 0
            If dst.Tables(0).Rows(x)(9) = "nog te bepalen" Then

                Select Case cod
                    Case "OV"
                        transtype = IIf(InStr(dst.Tables(0).Rows(x)(12), "Excasso-") > 0, "Bank", "Internal")
                        bname = IIf(transtype = "Bank", "Overschrijving uitkering", "Saldosteun")
                        'If IsDBNull(dst.Tables(0).Rows(x)(10)) Then payment_rate = 1 Else payment_rate = Decimal.Round(dst.Tables(0).Rows(x)(10), 2)
                        'boek euro tegenwaarde

                        sql1 &= vbCrLf & "UPDATE journal SET fk_account='" & acc_euro_tegenwaarde & "', name='" & bname & "', type='" & transtype & "' WHERE fk_bank='" & id & "';"
                        'vervang 'nog te bepalen met 'euro tegenwaarde
                        sql1 &= vbCrLf & "UPDATE bank SET fk_journal_name='Intrabank boeking' WHERE id='" & id & "';"

                    Case "GM"
                        'bereken euro tegenwaarde
                        If payment_rate < -1 Then payment_rate = InputBox("De wisselkoers van de voorafgaande uitkeringslijst (" & iban & ") kon niet bepaald worden. Geef deze s.v.p. op (MLD per €)")
                        trans_cost = -Decimal.Round(dst.Tables(0).Rows(x)(7), 2)
                        If Decimal.Round(dst.Tables(0).Rows(x)(6)) <> 0 Then
                            euro_tegenwaarde = -Decimal.Round(dst.Tables(0).Rows(x)(6) / payment_rate, 2) '+ trans_cost
                        Else
                            euro_tegenwaarde = -Decimal.Round(dst.Tables(0).Rows(x)(3), 2) - trans_cost
                        End If

                        diff = -euro_tegenwaarde - Decimal.Round(dst.Tables(0).Rows(x)(3), 2) - trans_cost

                        'boek euro tegenwaarde

                        sql1 &= vbCrLf & " UPDATE journal SET fk_account='" & acc_euro_tegenwaarde & "',name='Uitkeringstransactie', 
                                           type='Bank', iban='" & Trim(iban) & "',amt1='" & Cur2(euro_tegenwaarde) & "'::MONEY WHERE fk_bank='" & id & "';"
                        'boek wisselkoersverschil
                        If diff <> 0 Then
                            sql2 &= vbCrLf & "('" & Cur2(diff) & "'::MONEY,'" & acc_wisselkoers_verschil & "','" & dat &
                            "','Verwerkt','Wisselkoersverschil','Bank','" & id & "','Wisselkoersverschil','Bank','" & Trim(iban) & "'),"
                        End If
                        'boek transactiekosten
                        If trans_cost <> 0 Then
                            sql2 &= vbCrLf & "('" & Cur2(trans_cost) & "'::MONEY,'" &
                            bank_transactie_kosten & "','" & dat &
                            "','Verwerkt','Banktransactiekosten','Bank','" & id & "','Transactiekosten','Bank','" & Trim(iban) & "'),"
                        End If
                        'vervang 'nog te bepalen met 'euro tegenwaarde/wisselkoers'
                        sql1 &= vbCrLf & "UPDATE bank SET fk_journal_name='Lokale opname' WHERE id='" & id & "';"
                    Case "DV", "VZ"
                        'boek bankkosten
                        sql1 &= vbCrLf & "UPDATE journal SET fk_account='" & bank_kosten & "', name='Boeking bankkosten',type='Bank' WHERE fk_bank='" & id & "';"
                        'vervang 'nog te bepalen met 'euro tegenwaarde/wisselkoers'
                        sql1 &= vbCrLf & "UPDATE bank SET fk_journal_name='Bankkosten' WHERE id='" & id & "';"

                End Select


            End If
        Next x

        If sql2 <> "" Then
            sql2 = "INSERT INTO journal (amt1,fk_account,date, status,description,source,fk_bank,name,type,iban) VALUES " &
                Strings.Left(sql2, Strings.Len(sql2) - 1) & ";" 'remove the last comma
        End If
        Clipboard.Clear()
        Clipboard.SetText(sql1 & vbCrLf & sql2)  'sql1 & vbCrLf & 

        If sql1 <> "" Or sql2 <> "" Then
            RunSQL(sql1 & vbCrLf & sql2, "NULL", "Calculate_Exchange_Rates")
        Else
            'MsgBox("No transactions categorized.")
        End If

    End Sub


    Sub Mark_rows_Dgv_Bank()
        For x As Integer = 0 To SPAS.Dgv_Bank.Rows.Count - 1

            Dim cnt As Integer = SPAS.Dgv_Bank.Rows(x).Cells(17).Value
            If cnt > 0 Then
                SPAS.Dgv_Bank.Rows(x).DefaultCellStyle.ForeColor = Color.DarkRed
                'SPAS.Btn_Bank_Split.Enabled = True
            Else
                SPAS.Dgv_Bank.Rows(x).DefaultCellStyle.ForeColor = Color.DarkGreen
                'SPAS.Btn_Bank_Split.Enabled = False
            End If

        Next

    End Sub
    Sub Fill_Journals_by_bank(ByVal journal_name As Integer)

        'If Strings.Left(journal_name, 1) = "0" Then Exit Sub

        Dim SQLstr = "SELECT account.id, account.name, journal.amt1, journal.type FROM journal
                     JOIN account ON journal.fk_account = account.id
                     JOIN bank ON bank.id = journal.fk_bank
                     WHERE bank.id =" & journal_name

        Collect_bankdata(SQLstr)
        SPAS.ToClipboard(SQLstr, True)

        Dim Amt_In = CDec(SPAS.Dgv_Bank.SelectedCells(4).Value)
        Dim cod As String = SPAS.Dgv_Bank.SelectedCells(6).Value
        Dim cnt As Integer = SPAS.Dgv_Bank.SelectedCells(17).Value

        SPAS.Btn_Bank_Type.Visible = (Trim(cod) = "cb")

        SPAS.Dgv_Bank_Account.DataSource = dstbank.Tables(0)
        SPAS.Btn_Bank_Type.Text = ""
        'SPAS.Btn_Bank_Split.Enabled = cnt > 0
        If Trim(cod) = "cb" Then
            SPAS.Pan_Bank_jtype.Visible = True
            If Not IsDBNull(dstbank.Tables(0).Rows(0)(3)) Then
                SPAS.Btn_Bank_Type.Text = Left(dstbank.Tables(0).Rows(0)(3), 1)
            End If
            If SPAS.Btn_Bank_Type.Text = "" Then SPAS.Btn_Bank_Type.Text = "?"
            Dim jtype = dstbank.Tables(0).Rows(0)(3)
            SPAS.Rbn_Bank_jtype_con.Checked = False
            SPAS.Rbn_Bank_jtype_ext.Checked = False
            SPAS.Rbn_Bank_jtype_int.Checked = False
            SPAS.Btn_Bank_Add_Journal.Enabled = False
            If Not IsDBNull(jtype) Then
                Select Case Trim(jtype)
                    Case "Contract"
                        SPAS.Rbn_Bank_jtype_con.Checked = True
                        SPAS.Btn_Bank_Add_Journal.Enabled = True
                    Case "Extra"
                        SPAS.Rbn_Bank_jtype_ext.Checked = True
                        SPAS.Btn_Bank_Add_Journal.Enabled = True
                    Case "Internal"
                        SPAS.Rbn_Bank_jtype_int.Checked = True
                        SPAS.Btn_Bank_Add_Journal.Enabled = True
                End Select
            End If
        Else

            SPAS.Pan_Bank_jtype.Visible = False

        End If


    End Sub

    Sub Calculate_Bank_Balance()
        If Strings.InStr(SPAS.Cmx_Bank_bankacc.Text, "NL") = 0 Then Exit Sub

        Dim balance As Decimal = QuerySQL("
         select case when sum(credit)-sum(debit)::money isnull then 0::money else sum(credit-debit)::money end 
  		from bank ba WHERE iban = '" & Strings.Right(SPAS.Cmx_Bank_bankacc.Text, 18) & "' 
")
        SPAS.Lbl_Bank_Saldo.Text = Format(balance, "#,##0.00")

    End Sub


    Sub Update_Category_Status()
        Dim currow As Integer = SPAS.Dgv_Bank.SelectedCells(3).RowIndex

        SPAS.Dgv_Bank.Rows(currow).Cells(17).Value = 0
        SPAS.Dgv_Bank.Rows(currow).DefaultCellStyle.ForeColor = Color.DarkGreen

        For x = 0 To SPAS.Dgv_Bank_Account.Rows.Count - 1
            If SPAS.Dgv_Bank_Account.Rows(x).Cells(0).Value = nocat And SPAS.Dgv_Bank_Account.Rows(x).Cells(2).Value <> 0 Then
                SPAS.Dgv_Bank.Rows(currow).Cells(17).Value = 1
                SPAS.Dgv_Bank.Rows(currow).DefaultCellStyle.ForeColor = Color.DarkRed
                Exit For

            End If
        Next x

    End Sub

    Sub Obsoleet()
        ' automatische incasso
        Dim sqlupd As String = ""
        Dim sqlins As String = ""
        Dim SQLinsert_hdr As String = ""

        Collect_data("Select b.credit - Sum(j.amt1), b.id, j.date, j.name, b.iban 
                        FROM bank b
                        JOIN journal j ON b.description = j.name 
                        WHERE b.code = 'ei'
                        AND j.status = 'Open'
                        AND j.source = 'Incasso'
                        GROUP BY b.credit, b.id, j.date, j.name")


        For x As Integer = 0 To dst.Tables(0).Rows.Count - 1
            If dst.Tables(0).Rows.Count > 0 Then '.Rows(0)(0)) Then 'bank is automated incasso?
                'Delete previous standard journal post on bank transaction
                RunSQL("DELETE FROM journal WHERE fk_bank='" & dst.Tables(0).Rows(x)(1) & "'", "NULL",
                                   "Categorize banktransactions D1")
                'Update journal posts with bankid, set type on contract
                RunSQL("Update Journal 
                                    SET status='Verwerkt',type='Contract',fk_bank='" & dst.Tables(0).Rows(x)(1) & "' 
                                    ,iban='" & dst.Tables(0).Rows(x)(4) & "' 
                                    WHERE name='" & dst.Tables(0).Rows(x)(3) & "'", "NULL", "Categorize banktransactions D2")
                Dim updsql As String = "UPDATE bank set fk_journal_name='" & dst.Tables(0).Rows(x)(3) & "' where id=" & dst.Tables(0).Rows(x)(1)
                RunSQL(updsql, "NULL", "Categorize banktransactions D4")


                If dst.Tables(0).Rows(x)(0) <> 0 Then
                    'create an extra journal item 'nocat' for the difference
                    '(amt1,fk_account,date, status,description,source,fk_relation,fk_bank,name,type) 
                    sqlins =
                                 "('" & dst.Tables(0).Rows(x)(0) & "',0,'" & 'amt1
                                 nocat & "','" & 'fk_account
                                 dst.Tables(0).Rows(x)(2) & 'date
                                 "','Verwerkt','Verschil tussen bankincasso en incassoboeking','Incasso','','" &  'status/descr/source/relat
                                 dst.Tables(0).Rows(x)(1) & "','" & 'fk_bank
                                 dst.Tables(0).Rows(x)(3) & "','Contract','" & dst.Tables(0).Rows(x)(16) & "');" 'name/type
                    RunSQL(SQLinsert_hdr & sqlins, "NULL", "Categorize banktransactions D3-" & x)
                End If
            End If
        Next x


    End Sub

End Module
