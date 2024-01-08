
Imports System.Windows
Imports Npgsql
'Imports PdfSharp.Pdf
Imports System.IO
Imports System.Diagnostics.Tracing
Imports Microsoft.EntityFrameworkCore.Update.Internal
Imports System.ComponentModel
Imports Microsoft.EntityFrameworkCore.Metadata.Internal
Imports System.Dynamic

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
        'Categorize_Bank_Transactions()
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
            If Strings.Right(f.Name, 4) = ".csv" Then Upload_CSV(SelectFolder.SelectedPath & "\" & f.Name)
        Next
        Categorize_Bank_Transactions(True, True, True, True, True, True, True)
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


    End Sub

    Sub Categorize_Bank_Transactions(ByVal contr As Boolean, uitk As Boolean, inc As Boolean, bcode As Boolean, omschr As Boolean, extrag As Boolean, ing As Boolean)

        Dim nocat As String = QuerySQL("Select value from settings where label='nocat'")


        'controle op null toevoegen

        If inc Then RunQuery("Categoriseer contractincasso")


        If uitk Then
            RunQuery("Categoriseer uitkering")
            Fill_Cmx_Excasso_Select_Combined()
        End If
        If contr Then RunQuery("Categoriseer contractbetaling")
        If bcode Then RunQuery("Categoriseer obv bankcode")
        If omschr Then RunQuery("Categoriseer obv omschrijving")
        If extrag Then RunQuery("Categoriseer extra gift")
        If ing Then RunQuery("Categoriseer ingbank")


    End Sub



    Public Sub Mark_rows_Dgv_Bank()
        For x As Integer = 0 To SPAS.Dgv_Bank.Rows.Count - 1
            Dim cnt As Integer = SPAS.Dgv_Bank.Rows(x).Cells(17).Value
            SPAS.Dgv_Bank.Rows(x).DefaultCellStyle.ForeColor = IIf(cnt > 0, Color.DarkRed, Color.DarkGreen)
            If SPAS.Dgv_Bank.Rows(x).Cells(12).Value = "Auto-cat" Then SPAS.Dgv_Bank.Rows(x).DefaultCellStyle.ForeColor = Color.DarkGoldenrod
        Next

    End Sub
    Sub Fill_Journals_by_bank(ByVal journal_name As Integer)

        'If Strings.Left(journal_name, 1) = "0" Then Exit Sub

        Dim SQLstr = "SELECT account.id, account.name, journal.amt1, journal.type FROM journal
                     JOIN account ON journal.fk_account = account.id
                     JOIN bank ON bank.id = journal.fk_bank
                     WHERE bank.id =" & journal_name

        Collect_bankdata(SQLstr)
        ToClipboard(SQLstr, True)

        Dim Amt_In = CDec(SPAS.Dgv_Bank.SelectedCells(4).Value)
        Dim cod As String = SPAS.Dgv_Bank.SelectedCells(6).Value
        Dim cnt As Integer = SPAS.Dgv_Bank.SelectedCells(17).Value

        SPAS.Dgv_Bank_Account.DataSource = dstbank.Tables(0)

        If Trim(cod) = "cb" Then
            SPAS.Pan_Bank_jtype.Visible = True
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



End Module
