
Imports System.Windows
Imports Npgsql
'Imports PdfSharp.Pdf
Imports System.IO
Imports System.Text.RegularExpressions
Imports Microsoft.VisualBasic.FileIO
Imports System.Diagnostics.Tracing
Imports Microsoft.EntityFrameworkCore.Update.Internal
Imports System.ComponentModel
Imports Microsoft.EntityFrameworkCore.Metadata.Internal
Imports System.Dynamic
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports PdfSharp.Pdf.Content.Objects
Imports System.Linq ' <-- REQUIRED FOR DB MAPPING SPLIT

Module bank
    Public Class BankTransactionLoader
        Private ReadOnly _connectionString As String

        ' Helper class to store Relation data in memory
        Private Class RelationInfo
            Public Property Id As Integer
            Public Property Name As String
            Public Property NameAdd As String
        End Class

        Public Sub New(connectionString As String)
            _connectionString = connectionString
        End Sub

        ''' <summary>
        ''' Loads bank transactions from a CSV file into the PostgreSQL 'bank' table.
        ''' </summary>
        Public Function Load_Bank_Transactions(csvFilePath As String, ByRef statusMessage As String) As Boolean
            Try
                Dim lastDbSeqOrder As Integer? = GetLatestSeqOrder()
                Dim newRecords As New List(Of Dictionary(Of String, Object))
                Dim fileMinNewSeqOrder As Integer = Integer.MaxValue

                Dim detectedDelimiter As String = DetectDelimiter(csvFilePath)

                Using parser As New TextFieldParser(csvFilePath)
                    parser.TextFieldType = FieldType.Delimited
                    parser.SetDelimiters(detectedDelimiter)
                    parser.HasFieldsEnclosedInQuotes = True

                    If parser.EndOfData Then
                        statusMessage = "Het CSV-bestand is leeg."
                        Return False
                    End If

                    Dim headers As String() = parser.ReadFields()

                    ' 1. Detect the Bank from the headers
                    Dim bankCode As String = DetectBankFromHeaders(headers)
                    If bankCode = "UNKNOWN" Then
                        statusMessage = "Kan de bank niet identificeren op basis van de CSV kolomnamen."
                        Return False
                    End If

                    ' 2. Load mapping and relations from the PostgreSQL database
                    Dim columnMapping As Dictionary(Of String, String())
                    Dim relationsDict As Dictionary(Of String, RelationInfo)

                    Using conn As New NpgsqlConnection(_connectionString)
                        conn.Open()
                        columnMapping = GetColumnMappingFromDB(bankCode, conn)
                        relationsDict = GetRelationsFromDB(conn) ' <-- Load all relations once
                    End Using

                    ' 3. Map header positions for this specific file
                    Dim headerIndexMap As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
                    For i As Integer = 0 To headers.Length - 1
                        headerIndexMap(headers(i).Trim()) = i
                    Next

                    Dim seqColName As String = columnMapping("seqorder")(0)
                    If Not headerIndexMap.ContainsKey(seqColName) Then
                        statusMessage = $"De vereiste CSV-kolom '{seqColName}' is niet gevonden. Gevonden kolommen: {String.Join(", ", headers)}"
                        Return False
                    End If

                    While Not parser.EndOfData
                        Dim currentRow As String() = parser.ReadFields()
                        Dim record As New Dictionary(Of String, Object)

                        Dim currentSeqOrder As Integer
                        Dim seqStr As String = currentRow(headerIndexMap(seqColName))

                        If Not Integer.TryParse(seqStr, currentSeqOrder) Then
                            Continue While
                        End If

                        If lastDbSeqOrder.HasValue AndAlso currentSeqOrder <= lastDbSeqOrder.Value Then
                            Continue While
                        End If

                        If currentSeqOrder < fileMinNewSeqOrder Then
                            fileMinNewSeqOrder = currentSeqOrder
                        End If

                        ' Map columns dynamically, concatenate if needed, and convert types
                        For Each kvp In columnMapping
                            Dim dbColName As String = kvp.Key.ToLower()
                            Dim csvColNames As String() = kvp.Value
                            Dim combinedValue As String = ""

                            ' 1 & 2. Check for Constant OR Extract and Concatenate CSV columns
                            If csvColNames.Length > 0 AndAlso csvColNames(0).StartsWith("CONSTANT:", StringComparison.OrdinalIgnoreCase) Then
                                combinedValue = csvColNames(0).Substring(9)
                            Else
                                For Each csvColName In csvColNames
                                    If headerIndexMap.ContainsKey(csvColName) AndAlso headerIndexMap(csvColName) < currentRow.Length Then
                                        Dim val As String = currentRow(headerIndexMap(csvColName)).Trim()
                                        If Not String.IsNullOrWhiteSpace(val) Then
                                            combinedValue &= val & " "
                                        End If
                                    End If
                                Next
                                combinedValue = Regex.Replace(combinedValue.Trim(), "\s+", " ")
                            End If

                            ' 3. Process the combined value
                            If String.IsNullOrWhiteSpace(combinedValue) Then
                                If dbColName = "debit" OrElse dbColName = "credit" Then
                                    record(dbColName) = 0D
                                Else
                                    record(dbColName) = DBNull.Value
                                End If
                            Else
                                Try
                                    ' --- Special Handling for Dutch Bank Amounts ("Bedrag") ---
                                    If csvColNames.Contains("Bedrag", StringComparer.OrdinalIgnoreCase) Then
                                        Dim cleanNum As String = combinedValue.Replace(".", "").Replace(",", ".")
                                        Dim amount As Decimal = Convert.ToDecimal(cleanNum, Globalization.CultureInfo.InvariantCulture)

                                        If amount < 0 Then
                                            If dbColName = "debit" Then record(dbColName) = Math.Abs(amount)
                                            If dbColName = "credit" Then record(dbColName) = 0D
                                        ElseIf amount > 0 Then
                                            If dbColName = "credit" Then record(dbColName) = amount
                                            If dbColName = "debit" Then record(dbColName) = 0D
                                        Else
                                            If dbColName = "debit" OrElse dbColName = "credit" Then record(dbColName) = 0D
                                        End If
                                        Continue For
                                    End If

                                    ' Convert string to the correct .NET type
                                    Select Case dbColName
                                        Case "seqorder", "amt_cur"
                                            record(dbColName) = Convert.ToInt32(combinedValue)
                                        Case "date"
                                            record(dbColName) = Convert.ToDateTime(combinedValue)
                                        Case "debit", "credit", "exch_rate", "cost"
                                            Dim cleanNum As String = combinedValue.Replace(".", "").Replace(",", ".")
                                            record(dbColName) = Convert.ToDecimal(cleanNum, Globalization.CultureInfo.InvariantCulture)
                                        Case Else
                                            record(dbColName) = combinedValue
                                    End Select
                                Catch ex As Exception
                                    If dbColName = "debit" OrElse dbColName = "credit" Then
                                        record(dbColName) = 0D
                                    Else
                                        record(dbColName) = DBNull.Value
                                    End If
                                End Try
                            End If
                        Next

                        ' =====================================================================
                        ' RELATION LOOKUP LOGIC
                        ' =====================================================================
                        ' We check iban2 (Tegenrekening) first, if not available we check iban.
                        Dim lookupIban As String = ""
                        If record.ContainsKey("iban2") AndAlso Not IsDBNull(record("iban2")) Then
                            lookupIban = record("iban2").ToString()
                        ElseIf record.ContainsKey("iban") AndAlso Not IsDBNull(record("iban")) Then
                            lookupIban = record("iban").ToString()
                        End If

                        If Not String.IsNullOrWhiteSpace(lookupIban) Then
                            ' Clean IBAN (remove spaces, uppercase) for accurate matching
                            lookupIban = lookupIban.Replace(" ", "").ToUpper()

                            If relationsDict.ContainsKey(lookupIban) Then
                                Dim rel = relationsDict(lookupIban)

                                ' 1) Format the name (relation.name + ", " + relation.name_add)
                                Dim newName As String = rel.Name
                                If Not String.IsNullOrWhiteSpace(rel.NameAdd) Then
                                    newName &= ", " & rel.NameAdd
                                End If

                                ' Overwrite the bank.name with the formatted relation name
                                record("name") = newName

                                ' 2) Store the relation ID for later use
                                Dim rel_id As String = rel.Id.ToString()
                                record("rel_id") = rel_id ' <--- This is now available in your record dictionary
                            End If
                        End If
                        ' =====================================================================

                        record("filename") = Path.GetFileName(csvFilePath)
                        newRecords.Add(record)
                    End While
                End Using

                If newRecords.Count = 0 Then
                    statusMessage = "Er zijn geen nieuwe transacties gevonden om te importeren (alle 'seqorder' nummers in dit bestand bestaan al in de database)."
                    Return False
                End If

                If lastDbSeqOrder.HasValue AndAlso fileMinNewSeqOrder > (lastDbSeqOrder.Value + 1) Then
                    statusMessage = $"Import geannuleerd: Gat in de volgorde ontdekt. Laatste 'seqorder' in database is {lastDbSeqOrder.Value}, maar nieuwe data in CSV begint bij {fileMinNewSeqOrder}."
                    Return False
                End If

                InsertTransactions(newRecords)
                statusMessage = $"{newRecords.Count} nieuwe transactie(s) succesvol geïmporteerd."
                Return True

            Catch ex As Exception
                statusMessage = $"Er is een technische fout opgetreden tijdens het inlezen: {ex.Message}"
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Fetches all relations with an IBAN from the database into memory for quick lookup.
        ''' </summary>
        Private Function GetRelationsFromDB(conn As NpgsqlConnection) As Dictionary(Of String, RelationInfo)
            Dim dict As New Dictionary(Of String, RelationInfo)(StringComparer.OrdinalIgnoreCase)

            ' Only pull records where the iban actually has characters
            Dim sql As String = "SELECT id, name, name_add, iban FROM public.relation WHERE iban IS NOT NULL AND iban <> ''"

            Using cmd As New NpgsqlCommand(sql, conn)
                Using reader As NpgsqlDataReader = cmd.ExecuteReader()
                    While reader.Read()
                        ' Strip spaces and uppercase to ensure matching works perfectly
                        Dim dbIban As String = reader("iban").ToString().Replace(" ", "").ToUpper()

                        Dim info As New RelationInfo With {
                        .Id = Convert.ToInt32(reader("id")),
                        .Name = If(IsDBNull(reader("name")), "", reader("name").ToString().Trim()),
                        .NameAdd = If(IsDBNull(reader("name_add")), "", reader("name_add").ToString().Trim())
                    }

                        ' Add or update the dictionary
                        dict(dbIban) = info
                    End While
                End Using
            End Using

            Return dict
        End Function

        Private Function DetectBankFromHeaders(headers As String()) As String
            Dim headerLine As String = String.Join("|", headers).ToLower()

            If headerLine.Contains("volgnr") AndAlso headerLine.Contains("reden retour") Then
                Return "RABO"
            ElseIf headerLine.Contains("af bij") AndAlso headerLine.Contains("mededelingen") Then
                Return "INGB"
            ElseIf headerLine.Contains("muntsoort") AndAlso headerLine.Contains("tegenrekening") Then
                Return "ABNA"
            End If

            Return "UNKNOWN"
        End Function

        Private Function GetColumnMappingFromDB(bankCode As String, conn As NpgsqlConnection) As Dictionary(Of String, String())
            Dim mapping As New Dictionary(Of String, String())(StringComparer.OrdinalIgnoreCase)

            Dim sql As String = "SELECT db_column, csv_headers FROM public.bank_mappings WHERE bank_code = @bank"

            Using cmd As New NpgsqlCommand(sql, conn)
                cmd.Parameters.AddWithValue("@bank", bankCode)

                Using reader As NpgsqlDataReader = cmd.ExecuteReader()
                    While reader.Read()
                        Dim dbCol As String = reader.GetString(0)
                        Dim rawHeaders As String = reader.GetString(1)
                        Dim csvCols As String() = rawHeaders.Split(","c).Select(Function(s) s.Trim()).ToArray()
                        mapping(dbCol) = csvCols
                    End While
                End Using
            End Using

            If mapping.Count = 0 Then
                Throw New Exception($"Geen mapping gevonden in de database voor bank code: {bankCode}")
            End If

            Return mapping
        End Function

        Private Function DetectDelimiter(csvFilePath As String) As String
            Dim possibleDelimiters() As String = {";", ",", "|", vbTab}
            Dim bestDelimiter As String = ","
            Dim maxFieldCount As Integer = 0

            For Each delim In possibleDelimiters
                Using parser As New TextFieldParser(csvFilePath)
                    parser.TextFieldType = FieldType.Delimited
                    parser.SetDelimiters(delim)
                    parser.HasFieldsEnclosedInQuotes = True

                    Try
                        If Not parser.EndOfData Then
                            Dim fields As String() = parser.ReadFields()
                            If fields IsNot Nothing AndAlso fields.Length > maxFieldCount Then
                                maxFieldCount = fields.Length
                                bestDelimiter = delim
                            End If
                        End If
                    Catch ex As Exception
                    End Try
                End Using
            Next
            Return bestDelimiter
        End Function

        Private Function GetLatestSeqOrder() As Integer?
            Using conn As New NpgsqlConnection(_connectionString)
                conn.Open()
                Using cmd As New NpgsqlCommand("SELECT MAX(seqorder) FROM public.bank", conn)
                    Dim result = cmd.ExecuteScalar()
                    If result IsNot Nothing AndAlso Not DBNull.Value.Equals(result) Then
                        Return Convert.ToInt32(result)
                    End If
                End Using
            End Using
            Return Nothing
        End Function

        Private Sub InsertTransactions(records As List(Of Dictionary(Of String, Object)))
            Using conn As New NpgsqlConnection(_connectionString)
                conn.Open()
                Using trans = conn.BeginTransaction()
                    For Each rec In records
                        Dim columns As New List(Of String)
                        Dim parameters As New List(Of String)
                        Dim cmd As New NpgsqlCommand() With {
                        .Connection = conn,
                        .Transaction = trans
                    }

                        For Each kvp In rec
                            ' SKIP inserting "rel_id" into the public.bank table.
                            ' Since "rel_id" belongs to the future "journal" table, trying to insert it here will cause a PostgreSQL error.
                            If kvp.Key.Equals("rel_id", StringComparison.OrdinalIgnoreCase) Then Continue For

                            columns.Add($"""{kvp.Key}""")
                            parameters.Add($"@{kvp.Key}")
                            cmd.Parameters.AddWithValue($"@{kvp.Key}", kvp.Value)
                        Next

                        ' Insert into public.bank
                        cmd.CommandText = $"INSERT INTO public.bank ({String.Join(", ", columns)}) VALUES ({String.Join(", ", parameters)})"
                        cmd.ExecuteNonQuery()

                        ' Note: When you are ready to insert into the "journal" table, 
                        ' you can retrieve rec("rel_id") here and execute a second INSERT statement inside this same transaction loop!
                    Next
                    trans.Commit()
                End Using
            End Using
        End Sub
    End Class

    Sub Download_Bank_Transactions()
        Dim csv As String = ""

        ' Get CSV
        SPAS.OpenFileDialog1.Title = "Selecteer een bankafschrift"
        SPAS.OpenFileDialog1.FileName = ""
        SPAS.OpenFileDialog1.InitialDirectory = "" ' My.Settings._bankpath
        SPAS.OpenFileDialog1.Filter = "Bank bestanden|*.csv"

        If SPAS.OpenFileDialog1.ShowDialog() = DialogResult.OK Then
            csv = SPAS.OpenFileDialog1.FileName
        End If

        If csv = "" Then Exit Sub

        ' Use your global connection_string variable instead of the dummy string
        Dim loader As New BankTransactionLoader(connect_string)

        Dim statusMsg As String = ""

        ' Pass the statusMsg variable so the function can fill it
        Dim isSuccess As Boolean = loader.Load_Bank_Transactions(csv, statusMsg)

        If isSuccess Then
            ' Shows the success message (e.g. "15 nieuwe transactie(s) succesvol geïmporteerd.")
            MessageBox.Show(statusMsg, "Import Geslaagd", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            ' Shows the exact reason it failed or found nothing
            Clipboard.Clear()
            Clipboard.SetText(statusMsg)
            MessageBox.Show(statusMsg, "Import Geannuleerd / Mislukt", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End If

        'load journal transactions (with category' niet toegewezen) to maintain bank-journal consistency

        Dim sqlstr As String = "
        INSERT INTO Public.journal (date, amt1, description, source, fk_account, fk_bank, iban, status, name,fk_relation)
        Select b.date, COALESCE(b.credit, 0::money) - COALESCE(b.debit, 0::money) AS amt1,'niet toegewezen','Bank'," & nocat & ", b.id, b.iban, 'Verwerkt', 'nog te bepalen',
        (SELECT id from public.relation r where b.iban2 = r.iban)
        FROM public.bank b WHERE Not EXISTS (SELECT 1 FROM public.journal j WHERE j.fk_bank = b.id);"
        RunSQL(sqlstr, "NULL", "BankTransactionLoader")


        Fill_bank_transactions("Download_Bank_Transactions", Nothing)
        Categorize_Bank_Transactions(True, True, True, True, True, True, True)
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
            'If Strings.Right(f.Name, 4) = ".csv" Then Upload_CSV(SelectFolder.SelectedPath & "\" & f.Name)
        Next
        Categorize_Bank_Transactions(True, True, True, True, True, True, True)
        Fill_bank_transactions("Load_Bank_csv_from_folder", Nothing)

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
        'If ing Then RunQuery("Categoriseer ingbank")


    End Sub


    Sub Fill_Journals_by_bank(ByVal journal_name As Integer)

        SPAS.isManualChange = False
        'If Strings.Left(journal_name, 1) = "0" Then Exit Sub

        Dim SQLstr = "SELECT a.id, a.name As Accountnaam, j.amt1 As Bedrag, j.type As Type, j.source As Bron FROM journal j
                     JOIN account a ON j.fk_account = a.id
                     JOIN bank b ON b.id = j.fk_bank
                     WHERE b.id =" & journal_name


        SPAS.Prepare_Datagridview(SPAS.Dgv_Bank_Account, SQLstr, {"HZ010", "TZ200", "NB075", "HZ040", "HZ040"})

        Dim cod As String = SPAS.Dgv_Bank.SelectedCells(6).Value

        'Binnen een banktransacties hebben alle journaalposten hetzelfde type
        Dim jtype = SPAS.Dgv_Bank_Account.Rows(0).Cells(3).Value

        'SPAS.Dgv_Bank_Account.DataSource = bankdata

        If Trim(cod) = "cb" Then
            SPAS.Pan_Bank_jtype.Visible = True
            'Dim jtype = bankdata.Rows(0)(3)
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

        SPAS.isManualChange = True
    End Sub

    Sub Calculate_Bank_Balance()
        If Strings.InStr(SPAS.Cmx_Bank_bankacc.Text, "NL") = 0 Then Exit Sub

        Dim balance As Decimal = QuerySQL($"
         select case when sum(credit)-sum(debit)::money isnull then 0::money else sum(credit-debit)::money end 
  		from bank ba WHERE iban = '{Strings.Right(SPAS.Cmx_Bank_bankacc.Text, 18)}'")
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



    Sub Calculate_Total_Booked(sender)

        Dim Amt_In = CDec(SPAS.Dgv_Bank.SelectedCells(4).Value)
        Dim Amt_Out = CDec(SPAS.Dgv_Bank.SelectedCells(5).Value)
        Dim total As Decimal = 0
        Dim nill As Integer = -1
        Dim or_amt = Amt_In - Amt_Out
        Dim bankdata = SPAS.Dgv_Bank_Account.DataSource

        If bankdata.Rows.Count <> 0 Then

            Dim amt As Decimal
            For x As Integer = 0 To bankdata.Rows.Count - 1
                If bankdata.Rows(x)(0) = nocat Then
                    nill = x
                Else
                    If IsDBNull(bankdata.Rows(x)(2)) Then amt = 0 Else amt = CDec(bankdata.Rows(x)(2))
                    total = total + amt
                End If
            Next
            Dim diff = or_amt - total
            If nill = -1 Then

                If diff <> 0 Then  'account 'uncategorized not present
                    Dim R As DataRow
                    R = bankdata.Rows.Add
                    R(0) = nocat
                    R(1) = QuerySQL("SELECT name FROM account WHERE id='" & nocat & "'")
                    R(2) = diff
                End If
            Else
                bankdata.Rows(nill)(2) = or_amt - total
            End If
            SPAS.Tbx_Bank_Amount.Text = diff

        End If

    End Sub

    Sub Add_Journal_post_to_banktransaction()
        Dim bankdata = SPAS.Dgv_Bank_Account.DataSource
        If Check_Change_Bank_Categories(True) = False Then Exit Sub
        SPAS.isManualChange = True
        If (Not SPAS.Rbn_Bank_jtype_con.Checked And Not SPAS.Rbn_Bank_jtype_ext.Checked And Not SPAS.Rbn_Bank_jtype_int.Checked) And SPAS.Pan_Bank_jtype.Visible Then
            MsgBox("Selecteer eerst of dit een contractgift, extra gift of een andere banktransactie betreft")
            'Exit Sub
        End If

        If SPAS.Cmx_Bank_Account.Text = "" Or (Not IsNumeric(SPAS.Tbx_Bank_Amount.Text)) Or SPAS.Tbx_Bank_Amount.Text = "" Or SPAS.Cmx_Bank_Account.SelectedIndex = -1 Then
            'MsgBox("Nieuwe categorie: Ongeldige invoer")
            Exit Sub
        Else
            If SPAS.Cmx_Bank_Account.SelectedValue = QuerySQL("Select value from settings where label='nocat'") Then Exit Sub
            Dim R As DataRow
            R = bankdata.Rows.Add
            R(0) = SPAS.Cmx_Bank_Account.SelectedValue
            R(1) = SPAS.Cmx_Bank_Account.Text
            R(2) = SPAS.Tbx_Bank_Amount.Text
            Dim newRowIndex As Integer = SPAS.Dgv_Bank_Account.Rows.Count - 1
            SPAS.Dgv_Bank_Account.Rows(newRowIndex).Tag = "Modified"

            Calculate_Total_Booked("Btn_Bank_Add_Journal_Click")

            'Save_Banktransaction_Accounts()
            'Update_Category_Status()
        End If

    End Sub


    Sub Save_Banktransaction_Accounts()
        'Opslaan van aanpasbare banktransactiedata (description)
        ' Dit mag in alle gevallen worden aanpast

        Dim SQLstr As String
        Dim bankid = SPAS.Dgv_Bank.SelectedCells(0).Value

        '' 1) banktransactieomschrijving opslaan
        SQLstr = $"UPDATE bank SET description='{SPAS.Tbx_Bank_Description.Text}' WHERE id='{bankid}'"
        RunSQL(SQLstr, "NULL", "Save_Banktransaction_Accounts")
        SPAS.Dgv_Bank.SelectedCells(3).Value = SPAS.Tbx_Bank_Description.Text

        '' 2) Opslaan van de journaalposten waarmee de banktransacties gecategoriseerd zijn
        '' Hierbij wordt ook fk_journal_name aangepast
        '' Hiervoor vindt een check plaats of een banktransactie een incasso of excasso betreft, deze mogen niet
        '' aangepast c.q. opgeslagen worden


        Dim modified As Boolean = False
        Dim red As Boolean = False
        For Each row As DataGridViewRow In SPAS.Dgv_Bank_Account.Rows
            If row.Tag IsNot Nothing Then
                If row.Tag.ToString = "Modified" Then
                    modified = True
                End If
            End If
            If row.Cells(0).Value = nocat And row.Cells(2).Value <> 0 Then red = True Else red = False
            'Eerste blokkade voor het corrumperen van uitkering/incassoboekingen
            If Not IsDBNull(row.Cells(4).Value) Then
                If Trim(row.Cells(4).Value) <> "Bank" Then
                    Exit Sub
                End If
            End If
        Next


        If modified Then
            Dim bid As Integer = SPAS.Dgv_Bank.SelectedCells(0).Value
            Dim _dat As Date = SPAS.Dgv_Bank.SelectedCells(1).Value
            Dim dat As String = _dat.Year & "-" & _dat.Month & "-" & _dat.Day
            Dim des As String = SPAS.Dgv_Bank.SelectedCells(3).Value  'dit gaat fout met een bestaande excassojob waar al een beschrijving aanwezig is
            Dim afschrift As String = SPAS.Dgv_Bank.SelectedCells(9).Value
            Dim typ As String = "---"
            Dim nam As String
            Dim iban As String = Strings.Right(SPAS.Cmx_Bank_bankacc.Text, 18)
            Dim source As String = SPAS.Dgv_Bank.SelectedCells(12).Value
            Dim bankdata = SPAS.Dgv_Bank_Account.DataSource


            If source = "Uitkering" Or source = "Incasso" Then
                'Tweede blokkade voor het corrumperen van uitkering/incassoboekingen
                MsgBox("Incasso- & uitkeringslijsten kunnen niet in de bankapplicatie aangepast worden")
                Exit Sub
            End If

            If SPAS.Rbn_Bank_jtype_con.Checked Then
                typ = "Contract"
                nam = $"Contractgift ({afschrift})"
            ElseIf SPAS.Rbn_Bank_jtype_ext.Checked Then
                typ = "Extra"
                nam = $"Extra gift ({afschrift})"
            Else
                typ = "Internal"
                nam = $"Fondsgift ({afschrift})"
            End If

            SQLstr = $"DELETE FROM journal WHERE fk_bank={bid};" &
                         "INSERT INTO journal(date,status,amt1,description,source, fk_account,fk_bank,name,type,iban) VALUES "

            For x As Integer = 0 To bankdata.Rows.Count - 1
                If Not IsDBNull(bankdata.Rows(x)(2)) Then
                    nam = IIf(bankdata.Rows(x)(0) = nocat, "nog te bepalen", nam)
                    If nam = "Betaling intern account" Then nam = nam & "/" & bankdata.Rows(x)(0)
                    If bankdata.Rows(x)(2) <> 0 Then
                        SQLstr &= $"('{dat}','Verwerkt','{Cur2(bankdata.Rows(x)(2))}','{des}','Bank',{bankdata.Rows(x)(0)},{bid},'{nam}','{typ}','{iban}'),"
                    End If
                End If
            Next

            SQLstr = Strings.Left(SQLstr, Strings.Len(SQLstr) - 1) 'remove the last comma
            If SPAS.Chbx_test.Checked Then MsgBox(SQLstr)
            RunSQL(SQLstr, "NULL", "")

            RunSQL("update bank b set fk_journal_name = j.source from journal j where b.id = j.fk_bank and j.fk_account !=" & nocat & " and b.fk_journal_name='nog te bepalen';
            update bank b set fk_journal_name='nog te bepalen' from journal j where b.id = j.fk_bank and j.fk_account =" & nocat, "NULL", "Categorize_Bank_Transactions / Set journal Name")

        End If
        'tekstkleur aanpassen op basis van aanwezigheid nocat (niet toegewezen)'
        SPAS.Dgv_Bank.Rows(SPAS.Dgv_Bank.CurrentRow.Index).DefaultCellStyle.ForeColor = IIf(red, Color.DarkRed, Color.DarkGreen)


    End Sub
    Sub Fill_bank_transactions(sender, rowindex)

        SPAS.isManualChange = False
        If SPAS.Cmx_Bank_bankacc.SelectedIndex = -1 Then SPAS.Cmx_Bank_bankacc.SelectedIndex = 0
        Calculate_Bank_Balance()
        If Strings.InStr(SPAS.Cmx_Bank_bankacc.Text, "NL") = 0 Then Exit Sub

        Dim bankacc = Strings.Right(SPAS.Cmx_Bank_bankacc.Text, 18)

        Dim SQLstr = $"SELECT id, date As Datum, name As Naam, description As Omschrijving, 
                      credit As bij, debit as Af, code As cod, exch_rate, iban2, seqorder As Afschrift,
                      batchid, amt_cur, fk_journal_name As Journaalnaam,filename,cost,iban, id As Bankid,
                      (select count(j.id) from journal j left join bank b2 on b2.id=j.fk_bank where j.fk_account='" & nocat & "' and b.id = b2.id)
                      FROM bank b WHERE iban ='" & bankacc & "' ORDER BY seqorder DESC, date DESC"

        SPAS.Prepare_Datagridview(SPAS.Dgv_Bank, SQLstr, {"HZ010", "FZ050", "TZ150", "TZ300", "NZ070", "NZ070", "HZ030", "HZ030", "HZ030", "TZ040", "HZ030", "HZ030", "TZ070", "HZ030", "HZ030", "HZ030", "TZ040", "TZ020"})
        SPAS.Format_dvg_bank()

        If rowindex = Nothing Then rowindex = 0
        If SPAS.Dgv_Bank.RowCount > 0 Then SPAS.Dgv_Bank.Rows(rowindex).Selected = True
        SPAS.Dgv_Bank.Enabled = True
        SPAS.isManualChange = True

    End Sub
    Function Check_Change_Bank_Categories(ByVal msg As Boolean)
        If SPAS.Dgv_Bank.Rows.Count = 0 Or SPAS.Dgv_Bank_Account.Rows.Count = 0 Then
            Return False
            Exit Function
        End If
        If Not IsDBNull(SPAS.Dgv_Bank.SelectedCells(12).Value) Then
            If SPAS.Dgv_Bank.SelectedCells(12).Value = "Uitkering" Or SPAS.Dgv_Bank.SelectedCells(12).Value = "Incasso" Then
                If msg Then MsgBox("Incasso- & uitkeringslijsten kunnen niet in de bankapplicatie aangepast worden")
                Fill_Journals_by_bank(SPAS.Dgv_Bank.SelectedCells(0).Value)
                Return False
            End If
        End If
        Return True

    End Function
End Module
