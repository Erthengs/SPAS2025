Public Class Banksplit


    Sub Calculate_Split_Totals()
        Dim tot As Decimal = 0
        Dim amt As Decimal = 0
        Dim diff As Decimal = 0
        If Me.Dgv_Split.Rows.Count = 2 Then
            MsgBox("U heeft niets gesplitst")
            Exit Sub
        End If


        For x As Integer = 0 To Me.Dgv_Split.Rows.Count - 1
            If IsDBNull(Me.Dgv_Split.Rows(x).Cells(1).Value) Then amt = 0 Else amt = CDec(Me.Dgv_Split.Rows(x).Cells(1).Value)
            tot = tot + amt
        Next x
        diff = Tbx2Dec(Lbl_Split_Amount.Text) - tot
        Lbl_Split_Diff.Text = diff
        Lbl_Split_Diff.ForeColor = IIf(diff = 0, Color.Black, Color.Red)

    End Sub
    Private Sub Dgv_Split_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles Dgv_Split.CellEndEdit
        Calculate_Split_Totals()
    End Sub

    Private Sub Btn_Split_Save_Click(sender As Object, e As EventArgs) Handles Btn_Split_Save.Click
        Calculate_Split_Totals()
        If Tbx2Dec(Lbl_Split_Diff.Text) <> 0 Then
            MsgBox("Deze banktransactie is onjuist verdeeld")
            Exit Sub
        End If

        If CInt(QuerySQL("select count(distinct(fk_account)) from journal j where j.fk_bank In (Select id from bank where seqorder ='" & Lbl_Split_seqorder.Text & "')")) > 1 Then
            If MsgBox("Momenteel is deze transacties verdeeld over meerdere categories." & vbCr &
                   "Door deze wijziging te bewaren krijgen alle subtransacties de categorie " & Lbl_SplitBank_Accountnr.Text & vbCr &
                   "Wilt u dat?", vbYesNo, vbExclamation) = vbNo Then
                Exit Sub
            End If
        End If


        'Bewaar banktransactie
        'geldig voor alle colommen binnen een transactie
        Dim iba As String = SPAS.Dgv_Bank.SelectedCells(15).Value
        Dim cur As String = "EUR"
        Dim _dat As Date = SPAS.Dgv_Bank.SelectedCells(1).Value
        Dim dat As String = _dat.Year & "-" & _dat.Month & "-" & _dat.Day
        Dim seq As Integer = Lbl_Split_seqorder.Text
        Dim ib2 As String = SPAS.Dgv_Bank.SelectedCells(8).Value
        Dim nam As String = SPAS.Dgv_Bank.SelectedCells(2).Value
        Dim cod As String = Strings.Trim(SPAS.Dgv_Bank.SelectedCells(6).Value)
        Dim bat As String = SPAS.Dgv_Bank.SelectedCells(3).Value
        Dim exc As Decimal = SPAS.Dgv_Bank.SelectedCells(7).Value
        Dim amt_cur As Decimal = 0
        Dim fkj As String = Lbl_SplitBank_journal_name.Text 'SPAS.n20(SPAS.Dgv_Bank.SelectedCells(12).Value)

        Dim fil As String = SPAS.Dgv_Bank.SelectedCells(13).Value
        Dim cst As Decimal = 0
        Dim new_id = QuerySQL("Select Max(id) FROM Bank")
        Dim accountid = Strings.Left(Lbl_SplitBank_Accountnr.Text, InStr(Lbl_SplitBank_Accountnr.Text, " [") - 1)
        Dim typ As String = Lbl_SplitBank_Type.Text

        'specifiek voor de gesplitste transactie
        Dim val As Decimal
        Dim deb As Decimal
        Dim cre As Decimal
        Dim des As String
        Dim rst As Decimal = Lbl_Split_Diff.Text

        Dim Sqlstr As String = "DELETE FROM bank WHERE seqorder='" & seq & "';"
        Sqlstr &= "INSERT INTO BANK(iban,currency,date,seqorder,iban2,name,code,batchid,exch_rate, 
                   amt_cur,fk_journal_name,filename,cost,debit,credit,description) VALUES"
        Dim SQLstr2 As String = "INSERT INTO journal(name,date,status,description,source,amt1,fk_account,
                                fk_bank,fk_relation,type, iban) VALUES "


        For x As Integer = 0 To Me.Dgv_Split.Rows.Count - 1
            des = Me.Dgv_Split.Rows(x).Cells(0).Value &
              IIf(Me.Dgv_Split.Rows(x).Cells(0).Value <> Lbl_Split_Description.Text,
              " | " & Lbl_Split_Description.Text, "")

            If Not (IsDBNull(Me.Dgv_Split.Rows(x).Cells(1).Value)) Then val = Me.Dgv_Split.Rows(x).Cells(1).Value Else val = 0
            'val = SPAS.n20(val)
            deb = IIf(val < 0, val * -1, 0)


            cre = IIf(val > 0, val, 0)
            If val = 0 Then GoTo skipit  'by this way splitting can be undone, empty bank transactions are not stored
            new_id = new_id + 1
            Sqlstr &= "('" & iba & "','" & cur & "','" & dat & "','" & seq & "','" & ib2 & "','" &
                        nam & "','" & cod & "','" & bat & "','" & exc & "','" & amt_cur & "','" &
                        fkj & "','" & fil & "','" & cst & "','" &
                        Cur2(deb) & "','" & Cur2(cre) & "','" & des & "'),"
            SQLstr2 &= "('" & nam & "','" & dat & "','Open','" & des & "','Bank','" &
                        Cur2(val) & "','" & accountid & "'," & new_id & ",0,'" & typ & "','" & iba & "')," 'FK_BANK NOT CORRECT, THEREFORE NOT LINKED 
skipit:
        Next

        Sqlstr = Strings.Left(Sqlstr, Strings.Len(Sqlstr) - 1)
        SQLstr2 = Strings.Left(SQLstr2, Strings.Len(SQLstr2) - 1)
        ToClipboard(SQLstr2, True)

        RunSQL(Sqlstr, "NULL", "Btn_Split_Save_Click")
        RunSQL(SQLstr2, "NULL", "Btn_Split_Save_Click")


        'delete function vervangen door update van de originele transactie met het restbedrag
        '---> "UPDATE journal j set status = 'Verwerkt', amt1 = '" & Lbl_Split_Diff.Text & "' WHERE fk_bank = '" & Lbl__Split_Bank_id.Text & "';"
        'restant 
        If rst = 0 Then
            RunSQL("DELETE from journal WHERE fk_bank = '" & Lbl_Split_Bank_id.Text & "';", "NULL", "Btn_Split_Save_Click2")
        Else
            RunSQL("UPDATE journal j set status = 'Verwerkt', amt1 = '" & Lbl_Split_Diff.Text & "' WHERE fk_bank = '" & Lbl_Split_Bank_id.Text & "';", "NULL", "Btn_Split_Save_Click_3")
        End If

        Me.Close()
        Categorize_Bank_Transactions(False, True, False, False, True, False, False)
        SPAS.Fill_bank_transactions("Btn_Split_Save_Click")
        Fill_Cmx_Excasso_Select_Combined()

    End Sub

    Private Sub Btn_Split_Cancel_Click(sender As Object, e As EventArgs) Handles Btn_Split_Cancel.Click
        Me.Close()
    End Sub

    Private Sub Banksplit_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'uitvoeren van checks


        Collect_data1("SELECT description As Omschrijving, credit-debit As Bedrag FROM bank 
                     WHERE seqorder='" & Lbl_Split_seqorder.Text & "'")

        Dgv_Split.DataSource = dst1.Tables(0)


        SPAS.Format_Datagridview(Dgv_Split, {"T360", "N080"}, True)

        Exit Sub
        Try
            With Dgv_Split

                .Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                .Columns(1).DefaultCellStyle.Format = "N2"
                .Columns(1).DefaultCellStyle.ForeColor = Color.Blue
                .Columns(1).Width = 75

            End With
        Catch
        End Try
    End Sub

    Private Sub Dgv_Split_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles Dgv_Split.DataError
        MsgBox("Ongeldige invoer")
        e.ThrowException = False
    End Sub

    Private Sub Btn_Prefill_Split_Click(sender As Object, e As EventArgs) Handles Btn_Prefill_Split.Click
        Dim errmsg As String = ""
        If SPAS.Dgv_Bank.SelectedCells(12).Value <> "nog te bepalen" Then
            MsgBox("Laden van openstaande uitkeringen kan alleen bij ongecategoriseerde banktransacties.")
            Exit Sub
        End If

        Dim sql As String = "
                    SELECT name As Omschrijving, SUM(AMt1) AS Bedrag FROM journal
                    WHERE name ILIKE 'Excasso%' AND status = 'Open' GROUP By name, status
"
        Load_Datagridview(Dgv_Split, sql, "Btn_Prefill_Split")
        SPAS.Format_Datagridview(Dgv_Split, {"T360", "N080"}, True)
        Calculate_Split_Totals()
    End Sub

    Private Sub Dgv_Split_UserDeletedRow(sender As Object, e As DataGridViewRowEventArgs) Handles Dgv_Split.UserDeletedRow
        Calculate_Split_Totals()
    End Sub

End Class