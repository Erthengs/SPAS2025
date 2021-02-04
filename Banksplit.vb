Public Class Banksplit
    Private Sub Dgv_Split_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles Dgv_Split.CellContentClick

    End Sub



    Sub Calculate_Split_Totals()
        Dim tot As Decimal = 0
        Dim amt As Decimal = 0
        Dim diff As Decimal = 0
        For x As Integer = 0 To dst1.Tables(0).Rows.Count - 1
            'MsgBox(dst1.Tables(0).Rows(x)(1))
            If IsDBNull(dst1.Tables(0).Rows(x)(1)) Then
                amt = 0
            Else
                amt = CDec(dst1.Tables(0).Rows(x)(1))
            End If

            tot = tot + amt
        Next

        diff = Tbx2Dec(Tbx_Split_Amount.Text) - tot
        Tbx_Split_Diff.Text = diff

        'dst.Tables(0).Rows(0)(1) = Tbx2Dec(Tbx_Split_Amount.Text) + diff
    End Sub



    Private Sub Dgv_Split_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles Dgv_Split.CellEndEdit
        Calculate_Split_Totals()
    End Sub

    Private Sub Btn_Split_Save_Click(sender As Object, e As EventArgs) Handles Btn_Split_Save.Click
        Calculate_Split_Totals()
        If Tbx2Dec(Tbx_Split_Diff.Text) <> 0 Then
            MsgBox("Deze banktransactie is onjuist verdeeld")
            Exit Sub
        End If

        'Bewaar banktransactie
        'geldig voor alle colommen binnen een transactie
        Dim iba As String = SPAS.Dgv_Bank.SelectedCells(15).Value
        Dim cur As String = "EUR"
        Dim _dat As Date = SPAS.Dgv_Bank.SelectedCells(1).Value
        Dim dat As String = _dat.Year & "-" & _dat.Month & "-" & _dat.Day
        Dim seq As Integer = Tbx_Split_seqorder.Text
        Dim ib2 As String = SPAS.Dgv_Bank.SelectedCells(8).Value
        Dim nam As String = SPAS.Dgv_Bank.SelectedCells(2).Value
        Dim cod As String = Strings.Trim(SPAS.Dgv_Bank.SelectedCells(6).Value)
        Dim bat As String = SPAS.Dgv_Bank.SelectedCells(3).Value
        Dim exc As Decimal = SPAS.Dgv_Bank.SelectedCells(7).Value
        Dim amt_cur As Decimal = 0
        Dim fkj As String = SPAS.Dgv_Bank.SelectedCells(12).Value
        Dim fil As String = SPAS.Dgv_Bank.SelectedCells(13).Value
        Dim cst As Decimal = 0

        'specifiek voor de gesplitste transactie
        Dim val As Decimal
        Dim deb As Decimal
        Dim cre As Decimal
        Dim des As String


        Dim Sqlstr As String = "
        DELETE FROM bank WHERE seqorder='" & seq & "';"
        Sqlstr &= "INSERT INTO BANK(iban,currency,date,seqorder,iban2,name,code,batchid,exch_rate, 
        amt_cur,fk_journal_name,filename,cost,debit,credit,description) VALUES"


        For x As Integer = 0 To Me.Dgv_Split.Rows.Count - 1
            des = Me.Dgv_Split.Rows(x).Cells(0).Value
            val = Me.Dgv_Split.Rows(x).Cells(1).Value
            deb = IIf(val < 0, val * -1, 0)
            cre = IIf(val > 0, val, 0)
            If val = 0 Then GoTo skipit  'by this way splitting can be undone, empty bank transactions are not stored
            Sqlstr &= "('" & iba & "','" & cur & "','" & dat & "','" & seq & "','" & ib2 & "','" &
                        nam & "','" & cod & "','" & bat & "','" & exc & "','" & amt_cur & "','" &
                        fkj & "','" & fil & "','" & cst & "','" &
                        Cur2(deb) & "','" & Cur2(cre) & "','" & des & "'),"
skipit:
        Next
        Sqlstr = Strings.Left(Sqlstr, Strings.Len(Sqlstr) - 1)
        'Clipboard.Clear()
        'Clipboard.SetText(Sqlstr)
        RunSQL(Sqlstr, "NULL", "Btn_Split_Save_Click")
        'Now updating journal
        Dim Sqlstr2 As String =
            "Update journal j
            Set fk_bank = b.id, status = 'Verwerkt' from bank b 
            WHERE b.description = j.name
            AND b.seqorder = '" & Tbx_Split_seqorder.Text & "' AND j.status = 'Open';
            DELETE from journal WHERE fk_bank = '" & Tbx_Split_Bank_id.Text & "';"
        RunSQL(Sqlstr2, "NULL", "Btn_Split_Save_Click-update journal")
        Me.Close()
        SPAS.Fill_bank_transactions()
        Fill_Cmx_Excasso_Select_Combined()
    End Sub

    Private Sub Btn_Split_Cancel_Click(sender As Object, e As EventArgs) Handles Btn_Split_Cancel.Click
        Me.Close()
    End Sub

    Private Sub Banksplit_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Tbx_Split_Description.Text = SPAS.Dgv_Bank.SelectedCells(3).Value
        Tbx_Split_seqorder.Text = SPAS.Dgv_Bank.SelectedCells(9).Value
        Tbx_Split_Bank_id.Text = SPAS.Dgv_Bank.SelectedCells(0).Value
        Tbx_Split_Amount.Text = QuerySQL("Select sum(credit) - sum(debit) from bank where seqorder = '" & Tbx_Split_seqorder.Text & "';")

        Collect_data1("SELECT description, credit-debit FROM bank 
                     WHERE seqorder='" & Tbx_Split_seqorder.Text & "'")
        'Clipboard.Clear()
        'Clipboard.SetText(SQLstr)
        Dgv_Split.DataSource = dst1.Tables(0)
        Try
            With Dgv_Split

                .Columns(0).HeaderText = "Omschrijving"
                .Columns(0).DefaultCellStyle.ForeColor = Color.Blue
                .Columns(0).Width = 320

                .Columns(1).HeaderText = "Bedrag"
                .Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                .Columns(1).DefaultCellStyle.Format = "N2"
                .Columns(1).DefaultCellStyle.ForeColor = Color.Blue
                .Columns(1).Width = 75

            End With
        Catch
        End Try
    End Sub
End Class