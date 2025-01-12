Imports System.ComponentModel.DataAnnotations
Imports System.Data.Entity
Imports System.Data.Entity.Core.Common.CommandTrees.ExpressionBuilder
Imports System.Data.Entity.Migrations
Imports System.IO
Imports System.Management.Instrumentation
Imports System.Reflection
Imports System.Security.Cryptography
Imports System.Windows.Forms.VisualStyles
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Rebar
Imports System.Xml
Imports Microsoft.EntityFrameworkCore.Metadata.Internal
Imports Microsoft.EntityFrameworkCore.Query.Internal
Imports Microsoft.EntityFrameworkCore.Query.SqlExpressions
Imports Npgsql
Imports NpgsqlTypes

Public Class SPAS
    Private Const V As Boolean = False
    Private PreviousTab As Integer
    Private PreviousTabMain As Integer
    Private oldend_date As Date
    'bekende fouten

    Private Sub SPAS_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'Login.Cmx_Login_Database.Text = "Productie"
        Login.ShowDialog()
        'InitLoad()
        'ook gebruiken na bewaren van nieuwe cp, bankacc, relation en target
    End Sub
    Sub InitLoad()
        If username = "" Then Exit Sub

        RunSQL("Update contract Set Active ='false' where enddate < current_date", "NULL", "SPAS_Load")
        Load_Comboboxes()
        TC_Object.SelectedIndex = 0
        PreviousTabMain = 0
        PreviousTab = 0
        'Select_Obj2("InitLoad")
        Load_Table()
        If Lbx_Basis.Items.Count = 0 Then Empty_Tabpage()
        nocat = QuerySQL("SELECT value FROM settings WHERE label='nocat'")
        Load_Account_Settings()
        report_year = QuerySQL("select min(extract (year from date)) from journal")

        Dim sql = $"SELECT module, name, sql from query where category = 'Overzicht' order by module, name;"
        Populate_DataTree(sql, BankTree)



    End Sub


    Sub Load_Comboboxes()
        'can go wrong if tables are empty

        Load_Combobox(Cmx_01_cp__fk_bankacc_id, "id", "name", "SELECT id, Name||'/'||accountno as name FROM bankacc WHERE expense=True AND active=TRUE ORDER BY name")
        Load_Combobox(Cmx_Incasso_Bankaccount, "id", "name", "SELECT id, accountno AS name FROM bankacc WHERE expense=FALSE AND active=TRUE ORDER BY name")
        Load_Combobox(Cmx_01_Target__fk_cp_id, "id", "name", "SELECT id, CONCAT(name, ', ', name_add) as name FROM cp WHERE active=True ORDER BY name")
        Load_Combobox(Cmx_00_contract__fk_relation_id, "id", "name", "SELECT id, CONCAT(name, ', ', name_add) as name FROM relation WHERE active=TRUE ORDER BY name")
        Load_Combobox(Cmbx_journaalposten_relatie, "id", "name", "SELECT id, CONCAT(name, ', ', name_add) as name FROM relation ORDER BY name")
        Load_Combobox(Cmbx_journaalposten_account, "id", "name", "SELECT id, name FROM account ORDER BY name")
        Load_Combobox(Cmx_00_Account__accgroup, "accgroup", "name", "SELECT DISTINCT accgroup FROM account WHERE accgroup <> '' ORDER BY accgroup")
        Load_Combobox(Cmx_Bank_bankacc, "id", "name", "SELECT id, CONCAT(Name, '/', accountno) as name FROM bankacc ORDER BY name DESC")
        Load_Combobox(Cmx_Bank_Account, "id", "name", "SELECT id, name FROM account WHERE active = True ORDER BY source, name")
        Load_Combobox(Cmx_00_Contract__fk_account_id, "id", "name", "SELECT id, CONCAT(id, ' ',name) As name FROM account 
                                          WHERE active=TRUE AND source='cat' AND type = 'Inkomsten' ORDER BY name")
        Load_Combobox(Cmx_01_account__fk_accgroup_id, "id", "name", "SELECT id, name FROM accgroup WHERE active=True ORDER BY name")
        Load_Combobox(Cmx_Incasso_Jobs, "name", "name", "select distinct name, date, status from journal where source='Incasso' and status !='Verwerkt'")
        Populate_Single_Combobox(Cmbx_Reporting_Year, "

select distinct extract (year from date) As Year from journal_archive 
                                            union select distinct min(extract (year from date)) from journal")

        'Clipboard.Clear()
        'Clipboard.SetText("SELECT id, CONCAT(name, ', ', name_add) as name FROM target WHERE active=TRUE ORDER BY name")
        If Me.Dgv_Mgnt_Tables.Rows(8).Cells(1).Value > 0 Then

            Load_Combobox(Cmx_01_contract__fk_target_id, "id", "name", "SELECT id, name||', '||name_add as name FROM target WHERE active=TRUE ORDER BY name")
        End If
        '@@@ hier gaat iets fout
        Fill_Cmx_Excasso_Select_Combined()
        Me.Cmbx_journaalposten_account.SelectedIndex = -1
        Me.Cmbx_journaalposten_relatie.SelectedIndex = -1


    End Sub
    Private Sub TC_Object_Click(sender As Object, e As EventArgs) Handles TC_Object.Click
        'If Edit_Mode Or Add_Mode Then
        'If Btn_Basis_Save.Enabled Then
        'MsgBox("TC_Object_Click")
        If MenuSave.Enabled Then
            MsgBox("U bent nog bezig met het " & IIf(Edit_Mode, "bewerken", "aanmaken") & " van een " & TC_Object.TabPages(PreviousTab).Text.ToLower & ".")
            TC_Object.SelectedIndex = PreviousTab
            TC_Main.SelectedIndex = PreviousTabMain

        Else
            Load_Table()
        End If
    End Sub

    Private Sub TC_Object_Selecting(sender As Object, e As TabControlCancelEventArgs) Handles TC_Object.Selecting
        If MenuSave.Enabled Then
            TC_Main.SelectedIndex = PreviousTabMain
            TC_Object.SelectedIndex = PreviousTab
        End If
    End Sub
    Private Sub Rbtn_Target_Child_CheckedChanged(sender As Object, e As EventArgs) Handles Rbtn_Target_Child.CheckedChanged
        If MenuSave.Enabled Then Tbx_01_Target__ttype.Text = Rbtn_Target_Child.Text
    End Sub

    Private Sub Rbtn_Target_Elder_CheckedChanged(sender As Object, e As EventArgs) Handles Rbtn_Target_Elder.CheckedChanged
        If MenuSave.Enabled Then Tbx_01_Target__ttype.Text = Rbtn_Target_Elder.Text
    End Sub

    Private Sub Rbtn_Target_Other_CheckedChanged(sender As Object, e As EventArgs) Handles Rbtn_Target_Other.CheckedChanged
        If MenuSave.Enabled Then Tbx_01_Target__ttype.Text = Rbtn_Target_Other.Text
    End Sub

    Private Sub Btn_Basis_Add_Click(sender As Object, e As EventArgs) Handles Btn_Basis_Add.Click
        Basis_Add()
    End Sub
    Sub Basis_Add()

        Dim t As String = TC_Object.SelectedIndex.ToString

        Add_Mode = True
        Manage_Buttons_Target(False, False, False, True, True, "Btn_Basis_Add_Click")
        Empty_Tabpage()


        If TC_Object.SelectedIndex = 0 Then  'additional functionality for contract management

            Dtp_31_contract__startdate.Value = Date.Today
            Me.Rbn_00_contract_child.Checked = True
            Rbn_00_contract_child.Checked = True
            '---------------- Temp solution of error
            Lbl_00_Contract__name.Text = Contract_number("K")
            Load_Combobox(Cmx_01_contract__fk_target_id, "id", "name", "Select id, Name||', '||name_add as name FROM target
                                                        WHERE ttype='" & Rbn_00_contract_child.Text & "' AND active= TRUE ORDER BY name")
            '-------standaard_waarden ophalen

            Tbx_11_Contract__donation.Text = QuerySQL("select value from settings where label ilike 'standaard_bedrag_kind'")
            Tbx_11_contract__overhead.Text = QuerySQL("select value from settings where label ilike 'standaard_overhead_kind'")
            Dtp_31_contract__startdate.Value = New DateTime(Date.Today.Year, Date.Today.Month, 1).AddMonths(1)
            '----------------

            Handle_Contract_Fields()
            Cmx_02_Contract__term.Text = 12
            Pan_Contract_Date_New.Visible = False
            Cbx_00_contract__active.Checked = True
            Rbn_00_contract_child.Checked = True
            Pan_contract_select_target.Enabled = True
            Lbl_00_Contract__name.Text = Contract_number("K")
            Load_Combobox(Cmx_01_contract__fk_target_id, "id", "name", "SELECT id, 
            name||', '||name_add) as name FROM target
            WHERE ttype='" & Rbn_00_contract_child.Text & "' 
            And active=true ORDER BY name")
            Cmx_01_contract__fk_target_id.Text = ""
            Chx_00_contract__autcol.Enabled = False

        End If
        If TC_Object.SelectedIndex = 1 Then
            Pan_Target.Enabled = True
            Cbx_00_target__active.Checked = True
            Dtp_00_Target__birthday.Value = Date.Today
        ElseIf TC_Object.SelectedIndex = 4 Then
            ' = True
            Cbx_00_Account__active.Checked = True
            Lbl_00_Account__source.Text = "cat"
            Lbl_20_Account__f_key.Text = QuerySQL("SELECT Max(f_key) FROM account Where source='cat'") + 1
            Tbx_01_Account__name.Enabled = True
            Lbl_00_pkid.Text = ""

        End If

    End Sub
    Public Sub Btn_Basis_Cancel_Click(sender As Object, e As EventArgs) Handles Btn_Basis_Cancel.Click
        Cancel()

    End Sub
    Sub Cancel()
        Select_Obj2("Cancel")
        Manage_Buttons_Target(True, True, True, False, False, "Cancel")
        Edit_Mode = False
        Add_Mode = False
        Pan_Target.Enabled = False
        Pan_contract_select_target.Enabled = False

        If TC_Object.SelectedIndex = 0 Then  'additional functionality for contract management
            Handle_Contract_Fields()
            Pan_Contract_Date_New.Visible = False
        End If
        If TC_Object.SelectedIndex = 4 Then
            Lbl_Account_Budget_Difference.Text = ""
        End If


    End Sub

    Private Sub Btn_Basis_Save_Click(sender As Object, e As EventArgs) Handles Btn_Basis_Save.Click
        Basis_Save()
    End Sub
    Sub Basis_Save()
        Dim tbl As String = Me.TC_Object.TabPages(Me.TC_Object.SelectedIndex).Name
        Dim val, val2 As Integer
        Dim errmsg = Handle_errors("")
        If errmsg <> "" Then
            MsgBox(errmsg)
            Exit Sub
        End If
        If Lbx_Basis.SelectedIndex = 5 Then
            If Cbx_00_BankAcc__income.Checked And (Tbx_00_BankAcc__bic.Text = "" Or Tbx_00_BankAcc__id2.Text = "") Then
                MsgBox("Voor inkomstenaccounts is het invullen van BIC en bankidnummer verplicht")
                Exit Sub
            End If
        End If
        'check uitvoeren op overlappende contracten met hetzelfde sponsordoel...

        If Lbx_Basis.SelectedIndex <> -1 Then val = Lbx_Basis.SelectedItem(Me.Lbx_Basis.ValueMember)

        Select Case TC_Object.SelectedIndex
            Case 0
                If Add_Mode Then

                    Insert_into_table() 'regular adding to database
                    val = Convert.ToInt32(QuerySQL("Select MAX(id) FROM " & tbl))
                    reload = True

                Else 'change mode

                    'relation, target and target type can never be changed; this would imply another contract
                    'description may be changed freely -- now not possible

                    'Handle_Contract_Fields()
                    If Dtp_30_Contract_Change.Visible = True Then   'new version of the contract / edit_mode

                        '1 Close current contract by updating enddate and active if applicable
                        Dim d1, d2 As DateTime
                        Dim act As Boolean

                        d1 = Me.Dtp_30_Contract_Change.Value
                        Dim _d1 As String = d1.Year & "-" & d1.Month & "-" & d1.Day
                        d2 = New DateTime(d1.Year, d1.Month, d1.Day).AddDays(-1)
                        act = d2 > Date.Today
                        Dim _d2 As String = d2.Year & "-" & d2.Month & "-" & d2.Day
                        Dim sqlstr, msg As String
                        sqlstr = "UPDATE contract SET enddate='" & _d2 & "', active=" & act & " WHERE id=" & val & ";"

                        '2 Create a new contractversion 
                        sqlstr &= "INSERT INTO public.contract(fk_target_id, fk_relation_id, 
                                    donation, overhead, description, autcol, name, term,intern, fk_account_id) 
                                    SELECT fk_target_id, fk_relation_id, 
                                    donation, overhead, description, autcol, name, term,  
                                    intern, fk_account_id FROM contract WHERE id=" & val & ";"

                        'Clipboard.Clear()
                        'Clipboard.SetText(sqlstr)
                        RunSQL(sqlstr, "NULL", "MenuSave.Click upsert new version")

                        val2 = Convert.ToInt32(QuerySQL("Select MAX(id) FROM " & tbl))
                        '3 update new version with new values, startdate / enddate and active
                        sqlstr = "UPDATE contract SET startdate='" & _d1 & "', 
                           donation='" & Cur2(Replace(Tbx_11_Contract__donation.Text, ".", "")) & "', 
                           overhead='" & Cur2(Replace(Tbx_11_contract__overhead.Text, ".", "")) & "', 
                           enddate ='2999-12-31',active=true  
                           WHERE id=" & val2 & ";"
                        Clipboard.Clear()
                        Clipboard.SetText(sqlstr)
                        MsgBox(Cur2(Replace(Tbx_11_Contract__donation.Text, ".", "")))
                        RunSQL(sqlstr, "NULL", "MenuSave.Click update New version")
                        'reload = True
                        msg = "Een nieuwe versie van het contract is aangemaakt."
                        If act Then msg &= "De wijziging gaat in in de toekomst (nu nog inactief); wilt u de laatste versie nu bekijken?"

                        val = val2
                        reload = True
                        Pan_Contract_Date_New.Visible = False

                    Else
                        'updating description in the regular way
                        val = Lbx_Basis.SelectedItem(Me.Lbx_Basis.ValueMember)
                        Update_table()
                    End If
                    Dim acc_id As Integer = QuerySQL("select id from account where source = 'Doel' and f_key=" & Cmx_01_contract__fk_target_id.SelectedValue)
                    Calculate_Budget(acc_id)
                End If
            Case 1
                Dim tmp_cp = Cmx_01_Target__fk_cp_id.SelectedText

                Pan_contract_select_target.Enabled = False
                Pan_Target.Enabled = False
                If Add_Mode Then
                    Insert_into_table()
                    val = Convert.ToInt32(QuerySQL("Select MAX(id) FROM " & tbl))
                    reload = True

                Else
                    val = Lbx_Basis.SelectedItem(Me.Lbx_Basis.ValueMember)
                    Update_table()
                End If
                Cmx_01_Target__fk_cp_id.SelectedText = tmp_cp


            Case Else
                Pan_contract_select_target.Enabled = False
                Pan_Target.Enabled = False
                If Add_Mode Then
                    Insert_into_table()
                    val = Convert.ToInt32(QuerySQL("Select MAX(id) FROM " & tbl))
                    reload = True

                Else
                    val = Lbx_Basis.SelectedItem(Me.Lbx_Basis.ValueMember)
                    Update_table()
                End If
        End Select

        If reload Then
            Load_Table()
            Locate_Listbox_Position(val)

        End If
        'finalizing
        'Load_Comboboxes()
        Manage_Buttons_Target(True, True, True, False, False, "MenuSave.Click")
        Edit_Mode = False
        Add_Mode = False
        reload = False
    End Sub


    Private Sub Lbx_Basis_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Lbx_Basis.SelectedIndexChanged

        If InStr(sender.ToString, "System.Data.DataRowView") > 0 Then Exit Sub
        Click_Lbx_Basis()

    End Sub
    Private Sub Lbx_Basis_Click(sender As Object, e As EventArgs) Handles Lbx_Basis.Click
        If InStr(sender.ToString, "System.Data.DataRowView") > 0 Then Exit Sub
        'Click_Lbx_Basis()
    End Sub
    Sub Click_Lbx_Basis()
        If Lbx_Basis.Items.Count > 0 Then Select_Obj2("Lbx_Basis_SelectedIndexChanged") Else Empty_Tabpage()
    End Sub
    Private Sub Tbx_Target__ttype_TextChanged(sender As Object, e As EventArgs) Handles Tbx_01_Target__ttype.TextChanged
        Rbtn_Target_Child.Checked = Strings.Trim(Tbx_01_Target__ttype.Text) = "Kind"
        Rbtn_Target_Elder.Checked = Strings.Trim(Tbx_01_Target__ttype.Text) = "Oudere"
        Rbtn_Target_Other.Checked = Strings.Trim(Tbx_01_Target__ttype.Text) = "Overig"
        '@@@ hard value vervangen door tt_type.Text
    End Sub
    Private Sub Rbtn_Target_Alone_CheckedChanged(sender As Object, e As EventArgs)
        If MenuSave.Enabled Then Tbx_00_Target__living.Text = Rbtn_Target_Alone.Text
    End Sub

    Private Sub Rbtn_Target_Institution_CheckedChanged(sender As Object, e As EventArgs)
        If MenuSave.Enabled Then Tbx_00_Target__living.Text = Rbtn_Target_Institution.Text
    End Sub

    Private Sub Rbtn_Target_OtherHousing_CheckedChanged(sender As Object, e As EventArgs)
        If MenuSave.Enabled Then Tbx_00_Target__living.Text = Rbtn_Target_OtherHousing.Text
    End Sub

    Private Sub Tbx_Target__living_TextChanged(sender As Object, e As EventArgs) Handles Tbx_00_Target__living.TextChanged
        Rbtn_Target_Alone.Checked = Strings.Trim(Tbx_00_Target__living.Text) = "Alleen"
        Rbtn_Target_Institution.Checked = Strings.Trim(Tbx_00_Target__living.Text) = "Tehuis"
        Rbtn_Target_OtherHousing.Checked = Strings.Trim(Tbx_00_Target__living.Text) = "Overig"
    End Sub

    Private Sub Tbx_Target__income_TextChanged(sender As Object, e As EventArgs) Handles _
        Tbx_10_Target__income.TextChanged, Tbx_10_Target__pension.TextChanged, Tbx_10_Target__benefit.TextChanged,
        Tbx_10_Target__allowance.TextChanged, Tbx_10_Target__otherincome.TextChanged,
        Tbx_10_Target__rent.TextChanged, Tbx_10_Target__heating.TextChanged, Tbx_10_Target__heating.TextChanged,
        Tbx_10_Target__gaselectra.TextChanged, Tbx_10_Target__water.TextChanged, Tbx_10_Target__food.TextChanged,
        Tbx_10_Target__medicine.TextChanged
        If Edit_Mode Then Manage_Buttons_Target(False, False, False, True, True, "Tbx_Target__income_TextChanged")
        Calculate_Target_Totals()
    End Sub

    Private Sub Tbx_Target__income_Leave(sender As Object, e As EventArgs) Handles Tbx_10_Target__income.Leave
        Tbx_10_Target__income.Text = Tbx2Dec(Tbx_10_Target__income.Text)
    End Sub
    Private Sub Tbx_Target__pension_Leave(sender As Object, e As EventArgs) Handles Tbx_10_Target__pension.Leave
        Tbx_10_Target__pension.Text = Tbx2Dec(Tbx_10_Target__pension.Text)
    End Sub
    Private Sub Tbx_Target__benefit_Leave(sender As Object, e As EventArgs) Handles Tbx_10_Target__benefit.Leave
        Tbx_10_Target__benefit.Text = Tbx2Dec(Tbx_10_Target__benefit.Text)
    End Sub
    Private Sub Tbx_Target__allowance_Leave(sender As Object, e As EventArgs) Handles Tbx_10_Target__allowance.Leave
        Tbx_10_Target__allowance.Text = Tbx2Dec(Tbx_10_Target__allowance.Text)
    End Sub
    Private Sub Tbx_Target__otherincome_Leave(sender As Object, e As EventArgs) Handles Tbx_10_Target__otherincome.Leave
        Tbx_10_Target__otherincome.Text = Tbx2Dec(Tbx_10_Target__otherincome.Text)
    End Sub
    Private Sub Tbx_Target__rent_Leave(sender As Object, e As EventArgs) Handles Tbx_10_Target__rent.Leave
        Tbx_10_Target__rent.Text = Tbx2Dec(Tbx_10_Target__rent.Text)
    End Sub
    Private Sub Tbx_Target__heating_Leave(sender As Object, e As EventArgs) Handles Tbx_10_Target__heating.Leave
        Tbx_10_Target__heating.Text = Tbx2Dec(Tbx_10_Target__heating.Text)
    End Sub
    Private Sub Tbx_Target__gaselectra_Leave(sender As Object, e As EventArgs) Handles Tbx_10_Target__gaselectra.Leave
        Tbx_10_Target__gaselectra.Text = Tbx2Dec(Tbx_10_Target__gaselectra.Text)
    End Sub
    Private Sub Tbx_Target__water_Leave(sender As Object, e As EventArgs) Handles Tbx_10_Target__water.Leave
        Tbx_10_Target__water.Text = Tbx2Dec(Tbx_10_Target__water.Text)
    End Sub
    Private Sub Tbx_Target__food_Leave(sender As Object, e As EventArgs) Handles Tbx_10_Target__food.Leave
        Tbx_10_Target__food.Text = Tbx2Dec(Tbx_10_Target__food.Text)
    End Sub
    Private Sub Tbx_Target__medicine_Leave(sender As Object, e As EventArgs) Handles Tbx_10_Target__medicine.Leave
        Tbx_10_Target__medicine.Text = Tbx2Dec(Tbx_10_Target__medicine.Text)
    End Sub

    Private Sub Tbx_Target__name_Leave(sender As Object, e As EventArgs) Handles Tbx_01_Target__name.Leave
        If Lbx_Basis.Items.Count <> 0 Then
            ind1 = Lbx_Basis.SelectedItem(Me.Lbx_Basis.ValueMember)
        End If

    End Sub

    Private Sub Tbx_Target__name_add_Leave(sender As Object, e As EventArgs) Handles Tbx_01_Target__name_add.Leave
        If Lbx_Basis.Items.Count <> 0 Then
            ind1 = Lbx_Basis.SelectedItem(Me.Lbx_Basis.ValueMember)
        End If
    End Sub

    Private Sub Tbx_CP__name_TextChanged(sender As Object, e As EventArgs) Handles Tbx_01_CP__name.TextChanged
        If Edit_Mode Then Manage_Buttons_Target(False, False, False, True, True, "Tbx_CP__name_TextChanged")
        reload = True
        If Lbx_Basis.Items.Count = 0 Then Add_Mode = True
    End Sub


    Private Sub Button4_Click(sender As Object, e As EventArgs)
        MsgBox(sender.ToString)

        Exit Sub
        Dim str As String = "one,two,three"
        Dim str2() As String = Split(str, ",")

        'Load_Listbox(Me.Lbx_Basis, "Select id, name FROM Bankacc WHERE name ilike '%%' AND active=True ORDER BY name")
        'MsgBox(Me.Dgv_Mgnt_Tables.Rows(1).Cells(0).Value)
        Dim amt1 = QuerySQL("select max(startbalance) from bankacc")


        Dim LocalFilePath As String = "C:\temp\lcal.html"
        Dim objWebClient As New System.Net.WebClient
        'objWebClient.DownloadFile("https://www.google.com/search?newwindow=1&sxsrf=ALeKk00tqujhzWGn2oO1UiVUC8hWGsGjvw%3A1596922685176&ei=PRsvX5KlCoHisAfjxqjQAQ&q=exchange+rate+eur+mdl&oq=exchange+ra&gs_lcp=CgZwc3ktYWIQARgAMgYIIxAnEBMyBQgAELEDMgIIADICCAAyAggAMgIIADICCAAyAggAMgIIADICCAA6BAgjECc6AgguOggIABCxAxCDAToECAAQAzoECAAQQzoHCAAQsQMQQzoICC4QsQMQgwFQ8f8YWIORGWCLnxloAHAAeACAAdoBiAGVCJIBBjExLjAuMZgBAKABAaoBB2d3cy13aXrAAQE&sclient=psy-ab", LocalFilePath)
        objWebClient.DownloadFile("https://eur.fxexchangerate.com/mdl-exchange-rates-history.html", LocalFilePath)

        Dim text As String = File.ReadAllText("C:\Temp\lcal.html")
        Dim index As Integer = text.IndexOf("<td>1 EUR =</td>")
        If index >= 0 Then
            MsgBox(Strings.Mid(text, index + 22, 8))
        Else
            MsgBox("Wisselkoers niet gevonden.")
        End If

    End Sub

    Private Sub navigation_complete(ByVal sender As System.Object,
           ByVal e As System.Windows.Forms.WebBrowserDocumentCompletedEventArgs)

        Dim HTMlAuthorCode As String = sender.DocumentText
        My.Computer.FileSystem.WriteAllText("C:\temp\xe.html", HTMlAuthorCode, True)

        Dim strAuthorCode As String = sender.Document.Body.InnerText
        My.Computer.FileSystem.WriteAllText("c:\temp\xe.txt", strAuthorCode, True)
        sender.Dispose()
    End Sub


    Private Sub Rbtn_Income_CheckedChanged(sender As Object, e As EventArgs) Handles Rbtn_Account_Income.CheckedChanged
        If MenuSave.Enabled Then Tbx_00_Account__type.Text = Rbtn_Account_Income.Text
    End Sub

    Private Sub Rbtn_Account_Transit_CheckedChanged(sender As Object, e As EventArgs) Handles Rbtn_Account_Transit.CheckedChanged
        If MenuSave.Enabled Then Tbx_00_Account__type.Text = Rbtn_Account_Transit.Text
    End Sub

    Private Sub Rbtn_Account_Expense_CheckedChanged(sender As Object, e As EventArgs) Handles Rbtn_Account_Expense.CheckedChanged
        If MenuSave.Enabled Then Tbx_00_Account__type.Text = Rbtn_Account_Expense.Text
    End Sub

    Private Sub Tbx_Account__type_TextChanged(sender As Object, e As EventArgs) Handles Tbx_00_Account__type.TextChanged
        Rbtn_Account_Income.Checked = Tbx_00_Account__type.Text = "Generiek (fonds)"
        Rbtn_Account_Expense.Checked = Tbx_00_Account__type.Text = "Specifiek (doel)"
        Rbtn_Account_Transit.Checked = Tbx_00_Account__type.Text = "Anders"
    End Sub
    Private Sub Tbx_BankAcc__accountno_Leave(sender As Object, e As EventArgs) Handles Tbx_01_BankAcc__accountno.Leave
        If Tbx_01_BankAcc__accountno.Text = "" Then Exit Sub
        Tbx_01_BankAcc__accountno.Text = Tbx_01_BankAcc__accountno.Text.ToUpper
        If IBANcheck(Tbx_01_BankAcc__accountno.Text) <> 1 Then
            MsgBox("Bankrekeningnummer Is niet correct", vbCritical)
            Tbx_01_BankAcc__accountno.Focus()
        End If
    End Sub


    Private Sub Tbx_10_Relation__name_TextChanged(sender As Object, e As EventArgs)
        If Edit_Mode Then
            Manage_Buttons_Target(False, False, False, True, True, "Tbx_10_Relation__name_TextChanged")
            reload = True
        End If

        If Add_Mode Then Generate_Reference()
        If Lbx_Basis.Items.Count = 0 Then Add_Mode = True
    End Sub
    Private Sub Tbx_00_Relation__iban_Leave(sender As Object, e As EventArgs) Handles Tbx_00_Relation__iban.Leave
        If Tbx_00_Relation__iban.Text = "" Then Exit Sub
        Tbx_00_Relation__iban.Text = Tbx_00_Relation__iban.Text.ToUpper
        If IBANcheck(Tbx_00_Relation__iban.Text) <> 1 Then
            MsgBox("Bankrekeningnummer Is niet correct", vbCritical)
            Tbx_00_Relation__iban.Focus()
        End If
    End Sub
    Private Sub Rbn_00_contract_child_Click(sender As Object, e As EventArgs) Handles Rbn_00_contract_child.Click
        Tbx_Contract_ttype.Text = "Kind"
        Lbl_00_Contract__name.Text = Contract_number("K")
        Load_Combobox(Cmx_01_contract__fk_target_id, "id", "name", "Select id, Name||', '||name_add as name FROM target
                                                        WHERE ttype='" & Rbn_00_contract_child.Text & "' ORDER BY name")
        '-------standaard_waarden ophalen
        Collect_data("select value from settings where label ilike 'standaard_%_kind' order by label")
        Tbx_11_Contract__donation.Text = dst.Tables(0).Rows(0)(0)
        Tbx_11_contract__overhead.Text = dst.Tables(0).Rows(1)(0)

        '----------------------------

    End Sub
    Private Sub Tbx_10_Contract__transport_TextChanged(sender As Object, e As EventArgs)
        If Edit_Mode Then
            Manage_Buttons_Target(False, False, False, True, True, "Tbx_10_Contract__transport_TextChanged")
        End If
        'Pan_Contract_Date_New.Visible = Not Add_Mode
        Calculate_contract_amounts()
    End Sub
    Private Sub Tbx_11_Contract__donation_TextChanged(sender As Object, e As EventArgs) Handles Tbx_11_Contract__donation.TextChanged
        If Edit_Mode Then
            Manage_Buttons_Target(False, False, False, True, True, "Tbx_11_Contract__donation_TextChanged")
        End If

        'Pan_Contract_Date_New.Visible = Not Add_Mode
        Calculate_contract_amounts()
    End Sub

    Private Sub Tbx_11_contract__overhead_TextChanged(sender As Object, e As EventArgs) Handles Tbx_11_contract__overhead.TextChanged
        If Edit_Mode Then
            Manage_Buttons_Target(False, False, False, True, True, "Tbx_11_contract__overhead_TextChanged")
        End If
        Calculate_contract_amounts()
    End Sub
    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs)
        Calculate_contract_amounts()
    End Sub

    Private Sub Cbx_00_BankAcc__active_CheckedChanged(sender As Object, e As EventArgs) Handles _
            Cbx_00_BankAcc__active.CheckedChanged, Rbtn_Account_Income.CheckedChanged,
            Rbtn_Account_Expense.CheckedChanged, Rbtn_Account_Transit.CheckedChanged
        If Edit_Mode Then Manage_Buttons_Target(False, False, False, True, True, "Cbx_00_BankAcc__active_CheckedChanged")
    End Sub

    Private Sub Cbx_00_Account__active_CheckedChanged(sender As Object, e As EventArgs) Handles Cbx_00_Account__active.CheckedChanged
        If Edit_Mode Then Manage_Buttons_Target(False, False, False, True, True, "Cbx_00_Account__active_CheckedChanged")
    End Sub

    Private Sub Cbx_00_cp__active_CheckedChanged(sender As Object, e As EventArgs) Handles Cbx_00_cp__active.CheckedChanged
        If Edit_Mode Then Manage_Buttons_Target(False, False, False, True, True, "Cbx_00_cp__active_CheckedChanged")


    End Sub

    Private Sub Cbx_00_relation__active_CheckedChanged(sender As Object, e As EventArgs) Handles Cbx_00_relation__active.CheckedChanged
        'If Edit_Mode Then Manage_Buttons_Target(False, False, False, True, True)
    End Sub
    Private Sub Pic_cp__photo_DoubleClick(sender As Object, e As EventArgs) Handles Pic_cp__photo.DoubleClick
        Save_Image(Pic_cp__photo)
    End Sub

    Private Sub Pic_Target__photo_DoubleClick(sender As Object, e As EventArgs) Handles Pic_Target__photo.DoubleClick
        Save_Image(Pic_Target__photo)
    End Sub

    Private Sub Tbx_11_Contract__donation_Leave(sender As Object, e As EventArgs) Handles Tbx_11_Contract__donation.Leave
        Tbx_11_Contract__donation.Text = Tbx2Dec(Tbx_11_Contract__donation.Text)
    End Sub

    Private Sub Tbx_11_contract__overhead_Leave(sender As Object, e As EventArgs) Handles Tbx_11_contract__overhead.Leave
        Tbx_11_contract__overhead.Text = Tbx2Dec(Tbx_11_contract__overhead.Text)
    End Sub

    Private Sub Cmx_01_contract_fk_target_id_Leave(sender As Object, e As EventArgs) Handles Cmx_01_contract__fk_target_id.Leave
        If (Cmx_01_contract__fk_target_id.SelectedIndex = -1) Then
            Cmx_01_contract__fk_target_id.Focus()
            Exit Sub
        End If
        Exit Sub
        Dim id = Cmx_01_contract__fk_target_id.SelectedValue
        Try
            Pic_Contract_Target_photo.Image = BlobToImage(QuerySQL("SELECT photo FROM target WHERE id='" & id & "'"))

        Catch ex As Exception
            Pic_Contract_Target_photo.Image = Nothing
        End Try
    End Sub

    Private Sub Cmx_01_contract_fk_target_id_SelectedValueChanged(sender As Object, e As EventArgs) Handles Cmx_01_contract__fk_target_id.SelectedValueChanged
        Dim id = Cmx_01_contract__fk_target_id.SelectedValue
        'Tbx_Contract_ttype.Text = QuerySQL("Select ttype FROM target WHERE id=" & id)
        Try
            Pic_Contract_Target_photo.Image = BlobToImage(QuerySQL("SELECT photo FROM target WHERE id='" & id & "'"))

        Catch ex As Exception
            Pic_Contract_Target_photo.Image = Nothing
        End Try
    End Sub

    Private Sub Cmx_01_contract_fk_relation_id_Leave(sender As Object, e As EventArgs) Handles Cmx_00_contract__fk_relation_id.Leave
        If (Cmx_00_contract__fk_relation_id.SelectedIndex = -1) Then
            Cmx_00_contract__fk_relation_id.Focus()
            Exit Sub
        End If

        Get_Sponsor_data()
    End Sub

    Private Sub Dtp_01_contract__enddate_Enter(sender As Object, e As EventArgs) Handles Dtp_31_contract__enddate.Enter
        oldend_date = Dtp_31_contract__enddate.Value
        Dim newEndDate As Date = Date.Today.AddMonths(1)
        'Me.Dtp_Incasso_start.Value = CDate("01-" & newDate.Month & "-" & newDate.Year)
        Edit_Mode = True
        If Not Add_Mode Then Dtp_31_contract__enddate.Value = New DateTime(newEndDate.Year, newEndDate.Month, 1).AddDays(-1) 'end of current month
    End Sub

    Private Sub Rbn_00_contract_elder_Click(sender As Object, e As EventArgs) Handles Rbn_00_contract_elder.Click
        Tbx_Contract_ttype.Text = "Oudere"
        Lbl_00_Contract__name.Text = Contract_number("O")
        Load_Combobox(Cmx_01_contract__fk_target_id, "id", "name", "SELECT id, name||', '||name_add as name FROM target
                                                        WHERE ttype='" & Rbn_00_contract_elder.Text & "' ORDER BY name")
        Collect_data("select value from settings where label ilike 'standaard_%_oudere' order by label")
        Tbx_11_Contract__donation.Text = dst.Tables(0).Rows(0)(0)
        Tbx_11_contract__overhead.Text = dst.Tables(0).Rows(1)(0)
    End Sub

    Private Sub Rbn_00_contract_other_Click(sender As Object, e As EventArgs) Handles Rbn_00_contract_other.Click
        Tbx_Contract_ttype.Text = "Overig"
        Load_Combobox(Cmx_01_contract__fk_target_id, "id", "name", "SELECT id, name||', '||name_add as name FROM target
                                                        WHERE ttype='" & Rbn_00_contract_other.Text & "' ORDER BY name")
        Lbl_00_Contract__name.Text = Contract_number("V")
        Tbx_11_Contract__donation.Text = 0
        Tbx_11_contract__overhead.Text = 0
    End Sub
    Sub Check_Contract_Status()
        'check that contract is not already ended or has a newer version
        Dim sd As Date = QuerySQL("SELECT MAX(startdate) FROM contract 
                                       WHERE name='" & Lbl_00_Contract__name.Text & "'")

        If Me.Dtp_31_contract__enddate.Value < Date.Today Or  '@@@eigenlijk: de eerste dag van de volgende maand
            Me.Dtp_31_contract__enddate.Value < sd Then
            MsgBox("Een contract dat beeindigd is of niet de laatste versie is kan niet gewijzigd worden.")
            Select_Obj2("Check_Contract_Status")
            Manage_Buttons_Target(True, True, True, False, False, "Check_Contract_Status")
            Edit_Mode = False
            Add_Mode = False
            Pan_Target.Enabled = False

            If TC_Object.SelectedIndex = 0 Then  'additional functionality for contract management
                Handle_Contract_Fields()
                Pan_Contract_Date_New.Visible = False
            End If
            Exit Sub
        End If
    End Sub
    Private Sub Tbx_01_contract_yeartotal_TextChanged(sender As Object, e As EventArgs) Handles Tbx_01_contract_yeartotal.TextChanged

        If Edit_Mode And Not Add_Mode Then
            Dim name = Lbl_00_Contract__name.Text
            '0. Retrieve all versions of the contract 

            '1 determine whether there is a future version. If so then a change is not allowed (first
            'delete that version

            Dim next_date As Date = QuerySQL("select max(startdate) from contract where name ='" & name & "'")
            If next_date > Date.Today Then
                MsgBox("Er is een nieuwere versie die nog niet is ingegaan (" & next_date & ")" & vbCrLf &
                   "S.v.p. deze eerst verwijderen. " & next_date)
                Cancel()
                Exit Sub

            End If

            'Set new start dates
            Dim mindate As DateTime
            'a new version must start 1 month later 
            mindate = Dtp_31_contract__startdate.Value
            Dtp_30_Contract_Change.MinDate = New DateTime(mindate.Year, mindate.Month, 1).AddMonths(1)

            'set default new startdate to first day of the next month
            Dim m_add, m_year As Integer
            If Me.Dtp_30_Contract_Change.Value.Month = 12 Then
                m_add = -11
                m_year = 1
            Else
                m_add = IIf(Me.Dtp_30_Contract_Change.Value.Day > 1, 1, 0)
                m_year = 0
            End If
            Me.Dtp_30_Contract_Change.Value = CDate("01-" & Me.Dtp_30_Contract_Change.Value.Month +
            m_add & "-" & Me.Dtp_30_Contract_Change.Value.Year + m_year)

            Pan_Contract_Date_New.Visible = True

        End If
        If Add_Mode Then
            Pan_Contract_Date_New.Visible = False
        End If
    End Sub

    Private Sub Chx_00_contract__autcol_Click(sender As Object, e As EventArgs) Handles Chx_00_contract__autcol.Click
        'check if the sponsor provided an authorization for automatic collection
        If Chx_00_contract__autcol.Checked = False Then
            Dim SQLstr As String = "UPDATE contract SET autcol=False WHERE id=" & Lbl_Contract_pkid.Text
            If Chbx_test.Checked Then MsgBox(SQLstr)
            RunSQL(SQLstr, "NULL", "")
            Exit Sub
        End If

        Dim dtp As String
        Dim ttype As String
        Dim rel_id = Cmx_00_contract__fk_relation_id.SelectedValue

        If Rbn_00_contract_child.Checked Then
            dtp = "date1"
            ttype = "kindersponsoring"
        ElseIf Rbn_00_contract_elder.Checked Then
            dtp = "date2"
            ttype = "ouderensponsoring"
        Else
            dtp = "date3"
            ttype = "algemene sponsoring"
        End If

        Dim autcol_date As Date = QuerySQL("SELECT " & dtp & " FROM relation WHERE id=" & rel_id)

        If autcol_date > Date.Now Then
            MsgBox("De sponsor heeft nog geen geldige incassomachtiging voor " & ttype &
                   "; Automatische incasso kan (nog) niet geactiveerd worden voor dit contract.", vbCritical)
            'Dim ans = MsgBox("Weet u zeker dat u automatische incasso wilt instellen?", vbYesNo)
            'If vbYes Then
            'RunSQL("UPDATE contract SET autcol=True WHERE id=" & Lbl_Contract_pkid.Text, "NULL", "")
            'Else
            Chx_00_contract__autcol.Checked = False
            'End If
        Else
            RunSQL("UPDATE contract SET autcol=True WHERE id=" & Lbl_Contract_pkid.Text, "NULL", "")
        End If

    End Sub
    Private Sub Cbx_00_contract__autcol_CheckedChanged(sender As Object, e As EventArgs) Handles Chx_00_contract__autcol.CheckedChanged
        Dim rel_id = Cmx_00_contract__fk_relation_id.SelectedValue
        Dim dtp = IIf(Rbn_00_contract_child.Checked, "date1", IIf(Rbn_00_contract_elder.Checked, "date2", "date3"))
        Lbl_00_contract_autcol.Visible = Chx_00_contract__autcol.Checked
        Lbl_00_contract_autcol.Text = QuerySQL("SELECT reference FROM relation WHERE id=" & rel_id)
        dtp_contract_relation_date.Visible = Chx_00_contract__autcol.Checked
        Lbl_contract_mach_datum.Visible = Chx_00_contract__autcol.Checked
        dtp_contract_relation_date.Value = QuerySQL("SELECT " & dtp & " FROM relation WHERE id=" & rel_id)
        Lbl_contract_macht_kenm.Visible = Chx_00_contract__autcol.Checked
        '@@@ 
    End Sub

    Private Sub Cbx_00_relation__active_Click(sender As Object, e As EventArgs) Handles Cbx_00_relation__active.Click
        CheckActive(Cbx_00_relation__active, Lbl_relation_pkid, "contract")
    End Sub

    Private Sub Cbx_00_target__active_Click(sender As Object, e As EventArgs) Handles Cbx_00_target__active.Click
        CheckActive(Cbx_00_target__active, Lbl_Target_pkid, "contract")
    End Sub

    Private Sub Dtp_30_Contract_Change_ValueChanged(sender As Object, e As EventArgs) Handles Dtp_30_Contract_Change.ValueChanged
        Exit Sub
        If Dtp_30_Contract_Change.Visible Then
            Dim d As DateTime
            d = Me.Dtp_30_Contract_Change.Value
            'Determine enddate previous version
            Me.Dtp_31_contract__enddate.Value = New DateTime(d.Year, d.Month, 1).AddDays(-1)

            Add_Mode = True
        End If
    End Sub


    Private Sub Tbx_01_Account__name_Enter(sender As Object, e As EventArgs) Handles _
        Tbx_10_Target__income.Enter, Tbx_10_Target__pension.Enter, Tbx_10_Target__benefit.Enter,
        Tbx_10_Target__allowance.Enter, Tbx_10_Target__otherincome.Enter, Tbx_10_Target__rent.Enter,
        Tbx_10_Target__heating.Enter, Tbx_10_Target__pension.Enter, Tbx_10_Target__gaselectra.Enter,
        Tbx_10_Target__water.Enter, Tbx_10_Target__food.Enter, Tbx_10_Target__medicine.Enter,
        Cmx_01_Target__fk_cp_id.Click, Dtp_00_Target__birthday.Click,
        Tbx_01_Target__name_add.Enter, Tbx_01_Target__name.Enter, Tbx_00_Target__zip.Enter,
        Tbx_00_Target__address.Enter, Tbx_00_Target__city.Enter, Tbx_00_Target__country.EnabledChanged, Tbx_00_Target__country.Enter,
        Tbx_00_Target__description.Enter, Tbx_01_CP__name.Enter, Tbx_01_CP__name_add.Enter,
        Tbx_01_BankAcc__accountno.Enter, Tbx_01_BankAcc__name.Enter, Tbx_01_BankAcc__owner.Enter,
        Tbx_BankAcc_startbalance.Enter,
        Tbx_01_Account__name.Enter, Cmx_00_Account__accgroup.Enter, Cmx_01_account__fk_accgroup_id.Enter,
        Tbx_10_Account__b_jan.Enter, Tbx_10_Account__b_feb.Enter, Tbx_10_Account__b_mar.Enter, Tbx_10_Account__b_apr.Enter, Tbx_10_Account__b_may.Enter, Tbx_10_Account__b_jun.Enter,
        Tbx_10_Account__b_jul.Enter, Tbx_10_Account__b_aug.Enter, Tbx_10_Account__b_sep.Enter, Tbx_10_Account__b_oct.Enter, Tbx_10_Account__b_nov.Enter, Tbx_10_Account__b_dec.Enter,
        Cbx_00_BankAcc__income.Enter,
        Tbx_00_contract__description.Enter, Tbx_00_BankAcc__id2.Enter, Tbx_00_BankAcc__bic.Enter,
        Chx_00_BankAcc__expense.Enter, Cmx_01_BankAcc__currency.Enter, Tbx_00_BankAcc__description.Enter,
        Rbtn_Account_Income.Enter, Rbtn_Account_Expense.Enter, Rbtn_Account_Transit.Enter,
        Cmx_01_cp__fk_bankacc_id.Enter, Cmx_01_account__fk_accgroup_id.Click, Tbx_00_Account__description.Enter, Tbx_00_Account__searchword.Enter,
        Dtp_00_relation__date1.Enter, Dtp_00_relation__date2.Enter, Dtp_00_relation__date3.Enter, Tbx_00_Account__bankcode.Enter,
        Dtp_31_contract__startdate.Enter, Tbx_00_cp__description.Enter, Cmx_01_cp__fk_bankacc_id.Click, Tbx_00_CP__telephone.Enter,
        Tbx_00_Relation__iban.Enter, Tbx_11_Contract__donation.Enter, Tbx_11_contract__overhead.Enter, Cbx_00_BankAcc__active.Enter,
        Cbx_00_Account__active.Enter, Cbx_00_cp__active.Enter, Tbx_00_CP__telephone.Enter, Tbx_00_CP__address.Enter, Tbx_00_CP__zip.Enter,
        Tbx_00_CP__city.Enter, Tbx_00_CP__country.Enter, Tbx_00_CP__email.Enter,
        Tbx_00_Accgroup__subtype.Enter, Tbx_00_Accgroup__description.Enter, Tbx_01_Accgroup__name.Enter, Tbx_01_Accgroup__type.Enter, Cmx_00_Account__accgroup.Click
        ',        Rbtn_accgroup_Income.Click, Rbtn_accgroup_expense.Click, Rbtn_accgroup_transit.Click

        Edit_Mode = True
    End Sub
    Sub Manage_Buttons_Target(ByVal _add As Boolean, _searchbox As Boolean, d As Boolean, _menusave As Boolean, _cancel As Boolean, sender As String)

        If Cbx_LifeCycle.Text = "Inactief" And Edit_Mode Then
            MsgBox("Inactieve objecten kunnen niet gewijzigd worden.")
            Exit Sub
        End If
        Lbx_Basis.Enabled = _add
        MenuAdd.Enabled = _add
        MenuDelete.Enabled = _add
        MenuFilter.Enabled = _searchbox
        Searchbox.Enabled = _searchbox
        Cbx_LifeCycle.Enabled = _searchbox
        MenuSave.Enabled = _menusave
        MenuCancel.Enabled = _cancel
    End Sub
    Private Sub Tbx_BankAcc__accountno_TextChanged(sender As Object, e As EventArgs) Handles _
          Tbx_01_BankAcc__accountno.TextChanged, Tbx_01_BankAcc__name.TextChanged, Tbx_01_Accgroup__name.TextChanged, Tbx_01_Target__name.TextChanged, Cmx_01_account__fk_accgroup_id.TextUpdate,
          Tbx_01_Target__name_add.TextChanged, Tbx_01_Account__name.TextChanged, Tbx_01_CP__name_add.TextChanged,
          Tbx_00_Accgroup__subtype.TextChanged, Tbx_00_Accgroup__description.TextChanged, Tbx_01_Accgroup__name.TextChanged, Tbx_01_Accgroup__type.TextChanged,
          Rbtn_accgroup_Income.CheckedChanged, Rbtn_accgroup_expense.CheckedChanged, Rbtn_accgroup_transit.CheckedChanged, Cmx_00_Account__accgroup.TextUpdate, Cmx_00_Account__accgroup.SelectedValueChanged



        If Edit_Mode Then Manage_Buttons_Target(False, False, False, True, True, "Tbx_BankAcc__accountno_TextChanged")
        reload = True
    End Sub

    Private Sub Tbx_BankAcc__owner_TextChanged(sender As Object, e As EventArgs) Handles Tbx_01_BankAcc__owner.TextChanged,
Tbx_00_BankAcc__id2.TextChanged, Tbx_00_BankAcc__bic.TextChanged, Tbx_BankAcc_startbalance.TextChanged,
Tbx_00_BankAcc__description.TextChanged, Cbx_00_BankAcc__income.CheckedChanged, Chx_00_BankAcc__expense.CheckedChanged,
Cmx_01_BankAcc__currency.SelectedIndexChanged, Cmx_01_cp__fk_bankacc_id.SelectedIndexChanged,
Cmx_01_account__fk_accgroup_id.SelectedIndexChanged,
Tbx_00_cp__description.TextChanged, Tbx_00_Account__bankcode.TextChanged, Dtp_00_relation__date1.ValueChanged, Dtp_00_relation__date2.ValueChanged,
Dtp_00_relation__date3.ValueChanged, Tbx_00_Target__address.TextChanged, Dtp_31_contract__startdate.TextChanged,
Tbx_00_Target__zip.TextChanged, Tbx_00_Target__city.TextChanged, Tbx_00_Target__country.TextChanged, Tbx_00_Target__description.TextChanged,
Dtp_00_Target__birthday.ValueChanged, Cmx_01_Target__fk_cp_id.SelectedIndexChanged, Tbx_00_Account__description.TextChanged, Tbx_00_Account__searchword.TextChanged,
Tbx_00_Relation__iban.TextChanged, Tbx_00_contract__description.TextChanged, Dtp_31_contract__enddate.ValueChanged,
Tbx_00_CP__telephone.TextChanged, Tbx_00_CP__address.TextChanged, Tbx_00_CP__zip.TextChanged,
Tbx_00_CP__city.TextChanged, Tbx_00_CP__country.TextChanged, Tbx_00_CP__email.TextChanged, Cmx_00_Account__accgroup.SelectedIndexChanged


        If Edit_Mode Then Manage_Buttons_Target(False, False, False, True, True, "Tbx_BankAcc__owner_TextChanged")
    End Sub

    Private Sub Tbx_10_Account__b_jan_TextChanged(sender As Object, e As EventArgs) Handles _
        Tbx_10_Account__b_jan.TextChanged, Tbx_10_Account__b_feb.TextChanged, Tbx_10_Account__b_mar.TextChanged,
        Tbx_10_Account__b_apr.TextChanged, Tbx_10_Account__b_may.TextChanged, Tbx_10_Account__b_jun.TextChanged,
        Tbx_10_Account__b_jul.TextChanged, Tbx_10_Account__b_aug.TextChanged, Tbx_10_Account__b_sep.TextChanged,
        Tbx_10_Account__b_oct.TextChanged, Tbx_10_Account__b_nov.TextChanged, Tbx_10_Account__b_dec.TextChanged

        If Edit_Mode Then
            Manage_Buttons_Target(False, False, False, True, True, "Tbx_10_Account__b_jan_TextChanged")
            Calculate_Manual_Budgets()
        End If
    End Sub

    Sub Cmx_Bank_bankacc_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Cmx_Bank_bankacc.SelectedIndexChanged
        'MsgBox("A:" & Cmx_Bank_bankacc.SelectedIndex)
        If Cmx_Bank_bankacc.SelectedIndex = -1 Then Cmx_Bank_bankacc.SelectedIndex = 0
        Fill_bank_transactions("Cmx_Bank_bankacc.SelectedIndexChanged")

    End Sub
    Sub Dgv_Bank_Click(sender As Object, e As EventArgs) Handles Dgv_Bank.Click, Dgv_Bank.SelectionChanged

        If Dgv_Bank.Rows.Count = 0 Or Dgv_Bank.DataSource Is Nothing Then Exit Sub




        Try
            If Not IsDBNull(Dgv_Bank.SelectedCells(3).Value) Then
                If Strings.Left(Dgv_Bank.SelectedCells(3).Value, 16) = "Contract incasso" _
                    Or Strings.Left(Dgv_Bank.SelectedCells(3).Value, 7) = "Excasso" Then
                    Dgv_Bank_Account.EditMode = DataGridViewEditMode.EditProgrammatically
                Else
                    Dgv_Bank_Account.EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2
                End If
            End If

            Tbx_Bank_Relation.Text = Dgv_Bank.SelectedCells(2).Value
            Tbx_Bank_Description.Text = Dgv_Bank.SelectedCells(3).Value
            If Chbx_Bank_ExtraInfo_voor.Checked Then
                'MsgBox(Tbx_Bank_Description.Text & "---" & Strings.InStr(Tbx_Bank_Afschrift.Text, " | "))

                If Strings.InStr(Tbx_Bank_Description.Text, " | ") > 0 Then
                    Tbx_Bank_Extra_Info.Text = Strings.Left(Tbx_Bank_Description.Text, Strings.InStr(Tbx_Bank_Description.Text, " | ") - 1)
                Else
                    Tbx_Bank_Extra_Info.Text = ""
                End If
            End If


            If Not IsDBNull(Dgv_Bank.SelectedCells(8).Value) Then Tbx_Bank_Relation_account.Text = Dgv_Bank.SelectedCells(8).Value
            If Not IsDBNull(Dgv_Bank.SelectedCells(6).Value) Then
                Tbx_Bank_Code.Text = Dgv_Bank.SelectedCells(6).Value

            End If
            If Not IsDBNull(Dgv_Bank.SelectedCells(9).Value) Then
                Tbx_Bank_Afschrift.Text = Dgv_Bank.SelectedCells(9).Value
                Dim qry As String = $"Select sum(credit-debit) from bank where seqorder ='{Tbx_Bank_Afschrift.Text}' "
                Tbx_Transactie_totaal.Text = QuerySQL(qry)
                'MsgBox(qry)
            Else
                MsgBox(Dgv_Bank.SelectedCells(9).Value)
            End If



            'Dim ink As Boolean = Dgv_Bank.SelectedCells(4).Value > 0 And
            'Strings.Left(Dgv_Bank.SelectedCells(3).Value, 8) <> "Contract"

            Fill_Journals_by_bank(Dgv_Bank.SelectedCells(0).Value)


            If Dgv_Bank.Rows(Dgv_Bank.SelectedCells(2).RowIndex).DefaultCellStyle.ForeColor = Color.DarkRed And Trim(Tbx_Bank_Code.Text) = "cb" Then

                Dim sqlstr = "
                Select ac.name From account ac
                Left Join target t on t.id = ac.f_key And source='Doel'
                Left Join contract c on c.fk_target_id = t.id
                Left Join relation r on r.id = c.fk_relation_id
                Where R.iban = '" & Tbx_Bank_Relation_account.Text & "' 
                And R.active = True limit 1
                "
                Cmx_Bank_Account.Text = QuerySQL(sqlstr)

            Else
                Cmx_Bank_Account.Text = ""

            End If
            Tbx_Bank_Amount.Text = 0

            For x = 0 To Dgv_Bank_Account.Rows.Count - 1
                If Dgv_Bank_Account.Rows(x).Cells(1).Value = "[Niet toegewezen]" Then
                    Tbx_Bank_Amount.Text = Dgv_Bank_Account.Rows(x).Cells(2).Value
                    Exit For
                End If
            Next x
            If Dgv_Bank.Rows(Dgv_Bank.SelectedCells(2).RowIndex).Cells(12).Value = "Auto-cat" Then
                RunSQL("Update Bank set fk_journal_name='Bank' where id='" & Dgv_Bank.SelectedCells(0).Value & "'", "NULL", "auto_cat")
                Dgv_Bank.Rows(Dgv_Bank.SelectedCells(2).RowIndex).DefaultCellStyle.ForeColor = Color.DarkGreen
                Dgv_Bank.Rows(Dgv_Bank.SelectedCells(2).RowIndex).Cells(12).Value = "Bank" '
            End If
        Catch ex As Exception
            'MsgBox(ex.ToString)
        End Try

    End Sub

    Private Sub Btn_Bank_Add_Journal_Click(sender As Object, e As EventArgs) Handles Btn_Bank_Add_Journal.Click ', Cmx_Bank_Account.SelectedValueChanged
        'Exit Sub
        'Cmx_Bank_Account.SelectedIndexChanged,
        If Check_Change_Bank_Categories(True) = False Then Exit Sub
        If (Not Rbn_Bank_jtype_con.Checked And Not Rbn_Bank_jtype_ext.Checked And Not Rbn_Bank_jtype_int.Checked) And Pan_Bank_jtype.Visible Then
            MsgBox("Selecteer eerst of dit een contractgift, extra gift of een andere banktransactie betreft")
            'Exit Sub
        End If

        If Cmx_Bank_Account.Text = "" Or (Not IsNumeric(Tbx_Bank_Amount.Text)) Or Tbx_Bank_Amount.Text = "" Or Cmx_Bank_Account.SelectedIndex = -1 Then
            MsgBox("Ongeldige invoer")
            Exit Sub
        Else
            If Cmx_Bank_Account.SelectedValue = QuerySQL("Select value from settings where label='nocat'") Then Exit Sub
            Dim R As DataRow
            R = dstbank.Tables(0).Rows.Add
            R(0) = Cmx_Bank_Account.SelectedValue
            R(1) = Cmx_Bank_Account.Text
            R(2) = Tbx_Bank_Amount.Text

            Calculate_Total_Booked("Btn_Bank_Add_Journal_Click")
            Save_Banktransaction_Accounts()
            Update_Category_Status()
        End If
    End Sub

    Private Sub Dgv_Test_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles _
        Dgv_Bank_Account.CellValueChanged ', Dgv_Bank_Account.Leave
        If Dgv_Bank_Account.Rows.Count = 0 Then  'dit kan alleen voorkomen als er een error is opgetreden. 
            MsgBox("Er is een fout opgetreden, u kunt wel doorgaan")
            Exit Sub
        End If

        Try
            If IsDBNull(Dgv_Bank_Account.CurrentCell.Value) Then
                MsgBox("Ongeldige invoer")
                Exit Sub
            End If
        Catch
            Exit Sub
        End Try


        If Not IsNumeric(Dgv_Bank_Account.CurrentCell.Value) Then
            MsgBox("Ongeldige invoer")
            Exit Sub
        End If

        If Check_Change_Bank_Categories(True) = False Then Exit Sub

        Calculate_Total_Booked("Dgv_Test_CellValueChanged")
        Save_Banktransaction_Accounts()
        Update_Category_Status()

    End Sub

    Private Sub Tbx_Bank_Search_TextChanged(sender As Object, e As EventArgs) Handles Tbx_Bank_Search.TextChanged
        Fill_bank_transactions("Tbx_Bank_Search.TextChanged")
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs)
        RunSQL("TRUNCATE TABLE bank", "NULL", "")
        RunSQL("Delete From journal WHERE source='Bank'", "NULL", "")
    End Sub

    Private Sub Rbn_Relation_1_CheckedChanged(sender As Object, e As EventArgs) Handles Rbn_Relation_1.Click
        Edit_Mode = True
        Manage_Buttons_Target(False, False, False, True, True, "Rbn_Relation_1_CheckedChanged")
        If Rbn_Relation_1.Checked Then Tbx_01_Relation__title.Text = Rbn_Relation_1.Text
    End Sub

    Private Sub Rbn_Relation_2_CheckedChanged(sender As Object, e As EventArgs) Handles Rbn_Relation_2.Click
        Edit_Mode = True
        Manage_Buttons_Target(False, False, False, True, True, "Rbn_Relation_2_CheckedChanged")
        If Rbn_Relation_2.Checked Then Tbx_01_Relation__title.Text = Rbn_Relation_2.Text
    End Sub

    Private Sub Rbn_Relation_3_CheckedChanged(sender As Object, e As EventArgs) Handles Rbn_Relation_3.Click
        Edit_Mode = True
        Manage_Buttons_Target(False, False, False, True, True, "Rbn_Relation_3_CheckedChanged")
        If Rbn_Relation_3.Checked Then Tbx_01_Relation__title.Text = Rbn_Relation_3.Text
    End Sub

    Private Sub Rbn_Relation_4_CheckedChanged(sender As Object, e As EventArgs) Handles Rbn_Relation_4.Click
        Edit_Mode = True
        Manage_Buttons_Target(False, False, False, True, True, "Rbn_Relation_4_CheckedChanged")
        If Rbn_Relation_4.Checked Then Tbx_01_Relation__title.Text = Rbn_Relation_4.Text
    End Sub
    Private Sub Rbn_Relation_5_CheckedChanged(sender As Object, e As EventArgs) Handles Rbn_Relation_5.Click
        Edit_Mode = True
        Manage_Buttons_Target(False, False, False, True, True, "Rbn_Relation_5_CheckedChanged")
        If Rbn_Relation_5.Checked Then Tbx_01_Relation__title.Text = Rbn_Relation_5.Text
    End Sub
    Private Sub Rbn_Relation_6_CheckedChanged(sender As Object, e As EventArgs) Handles Rbn_Relation_6.Click
        Edit_Mode = True
        Manage_Buttons_Target(False, False, False, True, True, "Rbn_Relation_6_CheckedChanged")
        If Rbn_Relation_6.Checked Then Tbx_01_Relation__title.Text = Rbn_Relation_6.Text
    End Sub

    Private Sub Tbx_01_Relation__title_TextChanged(sender As Object, e As EventArgs) Handles Tbx_01_Relation__title.TextChanged
        Rbn_Relation_1.Checked = Strings.Trim(Tbx_01_Relation__title.Text) = Rbn_Relation_1.Text
        Rbn_Relation_2.Checked = Strings.Trim(Tbx_01_Relation__title.Text) = Rbn_Relation_2.Text
        Rbn_Relation_3.Checked = Strings.Trim(Tbx_01_Relation__title.Text) = Rbn_Relation_3.Text
        Rbn_Relation_4.Checked = Strings.Trim(Tbx_01_Relation__title.Text) = ""
        Rbn_Relation_5.Checked = Strings.Trim(Tbx_01_Relation__title.Text) = Rbn_Relation_5.Text
        Rbn_Relation_6.Checked = Strings.Trim(Tbx_01_Relation__title.Text) = Rbn_Relation_6.Text

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles Tbx_Contract_ttype.TextChanged

        Dim rel_id = Cmx_00_contract__fk_relation_id.SelectedValue
        Dim dtp = IIf(Tbx_Contract_ttype.Text = "Kind", "date1",
                     IIf(Tbx_Contract_ttype.Text = "Oudere", "date2", "date3"))

        If Rbn_00_contract_child.Checked Then
            dtp = "date1"
        ElseIf Rbn_00_contract_elder.Checked Then
            dtp = "date2"
        Else
            dtp = "date3"
        End If

        dtp_contract_relation_date.Value = QuerySQL("SELECT " & dtp & " FROM relation WHERE id=" & rel_id)

    End Sub

    Private Sub Dtp_Incasso_start_ValueChanged(sender As Object, e As EventArgs) Handles Dtp_Incasso_start.ValueChanged
        If TC_Main.SelectedIndex <> 2 Then Exit Sub
        Create_Incassolist()
        If Rbn_Incasso_SEPA.Checked Then Format_dvg_incasso() Else Format_dvg_incasso_bookings()
    End Sub

    Private Sub Rbn_Incasso_SEPA_CheckedChanged(sender As Object, e As EventArgs) Handles Rbn_Incasso_SEPA.CheckedChanged
        If TC_Main.SelectedIndex <> 2 Then Exit Sub
        If Rbn_Incasso_SEPA.Checked Then
            Load_Datagridview(Me.Dgv_Incasso, Create_Incasso(Dtp_Incasso_start.Value), "Dtp_Incasso_start.ValueChanged")
            Format_dvg_incasso()
        Else
            Load_Datagridview(Me.Dgv_Incasso, Create_Incasso_Bookings(Dtp_Incasso_start.Value), "Dtp_Incasso_start.ValueChanged")
            Format_dvg_incasso_bookings()
        End If
    End Sub



    Private Sub Cbx_Uitkering_Kind_Click(sender As Object, e As EventArgs) Handles Cbx_Uitkering_Kind.Click,
            Cbx_Uitkering_Oudere.Click, Cbx_Uitkering_Overig.Click

        If Not Cbx_Uitkering_Oudere.Checked And Not Cbx_Uitkering_Kind.Checked And Not Cbx_Uitkering_Overig.Checked Then
            Empty_Excasso_Window()
        Else
            Call_Excasso_form(sender)
            Calculate_CP_Allowance()
        End If

    End Sub

    Private Sub Dtp_Excasso_Start_ValueChanged(sender As Object, e As EventArgs) Handles Dtp_Excasso_Start.ValueChanged
        'Dtp_Excasso_Start.Value = CDate("01-" & Dtp_Excasso_Start.Value.Month & "-" & Dtp_Excasso_Start.Value.Year)
        Dtp_Excasso_Start.MaxDate = Date.Today


    End Sub

    Private Sub Dgv_Excasso2_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles Dgv_Excasso2.CellEndEdit

        'checks of ingevoerde waarden ook beschikbaar zijn

        Dim i As Integer = Me.Dgv_Excasso2.CurrentCell.RowIndex  'Me.Dgv_Excasso2.CurrentRow.Index
        Dim j As Decimal = Me.Dgv_Excasso2.CurrentCell.ColumnIndex
        'Dim fin As VariantType
        Dim ruimte_contract As Integer = Math.Max(Me.Dgv_Excasso2.Rows(i).Cells(2).Value, Me.Dgv_Excasso2.Rows(i).Cells(3).Value)
        Dim bedrag As Integer = Me.Dgv_Excasso2.CurrentCell.Value

        Dim int As Integer
        'If Not Integer.TryParse(bedrag, int) Then GoTo fin
        On Error GoTo fin
        Select Case j
            Case 6
                If Me.Dgv_Excasso2.Rows(i).Cells(6).Value > ruimte_contract Then
                    MsgBox("Uitkering is maximaal het beschikbare contractbedrag")
                    Me.Dgv_Excasso2.Rows(i).Cells(6).Value = ruimte_contract
                End If

            Case 7
                If Me.Dgv_Excasso2.Rows(i).Cells(7).Value > Me.Dgv_Excasso2.Rows(i).Cells(4).Value Then
                    MsgBox("Extra gift is hoger dan binnengekomen extra gift")
                    Me.Dgv_Excasso2.Rows(i).Cells(7).Value = Me.Dgv_Excasso2.Rows(i).Cells(4).Value
                End If
            Case 8
                If Me.Dgv_Excasso2.Rows(i).Cells(8).Value > Me.Dgv_Excasso2.Rows(i).Cells(5).Value Then
                    MsgBox("Interne gift is hoger dan interne boeking")
                    Me.Dgv_Excasso2.Rows(i).Cells(8).Value = Me.Dgv_Excasso2.Rows(i).Cells(5).Value
                End If
        End Select
        Calculate_Excasso_Totals2()

        Exit Sub
fin:
        MsgBox("Het formulier accepteert alleen uitkeringen in hele euro's")
    End Sub
    Private Sub Dgv_Excasso2_DataError(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) _
    Handles Dgv_Excasso2.DataError, Dgv_Bank_Account.DataError, Dgv_Bank_Account.DataError

        MsgBox("Ongeldige invoer")
        e.ThrowException = False

    End Sub



    Private Sub Tbx_Excasso_CP2_TextChanged(sender As Object, e As EventArgs) Handles Tbx_Excasso_CP2.TextChanged,
        Tbx_Excasso_CP3.TextChanged, Tbx_Excasso_CP1.TextChanged
        'Calculate_CP_Allowance()
        'Exit Sub
        'If Not IsNumeric(Tbx_Excasso_CP2) Then Exit Sub
        Try
            Me.Lbl_Excasso_CP_Totaal.Text = Tbx2Int(GetDouble(Me.Tbx_Excasso_CP1.Text) _
            + GetDouble(Me.Tbx_Excasso_CP2.Text) + GetDouble(Me.Tbx_Excasso_CP3.Text))
            Me.Lbl_Excasso_CP_Totaal_MDL.Text = (Tbx2Int(GetDouble(Me.Tbx_Excasso_CP1.Text) _
            + GetDouble(Me.Tbx_Excasso_CP2.Text) + GetDouble(Me.Tbx_Excasso_CP3.Text)) * Tbx2Dec(Tbx_Excasso_Exchange_rate.Text))
            Me.Lbl_Excasso_Tot_Gen.Text = CInt(Me.Lbl_Excasso_CP_Totaal.Text) + CInt(Lbl_Excasso_Totaal.Text)
            Lbl_Excasso_Tot_Gen_MLD.Text = Math.Round(CInt(Me.Lbl_Excasso_Tot_Gen.Text) * Tbx2Dec(Tbx_Excasso_Exchange_rate.Text), 2)
        Catch
            MsgBox("Geen geldige invoer.")
        End Try
    End Sub

    Private Sub GroupBox5_Leave(sender As Object, e As EventArgs)
        If IsNumeric(Tbx_Excasso_Exchange_rate.Text) Then

        Else
            MsgBox("Ongeldige inhoud")
        End If
    End Sub
    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Btn_Excasso_Exchrate.Click

        Btn_Excasso_Exchrate.Enabled = False
        Calculate_Excasso_Totals2()
    End Sub
    Private Sub Btn_Excasso_Print_Click(sender As Object, e As EventArgs) Handles Btn_Excasso_Print.Click
        If Cmx_Excasso_Select.SelectedIndex = -1 Then Exit Sub
        Print_Excasso_form()
    End Sub

    Private Sub Btn_Excasso_Delete_Click(sender As Object, e As EventArgs) Handles Btn_Excasso_Delete.Click
        MenuExcassoDelete()
    End Sub
    Sub MenuExcassoDelete()
        If Cmx_Excasso_Select.SelectedIndex = -1 Then Exit Sub
        If MsgBox("Wilt u de uitkeringslijst verwijderen?", vbYesNo) = vbYes Then
            RunSQL("DELETE FROM journal WHERE name ilike '%" & Me.Cmx_Excasso_Select.SelectedItem & "'", "NULL", "Delete_Excasso_Job")
            'RunSQL("DELETE FROM journal WHERE name='Intern tbv " & Me.Cmx_Excasso_Select.SelectedItem & "'", "NULL", "Delete_Excasso_Job")
            Fill_Cmx_Excasso_Select_Combined()

            'Calculate_CP_Allowance()
            Empty_Excasso_Window()
        End If

    End Sub

    Private Sub Btn_Excasso_CP_Calculate_Click(sender As Object, e As EventArgs) Handles Btn_Excasso_CP_Calculate.Click
        Calculate_CP_Allowance()
        'Btn_Excasso_CP_Calculate.Enabled = False
    End Sub

    Private Sub Btn_Excasso_Save_Click(sender As Object, e As EventArgs) Handles Btn_Excasso_Save.Click
        If Cmx_Excasso_Select.SelectedIndex = -1 Then Exit Sub
        Save_Excasso_job()
    End Sub
    Private Sub Tbx_Excasso_Norm1_Enter(sender As Object, e As EventArgs) Handles _
        Tbx_Excasso_Norm1.Enter, Tbx_Excasso_Norm2.Enter, Tbx_Excasso_Norm3.Enter
        Btn_Excasso_CP_Calculate.Enabled = True
    End Sub
    Private Sub Tbx_Excasso_Exchange_rate_Enter(sender As Object, e As EventArgs) Handles Tbx_Excasso_Exchange_rate.Enter
        Btn_Excasso_Exchrate.Enabled = True
    End Sub


    Private Sub Btn_Select_Bulk_Click(sender As Object, e As EventArgs) Handles Btn_Select_Bulk.Click
        Select_Target_Account()
    End Sub




    Private Sub Tbx_10_Account__b_jan_Leave(sender As Object, e As EventArgs) Handles _
            Tbx_10_Account__b_jan.Leave, Tbx_10_Account__b_feb.Leave, Tbx_10_Account__b_mar.Leave,
            Tbx_10_Account__b_apr.Leave, Tbx_10_Account__b_may.Leave, Tbx_10_Account__b_jun.Leave,
            Tbx_10_Account__b_jul.Leave, Tbx_10_Account__b_aug.Leave, Tbx_10_Account__b_sep.Leave,
            Tbx_10_Account__b_oct.Leave, Tbx_10_Account__b_nov.Leave, Tbx_10_Account__b_dec.Leave

        Calculate_Manual_Budgets()
    End Sub

    Private Sub Button9_Click_1(sender As Object, e As EventArgs) Handles Btn_Journal_Add_Source.Click
        Select_Source_Account()


    End Sub


    Private Sub Tbx_Journal_Source_Amt_TextChanged(sender As Object, e As EventArgs) Handles Tbx_Journal_Source_Amt.TextChanged
        Dim s As Decimal = Tbx2Dec(Me.Tbx_Journal_Source_Amt.Text)
        Dim m As Decimal = Tbx2Dec(Me.Lbl_Journal_Source_Saldo.Text)
        If (s <= 0 Or s > m) And (Tbx2Dec(Lbl_Journal_Source_Saldo.Text) <> 0) Then
            MsgBox("Bedrag (" & s & ") moet groter zijn dan nul en kleiner dan het saldo van de bronaccount (" & m & ")")
            Tbx_Journal_Source_Amt.Text = Tbx2Dec(m)
        End If
    End Sub

    Private Sub Tbx_Journal_Source_Amt_Leave(sender As Object, e As EventArgs) Handles Tbx_Journal_Source_Amt.Leave
        Calculate_Journal_Booking_Data()
    End Sub

    Private Sub Dgv_Journal_Intern_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles Dgv_Journal_Intern.CellEndEdit
        Dim i As Integer = Me.Dgv_Journal_Intern.CurrentRow.Index
        Dim s As Decimal = Me.Dgv_Journal_Intern.Rows(i).Cells(2).Value

        If s < 0 Then
            MsgBox("Doelbedrag mag niet negatief zijn.")
            Me.Dgv_Journal_Intern.Rows(i).Cells(2).Value = 0
        End If
        Calculate_Journal_Booking_Data()
    End Sub

    Sub Btn_Journals_Cancel_Click(sender As Object, e As EventArgs) Handles Btn_Journals_Cancel.Click
        Lbl_Journal_Source_Saldo.Text = 0
        Lbl_Journal_Source_Name.Text = ""
        Tbx_Journal_Source_Amt.Text = 0
        Dgv_Journal_Intern.Rows.Clear()
        Lbl_Journal_Source_Restamt.Text = 0

    End Sub

    Private Sub Btn_Journal_Recalculate_Click(sender As Object, e As EventArgs) Handles Btn_Journal_Recalculate.Click
        Divide_among_targets()
    End Sub

    Private Sub Btn_Journal_Intern_Save_Click(sender As Object, e As EventArgs) Handles Btn_Journal_Intern_Save.Click
        Save_Internal_Booking()
    End Sub

    Private Sub Lv_Journal_List_DoubleClick(sender As Object, e As EventArgs) Handles Lv_Journal_List.DoubleClick


        If TC_Boeking.SelectedIndex = 1 Then 'overboeking
            If Lbl_Journal_Source_Name.Text = "" Then
                Select_Source_Account()
            Else
                Select_Target_Account()
            End If

        End If

    End Sub

    Sub Btn_Account_Budget_Id_Click(sender As Object, e As EventArgs) Handles Btn_Account_Budget_Id.Click
        Calculate_Budget(Lbl_00_pkid.Text)
        Select_Obj2("Btn_Account_Budget_Id_Click")
    End Sub

    Private Sub Btn_Account_Budget_All_Click(sender As Object, e As EventArgs) Handles Btn_Account_Budget_All.Click
        Calculate_Budget("")
        Select_Obj2("Btn_Account_Budget_All_Click")

    End Sub
    Sub Create_Incassolist()


        Dim d As DateTime
        Dim t1 As String
        Dim t2 As String
        Dim newDate As Date = Date.Now.AddMonths(1)
        Dim maxDate As Date = Date.Now.AddMonths(2)
        Dim minDate1 As Date = Date.Now.AddMonths(-1)

        Me.Dtp_Incasso_start.MinDate = CDate("01-" & minDate1.Month & "-" & minDate1.Year)

        Me.Dtp_Incasso_start.Value = CDate("01-" & Me.Dtp_Incasso_start.Value.Month & "-" & Me.Dtp_Incasso_start.Value.Year)
        If Me.Dtp_Incasso_start.Value.Year <> Date.Today.Year Then
            Me.Dtp_Incasso_start.Value = CDate("01-" & newDate.Month & "-" & newDate.Year)

        End If


        d = Me.Dtp_Incasso_start.Value.AddMonths(1)
        Me.Dtp_Incasso_end.Value = New DateTime(d.Year, d.Month, 1).AddDays(-1)
        'Me.Dtp_Incasso_start.MinDate = New Date(minDate1.Year, 1, 1)
        Me.Dtp_Incasso_start.MaxDate = New Date(maxDate.Year, maxDate.Month, 1)
        'MsgBox(Dtp_Incasso_start.MinDate)


        Dim isd As Date = Me.Dtp_Incasso_start.Value
        Dim MsgId = "Contract incasso " & Month(isd) & "-" & Year(isd)
        Me.Lbl_Incasso_job_name.Text = MsgId
        Dim qtopen, qtverwerkt As Integer
        'Dim d1 As String = Year(Me.Dtp_Incasso_start.Value, "yyyy-MM-dd")
        'Dim d2 As Date = Format(Me.Dtp_Incasso_end.Value, "yyyy-MM-dd")
        t1 = Year(Me.Dtp_Incasso_start.Value) & "-" & Month(Me.Dtp_Incasso_start.Value) & "-01"
        t2 = Year(Me.Dtp_Incasso_end.Value) & "-" &
            Month(Me.Dtp_Incasso_end.Value) & "-" & Me.Dtp_Incasso_end.Value.Day

        'load lists and overview
        If Me.Rbn_Incasso_SEPA.Checked Then

            Load_Datagridview(Me.Dgv_Incasso, Create_Incasso(t1), "Me.Dtp_Incasso_start.ValueChanged")
            Format_dvg_incasso()
        Else
            Load_Datagridview(Me.Dgv_Incasso, Create_Incasso_Bookings(t1), "Me.Dtp_Incasso_start.ValueChanged")
            Format_dvg_incasso_bookings()

        End If
        Load_Listview(Me.Lv_Incasso_Overview, Create_Incasso_Totals(t1))
        Me.Lv_Incasso_Overview.Columns(0).Text = "Type"
        Me.Lv_Incasso_Overview.Columns(1).Text = "Aantal"
        Me.Lv_Incasso_Overview.Columns(2).Text = "Bedrag"
        Me.Lv_Incasso_Overview.Columns(1).TextAlign = HorizontalAlignment.Center
        Me.Lv_Incasso_Overview.Columns(2).TextAlign = HorizontalAlignment.Right
        Me.Lv_Incasso_Overview.Items(3).BackColor = Color.LightBlue

        Dim Tot_amt = CDec(Me.Lv_Incasso_Overview.Items(3).SubItems(2).Text).ToString("#,##")


        'Check_Existing_Incasso()
        Me.Lbl_Incasso_Error.Visible = False
        Dim journal_name As String = Me.Lbl_Incasso_job_name.Text
        qtopen = QuerySQL("select count(id) from journal where status = 'Open' and name ='" & journal_name & "'")
        qtverwerkt = QuerySQL("select count(id) from journal where status = 'Verwerkt' and name ='" & journal_name & "'")

        If qtopen > 0 Then
            Me.Lbl_Incasso_Status.Text = "Open"

            Me.Btn_Incasso_Delete.Enabled = True
            Me.Btn_Run_Incasso.Enabled = False
            Me.Btn_Incasso_Print.Enabled = True

            MenuDelete.Enabled = True
            MenuSave.Enabled = False
            Menu_Print.Enabled = True


            Dim Checksum = QuerySQL("Select Sum(amt1) from journal where name ='" & journal_name & "'")
            If Tot_amt <> Checksum Then
                Dim msg = "Er is een fout opgetreden: het berekende totaal (" & Tot_amt &
                    ") komt niet overeen met de eerder gecreëerde incassojob (" &
                    Checksum & "). Als deze incassojob nog niet naar de bank is verstuurd" &
                    " wordt u geadviseerd deze job te verwijderen en een nieuwe aan te maken " &
                    "voor deze maand."
                Me.Lbl_Incasso_Error.Text = msg
                Me.Lbl_Incasso_Error.Visible = True
            End If
        ElseIf qtverwerkt > 0 Then
            Me.Lbl_Incasso_Status.Text = "Verwerkt"

            Me.Btn_Incasso_Delete.Enabled = False
            Me.Btn_Run_Incasso.Enabled = False
            Me.Btn_Incasso_Print.Enabled = True

            MenuDelete.Enabled = False
            MenuSave.Enabled = False
            Menu_Print.Enabled = True

            Dim Checksum = QuerySQL("SELECT Sum(amt1) from journal where name ='" & journal_name & "'")
            If Tot_amt <> Checksum Then
                Me.Lbl_Incasso_Error.Text = "Opgeslagen incassojob is niet in lijn met contractdata"
            End If
        Else
            Me.Lbl_Incasso_Status.Text = "Nieuw"

            Me.Btn_Incasso_Delete.Enabled = False
            Me.Btn_Run_Incasso.Enabled = True
            Me.Btn_Incasso_Print.Enabled = False

            MenuDelete.Enabled = False
            MenuSave.Enabled = True
            Menu_Print.Enabled = False


        End If
        Format_dvg_incasso()
    End Sub
    Private Sub Cmx_Excasso_Select_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Cmx_Excasso_Select.SelectedIndexChanged
        Load_Excasso_Form()
        Call_Excasso_form(sender)

    End Sub

    Sub Load_Excasso_Form()
        If Cmx_Excasso_Select.SelectedIndex = -1 Then Exit Sub
        'check of de budgetbedragen nog geldig zijn
        If QuerySQL("select extract (year from min(date)) from journal") < Now.Year Then Calculate_Budget("")



        Gbx_Excasso_Calculate.Enabled = True
        If Strings.Left(Cmx_Excasso_Select.SelectedItem, 5) = "Nieuw" Then

            '=================nieuwe incasso ===============================================
            'determine CP id
            Btn_Excasso_Delete.Enabled = False
            Btn_Excasso_Print.Enabled = False
            Dim pos1 As Integer = Strings.InStr(Me.Cmx_Excasso_Select.SelectedItem, "[")

            Me.Lbl_Excasso_CPid.Text = Strings.Mid(Me.Cmx_Excasso_Select.SelectedItem, pos1 + 1,
                                       Len(Me.Cmx_Excasso_Select.SelectedItem) - pos1 - 1)

            Dtp_Excasso_Start.MaxDate = Date.Today
            Dtp_Excasso_Start.Value = Date.Today
            Dtp_Excasso_Start.Enabled = False
            Tbx_Excasso_Exchange_rate.Text = Tbx2Dec(My.Settings._exrate)
            Cbx_Uitkering_Kind.Enabled = True
            Cbx_Uitkering_Oudere.Enabled = True
            Cbx_Uitkering_Overig.Enabled = True
            Cbx_Uitkering_Kind.Checked = False
            Cbx_Uitkering_Oudere.Checked = False
            Cbx_Uitkering_Overig.Checked = False
            Dtp_Excasso_Start.Enabled = True
            Tbx_Excasso_CP1.Text = ""
            Tbx_Excasso_CP2.Text = ""
            Tbx_Excasso_CP3.Text = ""
            Load_Excasso_Balances()
            Lbl_Excasso_CP_Totaal.Text = 0
            Lbl_Excasso_CP_Totaal_MDL.Text = 0
            Calculate_CP_Allowance()


        Else  '=============================existing excasso===============================
            Btn_Excasso_Delete.Enabled = True
            Btn_Excasso_Print.Enabled = True
            Dtp_Excasso_Start.Enabled = False
            'determine CP id
            'Rbn_Excasso_Maandbudget.Checked = False
            Dim str1() As String = Split(QuerySQL("SELECT cpinfo FROM journal
                                               WHERE name ='" & Cmx_Excasso_Select.SelectedItem & "'
                                                "), "-")
            'cpinfo: cpid-Tbx_Excasso_Norm1- ..2-..3-Tbx_Excasso_CP1-..2-..3
            Lbl_Excasso_CPid.Text = str1(0)
            Tbx_Excasso_Norm1.Text = str1(1)
            Tbx_Excasso_Norm2.Text = str1(2)
            Tbx_Excasso_Norm3.Text = str1(3)
            Tbx_Excasso_CP1.Text = str1(4)
            Tbx_Excasso_CP2.Text = str1(5)
            Tbx_Excasso_CP3.Text = str1(6)

            Btn_Excasso_Base1.Text = IIf(str1(7) = "1", "€", "%")
            Btn_Excasso_Base2.Text = IIf(str1(8) = "1", "€", "%")
            Btn_Excasso_Base3.Text = IIf(str1(9) = "1", "€", "%")


            'calculate actual exchange rate
            Dim exr = QuerySQL("SELECT sum(amt2)/sum(amt1) FROM journal WHERE name ='" & Cmx_Excasso_Select.SelectedItem & "'")
            If IsDBNull(exr) Then exr = 0
            Tbx_Excasso_Exchange_rate.Text = Math.Round(GetDouble(exr), 2)
            ' determine date
            Dtp_Excasso_Start.Value = CDate(QuerySQL("SELECT date FROM journal WHERE name='" _
                                & Cmx_Excasso_Select.SelectedItem & "'"))
            Dtp_Excasso_Start.Enabled = False
            'determine target type

            Cbx_Uitkering_Kind.Enabled = False
            Cbx_Uitkering_Oudere.Enabled = False
            Cbx_Uitkering_Overig.Enabled = False

            Dim str2() As String = Split(Cmx_Excasso_Select.SelectedItem, "-")
            Cbx_Uitkering_Kind.Checked = InStr(str2(1), "K") > 0
            Cbx_Uitkering_Oudere.Checked = InStr(str2(1), "O") > 0
            Cbx_Uitkering_Overig.Checked = InStr(str2(1), "V") > 0

            Dim cp_amount = QuerySQL("
            Select sum(amt1) FROM journal
            WHERE name ='" & Cmx_Excasso_Select.SelectedItem & "' 
                               AND type='CP'
                               AND amt1<='0.00'")

            If IsNumeric(cp_amount) Then
                Lbl_Excasso_CP_Totaal.Text = CInt(cp_amount * -1)
                Lbl_Excasso_CP_Totaal_MDL.Text = Tbx2Int(CInt(Lbl_Excasso_CP_Totaal.Text) * Tbx2Dec(Tbx_Excasso_Exchange_rate.Text))
            Else
                Lbl_Excasso_CP_Totaal.Text = 0
                Lbl_Excasso_CP_Totaal_MDL.Text = 0
            End If

        End If
    End Sub

    Sub Load_Excasso_Balances()
        'stap 1: laad de balansen zoals bekend in de administratie
        Dim t1 = IIf(Me.Cbx_Uitkering_Kind.Checked, "Kind", "--")
        Dim t2 = IIf(Me.Cbx_Uitkering_Oudere.Checked, "Oudere", "--")
        Dim t3 = IIf(Me.Cbx_Uitkering_Overig.Checked, "Overig", "--")
        Dim m As Integer = Month(Me.Dtp_Excasso_Start.Value)
        Dim cp = Lbl_Excasso_CPid.Text
        Dim d1 As String = Dtp_Excasso_Start.Value.Year & "-" & Dtp_Excasso_Start.Value.Month & "-" &
            Dtp_Excasso_Start.Value.Day
        Dim s As String = Create_Excasso(cp, t1, t2, t3, d1, m.ToString)
        Load_Datagridview(Me.Dgv_Excasso2, s, "Load_Existing_Excasso")
        Format_dvg_excasso()

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Btn_Excasso_Base1.Click
        If Btn_Excasso_Base1.Text = "%" Then Btn_Excasso_Base1.Text = "€" Else Btn_Excasso_Base1.Text = "%"
        Calculate_CP_Allowance()
    End Sub

    Private Sub Btn_Excasso_Base2_Click(sender As Object, e As EventArgs) Handles Btn_Excasso_Base2.Click
        If Btn_Excasso_Base2.Text = "%" Then Btn_Excasso_Base2.Text = "€" Else Btn_Excasso_Base2.Text = "%"
        Calculate_CP_Allowance()
    End Sub

    Private Sub Btn_Excasso_Base3_Click(sender As Object, e As EventArgs) Handles Btn_Excasso_Base3.Click
        If Btn_Excasso_Base3.Text = "%" Then Btn_Excasso_Base3.Text = "€" Else Btn_Excasso_Base3.Text = "%"
        Calculate_CP_Allowance()
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Process.Start("https://www.xe.com/currencyconverter/convert/?Amount=1&From=EUR&To=MDL")

    End Sub

    Private Sub Tbx_Excasso_Exchange_rate_Leave(sender As Object, e As EventArgs) Handles Tbx_Excasso_Exchange_rate.Leave
        My.Settings._exrate = Tbx2Dec(Tbx_Excasso_Exchange_rate.Text)
    End Sub

    Private Sub Btn_Incasso_Delete_Click(sender As Object, e As EventArgs) Handles Btn_Incasso_Delete.Click
        RunSQL("Delete From Journal where name ='" &
               Me.Lbl_Incasso_job_name.Text & "'", "NULL", "Btn_Incasso_Delete_Click")
        Me.Lbl_Incasso_Status.Text = "Nieuw"

        Me.Btn_Incasso_Delete.Enabled = False
        Me.Btn_Run_Incasso.Enabled = True
        Me.Btn_Incasso_Print.Enabled = False

        MenuDelete.Enabled = False
        MenuSave.Enabled = True
        Menu_Print.Enabled = False

        Me.Lbl_Incasso_Error.Visible = False
    End Sub

    Private Sub Dgv_Bank_ColumnHeaderMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles Dgv_Bank.ColumnHeaderMouseClick
        'Format_dvg_bank()
    End Sub

    Private Sub Dgv_Bank_RowHeaderMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles Dgv_Bank.RowHeaderMouseClick
        'Format_dvg_bank()
    End Sub
    Private Sub Tbx_Bank_Description_Leave(sender As Object, e As EventArgs) Handles Tbx_Bank_Description.Leave, Tbx_Bank_Extra_Info.Leave
        Dim SQLstr = "UPDATE bank SET description='" & Tbx_Bank_Description.Text &
               "' WHERE id='" & Dgv_Bank.SelectedCells(0).Value & "'"
        RunSQL(SQLstr, "NULL", "Tbx_Bank_Description.Leave")
        Dgv_Bank.SelectedCells(3).Value = Tbx_Bank_Description.Text

        If Me.Dgv_Bank.RowCount > 0 Then Me.Dgv_Bank.Rows(Dgv_Bank.SelectedCells(3).RowIndex).Selected = True

    End Sub

    Private Sub Btn_Excasso_Copy_to_clipboard_Click(sender As Object, e As EventArgs) Handles Btn_Excasso_Copy_to_clipboard.Click

        If Strings.Left(Cmx_Excasso_Select.SelectedItem, 6) = "Nieuwe" Then
            MsgBox("Bewaar deze uitkeringslijst eerst s.v.p.")
        Else
            If IsDBNull(Cmx_Excasso_Select.SelectedItem) Or Cmx_Excasso_Select.SelectedItem = "" Then Exit Sub
            Clipboard.Clear()
            Clipboard.SetText(Cmx_Excasso_Select.SelectedItem)
            'MsgBox("'" & Cmx_Excasso_Select.SelectedItem & "' gekopieerd naar het klembord;
            'Plak dit s.v.p. in de omschrijving van de bankoverschrijving.")

        End If

    End Sub


    Private Sub Cmx_00_contract__fk_relation_id_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Cmx_00_contract__fk_relation_id.SelectedIndexChanged
        If Not Add_Mode Then Exit Sub
        Dim int = QuerySQL("
                                        SELECT ba.id
                                        FROM relation r
                                        LEFT join bankacc ba ON ba.accountno = r.iban 
                                        WHERE r.id ='" & Cmx_00_contract__fk_relation_id.SelectedValue & "'
                                        ")

        Me.Lbl_Contract_Bronaccount.Visible = Not IsDBNull(int)
        Me.Cmx_00_Contract__fk_account_id.Visible = Not IsDBNull(int)
        Chx_00_contract__autcol.Enabled = IsDBNull(int)

    End Sub

    Private Sub Cmx_00_contract__fk_relation_id_Enter(sender As Object, e As EventArgs) Handles Cmx_00_contract__fk_relation_id.Enter
        Exit Sub
        Dim int As Integer = QuerySQL("
                                        SELECT ba.id
                                        FROM relation r
                                        LEFT join bankacc ba ON ba.accountno = r.iban 
                                        WHERE r.id ='" & Cmx_00_contract__fk_relation_id.SelectedValue & "'
                                        ")

        Me.Lbl_Contract_Bronaccount.Visible = Not IsDBNull(int)
        Me.Cmx_00_Contract__fk_account_id.Visible = Not IsDBNull(int)
        Chx_00_contract__autcol.Enabled = IsDBNull(int)

    End Sub

    Private Sub Cbx_00_cp__active_Click(sender As Object, e As EventArgs) Handles Cbx_00_cp__active.Click
        CheckActive(Cbx_00_cp__active, Lbl_CP_pkid, "target")
    End Sub
    Function Check_Change_Bank_Categories(ByVal msg As Boolean)
        If Me.Dgv_Bank.Rows.Count = 0 Or Dgv_Bank_Account.Rows.Count = 0 Then
            Return False
            Exit Function
        End If
        If Not IsDBNull(Me.Dgv_Bank.SelectedCells(12).Value) Then
            If Me.Dgv_Bank.SelectedCells(12).Value = "Uitkering" Or Me.Dgv_Bank.SelectedCells(12).Value = "Incasso" Then
                If msg Then MsgBox("Incasso- & uitkeringslijsten kunnen niet in de bankapplicatie aangepast worden")
                Fill_Journals_by_bank(Dgv_Bank.SelectedCells(0).Value)
                Return False
            End If
        End If
        Return True

    End Function


    Sub Save_Banktransaction_Accounts()

        If Check_Change_Bank_Categories(False) = False Then Exit Sub


        Dim bid As Integer = Me.Dgv_Bank.SelectedCells(0).Value
        Dim _dat As Date = Me.Dgv_Bank.SelectedCells(1).Value
        Dim dat As String = _dat.Year & "-" & _dat.Month & "-" & _dat.Day
        Dim des As String = Me.Dgv_Bank.SelectedCells(3).Value  'dit gaat fout met een bestaande excassojob waar al een beschrijving aanwezig is
        Dim Amt_In = CDec(Me.Dgv_Bank.SelectedCells(4).Value)
        Dim Amt_Out = CDec(Me.Dgv_Bank.SelectedCells(5).Value)
        Dim cod As String = Me.Dgv_Bank.SelectedCells(6).Value
        Dim typ As String = "---"
        Dim nam As String
        Dim iban As String = Strings.Right(Me.Cmx_Bank_bankacc.Text, 18)

        'als er al eerder een categorisering heeft plaatsgevonden moet deze popup box niet opnieuw getoond worden.
        'Dim categorized As Boolean
        'categorized = (dstbank.Tables(0).Rows(0)(3) <> "Internal" And dstbank.Tables(0).Rows(0)(3) <> "Contract" And dstbank.Tables(0).Rows(0)(3) <> "Extra")


        If Rbn_Bank_jtype_con.Checked Then
            typ = "Contract"
            nam = "Contractbetaling (handmatig)"
        ElseIf Rbn_Bank_jtype_ext.Checked Then
            typ = "Extra"
            nam = "Extra gift"
        Else
            typ = "Internal"
            nam = "Betaling intern account"
        End If



        Dim SQLstr = "DELETE FROM journal WHERE fk_bank=" & bid & ";" &
                     "INSERT INTO journal(date,status,amt1,description,source, fk_account,fk_bank,name,type,iban) VALUES "

        For x As Integer = 0 To dstbank.Tables(0).Rows.Count - 1
            If Not IsDBNull(dstbank.Tables(0).Rows(x)(2)) Then
                nam = IIf(dstbank.Tables(0).Rows(x)(0) = nocat, "nog te bepalen", nam)
                If nam = "Betaling intern account" Then nam = nam & "/" & dstbank.Tables(0).Rows(x)(0)
                If dstbank.Tables(0).Rows(x)(2) <> 0 Then
                    SQLstr &= "('" & dat & "','Verwerkt','" & Cur2(dstbank.Tables(0).Rows(x)(2)) & "','" &
                        des & "','Bank'," & dstbank.Tables(0).Rows(x)(0) & "," & bid & ",'" & nam & "','" & typ & "','" & iban & "'),"
                End If
            End If
        Next

        SQLstr = Strings.Left(SQLstr, Strings.Len(SQLstr) - 1) 'remove the last comma
        If Me.Chbx_test.Checked Then MsgBox(SQLstr)
        RunSQL(SQLstr, "NULL", "")

        RunSQL("update bank b set fk_journal_name = j.source from journal j where b.id = j.fk_bank and j.fk_account !=" & nocat & " and b.fk_journal_name='nog te bepalen';
        update bank b set fk_journal_name='nog te bepalen' from journal j where b.id = j.fk_bank and j.fk_account =" & nocat, "NULL", "Categorize_Bank_Transactions / Set journal Name")


    End Sub
    Sub Calculate_Total_Booked(sender)


        Dim Amt_In = CDec(Me.Dgv_Bank.SelectedCells(4).Value)
        Dim Amt_Out = CDec(Me.Dgv_Bank.SelectedCells(5).Value)
        Dim total As Decimal = 0
        Dim nill As Integer = -1
        Dim or_amt = Amt_In - Amt_Out

        If dstbank.Tables("Table").Rows.Count = 0 Then
            'SPAS.Tbx_Bank_Accounts_Total.Text = 0

        Else
            Dim amt As Decimal
            For x As Integer = 0 To dstbank.Tables(0).Rows.Count - 1
                If dstbank.Tables(0).Rows(x)(0) = nocat Then
                    nill = x
                    'amt2 = dstbank.Tables(0).Rows(x)(2)
                Else
                    If IsDBNull(dstbank.Tables(0).Rows(x)(2)) Then amt = 0 Else amt = CDec(dstbank.Tables(0).Rows(x)(2))

                    total = total + amt
                End If
            Next
            Dim diff = or_amt - total
            If nill = -1 Then

                If diff <> 0 Then  'account 'uncategorized not present
                    Dim R As DataRow
                    R = dstbank.Tables(0).Rows.Add
                    R(0) = nocat
                    R(1) = QuerySQL("SELECT name FROM account WHERE id='" & nocat & "'")
                    R(2) = diff
                End If
            Else
                dstbank.Tables(0).Rows(nill)(2) = or_amt - total
            End If
            Me.Tbx_Bank_Amount.Text = diff

            'SPAS.Tbx_Bank_Accounts_Total.Text = total
        End If


    End Sub
    Sub Format_dvg_bank()

        With Me.Dgv_Bank
            .Columns(1).HeaderText = "Datum"
            .Columns(2).HeaderText = "Naam"
            .Columns(3).HeaderText = "Omschrijving"
            .Columns(4).HeaderText = "Bij"
            .Columns(5).HeaderText = "Af"

            .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            .Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

            .Columns(1).Width = 75
            .Columns(2).Width = 150
            .Columns(3).Width = 300
            .Columns(4).Width = 70
            .Columns(5).Width = 70

            .Columns(0).Visible = False

        End With
        Dim seq As String = ""

        For x As Integer = 0 To Dgv_Bank.Rows.Count - 1
            Dim cnt As Integer = Dgv_Bank.Rows(x).Cells(17).Value
            Dim col As Color

            Dgv_Bank.Rows(x).DefaultCellStyle.ForeColor = IIf(cnt > 0, Color.DarkRed, Color.DarkGreen)
            If Dgv_Bank.Rows(x).Cells(12).Value = "Auto-cat" Then Dgv_Bank.Rows(x).DefaultCellStyle.ForeColor = Color.DarkGoldenrod

            If x > 0 Then

                seq = Dgv_Bank.Rows(x).Cells(9).Value
                col = Color.White

                If seq = Dgv_Bank.Rows(x - 1).Cells(9).Value Then
                    Dgv_Bank.Rows(x).DefaultCellStyle.BackColor = Dgv_Bank.Rows(x - 1).DefaultCellStyle.BackColor
                Else
                    col = IIf(Dgv_Bank.Rows(x - 1).DefaultCellStyle.BackColor = Color.LightSteelBlue, Color.White, Color.LightSteelBlue)
                    Dgv_Bank.Rows(x).DefaultCellStyle.BackColor = col
                End If
            End If
        Next
        Try
            With Me.Dgv_Bank_Account

                '.Columns(0).HeaderText = "Id"
                .Columns(1).HeaderText = "Account"
                .Columns(2).HeaderText = "Bedrag"

                .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                .Columns(2).DefaultCellStyle.Format = "N2"
                .Columns(2).DefaultCellStyle.ForeColor = Color.Blue

                .Columns(0).Visible = False
                .Columns(1).Width = 200
                .Columns(2).Width = 95
                .Columns(3).Visible = False

                .Columns(0).ReadOnly = True
                .Columns(1).ReadOnly = True
                .Columns(2).ReadOnly = False

            End With
        Catch
        End Try

    End Sub

    Sub Fill_bank_transactions(sender)

        If Cmx_Bank_bankacc.SelectedIndex = -1 Then Cmx_Bank_bankacc.SelectedIndex = 0
        Calculate_Bank_Balance()
        If Strings.InStr(Me.Cmx_Bank_bankacc.Text, "NL") = 0 Then Exit Sub

        Dim bankacc = Strings.Right(Me.Cmx_Bank_bankacc.Text, 18)

        Dim sv As String = Me.Searchbox.Text '  Me.Tbx_Bank_Search.Text

        Dim SQLstr = $"SELECT id, date, name, description As descr, 
                      credit, debit,code, exch_rate, iban2, seqorder,
                      batchid, amt_cur, fk_journal_name,filename,cost,iban, id,
                      (select count(j.id) from journal j left join bank b2 on b2.id=j.fk_bank where j.fk_account='" & nocat & "' and b.id = b2.id)
                      FROM bank b WHERE iban ='" & bankacc & "' ORDER BY seqorder DESC, date DESC"


        Load_Datagridview(Me.Dgv_Bank, SQLstr, "fill bank transactions")
        Format_dvg_bank()

        'Mark_rows_Dgv_Bank()

        If Me.Dgv_Bank.RowCount > 0 Then

            Me.Dgv_Bank.Rows(0).Selected = True
        End If

    End Sub

    Private Sub NieuwToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Login.Text = "Inloggen in productieomgeving"
        Login.Cmx_Login_Database.Text = "Productie"
        Login.Show()

    End Sub

    Private Sub OpenToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Login.Text = "Inloggen in testomgeving"
        Login.Cmx_Login_Database.Text = "Acceptatie"
        Login.Show()
    End Sub

    Sub Basis_Delete()
        Dim id As Integer
        Dim sqlstr As String = ""
        Dim t As Integer = TC_Object.SelectedIndex
        If Lbx_Basis.SelectedIndex <> -1 Then id = Lbx_Basis.SelectedItem(Me.Lbx_Basis.ValueMember) Else Exit Sub

        Select Case t
            Case 0
                If Me.Dtp_31_contract__startdate.Value <= Date.Today Then
                    MsgBox("Alleen contracten die nog niet zijn ingegaan kunnen verwijderd worden.")
                    Exit Sub
                Else
                    If MsgBox("Weet u zeker dat u dit contract wilt verwijderen (vergeet niet eventueel de einddatum van eerdere versie van dit contract terug te zetten)?", vbYesNo) = vbNo Then
                        Exit Sub
                    Else

                        QuerySQL("Update account set b_jan=0, b_feb=0, b_mar=0, b_apr=0, b_may=0, b_jun=0, b_jul=0, b_aug=0, b_sep=0, b_oct=0, b_nov=0, b_dec=0 
                        where source ilike 'Doel' and f_key=" & Cmx_01_contract__fk_target_id.SelectedValue)

                        sqlstr = "DELETE FROM contract WHERE id=" & id

                    End If
                End If

            Case 1
                Collect_data("SELECT t.id, t.name, t.active, ac.name, ac.id, j.id As journal, c.id As Contract
                                From target t
                                LEFT join account ac on t.id= ac.f_key
                                LEFT join journal j on j.fk_account = ac.id
                                LEFT join contract c on c.fk_target_id = t.id
                                WHERE (j.id is null or c.id is null)
                                AND t.id =" & id)
                If dst.Tables(0).Rows.Count = 0 Then
                    MsgBox("Dit doel maakt nog onderdeel uit van een contract waarop transacties hebben plaatsgevonden." & vbCrLf &
                           "U kunt het niet verwijderen, maar wel inactief maken zodat er geen contract meer voor kan worden afgesloten of giften aan gegeven.")
                    Exit Sub
                End If
                Dim account_id = dst.Tables(0).Rows(0)(4)
                Dim journal_id = dst.Tables(0).Rows(0)(5)
                Dim contract_id = dst.Tables(0).Rows(0)(6)

                Dim Msg As String = "Dit doel: "
                If Not IsDBNull(contract_id) Then Msg &= vbCrLf & "- maakt onderdeel uit van contract " & contract_id
                If Not IsDBNull(journal_id) Then Msg &= vbCrLf & "- komt voor in journaalposten"
                If Len(Msg) > 10 Then
                    Msg &= vbCrLf & "en kan daarom niet verwijderd worden. U kunt het wel als [inactief] markeren."
                    MsgBox(Msg)
                Else
                    If MsgBox("Weet u zeker dat u het doel " & Tbx_01_Target__name.Text & "," & Tbx_01_Target__name_add.Text &
                        " wilt verwijderen?") Then
                        sqlstr = "Delete from target where id=" & id & ";DELETE from account WHERE id=" & account_id
                    End If

                End If

            Case 2
                If QuerySQL("select count(id) from contract where fk_relation_id = " & id) > 0 Then
                    MsgBox("Deze relatie staat geregistreerd bij contracten; deze moeten eerst verwijderd worden")
                Else
                    If MsgBox("Weet u zeker dat u deze relatie wilt verwijderen?", vbYesNo) = vbNo Then
                        Exit Sub
                    Else
                        sqlstr = "delete from relation where id = " & id
                    End If
                End If
            Case 3
                Collect_data("SELECT cp.name, ac.name, j.name As journal, t.name FROM CP
                                LEFT join account ac on cp.id = ac.f_key
                                LEFT JOIN journal j on ac.id = j.fk_account
                                LEFT JOIN target t on t.fk_cp_id = cp.id 
                                WHERE ac.id is not distinct from null or j.id is not distinct from null 
                                or cp.id is not distinct from null AND cp.id =" & id)
                Dim account_id = dst.Tables(0).Rows(0)(1)
                If dst.Tables(0).Rows.Count = 0 Then
                    MsgBox("Deze staat nog geregistreerd bij doel(en) en/of journaalposten." & vbCrLf &
                           "U kunt het niet verwijderen, maar wel inactief maken zodat deze niet mmer gebruikt kan worden.")
                    Exit Sub
                Else
                    sqlstr = "Delete from cp where id=" & id & ";DELETE from account WHERE id=" & account_id
                End If
            Case Else
                MsgBox("Deze functie Is nog niet voor dit object gedefinieerd")

        End Select

        If sqlstr <> "" Then
            RunSQL(sqlstr, "NULL", "Menu_Delete_Click")
            Load_Table()
            MsgBox("Het object is verwijderd.")
        End If

    End Sub

    Private Sub Btn_Excasso_Calculate_Exchrate_Click(sender As Object, e As EventArgs) Handles Btn_Excasso_Calculate_Exchrate.Click
        MsgBox("Deze functie is niet langer beschikbaar")
        Exit Sub
        Dim LocalFilePath As String = "C:\temp\lcal.html"
        Dim objWebClient As New System.Net.WebClient
        'objWebClient.DownloadFile("https://www.google.com/search?newwindow=1&sxsrf=ALeKk00tqujhzWGn2oO1UiVUC8hWGsGjvw%3A1596922685176&ei=PRsvX5KlCoHisAfjxqjQAQ&q=exchange+rate+eur+mdl&oq=exchange+ra&gs_lcp=CgZwc3ktYWIQARgAMgYIIxAnEBMyBQgAELEDMgIIADICCAAyAggAMgIIADICCAAyAggAMgIIADICCAA6BAgjECc6AgguOggIABCxAxCDAToECAAQAzoECAAQQzoHCAAQsQMQQzoICC4QsQMQgwFQ8f8YWIORGWCLnxloAHAAeACAAdoBiAGVCJIBBjExLjAuMZgBAKABAaoBB2d3cy13aXrAAQE&sclient=psy-ab", LocalFilePath)
        objWebClient.DownloadFile("https://eur.fxexchangerate.com/mdl-exchange-rates-history.html", LocalFilePath)

        Dim text As String = File.ReadAllText("C:\Temp\lcal.html")
        Dim index As Integer = text.IndexOf("<td>1 EUR =</td>")
        If index >= 0 Then
            Tbx_Excasso_Exchange_rate.Text = Math.Round(CDbl(Replace((Strings.Mid(text, index + 22, 8)), ".", ",")), 1)

        Else
            MsgBox("Wisselkoers niet gevonden.")
        End If
    End Sub

    Private Sub TestToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Login.Text = "Inloggen in productieomgeving"
        Login.Cmx_Login_Database.Text = "Test"
        Login.Show()
    End Sub

    Private Sub Btn_Excasso_Cancel_Click(sender As Object, e As EventArgs) Handles Btn_Excasso_Cancel.Click
        If Cmx_Excasso_Select.SelectedIndex = -1 Then Exit Sub
    End Sub


    Private Sub Btn_Bank_Split_Click(sender As Object, e As EventArgs) Handles Btn_Bank_Split.Click, Dgv_Bank.DoubleClick


        Banksplit.Lbl_Split_Description.Text = Dgv_Bank.SelectedCells(3).Value
        Banksplit.Lbl_Split_seqorder.Text = Dgv_Bank.SelectedCells(9).Value
        Banksplit.Lbl_Split_Bank_id.Text = Dgv_Bank.SelectedCells(0).Value
        Banksplit.Lbl_SplitBank_journal_name.Text = Dgv_Bank.SelectedCells(12).Value

        Banksplit.Lbl_Split_Amount.Text = QuerySQL("Select sum(credit) - sum(debit) from bank where seqorder = '" & Banksplit.Lbl_Split_seqorder.Text & "';")

        If Not Check_Change_Bank_Categories(False) Then Exit Sub
        Dim cnt = QuerySQL("select count(j.fk_account) from bank b left join journal j on j.fk_bank = b.id where b.id=" & Banksplit.Lbl_Split_Bank_id.Text)
        If cnt <> 1 Then
            MsgBox("Splitsen van een banktransactie met meerdere categoriëen is niet mogelijk")
            Exit Sub
        End If

        Banksplit.Lbl_SplitBank_Accountnr.Text = QuerySQL("select j.fk_account||' ['||a.name||']' from bank b left join journal j on j.fk_bank = b.id 
            left join account a on a.id = j.fk_account where b.id=" & Banksplit.Lbl_Split_Bank_id.Text)
        Dim jtype = QuerySQL("select j.type from bank b left join journal j on j.fk_bank = b.id 
            left join account a on a.id = j.fk_account where b.id=" & Banksplit.Lbl_Split_Bank_id.Text)
        If Not IsDBNull(jtype) Then Banksplit.Lbl_SplitBank_Type.Text = jtype

        Banksplit.Show()

    End Sub

    Private Sub ToolTip1_Popup(sender As Object, e As PopupEventArgs) Handles ToolTip1.Popup

        ' ToolTip1.SetToolTip(Btn_Bank_Categorize, "Categoriseer transacties")
    End Sub

    Private Sub Rbn_00_contract_child_CheckedChanged(sender As Object, e As EventArgs) Handles Rbn_00_contract_child.CheckedChanged
        Tbx_Contract_ttype.Text = "Kind"
    End Sub

    Private Sub Rbn_00_contract_elder_CheckedChanged(sender As Object, e As EventArgs) Handles Rbn_00_contract_elder.CheckedChanged
        Tbx_Contract_ttype.Text = "Oudere"
    End Sub

    Private Sub Rbn_00_contract_other_CheckedChanged(sender As Object, e As EventArgs) Handles Rbn_00_contract_other.CheckedChanged
        Tbx_Contract_ttype.Text = "Overig"
    End Sub

    Private Sub Btn_Settings_Cancel_Click(sender As Object, e As EventArgs)
        Load_Account_Settings()
    End Sub

    Private Sub Dgv_Excasso2_DoubleClick(sender As Object, e As EventArgs) Handles Dgv_Excasso2.DoubleClick


        If Dgv_Excasso2.CurrentCell.ColumnIndex <> 1 Then Exit Sub
        'Clipboard.SetText(Dgv_Excasso2.CurrentCell.Value)
        TC_Main.SelectedIndex = 4
        Cmx_Journal_List.SelectedIndex = 0
        Searchbox.Text = Dgv_Excasso2.CurrentCell.Value
        If Lv_Journal_List.Items.Count > 0 Then
            Lv_Journal_List.Items(0).Focused = True
            Lv_Journal_List.Items(0).Selected = True


        End If

    End Sub

    Private Sub Dgv_Excasso2_Click(sender As Object, e As EventArgs) Handles Dgv_Excasso2.Click
        'Clipboard.Clear()
        'Clipboard.SetText(Dgv_Excasso2.CurrentCell.Value)
    End Sub

    Private Sub Lbl_Excasso_Items_Contract_TextChanged(sender As Object, e As EventArgs) Handles Lbl_Excasso_Items_Contract.TextChanged,
            Lbl_Excasso_Items_Extra.TextChanged, Lbl_Excasso_Items_Intern.TextChanged, Lbl_Excasso_Contractwaarde.TextChanged,
            Lbl_Excasso_Extra.TextChanged, Lbl_Excasso_Intern.TextChanged
        Lbl_Excasso_Contract.Text = Lbl_Excasso_Items_Contract.Text & " gepland, €" _
        & Lbl_Excasso_Contractwaarde.Text & " à"
        Lbl_Excasso_Extr.Text = Lbl_Excasso_Items_Extra.Text & " extra, €" _
        & Lbl_Excasso_Extra.Text & " à"
        Lbl_Excasso_Internal.Text = Lbl_Excasso_Items_Intern.Text & " intern, €" _
        & Lbl_Excasso_Intern.Text & " à"
    End Sub

    Private Sub Btn_Incasso_Export_Click(sender As Object, e As EventArgs) Handles Btn_Incasso_Export.Click
        Export_2_Excel(Me.Dgv_Incasso)
    End Sub

    Private Sub ToolStripTextBox1_TextChanged(sender As Object, e As EventArgs) Handles Searchbox.TextChanged
        'Dim dt As New DataTable()
        Select Case TC_Main.SelectedIndex
            Case 0
                Load_Table()
            Case 1
                'Fill_bank_transactions("Searchbox.TextChanged")
                If Dgv_Bank.DataSource IsNot Nothing Then
                    ApplyFilter(Dgv_Bank.DataSource)
                    Format_dvg_bank()
                End If

            Case 4
                Fill_Cmx_Journal_List()
            Case 5

                ' dt = Dgv_Rapportage_Overzicht.DataSource
                If Dgv_Rapportage_Overzicht.DataSource IsNot Nothing Then
                    ApplyFilter(Dgv_Rapportage_Overzicht.DataSource)
                    Format_Datagridview(Dgv_Rapportage_Overzicht, LbL_Formatting.Text.Split(","c), False)
                End If


        End Select


    End Sub

    Sub ApplyFilter(ByVal dt As DataTable)
        If String.IsNullOrWhiteSpace(Searchbox.Text) Then
            dt.DefaultView.RowFilter = "" ' Clear filter if search box is empty
            Return
        End If

        ' Split search terms by spaces
        Dim searchTerms As String() = Searchbox.Text.Split(New Char() {" "c}, StringSplitOptions.RemoveEmptyEntries)

        Dim filterParts As New List(Of String)

        For Each term As String In searchTerms
            Dim termFilter As String = ""
            For Each col As DataColumn In dt.Columns
                If Not String.IsNullOrEmpty(col.ColumnName) Then
                    If termFilter.Length > 0 Then termFilter &= " OR "
                    termFilter &= $"CONVERT([{col.ColumnName}], 'System.String') LIKE '%{term}%'"
                End If
            Next
            ' Wrap each term's filter in parentheses and add to the list
            filterParts.Add($"({termFilter})")
        Next

        ' Combine all term filters with AND
        dt.DefaultView.RowFilter = String.Join(" AND ", filterParts)
    End Sub

    Private Sub ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles MenuFilter.Click
        Searchbox.Text = ""
    End Sub

    Private Sub Lv_Journal_List_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Lv_Journal_List.SelectedIndexChanged
        If Cmx_Journal_List.Text = "Journaalnaam" Then
            TC_Boeking.SelectedIndex = 2
            Fill_Journal_List_journaalposten()
        Else
            Fill_Journal_List()
        End If

    End Sub

    Private Sub Cbx_LifeCycle_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Cbx_LifeCycle.SelectedIndexChanged

        Select Case TC_Main.SelectedIndex
            Case 0
                Try
                    MenuDelete.Enabled = (Cbx_LifeCycle.Text = "Inactief") Or Dtp_31_contract__startdate.Value > Date.Today
                    Load_Table()
                Catch ex As Exception
                End Try
            Case 1
                'Fill_bank_transactions()
            Case 4
                Fill_Cmx_Journal_List()
        End Select
    End Sub


    Private Sub MenuSave_Click(sender As Object, e As EventArgs) Handles MenuSave.Click
        Select Case TC_Main.SelectedIndex
            Case 0 'basisadministratie
                Basis_Save()
            Case 1 'Bank
                'Save_Banktransaction_Accounts()
                'Mark_rows_Dgv_Bank()
            Case 2 'Incasso
                Create_Incasso_Journals()
                Create_SEPA_XML()
                Me.Lbl_Incasso_Status.Text = "Open"
                Me.Btn_Incasso_Delete.Enabled = True
                Me.Btn_Run_Incasso.Enabled = False
                Me.Btn_Incasso_Print.Enabled = True
                MenuDelete.Enabled = True
                Menu_Print.Enabled = True
                MenuSave.Enabled = False
            Case 3 'Uitkeringen
                If Cmx_Excasso_Select.SelectedIndex = -1 Then Exit Sub
                Save_Excasso_job()
            Case 4
                Load_Combobox(Cmx_Bank_Account, "id", "name", "SELECT id, name FROM account WHERE active = True ORDER BY source, name")
                Select Case TC_Boeking.SelectedIndex
                    Case 0
                        MsgBox(0)
                    Case 1
                        Save_Internal_Booking()
                    Case 2
                        Save_modified_journaalposts()
                    Case 6
                        Save_Settings()
                End Select
                Exit Sub
                If TC_Boeking.SelectedIndex = 1 Then

                    Save_Internal_Booking()
                ElseIf TC_Boeking.SelectedIndex = 0 Then
                    Dim id = Dgv_journaalposten.SelectedCells(10).Value
                    RunSQL("Update Journal Set description='" & Tbx_Journal_Descr.Text & "' Where id='" & id & "'", "NULL", "Tbx_Journal_Descr_Leave")
                    Load_Datagridview(Me.Dgv_Journal_items, Create_Journal_SQL, "Lv_Journal_List_Click")
                    Me.Tbx_.Text = Me.Dgv_Journal_items.Rows(0).Cells(1).Value
                    Me.Tbx_Journal_Descr.Text = Me.Dgv_Journal_items.Rows(0).Cells(4).Value
                    MenuSave.Enabled = False

                End If

        End Select

    End Sub

    Private Sub MenuAdd_Click(sender As Object, e As EventArgs) Handles MenuAdd.Click
        Select Case TC_Main.SelectedIndex
            Case 0
                PreviousTabMain = TC_Main.SelectedIndex
                PreviousTab = TC_Object.SelectedIndex
                Basis_Add()
            Case 3
        End Select
    End Sub
    Private Sub Menu_Export_Click(sender As Object, e As EventArgs) Handles MenuCancel.Click
        Select Case TC_Main.SelectedIndex
            Case 2

            Case 3

            Case 6
                Load_Account_Settings()

            Case Else
        End Select
    End Sub
    Private Sub MenuCancel_Click(sender As Object, e As EventArgs) Handles MenuCancel.Click
        Select Case TC_Main.SelectedIndex
            Case 0
                Cancel()
            Case 3
                If Cmx_Excasso_Select.SelectedIndex = -1 Then Exit Sub
            Case 4
                If TC_Boeking.SelectedIndex = 1 Then
                    Lbl_Journal_Source_Saldo.Text = 0
                    Lbl_Journal_Source_Name.Text = ""
                    Tbx_Journal_Source_Amt.Text = 0
                    Dgv_Journal_Intern.Rows.Clear()
                    Lbl_Journal_Source_Restamt.Text = 0
                End If
        End Select
    End Sub

    Private Sub MenuDelete_Click(sender As Object, e As EventArgs) Handles MenuDelete.Click


        Select Case TC_Main.SelectedIndex
            Case 0
                Basis_Delete()
            Case 2
                RunSQL("Delete From Journal where name ='" &
                 Me.Lbl_Incasso_job_name.Text & "'", "NULL", "Btn_Incasso_Delete_Click")
                Me.Lbl_Incasso_Status.Text = "Nieuw"
                Me.Btn_Incasso_Delete.Enabled = False
                Me.Btn_Run_Incasso.Enabled = True
                Me.Btn_Incasso_Print.Enabled = False

                MenuDelete.Enabled = False
                MenuSave.Enabled = True
                Menu_Print.Enabled = False

                Me.Lbl_Incasso_Error.Visible = False
            Case 3
                MenuExcassoDelete()
        End Select
    End Sub

    Private Sub MenuBanktransactie_Click(sender As Object, e As EventArgs) Handles MenuBanktransactie.Click
        Download_Bank_Transactions()
    End Sub

    Private Sub MenuUploadAlles_Click(sender As Object, e As EventArgs) Handles MenuUploadAlles.Click
        Load_Bank_csv_from_folder()
    End Sub

    Private Sub MenuCategoriseer_Click(sender As Object, e As EventArgs) Handles MenuCategoriseer.Click
        Categorize_Bank_Transactions(True, True, True, True, True, True, True)
        Fill_bank_transactions("MenuCategoriseer")
    End Sub

    Private Sub Menu_Print_Click(sender As Object, e As EventArgs) Handles Menu_Print.Click
        Select Case TC_Main.SelectedIndex
            Case 2
                Create_SEPA_XML()
            Case 3
                If Cmx_Excasso_Select.SelectedIndex = -1 Then Exit Sub
                Save_Excasso_job()
                Print_Excasso_form()
        End Select
    End Sub
    Sub ShowButtons()

        Dim i = TC_Main.SelectedIndex
        MenuBanktransactie.Visible = (i = 1)
        MenuUploadAlles.Visible = (i = 1)
        MenuBanktransactie.Visible = (i = 1)
        MenuCategoriseer.Visible = (i = 1)
        Menu_Export.Visible = (i = 2 Or i = 4 Or i = 5) Or TC_Boeking.SelectedIndex = 0
        Menu_Print.Visible = (i = 2 Or i = 3)
        MenuDelete.Visible = (i = 0 Or i = 2 Or i = 3)
        MenuSave.Visible = (i = 0 Or i = 2 Or i = 3 Or i = 4 Or i = 6)
        MenuAdd.Visible = (i = 0) Or (i = 4 And TC_Boeking.SelectedIndex = 2)
        MenuCancel.Visible = (i = 0 Or i = 3 Or (i = 4 And TC_Boeking.SelectedIndex > 0)) Or i = 6
        ZoekenToolStripMenuItem.Visible = (i < 2 Or i = 4 Or i = 5)
        Searchbox.Visible = ZoekenToolStripMenuItem.Visible
        MenuFilter.Visible = ZoekenToolStripMenuItem.Visible
        ToolStripTextBox1.Visible = (i <> 1 Or i = 4)
        Cbx_LifeCycle.Visible = ToolStripTextBox1.Visible
        MenuAdd.Visible = IIf(InStr(Text, "(ALLEEN LEZEN)") = 0, MenuAdd.Visible, False)
        MenuSave.Visible = IIf(InStr(Text, "(ALLEEN LEZEN)") = 0, MenuSave.Visible, False)
        MenuDelete.Visible = IIf(InStr(Text, "(ALLEEN LEZEN)") = 0, MenuDelete.Visible, False)
        Menu_Help.Visible = (i = 4)


    End Sub
    Private Sub TC_Main_Click(sender As Object, e As EventArgs) Handles TC_Main.Click


        Select Case TC_Main.SelectedIndex

            Case 0

                Manage_Buttons_Target(True, True, True, False, False, "TC_Main_SelectedIndexChanged")
                If Searchbox.Text <> "" Then Load_Table()
            Case 1  'bank
                Searchbox.Text = ""
                Manage_Buttons_Target(False, True, False, False, False, "TC_Main_SelectedIndexChanged")

                'only load the bank data if datagridview is still empty
                If Dgv_Bank.Rows.Count = 0 Or Dgv_Bank.DataSource Is Nothing Then
                    If Me.Dgv_Mgnt_Tables.Rows(1).Cells(1).Value > 0 Then
                        Fill_bank_transactions("Cmx_Bank_bankacc.SelectedIndexChanged")
                    End If
                End If

            Case 2 'incasso

                If Me.Dgv_Mgnt_Tables.Rows(3).Cells(1).Value > 0 And
                    Me.Dgv_Mgnt_Tables.Rows(5).Cells(1).Value > 0 And
                    Me.Dgv_Mgnt_Tables.Rows(8).Cells(1).Value > 0 Then
                    Create_Incassolist()
                    Dtp_Incasso_start.Format = DateTimePickerFormat.Custom
                    Dtp_Incasso_start.CustomFormat = "MMMM yyyy"
                    Dtp_Incasso_start.ShowUpDown = True


                    Format_dvg_incasso()
                End If

            Case 3

                MenuSave.Enabled = True
                MenuCancel.Enabled = True
                MenuDelete.Enabled = True
                MenuAdd.Enabled = False
                MenuBanktransactie.Visible = False ' SPAS.SPAS.V
                MenuUploadAlles.Visible = False
                MenuBanktransactie.Visible = False
                MenuCategoriseer.Visible = False

                If Me.Dgv_Mgnt_Tables.Rows(3).Cells(1).Value > 0 And
                    Me.Dgv_Mgnt_Tables.Rows(5).Cells(1).Value > 0 And
                    Me.Dgv_Mgnt_Tables.Rows(8).Cells(1).Value > 0 And
                    Me.Cmx_Excasso_Select.SelectedItem = "" Then

                    Dtp_Excasso_Start.ShowUpDown = False
                    Dtp_Excasso_Start.Value = CDate(Date.Today.Year & "-" & Date.Today.Month & "-" & Date.Today.Day)
                    Dtp_Excasso_Start.MaxDate = CDate(Date.Today.Year & "-" & Date.Today.Month & "-" & Date.Today.Day)
                    Fill_Cmx_Excasso_Select_Combined()
                End If

            Case 4
                Manage_Buttons_Target(False, True, False, False, False, "TC_Main_SelectedIndexChanged")
                If Cmx_Journal_List.Text = "" Or IsDBNull(Cmx_Journal_List.Text) Then Cmx_Journal_List.Text = "Accounts"
                Me.Dtp_Journal_intern.Value = CDate(Date.Today.Year & "-" & Date.Today.Month & "-" & Date.Today.Day)
                Dim sql As String = "update journal j set name = 
                (select left(replace(replace(replace(replace(replace(replace(replace(b.name,' van der',''),' van de',''),'Hr ',''),'Mw ',''),' de ',''),' van ',''),'.',''),14) 
                from bank b where b.id = j.fk_bank)||'/'||(select a.name from account a where a.id = j.fk_account)
                where name='nog te bepalen' and fk_account != (select value::integer from settings where label='nocat') and source = 'Bank'"
                RunSQL(sql, "NULL", "TC_Main_Click")

            Case 5
                Manage_Buttons_Target(False, True, False, False, False, "TC_Main_SelectedIndexChanged")

            Case 6
                MenuSave.Enabled = True
                MenuCancel.Enabled = True
                Load_Account_Settings()



            Case Else


        End Select
        ShowButtons()


    End Sub
    Private Sub TC_Boeking_Click(sender As Object, e As EventArgs) Handles TC_Boeking.Click
        ShowButtons()
        MenuSave.Enabled = True
        MenuCancel.Enabled = TC_Boeking.SelectedIndex = 1
        Menu_Export.Enabled = TC_Boeking.SelectedIndex = 0

        If TC_Boeking.SelectedIndex = 0 And Cmx_Journal_List.Text = "Journaalnaam" Then 'boekingen
            Cmx_Journal_List.Text = "Accounts"
        ElseIf TC_Boeking.SelectedIndex = 2 And Cmx_Journal_List.Text <> "Journaalnaam" Then
            Cmx_Journal_List.Text = "Journaalnaam"
            'Me.Cmbx_journaalposten_account.SelectedIndex = -1
            'Me.Cmbx_journaalposten_relatie.SelectedIndex = -1
        Else 'overboekingen
            Searchbox.Text = ""
            If Lbl_Journal_Source_Name.Text = "" Then
                Tbx_Journal_Name.Text = ""
                Tbx_Journal_Description.Text = ""
                Rbn_Journal_Intern.Checked = True
            End If
            If Cmx_Journal_List.Text = "CP" Or Cmx_Journal_List.Text = "Journaalnaam" Or Cmx_Journal_List.Text = "Relaties" Then
                Cmx_Journal_List.Text = "Accounts"
            End If

        End If


    End Sub

    Private Sub Dgv_Bank_Account_Leave(sender As Object, e As EventArgs) Handles Dgv_Bank_Account.Leave

        If Check_Change_Bank_Categories(False) = False Then Exit Sub

        Calculate_Total_Booked("Dgv_Bank_Account_Leave")
        Save_Banktransaction_Accounts()
        Update_Category_Status()

    End Sub

    Private Sub Menu_Export_Click_1(sender As Object, e As EventArgs) Handles Menu_Export.Click
        Select Case TC_Main.SelectedIndex
            Case 1
                Export_2_Excel(Me.Dgv_Bank)
            Case 2
                Export_2_Excel(Me.Dgv_Incasso)
            Case 3
                Export_2_Excel(Dgv_Excasso2)
            Case 4
                Export_2_Excel(Dgv_Journal_items)

            Case 5
                Select Case TC_Rapportage.SelectedTab.Name
                    Case "Journaal"
                        Export_2_Excel(Dgv_Rapportage_Overzicht)
                    Case "TC_Boekingen"
                        Export_2_Excel(Dgv_Report_6)
                    Case "TC_Jaarafsluiting"
                        Export_2_Excel(Dgv_Report_Year_Closing)
                End Select
            Case Else
        End Select
    End Sub

    Private Sub Lbl_Incasso_Error_Click(sender As Object, e As EventArgs) Handles Lbl_Incasso_Error.Click

    End Sub
    Private Sub Rbn_Bank_jtype_con_Click(sender As Object, e As EventArgs) Handles Rbn_Bank_jtype_con.Click, Rbn_Bank_jtype_ext.Click, Rbn_Bank_jtype_int.Click
        'Aanklikken mag, alleen bewaren als er een andere categorie als "nocat" is
        If Dgv_Bank.Rows(Dgv_Bank.SelectedCells(2).RowIndex).DefaultCellStyle.ForeColor <> Color.DarkRed Then Save_Banktransaction_Accounts()

    End Sub

    Private Sub Cbx_Journal_Status_Click(sender As Object, e As EventArgs) Handles Cbx_Journal_Status_Open.Click, Cbx_Journal_Status_Verwerkt.Click,
            Cbx_Journal_Saldo_Open.Click, Cmx_Journal_List.SelectedIndexChanged, Chbx_Journal_Inactive.CheckedChanged
        Fill_Cmx_Journal_List()
    End Sub

    Sub Call_Excasso_form(sender As Object)


        If Cmx_Excasso_Select.SelectedItem = "" Then Exit Sub 'alleen lijst genereren als gekozen is voor een bestaande of nieuwe lijst
        Dim t1 = IIf(Me.Cbx_Uitkering_Kind.Checked, "Kind", "--")
        Dim t2 = IIf(Me.Cbx_Uitkering_Oudere.Checked, "Oudere", "--")
        Dim t3 = IIf(Me.Cbx_Uitkering_Overig.Checked, "Overig", "--")
        If t1 & t2 & t3 = "------" Then Exit Sub

        Dim dat As String = Me.Dtp_Excasso_Start.Value.Year & "-" & Me.Dtp_Excasso_Start.Value.Month & "-" & Me.Dtp_Excasso_Start.Value.Day
        'Dim _dat2 As Date = Me.Dtp_Excasso_Start.Value.AddDays(-1)
        'Dim dat2 As String = _dat2.Year & "-" & _dat2.Month & "-" & _dat2.Day

        Dim cp As String = ""
        Dim nieuw As Boolean = False
        Dim prefill_contract As Boolean = Rbn_uitkering_budget.Checked
        Dim naam1 As String = Cmx_Excasso_Select.SelectedItem
        Dim naam2 As String

        If Strings.Left(Cmx_Excasso_Select.SelectedItem, 5) = "Nieuw" Then
            Dim pos1 As Integer = Strings.InStr(Me.Cmx_Excasso_Select.SelectedItem, "[")
            cp = Strings.Mid(Me.Cmx_Excasso_Select.SelectedItem, pos1 + 1, Len(Me.Cmx_Excasso_Select.SelectedItem) - pos1 - 1)
            nieuw = True
            naam2 = ""
            Pan_Excasso_preset.Enabled = True
            '@@@ cp calculatie klopt nog niet 
        Else
            nieuw = False
            naam2 = naam1
            cp = Lbl_Excasso_CPid.Text
            Pan_Excasso_preset.Enabled = False
        End If

        'Dim s As String = Get_Excasso_data(cp, t1, t2, t3, Cmx_Excasso_Select.SelectedItem, dat, nieuw, prefill_contract)
        Dim s2 As String = Get_Excasso_data2(cp, t1, t2, t3, naam1, naam2, dat)
        'If s = "" Then Exit Sub
        If s2 = "" Then Exit Sub

        Clipboard.Clear()
        Clipboard.SetText(s2)


        'Load_Datagridview(Me.Dgv_Excasso2, s, "Call_Excasso_form")
        Load_Datagridview(Me.Dgv_Excasso2, s2, "Call_Excasso_form2")

        If nieuw Then Prefill_Excasso_Form()

        Format_dvg_excasso2()

        Calculate_Excasso_Totals2()



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
end as e_cont,
case 
 when (select -round(sum(amt1)::numeric) from journal j where j.fk_account = ac.id and j.type = 'Extra' and j.name ilike '" & naam1 & "'and j.date <='" & dat & "') is not distinct from null then 0::numeric
 else (select -round(sum(amt1)::numeric) from journal j where j.fk_account = ac.id and j.type = 'Extra' and j.name ilike '" & naam1 & "'and j.date <='" & dat & "')
end as e_extra,
case 
 when (select -round(sum(amt1)::numeric) from journal j where j.fk_account = ac.id and j.type = 'Internal' and j.name ilike '" & naam1 & "'and j.date <='" & dat & "') is not distinct from null then 0::numeric
 else (select -round(sum(amt1)::numeric) from journal j where j.fk_account = ac.id and j.type = 'Internal' and j.name ilike '" & naam1 & "'and j.date <='" & dat & "')
end as e_intern,
    0::numeric as e_tot,
    0::numeric as m_tot

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
        Clipboard.Clear()
        Clipboard.SetText(Sqlstr)
        Return Sqlstr



    End Function
    Function Get_Excasso_data(ByVal cp As String, type1 As String, type2 As String, type3 As String, u_form As String, dat As String, nieuw As Boolean, prefill_contract As Boolean)

        Dim Sqlstr As String =
            "
        -- UITKERINGEN

        select id, name, case when (round(max(plan::numeric)) is distinct from null and not " & nieuw & ") or 
            (" & nieuw & " and (select count(*) from journal where extract(month from date) = extract(month from current_date) 
                and source = 'Uitkering' and fk_account = it1.id and type = 'Contract')=0)

            
            then round(max(plan::numeric)) else 0 end as plan
	        ,round(sum(contr::numeric)) as saldo, round(sum(extra::numeric))as extra, round(sum(intern::numeric)) as intern 
	        ,case when not " & nieuw & " --ophalen incassoformulier
			        then 
			        case when (select round(sum(j2.amt1::numeric)) from journal j2 where j2.name like '" & u_form & "' and j2.fk_account = it1.id and j2.type ilike '%contract%') is distinct from null 
			     	        then (select -round(sum(j2.amt1::numeric)) from journal j2 where j2.name like '" & u_form & "' and j2.fk_account = it1.id and j2.type ilike '%contract%') else 0 end
			        else
			        case --lijst vullen met contractbudget of saldo contractbetalingen
				        when " & prefill_contract & " 
					        then case when round(max(plan::numeric)) is distinct from null  
                        
                        and (select count(*) from journal where extract(month from date) = extract(month from current_date) 
                and source = 'Uitkering' and fk_account = it1.id and type = 'Contract')=0
                        
                        
                        then round(max(plan::numeric)) else 0 end else 
				        case when round(sum(contr::numeric)) is distinct from null then 
                        case when round(sum(contr::numeric))>0 then round(sum(contr::numeric)) else 0 end  else 0 end
			        end	end as e_cont
	        ,case when not " & nieuw & " --ophalen incassoformulier
			    then (case when (select round(sum(j2.amt1::numeric)) from journal j2 where j2.name like '%" & u_form & "%' and j2.fk_account = it1.id and (j2.type ilike '%extra%')) is distinct from null 
	 					    then (select -round(sum(j2.amt1::numeric)) from journal j2 where j2.name like '" & u_form & "' and j2.fk_account = it1.id and (j2.type ilike '%extra%')) else 0::numeric end) 
	 				        else round(sum(extra::numeric))
	            end as e_extra 
	 	    ,case when not " & nieuw & " --ophalen incassoformulier
			    then (case when (select round(sum(j2.amt1::numeric)) from journal j2 where j2.name like '%" & u_form & "%' and j2.fk_account = it1.id and (j2.type ilike '%intern%')) is distinct from null 
	 					    then (select -round(sum(j2.amt1::numeric)) from journal j2 where j2.name like '%" & u_form & "%' and j2.fk_account = it1.id and (j2.type ilike '%intern%')) else 0::numeric end)
	 					    else round(sum(intern::numeric))
	            end as e_intern
            ,0::numeric as e_tot
	        ,0::numeric as m_tot

	        FROM
		        (
                SELECT 
                    ac.id as id, ac.name as name
                    ,CASE 
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

                     END As Plan
			        ,CASE when j.type ilike '%contract%' and j.name not like '" & u_form & "' then sum(j.amt1) else 0::money end as Contr
			        ,CASE when j.type ilike '%extra%' and j.name not like '" & u_form & "' then sum(j.amt1) else 0::money end as Extra
			        ,CASE when j.type ilike '%intern%' and j.name not like '" & u_form & "' then sum(j.amt1) else 0::money end as Intern
			        ,case when (select sum(j2.amt1) from journal j2 where j2.name not like '" & u_form & "' and j2.fk_account = ac.id and j2.type ilike '%contract%') is distinct from null 
			        then (select sum(j2.amt1) from journal j2 where j2.name not like '" & u_form & "' and j2.fk_account = ac.id and j2.type ilike '%contract%') else 0::money end
			        as ctr1
                FROM account ac
                    LEFT JOIN journal j ON j.fk_account = ac.id
                    LEFT JOIN target ta ON ta.id = ac.f_key
                    LEFT JOIN cp ON cp.id = ta.fk_cp_id
                    WHERE (ta.ttype='" & type1 & "' or  ta.ttype='" & type2 & "' or ta.ttype='" & type3 & "')
                    AND ta.active=true 
                    and cp.id=" & cp & "
			        GROUP BY ac.id, ac.name, j.name, j.type ORDER BY ac.name asc
                    ) as it1
        group by id, name
        order by name
        "
        Clipboard.SetText(Sqlstr)

        Return Sqlstr
        '
    End Function


    Private Sub Rbn_uitkering_saldo_Click(sender As Object, e As EventArgs) Handles Rbn_uitkering_saldo.Click
        Prefill_Excasso_Form()
        Calculate_Excasso_Totals2()

        'Call_Excasso_form("saldo")
    End Sub

    Private Sub Rbn_uitkering_budget_Click(sender As Object, e As EventArgs) Handles Rbn_uitkering_budget.Click
        Prefill_Excasso_Form()
        Calculate_Excasso_Totals2()
        'Call_Excasso_form("budget")
    End Sub

    Private Sub Rbn_uitkering_nul_Click(sender As Object, e As EventArgs) Handles Rbn_uitkering_nul.Click
        Set_Excasso_Nullvalues2()
    End Sub
    Sub Format_dvg_excasso2()

        If Dgv_Excasso2.Rows.Count = 0 Then Exit Sub  'de vraag is of dit correct is
        Try
            With Dgv_Excasso2

                .Columns(0).HeaderText = "Id"
                .Columns(1).HeaderText = "Account"
                .Columns(2).HeaderText = "Plan"
                .Columns(3).HeaderText = "Saldo"
                .Columns(4).HeaderText = "Extra"
                .Columns(5).HeaderText = "Intern"
                .Columns(6).HeaderText = "Plan"
                .Columns(7).HeaderText = "Extra"
                .Columns(8).HeaderText = "Intern"
                .Columns(9).HeaderText = "Tot EUR"
                .Columns(10).HeaderText = "Tot MLD"

                For c = 2 To 10
                    .Columns(c).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter  '.MiddleRight
                    '.Columns(c).DefaultCellStyle.Format = "N2"
                    .Columns(c).Width = 58
                Next

                .Columns(0).Width = 45
                .Columns(1).Width = 140
                For c = 0 To 10 : .Columns(c).ReadOnly = True : Next
                For c = 6 To 8
                    .Columns(c).DefaultCellStyle.ForeColor = Color.Blue
                    .Columns(c).ReadOnly = False
                Next
                For c = 9 To 10
                    .Columns(c).DefaultCellStyle.ForeColor = Color.Green
                Next
                '.Columns(6).DefaultCellStyle.ForeColor = Color.Blue
                '.Columns(6).ReadOnly = False
                '.Columns(6).DefaultCellStyle.Format = "G"
                '.Columns(6).ValueType = GetType(Decimal)
                '.Columns(7).DefaultCellStyle.ForeColor = Color.Green
                '.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                '.Columns(7).Width = 80


            End With
        Catch
        End Try
        Exit Sub
        For x = 0 To Dgv_Excasso2.Rows.Count - 1
            For y = 2 To Dgv_Excasso2.Columns.Count - 3
                If IsDBNull(Dgv_Excasso2.Rows(x).Cells(y).Value) Then
                    Dgv_Excasso2.Rows(x).Cells(y).Value = 0
                    Dgv_Excasso2.Rows(x).Cells(y).Style.ForeColor = Color.LightGray
                Else
                    If Dgv_Excasso2.Rows(x).Cells(y).Value = 0 Then
                        Dgv_Excasso2.Rows(x).Cells(y).Style.ForeColor = Color.LightGray
                    End If
                End If
            Next y
        Next


    End Sub
    Sub Set_Excasso_Nullvalues2()
        If Dgv_Excasso2.Rows.Count > 0 Then
            'Convert_Null_to_0()
            For x As Integer = 0 To Dgv_Excasso2.Rows.Count - 1
                For y = 6 To 10
                    Dgv_Excasso2.Rows(x).Cells(y).Value = 0
                Next
            Next
            Lbl_Excasso_Items_Totaal.Text = 0
            Lbl_Excasso_Items_Contract.Text = 0
            Lbl_Excasso_Items_Extra.Text = 0
            Lbl_Excasso_Items_Intern.Text = 0
            Lbl_Excasso_Contractwaarde.Text = 0
            Lbl_Excasso_Extra.Text = 0
            Lbl_Excasso_Intern.Text = 0
            Lbl_Excasso_Totaal.Text = 0
            Lbl_Excasso_Tot_Gen_MLD.Text = 0
            Lbl_Excasso_Tot_Gen.Text = 0
            Lbl_Excasso_Totaal_MDL.Text = 0
            Calculate_CP_Allowance()
        End If
    End Sub
    Sub Prefill_Excasso_Form()

        Dim r = IIf(Rbn_uitkering_budget.Checked, 2, 3)

        For x As Integer = 0 To Dgv_Excasso2.Rows.Count - 1
            Dgv_Excasso2.Rows(x).Cells(6).Value = IIf(Dgv_Excasso2.Rows(x).Cells(r).Value > 0, Dgv_Excasso2.Rows(x).Cells(r).Value, 0)
            Dgv_Excasso2.Rows(x).Cells(7).Value = Dgv_Excasso2.Rows(x).Cells(4).Value
            Dgv_Excasso2.Rows(x).Cells(8).Value = Dgv_Excasso2.Rows(x).Cells(5).Value

        Next x



    End Sub

    Sub Calculate_Excasso_Totals2()


        If Dgv_Excasso2.Rows.Count > 0 Then
            Dim items_contract = 0
            Dim items_extra = 0
            Dim items_intern = 0
            Dim amount_contract = 0
            Dim amount_extra = 0
            Dim amount_intern = 0
            Dim total_eur = 0
            Dim tot_gen As Integer = 0
            Dim cnt As Integer = 0  'teller voor het totaal aantal begunstigden
            Dim begunstigde As Boolean = False
            Dim amount As Integer
            For x As Integer = 0 To Dgv_Excasso2.Rows.Count - 1
                total_eur = 0
                For y = 6 To 8
                    amount = IIf(IsDBNull(Dgv_Excasso2.Rows(x).Cells(y).Value), 0, Dgv_Excasso2.Rows(x).Cells(y).Value)
                    amount = Dgv_Excasso2.Rows(x).Cells(y).Value
                    If amount > 0 Then
                        Dgv_Excasso2.Rows(x).Cells(9).Value = Dgv_Excasso2.Rows(x).Cells(9).Value + Dgv_Excasso2.Rows(x).Cells(y).Value
                        total_eur = total_eur + amount
                        begunstigde = True
                        Select Case y
                            Case 6
                                items_contract = items_contract + 1
                                amount_contract = amount_contract + amount
                            Case 7
                                items_extra = items_extra + 1
                                amount_extra = amount_extra + amount
                            Case 8
                                items_intern = items_intern + 1
                                amount_intern = amount_intern + amount

                        End Select
                    End If

                Next
                cnt = cnt + IIf(begunstigde, 1, 0)
                begunstigde = False
                Dgv_Excasso2.Rows(x).Cells(9).Value = total_eur
                Dgv_Excasso2.Rows(x).Cells(10).Value = Math.Round(Dgv_Excasso2.Rows(x).Cells(9).Value * Tbx_Excasso_Exchange_rate.Text, 0)
                'Dgv_Excasso2.Rows(x).Cells(9).Value = Dgv_Excasso2.Rows(x).Cells(6).Value + Dgv_Excasso2.Rows(x).Cells(7).Value + Dgv_Excasso2.Rows(x).Cells(8).Value
            Next
            Lbl_Excasso_Items_Contract.Text = items_contract
            Lbl_Excasso_Contractwaarde.Text = amount_contract
            Lbl_Excasso_Items_Extra.Text = items_extra
            Lbl_Excasso_Extra.Text = amount_extra
            Lbl_Excasso_Items_Intern.Text = items_intern
            Lbl_Excasso_Intern.Text = amount_intern
            Lbl_Excasso_Items_Totaal.Text = cnt
            Lbl_Excasso_Totaal.Text = amount_contract + amount_extra + amount_intern

            tot_gen = CInt(Lbl_Excasso_Totaal.Text) + CInt(Lbl_Excasso_CP_Totaal.Text)
            'MsgBox(tot_gen)
            Lbl_Excasso_Tot_Gen.Text = tot_gen
            Lbl_Excasso_Totaal_MDL.Text = Math.Round(Lbl_Excasso_Totaal.Text * Tbx_Excasso_Exchange_rate.Text, 0)
            Lbl_Excasso_Tot_Gen_MLD.Text = Math.Round(Lbl_Excasso_Tot_Gen.Text * Tbx_Excasso_Exchange_rate.Text, 0)
            Btn_Excasso_CP_Calculate.Enabled = True


        End If
        If Tbx_Excasso_CP1.Text = "" Then Tbx_Excasso_CP1.Text = 0
        Lbl_Excasso_CP_Totaal.Text = CInt(Tbx_Excasso_CP1.Text) + CInt(Tbx_Excasso_CP2.Text) + CInt(Tbx_Excasso_CP3.Text)
        Lbl_Excasso_Tot_Gen.Text = CInt(Lbl_Excasso_Totaal.Text) + CInt(Lbl_Excasso_CP_Totaal.Text)
        Lbl_Excasso_CP_Totaal_MDL.Text = Math.Round(CInt(Lbl_Excasso_CP_Totaal.Text) * Tbx_Excasso_Exchange_rate.Text, 0)
        Lbl_Excasso_Tot_Gen_MLD.Text = CInt(Lbl_Excasso_CP_Totaal_MDL.Text) + CInt(Lbl_Excasso_Totaal_MDL.Text)
    End Sub


    Sub Empty_Excasso_Window()

        Dgv_Excasso2.DataSource = Nothing
        Dgv_Excasso2.Rows.Clear()
        Me.Dgv_Excasso2.Columns.Clear()
        Lbl_Excasso_Items_Contract.Text = 0
        Lbl_Excasso_Items_Extra.Text = 0
        Lbl_Excasso_Items_Intern.Text = 0
        Lbl_Excasso_Extra.Text = 0
        Lbl_Excasso_Contractwaarde.Text = 0
        Lbl_Excasso_Intern.Text = 0
        Lbl_Excasso_Totaal.Text = 0
        Lbl_Excasso_Items_Totaal.Text = 0
        Lbl_Excasso_Totaal_MDL.Text = 0
        Lbl_Excasso_Tot_Gen_MLD.Text = 0
        Lbl_Excasso_Tot_Gen.Text = 0
        Cbx_Uitkering_Kind.Checked = False
        Cbx_Uitkering_Oudere.Checked = False
        Cbx_Uitkering_Overig.Checked = False
        Tbx_Excasso_CP1.Text = 0
        Tbx_Excasso_CP2.Text = 0
        Tbx_Excasso_CP3.Text = 0
    End Sub

    Private Sub Button1_Click_2(sender As Object, e As EventArgs)
        MsgBox("Dit is een handmatige activiteit die door de databasebeheerder moet worden uitgevoerd")
    End Sub

    Private Sub Tbx_01_relation__name_Enter(sender As Object, e As EventArgs) Handles Tbx_01_relation__name.Enter,
            Tbx_00_Relation__description.Enter, Tbx_01_Relation__name_add.Enter, Tbx_00_Relation__email.Enter,
            Tbx_00_Relation__phone.Enter, Tbx_00_Relation__address.Enter, Tbx_00_Relation__zip.Enter, Tbx_00_Relation__city.Enter

        Edit_Mode = True
        Manage_Buttons_Target(False, False, False, True, True, "Rbn_Relation_1_CheckedChanged")
    End Sub

    Private Sub TC_Rapportage_Click(sender As Object, e As EventArgs) Handles TC_Rapportage.Click
        'SelectNodeByName(BankTree, "Financieel")

        If TC_Rapportage.SelectedTab.Tag = "report_closing" Then Report_Closing()

    End Sub

    Private Sub Dgv_Rapportage_Overzicht_CellContentDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles Dgv_Rapportage_Overzicht.CellContentDoubleClick



    End Sub
    Function Check_administratie()


        Collect_data(QuerySQL("Select sql from query where name= 'Haal checks op'"))
        Dim result As VariantType
        Dim response As String = "Uitkomsten controles:" & vbCr
        Dim Res As Boolean = True


        For c = 0 To dst.Tables(0).Rows.Count - 1
            'dst.Tables(0).Rows(c)(1) = dst.Tables(0).Rows(c)(1).Replace("p1", Bank_table(rep_year))
            'dst.Tables(0).Rows(c)(1) = dst.Tables(0).Rows(c)(1).Replace("p2", Report_table(rep_year))
            dst.Tables(0).Rows(c)(1) = dst.Tables(0).Rows(c)(1).Replace("[year]", report_year)
            Debug.Print(dst.Tables(0).Rows(c)(1))
            'result = QuerySQL(dst.Tables(0).Rows(c)(1))

            If Not IsDBNull(QuerySQL(dst.Tables(0).Rows(c)(1))) Then

                result = QuerySQL(dst.Tables(0).Rows(c)(1))
                If result <> 0 Then
                    dst.Tables(0).Rows(c)(2) = dst.Tables(0).Rows(c)(2).Replace("#", result)
                    dst.Tables(0).Rows(c)(2) = "- FOUT: " & dst.Tables(0).Rows(c)(2)
                    response &= dst.Tables(0).Rows(c)(2) & vbCr
                    Res = False

                End If
            End If
            If Strings.Left(dst.Tables(0).Rows(c)(2), 6) <> "- FOUT" Then
                dst.Tables(0).Rows(c)(3) = "- " & dst.Tables(0).Rows(c)(3) & " OK"
                response &= dst.Tables(0).Rows(c)(3) & vbCr
            End If

        Next c

        MsgBox(response, IIf(Res, vbInformation, vbCritical))
        If Res Then Return True Else Return False

    End Function



    Private Sub Btn_Report_YearEnd_Post_Click(sender As Object, e As EventArgs) Handles Btn_Report_YearEnd_Post.Click

        '----------- uitvoeren controles

        If Check_administratie() = False Then Exit Sub
        If MsgBox("Wilt u het jaar " & report_year & " definitief afsluiten? Dit kan niet meer worden teruggedraaid!", vbYesNo) = vbNo Then Exit Sub

        'Ophalen transitieposten ten behoeve van table t1
        Dim Sql1 As String = QuerySQL("Select sql from query where category = 'Overzicht' and name='Transitieposten'")
        Dim Sql2 As String = QuerySQL("Select sql from query where category = 'Transaction' and name='Jaarafsluiting'")
        Dim Sql As String = Sql1 & vbCr & Sql2
        Sql = Sql.Replace("[Year]", report_year)
        'Uitvoeren jaarafsluiting
        RunSQL(Sql, "NULL", "Btn_Report_YearEnd_Post/Transitieposten")

        If MsgBox("Wilt u de budgetten voor " & Now.Year & " berekenen (eventuele handmatige aanpassen gaan verloren)? ", vbYesNo) = vbYes Then
            Calculate_Budget("")
        End If
    End Sub

    Private Sub Tbx_Bank_Description_TextChanged(sender As Object, e As EventArgs) Handles Tbx_Bank_Description.TextChanged
        Dgv_Bank.SelectedCells(3).Value = Tbx_Bank_Description.Text
    End Sub
    Private Sub Tbx_01_Accgroup__type_TextChanged(sender As Object, e As EventArgs) Handles Tbx_01_Accgroup__type.TextChanged
        Rbtn_accgroup_Income.Checked = Strings.Trim(Tbx_01_Accgroup__type.Text) = "Inkomsten"
        Rbtn_accgroup_expense.Checked = Strings.Trim(Tbx_01_Accgroup__type.Text) = "Uitgaven"
        Rbtn_accgroup_transit.Checked = Strings.Trim(Tbx_01_Accgroup__type.Text) = "Transit"
        '@@@ hard value vervangen door tt_type.Text
    End Sub

    Private Sub Rbtn_accgroup_Income_CheckedChanged(sender As Object, e As EventArgs) Handles Rbtn_accgroup_Income.Click
        If Not Add_Mode Then Edit_Mode = True 'Manage_Buttons_Target(False, False, False, True, True, "Radiobutton")
        If MenuSave.Enabled Then Tbx_01_Accgroup__type.Text = Rbtn_accgroup_Income.Text
    End Sub

    Private Sub Rbtn_accgroup_expense_CheckedChanged(sender As Object, e As EventArgs) Handles Rbtn_accgroup_expense.Click
        If Not Add_Mode Then Edit_Mode = True 'Manage_Buttons_Target(False, False, False, True, True, "Radiobutton")
        If MenuSave.Enabled Then Tbx_01_Accgroup__type.Text = Rbtn_accgroup_expense.Text
    End Sub

    Private Sub Rbtn_accgroup_transit_CheckedChanged(sender As Object, e As EventArgs) Handles Rbtn_accgroup_transit.Click
        If Not Add_Mode Then Edit_Mode = True 'Manage_Buttons_Target(False, False, False, True, True, "Radiobutton")
        If MenuSave.Enabled Then Tbx_01_Accgroup__type.Text = Rbtn_accgroup_transit.Text
    End Sub


    Private Sub Dgv_Journal_items_Click(sender As Object, e As EventArgs) Handles Dgv_Journal_items.Click
        If Dgv_Journal_items.RowCount = 0 Then Exit Sub

        Dim a = Dgv_Journal_items.SelectedCells(0).Value
        Dim i As Integer = Me.Dgv_Journal_items.CurrentRow.Index
        Dim name As String = Me.Dgv_Journal_items.Rows(i).Cells(1).Value
        Dim id = Dgv_Journal_items.SelectedCells(10).Value
        Tbx_Journal_Descr.Text = Dgv_Journal_items.SelectedCells(4).Value 'omschrijving

    End Sub



    Private Sub Tbx_Journal_Descr_Enter(sender As Object, e As EventArgs) Handles Tbx_Journal_Descr.Enter
        MenuSave.Enabled = True
    End Sub


    Private Sub Dgv_Report_6_Click(sender As Object, e As DataGridViewCellEventArgs)
        Dim jid As String = Me.Dgv_Report_6.Rows(Me.Dgv_Report_6.CurrentRow.Index).Cells(9).Value

        Dim s As String = "
            select 
                j.id, j.source, j.status, j.type, j.iban,
                b.id, b.date, b.debit, b.credit, b.name, b.description, b.iban2,
                r.id, r.name||', '||r.name_add,
                a.id, a.name
                From journal j 
                Left Join bank b on b.id = j.fk_bank
                Left Join relation r on r.id = j.fk_relation
                Left Join account a on a.id = j.fk_account
            where j.id = " & jid
        ToClipboard(s, True)

        Collect_data(s)

        Me.Tbx_Report6_Add.Text = "[JOURNAAL]   id: <" & dst.Tables(0).Rows(0)(0) & ">  bron: <" & Trim(dst.Tables(0).Rows(0)(1)) & ">  status: <" & Trim(dst.Tables(0).Rows(0)(2)) _
            & ">  type: <" & dst.Tables(0).Rows(0)(3) & ">  iban: <" & dst.Tables(0).Rows(0)(4) & ">  account id:<" & dst.Tables(0).Rows(0)(14) & ">" & vbCrLf _
            & "[BANK]  id:<" & dst.Tables(0).Rows(0)(5) & ">  datum: <" & dst.Tables(0).Rows(0)(6) & ">  bedrag: <" & dst.Tables(0).Rows(0)(8) - dst.Tables(0).Rows(0)(7) _
            & ">  naam: <" & dst.Tables(0).Rows(0)(9) & ">  tegenrekening: <" & dst.Tables(0).Rows(0)(11) & ">  omschrijving: <" & dst.Tables(0).Rows(0)(10)



    End Sub


    Sub Format_Datagridview(dgv As DataGridView, arr As Array, Editable As Boolean)

        'formatarray
        '[letter][getal][getal][getal]
        'T = Standaardformaat
        'N = Numeriek / 2 cijfers achter de komma
        'H = Verberg kolom
        'Getallen: kolombreedte
        'Rijen: Tota betekent totaalkolom

        Dim c As Integer
        Dim f As String
        Dim tstr1, tstr2 As String

        Try
            With dgv
                ' formatteer kolommen
                For x = 0 To UBound(arr)
                    'Debug.Print(arr(x))
                    c = CInt(Mid(arr(x), 2))
                    f = Strings.Left(arr(x), 1)

                    .Columns(x).ReadOnly = Not Editable
                    .Columns(x).Width = c
                    .Columns(x).HeaderText = Strings.Left(.Columns(x).HeaderText, 1).ToUpper & Strings.Mid(.Columns(x).HeaderText, 2).ToLower

                    If f = "N" Then
                        .Columns(x).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        .Columns(x).DefaultCellStyle.Format = "N2"
                    ElseIf f = "H" Then
                        .Columns(x).Visible = False
                    End If
                Next

                'formatter rijen
                'MsgBox(dgv.Name.ToString)
                For r As Integer = 0 To .Rows.Count - 1

                    For x = 0 To UBound(arr)
                        'Hide_Zero_values(dgv.Rows(r).Cells(x).Value)
                        If IsDBNull(.Rows(r).Cells(x).Value) Then .Rows(r).Cells(x).Value = 0
                        If .Rows(r).Cells(x).Value IsNot Nothing Then
                            If .Rows(r).Cells(x).Value.ToString = "0,00" Or .Rows(r).Cells(x).Value.ToString = "0" Then
                                .Rows(r).Cells(x).Style.ForeColor = Color.LightGray
                            End If
                        End If

                        tstr1 = CStr(.Rows(r).Cells(x).Value)
                        tstr2 = Strings.Mid(CStr(.Rows(r).Cells(x).Value), 6)


                        If InStr(tstr1, "Tota") > 0 Then
                            .Rows(r).DefaultCellStyle.BackColor = Color.Khaki

                        ElseIf InStr(tstr1, "Afschrift") > 0 Then
                            .Rows(r).DefaultCellStyle.BackColor = Color.DarkSeaGreen
                        ElseIf InStr(tstr1, "(Excasso)") > 0 Then
                            .Rows(r).DefaultCellStyle.BackColor = Color.White
                        ElseIf InStr(tstr1, "(Tussenrekening)") > 0 Then
                            .Rows(r).DefaultCellStyle.BackColor = Color.Gainsboro
                        ElseIf InStr(tstr1, "#") > 0 Then
                            .Rows(r).DefaultCellStyle.BackColor = Color.DarkSeaGreen
                        End If

                        'extra formatting 
                        If InStr(tstr1, "generaal") > 0 Then
                            .Rows(r).DefaultCellStyle.Font = New Font("Calibri", 12, FontStyle.Bold)
                            .Rows(r).DefaultCellStyle.ForeColor = Color.Blue
                        End If
                        If InStr(tstr1, "Vergelijking") > 0 Then
                            .Rows(r).DefaultCellStyle.Font = New Font("Calibri", 10, FontStyle.Italic)
                            .Rows(r).DefaultCellStyle.ForeColor = Color.Blue
                        End If

                    Next x
                Next r
            End With
        Catch ex As Exception
            MsgBox(ex.ToString & "-x-")
        End Try

    End Sub

    Sub Hide_Zero_values(ByVal value)
        If IsDBNull(value) Then value = 0
        If value IsNot Nothing Then
            If value.ToString = "0,00" Or value.ToString = "0" Then
                value.Style.ForeColor = Color.LightGray
            End If
        End If
    End Sub

    Private Sub Rbn_Bank_jtype_con_CheckedChanged(sender As Object, e As EventArgs) Handles Rbn_Bank_jtype_con.CheckedChanged, Rbn_Bank_jtype_ext.CheckedChanged, Rbn_Bank_jtype_int.CheckedChanged
        Btn_Bank_Add_Journal.Enabled = True
    End Sub


    Private Sub Btn_Report_YearEnd_Check_Click(sender As Object, e As EventArgs) Handles Btn_Report_YearEnd_Check.Click
        Dim ans = Check_administratie()

    End Sub

    Private Sub Btn_Query_Test_Click(sender As Object, e As EventArgs)

        If UCase(Strings.Left(Tbx_Query_SQL.Text, 6)) <> "SELECT" Then
            MsgBox("Alleen select-statements zijn toegestaan")
        Else
            Load_Datagridview(Dgv_Query_Test, Tbx_Query_SQL.Text, "Btn_Query_Test.Click")
            'MsgBox("Query is niet correct")
        End If


    End Sub

    Private Sub Button3_Click_1(sender As Object, e As EventArgs) Handles Button3.Click
        Dim p1 = InputBox("maand:")
        Dim sql = QuerySQL("Select sql from query where category ilike 'Transaction' and name='Verwijder maand'")
        sql = sql.Replace("p1", p1)
        ToClipboard(sql, True)
        RunSQL(sql, "NULL", "Testbutton verwijder maand")
        Fill_bank_transactions("Button3")

    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim p1 = InputBox("maand:")
        Dim sql = QuerySQL("Select sql from query where category ilike 'Transaction' and name='Verwijder maand'")
        sql = sql.Replace("p1", p1)
        ToClipboard(sql, True)
        RunSQL(sql, "NULL", "Testbutton verwijder maand")
        Fill_bank_transactions("Button1")
    End Sub


    Private Sub Tbx_Extra_Info_TextChanged(sender As Object, e As EventArgs) Handles Tbx_Bank_Extra_Info.TextChanged
        Dim des As String = Tbx_Bank_Description.Text
        If Chbx_Bank_ExtraInfo_voor.Checked Then
            If Strings.InStr(des, " | ") = 0 And Tbx_Bank_Extra_Info.Text <> "" Then des = " | " & des
            Try
                des = Tbx_Bank_Extra_Info.Text & Strings.Mid(des, Strings.InStr(des, " | "))
            Catch
            End Try
        Else
            'Tbx_Bank_Description.Text = Tbx_Bank_Description.Text & " [Extra info] " & Tbx_Bank_Extra_Info.Text
            'Tbx_Bank_Extra_Info.Text = Strings.Mid(Tbx_Bank_Description.Text, Strings.InStr(Tbx_Bank_Description.Text, "[Extra info]") + 14)
        End If
        If Tbx_Bank_Extra_Info.Text = "" And Strings.InStr(des, " | ") > 0 Then des = Mid(des, Strings.InStr(des, " | ") + 3)
        Tbx_Bank_Description.Text = des


    End Sub




    Private Sub Menu_Help_Click(sender As Object, e As EventArgs) Handles Menu_Help.Click

        Select Case TC_Main.SelectedIndex
            Case 4
                Helptext.TextBox1.Text = QuerySQL("Select value from settings where label = 'journalposten' limit 1")
                Helptext.LbL_Onderwerp2.Text = "Journalposten"
            Case Else
                Helptext.TextBox1.Text = "Nog geen helptekst beschikbaar"
        End Select
        Helptext.Show()


    End Sub
    '========================================================================================================
    '======                                                                                            ======
    '======                                B O E K I N G E N                                           ======
    '======                                                                                            ======
    '========================================================================================================

    Sub Lv_Journal_List_Click(sender As Object, e As EventArgs) Handles Lv_Journal_List.Click
        If Cmx_Journal_List.Text <> "Journaalnaam" Then
            'TC_Boeking.SelectedIndex = 2
            Fill_Journal_List()
        Else
            'TC_Boeking.SelectedIndex = 0
            Try
                Dim selectedItem As ListViewItem = Lv_Journal_List.SelectedItems(0)
                Collect_data(Create_Journal_SQL)

                Me.Lbl_Journaalposten_datum.Text = IIf(IsDBNull(dst.Tables(0).Rows(0)(0)), "", dst.Tables(0).Rows(0)(0))
                Me.Lbl_Journaalposten_header.Text = IIf(IsDBNull(dst.Tables(0).Rows(0)(1)), "", dst.Tables(0).Rows(0)(1))
                'Me.Tbx_journaalposten_omschr.Text = IIf(IsDBNull(dst.Tables(0).Rows(0)(4)), "", dst.Tables(0).Rows(0)(4))
                Me.Lbl_journaalposten_status.Text = IIf(IsDBNull(dst.Tables(0).Rows(0)(6)), "", dst.Tables(0).Rows(0)(6))
                Me.Lbl_Journaalposten_bron.Text = IIf(IsDBNull(dst.Tables(0).Rows(0)(7)), "", dst.Tables(0).Rows(0)(7))
                Me.Lbl_journaalposten_iban.Text = IIf(IsDBNull(dst.Tables(0).Rows(0)(8)), "", dst.Tables(0).Rows(0)(8))
                Me.Lbl_journaalposten_type.Text = IIf(IsDBNull(dst.Tables(0).Rows(0)(9)), "", dst.Tables(0).Rows(0)(9))
                Me.Lbl_journaalposten_cpinfo.Text = IIf(IsDBNull(dst.Tables(0).Rows(0)(14)), "", dst.Tables(0).Rows(0)(14))
                Me.Lbl_journaalposten_wisselkoers.Text = IIf(IsDBNull(dst.Tables(0).Rows(0)(15)), "", dst.Tables(0).Rows(0)(15))
                Me.Banklink.Text = IIf(IsDBNull(dst.Tables(0).Rows(0)(16)), 0, dst.Tables(0).Rows(0)(16).ToString)               'Me.Cmbx_journaalposten_relatie.SelectedIndex = -1

                Fill_Journal_List_journaalposten()


                If Dgv_journaalposten.Rows.Count > 0 Then
                    ' Clear any previous selection
                    Dgv_journaalposten.ClearSelection()
                    Dgv_journaalposten.Rows(0).Selected = True
                    Dgv_journaalposten_Click("a", e)
                    ' Optionally, scroll to the first row if it is out of view
                    Dgv_journaalposten.FirstDisplayedScrollingRowIndex = 0
                End If
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try


        End If
    End Sub


    '========================================================================================================
    '======                                                                                            ======
    '======                B O E K I N G E N   - J O U R N A A L P O S T E N                           ======
    '======                                                                                            ======
    '========================================================================================================

    Private Sub Dgv_Journal_items_DoubleClick(sender As Object, e As EventArgs) Handles Dgv_Journal_items.DoubleClick
        Dim i As Integer = Me.Dgv_Journal_items.CurrentRow.Index
        Dim name As String = Me.Dgv_Journal_items.Rows(i).Cells(1).Value
        Dim dat_ = Me.Dgv_Journal_items.Rows(i).Cells(0).Value
        Dim dat As String = dat_.Year & "-" & dat_.Month & "-" & dat_.Day
        Dim selectedItem As ListViewItem = Lv_Journal_List.SelectedItems(0)

        Me.Tbx_.Text = name
        journalid2.Text = Me.Dgv_Journal_items.Rows(i).Cells(10).Value.ToString
        Me.Lbl_accountname.Text = selectedItem.SubItems(1).Text
        Select_in_Lv_Journal_list_ByNameAndDate(name, dat, 2, "Journaalnaam")
        'Fill_Journal_List_journaalposten()
        Lv_Journal_List_Click(sender, e)

    End Sub


    Private Sub Menu_Back_Click(sender As Object, e As EventArgs) Handles Menu_Back.Click
        ' Loop through each item in the ListView


        TC_Boeking.SelectedIndex = 0
        Cmx_Journal_List.Text = "Accounts"
        For Each item As ListViewItem In Lv_Journal_List.Items
            If item.SubItems(1).Text = Lbl_accountname.Text Then
                item.Selected = True
                item.EnsureVisible()
                Exit For
            End If
        Next
        For Each row As DataGridViewRow In Dgv_Journal_items.Rows
            ' Compare the name (first column) and date (third column)
            If row.Cells(10).Value.ToString() = Me.journalid2.Text Then
                ' Select the row
                row.Selected = True
                ' Optionally, scroll the DataGridView to the selected row
                Dgv_Journal_items.FirstDisplayedScrollingRowIndex = row.Index
                ' Stop the loop once the matching row is found
                Exit For
            End If
        Next

    End Sub



    Sub Dgv_journaalposten_Click(sender As Object, e As EventArgs) Handles Dgv_journaalposten.Click

        Try
            Dim selectedRow As DataGridViewRow = Dgv_journaalposten.CurrentRow
            If selectedRow Is Nothing Then selectedRow = Dgv_journaalposten.Rows(0)


            Me.Cmbx_journaalposten_account.SelectedValue = selectedRow.Cells("Accountnr").Value
            Me.Tbx_journaalposten_omschr.Text =
                If(IsDBNull(selectedRow.Cells("Omschrijving").Value), "", selectedRow.Cells("Omschrijving").Value)
            If Not IsDBNull(Me.Dgv_journaalposten.Rows(0).Cells(13).Value) Then
                Me.Cmbx_journaalposten_relatie.SelectedValue = selectedRow.Cells("relatie").Value
            Else
                Me.Cmbx_journaalposten_relatie.SelectedIndex = -1
            End If
        Catch ex As Exception

            MsgBox(ex.ToString)
        End Try

    End Sub

    Private Sub Dgv_journaalposten_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles Dgv_journaalposten.CellValueChanged
        Try

            If Not Dgv_journaalposten.Rows(e.RowIndex).IsNewRow Then
                Dgv_journaalposten.Rows(e.RowIndex).Tag = "Modified"

                Calculate_Journaalposten_totalen(Dgv_journaalposten)
            End If
        Catch ex As Exception
        End Try
    End Sub
    Private Sub Dgv_journaalposten_UserAddedRow(sender As Object, e2 As DataGridViewRowEventArgs) Handles Dgv_journaalposten.UserAddedRow

        Dgv_journaalposten.Rows(Dgv_journaalposten.RowCount - 2).Cells(2).Value = 0
        Dgv_journaalposten.Rows(Dgv_journaalposten.RowCount - 2).Cells(3).Value = 0
        Dgv_journaalposten.Rows(Dgv_journaalposten.RowCount - 2).Cells(4).Value = "handmatig toegevoegde journaalpost"
        'Calculate_Journaalposten_totalen(Dgv_journaalposten)

    End Sub

    Private Sub Cmbx_journaalposten_account_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Cmbx_journaalposten_account.SelectedIndexChanged
        If Cmbx_journaalposten_account.SelectedIndex = -1 Or
            TC_Boeking.SelectedIndex <> 2 _
            Or Dgv_journaalposten.RowCount = 0 Then Exit Sub


        Try
            Dgv_journaalposten.SelectedCells(11).Value = Cmbx_journaalposten_account.SelectedValue
            Dgv_journaalposten.SelectedCells(5).Value = Cmbx_journaalposten_account.Text

        Catch ex As Exception
            'MsgBox(ex.ToString)
        End Try


    End Sub

    Private Sub Cmbx_journaalposten_relatie_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Cmbx_journaalposten_relatie.SelectedIndexChanged
        If Cmbx_journaalposten_relatie.SelectedIndex = -1 Then Exit Sub
        Try
            Dgv_journaalposten.SelectedCells(13).Value = Cmbx_journaalposten_relatie.SelectedValue
            Dgv_journaalposten.SelectedCells(17).Value = Cmbx_journaalposten_relatie.Text
        Catch ex As Exception
            'MsgBox(ex.ToString)
        End Try
    End Sub
    Private Sub Tbx_journaalposten_omschr_TextChanged(sender As Object, e As EventArgs) Handles Tbx_journaalposten_omschr.TextChanged
        Try
            Dgv_journaalposten.SelectedCells(4).Value = Tbx_journaalposten_omschr.Text
        Catch ex As Exception

        End Try
    End Sub
    Private Sub Dgv_journaalposten_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles Dgv_journaalposten.CellClick
        Dim selectedRowIndex As Integer = Dgv_journaalposten.CurrentCell.RowIndex
        With Dgv_journaalposten
            If .CurrentCell.ColumnIndex = 5 Or .CurrentCell.ColumnIndex = 17 Then
                .Rows(selectedRowIndex).Selected = True
            End If
        End With

        'MsgBox(Dgv_journaalposten.CurrentCell.ColumnIndex.ToString & " Row:" & Dgv_journaalposten.CurrentCell.RowIndex.ToString)
    End Sub


    Sub Save_modified_journaalposts()
        Dim name As String = Lbl_Journaalposten_header.Text
        Dim datum As Date = Lbl_Journaalposten_datum.Text
        Dim status As String = Trim(Lbl_journaalposten_status.Text)
        Dim amt1 As Decimal
        Dim bij As Decimal
        Dim af As Decimal
        Dim amt2 As Decimal
        Dim description As String
        Dim source As String = Trim(Lbl_Journaalposten_bron.Text)
        Dim id As Integer
        Dim fk_account As Integer
        Dim fk_relation As Integer
        Dim fk_bank As Integer = Integer.Parse(Banklink.Text)
        Dim Type As String = Trim(Lbl_journaalposten_type.Text)
        Dim cpinfo As String = Trim(Lbl_journaalposten_cpinfo.Text)
        Dim iban As String = Trim(Lbl_journaalposten_iban.Text)
        Dim transactiesaldo = CInt(Math.Round(Decimal.Parse(Tbx_Journal_Saldo.Text)))

        Dim errmsg As String = ""


        ' ---------------------1 Uitvoeren van controles ----------
        'a) Intern: saldo moet altijd 0 zijn

        If source = "Intern" And Tbx_Journal_Saldo.Text <> "0,00" Then
            errmsg &= "- Interne transacties moeten altijd een nulsaldo hebben." & vbCr
        End If
        'b) Bank: bankbedrag moet altijd gelijk zijn aan journaaltransactie
        If source <> "Intern" And status <> "Open" Then
            'Dim bankcheck QuerySQL("select sum(credit-debit) from bank where id=" & fk_bank) & "---" & transactiesaldo)
            If QuerySQL("select sum(credit-debit) from bank where id=" & fk_bank) - transactiesaldo <> 0 Then
                errmsg &= "- Mismatch tussen bankbedrag and journaaltransactiesaldo" & vbCr
            End If
        End If
        'c) Open: geen bewerkingen toestaan
        If status = "Open" Then
            errmsg &= "- Deze transactie is nog niet verwerkt, doe eventuele aanpassingen in het tabblad '" _
                & source & "'" & vbCr
        End If
        '-- de volgende controls betreffende de aanpasbare inhoud 
        'For Each row As DataGridViewRow In .
        For x As Integer = 0 To Dgv_journaalposten.Rows.Count - 2
            With Dgv_journaalposten
                bij = .Rows(x).Cells(2).Value
                af = .Rows(x).Cells(3).Value
                fk_account = IIf(IsDBNull(.Rows(x).Cells(11).Value), "0", .Rows(x).Cells(11).Value)
                'd) Als bij gevuld is moet af 0 zijn en viceversa
                If bij <> 0 And af <> 0 Then
                    errmsg &= "- BIJ en AF kunnen niet beide een bedrag zijn" & vbCr
                End If
                'e) Fk_account moet altijd ingevuld zijn (standaard "Niet toegewezen"?
                If fk_account = 0 Then
                    errmsg &= "- Account ontbreekt in journaalpost, dit is verplicht" & vbCr
                End If
            End With
        Next
        If errmsg = "" Then
            If MsgBox("Weet u zeker dat u deze handmatige aanpassingen in het grootboek wil aanbrengen?", vbYesNo) = vbNo Then Exit Sub
        Else
            MsgBox("De wijzigingen kunnen niet worden opgeslagen vanwege:" & vbCr & errmsg)
            Exit Sub
        End If

        For Each row As DataGridViewRow In Dgv_journaalposten.Rows
            If row.Tag IsNot Nothing Then
                id = IIf(IsDBNull(row.Cells("id").Value), Nothing, row.Cells("id").Value.ToString)

                amt2 = 0  '@@@ nog aanpassen
                If row.Cells("bij").Value > 0 Then
                    amt1 = row.Cells("bij").Value.ToString
                Else
                    amt1 = "-" & row.Cells("af").Value.ToString
                End If
                description = row.Cells("omschrijving").Value.ToString
                fk_account = row.Cells("accountnr").Value.ToString
                fk_relation = IIf(IsDBNull(row.Cells("relatie").Value), Nothing, row.Cells("relatie").Value.ToString)
                Dim ops As String
                If row.Tag.ToString = "Modified" Then
                    If Len(row.Cells("id").Value.ToString) = 0 Then 'new record
                        ops = "INSERT"
                    Else 'updated record
                        ops = "UPDATE"

                    End If
                    'MsgBox(row.Tag.ToString & "---" & description)
                    Run_SQL_Journal("Save_modified_journaalposts", ops, id, name, datum, status, amt1, amt2, description,
                                             source, fk_account, fk_relation, fk_bank, Type, cpinfo, iban)

                End If
            End If

        Next


    End Sub

    Sub Calculate_Journaalposten_totalen(dgv As DataGridView)
        Dim cred, deb As Decimal

        Try
            For r = 0 To dgv.RowCount - 1
                cred += dgv.Rows(r).Cells(2).Value
                deb += dgv.Rows(r).Cells(3).Value
                'cred += IIf(IsDBNull(dgv.Rows(r).Cells(2).Value) = 0, 0, dgv.Rows(r).Cells(2).Value)
                'deb += IIf(IsDBNull(dgv.Rows(r).Cells(3).Value) = 0, 0, dgv.Rows(r).Cells(3).Value)
            Next
            Tbx_Journal_Credit.Text = cred.ToString("#0.00")
            Tbx_Journal_Debit.Text = deb.ToString("#0.00")
            Tbx_Journal_Saldo.Text = cred - deb
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Grp_Journaalposten_Enter(sender As Object, e As EventArgs) Handles Grp_Journaalposten.Enter

    End Sub

    Private Sub Banklink_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles Banklink.LinkClicked
        Dim bankid As Integer = Integer.Parse(Banklink.Text)
        If bankid = 0 Or Len(bankid) = 0 Then Exit Sub
        TC_Main.SelectedIndex = 1
        Fill_bank_transactions("TC_Main.SelectedIndex")

        SelectRowById(Dgv_Bank, bankid)


    End Sub

    Private Sub Overboekingen_Click(sender As Object, e As EventArgs) Handles Overboekingen.Click
        With Dgv_Journal_Intern
            .Columns(0).Visible = False
            .Columns(1).Width = 160
            .Columns(1).ReadOnly = True
            .Columns(2).Width = 70
            .Columns(2).DefaultCellStyle.Format = "N2"
            .Columns(2).DefaultCellStyle.ForeColor = Color.Blue
            .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

        End With
    End Sub

    Private Sub Btn_Run_Incasso_Click(sender As Object, e As EventArgs) Handles Btn_Run_Incasso.Click
        Create_Incassolist()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Count_Occurences()
    End Sub

    Sub Run_BankTree(ByVal rep As String)
        Dim arr_format() As String = Nothing
        Dim sql As String = ""
        Dim formatting As String

        'ElseIf e.Node.Level = 2 Then
        '   rep = e.Node.Parent.Text


        sql = QuerySQL($"Select sql from query where category ilike 'Overzicht%' and name='{rep}'")
        If IsNothing(Sql) Then Exit Sub

        Formatting = QuerySQL($"Select formatting from query where category ilike 'Overzicht%' and name='{rep}'")
        LbL_Formatting.Text = Formatting
        If Not IsNothing(LbL_Formatting.Text) Then arr_format = LbL_Formatting.Text.Split(","c)

        Sql = Sql.Replace("[year]", report_year)
        If Cmbx_Reporting_Year.SelectedIndex > 0 Then
            Sql = Sql.Replace("from bank ", "from bank_archive ")
            Sql = Sql.Replace("from journal ", "from journal_archive ")
        End If
        Load_Datagridview(Dgv_Rapportage_Overzicht, Sql, "BankTree.NodeMouseClick-level2")
        Format_Datagridview(Dgv_Rapportage_Overzicht, arr_format, False)
    End Sub

    Sub BankTree_NodeMouseClick(sender As Object, e As TreeNodeMouseClickEventArgs) Handles BankTree.NodeMouseClick

        Dim rep As String = ""

        report_year = Cmbx_Reporting_Year.SelectedItem
        'MsgBox(report_year)


        If e.Node.Level = 1 Then
            rep = e.Node.Text
            Lbl_Rapportage.Text = rep
            Run_BankTree(rep)
        End If





        Exit Sub
        Dim arr_format() As String = Nothing
        Dim sql As String = ""
        Dim formatting As String
        If e.Node.Level = 1 Then
            'report_year = QuerySQL("select min(extract (year from date)) from journal")
            rep = e.Node.Text
            Lbl_Rapportage.Text = rep
            sql = QuerySQL($"Select sql from query where category ilike 'Overzicht%' and name='{rep}'")
            If IsNothing(sql) Then Exit Sub

            formatting = QuerySQL($"Select formatting from query where category ilike 'Overzicht%' and name='{e.Node.Text}'")
            LbL_Formatting.Text = formatting
            If Not IsNothing(LbL_Formatting.Text) Then arr_format = LbL_Formatting.Text.Split(","c)
            sql = sql.Replace("[year]", report_year)

            Load_Datagridview(Dgv_Rapportage_Overzicht, sql, "BankTree.NodeMouseClick-level1")
            Format_Datagridview(Dgv_Rapportage_Overzicht, arr_format, False)

        ElseIf e.Node.Level = 2 Then
            'raadpleging van archief, eerder rapportagejaar

            'report_year = e.Node.Text
            rep = e.Node.Parent.Text
            Lbl_Rapportage.Text = rep
            sql = QuerySQL($"Select sql from query where category ilike 'Overzicht%' and name='{rep}'")
            If IsNothing(sql) Then Exit Sub

            formatting = QuerySQL($"Select formatting from query where category ilike 'Overzicht%' and name='{rep}'")
            LbL_Formatting.Text = formatting
            If Not IsNothing(LbL_Formatting.Text) Then arr_format = LbL_Formatting.Text.Split(","c)

            sql = sql.Replace("from bank ", "from bank_archive ")
            sql = sql.Replace("from journal ", "from journal_archive ")
            sql = sql.Replace("[year]", report_year)

            Load_Datagridview(Dgv_Rapportage_Overzicht, sql, "BankTree.NodeMouseClick-level2")
            Format_Datagridview(Dgv_Rapportage_Overzicht, arr_format, False)

        End If

    End Sub



    Private Sub Btn_Rap_Expand_Collapse_Click(sender As Object, e As EventArgs) Handles Btn_Rap_Expand_Collapse.Click
        If Btn_Rap_Expand_Collapse.Text = "Alles uitklappen" Then
            BankTree.ExpandAll()
            Btn_Rap_Expand_Collapse.Text = "Alles inklappen"
        Else
            BankTree.CollapseAll()
            Btn_Rap_Expand_Collapse.Text = "Alles uitklappen"
        End If

    End Sub



    Private Sub Dgv_Rapportage_Overzicht_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles Dgv_Rapportage_Overzicht.CellContentClick
        Dim selectedNode As TreeNode = BankTree.SelectedNode

        If selectedNode.Text = "Jaaroverzicht Bank" Then
            If Dgv_Rapportage_Overzicht.CurrentCell.ColumnIndex = 1 Then
                'Clipboard.SetText(Dgv_Rapportage_Overzicht.CurrentRow.Cells(1).Value)
                'TC_Main.SelectedIndex = 4
                'Cmx_Journal_List.Text = "Accountgroep"
                'Searchbox.Text = Dgv_Rapportage_Overzicht.CurrentRow.Cells(1).Value
            Else
                Drill_down_Bank_overview(Me.Dgv_Rapportage_Overzicht.CurrentCell.RowIndex, Me.Dgv_Rapportage_Overzicht.CurrentCell.ColumnIndex)
            End If
        ElseIf selectedNode.Text = "Jaarrapportage" Then
            Drill_down_Report_overview(Dgv_Rapportage_Overzicht.CurrentCell.RowIndex, Dgv_Rapportage_Overzicht.CurrentCell.ColumnIndex)
        End If
    End Sub


    Private Sub dgv_1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles Dgv_Rapportage_Overzicht.CellDoubleClick
        ' Check if the clicked cell is valid
        If e.RowIndex >= 0 AndAlso e.ColumnIndex >= 0 Then
            ' Get the column header text
            Dim columnName As String = Dgv_Rapportage_Overzicht.Columns(e.ColumnIndex).HeaderText

            If columnName = "Accountname" Then

                Dim sql = $"Select j.date As Datum, j.name As Journaalnaam
                    ,case when amt1>0::money then amt1 else 0::money end As Bij
                    ,case when amt1<=0::money then amt1 else 0::money end As Af
                    ,j.status As Status, j.type As Journaaltype, j.source AS Bron
                    from journal j left join account a on a.id = j.fk_account where a.name = '{Dgv_Rapportage_Overzicht.CurrentCell.Value}'
                    order by date
"
                Load_Datagridview(Dgv_Report_6, sql, "Dgv_Rapportage_Overzicht.DoubleClick")
            End If

        End If
    End Sub

    Private Sub Cmbx_Reporting_Year_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Cmbx_Reporting_Year.SelectedIndexChanged
        Try
            report_year = Cmbx_Reporting_Year.SelectedItem
            Run_BankTree(Lbl_Rapportage.Text)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub


End Class



