Imports System.IO
Imports System.Xml
Imports Npgsql
Public Class SPAS
    Private PreviousTab As Integer
    Private oldend_date As Date
    'bekende fouten

    Private Sub SPAS_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Login.Cmx_Login_Database.Text = "Productie"

        Login.ShowDialog()
        'InitLoad()
        'ook gebruiken na bewaren van nieuwe cp, bankacc, relation en target
    End Sub
    Sub InitLoad()
        If username = "" Then Exit Sub
        RunSQL("Update contract Set Active ='false' 
                where enddate < current_date", "NULL", "SPAS_Load")
        Load_Comboboxes()
        TC_Object.SelectedIndex = 1
        Select_Obj2()
        Load_Table()
        If Lbx_Basis.Items.Count = 0 Then Empty_Tabpage()
        nocat = QuerySQL("SELECT value FROM settings WHERE label='nocat'")
        Get_Settings_Data()
        'Clipboard.Clear()
        'Clipboard.SetText(totsql)
        Load_Datagridview(Dgv_Rapportage, "select * from public.reports", "rapportagefout")
        Format_dvg_reports()

    End Sub


    Sub Load_Comboboxes()
        'can go wrong if tables are empty

        Load_Combobox(Cmx_01_cp__fk_bankacc_id, "id", "name", "SELECT id, CONCAT(Name, '/', accountno) as name FROM bankacc WHERE expense=False AND active=TRUE ORDER BY name")
        Load_Combobox(Cmx_Incasso_Bankaccount, "id", "name", "SELECT id, accountno AS name FROM bankacc WHERE expense=FALSE AND active=TRUE ORDER BY name")
        Load_Combobox(Cmx_01_Target__fk_cp_id, "id", "name", "SELECT id, CONCAT(name, ', ', name_add) as name FROM cp WHERE active=True ORDER BY name")
        Load_Combobox(Cmx_00_contract__fk_relation_id, "id", "name", "SELECT id, CONCAT(name, ', ', name_add) as name FROM relation WHERE active=TRUE ORDER BY name")
        Load_Combobox(Cmx_00_Account__accgroup, "accgroup", "name", "SELECT DISTINCT accgroup FROM account WHERE accgroup <> '' ORDER BY accgroup")
        Load_Combobox(Cmx_Bank_bankacc, "id", "name", "SELECT id, CONCAT(Name, '/', accountno) as name FROM bankacc ORDER BY name DESC")
        Load_Combobox(Cmx_Bank_Account, "id", "name", "SELECT id, name FROM account ORDER BY source, name")
        Load_Combobox(Cmx_00_Contract__fk_account_id, "id", "name", "SELECT id, CONCAT(id, ' ',name) As name FROM account 
                                          WHERE active=TRUE AND source='cat' AND type = 'Inkomsten' ORDER BY name")
        'Clipboard.Clear()
        'Clipboard.SetText("SELECT id, CONCAT(name, ', ', name_add) as name FROM target WHERE active=TRUE ORDER BY name")
        If Me.Dgv_Mgnt_Tables.Rows(8).Cells(1).Value > 0 Then

            Load_Combobox(Cmx_01_contract__fk_target_id, "id", "name", "SELECT id, name||', '||name_add as name FROM target WHERE active=TRUE ORDER BY name")
        End If
        '@@@ hier gaat iets fout
        Fill_Cmx_Excasso_Select_Combined()
        'fill other account comboboxes based on Cmx_Bank_Account -later to be added

    End Sub
    Private Sub TC_Object_Click(sender As Object, e As EventArgs) Handles TC_Object.Click

        'If Edit_Mode Or Add_Mode Then
        If Btn_Basis_Save.Enabled Then
            MsgBox("U bent nog bezig met het " & IIf(Edit_Mode, "bewerken", "aanmaken") & " van een " & TC_Object.TabPages(PreviousTab).Text.ToLower & ".")
            PreviousTab = TC_Object.SelectedIndex
        Else
            Load_Table()
        End If
    End Sub
    Private Sub TC_Object_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TC_Object.SelectedIndexChanged
        If (Not Edit_Mode) And (Not Add_Mode) Then
            'Load_Table()

        Else
            If TC_Object.SelectedIndex <> PreviousTab Then
                'MsgBox("U hebt de vorige bewerking nog niet afgerond.")
            End If
        End If

    End Sub
    Private Sub TC_Object_Selecting(sender As Object, e As TabControlCancelEventArgs) Handles TC_Object.Selecting
        If Btn_Basis_Save.Enabled Then TC_Object.SelectedIndex = PreviousTab
    End Sub
    Private Sub Rbtn_Target_Child_CheckedChanged(sender As Object, e As EventArgs) Handles Rbtn_Target_Child.CheckedChanged
        If Btn_Basis_Save.Enabled Then Tbx_01_Target__ttype.Text = Rbtn_Target_Child.Text
    End Sub

    Private Sub Rbtn_Target_Elder_CheckedChanged(sender As Object, e As EventArgs) Handles Rbtn_Target_Elder.CheckedChanged
        If Btn_Basis_Save.Enabled Then Tbx_01_Target__ttype.Text = Rbtn_Target_Elder.Text
    End Sub

    Private Sub Rbtn_Target_Other_CheckedChanged(sender As Object, e As EventArgs) Handles Rbtn_Target_Other.CheckedChanged
        If Btn_Basis_Save.Enabled Then Tbx_01_Target__ttype.Text = Rbtn_Target_Other.Text
    End Sub

    Private Sub Btn_Basis_Add_Click(sender As Object, e As EventArgs) Handles Btn_Basis_Add.Click
        Dim t As String = TC_Object.SelectedIndex.ToString


        Empty_Tabpage()
        Add_Mode = True
        Manage_Buttons_Target(False, False, False, True, True)

        If TC_Object.SelectedIndex = 0 Then  'additional functionality for contract management

            Dtp_31_contract__startdate.Value = Date.Today
            Me.Rbn_00_contract_child.Checked = True
            Rbn_00_contract_child.Checked = True
            '---------------- Temp solution of error
            Lbl_00_Contract__name.Text = Contract_number("K")
            Load_Combobox(Cmx_01_contract__fk_target_id, "id", "name", "Select id, Name||', '||name_add as name FROM target
                                                        WHERE ttype='" & Rbn_00_contract_child.Text & "' ORDER BY name")
            '-------standaard_waarden ophalen

            Tbx_11_Contract__donation.Text = Tbx_Settings_Bedrag_Kind.Text
            Tbx_11_contract__overhead.Text = Tbx_Settings_Overhead_Kind.Text

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
        Select_Obj2()
        Manage_Buttons_Target(True, True, True, False, False)
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

    Private Sub Btn_Basis_Delete_Click(sender As Object, e As EventArgs)
        Edit_Mode = False
    End Sub

    Private Sub Btn_Basis_Save_Click(sender As Object, e As EventArgs) Handles Btn_Basis_Save.Click
        Dim tbl As String = Me.TC_Object.TabPages(Me.TC_Object.SelectedIndex).Name
        Dim val, val2 As Integer
        Dim errmsg = Handle_errors("")
        If errmsg <> "" Then
            MsgBox(errmsg)
            Exit Sub
        End If

        If Lbx_Basis.SelectedIndex <> -1 Then val = Lbx_Basis.SelectedItem(Me.Lbx_Basis.ValueMember)

        If TC_Object.SelectedIndex > 0 Then
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
        Else   'contracten kennen versies en wijken daardoor af van de standaardprocedures
            Dim acid = QuerySQL("SELECT id FROM account WHERE f_key='" & Cmx_01_contract__fk_target_id.SelectedValue & "'")
            MsgBox(acid)
            Calculate_Budget(acid)

            If Add_Mode Then
                Insert_into_table() 'regular adding to database

                If Strings.Len(Cmx_00_Contract__fk_account_id.Text) > 0 Then 'adding for internal contracts...
                    'Dim Source_Account = QuerySQL("SELECT id FROM account WHERE f_key='" & Cmx_01_contract__fk_target_id.SelectedValue & "'")
                    'Add_Internal_Contract_Bookings(Cmx_00_Contract__fk_account_id.SelectedValue, Source_Account,
                    'Tbx_contract_period_amt.Text, Cmx_01_contract__fk_target_id.Text, Cmx_00_contract__fk_relation_id.SelectedValue)
                End If

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

                    Clipboard.Clear()
                    Clipboard.SetText(sqlstr)
                    RunSQL(sqlstr, "NULL", "Btn_Basis_Save.Click upsert new version")

                    val2 = Convert.ToInt32(QuerySQL("Select MAX(id) FROM " & tbl))
                    '3 update new version with new values, startdate / enddate and active
                    sqlstr = "UPDATE contract SET startdate='" & d1 & "', 
                           donation='" & Tbx2Dec(Tbx_11_Contract__donation.Text) & "', 
                           overhead='" & Tbx2Dec(Tbx_11_contract__overhead.Text) & "', 
                           enddate ='2999-12-31',active=not " & act & " 
                           WHERE id=" & val2 & ";"
                    'Clipboard.Clear()
                    'Clipboard.SetText(sqlstr)

                    RunSQL(sqlstr, "NULL", "Btn_Basis_Save.Click update New version")
                    reload = True
                    msg = "Een nieuwe versie van het contract is aangemaakt."
                    If act Then
                        msg &= "De wijziging gaat in in de toekomst (nu nog inactief); wilt u de laatste versie nu bekijken?"

                    End If

                    If MsgBox(msg, IIf(act, vbYesNo, vbOK)) = vbYes Then
                        Rbn_contract_inactive.Checked = True
                        Tbx_Basis_Filter.Text = Lbl_00_Contract__name.Text

                        reload = True
                    Else
                        val = val2
                        reload = True
                    End If

                    Pan_Contract_Date_New.Visible = False
                Else
                    'updating description in the regular way
                    val = Lbx_Basis.SelectedItem(Me.Lbx_Basis.ValueMember)
                    Update_table()
                End If
            End If
        End If

        If reload Then
            Load_Table()
            Locate_Listbox_Position(val)
        End If
        'finalizing
        Load_Comboboxes()
        Manage_Buttons_Target(True, True, True, False, False)
        Edit_Mode = False
        Add_Mode = False
        reload = False
    End Sub

    Sub Manage_Buttons_Target(ByVal a As Boolean, e As Boolean, d As Boolean, s As Boolean, c As Boolean)

        If Rbn_contract_inactive.Checked And Edit_Mode Then
            MsgBox("Inactieve objecten kunnen niet gewijzigd worden.")
            Exit Sub
        End If
        Btn_Basis_Add.Enabled = a
        Btn_Basis_Save.Enabled = s
        Btn_Basis_Cancel.Enabled = c
        Lbx_Basis.Enabled = a

    End Sub

    Private Sub TC_Main_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TC_Main.SelectedIndexChanged


        Select Case TC_Main.SelectedIndex
            Case 0
                'Pan_contract_select_target.Enabled = False
            Case 1
                Calculate_Bank_Balance()
                If Dgv_Bank.Rows.Count = 0 Or Dgv_Bank.DataSource Is Nothing Then
                    If Me.Dgv_Mgnt_Tables.Rows(1).Cells(1).Value > 0 Then
                        Cmx_Bank_bankacc.SelectedIndex = 0
                        Fill_bank_transactions()

                        Format_dvg_bank_journal()
                    End If
                End If

            Case 2
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

                If Me.Dgv_Mgnt_Tables.Rows(3).Cells(1).Value > 0 And
                    Me.Dgv_Mgnt_Tables.Rows(5).Cells(1).Value > 0 And
                    Me.Dgv_Mgnt_Tables.Rows(8).Cells(1).Value > 0 And
                    Me.Cmx_Excasso_Select.SelectedItem = "" Then

                    Dtp_Excasso_Start.ShowUpDown = False
                    Dtp_Excasso_Start.Value = Date.Today
                    Dtp_Excasso_Start.MaxDate = Date.Today
                    Fill_Cmx_Excasso_Select_Combined()
                End If

            Case 4
                Cmx_Journal_List.Text = "Alle accounts"
                Me.Dtp_Journal_intern.Value = Date.Now

            Case 5
                Format_dvg_reports()
            Case 6



        End Select

    End Sub
    Sub Get_Settings_Data()
        Collect_data("SELECT * FROM settings ORDER BY label")
        Tbx_Settings_Bedrag_Kind.Text = dst.Tables(0).Rows(7)(1)
        Tbx_Settings_Bedrag_Oudere.Text = dst.Tables(0).Rows(8)(1)
        Tbx_Settings_Overhead_Kind.Text = dst.Tables(0).Rows(9)(1)
        Tbx_Settings_Overhead_Oudere.Text = dst.Tables(0).Rows(10)(1)
        Tbx_Settings_Banktext_Kind.Text = dst.Tables(0).Rows(0)(1)
        Tbx_Settings_Banktext_Oudere.Text = dst.Tables(0).Rows(2)(1)
    End Sub
    Private Sub Lbx_Basis_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Lbx_Basis.SelectedIndexChanged
        If Lbx_Basis.Items.Count > 0 Then Select_Obj2() Else Empty_Tabpage()

    End Sub
    Private Sub Tbx_Target__ttype_TextChanged(sender As Object, e As EventArgs) Handles Tbx_01_Target__ttype.TextChanged
        Rbtn_Target_Child.Checked = Strings.Trim(Tbx_01_Target__ttype.Text) = "Kind"
        Rbtn_Target_Elder.Checked = Strings.Trim(Tbx_01_Target__ttype.Text) = "Oudere"
        Rbtn_Target_Other.Checked = Strings.Trim(Tbx_01_Target__ttype.Text) = "Overig"
        '@@@ hard value vervangen door tt_type.Text
    End Sub
    Private Sub Rbtn_Target_Alone_CheckedChanged(sender As Object, e As EventArgs) Handles Rbtn_Target_Alone.CheckedChanged
        If Btn_Basis_Save.Enabled Then Tbx_00_Target__living.Text = Rbtn_Target_Alone.Text
    End Sub

    Private Sub Rbtn_Target_Institution_CheckedChanged(sender As Object, e As EventArgs) Handles Rbtn_Target_Institution.CheckedChanged
        If Btn_Basis_Save.Enabled Then Tbx_00_Target__living.Text = Rbtn_Target_Institution.Text
    End Sub

    Private Sub Rbtn_Target_OtherHousing_CheckedChanged(sender As Object, e As EventArgs) Handles Rbtn_Target_OtherHousing.CheckedChanged
        If Btn_Basis_Save.Enabled Then Tbx_00_Target__living.Text = Rbtn_Target_OtherHousing.Text
    End Sub

    Private Sub Tbx_Target__living_TextChanged(sender As Object, e As EventArgs) Handles Tbx_00_Target__living.TextChanged
        Rbtn_Target_Alone.Checked = Strings.Trim(Tbx_00_Target__living.Text) = "Alleen"
        Rbtn_Target_Institution.Checked = Strings.Trim(Tbx_00_Target__living.Text) = "Tehuis"
        Rbtn_Target_OtherHousing.Checked = Strings.Trim(Tbx_00_Target__living.Text) = "Overig"
    End Sub
    Private Sub Tbx_Target__name_TextChanged(sender As Object, e As EventArgs) Handles Tbx_01_Target__name.TextChanged
        If Edit_Mode Then Manage_Buttons_Target(False, False, False, True, True)
        reload = True
    End Sub

    Private Sub Tbx_Target__name_add_TextChanged(sender As Object, e As EventArgs) Handles Tbx_01_Target__name_add.TextChanged
        If Edit_Mode Then Manage_Buttons_Target(False, False, False, True, True)
        reload = True
    End Sub
    Private Sub Tbx_Target__address_TextChanged(sender As Object, e As EventArgs) Handles _
        Tbx_00_Target__address.TextChanged, Dtp_31_contract__startdate.TextChanged,
        Tbx_00_Target__zip.TextChanged, Tbx_00_Target__city.TextChanged, Tbx_00_Target__country.TextChanged,
        Tbx_20_Target__children.TextChanged, Tbx_20_Target__childnearby.TextChanged, Tbx_00_Target__description.TextChanged,
        Dtp_00_Target__birthday.ValueChanged, Cmx_01_Target__fk_cp_id.SelectedIndexChanged

        If Edit_Mode Then Manage_Buttons_Target(False, False, False, True, True)
    End Sub


    Private Sub Tbx_Target__income_Enter(sender As Object, e As EventArgs) Handles _
        Tbx_10_Target__income.Enter, Tbx_10_Target__pension.Enter, Tbx_10_Target__benefit.Enter,
        Tbx_10_Target__allowance.Enter, Tbx_10_Target__otherincome.Enter, Tbx_10_Target__rent.Enter,
        Tbx_10_Target__heating.Enter, Tbx_10_Target__pension.Enter, Tbx_10_Target__gaselectra.Enter,
        Tbx_10_Target__water.Enter, Tbx_10_Target__food.Enter, Tbx_10_Target__medicine.Enter,
        Cmx_01_Target__fk_cp_id.Click, Dtp_00_Target__birthday.Click,
        Tbx_01_Target__name_add.Enter, Tbx_01_Target__name.Enter, Tbx_00_Target__zip.Enter,
        Tbx_00_Target__address.Enter, Tbx_00_Target__city.Enter, Tbx_00_Target__country.EnabledChanged,
        Tbx_20_Target__children.Enter, Tbx_20_Target__childnearby.Enter, Tbx_00_Target__country.Enter,
        Tbx_00_Target__description.Enter, Tbx_01_CP__name.Enter, Tbx_01_CP__name_add.Enter,
        Tbx_01_BankAcc__accountno.Enter, Tbx_01_BankAcc__name.Enter, Tbx_01_BankAcc__owner.Enter,
        Tbx_10_BankAcc__startbalance.Enter

        Edit_Mode = True
    End Sub


    Private Sub Tbx_Target__income_TextChanged(sender As Object, e As EventArgs) Handles _
        Tbx_10_Target__income.TextChanged, Tbx_10_Target__pension.TextChanged, Tbx_10_Target__benefit.TextChanged,
        Tbx_10_Target__allowance.TextChanged, Tbx_10_Target__otherincome.TextChanged,
        Tbx_10_Target__rent.TextChanged, Tbx_10_Target__heating.TextChanged, Tbx_10_Target__heating.TextChanged,
        Tbx_10_Target__gaselectra.TextChanged, Tbx_10_Target__water.TextChanged, Tbx_10_Target__food.TextChanged,
        Tbx_10_Target__medicine.TextChanged
        If Edit_Mode Then Manage_Buttons_Target(False, False, False, True, True)
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
    Private Sub Tbx_Basis_Filter_TextChanged(sender As Object, e As EventArgs) Handles Tbx_Basis_Filter.TextChanged
        Load_Table()
    End Sub

    Private Sub Tbx_CP__bankaccount_Enter(sender As Object, e As EventArgs)
        Edit_Mode = True
    End Sub

    Private Sub Tbx_CP__email_Enter(sender As Object, e As EventArgs) Handles Tbx_00_CP__email.Enter,
        Tbx_00_CP__telephone.Enter, Tbx_00_CP__address.Enter, Tbx_00_CP__zip.Enter, Tbx_00_CP__city.Enter,
        Tbx_00_CP__country.Enter
        Edit_Mode = True
    End Sub

    Private Sub Tbx_CP__telephone_Enter(sender As Object, e As EventArgs) Handles Tbx_00_CP__telephone.Enter
        Edit_Mode = True
        'StatusText.Text = Edit_Mode
    End Sub

    Private Sub Tbx_CP__name_TextChanged(sender As Object, e As EventArgs) Handles Tbx_01_CP__name.TextChanged
        If Edit_Mode Then Manage_Buttons_Target(False, False, False, True, True)
        reload = True
        If Lbx_Basis.Items.Count = 0 Then Add_Mode = True
    End Sub

    Private Sub Tbx_CP__name_add_TextChanged(sender As Object, e As EventArgs) Handles Tbx_01_CP__name_add.TextChanged
        If Edit_Mode Then Manage_Buttons_Target(False, False, False, True, True)
        reload = True
    End Sub

    Private Sub Tbx_CP__bankaccount_TextChanged(sender As Object, e As EventArgs)
        If Edit_Mode Then Manage_Buttons_Target(False, False, False, True, True)
    End Sub

    Private Sub Tbx_CP__email_TextChanged(sender As Object, e As EventArgs) Handles Tbx_00_CP__email.TextChanged,
            Tbx_00_CP__telephone.TextChanged, Tbx_00_CP__address.TextChanged, Tbx_00_CP__zip.TextChanged,
            Tbx_00_CP__city.TextChanged, Tbx_00_CP__country.TextChanged
        If Edit_Mode Then Manage_Buttons_Target(False, False, False, True, True)
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        Lv_Journal_List.Items(0).Selected = True

        Lbx_Basis.SetSelected(2, True)


        Exit Sub
        Dim str As String = "one,two,three"
        Dim str2() As String = Split(str, ",")

        'Load_Listbox(Me.Lbx_Basis, "Select id, name FROM Bankacc WHERE name ilike '%%' AND active=True ORDER BY name")
        'MsgBox(Me.Dgv_Mgnt_Tables.Rows(1).Cells(0).Value)
        Dim amt1 = QuerySQL("select max(startbalance) from bankacc")
        'MsgBox(amt1)
        'Dim amt = InputBox("enter amount: ")
        'MsgBox(FormatCurrency(amt1 / 100, 2, TriState.True, TriState.True))

        'Dim a1 = QuerySQL("SELECT startbalance FROM bankacc WHERE id=3")
        'MsgBox(a1)
        'Dim objwebbrowser As New WebBrowser
        'objwebbrowser.Navigate("https://www.xe.com/currencyconverter/convert/?Amount=1&From=EUR&To=MDL")
        'AddHandler objwebbrowser.DocumentCompleted, AddressOf navigation_complete

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
        If Btn_Basis_Save.Enabled Then Tbx_01_Account__type.Text = Rbtn_Account_Income.Text
    End Sub

    Private Sub Rbtn_Account_Transit_CheckedChanged(sender As Object, e As EventArgs) Handles Rbtn_Account_Transit.CheckedChanged
        If Btn_Basis_Save.Enabled Then Tbx_01_Account__type.Text = Rbtn_Account_Transit.Text
    End Sub

    Private Sub Rbtn_Account_Expense_CheckedChanged(sender As Object, e As EventArgs) Handles Rbtn_Account_Expense.CheckedChanged
        If Btn_Basis_Save.Enabled Then Tbx_01_Account__type.Text = Rbtn_Account_Expense.Text
    End Sub

    Private Sub Tbx_Account__type_TextChanged(sender As Object, e As EventArgs) Handles Tbx_01_Account__type.TextChanged
        Rbtn_Account_Expense.Checked = Tbx_01_Account__type.Text = "Uitgaven"
        Rbtn_Account_Income.Checked = Tbx_01_Account__type.Text = "Inkomsten"
        Rbtn_Account_Transit.Checked = Tbx_01_Account__type.Text = "Transit"
    End Sub
    Private Sub Tbx_BankAcc__accountno_Leave(sender As Object, e As EventArgs) Handles Tbx_01_BankAcc__accountno.Leave
        If Tbx_01_BankAcc__accountno.Text = "" Then Exit Sub
        Tbx_01_BankAcc__accountno.Text = Tbx_01_BankAcc__accountno.Text.ToUpper
        If IBANcheck(Tbx_01_BankAcc__accountno.Text) <> 1 Then
            MsgBox("Bankrekeningnummer Is niet correct", vbCritical)
            Tbx_01_BankAcc__accountno.Focus()
        End If
    End Sub



    Private Sub Tbx_BankAcc__accountno_TextChanged(sender As Object, e As EventArgs) Handles Tbx_01_BankAcc__accountno.TextChanged
        If Edit_Mode Then Manage_Buttons_Target(False, False, False, True, True)
        reload = True
    End Sub
    Private Sub Tbx_BankAcc__label_TextChanged(sender As Object, e As EventArgs) Handles Tbx_01_BankAcc__name.TextChanged
        If Edit_Mode Then
            Manage_Buttons_Target(False, False, False, True, True)
            reload = True
        End If
    End Sub
    Private Sub Tbx_BankAcc__owner_TextChanged(sender As Object, e As EventArgs) Handles Tbx_01_BankAcc__owner.TextChanged,
Tbx_00_BankAcc__id2.TextChanged, Tbx_00_BankAcc__bic.TextChanged, Tbx_10_BankAcc__startbalance.TextChanged, Tbx_00_BankAcc__description.TextChanged
        If Edit_Mode Then Manage_Buttons_Target(False, False, False, True, True)
    End Sub
    Private Sub Chbx_BankAcc__income_CheckedChanged(sender As Object, e As EventArgs) Handles Cbx_00_BankAcc__income.CheckedChanged,
            Chx_00_BankAcc__expense.CheckedChanged
        If Edit_Mode Then Manage_Buttons_Target(False, False, False, True, True)
    End Sub

    Private Sub Cmbx_BankAcc__currency_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Cmx_01_BankAcc__currency.SelectedIndexChanged
        If Edit_Mode Then Manage_Buttons_Target(False, False, False, True, True)
    End Sub
    Private Sub Chbx_BankAcc__income_Enter(sender As Object, e As EventArgs) Handles Cbx_00_BankAcc__income.Enter,
        Tbx_00_contract__description.Enter, Tbx_00_BankAcc__id2.Enter, Tbx_00_BankAcc__bic.Enter,
        Chx_00_BankAcc__expense.Enter, Cmx_01_BankAcc__currency.Enter, Tbx_00_BankAcc__description.Enter,
        Rbtn_Account_Income.Enter, Rbtn_Account_Expense.Enter, Rbtn_Account_Transit.Enter

        Edit_Mode = True
    End Sub
    Private Sub Cmx_00_Account__accgroup_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Cmx_00_Account__accgroup.SelectedIndexChanged
        If Edit_Mode Then Manage_Buttons_Target(False, False, False, True, True)
        reload = True
    End Sub

    Private Sub Cmx_01_cp__bankaccount_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Cmx_01_cp__fk_bankacc_id.SelectedIndexChanged
        If Edit_Mode Then Manage_Buttons_Target(False, False, False, True, True)
    End Sub
    Private Sub Tbx_00_cp__description_TextChanged(sender As Object, e As EventArgs) Handles Tbx_00_cp__description.TextChanged
        If Edit_Mode Then Manage_Buttons_Target(False, False, False, True, True)
    End Sub
    Private Sub Tbx_00_cp__description_Enter(sender As Object, e As EventArgs) Handles Tbx_00_cp__description.Enter
        Edit_Mode = True
    End Sub

    Private Sub Cmx_01_cp__bankaccount_Click(sender As Object, e As EventArgs) Handles Cmx_01_cp__fk_bankacc_id.Click
        Edit_Mode = True
    End Sub


    Private Sub Tbx_10_Relation__name_TextChanged(sender As Object, e As EventArgs) Handles Tbx_01_relation__name.TextChanged
        If Edit_Mode Then
            Manage_Buttons_Target(False, False, False, True, True)
            reload = True
        End If

        If Add_Mode Then Generate_Reference()
        If Lbx_Basis.Items.Count = 0 Then Add_Mode = True
    End Sub

    Private Sub Tbx_10_Relation__name_add_TextChanged(sender As Object, e As EventArgs) Handles Tbx_01_Relation__name_add.TextChanged
        If Edit_Mode Then
            Manage_Buttons_Target(False, False, False, True, True)
            reload = True
        End If
    End Sub
    Private Sub Tbx_00_Relation__iban_TextChanged(sender As Object, e As EventArgs) Handles _
            Tbx_00_Relation__iban.TextChanged, Tbx_00_Relation__email.TextChanged, Tbx_00_Relation__phone.TextChanged,
            Tbx_00_Relation__address.TextChanged, Tbx_00_Relation__zip.TextChanged, Tbx_00_Relation__city.TextChanged,
            Tbx_00_Relation__description.TextChanged, Tbx_00_contract__description.TextChanged
        If Edit_Mode Then Manage_Buttons_Target(False, False, False, True, True)
    End Sub
    Private Sub Tbx_10_Relation__name_add_Enter(sender As Object, e As EventArgs) Handles _
        Tbx_01_Relation__name_add.Enter, Tbx_00_Relation__iban.Enter, Tbx_00_Relation__email.Enter,
        Tbx_00_Relation__phone.Enter, Tbx_00_Relation__address.Enter, Tbx_00_Relation__zip.Enter,
        Tbx_00_Relation__city.Enter, Tbx_00_Relation__description.Enter, Tbx_01_relation__name.Enter,
        Dtp_31_contract__startdate.Enter
        Edit_Mode = True
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
        Tbx_11_Contract__donation.Text = Tbx_Settings_Bedrag_Kind.Text
        Tbx_11_contract__overhead.Text = Tbx_Settings_Overhead_Kind.Text

        '----------------------------

    End Sub
    Private Sub Tbx_10_Contract__transport_TextChanged(sender As Object, e As EventArgs)
        If Edit_Mode Then
            Manage_Buttons_Target(False, False, False, True, True)
        End If
        'Pan_Contract_Date_New.Visible = Not Add_Mode
        Calculate_contract_amounts()
    End Sub
    Private Sub Tbx_11_Contract__donation_TextChanged(sender As Object, e As EventArgs) Handles Tbx_11_Contract__donation.TextChanged
        If Edit_Mode Then
            Manage_Buttons_Target(False, False, False, True, True)
        End If

        'Pan_Contract_Date_New.Visible = Not Add_Mode
        Calculate_contract_amounts()
    End Sub

    Private Sub Tbx_11_contract__overhead_TextChanged(sender As Object, e As EventArgs) Handles Tbx_11_contract__overhead.TextChanged
        If Edit_Mode Then
            Manage_Buttons_Target(False, False, False, True, True)
        End If
        Calculate_contract_amounts()
    End Sub
    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs)
        Calculate_contract_amounts()
    End Sub
    Private Sub Rbn_contract_active_Click(sender As Object, e As EventArgs) Handles Rbn_contract_active.Click,
            Rbn_contract_all.Click, Rbn_contract_inactive.Click
        Btn_Basis_Delete.Enabled = Rbn_contract_inactive.Checked
        Load_Table()
    End Sub
    Private Sub Cbx_00_BankAcc__active_Enter(sender As Object, e As EventArgs) Handles _
        Cbx_00_BankAcc__active.Enter, Cbx_00_Account__active.Enter, Cbx_00_cp__active.Enter
        Edit_Mode = True
    End Sub
    Private Sub Cbx_00_Account__active_Enter(sender As Object, e As EventArgs) Handles Cbx_00_Account__active.Enter
        'Edit_Mode = True
    End Sub
    Private Sub Cbx_00_cp__active_Enter(sender As Object, e As EventArgs) Handles Cbx_00_cp__active.Enter
        'Edit_Mode = True
    End Sub
    Private Sub Cbx_00_relation__active_Enter(sender As Object, e As EventArgs) Handles Cbx_00_relation__active.Enter
        'Edit_Mode = True
    End Sub
    Private Sub Cbx_00_target__active_Enter(sender As Object, e As EventArgs) Handles Cbx_00_target__active.Enter
        'Edit_Mode = True
    End Sub
    Private Sub Cbx_00_BankAcc__active_CheckedChanged(sender As Object, e As EventArgs) Handles _
            Cbx_00_BankAcc__active.CheckedChanged, Rbtn_Account_Income.CheckedChanged,
            Rbtn_Account_Expense.CheckedChanged, Rbtn_Account_Transit.CheckedChanged
        If Edit_Mode Then Manage_Buttons_Target(False, False, False, True, True)
    End Sub

    Private Sub Cbx_00_Account__active_CheckedChanged(sender As Object, e As EventArgs) Handles Cbx_00_Account__active.CheckedChanged
        If Edit_Mode Then Manage_Buttons_Target(False, False, False, True, True)
    End Sub

    Private Sub Cbx_00_cp__active_CheckedChanged(sender As Object, e As EventArgs) Handles Cbx_00_cp__active.CheckedChanged
        If Edit_Mode Then Manage_Buttons_Target(False, False, False, True, True)


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

    Private Sub Tbx_10_Contract__transport_Enter(sender As Object, e As EventArgs)
        Edit_Mode = True
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
        Dtp_31_contract__enddate.Value = Date.Today
    End Sub

    Private Sub Dtp_01_contract__enddate_Leave(sender As Object, e As EventArgs) Handles Dtp_31_contract__enddate.Leave

        Dim ans = MsgBox("Weet u zeker dat u dit contract wilt beëindigen per " & Dtp_31_contract__enddate.Value & "?", vbYesNo)
        If ans = vbNo Then
            Dtp_31_contract__enddate.Value = oldend_date
        End If

    End Sub
    Private Sub Tbx_11_Contract__donation_Enter(sender As Object, e As EventArgs) Handles Tbx_11_Contract__donation.Enter
        Edit_Mode = True
    End Sub

    Private Sub Tbx_11_contract__overhead_Enter(sender As Object, e As EventArgs) Handles Tbx_11_contract__overhead.Enter
        Edit_Mode = True
    End Sub

    Private Sub Rbn_00_contract_elder_Click(sender As Object, e As EventArgs) Handles Rbn_00_contract_elder.Click
        Tbx_Contract_ttype.Text = "Oudere"
        Lbl_00_Contract__name.Text = Contract_number("O")
        Load_Combobox(Cmx_01_contract__fk_target_id, "id", "name", "SELECT id, name||', '||name_add as name FROM target
                                                        WHERE ttype='" & Rbn_00_contract_elder.Text & "' ORDER BY name")
        Tbx_11_Contract__donation.Text = Tbx_Settings_Bedrag_Oudere.Text
        Tbx_11_contract__overhead.Text = Tbx_Settings_Overhead_Oudere.Text
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
            Select_Obj2()
            Manage_Buttons_Target(True, True, True, False, False)
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
            Collect_data("
                        SELECT name, startdate, enddate, donation, overhead, active 
                        FROM contract WHERE name ='" & name & "' ORDER BY startdate DESC
                        ")

            '1 determine whether there is a future version. If so then a change is not allowed (first
            'delete that version
            If CDate(dst.Tables(0).Rows(0)(1)) > Date.Today Then
                MsgBox("Er is een nieuwere versie die nog niet is ingegaan" & vbCrLf &
                       "S.v.p. deze eerst verwijderen (selecteer 'inactief' boven linkerlijst). ")
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
                   "; Automatische incasso kan niet niet geactiveerd worden voor dit contract.", vbCritical)
            Chx_00_contract__autcol.Checked = False
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

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Btn_Bank_Download.Click
        Download_Bank_Transactions()

    End Sub
    Private Sub Tbx_00_Account__searchword_TextChanged(sender As Object, e As EventArgs) Handles Tbx_00_Account__searchword.TextChanged
        If Edit_Mode Then Manage_Buttons_Target(False, False, False, True, True)
    End Sub
    Private Sub Tbx_00_Account__searchword_Enter(sender As Object, e As EventArgs) Handles Tbx_00_Account__searchword.Enter
        Edit_Mode = True
    End Sub
    Private Sub Tbx_00_Account__description_TextChanged(sender As Object, e As EventArgs) Handles Tbx_00_Account__description.TextChanged
        If Edit_Mode Then Manage_Buttons_Target(False, False, False, True, True)
    End Sub
    Private Sub Tbx_00_Account__description_Enter(sender As Object, e As EventArgs) Handles Tbx_00_Account__description.Enter
        Edit_Mode = True
    End Sub

    Private Sub Cmx_00_Account__accgroup_Click(sender As Object, e As EventArgs) Handles Cmx_00_Account__accgroup.Click
        Edit_Mode = True
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

    Private Sub Tbx_01_Account__name_TextChanged(sender As Object, e As EventArgs) Handles Tbx_01_Account__name.TextChanged
        If Edit_Mode Then Manage_Buttons_Target(False, False, False, True, True)
        reload = True
    End Sub

    Private Sub Tbx_01_Account__name_Enter(sender As Object, e As EventArgs) Handles Tbx_01_Account__name.Enter
        Edit_Mode = True
    End Sub

    Private Sub Cmx_00_Account__accgroup_Enter(sender As Object, e As EventArgs) Handles Cmx_00_Account__accgroup.Enter
        Edit_Mode = True
    End Sub

    Private Sub Cmx_01_cp__fk_bankacc_id_Enter(sender As Object, e As EventArgs) Handles Cmx_01_cp__fk_bankacc_id.Enter
        Edit_Mode = True
    End Sub

    Private Sub Cmx_00_Account__accgroup_SelectedValueChanged(sender As Object, e As EventArgs) Handles Cmx_00_Account__accgroup.SelectedValueChanged
        If Edit_Mode Then Manage_Buttons_Target(False, False, False, True, True)
        reload = True
    End Sub

    Private Sub Cmx_00_Account__accgroup_TextUpdate(sender As Object, e As EventArgs) Handles Cmx_00_Account__accgroup.TextUpdate
        If Edit_Mode Then Manage_Buttons_Target(False, False, False, True, True)
        reload = True
    End Sub

    Private Sub Cmx_Bank_bankacc_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Cmx_Bank_bankacc.SelectedIndexChanged
        If Cmx_Bank_bankacc.SelectedIndex <> -1 Then
            Fill_bank_transactions()
            Calculate_Bank_Balance()
        End If
    End Sub

    Private Sub Dgv_Bank_SelectionChanged(sender As Object, e As EventArgs) Handles Dgv_Bank.SelectionChanged

        If Dgv_Bank.Rows.Count = 0 Or Dgv_Bank.DataSource Is Nothing Then Exit Sub

    End Sub
    Private Sub Dgv_Bank_Click(sender As Object, e As EventArgs) Handles Dgv_Bank.Click
        Format_dvg_bank_journal()
        Tbx_Bank_Relation.Text = Dgv_Bank.SelectedCells(2).Value
        Tbx_Bank_Description.Text = Dgv_Bank.SelectedCells(3).Value
        If Not IsDBNull(Dgv_Bank.SelectedCells(8).Value) Then Tbx_Bank_Relation_account.Text = Dgv_Bank.SelectedCells(8).Value
        If Not IsDBNull(Dgv_Bank.SelectedCells(6).Value) Then Tbx_Bank_Code.Text = Dgv_Bank.SelectedCells(6).Value
        Dim ink As Boolean = Dgv_Bank.SelectedCells(4).Value > 0 And
            Strings.Left(Dgv_Bank.SelectedCells(3).Value, 8) <> "Contract"

        Rbn_Bank_Contract.Enabled = ink
        Rbn_Bank_Extra.Enabled = ink
        Rbn_Bank_Other.Enabled = ink
        Rbn_Bank_Contract.Checked = False
        Rbn_Bank_Extra.Checked = False
        Rbn_Bank_Other.Checked = False

        Fill_Journals_by_bank(Dgv_Bank.SelectedCells(0).Value)
        'Tbx_Bank_Accounts_Total.Text = 0
        Calculate_Total_Booked()

        'Tbx_Bank_Amount.Text = Amt_In - Amt_Out - CDec(Tbx_Bank_Accounts_Total.Text)
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
            'MsgBox(QuerySQL(sqlstr))
        Else
            Cmx_Bank_Account.Text = ""
        End If


    End Sub

    Private Sub Dtp_00_relation__date1_ValueChanged(sender As Object, e As EventArgs) Handles _
        Dtp_00_relation__date1.ValueChanged, Dtp_00_relation__date2.ValueChanged, Dtp_00_relation__date3.ValueChanged
        If Edit_Mode Then Manage_Buttons_Target(False, False, False, True, True)
    End Sub

    Private Sub Dtp_00_relation__date1_Enter(sender As Object, e As EventArgs) Handles _
        Dtp_00_relation__date1.Enter, Dtp_00_relation__date2.Enter, Dtp_00_relation__date3.Enter
        Edit_Mode = True
    End Sub

    Private Sub Btn_Bank_Add_Journal_Click(sender As Object, e As EventArgs) Handles Btn_Bank_Add_Journal.Click
        If Tbx_Bank_Amount.Text = "" Then Exit Sub
        If Cmx_Bank_Account.SelectedValue = QuerySQL("Select value from settings where label='nocat'") Then Exit Sub

        Dim R As DataRow
        R = dst.Tables(0).Rows.Add
        R(0) = Cmx_Bank_Account.SelectedValue
        R(1) = Cmx_Bank_Account.Text
        R(2) = Tbx_Bank_Amount.Text

        Calculate_Total_Booked()

    End Sub

    Private Sub Dgv_Test_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles Dgv_Bank_Account.CellContentClick
        Calculate_Total_Booked()
    End Sub

    Private Sub Dgv_Test_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles _
        Dgv_Bank_Account.CellValueChanged  ', Dgv_Bank_Account.Leave
        Calculate_Total_Booked()
    End Sub

    Private Sub Btn_Bank_Save_Accounts_Click(sender As Object, e As EventArgs) Handles Btn_Bank_Save_Accounts.Click
        Save_Banktransaction_Accounts()
        Mark_rows_Dgv_Bank()
    End Sub

    Private Sub Tbx_Bank_Search_TextChanged(sender As Object, e As EventArgs) Handles Tbx_Bank_Search.TextChanged
        Fill_bank_transactions()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        RunSQL("TRUNCATE TABLE bank", "NULL", "")
        RunSQL("Delete From journal WHERE source='Bank'", "NULL", "")
    End Sub

    Private Sub Rbn_Relation_1_CheckedChanged(sender As Object, e As EventArgs) Handles Rbn_Relation_1.Click
        Edit_Mode = True
        Manage_Buttons_Target(False, False, False, True, True)
        If Rbn_Relation_1.Checked Then Tbx_01_Relation__title.Text = Rbn_Relation_1.Text
    End Sub

    Private Sub Rbn_Relation_2_CheckedChanged(sender As Object, e As EventArgs) Handles Rbn_Relation_2.Click
        Edit_Mode = True
        Manage_Buttons_Target(False, False, False, True, True)
        If Rbn_Relation_2.Checked Then Tbx_01_Relation__title.Text = Rbn_Relation_2.Text
    End Sub

    Private Sub Rbn_Relation_3_CheckedChanged(sender As Object, e As EventArgs) Handles Rbn_Relation_3.Click
        Edit_Mode = True
        Manage_Buttons_Target(False, False, False, True, True)
        If Rbn_Relation_3.Checked Then Tbx_01_Relation__title.Text = Rbn_Relation_3.Text
    End Sub

    Private Sub Rbn_Relation_4_CheckedChanged(sender As Object, e As EventArgs) Handles Rbn_Relation_4.Click
        Edit_Mode = True
        Manage_Buttons_Target(False, False, False, True, True)
        If Rbn_Relation_4.Checked Then Tbx_01_Relation__title.Text = ""
    End Sub
    Private Sub Rbn_Relation_5_CheckedChanged(sender As Object, e As EventArgs) Handles Rbn_Relation_5.Click
        Edit_Mode = True
        Manage_Buttons_Target(False, False, False, True, True)
        If Rbn_Relation_5.Checked Then Tbx_01_Relation__title.Text = Rbn_Relation_5.Text
    End Sub
    Private Sub Rbn_Relation_6_CheckedChanged(sender As Object, e As EventArgs) Handles Rbn_Relation_6.Click
        Edit_Mode = True
        Manage_Buttons_Target(False, False, False, True, True)
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
        'Rbn_00_contract_child.Checked = Strings.Trim(Tbx_Contract_ttype.Text) = "Kind"
        'Rbn_00_contract_elder.Checked = Strings.Trim(Tbx_Contract_ttype.Text) = "Oudere"
        'Rbn_00_contract_other.Checked = Strings.Trim(Tbx_Contract_ttype.Text) = "Overig"
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
    End Sub

    Private Sub Rbn_Incasso_SEPA_CheckedChanged(sender As Object, e As EventArgs) Handles Rbn_Incasso_SEPA.CheckedChanged
        If TC_Main.SelectedIndex <> 2 Then Exit Sub
        If Rbn_Incasso_SEPA.Checked Then
            Load_Datagridview(Me.Dgv_Incasso, Create_Incasso(Dtp_Incasso_start.Value), "Dtp_Incasso_start.ValueChanged")
        Else
            Load_Datagridview(Me.Dgv_Incasso, Create_Incasso_Bookings(Dtp_Incasso_start.Value), "Dtp_Incasso_start.ValueChanged")
        End If
    End Sub

    Private Sub Btn_Run_Incasso_Click(sender As Object, e As EventArgs) Handles Btn_Run_Incasso.Click
        Create_Incasso_Journals()
        Create_SEPA_XML()
        Me.Lbl_Incasso_Status.Text = "Open"
        Me.Btn_Incasso_Delete.Enabled = True
        Me.Btn_Run_Incasso.Enabled = False
        Me.Btn_Incasso_Print.Enabled = True
    End Sub

    Private Sub Cbx_Uitkering_Kind_Click(sender As Object, e As EventArgs) Handles Cbx_Uitkering_Kind.Click
        Get_New_Excasso_Data()
        Calculate_CP_Allowance()
    End Sub

    Private Sub Cbx_Uitkering_Oudere_Click(sender As Object, e As EventArgs) Handles Cbx_Uitkering_Oudere.Click
        Get_New_Excasso_Data()
        Calculate_CP_Allowance()
    End Sub

    Private Sub Cbx_Uitkering_Overig_Click(sender As Object, e As EventArgs) Handles Cbx_Uitkering_Overig.Click
        Get_New_Excasso_Data()
        Calculate_CP_Allowance()
    End Sub

    Private Sub Dtp_Excasso_Start_ValueChanged(sender As Object, e As EventArgs) Handles Dtp_Excasso_Start.ValueChanged
        'Dtp_Excasso_Start.Value = CDate("01-" & Dtp_Excasso_Start.Value.Month & "-" & Dtp_Excasso_Start.Value.Year)
        Dtp_Excasso_Start.MaxDate = Date.Today

        'Get_Excasso_Data()
        'dit moet nog aangepast worden

    End Sub

    Private Sub Dgv_Excasso2_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles Dgv_Excasso2.CellEndEdit
        'MsgBox(Dgv_Excasso2.Items)
        Dim i As Integer = Me.Dgv_Excasso2.CurrentRow.Index
        Dim s As Decimal
        Dim j As Integer
        If Tbx2Dec(Me.Dgv_Excasso2.Rows(i).Cells(2).Value) > Tbx2Dec(Me.Dgv_Excasso2.Rows(i).Cells(3).Value) Then
            j = 2
        Else
            j = 3
        End If

        s = Me.Dgv_Excasso2.Rows(i).Cells(5).Value + Me.Dgv_Excasso2.Rows(i).Cells(4).Value +
        IIf(Me.Dgv_Excasso2.Rows(i).Cells(2).Value > Me.Dgv_Excasso2.Rows(i).Cells(3).Value,
            Me.Dgv_Excasso2.Rows(i).Cells(2).Value, Me.Dgv_Excasso2.Rows(i).Cells(3).Value)

        If s < Me.Dgv_Excasso2.Rows(i).Cells(6).Value Then
            MsgBox("Het uit te keren bedrag mag niet hoger zijn dan de som van de saldi voor contracten, extra giften en fondsen.")
            Me.Dgv_Excasso2.Rows(i).Cells(6).Value = s
        Else
            Me.Dgv_Excasso2.Rows(i).Cells(7).Value = Me.Dgv_Excasso2.Rows(i).Cells(6).Value * Tbx_Excasso_Exchange_rate.Text
        End If

        Calculate_Excasso_Totals()

    End Sub

    Private Sub Tbx_Excasso_CP1_TextChanged(sender As Object, e As EventArgs) Handles Tbx_Excasso_CP1.TextChanged

        Try
            Me.Lbl_Excasso_CP_Totaal.Text = Tbx2Int(GetDouble(Me.Tbx_Excasso_CP1.Text) _
            + GetDouble(Me.Tbx_Excasso_CP2.Text) + GetDouble(Me.Tbx_Excasso_CP3.Text))
        Catch
            MsgBox("Geen geldige invoer.")
        End Try

    End Sub

    Private Sub Tbx_Excasso_CP2_TextChanged(sender As Object, e As EventArgs) Handles Tbx_Excasso_CP2.TextChanged
        'If Not IsNumeric(Tbx_Excasso_CP2) Then Exit Sub
        Me.Lbl_Excasso_CP_Totaal.Text = Tbx2Int(GetDouble(Me.Tbx_Excasso_CP1.Text) _
            + GetDouble(Me.Tbx_Excasso_CP2.Text) + GetDouble(Me.Tbx_Excasso_CP3.Text))
    End Sub
    Private Sub Tbx_Excasso_CP3_TextChanged(sender As Object, e As EventArgs) Handles Tbx_Excasso_CP3.TextChanged
        'If Not IsNumeric(Tbx_Excasso_CP3) Then Exit Sub
        Me.Lbl_Excasso_CP_Totaal.Text = Tbx2Int(GetDouble(Me.Tbx_Excasso_CP1.Text) _
            + GetDouble(Me.Tbx_Excasso_CP2.Text) + GetDouble(Me.Tbx_Excasso_CP3.Text))
    End Sub
    Private Sub Rbn__Excasso_Maandbudget_CheckedChanged(sender As Object, e As EventArgs)
        'If Rbn_Excasso_Maandbudget.Checked And Dgv_Excasso2.Rows.Count > 0 Then
        'Calculate_Excasso_Totals(False)
        'End If
    End Sub
    Private Sub GroupBox5_Leave(sender As Object, e As EventArgs)
        If IsNumeric(Tbx_Excasso_Exchange_rate.Text) Then

        Else
            MsgBox("Ongeldige inhoud")
        End If
    End Sub
    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Btn_Excasso_Exchrate.Click
        For x As Integer = 0 To Me.Dgv_Excasso2.Rows.Count - 1
            Me.Dgv_Excasso2.Rows(x).Cells(7).Value = CInt(Me.Dgv_Excasso2.Rows(x).Cells(6).Value *
                (Me.Tbx_Excasso_Exchange_rate.Text))
        Next

        Btn_Excasso_Exchrate.Enabled = False
        Calculate_Excasso_Totals()

    End Sub
    Private Sub Btn_Excasso_Print_Click(sender As Object, e As EventArgs) Handles Btn_Excasso_Print.Click
        If Cmx_Excasso_Select.SelectedIndex = -1 Then Exit Sub
        Print_Excasso_form()
    End Sub
    Private Sub Btn_Excasso_Delete_Click(sender As Object, e As EventArgs) Handles Btn_Excasso_Delete.Click
        If Cmx_Excasso_Select.SelectedIndex = -1 Then Exit Sub
        If MsgBox("Wilt u de uitkeringslijst verwijderen?", vbYesNo) = vbYes Then
            RunSQL("DELETE FROM journal WHERE name ilike '%" & Me.Cmx_Excasso_Select.SelectedItem & "'", "NULL", "Delete_Excasso_Job")
            'RunSQL("DELETE FROM journal WHERE name='Intern tbv " & Me.Cmx_Excasso_Select.SelectedItem & "'", "NULL", "Delete_Excasso_Job")
            Fill_Cmx_Excasso_Select_Combined()
            Me.Dgv_Excasso2.Columns.Clear()
            Lbl_Excasso_Totaal.Text = ""
            Lbl_Excasso_Items_Totaal.Text = ""
            Lbl_Excasso_Totaal_MDL.Text = ""
            Lbl_Excasso_Tot_Gen_MLD.Text = ""
            Lbl_Excasso_Tot_Gen.Text = ""
            Calculate_CP_Allowance()

        End If

    End Sub

    Private Sub Btn_Excasso_CP_Calculate_Click(sender As Object, e As EventArgs) Handles Btn_Excasso_CP_Calculate.Click
        Calculate_CP_Allowance()
        Btn_Excasso_CP_Calculate.Enabled = False
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
    Private Sub Cmx_Journal_List_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Cmx_Journal_List.SelectedIndexChanged
        Fill_Cmx_Journal_List()
    End Sub

    Private Sub Btn_Select_Bulk_Click(sender As Object, e As EventArgs) Handles Btn_Select_Bulk.Click
        Select_Target_Account()
    End Sub

    Private Sub Lv_Journal_List_Click(sender As Object, e As EventArgs) Handles Lv_Journal_List.Click
        Fill_Journal_List()
    End Sub
    Private Sub Btn_Incasso_Print_Click(sender As Object, e As EventArgs) Handles Btn_Incasso_Print.Click
        Create_SEPA_XML()
    End Sub

    Private Sub Tbx_Journal_Filter_TextChanged(sender As Object, e As EventArgs) Handles Tbx_Journal_Filter.TextChanged
        Fill_Cmx_Journal_List()
    End Sub

    Private Sub Lbl_00_Account__startsaldo_Enter(sender As Object, e As EventArgs)
        Edit_Mode = True
    End Sub

    Private Sub Lbl_00_Account__startsaldo_TextChanged(sender As Object, e As EventArgs)
        If Edit_Mode Then Manage_Buttons_Target(False, False, False, True, True)
    End Sub

    Private Sub Tbx_10_Account__startsaldo_TextChanged(sender As Object, e As EventArgs) Handles Tbx_10_Account__startsaldo.TextChanged
        If Edit_Mode Then Manage_Buttons_Target(False, False, False, True, True)
    End Sub

    Private Sub Tbx_10_Account__startsaldo_Enter(sender As Object, e As EventArgs) Handles Tbx_10_Account__startsaldo.Enter
        Edit_Mode = True
    End Sub

    Private Sub Tbx_10_Account__b_jan_TextChanged(sender As Object, e As EventArgs) Handles _
        Tbx_10_Account__b_jan.TextChanged, Tbx_10_Account__b_feb.TextChanged, Tbx_10_Account__b_mar.TextChanged,
        Tbx_10_Account__b_apr.TextChanged, Tbx_10_Account__b_may.TextChanged, Tbx_10_Account__b_jun.TextChanged,
        Tbx_10_Account__b_jul.TextChanged, Tbx_10_Account__b_aug.TextChanged, Tbx_10_Account__b_sep.TextChanged,
        Tbx_10_Account__b_oct.TextChanged, Tbx_10_Account__b_nov.TextChanged, Tbx_10_Account__b_dec.TextChanged

        If Edit_Mode Then
            Manage_Buttons_Target(False, False, False, True, True)
            Calculate_Manual_Budgets()
        End If
    End Sub

    Private Sub Tbx_10_Account__b_jan_Enter(sender As Object, e As EventArgs) Handles _
            Tbx_10_Account__b_jan.Enter, Tbx_10_Account__b_feb.Enter, Tbx_10_Account__b_mar.Enter,
            Tbx_10_Account__b_apr.Enter, Tbx_10_Account__b_may.Enter, Tbx_10_Account__b_jun.Enter,
            Tbx_10_Account__b_jul.Enter, Tbx_10_Account__b_aug.Enter, Tbx_10_Account__b_sep.Enter,
            Tbx_10_Account__b_oct.Enter, Tbx_10_Account__b_nov.Enter, Tbx_10_Account__b_dec.Enter

        Edit_Mode = True
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
    Private Sub Chbx_Journal_Inactive_CheckedChanged(sender As Object, e As EventArgs) Handles Chbx_Journal_Inactive.CheckedChanged
        Fill_Cmx_Journal_List()
    End Sub

    Private Sub Cbx_Journal_DeSelect_All_Click(sender As Object, e As EventArgs) Handles Cbx_Journal_DeSelect_All.Click
        Cbx_Journal_Select_All.Checked = False
        Select_Deselect_Accounts(False)
        Load_Datagridview(Me.Dgv_Journal_items, "SELECT * FROM journal WHERE name='x!x!x'", "Lv_Journal_List_Click")
    End Sub

    Private Sub Cbx_Journal_Select_All_Click(sender As Object, e As EventArgs) Handles Cbx_Journal_Select_All.Click
        Cbx_Journal_DeSelect_All.Checked = False
        Select_Deselect_Accounts(Cbx_Journal_Select_All.Checked)
        'Fill_Journal_List()

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
        Dim i As Integer = Me.Dgv_Journal_items.CurrentRow.Index
        Dim name As String = Me.Dgv_Journal_items.Rows(i).Cells(1).Value
        If Cmx_Journal_List.Text <> "Journaalnaam" Then
            Exit Sub
        End If
        If InStr(name, "Intern") = 0 Then
            'Clipboard.Clear()
            'Clipboard.SetText(Lv_Journal_Li)
            'MsgBox("'" & Lv_Journal_List.Text & "' gekopieerd naar het klembord.")
            Exit Sub
        End If

        If MsgBox("Weet u zeker dat u deze interne boeking wilt verwijderen?", vbYesNo) = vbYes Then
            RunSQL("DELETE from journal WHERE name='" & name & "'", "NULL", "Lv_Journal_List.DoubleClick")
            Me.Tbx_Journal_Filter.Text = ""
            Fill_Cmx_Journal_List()
        End If
    End Sub

    Private Sub Btn_Account_Budget_Id_Click(sender As Object, e As EventArgs) Handles Btn_Account_Budget_Id.Click
        Calculate_Budget(Lbl_00_pkid.Text)
        Select_Obj2()
    End Sub

    Private Sub Btn_Account_Budget_All_Click(sender As Object, e As EventArgs) Handles Btn_Account_Budget_All.Click
        Calculate_Budget("")
        Select_Obj2()

    End Sub
    Sub Create_Incassolist()
        Dim d As DateTime
        Dim t1 As String
        Dim t2 As String
        Me.Dtp_Incasso_start.Value = CDate("01-" & Me.Dtp_Incasso_start.Value.Month & "-" & Me.Dtp_Incasso_start.Value.Year)
        Me.Dtp_Incasso_start.MaxDate = New Date(Today.Year, Today.Month + 1, 1)
        d = Me.Dtp_Incasso_start.Value.AddMonths(1)
        Me.Dtp_Incasso_end.Value = New DateTime(d.Year, d.Month, 1).AddDays(-1)
        Me.Dtp_Incasso_start.MinDate = New Date(Today.Year, 1, 1)
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
        Else
            Load_Datagridview(Me.Dgv_Incasso, Create_Incasso_Bookings(t1), "Me.Dtp_Incasso_start.ValueChanged")

        End If
        Load_Listview(Me.Lv_Incasso_Overview, Create_Incasso_Totals(t1))
        Me.Lv_Incasso_Overview.Columns(0).Text = "Type"
        Me.Lv_Incasso_Overview.Columns(1).Text = "Aantal"
        Me.Lv_Incasso_Overview.Columns(2).Text = "Bedrag"


        Me.Lv_Incasso_Overview.Items.Add("Totaal")
        Dim Tot_nr As Integer = CInt(Me.Lv_Incasso_Overview.Items(0).SubItems(1).Text) +
                                        CInt(Me.Lv_Incasso_Overview.Items(1).SubItems(1).Text)
        Dim Tot_amt = Format(CDec(Me.Lv_Incasso_Overview.Items(0).SubItems(2).Text) +
                                        CDec(Me.Lv_Incasso_Overview.Items(1).SubItems(2).Text), "€#.##")
        Me.Lv_Incasso_Overview.Items(2).SubItems.Add(Tot_nr)
        Me.Lv_Incasso_Overview.Items(2).SubItems.Add(Tot_amt)

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
            Dim Checksum = QuerySQL("Select Sum(amt1) from journal where name ='" & journal_name & "'")
            If Tot_amt <> Checksum Then
                Dim msg = "Er is een fout opgetreden: het berekende totaal (" & Tot_amt &
                    ") komt niet overeen met de eerder gecreëerde incassojob (" &
                    Checksum & "). Als deze incassojob nog niet naar de bank is verstuurd" &
                    " wordt u geadviseerd  deze job te verwijderen en een nieuwe aan te maken " &
                    "voor deze maand."
                Me.Lbl_Incasso_Error.Text = msg
                Me.Lbl_Incasso_Error.Visible = True
            End If
        ElseIf qtverwerkt > 0 Then
            Me.Lbl_Incasso_Status.Text = "Verwerkt"
            Me.Btn_Incasso_Delete.Enabled = False
            Me.Btn_Run_Incasso.Enabled = False
            Me.Btn_Incasso_Print.Enabled = True
            Dim Checksum = QuerySQL("SELECT Sum(amt1) from journal where name ='" & journal_name & "'")
            If Tot_amt <> Checksum Then
                Me.Lbl_Incasso_Error.Text = "Opgeslagen incassojob is niet in lijn met contractdata"
            End If
        Else
            Me.Lbl_Incasso_Status.Text = "Nieuw"
            Me.Btn_Incasso_Delete.Enabled = False
            Me.Btn_Run_Incasso.Enabled = True
            Me.Btn_Incasso_Print.Enabled = False
        End If

    End Sub
    Private Sub Cmx_Excasso_Select_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Cmx_Excasso_Select.SelectedIndexChanged

        Load_Excasso_Form()
    End Sub

    Sub Load_Excasso_Form()
        If Cmx_Excasso_Select.SelectedIndex = -1 Then Exit Sub

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
            Dtp_Excasso_Start.Enabled = True

            Load_Excasso_Balances()
            Lbl_Excasso_CP_Totaal.Text = 0
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
            Lbl_Excasso_CPid.Text = str1(0)
            'calculate actual exchange rate
            Tbx_Excasso_Exchange_rate.Text =
                            GetDouble(QuerySQL("
                                    SELECT sum(amt2)/sum(amt1) FROM journal
                                    WHERE name ='" & Cmx_Excasso_Select.SelectedItem & "'"))
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
            Else
                Lbl_Excasso_CP_Totaal.Text = 0
            End If

            ' Lbl_Excasso_CP_Totaal.Text =
            'GetDouble(QuerySQL("
            'Select Case sum(amt1) FROM journal
            'WHERE name ='" & Cmx_Excasso_Select.SelectedItem & "' 
            '                  AND type='CP'
            '                 AND amt1<='0.00'") * -1)
            'retrieve the data from journal posts
            Load_Excasso_Balances()
            Collect_data(Existing_Excasso(Cmx_Excasso_Select.SelectedItem))
            Dim c


            For i = 0 To Me.Dgv_Excasso2.Rows.Count - 1  'go through all posts
                For j = 0 To dst.Tables(0).Rows.Count - 1
                    'MsgBox(Me.Dgv_Excasso2.Rows(i).Cells(0).Value & " - " & dst.Tables(0).Rows(j)(0))
                    If Me.Dgv_Excasso2.Rows(i).Cells(0).Value = dst.Tables(0).Rows(j)(0) Then 'same account id

                        Me.Dgv_Excasso2.Rows(i).Cells(6).Value = Me.Dgv_Excasso2.Rows(i).Cells(6).Value +
                                                                    dst.Tables(0).Rows(j)(5)
                        Me.Dgv_Excasso2.Rows(i).Cells(7).Value = Me.Dgv_Excasso2.Rows(i).Cells(7).Value +
                                                                    dst.Tables(0).Rows(j)(6)

                        If Strings.Trim(dst.Tables(0).Rows(j)(7)) = "Contract" Then
                            c = IIf(IsDBNull(dst.Tables(0).Rows(j)(2)), 0, dst.Tables(0).Rows(j)(2))
                            Me.Dgv_Excasso2.Rows(i).Cells(3).Value =
                            Me.Dgv_Excasso2.Rows(i).Cells(3).Value + c

                        End If

                        If Strings.Trim(dst.Tables(0).Rows(j)(7)) = "Extra" Then
                            c = IIf(IsDBNull(dst.Tables(0).Rows(j)(3)), 0, dst.Tables(0).Rows(j)(3))
                            Me.Dgv_Excasso2.Rows(i).Cells(4).Value =
                            Me.Dgv_Excasso2.Rows(i).Cells(4).Value + c
                        End If
                        If Strings.Trim(dst.Tables(0).Rows(j)(7)) = "Internal" Then
                            c = IIf(IsDBNull(dst.Tables(0).Rows(j)(4)), 0, dst.Tables(0).Rows(j)(4))
                            Me.Dgv_Excasso2.Rows(i).Cells(5).Value =
                            Me.Dgv_Excasso2.Rows(i).Cells(5).Value + c
                        End If
                        'Me.Dgv_Excasso2.Rows(i).Cells(7).Value = 100

                    End If
                Next j

            Next i

        End If
        Calculate_Excasso_Totals()
        'Tbx_Excasso_CP1.Text = "0"
        'Tbx_Excasso_CP2.Text = "0"
        'Tbx_Excasso_CP2.Text = "0"
        'Lbl_Excasso_CP_Totaal.Text = "0"
    End Sub
    Sub Get_New_Excasso_Data()


        If Me.Cmx_Excasso_Select.SelectedIndex = -1 Then Exit Sub

        Dim t1 = IIf(Me.Cbx_Uitkering_Kind.Checked, "Kind", "--")
        Dim t2 = IIf(Me.Cbx_Uitkering_Oudere.Checked, "Oudere", "--")
        Dim t3 = IIf(Me.Cbx_Uitkering_Overig.Checked, "Overig", "--")
        Dim m As Integer = Month(Me.Dtp_Excasso_Start.Value)


        Dim pos1 As Integer = Strings.InStr(Me.Cmx_Excasso_Select.SelectedItem, "[")
        Dim CP As String = Strings.Mid(Me.Cmx_Excasso_Select.SelectedItem, pos1 + 1,
        Len(Me.Cmx_Excasso_Select.SelectedItem) - pos1 - 1)
        'MsgBox(CP)

        Dim d1 = Year(Me.Dtp_Excasso_Start.Value) & "-" & Month(Me.Dtp_Excasso_Start.Value) & "-" &
             Me.Dtp_Excasso_Start.Value.Day

        Dim s As String = Create_Excasso(CP, t1, t2, t3, d1, m.ToString)
        If s = "" Then Exit Sub
        'MsgBox(s)

        'Dim d1 = Format(Me.Dtp_Excasso_Start.Value, "dd-MM-yyyy")
        Load_Datagridview(Me.Dgv_Excasso2, s, "Get_New_Excasso_Data")

        Format_dvg_excasso()
        Convert_Null_to_0()
        Dim arg As Integer = Me.Lbl_Excasso_LastCalc.Text
        If arg = 2 Or arg = 3 Then Calculate_Excasso_Amounts(Tbx2Int(Me.Lbl_Excasso_LastCalc.Text))
        'Calculate_Excasso_Totals(False)
        Calculate_Excasso_Totals()
        Tbx_Excasso_CP1.Text = "0"
        Tbx_Excasso_CP2.Text = "0"
        Tbx_Excasso_CP2.Text = "0"
        Lbl_Excasso_CP_Totaal.Text = "0"


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

    Private Sub Btn_Excasso_Nullvalues_Click(sender As Object, e As EventArgs) Handles Btn_Excasso_Nullvalues.Click
        Set_Excasso_Nullvalues()
        Lbl_Excasso_LastCalc.Text = 0
        Calculate_Excasso_Totals()
        Calculate_CP_Allowance()
    End Sub

    Private Sub Btn_Excasso_Maandbudget_Click(sender As Object, e As EventArgs) Handles Btn_Excasso_Maandbudget.Click
        Calculate_Excasso_Amounts(2)
        Lbl_Excasso_LastCalc.Text = 2
        Calculate_Excasso_Totals()
        Calculate_CP_Allowance()

    End Sub

    Private Sub Btn_Excasso_Act_Saldo_Click(sender As Object, e As EventArgs) Handles Btn_Excasso_Act_Saldo.Click
        Calculate_Excasso_Amounts(3)
        Lbl_Excasso_LastCalc.Text = 3
        Calculate_Excasso_Totals()
        Calculate_CP_Allowance()
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
        Me.Lbl_Incasso_Error.Visible = False

    End Sub

    Private Sub Btn_Bank_Folder_Click(sender As Object, e As EventArgs) Handles Btn_Bank_Folder.Click
        Load_Bank_csv_from_folder()
    End Sub

    Private Sub Dgv_Bank_ColumnHeaderMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles Dgv_Bank.ColumnHeaderMouseClick
        Format_dvg_bank()
    End Sub

    Private Sub Dgv_Bank_RowHeaderMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles Dgv_Bank.RowHeaderMouseClick
        Format_dvg_bank()
    End Sub

    Private Sub Btn_Bank_Categorize_Click(sender As Object, e As EventArgs) Handles Btn_Bank_Categorize.Click
        Categorize_Bank_Transactions()
        Fill_bank_transactions()
    End Sub

    Private Sub Tbx_Bank_Description_Leave(sender As Object, e As EventArgs) Handles Tbx_Bank_Description.Leave
        Dim SQLstr = "UPDATE bank SET description='" & Tbx_Bank_Description.Text &
               "' WHERE id='" & Dgv_Bank.SelectedCells(0).Value & "'"
        RunSQL(SQLstr, "NULL", "Tbx_Bank_Description.Leave")
        Dgv_Bank.SelectedCells(3).Value = Tbx_Bank_Description.Text
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

    Private Sub Cbx_Uitkering_Kind_CheckedChanged(sender As Object, e As EventArgs) Handles Cbx_Uitkering_Kind.CheckedChanged

    End Sub

    Private Sub Cbx_00_cp__active_Click(sender As Object, e As EventArgs) Handles Cbx_00_cp__active.Click
        CheckActive(Cbx_00_cp__active, Lbl_CP_pkid, "target")
    End Sub

    Sub Save_Banktransaction_Accounts()
        'remove old transactions -> later to be implemented
        If Me.Dgv_Bank.Rows.Count = 0 Or Dgv_Bank_Account.Rows.Count = 0 Then Exit Sub


        Dim bid As Integer = Me.Dgv_Bank.SelectedCells(0).Value
        Dim _dat As Date = Me.Dgv_Bank.SelectedCells(1).Value
        Dim dat As String = _dat.Year & "-" & _dat.Month & "-" & _dat.Day
        Dim des As String = Me.Dgv_Bank.SelectedCells(3).Value
        Dim Amt_In = CDec(Me.Dgv_Bank.SelectedCells(4).Value)
        Dim Amt_Out = CDec(Me.Dgv_Bank.SelectedCells(5).Value)
        Dim nam As String
        Dim typ As String = ""
        If Rbn_Bank_Contract.Checked Then
            typ = "Contract"
        ElseIf Rbn_Bank_Extra.Checked Then
            typ = "Extra"
        ElseIf Rbn_Bank_Other.Checked Then
            typ = "Anders"
        ElseIf Amt_In > 0 And Strings.Left(des, 8) <> "Contract" Then
            MsgBox("Geef s.v.p. aan of dit een contractbetaling, extra gift of andere betaling betreft.")
            Exit Sub
        End If

        Dim SQLstr = "DELETE FROM journal WHERE fk_bank=" & bid & ";" &
                     "INSERT INTO journal(date,status,amt1,description,source, fk_account,fk_bank,name,type) VALUES "

        For x As Integer = 0 To dst.Tables(0).Rows.Count - 1
            If dst.Tables(0).Rows(x)(2) <> 0 Then
                SQLstr &= "('" & dat & "','Verwerkt','" & Cur2(dst.Tables(0).Rows(x)(2)) & "','" &
                  des & "','Bank'," & dst.Tables(0).Rows(x)(0) & "," & bid & ",'" & typ & "','" & typ & "'),"
            End If
        Next

        SQLstr = Strings.Left(SQLstr, Strings.Len(SQLstr) - 1) 'remove the last comma
        If Me.Chbx_test.Checked Then MsgBox(SQLstr)
        RunSQL(SQLstr, "NULL", "")


    End Sub
    Sub Calculate_Total_Booked()


        Dim Amt_In = CDec(Me.Dgv_Bank.SelectedCells(4).Value)
        Dim Amt_Out = CDec(Me.Dgv_Bank.SelectedCells(5).Value)
        Dim total As Decimal = 0
        Dim nill As Integer = -1
        Dim or_amt = Amt_In - Amt_Out

        If dst.Tables("Table").Rows.Count = 0 Then
            'SPAS.Tbx_Bank_Accounts_Total.Text = 0

        Else
            Dim amt As Decimal
            For x As Integer = 0 To dst.Tables(0).Rows.Count - 1
                If dst.Tables(0).Rows(x)(0) = nocat Then
                    nill = x
                    'amt2 = dst.Tables(0).Rows(x)(2)
                Else
                    amt = CDec(dst.Tables(0).Rows(x)(2))
                    total = total + amt
                End If
            Next
            Dim diff = or_amt - total
            If nill = -1 Then

                If diff <> 0 Then  'account 'uncategorized not present
                    Dim R As DataRow
                    R = dst.Tables(0).Rows.Add
                    R(0) = nocat
                    R(1) = QuerySQL("SELECT name FROM account WHERE id='" & nocat & "'")
                    R(2) = diff
                End If
            Else
                dst.Tables(0).Rows(nill)(2) = or_amt - total
            End If
            Me.Tbx_Bank_Amount.Text = diff

            'SPAS.Tbx_Bank_Accounts_Total.Text = total
        End If


    End Sub

    Sub Format_dvg_reports()

        With Me.Dgv_Rapportage

            .Columns(1).HeaderText = "Account"
            .Columns(1).HeaderText = "Account detail"
            .Columns(2).HeaderText = "Startsaldo"
            .Columns(3).HeaderText = "Bij"
            .Columns(4).HeaderText = "Af"
            .Columns(5).HeaderText = "Saldo"
            .Columns(0).Width = 230
            .Columns(1).Width = 200

            For k = 2 To 5
                .Columns(k).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                .Columns(k).Width = 100
                .Columns(k).DefaultCellStyle.Format = "N2"
                .Columns(k).ReadOnly = False

            Next


            '.Rows(1).DefaultCellStyle.Font


            For r = 0 To .Rows.Count - 1
                If IsDBNull(.Rows(r).Cells(2).Value) Then
                    '.Rows(r).DefaultCellStyle.ForeColor = Color.DarkGreen
                    .Rows(r).DefaultCellStyle.Font = New Font("Arial", 12, FontStyle.Bold)
                ElseIf InStr(UCase(.Rows(r).Cells(0).Value), "TOTAAL") > 0 Then
                    .Rows(r).DefaultCellStyle.ForeColor = Color.Blue
                    .Rows(r).DefaultCellStyle.Font = New Font("Calibri", 12, FontStyle.Bold)
                Else
                    If Not IsDBNull(.Rows(r).Cells(4).Value) Then
                        If .Rows(r).Cells(4).Value = 0 Then
                            .Rows(r).Cells(4).Style.ForeColor = Color.Gray            '.Font = New Font("Arial", 12, FontStyle.Bold)
                        End If
                    End If
                End If


            Next



            '.Columns(0).Visible = False

        End With

    End Sub

    Sub Format_dvg_bank()

        With Me.Dgv_Bank
            .Columns(1).HeaderText = "Datum"
            .Columns(2).HeaderText = "Name"
            .Columns(3).HeaderText = "Omschrijving"
            .Columns(4).HeaderText = "Bij"
            .Columns(5).HeaderText = "Af"

            .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            .Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

            .Columns(1).Width = 80
            .Columns(2).Width = 170
            .Columns(3).Width = 340
            .Columns(4).Width = 75
            .Columns(5).Width = 75

            .Columns(0).Visible = False

        End With

    End Sub

    Sub Format_dvg_bank_journal()

        If Me.Dgv_Bank.Rows.Count = 0 Then Exit Sub  'de vraag is of dit correct is
        Try
            With Me.Dgv_Bank_Account

                .Columns(0).HeaderText = "Id"
                .Columns(1).HeaderText = "Account"
                .Columns(2).HeaderText = "Bedrag"

                .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                .Columns(2).DefaultCellStyle.Format = "N2"
                .Columns(2).DefaultCellStyle.ForeColor = Color.Blue

                .Columns(0).Visible = False
                .Columns(1).Width = 190
                .Columns(2).Width = 70


                .Columns(0).ReadOnly = True
                .Columns(1).ReadOnly = True
                .Columns(2).ReadOnly = False

            End With
        Catch
        End Try

    End Sub
    Sub Fill_bank_transactions()


        If Strings.InStr(Me.Cmx_Bank_bankacc.Text, "NL") = 0 Then Exit Sub
        'MsgBox(Strings.Right(SPAS.Cmx_Bank_bankacc.Text, 18))
        Dim bankacc = Strings.Right(Me.Cmx_Bank_bankacc.Text, 18)
        Dim sv As String = Me.Tbx_Bank_Search.Text
        Dim filter As String
        If Not IsDBNull(sv) Then
            filter = "AND description ILIKE '%" & sv & "%'"
        Else
            filter = ""
        End If

        Dim SQLstr = "SELECT id, date, name, description As descr, 
                      credit, debit,code, exch_rate, iban2, seqorder,
                      batchid, amt_cur, fk_journal_name,filename,cost,iban
                      FROM bank WHERE iban ='" & bankacc & "' " & filter & " 
                      ORDER BY seqorder DESC, date DESC"

        If Me.Chbx_test.Checked Then MsgBox(SQLstr)

        Load_Datagridview(Me.Dgv_Bank, SQLstr, "fill bank transactions")
        Format_dvg_bank()
        Mark_rows_Dgv_Bank()
        Calculate_Bank_Balance()

    End Sub

    Private Sub NieuwToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NieuwToolStripMenuItem.Click
        Login.Text = "Inloggen in productieomgeving"
        Login.Cmx_Login_Database.Text = "Productie"
        Login.Show()

    End Sub

    Private Sub OpenToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenToolStripMenuItem.Click
        Login.Text = "Inloggen in testomgeving"
        Login.Cmx_Login_Database.Text = "Acceptatie"
        Login.Show()
    End Sub

    Private Sub Btn_Basis_Delete_Click_1(sender As Object, e As EventArgs) Handles Btn_Basis_Delete.Click
        Dim id As Integer
        Dim sqlstr As String = ""

        If Me.Dtp_31_contract__enddate.Value <= Date.Today Then
            MsgBox("Alleen contracten die nog niet zijn ingegaan kunnen verwijderd worden.")
            Exit Sub
        Else
            If MsgBox("Weet u zeker dat u dit contract wilt verwijderen?", vbYesNo) = vbNo Then
                Exit Sub
            End If
        End If
        If Lbx_Basis.SelectedIndex <> -1 Then id = Lbx_Basis.SelectedItem(Me.Lbx_Basis.ValueMember)

        Select Case TC_Object.SelectedIndex
            Case 0
                sqlstr = "DELETE FROM contract WHERE id=" & id
            Case Else
                MsgBox("Deze functie is nu nog alleen voor contracten gedefnieerd")
        End Select

        If sqlstr <> "" Then
            RunSQL(sqlstr, "NULL", "Btn_Basis_Delete_Click")
            Load_Table()
            MsgBox("Het contract is verwijderd.")
        End If

    End Sub

    Private Sub Btn_Excasso_Calculate_Exchrate_Click(sender As Object, e As EventArgs) Handles Btn_Excasso_Calculate_Exchrate.Click
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

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Cmx_Settings_Overhead.SelectedIndexChanged

    End Sub

    Private Sub TestToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TestToolStripMenuItem.Click
        Login.Text = "Inloggen in productieomgeving"
        Login.Cmx_Login_Database.Text = "Test"
        Login.Show()
    End Sub

    Private Sub Btn_Excasso_Cancel_Click(sender As Object, e As EventArgs) Handles Btn_Excasso_Cancel.Click
        If Cmx_Excasso_Select.SelectedIndex = -1 Then Exit Sub
    End Sub

    Private Sub Btn_Beheer_YearClose_Click(sender As Object, e As EventArgs) Handles Btn_Beheer_YearClose.Click
        Manage_StartSaldo()
    End Sub


    Private Sub Btn_Bank_Split_Click(sender As Object, e As EventArgs) Handles Btn_Bank_Split.Click



        Banksplit.Show()

    End Sub

    Private Sub Btn_Account_Startsaldi_All_Click(sender As Object, e As EventArgs) Handles Btn_Account_Startsaldi_All.Click

        Dim bookyear As String = InputBox("Voor boekjaar in (bijv. 2021)...")
        Dim SQLstr As String =
         "
        DELETE FROM public.journal WHERE description Like '[Startsaldo %' and description like '%]';
        INSERT INTO Public.journal (name, Date, status, Amt1, description, source, fk_account, Type) 
        SELECT 'Startsaldo', '" & bookyear & "-01-01','Verwerkt', startsaldo,'Startsaldo '||name, 'Intern', id, 'Extra'  from account where startsaldo is distinct from NULL;
          "
        Clipboard.Clear()
        Clipboard.SetText(SQLstr)
        RunSQL(SQLstr, "NULL", "Btn_Account_Startsaldi_All.Click")
    End Sub

    Private Sub Btn_Bankacc_UpdateStartsaldi_Click(sender As Object, e As EventArgs) Handles Btn_Bankacc_UpdateStartsaldi.Click
        Dim bookyear As String = InputBox("Voor boekjaar in (bijv. 2021)...")
        If Not IsNumeric(bookyear) Or Len(bookyear) <> 4 Or bookyear < "2020" Then
            MsgBox("Ongeldige invoer")
            Exit Sub
        End If

        Dim SQLstr As String =
         "
        DELETE from bank WHERE name = 'Startsaldo';
        INSERT INTO bank(iban,currency,date,debit,credit,name,description,fk_journal_name)
        SELECT accountno, currency, '2021-01-01', 
        CASE WHEN startbalance::numeric < 0.00 Then startbalance else '0' END,
        CASE WHEN startbalance::numeric > 0.00  Then startbalance ELSE '0' END,
        'Startsaldo' As name,'Startsaldo '||accountno, 'Startsaldo'
        FROM public.bankacc
        WHERE startbalance is distinct from NULL;
"
        RunSQL(SQLstr, "NULL", "Btn_Bankacc_UpdateStartsaldi.Click")
    End Sub

    Private Sub Cmx_00_Contract__fk_account_id_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Cmx_00_Contract__fk_account_id.SelectedIndexChanged

    End Sub

    Private Sub ToolTip1_Popup(sender As Object, e As PopupEventArgs) Handles ToolTip1.Popup

        ToolTip1.SetToolTip(Btn_Bank_Categorize, "Categoriseer transacties")
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

    Private Sub Btn_Settings_Cancel_Click(sender As Object, e As EventArgs) Handles Btn_Settings_Cancel.Click
        Get_Settings_Data()
    End Sub

    Private Sub Btn_Settings_Save_Click(sender As Object, e As EventArgs) Handles Btn_Settings_Save.Click
        Dim sqlstr As String
        sqlstr = "
        Update public.settings SET value = " & Tbx2Int(Tbx_Settings_Overhead_Oudere.Text) & " WHERE label = 'standaard_overhead_oudere';
        Update public.settings SET value = " & Tbx2Int(Tbx_Settings_Overhead_Kind.Text) & " WHERE label = 'standaard_overhead_kind';
        Update public.settings SET value = " & Tbx2Int(Tbx_Settings_Bedrag_Oudere.Text) & " WHERE label = 'standaard_bedrag_oudere';
        Update public.settings SET value = " & Tbx2Int(Tbx_Settings_Bedrag_Kind.Text) & " WHERE label = 'standaard_bedrag_kind';
        Update public.settings SET value = '" & Tbx_Settings_Banktext_Kind.Text & "' WHERE label = 'bank_kind';
        Update public.settings SET value = '" & Tbx_Settings_Banktext_Oudere.Text & "' WHERE label = 'bank_oudere';
"
        RunSQL(sqlstr, "NULL", "Btn_Settings_Save")
        Clipboard.Clear()
        Clipboard.SetText(sqlstr)

    End Sub

    Private Sub Dgv_Bank_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles Dgv_Bank.CellContentClick

    End Sub

    Private Sub Dgv_Excasso2_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles Dgv_Excasso2.CellContentClick
        ''Tbx_Journal_Filter.Text = Dgv_Excasso2.CurrentCell.Value
    End Sub

    Private Sub Dgv_Excasso2_DoubleClick(sender As Object, e As EventArgs) Handles Dgv_Excasso2.DoubleClick


        If Dgv_Excasso2.CurrentCell.ColumnIndex <> 1 Then Exit Sub
        'Clipboard.SetText(Dgv_Excasso2.CurrentCell.Value)
        TC_Main.SelectedIndex = 4
        Tbx_Journal_Filter.Text = Dgv_Excasso2.CurrentCell.Value
        If Lv_Journal_List.Items.Count > 0 Then
            Lv_Journal_List.Items(0).Focused = True
            Lv_Journal_List.Items(0).Selected = True
            Lv_Journal_List.Items(0).Focused = True

        End If

    End Sub

    Private Sub Dgv_Excasso2_CellEnter(sender As Object, e As DataGridViewCellEventArgs) Handles Dgv_Excasso2.CellEnter


    End Sub

    Private Sub Dgv_Excasso2_Click(sender As Object, e As EventArgs) Handles Dgv_Excasso2.Click
        Clipboard.Clear()
        Clipboard.SetText(Dgv_Excasso2.CurrentCell.Value)
    End Sub

    Private Sub Dgv_Rapportage_DoubleClick(sender As Object, e As EventArgs) Handles Dgv_Rapportage.DoubleClick
        Clipboard.SetText(Dgv_Rapportage.CurrentRow.Cells(1).Value)
        TC_Main.SelectedIndex = 4
        Cmx_Journal_List.Text = "Accountgroep"
        Tbx_Journal_Filter.Text = Dgv_Rapportage.CurrentRow.Cells(1).Value
    End Sub

    Private Sub Btn_Rapportage_Ververs_Click(sender As Object, e As EventArgs) Handles Btn_Rapportage_Ververs.Click
        Load_Datagridview(Dgv_Rapportage, "select * from public.reports", "rapportagefout")
        Format_dvg_reports()
    End Sub

    Private Sub Btn_Boeking_Intern_Click(sender As Object, e As EventArgs) Handles Btn_Journal_Intern.Click
        Toggle_Internal_Bookings()

    End Sub

    Sub Toggle_Internal_Bookings()
        Dim vis = Gbx_Boekingen_Overboeking.Visible

        Gbx_Boekingen_Overboeking.Visible = IIf(vis, False, True)
        Lbl_Boeking_Selecteer.Left = IIf(vis, 10, +272)
        Lbl_Journal_Filter.Left = IIf(vis, 10, 272)
        Cmx_Journal_List.Left = IIf(vis, 82, 354)
        Tbx_Journal_Filter.Left = IIf(vis, 82, 354)
        Lv_Journal_List.Left = Lv_Journal_List.Left - -IIf(vis, -262, +262)
        Lv_Journal_List.Width = IIf(vis, 370, 225)
        Cbx_Journal_Select_All.Left = IIf(vis, 10, 272)
        Cbx_Journal_DeSelect_All.Left = IIf(vis, 100, 340)
        Chbx_Journal_Inactive.Left = IIf(vis, 180, 420)
        Gbx_Journal_Totals.Left = IIf(vis, 404, 536)
        Dgv_Journal_items.Left = IIf(vis, 404, 536)
        Dgv_Journal_items.Width = IIf(vis, 700, 571)
        Dgv_Journal_items.Columns(5).Width = Dgv_Journal_items.Columns(5).Width + IIf(vis, 100, -100)
        Dgv_Journal_items.Columns(6).Width = IIf(vis, 200, 110)
        Lbl_Journal_Status.Left = IIf(vis, 405, 535)
        Cbx_Journal_Status_Open.Left = IIf(vis, 460, 590)
        Cbx_Journal_Status_Verwerkt.Left = IIf(vis, 530, 660)
    End Sub

    Private Sub Btn_Excasso_FormRefresh_Click(sender As Object, e As EventArgs) Handles Btn_Excasso_FormRefresh.Click
        Load_Excasso_Form()
    End Sub
End Class

