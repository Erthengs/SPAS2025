Imports System.ComponentModel
Imports System.ComponentModel.DataAnnotations
Imports System.Data.Entity
Imports System.Data.Entity.Core.Common.CommandTrees.ExpressionBuilder
Imports System.Data.Entity.Core.Common.EntitySql
Imports System.Data.Entity.Migrations
Imports System.Diagnostics
Imports System.IO
Imports System.Linq.Expressions
Imports System.Management.Instrumentation
'Imports System.Refection
Imports System.Reflection
Imports System.Security.Cryptography

Imports System.Windows.Forms.VisualStyles
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.ProgressBar
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Rebar
Imports System.Xml
Imports Microsoft.EntityFrameworkCore.Metadata.Internal
Imports Microsoft.EntityFrameworkCore.Query.Internal
Imports Microsoft.EntityFrameworkCore.Query.SqlExpressions
Imports Microsoft.VisualBasic.Devices
Imports Npgsql
Imports NpgsqlTypes
Imports OpenTK.Graphics.OpenGL
Imports OpenTK.Platform
Imports PdfSharp.Pdf.Content


Public Class SPAS
    Private Const V As Boolean = False
    Private oldend_date As Date
    Private current_tabpage As Integer
    Private current_tabpage_basis As Integer
    Public isManualChange As Boolean = False
    Private isProgrammaticChange As Boolean = True
    Private originalValues As New Dictionary(Of String, Object)
    Private editingControl As System.Windows.Forms.TextBox
    Private originalValue As Object ' Stores the original cell value
    Private isCanceling As Boolean = False
    Public MustWarn As Boolean = False


    'bekende fouten

    '========================================================================================================
    '======                                                                                            ======
    '======                                 nieuwe voor tussenrekening                                 ======
    '======                                                                                            ======
    '========================================================================================================

    Private Sub Dgv_Tussenrekening_Uitk_SelectionChanged(sender As Object, e As EventArgs)
        ' Check if a valid row is currently selected
        If Dgv_Tussenrekening_Uitk.CurrentRow IsNot Nothing AndAlso
       Not Dgv_Tussenrekening_Uitk.CurrentRow.IsNewRow AndAlso
       Dgv_Tussenrekening_Uitk.CurrentRow.Cells(0).Value IsNot Nothing Then

            Btn_Tussenrekening_Save.Enabled = True
        Else
            Btn_Tussenrekening_Save.Enabled = False
        End If
    End Sub

    Private Sub Btn_Tussenrekening_Save_Click(sender As Object, e As EventArgs) Handles Btn_Tussenrekening_Save.Click
        Dim clearanceAccountId As String = "750"

        ' Traffic Cop Check: Is the manual amount textbox filled?
        If Not String.IsNullOrWhiteSpace(Tbx_Tussenrekening_bedrag.Text) Then
            ' ==========================================
            ' VALIDATE & PROCESS MANUAL ENTRY
            ' ==========================================
            If Cmbx_Tussenrekening.SelectedItem Is Nothing Then
                MsgBox("Selecteer een doelaccount uit de lijst.")
                Exit Sub
            End If

            Dim amount As Decimal
            If Not Decimal.TryParse(Tbx_Tussenrekening_bedrag.Text, amount) Then
                MsgBox("Voer een geldig numeriek bedrag in (gebruik een minteken voor uitgaven).")
                Exit Sub
            End If

            Dim selectedAccItem As ComboBoxItem = TryCast(Cmbx_Tussenrekening.SelectedItem, ComboBoxItem)
            Dim targetAccountId As String = selectedAccItem.Column1

            Dim transDate As Date = Dtp_tussenrekening.Value
            Dim desc As String = Tbx_Tussenrekening_toelichting.Text
            selectedAccItem = TryCast(Cmbx_Tussenrekening.SelectedItem, ComboBoxItem)
            targetAccountId = selectedAccItem.Column1
            transDate = Dtp_tussenrekening.Value
            desc = Tbx_Tussenrekening_toelichting.Text

            Dim actionText As String = IIf(amount > 0, "afboeken als uitgave", "boeken als teruggave/inkomst")
            Dim conf = MsgBox($"Weet u zeker dat u €{Math.Abs(amount)} wilt {actionText} op de geselecteerde account?", MsgBoxStyle.YesNo Or MsgBoxStyle.Question, "Bevestig Afletteren")

            If conf = MsgBoxResult.Yes Then
                ' Call procedure with manual parameters
                Net_Distribution_List(transDate, clearanceAccountId, "", amount, targetAccountId, desc)

                ' Reset UI fields after success
                Tbx_Tussenrekening_bedrag.Clear()
                Tbx_Tussenrekening_toelichting.Clear()
                Cmbx_Tussenrekening.SelectedIndex = -1
            End If

        Else
            ' ==========================================
            ' VALIDATE & PROCESS DISTRIBUTION LIST
            ' ==========================================
            If Dgv_Tussenrekening_Uitk.CurrentRow Is Nothing Then
                MsgBox("Selecteer een uitkeringslijst in de tabel of vul een handmatig bedrag in om af te letteren.")
                Exit Sub
            End If

            ' Extract from DataGridView
            Dim listDate As Date = CDate(Dgv_Tussenrekening_Uitk.CurrentRow.Cells(0).Value)
            Dim listName As String = Dgv_Tussenrekening_Uitk.CurrentRow.Cells(1).Value.ToString()

            Dim conf = MsgBox($"Weet u zeker dat u uitkeringslijst '{listName}' van {listDate.ToShortDateString()} wilt afletteren?", MsgBoxStyle.YesNo Or MsgBoxStyle.Question, "Bevestig Afletteren")
            If conf = MsgBoxResult.Yes Then

                ' Call procedure with list parameters
                Net_Distribution_List(listDate, clearanceAccountId, listName)
                Fill_Cmx_Excasso_Select_Combined()
            End If
        End If

        ' Refresh calls to reload the DataGridViews
        Prepare_Datagridview(Dgv_Tussenrekening, Fill_Afletterbox, {"TZ080", "TZ300", "NZ080", "NZ080", "NZ080"})

        Prepare_Datagridview(Dgv_Tussenrekening_Uitk,
                 "SELECT date, name, SUM(amt1) FROM journal WHERE source = 'Uitkering' AND status = 'Open' GROUP BY date, name order by date asc",
                 {"HZ000", "TZ200", "NZ080"})

        Lbl_Tussenrekening_3.Text = $"Openstaande uitkeringslijsten ({Dgv_Tussenrekening_Uitk.RowCount})"
        Initialize_Tussenrekening_DatePicker()
    End Sub




    Private Sub Dgv_Tussenrekening_Uitk_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles Dgv_Tussenrekening_Uitk.CellContentClick
        Btn_Tussenrekening_Save.Enabled = True
    End Sub

    Sub Initialize_Tussenrekening_DatePicker()
        ' 1. Fetch the most recent date from the journal
        Dim sqlMaxDate As String = "SELECT MAX(""date"") FROM public.journal;"
        Dim maxDateStr As String = QuerySQL(sqlMaxDate)

        Dim minAllowedDate As Date

        ' 2. Parse the result and apply the limits safely
        If Date.TryParse(maxDateStr, minAllowedDate) Then
            ' Set the lowest possible date the user can select
            Dtp_tussenrekening.MinDate = minAllowedDate
        Else
            ' Fallback if the journal is completely empty
            Dtp_tussenrekening.MinDate = DateTime.Today
        End If

        ' 3. Set the default value to Today. 
        ' Safety check: If the database somehow contains a future-dated record, 
        ' setting the Value to Today would throw an exception because Today < MinDate.
        If DateTime.Today >= Dtp_tussenrekening.MinDate Then
            Dtp_tussenrekening.Value = DateTime.Today
        Else
            Dtp_tussenrekening.Value = Dtp_tussenrekening.MinDate
        End If
    End Sub




    Private Sub TP_Analyse_Click(sender As Object, e As EventArgs) Handles TP_Analyse.Click

        'Load_NetFlow_LiveChart()
    End Sub

    '========================================================================================================
    '======                                                                                            ======
    '======                                 nieuwe voor contract                                       ======
    '======                                                                                            ======
    '========================================================================================================

    Private originalContract As ContractModel ' Bewaar deze op class-niveau in main.vb

    Private Sub MapContractToUI(contract As ContractModel)
        Lbl_Contract_pkid.Text = contract.Id.ToString()
        Lbl_00_Contract__name.Text = contract.Name

        ' --- FIX: WinForms SelectedValue Quirk ---
        ' Always set SelectedIndex to -1 before assigning SelectedValue.
        ' This forces WinForms to fetch the DisplayMember, even if you click the exact same contract twice!

        ' Map Target
        Cmx_01_contract_fk_target_id.SelectedIndex = -1
        Cmx_01_contract_fk_target_id.SelectedValue = contract.FkTargetId

        ' If it's inactive and not in the active list, fetch the text directly
        If Cmx_01_contract_fk_target_id.SelectedIndex = -1 AndAlso contract.FkTargetId > 0 Then
            Dim targetName = QuerySQL($"SELECT CONCAT(name, ', ', name_add) FROM target WHERE id = {contract.FkTargetId}")
            If targetName IsNot Nothing AndAlso Not IsDBNull(targetName) Then
                Cmx_01_contract_fk_target_id.Text = targetName.ToString()
            End If
        End If

        ' Map Relation
        Cmx_00_contract_fk_relation_id.SelectedIndex = -1
        Cmx_00_contract_fk_relation_id.SelectedValue = contract.FkRelationId

        ' If it's inactive and not in the active list, fetch the text directly
        If Cmx_00_contract_fk_relation_id.SelectedIndex = -1 AndAlso contract.FkRelationId > 0 Then
            Dim relationName = QuerySQL($"SELECT CONCAT(name, ', ', name_add) FROM relation WHERE id = {contract.FkRelationId}")
            If relationName IsNot Nothing AndAlso Not IsDBNull(relationName) Then
                Cmx_00_contract_fk_relation_id.Text = relationName.ToString()
            End If
        End If

        ' Map Internal Account (Fonds)
        Cmx_Contract_fk_account_id.SelectedIndex = -1
        Cmx_Contract_fk_account_id.SelectedValue = contract.FkAccountId

        ' If it's inactive and not in the active list, fetch the text directly
        If Cmx_Contract_fk_account_id.SelectedIndex = -1 AndAlso contract.FkAccountId > 0 Then
            Dim accName = QuerySQL($"SELECT CONCAT(id, ' ', name) FROM account WHERE id = {contract.FkAccountId}")
            If accName IsNot Nothing AndAlso Not IsDBNull(accName) Then
                Cmx_Contract_fk_account_id.Text = accName.ToString()
            End If
        End If

        ' Map Target Type & Radio Buttons
        If contract.FkTargetId > 0 Then
            Dim targetTypeObj = QuerySQL($"SELECT ttype FROM target WHERE id = {contract.FkTargetId}")
            If targetTypeObj IsNot Nothing AndAlso Not IsDBNull(targetTypeObj) Then
                Dim targetType As String = targetTypeObj.ToString()
                Tbx_Contract_ttype.Text = targetType

                If targetType = "Kind" Then Rbn_00_contract_child.Checked = True
                If targetType = "Oudere" Then Rbn_00_contract_elder.Checked = True
                If targetType = "Overig" Then Rbn_00_contract_other.Checked = True
            End If
        End If

        ' Set Term FIRST to prevent division-by-zero during TextChanged events
        Cmx_02_Contract__term.Text = contract.Term.ToString()

        Tbx_11_Contract__donation.Text = contract.Donation.ToString("N2")
        Tbx_11_contract__overhead.Text = contract.Overhead.ToString("N2")

        ' Explicitly calculate the UI-only fields
        Tbx_01_contract_yeartotal.Text = (contract.Donation + contract.Overhead).ToString("N2")
        If contract.Term > 0 Then
            Tbx_contract_period_amt.Text = ((contract.Donation + contract.Overhead) / contract.Term).ToString("N2")
        End If

        Dtp_31_contract__startdate.Value = contract.StartDate
        Dtp_31_contract__enddate.Value = contract.EndDate
        Tbx_00_contract__description.Text = contract.Description
        Chx_00_contract__autcol.Checked = contract.Autcol

        ' Ensure the autcol visibility rules synchronize
        Lbl_00_contract_autcol.Visible = contract.Autcol
        dtp_contract_relation_date.Visible = contract.Autcol
        Lbl_contract_mach_datum.Visible = contract.Autcol
        Lbl_contract_macht_kenm.Visible = contract.Autcol
    End Sub


    Private Function MapUIToContract() As ContractModel
        Dim contract As New ContractModel()
        Integer.TryParse(Lbl_Contract_pkid.Text, contract.Id)
        contract.Name = Lbl_00_Contract__name.Text

        contract.FkTargetId = CInt(Cmx_01_contract_fk_target_id.SelectedValue)
        contract.FkRelationId = CInt(Cmx_00_contract_fk_relation_id.SelectedValue)

        contract.Donation = Tbx2Dec(Tbx_11_Contract__donation.Text)
        contract.Overhead = Tbx2Dec(Tbx_11_contract__overhead.Text)
        contract.Term = CInt(Cmx_02_Contract__term.Text)

        contract.StartDate = Dtp_31_contract__startdate.Value
        contract.EndDate = Dtp_31_contract__enddate.Value
        contract.Description = Tbx_00_contract__description.Text
        contract.Autcol = Chx_00_contract__autcol.Checked
        Return contract
    End Function




    Private Sub ApplyContractUIRules(contract As ContractModel)
        ' Rename the variable to avoid conflicting with the function name
        Dim hasFuture As Boolean = HasFutureVersion(contract.Name, contract.StartDate)
        Dim isClosed As Boolean = (contract.EndDate < Date.Today)

        Pan_Contract_Date_New.Visible = False

        ' Reset fields to 'Enabled' by default
        Tbx_11_Contract__donation.Enabled = True
        Tbx_11_contract__overhead.Enabled = True
        Chx_00_contract__autcol.Enabled = True
        Cmx_02_Contract__term.Enabled = True


        Dtp_31_contract__startdate.Enabled = False ' Startdate is never editable after creation
        Dtp_31_contract__enddate.Enabled = True    ' Editable by default to allow termination

        ' Handle the Status Label and lock the UI using the renamed variable
        If hasFuture Then
            Lbl_Contract_Status.Text = "Contract geblokkeerd, er bestaat een nieuwe versie"
            Lbl_Contract_Status.Visible = True
            Lbl_Contract_Status.ForeColor = Color.DarkRed

            ' Block input
            Tbx_11_Contract__donation.Enabled = False
            Tbx_11_contract__overhead.Enabled = False
            Chx_00_contract__autcol.Enabled = False
            Cmx_02_Contract__term.Enabled = False
            Dtp_31_contract__enddate.Enabled = False ' Locked because the future version dictates the timeline

        ElseIf isClosed Then
            Lbl_Contract_Status.Text = "Contract is beëindigd"
            Lbl_Contract_Status.Visible = True
            Lbl_Contract_Status.ForeColor = Color.DarkRed

            ' Block input
            Tbx_11_Contract__donation.Enabled = False
            Tbx_11_contract__overhead.Enabled = False
            Chx_00_contract__autcol.Enabled = False
            Cmx_02_Contract__term.Enabled = False
            Dtp_31_contract__enddate.Enabled = False ' Already terminated

        Else
            ' Hide label if everything is fine and active
            Lbl_Contract_Status.Visible = False
        End If

        ' Description is always editable
        Tbx_00_contract__description.Enabled = True
    End Sub


    Private Sub DeleteContract()
        If originalContract Is Nothing Then Exit Sub

        ' 1. Ensure it's actually a future contract
        If originalContract.StartDate <= Date.Today Then
            MsgBox("Alleen toekomstige wijzigingen (contracten met een startdatum in de toekomst) kunnen worden verwijderd.", vbExclamation)
            Exit Sub
        End If

        ' 2. Ask for confirmation
        If MsgBox("Weet u zeker dat u deze geplande wijziging wilt verwijderen? De einddatum van de huidige actieve versie zal worden hersteld.", vbYesNo + vbQuestion, "Toekomstig contract verwijderen") = vbYes Then

            ' 3. Execute the database transaction
            DeleteFutureContract(originalContract.Id, originalContract.Name)

            MsgBox("Toekomstige wijziging verwijderd. De vorige versie is succesvol hersteld.")

            ' 4. Retrieve the ID of the restored contract (which is now the max ID for this contract name)
            Dim restoredId As Integer = 0
            Try
                Dim result = QuerySQL($"SELECT MAX(id) FROM contract WHERE name = '{originalContract.Name}'")
                If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                    restoredId = Convert.ToInt32(result)
                End If
            Catch ex As Exception
                ' Safely ignore
            End Try

            ' 5. Refresh UI and focus the restored contract
            RefreshContractList(restoredId)
        End If
    End Sub


    Private Sub AddContractUI()
        isManualChange = False
        Add_Mode = True
        Empty_Tabpage()

        ' 1. Set default form values for a new Contract
        Rbn_00_contract_child.Checked = True
        Lbl_00_Contract__name.Text = Contract_number("K")

        ' Fetch default donation/overhead from settings
        Tbx_11_Contract__donation.Text = QuerySQL("select value from settings where label ilike 'standaard_bedrag_kind'")
        Tbx_11_contract__overhead.Text = QuerySQL("select value from settings where label ilike 'standaard_overhead_kind'")

        ' Default start date is the 1st of the next month
        Dtp_31_contract__startdate.Value = New DateTime(Date.Today.Year, Date.Today.Month, 1).AddMonths(1)
        Cmx_02_Contract__term.Text = "12"
        Cbx_00_contract__active.Checked = True

        ' Load targets corresponding to the default 'Kind' selection
        Load_Combobox(Cmx_01_contract_fk_target_id, "id", "name",
                  $"SELECT t.id, t.name||', '||t.name_add as name FROM target t WHERE t.ttype='Kind' And t.active=true ORDER BY t.name")
        Cmx_01_contract_fk_target_id.Text = ""

        ' 2. Enable/Disable specific controls for Add Mode
        Pan_Contract_Date_New.Visible = False
        Pan_contract_select_target.Enabled = True
        Dtp_31_contract__startdate.Enabled = True
        Chx_00_contract__autcol.Enabled = False

        ' Ensure comboboxes are accessible
        Cmx_01_contract_fk_target_id.Enabled = True
        Cmx_00_contract_fk_relation_id.Enabled = True
        Cmx_Contract_fk_account_id.Enabled = True

        Lbl_Add_mode.Text = "Add_mode=True"
        isManualChange = True
    End Sub

    Private Sub SaveContract()
        ' 1. Validate
        If Cmx_Contract_fk_account_id.Text = "" And Cmx_00_contract_fk_relation_id.Text = "" Then
            MsgBox("Kies ofwel een externe sponsor ofwel een intern fondsaccount.", vbExclamation)
            Exit Sub
        End If

        ' 2. Read the UI into our model
        Dim uiContract As ContractModel = MapUIToContract()
        Dim savedId As Integer = 0

        ' 3. Save to database
        If Add_Mode Then
            If uiContract.StartDate < Date.Today Then
                MsgBox("De startdatum van een nieuw contract mag niet in het verleden liggen.", vbExclamation)
                Exit Sub
            End If

            Dim overlappingName As String = Basisadmin.GetOverlappingContract(uiContract.FkTargetId, uiContract.FkRelationId, uiContract.StartDate)
            If overlappingName <> String.Empty Then
                MsgBox($"Er loopt al een contract ({overlappingName}) voor deze combinatie van sponsor en doel. " &
                   "Beëindig deze eerst alvorens een nieuw contract af te sluiten.", vbExclamation)
                Exit Sub
            End If

            ' --- FORCE ACTIVE FOR ALL NEW CONTRACTS ---
            uiContract.Active = True

            savedId = Basisadmin.InsertNewContract(uiContract)
            MsgBox("Nieuw contract succesvol opgeslagen!")

        Else
            ' Edit Mode: Check if financial/structural fields changed requiring a new version
            If originalContract.RequiresNewVersion(uiContract) Then
                Dim newVersionStartDate As Date = Dtp_30_Contract_Change.Value

                If newVersionStartDate <= originalContract.StartDate Then
                    MsgBox("De ingangsdatum van de nieuwe versie moet na de ingangsdatum van de huidige versie liggen.", vbExclamation)
                    Exit Sub
                End If

                uiContract.StartDate = newVersionStartDate

                ' New future versions are also always active 
                uiContract.Active = True

                Basisadmin.CreateNewContractVersion(originalContract.Id, uiContract)
                MsgBox("Er is een nieuwe versie van het contract aangemaakt!")

                ' Fetch the new ID by querying the max ID for this contract name
                savedId = Convert.ToInt32(QuerySQL($"SELECT MAX(id) FROM contract WHERE name = '{uiContract.Name}'"))
            Else
                ' Update basic info (Description and EndDate)
                ' Automatically terminate (set Active = False) if the new EndDate is in the past
                uiContract.Active = (uiContract.EndDate >= Date.Today)

                Basisadmin.UpdateContractBasicInfo(uiContract)
                savedId = uiContract.Id
                MsgBox("Contractbijwerking succesvol opgeslagen.")
            End If
        End If
        ' --- NEW: UPDATE THE BUDGET FOR THE TARGET ACCOUNT ---

        Try
            ' Fix: Use ILIKE to ignore case-sensitivity issues ('Doel' vs 'doel')
            Dim query As String = $"SELECT id FROM account WHERE source ILIKE 'doel' AND f_key = {uiContract.FkTargetId}"
            Dim targetAccountIdObj = QuerySQL(query)

            ' Ensure the result isn't Nothing AND isn't an empty string
            If targetAccountIdObj IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(targetAccountIdObj.ToString()) Then
                Dim targetAccountId As Integer = Convert.ToInt32(targetAccountIdObj)
                Calculate_Budget(targetAccountId)
            End If
        Catch ex As Exception
            MsgBox($"Budget not (re)calculated, for error {ex.Message}", vbExclamation)
        End Try
        ' -----------------------------------------------------
        ' -----------------------------------------------------
        ' 4. Finalize UI State
        Add_Mode = False
        Lbl_Add_mode.Text = "Add_mode=False"

        ' Lock down the target/relation comboboxes again
        Pan_contract_select_target.Enabled = False
        Cmx_01_contract_fk_target_id.Enabled = False
        Cmx_00_contract_fk_relation_id.Enabled = False
        Cmx_Contract_fk_account_id.Enabled = False
        Pan_Contract_Date_New.Visible = False

        ' --- RELOAD AND FOCUS THE NEW/UPDATED CONTRACT ---
        RefreshContractList(savedId)
    End Sub


    Private Sub RefreshContractList(Optional forceSelectId As Integer = 0)
        If username = "" Then Exit Sub

        ' 1. Remember the current state and disable the global dirty-tracker
        Dim previousState As Boolean = isManualChange
        isManualChange = False

        ' 2. Determine which ID to select after binding
        Dim currentId As Integer = forceSelectId
        If currentId = 0 AndAlso Lbx_Basis.SelectedValue IsNot Nothing AndAlso TypeOf Lbx_Basis.SelectedValue Is Integer Then
            currentId = CInt(Lbx_Basis.SelectedValue)
        End If

        ' Fetch the secure data
        Dim dt As DataTable = Basisadmin.GetContractList(Searchbox2.Text, Cbx_LifeCycle2.Text)

        ' Bind it directly
        Lbx_Basis.DataSource = dt
        Lbx_Basis.DisplayMember = "name"
        Lbx_Basis.ValueMember = "id"

        ' Restore or force selection
        If currentId > 0 Then
            Try
                Lbx_Basis.SelectedValue = currentId
            Catch
            End Try
        End If

        ' 2. Safely restore the global dirty-tracker
        isManualChange = previousState
    End Sub



    '========================================================================================================
    '======                                                                                            ======
    '======                                 A L G E M E E N                                            ======
    '======                                                                                            ======
    '========================================================================================================
    Private Sub SPAS_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Login.ShowDialog()
        AttachHandlers(Me.Controls)
    End Sub

    Private Sub FieldChangedHandler(sender As Object, e As EventArgs)


        If Not isManualChange Then Return

        ' FIX 3: Verify the value actually changed (prevents same-value re-selection triggering edit mode)
        Dim ctrl As Control = TryCast(sender, Control)
        'MsgBox($"Event fired by: {If(ctrl IsNot Nothing, ctrl.Name, "Unknown")} - Type: {sender.GetType().Name}")

        If ctrl IsNot Nothing AndAlso ctrl.Tag IsNot Nothing Then
            If ctrl.Text = ctrl.Tag.ToString() Then Return ' Value hasn't actually changed
        End If

        ' Enable Save and Cancel, disable other buttons
        If TC_Main.SelectedTab.Name <> "Incasso" And TC_Main.SelectedTab.Name <> "Tussenrekening" Then Enable_Buttons(True, False)

        If TC_Main.SelectedIndex = 0 Then Lbx_Basis.Enabled = False
        If TC_Main.SelectedIndex = 1 Then Dgv_Bank.Enabled = False

        ' Optional: Update the Tag to the new value so it doesn't keep triggering unnecessarily
        If ctrl IsNot Nothing Then ctrl.Tag = ctrl.Text
    End Sub
    Private Sub AttachHandlers(controls As Control.ControlCollection)
        Dim excludedPanels As New List(Of String) From {"Testpanel", "Pan_bank", "Pan_Bank2", "Pan_Incasso", "Pan_Incasso_Views"}

        ' FIX 1: Exclude navigation/search controls from being tracked for changes
        Dim excludedControls As New List(Of String) From {"Cmx_Excasso_Select", "Dgv_Uitkering_Account_Details"}

        For Each ctrl As Control In controls

            ' Skip this control if it's in the excluded controls list
            If excludedControls.Contains(ctrl.Name) Then Continue For

            If TypeOf ctrl Is System.Windows.Forms.TextBox AndAlso Not excludedPanels.Contains(ctrl.Parent.Name) Then
                AddHandler DirectCast(ctrl, System.Windows.Forms.TextBox).TextChanged, AddressOf FieldChangedHandler

            ElseIf TypeOf ctrl Is System.Windows.Forms.ComboBox AndAlso Not excludedPanels.Contains(ctrl.Parent.Name) Then
                ' FIX 2: Use SelectionChangeCommitted instead of SelectedIndexChanged
                AddHandler DirectCast(ctrl, System.Windows.Forms.ComboBox).SelectionChangeCommitted, AddressOf FieldChangedHandler

                ' Optional: Keep TextChanged only if users are allowed to type custom text into the comboboxes
                AddHandler DirectCast(ctrl, System.Windows.Forms.ComboBox).TextChanged, AddressOf FieldChangedHandler

            ElseIf TypeOf ctrl Is System.Windows.Forms.RadioButton AndAlso Not excludedPanels.Contains(ctrl.Parent.Name) Then
                AddHandler DirectCast(ctrl, System.Windows.Forms.RadioButton).CheckedChanged, AddressOf FieldChangedHandler
            ElseIf TypeOf ctrl Is System.Windows.Forms.DateTimePicker AndAlso Not excludedPanels.Contains(ctrl.Parent.Name) Then
                AddHandler DirectCast(ctrl, System.Windows.Forms.DateTimePicker).ValueChanged, AddressOf FieldChangedHandler
            ElseIf TypeOf ctrl Is System.Windows.Forms.CheckBox AndAlso Not excludedPanels.Contains(ctrl.Parent.Name) Then
                AddHandler DirectCast(ctrl, System.Windows.Forms.CheckBox).CheckedChanged, AddressOf FieldChangedHandler
            ElseIf TypeOf ctrl Is System.Windows.Forms.DataGridView AndAlso Not excludedPanels.Contains(ctrl.Parent.Name) Then
                AddHandler DirectCast(ctrl, System.Windows.Forms.DataGridView).CellValueChanged, AddressOf FieldChangedHandler
            End If

            ' Recursively handle nested controls
            If ctrl.HasChildren Then
                AttachHandlers(ctrl.Controls)
            End If
        Next
    End Sub

    Private Sub UpdateControlTags(parentControl As Control)
        For Each ctrl As Control In parentControl.Controls
            ' Use fully qualified names to avoid the VisualStyleElement conflict
            If TypeOf ctrl Is System.Windows.Forms.TextBox OrElse TypeOf ctrl Is System.Windows.Forms.ComboBox Then
                ctrl.Tag = ctrl.Text
            End If

            ' Recursively search inside panels, groupboxes, or tab pages
            If ctrl.HasChildren Then
                UpdateControlTags(ctrl)
            End If
        Next
    End Sub

    Private Sub TabControl_Selected(sender As Object, e As TabControlEventArgs) Handles TC_Main.Selected
        'dit is tijdelijk, zolang de save-button niet enabled/disabled moet worden
        current_tabpage = TC_Main.SelectedIndex
    End Sub

    Sub Enable_Buttons(ByVal SaveCancel As Boolean, AddDelete As Boolean)


        MenuSave.Enabled = SaveCancel
        MenuCancel.Enabled = SaveCancel

        MenuAdd.Enabled = AddDelete
        MenuDelete.Enabled = AddDelete

    End Sub

    Private Sub TC_Main_Selecting(sender As Object, e As TabControlCancelEventArgs) Handles TC_Main.Selecting, TC_Object.Selecting
        If current_tabpage = 2 Then 'workaround omdat in tabpage bank het uitsc
        Else
            If MenuSave.Enabled Or MenuCancel.Enabled Then
                MessageBox.Show("Bewaar of annuleer de huidige bewerking.", "Actie Nodig", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                e.Cancel = True
            End If
            'Enable_Buttons(False, False)
        End If
    End Sub



    Sub StoreInitialValues(controls As Control.ControlCollection)
        'MsgBox("original values stored")
        For Each ctrl As Control In controls
            If TypeOf ctrl Is System.Windows.Forms.TextBox Then
                originalValues(ctrl.Name) = DirectCast(ctrl, System.Windows.Forms.TextBox).Text
            ElseIf TypeOf ctrl Is System.Windows.Forms.ComboBox Then
                originalValues(ctrl.Name) = DirectCast(ctrl, System.Windows.Forms.ComboBox).SelectedIndex
            ElseIf TypeOf ctrl Is System.Windows.Forms.CheckBox Then
                originalValues(ctrl.Name) = DirectCast(ctrl, CheckBox).Checked
            ElseIf TypeOf ctrl Is System.Windows.Forms.RadioButton Then
                originalValues(ctrl.Name) = DirectCast(ctrl, RadioButton).Checked
            ElseIf TypeOf ctrl Is System.Windows.Forms.DateTimePicker Then
                originalValues(ctrl.Name) = DirectCast(ctrl, DateTimePicker).Value
            End If

            ' Recursively handle nested controls
            If ctrl.HasChildren Then
                StoreInitialValues(ctrl.Controls)
            End If
        Next
    End Sub

    Sub Show_buttons()
        'Knoppen
        Dim i = TC_Main.SelectedIndex
        Dim j = TC_Boeking.SelectedIndex
        Dim a = TC_Management.SelectedIndex
        'CRUD-knoppen
        MenuAdd.Visible = (i = 0) Or (i = 4 And j = 2) Or (a = 0)
        MenuSave.Visible = i = 0 Or i = 2 Or i = 3 Or (i = 4 And j < 3) Or i = 6 Or i = 1
        MenuCancel.Visible = i = 0 Or i = 3 Or (i = 4 And j < 3) Or i = 6 Or i = 1 Or i = 7
        MenuDelete.Visible = i = 0 Or i = 2 Or i = 3 Or (i = 4 And j < 3 And j <> 1) Or i = 6
        'Outputknoppen
        Menu_Print.Visible = (i = 2 Or i = 3)
        Menu_Export.Visible = (i > 0 And i <> 4) Or (i = 4 And (j = 1 Or j = 3))
        'Knoppen specifiek voor bank
        MenuCategoriseer.Visible = (i = 1)
        MenuBanktransactie.Visible = (i = 1)
        MenuUploadAlles.Visible = (i = 1)
    End Sub

    Sub InitLoad()
        If username = "" Then Exit Sub

        RunSQL("Update contract Set Active ='false' where enddate < current_date", "NULL", "SPAS_Load")
        Load_Comboboxes()
        TC_Object.SelectedIndex = 0

        Load_Table()
        If Lbx_Basis.Items.Count = 0 Then Empty_Tabpage()
        nocat = QuerySQL("SELECT value FROM settings WHERE label='nocat'")
        Load_Account_Settings()
        report_year = QuerySQL("select min(extract (year from date)) from journal")

        Dim sql = $"SELECT module, name, sql from query where category = 'Overzicht' order by module, name;"
        Populate_DataTree(sql, ReportTree)
        sql = $"select g.name, a.name  from account a left join accgroup g on g.id = a.fk_accgroup_id where a.active = true and g.active = true order by g.name, a.name"
        'Populate_DataTree_New(sql, AccountTree)
        Show_buttons()


    End Sub


    '========================================================================================================
    '======                                                                                            ======
    '======                              C O M B O B O X E N                                           ======
    '======                                                                                            ======
    '========================================================================================================
    Sub Load_Cmx_Bank_Account()
        Load_Combobox(Cmx_Bank_Account, "id", "name", "SELECT a.id, a.name FROM account a WHERE a.active = True ORDER BY a.source, a.name")
    End Sub


    Sub Load_Comboboxes()
        'can go wrong if tables are empty
        Load_Cmx_Bank_Account()

        Load_Combobox(Cmx_01_cp__fk_bankacc_id, "id", "name", "SELECT ba.id, ba.name||'/'||ba.accountno as name FROM bankacc ba WHERE ba.expense=True AND ba.active=TRUE ORDER BY ba.name")
        Load_Combobox(Cmx_Incasso_Bankaccount, "id", "name", "SELECT ba.id, ba.accountno AS name FROM bankacc ba WHERE ba.expense=FALSE AND ba.active=TRUE ORDER BY name")
        Load_Combobox(Cmx_01_Target__fk_cp_id, "id", "name", "SELECT cp.id, CONCAT(cp.name, ', ', cp.name_add) as name FROM cp WHERE cp.active=True ORDER BY cp.name")
        Load_Combobox(Cmx_00_contract_fk_relation_id, "id", "name", "SELECT r.id, CONCAT(r.name, ', ', r.name_add) as name FROM relation r WHERE r.active=TRUE ORDER BY r.name")
        Load_Combobox(Cmbx_journaalposten_relatie, "id", "name", "SELECT r.id, CONCAT(r.name, ', ', r.name_add) as name FROM relation r ORDER BY name")
        Load_Combobox(Cmbx_journaalposten_account, "id", "name", "SELECT a.id, a.name FROM account a ORDER BY name")
        Load_Combobox(Cmx_Bank_bankacc, "id", "name", "SELECT ba.id, CONCAT(ba.Name, '/', ba.accountno) as name FROM bankacc ba ORDER BY ba.name DESC")

        Load_Combobox(Cmx_Contract_fk_account_id, "id", "name", "SELECT a.id, CONCAT(a.id, ' ',a.name) As name FROM account a
                                          WHERE a.active=TRUE AND a.type = 'Generiek (fonds)' ORDER BY a.name")
        Load_Combobox(Cmx_01_account__fk_accgroup_id, "id", "name", "SELECT ag.id, ag.name FROM accgroup ag WHERE ag.active=True ORDER BY ag.name")
        Load_Combobox(Cmbx_Beheer_Accgroup, "id", "name", "SELECT ag.id, ag.name FROM accgroup ag WHERE ag.active=True ORDER BY ag.name")
        Populate_Single_Combobox(Cmbx_Reporting_Year, "select distinct extract (year from date) As Year from journal_archive 
                                            union select distinct min(extract (year from date)) from journal")

        Call Populate_Cmbx_Overboeking()

        'Populate_Cmx_Incasso_IncassoForm()
        'Cmx_Incasso_IncassoForm.SelectedIndex = -1
        Fill_Cmx_Journal_List()

        If Me.Dgv_Mgnt_Tables.Rows(8).Cells(1).Value > 0 Then
            Load_Combobox(Cmx_01_contract_fk_target_id, "id", "name", "SELECT t.id, t.name||', '||t.name_add as name FROM target t WHERE t.active=TRUE ORDER BY t.name")
        End If
        '@@@ hier gaat iets fout
        Fill_Cmx_Excasso_Select_Combined()
        Me.Cmbx_journaalposten_account.SelectedIndex = -1
        Me.Cmbx_journaalposten_relatie.SelectedIndex = -1
        Cmbx_Overboeking_Bron.SelectedIndex = -1
        Cmbx_Overboeking_Target.SelectedIndex = -1

    End Sub

    Sub Populate_Cmbx_Overboeking()
        Populate_Combobox(Cmbx_Overboeking_Bron, "select a.id, a.name, sum(j.amt1) from journal j left join account a on a.id=j.fk_account 
        WHERE a.active=True group by a.id, a.name having sum(amt1)>0::money ORDER BY a.name")
        Populate_Combobox(Cmbx_Overboeking_Target, "select a.id,a.name,COALESCE(SUM(j.amt1), 0::money) AS total_amt1
        from account a LEFT join journal j ON a.id = j.fk_account where a.active = TRUE GROUP by  a.id,  a.name ORDER by a.name;")
    End Sub

    Sub Populate_Cmx_Incasso_IncassoForm()
        Populate_Combobox(Cmx_Incasso_IncassoForm, "select CURRENT_DATE as Date, 'Nieuwe incasso' as name, 'Nieuw' as Status union
	    select distinct date, 'I'|| Substring(name,11,15) as name, trim(status) as Status from journal j where source= 'Incasso' order by status, date desc")
    End Sub

    '========================================================================================================
    '======                                                                                            ======
    '======                         B A S I S A D M I N I S T R A T I E                                ======
    '======                                                                                            ======
    '========================================================================================================

    Sub TC_Object_Click(sender As Object, e As EventArgs) Handles TC_Object.Click, TC_Object.SelectedIndexChanged
        If Not MenuSave.Enabled Then
            isManualChange = False

            If TC_Object.SelectedIndex = 0 Then
                RefreshContractList()
            Else
                Load_Table()
            End If

            isManualChange = True
            Enable_Buttons(False, True)
        End If
    End Sub

    '************************************      TARGET    *****************************************
    Private Sub Rbtn_Target_Child_CheckedChanged(sender As Object, e As EventArgs) Handles Rbtn_Target_Child.CheckedChanged
        If MenuSave.Enabled Then Tbx_01_Target__ttype.Text = Rbtn_Target_Child.Text
    End Sub

    Private Sub Rbtn_Target_Elder_CheckedChanged(sender As Object, e As EventArgs) Handles Rbtn_Target_Elder.CheckedChanged
        If MenuSave.Enabled Then Tbx_01_Target__ttype.Text = Rbtn_Target_Elder.Text
    End Sub

    Private Sub Rbtn_Target_Other_CheckedChanged(sender As Object, e As EventArgs) Handles Rbtn_Target_Other.CheckedChanged
        If MenuSave.Enabled Then Tbx_01_Target__ttype.Text = Rbtn_Target_Other.Text
    End Sub


    Sub Basis_Add()

        Dim t As String = TC_Object.SelectedIndex.ToString

        Add_Mode = True
        Manage_Buttons_Target(False, False, False, True, True, "Menu_Add_Click")
        Empty_Tabpage()


        If TC_Object.SelectedIndex = 0 Then  'additional functionality for contract management

            Dtp_31_contract__startdate.Value = Date.Today
            Me.Rbn_00_contract_child.Checked = True
            Rbn_00_contract_child.Checked = True
            '---------------- Temp solution of error
            Lbl_00_Contract__name.Text = Contract_number("K")
            Load_Combobox(Cmx_01_contract_fk_target_id, "id", "name", "Select t.id, t.Name||', '||t.name_add as name FROM target t
                                                        WHERE t.ttype='" & Rbn_00_contract_child.Text & "' AND t.active= TRUE ORDER BY t.name")
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
            Dtp_31_contract__startdate.Enabled = True
            Lbl_00_Contract__name.Text = Contract_number("K")
            Load_Combobox(Cmx_01_contract_fk_target_id, "id", "name", $"SELECT t.id, 
            t.name||', '||t.name_add) as name FROM target t WHERE t.ttype='{Rbn_00_contract_child.Text}' And t.active=true ORDER BY t.name")
            Cmx_01_contract_fk_target_id.Text = ""
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
        Lbl_Add_mode.Text = IIf(Add_Mode, "Add_mode=True", "Add_mode=False")
    End Sub

    Sub Cancel()
        If TC_Object.SelectedIndex = 0 Then
            ' --- NEW: Safely Restore Contract ---
            If originalContract IsNot Nothing Then
                isManualChange = False
                MapContractToUI(originalContract)
                ApplyContractUIRules(originalContract)
                StoreInitialValues(Me.Controls)
                isManualChange = True
            Else
                Empty_Tabpage()
            End If

            Add_Mode = False
            Lbl_Add_mode.Text = "Add_mode=False"
            Manage_Buttons_Target(True, True, True, False, False, "Cancel")

            ' Lock down UI fields
            Pan_contract_select_target.Enabled = False
            Cmx_01_contract_fk_target_id.Enabled = False
            Cmx_00_contract_fk_relation_id.Enabled = False
            Cmx_Contract_fk_account_id.Enabled = False
            Pan_Contract_Date_New.Visible = False
        Else
            ' --- OLD: Legacy Cancel ---
            Select_Obj2("Cancel")
            Manage_Buttons_Target(True, True, True, False, False, "Cancel")
            Add_Mode = False
            Pan_Target.Enabled = False
            Pan_contract_select_target.Enabled = False

            If TC_Object.SelectedIndex = 4 Then
                Lbl_Account_Budget_Difference.Text = ""
            End If
            Lbl_Add_mode.Text = IIf(Add_Mode, "Add_mode=True", "Add_mode=False")
        End If
    End Sub



    Sub Basis_Save()
        Dim tbl As String = Me.TC_Object.TabPages(Me.TC_Object.SelectedIndex).Name
        Dim val, val2 As Integer
        Dim errmsg = Handle_errors("")
        If errmsg <> "" Then
            MsgBox(errmsg)
            Exit Sub
        End If
        If Lbx_Basis.SelectedIndex = 6 Then
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
                        sqlstr &= $"INSERT INTO public.contract(fk_target_id, fk_relation_id, 
                                    donation, overhead, description, autcol, name, term,intern, fk_account_id) 
                                    SELECT fk_target_id, fk_relation_id, 
                                    donation, overhead, description, autcol, name, term,  
                                    intern, fk_account_id FROM contract WHERE id={val};"


                        RunSQL(sqlstr, "NULL", "MenuSave.Click upsert new version")

                        val2 = Convert.ToInt32(QuerySQL("Select MAX(id) FROM " & tbl))
                        '3 update new version with new values, startdate / enddate and active
                        sqlstr = "UPDATE contract SET startdate='" & _d1 & "', 
                           donation='" & Cur2(Replace(Tbx_11_Contract__donation.Text, ".", "")) & "', 
                           overhead='" & Cur2(Replace(Tbx_11_contract__overhead.Text, ".", "")) & "', 
                           enddate ='2999-12-31',active=true  
                           WHERE id=" & val2 & ";"

                        'MsgBox(Cur2(Replace(Tbx_11_Contract__donation.Text, ".", "")))
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
                    Dim acc_id As Integer = QuerySQL("select id from account where source = 'Doel' and f_key=" & Cmx_01_contract_fk_target_id.SelectedValue)
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
        Manage_Buttons_Target(True, True, True, False, False, "MenuSave.Click")
        Add_Mode = False
        reload = False
        Lbl_Add_mode.Text = IIf(Add_Mode, "Add_mode=True", "Add_mode=False")
    End Sub


    Private Sub Lbx_Basis_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Lbx_Basis.SelectedIndexChanged

        If InStr(sender.ToString, "System.Data.DataRowView") > 0 Then Exit Sub
        Click_Lbx_Basis()

    End Sub



    Sub Click_Lbx_Basis()
        If username = "" Then Exit Sub
        ' Abort immediately if the ListBox is still in the middle of data binding
        If String.IsNullOrEmpty(Lbx_Basis.ValueMember) Then Exit Sub

        ' Safely lock the dirty-tracker for nested operations
        Dim previousState As Boolean = isManualChange
        isManualChange = False

        If Lbx_Basis.Items.Count > 0 AndAlso Lbx_Basis.SelectedIndex <> -1 Then

            If TC_Object.SelectedIndex = 0 Then
                ' --- THE NEW CODE FOR CONTRACTS ---
                Empty_Tabpage()

                Dim contractId As Integer
                Try
                    If TypeOf Lbx_Basis.SelectedItem Is DataRowView Then
                        contractId = Convert.ToInt32(DirectCast(Lbx_Basis.SelectedItem, DataRowView)(Lbx_Basis.ValueMember))
                    Else
                        contractId = Convert.ToInt32(Lbx_Basis.SelectedValue)
                    End If
                Catch ex As Exception
                    ' Fallback: restore state and abort if extraction fails
                    isManualChange = previousState
                    Exit Sub
                End Try

                originalContract = Basisadmin.GetContractById(contractId)

                If originalContract IsNot Nothing Then
                    MapContractToUI(originalContract)
                    ApplyContractUIRules(originalContract)
                End If

                StoreInitialValues(Me.Controls)

            Else
                ' --- THE OLD CODE FOR EVERYTHING ELSE ---
                Select_Obj2("Lbx_Basis_SelectedIndexChanged")
            End If

        Else
            ' This ensures emptying the form on 0 search results DOES NOT trigger Save/Cancel!
            Empty_Tabpage()
        End If

        ' Safely restore the dirty-tracker
        isManualChange = previousState
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
        If Lbx_Basis.Items.Count <> 0 Then ind1 = Lbx_Basis.SelectedItem(Me.Lbx_Basis.ValueMember)
    End Sub
    Private Sub Tbx_Target__name_add_Leave(sender As Object, e As EventArgs) Handles Tbx_01_Target__name_add.Leave
        If Lbx_Basis.Items.Count <> 0 Then ind1 = Lbx_Basis.SelectedItem(Me.Lbx_Basis.ValueMember)
    End Sub

    Private Sub Tbx_CP__name_TextChanged(sender As Object, e As EventArgs) Handles Tbx_01_CP__name.TextChanged
        reload = True
        If Lbx_Basis.Items.Count = 0 Then Add_Mode = True
        Lbl_Add_mode.Text = IIf(Add_Mode, "Add_mode=True", "Add_mode=False")
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

    Private Sub Tbx_10_Relation__name_TextChanged(sender As Object, e As EventArgs) Handles Tbx_01_relation__name.TextChanged
        If Add_Mode Then Generate_Reference()
        If Lbx_Basis.Items.Count = 0 Then Add_Mode = True
        Lbl_Add_mode.Text = IIf(Add_Mode, "Add_mode=True", "Add_mode=False")
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
        Load_Combobox(Cmx_01_contract_fk_target_id, "id", "name", "Select t.id, t.Name||', '||t.name_add as name FROM target t
                                                        WHERE t.ttype='" & Rbn_00_contract_child.Text & "' ORDER BY t.name")
        '-------standaard_waarden ophalen
        Dim settingdata = Collect_data2("select value from settings where label ilike 'standaard_%_kind' order by label")
        Tbx_11_Contract__donation.Text = settingdata.Rows(0)(0)
        Tbx_11_contract__overhead.Text = settingdata.Rows(1)(0)

        '----------------------------

    End Sub

    Private Sub Tbx_11_contract__overhead_TextChanged(sender As Object, e As EventArgs) Handles Tbx_11_contract__overhead.TextChanged, Tbx_11_Contract__donation.TextChanged
        Calculate_contract_amounts()
    End Sub

    Private Sub Pic_Target__photo_DoubleClick(sender As Object, e As EventArgs) Handles Pic_Target__photo.DoubleClick, Pic_cp__photo.DoubleClick
        Save_Image(Pic_Target__photo)
    End Sub

    Private Sub Tbx_11_Contract__donation_Leave(sender As Object, e As EventArgs) Handles Tbx_11_Contract__donation.Leave
        Tbx_11_Contract__donation.Text = Tbx2Dec(Tbx_11_Contract__donation.Text)
    End Sub
    Private Sub Tbx_11_contract__overhead_Leave(sender As Object, e As EventArgs) Handles Tbx_11_contract__overhead.Leave
        Tbx_11_contract__overhead.Text = Tbx2Dec(Tbx_11_contract__overhead.Text)
    End Sub

    Private Sub Cmx_01_contract_fk_target_id_Leave(sender As Object, e As EventArgs) Handles Cmx_01_contract_fk_target_id.Leave
        If (Cmx_01_contract_fk_target_id.SelectedIndex = -1) Then
            Cmx_01_contract_fk_target_id.Focus()
            Exit Sub
        End If
        Exit Sub
        Dim id = Cmx_01_contract_fk_target_id.SelectedValue
        Try
            Pic_Contract_Target_photo.Image = BlobToImage(QuerySQL("SELECT photo FROM target WHERE id='" & id & "'"))

        Catch ex As Exception
            Pic_Contract_Target_photo.Image = Nothing
        End Try
    End Sub

    Private Sub Cmx_01_contract_fk_target_id_SelectedValueChanged(sender As Object, e As EventArgs) Handles Lbl_11_contract__fk_target_id.TextChanged

        Exit Sub
        Dim id = Lbl_11_contract__fk_target_id.Text
        'Tbx_Contract_ttype.Text = QuerySQL("Select ttype FROM target WHERE id=" & id)
        Lbl_11_contract__fk_target_id.Text = id.ToString
        Try
            Pic_Contract_Target_photo.Image = BlobToImage(QuerySQL("SELECT photo FROM target WHERE id='" & id & "'"))
        Catch ex As Exception
            Pic_Contract_Target_photo.Image = Nothing
        End Try
    End Sub

    Private Sub Cmx_01_contract_fk_relation_id_Leave(sender As Object, e As EventArgs) Handles Cmx_00_contract_fk_relation_id.SelectedIndexChanged
        If Not Add_Mode Then Exit Sub
        If (Cmx_00_contract_fk_relation_id.SelectedIndex = -1) Then
            Cmx_00_contract_fk_relation_id.Focus()
            Exit Sub
        End If

        Get_Sponsor_data()
    End Sub

    Private Sub Dtp_01_contract__enddate_Enter(sender As Object, e As EventArgs) Handles Dtp_31_contract__enddate.Enter
        oldend_date = Dtp_31_contract__enddate.Value
        Dim newEndDate As Date = Date.Today.AddMonths(1)

        If Not Add_Mode Then Dtp_31_contract__enddate.Value = New DateTime(newEndDate.Year, newEndDate.Month, 1).AddDays(-1) 'end of current month
    End Sub

    Private Sub Rbn_00_contract_elder_Click(sender As Object, e As EventArgs) Handles Rbn_00_contract_elder.Click
        Tbx_Contract_ttype.Text = "Oudere"
        Lbl_00_Contract__name.Text = Contract_number("O")
        Load_Combobox(Cmx_01_contract_fk_target_id, "id", "name",
                      $"SELECT t.id, t.name||', '||t.name_add as name FROM target t WHERE t.ttype='{Rbn_00_contract_elder.Text}' ORDER BY t.name")
        Dim presetdata = Collect_data2("select value from settings where label ilike 'standaard_%_oudere' order by label")
        Tbx_11_Contract__donation.Text = presetdata.Rows(0)(0)
        Tbx_11_contract__overhead.Text = presetdata.Rows(1)(0)
    End Sub

    Private Sub Rbn_00_contract_other_Click(sender As Object, e As EventArgs) Handles Rbn_00_contract_other.Click
        Tbx_Contract_ttype.Text = "Overig"
        Load_Combobox(Cmx_01_contract_fk_target_id, "id", "name", "SELECT t.id, t.name||', '||t.name_add as name FROM target t
                                                        WHERE t.ttype='" & Rbn_00_contract_other.Text & "' ORDER BY t.name")
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
            Add_Mode = False
            Pan_Target.Enabled = False

            If TC_Object.SelectedIndex = 0 Then  'additional functionality for contract management
                Handle_Contract_Fields()
                Pan_Contract_Date_New.Visible = False
            End If
            Exit Sub
        End If
        Lbl_Add_mode.Text = IIf(Add_Mode, "Add_mode=True", "Add_mode=False")
    End Sub
    Private Sub Tbx_01_contract_yeartotal_TextChanged(sender As Object, e As EventArgs) Handles Tbx_01_contract_yeartotal.TextChanged, Chx_00_contract__autcol.Click
        If Not isManualChange Then Exit Sub
        If isCanceling Then Exit Sub
        Try
            If Not Add_Mode And MenuSave.Enabled Then
                Dim firstDayOfNextMonth As DateTime = New DateTime(DateTime.Today.Year, DateTime.Today.Month, 1).AddMonths(1)
                Dtp_30_Contract_Change.Value = firstDayOfNextMonth

                Dim daysInThisMonth As Integer = DateTime.DaysInMonth(DateTime.Today.Year, DateTime.Today.Month)
                Dim lastDayOfCurrentMonth As New DateTime(DateTime.Today.Year, DateTime.Today.Month, daysInThisMonth)
                Dtp_31_contract__enddate.Value = lastDayOfCurrentMonth

                Pan_Contract_Date_New.Visible = True

                ' --- FIX: Lock the end date when creating a new version ---
                Dtp_31_contract__enddate.Enabled = False
            End If
            If Add_Mode Then
                Pan_Contract_Date_New.Visible = False
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub Dtp_30_Contract_Change_ValueChanged(sender As Object, e As EventArgs) Handles Dtp_30_Contract_Change.ValueChanged
        If Not isManualChange Then Exit Sub

        ' Automatically snap the old contract's end date to 1 day before the new version starts
        Dtp_31_contract__enddate.Value = Dtp_30_Contract_Change.Value.AddDays(-1)
    End Sub

    Private Sub Chx_00_contract__autcol_Click(sender As Object, e As EventArgs) Handles Chx_00_contract__autcol.Click
        If Not isManualChange Then Exit Sub

        ' ONLY check authorization if the user is enabling the checkbox
        If Chx_00_contract__autcol.Checked Then
            Dim dtp As String
            Dim rel_id = Cmx_00_contract_fk_relation_id.SelectedValue

            If Rbn_00_contract_child.Checked Then
                dtp = "date1"
            ElseIf Rbn_00_contract_elder.Checked Then
                dtp = "date2"
            Else
                dtp = "date3"
            End If

            Dim autcol_date As Date = QuerySQL("SELECT " & dtp & " FROM relation WHERE id=" & rel_id)

            If autcol_date > Date.Now Then
                ' FIX: Added .Text to Tbx_Contract_ttype
                MsgBox($"De sponsor heeft nog geen geldige incassomachtiging voor '{Tbx_Contract_ttype.Text}'; Automatische incasso kan (nog) niet geactiveerd worden voor dit contract.", vbCritical)
                Chx_00_contract__autcol.Checked = False
            End If
        End If
    End Sub


    Private Sub Cbx_00_contract__autcol_CheckedChanged(sender As Object, e As EventArgs) Handles Chx_00_contract__autcol.CheckedChanged
        Dim rel_id = Cmx_00_contract_fk_relation_id.SelectedValue
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



    Sub Manage_Buttons_Target(ByVal _add As Boolean, _searchbox As Boolean, d As Boolean, _menusave As Boolean, _cancel As Boolean, sender As String)
        Exit Sub
        If Cbx_LifeCycle2.Text = "Inactief" And MenuSave.Enabled Then
            MsgBox("Inactieve objecten kunnen niet gewijzigd worden.")
            Exit Sub
        End If
        Lbx_Basis.Enabled = _add
        MenuAdd.Enabled = _add
        MenuDelete.Enabled = _add
        MenuFilter.Enabled = _searchbox

        MenuSave.Enabled = _menusave
        MenuCancel.Enabled = _cancel
    End Sub
    Private Sub Tbx_BankAcc__accountno_TextChanged(sender As Object, e As EventArgs) Handles _
          Tbx_01_BankAcc__accountno.TextChanged, Tbx_01_BankAcc__name.TextChanged, Tbx_01_Accgroup__name.TextChanged, Tbx_01_Target__name.TextChanged, Cmx_01_account__fk_accgroup_id.TextUpdate,
          Tbx_01_Target__name_add.TextChanged, Tbx_01_Account__name.TextChanged, Tbx_01_CP__name_add.TextChanged,
          Tbx_00_Accgroup__subtype.TextChanged, Tbx_00_Accgroup__description.TextChanged, Tbx_01_Accgroup__name.TextChanged, Tbx_01_Accgroup__type.TextChanged,
          Rbtn_accgroup_Income.CheckedChanged, Rbtn_accgroup_expense.CheckedChanged, Rbtn_accgroup_transit.CheckedChanged

        reload = True
    End Sub


    Private Sub Tbx_10_Account__b_jan_TextChanged(sender As Object, e As EventArgs) Handles _
        Tbx_10_Account__b_jan.TextChanged, Tbx_10_Account__b_feb.TextChanged, Tbx_10_Account__b_mar.TextChanged,
        Tbx_10_Account__b_apr.TextChanged, Tbx_10_Account__b_may.TextChanged, Tbx_10_Account__b_jun.TextChanged,
        Tbx_10_Account__b_jul.TextChanged, Tbx_10_Account__b_aug.TextChanged, Tbx_10_Account__b_sep.TextChanged,
        Tbx_10_Account__b_oct.TextChanged, Tbx_10_Account__b_nov.TextChanged, Tbx_10_Account__b_dec.TextChanged

        If MenuSave.Enabled Then
            Calculate_Manual_Budgets()
        End If
    End Sub

    Sub Cmx_Bank_bankacc_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Cmx_Bank_bankacc.SelectedIndexChanged

        If Cmx_Bank_bankacc.SelectedIndex = -1 Then Cmx_Bank_bankacc.SelectedIndex = 0
        Fill_bank_transactions("Cmx_Bank_bankacc.SelectedIndexChanged", 0)
        If TC_Main.SelectedTab.Name = "Tab_Bank" Then isManualChange = False

    End Sub
    Sub Dgv_Bank_Click(sender As Object, e As EventArgs) Handles Dgv_Bank.Click, Dgv_Bank.SelectionChanged
        isManualChange = False
        If Dgv_Bank.Rows.Count = 0 Or Dgv_Bank.DataSource Is Nothing Then Exit Sub

        Try
            'Voorkomen dat de gebruiker incassso- of excassojobs gaat editen
            If Not IsDBNull(Dgv_Bank.SelectedCells(3).Value) Then
                If Strings.Left(Dgv_Bank.SelectedCells(3).Value, 16) = "Contract incasso" _
                    Or Strings.Left(Dgv_Bank.SelectedCells(3).Value, 7) = "Excasso" Then
                    Dgv_Bank_Account.EditMode = DataGridViewEditMode.EditProgrammatically
                    Cmx_Bank_Account.Enabled = False
                Else
                    Dgv_Bank_Account.EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2
                    Cmx_Bank_Account.Enabled = True
                End If
            End If

            Dim bankdata = Dgv_Bank.DataSource
            'vullen van niet aanpasbare velden
            Lbl_Bank_Relation.Text = Dgv_Bank.SelectedCells(2).Value
            If Not IsDBNull(Dgv_Bank.SelectedCells(8).Value) Then Lbl_Bank_Relation_account.Text = Dgv_Bank.SelectedCells(8).Value
            If Not IsDBNull(Dgv_Bank.SelectedCells(6).Value) Then Lbl_Bank_Code.Text = Dgv_Bank.SelectedCells(6).Value
            If Not IsDBNull(Dgv_Bank.SelectedCells(9).Value) Then Lbl_Bank_Afschrift.Text = Dgv_Bank.SelectedCells(9).Value
            Lbl_Transactie_totaal.Text = QuerySQL($"Select sum(credit-debit) from bank where seqorder ='{Lbl_Bank_Afschrift.Text}'")

            'Vullen van aanpasbare velden
            Tbx_Bank_Description.Text = Dgv_Bank.SelectedCells(3).Value
            If Chbx_Bank_ExtraInfo_voor.Checked Then
                If Strings.InStr(Tbx_Bank_Description.Text, " | ") > 0 Then
                    Tbx_Bank_Extra_Info.Text = Strings.Left(Tbx_Bank_Description.Text, Strings.InStr(Tbx_Bank_Description.Text, " | ") - 1)
                Else
                    Tbx_Bank_Extra_Info.Text = ""
                End If
            End If

            Fill_Journals_by_bank(Dgv_Bank.SelectedCells(0).Value)

            '' vul de combobox alvast met het doel waarmee de bankrelatie een sponsorcontract heeft - dan betreft het waarschijnlijk een extra gift
            If Dgv_Bank.Rows(Dgv_Bank.SelectedCells(2).RowIndex).DefaultCellStyle.ForeColor = Color.DarkRed And Trim(Lbl_Bank_Code.Text) = "cb" Then
                Dim sqlstr = $"
                Select ac.name From account ac
                Left Join target t on t.id = ac.f_key And source='Doel'
                Left Join contract c on c.fk_target_id = t.id
                Left Join relation r on r.id = c.fk_relation_id
                Where R.iban = '{Lbl_Bank_Relation_account.Text}' 
                And R.active = True limit 1
                "
                Cmx_Bank_Account.Text = QuerySQL(sqlstr)
            Else
                Cmx_Bank_Account.Text = ""
            End If

            Tbx_Bank_Amount.Text = 0
            'Vul het bedrag alvast in met het nog niet toegewezen bedrag
            For x = 0 To Dgv_Bank_Account.Rows.Count - 1
                If Dgv_Bank_Account.Rows(x).Cells(1).Value = "[Niet toegewezen]" Then
                    Tbx_Bank_Amount.Text = Dgv_Bank_Account.Rows(x).Cells(2).Value
                    Exit For
                End If
            Next x
            'Een auto-cat banktransactie moet handmatig gecontroleerd worden. Het aanklikken wordt beschouwd als controle, waarna de registratie van fk_journal_name wordt afgerond
            If Dgv_Bank.Rows(Dgv_Bank.SelectedCells(2).RowIndex).Cells(12).Value = "Auto-cat" Then
                RunSQL("Update Bank set fk_journal_name='Bank' where id='" & Dgv_Bank.SelectedCells(0).Value & "'", "NULL", "auto_cat")
                Dgv_Bank.Rows(Dgv_Bank.SelectedCells(2).RowIndex).DefaultCellStyle.ForeColor = Color.DarkGreen
                Dgv_Bank.Rows(Dgv_Bank.SelectedCells(2).RowIndex).Cells(12).Value = "Bank" '
            End If
        Catch ex As Exception
            'MsgBox(ex.ToString)
        End Try
        isManualChange = True
        Enable_Buttons(False, False)
        Dgv_Bank.Enabled = True

    End Sub

    Private Sub Btn_Bank_Add_Journal_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles Cmx_Bank_Account.SelectionChangeCommitted, Cmx_Bank_Account.Click

        isManualChange = False
        'Add_Journal_post_to_banktransaction()
    End Sub







    Private Sub Dgv_Test_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles _
        Dgv_Bank_Account.CellValueChanged


        If Not isManualChange Then Exit Sub
        'isManualChange = True
        If Dgv_Bank_Account.Rows.Count = 0 Then  'dit kan alleen voorkomen als er een error is opgetreden. 
            'MsgBox("Er is een fout opgetreden, u kunt wel doorgaan")
            Exit Sub
        End If

        Try
            If IsDBNull(Dgv_Bank_Account.CurrentCell.Value) Then
                'MsgBox("Ongeldige invoer")
                Exit Sub
            End If
        Catch
            'Exit Sub
        End Try

        If Not IsNumeric(Dgv_Bank_Account.CurrentCell.Value) Then
            'MsgBox("Ongeldige invoer")
            Exit Sub
        End If


        If Check_Change_Bank_Categories(True) = False Then Exit Sub

        Calculate_Total_Booked("Dgv_Test_CellValueChanged")
        'Save_Banktransaction_Accounts()
        'Update_Category_Status()
        Dgv_Bank_Account.Rows(e.RowIndex).Tag = "Modified"
    End Sub



    Private Sub Tbx_Bank_Search_TextChanged(sender As Object, e As EventArgs) Handles Tbx_Bank_Search.TextChanged
        Fill_bank_transactions("Tbx_Bank_Search.TextChanged", 0)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs)
        RunSQL("TRUNCATE TABLE bank", "NULL", "")
        RunSQL("Delete From journal WHERE source='Bank'", "NULL", "")
    End Sub

    Private Sub Rbn_Relation_1_CheckedChanged(sender As Object, e As EventArgs) Handles Rbn_Relation_1.Click
        If Rbn_Relation_1.Checked Then Tbx_01_Relation__title.Text = Rbn_Relation_1.Text
    End Sub

    Private Sub Rbn_Relation_2_CheckedChanged(sender As Object, e As EventArgs) Handles Rbn_Relation_2.Click
        If Rbn_Relation_2.Checked Then Tbx_01_Relation__title.Text = Rbn_Relation_2.Text
    End Sub

    Private Sub Rbn_Relation_3_CheckedChanged(sender As Object, e As EventArgs) Handles Rbn_Relation_3.Click
        If Rbn_Relation_3.Checked Then Tbx_01_Relation__title.Text = Rbn_Relation_3.Text
    End Sub

    Private Sub Rbn_Relation_4_CheckedChanged(sender As Object, e As EventArgs) Handles Rbn_Relation_4.Click
        If Rbn_Relation_4.Checked Then Tbx_01_Relation__title.Text = Rbn_Relation_4.Text
    End Sub
    Private Sub Rbn_Relation_5_CheckedChanged(sender As Object, e As EventArgs) Handles Rbn_Relation_5.Click
        If Rbn_Relation_5.Checked Then Tbx_01_Relation__title.Text = Rbn_Relation_5.Text
    End Sub
    Private Sub Rbn_Relation_6_CheckedChanged(sender As Object, e As EventArgs) Handles Rbn_Relation_6.Click
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

        Dim rel_id = Cmx_00_contract_fk_relation_id.SelectedValue
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
        'tijdelijk uitgezet: Create_Incassolist()
        If Rbn_Incasso_SEPA.Checked Then
            Prepare_Datagridview(Dgv_Incasso, Nothing, {"TZ205", "NG080", "TZ160", "TZ080", "TZ090", "DZ090"})
        Else
            Prepare_Datagridview(Dgv_Incasso, Nothing, {"TZ210", "TZ210", "TZ070", "TZ070", "NG080", "NG080", "HZ080", "HZ080"})
        End If

        Rbn_Incasso_SEPA.Checked = True
    End Sub

    Private Sub Rbn_Incasso_SEPA_CheckedChanged(sender As Object, e As EventArgs) Handles Rbn_Incasso_SEPA.CheckedChanged, Rbn_Incasso_journal.CheckedChanged, Rbn_Incasso_SEPA.Click, Rbn_Incasso_journal.Click
        If TC_Main.SelectedIndex <> 2 Then Exit Sub
        If Lbl_Incasso_Status.Text = "New" Or Lbl_Incasso_Status.Text = "Open" Then

            If Rbn_Incasso_SEPA.Checked Then
                Prepare_Datagridview(Dgv_Incasso, Create_Incasso(Dtp_Incasso_start.Value.ToString("yyyy-MM-dd")), {"TZ205", "NG080", "TZ160", "TZ080", "TZ090", "DZ090"})
            ElseIf Rbn_Incasso_journal.Checked Then
                Prepare_Datagridview(Dgv_Incasso, Create_Incasso_Bookings(Dtp_Incasso_start.Value.ToString("yyyy-MM-dd")), {"TZ210", "TZ210", "TZ070", "TZ070", "NG080", "NG080", "HZ080", "HZ080"})
            End If
        Else
            If Rbn_Incasso_SEPA.Checked Then
                Prepare_Datagridview(Dgv_Incasso, Create_Incasso_Processed(Dtp_Incasso_start.Value.ToString("yyyy-MM-dd")), {"TZ205", "NG080", "TZ160", "TZ080", "TZ090", "DZ090"})
            ElseIf Rbn_Incasso_journal.Checked Then
                Prepare_Datagridview(Dgv_Incasso, Create_Incasso_Bookings_Processed(Dtp_Incasso_start.Value.ToString("yyyy-MM-dd")), {"TZ210", "TZ210", "TZ070", "TZ070", "NG080", "NG080", "HZ080", "HZ080"})
            End If
        End If

    End Sub

    Private Sub Dtp_Excasso_Start_ValueChanged(sender As Object, e As EventArgs)
        Dtp_Excasso_Start.MaxDate = Date.Today
    End Sub

    Private Sub Dgv_Excasso2_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles Dgv_Excasso2.CellEndEdit
        Edit_Dgv_Excasso()
    End Sub


    Private Sub Dgv_Excasso2_DataError(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) _
    Handles Dgv_Excasso2.DataError ', Dgv_Bank_Account.DataError

        MsgBox("Datafout: ongeldige invoer")
        e.ThrowException = False

    End Sub


    Private Sub GroupBox5_Leave(sender As Object, e As EventArgs)
        If IsNumeric(Tbx_Excasso_Exchange_rate.Text) Then

        Else
            MsgBox("Ongeldige inhoud")
        End If
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
            RunSQL("DELETE FROM journal WHERE name ilike '%" & Me.Cmx_Excasso_Select.DisplayMember & "'", "NULL", "Delete_Excasso_Job")
            Fill_Cmx_Excasso_Select_Combined()
            Empty_Excasso_Window()
        End If

    End Sub


    Private Sub Btn_Excasso_Save_Click(sender As Object, e As EventArgs) Handles Btn_Excasso_Save.Click
        If Cmx_Excasso_Select.SelectedIndex = -1 Then Exit Sub
        Save_Excasso_job()
    End Sub
    Private Sub Tbx_Excasso_Norm1_Enter(sender As Object, e As EventArgs)
        'Btn_Excasso_CP_Calculate.Enabled = True
    End Sub

    Private Sub Tbx_10_Account__b_jan_Leave(sender As Object, e As EventArgs) Handles _
            Tbx_10_Account__b_jan.Leave, Tbx_10_Account__b_feb.Leave, Tbx_10_Account__b_mar.Leave,
            Tbx_10_Account__b_apr.Leave, Tbx_10_Account__b_may.Leave, Tbx_10_Account__b_jun.Leave,
            Tbx_10_Account__b_jul.Leave, Tbx_10_Account__b_aug.Leave, Tbx_10_Account__b_sep.Leave,
            Tbx_10_Account__b_oct.Leave, Tbx_10_Account__b_nov.Leave, Tbx_10_Account__b_dec.Leave
        Calculate_Manual_Budgets()
    End Sub

    Private Sub Tbx_Journal_Source_Amt_TextChanged(sender As Object, e As EventArgs) Handles Tbx_Journal_Source_Amt.TextChanged
        Dim s As Decimal = Tbx2Dec(Me.Tbx_Journal_Source_Amt.Text)
        Dim m As Decimal = Tbx2Dec(Me.Lbl_Journal_Source_Saldo.Text)
        If (s <= 0 Or s > m) And (Tbx2Dec(Lbl_Journal_Source_Saldo.Text) <> 0) Then
            MsgBox("Bedrag (" & s & ") moet groter zijn dan nul en kleiner dan het saldo van de bronaccount (" & m & ")")
            Tbx_Journal_Source_Amt.Text = Tbx2Dec(m)
            Lbl_Journal_Source_Restamt.Text = Tbx_Journal_Source_Amt.Text
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

    Sub Btn_Account_Budget_Id_Click(sender As Object, e As EventArgs)
        Calculate_Budget(Lbl_00_pkid.Text)
        Select_Obj2("Btn_Account_Budget_Id_Click")
    End Sub

    Private Sub Btn_Account_Budget_All_Click(sender As Object, e As EventArgs)
        Calculate_Budget("")
        Select_Obj2("Btn_Account_Budget_All_Click")
    End Sub


    Sub Load_Excasso_Form()
        If Cmx_Excasso_Select.SelectedIndex = -1 Then Exit Sub

        'check of de budgetbedragen nog geldig zijn -----------------------------------------------------------
        If QuerySQL("select extract (year from min(date)) from journal") < Now.Year Then Calculate_Budget("")

        If Strings.Left(Cmx_Excasso_Select.Text, 5) = "Nieuw" Then


        Else  '=============================existing excasso===============================
            Load_Existing_Excasso()

        End If

    End Sub



    Private Sub Button8_Click(sender As Object, e As EventArgs)
        Process.Start("https://www.xe.com/currencyconverter/convert/?Amount=1&From=EUR&To=MDL")

    End Sub

    Private Sub Tbx_Excasso_Exchange_rate_Leave(sender As Object, e As EventArgs)
        My.Settings._exrate = Tbx2Dec(Tbx_Excasso_Exchange_rate.Text)
    End Sub


    Private Sub Tbx_Bank_Description_Leave(sender As Object, e As EventArgs) Handles Tbx_Bank_Extra_Info.Leave
        Exit Sub

        Dim SQLstr = "UPDATE bank SET description='" & Tbx_Bank_Description.Text &
               "' WHERE id='" & Dgv_Bank.SelectedCells(0).Value & "'"
        RunSQL(SQLstr, "NULL", "Tbx_Bank_Description.Leave")
        Dgv_Bank.SelectedCells(3).Value = Tbx_Bank_Description.Text

        If Me.Dgv_Bank.RowCount > 0 Then Me.Dgv_Bank.Rows(Dgv_Bank.SelectedCells(3).RowIndex).Selected = True

    End Sub

    Private Sub Btn_Excasso_Copy_to_clipboard_Click(sender As Object, e As EventArgs)

        If Strings.Left(Cmx_Excasso_Select.DisplayMember, 5) = "Nieuw" Then
            MsgBox("Bewaar deze uitkeringslijst eerst s.v.p.")
        Else
            If IsDBNull(Cmx_Excasso_Select.DisplayMember) Or Cmx_Excasso_Select.DisplayMember = "" Then Exit Sub
            Clipboard.Clear()
            Clipboard.SetText(Cmx_Excasso_Select.DisplayMember)

        End If

    End Sub


    Private Sub Cmx_00_contract__fk_relation_id_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Cmx_00_contract_fk_relation_id.SelectedIndexChanged
        If Not Add_Mode Then Exit Sub
        Exit Sub
        ''' Dit is specifieke functionaliteit voor interne contracten
        MsgBox("Lbl_00_contract__fk_relation_id.Text")
        Dim int = QuerySQL($"
                                        SELECT ba.id
                                        FROM relation r
                                        LEFT join bankacc ba ON ba.accountno = r.iban 
                                        WHERE r.id ={Lbl_11_contract__fk_relation_id.Text}
                                        ")

        MsgBox("SelectedIndexChanged2")

        Me.Lbl_Contract_Bronaccount.Visible = Not IsDBNull(int)
        Me.Cmx_Contract_fk_account_id.Visible = Not IsDBNull(int)
        Chx_00_contract__autcol.Enabled = IsDBNull(int)
        'Lbl_00_contract__fk_relation_id.Text = Cmx_00_contract_fk_relation_id.SelectedValue
    End Sub


    Private Sub Cbx_00_cp__active_Click(sender As Object, e As EventArgs) Handles Cbx_00_cp__active.Click
        CheckActive(Cbx_00_cp__active, Lbl_CP_pkid, "target")
    End Sub

    Sub Format_dvg_bank()

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


    Private Sub Dgv_Excasso2_Click(sender As Object, e As EventArgs) Handles Dgv_Excasso2.Click
        If Dgv_Excasso2.CurrentCell.ColumnIndex <> 1 Then
            'MsgBox(Dgv_Excasso2.CurrentCell.ColumnIndex)
            'Exit Sub
        End If

        Dim i As Integer = Me.Dgv_Excasso2.CurrentRow.Index

        'Dim name As String = Me.Dgv_Excasso2.Rows(i).Cells(1).Value
        Dim id = Me.Dgv_Excasso2.Rows(i).Cells(0).Value



        Dim sql As String = $"
         Select j.date As dat, j.name As Journaalnaam, amt1 As Bedr, j.description As Omschrijving from journal j
         where j.fk_account = '{id}'
         order by j.date desc, abs(amt1::decimal) 
"
        Prepare_Datagridview(Dgv_Uitkering_Account_Details, sql, {"FZ048", "TZ190", "NZ056", "TZ250"})

        With Dgv_Uitkering_Account_Details

            For Each row As DataGridViewRow In .Rows
                Dim cellValue As Object = row.Cells(2).Value ' Assuming you want to check column 1
                If IsNumeric(cellValue) Then
                    Dim value As Double = Convert.ToDouble(cellValue)
                    If value < 0 Then
                        row.Cells(1).Style.ForeColor = Color.DarkRed
                        row.Cells(2).Style.ForeColor = Color.DarkRed
                    ElseIf value > 0 Then
                        row.Cells(1).Style.ForeColor = Color.Green
                        row.Cells(2).Style.ForeColor = Color.Green
                    Else
                        row.Cells(1).Style.ForeColor = Color.Black ' Default color for zero
                        row.Cells(2).Style.ForeColor = Color.Black
                    End If
                End If
            Next

        End With

    End Sub



    Private Sub ToolStripTextBox1_TextChanged(sender As Object, e As EventArgs) Handles Searchbox2.TextChanged
        Select Case TC_Main.SelectedIndex
            Case 0
                If TC_Object.SelectedIndex = 0 Then
                    ' --- NEW CODE FOR CONTRACTS ---
                    RefreshContractList()
                Else
                    ' --- OLD CODE ---
                    Load_Table()
                End If
            Case 1
                If Dgv_Bank.DataSource IsNot Nothing Then
                    ApplyFilter(Dgv_Bank.DataSource)
                    Format_dvg_bank()
                End If
            Case 4
                Select Case TC_Boeking.SelectedIndex
                    Case 1
                        Fill_Cmx_Journal_List()
                End Select
            Case 5
                If Dgv_Rapportage_Overzicht.DataSource IsNot Nothing Then
                    ApplyFilter(Dgv_Rapportage_Overzicht.DataSource)
                    Prepare_Datagridview(Dgv_Rapportage_Overzicht, Nothing, LbL_Formatting.Text.Split(","c))
                End If
            Case 6 'Beheer
                'MsgBox(TC_Management.SelectedTab.Name)
                If TC_Management.SelectedTab.Name = "TP_Accounts" Then
                    LoadAccountTree()

                End If
        End Select
    End Sub




    Sub ApplyFilter(ByVal dt As DataTable)
        If String.IsNullOrWhiteSpace(Searchbox2.Text) Then
            dt.DefaultView.RowFilter = "" ' Clear filter if search box is empty
            Return
        End If

        ' Split search terms by spaces
        Dim searchTerms As String() = Searchbox2.Text.Split(New Char() {" "c}, StringSplitOptions.RemoveEmptyEntries)

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
        Searchbox2.Text = ""
    End Sub

    Private Sub Lv_Journal_List_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Lv_Journal_List.SelectedIndexChanged
        Fill_Journal_List_journaalposten()
    End Sub

    Private Sub Cbx_LifeCycle_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Cbx_LifeCycle2.SelectedIndexChanged
        Select Case TC_Main.SelectedIndex
            Case 0
                Try
                    MenuDelete.Enabled = (Cbx_LifeCycle2.Text = "Inactief") Or Dtp_31_contract__startdate.Value > Date.Today
                    If TC_Object.SelectedIndex = 0 Then
                        RefreshContractList()
                    Else
                        Load_Table()
                    End If
                Catch ex As Exception
                End Try
            Case 1
            'Fill_bank_transactions()
            Case 4
                Fill_Cmx_Journal_List()
        End Select
    End Sub


    Private Sub MenuSave_Click(sender As Object, e As EventArgs) Handles MenuSave.Click
        Dim saveSuccess As Boolean = True ' Default to true for the other tabs so their behavior stays identical

        Select Case TC_Main.SelectedIndex
            Case 0 ' Basisadministratie
                If TC_Object.SelectedIndex = 0 Then
                    SaveContract()
                Else
                    Basis_Save()
                End If
                Lbx_Basis.Enabled = True

            Case 1 ' Bank
                Save_Banktransaction_Accounts()
                MustWarn = True
                Dgv_Bank.Enabled = True

            Case 2 ' Incasso
                Create_Incasso_Journals()
                Create_SEPA_XML()
                Populate_Cmx_Incasso_IncassoForm()
                Me.Lbl_Incasso_Status.Text = "Open"
                Menu_Print.Enabled = True

            Case 3 ' Uitkering
                If Cmx_Excasso_Select.SelectedIndex = -1 Then Exit Sub
                Save_Excasso_job()
                MustWarn = True

            Case 4 ' Boekingen
                Select Case TC_Boeking.SelectedIndex
                    Case 0
                        Save_Internal_Booking()
                    Case 1
                        Save_modified_journaalposts()
                End Select
                Load_Cmx_Bank_Account()

            Case 6 ' Beheer (Instellingen)
                If TC_Management.SelectedTab.Name = "TP_Accounts" Then
                    Dim nodeName As String = ""
                    If AccountTree.SelectedNode IsNot Nothing Then nodeName = AccountTree.SelectedNode.Name

                    If nodeName = "AccountType" Then
                        saveSuccess = SaveCurrentAccountGroup()

                    ElseIf nodeName = "AccountGroup" Then
                        If Add_Mode Then
                            saveSuccess = SaveCurrentAccount()
                        Else
                            saveSuccess = SaveCurrentAccountGroup()
                        End If

                    ElseIf nodeName = "Account" Then
                        saveSuccess = SaveCurrentAccount()
                    End If
                End If
        End Select

        ' ---> FIX: Only disable the buttons if the save procedure returned True
        If saveSuccess Then
            Enable_Buttons(False, True)
        End If
    End Sub



    Sub Leeg_overboeking_scherm()
        If TC_Boeking.SelectedIndex = 0 Then
            Lbl_Journal_Source_Saldo.Text = 0
            Lbl_Journal_Source_Name.Text = ""
            Tbx_Journal_Source_Amt.Text = 0
            Dgv_Journal_Intern.Rows.Clear()
            Lbl_Journal_Source_Restamt.Text = 0
            Cmbx_Overboeking_Bron.SelectedIndex = -1
            Cmbx_Overboeking_Target.SelectedIndex = -1
            Tbx_Journal_Description.Text = ""
            Dtp_Journal_intern.Value = Date.Today
            Tbx_Journal_Name.Text = ""
            Rbn_Journal_Extra.Checked = False
            Rbn_Journal_Intern.Checked = False
            Rbn_Journal_Contract.Checked = False
        End If
    End Sub


    Private Sub MenuAdd_Click(sender As Object, e As EventArgs) Handles MenuAdd.Click
        Select Case TC_Main.SelectedIndex
            Case 0
                If TC_Object.SelectedIndex = 0 Then
                    ' --- NEW CODE FOR CONTRACTS ---
                    AddContractUI()
                Else
                    ' --- OLD CODE ---
                    Basis_Add()
                End If
            Case 6
                Select Case TC_Management.SelectedTab.Name
                    Case "TP_Accounts"
                        Dim selectedNode = AccountTree.SelectedNode
                        If selectedNode Is Nothing Then Exit Sub

                        Add_Mode = True

                        If selectedNode.Name = "AccountType" Then
                            ' --- Prepare to add a new AccountGroup ---
                            Grbx_Beheer_Accountgroep.Enabled = True
                            Grbx_Beheer_Account.Enabled = False

                            Tbx_Beheer_Accountgroepnaam.Text = ""
                            Tbx_Beheer_Accgroup_Description.Text = ""
                            Cmbox_Beheer_Accgroup_Subtype.SelectedIndex = -1
                            Cmbox_Beheer_Accgroup_Subtype.Text = ""

                            Chbx_Beheer_Accgroup_Active.Checked = True
                            Chbx_Beheer_Accgroup_Active.Enabled = True
                            Lbl_Beheer_Accgroup_posts.Text = "0"

                            ' Crucial: Set ID to 0 so the Repository executes an INSERT
                            Lbl_Beheer_Accgroup_id.Text = "0"

                            ' Set radio buttons based on the selected AccountType node text
                            Dim accType As String = selectedNode.Text
                            Rbtn_Beheer_Accounttype1.Checked = (accType = "Inkomsten")
                            Rbtn_Beheer_Accounttype2.Checked = (accType = "Uitgaven")
                            Rbtn_Beheer_Accounttype3.Checked = (accType = "Transit")

                            Tbx_Beheer_Accountgroepnaam.Focus()
                        ElseIf selectedNode.Name = "AccountGroup" Then
                            ' --- Prepare to add a new Account ---

                            ' ---> FIX 1: Lock the AccountGroup container and the Combobox
                            Grbx_Beheer_Accountgroep.Enabled = False
                            Cmbx_Beheer_Accgroup.Enabled = False

                            Grbx_Beheer_Account.Enabled = True

                            Clear_Account()

                            Lbl_Beheer_Account_id.Text = "0"
                            Lbl_Beheer_Account_posts.Text = "0"
                            Chbx_Beheer_Account_Active.Checked = True
                            Chbx_Beheer_Account_Active.Enabled = True

                            ' Match the ID to bypass strict type-casting failures
                            Dim targetGroupId As String = selectedNode.Tag.ToString()
                            For i As Integer = 0 To Cmbx_Beheer_Accgroup.Items.Count - 1
                                Dim rowView As DataRowView = TryCast(Cmbx_Beheer_Accgroup.Items(i), DataRowView)
                                If rowView IsNot Nothing AndAlso rowView("id").ToString() = targetGroupId Then
                                    Cmbx_Beheer_Accgroup.SelectedIndex = i
                                    Exit For
                                End If
                            Next

                            Tbx_Beheer_Accountbron.Text = "cat"
                            Tbx_Beheer_Accountnaam.Focus()
                        End If
                End Select
        End Select
                Enable_Buttons(True, False)
        Lbx_Basis.Enabled = False
    End Sub

    Private Sub MenuCancel_Click(sender As Object, e As EventArgs) Handles MenuCancel.Click
        isCanceling = True
        Select Case TC_Main.SelectedIndex
            Case 0
                Cancel()
            Case 1
                'Dgv_Bank.Click()
                Fill_bank_transactions("MenuCancel_Click", Me.Dgv_Bank.SelectedCells(3).RowIndex)
                Fill_Journals_by_bank(Me.Dgv_Bank.SelectedCells(0).Value)
                SelectRowById(Dgv_Bank, Dgv_Bank.SelectedCells(0).Value)

            Case 3
                If Cmx_Excasso_Select.SelectedIndex = -1 Then Exit Sub
                Load_Excasso_Form()
            Case 4
                Leeg_overboeking_scherm()
            Case 6
                Load_Account_Settings()
                If TC_Management.SelectedTab.Name = "TP_Accounts" Then
                    Add_Mode = False

                    ' Re-fire the AfterSelect event to restore the original data of the selected node
                    If AccountTree.SelectedNode IsNot Nothing Then
                        Dim args As New TreeViewEventArgs(AccountTree.SelectedNode, TreeViewAction.Unknown)
                        AccountTree_AfterSelect(AccountTree, args)
                    End If
                End If
            Case 7

        End Select
        Enable_Buttons(False, True)
        Lbx_Basis.Enabled = True
        Dgv_Bank.Enabled = True

        isCanceling = False
    End Sub

    Private Sub MenuDelete_Click(sender As Object, e As EventArgs) Handles MenuDelete.Click


        Select Case TC_Main.SelectedIndex
            Case 0
                If TC_Object.SelectedIndex = 0 Then
                    ' --- NEW CODE FOR CONTRACTS ---
                    DeleteContract()
                Else
                    ' --- OLD CODE ---
                    Basis_Delete()
                End If
            Case 2

                Dim Incasso2Delete As String = $"Delete From Journal where name ilike '%{Lbl_Incasso_job_name.Text}%' and source='Incasso'"
                'MsgBox(Incasso2Delete)
                'Clipboard.SetText(Incasso2Delete)
                RunSQL(Incasso2Delete, "NULL", "Btn_Incasso_Delete_Click")
                Populate_Cmx_Incasso_IncassoForm()
                Me.Lbl_Incasso_Status.Text = "Nieuw"
                Menu_Print.Enabled = False
                Me.Lbl_Incasso_Error.Visible = False
            Case 3
                MenuExcassoDelete()
        End Select
        Enable_Buttons(False, True)
    End Sub

    Private Sub MenuBanktransactie_Click(sender As Object, e As EventArgs) Handles MenuBanktransactie.Click
        Download_Bank_Transactions()
    End Sub

    Private Sub MenuUploadAlles_Click(sender As Object, e As EventArgs) Handles MenuUploadAlles.Click
        Load_Bank_csv_from_folder()
    End Sub

    Private Sub MenuCategoriseer_Click(sender As Object, e As EventArgs) Handles MenuCategoriseer.Click
        Categorize_Bank_Transactions(True, True, True, True, True, True, True)
        Fill_bank_transactions("MenuCategoriseer", Me.Dgv_Bank.SelectedCells(3).RowIndex)
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

    Private Sub TC_Main_Click(sender As Object, e As EventArgs) Handles TC_Main.SelectedIndexChanged
        'MsgBox($" begin TC_Main.Selectedindechanged: {isManualChange}")
        Enable_Buttons(False, True)
        Select Case TC_Main.SelectedIndex

            Case 0
                isManualChange = True
                If Searchbox2.Text <> "" Then Load_Table()

            Case 1  'bank
                isManualChange = True
                Searchbox2.Text = ""


                'only load the bank data if datagridview is still empty
                If Dgv_Bank.Rows.Count = 0 Or Dgv_Bank.DataSource Is Nothing Then
                    If Me.Dgv_Mgnt_Tables.Rows(1).Cells(1).Value > 0 Then
                        Fill_bank_transactions("Cmx_Bank_bankacc.SelectedIndexChanged", 0)
                    End If
                End If

                Enable_Buttons(False, False)
            Case 2 'incasso
                isManualChange = False
                Populate_Cmx_Incasso_IncassoForm()
                Cmx_Incasso_IncassoForm.SelectedIndex = -1
                If Me.Dgv_Mgnt_Tables.Rows(3).Cells(1).Value > 0 And
                    Me.Dgv_Mgnt_Tables.Rows(5).Cells(1).Value > 0 And
                    Me.Dgv_Mgnt_Tables.Rows(8).Cells(1).Value > 0 Then

                    Prepare_Datagridview(Dgv_Incasso, Nothing, {"TZ205", "NG080", "TZ160", "TZ080", "TZ090", "DZ090"})
                End If
                isManualChange = True
                Enable_Buttons(False, False)
            Case 3 'uitkering
                isManualChange = False
                'Fill_Cmx_Excasso_Select_Combined()

                Enable_Buttons((Dgv_Excasso2.Rows.Count > 0), (Dgv_Excasso2.Rows.Count = 0))

                MenuBanktransactie.Visible = False '
                MenuUploadAlles.Visible = False
                MenuBanktransactie.Visible = False
                MenuCategoriseer.Visible = False

                If Me.Dgv_Mgnt_Tables.Rows(3).Cells(1).Value > 0 And
                    Me.Dgv_Mgnt_Tables.Rows(5).Cells(1).Value > 0 And
                    Me.Dgv_Mgnt_Tables.Rows(8).Cells(1).Value > 0 And
                    Me.Cmx_Excasso_Select.DisplayMember = "" Then

                    Dtp_Excasso_Start.ShowUpDown = False
                    Dtp_Excasso_Start.Value = CDate(Date.Today.Year & "-" & Date.Today.Month & "-" & Date.Today.Day)
                    Dtp_Excasso_Start.MaxDate = CDate(Date.Today.Year & "-" & Date.Today.Month & "-" & Date.Today.Day)
                    Fill_Cmx_Excasso_Select_Combined()
                End If
                isManualChange = True
            Case 4
                isManualChange = False
                Enable_Buttons(False, False)
                Menu_Export.Enabled = True


                Me.Dtp_Journal_intern.Value = CDate(Date.Today.Year & "-" & Date.Today.Month & "-" & Date.Today.Day)
                Dim sql As String = "update journal j set name = 
                (select left(replace(replace(replace(replace(replace(replace(replace(b.name,' van der',''),' van de',''),'Hr ',''),'Mw ',''),' de ',''),' van ',''),'.',''),14) 
                from bank b where b.id = j.fk_bank)||'/'||(select a.name from account a where a.id = j.fk_account)
                where name='nog te bepalen' and fk_account != (select value::integer from settings where label='nocat') and source = 'Bank'"
                RunSQL(sql, "NULL", "TC_Main_Click")
                isManualChange = True
            Case 5
                isManualChange = False
                'Enable_Buttons(False, False)

            Case 6
                isManualChange = False
                Load_Account_Settings()


                Select Case TC_Management.SelectedTab.Name
                    Case "TP_Accounts"
                        LoadAccountTree()
                        MenuAdd.Enabled = True
                        ExpandSpecificAccountType("Inkomsten")
                        'ExpandOnlyLevel1()

                End Select


                isManualChange = True
                    Case 7
                        isManualChange = False
                        Populate_Combobox(Cmbx_Tussenrekening, "select * from account where source = 'cat'  and type = 'Anders' and name not in ('[Niet toegewezen]','Euro tegenwaarde', 'Overhead') and name not ilike 'Bank%'")
                        Cmbx_Tussenrekening.SelectedIndex = -1
                        Prepare_Datagridview(Dgv_Tussenrekening, Fill_Afletterbox, {"TZ080", "TZ300", "NZ080", "NZ080", "NZ080"})
                        Prepare_Datagridview(Dgv_Tussenrekening_Uitk,
                     "SELECT date, name, SUM(amt1) FROM journal WHERE source = 'Uitkering' AND status = 'Open' GROUP BY date, name order by date desc",
                     {"HZ000", "TZ200", "NZ080"})
                        Lbl_Tussenrekening_3.Text = $"Openstaande uitkeringslijsten ({Dgv_Tussenrekening_Uitk.RowCount})"
                        Initialize_Tussenrekening_DatePicker()
                        isManualChange = True


                End Select

                Show_buttons()

    End Sub
    Private Sub TC_Boeking_Click(sender As Object, e As EventArgs) Handles TC_Boeking.Click

        Enable_Buttons(False, False)
        Menu_Export.Enabled = True

        Searchbox2.Text = ""
        If Lbl_Journal_Source_Name.Text = "" Then
            Tbx_Journal_Name.Text = ""
            Rbn_Journal_Intern.Checked = True
        End If
        Select Case TC_Boeking.SelectedIndex
            Case 2
                Report_Closing()
        End Select
        Show_buttons()

    End Sub



    Private Sub Menu_Export_Click_1(sender As Object, e As EventArgs) Handles Menu_Export.Click
        Export2Excel()
    End Sub


    Sub Empty_Excasso_Window()

        Dgv_Excasso2.DataSource = Nothing
        Dgv_Excasso2.Rows.Clear()
        Me.Dgv_Excasso2.Columns.Clear()

    End Sub

    Private Sub Button1_Click_2(sender As Object, e As EventArgs)
        MsgBox("Dit is een handmatige activiteit die door de databasebeheerder moet worden uitgevoerd")
    End Sub

    Private Sub Btn_Report_YearEnd_Post_Click(sender As Object, e As EventArgs) Handles Btn_Report_YearEnd_Post.Click
        Close_Year()
    End Sub

    Private Sub Tbx_Bank_Description_TextChanged(sender As Object, e As EventArgs)
        Dgv_Bank.SelectedCells(3).Value = Tbx_Bank_Description.Text
    End Sub
    Private Sub Tbx_01_Accgroup__type_TextChanged(sender As Object, e As EventArgs) Handles Tbx_01_Accgroup__type.TextChanged
        Rbtn_accgroup_Income.Checked = Strings.Trim(Tbx_01_Accgroup__type.Text) = "Inkomsten"
        Rbtn_accgroup_expense.Checked = Strings.Trim(Tbx_01_Accgroup__type.Text) = "Uitgaven"
        Rbtn_accgroup_transit.Checked = Strings.Trim(Tbx_01_Accgroup__type.Text) = "Transit"
        '@@@ hard value vervangen door tt_type.Text
    End Sub

    Private Sub Rbtn_accgroup_Income_CheckedChanged(sender As Object, e As EventArgs) Handles Rbtn_accgroup_Income.Click
        If MenuSave.Enabled Then Tbx_01_Accgroup__type.Text = Rbtn_accgroup_Income.Text
    End Sub

    Private Sub Rbtn_accgroup_expense_CheckedChanged(sender As Object, e As EventArgs) Handles Rbtn_accgroup_expense.Click
        If MenuSave.Enabled Then Tbx_01_Accgroup__type.Text = Rbtn_accgroup_expense.Text
    End Sub

    Private Sub Rbtn_accgroup_transit_CheckedChanged(sender As Object, e As EventArgs) Handles Rbtn_accgroup_transit.Click
        If MenuSave.Enabled Then Tbx_01_Accgroup__type.Text = Rbtn_accgroup_transit.Text
    End Sub


    Sub Prepare_Datagridview(dgv As DataGridView, sql As String, arr As Array)

        If sql IsNot Nothing Then dgv.DataSource = Collect_data2(sql)

        'formatarray
        '[datatypeletter(1)][kleurletter(1)][kolombreedte(3)]'
        '*** datatype ***
        'T = Standaardformaat
        'N = Numeriek / 2 cijfers achter de komma
        'I = Integer /Rechts uitgelijnd
        'J = Integer /Rechts gecentreerd
        'H = Verberg kolom
        'D = Datum

        '*** kleur ***
        'Z = Black
        'B = Blue = Editable
        'G = Green
        'R = DarkRed

        For i As Integer = 0 To Math.Min(arr.Length, dgv.Columns.Count) - 1

            Dim formatStr As String = arr(i).ToString()

            'column formatting
            If formatStr.Length >= 5 Then

                ' Extract components from format string
                Dim dataType As Char = formatStr(0) ' First character (Data Type)
                Dim colorCode As Char = formatStr(1) ' Second character (Color)
                Dim columnWidth As Integer

                ' Extract and convert column width (last 3 characters)
                If Integer.TryParse(formatStr.Substring(2, 3), columnWidth) Then dgv.Columns(i).Width = columnWidth

                ' Set column data type 
                Try
                    dgv.Columns(i).HeaderText = Strings.Left(dgv.Columns(i).HeaderText, 1).ToUpper & Strings.Mid(dgv.Columns(i).HeaderText, 2).ToLower
                Catch
                End Try

                Select Case dataType
                    Case "D"c : dgv.Columns(i).DefaultCellStyle.Format = "dd-MM-yyyy"
                    Case "E"c : dgv.Columns(i).DefaultCellStyle.Format = "MM-yy"
                    Case "F"c : dgv.Columns(i).DefaultCellStyle.Format = "dd-MM"
                    Case "N"c
                        dgv.Columns(i).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dgv.Columns(i).DefaultCellStyle.Format = "N2"
                    Case "I"c
                        dgv.Columns(i).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        dgv.Columns(i).DefaultCellStyle.Format = "N0"
                    Case "J"c
                        dgv.Columns(i).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                        dgv.Columns(i).DefaultCellStyle.Format = "N0"
                    Case "H"c : dgv.Columns(i).Visible = False
                    Case Else : dgv.Columns(i).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                End Select

                Select Case colorCode
                    Case "R"c : dgv.Columns(i).DefaultCellStyle.ForeColor = Color.DarkRed
                    Case "G"c : dgv.Columns(i).DefaultCellStyle.ForeColor = Color.Green
                    Case "B"c
                        dgv.Columns(i).DefaultCellStyle.ForeColor = Color.Blue
                        dgv.Columns(i).ReadOnly = False
                End Select

            End If

            'row formatting
            For r As Integer = 0 To dgv.Rows.Count - 1
                ' (1) Set row background color based on first column text
                If dgv.Columns.Count > 0 Then
                    For c = 0 To dgv.Columns.Count - 1
                        Dim firstColValue As String = dgv.Rows(r).Cells(c).Value?.ToString()

                        If Not String.IsNullOrEmpty(firstColValue) Then
                            If firstColValue.Contains("Total") Then : dgv.Rows(r).DefaultCellStyle.BackColor = Color.Khaki
                            ElseIf firstColValue.Contains("🚨") Then : dgv.Rows(r).DefaultCellStyle.BackColor = Color.LightSalmon
                            ElseIf firstColValue.Contains("⚠️") Then : dgv.Rows(r).DefaultCellStyle.BackColor = Color.LightYellow
                            ElseIf firstColValue.Contains("#") Then : dgv.Rows(r).DefaultCellStyle.BackColor = Color.DarkSeaGreen
                            ElseIf firstColValue.Contains("Afschrift") Then : dgv.Rows(r).DefaultCellStyle.BackColor = Color.DarkSeaGreen
                            ElseIf firstColValue.Contains("(Excasso)") Then : dgv.Rows(r).DefaultCellStyle.BackColor = Color.White
                            ElseIf firstColValue.Contains("Tussenrekening") Then : dgv.Rows(r).DefaultCellStyle.BackColor = Color.Gainsboro
                            ElseIf firstColValue.Contains("generaal") Then
                                dgv.Rows(r).DefaultCellStyle.BackColor = Color.DarkSeaGreen
                                dgv.Rows(r).DefaultCellStyle.ForeColor = Color.Blue
                            ElseIf firstColValue.Contains("vergelijking") Then
                                dgv.Rows(r).DefaultCellStyle.BackColor = Color.DarkSeaGreen
                                dgv.Rows(r).DefaultCellStyle.ForeColor = Color.Blue
                            End If
                        End If
                    Next c
                End If

                ' (2) Zero Value Coloring (for numeric columns)

                If IsDBNull(dgv.Rows(r).Cells(i).Value) Then dgv.Rows(r).Cells(i).Value = 0
                If dgv.Rows(r).Cells(i).Value IsNot Nothing Then
                    If dgv.Rows(r).Cells(i).Value.ToString = "0,00" Or dgv.Rows(r).Cells(i).Value.ToString = "0" Then
                        dgv.Rows(r).Cells(i).Style.ForeColor = Color.LightGray 'dgv.Rows(r).DefaultCellStyle.BackColor ' Make it blend
                    Else
                        dgv.Rows(r).Cells(i).Style.ForeColor = dgv.Columns(i).DefaultCellStyle.ForeColor
                    End If
                End If
            Next
        Next
    End Sub


    Private Sub Rbn_Bank_jtype_con_CheckedChanged(sender As Object, e As EventArgs) Handles Rbn_Bank_jtype_con.CheckedChanged, Rbn_Bank_jtype_ext.CheckedChanged, Rbn_Bank_jtype_int.CheckedChanged

        Btn_Bank_Add_Journal.Enabled = True
    End Sub
    Private Sub Dgv_Bank_Sorted(sender As Object, e As EventArgs) Handles Dgv_Bank.Sorted
        Format_dvg_bank()
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

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim p1 = InputBox("maand:")
        Dim sql = QuerySQL("Select sql from query where category ilike 'Transaction' and name='Verwijder maand'")
        sql = sql.Replace("p1", p1)
        ToClipboard(sql, True)
        RunSQL(sql, "NULL", "Testbutton verwijder maand")
        Fill_bank_transactions("Button1", 0)
    End Sub


    Private Sub Tbx_Extra_Info_TextChanged(sender As Object, e As EventArgs) Handles Tbx_Bank_Extra_Info.TextChanged
        Dim des As String = Tbx_Bank_Description.Text
        If Chbx_Bank_ExtraInfo_voor.Checked Then
            If Strings.InStr(des, " | ") = 0 And Tbx_Bank_Extra_Info.Text <> "" Then des = " | " & des
            Try
                des = Tbx_Bank_Extra_Info.Text & Strings.Mid(des, Strings.InStr(des, " | "))
            Catch
            End Try
        End If
        If Tbx_Bank_Extra_Info.Text = "" And Strings.InStr(des, " | ") > 0 Then des = Mid(des, Strings.InStr(des, " | ") + 3)
        Tbx_Bank_Description.Text = des
    End Sub

    Private Sub Menu_Help_Click(sender As Object, e As EventArgs) Handles Menu_Help.Click
        Select Case TC_Main.SelectedIndex
            Case 0 : Process.Start("https://github.com/Erthengs/SPAS2025/wiki/Stappenplan-(Maandelijks)")
            Case 1 : Process.Start("https://github.com/Erthengs/SPAS2025/wiki/Bank")
            Case 2 : Process.Start("https://github.com/Erthengs/SPAS2025/wiki/Incasso")
            Case 3 : Process.Start("https://github.com/Erthengs/SPAS2025/wiki/Uitkering")
            Case 5 : Process.Start("https://github.com/Erthengs/SPAS2025/wiki/SPAS:-Inleiding")
            Case 6 : Process.Start("https://github.com/Erthengs/SPAS2025/wiki/SPAS:-Inleiding")
            Case 7 : Process.Start("https://github.com/Erthengs/SPAS2025/wiki/Tussenrekening")

        End Select




        Process.Start("https://github.com/Erthengs/SPAS2025/wiki/Stappenplan-(Maandelijks)")
    End Sub
    '========================================================================================================
    '======                                                                                            ======
    '======                                B O E K I N G E N                                           ======
    '======                                                                                            ======
    '========================================================================================================

    Sub Lv_Journal_List_Click(sender As Object, e As EventArgs) Handles Lv_Journal_List.Click

        Try
            Dim selectedItem As ListViewItem = Lv_Journal_List.SelectedItems(0)
            Dim journaaldata = Collect_data2(Create_Journal_SQL)

            Me.Lbl_Journaalposten_datum.Text = IIf(IsDBNull(journaaldata.Rows(0)(0)), "", journaaldata.Rows(0)(0))
            Me.Lbl_Journaalposten_header.Text = IIf(IsDBNull(journaaldata.Rows(0)(1)), "", journaaldata.Rows(0)(1))
            'Me.Tbx_journaalposten_omschr.Text = IIf(IsDBNull(journaaldata.Rows(0)(4)), "", journaaldata.Rows(0)(4))
            Me.Lbl_journaalposten_status.Text = IIf(IsDBNull(journaaldata.Rows(0)(6)), "", journaaldata.Rows(0)(6))
            Me.Lbl_Journaalposten_bron.Text = IIf(IsDBNull(journaaldata.Rows(0)(7)), "", journaaldata.Rows(0)(7))
            Me.Lbl_journaalposten_iban.Text = IIf(IsDBNull(journaaldata.Rows(0)(8)), "", journaaldata.Rows(0)(8))
            Me.Lbl_journaalposten_type.Text = IIf(IsDBNull(journaaldata.Rows(0)(9)), "", journaaldata.Rows(0)(9))
            Me.Lbl_journaalposten_cpinfo.Text = IIf(IsDBNull(journaaldata.Rows(0)(14)), "", journaaldata.Rows(0)(14))
            Me.Lbl_journaalposten_wisselkoers.Text = IIf(IsDBNull(journaaldata.Rows(0)(15)), "", journaaldata.Rows(0)(15))
            Me.Banklink.Text = IIf(IsDBNull(journaaldata.Rows(0)(16)), 0, journaaldata.Rows(0)(16).ToString)               'Me.Cmbx_journaalposten_relatie.SelectedIndex = -1

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

    End Sub


    '========================================================================================================
    '======                                                                                            ======
    '======                B O E K I N G E N   - J O U R N A A L P O S T E N                           ======
    '======                                                                                            ======
    '========================================================================================================

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

    End Sub

    Private Sub Cmbx_journaalposten_account_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Cmbx_journaalposten_account.SelectedIndexChanged
        If Cmbx_journaalposten_account.SelectedIndex = -1 Or
            TC_Boeking.SelectedIndex <> 1 Or Dgv_journaalposten.RowCount = 0 Then Exit Sub
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
    End Sub

    Private Sub Banklink_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles Banklink.LinkClicked
        Dim bankid As Integer = Integer.Parse(Banklink.Text)
        If bankid = 0 Or Len(bankid) = 0 Then Exit Sub
        TC_Main.SelectedIndex = 1
        Fill_bank_transactions("TC_Main.SelectedIndex", 0)
        SelectRowById(Dgv_Bank, bankid)

    End Sub

    Private Sub Overboekingen_Click(sender As Object, e As EventArgs) Handles Overboekingen.Click
        Prepare_Datagridview(Dgv_Journal_Intern, Nothing, {"HZ010", "TZ160", "NB070"})
    End Sub


    Private Sub Cmbx_Overboeking_Bron_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Cmbx_Overboeking_Bron.SelectedIndexChanged
        If TC_Boeking.SelectedTab.Text <> "Overboekingen" Or Cmbx_Overboeking_Bron.SelectedIndex = -1 Then Exit Sub
        Try
            Dim selectedItem As ComboBoxItem = TryCast(Cmbx_Overboeking_Bron.SelectedItem, ComboBoxItem)

            If selectedItem IsNot Nothing Then
                Tbx_Journal_Source_Amt.Text = selectedItem.Column3
                Calculate_Journal_Booking_Data()
            End If
        Catch
        End Try


    End Sub
    Private Sub Cmbx_Overboeking_Bron_TextChanged(sender As Object, e As EventArgs) Handles Cmbx_Overboeking_Bron.TextChanged
        If Cmbx_Overboeking_Bron.SelectedIndex = -1 Then Tbx_Journal_Source_Amt.Text = "0"
    End Sub
    Private Sub Cmbx_Overboeking_Target_Changed(sender As Object, e As EventArgs) Handles Cmbx_Overboeking_Target.SelectedIndexChanged

        If isProgrammaticChange Then Exit Sub
        If Cmbx_Overboeking_Target.SelectedIndex = -1 Then Exit Sub

        If Cmbx_Overboeking_Bron.SelectedIndex = -1 And TC_Boeking.SelectedTab.Text <> "Overboekingen" Then
            If Tbx2Dec(Lbl_Journal_Source_Restamt.Text) <= 0 Then
                MsgBox($"Selecteer eerst een bronaccount.", vbInformation)
                Exit Sub
            End If
        End If
        Dim selectedItem As ComboBoxItem = TryCast(Cmbx_Overboeking_Target.SelectedItem, ComboBoxItem)
        Dim tgt_tot As Decimal = 0

        If selectedItem IsNot Nothing Then
            If Cmbx_Overboeking_Bron.SelectedIndex > -1 Then
                'MsgBox($"1) {selectedItem.Column2}{vbCr} 2) {selectedItem.Column2}{vbCr} 3) {selectedItem.Column3}")
                With Dgv_Journal_Intern
                    .Rows.Add(selectedItem.Column1)
                    .Rows(.Rows.Count - 1).Cells(1).Value = selectedItem.Column2
                    .Rows(.Rows.Count - 1).Cells(2).Value = Tbx2Dec(Lbl_Journal_Source_Restamt.Text)

                    'calculation of rest amount
                    For i = 0 To .Rows.Count - 1
                        tgt_tot = tgt_tot + .Rows(i).Cells(2).Value
                    Next
                End With
            End If

        End If

        Prepare_Datagridview(Dgv_Journal_Intern, Nothing, {"HZ010", "TZ160", "NB070"})
        Calculate_Journal_Booking_Data()
        isProgrammaticChange = True
    End Sub

    Private Sub Cmbx_Overboeking_Target_Enter(sender As Object, e As EventArgs) Handles Cmbx_Overboeking_Target.Enter, Cmbx_Overboeking_Target.Click
        isProgrammaticChange = False
    End Sub

    Private Sub Rbtn_Overboekingen_Kind_CheckedChanged(sender As Object, e As EventArgs) Handles Rbtn_Overboekingen_Kind.CheckedChanged,
        Rbtn_Overboekingen_Oudere.CheckedChanged, Rbtn_Overboekingen_alles.CheckedChanged
        If isProgrammaticChange = True Then Exit Sub
        Dim sql As String = "select a.id,a.name,COALESCE(SUM(j.amt1), 0::money) AS total_amt1
                             from account a LEFT join     journal j ON a.id = j.fk_account
                            where a.active = TRUE GROUP by  a.id,  a.name ORDER by a.name;"
        If Rbtn_Overboekingen_Kind.Checked Then
            sql = sql.Replace("True", "True and g.name = 'Kindersponsoring'")
        ElseIf Rbtn_Overboekingen_Oudere.Checked Then
            sql = sql.Replace("True", "True and g.name = 'Ouderensponsoring'")

        End If
        isProgrammaticChange = True
        Populate_Combobox(Cmbx_Overboeking_Target, sql)

    End Sub

    Private Sub Rbtn_Overboekingen_Oudere_Click(sender As Object, e As EventArgs) Handles Rbtn_Overboekingen_Oudere.Click,
            Rbtn_Overboekingen_Oudere.Click, Rbtn_Overboekingen_alles.Click
        isProgrammaticChange = False
    End Sub

    Private Sub Btn_Boeking_Expand_Collapse_Click(sender As Object, e As EventArgs)
        If Btn_Boeking_Expand_Collapse.Text = "Alles uitklappen" Then
            AccountTree.ExpandAll()
            Btn_Boeking_Expand_Collapse.Text = "Alles inklappen"
        Else
            AccountTree.CollapseAll()
            Btn_Boeking_Expand_Collapse.Text = "Alles uitklappen"
        End If
    End Sub
    '========================================================================================================
    '======                                                                                            ======
    '======                               R A P P O R T A G E                                          ======
    '======                                                                                            ======
    '========================================================================================================

    Sub ReportTree_NodeMouseClick(sender As Object, e As TreeNodeMouseClickEventArgs) Handles ReportTree.NodeMouseClick

        Dim rep As String = ""
        report_year = Cmbx_Reporting_Year.SelectedItem

        If e.Node.Level = 1 Then
            rep = e.Node.Text
            Lbl_Rapportage.Text = rep
            Run_ReportTree(rep)
        End If

    End Sub

    Private Sub Btn_Rap_Expand_Collapse_Click(sender As Object, e As EventArgs) Handles Btn_Rap_Expand_Collapse.Click
        If Btn_Rap_Expand_Collapse.Text = "Alles uitklappen" Then
            ReportTree.ExpandAll()
            Btn_Rap_Expand_Collapse.Text = "Alles inklappen"
        Else
            ReportTree.CollapseAll()
            Btn_Rap_Expand_Collapse.Text = "Alles uitklappen"
        End If
    End Sub

    Private Sub Cmbx_Reporting_Year_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Cmbx_Reporting_Year.SelectedIndexChanged
        Try
            report_year = Cmbx_Reporting_Year.SelectedItem
            Run_ReportTree(Lbl_Rapportage.Text)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub Dgv_Report_6_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles Dgv_Report_6.CellContentClick
        Dim columnName As String = Dgv_Report_6.Columns(e.ColumnIndex).HeaderText
        If columnName = "Journaalnaam" Then

            SelectNodeByName(ReportTree, "Posten per boeking")
            With Dgv_Report_6
                Searchbox2.Text = $"{ .CurrentCell.Value} { .Rows(.CurrentCell.RowIndex).Cells(0).Value}"
            End With
        End If
    End Sub

    Private Sub dgv_1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles Dgv_Rapportage_Overzicht.CellContentClick
        Dim selectedNode As TreeNode = ReportTree.SelectedNode

        Select Case selectedNode.Text
            Case "Jaaroverzicht Bank"
                'If Dgv_Rapportage_Overzicht.CurrentCell.ColumnIndex = 1 Then
                'Else
                Drill_down_Bank_overview(Me.Dgv_Rapportage_Overzicht.CurrentCell.RowIndex, Me.Dgv_Rapportage_Overzicht.CurrentCell.ColumnIndex)
                'End If
            Case "Jaarrapportage"
                Drill_down_Report_overview(Dgv_Rapportage_Overzicht.CurrentCell.RowIndex, Dgv_Rapportage_Overzicht.CurrentCell.ColumnIndex)
            Case Else
                Dim columnName As String = Dgv_Rapportage_Overzicht.Columns(e.ColumnIndex).HeaderText
                Dim formatting As String = Nothing
                Dim arr_format() As String = Nothing
                formatting = QuerySQL($"Select formatting from query where category = 'Transaction' and name='Detail journaalposten'")
                If e.RowIndex >= 0 AndAlso e.ColumnIndex >= 0 AndAlso
                    (columnName = "Accountnaam" Or columnName = "Journaalnaam" Or columnName = "Accountgroep" Or columnName = "Relatienaam") Then
                    ' Get the column header text
                    Dim sql As String = QuerySQL("SELECT sql from query where name = 'Detail journaalposten';")

                    If Not IsNothing(LbL_Formatting.Text) Then arr_format = LbL_Formatting.Text.Split(","c)

                    Select Case columnName
                        Case "Accountnaam"
                            sql = sql.Replace("a.name like '%%'", $"a.name like '{Dgv_Rapportage_Overzicht.CurrentCell.Value}'")
                            formatting = formatting.Replace("T150", "H150")
                        Case "Journaalnaam"
                            sql = sql.Replace("j.name like '%%'", $"j.name like '{Dgv_Rapportage_Overzicht.CurrentCell.Value}'")
                            formatting = formatting.Replace("T250", "H250")
                        Case "Accountgroep"
                            sql = sql.Replace("c.name like '%%'", $"c.name like '{Dgv_Rapportage_Overzicht.CurrentCell.Value}'")
                            formatting = formatting.Replace("T149", "H149")
                        Case "Relatienaam"
                            sql = sql.Replace("concat(r.name||','||r.name_add) like'%%'", $"concat(r.name||','||r.name_add) like '{Dgv_Rapportage_Overzicht.CurrentCell.Value}'")
                            formatting = formatting.Replace("T151", "H151")
                        Case Else
                            'Do nothing
                    End Select

                    Prepare_Datagridview(Dgv_Report_6, sql, formatting.Split(","))
                    'End If

                    Lbl_Rapportage_Detail.Text = $"Details {Dgv_Rapportage_Overzicht.CurrentCell.Value}"
                End If
        End Select

    End Sub

    Private Sub Cmx_Excasso_Select_SelectedIndexChanged_3(sender As Object, e As EventArgs) Handles Cmx_Excasso_Select.SelectedIndexChanged
        Dim previousState As Boolean = isManualChange
        isManualChange = False

        If Cmx_Excasso_Select.Items.Count > 0 Then
            If Strings.Left(Cmx_Excasso_Select.Text, 5).ToString <> "Nieuw" Then
                Load_Existing_Excasso()
            Else
                Load_New_Excasso(True)
            End If
        End If

        isManualChange = previousState
    End Sub

    Function CalculateColumnSum(dgv As DataGridView, columnIndex As Integer) As Double
        Dim sum As Double = 0

        For Each row As DataGridViewRow In dgv.Rows
            ' Skip the new row (for adding new entries)
            If Not row.IsNewRow Then
                ' Check for null or empty values to avoid errors
                Dim cellValue = row.Cells(columnIndex).Value
                'Dim accountnr = row.Cells(0).Value
                If cellValue IsNot Nothing AndAlso IsNumeric(cellValue) Then

                    sum += Convert.ToDouble(cellValue)

                End If
            End If
        Next
        Return sum

    End Function

    Function CalculateColumnCount(dgv As DataGridView, columnIndex As Integer) As Double
        Dim cnt As Integer = 0

        For Each row As DataGridViewRow In dgv.Rows
            ' Skip the new row (for adding new entries)
            If Not row.IsNewRow Then
                Dim cellValue = row.Cells(columnIndex).Value
                If cellValue IsNot Nothing AndAlso IsNumeric(cellValue) AndAlso cellValue > 0 Then
                    cnt += 1
                End If
            End If
        Next
        Return cnt

    End Function






    Private Sub Dgv_Excasso_numbers_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles Dgv_Excasso_numbers.CellEndEdit

        If Not Me.isManualChange Then Exit Sub
        If Me.Dgv_Excasso2.Rows.Count = 0 Then Exit Sub
        If IsDBNull(Me.Dgv_Excasso_numbers.CurrentCell.Value) Then Exit Sub
        Dim colindex As Integer = e.ColumnIndex
        Dim rowindex As Integer = e.RowIndex
        If colindex <> 6 And colindex <> 8 Then Exit Sub

        If colindex = 6 Then Me.Dgv_Excasso_numbers.Rows(rowindex).Cells(8).Value = Me.Dgv_Excasso_numbers.CurrentCell.Value * Me.Dgv_Excasso_numbers.Rows(rowindex).Cells(4).Value
        Me.Dgv_Excasso_numbers.Rows(1).Cells(10).Value = CInt(Me.Dgv_Excasso_numbers.Rows(0).Cells(8).Value) + CInt(Me.Dgv_Excasso_numbers.Rows(1).Cells(8).Value) + CInt(Me.Dgv_Excasso_numbers.Rows(2).Cells(8).Value)
        Me.Dgv_Excasso_numbers.Rows(2).Cells(10).Value = CInt(Me.Dgv_Excasso_numbers.Rows(0).Cells(10).Value) + CInt(Me.Dgv_Excasso_numbers.Rows(1).Cells(10).Value)


    End Sub



    Private Sub LinkLabel_Wisselkoers_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel_Wisselkoers.LinkClicked
        Process.Start("https://www.xe.com/currencyconverter/convert/?Amount=1&From=EUR&To=MDL")
    End Sub



    Private Sub Tbx_Excasso_Exchange_rate_TextChanged(sender As Object, e As EventArgs) Handles Tbx_Excasso_Exchange_rate.TextChanged
        Calculate_Excasso_Totals(False)
    End Sub

    Private Sub Tbx_Excasso_Exchange_rate_Validating(sender As Object, e As CancelEventArgs) Handles Tbx_Excasso_Exchange_rate.Validating

        Dim numericValue As Integer
        If Not Decimal.TryParse(Tbx_Excasso_Exchange_rate.Text, numericValue) Then
            Tbx_Excasso_Exchange_rate.Text = "1" ' Set default value to 1
        End If
    End Sub


    Private Sub Dgv_Excasso_numbers_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles Dgv_Excasso_numbers.CellClick
        'MsgBox("Dgv_Excasso_numbers_CellClick")
        If e.RowIndex < 0 Then Exit Sub

        If e.ColumnIndex = 2 Then  ' Column 2 is the button column
            Dgv_Excasso_numbers.Rows(0).Cells(2).Style.BackColor = Color.White
            Dgv_Excasso_numbers.Rows(1).Cells(2).Style.BackColor = Color.White
            Dgv_Excasso_numbers.Rows(2).Cells(2).Style.BackColor = Color.White
            Dgv_Excasso_numbers.Rows(e.RowIndex).Cells(2).Style.BackColor = Color.CornflowerBlue
            Calculate_Excasso_Totals(2)

        End If


    End Sub



    Private Sub Cbx_CP_Automatisch_CheckedChanged(sender As Object, e As EventArgs) Handles Cbx_CP_Automatisch.Click

        'If Cbx_CP_Automatisch.Checked Then
        Calculate_Excasso_Totals(False)
        'End If
    End Sub

    Sub Dgv_Excasso_numbers_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles Dgv_Excasso_numbers.CellContentClick


        If TypeOf Dgv_Excasso_numbers.Rows(e.RowIndex).Cells(e.ColumnIndex) Is DataGridViewCheckBoxCell Then

            If Strings.Left(Cmx_Excasso_Select.DisplayMember, 5) <> "Nieuw" Then Exit Sub

            Dgv_Excasso_numbers.CommitEdit(DataGridViewDataErrorContexts.Commit)
            Dgv_Excasso_numbers.Rows(0).Cells(2).Style.BackColor = Color.CornflowerBlue
            Dgv_Excasso_numbers.Rows(1).Cells(2).Style.BackColor = Color.White
            Dgv_Excasso_numbers.Rows(2).Cells(2).Style.BackColor = Color.White
            Load_New_Excasso(False)

        End If
    End Sub

    Private Sub Butn_Settings_Whatsnew_Click(sender As Object, e As EventArgs)
        If MsgBox("Nieuwe versie?", vbYesNoCancel) = vbYes Then My.Settings._whatsnew = "Ja" Else My.Settings._whatsnew = "Nee"
    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged

        My.Settings._whatsnew = DateTimePicker1.Value  '.ToString("yyyy-MM-dd")
        MsgBox(My.Settings._whatsnew)

    End Sub


    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Add_Journal_post_to_banktransaction()
    End Sub

    Private Sub Check_relid_Click(sender As Object, e As EventArgs)
        Lbl_11_contract__fk_relation_id.Text = Cmx_00_contract_fk_relation_id.SelectedValue
    End Sub

    Private Sub Cmx_01_contract_fk_target_id_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Cmx_01_contract_fk_target_id.SelectedIndexChanged
        If Not Add_Mode Then Exit Sub
        Dim target_id = QuerySQL($"Select id from target where concat(name, ', ', name_add) ='{Cmx_01_contract_fk_target_id.Text}'")
        Lbl_11_contract__fk_target_id.Text = target_id
    End Sub

    Private Sub Cmx_00_Contract_fk_account_id_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Cmx_Contract_fk_account_id.SelectedIndexChanged
        'If Not Add_Mode Then Exit Sub
        Lbl_10_Contract__fk_account_id.Text = Strings.Trim(Strings.Left(Cmx_Contract_fk_account_id.Text, 4))
    End Sub

    Private Sub Btn_Incasso_Click(sender As Object, e As EventArgs) Handles Btn_Incasso_Open_XML.Click
        Dim openFileDialog As New OpenFileDialog()
        openFileDialog.Filter = "XML Files|*.xml"

        If openFileDialog.ShowDialog() = DialogResult.OK Then
            Dim parser As New SepaParser()
            Try
                ' Roep de functie aan en bind het resultaat direct aan je Grid
                Dim table As DataTable = parser.ConvertSepaXmlToDataTable(openFileDialog.FileName)

                Dgv_Incasso.DataSource = table

                MessageBox.Show($"Er zijn {table.Rows.Count} transacties gevonden.", "Succes")
                Call Fill_Incasso_Overview(table)
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Fout")
            End Try
        End If
    End Sub
    Sub Count_Differences()
        Dim nieuw_deze_maand As Integer = 0
        Dim verdwenen_deze_maand As Integer = 0
        Dim ontbrekend_journaal As Integer = 0
        Dim ontbrekend_contract As Integer = 0
        Dim afwijkend_bedrag As Integer = 0

        Dim sql = QuerySQL($"select sql from query where name='Check_incasso'")
        If IsNothing(sql) Then Exit Sub
        Dim isd = Dtp_Incasso_start.Value
        sql = sql.replace("[date]", $"'{isd.Year}-{isd.Month}-{isd.Day}'")
        Dim incassodata = Collect_data2(sql)


        For x As Integer = 0 To incassodata.Rows.Count - 1
            Dim i = incassodata.Rows(x)(0)
            If i = "Toegevoegd t.o.v. vorige maand" Then nieuw_deze_maand += 1
            If i = "Verdwenen t.o.v. vorige maand" Then verdwenen_deze_maand += 1
            If i = "Ontbrekend in journaal" Then ontbrekend_journaal += 1
            If i = "Ontbrekend in contract" Then ontbrekend_contract += 1
            If i = "Afwijkend bedrag" Then afwijkend_bedrag += 1

        Next x


        Dgv_Incasso_Analyse.Columns.Clear()
        Dgv_Incasso_Analyse.Rows.Clear()

        Dgv_Incasso_Analyse.Columns.Add("Controle", "Controle")
        Dgv_Incasso_Analyse.Columns.Add("Aant", "Aant.")
        Dgv_Incasso_Analyse.Columns("Controle").Width = 146
        Dgv_Incasso_Analyse.Columns("Aant").Width = 50

        Dgv_Incasso_Analyse.Rows.Add("Nieuw deze maand", nieuw_deze_maand)
        Dgv_Incasso_Analyse.Rows.Add("Gestopt deze maand", verdwenen_deze_maand)
        'Dgv_Incasso_Analyse.Rows.Add("Afwezig in journal", ontbrekend_journaal)

        If Lbl_Incasso_Status.Text <> "Nieuw" Then

            Dim idx0 As Integer = Dgv_Incasso_Analyse.Rows.Add("Afwezig in journal", ontbrekend_journaal)

            If ontbrekend_journaal > 0 Then
                Dgv_Incasso_Analyse.Rows(idx0).DefaultCellStyle.Font = New Font(Dgv_Incasso_Analyse.Font, FontStyle.Bold)
                Dgv_Incasso_Analyse.Rows(idx0).DefaultCellStyle.ForeColor = Color.DarkRed
            End If

            Dim idx1 As Integer = Dgv_Incasso_Analyse.Rows.Add("Afwezig in contract", ontbrekend_contract)

            If ontbrekend_contract > 0 Then
                Dgv_Incasso_Analyse.Rows(idx1).DefaultCellStyle.Font = New Font(Dgv_Incasso_Analyse.Font, FontStyle.Bold)
                Dgv_Incasso_Analyse.Rows(idx1).DefaultCellStyle.ForeColor = Color.DarkRed
            End If

            ' 3. "Afwijkend bedrag" toevoegen en opmaken
            Dim idx2 As Integer = Dgv_Incasso_Analyse.Rows.Add("Afwijkend bedrag", afwijkend_bedrag)

            If afwijkend_bedrag > 0 Then
                Dgv_Incasso_Analyse.Rows(idx2).DefaultCellStyle.Font = New Font(Dgv_Incasso_Analyse.Font, FontStyle.Bold)
                Dgv_Incasso_Analyse.Rows(idx2).DefaultCellStyle.ForeColor = Color.DarkRed
            End If
        End If

        ' Zorg dat er niets geselecteerd is na het vullen (oogt rustiger)
        Dgv_Incasso_Analyse.ClearSelection()

    End Sub


    Sub Fill_Incasso_Overview(t As DataTable)
        'If Rbn_Incasso_Verschillen.Checked Then Exit Sub
        ' 1. Koppel de datasource los
        Dgv_incasso_totals.DataSource = Nothing

        ' 2. Wis ALLE bestaande kolommen en rijen om schoon te beginnen
        Dgv_incasso_totals.Columns.Clear()
        Dgv_incasso_totals.Rows.Clear()

        ' 3. Voeg de kolommen handmatig toe
        ' Structuur: .Add("UniekeNaam", "Koptekst")
        Dgv_incasso_totals.Columns.Add("colDoel", "Doel")
        Dgv_incasso_totals.Columns.Add("colAantal", "Aantal")
        Dgv_incasso_totals.Columns.Add("colTotaal", "Totaal")

        ' 4. (Optioneel) Opmaak van de kolommen instellen
        ' Kolom Aantal: rechts uitlijnen
        Dgv_incasso_totals.Columns("colAantal").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

        ' Kolom Totaal: rechts uitlijnen én als Euro weergeven (c2 = Currency met 2 decimalen)
        Dgv_incasso_totals.Columns("colTotaal").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Dgv_incasso_totals.Columns("colDoel").Width = 60
        Dgv_incasso_totals.Columns("colAantal").Width = 60
        Dgv_incasso_totals.Columns("colTotaal").Width = 80
        Dgv_incasso_totals.Columns("colTotaal").DefaultCellStyle.Format = "N2"
        'Dgv_incasso_totals.Columns("colTotaal").DefaultCellStyle.ForeColor = Color.Blue

        ' ---  DE BEREKENING  ---

        Dim kindAantal As Integer = 0
        Dim kindTotaal As Decimal = 0
        Dim oudereAantal As Integer = 0
        Dim oudereTotaal As Decimal = 0
        Dim overigAantal As Integer = 0
        Dim overigTotaal As Decimal = 0

        For Each row As DataRow In t.Rows
            ' ... (Dezelfde logica als in het vorige antwoord) ...
            Dim kenmerk As String = row("Mandaatcode").ToString().ToLower()
            Dim bedrag As Integer = 0
            If Not IsDBNull(row("Bedrag")) Then Decimal.TryParse(row("Bedrag").ToString(), bedrag)

            If kenmerk.StartsWith("k") Then
                kindAantal += 1
                kindTotaal += bedrag
            ElseIf kenmerk.StartsWith("o") Then
                oudereAantal += 1
                oudereTotaal += bedrag
            ElseIf kenmerk.StartsWith("v") Then
                overigAantal += 1
                overigTotaal += bedrag
            End If
        Next

        ' --- RIJEN TOEVOEGEN ---

        Dgv_incasso_totals.Rows.Add("Kind", kindAantal, kindTotaal)
        Dgv_incasso_totals.Rows.Add("Oudere", oudereAantal, oudereTotaal)
        Dgv_incasso_totals.Rows.Add("Overig", overigAantal, overigTotaal)

        Dim eindAantal As Integer = kindAantal + oudereAantal + overigAantal
        Dim eindTotaal As Decimal = kindTotaal + oudereTotaal + overigTotaal

        Dgv_incasso_totals.Rows.Add("Totaal", eindAantal, eindTotaal)

        ' Eventueel de 'Totaal' rij vetgedrukt maken
        Dim laatsteRijIndex As Integer = Dgv_incasso_totals.Rows.Count - 1
        Dgv_incasso_totals.Rows(laatsteRijIndex).DefaultCellStyle.Font = New Font(Dgv_incasso_totals.Font, FontStyle.Bold)


    End Sub


    Private Sub MonthCalendar1_DateSelected(sender As Object, e As DateRangeEventArgs) Handles MonthCalendar1.DateSelected
        ' 1. Haal de datum op
        Dim geselecteerdeDatum As Date = MonthCalendar1.SelectionStart

        '2 Fill the values
        Dtp_Incasso_start.Value = geselecteerdeDatum
        Lbl_Incasso_job_name.Text = "Contract incasso " & Dtp_Incasso_start.Value.ToString("d-M-yyyy")
        Lbl_Incasso_Status.Text = "Nieuw"
        ' 3. Verberg de kalender weer

        MonthCalendar1.Visible = False

        Prepare_Datagridview(Dgv_Incasso, Create_Incasso(Dtp_Incasso_start.Value.ToString), {"TZ250", "NG080", "TZ160", "TZ080", "TZ090", "DZ090"})


        Dim table As DataTable = CType(Dgv_Incasso.DataSource, DataTable)
        Fill_Incasso_Overview(table)
        Count_Differences()
        Pan_Incasso_views.Enabled = True
        MenuSave.Enabled = (Lbl_Incasso_Status.Text = "Nieuw")
        MenuDelete.Enabled = (Lbl_Incasso_Status.Text = "Open")
    End Sub

    Private Sub Cmx_Incasso_IncassoForm_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Cmx_Incasso_IncassoForm.SelectedIndexChanged

        If TC_Main.SelectedTab.Name <> "Incasso" Then Exit Sub


        If Cmx_Incasso_IncassoForm.SelectedIndex = -1 Then Exit Sub
        Dgv_incasso_totals.Columns.Clear()
        Dgv_incasso_totals.Rows.Clear()
        Dgv_Incasso_Analyse.Columns.Clear()
        Dgv_Incasso_Analyse.Rows.Clear()
        Lbl_Incasso_job_name.Text = ""
        Lbl_Incasso_Status.Text = " "

        'Dgv_incasso_totals.Columns.Clear()
        Dgv_Incasso.DataSource = Nothing
        Dgv_Incasso.Rows.Clear()
        Pan_Incasso_views.Enabled = False


        If Cmx_Incasso_IncassoForm.SelectedIndex = 0 Then  'nieuwe incasso
            Pan_Incasso_views.Enabled = False

            StelKalenderIn()

            MonthCalendar1.Location = New Point(Cmx_Incasso_IncassoForm.Left, Cmx_Incasso_IncassoForm.Bottom + 5)
            MonthCalendar1.BringToFront()
            MonthCalendar1.Visible = Not MonthCalendar1.Visible

            'MsgBox("2) " & Dtp_Incasso_start.Value)
            'Pan_Incasso_views.Enabled = True
        Else
            Dim geselecteerdItem As ComboBoxItem = CType(Cmx_Incasso_IncassoForm.SelectedItem, ComboBoxItem)
            Dim waarde1 As String = geselecteerdItem.Column1
            Dtp_Incasso_start.Value = waarde1
            Lbl_Incasso_job_name.Text = geselecteerdItem.Column2
            Lbl_Incasso_Status.Text = geselecteerdItem.Column3

            If Lbl_Incasso_Status.Text = "Nieuw" Or Lbl_Incasso_Status.Text = "Open" Then

                'Vullen van op contract gebaseerd incassoformulier 
                Prepare_Datagridview(Dgv_Incasso, Create_Incasso(Dtp_Incasso_start.Value.ToString("yyyy-MM-dd")), {"TZ250", "NG080", "TZ160", "TZ080", "TZ090", "DZ090"})
            Else
                Prepare_Datagridview(Dgv_Incasso, Create_Incasso_Processed(Dtp_Incasso_start.Value.ToString("yyyy-MM-dd")), {"TZ250", "NG080", "TZ160", "TZ080", "TZ090", "DZ090"})
            End If


            'vullen van datatable met totalen
            Dim table As DataTable = CType(Dgv_Incasso.DataSource, DataTable)
            Fill_Incasso_Overview(table)
            Count_Differences()
            Pan_Incasso_views.Enabled = True

        End If


        MenuSave.Enabled = (Lbl_Incasso_Status.Text = "Nieuw")
        MenuDelete.Enabled = (Lbl_Incasso_Status.Text = "Open")


    End Sub

    Private Sub Cmx_Incasso_IncassoForm_Click(sender As Object, e As EventArgs) Handles Cmx_Incasso_IncassoForm.Click
        If MonthCalendar1.Visible Then
            MonthCalendar1.Visible = False
            Cmx_Incasso_IncassoForm.SelectedIndex = -1
        End If
    End Sub

    Private Sub StelKalenderIn()
        ' 1. Stel de minimale datum in op Vandaag
        MonthCalendar1.MinDate = DateTime.Today
        Dim laatsteDatum As Date = DateTime.MinValue
        Dim sql As String = "SELECT MAX(date) FROM journal WHERE source = 'Incasso'"

        Try
            Connect(sql)
            Dim cmd As New NpgsqlCommand(sql, connection)

            ' ExecuteScalar is sneller voor 1 enkele waarde
            Dim result As Object = cmd.ExecuteScalar()

            ' Check of er wel een resultaat is (voor de allereerste keer draaien)
            If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                laatsteDatum = Convert.ToDateTime(result)
            End If

            connection.Close()

        Catch ex As Exception
            MsgBox("Fout bij ophalen datum: " & ex.Message)
        Finally
            If connection.State = ConnectionState.Open Then connection.Close()
        End Try

        ' 3. Bereken de nieuwe datum
        If laatsteDatum <> DateTime.MinValue Then
            Dim voorgesteldeDatum As Date = laatsteDatum.AddMonths(1)

            While voorgesteldeDatum < DateTime.Today
                voorgesteldeDatum = voorgesteldeDatum.AddMonths(1)
            End While

            ' Stel de kalender in
            MonthCalendar1.SelectionStart = voorgesteldeDatum
            MonthCalendar1.SelectionEnd = voorgesteldeDatum
        Else
            ' Als er nog nooit een incasso is geweest, pak dan vandaag
            MonthCalendar1.SelectionStart = DateTime.Today
        End If
    End Sub

    Private Sub Btn_Incasso_Analyseer_Click(sender As Object, e As EventArgs) Handles Btn_Incasso_Analyseer.Click
        Dim arr_format() As String = Nothing
        Dim sql = QuerySQL($"select sql from query where name='Check_incasso'")
        If IsNothing(sql) Then Exit Sub
        Dim formatting = QuerySQL($"select sql from query where name='Check_incasso'")
        If Not IsNothing(formatting) Then arr_format = formatting.Split(",")
        Dim isd = Dtp_Incasso_start.Value
        sql = sql.replace("[date]", $"'{isd.Year}-{isd.Month}-{isd.Day}'")

        Prepare_Datagridview(Dgv_Incasso, sql, {"TZ350", "TZ200", "NZ080", "NZ080"})
    End Sub

    Private Sub Dtp_31_contract__enddate_ValueChanged(sender As Object, e As EventArgs) Handles Dtp_31_contract__enddate.ValueChanged

    End Sub

    Private Sub Dgv_Excasso2_SelectionChanged(sender As Object, e As EventArgs) Handles Dgv_Excasso2.SelectionChanged
        ' 1. Remember the current state (was it already False because a Load is running?)
        Dim previousState As Boolean = isManualChange

        ' 2. Disable change tracking
        isManualChange = False

        ' 3. Update the Tags for the new row
        UpdateControlTags(Me)

        ' 4. Restore tracking to whatever it was BEFORE this event fired.
        ' (If Load_Existing_Excasso is running, previousState is False, so it safely stays False!)
        isManualChange = previousState
    End Sub

    Private Sub TC_Tussenrekening_Click(sender As Object, e As EventArgs) Handles TC_Tussenrekening.Click


        Dim lab0 = "Tussenrekening Rapportage"
        Dim lab1 = "Netting Ouderdom"
        Dim lab2 = "Netting Flowtrend"
        Dim lab3 = "Netting Volatiliteit"
        Dim lab4 = "Netting Volume"
        Lbl_netting0.Text = lab0
        Lbl_netting1.Text = lab1
        Lbl_Netting2.Text = lab2
        Lbl_Netting3.Text = lab3
        Lbl_Netting4.Text = lab4

        Dim Sql0 As String = QuerySQL($"Select sql from query where category = 'Overzicht' and name='{lab0}'")
        Prepare_Datagridview(Dgv_Tussenrekening_0, Sql0, {"TZ200", "DZ020", "TZ150"})
        Dim Sql1 As String = QuerySQL($"Select sql from query where category = 'Overzicht' and name='{lab1}'")
        Prepare_Datagridview(Dgv_Tussenrekening_1, Sql1, {"TZ150", "DZ070", "JZ030", "TZ030", "NZ070"})
        Dim Sql2 As String = QuerySQL($"Select sql from query where category = 'Overzicht' and name='{lab2}'")
        Prepare_Datagridview(Dgv_Tussenrekening_2, Sql2, {"DZ080", "NZ080", "NZ080", "NZ080"})
        Dim Sql3 As String = QuerySQL($"Select sql from query where category = 'Overzicht' and name='{lab3}'")
        Prepare_Datagridview(Dgv_Tussenrekening_3, Sql3, {"DZ080", "NZ080", "NZ080", "TZ080"})
        Dim Sql4 As String = QuerySQL($"Select sql from query where category = 'Overzicht' and name='{lab4}'")
        Prepare_Datagridview(Dgv_Tussenrekening_4, Sql4, {"TZ200", "DZ070", "JZ030", "NZ070"})


    End Sub

    Private Sub SPAS_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If Not MustWarn Then Exit Sub

        Dim alert As Integer = QuerySQL(QuerySQL($"Select sql from query where category = 'Check' and name='CountAlert'"))

        If alert > 0 Then
            Dim result As MsgBoxResult = MsgBox($"Er zijn nog {alert} aandachtsgebieden m.b.t. de tussenrekening. Zie hiervoor het tabblad 'Tussenrekening > Analyse'{vbCr}
{vbCr}Weet u zeker dat u wilt stoppen?", vbYesNo + vbExclamation, "Waarschuwing")

            If result = vbNo Then
                ' This is the magic line that prevents the application from closing
                e.Cancel = True
            End If
        End If
    End Sub

    ''' <summary>
    ''' Fetches data from the repository and loads it into the AccountTree.
    ''' </summary>
    Private Sub LoadAccountTree()
        ' 1. Ask the repository (in Beheer.vb) for the data, passing the current search text
        Dim treeData As DataTable = AccountRepository.GetAccountHierarchyData(Me.Searchbox2.Text)

        ' 2. Hand the data and the UI control to the generic mapper
        TreeViewMapper.Populate3LevelTree(Me.AccountTree, treeData, "AccountType", "AccountGroup", "Account")

        ' 3. Auto-expand if the user is actively searching
        If Not String.IsNullOrWhiteSpace(Me.Searchbox2.Text) Then
            Me.AccountTree.ExpandAll()
        End If
    End Sub


    Private Sub ExpandOnlyLevel1()
        AccountTree.BeginUpdate()
        AccountTree.CollapseAll() ' Reset the tree to fully closed

        ' Loop only through the root nodes (Level 1)
        For Each rootNode As TreeNode In AccountTree.Nodes
            rootNode.Expand()
        Next

        AccountTree.EndUpdate()
    End Sub

    ''' <summary>
    ''' Handles UI updates when a node is clicked in the AccountTree.
    ''' </summary>
    Private Sub AccountTree_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles AccountTree.AfterSelect
        If e.Node Is Nothing Then Exit Sub

        ' 1. Suspend change tracking to prevent false triggers while populating fields
        Dim previousState As Boolean = isManualChange
        isManualChange = False

        ' 2. Menu Button Logic
        ' Enable Add only for Level 1 (AccountType) or Level 2 (AccountGroup)
        MenuAdd.Enabled = (e.Node.Name = "AccountType" Or e.Node.Name = "AccountGroup")

        ' Disable Save and Cancel (they will only be enabled by MenuAdd_Click or FieldChangedHandler)
        MenuSave.Enabled = False
        MenuCancel.Enabled = False

        ' 3. Reset UI to baseline
        Grbx_Beheer_Accountgroep.Enabled = False
        Grbx_Beheer_Account.Enabled = False

        AccountTree.BeginUpdate()

        ' 4. Accordion Logic: Find the top-level root node
        Dim currentRoot As TreeNode = e.Node
        While currentRoot.Parent IsNot Nothing
            currentRoot = currentRoot.Parent
        End While

        ' Collapse all Level 1 nodes we aren't currently inside
        For Each rootNode As TreeNode In AccountTree.Nodes
            If rootNode IsNot currentRoot Then
                rootNode.Collapse()
            End If
        Next

        ' Expand the clicked node
        e.Node.Expand()
        AccountTree.EndUpdate()

        ' 5. UI Routing Logic
        Select Case e.Node.Name
            Case "AccountType"
                Tbx_Beheer_Accountgroepnaam.Text = ""
                Tbx_Beheer_Accountnaam.Text = ""
                Clear_Account()

            Case "AccountGroup"
                Grbx_Beheer_Accountgroep.Enabled = True
                Tbx_Beheer_Accountgroepnaam.Text = e.Node.Text
                Clear_Account()

                Dim groupId As String = e.Node.Tag.ToString()
                Lbl_Beheer_Accgroup_id.Text = groupId
                Tbx_Beheer_Accountnaam.Text = ""

                Dim groupDetails As DataRow = AccountRepository.GetAccountGroupDetails(groupId)
                If groupDetails IsNot Nothing Then
                    Tbx_Beheer_Accgroup_Description.Text = If(IsDBNull(groupDetails("description")), "", groupDetails("description").ToString())
                    Dim accType As String = If(IsDBNull(groupDetails("type")), "", groupDetails("type").ToString())
                    Tbx_Beheer_Account_Type.Text = accType
                    Chbx_Beheer_Accgroup_Active.Checked = If(IsDBNull(groupDetails("active")), False, Convert.ToBoolean(groupDetails("active")))
                    Rbtn_Beheer_Accounttype1.Checked = (accType = "Inkomsten")
                    Rbtn_Beheer_Accounttype2.Checked = (accType = "Uitgaven")
                    Rbtn_Beheer_Accounttype3.Checked = (accType = "Transit")
                    Cmbox_Beheer_Accgroup_Subtype.Text = If(IsDBNull(groupDetails("subtype")), "", groupDetails("subtype").ToString())

                    ' If 'posts' doesn't exist in GetAccountGroupDetails, ensure it's handled safely
                    Lbl_Beheer_Accgroup_posts.Text = If(groupDetails.Table.Columns.Contains("posts") AndAlso Not IsDBNull(groupDetails("posts")), groupDetails("posts").ToString(), "0")
                    Chbx_Beheer_Accgroup_Active.Enabled = (Lbl_Beheer_Accgroup_posts.Text = "0")
                Else
                    Tbx_Beheer_Accgroup_Description.Text = ""
                    Tbx_Beheer_Account_Type.Text = ""
                    Chbx_Beheer_Accgroup_Active.Checked = False
                    Rbtn_Beheer_Accounttype1.Checked = False
                    Rbtn_Beheer_Accounttype2.Checked = False
                    Rbtn_Beheer_Accounttype3.Checked = False
                End If

            Case "Account"
                Grbx_Beheer_Accountgroep.Enabled = False
                Grbx_Beheer_Account.Enabled = True
                Cmbx_Beheer_Accgroup.Enabled = True
                Dim accountId As String = e.Node.Tag.ToString()
                Tbx_Beheer_Accountnaam.Text = e.Node.Text

                ' ---> FIX: Explicitly cast the Tag to an Integer to satisfy WinForms binding
                If e.Node.Parent IsNot Nothing AndAlso e.Node.Parent.Tag IsNot Nothing Then
                    Dim parentId As Integer
                    If Integer.TryParse(e.Node.Parent.Tag.ToString(), parentId) Then
                        Cmbx_Beheer_Accgroup.SelectedValue = parentId
                    End If
                End If

                If e.Node.Parent IsNot Nothing Then
                    Tbx_Beheer_Accountgroepnaam.Text = e.Node.Parent.Text
                    Lbl_Beheer_Accgroup_id.Text = e.Node.Parent.Tag.ToString()

                    Dim accountDetails As DataRow = AccountRepository.GetAccountDetails(accountId)
                    If accountDetails IsNot Nothing Then
                        Lbl_Beheer_Account_id.Text = accountId
                        Tbx_Beheer_Account_Description1.Text = If(accountDetails.Table.Columns.Contains("description") AndAlso Not IsDBNull(accountDetails("description")), accountDetails("description").ToString(), "")
                        Tbx_Beheer_Accounttype.Text = Trim(If(IsDBNull(accountDetails("type")), "", accountDetails("type").ToString()))
                        Tbx_Beheer_Account_Code.Text = If(accountDetails.Table.Columns.Contains("bankcode") AndAlso Not IsDBNull(accountDetails("bankcode")), accountDetails("bankcode").ToString(), "")
                        Tbx_Beheer_AccountTrefwoorden.Text = If(accountDetails.Table.Columns.Contains("searchword") AndAlso Not IsDBNull(accountDetails("searchword")), accountDetails("searchword").ToString(), "")
                        Lbl_Beheer_Account_posts.Text = If(accountDetails.Table.Columns.Contains("posts") AndAlso Not IsDBNull(accountDetails("posts")), accountDetails("posts").ToString(), "0")
                        Chbx_Beheer_Account_Active.Checked = If(IsDBNull(accountDetails("active")), False, Convert.ToBoolean(accountDetails("active")))
                        Cmbx_Beheer_Accgroup.Text = Tbx_Beheer_Accountgroepnaam.Text
                        Tbx_Beheer_Accountbron.Text = If(IsDBNull(accountDetails("source")), "", accountDetails("source").ToString())

                        Rbtn_Beheer_Account_1.Checked = (Rbtn_Beheer_Account_1.Tag IsNot Nothing AndAlso Rbtn_Beheer_Account_1.Tag.ToString() = Tbx_Beheer_Accounttype.Text)
                        Rbtn_Beheer_Account_2.Checked = (Rbtn_Beheer_Account_2.Tag IsNot Nothing AndAlso Rbtn_Beheer_Account_2.Tag.ToString() = Tbx_Beheer_Accounttype.Text)
                        Rbtn_Beheer_Account_3.Checked = (Rbtn_Beheer_Account_3.Tag IsNot Nothing AndAlso Rbtn_Beheer_Account_3.Tag.ToString() = Tbx_Beheer_Accounttype.Text)

                        Chbx_Beheer_Account_Active.Enabled = (Lbl_Beheer_Account_posts.Text = "0")
                    Else
                        Clear_Account()
                    End If
                End If
        End Select

        ' 6. Critical Step: Update the tags to the new values so the FieldChangedHandler establishes this as the new baseline
        UpdateControlTags(Me)

        ' 7. Restore change tracking 
        isManualChange = previousState
    End Sub


    Sub Clear_Account()
        Tbx_Beheer_Account_Description1.Text = ""
        Tbx_Beheer_Accounttype.Text = ""
        Tbx_Beheer_Account_Code.Text = ""
        Tbx_Beheer_AccountTrefwoorden.Text = ""
        Chbx_Beheer_Accgroup_Active.Checked = False
        Rbtn_Beheer_Account_1.Checked = False
        Rbtn_Beheer_Account_2.Checked = False
        Rbtn_Beheer_Account_3.Checked = False
        Cmbx_Beheer_Accgroup.SelectedIndex = -1
    End Sub
    ''' <summary>
    ''' Expands a specific Level 1 node by its text value (e.g., "Inkomsten"), collapsing all others.
    ''' </summary>
    Private Sub ExpandSpecificAccountType(targetTypeName As String)
        AccountTree.BeginUpdate()

        ' 1. Reset the tree to a fully collapsed state
        AccountTree.CollapseAll()

        ' 2. Loop through only the Level 1 nodes (the root nodes)
        For Each rootNode As TreeNode In AccountTree.Nodes

            ' 3. Perform a case-insensitive check against the text
            If rootNode.Text.Equals(targetTypeName, StringComparison.OrdinalIgnoreCase) Then

                ' Expand this specific node
                rootNode.Expand()

                ' Scroll the tree to ensure it is visible to the user
                rootNode.EnsureVisible()

                ' Exit the loop early since we found what we were looking for
                Exit For
            End If
        Next

        AccountTree.EndUpdate()
    End Sub
    Private Sub Lbl_00_Account__source_Click(sender As Object, e As EventArgs) Handles Lbl_00_Account__source.Click

    End Sub

    Private Sub Btn_Account_Budget_All_Click_1(sender As Object, e As EventArgs) Handles Btn_Account_Budget_All.Click
        Calculate_Budget("")
    End Sub

    Private Sub Chbx_Beheer_Account_Active_CheckedChanged(sender As Object, e As EventArgs) Handles Chbx_Beheer_Account_Active.CheckedChanged
        If Lbl_Beheer_Account_posts.Text = "0" And Chbx_Beheer_Account_Active.Checked = False Then
            MsgBox($"Deactivering niet mogelijk: er zijn {Lbl_Beheer_Account_posts} boeking op dit account")
            Chbx_Beheer_Account_Active.Checked = True
        End If
    End Sub
    Private Function SaveCurrentAccountGroup() As Boolean
        Try
            Dim selectedType As String = ""
            If Rbtn_Beheer_Accounttype1.Checked Then selectedType = "Inkomsten"
            If Rbtn_Beheer_Accounttype2.Checked Then selectedType = "Uitgaven"
            If Rbtn_Beheer_Accounttype3.Checked Then selectedType = "Transit"

            Dim groupId As Integer = 0
            Integer.TryParse(Lbl_Beheer_Accgroup_id.Text, groupId)

            Dim groupModel As New AccountGroupModel() With {
            .Id = groupId,
            .Name = Tbx_Beheer_Accountgroepnaam.Text,
            .Type = selectedType,
            .Subtype = Cmbox_Beheer_Accgroup_Subtype.Text,
            .Description = Tbx_Beheer_Accgroup_Description.Text,
            .Active = Chbx_Beheer_Accgroup_Active.Checked
        }

            AccountRepository.SaveAccountGroup(groupModel)
            MsgBox("Accountgroep succesvol opgeslagen.", MsgBoxStyle.Information)

            Add_Mode = False

            ' ---> FIX 3: Refresh the combobox so the newly saved group is immediately added to the dropdown!
            Load_Combobox(Cmbx_Beheer_Accgroup, "id", "name", "SELECT id, name FROM accgroup WHERE active=True ORDER BY name")

            Dim savedId As String = groupModel.Id.ToString()
            If groupModel.Id = 0 Then
                savedId = QuerySQL("SELECT MAX(id) FROM accgroup;").ToString()
            End If

            LoadAccountTree()
            RestoreTreeSelection("AccountGroup", savedId)

            Return True ' <--- Save was successful!

        Catch ex As ArgumentException
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Validatie Fout")
            Return False ' <--- Save failed
        Catch ex As Exception
            MsgBox($"Er is een onverwachte fout opgetreden: {ex.Message}", MsgBoxStyle.Critical, "Systeem Fout")
            Return False ' <--- Save failed
        End Try
    End Function

    ''' <summary>
    ''' Reads the UI, maps to the AccountModel, and triggers the repository save.
    ''' Returns True if successful, False if validation fails.
    ''' </summary>
    Private Function SaveCurrentAccount() As Boolean
        Try
            ' 1. Extract the Account Group ID safely using Index 0
            Dim groupId As Integer = 0

            If Cmbx_Beheer_Accgroup.SelectedItem IsNot Nothing Then
                If TypeOf Cmbx_Beheer_Accgroup.SelectedItem Is DataRowView Then
                    Dim rowView As DataRowView = DirectCast(Cmbx_Beheer_Accgroup.SelectedItem, DataRowView)
                    Integer.TryParse(rowView(0).ToString(), groupId)
                ElseIf Cmbx_Beheer_Accgroup.SelectedValue IsNot Nothing Then
                    Integer.TryParse(Cmbx_Beheer_Accgroup.SelectedValue.ToString(), groupId)
                End If
            End If

            If groupId <= 0 Then
                MsgBox("Selecteer een geldige accountgroep uit de lijst.", MsgBoxStyle.Exclamation, "Validatie Fout")
                Return False
            End If

            ' 2. Extract the account type directly from the radio buttons
            Dim accType As String = ""
            If Rbtn_Beheer_Account_1.Checked Then accType = Rbtn_Beheer_Account_1.Text
            If Rbtn_Beheer_Account_2.Checked Then accType = Rbtn_Beheer_Account_2.Text
            If Rbtn_Beheer_Account_3.Checked Then accType = Rbtn_Beheer_Account_3.Text

            ' 3. Map the data to the model (Comments removed to prevent line continuation errors)
            Dim accModel As New AccountModel() With {
            .Id = If(Integer.TryParse(Lbl_Beheer_Account_id.Text, Nothing), Convert.ToInt32(Lbl_Beheer_Account_id.Text), 0),
            .FkAccGroupId = groupId,
            .Name = Tbx_Beheer_Accountnaam.Text,
            .Type = accType,
            .Source = Tbx_Beheer_Accountbron.Text,
            .FKey = If(Integer.TryParse(Lbl_20_Account__f_key.Text, Nothing), Convert.ToInt32(Lbl_20_Account__f_key.Text), 0),
            .Active = Chbx_Beheer_Account_Active.Checked,
            .Description = Tbx_Beheer_Account_Description1.Text,
            .Bankcode = Tbx_Beheer_Account_Code.Text,
            .Searchword = Tbx_Beheer_AccountTrefwoorden.Text
        }

            ' 4. Save to database
            AccountRepository.SaveAccount(accModel)
            MsgBox("Account succesvol opgeslagen.", MsgBoxStyle.Information)

            Add_Mode = False

            ' 5. Fetch ID if it was a newly inserted account
            Dim savedId As String = accModel.Id.ToString()
            If accModel.Id = 0 Then
                savedId = QuerySQL("SELECT MAX(id) FROM account;").ToString()
            End If

            ' 6. Reload the tree and focus the newly saved item
            LoadAccountTree()
            RestoreTreeSelection("Account", savedId)

            Return True ' <--- Save was successful!

        Catch ex As ArgumentException
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Validatie Fout")
            Return False ' <--- Save failed
        Catch ex As Exception
            MsgBox($"Er is een onverwachte fout opgetreden: {ex.Message}", MsgBoxStyle.Critical, "Systeem Fout")
            Return False ' <--- Save failed
        End Try
    End Function

    Private Sub Btn_Account_Budget_Id_Click_1(sender As Object, e As EventArgs) Handles Btn_Account_Budget_Id.Click
        Calculate_Budget(Lbl_Beheer_Account_id.Text)
    End Sub

    ''' <summary>
    ''' Clears the group fields, then searches the tree to re-select the newly saved node.
    ''' </summary>
    Private Sub RestoreTreeSelection(targetName As String, targetTag As String)
        ' 1. Wipe the stale data from the screen
        Tbx_Beheer_Accountgroepnaam.Text = ""
        Tbx_Beheer_Accgroup_Description.Text = ""
        Cmbox_Beheer_Accgroup_Subtype.SelectedIndex = -1
        Cmbox_Beheer_Accgroup_Subtype.Text = ""
        Rbtn_Beheer_Accounttype1.Checked = False
        Rbtn_Beheer_Accounttype2.Checked = False
        Rbtn_Beheer_Accounttype3.Checked = False
        Clear_Account()

        ' 2. Search the newly built tree for the node we just saved
        Dim foundNode As TreeNode = FindNodeByNameAndTag(AccountTree.Nodes, targetName, targetTag)

        If foundNode IsNot Nothing Then
            AccountTree.SelectedNode = foundNode
            foundNode.EnsureVisible()
        End If
    End Sub

    Private Function FindNodeByNameAndTag(nodes As TreeNodeCollection, name As String, tag As String) As TreeNode
        For Each node As TreeNode In nodes
            If node.Name = name AndAlso node.Tag IsNot Nothing AndAlso node.Tag.ToString() = tag Then
                Return node
            End If
            Dim foundChild As TreeNode = FindNodeByNameAndTag(node.Nodes, name, tag)
            If foundChild IsNot Nothing Then Return foundChild
        Next
        Return Nothing
    End Function
End Class
