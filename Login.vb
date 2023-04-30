Imports Npgsql
Public Class Login

    ' TODO: Insert code to perform custom authentication using the provided username and password 
    ' (See https://go.microsoft.com/fwlink/?LinkId=35339).  
    ' The custom principal can then be attached to the current thread's principal as follows: 
    '     My.User.CurrentPrincipal = CustomPrincipal
    ' where CustomPrincipal is the IPrincipal implementation used to perform authentication. 
    ' Subsequently, My.User will return identity information encapsulated in the CustomPrincipal object
    ' such as the username, display name, etc.

    Private Sub OK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK.Click
        username = Me.Tbx_Login_username.Text
        Select Case Cmx_Login_Database.Text
            Case "Productie"
                connect_string = "User ID=" & username & ";Password=" & Me.Tbx_login_password.Text & ";Host=hw26607-001.dbaas.ovh.net;Port=35263;Database=SPAS-PROD"
                SPAS.Pan_Test.Visible = False
                SPAS.Text = "SPAS " & username
                SPAS.BackColor = Color.WhiteSmoke
                SPAS.MenuStrip1.BackColor = Color.LightSteelBlue
                SPAS.ToolStripTextBox1.BackColor = Color.LightSteelBlue
                SPAS.TC_Main.TabPages.Remove(SPAS.TC_Main.TabPages(7))
            Case "Acceptatie"
                connect_string = "User ID=" & username & ";Password=" & Me.Tbx_login_password.Text & ";Host=hw26607-001.dbaas.ovh.net;Port=35263;Database=ACC" 'My.Settings._accovh
                SPAS.Pan_Test.Visible = True
                SPAS.Lbl_Excasso_Items_Contract.Visible = True
                SPAS.Lbl_Excasso_Items_Extra.Visible = True
                SPAS.Lbl_Excasso_Items_Intern.Visible = True
                SPAS.Lbl_Excasso_Contractwaarde.Visible = True
                SPAS.Lbl_Excasso_Extra.Visible = True
                SPAS.Lbl_Excasso_Intern.Visible = True
                SPAS.Text = "SPAS " & username & " (TIJDELIJKE ACCEPTATIE DATABASE)"
                SPAS.BackColor = Color.YellowGreen
                SPAS.MenuStrip1.BackColor = Color.GreenYellow
                SPAS.ToolStripTextBox1.BackColor = Color.GreenYellow
            Case "Test"
                connect_string = "User ID=" & username & ";Password=" & Me.Tbx_login_password.Text & ";Host=hw26607-001.dbaas.ovh.net;Port=35263;Database=SPAS-TEST" 'My.Settings._tstovh
                SPAS.Pan_Test.Visible = True
                SPAS.Pan_Test.Visible = True
                SPAS.Lbl_Excasso_Items_Contract.Visible = True
                SPAS.Lbl_Excasso_Items_Extra.Visible = True
                SPAS.Lbl_Excasso_Items_Intern.Visible = True
                SPAS.Lbl_Excasso_Contractwaarde.Visible = True
                SPAS.Lbl_Excasso_Extra.Visible = True
                SPAS.Lbl_Excasso_Intern.Visible = True
                SPAS.Text = "SPAS " & username & " (TEST DATABASE)"

        End Select


        'test connectie
        Dim connection As NpgsqlConnection
        Dim ex As Exception = Nothing
        Try
            connection = New NpgsqlConnection(connect_string)
            connection.Open()
        Catch ex
            MsgBox("Inloggen niet gelukt, probeer het nogmaals (controleer of gebruikersnaam en wachtwoord correct zijn). ")
        End Try
        If ex Is Nothing Then

            My.Settings._produser = Me.Tbx_Login_username.Text
            My.Settings._lastdb = Cmx_Login_Database.Text
            My.Settings._prodpwd = IIf(Chbx_Login_Save_Password.Checked, Me.Tbx_login_password.Text, "")

            Count_Occurences()
            Me.Close()
            SPAS.InitLoad()
        End If

    End Sub

    Private Sub Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel.Click
        Me.Close()
        Application.Exit()
    End Sub

    Private Sub Login_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'MsgBox(sender.ToString)

        Cmx_Login_Database.Text = My.Settings._lastdb
        Me.Tbx_Login_username.Text = My.Settings._produser
        If My.Settings._prodpwd <> "" Then
            Me.Tbx_login_password.Text = My.Settings._prodpwd
            Chbx_Login_Save_Password.Checked = True
        Else
            Chbx_Login_Save_Password.Checked = False
        End If
        'My.Settings._whatsnew = "Ja"
    End Sub

    Private Sub Cmx_Login_Database_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Cmx_Login_Database.SelectedIndexChanged

    End Sub
End Class
