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
        If Cmx_Login_Database.Text = "Productie" Then
            connect_string = "User ID=" & username & ";Password=" & Me.Tbx_login_password.Text & My.Settings._prod
            SPAS.Pan_Test.Visible = False
            SPAS.Text = "SPAS " & username
            SPAS.BackColor = Color.WhiteSmoke
        ElseIf Cmx_Login_Database.Text = "Acceptatie" Then
            connect_string = "User ID=" & username & ";Password=" & Me.Tbx_login_password.Text & My.Settings._tstovh
            SPAS.Pan_Test.Visible = True
            SPAS.Text = "SPAS " & username & " (CLOUD ACCEPATIETDATABASE)"
            SPAS.BackColor = Color.YellowGreen
        Else
            connect_string = "User ID=" & username & ";Password=" & Me.Tbx_login_password.Text & My.Settings._accovh
            SPAS.Pan_Test.Visible = True
            SPAS.Text = "SPAS " & username & " (CLOUD TESTDATABASE)"
        End If
        Me.Close()
        Count_Occurences()
        SPAS.InitLoad()

    End Sub

    Private Sub Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel.Click
        Me.Close()
        Application.Exit()
    End Sub

    Private Sub Login_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'MsgBox(sender.ToString)


    End Sub

    Private Sub Cmx_Login_Database_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Cmx_Login_Database.SelectedIndexChanged
        Me.Tbx_Login_username.Text = My.Settings._produser
        Me.Tbx_login_password.Text = My.Settings._prodpwd

        Exit Sub
        If Cmx_Login_Database.Text = "Productie" Or Cmx_Login_Database.Text = "Acceptatie" Then

            Me.Tbx_Login_username.Text = My.Settings._produser
            Me.Tbx_login_password.Text = My.Settings._prodpwd
        Else
            Me.Tbx_Login_username.Text = My.Settings._testuser
            Me.Tbx_login_password.Text = My.Settings._testpwd

        End If
    End Sub
End Class
