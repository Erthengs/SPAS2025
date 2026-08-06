Imports System.Net.Http
Imports System.Text.RegularExpressions
Imports System.Threading.Tasks
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
                db ='PROD'
                SPAS.Pan_Test.Visible = False
                SPAS.Text &= " " & username
                SPAS.BackColor = Color.LightSteelBlue  'Color.WhiteSmoke
                SPAS.MenuStrip1.BackColor = Color.LightSteelBlue
                SPAS.Testpanel.BackColor = Color.LightSteelBlue
                SPAS.ToolStripTextBox1.BackColor = Color.LightSteelBlue

            Case "Acceptatie"
                'db ='ACC'
                connect_string = "User ID=" & username & ";Password=" & Me.Tbx_login_password.Text & ";Host=hw26607-001.dbaas.ovh.net;Port=35263;Database=ACC" 'My.Settings._accovh
                SPAS.Pan_Test.Visible = True

                SPAS.Text &= " " & username & " (TIJDELIJKE ACCEPTATIE DATABASE)"
                SPAS.BackColor = Color.YellowGreen
                SPAS.Testpanel.BackColor = Color.GreenYellow
                SPAS.MenuStrip1.BackColor = Color.GreenYellow
                SPAS.ToolStripTextBox1.BackColor = Color.GreenYellow
            Case "Test"
                'db ='TEST'
                connect_string = "User ID=" & username & ";Password=" & Me.Tbx_login_password.Text & ";Host=hw26607-001.dbaas.ovh.net;Port=35263;Database=SPAS-TEST" 'My.Settings._tstovh
                SPAS.Pan_Test.Visible = True
                SPAS.Pan_Test.Visible = True
                SPAS.Text &= " " & username & " (TEST DATABASE)"
                SPAS.BackColor = Color.Orange
                SPAS.MenuStrip1.BackColor = Color.Orange
                SPAS.Testpanel.BackColor = Color.Orange
                SPAS.ToolStripTextBox1.BackColor = Color.Orange

        End Select


        'test connectie
        Dim connection As NpgsqlConnection
        Dim ex As Exception = Nothing
        Try
            connection = New NpgsqlConnection(connect_string)
            connection.Open()
        Catch ex
            MsgBox("Inloggen niet gelukt, probeer het nogmaals (controleer of gebruikersnaam, wachtwoord en IP-adres correct zijn). ")
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

        Lbl_Login_lastlogin.Text = $"Vorige login: {My.Settings._whatsnew}"
        Lbl_login_version.Text = $"Versie {Me.Tag}"
        Cmx_Login_Database.Text = My.Settings._lastdb
        Me.Tbx_Login_username.Text = My.Settings._produser
        If My.Settings._prodpwd <> "" Then
            Me.Tbx_login_password.Text = My.Settings._prodpwd
            Chbx_Login_Save_Password.Checked = True
        Else
            Chbx_Login_Save_Password.Checked = False
        End If
        CheckForUpdate()


    End Sub
    Private Async Sub CheckForUpdate()


        Dim latestVersion As String = Await VersionChecker.GetLatestVersionDateAsync()
        Dim str() = Split(latestVersion, ",")
        Dim version = str(0)
        Dim latestVersionDate As Date = str(1)
        Dim _date As Date

        If My.Settings._whatsnew IsNot Nothing And My.Settings._whatsnew <> "" Then
            _date = CDate(My.Settings._whatsnew).ToShortDateString
        Else
            _date = CDate("15-01-2000").ToShortDateString
        End If

        If version > Me.Tag Then
            MsgBox("Er is een nieuwe versie beschikbaar, synchroniseer eerst Dropbox s.v.p.", vbExclamation)
        ElseIf CDate(latestVersionDate).ToShortDateString() > _date Then
            Dim answer = MsgBox("Er is een nieuwe versie geïnstalleerd, wilt u de wijzigingen bekijken?", vbYesNo)
            If answer = vbYes Then Process.Start("https://github.com/Erthengs/SPAS2025/wiki/Aanpassingen-per-versie")
            My.Settings._whatsnew = latestVersionDate
        End If

    End Sub

    Private Sub LinkLabel_Wisselkoers_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles Lbl_login_version.LinkClicked
        Process.Start("https://github.com/Erthengs/SPAS2025/wiki/Aanpassingen-per-versie")
    End Sub


End Class


Public Class VersionChecker
    Private Shared ReadOnly wikiUrl As String = "https://github.com/Erthengs/SPAS2025/wiki/Aanpassingen-per-versie"

    Public Shared Async Function GetLatestVersionDateAsync() As Task(Of String)
        Try
            Using client As New HttpClient()
                ' Fetch the HTML content of the wiki page
                Dim pageContent As String = Await client.GetStringAsync(wikiUrl)

                ' Find the first occurrence of "Huidige versie: X.X (DD-M-YYYY)"
                Dim match As Match = Regex.Match(pageContent, "Huidige versie:\s*([\d.]+)\s*\((\d{1,2}-\d{1,2}-\d{4})\)")

                If match.Success Then
                    ' Extract the date string from the regex match
                    Return match.Groups(1).Value & "," & match.Groups(2).Value

                Else
                    MessageBox.Show("No matching version date found on the wiki page.")
                End If
            End Using
        Catch ex As Exception
            MessageBox.Show("Error checking latest version: " & ex.Message)
        End Try

        Return Nothing
    End Function

End Class

