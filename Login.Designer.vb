<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
<Global.System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Naming", "CA1726")> _
Partial Class Login
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub
    Friend WithEvents LogoPictureBox As System.Windows.Forms.PictureBox
    Friend WithEvents UsernameLabel As System.Windows.Forms.Label
    Friend WithEvents PasswordLabel As System.Windows.Forms.Label
    Friend WithEvents Tbx_Login_username As System.Windows.Forms.TextBox
    Friend WithEvents Tbx_login_password As System.Windows.Forms.TextBox
    Friend WithEvents OK As System.Windows.Forms.Button
    Friend WithEvents Cancel As System.Windows.Forms.Button

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Login))
        Me.LogoPictureBox = New System.Windows.Forms.PictureBox()
        Me.UsernameLabel = New System.Windows.Forms.Label()
        Me.PasswordLabel = New System.Windows.Forms.Label()
        Me.Tbx_Login_username = New System.Windows.Forms.TextBox()
        Me.Tbx_login_password = New System.Windows.Forms.TextBox()
        Me.OK = New System.Windows.Forms.Button()
        Me.Cancel = New System.Windows.Forms.Button()
        Me.Cmx_Login_Database = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Chbx_Login_Save_Password = New System.Windows.Forms.CheckBox()
        CType(Me.LogoPictureBox, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'LogoPictureBox
        '
        Me.LogoPictureBox.Image = CType(resources.GetObject("LogoPictureBox.Image"), System.Drawing.Image)
        Me.LogoPictureBox.InitialImage = CType(resources.GetObject("LogoPictureBox.InitialImage"), System.Drawing.Image)
        Me.LogoPictureBox.Location = New System.Drawing.Point(24, 68)
        Me.LogoPictureBox.Name = "LogoPictureBox"
        Me.LogoPictureBox.Size = New System.Drawing.Size(294, 163)
        Me.LogoPictureBox.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.LogoPictureBox.TabIndex = 0
        Me.LogoPictureBox.TabStop = False
        '
        'UsernameLabel
        '
        Me.UsernameLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UsernameLabel.Location = New System.Drawing.Point(344, 11)
        Me.UsernameLabel.Name = "UsernameLabel"
        Me.UsernameLabel.Size = New System.Drawing.Size(165, 20)
        Me.UsernameLabel.TabIndex = 0
        Me.UsernameLabel.Text = "Gebruikersnaam"
        Me.UsernameLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'PasswordLabel
        '
        Me.PasswordLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PasswordLabel.Location = New System.Drawing.Point(344, 68)
        Me.PasswordLabel.Name = "PasswordLabel"
        Me.PasswordLabel.Size = New System.Drawing.Size(165, 20)
        Me.PasswordLabel.TabIndex = 2
        Me.PasswordLabel.Text = "Wachtwoord"
        Me.PasswordLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Tbx_Login_username
        '
        Me.Tbx_Login_username.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Tbx_Login_username.Location = New System.Drawing.Point(344, 34)
        Me.Tbx_Login_username.Name = "Tbx_Login_username"
        Me.Tbx_Login_username.Size = New System.Drawing.Size(271, 28)
        Me.Tbx_Login_username.TabIndex = 1
        '
        'Tbx_login_password
        '
        Me.Tbx_login_password.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Tbx_login_password.Location = New System.Drawing.Point(344, 91)
        Me.Tbx_login_password.Name = "Tbx_login_password"
        Me.Tbx_login_password.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.Tbx_login_password.Size = New System.Drawing.Size(271, 28)
        Me.Tbx_login_password.TabIndex = 3
        '
        'OK
        '
        Me.OK.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.OK.Location = New System.Drawing.Point(344, 234)
        Me.OK.Name = "OK"
        Me.OK.Size = New System.Drawing.Size(123, 45)
        Me.OK.TabIndex = 4
        Me.OK.Text = "Inloggen"
        '
        'Cancel
        '
        Me.Cancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Cancel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Cancel.Location = New System.Drawing.Point(494, 234)
        Me.Cancel.Name = "Cancel"
        Me.Cancel.Size = New System.Drawing.Size(123, 45)
        Me.Cancel.TabIndex = 5
        Me.Cancel.Text = "Annuleren"
        '
        'Cmx_Login_Database
        '
        Me.Cmx_Login_Database.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Cmx_Login_Database.FormattingEnabled = True
        Me.Cmx_Login_Database.Items.AddRange(New Object() {"Productie", "Acceptatie", "Test"})
        Me.Cmx_Login_Database.Location = New System.Drawing.Point(344, 196)
        Me.Cmx_Login_Database.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Cmx_Login_Database.Name = "Cmx_Login_Database"
        Me.Cmx_Login_Database.Size = New System.Drawing.Size(268, 30)
        Me.Cmx_Login_Database.TabIndex = 6
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(344, 171)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(165, 20)
        Me.Label1.TabIndex = 7
        Me.Label1.Text = "Database"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Chbx_Login_Save_Password
        '
        Me.Chbx_Login_Save_Password.AutoSize = True
        Me.Chbx_Login_Save_Password.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.875!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Chbx_Login_Save_Password.Location = New System.Drawing.Point(344, 125)
        Me.Chbx_Login_Save_Password.Name = "Chbx_Login_Save_Password"
        Me.Chbx_Login_Save_Password.Size = New System.Drawing.Size(205, 24)
        Me.Chbx_Login_Save_Password.TabIndex = 8
        Me.Chbx_Login_Save_Password.Text = "Wachtwoord onthouden"
        Me.Chbx_Login_Save_Password.UseVisualStyleBackColor = True
        '
        'Login
        '
        Me.AcceptButton = Me.OK
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Cancel
        Me.ClientSize = New System.Drawing.Size(635, 303)
        Me.Controls.Add(Me.Chbx_Login_Save_Password)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Cmx_Login_Database)
        Me.Controls.Add(Me.Cancel)
        Me.Controls.Add(Me.OK)
        Me.Controls.Add(Me.Tbx_login_password)
        Me.Controls.Add(Me.Tbx_Login_username)
        Me.Controls.Add(Me.PasswordLabel)
        Me.Controls.Add(Me.UsernameLabel)
        Me.Controls.Add(Me.LogoPictureBox)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Login"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Login"
        CType(Me.LogoPictureBox, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Cmx_Login_Database As ComboBox
    Friend WithEvents Label1 As Label
    Friend WithEvents Chbx_Login_Save_Password As CheckBox
End Class
