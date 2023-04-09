<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Frm_Bank_Select_Type
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

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Frm_Bank_Select_Type))
        Me.Rbn_Bank_Contract = New System.Windows.Forms.RadioButton()
        Me.Rbn_Bank_Extra = New System.Windows.Forms.RadioButton()
        Me.Rbn_Bank_Other = New System.Windows.Forms.RadioButton()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Btn_Bank_Save_Accounts = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'Rbn_Bank_Contract
        '
        Me.Rbn_Bank_Contract.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.Rbn_Bank_Contract.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.125!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Rbn_Bank_Contract.Location = New System.Drawing.Point(15, 60)
        Me.Rbn_Bank_Contract.Margin = New System.Windows.Forms.Padding(4)
        Me.Rbn_Bank_Contract.Name = "Rbn_Bank_Contract"
        Me.Rbn_Bank_Contract.Size = New System.Drawing.Size(371, 31)
        Me.Rbn_Bank_Contract.TabIndex = 46
        Me.Rbn_Bank_Contract.Text = "Contractbetaling"
        Me.Rbn_Bank_Contract.UseVisualStyleBackColor = True
        '
        'Rbn_Bank_Extra
        '
        Me.Rbn_Bank_Extra.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.Rbn_Bank_Extra.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.125!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Rbn_Bank_Extra.Location = New System.Drawing.Point(15, 89)
        Me.Rbn_Bank_Extra.Margin = New System.Windows.Forms.Padding(4)
        Me.Rbn_Bank_Extra.Name = "Rbn_Bank_Extra"
        Me.Rbn_Bank_Extra.Size = New System.Drawing.Size(256, 32)
        Me.Rbn_Bank_Extra.TabIndex = 46
        Me.Rbn_Bank_Extra.Text = "Extra gift"
        Me.Rbn_Bank_Extra.UseVisualStyleBackColor = True
        '
        'Rbn_Bank_Other
        '
        Me.Rbn_Bank_Other.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.Rbn_Bank_Other.Checked = True
        Me.Rbn_Bank_Other.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.125!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Rbn_Bank_Other.Location = New System.Drawing.Point(15, 115)
        Me.Rbn_Bank_Other.Margin = New System.Windows.Forms.Padding(4)
        Me.Rbn_Bank_Other.Name = "Rbn_Bank_Other"
        Me.Rbn_Bank_Other.Size = New System.Drawing.Size(281, 34)
        Me.Rbn_Bank_Other.TabIndex = 46
        Me.Rbn_Bank_Other.TabStop = True
        Me.Rbn_Bank_Other.Text = "Anders"
        Me.Rbn_Bank_Other.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.125!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(12, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(191, 17)
        Me.Label1.TabIndex = 47
        Me.Label1.Text = "Kies het type banktransactie:"
        '
        'Btn_Bank_Save_Accounts
        '
        Me.Btn_Bank_Save_Accounts.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.Btn_Bank_Save_Accounts.Image = Global.SPAS.My.Resources.Resources.Save_16x16
        Me.Btn_Bank_Save_Accounts.Location = New System.Drawing.Point(381, 99)
        Me.Btn_Bank_Save_Accounts.Margin = New System.Windows.Forms.Padding(4)
        Me.Btn_Bank_Save_Accounts.Name = "Btn_Bank_Save_Accounts"
        Me.Btn_Bank_Save_Accounts.Size = New System.Drawing.Size(49, 50)
        Me.Btn_Bank_Save_Accounts.TabIndex = 48
        Me.Btn_Bank_Save_Accounts.UseVisualStyleBackColor = True
        '
        'Frm_Bank_Select_Type
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.ClientSize = New System.Drawing.Size(443, 169)
        Me.Controls.Add(Me.Btn_Bank_Save_Accounts)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Rbn_Bank_Contract)
        Me.Controls.Add(Me.Rbn_Bank_Extra)
        Me.Controls.Add(Me.Rbn_Bank_Other)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Frm_Bank_Select_Type"
        Me.Text = "Type transacties"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Rbn_Bank_Contract As RadioButton
    Friend WithEvents Rbn_Bank_Extra As RadioButton
    Friend WithEvents Rbn_Bank_Other As RadioButton
    Friend WithEvents Label1 As Label
    Friend WithEvents Btn_Bank_Save_Accounts As Button
End Class
